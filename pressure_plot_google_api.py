import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.cm as cm
import matplotlib.colors as mcolors
import io, os, json
from dotenv import load_dotenv
load_dotenv()

# Biblioteki do Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
CHART_FILE_ID = os.environ.get('CHART_FILE_ID')
CREDS_FILE = os.environ.get('CREDS_FILE')
SHEET_NAME_DATA = 'Pomiary'    # Nazwa zakładki z danymi
SHEET_NAME_CHART = 'Wykresy'  # Nazwa zakładki, gdzie wstawić wykres


# Parameter: how many months until a point is fully faded (transparent)
fade_months = 12.0


def authenticate_google_apis(creds_file):
    """Uwierzytelnia gspread (odczyt) i Drive API (zapis obrazu)."""
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    if creds_json := os.environ.get('GCP_CREDENTIALS'):
        creds_dict = json.loads(creds_json)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)

    gspread_client = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return gspread_client, drive_service

def load_data_from_sheets(gspread_client, ss_id, sheet_name):
    """Pobiera dane z Google Sheets uwzględniając scalenia w wierszach 5-8."""
    sh = gspread_client.open_by_key(ss_id)
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"BŁĄD: Nie znaleziono zakładki o nazwie '{sheet_name}'. Sprawdź wielkość liter.")
        return None
    
    all_values = ws.get_all_values()
    if len(all_values) < 9:
        print("BŁĄD: Arkusz ma za mało wierszy (dane powinny zaczynać się od 9 wiersza).")
        return None
    
    header = all_values[4]
    df = pd.DataFrame(all_values[5:], columns=header).iloc[3:].reset_index(drop=True)
    df = df.iloc[:, [4, 6, 7, 8]] 
    df.columns = ['Date', 'Time', 'SYS', 'DIA']
    
    df['SYS'] = pd.to_numeric(df['SYS'], errors='coerce')
    df['DIA'] = pd.to_numeric(df['DIA'], errors='coerce')
    df['Date'] = df['Date'].replace('', None).ffill()
    df = df.infer_objects(copy=False).dropna(subset=['Time', 'SYS', 'DIA'])
    
    df['datetime'] = pd.to_datetime(
        df['Date'].astype(str) + ' ' + df['Time'].astype(str), 
        dayfirst=True, 
        errors='coerce'
    )
    df = df.dropna(subset=['datetime'])
    
    most_recent = df['datetime'].max()
    df['days_diff'] = (most_recent - df['datetime']).dt.days
    df['months_diff'] = df['days_diff'] / 30.0
    df['age_norm'] = (df['months_diff'] / fade_months).clip(upper=1)
    
    return df

def generate_plot_image(df):
    """Generuje identyczny wykres i zwraca go jako obiekt BytesIO (w pamięci)."""
    
    x_data = df['DIA']
    x_lower, x_upper = (30, 110) if x_data.min() >= 30 and x_data.max() <= 110 else (min(30, x_data.min()), max(110, x_data.max()))
    y_data = df['SYS']
    y_lower, y_upper = (40, 180) if y_data.min() >= 40 and y_data.max() <= 180 else (min(40, y_data.min()), max(180, y_data.max()))

    fig, ax = plt.subplots(figsize=(8, 6))

    ax.add_patch(patches.Rectangle((0, 0), x_upper, y_upper, facecolor='red', zorder=1))
    ax.add_patch(patches.Rectangle((0, 0), 100, 160, facecolor='orange', zorder=2))
    ax.add_patch(patches.Rectangle((0, 0), 90, 140, facecolor='yellow', zorder=3))
    ax.add_patch(patches.Rectangle((0, 0), 85, 130, facecolor='green', zorder=4))
    ax.add_patch(patches.Rectangle((0, 0), 65, 110, facecolor='blue', zorder=5))

    cmap = plt.cm.gist_heat
    colors = cmap(df['age_norm'])

    for i in range(len(df)):
        ax.scatter(df['DIA'].iloc[i], df['SYS'].iloc[i], 
                   s=80, facecolors='none', edgecolors=[colors[i]], linewidth=1.5, zorder=6)

    norm = mcolors.Normalize(vmin=0, vmax=1)
    sm = cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    cbar = plt.colorbar(sm, ax=ax)
    cbar.set_label("Czas pomiaru [0 - aktualny; 1 - >=12 miesięcy temu]")

    ax.set_xlim(x_lower, x_upper)
    ax.set_ylim(y_lower, y_upper)
    ax.set_xlabel('ROZKURCZOWE (DIA)')
    ax.set_ylabel('SKURCZOWE (SYS)')
    ax.set_title('Ciśnienie krwi')

    img_data = io.BytesIO()
    # plt.savefig('blood_pressure_plot-google_api.pdf', format='pdf')  # TEST
    plt.savefig(img_data, format='png', dpi=100, bbox_inches='tight')
    img_data.seek(0)
    plt.close()
    
    return img_data

def upload_image_to_drive(drive_service, img_data):
    """Aktualizuje istniejący plik na Twoim Dysku Google."""
    
    # Przesuwamy wskaźnik na początek danych obrazu
    img_data.seek(0)
    
    # Media do wysłania
    media = MediaIoBaseUpload(img_data, mimetype='image/png', resumable=False)

    try:
        print(f"Aktualizacja pliku o ID: {CHART_FILE_ID}...")
        
        # Używamy metody UPDATE zamiast CREATE. 
        # Ponieważ plik już istnieje i Ty jesteś jego właścicielem, 
        # konto usługi tylko modyfikuje jego treść.
        file = drive_service.files().update(
            fileId=CHART_FILE_ID,
            media_body=media
        ).execute()

        # Upewniamy się, że uprawnienia są publiczne (dla formuły =IMAGE)
        # Robimy to raz, ale dla pewności możemy zostawić
        permission = {'type': 'anyone', 'role': 'reader'}
        try:
            drive_service.permissions().create(fileId=CHART_FILE_ID, body=permission).execute()
        except:
            pass # Jeśli już są nadane, Google rzuci błędem, więc go ignorujemy

        return f'https://drive.google.com/uc?export=view&id={CHART_FILE_ID}'

    except Exception as e:
        print(f"Błąd podczas aktualizacji pliku na Drive: {e}")
        return None

def insert_image_to_sheet(gspread_client, ss_id, sheet_name, image_url):
    """Wstawia obraz do komórki w Sheets za pomocą formuły =IMAGE."""
    sh = gspread_client.open_by_key(ss_id)
    
    # Sprawdź czy arkusz istnieje, jak nie, utwórz
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=20, cols=10)

    # Wstaw formułę IMAGE w komórkę P3
    # Tryb 4 pozwala na customowy rozmiar, ale Sheets często 
    # resetuje rozmiar komórki. Najbezpieczniej tryb 1 lub 2.
    # Tryb 2: dopasuj do rozmiaru komórki, zachowując proporcje
    formula = f'=IMAGE("{image_url}"; 2)'
    
    ws.update_acell('P3', formula)
    
    # Opcjonalnie: powiększ rządek i kolumnę, żeby wykres był widoczny
    ws.format("P3", {"numberFormat": {"type": "TEXT"}}) # zapobiega traktowaniu jako daty
    # Gspread nie ma super łatwej metody na resize, trzeba by użyć raw batchUpdate API
    # Ale formatowanie można zrobić ręcznie w Sheets raz.

# ==========================================
# GŁÓWNY PROGRAM
# ==========================================
if __name__ == '__main__':
    print("1. Autoryzacja w Google Cloud...")
    g_client, d_service = authenticate_google_apis(CREDS_FILE)
    
    print(f"2. Pobieranie danych z arkusza '{SHEET_NAME_DATA}'...")
    try:
        df_blood = load_data_from_sheets(g_client, SPREADSHEET_ID, SHEET_NAME_DATA)
    except Exception as e:
        print(f"Błąd podczas pobierania danych. Sprawdź ID arkusza i czy udostępniłeś go mailowi z JSONa. Błąd: {e}")
        exit()

    print("3. Generowanie wykresu w pamięci (Matplotlib)...")
    chart_image_storage = generate_plot_image(df_blood)
    
    print("4. Uploadowanie obrazu na tymczasowy Google Drive...")
    direct_image_url = upload_image_to_drive(d_service, chart_image_storage)
    
    print(f"5. Aktualizacja formuły IMAGE w arkuszu '{SHEET_NAME_CHART}'...")
    insert_image_to_sheet(g_client, SPREADSHEET_ID, SHEET_NAME_CHART, direct_image_url)
    
    print("GOTOWE. Wykres powinien być widoczny w Google Sheets.")
    # Uwaga: Pliki na drive konta usługi powinny być okresowo czyszczone, 
    # ale na potrzeby uruchamiania ręcznego to nie problem.