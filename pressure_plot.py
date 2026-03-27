import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Parameter: how many months until a point is fully faded (transparent)
fade_months = 12.0

# Read the Excel file
df = pd.read_excel(
    'ciśnienie.xlsx', 
    sheet_name='Pomiary', 
    usecols="E:I",   # Read only columns E to I
    skiprows=4,      # Start reading at row 5 (header row)
)
df = df.drop(df.columns[1], axis=1).iloc[3:].reset_index(drop=True)
df.columns = ['Date', 'Time', 'SYS', 'DIA']
df['Date'] = df['Date'].ffill()
df = df.infer_objects(copy=False).dropna(subset=['Time', 'SYS', 'DIA'])

# Combine Date and Time columns into a datetime column
df['datetime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))

# Determine the most recent measurement datetime
most_recent = df['datetime'].max()

# Compute the time difference (in days) and convert to months (approximate: 30 days per month)
df['days_diff'] = (most_recent - df['datetime']).dt.days
df['months_diff'] = df['days_diff'] / 30.0

# Normalize the age value to a [0, 1] scale for coloring: 0 is most recent, 1 is older than fade_months
df['age_norm'] = (df['months_diff'] / fade_months).clip(upper=1)

# Set baseline axis limits for DIA (x-axis) and SYS (y-axis)
x_data = df['DIA']
x_lower, x_upper = (30, 110) if x_data.min() >= 30 and x_data.max() <= 110 else (min(30, x_data.min()), max(110, x_data.max()))

y_data = df['SYS']
y_lower, y_upper = (40, 180) if y_data.min() >= 40 and y_data.max() <= 180 else (min(40, y_data.min()), max(180, y_data.max()))

# Create the plot
fig, ax = plt.subplots(figsize=(8, 6))

# Draw the background colored rectangles in descending order.
ax.add_patch(patches.Rectangle((0, 0), x_upper, y_upper, facecolor='red', zorder=1))  # (5) Red
ax.add_patch(patches.Rectangle((0, 0), 100, 160, facecolor='orange', zorder=2))      # (4) Orange
ax.add_patch(patches.Rectangle((0, 0), 90, 140, facecolor='yellow', zorder=3))      # (3) Yellow
ax.add_patch(patches.Rectangle((0, 0), 85, 130, facecolor='green', zorder=4))      # (2) Green
ax.add_patch(patches.Rectangle((0, 0), 65, 110, facecolor='blue', zorder=5))       # (1) Blue

# Get viridis colormap and apply to edge colors
cmap = plt.cm.gist_heat
colors = cmap(df['age_norm'])

# Scatter plot using open circles with colormap on edges
for i in range(len(df)):
    ax.scatter(df['DIA'].iloc[i], df['SYS'].iloc[i], 
               s=80, facecolors='none', edgecolors=[colors[i]], linewidth=1.5, zorder=6)

# Create a colorbar manually using a ScalarMappable
import matplotlib.cm as cm
import matplotlib.colors as mcolors
norm = mcolors.Normalize(vmin=0, vmax=1)
sm = cm.ScalarMappable(cmap=cmap, norm=norm)
sm.set_array([])
cbar = plt.colorbar(sm, ax=ax)
cbar.set_label("Czas pomiaru [0 - aktualny; 1 - >=12 miesięcy temu]")

# Set axis limits based on the computed ranges
ax.set_xlim(x_lower, x_upper)
ax.set_ylim(y_lower, y_upper)

# Label axes and add a title
ax.set_xlabel('ROZKURCZOWE (DIA)')
ax.set_ylabel('SKURCZOWE (SYS)')
ax.set_title('Ciśnienie krwi')

# Save the plot as a PDF file
plt.savefig('blood_pressure_plot.pdf', format='pdf')
plt.close()

print("Plot saved as 'blood_pressure_plot.pdf'.")
