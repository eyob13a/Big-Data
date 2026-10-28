import matplotlib.pyplot as plt
import matplotlib.patches as patches
from pathlib import Path

reports_dir = Path(r"E:\BigData_Traffic_Accident_Project\reports")
reports_dir.mkdir(parents=True, exist_ok=True)

# 1. Architecture Diagram
fig, ax = plt.subplots(figsize=(14, 7), dpi=300)
ax.set_xlim(0, 14)
ax.set_ylim(0, 7)
ax.axis("off")

# Styles
box_style_src = dict(boxstyle="round,pad=0.5", facecolor="#E3F2FD", edgecolor="#1565C0", linewidth=2)
box_style_proc = dict(boxstyle="round,pad=0.5", facecolor="#FFF3E0", edgecolor="#E65100", linewidth=2)
box_style_db = dict(boxstyle="round,pad=0.5", facecolor="#E8F5E9", edgecolor="#2E7D32", linewidth=2)
box_style_anal = dict(boxstyle="round,pad=0.5", facecolor="#F3E5F5", edgecolor="#6A1B9A", linewidth=2)
box_style_aux = dict(boxstyle="round,pad=0.4", facecolor="#ECEFF1", edgecolor="#455A64", linewidth=1.5, linestyle="--")

# Draw Boxes
# Source
ax.text(1.5, 4.5, "Source Dataset\n\nUS_Accidents_March23.csv\n(2.85 GB, 7.73M Rows)\nCSV / Flat File", 
        ha="center", va="center", fontsize=9, fontweight="bold", bbox=box_style_src)

# Extraction & Chunking
ax.text(4.5, 4.5, "Extraction Layer\n\nPandas pd.read_csv()\nchunksize = 10,000\nlow_memory = False\n773 Total Chunks", 
        ha="center", va="center", fontsize=9, fontweight="bold", bbox=box_style_proc)

# Transformation & Cleaning
ax.text(7.5, 4.5, "Transformation Layer\n\n- Header strip & Column mapping\n- ID validation & In-chunk dedup\n- Type coercion (datetime, float)\n- Boolean mapping (True/False)\n- String trimming & Empty->NA\n- Severity filtering [1..4]", 
        ha="center", va="center", fontsize=8.5, fontweight="bold", bbox=box_style_proc)

# Loading Layer
ax.text(10.5, 4.5, "Loading Layer\n\nSQLAlchemy + psycopg\nDataFrame.to_sql()\nchunksize = 500\nPer-Chunk Transaction\n(engine.begin())", 
        ha="center", va="center", fontsize=9, fontweight="bold", bbox=box_style_db)

# Database
ax.text(10.5, 1.8, "PostgreSQL Database\n\nDatabase: traffic_accidents_db\nTable: traffic_accidents\nRows Loaded: 7,728,394\nPrimary Key: id (VARCHAR 30)", 
        ha="center", va="center", fontsize=9, fontweight="bold", bbox=box_style_db)

# Analysis Layer
ax.text(4.5, 1.8, "SQL Analysis Layer\n\nAggregations & Analytical Queries\n- Severity distribution\n- State frequency rankings\n- Weather condition impact\n- Temporal/Monthly trends", 
        ha="center", va="center", fontsize=9, fontweight="bold", bbox=box_style_anal)

# Cross-cutting concerns
ax.text(1.5, 1.8, "Cross-Cutting System Services\n\n- Logging (etl_pipeline.log & Console)\n- Config Management (.env via python-dotenv)\n- Empty-table Pre-check Guard\n- Engine Resource Disposal", 
        ha="center", va="center", fontsize=8.5, fontweight="bold", bbox=box_style_aux)

# Arrows
arrow_kwargs = dict(arrowstyle="->", lw=2, color="#37474F")
# Source -> Extraction
ax.annotate("", xy=(3.3, 4.5), xytext=(2.7, 4.5), arrowprops=arrow_kwargs)
# Extraction -> Transformation
ax.annotate("", xy=(6.2, 4.5), xytext=(5.7, 4.5), arrowprops=arrow_kwargs)
# Transformation -> Loading
ax.annotate("", xy=(9.3, 4.5), xytext=(8.8, 4.5), arrowprops=arrow_kwargs)
# Loading -> Database
ax.annotate("", xy=(10.5, 2.7), xytext=(10.5, 3.5), arrowprops=arrow_kwargs)
# Database -> SQL Analysis
ax.annotate("", xy=(5.9, 1.8), xytext=(9.1, 1.8), arrowprops=arrow_kwargs)

# Dotted feedback / control lines
dash_arrow = dict(arrowstyle="->", lw=1.5, color="#78909C", linestyle=":")
ax.annotate("", xy=(4.5, 3.6), xytext=(2.6, 2.3), arrowprops=dash_arrow)
ax.annotate("", xy=(7.5, 3.5), xytext=(2.6, 2.2), arrowprops=dash_arrow)
ax.annotate("", xy=(9.4, 3.7), xytext=(2.6, 2.0), arrowprops=dash_arrow)

plt.title("System Architecture: Big Data ETL Pipeline for US Traffic Accidents", fontsize=14, fontweight="bold", pad=15)
plt.tight_layout()
arch_img_path = reports_dir / "system_architecture.png"
plt.savefig(arch_img_path, dpi=300)
plt.close()
print("Architecture diagram saved to:", arch_img_path)

# 2. Severity Chart
fig, ax = plt.subplots(figsize=(8, 4.5), dpi=300)
severities = ['Severity 1', 'Severity 2', 'Severity 3', 'Severity 4']
counts = [67366, 6156981, 1299337, 204710]
colors = ['#4CAF50', '#2196F3', '#FF9800', '#F44336']
bars = ax.bar(severities, counts, color=colors, width=0.55, edgecolor='#333333', linewidth=1)
ax.set_ylabel('Number of Accident Records', fontsize=11, fontweight='bold')
ax.set_title('Accident Records Distribution by Severity Level (Total: 7,728,394)', fontsize=12, fontweight='bold', pad=12)
ax.set_yscale('log')
ax.grid(axis='y', linestyle='--', alpha=0.7)

for bar in bars:
    height = bar.get_height()
    ax.annotate(f'{height:,}\n({height/sum(counts)*100:.1f}%)',
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 4), textcoords="offset points",
                ha='center', va='bottom', fontsize=9, fontweight='bold')

plt.tight_layout()
sev_img_path = reports_dir / "severity_distribution.png"
plt.savefig(sev_img_path, dpi=300)
plt.close()
print("Severity chart saved to:", sev_img_path)

# 3. Top 10 States Chart
fig, ax = plt.subplots(figsize=(9, 5), dpi=300)
states = ['California (CA)', 'Florida (FL)', 'Texas (TX)', 'South Carolina (SC)', 'New York (NY)', 
          'North Carolina (NC)', 'Virginia (VA)', 'Pennsylvania (PA)', 'Minnesota (MN)', 'Oregon (OR)']
state_counts = [1741433, 880192, 582837, 382557, 347960, 338199, 303301, 296620, 192084, 179660]
y_pos = range(len(states))
ax.barh(y_pos, state_counts[::-1], color='#3F51B5', height=0.65, edgecolor='#1A237E')
ax.set_yticks(y_pos)
ax.set_yticklabels(states[::-1], fontsize=10)
ax.set_xlabel('Number of Accident Records', fontsize=11, fontweight='bold')
ax.set_title('Top 10 US States by Recorded Accident Count', fontsize=12, fontweight='bold', pad=12)
ax.grid(axis='x', linestyle='--', alpha=0.7)

for i, v in enumerate(state_counts[::-1]):
    ax.text(v + 25000, i, f'{v:,}', va='center', fontsize=9, fontweight='bold', color='#1A237E')

ax.set_xlim(0, 2000000)
plt.tight_layout()
states_img_path = reports_dir / "top_states_distribution.png"
plt.savefig(states_img_path, dpi=300)
plt.close()
print("Top states chart saved to:", states_img_path)

# 4. Top Weather Conditions Chart
fig, ax = plt.subplots(figsize=(9, 5), dpi=300)
weather = ['Fair', 'Mostly Cloudy', 'Cloudy', 'Clear', 'Partly Cloudy', 
           'Overcast', 'Light Rain', 'Scattered Clouds', 'Light Snow', 'Fog']
weather_counts = [2560802, 1016195, 817082, 808743, 698972, 382866, 352957, 204829, 128680, 99238]
y_pos = range(len(weather))
ax.barh(y_pos, weather_counts[::-1], color='#009688', height=0.65, edgecolor='#004D40')
ax.set_yticks(y_pos)
ax.set_yticklabels(weather[::-1], fontsize=10)
ax.set_xlabel('Number of Accident Records', fontsize=11, fontweight='bold')
ax.set_title('Top 10 Recorded Weather Conditions in Accident Dataset', fontsize=12, fontweight='bold', pad=12)
ax.grid(axis='x', linestyle='--', alpha=0.7)

for i, v in enumerate(weather_counts[::-1]):
    ax.text(v + 35000, i, f'{v:,}', va='center', fontsize=9, fontweight='bold', color='#004D40')

ax.set_xlim(0, 2900000)
plt.tight_layout()
weather_img_path = reports_dir / "top_weather_distribution.png"
plt.savefig(weather_img_path, dpi=300)
plt.close()
print("Top weather chart saved to:", weather_img_path)

