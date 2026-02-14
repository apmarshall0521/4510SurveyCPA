import pandas as pd
import plotly.express as px
import os
import re

# Ensure output directory exists
os.makedirs('public', exist_ok=True)

# Load data
df = pd.read_excel('Grad Program Exit Survey Data 2024.xlsx', header=1)

# Remove metadata row (index 0 in the dataframe loaded with header=1)
# The first row contains import IDs like {"ImportId":"QID124"}
df = df.iloc[1:]

# Identify columns
core_cols = []
elective_cols = []

for col in df.columns:
    if "Please place each MAcc CORE course into rank order" in col:
        core_cols.append(col)
    elif "Rate " in col and "on a scale from 1-5" in col:
        elective_cols.append(col)

# Helper function to extract course name
def extract_course_name(col_name):
    # Split by " - " and take the last part
    parts = col_name.split(' - ')
    if len(parts) > 1:
        return parts[-1].strip()
    return col_name

# Process Core Courses (Rankings)
core_data = []
for col in core_cols:
    course_name = extract_course_name(col)
    # Convert to numeric, errors='coerce' turns non-numeric to NaN
    s = pd.to_numeric(df[col], errors='coerce')
    mean_rank = s.mean()
    core_data.append({'Course': course_name, 'Mean Rank': mean_rank, 'Type': 'Core'})

core_df = pd.DataFrame(core_data).sort_values('Mean Rank', ascending=True)
core_df['Rank'] = range(1, len(core_df) + 1)

# Save processed Core data
core_df.to_csv('public/core_courses_ranking.csv', index=False)

# Process Elective Courses (Ratings)
elective_data = []
for col in elective_cols:
    course_name = extract_course_name(col)
    s = pd.to_numeric(df[col], errors='coerce')
    mean_rating = s.mean()
    elective_data.append({'Course': course_name, 'Mean Rating': mean_rating, 'Type': 'Elective'})

elective_df = pd.DataFrame(elective_data).sort_values('Mean Rating', ascending=False)
elective_df['Rank'] = range(1, len(elective_df) + 1)

# Save processed Elective data
elective_df.to_csv('public/elective_courses_rating.csv', index=False)

# Visualization
# Core Chart (Lower is better, but bar charts usually show magnitude. Maybe invert axis?)
# Or just show the value.
fig_core = px.bar(core_df, x='Course', y='Mean Rank', title='Core Course Rankings (Lower is Better)',
                  text_auto='.2f', color='Mean Rank', color_continuous_scale='Viridis_r')
fig_core.update_layout(yaxis=dict(autorange="reversed")) # Invert y-axis so Rank 1 is at top

# Save Core Chart
fig_core.write_html('public/core_chart.html')

# Elective Chart (Higher is better)
fig_elective = px.bar(elective_df, x='Course', y='Mean Rating', title='Elective Course Ratings (Scale 1-5)',
                      text_auto='.2f', color='Mean Rating', color_continuous_scale='Viridis')
fig_elective.update_layout(yaxis=dict(range=[0, 5]))

# Save Elective Chart
fig_elective.write_html('public/elective_chart.html')

# Generate HTML
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>MAcc Program Course Rankings</title>
    <style>
        body {{ font-family: sans-serif; margin: 20px; }}
        h1 {{ text-align: center; }}
        .main-wrapper {{ display: flex; flex-wrap: wrap; justify-content: space-around; gap: 20px; }}
        .section {{ width: 45%; min-width: 400px; box-sizing: border-box; padding: 10px; border: 1px solid #eee; border-radius: 8px; }}
        .chart {{ width: 100%; margin-bottom: 20px; }}
        table {{ border-collapse: collapse; width: 100%; font-size: 0.9em; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        h2 {{ text-align: center; }}
        .download-links {{ text-align: center; margin-top: 10px; margin-bottom: 20px; }}
        .download-links a {{ margin: 0 10px; text-decoration: none; color: #007bff; }}
        .download-links a:hover {{ text-decoration: underline; }}
    </style>
</head>
<body>
    <h1>MAcc Program Course Rankings 2024</h1>

    <div class="main-wrapper">
        <div class="section">
            <h2>Core Courses (Ranked by Preference)</h2>
            <div class="chart">
                {fig_core.to_html(full_html=False, include_plotlyjs='cdn')}
            </div>
            <div class="download-links">
                <a href="core_courses_ranking.csv" download>Download Data (CSV)</a> |
                <a href="core_chart.html" download>Download Chart (HTML)</a>
            </div>
            {core_df[['Rank', 'Course', 'Mean Rank']].to_html(index=False, float_format='%.2f')}
        </div>

        <div class="section">
            <h2>Elective Courses (Rated 1-5)</h2>
            <div class="chart">
                {fig_elective.to_html(full_html=False, include_plotlyjs='cdn')}
            </div>
            <div class="download-links">
                <a href="elective_courses_rating.csv" download>Download Data (CSV)</a> |
                <a href="elective_chart.html" download>Download Chart (HTML)</a>
            </div>
            {elective_df[['Rank', 'Course', 'Mean Rating']].to_html(index=False, float_format='%.2f')}
        </div>
    </div>
</body>
</html>
"""

# Write to file
with open('public/index.html', 'w') as f:
    f.write(html_content)

print("Successfully generated public/index.html")
