import pandas as pd
import plotly.express as px
import os
import re

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

# Process Elective Courses (Ratings)
elective_data = []
for col in elective_cols:
    course_name = extract_course_name(col)
    s = pd.to_numeric(df[col], errors='coerce')
    mean_rating = s.mean()
    elective_data.append({'Course': course_name, 'Mean Rating': mean_rating, 'Type': 'Elective'})

elective_df = pd.DataFrame(elective_data).sort_values('Mean Rating', ascending=False)
elective_df['Rank'] = range(1, len(elective_df) + 1)

# Visualization
# Combine for a single chart? The scales are different.
# I'll make two subplots using plotly.graph_objects or just two separate charts in HTML.
# But plotly express is easier. Let's make a faceted chart? No, y-axis units differ.
# I'll create two bar charts and combine them in HTML.

# Core Chart (Lower is better, but bar charts usually show magnitude. Maybe invert axis?)
# Or just show the value.
fig_core = px.bar(core_df, x='Course', y='Mean Rank', title='Core Course Rankings (Lower is Better)',
                  text_auto='.2f', color='Mean Rank', color_continuous_scale='Viridis_r')
fig_core.update_layout(yaxis=dict(autorange="reversed")) # Invert y-axis so Rank 1 is at top

# Elective Chart (Higher is better)
fig_elective = px.bar(elective_df, x='Course', y='Mean Rating', title='Elective Course Ratings (Scale 1-5)',
                      text_auto='.2f', color='Mean Rating', color_continuous_scale='Viridis')
fig_elective.update_layout(yaxis=dict(range=[0, 5]))

# Generate HTML
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>MAcc Program Course Rankings</title>
    <style>
        body {{ font-family: sans-serif; margin: 20px; }}
        h1 {{ text-align: center; }}
        .container {{ display: flex; flex-wrap: wrap; justify-content: space-around; }}
        .chart {{ width: 100%; max-width: 800px; margin-bottom: 40px; }}
        table {{ border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px auto; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
</head>
<body>
    <h1>MAcc Program Course Rankings 2024</h1>

    <div class="container">
        <div class="chart">
            {fig_core.to_html(full_html=False, include_plotlyjs='cdn')}
        </div>
        <div>
            <h2>Core Courses (Ranked by Preference)</h2>
            {core_df[['Rank', 'Course', 'Mean Rank']].to_html(index=False, float_format='%.2f')}
        </div>
    </div>

    <hr>

    <div class="container">
        <div class="chart">
            {fig_elective.to_html(full_html=False, include_plotlyjs='cdn')}
        </div>
        <div>
            <h2>Elective Courses (Rated 1-5)</h2>
            {elective_df[['Rank', 'Course', 'Mean Rating']].to_html(index=False, float_format='%.2f')}
        </div>
    </div>
</body>
</html>
"""

# Ensure output directory exists
os.makedirs('public', exist_ok=True)

# Write to file
with open('public/index.html', 'w') as f:
    f.write(html_content)

print("Successfully generated public/index.html")
