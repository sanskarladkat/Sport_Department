# Add 'request' to handle URL parameters
from flask import Flask, jsonify, render_template, request
import pandas as pd

app = Flask(__name__)

@app.route('/')
def home():
    """Serves the main dashboard page."""
    return render_template('dashboard.html')

@app.route('/api/data')
def get_data():
    """Reads raw data, applies filters, and calculates all metrics."""
    try:
        excel_filename = 'New XLSX Worksheet.xlsx' 
        sheet_name = 'Sheet1'
        df = pd.read_excel(excel_filename, sheet_name=sheet_name)
        
        id_column = 'SR. NO'
        id_col = 'NAME OF STUDENT'

        # Clean key columns
        df[id_column] = df[id_column].astype(str).str.strip()
        df[id_col] = df[id_col].astype(str).str.strip()
        df['Sport'] = df['Sport'].str.strip().str.title()
        df['GENDER'] = df['GENDER'].str.strip().str.title()
        df['School'] = df['School'].str.strip()

        # --- CROSS-FILTERING LOGIC ---
        filter_gender = request.args.get('GENDER')
        filter_school = request.args.get('School')
        
        if filter_gender:
            df = df[df['GENDER'] == filter_gender]
        if filter_school:
            df = df[df['School'] == filter_school]
        
        # --- All calculations below are based on the filtered DataFrame ---
        
        kpi_metrics = {
            'totalAchievements': len(df),
            'totalAthletes': df[id_col].nunique(),
            'totalPoints': int(df['POINT'].sum()) if 'POINT' in df.columns and len(df) > 0 else 0,
            'uniqueSports': df['Sport'].nunique()
        }
        
        # --- Prepare data for charts ---
        
        school_counts = df['School'].value_counts().reset_index()
        school_counts.columns = ['School', 'Achievements']
        school_data = school_counts.to_dict(orient='records')

        unique_athletes = df.drop_duplicates(subset=[id_column])
        gender_counts = unique_athletes['GENDER'].value_counts().reset_index()
        gender_counts.columns = ['Gender', 'Count']
        gender_data = {
            'labels': gender_counts['Gender'].tolist(),
            'series': gender_counts['Count'].astype(int).tolist()
        }

        # Data for Top 5 Achievement Types Bar Chart
        df['Achievement_Type'] = df['RESULTS']
        achievement_counts = df['Achievement_Type'].value_counts().reset_index()
        achievement_counts.columns = ['Type', 'Count']
        achievement_data_bar = achievement_counts.head(5).to_dict(orient='records')
        
        # NEW: Data for Achievement Types Pie Chart (uses ALL types)
        achievement_data_pie = {
            'labels': achievement_counts['Type'].tolist(),
            'series': achievement_counts['Count'].astype(int).tolist()
        }
        
        # Data for Top 6 Popular Sports Bar Chart
        popular_sports_counts = df.groupby('Sport')[id_column].nunique().reset_index()
        popular_sports_counts.columns = ['Sport', 'Participants']
        popular_sports_counts = popular_sports_counts.sort_values(by='Participants', ascending=False)
        popular_sports_data_bar = popular_sports_counts.head(6).to_dict(orient='records')

        # NEW: Data for All Sports Pie Chart (uses ALL sports)
        sports_data_pie = {
            'labels': popular_sports_counts['Sport'].tolist(),
            'series': popular_sports_counts['Participants'].astype(int).tolist()
        }
        
        # (Sport by Gender data preparation remains the same)
        sport_gender_pivot = df.pivot_table(index='Sport', columns='GENDER', values=id_column, aggfunc='nunique').fillna(0)
        for gender_col in ['Boys', 'Girls']:
            if gender_col not in sport_gender_pivot.columns: sport_gender_pivot[gender_col] = 0
        sport_gender_pivot['Total'] = sport_gender_pivot.get('Boys', 0) + sport_gender_pivot.get('Girls', 0)
        sport_gender_pivot = sport_gender_pivot.sort_values(by='Total', ascending=False).drop(columns=['Total']).reset_index()
        sport_by_gender_data = {
            'categories': sport_gender_pivot['Sport'].tolist(),
            'series': [
                {'name': 'Boys', 'data': sport_gender_pivot.get('Boys', pd.Series(0, index=sport_gender_pivot.index)).astype(int).tolist()},
                {'name': 'Girls', 'data': sport_gender_pivot.get('Girls', pd.Series(0, index=sport_gender_pivot.index)).astype(int).tolist()}
            ]
        }

        dashboard_data = {
            'kpiMetrics': kpi_metrics,
            'schoolParticipation': school_data,
            'genderDistribution': gender_data,
            'achievementTypesBar': achievement_data_bar, # Renamed for clarity
            'achievementTypesPie': achievement_data_pie, # New
            'popularSportsBar': popular_sports_data_bar, # Renamed for clarity
            'sportsPie': sports_data_pie, # New
            'sportByGender': sport_by_gender_data
        }
        
        return jsonify(dashboard_data)

    except Exception as e:
        error_msg = f"An error occurred: {e}"
        print(error_msg)
        return jsonify({"error": error_msg}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')