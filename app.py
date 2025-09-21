# Add 'request' to handle URL parameters
from flask import Flask, jsonify, render_template, request
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import traceback

app = Flask(__name__)

@app.route('/')
def home():
    """Serves the main dashboard page."""
    return render_template('dashboard.html')

@app.route('/api/data')
def get_data():
    """Reads raw data from GOOGLE SHEETS, applies filters, and calculates all metrics."""
    try:
        # --- START: MODIFIED SECTION TO READ FROM GOOGLE SHEETS ---
        
        # Define the scope of the APIs
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        
        # Add credentials to the account
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        
        # Authorize the clientsheet 
        client = gspread.authorize(creds)
        
        # The name of your Google Sheet
        sheet_name = 'Sport_Achievers' 
        
        # Get the sheet
        sheet = client.open(sheet_name).sheet1
        
        # Get all of the records from the sheet and convert to a pandas DataFrame
        data = sheet.get_all_records()
        df = pd.DataFrame(data)

        # --- END: MODIFIED SECTION ---


        # vvv ALL OF YOUR EXISTING LOGIC BELOW THIS LINE REMAINS UNCHANGED vvv
        
        id_column = 'SR. NO'
        id_col = 'NAME OF STUDENT'

        # Clean key columns (ensure your Google Sheet columns are named correctly)
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
            'totalPoints': int(pd.to_numeric(df['POINT'], errors='coerce').sum()) if 'POINT' in df.columns and len(df) > 0 else 0,
            'uniqueSports': df['Sport'].nunique()
        }
        
        # (The rest of your file continues here, unchanged...)
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
            'achievementTypesBar': achievement_data_bar,
            'achievementTypesPie': achievement_data_pie,
            'popularSportsBar': popular_sports_data_bar,
            'sportsPie': sports_data_pie,
            'sportByGender': sport_by_gender_data
        }
        
        return jsonify(dashboard_data)

    except Exception as e:
        error_msg = f"An error occurred: {e}"
        print("--- DETAILED ERROR ---")
        traceback.print_exc() # This will print the full error details
        print("----------------------")
        return jsonify({"error": error_msg}), 500

# New endpoint to get student details by sport
@app.route('/api/students_by_sport')
def get_students_by_sport():
    """Filters and returns student details for a given sport."""
    sport_name = request.args.get('sport')

    if not sport_name:
        return jsonify({"error": "Sport name is required"}), 400

    try:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)
        sheet_url = 'https://docs.google.com/spreadsheets/d/1g_Uh2c4WxhjFiMRxcAIlTQ1obgS_ka8R5e8F76ntrgo/edit?gid=512331806#gid=512331806' # PASTE YOUR GOOGLE SHEET URL
        spreadsheet = client.open_by_url(sheet_url)
        sheet = spreadsheet.sheet1
        data = sheet.get_all_records()
        df = pd.DataFrame(data)

        df['Sport'] = df['Sport'].str.strip().str.title()
        df['NAME OF STUDENT'] = df['NAME OF STUDENT'].astype(str).str.strip()
        df['GENDER'] = df['GENDER'].astype(str).str.strip().str.title()
        df['School'] = df['School'].astype(str).str.strip()

        # Filter for the selected sport
        filtered_df = df[df['Sport'] == sport_name].drop_duplicates(subset=['NAME OF STUDENT'])
        
        # Select the columns we want
        student_details = filtered_df[['NAME OF STUDENT', 'GENDER', 'School']]
        
        # Convert the DataFrame to a list of dictionaries and return as JSON
        return jsonify(student_details.to_dict(orient='records'))

    except Exception as e:
        error_msg = f"An error occurred: {e}"
        print(error_msg)
        return jsonify({"error": error_msg}), 500 
    
    

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')