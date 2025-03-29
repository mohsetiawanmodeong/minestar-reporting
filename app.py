from flask import Flask, render_template, request, send_file
import pandas as pd
import numpy as np
import os
import io
import re
from datetime import datetime

app = Flask(__name__)

def clean_excel_data(df):
    # Create a new DataFrame with desired structure
    cleaned_df = pd.DataFrame(columns=[
        'Unit', 'Start', 'Finish', 'Dur', 'Desc', 'Delay Type', 'Category'
    ])
    
    # Map columns from original data to new format
    if 'Machine' in df.columns:
        cleaned_df['Unit'] = df['Machine']
    
    # Handle date/time columns
    if 'Start Date' in df.columns and 'Finish Date' in df.columns:
        # Helper function to parse WIT date format
        def parse_wit_date(date_str):
            if pd.isna(date_str):
                return None
                
            # Extract date components from strings like "Fri Mar 28 07:25:25 WIT 2025"
            pattern = r'(\w{3}) (\w{3}) (\d{1,2}) (\d{2}:\d{2}:\d{2}) WIT (\d{4})'
            match = re.match(pattern, str(date_str))
            
            if match:
                _, month, day, time, year = match.groups()
                
                # Convert month name to number
                month_map = {
                    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
                }
                month_num = month_map.get(month, 1)
                
                # Format as 'MM/DD/YYYY H:MM:SS'
                return f"{month_num}/{day}/{year} {time}"
            return None
        
        # Apply the parsing function
        cleaned_df['Start'] = df['Start Date'].apply(parse_wit_date)
        cleaned_df['Finish'] = df['Finish Date'].apply(parse_wit_date)
    
    # Process duration (convert from seconds to hours with 2 decimal places)
    if 'Duration' in df.columns:
        def convert_duration(dur_str):
            if pd.isna(dur_str):
                return None
                
            # Handle format like "9,545 s"
            match = re.match(r'([\d,]+\.?\d*)\s*s', str(dur_str))
            if match:
                # Remove commas and convert to float
                seconds = float(match.group(1).replace(',', ''))
                # Convert to hours
                return round(seconds / 3600, 2)
            
            # Try direct numeric conversion if no unit
            try:
                return round(float(str(dur_str).replace(',', '')) / 3600, 2)
            except:
                return None
        
        cleaned_df['Dur'] = df['Duration'].apply(convert_duration)
    
    # Copy description
    if 'Description' in df.columns:
        cleaned_df['Desc'] = df['Description']
    
    # Copy delay type
    if 'Delay Type' in df.columns:
        cleaned_df['Delay Type'] = df['Delay Type']
    
    # Determine category based on delay type
    if 'Delay Type' in df.columns:
        def get_category(delay_type):
            if pd.isna(delay_type):
                return None
                
            delay_type = str(delay_type)
            if delay_type.startswith('D-'):
                return 'DELAY'
            elif delay_type.startswith('S-'):
                return 'STANDBY'
            elif delay_type.startswith('UX-'):
                return 'UNPLANNED DOWN'
            elif delay_type.startswith('X-'):
                return 'PLANNED DOWN'
            elif delay_type.startswith('XX-'):
                return 'EXTENDED LOSS'
            return None
        
        cleaned_df['Category'] = df['Delay Type'].apply(get_category)
    
    return cleaned_df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        
        file = request.files['file']
        
        if file.filename == '':
            return render_template('index.html', error='No selected file')
        
        if file and file.filename.endswith(('.xlsx', '.xls')):
            try:
                # Read the Excel file
                df = pd.read_excel(file)
                
                # Process the data
                cleaned_df = clean_excel_data(df)
                
                # Save to a BytesIO object
                output = io.BytesIO()
                cleaned_df.to_excel(output, index=False, engine='openpyxl')
                output.seek(0)
                
                # Return the processed file
                return send_file(
                    output,
                    as_attachment=True,
                    download_name='cleaned_data.xlsx',
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            
            except Exception as e:
                return render_template('index.html', error=f'Error processing file: {str(e)}')
        
        else:
            return render_template('index.html', error='File must be an Excel file (.xlsx or .xls)')
    
    return render_template('index.html')

if __name__ == '__main__':
    # Ensure directories exist
    os.makedirs('static', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000) 