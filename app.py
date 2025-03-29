from flask import Flask, render_template, request, send_file
import pandas as pd
import numpy as np
import os
import io
import re
from datetime import datetime

app = Flask(__name__)

def clean_excel_data(df, data_type='delay'):
    """
    Clean Excel data based on the data type (delay or cycle)
    """
    if data_type == 'delay':
        return clean_delay_data(df)
    elif data_type == 'cycle':
        return clean_cycle_data(df)
    else:
        raise ValueError("Invalid data type. Must be 'delay' or 'cycle'")

def clean_delay_data(df):
    """
    Clean delay data from the Excel file
    """
    # Create a new DataFrame with desired structure
    cleaned_df = pd.DataFrame(columns=[
        'Unit', 'Start', 'Finish', 'Dur', 'Desc', 'Delay Type', 'Category'
    ])
    
    # Map columns from original data to new format
    if 'Machine' in df.columns:
        cleaned_df['Unit'] = df['Machine']
    elif 'Unit' in df.columns:
        cleaned_df['Unit'] = df['Unit']
    
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
    elif 'Desc' in df.columns:
        cleaned_df['Desc'] = df['Desc']
    
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
    elif 'Category' in df.columns:
        cleaned_df['Category'] = df['Category']
    
    return cleaned_df

def clean_cycle_data(df):
    """
    Clean cycle data from Excel file
    """
    # Create a new DataFrame that preserves original columns
    cleaned_df = pd.DataFrame()
    
    # Keep original Unit column
    if 'Unit' in df.columns:
        cleaned_df['Unit'] = df['Unit']
    
    # Keep original Operator column
    if 'Operator' in df.columns:
        cleaned_df['Operator'] = df['Operator']
    
    # Parse and store start time for duration calculation
    start_times = []
    
    # Handle Start time column, formatting to match delay data format
    if 'Start time' in df.columns:
        def format_date(date_str):
            if pd.isna(date_str):
                return None
            
            try:
                # Extract date components from strings like "Thu Mar 27 22:12:49 WIT 2025"
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
                    formatted_date = f"{month_num}/{day}/{year} {time}"
                    
                    # Parse for duration calculation
                    try:
                        dt_obj = pd.to_datetime(formatted_date)
                        start_times.append(dt_obj)
                        return formatted_date
                    except:
                        start_times.append(None)
                        return formatted_date
                
                # If pattern doesn't match, try using pandas
                date = pd.to_datetime(date_str)
                start_times.append(date)
                return date.strftime('%m/%d/%Y %H:%M:%S')
            except:
                start_times.append(None)
                return date_str
            
        cleaned_df['Start'] = df['Start time'].apply(format_date)
    elif 'Start' in df.columns:  # Fallback for different column name
        def format_start_date(x):
            if pd.notna(x):
                try:
                    date = pd.to_datetime(x)
                    start_times.append(date)
                    return date.strftime('%m/%d/%Y %H:%M:%S')
                except:
                    start_times.append(None)
                    return x
            else:
                start_times.append(None)
                return None
                
        cleaned_df['Start'] = df['Start'].apply(format_start_date)
    
    # Parse and store finish time for duration calculation
    finish_times = []
    
    # Handle Finish Time column
    if 'Finish Time' in df.columns:
        def format_date(date_str):
            if pd.isna(date_str):
                return None
            
            try:
                # Extract date components from strings like "Fri Mar 28 03:06:06 WIT 2025"
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
                    formatted_date = f"{month_num}/{day}/{year} {time}"
                    
                    # Parse for duration calculation
                    try:
                        dt_obj = pd.to_datetime(formatted_date)
                        finish_times.append(dt_obj)
                        return formatted_date
                    except:
                        finish_times.append(None)
                        return formatted_date
                
                # If pattern doesn't match, try using pandas
                date = pd.to_datetime(date_str)
                finish_times.append(date)
                return date.strftime('%m/%d/%Y %H:%M:%S')
            except:
                finish_times.append(None)
                return date_str
            
        cleaned_df['Finish'] = df['Finish Time'].apply(format_date)
    elif 'Finish' in df.columns:  # Fallback for different column name
        def format_finish_date(x):
            if pd.notna(x):
                try:
                    date = pd.to_datetime(x)
                    finish_times.append(date)
                    return date.strftime('%m/%d/%Y %H:%M:%S')
                except:
                    finish_times.append(None)
                    return x
            else:
                finish_times.append(None)
                return None
                
        cleaned_df['Finish'] = df['Finish'].apply(format_finish_date)
    
    # Calculate duration as difference between finish and start times (in hours)
    if len(start_times) == len(finish_times) and len(start_times) > 0:
        durations = []
        for i in range(len(start_times)):
            start = start_times[i]
            finish = finish_times[i]
            
            if pd.notna(start) and pd.notna(finish):
                try:
                    # Calculate duration in hours
                    duration_seconds = (finish - start).total_seconds()
                    duration_hours = duration_seconds / 3600
                    durations.append(round(duration_hours, 2))
                except:
                    # If calculation fails, use original Dur if available
                    if 'Dur' in df.columns and i < len(df['Dur']):
                        try:
                            durations.append(round(float(str(df['Dur'].iloc[i]).replace(',', '.')), 2))
                        except:
                            durations.append(df['Dur'].iloc[i] if i < len(df['Dur']) else None)
                    else:
                        durations.append(None)
            else:
                # If either start or finish is missing, use original Dur if available
                if 'Dur' in df.columns and i < len(df['Dur']):
                    try:
                        durations.append(round(float(str(df['Dur'].iloc[i]).replace(',', '.')), 2))
                    except:
                        durations.append(df['Dur'].iloc[i] if i < len(df['Dur']) else None)
                else:
                    durations.append(None)
        
        cleaned_df['Dur'] = durations
    else:
        # Process duration (ensure 2 decimal places)
        if 'Dur' in df.columns:
            def format_duration(dur_val):
                if pd.isna(dur_val):
                    return None
                    
                try:
                    # Convert to float and round to 2 decimal places
                    return round(float(str(dur_val).replace(',', '.')), 2)
                except:
                    return dur_val
            
            cleaned_df['Dur'] = df['Dur'].apply(format_duration)
    
    # Keep original Source column
    if 'Source' in df.columns:
        cleaned_df['Source'] = df['Source']
    
    # Keep original Destination column
    if 'Destination' in df.columns:
        cleaned_df['Destination'] = df['Destination']
    
    # Handle description column if present
    if 'Desc' in df.columns:
        cleaned_df['Desc'] = df['Desc']
    
    # Handle Delay Type and Category if present
    if 'Delay Type' in df.columns:
        cleaned_df['Delay Type'] = df['Delay Type']
    
    if 'Category' in df.columns:
        cleaned_df['Category'] = df['Category']
    
    return cleaned_df

def process_excel_file(file):
    """
    Process Excel file with multiple sheets
    """
    # Dictionary to store sheet data
    result_dict = {}
    
    try:
        # Check if the file has 'delay' and 'cycle' sheets
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        
        print(f"Sheets dalam file: {sheet_names}")  # Debug
        
        # Process all sheets
        for sheet_name in sheet_names:
            # Read the sheet data
            df = pd.read_excel(file, sheet_name=sheet_name)
            
            print(f"Kolom pada sheet {sheet_name}: {df.columns.tolist()}")  # Debug
            
            # Determine data type based on sheet name or column structure
            if sheet_name.lower() == "delay" or ('Start Date' in df.columns and 'Finish Date' in df.columns and 
                ('Machine' in df.columns or 'Unit' in df.columns) and 'Duration' in df.columns):
                data_type = 'delay'
                print(f"Sheet {sheet_name} terdeteksi sebagai delay")  # Debug
            elif sheet_name.lower() == "cycle" or (
                'Unit' in df.columns and 
                ('Start time' in df.columns or 'Start' in df.columns) and 
                ('Finish Time' in df.columns or 'Finish' in df.columns) and
                'Dur' in df.columns):
                data_type = 'cycle'
                print(f"Sheet {sheet_name} terdeteksi sebagai cycle")  # Debug
            else:
                # Skip sheets that don't match expected formats
                print(f"Sheet {sheet_name} tidak terdeteksi sebagai format yang valid")  # Debug
                continue
            
            # Clean the data based on type
            result_dict[sheet_name] = clean_excel_data(df, data_type)
    
    except Exception as e:
        print(f"Error saat memproses file: {str(e)}")  # Debug
        raise Exception(f"Error processing file: {str(e)}")
    
    print(f"Hasil sheets yang diproses: {list(result_dict.keys())}")  # Debug
    return result_dict

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
                # Process the Excel file with multiple sheets
                cleaned_data = process_excel_file(file)
                
                if not cleaned_data:
                    return render_template('index.html', error='No valid data sheets found in the file')
                
                # Create a new Excel file with the cleaned data
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet_name, df in cleaned_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
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

@app.route('/download-template')
def download_template():
    """
    Endpoint untuk mengunduh file template
    """
    template_path = os.path.join(os.getcwd(), 'Default Format Cycle & Delay.xlsx')
    
    try:
        return send_file(
            template_path,
            as_attachment=True,
            download_name='Template Cycle & Delay.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return render_template('index.html', error=f'Error downloading template: {str(e)}')

if __name__ == '__main__':
    # Ensure directories exist
    os.makedirs('static', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000) 