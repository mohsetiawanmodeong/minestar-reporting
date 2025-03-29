# Excel Data Cleaner for Mining Operations

This web application processes raw mining operation data Excel files and cleanses them according to specific formatting requirements.

## Features

- Converts machine data to a standardized format
- Reformats date/time information to MM/DD/YYYY H:MM:SS
- Converts duration from seconds to hours (2 decimal places)
- Categorizes delay types based on prefixes:
  - D- → DELAY
  - S- → STANDBY
  - UX- → UNPLANNED DOWN
  - X- → PLANNED DOWN
  - XX- → EXTENDED LOSS

## Installation

1. Clone or download this repository
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

1. Start the application:

```bash
python app.py
```

2. Open a web browser and go to `http://127.0.0.1:5000`
3. Upload your raw Excel file
4. The application will process the file and return a cleaned version for download

## Input Format

The application expects an Excel file with columns similar to:
- Machine
- Start Date (in format like "Fri Mar 28 07:25:25 WIT 2025")
- Finish Date (in format like "Fri Mar 28 10:04:30 WIT 2025")
- Duration (in format like "9,545 s")
- Delay Type (with prefixes like "D-", "S-", etc.)
- Description

## Output Format

The processed Excel file will have columns:
- Unit (from Machine)
- Start (reformatted as MM/DD/YYYY H:MM:SS)
- Finish (reformatted as MM/DD/YYYY H:MM:SS)
- Dur (duration in hours, 2 decimal places)
- Desc (description)
- Delay Type (preserved from input)
- Category (determined based on Delay Type prefix) 