# Lead Data Cleaner

A Streamlit web application for cleaning and processing lead data from various sources.

## Features

- Upload Excel (.xlsx) or CSV files
- Clean data from Outscraper (support for other tools planned)
- Remove unnecessary columns
- Adjust column widths for better readability
- Download processed files

## Usage

1. Upload your lead data file (Excel or CSV)
2. Select the lead generation tool used (currently only Outscraper)
3. Click "Process" to clean the data
4. Download the processed file

## Requirements

- Python 3.x
- Streamlit
- openpyxl
- pandas

## Installation

1. Clone this repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the app: `streamlit run app.py`

## Note

Make sure you have the necessary permissions to read and write files in the `Data/Upload` and `Data/Output` directories.