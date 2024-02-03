from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import re
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

def convert_to_mm_dd_yyyy(date_str):
    try:
        return pd.to_datetime(date_str, errors='coerce').strftime('%m/%d/%Y')
    except:
        return date_str

@app.route('/upload', methods=['POST'])
def upload_dirty_excel():
    if 'file' not in request.files:
        return render_template('index.html', message="No file selected. Please choose a dirty Excel file.")

    file = request.files['file']
    if file.filename == '':
        return render_template('index.html', message="No file selected. Please choose a dirty Excel file.")

    if file:
        df = pd.read_excel(file)

        # Check for missing commas in any cell
        df = df.applymap(lambda x: str(x).replace(',', '') if pd.notnull(x) else x)

        df = df.dropna(subset=['BRAC PIN'])

        # Remove special characters and numbers from the first and last name columns
        df['First Name'] = df['First Name'].apply(lambda x: re.sub('[^A-Za-z]+', '', str(x)))
        df['Last Name'] = df['Last Name'].apply(lambda x: re.sub('[^A-Za-z]+', '', str(x)))

        # Correct data of wrong city name and wrong country name columns
        # Assuming 'City' and 'Country' are the column names
        df['City'] = df['City'].replace({'OldCityName': 'CorrectCityName'})
        df['Country'] = df['Country'].replace({'OldCountryName': 'CorrectCountryName'})

        # Remove 'enter' from any cell data
        df = df.applymap(lambda x: str(x).replace('\n', '') if pd.notnull(x) else x)

        # Apply the conversion function to the 'Date of Birth (mm/dd/yyyy)' column
        df['Date of Birth (mm/dd/yyyy)'] = df['Date of Birth (mm/dd/yyyy)'].apply(convert_to_mm_dd_yyyy)

        # Find and remove duplicate data
        df.drop_duplicates(inplace=True)

        df['National Identification Number/Passport'] = df['National Identification Number/Passport'].astype(str)

        # Save the cleaned data to a CSV file
        cleaned_csv = BytesIO()
        df.to_csv(cleaned_csv, index=False)
        cleaned_csv.seek(0)

        return send_file(cleaned_csv, download_name='cleaned_data.csv', as_attachment=True)

    return render_template('index.html', message="Error processing the file.")

if __name__ == '__main__':
    app.run(debug=True)
