from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import os

app = Flask(__name__)

# Path to store the Excel file
EXCEL_FILE = 'data/employee_details.xlsx'

def ensure_directory_exists(file_path):
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

def save_to_excel(data, file_path):
    df = pd.DataFrame(data, columns=['Employee Name', 'Department Name', 'Date of Birth', 'Email ID', 'Mobile Number'])
    df.to_excel(file_path, index=False)

def read_from_excel(file_path):
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        return df.values.tolist()
    else:
        return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    department = request.form['department']
    dob = request.form['dob']
    email = request.form['email']
    mobile = request.form['mobile']

    # Ensure the directory exists
    ensure_directory_exists(EXCEL_FILE)

    # Read existing data
    existing_data = read_from_excel(EXCEL_FILE)

    # Append new data
    existing_data.append([name, department, dob, email, mobile])

    # Save updated data to Excel
    save_to_excel(existing_data, EXCEL_FILE)

    return redirect(url_for('index'))

@app.route('/download')
def download():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name='employee_details.xlsx')
    else:
        return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True)
