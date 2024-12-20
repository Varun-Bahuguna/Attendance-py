from flask import Flask, render_template, request, redirect, url_for, send_file, send_from_directory

import openpyxl
import os
from datetime import datetime
import qrcode  # Import qrcode for QR code generation
from io import BytesIO  # Import BytesIO for in-memory image handling




# Define the Flask app instance
app = Flask(__name__)

FILE_NAME = os.path.join(os.getcwd(), 'attendance.xlsx')

# Check if the Excel file exists, if not create one
if not os.path.exists(FILE_NAME):
    print("Excel file does not exist. Creating a new one.")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Attendance"
    sheet.append(["Student ID", "Session ID", "Subject", "Date", "Timestamp"])  # Header row
    workbook.save(FILE_NAME)
else:
    print("Excel file already exists.")

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate_qr')
def generate_qr():
    data = "https://attendance-py.onrender.com/"  # Data to encode in the QR code
    img = qrcode.make(data)  # Generate the QR code

    # Save the QR code image to the same directory as app.py
    qr_file_path = os.path.join(os.getcwd(), "generated_qr.png")
    img.save(qr_file_path)  # Save the image to disk
    print(f"QR code saved at: {qr_file_path}")

    return redirect(url_for('serve_qr_code'))  # Redirect to serve the QR code

@app.route('/qr_code')
def serve_qr_code():
    # Send the QR code from the correct directory
    return send_from_directory(os.getcwd(), "generated_qr.png") 

@app.route('/download/<filename>')
def download_file(filename):
    # Return files like attendance.xlsx from the working directory
    directory = os.getcwd()  # Current working directory
    return send_from_directory(directory, filename, as_attachment=True)

    
@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    if request.method == 'POST':
        student_id = request.form['student_id']
        session_id = request.form['session_id']
        subject = request.form['subject']
        date = request.form['date']

        # Record the attendance in the Excel file
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"Adding student ID: {student_id} to the attendance sheet.")
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook.active
        sheet.append([student_id, session_id, subject, date, timestamp])  # Append data
        workbook.save(FILE_NAME)
        print(f"Attendance for {student_id} saved.")

        return redirect(url_for('attendance_success'))


@app.route('/attendance_success')
def attendance_success():
    return render_template('success.html')

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))  # Get the PORT from environment or default to 5000
    app.run(host='0.0.0.0', port=port, debug=True)  # Bind to 0.0.0.0 to allow external access
