from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from datetime import datetime
import openpyxl
import os
import pytz  
import uuid
import qrcode  

# Define the Flask app instance
app = Flask(__name__)

FILE_NAME = os.path.join(os.getcwd(), 'attendance.xlsx')

# Set the IST timezone
IST = pytz.timezone('Asia/Kolkata')

# Check if the Excel file exists, if not create one
if not os.path.exists(FILE_NAME):
    print("Excel file does not exist. Creating a new one.")
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Attendance"
    sheet.append(["Student ID", "Session ID", "Subject", "Date", "Timestamp", "IP Address"])  # Added "IP Address" here
    workbook.save(FILE_NAME)
else:
    print("Excel file already exists.")


# Global variable to store current session details
current_session = {
    "subject": "Math",
    "session_id": "Session_101",
    "start_time": "09:00",  # Default start time
    "end_time": "10:00"     # Default end time
}

@app.route('/')
def home():
    global current_session
    current_date = datetime.now(IST).strftime('%Y-%m-%d')  # Use IST for the current date
    return render_template('index.html', current_session=current_session, current_date=current_date)

@app.route('/generate_qr')
def generate_qr():
    data = "https://attendance-pyth.onrender.com/"  # Data to encode in the QR code
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

@app.before_request
def assign_device_id():
    """
    Middleware to assign a unique device ID if not already set.
    """
    if not request.cookies.get('device_id'):
        # Generate a new unique device ID
        device_id = str(uuid.uuid4())
        print(f"Assigning new device ID: {device_id}")
        
        # Set the cookie in the response (done in the main route function)
        request.device_id = device_id  # Temporary storage for this request
    else:
        request.device_id = request.cookies.get('device_id')  # Existing device ID

@app.before_request
def assign_device_id():
    """
    Middleware to assign a unique device ID if not already set.
    """
    if not request.cookies.get('device_id'):
        # Generate a new unique device ID
        device_id = str(uuid.uuid4())
        print(f"Assigning new device ID: {device_id}")
        
        # Set the cookie in the response (done in the main route function)
        request.device_id = device_id  # Temporary storage for this request
    else:
        request.device_id = request.cookies.get('device_id')  # Existing device ID



@app.route('/submit_attendance', methods=['POST'])
def submit_attendance():
    if request.method == 'POST':
        student_id = request.form['student_id']
        session_id = request.form['session_id']
        subject = request.form['subject']
        date = request.form['date']

        # Retrieve the device ID from the request
        device_id = request.device_id

        # Load the workbook and check for duplicate device IDs
        workbook = openpyxl.load_workbook(FILE_NAME)
        sheet = workbook.active

        # Iterate through rows to check if the device ID already exists for the session
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip the header row
            existing_session_id, existing_device_id = row[1], row[5]
            if existing_session_id == session_id and existing_device_id == device_id:
                error_message = "Attendance has already been submitted from this device for the session."
                return render_template('index.html', current_session=current_session, current_date=date, error_message=error_message)

        # Get the current time in IST
        current_time = datetime.now(IST)

        # Convert the session start and end times into datetime objects for comparison (with IST)
        session_start = datetime.strptime(current_session['start_time'], '%H:%M').replace(year=current_time.year, month=current_time.month, day=current_time.day, tzinfo=IST)
        session_end = datetime.strptime(current_session['end_time'], '%H:%M').replace(year=current_time.year, month=current_time.month, day=current_time.day, tzinfo=IST)

        # Check if the current time is within the session's allowed range
        if current_time < session_start:
            error_message = f"Attendance cannot be recorded before the session starts at {current_session['start_time']}."
            return render_template('index.html', current_session=current_session, current_date=date, error_message=error_message)

        if current_time > session_end:
            error_message = f"Attendance cannot be recorded after the session ends at {current_session['end_time']}."
            return render_template('index.html', current_session=current_session, current_date=date, error_message=error_message)

        # If within the time range and no duplicate device ID, proceed with attendance recording
        timestamp = current_time.strftime('%H:%M:%S')  # Include seconds in the timestamp (HH:MM:SS)
        sheet.append([student_id, session_id, subject, date, timestamp, device_id])  # Append data
        workbook.save(FILE_NAME)

        response = redirect(url_for('attendance_success'))

        # Set the device ID cookie in the response
        response.set_cookie('device_id', device_id, max_age=30 * 24 * 60 * 60)  # Cookie expires in 30 days
        return response



@app.route('/set_session', methods=['GET', 'POST'])
def set_session():
    global current_session
    if request.method == 'POST':
        admin_id = request.form['admin_id']
        admin_password = request.form['admin_password']

        valid_id = "admin"
        valid_password = "admin@1234"

        if admin_id == valid_id and admin_password == valid_password:
            # Update session details including start_time and end_time
            current_session['subject'] = request.form['subject']
            current_session['session_id'] = request.form['session_id']
            current_session['start_time'] = request.form['start_time']
            current_session['end_time'] = request.form['end_time']
            
            return redirect(url_for('home'))
        else:
            error_message = "Invalid ID or password. Please try again."
            return render_template('set_session.html', current_session=current_session, error_message=error_message)

    return render_template('set_session.html', current_session=current_session)

@app.route('/attendance_success')
def attendance_success():
    return render_template('success.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Get the PORT from environment or default to 5000
    app.run(host='0.0.0.0', port=port, debug=True)  # Bind to 0.0.0.0 to allow external access
