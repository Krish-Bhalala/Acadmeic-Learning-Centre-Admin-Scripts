import os
import gzip
import http.client
import io
import ssl
import urllib.parse
from datetime import datetime, timedelta

import openpyxl
from bs4 import BeautifulSoup

from email_generator import emailDataGenerator  # Import from email_generator.py
from email_sender import sendEmail  # Import from email_sender.py

# Disable SSL verification
ssl._create_default_https_context = ssl._create_unverified_context

# Constants (update these to your actual values)
DIRECTORY_PATH = r"PATH TO FOLDER WITH TIMESHEET.XLXS"
FILE_PATH = r"PATH TO FILE TIMESHEET.XLXS"
TUTOR_ID = r"<ADD YOUR TUTOR ID HERE>"
EMPLOYEE_ENDPOINT = r"<ADD EMPLOYEE ENDPOINT HERE>"

# Configuration
FIRST_ROW = 15
COLUMN_SEQUENCE = "DEFGIK"
TRANSACTION_ID = "1030"
DESCRIPTION = "tutoring"

def calculate_week_range():
    """
    Calculate the start and end date for a timesheet period.
    We assume the timesheet is from Saturday to the following Friday.
    """
    today = datetime.now()
    # Calculate how many days to subtract to get the previous Saturday
    days_to_subtract = (today.weekday() + 2) % 7
    previous_saturday = today - timedelta(days=days_to_subtract)
    following_friday = previous_saturday + timedelta(days=6)
    
    start_date_str = previous_saturday.strftime("%d-%b-%y")
    end_date_str = following_friday.strftime("%d-%b-%y")
    return start_date_str, end_date_str


def fetch_raw_data(tutor_id, start_date, end_date):
    """
    Fetch raw appointment data for a given tutor and date range.
    """
    # Convert to format YYYY-MM-DD
    start_date_obj = datetime.strptime(start_date, "%d-%b-%y").strftime("%Y-%m-%d")
    end_date_obj = datetime.strptime(end_date, "%d-%b-%y").strftime("%Y-%m-%d")

    conn = http.client.HTTPSConnection("manitoba.mywconline.com")
    payload = f'rid%5B%5D={tutor_id}&sdate={start_date_obj}&edate={end_date_obj}'

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/png,image/svg+xml,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://manitoba.mywconline.com',
        'DNT': '1',
        'Sec-GPC': '1',
        'Connection': 'keep-alive',
        'Referer': f'https://manitoba.mywconline.com/{EMPLOYEE_ENDPOINT}?reset=1&sid={tutor_id}',
        'Upgrade-Insecure-Requests': '1',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Priority': 'u=0, i'
    }

    conn.request("POST", EMPLOYEE_ENDPOINT, payload, headers)
    response = conn.getresponse()
    data = response.read()

    # Decompress if needed
    if response.getheader('Content-Encoding') == 'gzip':
        buffer = io.BytesIO(data)
        with gzip.GzipFile(fileobj=buffer) as f:
            decompressed_data = f.read()
    else:
        decompressed_data = data

    # Decode and parse HTML
    try:
        decoded_data = decompressed_data.decode('utf-8')
        soup_data = BeautifulSoup(decoded_data, 'html.parser')
        return soup_data
    except UnicodeDecodeError:
        print("Failed to decode as UTF-8.")
        return None


def parse_raw_data(soup):
    """
    Parse the HTML to extract the appointment names, times, and dates.
    """
    if soup is None:
        return []
        
    names = soup.find_all("p", class_="modal_cursor fw-bold theblue fs-5 m-0")
    names = [[name.next_element] for name in names]

    dates = soup.find_all("span", class_="fw-bold fs-5")
    times = [date.parent.br.next_element for date in dates]
    times = [time.split(" to ") for time in times]

    for i in range(len(dates)):
        temp = dates[i].next_element
        dates[i] = temp.split(", ", maxsplit=1)

    data = format_data(names, times, dates)
    return data


def format_data(names, times, dates):
    """
    Format extracted data into a consistent structure:
    [Date, Start_Time, End_Time, Name_with_Date].
    """
    formatted = []
    for name, time, date in zip(names, times, dates):
        # date[:1] gives the first part of the date (likely the day)
        # time is [start, end]
        # name[0] + " dated: " + date[1] creates the combined name/date string
        formatted.append(date[:1] + time + [name[0] + " dated: " + date[1]])
    return formatted


def edit_row(sheet, row_number, row_data):
    """
    Edit a single row in the Excel sheet using COLUMN_SEQUENCE.
    """
    for i, value in enumerate(row_data):
        cell_id = COLUMN_SEQUENCE[i] + str(row_number)
        sheet[cell_id] = value


def fill_date(sheet, start_date):
    """
    Fill the start date in the specified cell.
    """
    START_DATE_CELL_ID = "F3"
    sheet[START_DATE_CELL_ID] = start_date


def save_workbook(workbook, new_filename):
    """
    Save the workbook to a new file name.
    """
    full_path = os.path.join(DIRECTORY_PATH, new_filename)
    if os.path.exists(full_path):
        print(f"Warning: {full_path} already exists. Choose a different name.")
    else:
        workbook.save(full_path)
        print(f"Workbook saved as {full_path}")


def edit_workbook(data, sheet, workbook, start_date, end_date):
    """
    Edit the workbook with fetched data, then save it.
    """
    current_row = FIRST_ROW
    for element in data:
        # element: [Date, Start_Time, End_Time, Name_with_Date]
        # Insert transaction ID and description between date and times.
        # rows = [Date, TransactionID, Start_Time, End_Time, Description, Name_with_Date]
        row_values = element[:1] + [TRANSACTION_ID] + element[1:3] + [DESCRIPTION] + [element[3]]
        edit_row(sheet, current_row, row_values)
        current_row += 1

    # Fill date in the sheet
    fill_date(sheet, start_date)
    print("Excel file has been edited successfully.")

    # Construct new filename based on the date range
    new_file_name = f"Krish Bhalala Timesheetssss {start_date} to {end_date}.xlsx"
    save_workbook(workbook, new_file_name)
    return new_file_name


def initiate_workbook():
    """
    Load the existing workbook from FILE_PATH.
    """
    try:
        workbook = openpyxl.load_workbook(FILE_PATH)
        return workbook
    except FileNotFoundError:
        print(f"Error: The file at {FILE_PATH} was not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def initiate_sheet(workbook):
    """
    Return the active sheet of the given workbook.
    """
    if workbook is None:
        return None
    try:
        return workbook.active
    except Exception as e:
        print(f"Error accessing active sheet: {e}")
        return None


def send_weekly_email(new_file_name):
    """
    Send the weekly email with the attached updated timesheet, if today is Monday.
    """
    today = datetime.now()
    # Let's say we send the email on Monday (weekday() == 0)
    if today.weekday() == 0:
        receiver = "admin_office@example.com"  # Replace with actual admin office email
        subject = "Weekly Timesheet Update"
        sender = "your_email@example.com"   # Replace with sender's email

        email_data = emailDataGenerator(receiver, subject, sender)
        attachment_path = os.path.join(DIRECTORY_PATH, new_file_name)

        result = sendEmail(email_data, receiverEmail=None, attachment_path=attachment_path)
        print("Email send result:", result)
    else:
        print("Not the scheduled day for sending weekly email. Skipping email send.")


def main():
    start_date, end_date = calculate_week_range()

    workbook = initiate_workbook()
    sheet = initiate_sheet(workbook)
    if sheet is None:
        return

    soup_data = fetch_raw_data(TUTOR_ID, start_date, end_date)
    parsed_data = parse_raw_data(soup_data)

    if not parsed_data:
        print("No data parsed. Exiting.")
        return

    new_file_name = edit_workbook(parsed_data, sheet, workbook, start_date, end_date)

    # Attempt to send weekly email if conditions are met
    send_weekly_email(new_file_name)


if __name__ == "__main__":
    main()
