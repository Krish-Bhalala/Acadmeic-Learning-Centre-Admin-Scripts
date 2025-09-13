from datetime import datetime, timedelta

from datetime import datetime, timedelta

def emailDataGenerator(receiver, subject, sender):
    today = datetime.now()
    current_date = today.strftime("%Y-%m-%d")

    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    current_week = f"{monday.strftime('%A %b %d, %Y')} to {sunday.strftime('%A %b %d, %Y')}"

    salutations = f"Warm regards,\n{sender}\nYour Contact Info Here"

    # You can modify the body to mention that the timesheet is attached.
    body = f"Hello,\n\nPlease find attached the updated timesheet for the week:\n{current_week}."

    # Return structure: [receiver email, subject, body, closing, sender email]
    return [receiver, subject, body, salutations, sender]
