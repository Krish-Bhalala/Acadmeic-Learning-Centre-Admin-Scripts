### `USER GUIDE`

```markdown
# User Guide for Deployment and Usage

## Overview
This user guide will help you set up, configure, and run the Academic-Learning-Centre-Admin-Scripts project smoothly in different environments. It covers installation steps, secure credential management, automated scheduling (e.g., weekly emailing), and best practices.

## Prerequisites

- **Python 3.x** installed on your system.
- **Excel Template:** An Excel timesheet template compatible with the script’s data formatting logic.
- **Access Credentials:**  
  - Tutor ID and Employee Endpoint from your academic learning centre’s scheduling system.
  - SMTP email credentials (if you plan to use the email-sending feature).

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/Krish-Bhalala/Academic-Learning-Centre-Admin-Scripts.git
   cd Academic-Learning-Centre-Admin-Scripts
   ```

2. **Install Dependencies:**
   ```bash
   pip install openpyxl beautifulsoup4 requests
   ```
   
   *Note:* For additional security and best practices, consider using a virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install openpyxl beautifulsoup4 requests
   ```

## Configuration

1. **Timesheet Paths:**
   - `directory_path`: Path to the folder containing the timesheet template.
   - `file_path`: Full path to the timesheet Excel file.

2. **Tutor and Endpoint Details:**
   - `TUTOR_ID`: Replace with your actual tutor ID.
   - `EMPLOYEE_ENDPOINT`: Endpoint URL for fetching the schedule data.

3. **Email Credentials:**
   - Update `email_sender.py` with your SMTP server details, email address, and a secure way to handle your email password.
   - For production, it’s recommended to use environment variables instead of hardcoding passwords.  
     For example:
     ```bash
     export EMAIL_USER="your_email@example.com"
     export EMAIL_PASS="some_secure_password"
     ```
     Then in your code:
     ```python
     import os
     sender_email = os.environ.get("EMAIL_USER")
     sender_password = os.environ.get("EMAIL_PASS")
     ```

4. **SSL Verification:**
   - Currently, `ssl._create_default_https_context = ssl._create_unverified_context` is used for development.
   - In production, remove this line or configure proper SSL certificates to ensure secure communication.

## Running the Script

1. **One-Time Run:**
   ```bash
   python timesheet.py
   ```
   This will:
   - Fetch last week’s appointment data.
   - Update the Excel timesheet.
   - Save a new Excel file.

   If configured, it will also send an email on the designated day (e.g., Mondays).

2. **Scheduled Runs:**
   To fully automate the weekly process, consider scheduling the script with:
   - **Cron (Linux/macOS):**
     ```bash
     crontab -e
     ```
     Add a line to run every Monday at 9 AM:
     ```bash
     0 9 * * 1 /usr/bin/python3 /path/to/Academic-Learning-Centre-Admin-Scripts/timesheet.py
     ```
   - **Windows Task Scheduler:**
     - Open Task Scheduler
     - Create a new basic task and choose weekly triggers.
     - Point it to your Python executable and `timesheet.py`.

3. **Verifying the Results:**
   - Check the output Excel file in the specified directory.
   - Verify that the email was received by the admin office.
   - If something isn’t right, use print/debug statements or logging to inspect variables.

## Secure Credential Handling

- **Environment Variables:**  
  Storing credentials in environment variables is more secure than hardcoding them.
  
- **Password Managers or Secret Stores:**  
  Consider using tools like `aws ssm`, HashiCorp Vault, or GitHub Actions secrets (for CI/CD) if you’re deploying this in a more controlled environment.

## Updating the Codebase

- **Refactoring Code:**
  If you plan on modifying the code (e.g., changing the date calculation logic or the email format):
  - Make changes in a feature branch.
  - Run the script locally to ensure correctness.
  - Possibly write automated tests (if time allows).

- **Coding Style & Linters:**
  Consider using a linter (e.g., `flake8`) or autoformatter (`black`) to maintain consistent coding style:
  ```bash
  pip install flake8 black
  black .
  flake8 .
  ```

## Troubleshooting

- **No Appointments Fetched:**  
  Check that `TUTOR_ID` and `EMPLOYEE_ENDPOINT` are correct.
  
- **Email Not Sent:**
  - Verify SMTP credentials and server details.
  - Check if you’re running on the designated day of the week (if you implemented weekly sending).

- **File Not Found Errors:**
  Ensure `file_path` and `directory_path` are set correctly and that the Excel template is present.

## Future Improvements

- Integration with a CI/CD pipeline for automated testing and deployments.
- Advanced logging instead of just `print` statements.
- More robust error handling and retry logic for network operations.

---

### Commit Message Example

```
docs: Add comprehensive USER_GUIDE.md for deployment and usage

- Introduced a detailed user guide covering setup, configuration, and running instructions
- Added guidance on environment variables, scheduling the script, and secure credential handling
- Provided best practices for code maintenance and troubleshooting steps
```

This commit message communicates that you’ve added a detailed user guide for deployment, addressing task #4 in the future features issue.

### Pull Request Description Example

**Title:**
"Add Comprehensive User Guide for Deployment and Usage"

**Description:**
This PR adds a `USER_GUIDE.md` providing detailed instructions on how to deploy and use the Academic-Learning-Centre-Admin-Scripts. It covers installation steps, environment configuration, scheduling for weekly automation, and secure credential management. It also offers best practices for code maintenance and troubleshooting tips. This addresses **task #4** in [Issue #1](https://github.com/Krish-Bhalala/Academic-Learning-Centre-Admin-Scripts/issues/1), making the project more accessible and easier to adopt.

**Included Changes:**
- `USER_GUIDE.md` file with comprehensive documentation.
- References to environment variables for secure credential handling.
- Steps for setting up automated tasks with cron and Windows Task Scheduler.

**Testing Steps:**
- Review the `USER_GUIDE.md` content and follow the instructions to run the script in a test environment.
- Verify that all necessary details (e.g., environment setup, scheduling, credentials) are covered.
