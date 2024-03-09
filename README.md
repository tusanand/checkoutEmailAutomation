# Checkout Email Automation by Tushar Anand

This script automates the process of sending checkout emails, providing a streamlined approach for your daily tasks. To use this script effectively, follow the guidelines below:

## Instructions:

1. **Folder Structure:**
   - Keep this script file and an updated Excel template (with a proper name) in the same folder.
   - Example: If today's date is February 16th, the Excel file should be named `Tushar_2023-02-14_Checkout` (any date that is behind today's date).

2. **Code Customization:**
   - Fill in the necessary details in the code to personalize it for your use.
   - Replace `<your first name>` with your actual first name.
   - Update email addresses, names, and the email body as per your requirements.

3. **Python Dependencies:**
   - Make sure to install all Python dependencies on your local machine. You can install them using:
     ```bash
     pip install openpyxl
     ```

4. **Running the Script:**
   - Execute the script by running the following command in your terminal or command prompt:
     ```bash
     python checkout_template.py
     ```

## Important Notes:

- The script automatically identifies the Excel file for the previous day, renames it, and attaches it to an email for sending.
- Ensure the correct file path is set in the `directory` variable.
- Set up your email server details and credentials for successful email sending.

## Script Details:

- The script utilizes Python, openpyxl, and smtplib for Excel file manipulation and email functionality.
- It automates the process of renaming the Excel file and sending it as an attachment to specified recipients.

Feel free to reach out if you have any questions or need further assistance.

Thank you,
Tushar Anand
