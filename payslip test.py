import os
import pandas as pd
from fpdf import FPDF
import yagmail
import getpass


def create_employee_folder(employee_name):
    folder_name = employee_name.replace(" ", "_")
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return folder_name


def generate_pdf(employee, folder):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for key, value in employee.items():
        pdf.cell(200, 10, txt=f"{key}: {value}", ln=True)

    pdf_path = os.path.join(folder, f"{employee['NAME']}_{employee['SURNAME']}.pdf")
    pdf.output(pdf_path)
    return pdf_path


def send_email(to_email, name, pdf_path, yag):
    subject = "Salary Collection Notification"
    body = f"""
    Dear {name},

    Your salary is ready for collection.

    Please see the attached salary slip and visit the HR department to collect your salary.

    Regards,
    HR Team
    """
    try:
        yag.send(to=to_email, subject=subject, contents=body, attachments=pdf_path)
        print(f"Email sent to {name} at {to_email}")
    except Exception as e:
        print(f"Failed to send email to {name}: {e}")


def main():
    # Prompt sender for Gmail credentials
    sender_email = input("Enter your Gmail address: ")
    sender_password = getpass.getpass("Enter your Gmail App Password (not your Gmail password): ")

    # Read employee Excel file
    file_path = input("Enter path to the employee Excel file: ")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Initialize Yagmail
    try:
        yag = yagmail.SMTP(sender_email, sender_password)
    except Exception as e:
        print(f"Login failed: {e}")
        return

    for _, row in df.iterrows():
        employee = row.to_dict()
        full_name = f"{employee['NAME']} {employee['SURNAME']}"
        folder = create_employee_folder(full_name)
        
        # Save PDF
        pdf_path = generate_pdf(employee, folder)

        # Copy Excel file into employee folder
        os.system(f'cp "{file_path}" "{folder}/Database.xlsx"')  # Use `copy` for Windows

        # Send email
        send_email(employee['E-mail'], full_name, pdf_path, yag)


if __name__ == "__main__":
    main()
