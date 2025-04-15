import os
import shutil  # ← used for file copying
import pandas as pd
from fpdf import FPDF
import yagmail
import getpass


def create_employees_folder():
    folder_name = "Employees"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    return folder_name


def generate_pdf(employee, folder):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Salary Payslip", ln=True, align='C')
    pdf.ln(10)

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
        print(f"[✓] Email sent to {name} at {to_email}")
    except Exception as e:
        print(f"[!] Failed to send email to {name}: {e}")


def main():
    # Prompt sender for Gmail credentials
    sender_email = input("Enter your Gmail address: ")
    sender_password = getpass.getpass("Enter your Gmail App Password (not your Gmail password): ")

    # Read employee Excel file
    file_path = input("Enter path to the employee Excel file: ")
    try:
        df = pd.read_excel(file_path)
        print(f"[✓] Loaded: {file_path}")
    except Exception as e:
        print(f"[!] Error reading Excel file: {e}")
        return

    # Initialize Yagmail
    try:
        yag = yagmail.SMTP(sender_email, sender_password)
    except Exception as e:
        print(f"[!] Login failed: {e}")
        return

    employees_folder = create_employees_folder()
    
    # Copy Excel file using cross-platform method
    db_copy_path = os.path.join(employees_folder, "Database.xlsx")
    try:
        shutil.copy(file_path, db_copy_path)
        print(f"[✓] Database copied to: {db_copy_path}")
    except Exception as e:
        print(f"[!] Failed to copy database: {e}")

    for _, row in df.iterrows():
        employee = row.to_dict()
        full_name = f"{employee['NAME']} {employee['SURNAME']}"
        
        # Save PDF
        pdf_path = generate_pdf(employee, employees_folder)

        # Send email
        send_email(employee['E-mail'], full_name, pdf_path, yag)


if __name__ == "__main__":
    main()