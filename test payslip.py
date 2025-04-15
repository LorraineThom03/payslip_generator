import pandas as pd
import yagmail


def send_email(to_email, name):
    try:
        yag = yagmail.SMTP("your-email@gmail.com", "your-password")
        subject = "Salary Collection Notification"
        body = f"Dear {name},\n\nYour salary is ready for collection.\n\nPlease visit the HR department to collect your salary.\n\nRegards,\nHR Team"
        yag.send(to=to_email, subject=subject, contents=body)
        print(f"Email sent to {name} at {to_email}")
    except Exception as e:
        print(f"Failed to send email to {name}: {e}")

def main():
    # Call the send_email function here
    send_email("recipient-email@gmail.com", "Recipient Name")

if __name__ == "__main__":
    main()

  
def main():
    file_path = input("employee xlsx: ")
    try:
        df = pd.read_excel(file_path)
        # Process the DataFrame
        print(df.head())
    except Exception as e:
        print(f"Error reading Excel file: {e}")

if __name__ == "__main__":
    main()
