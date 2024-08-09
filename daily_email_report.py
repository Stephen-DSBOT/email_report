import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import pyodbc

# Define email credentials and recipient
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = input("email address")
EMAIL_PASSWORD = input("password")
RECIPIENT_EMAIL = input("receiving email")

# Define your database connection parameters
DB_SERVER = input("server-name")
DB_DATABASE = input('db_name')
#DB_USERNAME = 'your_db_username'
#DB_PASSWORD = 'your_db_password'

# Define your data/report
def get_daily_report():
    # Replace this with your actual data fetching/generation logic
    data = {
        'Date': [datetime.today().strftime('%Y-%m-%d')],
        'Metric1': [100],
        'Metric2': [200]
    }
    df = pd.DataFrame(data)
    return df

def create_email(subject, body, to_email):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = to_email
    msg['Subject'] = subject
    
    msg.attach(MIMEText(body, 'html'))  # Changed 'plain' to 'html'
    return msg

def send_email(msg):
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

def store_report_in_db(df):
    conn_str = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={DB_SERVER};'
        f'DATABASE={DB_DATABASE};'
        f'Trusted_Connection=yes;'
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    
    for index, row in df.iterrows():
        try:
            cursor.execute(
                "INSERT INTO email_db (Date, Metric1, Metric2) VALUES (?, ?, ?)",
                row['Date'], row['Metric1'], row['Metric2']
            )
        except Exception as e:
            print(f"Failed to insert row {index}: {e}")
    
    conn.commit()
    cursor.close()
    conn.close()


def main():
    # Fetch the daily report
    report_df = get_daily_report()
    
    # Convert DataFrame to HTML for email body
    report_html = report_df.to_html(index=False)
    email_body = f"<h2>Daily Report</h2>{report_html}"
    
    # Create the email
    email_subject = f"Daily Report - {datetime.today().strftime('%Y-%m-%d')}"
    email_msg = create_email(email_subject, email_body, RECIPIENT_EMAIL)
    
    # Send the email
    send_email(email_msg)
    
    # Store the report in the database
    store_report_in_db(report_df)

if __name__ == "__main__":
    main()
