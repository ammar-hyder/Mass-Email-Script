import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Email credentials
SMTP_SERVER = "smtp.gmail.com"  # Change for Outlook/Yahoo
SMTP_PORT = 587
EMAIL_ADDRESS = "procom.net@nu.edu.pk"  # Replace with your email
EMAIL_PASSWORD = "tcza rhui bgxb fzjx"  # Use an App Password if required

# Read Excel file
file_path = "emails.xlsx"  # Change to your actual file path
df = pd.read_excel(file_path)

# Ensure correct column name
df.columns = df.columns.str.strip()  # Remove hidden spaces
print("Column Names:", df.columns.tolist())  # Debugging step

# Get the correct column name
email_column = df.columns[0]  # If only one column exists, use it

# Function to send email
def send_email(to_email):
    # Email Subject
    subject = "Register Now for Sheikhani Group Presents PROCOM’25 at FAST University – Pakistan’s Largest Tech Fest!"

    # Email Body (EXACTLY as provided)
    body = """\
    As we celebrate the 26th edition of PROCOM, carrying forward its legacy since 1998, this year’s event is set to be bigger, bolder, and more innovative than ever before! Whether you're a tech enthusiast, aspiring entrepreneur, or problem solver, PROCOM’25 at FAST National University of Computer and Emerging Sciences, Karachi is the ultimate platform to showcase your skills and ideas.

    Why PROCOM’25?
    🚀 Career Fair – Connect with top companies and kickstart your career.
    🏆 25+ Technical Competitions – Test your expertise in AI, Computing, Electrical Engineering, Business, and more!
    🎤 Panel Discussions – Gain insights from renowned industry leaders & tech experts.
    💡 Startup Showcase – Witness groundbreaking entrepreneurial innovations.

    With teams from across Pakistan, a 1.2 Million Rupee prize pool, and the prestigious Sheikhani Group PROCOM Trophy at stake, this is your chance to compete, innovate, and make an impact!

    📝 Registration Deadline: Monday, 17th February
    📍 Register Now: https://www.procom.com.pk/register

    Join us for two thrilling days of competition, networking, and learning. Your participation will contribute to the success of this grand event and provide invaluable opportunities for growth.

    Stay Updated:
    🌐 https://www.procom.com.pk/
    🔵 https://www.facebook.com/share/1DutW6eeje/?mibextid=wwXIfr
    📸 https://www.instagram.com/procom_fast?igsh=c3Y4dXhxNWdqNnZr
    💼 https://www.linkedin.com/company/procom-fast/
    
    Contact us on whatsapp:
    03702743866
    
    We look forward to welcoming you to Sheikhani Group presents PROCOM’25 at FAST University!

    Best Regards,  
    Team Sheikhani Group PROCOM’25
    """
    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure connection
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# Loop through each email and send
for _, row in df.iterrows():
    send_email(row[email_column])  # Use the detected email column

