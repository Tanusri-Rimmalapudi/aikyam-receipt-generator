import os
from pathlib import Path
import pandas as pd
import fitz  # PyMuPDF
import smtplib
from email.message import EmailMessage


# ==========================
# CONFIG – EDIT THESE
# ==========================
EXCEL_FILE = "donor_summary.xlsx"
TEMPLATE_PDF = "AIKYAM LetterHead.pdf"
OUTPUT_DIR = "receipts"

# Use your Gmail here
SENDER_EMAIL = "notifyaikyam@gmail.com"    # <-- CHANGE THIS
SENDER_NAME = "AIKYAM"                        # Name shown in From:
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_PASSWORD = "hagk cbyn xqxd dqaw"     # <-- CHANGE THIS (Gmail App Password)


REQUIRED_COLS = {
    "name",
    "email",
    "phone",
    "amount",
    "invoice number",
    "invoice date",
}


# ==========================
# HELPER FUNCTIONS
# ==========================

def load_data(excel_path: str) -> pd.DataFrame:
    print(f"Checking for Excel file: {excel_path}")
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"{excel_path} not found in {os.getcwd()}")

    print("Reading Excel file with pandas...")
    df = pd.read_excel(excel_path)

    df.columns = df.columns.str.strip().str.lower()
    print(f"Rows found in Excel: {len(df)}")
    print(f"Columns in file: {list(df.columns)}")

    if not REQUIRED_COLS.issubset(df.columns):
        raise ValueError(f"Excel file must contain columns: {REQUIRED_COLS}")

    return df


def sanitize_filename(name: str) -> str:
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip()
    return safe or "invoice"


def format_invoice_date(value) -> str:
    if pd.isna(value):
        return ""
    try:
        dt = pd.to_datetime(value)
        return dt.strftime("%m/%d/%Y")
    except Exception:
        return str(value)


def format_phone_number(phone):
    """Format phone number as +1(###)###-####."""
    digits = "".join(filter(str.isdigit, str(phone)))

    # Remove leading 1 if number is already in US format
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    if len(digits) != 10:
        return str(phone)  # return original if not a 10-digit US number

    area = digits[:3]
    first = digits[3:6]
    last = digits[6:]

    return f"+1({area}){first}-{last}"


def create_invoice_pdf(name, email, phone, amount, invoice_no, invoice_date) -> Path:
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(f"Template PDF '{TEMPLATE_PDF}' not found.")

    doc = fitz.open(TEMPLATE_PDF)
    page = doc[0]

    amount_val = float(amount)
    amount_str = f"${amount_val:,.2f}"
    invoice_date_str = format_invoice_date(invoice_date)
    formatted_phone = format_phone_number(phone)

    # --- TOP RIGHT: Invoice No + Date ---
    invoice_box = fitz.Rect(330, 120, 550, 200)
    invoice_text = f"Invoice No: {invoice_no}\nInvoice Date: {invoice_date_str}"
    page.insert_textbox(invoice_box, invoice_text, fontsize=12, fontname="helv", align=2)

    # --- ADDRESS LEFT ---
    address_box = fitz.Rect(72, 150, 350, 230)
    address_text = (
        "Mana AIKYAM\n"
        "11 Corte Rivera\n"
        "Lake Elsinore CA 92532\n"
        "USA\n"
    )
    page.insert_textbox(address_box, address_text, fontsize=12, fontname="helv", align=0)

    # --- BILL TO ---
    bill_to_box = fitz.Rect(72, 250, 350, 360)
    bill_to_text = f"Bill To:\n{name}\n{email}\n{formatted_phone}"
    page.insert_textbox(bill_to_box, bill_to_text, fontsize=12, fontname="helv", align=0)

    # --- LINE ITEM ---
    # Left label
    line_label_box = fitz.Rect(72, 380, 350, 410)
    page.insert_textbox(line_label_box, "Diwali 2025 Celebration", fontsize=12, fontname="helv", align=0)

    # Right amount
    line_amount_box = fitz.Rect(350, 380, 550, 410)
    page.insert_textbox(line_amount_box, amount_str, fontsize=12, fontname="helv", align=2)

    # --- INVOICE TOTAL (BOLD EFFECT) ---
    total_label_box = fitz.Rect(72, 430, 550, 470)
    total_amount_box = fitz.Rect(350, 430, 550, 470)

    # Bold effect using double print + larger font
    page.insert_textbox(total_label_box, "Invoice Total", fontsize=14, fontname="helv", align=1)
    page.insert_textbox(total_label_box, "Invoice Total", fontsize=14, fontname="helv", align=1)

    page.insert_textbox(total_amount_box, amount_str, fontsize=14, fontname="helv", align=2)
    page.insert_textbox(total_amount_box, amount_str, fontsize=14, fontname="helv", align=2)

    # --- THANK YOU MESSAGE UNDER TOTAL ---
    thanks_box = fitz.Rect(72, 480, 550, 540)
    thanks_text = (
        "Thank you for your generous support to AIKYAM and the Diwali 2025 Celebration."
    )
    page.insert_textbox(thanks_box, thanks_text, fontsize=12, fontname="helv", align=0)

    # Save
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"{sanitize_filename(str(invoice_no))}_{sanitize_filename(name)}.pdf"
    output_path = Path(OUTPUT_DIR) / filename

    doc.save(output_path)
    doc.close()
    return output_path


def build_email_message(sender_email, sender_name, recipient_email, recipient_name, amount, invoice_no, invoice_date, pdf_path):
    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = recipient_email
    msg["Subject"] = "Thank You for Registering – AIKYAM Diwali Event & Your Receipt"

    body = (
        f"Dear {recipient_name},\n\n"
        f"Thank you for registering for the AIKYAM Diwali Event 2025. We are delighted to share that the event was a grand success, and your contribution played a key role in making it memorable for our entire community. Please find your receipt attached for the amount you contributed during registration.\n\n"
        f"If you notice any discrepancy in the amount or have any questions, please reach out to us at:.\n"
        f"notify@aikyamusa.org.\n\n"
        f"Warm regards,\n"
        f"Team {sender_name}\n"
        f"AIKYAM — Together we are stronger\n"
        f"www.AIKYAMUSA.org\n"
    )

    msg.set_content(body)

    with open(pdf_path, "rb") as f:
        pdf_data = f.read()
    msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename=pdf_path.name)

    return msg


def send_invoices(df: pd.DataFrame):
    print(f"Invoices will be saved in folder: {OUTPUT_DIR}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"Connecting to SMTP server {SMTP_SERVER}:{SMTP_PORT} ...")
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        # LOGIN WITH YOUR GMAIL + APP PASSWORD
        smtp.login(SENDER_EMAIL, EMAIL_PASSWORD)
        print("SMTP login successful.")

        for idx, row in df.iterrows():
            name = str(row["name"]).strip()
            email = str(row["email"]).strip()
            phone = str(row["phone"]).strip()
            amount = row["amount"]
            invoice_no = str(row["invoice number"]).strip()
            invoice_date = row["invoice date"]

            print(f"\nProcessing Invoice {invoice_no} for {name}...")

            pdf_path = create_invoice_pdf(name, email, phone, amount, invoice_no, invoice_date)
            msg = build_email_message(
                SENDER_EMAIL, SENDER_NAME, email, name, amount, invoice_no, invoice_date, pdf_path
            )

            smtp.send_message(msg)
            print(f"Email sent to {email}")


def main():
    df = load_data(EXCEL_FILE)
    send_invoices(df)
    print("\nAll done!")


if __name__ == "__main__":
    main()
