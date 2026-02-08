import os
import subprocess
from datetime import datetime
import zipfile
import traceback
import time
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

import pandas as pd
from werkzeug.utils import secure_filename
from jinja2 import Environment, FileSystemLoader
from flask import Flask, request, jsonify, send_file, render_template

# Load environment variables
load_dotenv()

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "payslips")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")  # Put your logo.png in the root folder

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

COMPANY = {
    "name": "RS MAN-TECH",
    "address": "#14, 3rd Cross, Parappana Agrahara",
    "city": "Bengaluru-100"
}

# Email configuration from .env
EMAIL_CONFIG = {
    "smtp_server": os.getenv("SMTP_SERVER", "smtp.gmail.com"),
    "smtp_port": int(os.getenv("SMTP_PORT", "587")),
    "sender_email": os.getenv("SENDER_EMAIL", ""),
    "password": os.getenv("EMAIL_PASSWORD", ""),
}

def get_logo_base64():
    """Convert logo to base64 for embedding in HTML"""
    try:
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Could not load logo: {e}")
        return None

def send_email(to_email, emp_name, pdf_path, month):
    """Send payslip email using Zoho SMTP"""
    try:
        print(f"  Preparing email for {to_email}...")

        if not EMAIL_CONFIG["sender_email"] or not EMAIL_CONFIG["password"]:
            print("  ERROR: Email credentials not configured in .env file")
            return False

        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG["sender_email"]
        msg['To'] = to_email
        msg['Subject'] = f"Payslip for {month} - {COMPANY['name']}"

        # Email body
        body = f"""Dear {emp_name},

        Please find attached your payslip for the month of {month}.

        If you have any questions regarding your payslip, please contact HR.

        Best regards,
        {COMPANY['name']}
        HR Department

        ---
        This is an automated email. Please do not reply to this message.
        """

        msg.attach(MIMEText(body, 'plain'))

        # Attach PDF
        with open(pdf_path, 'rb') as file:
            pdf_attachment = MIMEApplication(file.read(), _subtype='pdf')
            pdf_attachment.add_header('Content-Disposition', 'attachment',
                                    filename=f'Payslip_{month}_{emp_name.replace(" ", "_")}.pdf')
            msg.attach(pdf_attachment)

        # Send email
        server = smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        server.starttls()
        server.login(EMAIL_CONFIG["sender_email"], EMAIL_CONFIG["password"])
        server.send_message(msg)
        server.quit()

        print(f"  ✓ Email sent successfully to {to_email}")
        return True

    except Exception as e:
        print(f"  ✗ Email failed for {to_email}: {str(e)}")
        return False

def get_numeric_value(val, default=0):
    try:
        return float(val) if pd.notna(val) else default
    except:
        return default

def number_to_words(num):
    try:
        num = int(float(num))
    except:
        return "Zero rupees only"

    if num == 0:
        return "Zero rupees only"

    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]

    def convert_below_thousand(n):
        if n == 0:
            return ""
        elif n < 10:
            return ones[n]
        elif n < 20:
            return teens[n - 10]
        elif n < 100:
            return tens[n // 10] + (" " + ones[n % 10] if n % 10 != 0 else "")
        else:
            return ones[n // 100] + " Hundred" + (" " + convert_below_thousand(n % 100) if n % 100 != 0 else "")

    if num < 1000:
        result = convert_below_thousand(num)
    elif num < 100000:
        thousands = num // 1000
        remainder = num % 1000
        result = convert_below_thousand(thousands) + " Thousand"
        if remainder > 0:
            result += " " + convert_below_thousand(remainder)
    elif num < 10000000:
        lakhs = num // 100000
        remainder = num % 100000
        result = convert_below_thousand(lakhs) + " Lakh"
        if remainder >= 1000:
            result += " " + convert_below_thousand(remainder // 1000) + " Thousand"
            if remainder % 1000 > 0:
                result += " " + convert_below_thousand(remainder % 1000)
        elif remainder > 0:
            result += " " + convert_below_thousand(remainder)
    else:
        crores = num // 10000000
        remainder = num % 10000000
        result = convert_below_thousand(crores) + " Crore"
        if remainder >= 100000:
            result += " " + convert_below_thousand(remainder // 100000) + " Lakh"
            remainder = remainder % 100000
        if remainder >= 1000:
            result += " " + convert_below_thousand(remainder // 1000) + " Thousand"
            if remainder % 1000 > 0:
                result += " " + convert_below_thousand(remainder % 1000)
        elif remainder > 0:
            result += " " + convert_below_thousand(remainder)

    return result.strip() + " rupees only"

@app.route("/")
def dashboard():
    return render_template("dashboard.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        print("\n" + "="*80)
        print("STARTING PAYSLIP GENERATION")
        print("="*80)

        if "csv_file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["csv_file"]
        month = request.form.get("month", "NA")

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_DIR, filename)
        file.save(file_path)

        print(f"File saved: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".csv":
            try:
                df = pd.read_csv(file_path, encoding="utf-8", engine="python")
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding="latin1", engine="python")
        elif ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path)
        else:
            return jsonify({"error": "Unsupported file type"}), 400

        df.columns = df.columns.str.strip().str.replace('\ufeff', '')

        print(f"\nTotal records in file: {len(df)}")

        env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=True)
        template = env.get_template("payslip.html")

        # Get logo as base64
        logo_base64 = get_logo_base64()

        preview = []
        success_count = 0
        error_count = 0

        for index, row in df.iterrows():
            try:
                print(f"\n--- Processing Employee {index + 1}/{len(df)} ---")

                emp_id = str(row.get("EMP_ID", f"EMP{index+1}")).strip()
                pay_month = str(row.get("Month", month)).strip()

                print(f"EMP_ID: {emp_id}")
                print(f"Name: {row.get('Name', 'N/A')}")

                salary_fixed = {
                    "basic": get_numeric_value(row.get("Fixed_Basic")),
                    "da": get_numeric_value(row.get("Fixed_DA")),
                    "hra": get_numeric_value(row.get("Fixed_HRA")),
                    "leave_wages": 0,
                    "others": 0,
                    "bonus": get_numeric_value(row.get("Fixed_Bonus")),
                    "total": get_numeric_value(row.get("Fixed_Total")),
                }

                salary_earned = {
                    "basic": get_numeric_value(row.get("Earned_Basic")),
                    "da": get_numeric_value(row.get("Earned_DA")),
                    "hra": get_numeric_value(row.get("Earned_HRA")),
                    "leave_wages": 0,
                    "others": get_numeric_value(row.get("Other_Allowance")),
                    "bonus": get_numeric_value(row.get("Earned_Bonus")),
                    "total": get_numeric_value(row.get("Earned_Total")),
                }

                deduction = {
                    "pf": get_numeric_value(row.get("PF")),
                    "esi": get_numeric_value(row.get("ESI")),
                    "pt": get_numeric_value(row.get("PT")),
                    "lwf": get_numeric_value(row.get("LWF")),
                    "adv": 0,
                    "total": get_numeric_value(row.get("Total_Deduction")),
                }

                net_pay = get_numeric_value(row.get("Net_Pay"))
                net_pay_words = number_to_words(net_pay)

                emp_data = {
                    "emp_id": emp_id,
                    "name": str(row.get("Name", "")).strip(),
                    "designation": str(row.get("Designation", "")).strip(),
                    "unit_name": str(row.get("Unit_Name", "")).strip(),
                    "uan": str(row.get("UAN_No", "")).strip(),
                    "esi": str(row.get("ESI_No", "")).strip(),
                    "doj": str(row.get("DOJ", "")).strip(),
                    "bank_ac": str(row.get("Bank_AC", "")).strip(),
                    "ifsc": str(row.get("IFSC_Code", "")).strip(),
                    "email": str(row.get("Email", "")).strip(),
                    "basic_days": str(row.get("Basic_Days", "31")).strip(),
                    "actual_days": str(row.get("Actual_Days", "31")).strip(),
                }

                print("Rendering HTML template...")

                html_content = template.render(
                    company=COMPANY,
                    emp=emp_data,
                    salary_fixed=salary_fixed,
                    salary_earned=salary_earned,
                    deduction=deduction,
                    net_pay=net_pay,
                    net_pay_words=net_pay_words,
                    month=pay_month,
                    generated_on=datetime.now().strftime("%d %b %Y"),
                    logo_base64=logo_base64
                )

                html_path = os.path.join(OUTPUT_DIR, f"{emp_id}.html")
                pdf_path = os.path.join(OUTPUT_DIR, f"{emp_id}.pdf")

                print(f"Writing HTML to: {html_path}")

                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(html_content)

                if not os.path.exists(WKHTMLTOPDF_CMD):
                    error_msg = f"wkhtmltopdf not found at {WKHTMLTOPDF_CMD}"
                    print(f"ERROR: {error_msg}")
                    return jsonify({"error": error_msg}), 500

                print(f"Converting to PDF: {pdf_path}")

                result = subprocess.run(
                    [
                        WKHTMLTOPDF_CMD,
                        "--enable-local-file-access",
                        "--page-size", "A4",
                        "--margin-top", "10mm",
                        "--margin-bottom", "10mm",
                        "--margin-left", "10mm",
                        "--margin-right", "10mm",
                        html_path,
                        pdf_path
                    ],
                    capture_output=True,
                    text=True,
                    timeout=30
                )

                if result.returncode != 0:
                    print(f"PDF generation FAILED for {emp_id}")
                    print(f"Error output: {result.stderr}")
                    error_count += 1
                    continue

                if not os.path.exists(pdf_path):
                    print(f"PDF file not found after generation: {pdf_path}")
                    error_count += 1
                    continue

                file_size = os.path.getsize(pdf_path)
                print(f"SUCCESS: PDF generated ({file_size} bytes)")

                preview.append({
                    "EMP_ID": emp_id,
                    "Name": emp_data["name"],
                    "Designation": emp_data["designation"],
                    "Email": emp_data["email"],
                    "Net_Pay": net_pay,
                    "PDF_Path": pdf_path
                })

                success_count += 1

            except subprocess.TimeoutExpired:
                print(f"TIMEOUT: PDF generation took too long for employee {index + 1}")
                error_count += 1
                continue
            except Exception as emp_error:
                print(f"ERROR processing employee {index + 1}: {str(emp_error)}")
                print(traceback.format_exc())
                error_count += 1
                continue

        print("\n" + "="*80)
        print(f"GENERATION COMPLETE")
        print(f"Success: {success_count}/{len(df)}")
        print(f"Errors: {error_count}/{len(df)}")
        print("="*80 + "\n")

        if success_count == 0:
            return jsonify({"error": "No payslips were generated successfully"}), 500

        return jsonify({
            "message": f"Generated {success_count} payslip(s) successfully",
            "preview": preview
        })

    except Exception as e:
        error_trace = traceback.format_exc()
        print("\n" + "="*80)
        print("FATAL ERROR OCCURRED:")
        print(error_trace)
        print("="*80 + "\n")
        return jsonify({"error": str(e), "details": error_trace}), 500

@app.route("/send-emails", methods=["POST"])
def send_emails():
    """Send payslips via email to all employees"""
    try:
        print("\n" + "="*80)
        print("SENDING EMAILS")
        print("="*80)

        data = request.get_json()
        employees = data.get("employees", [])
        month = data.get("month", "")

        if not employees:
            return jsonify({"error": "No employee data provided"}), 400

        sent_count = 0
        failed_count = 0
        results = []

        for emp in employees:
            emp_email = emp.get("Email")
            emp_name = emp.get("Name")
            emp_id = emp.get("EMP_ID")
            pdf_path = emp.get("PDF_Path")

            if not emp_email or not pdf_path:
                print(f"Skipping {emp_id}: Missing email or PDF path")
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Reason": "Missing email or PDF"})
                continue

            if not os.path.exists(pdf_path):
                pdf_path = os.path.join(OUTPUT_DIR, f"{emp_id}.pdf")
                if not os.path.exists(pdf_path):
                    print(f"Skipping {emp_id}: PDF not found")
                    failed_count += 1
                    results.append({"EMP_ID": emp_id, "Status": "Failed", "Reason": "PDF not found"})
                    continue

            success = send_email(emp_email, emp_name, pdf_path, month)

            if success:
                sent_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Sent", "Email": emp_email})
            else:
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Email": emp_email})

        print("\n" + "="*80)
        print(f"EMAIL SENDING COMPLETE")
        print(f"Sent: {sent_count}/{len(employees)}")
        print(f"Failed: {failed_count}/{len(employees)}")
        print("="*80 + "\n")

        return jsonify({
            "message": f"Sent {sent_count} email(s), {failed_count} failed",
            "sent_count": sent_count,
            "failed_count": failed_count,
            "results": results
        })

    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"\nEMAIL ERROR:\n{error_trace}\n")
        return jsonify({"error": str(e)}), 500

@app.route("/download", methods=["GET"])
def download_pdfs():
    try:
        print("\n" + "="*80)
        print("CREATING ZIP FILE")
        print("="*80)

        zip_path = os.path.join(BASE_DIR, "payslips.zip")

        if os.path.exists(zip_path):
            os.remove(zip_path)
            print("Removed old ZIP file")

        pdf_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pdf")]

        print(f"Found {len(pdf_files)} PDF files to zip")

        if not pdf_files:
            return jsonify({"error": "No PDF files found to download"}), 404

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for pdf_file in pdf_files:
                pdf_full_path = os.path.join(OUTPUT_DIR, pdf_file)
                file_size = os.path.getsize(pdf_full_path)
                print(f"  Adding: {pdf_file} ({file_size} bytes)")
                zipf.write(pdf_full_path, arcname=pdf_file)

        zip_size = os.path.getsize(zip_path)
        print(f"\nZIP created successfully: {zip_path}")
        print(f"ZIP size: {zip_size} bytes")
        print("="*80 + "\n")

        time.sleep(0.1)

        return send_file(zip_path, mimetype='application/zip', as_attachment=True, download_name='payslips.zip')

    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"\nZIP ERROR:\n{error_msg}\n")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("\n" + "="*80)
    print("PAYSLIP GENERATOR SERVER STARTING")
    print(f"Output directory: {OUTPUT_DIR}")
    print(f"Template directory: {TEMPLATE_DIR}")
    print(f"wkhtmltopdf path: {WKHTMLTOPDF_CMD}")
    print(f"Logo path: {LOGO_PATH}")
    print(f"Email configured: {bool(EMAIL_CONFIG['sender_email'] and EMAIL_CONFIG['password'])}")
    print("="*80 + "\n")
    app.run(debug=True)