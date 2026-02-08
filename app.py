import os
import subprocess
from datetime import datetime
import zipfile
import traceback
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv
import io

import pandas as pd
from werkzeug.utils import secure_filename
from jinja2 import Environment, FileSystemLoader
from flask import Flask, request, jsonify, send_file, render_template
from s3_utils import upload_to_s3, list_s3_pdfs, download_s3_file_to_memory

load_dotenv()
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "payslips")
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

COMPANY = {
    "name": "RS MAN-TECH",
    "address": "#14, 3rd Cross, Parappana Agrahara",
    "city": "Bengaluru-100"
}

EMAIL_CONFIG = {
    "smtp_server": os.getenv("SMTP_SERVER", "smtp.gmail.com"),
    "smtp_port": int(os.getenv("SMTP_PORT", "587")),
    "sender_email": os.getenv("SENDER_EMAIL", ""),
    "password": os.getenv("EMAIL_PASSWORD", ""),
}

current_session_pdfs = []

def get_logo_base64():
    try:
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Could not load logo: {e}")
        return None

def send_email(to_email, emp_name, pdf_path, month):
    try:
        print(f"  Preparing email for {to_email}...")
        if not EMAIL_CONFIG["sender_email"] or not EMAIL_CONFIG["password"]:
            print("  ERROR: Email credentials not configured")
            return False

        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG["sender_email"]
        msg['To'] = to_email
        msg['Subject'] = f"Payslip for {month} - {COMPANY['name']}"
        body = f"""Dear {emp_name},

Please find attached your payslip for the month of {month}.

Best regards,
{COMPANY['name']}
HR Department"""
        msg.attach(MIMEText(body, 'plain'))

        with open(pdf_path, 'rb') as file:
            pdf_attachment = MIMEApplication(file.read(), _subtype='pdf')
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename=f'Payslip_{month}_{emp_name.replace(" ", "_")}.pdf')
            msg.attach(pdf_attachment)

        server = smtplib.SMTP(EMAIL_CONFIG["smtp_server"], EMAIL_CONFIG["smtp_port"])
        server.starttls()
        server.login(EMAIL_CONFIG["sender_email"], EMAIL_CONFIG["password"])
        server.send_message(msg)
        server.quit()
        print(f"  ✓ Email sent to {to_email}")
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
        result = convert_below_thousand(num // 1000) + " Thousand"
        if num % 1000 > 0:
            result += " " + convert_below_thousand(num % 1000)
    elif num < 10000000:
        result = convert_below_thousand(num // 100000) + " Lakh"
        remainder = num % 100000
        if remainder >= 1000:
            result += " " + convert_below_thousand(remainder // 1000) + " Thousand"
            if remainder % 1000 > 0:
                result += " " + convert_below_thousand(remainder % 1000)
        elif remainder > 0:
            result += " " + convert_below_thousand(remainder)
    else:
        result = convert_below_thousand(num // 10000000) + " Crore"
        remainder = num % 10000000
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
    global current_session_pdfs
    current_session_pdfs = []
    
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
        env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=True)
        template = env.get_template("payslip.html")
        logo_base64 = get_logo_base64()

        preview = []
        success_count = 0
        error_count = 0

        for index, row in df.iterrows():
            try:
                emp_id = str(row.get("EMP_ID", f"EMP{index+1}")).strip()
                pay_month = month
                print(f"DEBUG: Storing to S3 with month folder: {pay_month}")


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

                html_content = template.render(
                    company=COMPANY, emp=emp_data, salary_fixed=salary_fixed,
                    salary_earned=salary_earned, deduction=deduction, net_pay=net_pay,
                    net_pay_words=net_pay_words, month=pay_month,
                    generated_on=datetime.now().strftime("%d %b %Y"), logo_base64=logo_base64
                )

                html_path = os.path.join(OUTPUT_DIR, f"{emp_id}.html")
                pdf_path = os.path.join(OUTPUT_DIR, f"{emp_id}.pdf")

                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(html_content)

                if not os.path.exists(WKHTMLTOPDF_CMD):
                    return jsonify({"error": f"wkhtmltopdf not found at {WKHTMLTOPDF_CMD}"}), 500

                result = subprocess.run([WKHTMLTOPDF_CMD, "--enable-local-file-access", "--page-size", "A4",
                    "--margin-top", "10mm", "--margin-bottom", "10mm", "--margin-left", "10mm",
                    "--margin-right", "10mm", html_path, pdf_path], capture_output=True, text=True, timeout=30)

                if result.returncode != 0 or not os.path.exists(pdf_path):
                    error_count += 1
                    continue

                try:
                    print(f"DEBUG: Storing to S3 with month folder: {pay_month}")
                    s3_key = upload_to_s3(pdf_path, month=pay_month)
                    print(f"DEBUG: S3 key created: {s3_key}")
                    current_session_pdfs.append(s3_key)
                except Exception as s3_error:
                    print(f"S3 upload failed: {s3_error}")

                preview.append({"EMP_ID": emp_id, "Name": emp_data["name"], "Designation": emp_data["designation"],
                    "Email": emp_data["email"], "Net_Pay": net_pay, "PDF_Path": pdf_path})
                success_count += 1

            except subprocess.TimeoutExpired:
                error_count += 1
                continue
            except Exception as emp_error:
                print(f"ERROR: {str(emp_error)}")
                error_count += 1
                continue

        print(f"\nGENERATION COMPLETE - Success: {success_count}/{len(df)}, Errors: {error_count}/{len(df)}\n")

        if success_count == 0:
            return jsonify({"error": "No payslips generated"}), 500

        return jsonify({"message": f"Generated {success_count} payslip(s)", "preview": preview})

    except Exception as e:
        print(f"\nFATAL ERROR: {traceback.format_exc()}\n")
        return jsonify({"error": str(e)}), 500

@app.route("/send-emails", methods=["POST"])
def send_emails():
    try:
        data = request.get_json()
        employees = data.get("employees", [])
        month = data.get("month", "")

        if not employees:
            return jsonify({"error": "No employee data"}), 400

        sent_count = 0
        failed_count = 0
        results = []

        for emp in employees:
            emp_email = emp.get("Email")
            emp_name = emp.get("Name")
            emp_id = emp.get("EMP_ID")
            pdf_path = emp.get("PDF_Path")

            if not emp_email or not pdf_path:
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Reason": "Missing data"})
                continue

            if not os.path.exists(pdf_path):
                pdf_path = os.path.join(OUTPUT_DIR, f"{emp_id}.pdf")
                if not os.path.exists(pdf_path):
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

        return jsonify({"message": f"Sent {sent_count}, failed {failed_count}", "sent_count": sent_count,
            "failed_count": failed_count, "results": results})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download-current", methods=["GET"])
def download_current_session():
    try:
        if not current_session_pdfs:
            return jsonify({"error": "No PDFs in current session"}), 404

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for s3_key in current_session_pdfs:
                pdf_data = download_s3_file_to_memory(s3_key)
                zipf.writestr(os.path.basename(s3_key), pdf_data.read())

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name='current_payslips.zip')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/download", methods=["GET"])
def download_pdfs():
    try:
        month = request.args.get("month")
        print(f"DEBUG: Searching S3 for month: {month}")
        s3_pdf_keys = list_s3_pdfs(month=month)
        print(f"DEBUG: Found keys: {s3_pdf_keys}")
        
        if not s3_pdf_keys:
            return jsonify({"error": "No PDF files found"}), 404

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for s3_key in s3_pdf_keys:
                pdf_data = download_s3_file_to_memory(s3_key)
                zipf.writestr(os.path.basename(s3_key), pdf_data.read())

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True,
            download_name=f'payslips_{month}.zip' if month else 'payslips.zip')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("\n" + "="*80)
    print("PAYSLIP GENERATOR STARTING")
    print("="*80 + "\n")
    app.run(debug=True)
