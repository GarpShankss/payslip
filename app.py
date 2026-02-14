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

# Detect wkhtmltopdf path (Docker vs Windows)
if os.path.exists('/usr/local/bin/wkhtmltopdf'):
    WKHTMLTOPDF_CMD = '/usr/local/bin/wkhtmltopdf'
elif os.path.exists('/usr/bin/wkhtmltopdf'):
    WKHTMLTOPDF_CMD = '/usr/bin/wkhtmltopdf'
else:
    WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
print(f"Using wkhtmltopdf at: {WKHTMLTOPDF_CMD}")

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
            # Find the correct header row by looking for EMP_ID or NAME column
            header_row = None
            for row_num in range(10):  # Check first 10 rows
                try:
                    test_df = pd.read_excel(file_path, header=row_num, nrows=1)
                    cols_lower = [str(c).lower() for c in test_df.columns]
                    if any('emp' in c or 'name' in c or 'id' in c for c in cols_lower if not c.startswith('unnamed')):
                        header_row = row_num
                        print(f"Found header row at: {row_num}")
                        break
                except:
                    continue
            
            if header_row is not None:
                df = pd.read_excel(file_path, header=header_row)
                
                # Handle multi-row headers by merging with next row if needed
                # Check if next row has additional column names
                try:
                    next_row_df = pd.read_excel(file_path, header=header_row+1, nrows=1)
                    next_cols = [str(c).lower() for c in next_row_df.columns]
                    # If next row has salary columns, merge the headers
                    if any('fixed' in c or 'earned' in c or 'deduction' in c for c in next_cols if not c.startswith('unnamed')):
                        print(f"Detected multi-row headers, merging row {header_row} and {header_row+1}")
                        # Read both rows as headers
                        df = pd.read_excel(file_path, header=[header_row, header_row+1])
                        # Flatten multi-level columns and remove duplicate prefixes
                        new_cols = []
                        for col in df.columns:
                            parts = [str(c).strip() for c in col if not str(c).startswith('Unnamed')]
                            # Remove duplicate words (e.g., FIXED_FIXED_BASIC -> FIXED_BASIC)
                            if len(parts) == 2 and parts[0].upper() == parts[1].split('_')[0].upper():
                                new_cols.append(parts[1])
                            else:
                                new_cols.append('_'.join(parts).strip('_'))
                        df.columns = new_cols
                except:
                    pass
            else:
                df = pd.read_excel(file_path)
        else:
            return jsonify({"error": "Unsupported file type"}), 400

        df.columns = df.columns.str.strip().str.replace('\ufeff', '')
        
        print(f"\n=== EXCEL COLUMNS DEBUG ===")
        print(f"Total columns in Excel: {len(df.columns)}")
        print(f"All columns: {list(df.columns)}")
        print(f"=== END DEBUG ===\n")
        
        # Remove empty rows
        df = df.dropna(how='all')
        
        # Validate ALL required columns BEFORE processing
        required_columns = [
            'Name', 'EMP_ID', 'Fixed_Basic', 'Fixed_DA', 'Fixed_HRA', 'Fixed_Total',
            'Earned_Basic', 'Earned_DA', 'Earned_HRA', 'Earned_Total',
            'PF', 'ESI', 'PT', 'Total_Deduction', 'Net_Pay'
        ]
        
        # Create case-insensitive column mapping for validation
        excel_cols_lower = {col.lower(): col for col in df.columns}
        missing_required = []
        
        for col in required_columns:
            col_lower = col.lower()
            # Check if column exists (exact match or with prefix)
            found = False
            if col_lower in excel_cols_lower:
                found = True
            elif col_lower == 'pf' and 'deductions_pf' in excel_cols_lower:
                found = True
            elif col_lower == 'esi' and 'deductions_esi' in excel_cols_lower:
                found = True
            elif col_lower == 'pt' and 'deductions_pt' in excel_cols_lower:
                found = True
            elif col_lower == 'total_deduction' and 'deductions_total' in excel_cols_lower:
                found = True
            elif col_lower.startswith('fixed_'):
                base_name = col_lower.replace('fixed_', '')
                if base_name in excel_cols_lower or f'fixed_{base_name}' in excel_cols_lower:
                    found = True
            elif col_lower.startswith('earned_'):
                base_name = col_lower.replace('earned_', '')
                if base_name in excel_cols_lower or f'earned_{base_name}' in excel_cols_lower:
                    found = True
            
            if not found:
                missing_required.append(col)
        
        if missing_required:
            error_msg = f"❌ Cannot generate payslips.\n\nRequired columns missing: {', '.join(missing_required)}\n\n"
            error_msg += f"Your Excel has: {', '.join(list(df.columns)[:20])}\n\n"
            error_msg += "Please add ALL required columns and try again."
            return jsonify({"error": error_msg}), 400
        
        # Create case-insensitive column mapping
        col_map = {col: col for col in df.columns}
        for col in df.columns:
            col_map[col.lower()] = col
            # Add aliases for common variations
            col_lower = col.lower().replace(' ', '_')
            col_map[col_lower] = col
            
            # Handle DEDUCTIONS_PF -> PF, DEDUCTIONS_ESI -> ESI, etc.
            if 'deductions_' in col_lower:
                col_map[col_lower.replace('deductions_', '')] = col
            # Handle NET_PAY_EMAIL -> EMAIL, NET_PAY_phone_no -> phone
            if 'net_pay_' in col_lower:
                col_map[col_lower.replace('net_pay_', '')] = col
            
            # Specific mappings
            if col_lower == 'total' or col_lower == 'deductions_total':
                col_map['total_deduction'] = col
            if col_lower == 'adv' or col_lower == 'deductions_adv':
                col_map['lwf'] = col
            if 'esi_no' in col_lower or 'esi no' in col_lower:
                col_map['esi_no'] = col
            if 'phone_no' in col_lower or 'phone no' in col_lower:
                col_map['phone'] = col
            if col_lower == 'email' or 'net_pay_email' in col_lower:
                col_map['email'] = col
            # Map EARNED_HRA.1 or EARNED_HRA.2 to Other_Allowance
            if 'earned_hra.1' in col_lower or 'earned_hra.2' in col_lower:
                col_map['other_allowance'] = col
            # Map EARNED_Other_Allowance to Other_Allowance
            if col_lower == 'earned_other_allowance':
                col_map['other_allowance'] = col
        
        def get_col(name):
            """Get column value with case-insensitive lookup"""
            result = col_map.get(name.lower(), name)
            if result == name and name.lower() not in col_map:
                missing_columns.add(name)
            return result
        
        missing_columns = set()
        
        print(f"Loaded {len(df)} employees from file")
        print(f"Columns found: {list(df.columns)[:15]}...")  # Show first 15 columns
        print(f"Total columns: {len(df.columns)}")
        
        env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=True)
        template = env.get_template("payslip.html")
        logo_base64 = get_logo_base64()

        preview = []
        success_count = 0
        error_count = 0

        for index, row in df.iterrows():
            try:
                # Skip rows with no EMP_ID or Name
                emp_id_val = row.get(get_col("EMP_ID"), f"EMP{index+1}")
                name_val = row.get(get_col("Name"), "")
                
                print(f"Row {index+1}: EMP_ID={emp_id_val}, Name={name_val}")
                
                if pd.isna(emp_id_val) or pd.isna(name_val) or str(name_val).strip() == "":
                    print(f"Skipping empty row {index+1} - EMP_ID: {emp_id_val}, Name: {name_val}")
                    continue
                
                emp_id = str(emp_id_val).strip()
                pay_month = month
                print(f"Processing employee {emp_id}...")


                salary_fixed = {
                    "basic": get_numeric_value(row.get(get_col("Fixed_Basic"))),
                    "da": get_numeric_value(row.get(get_col("Fixed_DA"))),
                    "hra": get_numeric_value(row.get(get_col("Fixed_HRA"))),
                    "leave_wages": 0,
                    "others": 0,
                    "bonus": get_numeric_value(row.get(get_col("Fixed_Bonus"))),
                    "total": get_numeric_value(row.get(get_col("Fixed_Total"))),
                }

                salary_earned = {
                    "basic": get_numeric_value(row.get(get_col("Earned_Basic"))),
                    "da": get_numeric_value(row.get(get_col("Earned_DA"))),
                    "hra": get_numeric_value(row.get(get_col("Earned_HRA"))),
                    "leave_wages": get_numeric_value(row.get(get_col("Earned_Leave_Wages"))),
                    "others": get_numeric_value(row.get(get_col("Other_Allowance"))),
                    "bonus": get_numeric_value(row.get(get_col("Earned_Bonus"))),
                    "total": get_numeric_value(row.get(get_col("Earned_Total"))),
                }

                deduction = {
                    "pf": get_numeric_value(row.get(get_col("PF"))),
                    "esi": get_numeric_value(row.get(get_col("ESI"))),
                    "pt": get_numeric_value(row.get(get_col("PT"))),
                    "lwf": get_numeric_value(row.get(get_col("LWF"))),
                    "adv": 0,
                    "total": get_numeric_value(row.get(get_col("Total_Deduction"))),
                }

                net_pay = get_numeric_value(row.get(get_col("Net_Pay")))
                net_pay_words = number_to_words(net_pay)

                emp_data = {
                    "emp_id": emp_id,
                    "name": str(row.get(get_col("Name"), "")).strip(),
                    "designation": str(row.get(get_col("Designation"), "")).strip(),
                    "unit_name": str(row.get(get_col("Unit_Name"), "")).strip(),
                    "uan": str(int(float(row.get(get_col("UAN_No"), 0)))) if pd.notna(row.get(get_col("UAN_No"))) else "",
                    "esi": str(row.get(get_col("ESI_No"), "")).strip(),
                    "doj": str(row.get(get_col("DOJ"), "")).strip(),
                    "bank_ac": str(int(float(row.get(get_col("Bank_AC"), 0)))) if pd.notna(row.get(get_col("Bank_AC"))) else "",
                    "ifsc": str(row.get(get_col("IFSC_Code"), "")).strip(),
                    "email": str(row.get(get_col("Email"), "")).strip(),
                    "phone": str(row.get(get_col("Phone"), "")).strip(),
                    "basic_days": str(row.get(get_col("Basic_Days"), "31")).strip(),
                    "actual_days": str(row.get(get_col("Actual_Days"), "31")).strip(),
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

                result = subprocess.run([WKHTMLTOPDF_CMD, "--enable-local-file-access", "--page-size", "A4",
                    "--margin-top", "10mm", "--margin-bottom", "10mm", "--margin-left", "10mm",
                    "--margin-right", "10mm", html_path, pdf_path], capture_output=True, text=True, timeout=30)

                if result.returncode != 0:
                    print(f"ERROR: wkhtmltopdf failed for {emp_id}")
                    print(f"STDOUT: {result.stdout}")
                    print(f"STDERR: {result.stderr}")
                    error_count += 1
                    continue
                    
                if not os.path.exists(pdf_path):
                    print(f"ERROR: PDF not created for {emp_id}")
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
                print(f"ERROR: Timeout for employee {emp_id}")
                error_count += 1
                continue
            except Exception as emp_error:
                print(f"ERROR processing {emp_id}: {str(emp_error)}")
                print(f"Traceback: {traceback.format_exc()}")
                error_count += 1
                continue

        print(f"\nGENERATION COMPLETE - Success: {success_count}/{len(df)}, Errors: {error_count}/{len(df)}\n")

        if success_count == 0:
            error_msg = "❌ No payslips generated.\n\n"
            if missing_columns:
                missing_list = sorted(list(missing_columns))
                error_msg += f"Missing columns in your Excel: {', '.join(missing_list)}\n\n"
                error_msg += "Please add these columns and try again."
            else:
                error_msg += "All rows were skipped. Check if your Excel has data."
            return jsonify({"error": error_msg}), 500
        
        # Show missing columns warning to user
        warning_msg = ""
        if missing_columns:
            missing_list = sorted(list(missing_columns))
            warning_msg = f"Warning: The following columns were not found in your Excel file: {', '.join(missing_list)}. These fields will be empty in the payslips."
            print(f"\n{warning_msg}\n")

        return jsonify({
            "message": f"Generated {success_count} payslip(s)", 
            "preview": preview,
            "warning": warning_msg if missing_columns else None
        })

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
