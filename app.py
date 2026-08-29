import os
import sys
import subprocess
from datetime import datetime
import zipfile
import traceback
import base64
import smtplib
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv
import io
from whatsapp_utils import send_payslip_whatsapp
from flask import Flask, request, jsonify, send_file, render_template, session, redirect, url_for

# Fix Windows console unicode encoding issues
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import pandas as pd
from werkzeug.utils import secure_filename
from jinja2 import Environment, FileSystemLoader
from s3_utils import upload_with_cleanup, list_s3_pdfs, download_s3_file_to_memory

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "change-this-secret")

# ── PATHS ──
BASE_DIR          = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR        = "/tmp/uploads"
PAYSLIPS_BASE_DIR = "/tmp/payslips"
TEMPLATE_DIR      = os.path.join(BASE_DIR, "templates")
LOGO_PATH         = os.path.join(BASE_DIR, "logo.png")
MAX_SESSIONS      = 2

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PAYSLIPS_BASE_DIR, exist_ok=True)

# ── GLOBALS ──
current_session_pdfs = []
current_output_dir   = PAYSLIPS_BASE_DIR  # safe now

# ── SESSION MANAGEMENT ──
def get_session_dir() -> str:
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    session_dir = os.path.join(PAYSLIPS_BASE_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)
    return session_dir

def cleanup_old_sessions():
    try:
        sessions = sorted([
            os.path.join(PAYSLIPS_BASE_DIR, d)
            for d in os.listdir(PAYSLIPS_BASE_DIR)
            if os.path.isdir(os.path.join(PAYSLIPS_BASE_DIR, d))
        ])
        sessions_to_delete = sessions[:-MAX_SESSIONS] if len(sessions) > MAX_SESSIONS else []
        for old_session in sessions_to_delete:
            shutil.rmtree(old_session, ignore_errors=True)
            print(f"[CLEANUP] Deleted old session: {old_session}")
    except Exception as e:
        print(f"Cleanup error: {e}")

TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")

os.makedirs(UPLOAD_DIR, exist_ok=True)
# os.makedirs(OUTPUT_DIR, exist_ok=True)

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

@app.route("/upload", methods=["POST"])
def upload_file():
    if not session.get("logged_in"):
        return jsonify({"error": "Session expired. Please log in again."}), 401
    global current_session_pdfs, current_output_dir
    current_session_pdfs = []

# Clean old sessions first, then create new one
    cleanup_old_sessions()
    current_output_dir = get_session_dir()
    print(f"Session directory: {current_output_dir}")
    try:
        print("\n" + "="*80)
        print("STARTING PAYSLIP GENERATION")
        print("="*80)

        if "csv_file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["csv_file"]
        month = request.form.get("month", "NA")
        year = request.form.get("year", str(datetime.now().year))
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
            # Find the correct header row by looking for EMP_ID, NAME, or Sl.No column
            header_row = None
            for row_num in range(10):  # Check first 10 rows
                try:
                    test_df = pd.read_excel(file_path, header=row_num, nrows=1)
                    cols_lower = [str(c).lower() for c in test_df.columns]
                    if any('emp' in c or 'name' in c or 'id' in c or 'sl.no' in c or 'sl no' in c for c in cols_lower if not c.startswith('unnamed')):
                        header_row = row_num
                        print(f"Found header row at: {row_num}")
                        break
                except:
                    continue
            
            if header_row is not None:
                try:
                    next_row_df = pd.read_excel(file_path, header=header_row+1, nrows=1)
                    next_cols = [str(c).lower() for c in next_row_df.columns]
                    # If next row has salary columns or subheaders, merge the headers
                    if any('fixed' in c or 'earned' in c or 'deduction' in c or 'basic' in c or 'uan' in c for c in next_cols if not c.startswith('unnamed')):
                        print(f"Detected multi-row headers, merging row {header_row} and {header_row+1}")
                        df = pd.read_excel(file_path, header=[header_row, header_row+1])
                        new_cols = []
                        for col in df.columns:
                            parts = [str(c).strip() for c in col if not str(c).startswith('Unnamed')]
                            if len(parts) == 2 and parts[0].upper() == parts[1].split('_')[0].upper():
                                new_cols.append(parts[1])
                            else:
                                new_cols.append('_'.join(parts).strip('_'))
                        df.columns = new_cols
                    else:
                        df = pd.read_excel(file_path, header=header_row)
                except:
                    df = pd.read_excel(file_path, header=header_row)
            else:
                df = pd.read_excel(file_path)
        else:
            return jsonify({"error": "Unsupported file type"}), 400

        df.columns = df.columns.str.strip().str.replace('\ufeff', '')
        df = df.dropna(how='all')
        
        print(f"\n=== EXCEL COLUMNS DEBUG ===")
        print(f"Total columns in Excel: {len(df.columns)}")
        print(f"All columns: {list(df.columns)}")
        print(f"=== END DEBUG ===\n")

        # Create case-insensitive column lookup dictionary
        col_map = {col: col for col in df.columns}
        for col in df.columns:
            col_map[col.lower()] = col
            col_lower = col.lower().replace(' ', '_')
            col_map[col_lower] = col
            if 'deductions_' in col_lower or 'deduction_' in col_lower:
                cleaned = col_lower.replace('deductions_', '').replace('deduction_', '')
                col_map[cleaned] = col
            if 'net_pay_' in col_lower:
                col_map[col_lower.replace('net_pay_', '')] = col

        def find_col(*candidates):
            for c in candidates:
                if c.lower() in col_map:
                    return col_map[c.lower()]
                c_slug = c.lower().replace(' ', '_')
                if c_slug in col_map:
                    return col_map[c_slug]
            return None

        # Detect columns in the uploaded Excel
        col_emp_id = find_col("EMP_ID", "EMP ID", "EMPLOYEE ID", "EMP CODE", "ID", "Sl.No", "Sl No")
        col_name = find_col("NAME", "EMPLOYEE NAME", "EMP NAME")
        col_designation = find_col("Designation", "DESIG")
        col_unit = find_col("Unit_Name", "UNIT NAME", "Unit", "Unit Name")
        col_uan = find_col("UAN_NO", "UAN NUMBER", "NAME_UAN", "UAN NO", "UAN")
        col_esi_no = find_col("ESI_NO", "ESI NO", "ESI NUMBER")
        col_doj = find_col("DOJ", "DATE OF JOINING")
        col_bank_ac = find_col("BANK_AC", "BANK A/C", "BANK AC", "ACCOUNT", "ACCOUNT NO")
        col_ifsc = find_col("IFSC_CODE", "IFSC CODE", "IFSC")
        col_phone = find_col("Phone", "phone no", "NAME_CONTACT", "MOBILE", "CONTACT")
        col_email = find_col("Email", "EMAIL")
        col_basic_days = find_col("BASIC_DAYS", "Basic Days", "TOTAL DAYS")
        col_actual_days = find_col("ACTUAL_DAYS", "Actual Days", "DAYS WORKED")

        # Fixed salary columns
        col_fix_basic_da = find_col("FIXED_BASIC & DA", "FIXED_BASIC_DA", "BASIC & DA")
        col_fix_basic = find_col("FIXED_BASIC", "FIXED_BASIC SALARY", "BASIC")
        col_fix_da = find_col("FIXED_DA", "DA")
        col_fix_hra = find_col("FIXED_HRA", "HRA")
        col_fix_leave = find_col("FIXED_LEAVE_WAGES", "FIXED_LEAVE WAGES", "Leave_Wages")
        col_fix_other = find_col("FIXED_OTHER ALLOWANCE", "FIXED_OTHER_ALLOWANCE", "FIXED_OTHERS", "OTHER ALLOWANCE", "OTHER_ALLOWANCE", "OTHERS")
        col_fix_special = find_col("FIXED_SPECIAL ALLOW", "FIXED_SPECIAL ALLOWANCE", "FIXED_SPECIAL_ALLOWANCE", "SPECIAL ALLOW", "SPECIAL ALLOWANCE")
        col_fix_bonus = find_col("FIXED_BONUS", "FIXED_STATUORY BONUS", "FIXED_STATUTORY BONUS", "STATUORY BONUS", "STATUTORY BONUS")
        col_fix_total = find_col("FIXED_TOTAL", "FIXED_GROSS", "TOTAL")

        # Earned salary columns
        col_earn_basic_da = find_col("EARNED_BASIC & DA", "EARNED_BASIC_DA")
        col_earn_basic = find_col("EARNED_BASIC")
        col_earn_da = find_col("EARNED_DA")
        col_earn_hra = find_col("EARNED_HRA")
        col_earn_leave = find_col("EARNED_LEAVE_WAGES", "Earned_Leave_Wages", "LEAVE_WAGES")
        col_earn_other = find_col("EARNED_OTHER ALLOWANCE", "EARNED_OTHER_ALLOWANCE", "EARNED_OTHERS")
        if not col_earn_other:
            # Check for generic other allowance
            col_earn_other = col_fix_other
        col_earn_special = find_col("EARNED_SPECIAL ALLOW", "EARNED_SPECIAL ALLOWANCE", "EARNED_SPECIAL_ALLOWANCE")
        if not col_earn_special:
            col_earn_special = col_fix_special
        col_earn_bonus = find_col("EARNED_BONUS", "EARNED_STATUORY BONUS", "EARNED_STATUTORY BONUS")
        col_earn_total = find_col("EARNED_TOTAL", "EARNED_GROSS")

        # Deductions
        col_ded_pf = find_col("DEDUCTIONS_PF", "DEDUCTION_PF 12%", "PF 12%", "PF")
        col_ded_esi = find_col("DEDUCTIONS_ESI", "DEDUCTION_ESI 0.75%", "ESI 0.75%", "ESI")
        col_ded_pt = find_col("DEDUCTIONS_PT", "DEDUCTION_PT", "PT")
        col_ded_adv = find_col("DEDUCTIONS_ADV", "DEDUCTION_ADV", "ADV", "ADVANCE")
        col_ded_lwf = find_col("DEDUCTIONS_LWF", "DEDUCTION_LWF", "LWF")
        col_ded_total = find_col("DEDUCTIONS_TOTAL", "DEDUCTION_TOTAL", "TOTAL_DEDUCTION")

        # Net pay
        col_net_pay = find_col("NET_PAY", "NET PAY")

        # Employer Contribution columns
        col_er_pf = find_col("EMPLOYER CONTRIBUTION_PF  13%", "EMPLOYER CONTRIBUTION_PF 13%", "EMPLOYER CONTRIBUTION_PF", "EMPLOYER_PF")
        col_er_esi = find_col("EMPLOYER CONTRIBUTION_ESI  3.25%", "EMPLOYER CONTRIBUTION_ESI 3.25%", "EMPLOYER CONTRIBUTION_ESI", "EMPLOYER_ESI")
        col_er_lww = find_col("EMPLOYER CONTRIBUTION_LWW", "EMPLOYER_LWW", "LWW")
        col_er_statu_bonus = find_col("EMPLOYER CONTRIBUTION_STATU BONUS", "EMPLOYER_STATU_BONUS", "STATU BONUS")
        col_er_total = find_col("EMPLOYER CONTRIBUTION_TOTAL", "EMPLOYER_TOTAL")

        has_employer_contribution = any([col_er_pf, col_er_esi, col_er_lww, col_er_statu_bonus, col_er_total])

        # Validate essential columns
        if not col_name:
            return jsonify({"error": "❌ Cannot generate payslips: 'NAME' column missing from uploaded file."}), 400
        if not (col_fix_basic or col_fix_basic_da or col_earn_basic or col_earn_basic_da):
            return jsonify({"error": "❌ Cannot generate payslips: Basic salary column missing from uploaded file."}), 400
        if not (col_net_pay or col_earn_total or col_fix_total):
            return jsonify({"error": "❌ Cannot generate payslips: Net Pay / Total column missing from uploaded file."}), 400

        print(f"Loaded {len(df)} employees from file")
        print(f"Has Employer Contribution Section: {has_employer_contribution}")
        print(f"  Earned other allowance col: {col_earn_other}")
        print(f"  Earned special allowance col: {col_earn_special}")
        print(f"  Employer LWW col: {col_er_lww}")
        print(f"  Employer STATU BONUS col: {col_er_statu_bonus}")
        
        env = Environment(loader=FileSystemLoader(TEMPLATE_DIR), autoescape=True)
        template = env.get_template("payslip.html")
        logo_base64 = get_logo_base64()

        preview = []
        success_count = 0
        error_count = 0
        missing_columns = set()

        for index, row in df.iterrows():
            try:
                name_val = row.get(col_name) if col_name else ""
                if pd.isna(name_val) or str(name_val).strip() == "" or str(name_val).strip().lower() in ['total', 'totals', 'grand total']:
                    print(f"Skipping row {index+1} (empty name or summary row)")
                    continue

                emp_id_val = row.get(col_emp_id) if col_emp_id else f"EMP{index+1}"
                emp_id = str(int(float(emp_id_val))) if str(emp_id_val).replace('.0','').isdigit() else str(emp_id_val).strip()
                
                pay_month = month
                print(f"Processing employee {emp_id} - {name_val}...")

                # Fixed basic / DA handling
                if col_fix_basic_da:
                    fix_basic_label = "Basic & DA"
                    fix_basic_val = get_numeric_value(row.get(col_fix_basic_da))
                    fix_da_val = 0
                    has_da_row = False
                else:
                    fix_basic_label = "Basic"
                    fix_basic_val = get_numeric_value(row.get(col_fix_basic)) if col_fix_basic else 0
                    fix_da_val = get_numeric_value(row.get(col_fix_da)) if col_fix_da else 0
                    has_da_row = bool(col_fix_da or col_earn_da)

                # Earned basic / DA handling
                if col_earn_basic_da:
                    earn_basic_val = get_numeric_value(row.get(col_earn_basic_da))
                    earn_da_val = 0
                else:
                    earn_basic_val = get_numeric_value(row.get(col_earn_basic)) if col_earn_basic else 0
                    earn_da_val = get_numeric_value(row.get(col_earn_da)) if col_earn_da else 0

                salary_fixed = {
                    "basic": fix_basic_val,
                    "da": fix_da_val,
                    "hra": get_numeric_value(row.get(col_fix_hra)) if col_fix_hra else 0,
                    "leave_wages": get_numeric_value(row.get(col_fix_leave)) if col_fix_leave else 0,
                    "others": get_numeric_value(row.get(col_fix_other)) if col_fix_other else 0,
                    "special_allowance": get_numeric_value(row.get(col_fix_special)) if col_fix_special else 0,
                    "bonus": get_numeric_value(row.get(col_fix_bonus)) if col_fix_bonus else 0,
                    "total": get_numeric_value(row.get(col_fix_total)) if col_fix_total else 0,
                }

                salary_earned = {
                    "basic": earn_basic_val,
                    "da": earn_da_val,
                    "hra": get_numeric_value(row.get(col_earn_hra)) if col_earn_hra else 0,
                    "leave_wages": get_numeric_value(row.get(col_earn_leave)) if col_earn_leave else 0,
                    "others": get_numeric_value(row.get(col_earn_other)) if col_earn_other else 0,
                    "special_allowance": get_numeric_value(row.get(col_earn_special)) if col_earn_special else 0,
                    "bonus": get_numeric_value(row.get(col_earn_bonus)) if col_earn_bonus else 0,
                    "total": get_numeric_value(row.get(col_earn_total)) if col_earn_total else 0,
                }

                # Calculate totals if not present
                if salary_fixed["total"] == 0:
                    salary_fixed["total"] = (salary_fixed["basic"] + salary_fixed["da"] + salary_fixed["hra"] +
                                            salary_fixed["leave_wages"] + salary_fixed["others"] +
                                            salary_fixed["special_allowance"] + salary_fixed["bonus"])

                if salary_earned["total"] == 0:
                    salary_earned["total"] = (salary_earned["basic"] + salary_earned["da"] + salary_earned["hra"] +
                                             salary_earned["leave_wages"] + salary_earned["others"] +
                                             salary_earned["special_allowance"] + salary_earned["bonus"])

                deduction = {
                    "pf": get_numeric_value(row.get(col_ded_pf)) if col_ded_pf else 0,
                    "esi": get_numeric_value(row.get(col_ded_esi)) if col_ded_esi else 0,
                    "pt": get_numeric_value(row.get(col_ded_pt)) if col_ded_pt else 0,
                    "adv": get_numeric_value(row.get(col_ded_adv)) if col_ded_adv else 0,
                    "lwf": get_numeric_value(row.get(col_ded_lwf)) if col_ded_lwf else 0,
                    "total": get_numeric_value(row.get(col_ded_total)) if col_ded_total else 0,
                }

                if deduction["total"] == 0:
                    deduction["total"] = deduction["pf"] + deduction["esi"] + deduction["pt"] + deduction["adv"] + deduction["lwf"]

                net_pay = get_numeric_value(row.get(col_net_pay)) if col_net_pay else (salary_earned["total"] - deduction["total"])
                net_pay_words = number_to_words(net_pay)

                # Dynamic earnings items (only included if found in uploaded excel or > 0)
                earnings_items = [
                    {"name": fix_basic_label, "fixed": salary_fixed["basic"], "earned": salary_earned["basic"]}
                ]
                if has_da_row:
                    earnings_items.append({"name": "DA", "fixed": salary_fixed["da"], "earned": salary_earned["da"]})
                if col_fix_hra or col_earn_hra or salary_fixed["hra"] > 0 or salary_earned["hra"] > 0:
                    earnings_items.append({"name": "HRA", "fixed": salary_fixed["hra"], "earned": salary_earned["hra"]})
                if col_fix_leave or col_earn_leave or salary_fixed["leave_wages"] > 0 or salary_earned["leave_wages"] > 0:
                    earnings_items.append({"name": "Leave with wages", "fixed": salary_fixed["leave_wages"], "earned": salary_earned["leave_wages"]})
                if col_fix_other or col_earn_other or salary_fixed["others"] > 0 or salary_earned["others"] > 0:
                    earnings_items.append({"name": "Other Allowance", "fixed": salary_fixed["others"], "earned": salary_earned["others"]})
                if col_fix_special or col_earn_special or salary_fixed["special_allowance"] > 0 or salary_earned["special_allowance"] > 0:
                    earnings_items.append({"name": "Special Allowance", "fixed": salary_fixed["special_allowance"], "earned": salary_earned["special_allowance"]})
                if col_fix_bonus or col_earn_bonus or salary_fixed["bonus"] > 0 or salary_earned["bonus"] > 0:
                    earnings_items.append({"name": "Bonus", "fixed": salary_fixed["bonus"], "earned": salary_earned["bonus"]})

                # Dynamic deductions items
                deductions_items = [
                    {"name": "Provident Fund", "amount": deduction["pf"]},
                    {"name": "ESI", "amount": deduction["esi"]},
                    {"name": "Professional Tax", "amount": deduction["pt"]},
                ]
                if col_ded_adv or deduction["adv"] > 0:
                    deductions_items.append({"name": "ADV", "amount": deduction["adv"]})
                if col_ded_lwf or deduction["lwf"] > 0:
                    deductions_items.append({"name": "LWF", "amount": deduction["lwf"]})

                # Pair earnings and deductions rows for balanced table
                max_rows = max(len(earnings_items), len(deductions_items))
                salary_rows = []
                for i in range(max_rows):
                    salary_rows.append({
                        "earning": earnings_items[i] if i < len(earnings_items) else None,
                        "deduction": deductions_items[i] if i < len(deductions_items) else None,
                    })

                # Employer Contribution values
                er_pf_val = get_numeric_value(row.get(col_er_pf)) if col_er_pf else 0
                er_esi_val = get_numeric_value(row.get(col_er_esi)) if col_er_esi else 0
                er_lww_val = get_numeric_value(row.get(col_er_lww)) if col_er_lww else 0
                er_statu_bonus_val = get_numeric_value(row.get(col_er_statu_bonus)) if col_er_statu_bonus else 0
                er_total_val = get_numeric_value(row.get(col_er_total)) if col_er_total else (er_pf_val + er_esi_val + er_lww_val + er_statu_bonus_val)

                employer_contribution = {
                    "has_data": has_employer_contribution,
                    "has_pf": bool(col_er_pf),
                    "pf": er_pf_val,
                    "has_esi": bool(col_er_esi),
                    "esi": er_esi_val,
                    "has_lww": bool(col_er_lww),
                    "lww": er_lww_val,
                    "has_statu_bonus": bool(col_er_statu_bonus),
                    "statu_bonus": er_statu_bonus_val,
                    "has_total": bool(col_er_total or has_employer_contribution),
                    "total": er_total_val,
                }

                emp_data = {
                    "emp_id": emp_id,
                    "name": str(name_val).strip(),
                    "designation": str(row.get(col_designation, "")).strip() if col_designation and pd.notna(row.get(col_designation)) else "",
                    "unit_name": str(row.get(col_unit, "")).strip() if col_unit and pd.notna(row.get(col_unit)) else "",
                    "uan": (lambda v: str(int(float(v))) if str(v).strip().replace('.','',1).isdigit() else str(v).strip())(row.get(col_uan, "")) if col_uan and pd.notna(row.get(col_uan)) else "",
                    "esi": str(row.get(col_esi_no, "")).strip() if col_esi_no and pd.notna(row.get(col_esi_no)) else "",
                    "doj": str(row.get(col_doj, "")).strip() if col_doj and pd.notna(row.get(col_doj)) else "",
                    "bank_ac": (lambda v: str(int(float(v))) if str(v).strip().replace('.','',1).isdigit() else str(v).strip())(row.get(col_bank_ac, "")) if col_bank_ac and pd.notna(row.get(col_bank_ac)) else "",
                    "ifsc": str(row.get(col_ifsc, "")).strip() if col_ifsc and pd.notna(row.get(col_ifsc)) else "",
                    "email": str(row.get(col_email, "")).strip() if col_email and pd.notna(row.get(col_email)) else "",
                    "phone": str(int(float(row.get(col_phone, 0)))).strip() 
                             if col_phone and pd.notna(row.get(col_phone)) 
                             and str(row.get(col_phone, "")).strip() not in ["", "nan", "0"] 
                             else "",
                    "basic_days": str(int(float(row.get(col_basic_days, 31)))) if col_basic_days and pd.notna(row.get(col_basic_days)) else "31",
                    "actual_days": str(int(float(row.get(col_actual_days, 31)))) if col_actual_days and pd.notna(row.get(col_actual_days)) else "31",
                }

                html_content = template.render(
                    company=COMPANY, emp=emp_data, salary_fixed=salary_fixed,
                    salary_earned=salary_earned, deduction=deduction,
                    salary_rows=salary_rows,
                    has_employer_contribution=has_employer_contribution,
                    employer_contribution=employer_contribution,
                    net_pay=net_pay, net_pay_words=net_pay_words, month=pay_month,
                    generated_on=datetime.now().strftime("%d %b %Y"), logo_base64=logo_base64
                )

                html_path = os.path.join(current_output_dir, f"{emp_id}.html")
                pdf_path  = os.path.join(current_output_dir, f"{emp_id}.pdf")

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

                # Upload to R2
                try:
                    emp_name = emp_data["name"]
                    unit_name = emp_data["unit_name"] or "NoUnit"
                    print(f"DEBUG: Uploading to R2 -> {year}/{pay_month}/{emp_name}_{unit_name}.pdf")
                    s3_key = upload_with_cleanup(
                        local_path=pdf_path,
                        employee_name=emp_name,
                        unit_name=unit_name,
                        month=pay_month,
                        year=year
                    )
                    print(f"[OK] Uploaded to R2: {s3_key}")
                    current_session_pdfs.append(s3_key)
                except Exception as s3_error:
                    print(f"[ERROR] R2 upload failed for {emp_id}: {s3_error}")
                    traceback.print_exc()

                preview.append({"EMP_ID": emp_id, "Name": emp_data["name"], "Designation": emp_data["designation"],
                    "Email": emp_data["email"], "Phone": emp_data["phone"],"Net_Pay": net_pay, "PDF_Path": pdf_path})
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
            error_msg = "No payslips generated.\n\n"
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
                pdf_path = os.path.join(current_output_dir, f"{emp_id}.pdf")
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

@app.route("/send-whatsapp", methods=["POST"])
def send_whatsapp():
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
            emp_name = emp.get("Name")
            emp_id = emp.get("EMP_ID")
            phone = emp.get("Phone")  # must exist in your Excel as Phone column
            pdf_path = emp.get("PDF_Path")

            if not phone or not pdf_path:
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Reason": "Missing phone or PDF"})
                continue

            if not os.path.exists(pdf_path):
                pdf_path = os.path.join(current_output_dir, f"{emp_id}.pdf")

            if not os.path.exists(pdf_path):
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Reason": "PDF not found"})
                continue

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

            pdf_filename = f"Payslip_{month}_{emp_name.replace(' ', '_')}.pdf"

            success = send_payslip_whatsapp(
                phone_number=str(phone),
                emp_name=emp_name,
                month=month,
                pdf_bytes=pdf_bytes,
                pdf_filename=pdf_filename
            )

            if success:
                sent_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Sent", "Phone": phone})
            else:
                failed_count += 1
                results.append({"EMP_ID": emp_id, "Status": "Failed", "Phone": phone})

        return jsonify({
            "message": f"WhatsApp sent: {sent_count}, failed: {failed_count}",
            "sent_count": sent_count,
            "failed_count": failed_count,
            "results": results
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500    
app.secret_key = os.getenv("SECRET_KEY", "change-this-secret")

ADMIN_USER = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASS = os.getenv("ADMIN_PASSWORD", "rsmantech123")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")
    data = request.get_json()
    if data.get("username") == ADMIN_USER and data.get("password") == ADMIN_PASS:
        session["logged_in"] = True
        return jsonify({"success": True})
    return jsonify({"error": "Invalid credentials"}), 401

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# Add this decorator to protect your dashboard route
@app.route("/")
def dashboard():
    if not session.get("logged_in"):
        return redirect(url_for("login"))
    return render_template("dashboard.html")    

@app.route("/download", methods=["GET"])
def download_pdfs():
    try:
        month = request.args.get("month")
        year = request.args.get("year")
        print(f"DEBUG: Searching S3 for year/month: {year}/{month}")
        s3_pdf_keys = list_s3_pdfs(month=month, year=year)
        print(f"DEBUG: Found keys: {s3_pdf_keys}")
        
        if not s3_pdf_keys:
            return jsonify({"error": "No PDF files found"}), 404

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for s3_key in s3_pdf_keys:
                pdf_data = download_s3_file_to_memory(s3_key)
                zipf.writestr(os.path.basename(s3_key), pdf_data.read())

        zip_buffer.seek(0)
        filename = f'payslips_{year}_{month}.zip' if year and month else f'payslips_{month}.zip' if month else 'payslips.zip'
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    print("\n" + "="*80)
    print("PAYSLIP GENERATOR STARTING")
    print("="*80 + "\n")
    app.run(debug=True)
