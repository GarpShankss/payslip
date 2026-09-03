import os
import sys
import subprocess
import base64
import re
import pandas as pd
from datetime import datetime
from jinja2 import Environment, FileSystemLoader

# Fix Windows console unicode issues
sys.stdout.reconfigure(encoding="utf-8")

# -----------------------------
# PATHS
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
OUTPUT_DIR = os.path.join(BASE_DIR, "payslips")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

if os.path.exists('/usr/local/bin/wkhtmltopdf'):
    WKHTMLTOPDF_CMD = '/usr/local/bin/wkhtmltopdf'
elif os.path.exists('/usr/bin/wkhtmltopdf'):
    WKHTMLTOPDF_CMD = '/usr/bin/wkhtmltopdf'
else:
    WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# -----------------------------
# JINJA SETUP
# -----------------------------
env = Environment(
    loader=FileSystemLoader(TEMPLATE_DIR),
    autoescape=True
)

COMPANY = {
    "name": "RS MAN-TECH",
    "address": "#14, 3rd Cross, Parappana Agrahara",
    "city": "Bengaluru-100"
}

def get_logo_base64():
    try:
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as img_file:
                return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Could not load logo: {e}")
    return None

def get_numeric_value(val, default=0.0):
    try:
        if pd.isna(val):
            return default
        return float(val)
    except:
        return default

def format_hours_value(val):
    if pd.isna(val) or str(val).strip() in ["", "nan", "None", "-", "0"]:
        return "0"
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return f"{f:.2f}".rstrip('0').rstrip('.')
    except:
        return str(val).strip()

# -----------------------------
# NUMBER TO WORDS CONVERSION
# -----------------------------
def number_to_words(num):
    """Convert number to Indian rupees format in words"""
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

# -----------------------------
# MAIN
# -----------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("CSV or XLSX file path required")
        sys.exit(1)

    file_path = sys.argv[1]

    # -----------------------------
    # READ CSV / EXCEL SAFELY
    # -----------------------------
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        try:
            df = pd.read_csv(file_path, engine="python", encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, engine="python", encoding="latin1")
    elif ext in [".xlsx", ".xls"]:
        header_row = None
        for row_num in range(10):
            try:
                test_df = pd.read_excel(file_path, header=row_num, nrows=1)
                cols_lower = [str(c).lower() for c in test_df.columns]
                if any('emp' in c or 'name' in c or 'id' in c or 'sl.no' in c or 'sl no' in c for c in cols_lower if not c.startswith('unnamed')):
                    header_row = row_num
                    break
            except:
                continue

        if header_row is not None:
            try:
                next_row_df = pd.read_excel(file_path, header=header_row+1, nrows=1)
                next_cols = [str(c).lower() for c in next_row_df.columns]
                if any('fixed' in c or 'earned' in c or 'deduction' in c or 'basic' in c or 'uan' in c for c in next_cols if not c.startswith('unnamed')):
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
        print("Unsupported file format")
        sys.exit(1)

    df.columns = df.columns.str.strip().str.replace('\ufeff', '')
    df = df.dropna(how='all')

    print(f"Loaded {len(df)} records")

    # Column mapping
    col_map = {}
    for col in df.columns:
        c_str = str(col).strip()
        col_map[c_str] = col
        col_map[c_str.lower()] = col
        col_lower_slug = c_str.lower().replace(' ', '_')
        col_map[col_lower_slug] = col
        norm_slug = re.sub(r'[^a-z0-9]+', '_', c_str.lower()).strip('_')
        if norm_slug:
            col_map[norm_slug] = col
        
        if 'deductions_' in col_lower_slug or 'deduction_' in col_lower_slug:
            cleaned = col_lower_slug.replace('deductions_', '').replace('deduction_', '')
            col_map[cleaned] = col
        if 'net_pay_' in col_lower_slug:
            col_map[col_lower_slug.replace('net_pay_', '')] = col

    def find_col(*candidates):
        for c in candidates:
            c_str = str(c).strip()
            if c_str in col_map:
                return col_map[c_str]
            if c_str.lower() in col_map:
                return col_map[c_str.lower()]
            c_slug = c_str.lower().replace(' ', '_')
            if c_slug in col_map:
                return col_map[c_slug]
            norm_slug = re.sub(r'[^a-z0-9]+', '_', c_str.lower()).strip('_')
            if norm_slug in col_map:
                return col_map[norm_slug]
        return None

    # Detect columns
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

    # Overtime (OT)
    col_ot_hrs = find_col("OT_HRS", "OT HRS", "OT_HOURS", "OT HOURS", "OT HR", "OT_HR", "OVERTIME_HRS", "OVERTIME HOURS", "OT_DAYS", "OT DAYS")
    col_ot_cost = find_col("OT_COST", "OT COST", "OT_AMT", "OT AMT", "OT_AMOUNT", "OT AMOUNT", "EARNED_OT", "EARNED_OVERTIME", "EARNED_OT_COST", "EARNED_OT_AMOUNT", "OVERTIME_COST", "OVERTIME_AMOUNT", "OVERTIME_AMT", "FIXED_OT", "FIXED_OVERTIME", "OT", "OVERTIME")
    has_ot_col = bool(col_ot_hrs or col_ot_cost)

    # Fixed salary columns
    col_fix_basic_da = find_col("FIXED_BASIC & DA", "FIXED_BASIC_DA", "BASIC & DA")
    col_fix_basic = find_col("FIXED_BASIC", "FIXED_BASIC SALARY", "BASIC")
    col_fix_da = find_col("FIXED_DA", "DA")
    col_fix_hra = find_col("FIXED_HRA", "HRA")
    col_fix_leave = find_col("FIXED_LEAVE_WAGES", "FIXED_LEAVE WAGES", "FIXED_LEAVE_WAGE", "FIXED_LEAVE", "FIXED_LWW", "LEAVE_WAGES", "LEAVE WAGES", "Leave_Wages", "LEAVE WITH WAGES", "LEAVE_WITH_WAGES", "LWW", "LEAVE")
    col_fix_other = find_col("FIXED_OTHER ALLOWANCE", "FIXED_OTHER_ALLOWANCE", "FIXED_OTHER_ALLOW", "FIXED_OTHER ALLOW", "FIXED_OTHERS", "FIXED_OTHER", "OTHER ALLOWANCE", "OTHER_ALLOWANCE", "OTHER ALLOW", "OTHER_ALLOW", "OTHERS", "OTHER", "OTHER_ALLOWANCES", "OTHER ALLOWANCES")
    col_fix_special = find_col("FIXED_SPECIAL ALLOW", "FIXED_SPECIAL ALLOWANCE", "FIXED_SPECIAL_ALLOWANCE", "FIXED_SPECIAL_ALLOW", "FIXED_SPECIAL", "SPECIAL ALLOW", "SPECIAL ALLOWANCE", "SPECIAL_ALLOW", "SPECIAL_ALLOWANCE", "SPECIAL", "SPL_ALLOW", "SPL_ALLOWANCE", "FIXED_SPL_ALLOW", "FIXED_SPL_ALLOWANCE")
    col_fix_tpt = find_col("FIXED_TPT", "FIXED_TPT_ALLOWANCE", "FIXED_TPT ALLOWANCE", "FIXED_TRANSPORT", "FIXED_TRANSPORT ALLOWANCE", "FIXED_TRANSPORT_ALLOWANCE", "FIXED_CONVEYANCE", "FIXED_CONVEYANCE ALLOWANCE", "TPT", "TPT_ALLOWANCE", "TPT ALLOWANCE", "TPT ALLOW", "TRANSPORT", "TRANSPORT ALLOWANCE", "TRANSPORT_ALLOWANCE", "CONVEYANCE", "CONVEYANCE ALLOWANCE", "CONVEYANCE_ALLOWANCE", "TRAVELLING ALLOWANCE", "TRAVEL ALLOWANCE")
    col_fix_bonus = find_col("FIXED_BONUS", "FIXED_STATUORY BONUS", "FIXED_STATUTORY BONUS", "STATUORY BONUS", "STATUTORY BONUS")
    col_fix_total = find_col("FIXED_TOTAL", "FIXED_GROSS", "TOTAL")

    # Earned salary columns
    col_earn_basic_da = find_col("EARNED_BASIC & DA", "EARNED_BASIC_DA")
    col_earn_basic = find_col("EARNED_BASIC")
    col_earn_da = find_col("EARNED_DA")
    col_earn_hra = find_col("EARNED_HRA")
    col_earn_leave = find_col("EARNED_LEAVE_WAGES", "EARNED_LEAVE WAGES", "Earned_Leave_Wages", "EARNED_LEAVE_WAGE", "EARNED_LEAVE", "EARNED_LWW", "EARNED_LEAVE WITH WAGES", "EARNED_LEAVE_WITH_WAGES")
    if not col_earn_leave and col_fix_leave:
        col_earn_leave = col_fix_leave
    col_earn_other = find_col("EARNED_OTHER ALLOWANCE", "EARNED_OTHER_ALLOWANCE", "EARNED_OTHER_ALLOW", "EARNED_OTHER ALLOW", "EARNED_OTHERS", "EARNED_OTHER", "EARNED_OTHER_ALLOWANCES", "EARNED_OTHER ALLOWANCES")
    if not col_earn_other and col_fix_other:
        col_earn_other = col_fix_other
    col_earn_special = find_col("EARNED_SPECIAL ALLOW", "EARNED_SPECIAL ALLOWANCE", "EARNED_SPECIAL_ALLOWANCE", "EARNED_SPECIAL_ALLOW", "EARNED_SPECIAL", "EARNED_SPL_ALLOW", "EARNED_SPL_ALLOWANCE")
    if not col_earn_special and col_fix_special:
        col_earn_special = col_fix_special
    col_earn_tpt = find_col("EARNED_TPT", "EARNED_TPT_ALLOWANCE", "EARNED_TPT ALLOWANCE", "EARNED_TRANSPORT", "EARNED_TRANSPORT ALLOWANCE", "EARNED_TRANSPORT_ALLOWANCE", "EARNED_CONVEYANCE", "EARNED_CONVEYANCE ALLOWANCE")
    if not col_earn_tpt and col_fix_tpt:
        col_earn_tpt = col_fix_tpt
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
    col_er_pf = find_col("EMPLOYER CONTRIBUTION_PF  13%", "EMPLOYER CONTRIBUTION_PF 13%", "EMPLOYER CONTRIBUTION_PF", "EMPLOYER_PF", "ER_PF")
    col_er_esi = find_col("EMPLOYER CONTRIBUTION_ESI  3.25%", "EMPLOYER CONTRIBUTION_ESI 3.25%", "EMPLOYER CONTRIBUTION_ESI", "EMPLOYER_ESI", "ER_ESI")
    col_er_lww = find_col("EMPLOYER CONTRIBUTION_LWW", "EMPLOYER CONTRIBUTION_LEAVE WAGES", "EMPLOYER CONTRIBUTION_LEAVE WITH WAGES", "EMPLOYER_LWW", "EMPLOYER_LEAVE_WAGES", "EMPLOYER_LEAVE_WITH_WAGES", "ER_LWW", "ER_LEAVE_WAGES")
    col_er_statu_bonus = find_col("EMPLOYER CONTRIBUTION_STATU BONUS", "EMPLOYER CONTRIBUTION_STATUTORY BONUS", "EMPLOYER_STATU_BONUS", "EMPLOYER_STATUTORY_BONUS", "ER_STATU_BONUS", "STATU BONUS")
    col_er_total = find_col("EMPLOYER CONTRIBUTION_TOTAL", "EMPLOYER_TOTAL", "ER_TOTAL")

    has_employer_contribution = any([col_er_pf, col_er_esi, col_er_lww, col_er_statu_bonus, col_er_total])

    print(f"Has OT columns: {has_ot_col} (Hours: {col_ot_hrs}, Cost: {col_ot_cost})")
    print(f"Has Employer Contribution Section: {has_employer_contribution}")
    print(f"  Earned other allowance col: {col_earn_other}")
    print(f"  Earned special allowance col: {col_earn_special}")
    print(f"  Earned TPT/Transport allowance col: {col_earn_tpt}")
    print(f"  Employer LWW col: {col_er_lww}")
    print(f"  Employer STATU BONUS col: {col_er_statu_bonus}")

    template = env.get_template("payslip.html")
    logo_base64 = get_logo_base64()

    # -----------------------------
    # GENERATE PAYSLIPS
    # -----------------------------
    for index, row in df.iterrows():
        name_val = row.get(col_name) if col_name else ""
        if pd.isna(name_val) or str(name_val).strip() == "" or str(name_val).strip().lower() in ['total', 'totals', 'grand total']:
            continue

        emp_id_val = row.get(col_emp_id) if col_emp_id else f"EMP{index+1}"
        emp_id = str(int(float(emp_id_val))) if str(emp_id_val).replace('.0','').isdigit() else str(emp_id_val).strip()
        month = str(row.get("Month", "NA")).strip()

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

        # OT extraction
        ot_hrs_str = format_hours_value(row.get(col_ot_hrs)) if col_ot_hrs else "0"
        ot_cost_val = get_numeric_value(row.get(col_ot_cost)) if col_ot_cost else 0
        has_ot_for_emp = bool(has_ot_col or ot_cost_val > 0 or (col_ot_hrs and ot_hrs_str not in ["0", ""]))
        ot_data = {
            "has_data": has_ot_for_emp,
            "hrs": ot_hrs_str,
            "cost": ot_cost_val
        }

        salary_fixed = {
            "basic": fix_basic_val,
            "da": fix_da_val,
            "hra": get_numeric_value(row.get(col_fix_hra)) if col_fix_hra else 0,
            "leave_wages": get_numeric_value(row.get(col_fix_leave)) if col_fix_leave else 0,
            "others": get_numeric_value(row.get(col_fix_other)) if col_fix_other else 0,
            "special_allowance": get_numeric_value(row.get(col_fix_special)) if col_fix_special else 0,
            "tpt": get_numeric_value(row.get(col_fix_tpt)) if col_fix_tpt else 0,
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
            "tpt": get_numeric_value(row.get(col_earn_tpt)) if col_earn_tpt else 0,
            "bonus": get_numeric_value(row.get(col_earn_bonus)) if col_earn_bonus else 0,
            "total": get_numeric_value(row.get(col_earn_total)) if col_earn_total else 0,
        }

        if salary_fixed["total"] == 0:
            salary_fixed["total"] = (salary_fixed["basic"] + salary_fixed["da"] + salary_fixed["hra"] +
                                    salary_fixed["leave_wages"] + salary_fixed["others"] +
                                    salary_fixed["special_allowance"] + salary_fixed["tpt"] + salary_fixed["bonus"])

        if salary_earned["total"] == 0:
            salary_earned["total"] = (salary_earned["basic"] + salary_earned["da"] + salary_earned["hra"] +
                                     salary_earned["leave_wages"] + salary_earned["others"] +
                                     salary_earned["special_allowance"] + salary_earned["tpt"] + salary_earned["bonus"])

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
        if (col_fix_tpt or col_earn_tpt or salary_fixed["tpt"] > 0 or salary_earned["tpt"] > 0) and (col_fix_tpt != col_fix_other and col_earn_tpt != col_earn_other):
            earnings_items.append({"name": "Transport Allowance", "fixed": salary_fixed["tpt"], "earned": salary_earned["tpt"]})
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

        has_employer_data = has_employer_contribution and (er_total_val > 0 or any([col_er_pf, col_er_esi, col_er_lww, col_er_statu_bonus]))

        employer_contribution = {
            "has_data": has_employer_data,
            "has_pf": bool(col_er_pf and (er_pf_val > 0 or has_employer_contribution)),
            "pf": er_pf_val,
            "has_esi": bool(col_er_esi and (er_esi_val > 0 or has_employer_contribution)),
            "esi": er_esi_val,
            "has_lww": bool(col_er_lww and (er_lww_val > 0 or has_employer_contribution)),
            "lww": er_lww_val,
            "has_statu_bonus": bool(col_er_statu_bonus and (er_statu_bonus_val > 0 or has_employer_contribution)),
            "statu_bonus": er_statu_bonus_val,
            "has_total": bool(col_er_total or (has_employer_data and er_total_val > 0)),
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
            "basic_days": str(int(float(row.get(col_basic_days, 31)))) if col_basic_days and pd.notna(row.get(col_basic_days)) else "31",
            "actual_days": str(int(float(row.get(col_actual_days, 31)))) if col_actual_days and pd.notna(row.get(col_actual_days)) else "31",
        }

        # -----------------------------
        # RENDER TEMPLATE
        # -----------------------------
        html_content = template.render(
            company=COMPANY,
            emp=emp_data,
            salary_fixed=salary_fixed,
            salary_earned=salary_earned,
            deduction=deduction,
            salary_rows=salary_rows,
            has_employer_contribution=has_employer_contribution,
            employer_contribution=employer_contribution,
            has_ot=has_ot_for_emp,
            ot_data=ot_data,
            net_pay=net_pay,
            net_pay_words=net_pay_words,
            month=month,
            generated_on=datetime.now().strftime("%d %b %Y"),
            logo_base64=logo_base64
        )

        html_path = os.path.join(OUTPUT_DIR, f"{emp_id}.html")
        pdf_path = os.path.join(OUTPUT_DIR, f"{emp_id}.pdf")

        # Write HTML
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        # Convert to PDF
        subprocess.run(
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
            check=True
        )

        print(f"Generated: {pdf_path}")

    print("All payslips generated successfully")