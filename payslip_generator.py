import os
import sys
import subprocess
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

os.makedirs(OUTPUT_DIR, exist_ok=True)

WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# -----------------------------
# JINJA SETUP
# -----------------------------
env = Environment(
    loader=FileSystemLoader(TEMPLATE_DIR),
    autoescape=True
)

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
        if file_path.lower().endswith(".csv"):
            try:
                df = pd.read_csv(
                    file_path,
                    engine="python",
                    encoding="utf-8"
                )
            except UnicodeDecodeError:
                df = pd.read_csv(
                    file_path,
                    engine="python",
                    encoding="latin1"
                )
        else:
            df = pd.read_excel(file_path)

            # Clean column names - remove BOM and extra spaces
            df.columns = df.columns.str.strip().str.replace('\ufeff', '')

            print(f"Loaded {len(df)} records")

            # -----------------------------
            # TEMPLATE
            # -----------------------------
            template = env.get_template("payslip.html")

            COMPANY = {
                "name": "RS MAN-TECH",
                "address": "#14, 3rd Cross, Parappana Agrahara",
                "city": "Bengaluru-100"
            }

            # -----------------------------
            # GENERATE PAYSLIPS
            # -----------------------------
            for index, row in df.iterrows():

                emp_id = str(row.get("EMP_ID", f"EMP{index+1}")).strip()
                month = str(row.get("Month", "NA")).strip()

                # Helper function to safely get numeric values
def get_num(key, default=0):
    val = row.get(key, default)
    try:
        return float(val) if pd.notna(val) else default
    except:
        return default

    # -----------------------------
    # SALARY BREAKUP (Fixed vs Earned)
    # -----------------------------
    salary_fixed = {
        "basic": get_num("Fixed_Basic"),
        "da": get_num("Fixed_DA"),
        "hra": get_num("Fixed_HRA"),
        "leave_wages": 0,  # Not in CSV
        "others": 0,  # Not in CSV
        "bonus": get_num("Fixed_Bonus"),
        "total": get_num("Fixed_Total"),
    }

    salary_earned = {
        "basic": get_num("Earned_Basic"),
        "da": get_num("Earned_DA"),
        "hra": get_num("Earned_HRA"),
        "leave_wages": 0,  # Not in CSV
        "others": get_num("Other_Allowance"),
        "bonus": get_num("Earned_Bonus"),
        "total": get_num("Earned_Total"),
    }

    deduction = {
        "pf": get_num("PF"),
        "esi": get_num("ESI"),
        "pt": get_num("PT"),
        "lwf": get_num("LWF"),
        "adv": 0,  # Not in CSV
        "total": get_num("Total_Deduction"),
    }

    net_pay = get_num("Net_Pay")
    net_pay_words = number_to_words(net_pay)

    # Employee data
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

    # -----------------------------
    # RENDER TEMPLATE
    # -----------------------------

    html_content = template.render(
        company=COMPANY,
        emp=emp_data,
        salary_fixed=salary_fixed,
        salary_earned=salary_earned,
        deduction=deduction,
        net_pay=net_pay,
        net_pay_words=net_pay_words,
        month=month,
        generated_on=datetime.now().strftime("%d %b %Y")
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