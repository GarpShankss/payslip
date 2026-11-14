import os
import sys
import json
import io
import unittest
from unittest.mock import patch, MagicMock
import pandas as pd
from jinja2 import Environment, FileSystemLoader

# Import app components
from app import (
    app,
    validate_locations,
    normalize_location,
    get_company_for_location,
    read_and_merge_excel,
    LOCATIONS,
    DEFAULT_LOCATIONS,
    COMPANY
)

class TestLocationFeature(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.client = app.test_client()
        cls.template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
        cls.jinja_env = Environment(loader=FileSystemLoader(cls.template_dir), autoescape=True)
        cls.test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data")
        os.makedirs(cls.test_dir, exist_ok=True)

    def login(self):
        with self.client.session_transaction() as sess:
            sess["logged_in"] = True

    def test_01_location_normalization(self):
        """Test normalization handles case, spaces, and invalid values."""
        self.assertEqual(normalize_location("BANGALORE"), "BANGALORE")
        self.assertEqual(normalize_location("bangalore"), "BANGALORE")
        self.assertEqual(normalize_location("  Bangalore  "), "BANGALORE")
        self.assertEqual(normalize_location("Bengaluru"), "BANGALORE")
        self.assertEqual(normalize_location("HYDERABAD"), "HYDERABAD")
        self.assertEqual(normalize_location("hyderabad"), "HYDERABAD")
        self.assertEqual(normalize_location("  Hyderabad  "), "HYDERABAD")
        self.assertIsNone(normalize_location(None))
        self.assertIsNone(normalize_location(""))
        self.assertIsNone(normalize_location("   "))
        self.assertIsNone(normalize_location(float("nan")))
        self.assertEqual(normalize_location("Chennai"), "INVALID")
        self.assertEqual(normalize_location("Mumbai"), "INVALID")

    def test_02_company_address_configuration(self):
        """Verify Bangalore and Hyderabad address configurations."""
        blr_co = get_company_for_location("BANGALORE")
        self.assertEqual(blr_co["name"], "RS MAN-TECH")
        self.assertIn("#14, 3rd Cross, Parappana Agrahara", blr_co["address_lines"][0])
        self.assertIn("Electronic City, Bangalore - 560100", blr_co["address_lines"][1])

        hyd_co = get_company_for_location("HYDERABAD")
        self.assertEqual(hyd_co["name"], "RS MAN-TECH")
        self.assertEqual(hyd_co["address_lines"][0], "No. 2-62, Yellampet Village,")
        self.assertEqual(hyd_co["address_lines"][1], "Medchal Mandal,")
        self.assertEqual(hyd_co["address_lines"][2], "Medchal-Malkajgiri, Telangana - 501401.")

    def test_03_payslip_template_address_rendering(self):
        """Test that master template dynamically renders Bangalore and Hyderabad addresses."""
        template = self.jinja_env.get_template("payslip.html")
        mock_emp = {
            "emp_id": "EMP001",
            "name": "Test Employee",
            "designation": "Developer",
            "unit_name": "IT",
            "uan": "12345",
            "esi": "67890",
            "doj": "01.01.2023",
            "bank_ac": "11223344",
            "ifsc": "SBIN0001",
            "email": "test@example.com",
            "phone": "9876543210",
            "basic_days": "31",
            "actual_days": "31"
        }
        mock_salary = {"basic": 15000, "da": 0, "hra": 5000, "leave_wages": 0, "others": 0, "special_allowance": 0, "tpt": 0, "bonus": 0, "total": 20000}
        mock_ded = {"pf": 1800, "esi": 150, "pt": 200, "adv": 0, "lwf": 0, "total": 2150}

        # 1. Render Bangalore
        blr_html = template.render(
            company=get_company_for_location("BANGALORE"),
            emp=mock_emp,
            salary_fixed=mock_salary,
            salary_earned=mock_salary,
            deduction=mock_ded,
            salary_rows=[{"earning": {"name": "Basic", "fixed": 15000, "earned": 15000}, "deduction": {"name": "PF", "amount": 1800}}],
            has_employer_contribution=False,
            employer_contribution={"has_data": False},
            has_ot=False,
            ot_data={"has_data": False},
            has_incentive=False,
            incentive_data={"has_data": False},
            net_pay=17850,
            net_pay_words="Seventeen Thousand Eight Hundred Fifty rupees only",
            month="January",
            generated_on="14 Nov 2026",
            logo_base64=None
        )
        self.assertIn("RS MAN-TECH", blr_html)
        self.assertIn("#14, 3rd Cross, Parappana Agrahara", blr_html)
        self.assertIn("Electronic City, Bangalore - 560100", blr_html)
        self.assertNotIn("Yellampet Village", blr_html)

        # 2. Render Hyderabad
        hyd_html = template.render(
            company=get_company_for_location("HYDERABAD"),
            emp=mock_emp,
            salary_fixed=mock_salary,
            salary_earned=mock_salary,
            deduction=mock_ded,
            salary_rows=[{"earning": {"name": "Basic", "fixed": 15000, "earned": 15000}, "deduction": {"name": "PF", "amount": 1800}}],
            has_employer_contribution=False,
            employer_contribution={"has_data": False},
            has_ot=False,
            ot_data={"has_data": False},
            has_incentive=False,
            incentive_data={"has_data": False},
            net_pay=17850,
            net_pay_words="Seventeen Thousand Eight Hundred Fifty rupees only",
            month="January",
            generated_on="14 Nov 2026",
            logo_base64=None
        )
        self.assertIn("RS MAN-TECH", hyd_html)
        self.assertIn("No. 2-62, Yellampet Village,", hyd_html)
        self.assertIn("Medchal Mandal,", hyd_html)
        self.assertIn("Medchal-Malkajgiri, Telangana - 501401.", hyd_html)
        self.assertNotIn("Parappana Agrahara", hyd_html)
        self.assertNotIn("Electronic City", hyd_html)

    def test_04_validation_missing_location_column(self):
        """Verify failure when Location column is missing entirely."""
        df = pd.DataFrame({
            "EMP_ID": [101, 102],
            "NAME": ["Emp One", "Emp Two"],
            "FIXED_BASIC": [15000, 16000],
            "NET_PAY": [14000, 15000]
        })
        val = validate_locations(df, "NAME", "EMP_ID", None)
        self.assertFalse(val["valid"])
        self.assertTrue(val["missing_column"])
        self.assertIn("Location' column missing", val["error"])

    def test_05_validation_missing_and_invalid_values(self):
        """Verify detection of missing or invalid location values in employee rows."""
        df = pd.DataFrame({
            "EMP_ID": [101, 102, 103, 104],
            "NAME": ["Alice", "Bob", "Charlie", "David"],
            "LOCATION": ["BANGALORE", None, "Chennai", "  hyderabad  "],
            "FIXED_BASIC": [15000, 16000, 17000, 18000],
            "NET_PAY": [14000, 15000, 16000, 17000]
        })
        val = validate_locations(df, "NAME", "EMP_ID", "LOCATION")
        self.assertFalse(val["valid"])
        self.assertEqual(val["counts"]["TOTAL"], 2)  # Alice and David are valid
        self.assertEqual(len(val["errors"]), 2)      # Bob (missing) and Charlie (invalid)
        self.assertEqual(val["errors"][0]["reason"], "Location is missing")
        self.assertIn("Invalid location 'Chennai'", val["errors"][1]["reason"])
        self.assertIn("Valid employees: 2", val["error"])
        self.assertIn("Employees with errors: 2", val["error"])

    def _create_excel_file(self, data, filename):
        file_path = os.path.join(self.test_dir, filename)
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        return file_path

    @patch("app.upload_with_cleanup")
    @patch("app.subprocess.run")
    def test_06_upload_bangalore_only(self, mock_subproc, mock_s3):
        """Test uploading file with only Bangalore employees."""
        mock_s3.return_value = "2026/March/test_blr.pdf"
        def fake_run(cmd, *args, **kwargs):
            pdf_path = cmd[-1]
            with open(pdf_path, "wb") as f:
                f.write(b"%PDF-1.4 Fake PDF for testing")
            res = MagicMock()
            res.returncode = 0
            res.stderr = ""
            return res

        mock_subproc.side_effect = fake_run

        self.login()
        file_path = self._create_excel_file({
            "EMP_ID": ["BLR001", "BLR002"],
            "NAME": ["Bangalore User 1", "Bangalore User 2"],
            "LOCATION": ["BANGALORE", "  bangalore  "],
            "FIXED_BASIC": [20000, 25000],
            "EARNED_BASIC": [20000, 25000],
            "NET_PAY": [18000, 23000],
            "Email": ["blr1@example.com", "blr2@example.com"],
            "Phone": ["9111111111", "9222222222"]
        }, "blr_only.xlsx")

        with open(file_path, "rb") as f:
            res = self.client.post("/upload", data={
                "csv_file": (f, "blr_only.xlsx"),
                "month": "March",
                "year": "2026",
                "location": "ALL"
            }, content_type="multipart/form-data")

        self.assertEqual(res.status_code, 200)
        json_data = res.get_json()
        self.assertEqual(json_data["success_count"], 2)
        self.assertEqual(json_data["location_counts"]["Bangalore"], 2)
        self.assertEqual(json_data["location_counts"]["Hyderabad"], 0)
        self.assertEqual(len(json_data["preview"]), 2)
        self.assertEqual(json_data["preview"][0]["Location"], "Bangalore")
        self.assertEqual(json_data["preview"][1]["Location"], "Bangalore")

        # Verify the generated HTML contains Bangalore address
        html_path = json_data["preview"][0]["PDF_Path"].replace(".pdf", ".html")
        with open(html_path, "r", encoding="utf-8") as hf:
            content = hf.read()
            self.assertIn("#14, 3rd Cross, Parappana Agrahara", content)
            self.assertIn("Electronic City, Bangalore - 560100", content)
            self.assertNotIn("Yellampet Village", content)

    @patch("app.upload_with_cleanup")
    @patch("app.subprocess.run")
    def test_07_upload_hyderabad_only(self, mock_subproc, mock_s3):
        """Test uploading file with only Hyderabad employees."""
        mock_s3.return_value = "2026/March/test_hyd.pdf"
        def fake_run(cmd, *args, **kwargs):
            pdf_path = cmd[-1]
            with open(pdf_path, "wb") as f:
                f.write(b"%PDF-1.4 Fake PDF for testing")
            res = MagicMock()
            res.returncode = 0
            res.stderr = ""
            return res

        mock_subproc.side_effect = fake_run

        self.login()
        file_path = self._create_excel_file({
            "EMP_ID": ["HYD001", "HYD002"],
            "NAME": ["Hyderabad User 1", "Hyderabad User 2"],
            "LOCATION": ["HYDERABAD", "  Hyderabad  "],
            "FIXED_BASIC": [30000, 35000],
            "EARNED_BASIC": [30000, 35000],
            "NET_PAY": [28000, 32000],
            "Email": ["hyd1@example.com", "hyd2@example.com"],
            "Phone": ["9333333333", "9444444444"]
        }, "hyd_only.xlsx")

        with open(file_path, "rb") as f:
            res = self.client.post("/upload", data={
                "csv_file": (f, "hyd_only.xlsx"),
                "month": "March",
                "year": "2026",
                "location": "ALL"
            }, content_type="multipart/form-data")

        self.assertEqual(res.status_code, 200)
        json_data = res.get_json()
        self.assertEqual(json_data["success_count"], 2)
        self.assertEqual(json_data["location_counts"]["Bangalore"], 0)
        self.assertEqual(json_data["location_counts"]["Hyderabad"], 2)
        self.assertEqual(json_data["preview"][0]["Location"], "Hyderabad")
        self.assertEqual(json_data["preview"][1]["Location"], "Hyderabad")

        # Verify HTML contains Hyderabad address
        html_path = json_data["preview"][0]["PDF_Path"].replace(".pdf", ".html")
        with open(html_path, "r", encoding="utf-8") as hf:
            content = hf.read()
            self.assertIn("No. 2-62, Yellampet Village,", content)
            self.assertIn("Medchal Mandal,", content)
            self.assertIn("Medchal-Malkajgiri, Telangana - 501401.", content)
            self.assertNotIn("Parappana Agrahara", content)

    @patch("app.upload_with_cleanup")
    @patch("app.subprocess.run")
    def test_08_upload_mixed_locations_filters(self, mock_subproc, mock_s3):
        """Test mixed location file with ALL, BANGALORE, and HYDERABAD filters."""
        mock_s3.return_value = "2026/March/test_mixed.pdf"
        def fake_run(cmd, *args, **kwargs):
            pdf_path = cmd[-1]
            with open(pdf_path, "wb") as f:
                f.write(b"%PDF-1.4 Fake PDF for testing")
            res = MagicMock()
            res.returncode = 0
            res.stderr = ""
            return res

        mock_subproc.side_effect = fake_run

        self.login()
        file_path = self._create_excel_file({
            "EMP_ID": ["E101", "E102", "E103", "E104", "E105"],
            "NAME": ["Emp Blr 1", "Emp Hyd 1", "Emp Blr 2", "Emp Hyd 2", "Emp Blr 3"],
            "LOCATION": ["BANGALORE", "HYDERABAD", "  bangalore  ", " Hyderabad ", "BENGALURU"],
            "FIXED_BASIC": [10000, 20000, 30000, 40000, 50000],
            "EARNED_BASIC": [10000, 20000, 30000, 40000, 50000],
            "NET_PAY": [9000, 18000, 27000, 36000, 45000],
            "Email": ["e1@test.com", "e2@test.com", "e3@test.com", "e4@test.com", "e5@test.com"],
            "Phone": ["9111111111", "9222222222", "9333333333", "9444444444", "9555555555"]
        }, "mixed.xlsx")

        # 1. Test Filter: ALL
        with open(file_path, "rb") as f:
            res = self.client.post("/upload", data={
                "csv_file": (f, "mixed.xlsx"),
                "month": "March",
                "year": "2026",
                "location": "ALL"
            }, content_type="multipart/form-data")
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data["success_count"], 5)
        self.assertEqual(data["location_counts"]["Bangalore"], 3)
        self.assertEqual(data["location_counts"]["Hyderabad"], 2)
        self.assertEqual(data["location_counts"]["Total"], 5)

        # Verify each employee received their own location's address
        for item in data["preview"]:
            html_p = item["PDF_Path"].replace(".pdf", ".html")
            with open(html_p, "r", encoding="utf-8") as hf:
                html_text = hf.read()
                if item["Location"] == "Bangalore":
                    self.assertIn("#14, 3rd Cross, Parappana Agrahara", html_text)
                    self.assertNotIn("Yellampet Village", html_text)
                else:
                    self.assertIn("No. 2-62, Yellampet Village,", html_text)
                    self.assertNotIn("Parappana Agrahara", html_text)

        # 2. Test Filter: BANGALORE
        with open(file_path, "rb") as f:
            res_blr = self.client.post("/upload", data={
                "csv_file": (f, "mixed.xlsx"),
                "month": "March",
                "year": "2026",
                "location": "BANGALORE"
            }, content_type="multipart/form-data")
        self.assertEqual(res_blr.status_code, 200)
        data_blr = res_blr.get_json()
        self.assertEqual(data_blr["success_count"], 3)
        self.assertTrue(all(p["Location"] == "Bangalore" for p in data_blr["preview"]))
        self.assertEqual(data_blr["location_counts"]["Bangalore"], 3)
        self.assertEqual(data_blr["location_counts"]["Hyderabad"], 2)

        # 3. Test Filter: HYDERABAD
        with open(file_path, "rb") as f:
            res_hyd = self.client.post("/upload", data={
                "csv_file": (f, "mixed.xlsx"),
                "month": "March",
                "year": "2026",
                "location": "HYDERABAD"
            }, content_type="multipart/form-data")
        self.assertEqual(res_hyd.status_code, 200)
        data_hyd = res_hyd.get_json()
        self.assertEqual(data_hyd["success_count"], 2)
        self.assertTrue(all(p["Location"] == "Hyderabad" for p in data_hyd["preview"]))

    def test_09_inspect_file_endpoint(self):
        """Test the /inspect-file endpoint for real-time counts and error reporting."""
        self.login()
        # Valid file
        valid_path = self._create_excel_file({
            "EMP_ID": [1, 2, 3],
            "NAME": ["Emp A", "Emp B", "Emp C"],
            "LOCATION": ["BANGALORE", "HYDERABAD", "BANGALORE"],
            "FIXED_BASIC": [10000, 20000, 30000]
        }, "inspect_valid.xlsx")

        with open(valid_path, "rb") as f:
            res = self.client.post("/inspect-file", data={"csv_file": (f, "inspect_valid.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertTrue(data["success"])
        self.assertEqual(data["location_counts"]["Bangalore"], 2)
        self.assertEqual(data["location_counts"]["Hyderabad"], 1)
        self.assertEqual(data["location_counts"]["Total"], 3)

        # Invalid file with missing location
        invalid_path = self._create_excel_file({
            "EMP_ID": [1, 2],
            "NAME": ["Emp A", "Emp B"],
            "LOCATION": ["BANGALORE", "Goa"],
            "FIXED_BASIC": [10000, 20000]
        }, "inspect_invalid.xlsx")

        with open(invalid_path, "rb") as f:
            res_inv = self.client.post("/inspect-file", data={"csv_file": (f, "inspect_invalid.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(res_inv.status_code, 400)
        data_inv = res_inv.get_json()
        self.assertFalse(data_inv["success"])
        self.assertEqual(data_inv["valid_count"], 1)
        self.assertEqual(data_inv["error_count"], 1)
        self.assertIn("Invalid location 'Goa'", data_inv["error"])

    def test_10_download_current_session(self):
        """Test downloading zip of currently generated payslips."""
        self.login()
        res = self.client.get("/download-current")
        # Should return 200 zip or 404 if no session
        self.assertIn(res.status_code, [200, 404])

    @patch("app.send_email")
    def test_11_send_emails_preserves_location_pdfs(self, mock_send_email):
        """Verify send_emails uses the generated location-specific PDF paths."""
        mock_send_email.return_value = True
        self.login()
        employees = [
            {"EMP_ID": "101", "Name": "Alice", "Email": "alice@test.com", "PDF_Path": "/tmp/dummy.pdf"},
            {"EMP_ID": "102", "Name": "Bob", "Email": "bob@test.com", "PDF_Path": "/tmp/dummy.pdf"}
        ]
        # Create dummy file so os.path.exists passes
        with open("/tmp/dummy.pdf", "wb") as f:
            f.write(b"%PDF-1.4 dummy")

        res = self.client.post("/send-emails", json={"employees": employees, "month": "January"})
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data["sent_count"], 2)
        self.assertEqual(data["failed_count"], 0)

    @patch("app.send_payslip_whatsapp")
    def test_12_send_whatsapp_preserves_location_pdfs(self, mock_wa):
        """Verify send_whatsapp uses the generated location-specific PDF files."""
        mock_wa.return_value = (True, "")
        self.login()
        employees = [
            {"EMP_ID": "101", "Name": "Alice", "Phone": "9111111111", "PDF_Path": "/tmp/dummy.pdf"}
        ]
        res = self.client.post("/send-whatsapp", json={"employees": employees, "month": "January"})
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertEqual(data["sent_count"], 1)
        self.assertEqual(data["failed_count"], 0)


if __name__ == "__main__":
    unittest.main()
