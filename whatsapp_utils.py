# whatsapp_utils.py
import os
import re
import requests
from dotenv import load_dotenv

load_dotenv()

def get_whatsapp_config():
    phone_id = os.getenv("WA_PHONE_NUMBER_ID", "").strip()
    token = os.getenv("WA_ACCESS_TOKEN", "").strip()
    return phone_id, token


# -------------------------------
# STEP 1: Upload PDF to WhatsApp Media
# -------------------------------
def upload_pdf_to_whatsapp(pdf_bytes: bytes, filename: str) -> tuple[str | None, str]:
    """
    Upload a PDF to WhatsApp's media endpoint.
    Returns (media_id, error_message).
    """
    phone_id, token = get_whatsapp_config()
    if not phone_id or not token:
        return None, "WhatsApp Phone Number ID or Access Token is missing in environment variables"

    upload_url = f"https://graph.facebook.com/v19.0/{phone_id}/media"

    try:
        response = requests.post(
            upload_url,
            headers={"Authorization": f"Bearer {token}"},
            files={
                "file": (filename, pdf_bytes, "application/pdf"),
                "messaging_product": (None, "whatsapp"),
                "type": (None, "application/pdf"),
            },
            timeout=30
        )
        data = response.json()
        if response.status_code == 200 and data.get("id"):
            media_id = data["id"]
            print(f"✓ PDF uploaded to WhatsApp, media_id: {media_id}")
            return media_id, ""
        else:
            err = data.get("error", {})
            err_msg = err.get("message", response.text)
            print(f"✗ WhatsApp media upload failed: {err_msg}")
            return None, f"Media upload failed: {err_msg}"

    except Exception as e:
        print(f"✗ WhatsApp media upload failed: {e}")
        return None, f"Media upload exception: {str(e)}"


# -------------------------------
# STEP 2: Send PDF via Template Message
# -------------------------------
def send_payslip_whatsapp(
    phone_number: str,
    emp_name: str,
    month: str,
    pdf_bytes: bytes,
    pdf_filename: str
) -> tuple[bool, str]:
    """
    Send payslip PDF to employee via WhatsApp.
    Uses a pre-approved template with document header.
    
    phone_number can be 10 digits (will auto-add 91) or international format.
    Returns (success: bool, message: str)
    """
    phone_id, token = get_whatsapp_config()
    if not phone_id or not token:
        return False, "WA_PHONE_NUMBER_ID or WA_ACCESS_TOKEN not configured"

    # Sanitize phone — remove +, spaces, dashes, dots, etc.
    phone = re.sub(r'[^\d]', '', str(phone_number).strip())

    if not phone:
        return False, "Phone number is empty"

    # Auto-format 10-digit Indian numbers (or 11-digit with leading 0)
    if len(phone) == 10:
        phone = "91" + phone
    elif len(phone) == 11 and phone.startswith("0"):
        phone = "91" + phone[1:]

    if len(phone) < 10:
        print(f"✗ Invalid phone number: {phone_number}")
        return False, f"Invalid phone number length: {phone_number}"

    # Upload PDF first to get media_id
    media_id, upload_err = upload_pdf_to_whatsapp(pdf_bytes, pdf_filename)
    if not media_id:
        return False, upload_err or "Failed to upload PDF to WhatsApp media server"

    # Send template message with PDF as document header
    payload = {
        "messaging_product": "whatsapp",
        "to": phone,
        "type": "template",
        "template": {
            "name": "payslip_notification",   # must match your approved template name
            "language": {"code": "en"},
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {
                            "type": "document",
                            "document": {
                                "id": media_id,
                                "filename": pdf_filename
                            }
                        }
                    ]
                },
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "text": emp_name},   # {{1}}
                        {"type": "text", "text": month}        # {{2}}
                    ]
                }
            ]
        }
    }

    try:
        api_url = f"https://graph.facebook.com/v19.0/{phone_id}/messages"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        response = requests.post(api_url, headers=headers, json=payload, timeout=30)
        data = response.json()

        print(f"  Send status code: {response.status_code}")
        print(f"  Send response: {data}")
        if response.status_code == 200:
            msg_id = data.get("messages", [{}])[0].get("id")
            print(f"✓ WhatsApp sent to {phone} ({emp_name}) — message ID: {msg_id}")
            return True, "Message sent successfully"
        else:
            error = data.get("error", {})
            err_code = error.get('code', 'unknown')
            err_msg = error.get('message', response.text)
            print(f"✗ WhatsApp failed: code={err_code} message={err_msg}")
            return False, f"Meta API Error ({err_code}): {err_msg}"

    except Exception as e:
        print(f"✗ WhatsApp failed for {phone} ({emp_name}): {e}")
        return False, f"Request exception: {str(e)}"
