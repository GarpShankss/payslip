import os
import re
import boto3
from botocore.config import Config
from dotenv import load_dotenv
import io
from datetime import datetime

load_dotenv()

S3_BUCKET = os.getenv("S3_BUCKET")

# R2 S3-compatible client
s3 = boto3.client(
    "s3",
    endpoint_url=os.getenv("AWS_ENDPOINT"),
    aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
    aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
    region_name="auto",
    config=Config(signature_version="s3v4")
)

# -------------------------------
# FILENAME UTILITIES
# -------------------------------

def sanitize_name(name: str) -> str:
    """
    Normalize a name for use in S3 keys:
    - Strip leading/trailing whitespace
    - Replace internal spaces/special chars with underscores
    - Remove anything that's not alphanumeric or underscore
    """
    name = name.strip()
    name = re.sub(r'[\s\-]+', '_', name)        # spaces/hyphens → underscore
    name = re.sub(r'[^\w]', '', name)            # strip non-alphanumeric
    name = re.sub(r'_+', '_', name)              # collapse multiple underscores
    return name


def build_pdf_filename(employee_name: str, unit_name: str) -> str:
    """
    Build a clean, deterministic PDF filename.
    e.g. 'John Doe' + 'Sales Unit' → 'John_Doe_Sales_Unit.pdf'
    """
    emp = sanitize_name(employee_name)
    unit = sanitize_name(unit_name)
    return f"{emp}_{unit}.pdf"


def build_s3_key(employee_name: str, unit_name: str, year: str, month: str) -> str:
    """
    Build the full S3 key with month-wise folder structure.
    e.g. '2026/March/John_Doe_Sales_Unit.pdf'
    """
    filename = build_pdf_filename(employee_name, unit_name)
    return f"{year}/{month}/{filename}"


# -------------------------------
# STORAGE UTILITIES
# -------------------------------

def get_bucket_size() -> int:
    """Get total bucket size in bytes."""
    total_size = 0
    paginator = s3.get_paginator('list_objects_v2')

    for page in paginator.paginate(Bucket=S3_BUCKET):
        if 'Contents' in page:
            for obj in page['Contents']:
                total_size += obj['Size']

    return total_size


def bytes_to_gb(bytes_size: int) -> float:
    """Convert bytes to GB."""
    return bytes_size / (1024 ** 3)


# -------------------------------
# CLEANUP LOGIC
# -------------------------------

def parse_month_key(month_key: str) -> datetime:
    """Convert '2026/March' → datetime for sorting."""
    year, month = month_key.split('/')
    return datetime.strptime(f"{year} {month}", "%Y %B")


def delete_oldest_month():
    """Delete the oldest month folder from the bucket."""
    paginator = s3.get_paginator('list_objects_v2')
    months = {}

    for page in paginator.paginate(Bucket=S3_BUCKET):
        if 'Contents' not in page:
            continue

        for obj in page['Contents']:
            key = obj['Key']
            parts = key.split('/')

            if len(parts) >= 2:
                month_key = f"{parts[0]}/{parts[1]}"
                months.setdefault(month_key, []).append(obj)

    if not months:
        print("No months found in bucket. Nothing to delete.")
        return

    sorted_months = sorted(months.keys(), key=parse_month_key)
    oldest_month = sorted_months[0]

    objects_to_delete = [{"Key": obj["Key"]} for obj in months[oldest_month]]

    print(f"Deleting oldest month: {oldest_month} ({len(objects_to_delete)} files)")

    # Batch delete (max 1000 per request — R2 supports this)
    for i in range(0, len(objects_to_delete), 1000):
        batch = objects_to_delete[i:i + 1000]
        s3.delete_objects(
            Bucket=S3_BUCKET,
            Delete={"Objects": batch, "Quiet": True}
        )


def ensure_storage_limit(limit_gb: float = 9.5):
    """
    Check current bucket size and delete oldest months
    until usage is safely below the limit.
    """
    size_bytes = get_bucket_size()
    size_gb = bytes_to_gb(size_bytes)

    print(f"Current storage: {size_gb:.3f} GB / {limit_gb} GB limit")

    while size_gb >= limit_gb:
        print(f"[WARN] Storage at {size_gb:.3f} GB - cleaning oldest month...")
        delete_oldest_month()

        size_bytes = get_bucket_size()
        size_gb = bytes_to_gb(size_bytes)
        print(f"Storage after cleanup: {size_gb:.3f} GB")


# -------------------------------
# UPLOAD
# -------------------------------

def upload_to_s3(
    local_path: str,
    employee_name: str,
    unit_name: str,
    month: str,
    year: str
) -> str:
    """
    Upload a PDF to R2 with a clean employee+unit filename
    under the YYYY/Month/ folder structure.

    Returns the final S3 key.
    """
    s3_key = build_s3_key(employee_name, unit_name, year, month)

    s3.upload_file(
        local_path,
        S3_BUCKET,
        s3_key,
        ExtraArgs={"ContentType": "application/pdf"}
    )

    print(f"Uploaded: {s3_key}")
    return s3_key


def upload_with_cleanup(
    local_path: str,
    employee_name: str,
    unit_name: str,
    month: str,
    year: str
) -> str:
    """
    Run storage cleanup check, then upload.
    Safe entry point for all uploads.
    """
    ensure_storage_limit()
    return upload_to_s3(local_path, employee_name, unit_name, month, year)


# -------------------------------
# DOWNLOAD
# -------------------------------

def download_from_s3(s3_key: str, local_path: str) -> str:
    """Download a file from R2 to local disk."""
    s3.download_file(S3_BUCKET, s3_key, local_path)
    return local_path


def download_s3_file_to_memory(s3_key: str) -> io.BytesIO:
    """Download a file from R2 directly into memory."""
    file_obj = io.BytesIO()
    s3.download_fileobj(S3_BUCKET, s3_key, file_obj)
    file_obj.seek(0)
    return file_obj


# -------------------------------
# LIST FILES
# -------------------------------

def list_s3_pdfs(month: str = None, year: str = None) -> list[str]:
    """
    List PDF files in the bucket, optionally filtered by year/month.
    """
    if year and month:
        prefix = f"{year}/{month}/"
    elif year:
        prefix = f"{year}/"
    elif month:
        prefix = f"{month}/"
    else:
        prefix = ""

    paginator = s3.get_paginator('list_objects_v2')
    files = []

    for page in paginator.paginate(Bucket=S3_BUCKET, Prefix=prefix):
        if 'Contents' not in page:
            continue
        for obj in page['Contents']:
            if obj['Key'].endswith('.pdf'):
                files.append(obj['Key'])

    return files


# -------------------------------
# GENERATE SIGNED URL
# -------------------------------

def generate_presigned_url(s3_key: str, expiry: int = 3600) -> str:
    """Generate a secure, temporary download URL for a given S3 key."""
    return s3.generate_presigned_url(
        'get_object',
        Params={
            'Bucket': S3_BUCKET,
            'Key': s3_key
        },
        ExpiresIn=expiry
    )