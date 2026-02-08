import os
import boto3
from botocore.exceptions import ClientError
from dotenv import load_dotenv
import io

load_dotenv()

S3_BUCKET = os.getenv("S3_BUCKET")
AWS_REGION = os.getenv("AWS_REGION")

s3 = boto3.client(
    "s3",
    region_name=AWS_REGION,
    aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
    aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
)

def upload_to_s3(local_path, s3_key=None):
    """Upload file to S3"""
    if s3_key is None:
        s3_key = os.path.basename(local_path)
    s3.upload_file(local_path, S3_BUCKET, s3_key, ExtraArgs={"ContentType": "application/pdf"})
    return s3_key

def download_from_s3(s3_key, local_path):
    """Download file from S3"""
    s3.download_file(S3_BUCKET, s3_key, local_path)
    return local_path

def list_s3_pdfs():
    """List all PDF files in S3 bucket"""
    response = s3.list_objects_v2(Bucket=S3_BUCKET, Prefix="")
    if 'Contents' not in response:
        return []
    return [obj['Key'] for obj in response['Contents'] if obj['Key'].endswith('.pdf')]

def download_s3_file_to_memory(s3_key):
    """Download S3 file to memory"""
    file_obj = io.BytesIO()
    s3.download_fileobj(S3_BUCKET, s3_key, file_obj)
    file_obj.seek(0)
    return file_obj
