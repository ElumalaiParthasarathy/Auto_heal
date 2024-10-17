import pymongo
import psycopg2
import requests
import json
import pandas as pd
import os
from datetime import datetime
import win32com.client as win32

# MongoDB connection strings
lead_collection_connection_string = "mongodb+srv://prf_intengg_mongo_prod_common_read_user:prod000read000bbnUXbu000mongo@piramal-sourcing-digital-prod-cluster-pl-0.y0ci9.mongodb.net/"
pan_collection_connection_string = "mongodb+srv://platform_common_read_user:platformreadTXykVbNGg7Qa8CEV@piramal-platform-prod-cluster-pl-0.y0ci9.mongodb.net/"

# PostgreSQL connection settings
postgres_host = "postgres-slave.piramalfinance.com"
postgres_port = 5432
postgres_user = "prod_postgres_read_user"
postgres_password = "readX3MlGOq7uvJcBUAHXF"
postgres_db = "auth_service_db"

# Connect to MongoDB
lead_client = pymongo.MongoClient(lead_collection_connection_string)
pan_client = pymongo.MongoClient(pan_collection_connection_string)

lead_db = lead_client['lead_db']
lead_collection = lead_db['leadDocument']

# Connect to PostgreSQL
postgres_conn = psycopg2.connect(
    host=postgres_host,
    port=postgres_port,
    user=postgres_user,
    password=postgres_password,
    database=postgres_db
)

# Define the Lead IDs and Customer IDs
data = [
    {"lead_id": "BLSA00047799", "customer_id": "f7495e52-d011-4004-815d-84fd3af934f8"},
    {"lead_id": "BLSA00047799", "customer_id": "f7495e52-d011-4004-815d-84fd3af934f8"},
    {"lead_id": "BLSA00047799", "customer_id": "f7495e52-d011-4004-815d-84fd3af934f8"},
    
    {"lead_id": "BLSA00048AC4", "customer_id": "ab9731c4-ef17-4f84-b780-c9b11f84a240"},
    {"lead_id": "BLSA00048AC4", "customer_id": "ab9731c4-ef17-4f84-b780-c9b11f84a240"},
    
    {"lead_id": "BLSA000497B8", "customer_id": "e6cd158a-b7e0-43e3-a447-cfe45a58c8f8"},
    
    {"lead_id": "BLSA000475F4", "customer_id": "f6d4e28b-3358-4731-b172-03bfa3c4ccc8"},
    {"lead_id": "BLSA000475F4", "customer_id": "f6d4e28b-3358-4731-b172-03bfa3c4ccc8"},
    
    {"lead_id": "HLSA000B292A", "customer_id": "b4b09cc1-f7bc-494f-b75b-369050dfa1ca"},
    
    {"lead_id": "SCUBL0076993", "customer_id": "075a1c75-4a55-4cd8-9cd7-0844f7907fa3"},
    
    {"lead_id": "HLSA000ACEF6", "customer_id": "489f7f62-08bb-4246-9026-6ac8dc7509cb"},
    
    {"lead_id": "BLSA0004554F", "customer_id": "29f2f9ee-1901-4005-b450-568b1350c340"},
    
    {"lead_id": "HLSA000B6171", "customer_id": "b6fed958-99ba-477b-b1f7-f924ac5925b6"},
    
    {"lead_id": "BLSA000496E5", "customer_id": "d9c16139-18c4-4660-ae0b-5be191a471a9"},
    
    {"lead_id": "HLSA000B6171", "customer_id": "b6fed958-99ba-477b-b1f7-f924ac5925b6"},
    
    {"lead_id": "BLSA000499D6", "customer_id": "1c5d8875-e01e-4072-80a3-d138a99ad52e"},
    
    {"lead_id": "HLSA000B1573", "customer_id": "57dd99f7-9fa0-461e-8445-80c4c64fe486"},
    
    {"lead_id": "BLSA000498F5", "customer_id": "6b3694f0-8939-4e99-81ed-ab0ea0836c7a"},
    
    {"lead_id": "HLSA000B567C", "customer_id": "010b0349-27f1-418e-b64d-1525bd7f0ab2"},
    
    {"lead_id": "MLAP00000C85", "customer_id": "d30c129e-a71d-440b-94ce-729fca46bf8c"},
    
    {"lead_id": "BLSA000499F6", "customer_id": "2cb901ba-8d47-47d2-9e4b-4ac47587d86a"},
]


results = []

# Function to create the directory if it doesn't exist
def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

# Function to get email from createdBy ID
def get_email(created_by_id):
    with postgres_conn.cursor() as cursor:
        cursor.execute("SELECT email FROM auth_schema.admin WHERE id=%s", (created_by_id,))
        result = cursor.fetchone()
        return result[0] if result else None

# Function to call KYC API
def check_pan_and_dob(pan):
    url = 'https://api.pchf.in/api/v1/kyc/pan/profile'
    headers = {
        'Content-Type': 'application/json',
        'x-code': '48cN4cvgPQrBkUW4WkmJ',
        'x-apikey': 'qK5CRIOVA1bH3LzlCDsb25OGApFMUwAS',
        'Authorization': 'Bearer gFa7MFrBd_kZeGf_JwdGB8ZZNurWxxq8QI-vjqIaraMdUunVsBcQDjLD4a5ywPzp4Lr0ks8AdoIuhGML11vMdw',
    }
    data = json.dumps({"pan": pan})
    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code == 200:
        response_data = response.json()
        actual_dob = response_data.get("result", {}).get("dob")
        return actual_dob
        
    return "no data"

# Function to send email
def send_email(to_email, subject, body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    mail.CC = "elumalai.p2@piramal.com; keerthana.m@piramal.com; vinod.kumar17@piramal.com"  # Add additional CC emails here
    try:
        mail.Send()
        print(f"Email sent to {to_email} with CC.")
    except Exception as e:
        print(f"Error sending email: {e}")

# Loop through each lead
for item in data:
    lead_id = item["lead_id"]
    
    # Fetch the createdBy ID from MongoDB
    lead_doc = lead_collection.find_one({"_id": lead_id}, {"createdBy": 1})
    
    if lead_doc:
        created_by_id = lead_doc.get("createdBy")
        email = get_email(created_by_id)
     
        # Fetch PAN card number from the KYC database
        pan_doc = pan_client['kyc_db']['Pan_Authentication_Data'].find_one({"referenceId": item["customer_id"]})
        
        if pan_doc and "panAuthenticationRequest" in pan_doc and "pan" in pan_doc["panAuthenticationRequest"]:
            pan_card_number = pan_doc["panAuthenticationRequest"]["pan"]
            entered_dob = pan_doc["panAuthenticationRequest"]["dob"]
            
            # Check with the KYC API
            actual_dob = check_pan_and_dob(pan_card_number)
            
            if actual_dob and actual_dob != entered_dob:
                results.append({
                    "lead_id": lead_id,
                    "created_by": created_by_id,
                    "customer_id": item["customer_id"],
                    "entered_dob": entered_dob,
                    "actual_dob": actual_dob,
                    "pan_card_number": pan_card_number,
                    "email": email
                })
                print(f"Discrepancy found for Lead ID: {lead_id} - Entered DOB: {entered_dob}, Actual DOB: {actual_dob}")

                # Send email notification
                subject = f"Discrepancy Alert for Lead ID: {lead_id}"
                body = (f"Dear User,\n\n"
                        f"A discrepancy has been found for Lead ID: {lead_id}.\n"
                        f"Entered DOB: {entered_dob}\n"
                        f"Actual DOB: {actual_dob}\n"
                        f"PAN Card Number: {pan_card_number}\n\n"
                        f"Please check the details and take necessary actions.\n\n"
                        f"Regards,\nYour System")
                send_email(email, subject, body)

# Check if results are empty
if not results:
    print("No discrepancies found.")

# Define the path for the output file
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
output_directory = os.path.join(desktop_path, "pan auto heal")
create_directory(output_directory)

# Generate a timestamp for the filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file_path = os.path.join(output_directory, f"dob_discrepancies_{timestamp}.xlsx")

# Write results to an Excel file
if results:
    df = pd.DataFrame(results)
    df.to_excel(output_file_path, index=False)
    print(f"Process completed. Discrepancies written to {output_file_path}.")
else:
    print("No discrepancies to write to the Excel file.")
