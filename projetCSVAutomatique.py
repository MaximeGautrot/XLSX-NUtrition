import os
import pickle
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Définissez la portée pour l'API Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def authenticate_and_upload(file_path, file_name, file_log, file_token, folder_id):
    creds = None
    if os.path.exists(file_token):
        with open(file_token, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                file_log, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(file_token, 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    file_metadata = {'name': file_name}
    if folder_id != "":
        file_metadata['parents'] = [folder_id]

    media = MediaFileUpload(file_path, mimetype='application/vnd.ms-excel')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    print(f"ID du fichier: {file.get('id')}")
