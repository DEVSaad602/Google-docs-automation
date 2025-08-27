from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os

SCOPES = ['https://www.googleapis.com/auth/documents']

def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_docs_api():
    creds = authenticate()

    from googleapiclient.discovery import build
    return build('docs', 'v1', credentials=creds)

def create_doc(service):
    document = service.documents().create(body={'title': 'My New Document'}).execute()
    document_id = document.get('documentId')
    print(f'Document created with ID: {document_id}')
    return document_id


def append_doc(service, document_id, content_h1=None, content_h2=None, content_text=''):
    doc = service.documents().get(documentId=document_id).execute()
    end_index = doc['body']['content'][-1]['endIndex'] if 'body' in doc and 'content' in doc['body'] and doc['body']['content'] else 1

    requests = []
    current_index = end_index - 1

    if content_h1:
        requests.append({
            'insertText': {
                'location': {'index': current_index},
                'text': content_h1 + '\n'
            }
        })
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': current_index,
                    'endIndex': current_index + len(content_h1) + 1
                },
                'paragraphStyle': {'namedStyleType': 'HEADING_1'},
                'fields': 'namedStyleType'
            }
        })
        current_index += len(content_h1) + 1

    if content_h2:
        requests.append({
            'insertText': {
                'location': {'index': current_index},
                'text': content_h2 + '\n'
            }
        })
        requests.append({
            'updateParagraphStyle': {
                'range': {
                    'startIndex': current_index,
                    'endIndex': current_index + len(content_h2) + 1
                },
                'paragraphStyle': {'namedStyleType': 'HEADING_2'},
                'fields': 'namedStyleType'
            }
        })
        current_index += len(content_h2) + 1

    if content_text:
        requests.append({
            'insertText': {
                'location': {'index': current_index},
                'text': content_text + '\n'
            }
        })

    if requests:
        service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
        print('Content appended successfully')


# Example usage:
service = get_docs_api()
document_id = create_doc(service)
append_doc(service, document_id, 'My Heading 1', 'My Heading 2', 'This is the appended content text.')

