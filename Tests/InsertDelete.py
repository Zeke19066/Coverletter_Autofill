from __future__ import print_function

import json

from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools



## SET MODE:
# 0 = Insert
# 1 = Delete

mode = 0

# Set doc ID, as found at `https://docs.google.com/document/d/YOUR_DOC_ID/edit`
DOCUMENT_ID = '1ldhIwp5d7RAIPv0o79YKIwGvJFVHYMoUgA3UbJJkuKY'

# Set the scopes and discovery info
# Safe Mode: 'https://www.googleapis.com/auth/documents.readonly' 
SCOPES = 'https://www.googleapis.com/auth/documents'
DISCOVERY_DOC = ('https://docs.googleapis.com/$discovery/rest?'
                 'version=v1')

# Initialize credentials and instantiate Docs API service
store = file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = discovery.build('docs', 'v1', http=creds.authorize(
    Http()), discoveryServiceUrl=DISCOVERY_DOC)

## Insert
if mode == 0:
    requests = [
            {
            'insertText': {
                'location': {
                    'index': 225,
                },
                'text': "Example of a custom paragraph"
            }
        },
                    {
            'insertText': {
                'location': {
                    'index': 184,
                },
                'text': "Jobboard Posting"
            }
        },
                    {
            'insertText': {
                'location': {
                    'index': 111,
                },
                'text': "Recipient Bit Here"
            }
        },
    ]

    result = service.documents().batchUpdate(
        documentId=DOCUMENT_ID, body={'requests': requests}).execute()
    print("Completed Insert")

## Delete
elif mode == 1:
    requests = [
        {
            'deleteContentRange': {
                'range': {
                    'startIndex': 10,
                    'endIndex': 24,
                }

            }
        },
    ]
    result = service.documents().batchUpdate(
        documentId=DOCUMENT_ID, body={'requests': requests}).execute()
    print("Completed Delete")
