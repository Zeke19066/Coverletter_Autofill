from __future__ import print_function

import json

from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools

# Set doc ID, as found at `https://docs.google.com/document/d/YOUR_DOC_ID/edit`
DOCUMENT_ID = '1ldhIwp5d7RAIPv0o79YKIwGvJFVHYMoUgA3UbJJkuKY'

# Set the scopes and discovery info
SCOPES = 'https://www.googleapis.com/auth/documents.readonly' 
DISCOVERY_DOC = ('https://docs.googleapis.com/$discovery/rest?'
                 'version=v1')

# Initialize credentials and instantiate Docs API service
store = file.Storage('token_docs.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = discovery.build('docs', 'v1', http=creds.authorize(
    Http()), discoveryServiceUrl=DISCOVERY_DOC)

# Do a document "get" request and print the results as formatted JSON
result = service.documents().get(documentId=DOCUMENT_ID).execute()

##Save JSON
result_json = json.dumps(result)
path = r'doc_sample.json'

with open(path, 'w') as f:
    f.write(result_json)
    f.close()
            
print(json.dumps(result, indent=4, sort_keys=False)) #True