import pickle
import os
from google_auth_oauthlib.flow import InstalledAppFlow


# Specify permissions to send and read/write messages
# Find more information at:
# https://developers.google.com/gmail/api/auth/scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/gmail.modify']


# Get the user's home directory
home_dir = os.path.expanduser('~')

# Recall that the credentials.json data is saved in our "Downloads" folder
json_path = os.path.join(home_dir, 'Downloads', 'credentials.json')

flow = InstalledAppFlow.from_client_secrets_file(json_path, SCOPES)
creds = flow.run_local_server(port=0)

# We are going to store the credentials in the user's home directory
pickle_path = os.path.join(home_dir, 'credentials.pickle')
with open(pickle_path, 'wb') as token:
    pickle.dump(creds, token)