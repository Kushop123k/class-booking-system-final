from google_auth_oauthlib.flow import InstalledAppFlow
import os

# Define the scopes your app needs
SCOPES = [
    "https://www.googleapis.com/auth/forms.body",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def generate_token():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)

    # Save the credentials to token.json
    with open("token.json", "w") as token:
        token.write(creds.to_json())

    print("✅ token.json created successfully!")

if __name__ == "__main__":
    generate_token()
