from pprint import pprint
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

scopes = [
    "Contacts.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite.All",
    "User.ReadWrite.All",
    "Notes.ReadWrite.All",
    "Directory.ReadWrite.All",
    "User.Read.All",
    "Directory.Read.All",
    "Directory.ReadWrite.All",
    "Mail.ReadWrite",
    "Sites.ReadWrite.All",
    "ExternalItem.Read.All",
]

# Initialize the Parser.
config = ConfigParser()

# Read the file.
config.read("configs/config.ini")

# Get the specified credentials.
client_id = config.get("graph_api", "client_id")
client_secret = config.get("graph_api", "client_secret")
redirect_uri = config.get("graph_api", "redirect_uri")

# Initialize the Client.
graph_client = MicrosoftGraphClient(
    client_id=client_id,
    client_secret=client_secret,
    redirect_uri=redirect_uri,
    scope=scopes,
    credentials="configs/ms_graph_state.jsonc",
)

# Login to the Client.
graph_client.login()

# Define a valid User ID.
USER_ID = "8bc640c57cda25b6"

# Define a folder ID.
FOLDER_ID = (
    "AQMkADAwATZiZmYAZC1hMDI2LTE3NTgtMDACLTAwCgAuAAADpjqwNb_d"
    + "ak68rN7703uffQEAFNKsLOjbGUuHHmYnyKdJiAAFAP8ORwAAAA=="
)

# Grab the Personal Contacts Service.
personal_contacts_service = graph_client.personal_contacts()

# Grab my contacts folders.
pprint(personal_contacts_service.list_my_contacts_folder())

# Grab a contact folder for a specific user and a specific ID.
pprint(
    personal_contacts_service.list_contacts_folder_by_id(
        user_id=USER_ID, folder_id=FOLDER_ID
    )
)

# Grab a contact folder for a specific user and a specific ID.
pprint(
    personal_contacts_service.get_contacts_folder_by_id(
        user_id=USER_ID, folder_id=FOLDER_ID
    )
)

# Grab the Contacts.
my_contacts = personal_contacts_service.list_my_contacts()

# Get a random contact id.
contact_id = my_contacts["value"][-1]["id"]

# Grab a specific contact from my contacts folder.
pprint(personal_contacts_service.get_my_contact_by_id(contact_id=contact_id))

# Create a new contact folder under the default profile.
pprint(
    personal_contacts_service.create_my_contact_folder(
        folder_resource={
            "parentFolderId": "sigma-coding-contacts",
            "displayName": "Sigma Coding - Contacts",
        }
    )
)

# Create a new contact folder under the specified user profile.
pprint(
    personal_contacts_service.create_user_contact_folder(
        user_id=USER_ID,
        folder_resource={
            "parentFolderId": "trading-robot-contacts",
            "displayName": "Trading Robot - Contacts",
        },
    )
)
