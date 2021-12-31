from pprint import pprint
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

# Needed Permissions
# Files.ReadWrite

scopes = [
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
config.read("config/config.ini")

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
    credentials="config/ms_graph_state.jsonc",
)

# Login to the Client.
graph_client.login()

# Grab the Workbooks Service.
workbooks_service = graph_client.workbooks()

# Create a new session for my Excel Workbook.
session_response = workbooks_service.create_session(
    item_path="Desktop/Personal Code/Repo - YouTube Channel Management/"
    + "youtube-channel-management/YouTube Video Description Database.xlsm"
)
pprint(session_response)

# List all the Tables in a Workbook.
table_objects = workbooks_service.list_tables(
    item_path="Desktop/Personal Code/Repo - YouTube Channel Management/"
    + "youtube-channel-management/YouTube Video Description Database.xlsm"
)
pprint(table_objects)

# List all the Worksheets in a Workbook.
worksheet_objects = workbooks_service.list_worksheets(
    item_path="Desktop/Personal Code/Repo - YouTube Channel Management/"
    + "youtube-channel-management/YouTube Video Description Database.xlsm"
)
pprint(worksheet_objects)

# List all the Named Objects in a Workbook.
name_objects = workbooks_service.list_names(
    item_path="Desktop/Personal Code/Repo - YouTube Channel Management/"
    + "youtube-channel-management/YouTube Video Description Database.xlsm"
)
pprint(name_objects)
