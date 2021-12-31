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

# Grab the Range Service.
range_service = graph_client.range()

# Grab a range.
range_object = range_service.get_range(
    item_path="Desktop/Personal Code/Repo - YouTube Channel Management/"
    + "youtube-channel-management/YouTube Video Description Database.xlsm",
    worksheet_name_or_id="Video_Database",
    address="A1:P374",
)
pprint(range_object)
