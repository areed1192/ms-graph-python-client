from pprint import pprint
from ms_graph.client import MicrosoftGraphClient
from configparser import ConfigParser

scopes = [
    'Calendars.ReadWrite',
    'Files.ReadWrite.All',
    'User.ReadWrite.All',
    'Notes.ReadWrite.All',
    'Directory.ReadWrite.All',
    'User.Read.All',
    'Directory.Read.All',
    'Directory.ReadWrite.All'
]

# Initialize the Parser.
config = ConfigParser()

# Read the file.
config.read('config/config.ini')

# Get the specified credentials.
client_id = config.get('graph_api', 'client_id')
client_secret = config.get('graph_api', 'client_secret')
redirect_uri = config.get('graph_api', 'redirect_uri')

# Initialize the Client.
graph_client = MicrosoftGraphClient(
    client_id=client_id,
    client_secret=client_secret,
    redirect_uri=redirect_uri,
    scope=scopes,
    credentials='config/ms_graph_state.jsonc'
)

# Login to the Client.
graph_client.login()

# Grab the Drive Services.
drive_services = graph_client.drives()

# List the Root Drive.
pprint(drive_services.get_root_drive())

# List the Root Drive Deltas.
pprint(drive_services.get_root_drive_delta())

# List the Root Drive Children.
pprint(drive_services.get_root_drive_children())

# List the Root Drive Followers
pprint(drive_services.get_root_drive_followed())

# Grab a Drive by id.
pprint(drive_services.get_drive_by_id(drive_id='8bc640c57cda25b6'))

# Grab MY Drives.
pprint(drive_services.get_my_drives())

# Grab User Drives.
pprint(drive_services.get_user_drives(user_id='8bc640c57cda25b6'))

# Grab Group Drives.
pprint(drive_services.get_user_drives(user_id='8bc640c57cda25b6'))