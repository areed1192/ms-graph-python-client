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

# Grab the Drive Items Service.
drive_item_services = graph_client.drive_item()

# Define a valid User ID.
user_id = '8bc640c57cda25b6'

# Define a valid Drive ID.
drive_id = '8bc640c57cda25b6'

# Define a valid Drive Item ID.
drive_item_id = '8BC640C57CDA25B6!3837'

# Grab a Drive Item, by ID.
pprint(
    drive_item_services.get_drive_item(
        drive_id=drive_id,
        item_id=drive_item_id
    )
)

# Grab a Drive Item, by path.
pprint(
    drive_item_services.get_drive_item_by_path(
        drive_id=drive_id,
        item_path='/Career - Certifications & Exams'
    )
)

# Grab a Drive Item, for a specific user in a specific Drive.
pprint(
    drive_item_services.get_user_drive_item(
        user_id=user_id,
        item_id=drive_item_id
    )
)

# Grab a Drive Item, by path for a specific user in a specific Drive.
pprint(
    drive_item_services.get_user_drive_item_by_path(
        user_id=user_id,
        item_path='/Career - Certifications & Exams'
    )
)

# Grab my Drive Item by ID.
pprint(
    drive_item_services.get_my_drive_item(
        item_id=drive_item_id
    )
)

# Grab my Drive Item, by path.
pprint(
    drive_item_services.get_my_drive_item_by_path(
        item_path='/Career - Certifications & Exams'
    )
)

# Define a valid Group ID.
group_id = 'GROUP_ID_GOES_HERE'

# Grab a group Drive Item by ID.
pprint(
    drive_item_services.get_group_drive_item(
        group_id=group_id,
        item_id=drive_item_id
    )
)

# Grab a group Drive Item, by path.
pprint(
    drive_item_services.get_group_drive_item_by_path(
        group_id=group_id,
        item_path='/Career - Certifications & Exams'
    )
)
