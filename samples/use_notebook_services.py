from pprint import pprint
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

# SCOPES NEEDED:
# ---------------
# "Notes.Create
# "Notes.Read"
# "Notes.ReadWrite"
# "Notes.Read.All",
# "Notes.ReadWrite.All"

# Define the Scopes needed to Login.
scopes = [
    "Calendars.ReadWrite",
    "Files.ReadWrite.All",
    "User.ReadWrite.All",
    "Notes.ReadWrite.All",
    "Directory.ReadWrite.All",
    "User.Read.All",
    "Directory.Read.All",
    "Directory.ReadWrite.All",
    "Group.Read.All",
    "Group.ReadWrite.All",
    "Notes.Create",
    "Notes.Read",
    "Notes.ReadWrite",
    "Notes.Read.All",
    "Notes.ReadWrite.All",
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

# Grab the Notes Services.
notes_services = graph_client.notes()

# List all Notes.
pprint(notes_services.list_my_notebooks())

# Grab the Notebook Sections.
notebook_sections = notes_services.list_my_notebook_sections(
    notebook_id="0-8BC640C57CDA25B6!71451"
)

# List all the sections for the Notebook.
pprint(notebook_sections)

# Grab all the Notebook Pages.
notebook_pages = notes_services.list_my_notebook_pages(
    section_id="0-8BC640C57CDA25B6!71455"
)

# List all the Page for the Notebook for the particular session.
pprint(notebook_sections)
