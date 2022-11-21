from pprint import pprint
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

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
    "ExternalItem.Read.All"
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
    credentials="configs/ms_graph_state.jsonc"
)

# Login to the Client.
graph_client.login()

# Grab the Search Service.
search_service = graph_client.search()

# Search for some documents.
search_response = search_service.query(
    search_request={
        "requests": [
            {
                "entityTypes": [
                    "message"
                ],
                "query": {
                    "queryString": "sigma"
                },
                "from": 0,
                "size": 25
            }
        ]
    }
)

# Print the Output.
pprint(search_response)
