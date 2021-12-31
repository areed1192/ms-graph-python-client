from pprint import pprint
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

# SCOPES NEEDED:
# ---------------
# 'Mail.Send',
# 'MailboxSettings.Read',
# 'MailboxSettings.ReadWrite'

# Define the Scopes needed to Login.
scopes = [
    'Calendars.ReadWrite',
    'Files.ReadWrite.All',
    'User.ReadWrite.All',
    'Notes.ReadWrite.All',
    'Directory.ReadWrite.All',
    'User.Read.All',
    'Directory.Read.All',
    'Directory.ReadWrite.All',
    'Group.Read.All',
    'Group.ReadWrite.All',
    'Notes.Create',
    'Notes.Read',
    'Notes.ReadWrite',
    'Notes.Read.All',
    'Notes.ReadWrite.All',
    'Mail.Send',
    'MailboxSettings.Read',
    'MailboxSettings.ReadWrite'
]

# Initialize the Parser.
config = ConfigParser()

# Read the file.
config.read('configs/config.ini')

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
    credentials='configs/ms_graph_state.jsonc'
)

# Login to the Client.
graph_client.login()

# Define a valid User ID.
USER_ID = '8bc640c57cda25b6'

# Define a mail Item ID.
MAIL_ID = 'AQMkADAwATZiZmYAZC1hMDI2LTE3NTgtMDACLTAwCgBGAAADpjqwNb_dak68rN7703u' + \
          'ffQcAFNKsLOjbGUuHHmYnyKdJiAAAAgEhAAAAFNKsLOjbGUuHHmYnyKdJiAAFBMTneQ' + \
          'AAAA=='

# Define a mail item ID with Attachments.
MAIL_ID_ATTACHMENTS = 'AQMkADAwATZiZmYAZC1hMDI2LTE3NTgtMDACLTAwCgBGAAADpjqwNb+' + \
                      'dak68rN7703uffQcAFNKsLOjbGUuHHmYnyKdJiAAAAgEMAAAAFNKsLO' + \
                      'jbGUuHHmYnyKdJiAAE9ucV+AAAAA=='

# Grab the Notes Services.
mail_services = graph_client.mail()

# Grab all my Messages.
pprint(
    mail_services.list_my_messages()
)

# Grab a specific message for the default user.
pprint(
    mail_services.get_my_messages(
        message_id=MAIL_ID
    )
)

# Get a Specific User's Message.
pprint(
    mail_services.get_user_messages(
        user_id=USER_ID,
        message_id=MAIL_ID
    )
)

# List the rules for a specific user..
pprint(
    mail_services.list_rules(user_id=USER_ID)
)

# List the rules for the default user.
pprint(
    mail_services.list_my_rules()
)

# List the overrides for a specific user.
pprint(
    mail_services.list_overrides(user_id=USER_ID)
)

# List the overrides for the default user.
pprint(
    mail_services.list_my_overrides()
)

# List the attachments for a specific message.
pprint(
    mail_services.list_my_attachements(
        message_id=MAIL_ID_ATTACHMENTS
    )
)


# Create a new message for the default user. Keep in mind this does not send the mail.
new_message_draft = mail_services.create_my_message(
    message={
        "subject": "Did you see last night's game?",
        "importance": "Low",
        "body": {
            "contentType": "HTML",
            "content": "They were <b>awesome</b>!"
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": "alexreed1192@gmail.com"
                }
            }
        ]
    }
)

# Check it out.
pprint(new_message_draft)

# grab the ID.
new_message_id = new_message_draft['id']

# Send the newly created message.
mail_services.send_my_message(message_id=new_message_id)

# Let's create a new message rule, this will help with things like incoming mail. We can
# control what happens to mail that meets certain conditions.
my_new_message_rule = mail_services.create_my_message_rule(
    rule={
        "displayName": "From partner",
        "sequence": 2,
        "isEnabled": True,
        "conditions": {
            "senderContains": [
                "youtube"
            ]
        },
        "actions": {
            "forwardTo": [
                {
                    "emailAddress": {
                        "name": "Alex Reed",
                        "address": "coding.sigma@gmail.com"
                    }
                }
            ],
            "stopProcessingRules": True
        }
    }
)

# Check it out.
pprint(my_new_message_rule)
