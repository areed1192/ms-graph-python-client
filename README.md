# Microsoft Graph Python Client Library

## Table of Contents

- [Overview](#overview)
- [Setup](#setup)
- [Usage](#usage)
- [Support These Projects](#support-these-projects)

## Overview

Microsoft Graph is the gateway to data and intelligence in Microsoft 365. It provides
a unified programmability model that you can use to access the tremendous amount of data
in Microsoft 365, Windows 10, and Enterprise Mobility + Security. This project utilizes python
to help users interact with and manage data on Microsoft Graph API.

## Setup

**Setup - Requirements Install:**

For this particular project, you only need to install the dependencies, to use the project. The dependencies
are listed in the `requirements.txt` file and can be installed by running the following command:

```console
pip install -r requirements.txt
```

After running that command, the dependencies should be installed.

**Setup - Local Install:**

If you are planning to make modifications to this project or you would like to access it
before it has been indexed on `PyPi`. I would recommend you either install this project
in `editable` mode or do a `local install`. For those of you, who want to make modifications
to this project. I would recommend you install the library in `editable` mode.

If you want to install the library in `editable` mode, make sure to run the `setup.py`
file, so you can install any dependencies you may need. To run the `setup.py` file,
run the following command in your terminal.

```console
pip install -e .
```

If you don't plan to make any modifications to the project but still want to use it across
your different projects, then do a local install.

```console
pip install .
```

This will install all the dependencies listed in the `setup.py` file. Once done
you can use the library wherever you want.

**Setup - PyPi Install:**

To **install** the library, run the following command from the terminal.

```console
pip install
```

**Setup - PyPi Upgrade:**

To **upgrade** the library, run the following command from the terminal.

```console
pip install --upgrade
```

## Usage

Here is a simple example of using the `ms_graph` library.

```python
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
    'Directory.ReadWrite.All',
    'offline_access',
    'openid',
    'profile'
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


# Grab the User Services.
user_services = graph_client.users()

# List the Users.
pprint(user_services.list_users())


# Grab the Drive Services.
drive_services = graph_client.drives()

# List the Root Drive.
pprint(drive_services.get_root_drive())
```

## Support These Projects

**Patreon:**
Help support this project and future projects by donating to my [Patreon Page](https://www.patreon.com/sigmacoding). I'm
always looking to add more content for individuals like yourself, unfortuantely some of the APIs I would require me to
pay monthly fees.

**YouTube:**
If you'd like to watch more of my content, feel free to visit my YouTube channel [Sigma Coding](https://www.youtube.com/c/SigmaCoding).

<!-- **Hire Me:**
If you have a project, you think I can help you with feel free to reach out at [coding.sigma@gmail.com](mailto:coding.sigma@gmail.com?subject=[GitHub]%20Project%20Proposal) or fill out the [contract request form](https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAa__aAmF1hURFg5ODdaVTg1TldFVUhDVjJHWlRWRzhZRy4u) -->
