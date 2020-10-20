# Using the Microsoft Graph API

## Overview

Microsoft Graph is the gateway to data and intelligence in Microsoft 365. It provides
a unified programmability model that you can use to access the tremendous amount of data
in Microsoft 365, Windows 10, and Enterprise Mobility + Security. Use the wealth of data
n Microsoft Graph to build apps for organizations and consumers that interact with millions
of users.

## Register an App

### Step 1: Go to your Azure Portal

For step 1 you'll need to go to your [Azure Portal](https://portal.azure.com/) and login.

### Step 2: Go to the Azure Active Directory Service

Once you've logged into Azure you'll need to register a new application so you can access
the Microsoft Graph API. To register a new application go to your **Azure Active Directory**
and once there go down to **App Registrations** a new window will pop up.

### Step 3: Register a New App

Select **New registration**. On the Register an application page, set the values as follows.

- Set Name to **Python Graph Tutorial**.
- Set **Supported account** types to Accounts in any organizational directory and personal Microsoft accounts.
- Under **Redirect URI**, set the first drop-down to Web and set the value to <http://localhost:8000/callback>.

### Step 4: Grab your Client ID

Select **Register**. On the *Python Graph Tutorial* page, copy the value of the **Application (client) ID** and
save it, you will need it in the next step.

### Step 5: Create a Client Secret

Select **Certificates & secrets** under Manage. Select the **New client secret** button. Enter a value in
Description and select one of the options for Expires and select Add. **MAKE SURE TO COPY THE CLIENT SECRET**
**BEFORE YOU LEAVE THE PAGE OR ELSE YOU WILL NOT BE ABLE TO GRAB IT AGAIN AND WILL NEED TO CREATE A NEW ONE.**

### Step 6: Save the Content to a Config File (Optional)

If you want you can save the content to a config file using python and then when you need to grab your credentials
you can simply read the config file into your script. Some alternative options is using an Azure Key Vault or storing
the credentials in environment variables. To save the content in a config file use the sample file in the [configs folder](https://github.com/areed1192/ms-graph-python-client/tree/master/samples/configs)
that is found in the samples folder.

Make sure to fill out the values in the file:

```python
from configparser import ConfigParser

# Initialize the Parser.
config = ConfigParser()

# Add the Section.
config.add_section('graph_api')

# Set the Values.
config.set('graph_api', 'client_id', '')
config.set('graph_api', 'client_secret', '')
config.set('graph_api', 'redirect_uri', '')

# Write the file.
with open(file='samples/configs/config_video.ini', mode='w+') as f:
    config.write(f)
```
