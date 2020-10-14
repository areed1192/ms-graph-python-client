import json
import time
import pprint
import urllib
import requests
import random
import string
import pathlib

from typing import List
from typing import Dict
from typing import Union

from ms_graph.users import Users
from ms_graph.session import GraphSession

from urllib.parse import urlencode, urlparse, quote_plus


class MicrosoftGraphClient():

    RESOURCE = 'https://graph.microsoft.com/'

    AUTHORITY_URL = 'https://login.microsoftonline.com/'
    AUTH_ENDPOINT = '/oauth2/v2.0/authorize?'
    TOKEN_ENDPOINT = '/oauth2/v2.0/token'

    OFFICE365_AUTHORITY_URL = 'https://login.live.com'
    OFFICE365_AUTH_ENDPOINT = '/oauth20_authorize.srf?'
    OFFICE365_TOKEN_ENDPOINT = '/oauth20_token.srf'

    def __init__(self, client_id: str, client_secret: str, redirect_uri: str, scope: List[str],
                 account_type: str = 'consumers', office365: bool = False, credentials: str = None):
        """Initializes the Graph Client.

        ### Parameters
        ----
        client_id : str
            The application Client ID assigned when creating a new Microsoft App.

        client_secret : str
            The application Client Secret assigned when creating a new Microsoft App.

        redirect_uri : str
            The application Redirect URI assigned when creating a new Microsoft App.

        scope : List[str]
            The list of scopes you want the application to have access to.

        account_type : str, optional
            [description], by default 'common'

        office365 : bool, optional
            [description], by default False
        """

        # printing lowercase
        letters = string.ascii_lowercase

        self.credentials = credentials
        self.token_dict = None

        self.client_id = client_id
        self.client_secret = client_secret
        self.api_version = 'v1.0'
        self.account_type = account_type
        self.redirect_uri = redirect_uri

        self.scope = scope
        self.state = ''.join(random.choice(letters) for i in range(10))

        self.access_token = None
        self.refresh_token = None
        self.graph_session = None

        self.base_url = self.RESOURCE + self.api_version + '/'
        self.office_url = self.OFFICE365_AUTHORITY_URL + self.OFFICE365_AUTH_ENDPOINT
        self.graph_url = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT
        self.office365 = office365

    def _state(self, action: str, token_dict: dict = None) -> bool:
        """Sets the session state for the Client Library.

        ### Arguments
        ----
        action : str
            Defines what action to take when determining the state. Either
            `load` or `save`.

        token_dict : dict, optional
            If the state is defined as `save` then pass through the
            token dictionary you want to save, by default None.

        ### Returns
        ----
        bool:
            If the state action was successful, then returns `True`
            otherwise it returns `False`.
        """

        # Determine if the Credentials file exists.
        does_exists = pathlib.Path(self.credentials).exists()

        # If it exists and we are loading it then proceed.
        if does_exists and action == 'load':

            # Load the file.
            with open(file=self.credentials, mode='r') as state_file:
                credentials = json.load(fp=state_file)

            # Grab the Token if it exists.
            if 'refresh_token' in credentials:

                self.refresh_token = credentials['refresh_token']
                self.access_token = credentials['access_token']
                self.token_dict = credentials

                return True

            else:
                return False

        # If we are saving the state then open the file and dump the dictionary.
        elif action == 'save':

            token_dict['expires_in'] = time.time() + int(token_dict['expires_in'])
            token_dict['ext_expires_in'] = time.time() + int(token_dict['ext_expires_in'])

            self.token_dict = token_dict

            with open(file=self.credentials, mode='w+') as state_file:
                json.dump(obj=token_dict, fp=state_file, indent=2)

    def _token_seconds(self, token_type: str = 'access_token') -> int:
        """Determines time till expiration for a token.

        Return the number of seconds until the current access token or refresh token
        will expire. The default value is access token because this is the most commonly used
        token during requests.

        ### Arguments:
        ----
        token_type {str} --  The type of token you would like to determine lifespan for. 
            Possible values are ['access_token', 'refresh_token'] (default: {access_token})

        ### Returns:
        ----
        {int} -- The number of seconds till expiration.
        """

        # if needed check the access token.
        if token_type == 'access_token':

            # if the time to expiration is less than or equal to 0, return 0.
            if not self.access_token or (time.time() + 60 >= self.token_dict['expires_in']):
                return 0

            # else return the number of seconds until expiration.
            token_exp = int(self.token_dict['expires_in'] - time.time() - 60)

        # if needed check the refresh token.
        elif token_type == 'refresh_token':

            # if the time to expiration is less than or equal to 0, return 0.
            if not self.refresh_token or (time.time() + 60 >= self.token_dict['ext_expires_in']):
                return 0

            # else return the number of seconds until expiration.
            token_exp = int(
                self.token_dict['ext_expires_in'] - time.time() - 60
            )

        return token_exp

    def _token_validation(self, nseconds: int = 60):
        """Checks if a token is valid.

        Verify the current access token is valid for at least N seconds, and
        if not then attempt to refresh it. Can be used to assure a valid token
        before making a call to the TD Ameritrade API.

        Arguments:
        ----
        nseconds {int} -- The minimum number of seconds the token has to be 
            valid for before attempting to get a refresh token. (default: {5})
        """

        if self._token_seconds(token_type='access_token') < nseconds:
            self.grab_refresh_token()

    def _silent_sso(self) -> bool:
        """Attempts a Silent Authentication using the Access Token and Refresh Token.

        Returns
        -------
        (bool)
            `True` if it was successful and `False` if it failed.
        """        

        # if the current access token is not expired then we are still authenticated.
        if self._token_seconds(token_type='access_token') > 0:
            return True

        # if the refresh token is expired then you have to do a full login.
        elif self._token_seconds(token_type='refresh_token') <= 0:
            return False

        # if the current access token is expired then try and refresh access token.
        elif self.refresh_token and self.grab_refresh_token():
            return True

        # More than likely a first time login, so can't do silent authenticaiton.
        return False

    def login(self) -> None:
        """Logs the user into the session."""        

        # Load the State.
        self._state(action='load')

        # Try a Silent SSO First.
        if self._silent_sso():

            # Set the Session.
            self.graph_session = GraphSession(client=self)
            return True
            
        else:

            # Build the URL.
            url = self.authorization_url()

            # aks the user to go to the URL provided, they will be prompted to authenticate themsevles.
            print('Please go to URL provided authorize your account: {}'.format(url))

            # ask the user to take the final URL after authentication and paste here so we can parse.
            my_response = input('Paste the full URL redirect here: ')

            # store the redirect URL
            self._redirect_code = my_response

            # this will complete the final part of the authentication process.
            self.grab_access_token()

            # Set the session.
            self.graph_session = GraphSession(client=self)

    def authorization_url(self):
        """Builds the authorization URL used to get an Authorization Code.

        ### Returns:
        ----
        A string.
        """

        params = {
            'client_id': self.client_id,
            'redirect_uri': self.redirect_uri,
            'scope': ' '.join(self.scope),
            'response_type': 'code',
            'response_mode': 'query',
            'state': self.state
        }

        if self.office365:
            response = self.office_url + urlencode(params).replace("+","%20")
        else:
            response = self.graph_url + urlencode(params).replace("+","%20")

        return response

    def grab_access_token(self) -> Dict:
        """Exchanges a code for an Access Token.

        ### Returns:
        ----
        Dict: A dictionary containing a new access token and refresh token.
        """

        # Parse the Code.
        query_dict = urllib.parse.parse_qs(self._redirect_code)

        # Grab the Code.
        code = query_dict[self.redirect_uri + "?code"]

        # Define the Arguments.
        data = {
            'client_id': self.client_id,
            'redirect_uri': self.redirect_uri,
            'client_secret': self.client_secret,
            'code': code,
            'grant_type': 'authorization_code',
        }

        # If we are doing a 365 request then change the endpoint.
        if self.office365:

            response = requests.post(
                self.OFFICE365_AUTHORITY_URL + self.OFFICE365_TOKEN_ENDPOINT,
                data=data
            )

        else:

            response = requests.post(
                self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT,
                data=data
            )

        token_dict = response.json()

        # Save the token dict.
        self._state(
            action='save',
            token_dict=token_dict
        )

        return token_dict

    def grab_refresh_token(self) -> Dict:
        """Grabs a new access token using a refresh token.

        ### Returns
        ----
        Dict:
            A token dictionary with a new access token.
        """

        data = {
            'client_id': self.client_id,
            'redirect_uri': self.redirect_uri,
            'client_secret': self.client_secret,
            'refresh_token': self.refresh_token,
            'grant_type': 'refresh_token',
        }

        if self.office365:

            response = requests.post(
                self.OFFICE365_AUTHORITY_URL + self.OFFICE365_TOKEN_ENDPOINT,
                data=data
            )

        else:

            response = requests.post(
                self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT,
                data=data
            )

        token_dict = response.json()

        self._state(
            action='save',
            token_dict=token_dict
        )

        return token_dict

    def users(self) -> Users:

        user_object: Users = Users(session=self.graph_session)

        return user_object
