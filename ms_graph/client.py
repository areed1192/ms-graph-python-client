import json
import time
import urllib
import random
import string
import pathlib

from typing import List
from typing import Dict

import msal

from ms_graph.users import Users
from ms_graph.drives import Drives
from ms_graph.groups import Groups
from ms_graph.notes import Notes
from ms_graph.session import GraphSession
from ms_graph.drive_items import DriveItems
from ms_graph.search import Search
from ms_graph.personal_contacts import PersonalContacts
from ms_graph.mail import Mail

from ms_graph.workbooks_and_charts.workbook import Workbooks
from ms_graph.workbooks_and_charts.range import Range

class MicrosoftGraphClient:

    """
    ### Overview:
    ----
    Used as the main entry point for the Microsoft Graph
    API Service.
    """

    RESOURCE = "https://graph.microsoft.com/"

    AUTHORITY_URL = "https://login.microsoftonline.com/"
    AUTH_ENDPOINT = "/oauth2/v2.0/authorize?"
    TOKEN_ENDPOINT = "/oauth2/v2.0/token"

    OFFICE365_AUTHORITY_URL = "https://login.live.com"
    OFFICE365_AUTH_ENDPOINT = "/oauth20_authorize.srf?"
    OFFICE365_TOKEN_ENDPOINT = "/oauth20_token.srf"

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        scope: List[str],
        account_type: str = "consumers",
        office365: bool = False,
        credentials: str = None,
    ):
        """Initializes the Graph Client.

        ### Parameters
        ----
        client_id : str
            The application Client ID assigned when
            creating a new Microsoft App.

        client_secret : str
            The application Client Secret assigned when
            creating a new Microsoft App.

        redirect_uri : str
            The application Redirect URI assigned when
            creating a new Microsoft App.

        scope : List[str]
            The list of scopes you want the application
            to have access to.

        account_type : str, optional
            [description], by default "common"

        office365 : bool, optional
            [description], by default False
        """

        # printing lowercase
        letters = string.ascii_lowercase

        self.credentials = credentials
        self.token_dict = None

        self.client_id = client_id
        self.client_secret = client_secret
        self.api_version = "v1.0"
        self.account_type = account_type
        self.redirect_uri = redirect_uri

        self.scope = scope
        self.state = "".join(random.choice(letters) for i in range(10))

        self.access_token = None
        self.refresh_token = None
        self.graph_session = None
        self.id_token = None

        self.base_url = self.RESOURCE + self.api_version + "/"
        self.office_url = self.OFFICE365_AUTHORITY_URL + self.OFFICE365_AUTH_ENDPOINT
        self.graph_url = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT
        self.office365 = office365
        self._redirect_code = None

        # Initialize the Credential App.
        self.client_app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=self.AUTHORITY_URL + self.account_type,
            client_credential=self.client_secret,
        )

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
        if does_exists and action == "load":

            # Load the file.
            with open(file=self.credentials, mode="r", encoding="utf-8") as state_file:
                credentials = json.load(fp=state_file)

            # Grab the Token if it exists.
            if "refresh_token" in credentials:

                self.refresh_token = credentials["refresh_token"]
                self.access_token = credentials["access_token"]
                self.id_token = credentials["id_token"]
                self.token_dict = credentials

                return True

            else:
                return False

        # If we are saving the state then open the file and dump the dictionary.
        elif action == "save":

            token_dict["expires_in"] = time.time() + int(token_dict["expires_in"])
            token_dict["ext_expires_in"] = time.time() + int(
                token_dict["ext_expires_in"]
            )

            self.refresh_token = token_dict["refresh_token"]
            self.access_token = token_dict["access_token"]
            self.id_token = token_dict["id_token"]
            self.token_dict = token_dict

            with open(file=self.credentials, mode="w+", encoding="utf-8") as state_file:
                json.dump(obj=token_dict, fp=state_file, indent=2)

    def _token_seconds(self, token_type: str = "access_token") -> int:
        """Determines time till expiration for a token.

        Return the number of seconds until the current access token or refresh token
        will expire. The default value is access token because this is the most commonly used
        token during requests.

        ### Arguments:
        ----
        token_type {str} --  The type of token you would like to determine lifespan for.
            Possible values are ["access_token", "refresh_token"] (default: {access_token})

        ### Returns:
        ----
        {int} -- The number of seconds till expiration.
        """

        # if needed check the access token.
        if token_type == "access_token":

            # if the time to expiration is less than or equal to 0, return 0.
            if not self.access_token or (
                time.time() + 60 >= self.token_dict["expires_in"]
            ):
                return 0

            # else return the number of seconds until expiration.
            token_exp = int(self.token_dict["expires_in"] - time.time() - 60)

        # if needed check the refresh token.
        elif token_type == "refresh_token":

            # if the time to expiration is less than or equal to 0, return 0.
            if not self.refresh_token or (
                time.time() + 60 >= self.token_dict["ext_expires_in"]
            ):
                return 0

            # else return the number of seconds until expiration.
            token_exp = int(self.token_dict["ext_expires_in"] - time.time() - 60)

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

        if self._token_seconds(token_type="access_token") < nseconds:
            self.grab_refresh_token()

    def _silent_sso(self) -> bool:
        """Attempts a Silent Authentication using the Access Token and Refresh Token.

        Returns
        ----
        (bool)
            `True` if it was successful and `False` if it failed.
        """

        # if the current access token is not expired then we are still authenticated.
        if self._token_seconds(token_type="access_token") > 0:
            return True

        # if the current access token is expired then try and refresh access token.
        elif self.refresh_token and self.grab_refresh_token():
            return True

        # More than likely a first time login, so can"t do silent authenticaiton.
        return False

    def login(self) -> None:
        """Logs the user into the session."""

        # Load the State.
        self._state(action="load")

        # Try a Silent SSO First.
        if self._silent_sso():

            # Set the Session.
            self.graph_session = GraphSession(client=self)

            return True

        else:

            # Build the URL.
            url = self.authorization_url()

            # aks the user to go to the URL provided, they will be prompted
            # to authenticate themsevles.
            print(f"Please go to URL provided authorize your account: {url}")

            # ask the user to take the final URL after authentication and
            # paste here so we can parse.
            my_response = input("Paste the full URL redirect here: ")

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

        # Build the Auth URL.
        auth_url = self.client_app.get_authorization_request_url(
            scopes=self.scope, state=self.state, redirect_uri=self.redirect_uri
        )

        return auth_url

    def grab_access_token(self) -> Dict:
        """Exchanges a code for an Access Token.

        ### Returns:
        ----
        dict : A dictionary containing a new access token and refresh token.
        """

        # Parse the Code.
        query_dict = urllib.parse.parse_qs(self._redirect_code)

        # Grab the Code.
        code = query_dict[self.redirect_uri + "?code"]

        # Grab the Token.
        token_dict = self.client_app.acquire_token_by_authorization_code(
            code=code, scopes=self.scope, redirect_uri=self.redirect_uri
        )

        # Save the token dict.
        self._state(action="save", token_dict=token_dict)

        return token_dict

    def grab_refresh_token(self) -> Dict:
        """Grabs a new access token using a refresh token.

        ### Returns
        ----
        dict :
            A token dictionary with a new access token.
        """

        # Grab a new token using our refresh token.
        token_dict = self.client_app.acquire_token_by_refresh_token(
            refresh_token=self.refresh_token, scopes=self.scope
        )

        if "error" in token_dict:
            print(token_dict)
            raise PermissionError(
                "Permissions not authorized, delete json file and run again."
            )

        # Save the Token.
        self._state(action="save", token_dict=token_dict)

        return token_dict

    def users(self) -> Users:
        """Used to access the User Services and metadata.

        ### Returns
        ---
        Users:
            The `Users` services Object.
        """

        # Grab the Users Object for the session.
        users_object: Users = Users(session=self.graph_session)

        return users_object

    def drives(self) -> Drives:
        """Used to access the Drives Services and metadata.

        ### Returns
        ---
        Drives:
            The `Drives` services Object.
        """

        # Grab the Drives Object for the session.
        drives_object: Drives = Drives(session=self.graph_session)

        return drives_object

    def drive_item(self) -> DriveItems:
        """Used to access the Drive Items Services and metadata.

        ### Returns
        ---
        DriveItems:
            The `DriveItems` services Object.
        """

        # Grab the Drive Items Object for the session.
        drive_items_object: Drives = DriveItems(session=self.graph_session)

        return drive_items_object

    def groups(self) -> Groups:
        """Used to access the Groups Services and metadata.

        ### Returns
        ---
        Groups:
            The `Groups` services Object.
        """

        # Grab the `Groups` Object for the session.
        groups_object: Groups = Groups(session=self.graph_session)

        return groups_object

    def notes(self) -> Notes:
        """Used to access the OneNotes Services and metadata.

        ### Returns
        ---
        Notes:
            The `Notes` services Object.
        """

        # Grab the `Notes` Object for the session.
        notes_object: Notes = Notes(session=self.graph_session)

        return notes_object

    def search(self) -> Search:
        """Used to access the Search Services and metadata.

        ### Returns
        ---
        Search:
            The `Search` services Object.
        """

        # Grab the `Search` Object for the session.
        search_object: Search = Search(session=self.graph_session)

        return search_object

    def personal_contacts(self) -> PersonalContacts:
        """Used to access the PersonalContacts Services and metadata.

        ### Returns
        ---
        PersonalContacts:
            The `PersonalContacts` services Object.
        """

        # Grab the `PersonalContacts` Object for the session.
        personal_contacts_object: PersonalContacts = PersonalContacts(
            session=self.graph_session
        )

        return personal_contacts_object

    def mail(self) -> Mail:
        """Used to access the Mail Services and metadata.

        ### Returns
        ---
        Mail:
            The `Mail` services Object.
        """

        # Grab the `Mail` Object for the session.
        mail_service: Mail = Mail(session=self.graph_session)

        return mail_service

    def workbooks(self) -> Workbooks:
        """Used to access the Workbooks Services and metadata.

        ### Returns
        ---
        Workbooks:
            The `Workbooks` services Object.
        """

        # Grab the `Workbooks` Object for the session.
        workbook_service: Workbooks = Workbooks(session=self.graph_session)

        return workbook_service

    def range(self) -> Range:
        """Used to access the Range Services and metadata.

        ### Returns
        ---
        Range:
            The `Range` services Object.
        """

        # Grab the `Range` Object for the session.
        range_service: Range = Range(session=self.graph_session)

        return range_service
