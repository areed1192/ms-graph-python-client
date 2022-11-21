from ms_graph.session import GraphSession


class PersonalContacts():

    """
    ## Overview:
    ----
    A contact is an item in Outlook where you can organize and save
    information about the people and organizations you communicate
    with. Contacts are contained in contact folders.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `PersonalContacts` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = "contacts"
        self.endpoint_folders = "contactFolders"

    def list_my_contacts(self) -> dict:
        """Retrieves all the contacts from the users mailbox.

        ### Returns
        ----
        dict :
            A List of `Contact` Resource Object.
        """

        # define the endpoints.
        endpoint = "me/" + self.endpoint

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def list_my_contacts_folder(self) -> dict:
        """Retrieves all the contacts folders from the users mailbox.

        ### Returns
        ----
        dict :
            A List of `ContactFolders` Resource Object.
        """

        # define the endpoints.
        endpoint = "me/" + self.endpoint_folders

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def list_contacts_folder_by_id(self, user_id: str, folder_id: str) -> dict:
        """Retrieves all the contacts folders from the users mailbox.

        ### Parameters
        ----
        user_id : str
            The User ID that the folder belongs to.

        folder_id : str
            The folder ID you want to retrieve.

        ### Returns
        ----
        dict :
            A List of `ContactFolders` Resource Object.
        """

        # define the endpoints.
        endpoint = f"users/{user_id}/" + self.endpoint_folders + f"/{folder_id}"

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def create_my_contact_folder(self, folder_resource: dict) -> dict:
        """Creates a new Contact Folder under the default users profile.

        ### Parameters
        ----
        folder_resource : dict
            A dictionary that specifies the folder resource
            attributes like the folder ID and folder display
            value.

        ### Returns
        ----
        dict :
            A `ContactFolder` Resource Object.
        """

        # define the endpoints.
        endpoint = "me/" + self.endpoint_folders

        content = self.graph_session.make_request(
            method="post",
            endpoint=endpoint,
            json=folder_resource
        )

        return content

    def create_user_contact_folder(self, user_id: str, folder_resource: dict) -> dict:
        """Creates a new Contact Folder under the specified users profile.

        ### Parameters
        ----
        user_id : str
            The User ID that the folder belongs to.

        folder_resource : dict
            A dictionary that specifies the folder resource
            attributes like the folder ID and folder display
            value.

        ### Returns
        ----
        dict :
            A `ContactFolder` Resource Object.
        """

        # define the endpoints.
        endpoint = f"users/{user_id}/" + self.endpoint_folders

        content = self.graph_session.make_request(
            method="post",
            endpoint=endpoint,
            json=folder_resource
        )

        return content

    def get_my_contacts_folder_by_id(self, folder_id: str) -> dict:
        """Retrieves a contactsFolder resource using the specified ID.

        ### Parameters
        ----
        folder_id : str
            The folder ID you want to retrieve.

        ### Returns
        ----
        dict :
            A `ContactFolder` Resource Object.
        """

        # define the endpoints.
        endpoint = "me/" + self.endpoint_folders + f"/{folder_id}"

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def get_contacts_folder_by_id(self, user_id: str, folder_id: str) -> dict:
        """Retrieves a contactsFolder resource using the specified ID for the
        specified user.

        ### Parameters
        ----
        user_id : str
            The User ID that the folder belongs to.

        folder_id : str
            The folder ID you want to retrieve.

        ### Returns
        ----
        dict :
            A `ContactFolder` Resource Object.
        """

        # define the endpoints.
        endpoint = f"users/{user_id}/" + self.endpoint_folders + f"/{folder_id}"

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def get_my_contact_by_id(self, contact_id: str) -> dict:
        """Retrieves the Contact Resource for the specified contact ID.

        ### Parameters
        ----
        contact_id : str
            An authenticated session for our Microsoft Graph Client.

        ### Returns
        ----
        dict :
            A List of `Contact` Resource Object.
        """

        # define the endpoints.
        endpoint = "me/" + self.endpoint + f"/{contact_id}"

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content
