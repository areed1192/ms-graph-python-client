from typing import List
from typing import Dict
from typing import Union

from ms_graph.session import GraphSession


class Drives():

    """
    ## Overview:
    ----
    The drive resource is the top level object representing a user's OneDrive or a 
    document library in SharePoint. OneDrive users will always have at least one drive
    available, their default drive. Users without a OneDrive license may not have a default
    drive available.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Drives` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'drive'
        self.collections_endpoint = 'drives'

    def get_root_drive(self) -> Dict:
        """Get root folder for user's default Drive.

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root"
        )

        return content

    def get_root_drive_children(self) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/children"
        )

        return content

    def get_root_drive_delta(self) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Object.
        """
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/delta"
        )

        return content

    def get_root_drive_followed(self) -> Dict:
        """List user's followed driveItems.

        ### Returns
        ----
        Dict:
            A List of DriveItem Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/followed"
        )

        return content

    def get_drive_by_id(self, drive_id: str) -> Dict:
        """Grab's a Drive Resource using the Drive ID.

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint + "/{id}".format(id=drive_id)
        )

        return content

    def get_my_drives(self) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint + "/me"
        )

        return content

    def get_user_drives(self, user_id: str) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Object.
        """
        content = self.graph_session.make_request(
            method='get',
            endpoint="users/{user_id}/drives".format(user_id=user_id)
        )

        return content

    def get_group_drives(self, group_id: str) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="groups/{group_id}/drives".format(group_id=group_id)
        )

        return content
