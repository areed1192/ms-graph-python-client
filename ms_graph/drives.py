from typing import Dict
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

    def get_recent_files(self) -> Dict:
        """List a set of items that have been recently used by the signed in user.

        ### Overview:
        ----
        This collection includes items that are in the user's drive
        as well as items they have access to from other drives.

        ### Returns
        ----
        Dict:
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="me/drive/recent"
        )

        return content

    def get_shared_files(self) -> Dict:
        """Retrieve a collection of DriveItem resources that have been shared with the owner of the Drive.

        ### Returns
        ----
        Dict:
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="me/drive/sharedWithMe"
        )

        return content

    def get_special_folder_by_name(self, folder_name: str) -> Dict:
        """Use the special collection to access a special folder by name.

        ### Overview:
        ----
        Special folders provide simple aliases to access well-known folders
        in OneDrive without the need to look up the folder by path (which
        would require localization), or reference the folder with an ID. If
        a special folder is renamed or moved to another location within the
        drive, this syntax will continue to find that folder. Special folders
        are automatically created the first time an application attempts to write
        to one, if it doesn't already exist. If a user deletes one, it is recreated
        when written to again. Note: If you have read-only permissions and request
        a special folder that doesn't exist, you'll receive a 403 Forbidden error.

        ### Returns
        ----
        Dict:
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="/me/drive/special/{folder_name}".format(
                folder_name=folder_name)
        )

        return content

    def get_special_folder_children_by_name(self, folder_name: str) -> Dict:
        """Use the special collection to access a collection of Children belonging to special folder
        by name.

        ### Returns
        ----
        Dict:
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="/me/drive/special/{folder_name}/children".format(
                folder_name=folder_name)
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

    def get_my_drive(self) -> Dict:
        """Get's the User's Current OneDrive.

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/me"
        )

        return content

    def get_my_drive_children(self, item_id: str) -> Dict:
        """Returns a list of DriveItem Resources for the User's Current OneDrive.

        ### Returns
        ----
        Dict:
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/me/drive/items/{item_id}/children".format(
                item_id=item_id
            )
        )

        return content

    def get_my_drives(self) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint + "/me"
        )

        return content

    def get_user_drive(self, user_id: str) -> Dict:
        """Returns the User's default OneDrive.

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="users/{user_id}/drive".format(user_id=user_id)
        )

        return content

    def get_user_drive_children(self, user_id: str, item_id: str) -> Dict:
        """Returns a list of DriveItem Resources for the Default User Drive.

        ### Returns
        ----
        Dict:
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="users/{user_id}/drive/items/{item_id}/children".format(
                user_id=user_id,
                item_id=item_id
            )
        )

        return content

    def get_user_drives(self, user_id: str) -> Dict:
        """Returns a List Drive Resource Objects for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="users/{user_id}/drives".format(user_id=user_id)
        )

        return content

    def get_group_drive(self, group_id: str) -> Dict:
        """Returns a Site Group default Drive..

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="groups/{group_id}/drive".format(group_id=group_id)
        )

        return content

    def get_group_drive_children(self, group_id: str, item_id: str) -> Dict:
        """Returns a list of DriveItems for the Specified Drive ID for the Specified Group.

        ### Returns
        ----
        Dict:
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="groups/{group_id}/drive/items/{item_id}/children".format(
                group_id=group_id,
                item_id=item_id
            )
        )

        return content

    def get_group_drives(self, group_id: str) -> Dict:
        """List children under the Drive for user's default Drive.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="groups/{group_id}/drives".format(group_id=group_id)
        )

        return content

    def get_sites_drive(self, site_id: str) -> Dict:
        """Returns the Default Drive Resource For the Specified Site ID.

        ### Returns
        ----
        Dict:
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="sites/{site_id}/drive".format(site_id=site_id)
        )

        return content

    def get_sites_drive_children(self, site_id: str, item_id: str) -> Dict:
        """Returns a list of DriveItems for the Specified Drive ID on the Specified Site.

        ### Returns
        ----
        Dict:
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="sites/{site_id}/drive/items/{item_id}/children".format(
                site_id=site_id,
                item_id=item_id
            )
        )

        return content

    def get_sites_drives(self, site_id: str) -> Dict:
        """Returns a List of Drive Resources for the Specified Site ID.

        ### Returns
        ----
        Dict:
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint="sites/{site_id}/drives".format(site_id=site_id)
        )

        return content
