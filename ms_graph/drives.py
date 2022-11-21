from ms_graph.session import GraphSession


class Drives():

    """
    ## Overview:
    ----
    The drive resource is the top level object representing a user"s OneDrive or a
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
        self.endpoint = "drive"
        self.collections_endpoint = "drives"

    def get_root_drive(self) -> dict:
        """Get root folder for user"s default Drive.

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + "/root"
        )

        return content

    def get_root_drive_children(self) -> dict:
        """List children under the Drive for user"s default Drive.

        ### Returns
        ----
        dict :
            A List of Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + "/root/children"
        )

        return content

    def get_root_drive_delta(self) -> dict:
        """List children under the Drive for user"s default Drive.

        ### Returns
        ----
        dict :
            A List of Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + "/root/delta"
        )

        return content

    def get_root_drive_followed(self) -> dict:
        """List user"s followed driveItems.

        ### Returns
        ----
        dict :
            A List of DriveItem Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + "/root/followed"
        )

        return content

    def get_recent_files(self) -> dict:
        """List a set of items that have been recently used by the signed in user.

        ### Overview:
        ----
        This collection includes items that are in the user"s drive
        as well as items they have access to from other drives.

        ### Returns
        ----
        dict :
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint="me/drive/recent"
        )

        return content

    def get_shared_files(self) -> dict:
        """Retrieve a collection of DriveItem resources that have been
        shared with the owner of the Drive.

        ### Returns
        ----
        dict :
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint="me/drive/sharedWithMe"
        )

        return content

    def get_special_folder_by_name(self, folder_name: str) -> dict:
        """Use the special collection to access a special folder by name.

        ### Overview:
        ----
        Special folders provide simple aliases to access well-known folders
        in OneDrive without the need to look up the folder by path (which
        would require localization), or reference the folder with an ID. If
        a special folder is renamed or moved to another location within the
        drive, this syntax will continue to find that folder. Special folders
        are automatically created the first time an application attempts to write
        to one, if it doesn"t already exist. If a user deletes one, it is recreated
        when written to again. Note: If you have read-only permissions and request
        a special folder that doesn"t exist, you"ll receive a 403 Forbidden error.

        ### Returns
        ----
        dict :
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/special/{folder_name}"
        )

        return content

    def get_special_folder_children_by_name(self, folder_name: str) -> dict:
        """Use the special collection to access a collection of Children belonging to special folder
        by name.

        ### Returns
        ----
        dict :
            A List of DriveItem Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/special/{folder_name}/children"
        )

        return content

    def get_drive_by_id(self, drive_id: str) -> dict:
        """Grab"s a Drive Resource using the Drive ID.

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.collections_endpoint + f"/{drive_id}"
        )

        return content

    def get_my_drive(self) -> dict:
        """Get"s the User"s Current OneDrive.

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + "/me"
        )

        return content

    def get_my_drive_children(self, item_id: str) -> dict:
        """Returns a list of DriveItem Resources for the User"s Current OneDrive.

        ### Returns
        ----
        dict :
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint + f"/me/drive/items/{item_id}/children"
        )

        return content

    def get_my_drives(self) -> dict:
        """List children under the Drive for user"s default Drive.

        ### Returns
        ----
        dict :
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.collections_endpoint + "/me"
        )

        return content

    def get_user_drive(self, user_id: str) -> dict:
        """Returns the User"s default OneDrive.

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"users/{user_id}/drive"
        )

        return content

    def get_user_drive_children(self, user_id: str, item_id: str) -> dict:
        """Returns a list of DriveItem Resources for the Default User Drive.

        ### Returns
        ----
        dict :
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"users/{user_id}/drive/items/{item_id}/children"
        )

        return content

    def get_user_drives(self, user_id: str) -> dict:
        """Returns a List Drive Resource Objects for user"s default Drive.

        ### Returns
        ----
        dict :
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"users/{user_id}/drives"
        )

        return content

    def get_group_drive(self, group_id: str) -> dict:
        """Returns a Site Group default Drive..

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"groups/{group_id}/drive"
        )

        return content

    def get_group_drive_children(self, group_id: str, item_id: str) -> dict:
        """Returns a list of DriveItems for the Specified Drive ID for the Specified Group.

        ### Returns
        ----
        dict :
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"groups/{group_id}/drive/items/{item_id}/children"
        )

        return content

    def get_group_drives(self, group_id: str) -> dict:
        """List children under the Drive for user"s default Drive.

        ### Returns
        ----
        dict :
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"groups/{group_id}/drives"
        )

        return content

    def get_sites_drive(self, site_id: str) -> dict:
        """Returns the Default Drive Resource For the Specified Site ID.

        ### Returns
        ----
        dict :
            A Drive Resource Object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"sites/{site_id}/drive"
        )

        return content

    def get_sites_drive_children(self, site_id: str, item_id: str) -> dict:
        """Returns a list of DriveItems for the Specified Drive ID on the Specified Site.

        ### Returns
        ----
        dict :
            A List of DriveChildren Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"sites/{site_id}/drive/items/{item_id}/children"
        )

        return content

    def get_sites_drives(self, site_id: str) -> dict:
        """Returns a List of Drive Resources for the Specified Site ID.

        ### Returns
        ----
        dict :
            A List of Drive Resource Objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"sites/{site_id}/drives"
        )

        return content
