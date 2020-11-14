from typing import List
from typing import Dict
from typing import Union

from ms_graph.session import GraphSession


class DriveItems():

    """
    ## Overview:
    ----
    The driveItem resource represents a file, folder, 
    or other item stored in a drive. All file system 
    objects in OneDrive and SharePoint are returned as 
    driveItem resources.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `DriveItems` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'drive'
        self.collections_endpoint = 'drives/'

    
    def get_drive_item(self, drive_id: str, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        drive_id : str
            The Drive ID in which the resource exist.

        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= self.collections_endpoint + "/{drive_id}/items/{item_id}".format(
                drive_id=drive_id,
                item_id=item_id
            )
        )

        return content

    def get_drive_item_by_path(self, drive_id: str, item_path: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        drive_id : str
            The Drive ID in which the resource exist.

        item_path : str
            The path to the Item.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= self.collections_endpoint + "/{drive_id}/root:/{path}".format(
                drive_id=drive_id,
                path=item_path
            )
        )

        return content

    def get_group_drive_item(self, group_id: str, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        group_id : str
            The Group ID in which the resource exist.

        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/groups/{group_id}/drive/items/{item_id}".format(
                group_id=group_id,
                item_id=item_id
            )
        )

        return content

    def get_group_drive_item_by_path(self, group_id: str, item_path: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        drive_id : str
            The Drive ID in which the resource exist.

        item_path : str
            The path to the Item.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/groups/{group_id}/drive/root:/{item_path}".format(
                group_id=group_id,
                item_path=item_path
            )
        )

        return content

    def get_my_drive_item(self, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/me/drive/items/{item_id}".format(
                item_id=item_id
            )
        )

        return content

    def get_my_drive_item_by_path(self, item_path: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        item_path : str
            The path to the Item.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/me/drive/root:/{item_path}".format(
                item_path=item_path
            )
        )

        return content

    def get_site_drive_item(self, site_id: str, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        site_id : str
            The site ID which to query the item from.

        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/sites/{site_id}/drive/items/{item_id}".format(
                site_id=site_id,
                item_id=item_id
            )
        )

        return content

    def get_site_drive_item_by_path(self, site_id: str, item_path: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        site_id : str
            The site ID which to query the item from.

        item_path : str
            The path to the Item.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/sites/{site_id}/drive/root:/{item_path}".format(
                site_id=site_id,
                item_path=item_path
            )
        )

        return content

    def get_site_drive_item_from_list(self, site_id: str, list_id: str, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        site_id : str
            The site ID which to query the item from.

        list_id : str
            The list ID which to query the item from.

        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/sites/{site_id}/lists/{list_id}/items/{item_id}/driveItem".format(
                site_id=site_id,
                list_id=list_id,
                item_id=item_id
            )
        )

        return content

    def get_user_drive_item(self, user_id: str, item_id: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        user_id : str
            The User ID which to query the item from.

        item_id : str
            The item ID of the object you want to
            return.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/users/{user_id}/drive/items/{item_id}".format(
                user_id=user_id,
                item_id=item_id
            )
        )

        return content

    def get_user_drive_item_by_path(self, user_id: str, item_path: str) -> Dict:     
        """Grab's a DriveItem Resource using the Item ID and Drive ID.

        ### Parameters
        ----
        site_id : str
            The User ID which to query the item from.

        item_path : str
            The path to the Item.

        ### Returns
        ----
        Dict:
            A DriveItem resource object.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint= "/users/{user_id}/drive/root:/{item_path}".format(
                user_id=user_id,
                item_path=item_path
            )
        )

        return content


# GET /drives/{drive-id}/items/{item-id}
# GET /drives/{drive-id}/root:/{item-path}
# GET /groups/{group-id}/drive/items/{item-id}
# GET /groups/{group-id}/drive/root:/{item-path}
# GET /me/drive/items/{item-id}
# GET /me/drive/root:/{item-path}

# GET /sites/{site-id}/drive/items/{item-id}
# GET /sites/{site-id}/drive/root:/{item-path}
# GET /sites/{site-id}/lists/{list-id}/items/{item-id}/driveItem

# GET /users/{user-id}/drive/items/{item-id}
# GET /users/{user-id}/drive/root:/{item-path}
