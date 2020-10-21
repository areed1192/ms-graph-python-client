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
        self.collections_endpoint = 'drives'
