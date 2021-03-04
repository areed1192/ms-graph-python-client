from typing import Dict
from ms_graph.session import GraphSession


class Workbooks():

    """
    ## Overview:
    ----
    You can use Microsoft Graph to allow web and mobile applications to
    read and modify Excel workbooks stored in OneDrive for Business, SharePoint
    site or Group drive. The Workbook (or Excel file) resource contains all the
    other Excel resources through relationships. You can access a workbook through
    the Drive API by identifying the location of the file in the URL.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Workbooks` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'workbook'
