from typing import Dict
from ms_graph.session import GraphSession


class Groups():

    def __init__(self, session: object) -> None:
        """Initializes the `Group` service.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'group'
        self.collections_endpoint = 'groups'

    def list_groups(self) -> Dict:
        """List all the groups in an organization, including but
        not limited to Microsoft 365 groups.

        ### Returns
        -------
        Dict
            If successful, this method returns a 200 OK
            response code and collection of group objects in
            the response body. The response includes only the
            default properties of each group.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint
        )

        return content
