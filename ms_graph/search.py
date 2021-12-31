from ms_graph.session import GraphSession


class Search():

    """
    ## Overview:
    ----
    You can use the Microsoft Search API to query Microsoft 365 data in your apps.
    Search requests run in the context of the signed-in user, identified using an
    access token with delegated permissions.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Query` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = "search"

    def query(self, search_request: dict) -> dict:
        """Runs the query specified in the request body. Search
        results are provided in the response.

        ### Returns
        ----
        dict :
            A `SearchResponse` collection.
        """

        # define the endpoints.
        endpoint = self.endpoint + "/query"

        content = self.graph_session.make_request(
            method="post",
            endpoint=endpoint,
            json=search_request
        )

        return content
