from ms_graph.session import GraphSession


class Users():

    """
    ## Overview:
    ----
    You can use Microsoft Graph to build compelling app experiences
    based on users, their relationships with other users and groups,
    and their mail, calendar, and files.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Users` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = "users"

    def list_users(self) -> dict:
        """Retrieve a list of user objects.

        ### Returns
        ----
        dict :
            If successful, this method returns a 200 OK response code
            and collection of user objects in the response body. If a
            large user collection is returned, you can use paging in your
            app.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=self.endpoint
        )

        return content
