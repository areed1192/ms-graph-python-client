from ms_graph.session import GraphSession


class WorkbookComments:

    """
    ## Overview:
    ----
    Represents a comment in workbook.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `WorkbookComment` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

    def list(self, item_id: str) -> dict:
        """Retrieve a list of workbookComment objects using the
        Item ID.

        ### Parameters
        ----
        item_id : str
            The drive item resource id.

        ### Returns
        ----
        dict:
            A collection of WorkbookComment objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/items/{item_id}/workbook/comments",
        )

        return content

    def get(self, item_id: str, comment_id: str) -> dict:
        """Retrieve the properties and relationships of
        a WorkbookComment object.

        ### Parameters
        ----
        item_id : str
            The drive item resource id.

        comment_id : str
            The comment resource id.

        ### Returns
        ----
        dict:
            A WorkbookComment object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/items/{item_id}/workbook/comments/{comment_id}",
        )

        return content

    def list_replies(self, item_id: str, comment_id) -> dict:
        """Retrieve a list of workbookCommentReply objects using the
        Item ID.

        ### Parameters
        ----
        item_id : str
            The drive item resource id.

        comment_id : str
            The comment resource id.

        ### Returns
        ----
        dict:
            A collection of WorkbookCommentReply objects.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/items/{item_id}/workbook/comments/{comment_id}/replies",
        )

        return content

    def create_reply(
        self, item_id: str, comment_id: str, content: str, content_type: str = "plain"
    ) -> dict:
        """Creates a new WorkbookCommentReply object using the
        specified Item ID.

        ### Parameters
        ----
        item_id : str
            The drive item resource id.

        comment_id : str
            The comment resource id.

        content : str
            The content of a comment reply.

        content_type : str (optional, Default='plain')
            Indicates the type for the comment reply.

        ### Returns
        ----
        dict:
            A WorkbookCommentReply object.
        """

        body = {"content": content, "contentType": content_type}

        content = self.graph_session.make_request(
            method="post",
            json=body,
            additional_headers={"Content-type": "application/json"},
            endpoint=f"/me/drive/items/{item_id}/workbook/comments/{comment_id}/replies",
        )

        return content

    def get_reply(self, item_id: str, comment_id: str, reply_id: str) -> dict:
        """Retrieve the properties and relationships of WorkbookCommentReply
        object using the specified Item ID.

        ### Parameters
        ----
        item_id : str
            The drive item resource id.

        comment_id : str
            The comment resource id.

        reply_id : str
            The comment reply resource id.

        ### Returns
        ----
        dict:
            A WorkbookCommentReply object.
        """

        content = self.graph_session.make_request(
            method="get",
            endpoint=f"/me/drive/items/{item_id}/workbook/comments/{comment_id}/replies/{reply_id}",
        )

        return content
