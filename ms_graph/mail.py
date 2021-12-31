from ms_graph.session import GraphSession


class Mail:

    """
    ## Overview:
    ----
    Microsoft Graph lets your app get authorized access to a user"s
    Outlook mail data in a personal or organization account. With the
    appropriate delegated or application mail permissions, your app can
    access the mail data of the signed-in user or any user in a tenant.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Mail` service.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = "mail"

    def list_my_messages(self) -> dict:
        """Get the messages in the signed-in user"s mailbox
        (including the Deleted Items and Clutter folders).

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response
            code and collection of `Message` objects in the response
            body.
        """

        content = self.graph_session.make_request(method="get", endpoint="/me/messages")

        return content

    def list_user_messages(self, user_id: str) -> dict:
        """Get the messages in the user"s mailbox
        (including the Deleted Items and Clutter folders).

        ### Parameters
        ----
        user_id : str
            The user for which to query messages for.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response
            code and collection of `Message` objects in the response
            body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/messages"
        )

        return content

    def create_my_message(self, message: dict) -> dict:
        """Use this API to create a draft of a new message.

        ### Overview:
        ----
        Drafts can be created in any folder and optionally
        updated before sending. To save to the Drafts folder,
        use the /messages shortcut.

        ### Parameters
        ----
        message : dict
            A JSON payload with the required message
            attributes.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created
            response code and `message` object in the response
            body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint="/me/messages", json=message
        )

        return content

    def create_user_message(self, user_id: str, message: dict) -> dict:
        """Use this API to create a draft of a new message for the specific
        user ID.

        ### Overview:
        ----
        Drafts can be created in any folder and optionally
        updated before sending. To save to the Drafts folder,
        use the /messages shortcut.

        ### Parameters
        ----
        message : dict
            A JSON payload with the required message
            attributes.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created
            response code and `message` object in the response
            body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/messages", json=message
        )

        return content

    def get_my_messages(self, message_id: str) -> dict:
        """Retrieve the properties and relationships of a message object for
        the default user.

        ### Parameters
        ----
        message_id : str
            The message ID you want to query.

        ### Returns
        ----
        dict
            If successful, this method returns 200 successful
            response code and `message` object in the response
            body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/me/messages/{message_id}"
        )

        return content

    def get_user_messages(self, user_id: str, message_id: str) -> dict:
        """Retrieve the properties and relationships of a message object for
        a specific user.

        ### Parameters
        ----
        user_id : str
            The User ID you want to query messages for.

        message_id : str
            The message ID you want to query.

        ### Returns
        ----
        dict
            If successful, this method returns 200 successful
            response code and `message` object in the response
            body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/messages/{message_id}"
        )

        return content

    def update_my_message(self, message: dict) -> dict:
        """Update the properties of a message object for the
        default user.

        ### Parameters
        ----
        message : str
            In the request body, supply the values for relevant
            fields that should be updated.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and updated `message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="patch", endpoint="/me/messages", json=message
        )

        return content

    def update_user_message(self, user_id: str, message: dict) -> dict:
        """Update the properties of a message object for the
        specified user.

        ### Parameters
        ----
        user_id : str
            The user for which to update a message for.

        message : str
            In the request body, supply the values for relevant
            fields that should be updated.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and updated `message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="patch", endpoint=f"/users/{user_id}/messages", json=message
        )

        return content

    def delete_my_message(self, message_id: str) -> dict:
        """Delete a message in the default user"s mailbox,
        or delete a relationship of the message.

        ### Parameters
        ----
        message_id : str
            The ID of the message you wish to delete.

        ### Returns
        ----
        dict
            If successful, this method returns a 204 No Content
            response code. It does not return anything in
            the response body..
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"me/messages/{message_id}"
        )

        return content

    def delete_user_message(self, user_id: str, message_id: str) -> dict:
        """Delete a message in the specified user"s mailbox,
        or delete a relationship of the message.

        ### Parameters
        ----
        user_id : str
            The user for which to delete a message for.

        message_id : str
            The ID of the message you wish to delete.

        ### Returns
        ----
        dict
            If successful, this method returns a 204 No Content
            response code. It does not return anything in
            the response body..
        """

        content = self.graph_session.make_request(
            method="delete", endpoint=f"/users/{user_id}/messages/{message_id}"
        )

        return content

    def send_my_message(self, message_id: str) -> dict:
        """Send a message in the draft folder for the default user.

        ### Overview:
        ----
        The draft message can be a new message draft,
        reply draft, reply-all draft, or a forward draft.
        The message is then saved in the Sent Items folder.

        ### Parameters
        ----
        message_id : str
            The ID of the message you wish to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted
            response code. It does not return anything in
            the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/send"
        )

        return content

    def send_user_message(self, user_id: str, message_id: str) -> dict:
        """Send a message in the draft folder for the specified user.

        ### Overview:
        ----
        The draft message can be a new message draft,
        reply draft, reply-all draft, or a forward draft.
        The message is then saved in the Sent Items folder.

        ### Parameters
        ----
        user_id : str
            The user for which to send a message for.

        message_id : str
            The ID of the message you wish to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted
            response code. It does not return anything in
            the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/users/{user_id}/messages/{message_id}/send"
        )

        return content

    def copy_my_message(self, message_id: str) -> dict:
        """Send a message in the draft folder for the specified user.

        ### Overview:
        ----
        The draft message can be a new message draft,
        reply draft, reply-all draft, or a forward draft.
        The message is then saved in the Sent Items folder.

        ### Parameters
        ----
        message_id : str
            The ID of the message you wish to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted
            response code. It does not return anything in
            the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/copy"
        )

        return content

    def copy_user_message(self, user_id: str, message_id: str) -> dict:
        """Send a message in the draft folder for the specified user.

        ### Overview:
        ----
        The draft message can be a new message draft,
        reply draft, reply-all draft, or a forward draft.
        The message is then saved in the Sent Items folder.

        ### Parameters
        ----
        user_id : str
            The user for which to send a message for.

        message_id : str
            The ID of the message you wish to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted
            response code. It does not return anything in
            the response body.
        """
        content = self.graph_session.make_request(
            method="post", endpoint=f"/users/{user_id}/messages/{message_id}/copy"
        )

        return content

    def move_my_message(self, message_id: str, destination_id: str) -> dict:
        """Move a message to another folder within the specified user"s mailbox.
        For the default user.

        ### Overview:
        ----
        This creates a new copy of the message in the destination folder and
        removes the original message.

        ### Parameters
        ----
        message_id : str
            The ID of the message you wish to move.

        destination_id : str
            The name of the folder you want to move it to.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/me/messages/{message_id}/move",
            json={"destinationId": destination_id},
        )

        return content

    def move_user_message(
        self, user_id: str, message_id: str, destination_id: str
    ) -> dict:
        """Move a message to another folder within the specified user"s mailbox.

        ### Overview:
        ----
        This creates a new copy of the message in the destination folder and
        removes the original message.

        ### Parameters
        ----
        user_id : str
            The user ID of the mailbox that you want to move a
            message for.

        message_id : str
            The ID of the message you wish to move.

        destination_id : str
            The name of the folder you want to move it to.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/move",
            json={"destinationId": destination_id},
        )

        return content

    def create_reply_my_message(self, message_id: str) -> dict:
        """Create a draft of the reply to the specified message. For
        the default user.

        ### Overview:
        ----
        You can then update the draft to add reply content to the
        body or change other message properties, or, simply send
        the draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to use
            as a reply message.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """
        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/createReply"
        )

        return content

    def create_reply_user_message(self, user_id: str, message_id: str) -> dict:
        """Create a draft of the reply to the specified message.

        ### Overview:
        ----
        You can then update the draft to add reply content to the
        body or change other message properties, or, simply send
        the draft.

        ### Parameters
        ----
        user_id : str
            The user ID of the mailbox that you want to create a
            reply message for.

        message_id : str
            The message ID for which you wish to use
            as a reply message.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/createReply",
        )

        return content

    def reply_to_my_message(self, message_id: str, message: dict) -> dict:
        """Reply to the sender of a message, add a comment or modify any updateable properties
        all in one reply call. The message is then saved in the Sent Items folder. For the
        default user.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to reply to.

        messgae : dict
            The message you want to reply with.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response
            code. It does not return anything in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/reply", json=message
        )

        return content

    def reply_to_user_message(
        self, user_id: str, message_id: str, message: dict
    ) -> dict:
        """Reply to the sender of a message, add a comment or modify any updateable properties
        all in one reply call. The message is then saved in the Sent Items folder.


        ### Parameters
        ----
        user_id : str
            The user ID of the mailbox that contains the message
            you want to reply to.

        message_id : str
            The message ID for which you wish to reply to.

        messgae : dict
            The message you want to reply with.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response
            code. It does not return anything in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/reply",
            json=message,
        )

        return content

    def create_reply_all_my_message(self, message_id: str) -> dict:
        """Create a draft to reply to the sender and all the recipients of the specified message.
        For the default user.

        ### Overview:
        ----
        You can then update the draft to add reply content to the body
        or change other message properties, or, simply send the draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to use
            as a repy message

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/createReplyAll"
        )

        return content

    def create_reply_all_user_message(self, user_id: str, message_id: str) -> dict:
        """Create a draft to reply to the sender and all the recipients of the specified message.

        ### Overview:
        ----
        You can then update the draft to add reply content to the body
        or change other message properties, or, simply send the draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to reply all
            to.

        user_id : str
            The User ID you want to reply all messages for.

        message : dict
            The message you want to respond with.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/createReplyAll",
        )

        return content

    def reply_all_my_message(self, message_id: str, message: dict) -> dict:
        """Create a draft to reply to the sender and all the recipients of the
        specified message for the default user.

        ### Overview:
        ----
        You can then update the draft to add reply content to the body
        or change other message properties, or, simply send the draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to reply all
            to.

        message : dict
            The message you want to respond with.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/replyAll", json=message
        )

        return content

    def reply_all_user_message(
        self, user_id: str, message_id: str, message: dict
    ) -> dict:
        """Create a draft to reply to the sender and all the recipients of the specified message.

        ### Overview:
        ----
        You can then update the draft to add reply content to the body
        or change other message properties, or, simply send the draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to reply all
            to.

        user_id : str
            The User ID you want to reply all messages for.

        message : dict
            The message you want to respond with.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/replyAll",
            json=message,
        )

        return content

    def create_forward_my_message(self, message_id: str) -> dict:
        """Create a draft to forward the specified message. For the default user.

        ### Overview:
        ----
        You can then update the draft to add content to the body
        or change other message properties, or, simply send the
        draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to forward.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and `Message` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/createForward"
        )

        return content

    def create_forward_user_message(self, user_id: str, message_id: str) -> dict:
        """Create a draft to forward the specified message.

        ### Overview:
        ----
        You can then update the draft to add content to the body
        or change other message properties, or, simply send the
        draft.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to forward.

        user_id : dict
            The User ID you want to create a new forward
            message for.

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response
            code and Message object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/createForward",
        )

        return content

    def forward_my_message(self, message_id: str, message: dict) -> dict:
        """Forward a message for the default user. The message is saved in the
        Sent Items folder.

        ### Parameters
        ----
        message_id : str
            The message ID for which you wish to forward.

        message : dict
            The message to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response code.
            It does not return anything in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint=f"/me/messages/{message_id}/forward", json=message
        )

        return content

    def forward_user_message(
        self, user_id: str, message_id: str, message: dict
    ) -> dict:
        """Forward a message. The message is saved in the Sent Items folder.

        ### Parameters
        ----
        user_id : str
            The user ID which contains the email you want to
            forward.

        message_id : str
            The message ID for which you wish to forward.

        message : dict
            The message to send.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response code.
            It does not return anything in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/messages/{message_id}/forward",
            json=message,
        )

        return content

    def send_my_mail(
        self, message_id: str, message: dict, save_to_send_items: bool = True
    ) -> dict:
        """Send the message specified in the request body for the default user.

        ### Overview:
        ----
        The message is saved in the Sent Items folder by default. You
        can include a file attachment in the same sendMail action call.

        ### Parameters
        ----
        message_id : dict
            The message ID for which you want to send.

        save_to_send_items : bool (optional)
            Indicates whether to save the message in Sent Items. Specify
            it only if the parameter is false; default is true.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response code.
            It does not return anything in the response body.
        """

        message["saveToSentItems"] = save_to_send_items

        content = self.graph_session.make_request(
            method="post", endpoint="/me/sendMail", json=message
        )

        return content

    def send_user_mail(
        self, user_id: str, message: dict, save_to_send_items: bool = True
    ) -> dict:
        """Send the message specified in the request body.

        ### Overview:
        ----
        The message is saved in the Sent Items folder by default. You
        can include a file attachment in the same sendMail action call.


        ### Parameters
        ----
        user_id : str
            The user for which to send a mailItem resource for.

        message : dict
            The message to send.

        save_to_send_items : bool, optional
            Indicates whether to save the message in Sent Items. Specify it
            only if the parameter is false, by default True.

        ### Returns
        ----
        dict
            If successful, this method returns 202 Accepted response code.
            It does not return anything in the response body.
        """

        message["saveToSentItems"] = save_to_send_items

        content = self.graph_session.make_request(
            method="post", endpoint=f"/users/{user_id}/sendMail", json=message
        )

        return content

    def list_my_attachements(self, message_id: str) -> dict:
        """Retrieve a list of `attachment` objects attached to a message.

        ### Parameters
        ----
        message_id : str
            The message Id of the mailItem resource that
            you want to query attachments for.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and collection of Attachment objects in the response body.
        """
        content = self.graph_session.make_request(
            method="get", endpoint=f"/me/messages/{message_id}/attachments"
        )

        return content

    def list_user_attachements(self, user_id: str, message_id: str) -> dict:
        """Get all the `messageRule` objects defined for the user"s Inbox. For
        the default user.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/messages/{message_id}/attachments"
        )

        return content

    def list_my_rules(self) -> dict:
        """Get all the `messageRule` objects defined for the user"s Inbox. For
        the default user.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint="/me/mailFolders/inbox/messageRules"
        )

        return content

    def list_rules(self, user_id: str) -> dict:
        """Get all the `messageRule` objects defined for the user"s Inbox. For
        the specific user.

        ### Parameters
        ----
        user_id : str
            The user ID for which to query `messageRules` for.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/mailFolders/inbox/messageRules"
        )

        return content

    def create_my_message_rule(self, rule: dict) -> dict:
        """Create a messageRule object by specifying a set of conditions and actions for
        the default user.

        ### Parameters
        ----
        rule : dict
            The parameters that are applicable to your rule.
            For more info:
            https://docs.microsoft.com/en-us/graph/api/mailfolder-post-messagerules?view=graph-rest-1.0&tabs=http#request-body

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response code and a
            `messageRule` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post", endpoint="/me/mailFolders/inbox/messageRules", json=rule
        )

        return content

    def create_message_rule(self, user_id: str, rule: dict) -> dict:
        """Create a messageRule object by specifying a set of conditions and actions
        for the specified User.

        ### Parameters
        ----
        user_id : str
            The User ID for which to create the message rule
            for.

        rule : dict
            The parameters that are applicable to your rule.
            For more info:
            https://docs.microsoft.com/en-us/graph/api/mailfolder-post-messagerules?view=graph-rest-1.0&tabs=http#request-body

        ### Returns
        ----
        dict
            If successful, this method returns 201 Created response code and a
            `messageRule` object in the response body.
        """

        content = self.graph_session.make_request(
            method="post",
            endpoint=f"/users/{user_id}/mailFolders/inbox/messageRules",
            json=rule,
        )

        return content

    def list_my_overrides(self) -> dict:
        """Get the overrides that a user has set up to always classify messages from
        certain senders in specific ways.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code and a collection
            of `inferenceClassificationOverride1 objects in the response body. An
            empty collection is returned if the user doesn"t have any overrides
            set up.
        """

        content = self.graph_session.make_request(
            method="get", endpoint="/me/inferenceClassification/overrides"
        )

        return content

    def list_overrides(self, user_id: str) -> dict:
        """Get the overrides that a user has set up to always classify messages from
        certain senders in specific ways.

        ### Parameters
        ----
        user_id : str
            The User ID for which to query `inferenceClassificationOverride`
            objects for.

        ### Returns
        ----
        dict
            If successful, this method returns a 200 OK response code and a collection
            of `inferenceClassificationOverride` objects in the response body. An
            empty collection is returned if the user doesn"t have any overrides
            set up.
        """

        content = self.graph_session.make_request(
            method="get", endpoint=f"/users/{user_id}/inferenceClassification/overrides"
        )

        return content
