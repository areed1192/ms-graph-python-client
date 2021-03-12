from typing import Dict
from ms_graph.session import GraphSession


class Mail():

    """
    ## Overview:
    ----
    Microsoft Graph lets your app get authorized access to a user's
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
        self.endpoint = 'mail'

    def list_my_messages(self) -> Dict:
        """Get the messages in the signed-in user's mailbox
        (including the Deleted Items and Clutter folders).

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response
            code and collection of `Message` objects in the response
            body.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages'
        )

        return content

    def list_user_messages(self, user_id: str) -> Dict:
        """Get the messages in the user's mailbox
        (including the Deleted Items and Clutter folders).

        ### Parameters
        ----
        user_id : str
            The user for which to query messages for.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response
            code and collection of `Message` objects in the response
            body.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id)
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
            method='post',
            endpoint='/me/messages',
            json=message
        )

        return content

    def create_user_message(self, user_id: str, message: dict) -> Dict:
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
            method='get',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id),
            json=message
        )

        return content

    def get_my_messages(self, message_id: str) -> Dict:
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
            method='get',
            endpoint='/me/messages/{message_id}'.format(
                message_id=message_id
            )
        )

        return content

    def get_user_messages(self, user_id: str, message_id: str) -> Dict:
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
            method='get',
            endpoint='/users/{user_id}/messages/{message_id}'.format(
                user_id=user_id,
                message_id=message_id
            )
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
            method='patch',
            endpoint='/me/messages',
            json=message
        )

        return content

    def update_user_message(self, user_id: str, message: dict) -> Dict:
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
            method='patch',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id),
            json=message
        )

        return content

    def delete_my_message(self, message_id: str) -> dict:
        """Delete a message in the default user's mailbox,
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
            method='get',
            endpoint='/me/messages/{message_id}'.format(
                message_id=message_id
            )
        )

        return content

    def delete_user_message(self, user_id: str, message_id: str) -> Dict:
        """Delete a message in the specified user's mailbox,
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
            method='delete',
            endpoint='/users/{user_id}/messages/{message_id}'.format(
                user_id=user_id,
                message_id=message_id
            )
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
            method='post',
            endpoint='/me/messages/{message_id}/send'.format(
                message_id=message_id
            )
        )

        return content

    def send_user_message(self, user_id: str, message_id: str) -> Dict:
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
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/send'.format(
                user_id=user_id,
                message_id=message_id
            )
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
            method='post',
            endpoint='/me/messages/{message_id}/copy'.format(
                message_id=message_id
            )
        )

        return content

    def copy_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/copy'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def move_my_message(self, message_id: str, destination_id: str) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/move'.format(
                message_id=message_id
            ),
            json={"destinationId": destination_id}
        )

        return content

    def move_user_message(self, user_id: str, message_id: str, destination_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/move'.format(
                user_id=user_id,
                message_id=message_id
            ),
            json={"destinationId": destination_id}
        )

        return content

    def create_reply_my_message(self, message_id: str) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/createReply'.format(
                message_id=message_id
            )
        )

        return content

    def create_reply_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/createReply'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def reply_to_my_message(self, message_id: str, message: dict) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/reply'.format(
                message_id=message_id
            ),
            json=message
        )

        return content

    def reply_to_user_message(self, user_id: str, message_id: str, message: dict) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/reply'.format(
                user_id=user_id,
                message_id=message_id
            ),
            json=message
        )

        return content

    def create_reply_all_my_message(self, message_id: str) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/createReplyAll'.format(
                message_id=message_id
            )
        )

        return content

    def create_reply_all_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/createReplyAll'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def reply_all_my_message(self, message_id: str, message: dict) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/replyAll'.format(
                message_id=message_id
            ),
            json=message
        )

        return content

    def reply_all_user_message(self, user_id: str, message_id: str, message: dict) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/replyAll'.format(
                user_id=user_id,
                message_id=message_id
            ),
            json=message
        )

        return content

    def create_forward_my_message(self, message_id: str) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/createForward'.format(
                message_id=message_id
            )
        )

        return content

    def create_forward_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/createForward'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def forward_my_message(self, message_id: str, message: dict) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/forward'.format(
                message_id=message_id
            ),
            json=message
        )

        return content

    def forward_user_message(self, user_id: str, message_id: str, message: dict) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/forward'.format(
                user_id=user_id,
                message_id=message_id
            ),
            json=message
        )

        return content

    def send_my_mail(self, message_id: str, message: dict, save_to_send_items: bool) -> dict:

        message['saveToSentItems'] = save_to_send_items

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/sendMail',
            json=message
        )

        return content

    def send_user_mail(self, user_id: str, message: dict, save_to_send_items: bool = True) -> Dict:
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
        Dict
            If successful, this method returns 202 Accepted response code.
            It does not return anything in the response body.
        """

        message['saveToSentItems'] = save_to_send_items

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/sendMail'.format(
                user_id=user_id
            ),
            json=message
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
        Dict
            If successful, this method returns a 200 OK response code
            and collection of Attachment objects in the response body.
        """
        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages/{message_id}/attachments'.format(
                message_id=message_id
            )
        )

        return content

    def list_user_attachements(self, user_id: str, message_id: str) -> Dict:
        """Get all the `messageRule` objects defined for the user's Inbox. For
        the default user.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages/{message_id}/attachments'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def list_my_rules(self) -> Dict:
        """Get all the `messageRule` objects defined for the user's Inbox. For
        the default user.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/mailFolders/inbox/messageRules'
        )

        return content

    def list_rules(self, user_id: str) -> Dict:
        """Get all the `messageRule` objects defined for the user's Inbox. For
        the specific user.

        ### Parameters
        ----
        user_id : str
            The user ID for which to query `messageRules` for.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response code
            and collection of messageRule objects in the response body.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/mailFolders/inbox/messageRules'.format(
                user_id=user_id
            )
        )

        return content

    def create_my_message_rule(self, rule: Dict) -> Dict:
        """Create a messageRule object by specifying a set of conditions and actions for
        the default user.

        ### Parameters
        ----
        rule : Dict
            The parameters that are applicable to your rule. For
            more info: https://docs.microsoft.com/en-us/graph/api/mailfolder-post-messagerules?view=graph-rest-1.0&tabs=http#request-body

        ### Returns
        ----
        Dict
            If successful, this method returns 201 Created response code and a
            `messageRule` object in the response body.
        """

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/mailFolders/inbox/messageRules',
            json=rule
        )

        return content

    def create_message_rule(self, user_id: str, rule: Dict) -> Dict:
        """Create a messageRule object by specifying a set of conditions and actions
        for the specified User.

        ### Parameters
        ----
        user_id : str
            The User ID for which to create the message rule
            for.

        rule : Dict
            The parameters that are applicable to your rule. For
            more info: https://docs.microsoft.com/en-us/graph/api/mailfolder-post-messagerules?view=graph-rest-1.0&tabs=http#request-body

        ### Returns
        ----
        Dict
            If successful, this method returns 201 Created response code and a
            `messageRule` object in the response body.
        """

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/mailFolders/inbox/messageRules'.format(
                user_id=user_id
            ),
            json=rule
        )

        return content

    def list_my_overrides(self) -> Dict:
        """Get the overrides that a user has set up to always classify messages from
        certain senders in specific ways.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response code and a collection
            of `inferenceClassificationOverride1 objects in the response body. An
            empty collection is returned if the user doesn't have any overrides
            set up.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/inferenceClassification/overrides'
        )

        return content

    def list_overrides(self, user_id: str) -> Dict:
        """Get the overrides that a user has set up to always classify messages from
        certain senders in specific ways.

        ### Parameters
        ----
        user_id : str
            The User ID for which to query `inferenceClassificationOverride`
            objects for.

        ### Returns
        ----
        Dict
            If successful, this method returns a 200 OK response code and a collection
            of `inferenceClassificationOverride` objects in the response body. An
            empty collection is returned if the user doesn't have any overrides
            set up.
        """

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/inferenceClassification/overrides'.format(
                user_id=user_id
            )
        )

        return content
