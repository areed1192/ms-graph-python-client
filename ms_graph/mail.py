from typing import Dict
from ms_graph.session import GraphSession


class Mail():

    def __init__(self, session: object) -> None:

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'mail'

    def list_my_messages(self) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages'
        )

        return content

    def list_user_messages(self, user_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id)
        )

        return content

    def create_my_message(self, message: dict) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages',
            json=message
        )

        return content

    def create_user_message(self, user_id: str, message: dict) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id),
            json=message
        )

        return content

    def get_my_messages(self, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages/{message_id}'.format(
                message_id=message_id
            )
        )

        return content

    def get_user_messages(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages/{message_id}'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def update_my_message(self, message: dict) -> dict:

        content = self.graph_session.make_request(
            method='patch',
            endpoint='/me/messages',
            json=message
        )

        return content

    def update_user_message(self, user_id: str, message: dict) -> Dict:

        content = self.graph_session.make_request(
            method='patch',
            endpoint='/users/{user_id}/messages'.format(user_id=user_id),
            json=message
        )

        return content

    def delete_my_message(self, message_id: str) -> dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages/{message_id}'.format(
                message_id=message_id
            )
        )

        return content

    def delete_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='delete',
            endpoint='/users/{user_id}/messages/{message_id}'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def send_my_message(self, message_id: str) -> dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/me/messages/{message_id}/send'.format(
                message_id=message_id
            )
        )

        return content

    def send_user_message(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='post',
            endpoint='/users/{user_id}/messages/{message_id}/send'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content

    def copy_my_message(self, message_id: str) -> dict:

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

    def send_user_mail(self, user_id: str, message: dict, save_to_send_items: bool) -> Dict:

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

        content = self.graph_session.make_request(
            method='get',
            endpoint='/me/messages/{message_id}/attachments'.format(
                message_id=message_id
            )
        )

        return content

    def list_user_attachements(self, user_id: str, message_id: str) -> Dict:

        content = self.graph_session.make_request(
            method='get',
            endpoint='/users/{user_id}/messages/{message_id}/attachments'.format(
                user_id=user_id,
                message_id=message_id
            )
        )

        return content
