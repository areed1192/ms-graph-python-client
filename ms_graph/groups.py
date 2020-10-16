import json
import requests

from typing import List
from typing import Dict
from typing import Union

from ms_graph.session import GraphSession

class Groups():


    def __init__(self, session: object ) -> None:

        # Set the session.
        self.graph_session: GraphSession = session

        # Set the endpoint.
        self.endpoint = 'group'
        self.collections_endpoint = 'groups'

    def list_groups(self) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint
        )

        return content

    def get_root_drive_children(self) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/children"
        )

        return content

    def get_root_drive_delta(self) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/delta"
        )

        return content

    def get_root_drive_followed(self) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.endpoint + "/root/followed"
        )

        return content


    def get_drive_by_id(self, drive_id: str) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint + "/{id}".format(id=drive_id)
        )

        return content

    def get_my_drives(self) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint=self.collections_endpoint + "/me"
        )

        return content

    def get_user_drives(self, user_id: str) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint="users/{user_id}/drives".format(user_id=user_id)
        )

        return content

    def get_group_drives(self, group_id: str) -> Dict:
        
        content = self.graph_session.make_request(
            method='get',
            endpoint="groups/{group_id}/drives".format(group_id=group_id)
        )

        return content