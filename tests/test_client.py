import unittest

from unittest import TestCase
from configparser import ConfigParser
from ms_graph.client import MicrosoftGraphClient

from ms_graph.mail import Mail
from ms_graph.notes import Notes
from ms_graph.users import Users
from ms_graph.search import Search
from ms_graph.drives import Drives
from ms_graph.groups import Groups
from ms_graph.drive_items import DriveItems
from ms_graph.personal_contacts import PersonalContacts


class MicrosoftGraphSessionTest(TestCase):

    """Will perform a unit test for the `MicrosoftGraphClient` session."""

    def setUp(self) -> None:
        """Set up the `MicrosoftGraphClient` Client."""

        scopes = [
            'Calendars.ReadWrite',
            'Files.ReadWrite.All',
            'User.ReadWrite.All',
            'Notes.ReadWrite.All',
            'Directory.ReadWrite.All',
            'User.Read.All',
            'Directory.Read.All',
            'Directory.ReadWrite.All'
        ]

        # Initialize the Parser.
        config = ConfigParser()

        # Read the file.
        config.read('config/config.ini')

        # Get the specified credentials.
        client_id = config.get('graph_api', 'client_id')
        client_secret = config.get('graph_api', 'client_secret')
        redirect_uri = config.get('graph_api', 'redirect_uri')

        # Initialize the Client.
        graph_client = MicrosoftGraphClient(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope=scopes,
            credentials='config/ms_graph_state.jsonc'
        )

        self.client = graph_client

    def test_creates_instance_of_session(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient`."""

        self.assertIsInstance(self.client, MicrosoftGraphClient)

    def test_creates_instance_of_mail(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Mail`."""

        self.assertIsInstance(self.client.mail(), Mail)

    def test_creates_instance_of_drive_items(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.DriveItems`."""

        self.assertIsInstance(self.client.drive_item(), DriveItems)

    def test_creates_instance_of_drives(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Drives`."""

        self.assertIsInstance(self.client.drives(), Drives)

    def test_creates_instance_of_users(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Users`."""

        self.assertIsInstance(self.client.users(), Users)

    def test_creates_instance_of_groups(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Groups`."""

        self.assertIsInstance(self.client.groups(), Groups)

    def test_creates_instance_of_notes(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Notes`."""

        self.assertIsInstance(self.client.notes(), Notes)

    def test_creates_instance_of_search(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.Search`."""

        self.assertIsInstance(self.client.search(), Search)

    def test_creates_instance_of_personal_contacts(self):
        """Create an instance and make sure it's a `MicrosoftGraphClient.PersonalContacts`."""

        self.assertIsInstance(
            self.client.personal_contacts(), PersonalContacts)

    def tearDown(self) -> None:
        """Teardown the `MicrosoftGraphClient` Client."""

        del self.client


if __name__ == '__main__':
    unittest.main()
