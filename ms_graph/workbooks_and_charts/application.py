from enum import Enum
from typing import Union
from ms_graph.session import GraphSession


class WorkbookApplication:

    """
    ## Overview:
    ----
    Represents the Excel application that manages
    the workbook.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `WorkbookApplication` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

    def get(self, item_id: str = None, item_path: str = None) -> dict:
        """Retrieve the properties and relationships of a workbookApplication
        object using the Item ID or Item Path.

        ### Parameters
        ----
        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A workbookApplication resource object.
        """

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/application",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/application",
            )

        return content

    def calculate(
        self,
        calculation_type: Union[str, Enum],
        item_id: str = None,
        item_path: str = None
    ) -> dict:
        """Recalculate all currently opened workbooks in Excel using the
        Item ID or Item Path.

        ### Parameters
        ----
        calculation_type :
            Specifies the calculation type to use. Possible
            values are: Recalculate, Full, FullRebuild.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A response object status code, 200 for success.
        """

        if isinstance(calculation_type, Enum):
            data = {"calculationType": calculation_type.value}
        else:
            data = {"calculationType": calculation_type}

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/application",
                json=data,
                additional_headers={"Content-type": "application/json"},
                expect_no_response=True,
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/application",
                json=data,
                additional_headers={"Content-type": "application/json"},
                expect_no_response=True,
            )

        return content
