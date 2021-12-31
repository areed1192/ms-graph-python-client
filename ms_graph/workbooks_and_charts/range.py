from enum import Enum
from typing import Union
from ms_graph.session import GraphSession
from ms_graph.utils.range import RangeProperties


class Range:

    """
    ## Overview:
    ----
    Range represents a set of one or more contiguous cells
    such as a cell, a row, a column, block of cells, etc.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Range` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

    def get_range(
        self,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve the properties and relationships of range object.

        ### Parameters
        ----
        address : str (optional, Default=None)
            The range address.

        name : str (optional, Default=None)

        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        column_name_or_id : str (optional, Default=None)
            The name of the table column or the resource id.
            This must be specified if you are grabbing a table
            range.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A Range object.
        """

        if item_id:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')"
                )
            elif name:
                endpoint = f"/me/drive/items/{item_id}/workbook/names/{name}/range"
            else:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range"
                )

        elif item_path:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')"
                )
            elif name:
                endpoint = f"/me/drive/root:/{item_path}:/workbook/names/{name}/range"
            else:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range"
                )

        content = self.graph_session.make_request(method="get", endpoint=endpoint)

        return content

    def update_range(
        self,
        range_properties: Union[dict, RangeProperties],
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve the properties and relationships of range object.

        ### Parameters
        ----
        address : str (optional, Default=None)
            The range address.

        name : str (optional, Default=None)

        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        column_name_or_id : str (optional, Default=None)
            The name of the table column or the resource id.
            This must be specified if you are grabbing a table
            range.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A Range object.
        """

        if item_id:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')"
                )
            elif name:
                endpoint = f"/me/drive/items/{item_id}/workbook/names/{name}/range"
            else:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range"
                )

        elif item_path:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')"
                )
            elif name:
                endpoint = f"/me/drive/root:/{item_path}:/workbook/names/{name}/range"
            else:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range"
                )

        if isinstance(range_properties, RangeProperties):
            range_properties = range_properties.to_dict()

        content = self.graph_session.make_request(
            method="patch",
            json=range_properties,
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def insert_range(
        self,
        shift: Union[str, Enum],
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve the properties and relationships of range object.

        ### Parameters
        ----
        shift : Union[str, Enum]
            Specifies which way to shift the cells. The
            possible values are: Down, Right.

        address : str (optional, Default=None)
            The range address.

        name : str (optional, Default=None)

        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        column_name_or_id : str (optional, Default=None)
            The name of the table column or the resource id.
            This must be specified if you are grabbing a table
            range.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A Range object.
        """

        if item_id:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')/insert"
                )
            elif name:
                endpoint = f"/me/drive/items/{item_id}/workbook/names/{name}/range/insert"
            else:
                endpoint = (
                    f"/me/drive/items/{item_id}/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range/insert"
                )

        elif item_path:

            if worksheet_name_or_id:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"worksheets/{worksheet_name_or_id}/range(address='{address}')/insert"
                )
            elif name:
                endpoint = f"/me/drive/root:/{item_path}:/workbook/names/{name}/range/insert"
            else:
                endpoint = (
                    f"/me/drive/root:/{item_path}:/workbook/"
                    + f"tables/{table_name_or_id}/columns/{column_name_or_id}/range/insert"
                )

        if isinstance(shift, Enum):
            shift = shift.value

        content = self.graph_session.make_request(
            method="post",
            json={"shift":shift},
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content
