from enum import Enum
from typing import Union
from ms_graph.session import GraphSession


def build_endpoint(inputs: dict) -> str:
    """Builds the endpoint for the Range object.

    ### Parameters
    ----
    inputs : dict
        The `locals()` of the function.

    ### Raises
    ----
    ValueError:
        If an item id or item path is not specified
        an error will be raised.

    ### Returns
    ----
    str:
        The full URL path.
    """

    item_id = inputs.get("item_id", None)
    item_path = inputs.get("item_path", None)
    worksheet_name_or_id = inputs.get("worksheet_name_or_id", None)
    address = inputs.get("address", None)
    name = inputs.get("name", None)
    table_name_or_id = inputs.get("table_name_or_id", None)
    column_name_or_id = inputs.get("column_name_or_id", None)

    if item_id:
        workbook_path = f"/me/drive/items/{item_id}/workbook/"
    elif item_path:
        workbook_path = f"/me/drive/root:/{item_path}:/workbook/"
    else:
        raise ValueError("Must specify an Item ID or Item Path.")

    if (worksheet_name_or_id and address):
        range_path = f"worksheets/{worksheet_name_or_id}/range(address='{address}')"
    elif name:
        range_path = f"names/{name}/range"
    elif (table_name_or_id and column_name_or_id):
        range_path = f"tables/{table_name_or_id}/columns/{column_name_or_id}/range"

    return workbook_path + range_path


class Table:

    """
    ### Overview:
    ----
    Represents an Excel Table Object, also called an
    Excel List Object.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Table` object.

        # Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

    def get_table(
        self,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve the properties and relationships of table object.

        ### Parameters
        ----
        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A WorkbookTable object.
        """

        if item_id:
            workbook_path = f"/me/drive/items/{item_id}/workbook/"
        elif item_path:
            workbook_path = f"/me/drive/root:/{item_path}:/workbook/"
        else:
            raise ValueError("Must specify an Item ID or Item Path.")

        if (worksheet_name_or_id and table_name_or_id):
            table_path = f"worksheets/{worksheet_name_or_id}/tables/{table_name_or_id}"
        elif table_name_or_id:
            table_path = f"tables/{table_name_or_id}"

        endpoint = workbook_path + table_path

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

        return content

    def add_table(
        self,
        address: str,
        has_headers: bool,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Create a new WorkbookTable Object.

        ### Overview
        ----
        The range source address determines the worksheet under
        which the table will be added. If the table cannot be
        added (e.g., because the address is invalid, or the table
        would overlap with another table), an error will be thrown.

        ### Parameters
        ----
        address : str
            Address or name of the range object representing the
            data source. If the address does not contain a sheet
            name, the currently-active sheet is used.

        has_headers : bool
            Boolean value that indicates whether the data being
            imported has column labels. If the source does not
            contain headers (i.e,. when this property set to
            `False`), Excel will automatically generate header
            shifting the data down by one row.

        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A WorkbookTable Object.
        """

        if item_id:
            workbook_path = f"/me/drive/items/{item_id}/workbook/"
        elif item_path:
            workbook_path = f"/me/drive/root:/{item_path}:/workbook/"
        else:
            raise ValueError("Must specify an Item ID or Item Path.")

        if (worksheet_name_or_id and table_name_or_id):
            table_path = f"worksheets/{worksheet_name_or_id}/tables/{table_name_or_id}"
        elif table_name_or_id:
            table_path = f"tables/{table_name_or_id}"

        body = {"address": address, "hasHeaders": has_headers}
        endpoint = workbook_path + table_path + "/add"

        content = self.graph_session.make_request(
            method="post",
            json=body,
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint
        )

        return content

    def update_table(
        self,
        name: str = None,
        show_headers: bool = None,
        show_totals: bool = None,
        style: Union[str, Enum] = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Update the properties of table object.

        ### Parameters
        ----
        name : str (optional, Default=None)
            The name of the table.

        show_headers : bool (optional, Default=None)
            Indicates whether the header row is visible or not. This
            value can be set to show or remove the header row.

        show_totals : bool (optional, Default=None)
            Indicates whether the total row is visible or not. This
            value can be set to show or remove the total row.

        style : Union[str, Enum] (optional, Default=None)
            Constant value that represents the Table style. The possible
            values are: `TableStyleLight1` through `TableStyleLight21`,
            `TableStyleMedium1` through `TableStyleMedium28`, `TableStyleDark1`
            through `TableStyleDark11`. A custom user-defined style
            present in the workbook can also be specified.

        table_name_or_id : str (optional, Default=None)
            The name of the table or the resource id.

        worksheet_name_or_id : str (optional, Default=None)
            The name of the worksheet or the resource id.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A WorkbookTable object.
        """

        if item_id:
            workbook_path = f"/me/drive/items/{item_id}/workbook/"
        elif item_path:
            workbook_path = f"/me/drive/root:/{item_path}:/workbook/"
        else:
            raise ValueError("Must specify an Item ID or Item Path.")

        if (worksheet_name_or_id and table_name_or_id):
            table_path = f"worksheets/{worksheet_name_or_id}/tables/{table_name_or_id}"
        elif table_name_or_id:
            table_path = f"tables/{table_name_or_id}"

        endpoint = workbook_path + table_path

        if isinstance(style, Enum):
            style = style.value

        body = {
            "name": name,
            "showHeaders": show_headers,
            "showTotals": show_totals,
            "style": style
        }

        content = self.graph_session.make_request(
            method="patch",
            json=body,
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint
        )

        return content
