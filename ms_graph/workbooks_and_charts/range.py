from enum import Enum
from typing import Union
from ms_graph.session import GraphSession
from ms_graph.utils.range import RangeProperties
from ms_graph.utils.range import RangeFormatProperties

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

class Range:

    """
    ### Overview:
    ----
    Range represents a set of one or more contiguous cells
    such as a cell, a row, a column, block of cells, etc.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Range` object.

        # Parameters
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

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)

        content = self.graph_session.make_request(
            method="get",
            endpoint=endpoint
        )

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

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)

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

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/insert"

        if isinstance(shift, Enum):
            shift = shift.value

        content = self.graph_session.make_request(
            method="post",
            json={"shift": shift},
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def get_range_format(
        self,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve the properties and relationships of rangeformat object.

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
            A WorkbookFormatRange object.
        """

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/format"

        content = self.graph_session.make_request(
            method="get", endpoint=endpoint)

        return content

    def update_range_format(
        self,
        range_format_properties: Union[dict, RangeProperties],
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
        range_format_properties : Union[dict, RangeProperties]
            Supply the values for relevant fields that should be
            updated. Existing properties that are not included in
            the request body will maintain their previous values
            or be recalculated based on changes to other property
            values. For best performance you shouldn't include
            existing values that haven't changed.

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

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/format"

        if isinstance(range_format_properties, RangeFormatProperties):
            range_format_properties = range_format_properties.to_dict()

        content = self.graph_session.make_request(
            method="patch",
            json=range_format_properties,
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def merge(
        self,
        across: bool = False,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Merge the range cells into one region in the worksheet.

        # Parameters
        ----
        across : bool (optional, Default=False)
            Set to `True` to merge cells in each row
            of the specified range as separate merged
            cells.

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
            A Response object.
        """

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/merge"

        body = {"across": across}

        content = self.graph_session.make_request(
            method="post",
            json=body,
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def unmerge(
        self,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Unmerge the range cells into seperate regions.

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
            A Response object.
        """

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/unmerge"

        content = self.graph_session.make_request(
            method="post",
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def clear(
        self,
        apply_to: Union[str, Enum] = None,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Clear range values such as format, fill, and border.

        ### Parameters
        ----
        apply_to : Union[str, Enum] (optional, Default=None)
            Determines the type of clear action. The possible
            values are: `All`, `Formats`, `Contents`.

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
            A Response object.
        """

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/clear"

        if isinstance(apply_to, Enum):
            apply_to = apply_to.value

        content = self.graph_session.make_request(
            method="post",
            json={"applyTo": apply_to},
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content

    def delete(
        self,
        shift: Union[str, Enum] = None,
        address: str = None,
        name: str = None,
        table_name_or_id: str = None,
        worksheet_name_or_id: str = None,
        column_name_or_id: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Deletes the cells associated with the range.

        ### Parameters
        ----
        shift : Union[str, Enum] (optional, Default=None)
            Specifies which way to shift the cells. The possible
            values are: `Up` and `Left`.

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
            A Response object.
        """

        inputs = {
            "address": address,
            "name": name,
            "table_name_or_id": table_name_or_id,
            "worksheet_name_or_id":worksheet_name_or_id,
            "column_name_or_id": column_name_or_id,
            "item_id": item_id,
            "item_path": item_path
        }

        endpoint = build_endpoint(inputs=inputs)
        endpoint = endpoint + "/delete"

        if isinstance(shift, Enum):
            shift = shift.value

        content = self.graph_session.make_request(
            method="post",
            json={"shift": shift},
            additional_headers={"Content-type": "application/json"},
            endpoint=endpoint,
        )

        return content
