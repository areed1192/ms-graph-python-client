from enum import Enum
from typing import Union
from ms_graph.session import GraphSession


class Worksheet:

    """
    ## Overview:
    ----
    Represents an Excel Worksheet object. An Excel Worksheet object
    is a grid of cells. It can contain data, tables, charts, etc.
    """

    def __init__(self, session: object) -> None:
        """Initializes the `Worksheet` object.

        ### Parameters
        ----
        session : object
            An authenticated session for our Microsoft Graph Client.
        """

        # Set the session.
        self.graph_session: GraphSession = session

    def add_worksheet(
        self,
        item_id: str = None,
        item_path: str = None,
        name: str = None,
        workbook_session_id: str = None,
    ) -> dict:
        """Adds a new worksheet to the workbook. The worksheet will be added
        at the end of existing worksheets. If you wish to activate the newly
        added worksheet, call .activate() on it.

        ### Parameters
        ----
        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        name : str
            The name of the worksheet to be added. If specified,
            name should be unqiue. If not specified, Excel
            determines the name of the new worksheet.

        workbook_session_id : str (optional, Default=None)
            Workbook session Id that determines if changes are
            persisted or not.

        ### Returns
        ----
        dict:
            A WorkbookWorksheet object.
        """

        if workbook_session_id:
            additional_headers = {"Workbook-Session-Id": workbook_session_id}
        else:
            additional_headers = None

        if name:
            body = {"name": name}
        else:
            body = None

        if item_id:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers=additional_headers,
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets/add",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers=additional_headers,
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/add",
            )

        return content

    def get_worksheet(
        self, worksheet_id_or_name: str, item_id: str = None, item_path: str = None
    ) -> dict:
        """Retrieve the properties and relationships of worksheet object.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A WorkbookWorksheet object.
        """

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets/{worksheet_id_or_name}",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/{worksheet_id_or_name}",
            )

        return content

    def get_used_range(
        self,
        worksheet_id_or_name: str,
        item_id: str = None,
        item_path: str = None,
        values_only: bool = True,
    ) -> dict:
        """The used range is the smallest range that encompasses any cells
        that have a value or formatting assigned to them. If the worksheet
        is blank, this function will return the top left cell.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        values_only: bool (optional, Default=True)
            Considers only cells with values as used cells
            (ignores formatting).

        ### Returns
        ----
        dict:
            A Range object.
        """

        if values_only:
            params = {"valuesOnly": values_only}
        else:
            params = None

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                params=params,
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets/"
                + f"{worksheet_id_or_name}/usedRange",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                params=params,
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/usedRange",
            )

        return content

    def update_worksheet(
        self,
        worksheet_id_or_name: str,
        item_id: str = None,
        item_path: str = None,
        name: str = None,
        position: int = None,
        visibility: Union[str, Enum] = None,
        workbook_session_id: str = None,
    ) -> dict:
        """Update the properties of worksheet object.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        name : str (optional, Default=None)
            The display name of the worksheet.

        position : int (optional, Default=None)
            The zero-based position of the worksheet within the workbook.

        visibility : Union[str, Enum] (optional, Default=None)
            The Visibility of the worksheet. The possible values
            are: Visible, Hidden, VeryHidden.

        workbook_session_id : str (optional, Default=None)
            Workbook session Id that determines if changes are persisted
            or not.

        ### Returns
        ----
        dict:
            The updated WorkbookWorksheet object.
        """

        body = {}
        additional_headers = {"Content-type": "application/json"}

        if name:
            body["name"] = name

        if visibility:
            if isinstance(visibility, Enum):
                body["visibility"] = visibility.value
            else:
                body["visibility"] = visibility

        if position:
            body["position"] = position

        if body == {}:
            return {
                "message": "No properties were requested to be updated, no request made."
            }

        if workbook_session_id:
            additional_headers["Workbook-Session-Id"] = workbook_session_id

        if item_id:
            content = self.graph_session.make_request(
                method="patch",
                json=body,
                additional_headers=additional_headers,
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets/{worksheet_id_or_name}",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="patch",
                json=body,
                additional_headers=additional_headers,
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/{worksheet_id_or_name}",
            )

        return content

    def delete_worksheet(
        self,
        worksheet_id_or_name: str,
        item_id: str = None,
        item_path: str = None,
        workbook_session_id: str = None,
    ) -> dict:
        """Deletes the worksheet from the workbook.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        workbook_session_id : str (optional, Default=None)
            Workbook session Id that determines if changes are persisted
            or not.

        ### Returns
        ----
        dict:
            The response status code, a 200 means success.
        """

        additional_headers = {}

        if workbook_session_id:
            additional_headers["Workbook-Session-Id"] = workbook_session_id

        if item_id:
            content = self.graph_session.make_request(
                method="delete",
                additional_headers=additional_headers,
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets/{worksheet_id_or_name}",
                expect_no_response=True,
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="delete",
                additional_headers=additional_headers,
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/{worksheet_id_or_name}",
                expect_no_response=True,
            )

        return content

    def get_cell(
        self,
        worksheet_id_or_name: str,
        row: int,
        column: int,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Gets the range object containing the single cell based on row and column
        numbers. The cell can be outside the bounds of its parent range, so long
        as it's stays within the worksheet grid.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        row : int
            Row number of the cell to be retrieved. Zero-indexed.

        column : int
            Column number of the cell to be retrieved. Zero-indexed.

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
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}/cell(row={row},column={column})",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/cell(row={row},column={column})",
            )

        return content

    def get_range(
        self,
        worksheet_id_or_name: str,
        address: str = None,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Gets the range object specified by the address or name.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        address : str (optional, Default=None)
            The address or the name of the range. If not specified,
            the entire worksheet range is returned.

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

        if address:
            endpoint = r"/range(address={address})"
        else:
            endpoint = "/range"

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}"
                + endpoint,
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}"
                + endpoint,
            )

        return content

    def list_tables(
        self,
        worksheet_id_or_name: str,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve a list of WorkbookTable objects.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A collection of WorkbookTable Objects.
        """

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}/tables",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/tables",
            )

        return content

    def add_table(
        self,
        worksheet_id_or_name: str,
        address: str,
        has_headers: bool,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Creates a new WorkbookTable Object.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        address : str
            The range address.

        has_headers : bool
            Boolean value that indicates whether the range has
            column labels. If the source does not contain headers
            (i.e,. when this property set to false), Excel will
            automatically generate header shifting the data down
            by one row.

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

        body = {"address": address, "hasHeaders": has_headers}

        if item_id:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers={"Content-type": "application/json"},
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}/tables/add",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers={"Content-type": "application/json"},
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/tables/add",
            )

        return content

    def list_charts(
        self,
        worksheet_id_or_name: str,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve a list of WorkbookChart Objects.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A collection of WorkbookChart Objects.
        """

        if item_id:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}/charts",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="get",
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/charts",
            )

        return content

    def add_chart(
        self,
        worksheet_id_or_name: str,
        name: str,
        height: float,
        top: float,
        left: float,
        width: float,
        item_id: str = None,
        item_path: str = None,
    ) -> dict:
        """Retrieve a list of table objects.

        ### Parameters
        ----
        worksheet_id_or_name : str
            The worksheet resource id or the worksheet name.

        height : float
            Represents the height, in points, of the chart object.

        top : float
            Represents the distance, in points, from the top edge
            of the object to the top of row 1 (on a worksheet) or
            the top of the chart area (on a chart).

        left : float
            The distance, in points, from the left side of the chart
            to the worksheet origin.

        width : float
            Represents the width, in points, of the chart object.

        name : str
            Represents the name of a chart object.

        item_id : str (optional, Default=None)
            The Drive Item Resource ID.

        item_path : str (optional, Default=None)
            The Item Path. An Example would be the following:
            `/TestFolder/TestFile.txt`

        ### Returns
        ----
        dict:
            A WorkbookChart object.
        """

        body = {
            "height": height,
            "top": top,
            "left": left,
            "width": width,
            "name": name,
        }

        if item_id:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers={"Content-type": "application/json"},
                endpoint=f"/me/drive/items/{item_id}/workbook/worksheets"
                + f"/{worksheet_id_or_name}/charts",
            )
        elif item_path:
            content = self.graph_session.make_request(
                method="post",
                json=body,
                additional_headers={"Content-type": "application/json"},
                endpoint=f"/me/drive/root:/{item_path}:/workbook/worksheets/"
                + f"{worksheet_id_or_name}/charts",
            )

        return content
