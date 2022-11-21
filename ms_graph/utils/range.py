from enum import Enum
from typing import Union
from dataclasses import fields
from dataclasses import dataclass
from dataclasses import is_dataclass


def _to_dict(data_class_obj: Union[dataclass, dict]) -> dict:
    """Converts a `dataclass` object to a normal python `dict`.

    ### Parameter
    ----
    data_class_obj : Union[dataclass, dict]
        The python `dataclass` object or a normal
        python `dict` that may contain `dataclass`
        objects.

    ### Returns
    ----
    dict :
        A python dict that can be sent to the Microsoft
        Graph API.
    """

    class_dict = {}

    if is_dataclass(data_class_obj):

        # Loop through each field and grab the value and key.
        for field in fields(data_class_obj):

            key = field.name
            value = getattr(data_class_obj, field.name)

            # Handle values that could be Enums.
            if isinstance(value, Enum):
                value = value.value

            if isinstance(value, dict):
                value = _to_dict(data_class_obj=value)

            if isinstance(value, list):
                value = [_to_dict(data_class_obj=item) for item in value]

            if value is not None:
                class_dict[key] = value

    elif isinstance(data_class_obj, dict):

        for key, value in data_class_obj.items():

            # Handle values that could be Enums.
            if isinstance(value, Enum):
                value = value.value

            if isinstance(value, dict):
                value = _to_dict(data_class_obj=value)

            if isinstance(value, list):
                value = [_to_dict(data_class_obj=item) for item in value]

            if value is not None:
                class_dict[key] = value

    return class_dict


@dataclass
class RangeProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Properties.
    The Microsoft Graph API allows users to update Range objects and
    this utility makes constructing those updates in a concise way that
    is python friendly.

    ### Parameters
    ----
    column_hidden : bool (optional, Default=None)
        Represents if all columns of the current range are hidden.

    formulas : list (optional, Default=None)
        Represents the formula in A1-style notation.

    formulas_local : list (optional, Default=None)
        Represents the formula in A1-style notation, in the user's
        language and number-formatting locale. For example, the
        English "=SUM(A1, 1.5)" formula would become
        "=SUMME(A1; 1,5)" in German.

    formulas_r1c1 : list (optional, Default=None)
        Represents the formula in R1C1-style notation.

    number_format : str (optional, Default=None)
        Represents Excel's number format code for the given cell.

    row_hidden : bool (optional, Default=None)
        Represents if all rows of the current range are hidden.

    values : list (optional, Default=None)
        Represents the raw values of the specified range. The
        data returned could be of type string, number, or a
        boolean. Cell that contain an error will return the
        error string.
    """

    column_hidden: bool
    row_hidden: bool
    formulas: list
    formulas_local: list
    formulas_r1c1: list
    number_format: str
    values: list

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)


@dataclass
class RangeFormatProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Format
    Properties. A format object encapsulating the range's font,
    fill, borders, alignment, and other properties.

    ### Parameters
    ----
    column_width : float (optional, Default=None)
        Sets the width of all columns within the range. If the
        column widths are not uniform, null will be returned.

    horizontal_alignment : Union[str, Enum] (optional, Default='General')
        Represents the horizontal alignment for the specified
        object. The possible values are: `General`, `Left`, `Center`,
        `Right`, `Fill`, `Justify`, `CenterAcrossSelection`,
        and `Distributed`.

    row_height : float (optional, Default=None)
        Sets the height of all rows in the range. If the
        row heights are not uniform, null will be returned.

    vertical_alignment : Union[str, Enum] (optional, Default='General')
        Represents the vertical alignment for the specified
        object. The possible values are: `Top`, `Center`, `Bottom`,
        `Justify`, and `Distributed`.

    wrap_text : bool (optional, Default=False)
        Indicates if Excel wraps the text in the object.
        A null value indicates that the entire range doesn't
        have uniform wrap setting
    """

    column_width: float = None
    horizontal_alignment: Union[str, Enum] = "General"
    row_height: float = None
    vertical_alignment: Union[str, Enum] = "General"
    wrap_text: bool = False

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)

@dataclass
class RangeFillProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Fill
    Properties. Represents the background of a range object.

    ### Parameters
    ----
    color : str (optional, Default=None)
        HTML color code representing the color of the border
        line, of the form #RRGGBB (e.g. "FFA500") or as a
        named HTML color (e.g. "orange").
    """

    column_width: str = None

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)

@dataclass
class RangeFontProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Font
    Properties. This object represents the font attributes
    (font name, font size, color, etc.) for an object.

    ### Parameters
    ----
    bold : bool (optional, Default=False)
        Represents the bold status of font. If set to `True`
        font will be bold, `False` it will not.

    color : str (optional, Default=None)
        HTML color code representation of the text color.
        E.g. #FF0000 represents Red.

    italic : bool (optional, Default=False)
        Represents the italic status of font. If set to `True`
        font will be italic, `False` it will not.

    name : str (optional, Default=None)
        The font name. For example, `Calibri`.

    size : float (optional, Default=None)
        Sets the size of the font.

    underline : Union[str, Enum] (optional, Default='None')
        Type of underline applied to the font. The possible
        values are: `None`, `Single`, `Double`,
        `SingleAccountant`, `DoubleAccountant`.
    """

    bold: bool = False
    color: str = None
    italic: bool = False
    name: str = None
    size: float = None
    underline: Union[str, Enum] = "None"

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)

@dataclass
class RangeBorderProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Border
    Properties. Represents the border of an object.

    ### Parameters
    ----
    color : str (optional, Default=None)
        HTML color code representing the color of the border
        line, of the form #RRGGBB (e.g. "FFA500") or as a
        named HTML color (e.g. "orange"). the text color.

    style : Union[str, Enum] (optional, Default="None")
        One of the constants of line style specifying the
        line style for the border. The possible values
        are: `None`, `Continuous`, `Dash`, `DashDot`,
        `DashDotDot`, `Dot`, `Double`, `SlantDashDot`.

    weight : Union[str, Enum] (optional, Default=None)
        Specifies the weight of the border around a
        range. The possible values are: `Hairline`,
        `Thin`, `Medium`, and `Thick`.
    """

    color: str = None
    style: Union[str, Enum] = "None"
    weight: Union[str, Enum] = None

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)

@dataclass
class RangeFormatProtectionProperties:

    """
    ### Overview
    ----
    A python dataclass which is used to represent Range Format
    Protection Properties. Represents the format protection
    of a range object.

    ### Parameters
    ----
    formula_hidden : bool (optional, Default=False)
        Indicates if Excel hides the formula for the cells
        in the range. A null value indicates that the entire
        range doesn't have uniform formula hidden setting.

    locked : bool (optional, Default=False)
        Indicates if Excel locks the cells in the object.
        A null value indicates that the entire range
        doesn't have uniform lock setting.
    """

    formula_hidden: bool = False
    locked: bool = False

    def to_dict(self) -> dict:
        """Generates a dictionary containing all the field
        names and values.

        ### Returns
        ----
        dict :
            A dictionary object where the Fieldnames
            are the keys and Field values are the
            values.
        """

        return _to_dict(data_class_obj=self)
