from enum import Enum


class CalculationTypes(Enum):
    """Specifies the calculation types used in the
    `WorkbookApplication` calculate method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import CalculationTypes
        >>> CalculationTypes.RECALCULATE.value
    """

    RECALCULATE = "Recaulcaute"
    FULL = "Full"
    FULLREBUILD = "FullRebuild"


class WorksheetVisibility(Enum):
    """Specifies the visibility types used in the
    `Worksheet` `update_worksheet` method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import WorksheetVisibility
        >>> WorksheetVisibility.VISIBLE.value
    """

    VISIBLE = "Visible"
    HIDDEN = "Hidden"
    VERYHIDDEN = "VeryHidden"


class RangeShift(Enum):
    """Specifies the shift directions used in the
    `Range` `insert_range` and `delete` method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import RangeShift
        >>> RangeShift.DOWN.value
    """

    # These are for inserts.
    DOWN = "Down"
    RIGHT = "Right"

    # These are for deletes.
    UP = "Up"
    LEFT = "Left"


class Underline(Enum):
    """Specifies the Underline property used in the
    `RangeFontProperties` object.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import Underline
        >>> RangeFontUnderline.SINGLE.value
    """

    NONE = "None"
    SINGLE = "Single"
    DOUBLE = "Double"
    SINGLE_ACCOUNTANT = "SingleAccountant"
    DOUBLE_ACCOUNTANT = "DoubleAccountant"


class VerticalAlignment(Enum):
    """Specifies the Vertical Alignment property used in the
    `RangeFormatProperties` object.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import VerticalAlignment
        >>> VerticalAlignment.TOP.value
    """

    TOP = "Top"
    CENTER = "Center"
    BOTTOM = "Bottom"
    JUSTIFY = "Justify"
    DISTRIBUTED = "Distributed"


class HorizontalAlignment(Enum):
    """Specifies the Horizontal Alignment property used in the
    `RangeFormatProperties` object.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import HorizontalAlignment
        >>> HorizontalAlignment.GENERAL.value
    """

    GENERAL = "General"
    LEFT = "Left"
    CENTER = "Center"
    RIGHT = "Right"
    FILL = "Fill"
    JUSTIFY = "Justify"
    CENTER_ACCROSS_SELECTION = "CenterAcrossSelection"
    DISTRIBUTED = "Distributed"


class BorderStyle(Enum):
    """Specifies the Border Style property used in the
    `RangeBorderProperties` object.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import BorderStyle
        >>> BorderStyle.CONTINUOUS.value
    """

    NONE = "None"
    CONTINUOUS = "Continuous"
    DASH = "Dash"
    DASH_DOT = "DashDot"
    DASH_DOT_DOT = "DashDotDot"
    DOT = "Dot"
    DOUBLE = "Double"
    SLANT_DASH_DOT = "SlantDashDot"


class BorderWeight(Enum):
    """Specifies the Border Weight property used in the
    `RangeBorderProperties` object.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import BorderWeight
        >>> BorderWeight.HAIRLINE.value
    """

    HAIRLINE = "Hairline"
    THIN = "Thin"
    MEDIUM = "Medium"
    THICK = "Thick"


class ApplyTo(Enum):
    """Specifies the type of Clear Action used in the `Range.clear()`
    method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import ApplyTo
        >>> ApplyTo.ALL.value
    """

    ALL = "All"
    FORMATS = "Formats"
    CONTENTS = "Contents"

class TableStyle(Enum):
    """Specifies the TableStyle property used in the `Table.update()`
    method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import TableStyle
        >>> TableStyle.TABLE_STYLE_LIGHT_1.value
    """

    TABLE_STYLE_LIGHT_1 = "TableStyleLight1"
    TABLE_STYLE_LIGHT_2 = "TableStyleLight2"
    TABLE_STYLE_LIGHT_3 = "TableStyleLight3"
    TABLE_STYLE_LIGHT_4 = "TableStyleLight4"
    TABLE_STYLE_LIGHT_5 = "TableStyleLight5"
    TABLE_STYLE_LIGHT_6 = "TableStyleLight6"
    TABLE_STYLE_LIGHT_7 = "TableStyleLight7"
    TABLE_STYLE_LIGHT_8 = "TableStyleLight8"
    TABLE_STYLE_LIGHT_9 = "TableStyleLight9"
    TABLE_STYLE_LIGHT_10 = "TableStyleLight10"
    TABLE_STYLE_LIGHT_11 = "TableStyleLight11"
    TABLE_STYLE_LIGHT_12 = "TableStyleLight12"
    TABLE_STYLE_LIGHT_13 = "TableStyleLight13"
    TABLE_STYLE_LIGHT_14 = "TableStyleLight14"
    TABLE_STYLE_LIGHT_15 = "TableStyleLight15"
    TABLE_STYLE_LIGHT_16 = "TableStyleLight16"
    TABLE_STYLE_LIGHT_17 = "TableStyleLight17"
    TABLE_STYLE_LIGHT_18 = "TableStyleLight18"
    TABLE_STYLE_LIGHT_19 = "TableStyleLight19"
    TABLE_STYLE_LIGHT_20 = "TableStyleLight20"
    TABLE_STYLE_LIGHT_21 = "TableStyleLight21"
    TABLE_STYLE_LIGHT_22 = "TableStyleLight22"
    TABLE_STYLE_LIGHT_23 = "TableStyleLight23"
    TABLE_STYLE_LIGHT_24 = "TableStyleLight24"
    TABLE_STYLE_LIGHT_25 = "TableStyleLight25"
    TABLE_STYLE_LIGHT_26 = "TableStyleLight26"
    TABLE_STYLE_LIGHT_27 = "TableStyleLight27"
    TABLE_STYLE_LIGHT_28 = "TableStyleLight28"

    TABLE_STYLE_MEDIUM_1 = "TableStyleMedium1"
    TABLE_STYLE_MEDIUM_2 = "TableStyleMedium2"
    TABLE_STYLE_MEDIUM_3 = "TableStyleMedium3"
    TABLE_STYLE_MEDIUM_4 = "TableStyleMedium4"
    TABLE_STYLE_MEDIUM_5 = "TableStyleMedium5"
    TABLE_STYLE_MEDIUM_6 = "TableStyleMedium6"
    TABLE_STYLE_MEDIUM_7 = "TableStyleMedium7"
    TABLE_STYLE_MEDIUM_8 = "TableStyleMedium8"
    TABLE_STYLE_MEDIUM_9 = "TableStyleMedium9"
    TABLE_STYLE_MEDIUM_10 = "TableStyleMedium10"
    TABLE_STYLE_MEDIUM_11 = "TableStyleMedium11"
    TABLE_STYLE_MEDIUM_12 = "TableStyleMedium12"
    TABLE_STYLE_MEDIUM_13 = "TableStyleMedium13"
    TABLE_STYLE_MEDIUM_14 = "TableStyleMedium14"
    TABLE_STYLE_MEDIUM_15 = "TableStyleMedium15"
    TABLE_STYLE_MEDIUM_16 = "TableStyleMedium16"
    TABLE_STYLE_MEDIUM_17 = "TableStyleMedium17"
    TABLE_STYLE_MEDIUM_18 = "TableStyleMedium18"
    TABLE_STYLE_MEDIUM_19 = "TableStyleMedium19"
    TABLE_STYLE_MEDIUM_20 = "TableStyleMedium20"
    TABLE_STYLE_MEDIUM_21 = "TableStyleMedium21"
    TABLE_STYLE_MEDIUM_22 = "TableStyleMedium22"
    TABLE_STYLE_MEDIUM_23 = "TableStyleMedium23"
    TABLE_STYLE_MEDIUM_24 = "TableStyleMedium24"
    TABLE_STYLE_MEDIUM_25 = "TableStyleMedium25"
    TABLE_STYLE_MEDIUM_26 = "TableStyleMedium26"
    TABLE_STYLE_MEDIUM_27 = "TableStyleMedium27"
    TABLE_STYLE_MEDIUM_28 = "TableStyleMedium28"

    TABLE_STYLE_DARK_1 = "TableStyleDark1"
    TABLE_STYLE_DARK_2 = "TableStyleDark2"
    TABLE_STYLE_DARK_3 = "TableStyleDark3"
    TABLE_STYLE_DARK_4 = "TableStyleDark4"
    TABLE_STYLE_DARK_5 = "TableStyleDark5"
    TABLE_STYLE_DARK_6 = "TableStyleDark6"
    TABLE_STYLE_DARK_7 = "TableStyleDark7"
    TABLE_STYLE_DARK_8 = "TableStyleDark8"
    TABLE_STYLE_DARK_9 = "TableStyleDark9"
    TABLE_STYLE_DARK_10 = "TableStyleDark10"
    TABLE_STYLE_DARK_11 = "TableStyleDark11"
    TABLE_STYLE_DARK_12 = "TableStyleDark12"
    TABLE_STYLE_DARK_13 = "TableStyleDark13"
    TABLE_STYLE_DARK_14 = "TableStyleDark14"
    TABLE_STYLE_DARK_15 = "TableStyleDark15"
    TABLE_STYLE_DARK_16 = "TableStyleDark16"
    TABLE_STYLE_DARK_17 = "TableStyleDark17"
    TABLE_STYLE_DARK_18 = "TableStyleDark18"
    TABLE_STYLE_DARK_19 = "TableStyleDark19"
    TABLE_STYLE_DARK_20 = "TableStyleDark20"
    TABLE_STYLE_DARK_21 = "TableStyleDark21"
    TABLE_STYLE_DARK_22 = "TableStyleDark22"
    TABLE_STYLE_DARK_23 = "TableStyleDark23"
    TABLE_STYLE_DARK_24 = "TableStyleDark24"
    TABLE_STYLE_DARK_25 = "TableStyleDark25"
    TABLE_STYLE_DARK_26 = "TableStyleDark26"
    TABLE_STYLE_DARK_27 = "TableStyleDark27"
    TABLE_STYLE_DARK_28 = "TableStyleDark28"
