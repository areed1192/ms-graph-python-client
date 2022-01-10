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
