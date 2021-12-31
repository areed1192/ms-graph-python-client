from enum import Enum

class CalculationTypes(Enum):
    """Specifies the calculation types used in the
    `WorkbookApplication` calculate method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import CalculationTypes
        >>> CalculationTypes.RECALCULATE.value
    """

    RECALCULATE = 'Recaulcaute'
    FULL = 'Full'
    FULLREBUILD = 'FullRebuild'


class WorksheetVisibility(Enum):
    """Specifies the visibility types used in the
    `Worksheet` `update_worksheet` method.

    ### Usage:
    ----
        >>> from ms_graph.workbooks_and_charts.enums import WorksheetVisibility
        >>> WorksheetVisibility.VISIBLE.value
    """

    VISIBLE = 'Visible'
    HIDDEN = 'Hidden'
    VERYHIDDEN = 'VeryHidden'
