from dataclasses import dataclass


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
        dict
            The Field Name and Values.
        """

        class_dict = {
            "columnHidden": self.column_hidden,
            "rowHidden": self.row_hidden,
            "formulas": self.formulas,
            "numberFormat": self.number_format,
            "formulasR1C1": self.formulas_r1c1,
            "formulasLocal": self.formulas_local,
            "values": self.values,
        }

        return class_dict
