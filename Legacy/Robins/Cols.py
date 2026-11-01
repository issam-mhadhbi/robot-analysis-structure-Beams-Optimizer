"""
robin_columns_simple.py
Two simple examples: read the columns, and add one column.
"""

import pythoncom
import win32com.client
from win32com.client import VARIANT


def connect():
    app = win32com.client.Dispatch("Robin.Application")
    app.Visible = True
    app.Activate()
    return app.Documents.ActiveDocument.Structure2


def doubles(*values):
    """Typed array of doubles (what COM hosts usually expect)."""
    return VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values)
    )


# ---------------------------------------------------------------------------
# Example 1: read the columns
# ---------------------------------------------------------------------------
def read_columns(structure):
    columns = structure.Columns
    print("Number of columns:", columns.Count)

    for i in range(1, columns.Count + 1):  # Robin collections are 1-based
        col = columns.Item(i)
        print(
            col.UniqueNumber,
            col.Name,
            "storey",
            col.Storey,
            "section",
            col.Section,
            "material",
            col.Material,
            "volume",
            round(col.Volume, 4),
        )


# ---------------------------------------------------------------------------
# Example 2: add one column
# ---------------------------------------------------------------------------
def add_column(structure, x, y, z_bottom, z_top, storey, section, material, name):
    columns = structure.Columns

    # The dump does not say what Add() expects, so try a few simple formats.
    attempts = [
        ("doubles(x, y)", doubles(x, y)),
        ("doubles(x, y, z0, z1)", doubles(x, y, z_bottom, z_top)),
        ("tuple (x, y, z0, z1)", (x, y, z_bottom, z_top)),
    ]
    for label, argument in attempts:
        before = columns.Count
        try:
            col = columns.Add(argument)
        except Exception as exc:
            print("rejected:", label, exc.args[2][5] if len(exc.args) > 2 else exc)
            continue
        if columns.Count > before:
            print("accepted:", label)
            col.Storey = storey
            col.Section = section
            col.Material = material
            col.Name = name
            return col

    print("No format worked - run robin_add_probe.py and send me the output.")
    return None


if __name__ == "__main__":
    structure = connect()
    read_columns(structure)
    add_column(
        structure,
        x=0.0,
        y=0.0,
        z_bottom=0.0,
        z_top=3.0,
        storey=1,
        section="COL 40x40",
        material="C25_30",
        name="C_test",
    )
    read_columns(structure)
