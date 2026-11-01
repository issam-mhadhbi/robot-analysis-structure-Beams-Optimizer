"""
robin_grids.py
==============

Read grids from a Robin document into JSON, and write grids from JSON into a
Robin document, using the Interop.Robin COM API through pywin32.

Two classes live in this file:

    RobinGridReader  -- Robin  -> Python dict / JSON file
    RobinGridWriter  -- Python dict / JSON file -> Robin

Both use the same JSON schema (see SCHEMA below), so a grid exported with the
reader can be loaded back with the writer.

Robin object path used by both classes:

    Robin.Application
      -> Documents.ActiveDocument      (IRobinDocument)
         -> Grids                      (GridSystem)
            -> GridsCart               collection of GridCart
                 .XgridSteps / .YgridSteps     (GridSteps -> GridStep)
            -> GridsCyl                collection of GridCyl
                 .RadiusSteps / .AngleSteps    (GridSteps -> GridStep)
            -> GridAxisGroups          collection of GridAxisGroup
                 .GridAxisLines        (GridAxisLine: x1, y1, x2, y2)
                 .GridAxisArcs         (GridAxisArc: X, Y, Radius,
                                        StartAngle, DeltaAngle)

Assumptions that the type dump cannot confirm (check them on a small test):
    * GridStep.Value is the spacing and GridStep.Count is how many times that
      spacing repeats.
    * Lengths use the document units; angles are probably degrees.
    * Item(i) on a collection may be 0-based or 1-based; the reader tries both.

SCHEMA
------
{
  "clear_existing": true,                       # writer only: wipe grids first
  "cartesian": [
    {"name": "Grid_A", "x": 0, "y": 0, "angle": 0,
     "x_steps": [{"value": 5.0, "count": 4}],
     "y_steps": [{"value": 4.0, "count": 2}, {"value": 6.0, "count": 1}]}
  ],
  "cylindrical": [
    {"name": "Cyl_1", "x": 30, "y": 0, "angle": 0,
     "radius_steps": [{"value": 3.0, "count": 3}],
     "angle_steps":  [{"value": 30.0, "count": 12}]}
  ],
  "axis_groups": [
    {"name": "Free axes",
     "lines": [{"name": "1", "x1": 0, "y1": -1, "x2": 0, "y2": 15}],
     "arcs":  [{"name": "R1", "x": 0, "y": 0, "radius": 5,
                "start_angle": 0, "delta_angle": 90}]}
  ]
}

Quick start
-----------
    from robin_grids import RobinGridReader, RobinGridWriter

    RobinGridReader().save_file("grid_export.json")   # Robin -> JSON
    RobinGridWriter().load_file("grid_export.json")   # JSON  -> Robin
"""

import json

import win32com.client

# ---------------------------------------------------------------------------
# Shared helpers (module level, used by both classes)
# ---------------------------------------------------------------------------


def _connect(robin=None):
    """Connect to Robin and return ``(application, document, grid_system)``.

    Parameters
    ----------
    robin : COM object, optional
        An existing ``Robin.Application`` object. If omitted, a new one is
        created with ``Dispatch("Robin.Application")``.

    Returns
    -------
    tuple
        ``(app, doc, gs)`` where ``doc`` is the active document and ``gs`` is
        ``doc.Grids`` (the GridSystem).

    Raises
    ------
    RuntimeError
        If Robin has no active document.

    Example
    -------
    >>> app, doc, gs = _connect()
    >>> print(doc.Title, gs.GridsCart.Count)
    """
    app = robin or win32com.client.Dispatch("Robin.Application")
    app.Visible = True
    app.Activate()
    doc = app.Documents.ActiveDocument
    if doc is None:
        raise RuntimeError("No active Robin document.")
    return app, doc, doc.Grids


def _items(coll):
    """Return the items of a Robin collection as a Python list.

    The type dump does not say if ``Item(i)`` is 0-based or 1-based, so this
    tries 0-based first and falls back to 1-based.

    Parameters
    ----------
    coll : COM collection
        Any Robin collection exposing ``Count`` and ``Item(i)`` (GridsCart,
        GridSteps, GridAxisLines, ...).

    Returns
    -------
    list
        The COM items in order (empty list if the collection is empty).

    Example
    -------
    >>> for g in _items(gs.GridsCart):
    ...     print(g.Name)
    """
    n = coll.Count
    if n == 0:
        return []
    try:
        return [coll.Item(i) for i in range(n)]  # 0-based
    except Exception:
        return [coll.Item(i) for i in range(1, n + 1)]  # 1-based


# ---------------------------------------------------------------------------
# Reader: Robin -> dict / JSON
# ---------------------------------------------------------------------------


class RobinGridReader:
    """Read all grids from the active Robin document.

    Reads Cartesian grids, cylindrical grids and free axis groups (lines and
    arcs) and returns them as a plain dict following the module SCHEMA, or
    saves them to a JSON file.

    Example
    -------
    >>> reader = RobinGridReader()
    >>> data = reader.read_all()
    >>> reader.save_file("grid_export.json")
    """

    def __init__(self, robin=None):
        """Connect to Robin and keep references to the document and grids.

        Parameters
        ----------
        robin : COM object, optional
            An existing ``Robin.Application``; one is created if omitted.

        Example
        -------
        >>> reader = RobinGridReader()          # new connection
        >>> reader2 = RobinGridReader(reader.app)   # reuse a connection
        """
        self.app, self.doc, self.gs = _connect(robin)

    # ---- low level ---------------------------------------------------------

    @staticmethod
    def _read_steps(steps):
        """Convert a ``GridSteps`` collection to a list of dicts.

        Parameters
        ----------
        steps : GridSteps
            For example ``grid.XgridSteps`` or ``grid.RadiusSteps``.

        Returns
        -------
        list of dict
            ``[{"value": spacing, "count": repeats}, ...]``

        Example
        -------
        >>> RobinGridReader._read_steps(g.XgridSteps)
        [{'value': 5.0, 'count': 4}]
        """
        return [{"value": s.Value, "count": s.Count} for s in _items(steps)]

    # ---- per grid type -----------------------------------------------------

    def read_cartesian(self):
        """Read every Cartesian grid (``GridsCart``).

        Returns
        -------
        list of dict
            One dict per grid with keys ``name, x, y, angle, x_steps,
            y_steps``.

        Example
        -------
        >>> RobinGridReader().read_cartesian()
        [{'name': 'Grid_A', 'x': 0.0, 'y': 0.0, 'angle': 0.0,
          'x_steps': [{'value': 5.0, 'count': 4}], 'y_steps': [...]}]
        """
        out = []
        for g in _items(self.gs.GridsCart):
            out.append(
                {
                    "name": g.Name,
                    "x": g.X,
                    "y": g.Y,
                    "angle": g.Angle,
                    "x_steps": self._read_steps(g.XgridSteps),
                    "y_steps": self._read_steps(g.YgridSteps),
                }
            )
        return out

    def read_cylindrical(self):
        """Read every cylindrical grid (``GridsCyl``).

        Returns
        -------
        list of dict
            One dict per grid with keys ``name, x, y, angle, radius_steps,
            angle_steps``.

        Example
        -------
        >>> RobinGridReader().read_cylindrical()
        [{'name': 'Cyl_1', 'x': 30.0, 'y': 0.0, 'angle': 0.0,
          'radius_steps': [...], 'angle_steps': [...]}]
        """
        out = []
        for g in _items(self.gs.GridsCyl):
            out.append(
                {
                    "name": g.Name,
                    "x": g.X,
                    "y": g.Y,
                    "angle": g.Angle,
                    "radius_steps": self._read_steps(g.RadiusSteps),
                    "angle_steps": self._read_steps(g.AngleSteps),
                }
            )
        return out

    def read_axis_groups(self):
        """Read every free axis group (``GridAxisGroups``) with its lines and arcs.

        Returns
        -------
        list of dict
            One dict per group: ``{"name", "lines": [...], "arcs": [...]}``.
            Lines carry ``name, x1, y1, x2, y2``; arcs carry ``name, x, y,
            radius, start_angle, delta_angle``.

        Example
        -------
        >>> groups = RobinGridReader().read_axis_groups()
        >>> groups[0]["lines"][0]
        {'name': '1', 'x1': 0.0, 'y1': -1.0, 'x2': 0.0, 'y2': 15.0}
        """
        out = []
        for grp in _items(self.gs.GridAxisGroups):
            lines = [
                {"name": a.Name, "x1": a.x1, "y1": a.y1, "x2": a.x2, "y2": a.y2}
                for a in _items(grp.GridAxisLines)
            ]
            arcs = [
                {
                    "name": a.Name,
                    "x": a.X,
                    "y": a.Y,
                    "radius": a.Radius,
                    "start_angle": a.StartAngle,
                    "delta_angle": a.DeltaAngle,
                }
                for a in _items(grp.GridAxisArcs)
            ]
            out.append({"name": grp.Name, "lines": lines, "arcs": arcs})
        return out

    # ---- public ------------------------------------------------------------

    def read_all(self):
        """Read all grid data into one dict following the module SCHEMA.

        ``clear_existing`` is set to ``True`` so that loading the result with
        :class:`RobinGridWriter` replaces the target model's grids.

        Returns
        -------
        dict
            Keys: ``clear_existing, cartesian, cylindrical, axis_groups``.

        Example
        -------
        >>> data = RobinGridReader().read_all()
        >>> print(len(data["cartesian"]), "cartesian grid(s)")
        """
        return {
            "clear_existing": True,
            "cartesian": self.read_cartesian(),
            "cylindrical": self.read_cylindrical(),
            "axis_groups": self.read_axis_groups(),
        }

    def save_file(self, path):
        """Read all grids and save them to a JSON file.

        Parameters
        ----------
        path : str
            Destination file path (UTF-8, indented JSON).

        Returns
        -------
        dict
            The data that was written (same as :meth:`read_all`).

        Example
        -------
        >>> RobinGridReader().save_file(r"C:\\temp\\grid_export.json")
        """
        data = self.read_all()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return data

    def summary(self):
        """Return the number of grids of each kind currently in Robin.

        Returns
        -------
        dict
            ``{"cartesian": n, "cylindrical": n, "axis_groups": n}``

        Example
        -------
        >>> RobinGridReader().summary()
        {'cartesian': 1, 'cylindrical': 0, 'axis_groups': 1}
        """
        return {
            "cartesian": self.gs.GridsCart.Count,
            "cylindrical": self.gs.GridsCyl.Count,
            "axis_groups": self.gs.GridAxisGroups.Count,
        }


# ---------------------------------------------------------------------------
# Writer: dict / JSON -> Robin
# ---------------------------------------------------------------------------


class RobinGridWriter:
    """Create grids in the active Robin document from a dict or JSON file.

    Accepts the module SCHEMA. Missing optional keys default to 0 / empty.

    Example
    -------
    >>> writer = RobinGridWriter()
    >>> writer.load({"cartesian": [{"name": "A", "x_steps": [{"value": 5, "count": 3}],
    ...                             "y_steps": [{"value": 4, "count": 2}]}]})
    >>> writer.load_file("grid_export.json")
    """

    def __init__(self, robin=None):
        """Connect to Robin and keep references to the document and grids.

        Parameters
        ----------
        robin : COM object, optional
            An existing ``Robin.Application``; one is created if omitted.

        Example
        -------
        >>> writer = RobinGridWriter()
        """
        self.app, self.doc, self.gs = _connect(robin)

    # ---- low level ---------------------------------------------------------

    @staticmethod
    def _fill_steps(steps, spec):
        """Append steps to a ``GridSteps`` collection.

        Parameters
        ----------
        steps : GridSteps
            Target collection, e.g. ``grid.XgridSteps``.
        spec : list of dict
            ``[{"value": spacing, "count": repeats}, ...]``. ``count``
            defaults to 1.

        Example
        -------
        >>> RobinGridWriter._fill_steps(g.XgridSteps, [{"value": 5.0, "count": 4}])
        """
        for item in spec or []:
            s = steps.Add()
            s.Value = float(item["value"])
            s.Count = int(item.get("count", 1))

    # ---- per grid type -----------------------------------------------------

    def add_cartesian(self, d):
        """Add one Cartesian grid.

        Parameters
        ----------
        d : dict
            Keys: ``name`` (str), ``x``, ``y``, ``angle`` (float, default 0),
            ``x_steps`` and ``y_steps`` (lists of ``{"value", "count"}``).

        Returns
        -------
        COM object
            The created ``GridCart``.

        Example
        -------
        >>> RobinGridWriter().add_cartesian({
        ...     "name": "Grid_A", "x": 0, "y": 0, "angle": 0,
        ...     "x_steps": [{"value": 5.0, "count": 4}],
        ...     "y_steps": [{"value": 4.0, "count": 2}]})
        """
        g = self.gs.GridsCart.Add()
        g.Name = str(d.get("name", ""))
        g.X = float(d.get("x", 0.0))
        g.Y = float(d.get("y", 0.0))
        g.Angle = float(d.get("angle", 0.0))
        self._fill_steps(g.XgridSteps, d.get("x_steps"))
        self._fill_steps(g.YgridSteps, d.get("y_steps"))
        return g

    def add_cylindrical(self, d):
        """Add one cylindrical grid.

        Parameters
        ----------
        d : dict
            Keys: ``name``, ``x``, ``y``, ``angle`` (default 0),
            ``radius_steps`` and ``angle_steps`` (lists of
            ``{"value", "count"}``).

        Returns
        -------
        COM object
            The created ``GridCyl``.

        Example
        -------
        >>> RobinGridWriter().add_cylindrical({
        ...     "name": "Cyl_1", "x": 30, "y": 0,
        ...     "radius_steps": [{"value": 3.0, "count": 3}],
        ...     "angle_steps": [{"value": 30.0, "count": 12}]})
        """
        g = self.gs.GridsCyl.Add()
        g.Name = str(d.get("name", ""))
        g.X = float(d.get("x", 0.0))
        g.Y = float(d.get("y", 0.0))
        g.Angle = float(d.get("angle", 0.0))
        self._fill_steps(g.RadiusSteps, d.get("radius_steps"))
        self._fill_steps(g.AngleSteps, d.get("angle_steps"))
        return g

    def add_axis_group(self, d):
        """Add one free axis group with its lines and arcs.

        Parameters
        ----------
        d : dict
            Keys: ``name``; ``lines`` (list of dicts with ``name, x1, y1,
            x2, y2``); ``arcs`` (list of dicts with ``name, x, y, radius,
            start_angle, delta_angle``).

        Returns
        -------
        COM object
            The created ``GridAxisGroup``.

        Example
        -------
        >>> RobinGridWriter().add_axis_group({
        ...     "name": "Axes",
        ...     "lines": [{"name": "1", "x1": 0, "y1": 0, "x2": 0, "y2": 10}],
        ...     "arcs": [{"name": "R1", "x": 0, "y": 0, "radius": 5,
        ...               "start_angle": 0, "delta_angle": 90}]})
        """
        grp = self.gs.GridAxisGroups.Add()
        grp.Name = str(d.get("name", ""))
        for ln in d.get("lines", []):
            a = grp.GridAxisLines.Add()
            a.Name = str(ln["name"])
            a.x1, a.y1 = float(ln["x1"]), float(ln["y1"])
            a.x2, a.y2 = float(ln["x2"]), float(ln["y2"])
        for ar in d.get("arcs", []):
            a = grp.GridAxisArcs.Add()
            a.Name = str(ar["name"])
            a.X, a.Y = float(ar["x"]), float(ar["y"])
            a.Radius = float(ar["radius"])
            a.StartAngle = float(ar.get("start_angle", 0.0))
            a.DeltaAngle = float(ar["delta_angle"])
        return grp

    # ---- public ------------------------------------------------------------

    def clear(self):
        """Remove all Cartesian grids, cylindrical grids and axis groups.

        Example
        -------
        >>> RobinGridWriter().clear()
        """
        self.gs.GridsCart.RemoveAll()
        self.gs.GridsCyl.RemoveAll()
        self.gs.GridAxisGroups.RemoveAll()

    def show_all(self):
        """Make every grid and axis visible, then refresh and zoom the view.

        Called automatically at the end of :meth:`load`.

        Example
        -------
        >>> RobinGridWriter().show_all()
        """
        self.gs.ShowAll(True)
        self.gs.GridAxisGroups.ShowAllLineAxis(True)
        self.gs.GridAxisGroups.ShowAllArcAxis(True)
        self.doc.UpdateViews()
        self.doc.ViewsZoomAll()

    def load(self, data):
        """Create all grids described by a dict following the module SCHEMA.

        If ``data["clear_existing"]`` is true, existing grids are removed
        first; otherwise the new grids are appended.

        Parameters
        ----------
        data : dict
            Keys (all optional): ``clear_existing, cartesian, cylindrical,
            axis_groups``.

        Example
        -------
        >>> RobinGridWriter().load({
        ...     "clear_existing": True,
        ...     "cartesian": [{"name": "A",
        ...                    "x_steps": [{"value": 5, "count": 3}],
        ...                    "y_steps": [{"value": 4, "count": 2}]}]})
        """
        if data.get("clear_existing", False):
            self.clear()
        for d in data.get("cartesian", []):
            self.add_cartesian(d)
        for d in data.get("cylindrical", []):
            self.add_cylindrical(d)
        for d in data.get("axis_groups", []):
            self.add_axis_group(d)
        self.show_all()

    def load_file(self, path):
        """Read a JSON file and create the grids it describes.

        Parameters
        ----------
        path : str
            Path to a JSON file following the module SCHEMA.

        Example
        -------
        >>> RobinGridWriter().load_file("grid_export.json")
        """
        with open(path, "r", encoding="utf-8") as f:
            self.load(json.load(f))

    def summary(self):
        """Return the number of grids of each kind currently in Robin.

        Returns
        -------
        dict
            ``{"cartesian": n, "cylindrical": n, "axis_groups": n}``

        Example
        -------
        >>> RobinGridWriter().summary()
        {'cartesian': 1, 'cylindrical': 0, 'axis_groups': 1}
        """
        return {
            "cartesian": self.gs.GridsCart.Count,
            "cylindrical": self.gs.GridsCyl.Count,
            "axis_groups": self.gs.GridAxisGroups.Count,
        }


# ---------------------------------------------------------------------------
# Demo
# ---------------------------------------------------------------------------

EXAMPLE_DATA = {
    "clear_existing": True,
    "cartesian": [
        {
            "name": "Grid_A",
            "x": 0,
            "y": 0,
            "angle": 0,
            "x_steps": [{"value": 5.0, "count": 4}],
            "y_steps": [{"value": 4.0, "count": 2}, {"value": 6.0, "count": 1}],
        }
    ],
    "cylindrical": [],
    "axis_groups": [
        {
            "name": "Free axes",
            "lines": [
                {"name": "1", "x1": 0, "y1": -1, "x2": 0, "y2": 15},
                {"name": "A", "x1": -1, "y1": 0, "x2": 21, "y2": 0},
            ],
            "arcs": [],
        }
    ],
}


if __name__ == "__main__":
    # 1) Write the example grid into the active Robin document.
    writer = RobinGridWriter()
    writer.load(EXAMPLE_DATA)
    print("After write:", writer.summary())

    # 2) Read it back and save it to JSON.
    reader = RobinGridReader()
    reader.save_file("grid_export.json")
    print("After read: ", reader.summary(), "-> grid_export.json")

    # 3) Later (or in another model) load the exported file again.
    # RobinGridWriter().load_file("grid_export.json")
