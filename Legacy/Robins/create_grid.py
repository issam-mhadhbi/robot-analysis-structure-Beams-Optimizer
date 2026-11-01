import json
import win32com.client


class RobinGridLoader:
    def __init__(self, robin=None):
        self.robin = robin or win32com.client.Dispatch("Robin.Application")
        self.robin.Visible = True
        self.robin.Activate()
        self.doc = self.robin.Documents.ActiveDocument
        if self.doc is None:
            raise RuntimeError("No active Robin document.")
        self.gs = self.doc.Grids          # Robin.GridSystem

    # ---------- helpers ----------
    @staticmethod
    def _fill_steps(steps, spec):
        """spec: list of {"value": float, "count": int}"""
        for item in spec or []:
            s = steps.Add()
            s.Value = float(item["value"])
            s.Count = int(item.get("count", 1))

    # ---------- builders ----------
    def _add_cartesian(self, d):
        g = self.gs.GridsCart.Add()
        g.Name = str(d.get("name", ""))
        g.X = float(d.get("x", 0.0))
        g.Y = float(d.get("y", 0.0))
        g.Angle = float(d.get("angle", 0.0))
        self._fill_steps(g.XgridSteps, d.get("x_steps"))
        self._fill_steps(g.YgridSteps, d.get("y_steps"))

    def _add_cylindrical(self, d):
        g = self.gs.GridsCyl.Add()
        g.Name = str(d.get("name", ""))
        g.X = float(d.get("x", 0.0))
        g.Y = float(d.get("y", 0.0))
        g.Angle = float(d.get("angle", 0.0))
        self._fill_steps(g.RadiusSteps, d.get("radius_steps"))
        self._fill_steps(g.AngleSteps, d.get("angle_steps"))

    def _add_axis_group(self, d):
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

    # ---------- public ----------
    def clear(self):
        self.gs.GridsCart.RemoveAll()
        self.gs.GridsCyl.RemoveAll()
        self.gs.GridAxisGroups.RemoveAll()

    def load(self, data):
        if data.get("clear_existing", False):
            self.clear()
        for d in data.get("cartesian", []):
            self._add_cartesian(d)
        for d in data.get("cylindrical", []):
            self._add_cylindrical(d)
        for d in data.get("axis_groups", []):
            self._add_axis_group(d)

        self.gs.ShowAll(True)
        self.gs.GridAxisGroups.ShowAllLineAxis(True)
        self.gs.GridAxisGroups.ShowAllArcAxis(True)
        self.doc.UpdateViews()
        self.doc.ViewsZoomAll()

    def load_file(self, path):
        with open(path, "r", encoding="utf-8") as f:
            self.load(json.load(f))

    def summary(self):
        return {
            "cartesian": self.gs.GridsCart.Count,
            "cylindrical": self.gs.GridsCyl.Count,
            "axis_groups": self.gs.GridAxisGroups.Count,
        }


grid_data = {
  "clear_existing": True,
  "cartesian": [
    {
      "name": "Grid_A", "x": 0, "y": 0, "angle": 0,
      "x_steps": [{"value": 5.0, "count": 4}],
      "y_steps": [{"value": 4.0, "count": 2}, {"value": 6.0, "count": 1}]
    }
  ],
  "cylindrical": [
    {
      "name": "Cyl_1", "x": 30, "y": 0, "angle": 0,
      "radius_steps": [{"value": 3.0, "count": 3}],
      "angle_steps": [{"value": 30.0, "count": 12}]
    }
  ],
  "axis_groups": [
    {
      "name": "Free axes",
      "lines": [
        {"name": "1", "x1": 0, "y1": -1, "x2": 0, "y2": 15},
        {"name": "A", "x1": -1, "y1": 0, "x2": 21, "y2": 0}
      ],
      "arcs": [
        {"name": "R1", "x": 0, "y": 0, "radius": 5, "start_angle": 0, "delta_angle": 90}
      ]
    }
  ]
}
if __name__ == "__main__":
    loader = RobinGridLoader()
    loader.load(grid_data)
    print(loader.summary())
