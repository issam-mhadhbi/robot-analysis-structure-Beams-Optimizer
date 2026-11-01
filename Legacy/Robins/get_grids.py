import win32com.client

robin = win32com.client.Dispatch("Robin.Application")
robin.Visible = True
robin.Activate()

doc = robin.Documents.ActiveDocument
gs = doc.Grids                      # Robin.GridSystem

# ---------- 1) Cartesian grid from step spacings ----------
def add_steps(steps, spacing_list):
    """spacing_list = [(value, count), ...]  -> GridStep.Value / GridStep.Count"""
    for value, count in spacing_list:
        s = steps.Add()
        s.Value = float(value)
        s.Count = int(count)

gs.GridsCart.RemoveAll()
g = gs.GridsCart.Add()
g.Name = "Grid_A"
g.X = 0.0
g.Y = 0.0
g.Angle = 0.0
add_steps(g.XgridSteps, [(5.0, 4)])           # 4 bays of 5.0 along X
add_steps(g.YgridSteps, [(4.0, 2), (6.0, 1)]) # 2 bays of 4.0, then 1 of 6.0 along Y

# ---------- 2) Free axis lines (explicit coordinates) ----------
gs.GridAxisGroups.RemoveAll()
grp = gs.GridAxisGroups.Add()
grp.Name = "Axes"

def add_line(name, x1, y1, x2, y2):
    ln = grp.GridAxisLines.Add()
    ln.Name = name
    ln.x1, ln.y1, ln.x2, ln.y2 = x1, y1, x2, y2

for i in range(5):                            # vertical axes 1..5
    add_line(str(i + 1), i * 5.0, -1.0, i * 5.0, 15.0)
for j, letter in enumerate("ABC"):            # horizontal axes A..C
    add_line(letter, -1.0, j * 5.0, 21.0, j * 5.0)

# ---------- 3) Optional: cylindrical grid and an arc axis ----------
# c = gs.GridsCyl.Add(); c.Name = "Cyl1"; c.X = c.Y = 0.0; c.Angle = 0.0
# add_steps(c.RadiusSteps, [(3.0, 3)]); add_steps(c.AngleSteps, [(30.0, 12)])
# arc = grp.GridAxisArcs.Add()
# arc.Name = "R1"; arc.X = 0.0; arc.Y = 0.0
# arc.Radius = 5.0; arc.StartAngle = 0.0; arc.DeltaAngle = 90.0

# ---------- Make it visible ----------
gs.ShowAll(True)
gs.GridAxisGroups.ShowAllLineAxis(True)
gs.GridAxisGroups.ShowAllArcAxis(True)
doc.UpdateViews()
doc.ViewsZoomAll()

print("Cart grids:", gs.GridsCart.Count,
      "| axis groups:", gs.GridAxisGroups.Count,
      "| lines:", grp.GridAxisLines.Count)
