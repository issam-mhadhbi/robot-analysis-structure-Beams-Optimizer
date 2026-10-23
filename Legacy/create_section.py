import win32com.client
import math

# Connect to Robot
robot = win32com.client.Dispatch("Robot.Application")
robot.Visible = True





def createRectSectionBeam(b , h ) : 
	labels = robot.Project.Structure.Labels #IRobotLabelServer
	section_name = f"hhhhhhhhhhhh{int(b * 100)}x{int(h * 100)}"
	section = labels.Create(3, section_name)  # I_LT_BAR_SECTION
	print(f"created section {section_name}")
	data = section.Data
	data.ShapeType = -1 #RobotBarSectionShapeType.I_BSST_CONCR_BEAM_RECT
	concrete = data.Concrete
	concrete.SetValue(1, b)  # Replace with I_BSCDV_BEAM_B
	concrete.SetValue(0, h)  # Replace with I_BSCDV_BEAM_H
	data.CalcNonstdGeometry()
	x = labels.Store(section)
	return section_name
sec = createRectSectionBeam(b = 0.22 , h = 0.45)
bar = robot.Project.Structure.Bars.Get(7)
print(dir(bar))
bar.SetLabel(
    win32com.client.constants.I_LT_BAR_SECTION,
    sec
)
print(bar.GetLabels())

