import json , os 
import math
from enum import IntEnum
import pythoncom
import win32com.client
from win32com.client import VARIANT, CastTo
from enum import IntEnum
from pathlib import Path 
import clr 
import importlib


robot = win32com.client.Dispatch("Robin.Application")
robot.Visible = True
clr_dll = os.path.join(str(Path(robot.Kernel.ProgramPath).parent) , "Interop.RobotOM.dll" )
clr.AddReference(clr_dll)
# # how to clr what exposed as modules after adding reference 
# RobotOM = importlib.import_module("RobotOM")

# Or list everything, including namespaces, via reflection
from System.Reflection import Assembly
asm = Assembly.LoadFrom(clr_dll)
for t in asm.GetExportedTypes():
    print(t.Namespace, t.Name)