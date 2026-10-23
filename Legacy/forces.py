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


robot = win32com.client.Dispatch("Robot.Application")
robot.Visible = True
clr_dll = os.path.join(str(Path(robot.Kernel.ProgramPath).parent) , "Interop.RobotOM.dll" )
clr.AddReference(clr_dll)
import RobotOM 



import json
import win32com.client

robot = win32com.client.Dispatch("Robot.Application")
robot.Visible = True

structure = robot.Project.Structure


def get_all_bar_forces(n_points=5):
    """Return {case_number: {bar_number: [ {x, fx, fy, fz, mx, my, mz}, ... ]}}"""
    # Results only exist if the structure has been calculated
    if not structure.Results.Available:
        robot.Project.CalcEngine.Calculate()

    bar_server = structure.Results.Bars
    forces_server = bar_server.Forces

    # --- all bar numbers ---
    bars = structure.Bars.GetAll()
    bar_numbers = [bars.Get(i).Number for i in range(1, bars.Count + 1)]

    # --- all case numbers (simple cases + combinations) ---
    cases = structure.Cases.GetAll()
    case_numbers = []
    for i in range(1, cases.Count + 1):
        case = cases.Get(i)
        try:
            if case.IsAnalyzed:      # skip cases without results
                case_numbers.append(case.Number)
        except Exception:
            # some case types don't expose IsAnalyzed -> keep them
            case_numbers.append(case.Number)

    results = {}
    for case_no in case_numbers:
        results[case_no] = {}
        for bar_no in bar_numbers:
            pts = []
            for i in range(1, n_points + 1):
                # x = relative position 0.0 -> 1.0 along the bar
                x = (i - 1) / (n_points - 1) if n_points > 1 else 0.5
                d = forces_server.Value(bar_no, case_no, x)
                pts.append({
                    "x": x,
                    "fx": d.FX, "fy": d.FY, "fz": d.FZ,   # N
                    "mx": d.MX, "my": d.MY, "mz": d.MZ,   # N·m
                })
            results[case_no][bar_no] = pts
    return results


if __name__ == "__main__":
    data = get_all_bar_forces(n_points=5)
    with open("bar_forces.json", "w") as f:
        json.dump(data, f, indent=2)
    print(f"{len(data)} cases exported")