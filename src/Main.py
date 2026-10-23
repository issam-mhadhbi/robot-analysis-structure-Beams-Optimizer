import json

import win32com.client
from win32com.client import CastTo
from BeamsOpt import BeamsOptPipeline

if __name__ == "__main__":
    robot = win32com.client.Dispatch("Robot.Application")
    robot.Visible = True
    print(dir(robot.Kernel.ProgramPath))
    print(robot.Kernel.ProgramPath)
    # locator = GridPipeline(robot)
    # ColumnsPipeline(robot , locator )
    # BeamsPipeline(robot, locator)
    # SlabsPipeline(robot, locator)