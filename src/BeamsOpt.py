import json
import math
from enum import IntEnum
import pythoncom
import win32com.client
from win32com.client import VARIANT, CastTo
from enum import IntEnum
from pathlib import Path 
import clr 
import importlib

class IRobotLabelType(IntEnum):
    I_LT_BAR_SECTION = 3  # Label describing a bar section.


class RobotBarSectionShapeType(IntEnum):
    I_BSST_CONCR_COL_R = -108  #vailable since version 2.5.
    I_BSST_CONCR_COL_T = -107 #Available since version 2.5.
    I_BSST_CONCR_COL_L = -106 #Available since version 2.5.
    I_BSST_CONCR_COL_Z = -105 #Available since version 2.5.
    I_BSST_CONCR_COL_P = -104 #Available since version 2.5.
    I_BSST_CONCR_COL_C = -103 #Available since version 2.5.
    I_BSST_CONCR_COL_CH = -102 #Available since version 2.5.
    I_BSST_CONCR_COL_CQ = -101 #Available since version 2.5.
    I_BSST_CONCR_BEAM_RECT = -3 #RC rectangular beam .
    I_BSST_CONCR_BEAM_T = -2 #Available since version 2.5.
    I_BSST_CONCR_BEAM = -1
    I_BSST_CONCR_BEAM_I = -4


class RobotBarSectionConcreteDataValue(IntEnum):
    I_BSCDV_BEAM_H = 0  # Beam height.
    I_BSCDV_BEAM_B = 1  # Beam width.
    I_BSCDV_BEAM_T_HF = 100  # T-beam flange height.
    I_BSCDV_BEAM_T_BF = 101  # T-beam flange width.

class StructureBars:
    BAR_STRUCTURAL_TYPE_BEAM = 1

    def __init__(self, robot):
        self.robot = robot
        self.nodes = dict()  # {node id: {"id", "x", "y", "z"}}
        self.bars = list()
        self.get_all_nodes()
        self.get_all_bars()

    def get_all_nodes(self):
        nodes = self.robot.Project.Structure.Nodes
        all_nodes = nodes.GetAll()
        for i in range(1, all_nodes.Count + 1):
            node = CastTo(all_nodes.Get(i), "IRobotNode")
            self.nodes[node.Number] = {
                "id": node.Number,
                "x": node.X,
                "y": node.Y,
                "z": node.Z,
            }

    def get_all_bars(self):
        bars = self.robot.Project.Structure.Bars
        all_bars = bars.GetAll()
        for idx in range(1, all_bars.Count + 1):
            bar = all_bars.Get(idx)
            if bar.StructuralType in [1,2]:

                if bar.StartNode not in self.nodes or bar.EndNode not in self.nodes:
                    print("Beam {}: node not found, skipped.".format(bar.Number))
                    continue
                try:
                    section = self.parseSection(bar)
                except Exception as exc:
                    print("Beam {}: {} -> default section".format(bar.Number, exc))
                    raise ValueError("Beam {}: {} -> default section".format(bar.Number, exc))
                    section = {"name": "?", "shape": "UNKNOWN", "goe": {"gamma": 0.0}}

                item = {
                    "id": bar.Number,
                    "start_xyz": self.nodes[bar.StartNode],
                    "end_xyz": self.nodes[bar.EndNode],
                    "section": section,
                }
                self.bars.append(item)
        print("{} beam(s) read from Robot.".format(len(self.bars)))

    # -- section -------------------------------------------------------------
    def parseSection(self, bar):
        SH = RobotBarSectionShapeType
        CB = RobotBarSectionConcreteDataValue

        # b = width, h = height of the section (T beam: b = web width,
        # hf / bf = flange height / width). In plan the beam is drawn with
        # its width b (or h when the section is rotated by 90 degrees).
        BEAM_PARAMS = {
            SH.I_BSST_CONCR_BEAM: {"b": CB.I_BSCDV_BEAM_B, "h": CB.I_BSCDV_BEAM_H},
            SH.I_BSST_CONCR_BEAM_RECT: {
                "b": CB.I_BSCDV_BEAM_B,
                "h": CB.I_BSCDV_BEAM_H,
            },
            SH.I_BSST_CONCR_BEAM_T: {
                "b": CB.I_BSCDV_BEAM_B,
                "h": CB.I_BSCDV_BEAM_H,
                "hf": CB.I_BSCDV_BEAM_T_HF,
                "bf": CB.I_BSCDV_BEAM_T_BF,
            },
            SH.I_BSST_CONCR_BEAM_I: {"b": CB.I_BSCDV_BEAM_B, "h": CB.I_BSCDV_BEAM_H},
        }

        LENGTH_SCALE = 1.0  # Robot API = metres
        if not bar.HasLabel(IRobotLabelType.I_LT_BAR_SECTION):
            raise ValueError("no section assigned")

        label = bar.GetLabel(IRobotLabelType.I_LT_BAR_SECTION)
        data = CastTo(label.Data, "IRobotBarSectionData")

        try:
            shape = SH(data.ShapeType)
            shape_name = shape.name
        except ValueError:
            shape = None
            shape_name = "UNKNOWN({})".format(data.ShapeType)
            raise ValueError(shape_name)


        result = {"name": label.Name, "shape": shape_name, "goe": {}}

        # rotation of the section (degrees), used only to know whether the
        # width seen in plan is b or h
        try:
            result["goe"]["gamma"] = data.GetValue(24)
        except Exception:
            result["goe"]["gamma"] = 0.0

        params = BEAM_PARAMS.get(shape)
        if params:
            concrete = CastTo(data.Concrete, "IRobotBarSectionConcreteData")
            for key, idx in params.items():
                result["goe"][key] = concrete.GetValue(int(idx)) * LENGTH_SCALE
        return result




class IRobotObjectType(IntEnum):
    I_OT_UNDEFINED = -1        # Auxiliary identifier denoting undefined or unknown object type. Available since version 2.0.
    I_OT_OBJECT = -2           # Shared auxiliary identifier for bar, slab, volumetric object and geometric object. Available since version 2.0.
    I_OT_NODE = 0              # Node.
    I_OT_BAR = 1               # Bar.
    I_OT_CASE = 2              # Load case.
    I_OT_FAMILY = 3            # Group.
    I_OT_PANEL = 4             # Panel.
    I_OT_FINITE_ELEMENT = 5    # Finite element.
    I_OT_GEOMETRY = 6          # Geometrical object.
    I_OT_VOLUME = 7            # Volumetric object. Available since version 3.


import os 

import matplotlib.pyplot as plt
import numpy as np


def plot_bar_results(data):
    """
    Plot all internal-force components for a bar in one figure.

    Diagram convention:
        Y-  -> UP
        Y+  -> DOWN

    Parameters
    ----------
    data : dict
        {
            "id": 1,
            "length": 3.2,
            "cases": {
                "Poids": [
                    {
                        "x": 0.0,
                        "fx": ...,
                        "fy": ...,
                        "fz": ...,
                        "mx": ...,
                        "my": ...,
                        "mz": ...
                    },
                    ...
                ]
            }
        }
    """

    cases = data["cases"]
    bar_id = data.get("id", "")
    length = data.get("length", "")

    components = [
        ("fx", "N — Effort normal", "kN"),
        ("fy", "Fy", "kN"),
        ("fz", "Fz", "kN"),
        ("mx", "Mx", "kN·m"),
        ("my", "My", "kN·m"),
        ("mz", "Mz", "kN·m"),
    ]

    fig, axes = plt.subplots(
        2,
        3,
        figsize=(18, 10),
        constrained_layout=True,
    )

    axes = axes.ravel()

    for ax, (component, title, unit) in zip(axes, components):

        global_max = 0.0
        max_x = None
        max_case = None
        max_value = None

        # -------------------------------------------------------
        # Plot each load case
        # -------------------------------------------------------
        for case_name, points in cases.items():

            x = np.array(
                [point["x"] for point in points],
                dtype=float,
            )

            values = np.array(
                [point[component] for point in points],
                dtype=float,
            )

            # Remove insignificant numerical noise
            values[np.abs(values) < 1e-8] = 0.0

            ax.plot(
                x,
                values,
                linewidth=1.8,
                label=case_name,
            )

            # Find maximum absolute value
            index = np.argmax(np.abs(values))

            if abs(values[index]) > global_max:
                global_max = abs(values[index])
                max_x = x[index]
                max_case = case_name
                max_value = values[index]

        # -------------------------------------------------------
        # Zero/reference axis
        # -------------------------------------------------------
        ax.axhline(
            0.0,
            linewidth=1.2,
        )

        # -------------------------------------------------------
        # Grid
        # -------------------------------------------------------
        ax.minorticks_on()

        ax.grid(
            which="major",
            linewidth=0.8,
            alpha=0.45,
        )

        ax.grid(
            which="minor",
            linewidth=0.4,
            alpha=0.20,
        )

        # -------------------------------------------------------
        # Y direction convention
        #
        # Mathematical positive Y normally points UP.
        #
        # Here:
        #       Y- = UP
        #       Y+ = DOWN
        #
        # Therefore invert the vertical axis.
        # -------------------------------------------------------
        ax.invert_yaxis()

        # -------------------------------------------------------
        # Maximum annotation
        # -------------------------------------------------------
        if max_x is not None:

            ylim = ax.get_ylim()
            yrange = abs(ylim[1] - ylim[0])

            offset = 0.06 * yrange

            if max_value >= 0:
                text_y = max_value + offset
                vertical_alignment = "bottom"
            else:
                text_y = max_value - offset
                vertical_alignment = "top"

            ax.scatter(
                [max_x],
                [max_value],
                s=45,
                zorder=5,
            )

            ax.annotate(
                f"Max = {max_value:.2f} {unit}\n"
                f"x = {max_x:.2f} m\n"
                f"{max_case}",
                xy=(max_x, max_value),
                xytext=(max_x, text_y),
                ha="center",
                va=vertical_alignment,
                fontsize=9,
                bbox=dict(
                    boxstyle="round,pad=0.35",
                    alpha=0.85,
                ),
                arrowprops=dict(
                    arrowstyle="->",
                    linewidth=0.9,
                ),
            )

        # -------------------------------------------------------
        # Axis labels / title
        # -------------------------------------------------------
        ax.set_title(
            title,
            fontsize=12,
            fontweight="bold",
        )

        ax.set_xlabel(
            "Position x (m)"
        )

        ax.set_ylabel(
            f"{unit}   (Y− ↑ / Y+ ↓)"
        )

        ax.set_xlim(
            0,
            length,
        )

        # -------------------------------------------------------
        # X ticks
        # -------------------------------------------------------
        ax.set_xticks(
            np.linspace(0, length, 9)
        )

    # -----------------------------------------------------------
    # Global title
    # -----------------------------------------------------------
    fig.suptitle(
        f"Bar {bar_id} — Internal Forces — L = {length} m",
        fontsize=17,
        fontweight="bold",
    )

    # -----------------------------------------------------------
    # One common legend
    # -----------------------------------------------------------
    handles, labels = axes[0].get_legend_handles_labels()

    fig.legend(
        handles,
        labels,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.965),
        ncol=min(len(labels), 6),
        frameon=True,
    )

    plt.show()












class BeamOptimzer : 
    def __init__(self) : 
        self.robot = win32com.client.Dispatch("Robot.Application")
        self.robot.Visible = True
        self.bars = StructureBars(self.robot).bars
        self.selection_results = dict()
        self.path = self.robot.Kernel.ProgramPath
        clr_dll = os.path.join(str(Path(self.robot.Kernel.ProgramPath).parent) , "Interop.RobotOM.dll" )
        clr.AddReference(clr_dll)
        import RobotOM
        # self.bars = BeamsAndColumns(robot).bars 
        self.get_load_cases()
        # self.get_bar_results()

        self.Text_info = f"{len(self.selection_results)} Beam(s) selected ....."



    def set_ratio_steel(self,value) : 
        self.ratio = 0.01 * value


    def get_bar_results(self) : 
        self.bars = StructureBars(self.robot).bars
        self.bars_results = dict()
        if self.robot.Project.Structure.IsCalcModelGenerated :
            bars = self.robot.Project.Structure.Bars.GetAll()
            forces_server = self.robot.Project.Structure.Results.Bars.Forces
            for bar_index in range(1, bars.Count + 1) : 
                bar_result = {
                    "id" : bars.Get(bar_index).Number , 
                    "type" : {2 : "col" , 1 : "beam"}.get(bars.Get(bar_index).StructuralType) , 
                    "division" : 50 , 
                    "length" : round(bars.Get(bar_index).Length ,2), 
                    "cases" : dict()
                }
  
                bar_coridnates = [i * bar_result["length"] / bar_result["division"] for i in range( bar_result["division"] + 1)]
                for case_id in self.load_cases :
                    bar_result["cases"][self.load_cases[case_id]] = list()
                    for idx_pt in range(0 , len(bar_coridnates) ) : 
                        force = forces_server.ValueByNPoints(bar_result["id"] , int(case_id) , bar_result["division"] , idx_pt )
                        itm = {"x": bar_coridnates[idx_pt], "fx": force.FX, "fy": force.FY, "fz": force.FZ,"mx": force.MX, "my": force.MY, "mz": force.MZ}
                        bar_result["cases"][self.load_cases[case_id]].append(itm)
                self.bars_results[bars.Get(bar_index).Number] = bar_result
            # print(json.dumps(self.bars_results , indent = 4) , file = open('data.json' , "w"))


    def get_load_cases(self) : 
        cases = self.robot.Project.Structure.Cases.GetAll()
        self.load_cases = dict()
        for i in range(1, cases.Count + 1):
            case = cases.Get(i)
            self.load_cases[case.Number] = case.Name
        # print(json.dumps(self.load_cases))

    def get_current_selection(self) : 
        selections = self.robot.Project.Structure.Selections
        objs = CastTo(selections.Get(IRobotObjectType.I_OT_BAR),"IRobotSelection")
        bars = self.robot.Project.Structure.Bars.GetAll()
        loaded_itm = 0 
        self.selection_results = dict()
        for i in range(1 , objs.Count + 1 ) : 
            obj_number = objs.Get(i)
            obj_type = {2 : "col" , 1 : "beam"}.get(self.robot.Project.Structure.Objects.GetStructuralType(obj_number)) 
            # print("Selected Object " ,obj_number  )
            if obj_type in ["col","beam"] : 
                # get the bar frm bars : 
                bar = [bars.Get(idx) for idx in range(1,bars.Count + 1 ) if bars.Get(idx).Number == obj_number ][0]
                forces_server = self.robot.Project.Structure.Results.Bars.Forces
                bar_result = {
                        "id" : obj_number , 
                        "type" : obj_type , 
                        "division" : 50 , 
                        "length" : round(bar.Length ,2), 
                        "cases" : dict()
                }
      
                bar_coridnates = [i * bar_result["length"] / bar_result["division"] for i in range( bar_result["division"] + 1)]
                for case_id in self.load_cases :
                    bar_result["cases"][self.load_cases[case_id]] = list()
                    for idx_pt in range(0 , len(bar_coridnates) ) : 
                        force = forces_server.ValueByNPoints(bar_result["id"] , int(case_id) , bar_result["division"] , idx_pt )
                        itm = {"x": bar_coridnates[idx_pt], "fx": force.FX, "fy": force.FY, "fz": force.FZ,"mx": force.MX, "my": force.MY, "mz": force.MZ}
                        bar_result["cases"][self.load_cases[case_id]].append(itm)
                self.selection_results[bar_result["id"]] = bar_result
            loaded_itm = loaded_itm + 1 
            yield int(100 * loaded_itm / objs.Count )
            #print(json.dumps(self.selection_results , indent = 4))

    def optimizeSelectedBeams(self,h_min,reduce_h) : 
        self.bars = StructureBars(self.robot).bars
        len_selected = len(self.selection_results)
        percent = 0 
        done = 0 
        for b_id in self.selection_results : 
            beam_goe_data = [i for i in self.bars if i["id"] == b_id][0]
            self.optmize_beam(beam_calc_data = self.selection_results.get(b_id)  ,beam_goe_data = beam_goe_data , h_min = h_min ,reduce_h = reduce_h )
            done = done + 1 
            yield round(100*done/len_selected)

    def optmize_beam(self,beam_calc_data ,beam_goe_data,h_min,reduce_h ) : 
        if beam_calc_data["type"] != "beam" : return
        beamMaterial = self.robot.Project.Structure.Bars.Get(beam_goe_data["id"]).GetLabel(IRobotLabelType.I_LT_BAR_SECTION).Data.MaterialName
        all_my = list()
        for idx , case  in self.load_cases.items() : 
            my = [i["my"] for i in beam_calc_data["cases"][case]]
            all_my.extend(my)
        max_my = max(all_my)
        my_inf = min(all_my) 
        my = max([abs(max_my) , abs(my_inf)])
        if beam_goe_data["section"]["shape"] == "I_BSST_CONCR_BEAM_RECT":
            b = beam_goe_data["section"]["goe"]["b"]  # m
            d, K = self.estimate_beam_depth_rect(my, b, fck=20, fyk=400, rho=self.ratio)
            d_round = math.ceil(d / 0.05) * 0.05      # round up to 5 cm
            h = int(100*(d_round + 0.05))*0.01                 # + cover (~5 cm) -> total height
            #print(f"My = {my} N·m, b = {b} m, K = {K:.3f} MPa, d = {d:.3f} m (use {d_round:.2f} m), h ≈ {h:.2f} m") 
            h_ret = None
            if h_min is not None:
                if reduce_h:
                    h_ret = max(h, h_min)
                else:
                    h_ret = max(h, h_min, beam_goe_data["section"]["goe"]["h"])
            else:
                if reduce_h:
                    h_ret = min(h, beam_goe_data["section"]["goe"]["h"])
                else:
                    h_ret = max(h, beam_goe_data["section"]["goe"]["h"])

            sec = self.createRectSectionBeam(beamMaterial , b , h_ret )
            bar = self.robot.Project.Structure.Bars.Get(beam_goe_data["id"]).SetLabel(IRobotLabelType.I_LT_BAR_SECTION,sec)



    def createRectSectionBeam(self ,material ,  b , h ) : 
        labels = self.robot.Project.Structure.Labels #IRobotLabelServer
        section_name = f"{int(b * 100)}x{int(h * 100)}"
        section = labels.Create(3, section_name)  # I_LT_BAR_SECTION
        # print(f"created section {section_name}")
        data = section.Data
        data.MaterialName = material
        data.ShapeType = RobotBarSectionShapeType.I_BSST_CONCR_BEAM_RECT
        concrete = data.Concrete
        concrete.SetValue(RobotBarSectionConcreteDataValue.I_BSCDV_BEAM_B, b)  # Replace with 
        concrete.SetValue(RobotBarSectionConcreteDataValue.I_BSCDV_BEAM_H, h)
        data.CalcNonstdGeometry()
        labels.Store(section)
        return section_name


    def estimate_beam_depth_rect(self,my, b, fck=20.0, fyk=400.0, rho=0.01,
                            gamma_c=1.5, gamma_s=1.15, my_unit="N.m"):
        """
        Estimate effective depth d (m) of a rectangular RC beam so that the
        steel ratio rho = As/(b*d) is the target value (good steel density).

        my   : design bending moment (N·m by default, or "kN.m")
        b    : beam width (m)
        fck  : concrete characteristic strength (MPa)
        fyk  : steel yield strength (MPa)
        rho  : target steel ratio (0.010 - 0.015 is a common economical range)
        """
        # Moment in MN·m
        m = abs(my) / (1e6 if my_unit == "N.m" else 1e3)

        fcd = fck / gamma_c            # MPa
        fyd = fyk / gamma_s            # MPa

        omega = rho * fyd / fcd        # mechanical steel ratio
        xi = omega / 0.8               # x/d (rectangular stress block, EC2)
        mu = omega * (1 - 0.5 * omega) # reduced moment M/(b d² fcd)
        # (EC2 rectangular block: z = d(1 - 0.4 x/d) = d(1 - 0.5 omega))

        if xi > 0.45:
            print(f"Warning: x/d = {xi:.2f} > 0.45 (ductility), lower rho")

        K = mu * fcd                   # MPa  -> M = K · b · d²
        d = math.sqrt(m / (K * b))     # m
        return d, K











def BeamsOptPipeline(robot) : 
    bopt = BeamOptimzer(robot)
    pass 


if __name__ == "__main__":

    BeamOptimzer()