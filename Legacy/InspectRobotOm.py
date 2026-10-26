"""
Describe the RobotOM interop assembly (Interop.RobotOM.dll) in a text file.

Uses .NET reflection (pythonnet) to write:
  1. assembly summary (counts per kind, namespaces)
  2. class inheritance tree
  3. interface hierarchy tree
  4. every enum with its values
  5. every interface: parents, implemented-by, properties, methods, events
  6. every class: base, interfaces, properties, methods, events, fields
  7. every struct

Nothing in Robot is called or modified; only the DLL metadata is read.
Run with Robot open:  python robotom_describe.py
"""
import os
import datetime
from pathlib import Path

import clr
import win32com.client as win32
from win32com.client import gencache

robapp = gencache.EnsureDispatch("Robot.Application")
c = win32.constants
structure = robapp.Project.Structure

clr_dll = os.path.join(str(Path(robapp.Kernel.ProgramPath).parent),
                       "Interop.RobotOM.dll")
clr.AddReference(clr_dll)
import RobotOM  # noqa: E402,F401
import System  # noqa: E402
from System.Reflection import Assembly, BindingFlags  # noqa: E402

OUT_DIR = r"C:\temp"
OUT_FILE = os.path.join(OUT_DIR, "robotom_description.txt")

FLAGS = (BindingFlags.Public | BindingFlags.Instance |
         BindingFlags.Static | BindingFlags.DeclaredOnly)


# ---------------------------------------------------------------- helpers
def tn(t):
    """Short readable type name."""
    if t is None:
        return "void"
    try:
        if t.IsByRef:
            return tn(t.GetElementType())
        return str(t.Name)
    except Exception:
        return "?"


def fmt_param(p):
    pt = p.ParameterType
    s = f"{tn(pt)} {p.Name}"
    if pt.IsByRef:
        s = ("out " if p.IsOut else "ref ") + s
    if p.IsOptional:
        s += " = optional"
    return s


def fmt_params(params):
    return ", ".join(fmt_param(p) for p in params)


def kind_of(t):
    if t.IsEnum:
        return "enum"
    if t.IsInterface:
        return "interface"
    if t.IsValueType:
        return "struct"
    if t.IsClass:
        return "class"
    return "type"


def direct_interfaces(t):
    """Interfaces declared directly (not inherited from another interface/base)."""
    all_if = list(t.GetInterfaces())
    inherited = set()
    for i in all_if:
        for j in i.GetInterfaces():
            inherited.add(str(j.FullName))
    if t.BaseType is not None:
        for j in t.BaseType.GetInterfaces():
            inherited.add(str(j.FullName))
    return [i for i in all_if if str(i.FullName) not in inherited]


def load_types(asm):
    try:
        return list(asm.GetTypes())
    except Exception as e:  # ReflectionTypeLoadException
        loaded = getattr(e, "Types", None)
        if loaded is None:
            raise
        return [t for t in loaded if t is not None]


# ------------------------------------------------------------- writers
class Out:
    def __init__(self, f):
        self.f = f

    def w(self, s=""):
        self.f.write(s + "\n")

    def title(self, s):
        self.w("\n" + "#" * 78)
        self.w("# " + s)
        self.w("#" * 78)


def write_members(o, t, indent="  "):
    props = sorted(t.GetProperties(FLAGS), key=lambda p: str(p.Name))
    if props:
        o.w(f"{indent}properties:")
        for p in props:
            acc = ("get" if p.CanRead else "") + \
                  ("/set" if p.CanWrite else "")
            idx = list(p.GetIndexParameters())
            name = p.Name + (f"[{fmt_params(idx)}]" if idx else "")
            o.w(f"{indent}  {tn(p.PropertyType)} {name} {{{acc.strip('/')}}}")

    methods = [m for m in t.GetMethods(FLAGS) if not m.IsSpecialName]
    methods.sort(key=lambda m: str(m.Name))
    if methods:
        o.w(f"{indent}methods:")
        for m in methods:
            static = "static " if m.IsStatic else ""
            o.w(f"{indent}  {static}{tn(m.ReturnType)} {m.Name}"
                f"({fmt_params(m.GetParameters())})")

    ctors = list(t.GetConstructors(BindingFlags.Public | BindingFlags.Instance
                                   | BindingFlags.DeclaredOnly))
    if ctors:
        o.w(f"{indent}constructors:")
        for k in ctors:
            o.w(f"{indent}  {tn(t)}({fmt_params(k.GetParameters())})")

    events = sorted(t.GetEvents(FLAGS), key=lambda e: str(e.Name))
    if events:
        o.w(f"{indent}events:")
        for e in events:
            o.w(f"{indent}  event {tn(e.EventHandlerType)} {e.Name}")

    fields = [f for f in t.GetFields(FLAGS) if not f.IsSpecialName]
    if fields:
        o.w(f"{indent}fields:")
        for f in sorted(fields, key=lambda f: str(f.Name)):
            o.w(f"{indent}  {tn(f.FieldType)} {f.Name}")


def write_enum(o, t):
    o.w(f"enum {t.FullName}")
    rows = []
    for f in t.GetFields(BindingFlags.Public | BindingFlags.Static):
        try:
            rows.append((int(f.GetRawConstantValue()), str(f.Name)))
        except Exception:
            rows.append((None, str(f.Name)))
    rows.sort(key=lambda r: (r[0] is None, r[0], r[1]))
    for val, name in rows:
        o.w(f"    {name} = {val}")


def write_tree(o, types_by_name, children, node, depth):
    t = types_by_name[node]
    o.w("  " * depth + f"{t.Name}")
    for ch in sorted(children.get(node, []), key=lambda n: types_by_name[n].Name):
        write_tree(o, types_by_name, children, ch, depth + 1)


# ------------------------------------------------------------------ main
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    asm = Assembly.LoadFrom(clr_dll)
    types = [t for t in load_types(asm) if t.IsPublic]
    types.sort(key=lambda t: str(t.FullName))
    by_name = {str(t.FullName): t for t in types}

    enums = [t for t in types if kind_of(t) == "enum"]
    ifaces = [t for t in types if kind_of(t) == "interface"]
    classes = [t for t in types if kind_of(t) == "class"]
    structs = [t for t in types if kind_of(t) == "struct"]

    # interface -> implementing classes
    implemented_by = {}
    for cl in classes:
        for i in cl.GetInterfaces():
            implemented_by.setdefault(str(i.FullName), []).append(tn(cl))

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        o = Out(f)
        o.w("ROBOTOM ASSEMBLY DESCRIPTION")
        o.w(f"generated : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
        o.w(f"dll       : {clr_dll}")
        o.w(f"assembly  : {asm.FullName}")

        # 1. summary
        o.title("1. SUMMARY")
        o.w(f"public types : {len(types)}")
        o.w(f"interfaces   : {len(ifaces)}")
        o.w(f"classes      : {len(classes)}")
        o.w(f"enums        : {len(enums)}")
        o.w(f"structs      : {len(structs)}")
        ns = sorted({str(t.Namespace) for t in types})
        o.w("namespaces   : " + ", ".join(ns))

        # 2. class tree
        o.title("2. CLASS INHERITANCE TREE")
        children, roots = {}, []
        for cl in classes:
            b = cl.BaseType
            bn = str(b.FullName) if b is not None else None
            if bn in by_name and kind_of(by_name[bn]) == "class":
                children.setdefault(bn, []).append(str(cl.FullName))
            else:
                roots.append(str(cl.FullName))
        for r in sorted(roots):
            write_tree(o, by_name, children, r, 0)

        # 3. interface tree
        o.title("3. INTERFACE HIERARCHY (direct parents)")
        for i in ifaces:
            parents = [tn(p) for p in direct_interfaces(i)]
            o.w(f"{tn(i)}" + (f"  <-  {', '.join(sorted(parents))}" if parents else ""))

        # 4. enums
        o.title(f"4. ENUMS ({len(enums)})")
        for t in enums:
            write_enum(o, t)
            o.w()

        # 5. interfaces
        o.title(f"5. INTERFACES ({len(ifaces)})")
        for t in ifaces:
            try:
                o.w("-" * 78)
                o.w(f"interface {t.FullName}")
                try:
                    o.w(f"  GUID: {t.GUID}")
                except Exception:
                    pass
                parents = [tn(p) for p in direct_interfaces(t)]
                if parents:
                    o.w(f"  extends: {', '.join(sorted(parents))}")
                impl = implemented_by.get(str(t.FullName))
                if impl:
                    o.w(f"  implemented by: {', '.join(sorted(impl))}")
                write_members(o, t)
            except Exception as e:
                o.w(f"  <failed: {e}>")

        # 6. classes
        o.title(f"6. CLASSES ({len(classes)})")
        for t in classes:
            try:
                o.w("-" * 78)
                abstract = "abstract " if t.IsAbstract else ""
                o.w(f"{abstract}class {t.FullName}")
                if t.BaseType is not None:
                    o.w(f"  base: {t.BaseType.Name}")
                ifs = [tn(i) for i in direct_interfaces(t)]
                if ifs:
                    o.w(f"  implements: {', '.join(sorted(ifs))}")
                write_members(o, t)
            except Exception as e:
                o.w(f"  <failed: {e}>")

        # 7. structs
        o.title(f"7. STRUCTS ({len(structs)})")
        for t in structs:
            try:
                o.w("-" * 78)
                o.w(f"struct {t.FullName}")
                write_members(o, t)
            except Exception as e:
                o.w(f"  <failed: {e}>")

    print("written:", OUT_FILE)
    print(f"types: {len(types)} (interfaces {len(ifaces)}, classes {len(classes)}, "
          f"enums {len(enums)}, structs {len(structs)})")


if __name__ == "__main__":
    main()
