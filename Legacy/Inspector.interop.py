



interop_dll = "C:\Program Files\Autodesk\Autodesk Robot Structural Analysis Professional 2015\System\Exe\Autodesk.ASD.Interop.dll"

"""
inspect_interop.py
Dump enums, interfaces, classes, structs and delegates of a .NET assembly
(e.g. Autodesk.ASD.Interop.dll) using reflection-only loading.
Nothing from the DLL is executed.

Requirements:
    pip install pythonnet
    (Windows 7: Python 3.8 or older is required by pythonnet)

Usage:
    python inspect_interop.py "C:\\path\\to\\Autodesk.ASD.Interop.dll"
    python inspect_interop.py "C:\\path\\to\\Autodesk.ASD.Interop.dll" -o dump.txt
"""

import argparse
import os
import sys

import clr  # noqa: F401  (pythonnet)
from System import AppDomain, ResolveEventHandler
from System.Reflection import Assembly, BindingFlags, ReflectionTypeLoadException


def main():
    # ap = argparse.ArgumentParser()
    # ap.add_argument("path", help="Path to the .NET assembly (.dll)")
    # ap.add_argument("-o", "--out", default="ASD_Interop_Dump.txt")
    # args = ap.parse_args()

    full = os.path.abspath(interop_dll)
    folder = os.path.dirname(full)
    args_out = os.path.join(os.path.dirname(__file__), os.path.basename(interop_dll) + ".txt")
    # Resolve dependencies from the same folder first, then from the GAC
    def resolve(sender, e):
        short = str(e.Name).split(",")[0]
        local = os.path.join(folder, short + ".dll")
        if os.path.exists(local):
            return Assembly.ReflectionOnlyLoadFrom(local)
        try:
            return Assembly.ReflectionOnlyLoad(e.Name)
        except Exception:
            return None

    handler = ResolveEventHandler(resolve)
    AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += handler

    asm = Assembly.ReflectionOnlyLoadFrom(full)

    # GetTypes can partly fail when dependencies are missing; keep what loaded
    try:
        types = list(asm.GetTypes())
    except ReflectionTypeLoadException as ex:
        types = [t for t in ex.Types if t is not None]
        print("WARNING: some types could not be loaded (missing dependencies).")

    flags = (BindingFlags.Public | BindingFlags.Instance |
             BindingFlags.Static | BindingFlags.DeclaredOnly)
    enum_flags = BindingFlags.Public | BindingFlags.Static

    out = []
    add = out.append

    add("Assembly : %s" % asm.FullName)
    add("Runtime  : %s" % asm.ImageRuntimeVersion)
    add("Types    : %d" % len(types))
    add("")

    public = sorted(
        (t for t in types if t.IsPublic or t.IsNestedPublic),
        key=lambda t: str(t.FullName),
    )
    enums = [t for t in public if t.IsEnum]
    ifaces = [t for t in public if t.IsInterface]
    others = [t for t in public if not t.IsEnum and not t.IsInterface]

    def members(t):
        result = []
        for m in t.GetMembers(flags):
            # skip property/event accessors (get_X, set_X, add_X, remove_X)
            if str(m.MemberType) == "Method" and getattr(m, "IsSpecialName", False):
                continue
            result.append("    %s %s" % (m.MemberType, m.ToString()))
        return sorted(result)

    add("=================== ENUMS ===================")
    for t in enums:
        add("")
        add("enum %s" % t.FullName)
        for f in t.GetFields(enum_flags):
            add("    %s = %s" % (f.Name, f.GetRawConstantValue()))

    add("")
    add("================= INTERFACES =================")
    for t in ifaces:
        add("")
        add("interface %s" % t.FullName)
        out.extend(members(t))

    add("")
    add("============ CLASSES / STRUCTS / DELEGATES ============")
    for t in others:
        kind = "class"
        base = t.BaseType
        if t.IsValueType:
            kind = "struct"
        elif base is not None and str(base.FullName) == "System.MulticastDelegate":
            kind = "delegate"
        add("")
        add("%s %s" % (kind, t.FullName))
        if base is not None:
            add("    base: %s" % base.FullName)
        out.extend(members(t))

    add("")
    add("================ REFERENCED ASSEMBLIES ================")
    for r in asm.GetReferencedAssemblies():
        add("    %s" % r.FullName)

    with open(args_out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out))

    print("Done. Dump written to: %s" % os.path.abspath(args_out))
    print("Public enums: %d  interfaces: %d  other: %d"
          % (len(enums), len(ifaces), len(others)))


if __name__ == "__main__":
    sys.exit(main())
