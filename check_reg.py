import winreg

def enum_tree(key, path, indent=0):
    try:
        r = winreg.OpenKey(key, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
    except FileNotFoundError:
        return
    # Values
    try:
        i = 0
        while True:
            name, val, vtype = winreg.EnumValue(r, i)
            print("  " * indent + f"  {name} = {val!r}")
            i += 1
    except OSError:
        pass
    # Subkeys
    try:
        i = 0
        while True:
            subname = winreg.EnumKey(r, i)
            print("  " * indent + subname + "\\")
            enum_tree(key, path + "\\" + subname, indent + 1)
            i += 1
    except OSError:
        pass
    winreg.CloseKey(r)

print("HKLM\\SOFTWARE\\MindMaker (32-bit view):")
enum_tree(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\MindMaker")
