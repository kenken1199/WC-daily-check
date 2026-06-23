import ctypes, string, sys
sys.stdout.reconfigure(encoding="utf-8")

USB_LABEL = "緑3F瓶充填"

kernel32 = ctypes.windll.kernel32
type_names = {0:"UNKNOWN", 1:"NO_ROOT", 2:"REMOVABLE", 3:"FIXED", 4:"REMOTE", 5:"CDROM", 6:"RAMDISK"}
bitmask = kernel32.GetLogicalDrives()
for letter in string.ascii_uppercase:
    if bitmask & 1:
        drive = f"{letter}:\\"
        dtype = kernel32.GetDriveTypeW(drive)
        label_buf = ctypes.create_unicode_buffer(1024)
        kernel32.GetVolumeInformationW(drive, label_buf, 1024, None, None, None, None, 0)
        label = label_buf.value
        match = (label == USB_LABEL)
        print(f'{drive}  type={dtype}({type_names.get(dtype,"?")})  label={repr(label)}  match={match}')
    bitmask >>= 1
