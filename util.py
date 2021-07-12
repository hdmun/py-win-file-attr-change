from datetime import datetime

import ntsecuritycon
import pythoncom
import pywintypes
import win32con
import win32file

from PIL import Image
from win32com.propsys import propsys, pscon
from win32comext.shell import shellcon

'''
pip install pypiwin32, Pillow
'''

def get_media_date_encoded(path: str):
    properties = propsys.SHGetPropertyStoreFromParsingName(path, None, shellcon.GPS_READWRITE, propsys.IID_IPropertyStore)
    prop_dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
    properties = None  # release
    format_dt = '%Y-%m-%d %H:%M:S'
    return datetime.strptime(prop_dt.strftime(format_dt), format_dt)


def set_media_date_encoded(path: str, timestamp: float):
    properties = propsys.SHGetPropertyStoreFromParsingName(path, None, shellcon.GPS_READWRITE, propsys.IID_IPropertyStore)
    prop_var_date = propsys.PROPVARIANTType(pywintypes.Time(timestamp), pythoncom.VT_DATE)
    properties.SetValue(pscon.PKEY_Media_DateEncoded, prop_var_date)
    properties.Commit()
    properties = None


def get_photo_take_time(path: str):
    try:
        img = Image.open(path)
        exif = img._getexif()
        date = exif[0x9003]
        return datetime.strptime(date, '%Y:%m:%d %H:%M:%S')
    except:
        return None


def change_file_time(filename: str, timestamp: float):
    winfile = win32file.CreateFile(
        filename,
        ntsecuritycon.FILE_WRITE_ATTRIBUTES,
        0,
        None,
        win32con.OPEN_EXISTING,
        0,
        None)

    wintime = pywintypes.Time(timestamp)
    create_time = wintime
    last_access_time = None
    modified_time = wintime
    win32file.SetFileTime(winfile, create_time, last_access_time, modified_time)
    winfile.close()
