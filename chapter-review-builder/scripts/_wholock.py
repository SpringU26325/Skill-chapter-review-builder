# -*- coding: utf-8 -*-
"""用 Restart Manager 查谁锁着文件。用法：python _wholock.py <file>"""
import ctypes, ctypes.wintypes as w, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [("dwProcessId", w.DWORD), ("ProcessStartTime", w.FILETIME)]


class RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [
        ("Process", RM_UNIQUE_PROCESS),
        ("strAppName", w.WCHAR * 256),
        ("strServiceShortName", w.WCHAR * 64),
        ("ApplicationType", ctypes.c_int),
        ("AppStatus", w.DWORD),
        ("TSSessionId", w.DWORD),
        ("bRestartable", w.BOOL),
    ]


rm = ctypes.WinDLL("rstrtmgr")
h = w.DWORD()
key = "wb-" + os.urandom(8).hex()
if rm.RmStartSession(ctypes.byref(h), 0, ctypes.c_wchar_p(key)) != 0:
    print("RmStartSession failed"); sys.exit(1)
path = os.path.abspath(sys.argv[1])
files = (ctypes.c_wchar_p * 1)(path)
if rm.RmRegisterResources(h, 1, files, 0, None, 0, None) != 0:
    print("RmRegisterResources failed"); sys.exit(1)

needed = w.UINT(0)
count = w.UINT(0)
reasons = w.DWORD(0)
rm.RmGetList(h, ctypes.byref(needed), ctypes.byref(count), None, ctypes.byref(reasons))
if needed.value == 0:
    print("no locker")
    rm.RmEndSession(h); sys.exit(0)

arr = (RM_PROCESS_INFO * needed.value)()
count = w.UINT(needed.value)
rc = rm.RmGetList(h, ctypes.byref(needed), ctypes.byref(count), arr, ctypes.byref(reasons))
print("rc", rc, "count", count.value)
for i in range(count.value):
    print("LOCKER pid=%d app=%s svc=%s" % (arr[i].Process.dwProcessId, arr[i].strAppName, arr[i].strServiceShortName))
rm.RmEndSession(h)
