import pyautogui
import time
import ctypes
import ctypes.wintypes


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [
        ('cbSize', ctypes.wintypes.UINT),
        ('dwTime', ctypes.wintypes.DWORD),
    ]


PLASTINPUTINFO = ctypes.POINTER(LASTINPUTINFO)

user32 = ctypes.windll.user32
GetLastInputInfo = user32.GetLastInputInfo
GetLastInputInfo.restype = ctypes.wintypes.BOOL
GetLastInputInfo.argtypes = [PLASTINPUTINFO]

kernel32 = ctypes.windll.kernel32
GetTickCount = kernel32.GetTickCount
Sleep = kernel32.Sleep


def wait_until_idle(idle_time=60):
    """Wait until no more user activity is detected.

    This function won't return until `idle_time` seconds have elapsed
    since the last user activity was detected.
    """

    idle_time_ms = int(idle_time*1000)
    liinfo = LASTINPUTINFO()
    liinfo.cbSize = ctypes.sizeof(liinfo)
    while True:
        GetLastInputInfo(ctypes.byref(liinfo))
        elapsed = GetTickCount() - liinfo.dwTime
        if elapsed >= idle_time_ms:
            break
        Sleep(idle_time_ms - elapsed or 1)


def wait_until_active(tol=5):
    """Wait until awakened by user activity.

    This function will block and wait until some user activity
    is detected. Because of the polling method used, it may return
    `tol` seconds (or less) after user activity actually began.
    """

    liinfo = LASTINPUTINFO()
    liinfo.cbSize = ctypes.sizeof(liinfo)
    lasttime = None
    delay = 1  # ms
    maxdelay = int(tol*1000)
    while True:
        GetLastInputInfo(ctypes.byref(liinfo))
        if lasttime is None:
            lasttime = liinfo.dwTime
        if lasttime != liinfo.dwTime:
            break
        delay = min(2*delay, maxdelay)
        Sleep(delay)


def test():
    print("Waiting for 10 seconds of no user input...")
    wait_until_idle(10)
    user32.MessageBeep(0)
    print("Ok. Now, do something!")
    wait_until_active(1)
    user32.MessageBeep(0)
    print("Done.")


def stay_alive():
    pyautogui.FAILSAFE = True
    titles = pyautogui.getAllTitles()
    msTeamWindowTitles = [x for x in titles if 'Microsoft Teams' in x]
    if(len(msTeamWindowTitles) == 0):
        print("Can't find MS Team window")
        return
    else:
        print("Found MS Team window")
        msTeamWindow = pyautogui.getWindowsWithTitle(msTeamWindowTitles[0])[0]
        msTeamWindow.maximize()
        msTeamWindow.activate()
        while True:
            for i in [100, 200, 300]:
                time.sleep(3)
                pyautogui.dragTo(0, i)
                pyautogui.click()


def main():
    while True:
        print("Waiting for 10 seconds of no user input...")
        wait_until_idle(10)
        stay_alive()


if __name__ == '__main__':
    main()
