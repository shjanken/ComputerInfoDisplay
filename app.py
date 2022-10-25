import sys
import os
from typing import List
from pathlib import Path

import tkinter as tk
from tkinter import StringVar, ttk
import wmi  # type: ignore
from PIL import Image  # type: ignore
from pystray import MenuItem, Icon, Menu  # type: ignore
import winshell  # type: ignore

PROGRAM_NAME = "ComputerInfo"  # the program name at runtime


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class AutoStarter():
    """ Get or sett the program autostart status."""

    statue: bool = False

    def __init__(self, wmi_client: wmi.WMI):
        self.wmi_client = wmi_client

    def __getProgramRunFolder(self) -> str:
        """use win32 wmi to get the ComputerInfo process path"""
        return __file__

    def setAutoStart(self):
        """Copy the program exe file to the autostart folder"""
        global status
        program_path = Path(self.__getProgramRunFolder())
        if not program_path.exists():
            print(f"{program_path.absolute()} not exists")
            for d in self.wmi_client.Win32_Process(Name=PROGRAM_NAME):
                print(d.Executable)
            return
        else:
            autostart_path = winshell.folder("Startup")
            winshell.copy_file(str(program_path),
                               autostart_path, no_confirm=True)
            self.statue = True


class MainPanel():
    """main panel display the info message list in the window"""

    # save the lastest labelframe's row position
    current_row = 0

    def __init__(self, autostarter: AutoStarter, *args, **kwargs) -> None:
        # create and init a main window
        self.root = tk.Tk()
        self.root.resizable(False, False)
        self.root.attributes("-alpha", 1)
        self.root.geometry(f"-100+50")
        self.root.title("Computer Info Display")
        # self.root.overrideredirect(True) # disable window decoration

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(column=0, row=0, sticky="ewns")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self.__hideWindow)

        # use autostarter instance to set program auto start
        self.starter = autostarter

    def addInfoList(self, title, msg_list):
        """add info list to main window and display these infos"""
        lbf = ttk.LabelFrame(self.main_frame, text=f"  {title}  ")
        lbf.grid(row=self.current_row, column=0, sticky="ew")
        lbf.grid_configure(padx=5, pady=5)

        for idx, msg in enumerate(msg_list):
            lb = ttk.Label(lbf, text=msg)
            lb.grid(column=0, row=idx, sticky="ew")
            lb.grid_configure(padx=10)

        self.current_row += 1  # set next labelframe display row index

    def start(self):
        self.root.mainloop()

    def __setAutoStart(self):
        self.starter.setAutoStart()

    def __hideWindow(self):
        """hide the window but not close it"""
        # init tray icon and run it
        icon_img = Image.open(resource_path('assets/tray_32x.ico'))
        menu = (MenuItem("Show", self.__showWindow),
                MenuItem("AutoStart", self.__setAutoStart,
                         checked=lambda i: self.starter.statue),
                MenuItem(Menu.SEPARATOR, None),
                MenuItem("Quit", self.__destroyWindow))
        self.icon = Icon("info", icon_img, menu=menu)

        self.root.withdraw()
        self.icon.run()

    def __destroyWindow(self):
        """close window and quti app"""
        self.icon.stop()
        self.root.destroy()

    def __showWindow(self):
        """display window from hidden status"""
        self.icon.stop()
        self.root.deiconify()


def fetchDiskInfo(wmi_client):
    """fetch the disk info with windows wmi"""

    return [f"{d.Caption} :: {d.SerialNumber}" for d in wmi_client.Win32_DiskDrive()]


def fetchNetWorkInfo(wmi_client):
    """fetch network adapter info with windows wmi"""

    return [f"{n.IPAddress[0]} :: {n.MACAddress}"
            for n in wmi_client.Win32_NetworkAdapterConfiguration(IPEnabled=True)]


w = wmi.WMI()

autostarter = AutoStarter(w)

main = MainPanel(autostarter)
main.addInfoList("磁盘序列号", fetchDiskInfo(w))
main.addInfoList("网卡适配器", fetchNetWorkInfo(w))

main.start()
