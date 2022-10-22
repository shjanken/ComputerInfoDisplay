import sys
import os
from typing import List

import tkinter as tk
from tkinter import StringVar, ttk
import wmi
from PIL import Image
from pystray import MenuItem, Icon


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class MainPanel():

    # save the lastest labelframe's row position
    current_row: int = 0

    def __init__(self, *args, **kwargs):
        # create and init a main window
        self.root = tk.Tk()
        self.root.resizable(False, False)
        self.root.attributes("-alpha", 1)
        self.root.geometry(f"-100+50")
        self.root.title("Computer Info Display")
        # self.root.overrideredirect(True)

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(column=0, row=0, sticky="ewns")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self.__hide_window)

    def addInfoList(self, title: str, msg_list: List[str]) -> None:
        lbf = ttk.LabelFrame(self.main_frame, text=f"  {title}  ")
        lbf.grid(row=self.current_row, column=0, sticky="ew")
        lbf.grid_configure(padx=5, pady=5)

        for idx, msg in enumerate(msg_list):
            lb = ttk.Label(lbf, text=msg)
            lb.grid(column=0, row=idx, sticky="ew")
            lb.grid_configure(padx=10)

        self.current_row += 1  # add net labelframe at next row position

    def start(self):
        self.root.mainloop()

    def __hide_window(self):
        """hide the window but not close it"""
        # init tray icon and run it
        icon_img = Image.open(resource_path('assets/tray_16x.ico'))
        menu = (MenuItem("Show", self.__showWindow),
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


def fetchDiskInfo(wmi_client) -> List[str]:
    """fetch the disk info with windows wmi"""
    info_list = []

    for d in wmi_client.Win32_DiskDrive():
        info_list.append(f"{d.Caption} :: {d.SerialNumber}")

    return info_list


def fetchNetWorkInfo(wmi_client) -> List[str]:
    """fetch network adapter info with windows wmi"""

    info_list = []
    for n in wmi_client.Win32_NetworkAdapterConfiguration(IPEnabled=True):
        info_list.append(f"{n.IPAddress[0]} :: {n.MACAddress}")

    return info_list


w = wmi.WMI()

main = MainPanel()
main.addInfoList("磁盘序列号", fetchDiskInfo(w))
main.addInfoList("网卡适配器", fetchNetWorkInfo(w))

main.start()
