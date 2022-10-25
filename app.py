import sys
import os
import platform
from pathlib import Path

import tkinter as tk
from tkinter import ttk
import wmi  # type: ignore
from PIL import Image  # type: ignore
from pystray import MenuItem, Icon, Menu  # type: ignore
import winshell  # type: ignore

PROGRAM_NAME = "ComputerInfo.exe"  # the program name at runtime


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

    def search_program_path(self) -> str:
        """use wmi win32_process find the program process.
        return the process.executablePath"""
        processes = self.wmi_client.Win32_Process(Name=PROGRAM_NAME)
        if not processes:
            raise OSError("program not running")
        return processes[0].ExecutablePath

    def set_auto_start(self) -> None:
        """Copy the program exe file to the autostart folder"""
        print(self.search_program_path())


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
        self.root.title("上海二手车交易专网助手v1.0")
        # self.root.overrideredirect(True) # disable window decoration

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(column=0, row=0, sticky="ewns")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self.__hide_window)

        # use autostarter instance to set program auto start
        self.starter = autostarter

    def add_info_list(self, title, msg_list):
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

    def __set_auto_start(self):
        self.starter.setAutoStart()

    def __hide_window(self):
        """hide the window but not close it"""
        # init tray icon and run it
        icon_img = Image.open(resource_path('assets/tray_32x.ico'))
        menu = (MenuItem("显示", self.__show_window),
                MenuItem("自启", self.__set_auto_start,
                         checked=lambda i: self.starter.statue),
                MenuItem(Menu.SEPARATOR, None),
                MenuItem("退出", self.__destroy_window))
        self.icon = Icon("info", icon_img, menu=menu)

        self.root.withdraw()
        self.icon.run()

    def __destroy_window(self):
        """close window and quti app"""
        self.icon.stop()
        self.root.destroy()

    def __show_window(self):
        """display window from hidden status"""
        self.icon.stop()
        self.root.deiconify()


def fetch_disk_info(wmi_client):
    """fetch the disk info with windows wmi"""

    return [f"{d.Caption} :: {d.SerialNumber}" for d in wmi_client.Win32_DiskDrive()]


def fetch_os():
    """fetch info from operating system.
    info msg like: 'system info(arch)'"""
    return [f"{platform.platform()} ({platform.machine()})"]


def fetch_network_info(wmi_client):
    """fetch network adapter info with windows wmi"""

    return [f" {n.Description}:{n.IPAddress[0]}:{n.MACAddress}"
            for n in wmi_client.Win32_NetworkAdapterConfiguration(IPEnabled=True)]


w = wmi.WMI()

starter = AutoStarter(w)

main = MainPanel(starter)
main.add_info_list("操作系统类型", fetch_os())
main.add_info_list("磁盘序列号", fetch_disk_info(w))
main.add_info_list("网卡适配器", fetch_network_info(w))

main.start()
