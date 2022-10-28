import os
import platform
import sys
import textwrap
import tkinter as tk
import winreg as reg
from pathlib import Path
from tkinter import ttk
from typing import List

import winshell  # type: ignore
import wmi  # type: ignore
from PIL import Image  # type: ignore
from pystray import MenuItem, Icon, Menu  # type: ignore

PROGRAM_NAME = "ComputerInfo.exe"  # the program name at runtime
STARTUP_FOLDER = winshell.folder("startup")


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class AutoStarter:
    """ Get or sett the program autostart status."""

    is_auto_start: bool = False

    __program_no_running_error = "没有找到程序运行信息"

    def __init__(self, wmi_client: wmi.WMI):
        self.wmi_client = wmi_client
        self.is_auto_start = (Path(STARTUP_FOLDER) / PROGRAM_NAME).exists()

    def search_program_path(self) -> str:
        """use wmi win32_process find the program process.
        return the process.executablePath"""
        processes = self.wmi_client.Win32_Process(Name=PROGRAM_NAME)
        if not processes:
            raise OSError(self.__program_no_running_error)
        return processes[0].ExecutablePath

    def set_auto_start(self) -> None:
        """Copy the program exe file to the autostart folder, no confirm.
        if cat not found the program execute path, throw os error"""
        try:
            p = self.search_program_path()
            winshell.copy_file(p, winshell.folder("startup"), no_confirm=True)
            self.is_auto_start = True  # copy file success. the program is auto started
        except OSError as ex:
            raise OSError from ex

    def unset_auto_start(self) -> None:
        """Delete the file from user's startup folder"""
        p = Path(STARTUP_FOLDER) / PROGRAM_NAME
        if p.exists() and self.is_auto_start:
            winshell.delete_file(
                f"{p.resolve().absolute()}", no_confirm=True, silent=True, allow_undo=False)
            self.is_auto_start = False


class MainPanel:
    """main panel display the info message list in the window"""

    # save the latest labelframe row position
    current_row = 0

    def __init__(self, auto_starter: AutoStarter) -> None:
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

        # use auto_starter instance to set program auto start
        self.starter = auto_starter

    def add_info_list(self, title, msg_list):
        """add info list to main window and display info messages"""
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
        try:
            if not self.starter.is_auto_start:
                self.starter.set_auto_start()
            else:
                self.starter.unset_auto_start()

        except OSError as ex:
            if self.icon.HAS_NOTIFICATION:
                self.icon.notify(f"创建开机启动失败：{ex}", "错误")

    def __hide_window(self):
        """hide the window but not close it"""
        # init tray icon and run it
        icon_img = Image.open(resource_path('assets/tray_32x.ico'))
        menu = (MenuItem("显示", self.__show_window),
                MenuItem("自启", self.__set_auto_start,
                         checked=lambda _: self.starter.is_auto_start),
                MenuItem(Menu.SEPARATOR, None),
                MenuItem("退出", self.__destroy_window))
        self.icon = Icon("info", icon_img, menu=menu)

        self.root.withdraw()
        self.icon.run()

    def __destroy_window(self):
        """close window and quit app"""
        self.icon.stop()
        self.root.destroy()

    def __show_window(self):
        """display window from hidden status"""
        self.icon.stop()
        self.root.deiconify()


def fetch_disk_info(wmi_client):
    """fetch the disk info with windows wmi"""

    return [f"{d.Caption} :: {d.SerialNumber}" for d in wmi_client.Win32_DiskDrive()]


def fetch_os() -> List[str]:
    """fetch info from operating system.
    info msg like: 'system info(arch)'"""
    return [f"{platform.platform()} ({platform.machine()})"]


def fetch_network_info(wmi_client: wmi.WMI) -> List[str]:
    """fetch network adapter info with windows wmi.
    Remove the private ip address (like: 127.0.0.1, 192.168.x.x)"""
    def is_public_network_address(address: str) -> bool:
        address_lst = address.split(".")
        return (address_lst[0] != "127"
                and address_lst[0] != "192"
                and address_lst[1] != "168")

    results = []
    for adapter in wmi_client.Win32_NetworkAdapterConfiguration(IPEnabled=True):
        if is_public_network_address(adapter.IPAddress[0]):
            msg = f"""\
            网卡: {adapter.Description}
              IP 地址: {adapter.IPAddress[0]}
              MAC 地址: {adapter.MACAddress}"""

            results.append(textwrap.dedent(msg))

    return results


def fetch_ie_version() -> List[str]:
    """fetch windows ie version from registry"""
    k = reg.OpenKey(reg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Internet Explorer", 0, reg.KEY_READ)
    v = reg.QueryValueEx(k, "Version")
    return [v[0]]


w = wmi.WMI()

starter = AutoStarter(w)

main = MainPanel(starter)
main.add_info_list("操作系统类型", fetch_os())
main.add_info_list("磁盘序列号", fetch_disk_info(w))
main.add_info_list("网卡适配器", fetch_network_info(w))
main.add_info_list("IE版本", fetch_ie_version())

main.start()
