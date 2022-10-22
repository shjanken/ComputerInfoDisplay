import tkinter as tk
from tkinter import StringVar, ttk
import wmi
from typing import List


class MainPanel():

    # save the lastest labelframe's row position
    current_row: int = 0

    def __init__(self, *args, **kwargs):
        self.root = tk.Tk()
        # self.title("Computer Info Panel")
        self.root.resizable(False, False)
        self.root.attributes("-alpha", 1)
        self.root.geometry(f"-100+50")
        # self.root.overrideredirect(True)

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(column=0, row=0, sticky="ewns")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

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
