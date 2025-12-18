from __future__ import annotations

import threading
from dataclasses import dataclass
from typing import Callable, Optional

import pythoncom
import win32com.client


@dataclass(frozen=True)
class DriveEvent:
    action: str  # "inserted" | "removed"
    drive_letter: str  # e.g. "G:"


def get_removable_drives() -> list[str]:
    """
    WMI 查询当前可移动盘（DriveType=2）
    返回如 ["G:", "H:"]（不带反斜杠）
    """
    pythoncom.CoInitialize()
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        items = wmi.ExecQuery("SELECT DeviceID FROM Win32_LogicalDisk WHERE DriveType = 2")
        return [i.DeviceID for i in items]
    finally:
        pythoncom.CoUninitialize()


class WmiDriveEventWatcher:
    """
    WMI 事件监听：Win32_VolumeChangeEvent
    为了减少 pywin32 “releasing IUnknown” 红字：
    - stop() 时通过 pythoncom.CoCancelCall 取消阻塞等待
    - 显式释放 COM 引用
    - join() 等待线程退出
    """

    def __init__(self, on_event: Callable[[DriveEvent], None]):
        self.on_event = on_event
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._thread_id: Optional[int] = None

        # 线程内 COM 引用（用于显式释放）
        self._service = None
        self._watcher = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="WmiDriveEventWatcher", daemon=True)
        self._thread.start()

    def stop(self, join_timeout_sec: float = 2.0) -> None:
        self._stop.set()

        # 尝试取消线程里 NextEvent 的阻塞等待（如果正在等待）
        if self._thread_id is not None:
            try:
                pythoncom.CoCancelCall(self._thread_id, 0)
            except Exception:
                # 某些情况下会失败（例如不在等待），忽略即可
                pass

        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=join_timeout_sec)

        # 主线程侧也清掉引用（双保险）
        self._service = None
        self._watcher = None
        self._thread = None
        self._thread_id = None

    def _run(self) -> None:
        pythoncom.CoInitialize()
        self._thread_id = threading.get_native_id()

        try:
            locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            service = locator.ConnectServer(".", "root\\cimv2")
            self._service = service

            query = "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2 OR EventType = 3"
            watcher = service.ExecNotificationQuery(query)
            self._watcher = watcher

            while not self._stop.is_set():
                try:
                    # 1000ms 超时：便于定期检查 stop
                    evt = watcher.NextEvent(1000)
                except Exception:
                    # 可能是超时/取消/临时错误，继续循环
                    continue

                drive_name = getattr(evt, "DriveName", None)  # e.g. "G:\"
                if not drive_name or len(drive_name) < 2:
                    continue
                drive_letter = drive_name[:2]

                event_type = int(getattr(evt, "EventType", 0))
                if event_type == 2:
                    action = "inserted"
                elif event_type == 3:
                    action = "removed"
                else:
                    continue

                # 插入过滤：避免光驱/网络盘；拔出不过滤（盘符可能已不存在）
                if action == "inserted":
                    try:
                        items = service.ExecQuery(
                            f"SELECT DriveType FROM Win32_LogicalDisk WHERE DeviceID='{drive_letter}'"
                        )
                        items = list(items)
                        if not items or int(items[0].DriveType) != 2:
                            continue
                    except Exception:
                        continue

                self.on_event(DriveEvent(action=action, drive_letter=drive_letter))

        finally:
            # 显式释放 COM 引用，减少“退出时释放”触发的噪声
            self._watcher = None
            self._service = None
            pythoncom.CoUninitialize()