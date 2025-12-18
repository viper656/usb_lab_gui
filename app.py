from __future__ import annotations

import getpass
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from file_ops import copy_with_progress, delete_path, write_text
from storage_monitor import WmiDriveEventWatcher, get_removable_drives
from usb_info import list_usb_devices


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("USB 总线与挂载设备测试（Windows/WMI事件版）")
        self.geometry("980x620")

        self.selected_usb_mount = tk.StringVar(value="")
        self.only_storage_var = tk.BooleanVar(value=True)

        self._refresh_timer_id = None

        self._build_ui()
        self._refresh_user()
        self._refresh_usb_devices()
        self._refresh_mounts()

        self.watcher = WmiDriveEventWatcher(on_event=self._on_drive_event_from_worker)
        self.watcher.start()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        # 关键：先停 watcher 并 join，让 COM 对象在线程内正确释放
        try:
            self.watcher.stop(join_timeout_sec=2.0)
        except Exception:
            pass
        self.destroy()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="当前登录用户：").pack(side="left")
        self.user_label = ttk.Label(top, text="(unknown)")
        self.user_label.pack(side="left")

        ttk.Button(top, text="刷新USB设备", command=self._refresh_usb_devices).pack(side="right")
        ttk.Button(top, text="刷新U盘列表", command=self._refresh_mounts).pack(side="right", padx=(0, 8))

        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=8)

        left = ttk.Frame(main)
        main.add(left, weight=1)

        header = ttk.Frame(left)
        header.pack(fill="x")
        ttk.Label(header, text="USB 设备列表").pack(side="left")

        ttk.Checkbutton(
            header,
            text="只显示U盘/存储设备(USBSTOR)",
            variable=self.only_storage_var,
            command=self._refresh_usb_devices,
        ).pack(side="right")

        cols = ("vid", "pid", "manufacturer", "product", "serial", "usb", "bus", "addr")
        self.usb_tree = ttk.Treeview(left, columns=cols, show="headings", height=14)
        headings = {
            "vid": "VID",
            "pid": "PID",
            "manufacturer": "制造商",
            "product": "产品",
            "serial": "序列号",
            "usb": "USB版本(bcd)",
            "bus": "Bus",
            "addr": "Addr",
        }
        for c in cols:
            self.usb_tree.heading(c, text=headings[c])
            self.usb_tree.column(c, width=120 if c in ("manufacturer", "product", "serial") else 90, anchor="w")
        self.usb_tree.pack(fill="both", expand=True, pady=(4, 8))

        right = ttk.Frame(main)
        main.add(right, weight=1)

        ttk.Label(right, text="U 盘（可移动盘符）").pack(anchor="w")

        sel_frame = ttk.Frame(right)
        sel_frame.pack(fill="x", pady=(4, 8))

        ttk.Label(sel_frame, text="选择盘符：").pack(side="left")
        self.mount_combo = ttk.Combobox(sel_frame, textvariable=self.selected_usb_mount, state="readonly")
        self.mount_combo.pack(side="left", fill="x", expand=True, padx=(8, 8))

        ttk.Button(sel_frame, text="打开目录", command=self._open_mount_dir).pack(side="right")

        ops = ttk.LabelFrame(right, text="基本操作")
        ops.pack(fill="x", pady=(0, 8))

        write_frame = ttk.Frame(ops)
        write_frame.pack(fill="x", padx=8, pady=6)
        ttk.Label(write_frame, text="写入文本到(相对路径)：").pack(side="left")
        self.write_rel = ttk.Entry(write_frame)
        self.write_rel.insert(0, "hello.txt")
        self.write_rel.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(write_frame, text="写入", command=self._write_text).pack(side="right")

        copy_frame = ttk.Frame(ops)
        copy_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(copy_frame, text="选择源文件并拷入U盘…", command=self._copy_file).pack(side="left")
        self.copy_progress = ttk.Label(copy_frame, text="进度：-")
        self.copy_progress.pack(side="left", padx=(12, 0))

        del_frame = ttk.Frame(ops)
        del_frame.pack(fill="x", padx=8, pady=6)
        ttk.Label(del_frame, text="删除(相对路径)：").pack(side="left")
        self.del_rel = ttk.Entry(del_frame)
        self.del_rel.insert(0, "hello.txt")
        self.del_rel.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(del_frame, text="删除", command=self._delete_path).pack(side="right")

        ttk.Label(right, text="日志").pack(anchor="w")
        self.log = tk.Text(right, height=12)
        self.log.pack(fill="both", expand=True, pady=(4, 0))

    def _log(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def _refresh_user(self):
        self.user_label.config(text=getpass.getuser())

    def _refresh_usb_devices(self):
        for item in self.usb_tree.get_children():
            self.usb_tree.delete(item)
        try:
            devs = list_usb_devices(only_storage=self.only_storage_var.get())
            for d in devs:
                self.usb_tree.insert(
                    "",
                    "end",
                    values=(
                        d.get("vendor_id"),
                        d.get("product_id"),
                        d.get("manufacturer"),
                        d.get("product"),
                        d.get("serial_number"),
                        d.get("usb_version_bcd"),
                        d.get("bus"),
                        d.get("address"),
                    ),
                )
            self._log(f"USB设备刷新完成：{len(devs)} 个设备")
        except Exception as e:
            self._log(f"USB设备刷新失败：{e}")
            messagebox.showerror("错误", f"USB设备刷新失败：\n{e}")

    def _refresh_mounts(self):
        drives = get_removable_drives()
        values = [d + "\\" for d in drives]
        self.mount_combo["values"] = values
        if values and self.selected_usb_mount.get() not in values:
            self.selected_usb_mount.set(values[0])
        if not values:
            self.selected_usb_mount.set("")
        self._log(f"U盘盘符刷新完成：{len(values)} 个")

    def _on_drive_event_from_worker(self, evt):
        self.after(0, lambda: self._handle_drive_event(evt.action, evt.drive_letter))

    def _handle_drive_event(self, action: str, drive_letter: str):
        mount = drive_letter + "\\"
        if action == "inserted":
            msg = f"检测到U盘插入：{mount}"
            self._log("[插入] " + msg)
            messagebox.showinfo("U盘插入", msg)
            threading.Thread(target=self._wait_ready_then_refresh, args=(mount,), daemon=True).start()
        elif action == "removed":
            msg = f"检测到U盘拔出：{mount}"
            self._log("[拔出] " + msg)
            messagebox.showwarning("U盘拔出", msg)
            self._schedule_single_refresh()

    def _wait_ready_then_refresh(self, mount: str):
        deadline = time.time() + 2.0
        while time.time() < deadline:
            if os.path.isdir(mount):
                break
            time.sleep(0.1)
        self.after(0, self._schedule_single_refresh)

    def _schedule_single_refresh(self):
        if self._refresh_timer_id is not None:
            self.after_cancel(self._refresh_timer_id)
            self._refresh_timer_id = None
        self._refresh_timer_id = self.after(200, self._do_refresh_after_event)

    def _do_refresh_after_event(self):
        self._refresh_timer_id = None
        self._refresh_mounts()
        self._refresh_usb_devices()

    def _require_mount(self) -> str:
        mp = self.selected_usb_mount.get()
        if not mp:
            raise RuntimeError("未选择U盘盘符，请先插入U盘并选择。")
        if not os.path.isdir(mp):
            raise RuntimeError(f"盘符不可用：{mp}")
        return mp

    def _open_mount_dir(self):
        try:
            mp = self._require_mount()
            if os.name == "nt":
                os.startfile(mp)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def _write_text(self):
        try:
            mp = self._require_mount()
            rel = self.write_rel.get().strip()
            if not rel:
                raise RuntimeError("相对路径不能为空。")
            target = write_text(mp, rel, "Hello USB!\n这是一段写入U盘的测试文本。\n")
            self._log(f"写入完成：{target}")
        except Exception as e:
            self._log(f"写入失败：{e}")
            messagebox.showerror("错误", str(e))

    def _copy_file(self):
        try:
            mp = self._require_mount()
            src = filedialog.askopenfilename(title="选择要拷入U盘的源文件")
            if not src:
                return
            dst = os.path.join(mp, os.path.basename(src))
            self.copy_progress.config(text="进度：0%")

            def worker():
                try:
                    def on_p(p):
                        pct = int((p.bytes_copied / max(p.total_bytes, 1)) * 100)
                        self.after(
                            0,
                            lambda: self.copy_progress.config(
                                text=f"进度：{pct}%  速度：{p.speed_bps/1024/1024:.2f} MB/s"
                            ),
                        )

                    copy_with_progress(src, dst, on_progress=on_p)
                    self.after(0, lambda: self._log(f"拷贝完成：{src} -> {dst}"))
                except Exception as e:
                    self.after(0, lambda: self._log(f"拷贝失败：{e}"))
                    self.after(0, lambda: messagebox.showerror("错误", str(e)))

            threading.Thread(target=worker, daemon=True).start()
        except Exception as e:
            self._log(f"拷贝启动失败：{e}")
            messagebox.showerror("错误", str(e))

    def _delete_path(self):
        try:
            mp = self._require_mount()
            rel = self.del_rel.get().strip()
            if not rel:
                raise RuntimeError("相对路径不能为空。")
            if not messagebox.askyesno("确认删除", f"确定删除 U盘中的：\n{rel}\n吗？"):
                return
            target = delete_path(mp, rel)
            self._log(f"删除完成：{target}")
        except Exception as e:
            self._log(f"删除失败：{e}")
            messagebox.showerror("错误", str(e))


if __name__ == "__main__":
    App().mainloop()