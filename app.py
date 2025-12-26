from __future__ import annotations

import getpass
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from file_ops import copy_with_progress, delete_path, write_text, list_files
from storage_monitor import WmiDriveEventWatcher, get_removable_drives
from usb_info import list_usb_devices


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("USB 总线与挂载设备测试（Windows/WMI事件版）")
        self.geometry("980x760")  # 稍微加高一点以适应内容

        self.selected_usb_mount = tk.StringVar(value="")
        # 默认只显示存储设备
        self.only_storage_var = tk.BooleanVar(value=True)
        self.show_hidden_var = tk.BooleanVar(value=True)

        self._refresh_timer_id = None

        self._build_ui()
        self._refresh_user()

        # 初始刷新
        self._refresh_usb_devices()
        self._refresh_mounts()
        self._refresh_file_list()

        # 绑定盘符变化事件，自动刷新文件列表
        self.selected_usb_mount.trace('w', lambda *args: self._refresh_file_list())

        self.watcher = WmiDriveEventWatcher(on_event=self._on_drive_event_from_worker)
        self.watcher.start()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
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

        # === 左侧区域：USB 设备总线列表 ===
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
            width = 120 if c in ("manufacturer", "product", "serial") else 80
            self.usb_tree.column(c, width=width, anchor="w")

        # 添加滚动条给左侧列表
        usb_scroll = ttk.Scrollbar(left, orient="vertical", command=self.usb_tree.yview)
        self.usb_tree.configure(yscrollcommand=usb_scroll.set)
        usb_scroll.pack(side="right", fill="y")
        self.usb_tree.pack(side="left", fill="both", expand=True, pady=(4, 8))

        # === 右侧区域：操作与文件 ===
        right = ttk.Frame(main)
        main.add(right, weight=2)

        # 1. 文件列表区域
        file_list_frame = ttk.LabelFrame(right, text="U 盘文件列表")
        file_list_frame.pack(fill="both", expand=True, pady=(0, 8))

        file_columns = ('name', 'size', 'type', 'modified', 'hidden')
        self.file_tree = ttk.Treeview(
            file_list_frame,
            columns=file_columns,
            show='headings',
            height=8
        )

        self.file_tree.heading('name', text='文件名')
        self.file_tree.heading('size', text='大小')
        self.file_tree.heading('type', text='类型')
        self.file_tree.heading('modified', text='修改时间')
        self.file_tree.heading('hidden', text='隐藏')

        self.file_tree.column('name', width=180)
        self.file_tree.column('size', width=70, anchor="e")
        self.file_tree.column('type', width=60)
        self.file_tree.column('modified', width=130)
        self.file_tree.column('hidden', width=40, anchor="center")

        tree_scroll = ttk.Scrollbar(file_list_frame, orient='vertical', command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side='right', fill='y')
        self.file_tree.pack(side='left', fill='both', expand=True)

        file_controls = ttk.Frame(file_list_frame)
        file_controls.pack(fill='x', padx=5, pady=5)

        ttk.Button(file_controls, text="刷新列表", command=self._refresh_file_list).pack(side='left', padx=2)
        ttk.Checkbutton(
            file_controls,
            text='显示隐藏文件',
            variable=self.show_hidden_var,
            command=self._refresh_file_list
        ).pack(side='left', padx=10)

        # 2. 进度条区域
        progress_frame = ttk.LabelFrame(right, text="文件传输进度")
        progress_frame.pack(fill="x", pady=(0, 8))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill='x', padx=10, pady=(10, 5))

        progress_info_frame = ttk.Frame(progress_frame)
        progress_info_frame.pack(fill='x', padx=10, pady=(0, 10))

        self.progress_text = ttk.Label(progress_info_frame, text="等待操作...")
        self.progress_text.pack(side='left')

        self.speed_label = ttk.Label(progress_info_frame, text="")
        self.speed_label.pack(side='left', padx=(10, 0))

        self.remaining_label = ttk.Label(progress_info_frame, text="")
        self.remaining_label.pack(side='left', padx=(10, 0))

        # 3. 操作区域
        ttk.Label(right, text="U 盘操作").pack(anchor="w")
        sel_frame = ttk.Frame(right)
        sel_frame.pack(fill="x", pady=(4, 8))

        ttk.Label(sel_frame, text="选择盘符：").pack(side="left")
        self.mount_combo = ttk.Combobox(sel_frame, textvariable=self.selected_usb_mount, state="readonly")
        self.mount_combo.pack(side="left", fill="x", expand=True, padx=(8, 8))

        ttk.Button(sel_frame, text="打开目录", command=self._open_mount_dir).pack(side="right")

        ops = ttk.LabelFrame(right, text="基本操作")
        ops.pack(fill="x", pady=(0, 8))

        # 写入
        write_frame = ttk.Frame(ops)
        write_frame.pack(fill="x", padx=8, pady=6)
        ttk.Label(write_frame, text="写入文本到(相对路径)：").pack(side="left")
        self.write_rel = ttk.Entry(write_frame)
        self.write_rel.insert(0, "hello.txt")
        self.write_rel.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(write_frame, text="写入", command=self._write_text).pack(side="right")

        # 拷贝
        copy_frame = ttk.Frame(ops)
        copy_frame.pack(fill="x", padx=8, pady=6)
        ttk.Button(copy_frame, text="选择源文件并拷入U盘…", command=self._copy_file).pack(side="left")

        # 删除
        del_frame = ttk.Frame(ops)
        del_frame.pack(fill="x", padx=8, pady=6)
        ttk.Label(del_frame, text="删除(相对路径)：").pack(side="left")
        self.del_rel = ttk.Entry(del_frame)
        self.del_rel.insert(0, "hello.txt")
        self.del_rel.pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(del_frame, text="删除", command=self._delete_path).pack(side="right")

        # 日志
        ttk.Label(right, text="日志").pack(anchor="w", pady=(5, 0))
        self.log = tk.Text(right, height=6)
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
            # 这里的 logic 是：only_storage_var.get() 返回 True/False
            # usb_info.list_usb_devices 内部会根据这个 bool 值过滤
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
            filter_status = " (仅存储)" if self.only_storage_var.get() else " (全部)"
            self._log(f"USB设备刷新完成：{len(devs)} 个设备{filter_status}")
        except Exception as e:
            self._log(f"USB设备刷新失败：{e}")
            messagebox.showerror("错误", f"USB设备刷新失败：\n{e}", parent=self)

    def _refresh_mounts(self):
        drives = get_removable_drives()
        values = [d + "\\" for d in drives]
        self.mount_combo["values"] = values

        current = self.selected_usb_mount.get()
        if values:
            if current not in values:
                self.selected_usb_mount.set(values[0])
        else:
            self.selected_usb_mount.set("")

        self._log(f"U盘盘符刷新完成：{len(values)} 个")
        # 刷新文件列表
        self._refresh_file_list()

    def _refresh_file_list(self, event=None):
        """刷新文件列表"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        mount = self.selected_usb_mount.get()
        if not mount or not os.path.isdir(mount):
            return

        try:
            files = list_files(mount, self.show_hidden_var.get())
            for f in files:
                f_type = '文件夹' if f['is_dir'] else '文件'

                # 转换大小显示
                size_val = f['size']
                if not f['is_dir']:
                    if size_val < 1024:
                        size_str = f"{size_val} B"
                    elif size_val < 1024 * 1024:
                        size_str = f"{size_val / 1024:.1f} KB"
                    else:
                        size_str = f"{size_val / (1024 * 1024):.1f} MB"
                else:
                    size_str = ""

                f_hidden = '√' if f['is_hidden'] else ''

                self.file_tree.insert(
                    '',
                    'end',
                    values=(f['name'], size_str, f_type, f['modified'], f_hidden)
                )

            # 更新日志状态
            # self._log(f"文件列表已更新")

        except Exception as e:
            self._log(f"刷新文件列表失败：{e}")

    def _on_drive_event_from_worker(self, evt):
        self.after(0, lambda: self._handle_drive_event(evt.action, evt.drive_letter))

    def _handle_drive_event(self, action: str, drive_letter: str):
        mount = drive_letter + "\\"
        if action == "inserted":
            msg = f"检测到U盘插入：{mount}"
            self._log("[插入] " + msg)
            messagebox.showinfo("U盘插入", msg, parent=self)
            threading.Thread(target=self._wait_ready_then_refresh, args=(mount,), daemon=True).start()
        elif action == "removed":
            msg = f"检测到U盘拔出：{mount}"
            self._log("[拔出] " + msg)
            messagebox.showwarning("U盘拔出", msg, parent=self)
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
        self._refresh_file_list()

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
                os.startfile(mp)
        except Exception as e:
            messagebox.showerror("错误", str(e), parent=self)

    def _write_text(self):
        try:
            mp = self._require_mount()
            rel = self.write_rel.get().strip()
            if not rel:
                raise RuntimeError("相对路径不能为空。")
            target = write_text(mp, rel, "Hello USB!\n这是一段写入U盘的测试文本。\n")
            self._log(f"写入完成：{target}")
            self._refresh_file_list()
        except Exception as e:
            self._log(f"写入失败：{e}")
            messagebox.showerror("错误", str(e), parent=self)

    def _copy_file(self):
        try:
            mp = self._require_mount()
            # 显式指定 parent=self，防止出现空白窗口
            src = filedialog.askopenfilename(title="选择要拷入U盘的源文件", parent=self)
            if not src:
                return
            dst = os.path.join(mp, os.path.basename(src))

            # 初始化UI
            self.progress_var.set(0)
            self.progress_text.config(text=f"正在复制: {os.path.basename(src)}")
            self.speed_label.config(text=" | 速率: -- MB/s")
            self.remaining_label.config(text=" | 剩余: --")
            self.progress_bar.config(mode='determinate', style="")

            def worker():
                try:
                    start_time = time.time()
                    last_update_time = start_time
                    last_copied = 0

                    def on_p(p):
                        nonlocal last_update_time, last_copied

                        current_time = time.time()
                        pct = int((p.bytes_copied / max(p.total_bytes, 1)) * 100)

                        # 降低刷新频率，避免卡顿
                        if current_time - last_update_time >= 0.1 or pct >= 100:
                            # 计算速度
                            time_diff = max(current_time - last_update_time, 0.001)
                            bytes_diff = p.bytes_copied - last_copied

                            # 瞬时速度
                            speed_mbps = (bytes_diff / time_diff) / (1024 * 1024)
                            # 平均速度（用于计算剩余时间更准）
                            total_time = current_time - start_time
                            avg_speed = (p.bytes_copied / total_time) / (1024 * 1024) if total_time > 0 else 0

                            # 估算剩余时间
                            rem_time_str = "--"
                            if avg_speed > 0 and pct < 100:
                                rem_bytes = p.total_bytes - p.bytes_copied
                                rem_sec = rem_bytes / (avg_speed * 1024 * 1024)
                                if rem_sec < 60:
                                    rem_time_str = f"{rem_sec:.0f}秒"
                                else:
                                    rem_time_str = f"{rem_sec / 60:.1f}分"

                            # 线程安全更新 UI
                            self.after(0, lambda: self._update_progress_ui(
                                pct, speed_mbps, avg_speed, rem_time_str
                            ))

                            last_update_time = current_time
                            last_copied = p.bytes_copied

                    copy_with_progress(src, dst, on_progress=on_p)

                    # 成功
                    self.after(0, lambda: self._copy_complete(src, dst))

                except Exception as e:
                    # [关键修复] 将异常转换为字符串，确保 lambda 绑定的是值而不是引用
                    err_msg = str(e)
                    self.after(0, lambda: self._copy_failed(err_msg))

            threading.Thread(target=worker, daemon=True).start()

        except Exception as e:
            self._log(f"拷贝启动失败：{e}")
            messagebox.showerror("错误", str(e), parent=self)

    def _update_progress_ui(self, percent, instant_speed, avg_speed, remaining):
        self.progress_var.set(percent)
        self.speed_label.config(text=f" | {instant_speed:.1f} MB/s")
        self.remaining_label.config(text=f" | 剩余: {remaining}")

    def _copy_complete(self, src, dst):
        self.progress_text.config(text="复制完成!")
        self.progress_var.set(100)
        self.speed_label.config(text="")
        self.remaining_label.config(text="")
        self.progress_bar.config(style="green.Horizontal.TProgressbar")

        self._log(f"拷贝完成：{src} -> {dst}")
        self._refresh_file_list()

        # 3秒后重置
        self.after(3000, self._reset_progress)

    def _copy_failed(self, error_msg):
        """处理复制失败"""
        self.progress_text.config(text="复制失败!")
        self.progress_bar.config(style="red.Horizontal.TProgressbar")
        self._log(f"拷贝失败：{error_msg}")
        messagebox.showerror("错误", f"文件复制失败：\n{error_msg}", parent=self)

        self.after(3000, self._reset_progress)

    def _reset_progress(self):
        self.progress_var.set(0)
        self.progress_text.config(text="等待操作...")
        self.speed_label.config(text="")
        self.remaining_label.config(text="")
        self.progress_bar.config(style="")

    def _delete_path(self):
        try:
            mp = self._require_mount()
            rel = self.del_rel.get().strip()
            if not rel:
                raise RuntimeError("相对路径不能为空。")
            if not messagebox.askyesno("确认删除", f"确定删除 U盘中的：\n{rel}\n吗？", parent=self):
                return
            target = delete_path(mp, rel)
            self._log(f"删除完成：{target}")
            self._refresh_file_list()
        except Exception as e:
            self._log(f"删除失败：{e}")
            messagebox.showerror("错误", str(e), parent=self)


if __name__ == "__main__":
    # 配置进度条样式
    try:
        # DPI 适配（可选）
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    app = App()

    style = ttk.Style()
    style.configure("green.Horizontal.TProgressbar", background='green')
    style.configure("red.Horizontal.TProgressbar", background='red')

    app.mainloop()