from __future__ import annotations

import os
import shutil
import time
import stat
from dataclasses import dataclass
from typing import Callable, Optional
from datetime import datetime


def write_text(usb_root: str, relative_path: str, text: str, encoding: str = "utf-8") -> str:
    target = os.path.join(usb_root, relative_path)
    os.makedirs(os.path.dirname(target) or usb_root, exist_ok=True)
    with open(target, "w", encoding=encoding) as f:
        f.write(text)
    return target


def delete_path(usb_root: str, relative_path: str) -> str:
    target = os.path.join(usb_root, relative_path)
    if os.path.isdir(target):
        shutil.rmtree(target)
    else:
        os.remove(target)
    return target


def list_files(drive_path: str, show_hidden: bool = True) -> list[dict]:
    """
    列出指定驱动器路径下的所有文件和目录。
    """
    files = []
    try:
        # 使用 scandir 获取更详细的文件信息
        with os.scandir(drive_path) as it:
            for entry in it:
                # 判断是否为隐藏文件
                is_hidden = False
                try:
                    if os.name == 'nt':
                        is_hidden = bool(entry.stat().st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)
                    else:
                        is_hidden = entry.name.startswith('.')
                except (OSError, AttributeError):
                    pass

                # 过滤隐藏文件
                if not show_hidden and is_hidden:
                    continue

                try:
                    info = entry.stat()
                    # 格式化时间
                    mtime = datetime.fromtimestamp(info.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

                    files.append({
                        'name': entry.name,
                        'path': entry.path,
                        'size': info.st_size if not entry.is_dir() else 0,
                        'is_dir': entry.is_dir(),
                        'is_hidden': is_hidden,
                        'modified': mtime
                    })
                except (OSError, PermissionError):
                    continue

    except (FileNotFoundError, PermissionError, OSError):
        pass

    # 排序：文件夹在前，文件在后
    files.sort(key=lambda x: (not x['is_dir'], x['name'].lower()))
    return files


@dataclass
class CopyProgress:
    bytes_copied: int
    total_bytes: int
    speed_bps: float


def copy_with_progress(
        src_file: str,
        dst_file: str,
        chunk_size: int = 1024 * 1024,
        on_progress: Optional[Callable[[CopyProgress], None]] = None,
) -> None:
    total = os.path.getsize(src_file)
    copied = 0
    t0 = time.time()

    os.makedirs(os.path.dirname(dst_file) or ".", exist_ok=True)

    with open(src_file, "rb") as fsrc, open(dst_file, "wb") as fdst:
        while True:
            chunk = fsrc.read(chunk_size)
            if not chunk:
                break
            fdst.write(chunk)
            copied += len(chunk)

            dt = max(time.time() - t0, 1e-6)
            speed = copied / dt
            if on_progress:
                on_progress(CopyProgress(bytes_copied=copied, total_bytes=total, speed_bps=speed))