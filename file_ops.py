from __future__ import annotations

import os
import shutil
import time
from dataclasses import dataclass
from typing import Callable, Optional


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