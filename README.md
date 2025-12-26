# USB 总线及挂载设备测试实验（Windows / Python / GUI）

本项目用于完成《汇编语言与接口技术》课程大作业中的 **USB 总线及挂载设备测试实验**：在 Windows 平台上通过 GUI 检测 USB 设备与 U 盘插拔，并对 U 盘进行读写/拷贝/删除等操作。

> 当前实现目标：先在 Windows 上完成“必选功能”的可运行版本；后续可扩展 Linux 版本与更多展示功能。

---

## 1. 当前已实现功能

### 1) GUI 界面
- 使用 **Tkinter**（Python 自带 GUI 库）构建窗口界面

### 2) 显示当前登录用户
- 界面顶部显示 Windows 当前登录用户名

### 3) 检测 USB 设备信息（面向实验：重点显示 U 盘）
- 左侧显示 USB 设备列表
- **默认只显示 USB 存储设备（USBSTOR）**：也就是 U 盘/移动硬盘这类“U盘相关设备”
- 可通过勾选框切换（扩展时可做成“显示全部 USB 设备”）

> 说明：本项目采用 **Windows 原生 WMI** 获取 USB 设备信息（不依赖 pyusb/libusb 后端），稳定可靠，适合课程作业展示。

### 4) 检测 U 盘插入 / 拔出（实时）
- 使用 **WMI 事件监听（Win32_VolumeChangeEvent）** 实现事件驱动检测
- 插入/拔出时弹出提示框，并写入右侧日志

### 5) 对 U 盘进行基本文件操作
- 写入简单文本（如 `hello.txt`）
- 选择文件拷贝到 U 盘（显示进度与速度）
- 删除 U 盘上的文件/目录

---

## 2. 技术方案

- **GUI**：Tkinter
- **U盘插拔检测**：WMI 事件监听（实时，低开销）
- **U盘盘符识别**：WMI `Win32_LogicalDisk` 的 `DriveType=2`（可移动盘）
- **USB 设备信息枚举**：WMI `Win32_PnPEntity`，并按 `Service=USBSTOR` 过滤显示 U 盘相关设备
- **文件拷贝速率**：应用层分块复制 + 计算吞吐（MB/s）

---

## 3. 环境要求

- Windows 10 / Windows 11
- Python 3.10+（建议）
- 需要安装依赖：
  - `psutil`
  - `pywin32`

---

## 4. 安装与运行（推荐使用虚拟环境）

在项目根目录打开 CMD：

```bat
py -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install -U pip
pip install -r requirements.txt
python app.py
```

> PyCharm 用户：请在 **Settings → Project → Python Interpreter** 选择项目的  
> `.venv\Scripts\python.exe` 作为解释器，否则会出现 “No module named psutil”等错误。

---

## 5. 项目结构说明

```
.
├─ app.py                 # 主程序：GUI + 事件响应 + 调用各模块
├─ usb_info.py            # USB设备信息枚举（WMI），默认只返回 USBSTOR（U盘类设备）
├─ storage_monitor.py     # WMI事件监听：检测U盘插入/拔出；查询可移动盘符
├─ file_ops.py            # U盘文件操作：写入文本/拷贝文件(含速率)/删除
├─ requirements.txt
├─ README.md
└─ .gitignore
```



1) **显示 U 盘文件列表（含隐藏文件）**  
- 建议加在：`file_ops.py` 新增 `list_files(root, show_hidden=True/False)`  
- GUI 展示加在：`app.py` 右侧新增 Tree/List 控件

2) **拷贝/读写时更详细的实时速率曲线、进度条**  
- 现有：`file_ops.copy_with_progress()` 已提供回调  
- 可在 `app.py` 增加 `ttk.Progressbar` + 速率曲线/历史数据

3) **USB 设备更完整字段（厂商、序列号更可靠、USB版本/速度）**
- Windows 下可扩展 WMI 关联查询：
  - `Win32_DiskDrive`（Model、InterfaceType=USB 等）
  - `Win32_DiskDriveToDiskPartition` → `Win32_LogicalDiskToPartition`（把物理U盘关联到盘符）
- Linux 下可新增 `linux_usb_info.py`（使用 `pyusb` / `lsusb -v`）

4) **只显示“当前插入的U盘”并显示其盘符/容量/文件系统**
- 建议改造方式：在 `usb_info.py` 中通过 WMI 关联把 USBSTOR 设备映射到逻辑盘符，再在 GUI 里合并显示

---



### 剩余需要添加功能
- 实现 USB 总线编号、接口编号、USB 协议版本的精确抓取,目前 usb_info.py 中的 bus、address 和 usb_version_bcd 字段都是 None
- 可加可不加：在 UI 上增加 U 盘总容量、剩余空间显示，并随文件操作动态更新。
- 可加可不加：增加“从 U 盘拷出文件”功能，并支持文件的重命名和批量删除，实现“安全弹出”功能，支持操作日志的一键导出。
---

## 6. 已知限制（可在报告中说明）

- Windows 用户态很难稳定获取“USB协议层真实传输速率”；本项目展示的是**文件拷贝的实际吞吐**，更贴近用户体验
- 部分设备可能不提供可解析的序列号/厂商字段（由系统/驱动/设备决定）

