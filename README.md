<div align="center">

# Windows Port Manager

Hyper-V / WSL 端口预留管理工具（Windows 10/11）

[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue?logo=windows)](https://www.microsoft.com/windows)
[![Python](https://img.shields.io/badge/Python-3.8+-yellow?logo=python)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

</div>



## 功能

- 查看当前被系统预留的端口范围
- 设置动态端口范围（需重启）
- 一键修复常用端口（49152-65535，需重启）
- 端口保护：添加/删除管理员排除（立即生效）
- Hyper-V / WSL 启用或禁用（需重启）
- 保存配置到 `config.json`





![](https://github.com/chxcodepro/windows-port-keeper/blob/master/img/SnowShot_2026-02-27_10-04-53.png?raw=true)


## 安装

### 下载发行版（推荐）

前往 [Releases](https://github.com/yourusername/windows-port-manager/releases) 下载最新版本，可选两种包：

- `PortManager-Portable-<版本号>.zip`：便携版，解压后直接运行 `PortManager.exe`
- `PortManager_Setup.exe`：安装版，按向导安装后运行

### 从源码运行

```bash
# 克隆仓库
git clone https://github.com/yourusername/windows-port-manager.git
cd windows-port-manager

# 运行（需要管理员权限）
python main.py
```

### 自行打包（可选）

```bash
build.bat
```

## 快速开始

### 释放常用开发端口

1. 以管理员身份运行程序
2. 点击 `一键修复常用端口`
3. 重启电脑

### 保护指定端口

1. 在 `端口保护` 输入 `3000` 或 `3000-3010`
2. 点击 `添加保护`
3. 立即生效，无需重启

## 原理说明

Windows 会使用动态端口范围给临时连接分配端口。Hyper-V 和 WSL 开启后，常见会占用低位端口区间，导致开发端口冲突。

本工具主要调用以下命令：

```powershell
# 查看被预留的端口
netsh interface ipv4 show excludedportrange protocol=tcp

# 设置动态端口范围（需重启）
netsh interface ipv4 set dynamic tcp start=49152 num=16384

# 添加管理员端口排除（立即生效）
netsh interface ipv4 add excludedportrange protocol=tcp startport=3000 numberofports=10
```

## 常见问题

- 为什么改动态端口后要重启：Windows 网络栈限制，重启后才会完全生效。
- 管理员排除和系统预留区别：管理员排除是你手动加的，可删；系统预留由系统服务管理。
- 禁用 Hyper-V 仍有占用：可能还有 WSL2 或其他虚拟化服务在占用。

## 许可证

[MIT License](LICENSE)
