"""
Windows 端口预留管理工具
管理 Hyper-V 和 WSL 的端口预留
"""
import subprocess
import re
import random
import ctypes
import sys


def is_admin():
    """检查是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    """请求管理员权限重新运行"""
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit()


def run_cmd(cmd, shell=True):
    """执行命令并返回输出"""
    try:
        result = subprocess.run(
            cmd,
            shell=shell,
            capture_output=True,
            text=True,
            encoding='gbk',
            errors='ignore'
        )
        return result.stdout, result.stderr, result.returncode
    except Exception as e:
        return "", str(e), 1


def get_excluded_ports():
    """获取当前被预留的端口列表"""
    stdout, stderr, code = run_cmd("netsh interface ipv4 show excludedportrange protocol=tcp")

    ports = []
    if code == 0:
        # 解析输出，提取端口范围
        lines = stdout.strip().split('\n')
        for line in lines:
            # 匹配格式: "起始端口    结束端口"
            match = re.match(r'\s*(\d+)\s+(\d+)\s*(\*)?', line)
            if match:
                start = int(match.group(1))
                end = int(match.group(2))
                is_admin_port = match.group(3) == '*'
                ports.append({
                    'start': start,
                    'end': end,
                    'is_admin': is_admin_port,
                    'count': end - start + 1
                })

    return ports, stderr if code != 0 else None


def get_dynamic_port_range():
    """获取当前动态端口范围设置"""
    stdout, stderr, code = run_cmd("netsh interface ipv4 show dynamicport tcp")

    if code == 0:
        # 解析起始端口和数量
        start_match = re.search(r'(\d+)', stdout.split('\n')[3] if len(stdout.split('\n')) > 3 else "")
        count_match = re.search(r'(\d+)', stdout.split('\n')[4] if len(stdout.split('\n')) > 4 else "")

        # 备用解析方式
        if not start_match:
            start_match = re.search(r'[Ss]tart.*?(\d+)', stdout)
        if not count_match:
            count_match = re.search(r'[Nn]um.*?(\d+)', stdout)

        start = int(start_match.group(1)) if start_match else 49152
        count = int(count_match.group(1)) if count_match else 16384

        return {'start': start, 'count': count}, None

    return None, stderr


def set_dynamic_port_range(start, count):
    """设置动态端口范围（需要重启生效）"""
    if start < 1025 or start > 65535:
        return False, "起始端口必须在 1025-65535 之间"
    if count < 255:
        return False, "端口数量至少为 255"
    if start + count > 65536:
        return False, f"端口范围超出上限，最大结束端口为 {start + count - 1}"

    cmd = f"netsh interface ipv4 set dynamic tcp start={start} num={count}"
    stdout, stderr, code = run_cmd(cmd)

    if code == 0:
        return True, "设置成功，需要重启电脑生效"
    return False, stderr or "设置失败"


def add_port_exclusion(start, end=None):
    """添加管理员端口排除（立即生效）"""
    if end is None:
        end = start

    if start < 1 or end > 65535 or start > end:
        return False, "端口范围无效"

    count = end - start + 1
    cmd = f"netsh interface ipv4 add excludedportrange protocol=tcp startport={start} numberofports={count}"
    stdout, stderr, code = run_cmd(cmd)

    if code == 0:
        return True, f"已保护端口 {start}-{end}"

    # 检查是否端口已被占用
    if "denied" in stderr.lower() or "拒绝" in stderr:
        return False, f"端口 {start}-{end} 已被占用，无法排除"

    return False, stderr or "添加失败"


def delete_port_exclusion(start, end=None):
    """删除管理员端口排除"""
    if end is None:
        end = start

    count = end - start + 1
    cmd = f"netsh interface ipv4 delete excludedportrange protocol=tcp startport={start} numberofports={count}"
    stdout, stderr, code = run_cmd(cmd)

    if code == 0:
        return True, f"已删除端口保护 {start}-{end}"
    return False, stderr or "删除失败（可能不是管理员排除的端口）"


def check_port_available(port):
    """检查端口是否可用"""
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', port))
            return True, "端口可用"
    except OSError as e:
        return False, f"端口被占用: {e}"


def check_ports_in_range(start, end):
    """检查范围内有多少端口被占用"""
    occupied = []
    for port in range(start, min(end + 1, start + 100)):  # 最多检查100个
        available, _ = check_port_available(port)
        if not available:
            occupied.append(port)
    return occupied


def get_hyperv_status():
    """获取 Hyper-V 状态"""
    stdout, stderr, code = run_cmd("dism /online /get-featureinfo /featurename:Microsoft-Hyper-V-All")

    if "enabled" in stdout.lower() or "启用" in stdout:
        return True, "已启用"
    elif "disabled" in stdout.lower() or "禁用" in stdout:
        return False, "已禁用"
    else:
        # 尝试另一种检测方式
        stdout2, _, _ = run_cmd("sc query vmms")
        if "RUNNING" in stdout2:
            return True, "已启用"
        return False, "未安装或已禁用"


def set_hyperv(enable):
    """开启或关闭 Hyper-V（需要重启）"""
    action = "Enable" if enable else "Disable"
    cmd = f"dism /online /{action}-Feature /FeatureName:Microsoft-Hyper-V-All /NoRestart"
    stdout, stderr, code = run_cmd(cmd)

    if code == 0:
        status = "启用" if enable else "禁用"
        return True, f"Hyper-V 已{status}，需要重启电脑生效"
    return False, stderr or stdout or "操作失败"


def get_wsl_status():
    """获取 WSL 状态"""
    stdout, stderr, code = run_cmd("dism /online /get-featureinfo /featurename:Microsoft-Windows-Subsystem-Linux")

    if "enabled" in stdout.lower() or "启用" in stdout:
        return True, "已启用"
    elif "disabled" in stdout.lower() or "禁用" in stdout:
        return False, "已禁用"
    return None, "无法检测"


def set_wsl(enable):
    """开启或关闭 WSL（需要重启）"""
    action = "Enable" if enable else "Disable"
    cmd = f"dism /online /{action}-Feature /FeatureName:Microsoft-Windows-Subsystem-Linux /NoRestart"
    stdout, stderr, code = run_cmd(cmd)

    if code == 0:
        status = "启用" if enable else "禁用"
        return True, f"WSL 已{status}，需要重启电脑生效"
    return False, stderr or stdout or "操作失败"


def generate_random_port_range(min_start=40000, max_start=55000, count=16384):
    """生成随机端口范围"""
    # 确保不会超出65535
    max_possible_start = 65536 - count
    max_start = min(max_start, max_possible_start)

    start = random.randint(min_start, max_start)
    return start, count


def fix_common_ports():
    """一键修复常用开发端口（3000-10000）"""
    # 将动态端口范围设置到高位
    return set_dynamic_port_range(49152, 16384)
