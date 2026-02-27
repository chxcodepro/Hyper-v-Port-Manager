"""
配置管理模块
"""
import json
import os

CONFIG_FILE = "config.json"


def get_config_path():
    """获取配置文件路径（与exe同目录）"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe
        base_path = os.path.dirname(sys.executable)
    else:
        # 开发环境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, CONFIG_FILE)


import sys


def load_config():
    """加载配置"""
    default_config = {
        "dynamic_port_start": 49152,
        "dynamic_port_count": 16384,
        "protected_ports": [],
        "last_random_range": None
    }

    config_path = get_config_path()
    try:
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 合并默认配置
                for key in default_config:
                    if key not in config:
                        config[key] = default_config[key]
                return config
    except Exception as e:
        print(f"加载配置失败: {e}")

    return default_config


def save_config(config):
    """保存配置"""
    config_path = get_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"保存配置失败: {e}")
        return False
