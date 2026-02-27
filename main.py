"""
Windows 端口预留管理工具 - GUI界面
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading

from port_manager import (
    is_admin, run_as_admin,
    get_excluded_ports, get_dynamic_port_range, set_dynamic_port_range,
    add_port_exclusion, delete_port_exclusion,
    check_port_available,
    get_hyperv_status, set_hyperv,
    get_wsl_status, set_wsl,
    fix_common_ports
)
from config_manager import load_config, save_config


class PortManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows 端口预留管理工具")
        self.root.geometry("900x540")
        self.root.minsize(820, 480)
        self.root.resizable(True, True)

        # 加载配置
        self.config = load_config()

        # 设置样式
        self.setup_styles()

        # 创建界面
        self.create_widgets()

        # 初始加载数据
        self.refresh_all()

    def setup_styles(self):
        """设置界面样式"""
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Microsoft YaHei", 12, "bold"))
        style.configure("Status.TLabel", font=("Microsoft YaHei", 9))
        style.configure("Warning.TLabel", foreground="orange")
        style.configure("Success.TLabel", foreground="green")
        style.configure("Error.TLabel", foreground="red")

    def create_widgets(self):
        """创建所有控件"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ===== 顶部状态栏 =====
        self.create_status_bar(main_frame)

        # ===== 中部内容区（左操作区 + 右列表区）=====
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        content_frame.columnconfigure(0, weight=1, uniform="main_cols")
        content_frame.columnconfigure(1, weight=1, uniform="main_cols")
        content_frame.rowconfigure(0, weight=1)

        left_panel = ttk.Frame(content_frame)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        right_panel = ttk.Frame(content_frame)
        right_panel.grid(row=0, column=1, sticky="nsew")

        # ===== Hyper-V 和 WSL 控制区 =====
        self.create_feature_controls(left_panel)

        # ===== 动态端口范围设置 =====
        self.create_port_range_section(left_panel)

        # ===== 端口保护管理 =====
        self.create_port_protection_section(left_panel)

        # ===== 被预留端口列表 =====
        self.create_excluded_ports_list(right_panel)

        # ===== 底部按钮 =====
        self.create_bottom_buttons(main_frame)

    def create_status_bar(self, parent):
        """创建状态栏"""
        status_frame = ttk.LabelFrame(parent, text="系统状态", padding="5")
        status_frame.pack(fill=tk.X, pady=(0, 10))

        # 管理员状态
        admin_status = "✓ 管理员权限" if is_admin() else "✗ 需要管理员权限"
        admin_color = "green" if is_admin() else "red"
        self.admin_label = ttk.Label(status_frame, text=admin_status, foreground=admin_color)
        self.admin_label.pack(side=tk.LEFT, padx=10)

        # 动态端口范围显示
        self.port_range_label = ttk.Label(status_frame, text="动态端口范围: 加载中...")
        self.port_range_label.pack(side=tk.LEFT, padx=20, fill=tk.X, expand=True)

        # 右侧操作按钮
        action_frame = ttk.Frame(status_frame)
        action_frame.pack(side=tk.RIGHT)
        ttk.Button(action_frame, text="保存配置", command=self.save_current_config, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="刷新", command=self.refresh_all, width=8).pack(side=tk.LEFT, padx=2)

    def create_feature_controls(self, parent):
        """创建 Hyper-V 和 WSL 控制区"""
        feature_frame = ttk.LabelFrame(parent, text="功能开关 (需要重启生效)", padding="10")
        feature_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Hyper-V 控制
        hyperv_frame = ttk.Frame(feature_frame)
        hyperv_frame.pack(fill=tk.X, pady=2)

        ttk.Label(hyperv_frame, text="Hyper-V:", width=12).pack(side=tk.LEFT)
        self.hyperv_status_label = ttk.Label(hyperv_frame, text="检测中...", width=15)
        self.hyperv_status_label.pack(side=tk.LEFT, padx=5)

        self.hyperv_enable_btn = ttk.Button(hyperv_frame, text="启用", width=8,
                                            command=lambda: self.toggle_feature("hyperv", True))
        self.hyperv_enable_btn.pack(side=tk.LEFT, padx=2)

        self.hyperv_disable_btn = ttk.Button(hyperv_frame, text="禁用", width=8,
                                             command=lambda: self.toggle_feature("hyperv", False))
        self.hyperv_disable_btn.pack(side=tk.LEFT, padx=2)

        # WSL 控制
        wsl_frame = ttk.Frame(feature_frame)
        wsl_frame.pack(fill=tk.X, pady=2)

        ttk.Label(wsl_frame, text="WSL:", width=12).pack(side=tk.LEFT)
        self.wsl_status_label = ttk.Label(wsl_frame, text="检测中...", width=15)
        self.wsl_status_label.pack(side=tk.LEFT, padx=5)

        self.wsl_enable_btn = ttk.Button(wsl_frame, text="启用", width=8,
                                         command=lambda: self.toggle_feature("wsl", True))
        self.wsl_enable_btn.pack(side=tk.LEFT, padx=2)

        self.wsl_disable_btn = ttk.Button(wsl_frame, text="禁用", width=8,
                                          command=lambda: self.toggle_feature("wsl", False))
        self.wsl_disable_btn.pack(side=tk.LEFT, padx=2)

    def create_port_range_section(self, parent):
        """创建动态端口范围设置区"""
        range_frame = ttk.LabelFrame(parent, text="动态端口范围设置 (需要重启生效)", padding="10")
        range_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # 输入行
        input_frame = ttk.Frame(range_frame)
        input_frame.pack(fill=tk.X, pady=5)

        ttk.Label(input_frame, text="起始端口:").pack(side=tk.LEFT)
        self.start_port_var = tk.StringVar(value=str(self.config.get("dynamic_port_start", 49152)))
        self.start_port_entry = ttk.Entry(input_frame, textvariable=self.start_port_var, width=10)
        self.start_port_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(input_frame, text="数量:").pack(side=tk.LEFT, padx=(20, 0))
        self.port_count_var = tk.StringVar(value=str(self.config.get("dynamic_port_count", 16384)))
        self.port_count_entry = ttk.Entry(input_frame, textvariable=self.port_count_var, width=10)
        self.port_count_entry.pack(side=tk.LEFT, padx=5)

        # 按钮行
        btn_row1 = ttk.Frame(range_frame)
        btn_row1.pack(fill=tk.X, pady=(5, 5))
        ttk.Button(btn_row1, text="应用设置", command=self.apply_port_range).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row1, text="一键修复常用端口", command=self.fix_common).pack(side=tk.LEFT, padx=2)

        # 提示
        tip_label = ttk.Label(range_frame, text="提示: 将起始端口设为 49152 可释放 1025-49151 的常用开发端口",
                              foreground="gray")
        tip_label.pack(anchor=tk.W)

    def create_port_protection_section(self, parent):
        """创建端口保护管理区"""
        protect_frame = ttk.LabelFrame(parent, text="端口保护 (立即生效)", padding="10")
        protect_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # 输入行
        input_frame = ttk.Frame(protect_frame)
        input_frame.pack(fill=tk.X, pady=5)

        ttk.Label(input_frame, text="端口/范围:").pack(side=tk.LEFT)
        self.protect_port_var = tk.StringVar()
        self.protect_port_entry = ttk.Entry(input_frame, textvariable=self.protect_port_var, width=15)
        self.protect_port_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(input_frame, text="(如: 3000 或 3000-3010)", foreground="gray").pack(side=tk.LEFT)

        action_frame = ttk.Frame(protect_frame)
        action_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(action_frame, text="添加保护", command=self.add_protection).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="删除保护", command=self.remove_protection).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="检测端口", command=self.check_single_port).pack(side=tk.LEFT, padx=2)

    def create_excluded_ports_list(self, parent):
        """创建被预留端口列表"""
        list_frame = ttk.LabelFrame(parent, text="当前被预留的端口", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # 创建表格
        columns = ("start", "end", "count", "type")
        self.ports_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        self.ports_tree.heading("start", text="起始端口")
        self.ports_tree.heading("end", text="结束端口")
        self.ports_tree.heading("count", text="数量")
        self.ports_tree.heading("type", text="类型")

        self.ports_tree.column("start", width=100, anchor=tk.CENTER)
        self.ports_tree.column("end", width=100, anchor=tk.CENTER)
        self.ports_tree.column("count", width=80, anchor=tk.CENTER)
        self.ports_tree.column("type", width=150, minwidth=130, anchor=tk.CENTER, stretch=True)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.ports_tree.yview)
        self.ports_tree.configure(yscrollcommand=scrollbar.set)

        self.ports_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 统计标签
        self.stats_label = ttk.Label(list_frame, text="")
        self.stats_label.pack(anchor=tk.W, pady=(5, 0))

    def create_bottom_buttons(self, parent):
        """创建底部按钮"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)

        # 状态消息
        self.status_msg = ttk.Label(btn_frame, text="")
        self.status_msg.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

    # ===== 功能方法 =====

    def refresh_all(self):
        """刷新所有数据"""
        self.show_status("正在刷新...")

        def do_refresh():
            # 刷新动态端口范围
            range_info, err = get_dynamic_port_range()
            if range_info:
                self.root.after(0, lambda: self.port_range_label.config(
                    text=f"动态端口范围: {range_info['start']} - {range_info['start'] + range_info['count'] - 1}"
                ))
                self.root.after(0, lambda: self.start_port_var.set(str(range_info['start'])))
                self.root.after(0, lambda: self.port_count_var.set(str(range_info['count'])))

            # 刷新 Hyper-V 状态
            hyperv_enabled, hyperv_msg = get_hyperv_status()
            color = "green" if hyperv_enabled else "gray"
            self.root.after(0, lambda: self.hyperv_status_label.config(text=hyperv_msg, foreground=color))

            # 刷新 WSL 状态
            wsl_enabled, wsl_msg = get_wsl_status()
            color = "green" if wsl_enabled else "gray"
            self.root.after(0, lambda: self.wsl_status_label.config(text=wsl_msg, foreground=color))

            # 刷新端口列表
            ports, err = get_excluded_ports()
            self.root.after(0, lambda: self.update_ports_list(ports))

            self.root.after(0, lambda: self.show_status("刷新完成"))

        threading.Thread(target=do_refresh, daemon=True).start()

    def update_ports_list(self, ports):
        """更新端口列表显示"""
        # 清空现有数据
        for item in self.ports_tree.get_children():
            self.ports_tree.delete(item)

        total_count = 0
        admin_count = 0

        for port in ports:
            port_type = "管理员排除 *" if port['is_admin'] else "系统预留"
            self.ports_tree.insert("", tk.END, values=(
                port['start'],
                port['end'],
                port['count'],
                port_type
            ))
            total_count += port['count']
            if port['is_admin']:
                admin_count += port['count']

        self.stats_label.config(
            text=f"共 {len(ports)} 个范围，{total_count} 个端口被预留 (其中 {admin_count} 个为管理员排除)"
        )

    def toggle_feature(self, feature, enable):
        """切换 Hyper-V 或 WSL"""
        action = "启用" if enable else "禁用"
        feature_name = "Hyper-V" if feature == "hyperv" else "WSL"

        if not messagebox.askyesno("确认", f"确定要{action} {feature_name} 吗？\n此操作需要重启电脑生效。"):
            return

        self.show_status(f"正在{action} {feature_name}...")

        def do_toggle():
            if feature == "hyperv":
                success, msg = set_hyperv(enable)
            else:
                success, msg = set_wsl(enable)

            self.root.after(0, lambda: self.show_result(success, msg))
            self.root.after(0, self.refresh_all)

        threading.Thread(target=do_toggle, daemon=True).start()

    def apply_port_range(self):
        """应用端口范围设置"""
        try:
            start = int(self.start_port_var.get())
            count = int(self.port_count_var.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        if not messagebox.askyesno("确认", f"设置动态端口范围为 {start} - {start + count - 1}？\n此操作需要重启电脑生效。"):
            return

        success, msg = set_dynamic_port_range(start, count)
        self.show_result(success, msg)

        if success:
            self.config["dynamic_port_start"] = start
            self.config["dynamic_port_count"] = count
            save_config(self.config)

    def fix_common(self):
        """一键修复常用端口"""
        if not messagebox.askyesno("确认", "将动态端口范围设为 49152-65535，释放常用开发端口？\n此操作需要重启电脑生效。"):
            return

        success, msg = fix_common_ports()
        self.show_result(success, msg)

        if success:
            self.start_port_var.set("49152")
            self.port_count_var.set("16384")
            self.refresh_all()

    def add_protection(self):
        """添加端口保护"""
        port_str = self.protect_port_var.get().strip()
        if not port_str:
            messagebox.showerror("错误", "请输入端口号")
            return

        try:
            if '-' in port_str:
                start, end = map(int, port_str.split('-'))
            else:
                start = end = int(port_str)
        except ValueError:
            messagebox.showerror("错误", "格式错误，请输入如 3000 或 3000-3010")
            return

        success, msg = add_port_exclusion(start, end)
        self.show_result(success, msg)

        if success:
            if [start, end] not in self.config["protected_ports"]:
                self.config["protected_ports"].append([start, end])
                save_config(self.config)
            self.refresh_all()

    def remove_protection(self):
        """删除端口保护"""
        port_str = self.protect_port_var.get().strip()
        if not port_str:
            messagebox.showerror("错误", "请输入端口号")
            return

        try:
            if '-' in port_str:
                start, end = map(int, port_str.split('-'))
            else:
                start = end = int(port_str)
        except ValueError:
            messagebox.showerror("错误", "格式错误")
            return

        success, msg = delete_port_exclusion(start, end)
        self.show_result(success, msg)

        if success:
            if [start, end] in self.config["protected_ports"]:
                self.config["protected_ports"].remove([start, end])
                save_config(self.config)
            self.refresh_all()

    def check_single_port(self):
        """检测单个端口"""
        port_str = self.protect_port_var.get().strip()
        if not port_str:
            messagebox.showerror("错误", "请输入端口号")
            return

        try:
            port = int(port_str.split('-')[0])
        except ValueError:
            messagebox.showerror("错误", "请输入有效的端口号")
            return

        available, msg = check_port_available(port)
        if available:
            messagebox.showinfo("检测结果", f"端口 {port} 可用")
        else:
            messagebox.showwarning("检测结果", f"端口 {port} 不可用\n{msg}")

    def save_current_config(self):
        """保存当前配置"""
        try:
            self.config["dynamic_port_start"] = int(self.start_port_var.get())
            self.config["dynamic_port_count"] = int(self.port_count_var.get())
        except ValueError:
            pass

        if save_config(self.config):
            self.show_status("配置已保存")
        else:
            self.show_status("保存失败", error=True)

    def show_status(self, msg, error=False):
        """显示状态消息"""
        color = "red" if error else "black"
        self.status_msg.config(text=msg, foreground=color)

    def show_result(self, success, msg):
        """显示操作结果"""
        if success:
            messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("失败", msg)


def main():
    # 检查管理员权限
    if not is_admin():
        if messagebox.askyesno("需要管理员权限", "此工具需要管理员权限才能修改系统设置。\n是否以管理员身份重新运行？"):
            run_as_admin()
        else:
            messagebox.showwarning("警告", "部分功能可能无法使用")

    root = tk.Tk()
    app = PortManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
