import psutil
import time
import os
import win32gui
import win32con
import logging

# 配置日志记录
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# 文件日志处理器
import os
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shutdown_log.txt')
file_handler = logging.FileHandler(log_path, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# 控制台日志处理器
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logger.addHandler(file_handler)
logger.addHandler(console_handler)

def get_disk_utilization():
    disk_io_counters_start = psutil.disk_io_counters()  #系统中所有磁盘的 I/O 计数器值
    time.sleep(1)  # 等待1秒钟，让磁盘活动发生变化
    disk_io_counters_end = psutil.disk_io_counters()
    
    read_count = disk_io_counters_end.read_count - disk_io_counters_start.read_count
    write_count = disk_io_counters_end.write_count - disk_io_counters_start.write_count
    
    disk_activity = read_count + write_count
    disk_total_time = (disk_io_counters_end.read_time + disk_io_counters_end.write_time -
                       disk_io_counters_start.read_time - disk_io_counters_start.write_time)
    
    if disk_total_time > 0:
        disk_utilization = (disk_activity / disk_total_time) * 100
    else:
        disk_utilization = 0
    
    return disk_utilization

def check_disk_utilization(threshold, num_checks):
    below_threshold_count = 0
    for _ in range(num_checks):
        disk_utilization = get_disk_utilization()
        if disk_utilization < threshold:
            below_threshold_count += 1
        else:
            below_threshold_count = 0  # 重置计数器
        if below_threshold_count >= num_checks:
            return True  # 连续 num_checks 次磁盘利用率小于阈值
        time.sleep(2)
    return False

def get_system_status():
    # 获取系统状态，包括磁盘、内存和 CPU 使用情况
    disk_usage = get_disk_utilization()
    memory_usage = psutil.virtual_memory().percent
    cpu_usage = psutil.cpu_percent(interval=1)
    
    return {
        "disk_usage": disk_usage,
        "memory_usage": memory_usage,
        "cpu_usage": cpu_usage
    }

def check_unsaved_documents():
    """检查是否有未保存的文档窗口"""
    def callback(hwnd, windows):
        title = win32gui.GetWindowText(hwnd)
        if "未保存" in title or "未命名" in title:
            windows.append(title)
        return True
    
    unsaved_windows = []
    win32gui.EnumWindows(callback, unsaved_windows)
    return unsaved_windows

def confirm_shutdown():
    """用户确认关机，3分钟无操作自动确认"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        # 创建变量存储用户选择
        user_choice = tk.BooleanVar(value=False)
        
        def on_timeout():
            """3分钟超时自动确认"""
            user_choice.set(True)
            root.quit()
        
        def on_yes():
            """用户点击是"""
            user_choice.set(True)
            root.quit()
        
        def on_no():
            """用户点击否"""
            user_choice.set(False)
            try:
                root.quit()
                root.destroy()
            except:
                pass
        
        # 设置3分钟超时定时器
        root.after(180000, on_timeout)
        
        # 创建自定义弹窗
        top = tk.Toplevel(root)
        top.title("关机确认")
        top.attributes("-topmost", True)
        
        # 设置窗口大小和居中
        window_width = 400
        window_height = 200
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        top.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置窗口样式
        top.configure(bg="#f0f0f0")
        top.resizable(False, False)
        
        # 添加消息
        msg_frame = tk.Frame(top, bg="#f0f0f0")
        msg_frame.pack(pady=20)
        
        # 使用更美观的字体
        try:
            custom_font = ("Microsoft YaHei", 12)
        except:
            custom_font = ("Arial", 12)
            
        msg = tk.Label(msg_frame, 
                      text="系统检测到可以安全关机，是否现在关机？\n(3分钟内无操作将自动确认)",
                      font=custom_font,
                      bg="#f0f0f0")
        msg.pack()
        
        # 倒计时标签
        countdown_var = tk.StringVar(value="剩余时间: 180秒")
        countdown_label = tk.Label(msg_frame, 
                                 textvariable=countdown_var,
                                 font=custom_font,
                                 fg="#ff0000",
                                 bg="#f0f0f0")
        countdown_label.pack(pady=10)
        
        # 更新倒计时
        def update_countdown(remaining=180):
            if remaining > 0:
                countdown_var.set(f"剩余时间: {remaining}秒")
                top.after(1000, update_countdown, remaining-1)
        
        update_countdown()
        
        # 按钮框架
        btn_frame = tk.Frame(top, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        
        # 美化按钮
        btn_style = {
            "font": custom_font,
            "width": 8,
            "height": 1,
            "bd": 0,
            "activebackground": "#e0e0e0"
        }
        
        yes_btn = tk.Button(btn_frame, 
                           text="是", 
                           command=on_yes,
                           bg="#4CAF50",
                           fg="white",
                           **btn_style)
        yes_btn.pack(side=tk.LEFT, padx=20)
        
        no_btn = tk.Button(btn_frame, 
                          text="否", 
                          command=on_no,
                          bg="#f44336",
                          fg="white",
                          **btn_style)
        no_btn.pack(side=tk.LEFT, padx=20)
        
        try:
            # 运行主循环
            root.mainloop()
            
            if user_choice.get():
                logging.info("用户确认或超时自动确认关机")
            else:
                logging.info("用户取消了关机操作")
                
            return user_choice.get()
        finally:
            try:
                root.destroy()
            except:
                pass
        
    except ImportError:
        # 回退到普通MessageBox
        try:
            win32gui.SetForegroundWindow(win32gui.GetDesktopWindow())
        except:
            pass
        response = win32gui.MessageBox(
            0,
            "系统检测到可以安全关机，是否现在关机？",
            "关机确认",
            win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST | win32con.MB_SETFOREGROUND
        )
        return response == win32con.IDYES

def shutdown_computer(test_mode=False):
    """执行关机操作
    :param test_mode: 测试模式，不会真正关机
    """
    # 检查未保存文档
    unsaved_docs = check_unsaved_documents()
    if unsaved_docs:
        logging.warning(f"检测到未保存文档: {unsaved_docs}")
        win32gui.MessageBox(
            0,
            f"检测到以下未保存文档:\n{chr(10).join(unsaved_docs)}\n请先保存您的工作",
            "有未保存的工作",
            win32con.MB_OK | win32con.MB_ICONWARNING | win32con.MB_TOPMOST
        )
        return False
    
    # 用户确认
    if not confirm_shutdown():
        logging.info("用户取消了关机操作")
        return False
    
    if test_mode:
        logging.info("测试模式: 满足关机条件但不会真正关机")
        win32gui.MessageBox(
            0,
            "测试模式: 满足所有关机条件\n实际运行时将执行关机",
            "测试模式",
            win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
        )
        return True
    else:
        # 记录关机
        logging.info("系统正在关机...")
        os.system("shutdown /s /t 30")  # 30秒后关机，给用户缓冲时间
        return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='ShutBuddy-关机小卫士')
    parser.add_argument('--test', action='store_true', help='测试模式，不会真正关机')
    args = parser.parse_args()

    try:
        logging.info("启动关机监测程序")
        while True:
            # 设定各项资源的阈值
            num_disk_util_checks = 10  # 连续监测磁盘利用率次数 
            disk_usage_threshold = 3  # 磁盘利用率阈值
            
            if check_disk_utilization(disk_usage_threshold, num_disk_util_checks):
                status = get_system_status()
                logging.info(f"系统状态: 磁盘 {status['disk_usage']:.1f}%, 内存 {status['memory_usage']}%, CPU {status['cpu_usage']}%")
                print(f"连续 {num_disk_util_checks} 次磁盘利用率小于 {disk_usage_threshold}%")
                print("满足关机条件，正在检查系统状态...")
                
                if shutdown_computer(test_mode=args.test):
                    if args.test:
                        print("测试完成，程序退出")
                    break
            else:
                print("未满足关机条件，不执行关机操作。")
                time.sleep(10)
                
    except Exception as e:
        logging.error(f"程序发生错误: {str(e)}")
        win32gui.MessageBox(
            0,
            f"关机程序发生错误:\n{str(e)}",
            "错误",
            win32con.MB_OK | win32con.MB_ICONERROR | win32con.MB_TOPMOST
        )
