"""
coding: utf-8
©2019-2024 GZYzhy Publish under Apache License 2.0
GitHub: https://github.com/gzyzhy/Name-Random-Picker
"""

print("=== 程序开始加载模块 ===")
import json
import os
import random
print(f"[DEBUG] random模块已加载，当前种子状态: {random.getstate() is not None}")
import threading
import time
from tkinter import *
from PIL import Image, ImageTk
import sys
import platform
import subprocess
import socket
from tkinter import font
from pystray import Icon, Menu as PystrayMenu, MenuItem

def resource_path(relative_path):
    """获取打包后资源的绝对路径"""
    try:
        # PyInstaller创建的临时文件夹路径
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

import pygame
from pygame import mixer

# 初始化音频模块
pygame.init()
mixer.init()

import pyttsx4
import platform
from comtypes import CoInitialize
import chardet
from tkinter import messagebox, filedialog, simpledialog

def ensure_single_instance():
    """
    确保程序只有一个实例在运行，如果检测到已有实例，则退出当前实例
    """
    # 端口可以是任意未使用的端口
    port = 28758
    
    try:
        # 创建一个socket并绑定到本地端口
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # 添加套接字重用选项，解决端口释放后立即重用的问题
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        # 设置非阻塞模式
        sock.setblocking(False)
        sock.bind(('localhost', port))
        # 开始监听，这一步很重要，确保端口真正被占用
        sock.listen(1)
        # 如果成功绑定，说明是第一个实例
        # 保持socket打开以维持端口占用
        # 将socket保存为全局变量以防止被垃圾回收
        global single_instance_socket
        single_instance_socket = sock
        print("程序启动成功，未检测到其他实例。")
        return True
    except socket.error as e:
        # 如果端口已被占用，说明已有一个实例在运行
        print(f"程序已经在运行，此实例将关闭。错误信息: {e}")
        # 可以选择通知用户
        if platform.system() == "Windows":
            ctypes.windll.user32.MessageBoxW(0, "程序已经在运行中", "提示", 0)
        return False
    
def edit_config():
    """
    编辑配置文件的函数
    尝试运行editor.exe，如果不存在则执行editor.py
    """
    editor_exe = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "editor.exe")
    editor_py = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "editor.py")
    
    if os.path.exists(editor_exe):
        # 使用子进程执行editor.exe
        subprocess.Popen([editor_exe])
    elif os.path.exists(editor_py):
        # 使用python执行editor.py
        if platform.system() == "Windows":
            subprocess.Popen([sys.executable, editor_py])
        else:
            subprocess.Popen(['python3', editor_py])
    else:
        # 如果都不存在，提示用户
        print("未找到editor.exe或editor.py文件")
        messagebox.showerror("错误", "未找到editor.exe或editor.py文件，请检查安装完整性")


# 全局变量初始化
root = Tk()
# 设置窗口为工具窗口样式
root.overrideredirect(1)
root.attributes('-toolwindow', True)  # 设置为工具窗口
if platform.system() == 'Windows':
    import ctypes
    # 设置扩展样式为工具窗口
    hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
    GWL_EXSTYLE = -20
    WS_EX_TOOLWINDOW = 0x00000080
    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
# 获取屏幕高度并动态计算位置
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 50
window_height = 100
x = screen_width - window_width  # 改为右下角
y = screen_height - window_height - 40  # 保留底部边距
root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.iconbitmap(resource_path('favicon.ico'))
root.title("随机抽签器")
# 初始隐藏标题栏
root.overrideredirect(1)  # 替换原来的toolwindow设置
root.attributes('-alpha', 0.4)  # 启动时使用待机透明度
root.attributes('-topmost', True)
# 设置窗口背景色
root.config(bg='white')
root.attributes("-transparentcolor", "white")
config_path = "config.json"
names = []
names_use = []
groups = []
groups_use = []
now_use = 'name'
have_w = False
name = ''
egg = True
have_img = False
leave_list = []
now_move = False
auto_close_enabled = True  # 自动关闭功能默认开启（会从配置文件读取）
auto_close_timer = None  # 自动关闭定时器
window = None  # 主显示窗口
window_image = None  # 图片窗口

# 在全局变量区域添加
voice_enabled = True
error_shown = False
first_read_successful = False  # 添加新变量：标记是否已成功执行过朗读
read_queue = []  # 朗读请求队列
is_reading = False  # 是否正在朗读

# 在全局变量区域添加拖动相关变量
drag_start_x = 0
drag_start_y = 0
drag_threshold = 5  # 拖动超过5像素视为移动操作
is_dragging = False

# 在全局变量区域添加动态透明度相关变量
normal_alpha = 0.9  # 正常透明度
idle_alpha = 0.4    # 待机透明度
last_click_time = 0  # 最后一次点击时间
transparency_timer = None  # 透明度定时器

# 在全局变量区域添加双击检测相关变量
last_button_click_time = 0  # 按钮最后一次点击时间
double_click_threshold = 500  # 双击检测阈值（毫秒）

# 在全局变量区域添加随机种子重新设置相关变量
# 种子重新设置间隔配置（单位：分钟，默认5分钟）
SEED_REFRESH_MINUTES = 5  # 默认值
seed_refresh_interval = SEED_REFRESH_MINUTES * 60000  # 转换为毫秒
seed_refresh_timer = None  # 种子重新设置定时器

# 在全局变量区域添加抽取模式相关变量
personal_mode = "rotation"  # 个人抽取模式：rotation 或 weighted
group_mode = "rotation"     # 小组抽取模式：rotation 或 weighted
personal_weights = {}       # 个人权重字典
group_weights = {}          # 小组权重字典
last_personal_selected = None  # 上一个抽到的个人
last_group_selected = None    # 上一个抽到的小组

# 设置Windows任务栏属性
if platform.system() == 'Windows':
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RandomPicker.1.0")
    root.attributes('-toolwindow', False)  # 主窗口保持正常状态

def show_error_popup(message, close_window=True, auto_close=False):
    """
    显示错误信息弹窗的函数
    :param message: 错误信息内容
    :param close_window: 是否在报错后退出程序
    :param auto_close: 是否自动关闭弹窗
    """
    error_window = Toplevel(root)
    error_window.iconbitmap(resource_path('favicon.ico'))
    error_window.title("随机抽签器-提示")
    error_window.transient(root)
    error_window.attributes('-toolwindow', True)
    error_window.minsize(300, 150)  # 设置最小尺寸

    # 添加Windows系统特定的窗口样式设置
    if platform.system() == 'Windows':
        import ctypes
        # 获取窗口句柄
        hwnd = ctypes.windll.user32.GetParent(error_window.winfo_id())
        # 设置扩展样式为工具窗口
        GWL_EXSTYLE = -20
        WS_EX_TOOLWINDOW = 0x00000080
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)

    # 主框架
    main_frame = Frame(error_window, padx=15, pady=10)
    main_frame.pack(fill=BOTH, expand=True)

    # 错误信息标签，支持自动换行
    error_label = Label(main_frame,
                       text=message,
                       justify=LEFT,
                       wraplength=350,  # 设置最大宽度，超过会自动换行
                       anchor=W)
    error_label.pack(fill=X, pady=(0, 10))

    # 按钮框架
    button_frame = Frame(main_frame)
    button_frame.pack(side=BOTTOM, pady=(5, 0))

    # 添加自动关闭功能
    if auto_close:
        def auto_close_window():
            if error_window.winfo_exists():
                error_window.destroy()
        error_window.after(4000, auto_close_window)

    ok_button = Button(button_frame, text="确定",
                      command=lambda: close(error_window, close_window=close_window),
                      width=8)
    ok_button.pack()

    # 居中显示窗口
    error_window.update_idletasks()
    width = error_window.winfo_width()
    height = error_window.winfo_height()
    x = (error_window.winfo_screenwidth() // 2) - (width // 2)
    y = (error_window.winfo_screenheight() // 2) - (height // 2)
    error_window.geometry(f'+{x}+{y}')

    # 不调用mainloop()，让Tkinter主事件循环处理
    # error_window.mainloop()  # 移除这行以避免阻塞
    
def read(name, voice):
    """
    纯本地语音合成方案
    """
    global voice_enabled, error_shown, first_read_successful, read_queue, is_reading
    
    # 彩蛋音频播放不受voice_enabled影响
    if voice and os.path.exists(voice):
        try:
            mixer.music.load(voice)
            mixer.music.play()
            while mixer.music.get_busy():
                pygame.time.Clock().tick(10)
        except:
            pass

    if not voice_enabled:
        return

    # 将朗读请求添加到队列
    read_queue.append(name)
    
    # 如果当前没有正在进行的朗读，则开始处理队列
    if not is_reading:
        process_read_queue()

def process_read_queue():
    """处理朗读队列中的请求"""
    global read_queue, is_reading, first_read_successful, voice_enabled, error_shown
    
    # 如果队列为空或朗读已禁用，直接返回
    if not read_queue or not voice_enabled:
        is_reading = False
        return
    
    # 标记为正在朗读状态
    is_reading = True
    
    # 获取队列中的第一个朗读请求
    name = read_queue.pop(0)
    
    try:
        engine = pyttsx4.init()
        
        # 平台特定配置
        if platform.system() == 'Windows':
            voices = engine.getProperty('voices')
            chinese_voices = [v for v in voices if 'Chinese' in v.name]
            if chinese_voices:
                engine.setProperty('voice', chinese_voices[0].id)
            else:
                raise Exception("未找到中文语音包")
                
        elif platform.system() == 'Darwin':
            engine.setProperty('voice', 'com.apple.speech.synthesis.voice.ting-ting.premium')
            
        else:  # Linux
            engine.setProperty('voice', 'chinese')
        
        engine.setProperty('rate', 150)
        engine.setProperty('volume', 0.9)
        
        if platform.system() == 'Windows':
            CoInitialize()
        
        # 移除回调函数设置，改用直接方式
        engine.say(name)
        engine.runAndWait()
        
        # 如果成功执行，标记为第一次朗读成功
        if not first_read_successful:
            first_read_successful = True
        
        # 朗读完成后，继续处理队列中的下一个请求
        root.after(10, process_read_queue)  # 添加小延迟，避免过快占用资源

    except Exception as e:
        # 在朗读失败后，继续处理队列中的下一个请求
        root.after(10, process_read_queue)
        
        # 只在第一次朗读失败时禁用朗读功能并显示错误
        if not first_read_successful:
            voice_enabled = False
            if not error_shown:
                error_shown = True
                show_error_popup(
                    f"""由于您的系统不支持，姓名朗读功能已禁用。报错：
{str(e)}
后续抽取中将禁用姓名朗读功能""", 
                    close_window=False, auto_close=True
                )
        # 如果不是首次朗读（已有成功记录），则不作任何处理，保持朗读功能启用


def auto_close_windows():
    """
    自动关闭展示窗口的函数
    """
    global have_w, window, window_image, have_img, auto_close_timer

    # 取消定时器
    if auto_close_timer is not None:
        root.after_cancel(auto_close_timer)
        auto_close_timer = None

    # 关闭图片窗口（如果有）
    if have_img and window_image is not None and window_image.winfo_exists():
        window_image.destroy()
        have_img = False

    # 关闭名字窗口（如果有）
    if have_w and window is not None and window.winfo_exists():
        window.destroy()
        have_w = False

    print("自动关闭展示窗口")


def set_window_transparency(alpha):
    """
    设置窗口透明度的函数
    :param alpha: 透明度值 (0.0-1.0)
    """
    try:
        root.attributes('-alpha', alpha)
        print(f"[DEBUG] 窗口透明度已设置为: {alpha}")
    except Exception as e:
        print(f"[DEBUG] 设置透明度失败: {e}")


def switch_to_idle_transparency():
    """
    切换到待机透明度的函数
    """
    global idle_alpha
    set_window_transparency(idle_alpha)
    print(f"切换到待机透明度: {idle_alpha}")


def switch_to_normal_transparency():
    """
    切换到正常透明度的函数
    """
    global normal_alpha
    set_window_transparency(normal_alpha)
    print(f"切换到正常透明度: {normal_alpha}")


def check_transparency_timeout():
    """
    检查是否需要切换到待机透明度的定时函数
    """
    global last_click_time, transparency_timer

    current_time = time.time()
    time_diff = current_time - last_click_time
    print(f"[DEBUG] 检查透明度超时 - 当前时间差: {time_diff:.1f}秒")

    if time_diff >= 10:  # 10秒超时
        print("[DEBUG] 达到10秒超时，切换到待机透明度")
        switch_to_idle_transparency()
    else:
        # 还没到10秒，继续检查
        print(f"[DEBUG] 未达到超时，继续等待 - 剩余{10 - time_diff:.1f}秒")
        transparency_timer = root.after(1000, check_transparency_timeout)


def update_last_click_time():
    """
    更新最后点击时间的函数
    """
    global last_click_time, transparency_timer

    last_click_time = time.time()
    print(f"[DEBUG] 更新最后点击时间: {time.ctime(last_click_time)}")
    print(f"[DEBUG] 当前时间戳: {last_click_time}")

    # 切换到正常透明度
    switch_to_normal_transparency()

    # 取消之前的定时器
    if transparency_timer is not None:
        print("[DEBUG] 取消之前的透明度定时器")
        root.after_cancel(transparency_timer)

    # 重新启动定时器
    print("[DEBUG] 重新启动透明度检查定时器")
    transparency_timer = root.after(1000, check_transparency_timeout)


def reseed_random():
    """
    重新设置随机种子的函数
    """
    global seed_refresh_timer, seed_refresh_interval

    # 使用当前时间作为新种子
    current_time = int(time.time() * 1000000)  # 微秒级时间戳
    random.seed(current_time)

    # 生成几个随机数来验证新种子
    test_numbers = [random.randint(1, 100) for _ in range(3)]

    print("=== 随机种子重新设置 ===")
    print(f"[DEBUG] 时间戳: {current_time}")
    print(f"[DEBUG] 测试随机数: {test_numbers}")
    print(f"[DEBUG] 下次重新设置将在 {SEED_REFRESH_MINUTES} 分钟后")
    print("=" * 30)

    # 重新启动定时器
    seed_refresh_timer = root.after(seed_refresh_interval, reseed_random)


def initialize_weights():
    """
    初始化权重字典的函数
    """
    global personal_weights, group_weights, names, groups

    # 初始化个人权重
    personal_weights = {}
    for name in names:
        personal_weights[name] = 1.0  # 初始权重都为1.0

    # 初始化小组权重
    group_weights = {}
    for group in groups:
        group_weights[group] = 1.0  # 初始权重都为1.0

    print(f"[DEBUG] 权重已初始化 - 个人: {len(personal_weights)}个, 小组: {len(group_weights)}个")


def weighted_choice(items, weights, exclude_last=None):
    """
    基于权重的随机选择函数
    :param items: 项目列表
    :param weights: 权重字典
    :param exclude_last: 要排除的上一个抽到的人
    :return: 选中的项目
    """
    import random

    # 获取有效项目（排除请假人员）
    valid_items = [item for item in items if item not in leave_list]

    if not valid_items:
        # 如果没有有效项目，返回None
        return None

    # 如果只有一个有效项目，直接返回（避免无限循环）
    if len(valid_items) == 1:
        return valid_items[0]

    # 排除上一个抽到的人
    if exclude_last and exclude_last in valid_items:
        valid_items = [item for item in valid_items if item != exclude_last]

        # 如果排除后没有有效项目了，重新选择（不排除）
        if not valid_items:
            valid_items = [item for item in items if item not in leave_list]

    # 获取对应权重
    valid_weights = [weights.get(item, 1.0) for item in valid_items]

    # 计算权重总和
    total_weight = sum(valid_weights)

    # 生成随机数
    rand = random.random() * total_weight

    # 根据权重选择项目
    cumulative_weight = 0.0
    for item, weight in zip(valid_items, valid_weights):
        cumulative_weight += weight
        if rand <= cumulative_weight:
            return item

    # 理论上不会到达这里，但为了安全起见
    return valid_items[-1]


def update_weight(weights, selected_item):
    """
    更新权重（抽到后权重减半）
    :param weights: 权重字典
    :param selected_item: 被抽中的项目
    """
    if selected_item in weights:
        old_weight = weights[selected_item]
        new_weight = old_weight * 0.5  # 减半
        weights[selected_item] = new_weight
        print(f"[DEBUG] 权重更新: {selected_item} ({old_weight:.2f} -> {new_weight:.4f})")


def handle_button_click(event, action_func):
    """
    处理按钮点击事件，检测双击并执行相应操作
    :param event: 事件对象
    :param action_func: 要执行的操作函数
    """
    global last_button_click_time, double_click_threshold

    current_time = time.time() * 1000  # 转换为毫秒
    time_diff = current_time - last_button_click_time

    print(f"[DEBUG] 按钮点击 - 时间差: {time_diff:.0f}ms")

    if time_diff < double_click_threshold:
        # 双击检测 - 视为单次操作
        print("[DEBUG] 检测到双击，执行单次操作")
        last_button_click_time = 0  # 重置时间戳，防止连续双击
    else:
        # 普通点击
        print("[DEBUG] 普通点击，执行操作")
        last_button_click_time = current_time
        action_func()  # 执行操作


def show_window(name, image_name, color, voice, s_read, s_read_str, parent_window=None):
    """
    显示包含名字等信息的窗口函数
    :param name: 要显示的名字
    :param image_name: 要显示的图片名称（路径）
    :param color: 名字的颜色
    :param voice: 语音文件路径
    :param s_read: 是否特殊读取
    :param s_read_str: 特殊读取时的名字字符串
    """
    global have_w, window, window_image, have_img, auto_close_enabled, auto_close_timer

    print(name)

    # 修改这里：直接使用主窗口获取屏幕尺寸
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 修改图片窗口创建方式
    if image_name != '':
        window_image = Toplevel()
        window_image.withdraw()  # 先隐藏窗口
        window_image.iconbitmap(resource_path('favicon.ico'))
        window_image.title("随机抽签器")
        window_image.overrideredirect(1)

        img = Image.open(image_name)
        img_w, img_h = img.size
        
        max_img_width = int(screen_width * 0.5)
        if img_w > max_img_width:
            ratio = max_img_width / img_w
            img_w = max_img_width
            img_h = int(img_h * ratio)
            img = img.resize((img_w, img_h), Image.LANCZOS)

        x = (screen_width - img_w) // 2
        imgfile = ImageTk.PhotoImage(img)
        label_img = Label(window_image, image=imgfile)
        label_img.image = imgfile
        window_image.geometry(f"{img_w}x{img_h}+{x}+0")
        label_img.pack()
        have_img = True

        window_image.deiconify()  # 完成所有设置后显示窗口

    # 修改名字窗口处理
    if parent_window:
        window = parent_window
        window.withdraw()  # 先隐藏窗口
        window.deiconify()  # 统一使用deiconify显示
    else:
        window = Toplevel()
        window.withdraw()  # 先隐藏窗口
        window.iconbitmap(resource_path('favicon.ico'))
        window.title("随机抽签器")
        window.attributes('-toolwindow', True)
        # Windows系统设置工具窗口样式
        if platform.system() == 'Windows':
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
    
    # 动态计算字体大小（基于屏幕宽度）
    base_font_size = 87  # 基准字体大小（1920x1080）
    font_size = int(base_font_size * (screen_width / 1920))
    font_size = max(40, min(font_size, 150))  # 设置合理范围
    
    # 替换测试标签的创建方式
    # 创建字体对象
    custom_font = font.Font(family='华文仿宋', size=font_size)
    # 计算文本尺寸
    text_width = custom_font.measure(name)
    text_height = custom_font.metrics("linespace")
    
    # 优化窗口尺寸计算（减少边距）
    window_width = int(text_width * 1.1 + 20)  # 10%边距 + 固定20px
    window_height = int(text_height * 1.2 + 20)  # 20%边距 + 固定20px
    
    # 确保窗口最小尺寸（根据字体大小动态调整）
    min_width = max(150, font_size * 2)  # 至少2个字的宽度
    min_height = max(80, font_size + 20)  # 字体高度+20px
    window_width = max(window_width, min_width)
    window_height = max(window_height, min_height)
    
    # 计算居中位置
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    window.overrideredirect(1)
    window.attributes('-topmost', True)
    window.attributes('-alpha', 1)

    # 实际显示标签（优化换行逻辑）
    l = Label(window, 
             text=name, 
             font=('华文仿宋', font_size), 
             fg=color,
             wraplength=int(text_width * 1.05),  # 仅比实际宽度多5%
             justify=CENTER)
    l.place(relx=0.5, rely=0.5, anchor=CENTER)

    if s_read:
        name = s_read_str
    thr_read = threading.Thread(target=read, args=(name, voice,))
    thr_read.start()
    have_w = True
    window.deiconify()  # 完成所有设置后显示窗口

    # 设置自动关闭定时器（如果功能开启且不是测试模式）
    if auto_close_enabled and parent_window is None:
        # 先取消之前的定时器（如果有）
        if auto_close_timer is not None:
            root.after_cancel(auto_close_timer)
        # 设置10秒后自动关闭
        auto_close_timer = root.after(10000, auto_close_windows)

def set_leave_list():
    """
    处理"请假名单"选项的函数，弹出可编辑的文本框窗口让用户输入请假者名单，并更新leave_list变量
    """
    global leave_list
    leave_window = Toplevel(root)
    leave_window.iconbitmap(resource_path('favicon.ico'))
    leave_window.title("随机抽签器 - 请假管理")
    leave_window.transient(root)
    leave_window.attributes('-toolwindow', True)
    leave_window.minsize(400, 350)  # 设置最小尺寸

    # 添加Windows系统特定的窗口样式设置
    if platform.system() == 'Windows':
        import ctypes
        hwnd = ctypes.windll.user32.GetParent(leave_window.winfo_id())
        GWL_EXSTYLE = -20
        WS_EX_TOOLWINDOW = 0x00000080
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)

    # 主容器
    main_frame = Frame(leave_window, padx=10, pady=10)
    main_frame.pack(expand=True, fill=BOTH)

    # 头部框架（包含说明和保存按钮）
    header_frame = Frame(main_frame)
    header_frame.pack(fill=X, pady=5)

    # 说明文字
    instructions = Label(header_frame,
                        text="请假名单设置说明：\n"
                             "1. 每行输入一个请假人员姓名，Ta将不会被抽取\n"
                             "2. 姓名必须与配置文件中的完全一致",
                        justify=LEFT,
                        foreground="blue",
                        anchor=W)
    instructions.pack(side=LEFT, expand=True, fill=X)
    # 保存按钮（移动到说明文字右侧）
    save_btn = Button(header_frame, 
                    text="保存", 
                    command=lambda: save_and_close(),
                    width=6,
                    bg='#4CAF50',
                    fg='white')
    save_btn.pack(side=RIGHT, padx=5)

    # 编辑区域框架
    edit_frame = Frame(main_frame)
    edit_frame.pack(expand=True, fill=BOTH)

    # 文本框
    text_widget = Text(edit_frame, 
                     width=30, 
                     height=10,
                     wrap=WORD,
                     font=('微软雅黑', 10))
    text_widget.pack(expand=True, fill=BOTH, pady=5)
    
    # 初始化内容
    for name in leave_list:
        text_widget.insert(END, name + "\n")

    # 底部按钮框架
    bottom_frame = Frame(main_frame)
    bottom_frame.pack(pady=5)

    def save_and_close():
        global leave_list
        new_list = text_widget.get("1.0", END).splitlines()
        # 过滤空行和纯空格行
        leave_list = [line.strip() for line in new_list if line.strip()]
        leave_window.destroy()
        show_error_popup("请假名单已更新！", close_window=False, auto_close=True)

    def check_unsaved_changes():
        current_content = text_widget.get("1.0", END).strip()
        original_content = '\n'.join(leave_list).strip()
        if current_content != original_content:
            return messagebox.askyesno("未保存的更改", "是否放弃未保存的修改？")
        return True

    # 取消按钮（保持在底部）
    cancel_btn = Button(bottom_frame,
                      text="取消",
                      command=lambda: check_unsaved_changes() and leave_window.destroy(),
                      width=10,
                      bg='#f44336',
                      fg='white')
    cancel_btn.pack(side=RIGHT, padx=5)

    # 绑定窗口关闭事件
    leave_window.protocol("WM_DELETE_WINDOW", 
                        lambda: check_unsaved_changes() and leave_window.destroy())

def show_about():
    """
    处理"关于"选项的函数，弹出一个窗口展示关于本软件的信息
    """
    about_window = Toplevel(root)
    about_window.iconbitmap(resource_path('favicon.ico'))
    about_window.title("随机抽签器 - 关于")
    about_window.transient(root)
    about_window.attributes('-toolwindow', True)
    about_window.minsize(450, 400)  # 设置最小尺寸

    if platform.system() == 'Windows':
        import ctypes
        hwnd = ctypes.windll.user32.GetParent(about_window.winfo_id())
        GWL_EXSTYLE = -20
        WS_EX_TOOLWINDOW = 0x00000080
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)

    # 主框架容器
    main_frame = Frame(about_window)
    main_frame.pack(padx=20, pady=10, fill=BOTH, expand=True)

    # 软件信息部分
    about_text = Label(main_frame, 
        text="随机抽签器 v4.0.0\n"
             "©2019-2024 GZYzhy\n"
             "遵循 Apache 2.0 许可协议发布",
        justify=LEFT)
    about_text.grid(row=0, column=0, sticky=W)

    # 在软件信息部分后添加可点击的超链接
    from tkinter import ttk
    def open_github():
        import webbrowser
        webbrowser.open("https://github.com/gzyzhy/Name-Random-Picker")
    
    link = ttk.Label(main_frame, 
                    text="GitHub仓库",
                    foreground="blue",
                    cursor="hand2")
    link.grid(row=1, column=0, sticky=W, pady=5)
    link.bind("<Button-1>", lambda e: open_github())
    
    # 使用帮助部分
    help_text = """【主要功能】
• 随机抽取姓名/分组
• 支持彩蛋特效（图片/语音/颜色）
• 请假名单管理
• 窗口置顶和移动控制

【操作方法】
1. 主界面操作：
   - 左键点击按钮进行抽取
   - 右键点击按钮打开设置菜单
   - 拖动窗口时需先通过右键菜单解锁

2. 特色功能：
   - 彩蛋配置：在config.json中添加特殊效果
   - 测试模式：预览特定姓名的显示效果
   - 语音控制：可禁用朗读功能

3. 配置文件：
   • 支持自定义姓名列表、分组配置
   • 可设置不同分辨率下的自适应显示
   • 建议使用UTF-8编码格式"""

    help_label = Label(main_frame, 
                     text=help_text,
                     justify=LEFT,
                     anchor=W)
    help_label.grid(row=2, column=0, pady=10, sticky=W)

    # 注意事项
    notice_text = Label(main_frame,
                      text="※ 注意事项 ※\n"
                           "图片建议使用PNG格式\n"
                           "语音文件支持MP3/WAV格式",
                      foreground="red",
                      justify=LEFT)
    notice_text.grid(row=3, column=0, sticky=W)

    # 按钮区域
    button_frame = Frame(main_frame)
    button_frame.grid(row=4, column=0, pady=10)
    
    gen_button = Button(button_frame, 
                      text="生成示例配置", 
                      command=lambda: create_sample_config(about_window, exit_after=False))
    gen_button.pack(side=LEFT, padx=5)

    gen_button = Button(button_frame, 
                      text="编辑配置", 
                      command=edit_config)
    gen_button.pack(side=LEFT, padx=5)
    
    ok_button = Button(button_frame, 
                     text="确定", 
                     command=about_window.destroy)
    ok_button.pack(side=LEFT, padx=5)



def egg_show(name, mode="name", _test_window=None):
    """
    根据彩蛋配置展示特殊效果的函数
    :param name: 要展示的名字
    :param mode: 要使用的读取模式
    """
    img = ''
    color = 'black'
    voice = ''
    special_read = False
    s_read_str = ''

    if egg:
        if mode=="name":
            for special_case in config['egg_cases']:
                if name == special_case['name']:
                    if ('new_name' in special_case) and (special_case['new_name'] != ""):
                        name = special_case['new_name']
                    if ('color' in special_case) and (special_case['color'] != ""):
                        # 检查颜色设置是否合规
                        if (special_case['color'] not in ['black', 'white', 'red', 'green', 'blue', 'yellow', 'purple']):
                            show_error_popup(f"彩蛋设置中的颜色 {special_case['color']} 不合规,请使用'black','white','red','green','blue','yellow','purple'中的一种")
                            return
                        else:
                            color = special_case['color']
                    if ('image' in special_case) and (special_case['image'] != ""):
                        # 检查图片文件是否存在
                        if (not os.path.exists(special_case['image'])):
                            show_error_popup(f"找不到彩蛋设置中的图片文件: {special_case['image']}")
                            return
                        else:
                            img = special_case['image']
                    if ('voice' in special_case) and (special_case['voice'] != ""):
                        # 检查语音文件是否存在
                        if (not os.path.exists(special_case['voice'])):
                            show_error_popup(f"找不到彩蛋设置中的语音文件: {special_case['voice']}")
                            return
                        else:
                            voice += special_case['voice']
                    if ('s_read_str' in special_case) and (special_case['s_read_str'] != ""):
                        s_read_str = special_case['s_read_str']
                    break

            if s_read_str != '':
                special_read = True
            show_window(name, img, color, voice, special_read, s_read_str, _test_window)

        if mode=="group":
            for special_case in config['egg_cases_group']:
                if name == special_case['name']:
                    if ('new_name' in special_case) and (special_case['new_name'] != ""):
                        name = special_case['new_name']
                    if ('color' in special_case) and (special_case['color'] != ""):
                        # 检查颜色设置是否合规
                        if (special_case['color'] not in ['black', 'white', 'red', 'green', 'blue', 'yellow', 'purple']):
                            show_error_popup(f"彩蛋设置中的颜色 {special_case['color']} 不合规,请使用'black','white','red','green','blue','yellow','purple'中的一种")
                            return
                        else:
                            color = special_case['color']
                    if ('image' in special_case) and (special_case['image'] != ""):
                        # 检查图片文件是否存在
                        if (not os.path.exists(special_case['image'])):
                            show_error_popup(f"找不到彩蛋设置中的图片文件: {special_case['image']}")
                            return
                        else:
                            img = special_case['image']
                    if ('voice' in special_case) and (special_case['voice'] != ""):
                        # 检查语音文件是否存在
                        if (not os.path.exists(special_case['voice'])):
                            show_error_popup(f"找不到彩蛋设置中的语音文件: {special_case['voice']}")
                            return
                        else:
                            voice += special_case['voice']
                    if ('s_read_str' in special_case) and (special_case['s_read_str'] != ""):
                        s_read_str = special_case['s_read_str']
                    break

            if s_read_str != '':
                special_read = True
            show_window(name, img, color, voice, special_read, s_read_str, _test_window)
    
    else:
        # 不播放任何音频
        show_window(name, img, color, voice, special_read, s_read_str, _test_window)


def openwindow():
    """
    打开抽取名字窗口的函数
    """
    global is_dragging, auto_close_timer, last_personal_selected

    # 更新点击时间和透明度
    update_last_click_time()

    if is_dragging:
        is_dragging = False
        return
    global have_w
    global name
    global names
    global names_use
    global egg
    global now_use
    global groups
    global groups_use
    global window_image
    global have_img

    if have_w:
        # 手动关闭时取消自动关闭定时器
        if auto_close_timer is not None:
            root.after_cancel(auto_close_timer)
            auto_close_timer = None
        if have_img and window_image is not None:
            window_image.destroy()
            have_img = False
        if window is not None:
            window.destroy()
        have_w = False
        return

    global personal_mode, personal_weights

    if personal_mode == "rotation":
        # 轮转模式（原有逻辑）
        if names_use == []:
            names_use = names[:]

        while True:
            name = random.choice(names_use)
            names_use.remove(name)
            if name not in leave_list:
                break
    elif personal_mode == "weighted":
        # 加权模式
        name = weighted_choice(names, personal_weights, last_personal_selected)
        if name is None:
            # 如果没有有效项目，提示错误
            show_error_popup("没有有效的抽取对象（可能所有人都请假了）", close_window=False)
            return
        # 更新权重
        update_weight(personal_weights, name)
        # 记录这次抽取的结果
        last_personal_selected = name
    else:
        # 默认使用轮转模式
        if names_use == []:
            names_use = names[:]

        while True:
            name = random.choice(names_use)
            names_use.remove(name)
            if name not in leave_list:
                break

    egg_show(name)


def openwindow_group():
    """
    打开抽取分组窗口的函数
    """
    global is_dragging, auto_close_timer, last_group_selected

    # 更新点击时间和透明度
    update_last_click_time()

    if is_dragging:
        is_dragging = False
        return
    global window, window_image
    global have_w, have_img
    global name
    global groups
    global groups_use

    if have_w:
        # 手动关闭时取消自动关闭定时器
        if auto_close_timer is not None:
            root.after_cancel(auto_close_timer)
            auto_close_timer = None
        # 关闭图片窗口（如果有）
        if have_img and window_image is not None and window_image.winfo_exists():
            window_image.destroy()
            have_img = False
        # 关闭名字窗口
        if window is not None and window.winfo_exists():
            window.destroy()
        have_w = False
        return

    global group_mode, group_weights

    if group_mode == "rotation":
        # 轮转模式（原有逻辑）
        if groups_use == []:
            groups_use = groups[:]

        name = random.choice(groups_use)
        groups_use.remove(name)
    elif group_mode == "weighted":
        # 加权模式
        name = weighted_choice(groups, group_weights, last_group_selected)
        if name is None:
            # 如果没有有效项目，提示错误
            show_error_popup("没有有效的小组抽取对象", close_window=False)
            return
        # 更新权重
        update_weight(group_weights, name)
        # 记录这次抽取的结果
        last_group_selected = name
    else:
        # 默认使用轮转模式
        if groups_use == []:
            groups_use = groups[:]

        name = random.choice(groups_use)
        groups_use.remove(name)

    egg_show(name,"group")


def reset():
    """
    重置名字列表的函数
    """
    print('已重置个人抽取记忆')
    global names, names_use, personal_weights
    names_use = names[:]

    # 如果是加权模式，重置权重
    if personal_mode == "weighted":
        initialize_weights()
        print('已重置个人权重')
        global last_personal_selected
        last_personal_selected = None  # 重置上一个抽取记录


def reset_group():
    """
    重置分组列表的函数
    """
    print('已重置小组抽取记忆')
    global groups, groups_use, group_weights
    groups_use = groups[:]

    # 如果是加权模式，重置权重
    if group_mode == "weighted":
        initialize_weights()
        print('已重置小组权重')
        global last_group_selected
        last_group_selected = None  # 重置上一个抽取记录


def egg_set():
    """
    设置彩蛋开关的函数
    """
    global egg
    global menu

    if egg:
        menu.entryconfig('关闭彩蛋', label='开启彩蛋')
    else:
        menu.entryconfig('开启彩蛋', label='关闭彩蛋')

    egg = not egg
    print(egg)


def auto_close_set():
    """
    设置自动关闭开关的函数
    """
    global auto_close_enabled
    global menu

    if auto_close_enabled:
        menu.entryconfig('关闭自动关闭', label='开启自动关闭')
    else:
        menu.entryconfig('开启自动关闭', label='关闭自动关闭')

    auto_close_enabled = not auto_close_enabled
    print(f"自动关闭功能: {'开启' if auto_close_enabled else '关闭'}")


def close(window,close_window=True):
    """
    关闭主窗口并退出程序的函数
    """
    
    window.destroy()
    if close_window:
        os._exit(0)


def test():
    """
    测试函数，用于根据输入的名字展示效果
    """
    entry_str = simpledialog.askstring(title='随机抽签器 - 测试', prompt='输入要测试效果的姓名/组')
    if (entry_str in names) or (entry_str in groups):
        test_window = Toplevel()
        test_window.withdraw()
        test_window.iconbitmap(resource_path('favicon.ico'))
        test_window.title("随机抽签器 - 测试模式")
        test_window.attributes('-toolwindow', True)
        if platform.system() == 'Windows':
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(test_window.winfo_id())
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW)
        egg_show(entry_str, _test_window=test_window)
    else:
        show_error_popup("所输入的姓名/组号不在配置文件给定列表之内", close_window=False)

def start_drag(event):
    global drag_start_x, drag_start_y, is_dragging
    drag_start_x = event.x
    drag_start_y = event.y
    is_dragging = False  # 初始化拖动状态

def do_drag(event):
    global is_dragging
    # 计算移动距离
    delta_x = abs(event.x - drag_start_x)
    delta_y = abs(event.y - drag_start_y)
    
    # 超过阈值视为拖动
    if delta_x > drag_threshold or delta_y > drag_threshold:
        is_dragging = True
        x = root.winfo_x() - drag_start_x + event.x
        y = root.winfo_y() - drag_start_y + event.y
        root.geometry(f"+{x}+{y}")

def move():
    global now_move, menu
    now_move = not now_move
    
    if now_move:
        menu.entryconfig('移动窗口', label='锁定窗口')
        root.config(highlightthickness=2, highlightbackground="red")
        root.bind("<Button-1>", start_drag)
        root.bind("<B1-Motion>", do_drag)
        root.bind("<ButtonRelease-1>", lambda e: is_dragging)  # 清除拖动状态
    else:
        menu.entryconfig('锁定窗口', label='移动窗口')
        root.config(highlightthickness=0)
        root.unbind("<Button-1>")
        root.unbind("<B1-Motion>")
        root.unbind("<ButtonRelease-1>")

# 修改按钮背景色和样式
button_name = Button(root,
                    width=50,
                    height=3,
                    bg='lightblue',  # 改为浅蓝色使按钮可见
                    activebackground='skyblue',
                    relief='raised')  # 添加凸起效果
button_name.pack()

button_group = Button(root,
                     width=50,
                     height=2,
                     bg='lightcoral',  # 改为浅珊瑚色
                     activebackground='coral',
                     relief='raised')  # 添加凸起效果
button_group.pack()

# 为按钮添加点击事件绑定（支持双击检测）
button_name.bind('<Button-1>', lambda event: handle_button_click(event, openwindow))
button_group.bind('<Button-1>', lambda event: handle_button_click(event, openwindow_group))

# 添加任务栏图标支持
import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("RandomPicker.1.0")

def showPopoutMenu(w, menu):
    """
    展示右键弹出菜单的函数
    :param w: 绑定菜单的控件3
    :param menu: 要展示的菜单
    """
    def popout(event):
        # 更新点击时间和透明度（右键菜单也是用户交互）
        update_last_click_time()
        menu.post(event.x + w.winfo_rootx(), event.y + w.winfo_rooty())
        w.update()

    w.bind('<Button-3>', popout)

def validate_config(config):
    """配置文件自检函数"""
    errors = []
    
    # 检查必要字段
    required_fields = ['names', 'groups', 'egg_cases', 'egg_cases_group']
    for field in required_fields:
        if field not in config:
            errors.append(f"缺少必要字段: {field}")
    
    # 检查姓名列表
    if not isinstance(config.get('names', []), list) or len(config['names']) < 1:
        errors.append("姓名列表(names)必须包含至少一个姓名")
    
    # 检查分组列表
    if not isinstance(config.get('groups', []), list):
        errors.append("分组列表(groups)必须为数组格式")
    
    # 检查彩蛋配置
    egg_sections = [
        ('egg_cases', '个人彩蛋'),
        ('egg_cases_group', '分组彩蛋')
    ]
    for section, name in egg_sections:
        for case in config.get(section, []):
            # 检查必要字段
            if 'name' not in case:
                errors.append(f"{name}配置缺少'name'字段")
            
            # 检查颜色值
            if 'color' in case and case['color'] not in ['black', 'white', 'red', 'green', 'blue', 'yellow', 'purple']:
                errors.append(f"{name}配置中的颜色值不合法: {case['color']}")
            
            # 检查文件存在性
            for file_type in ['image', 'voice']:
                if file_type in case and case[file_type] != "":
                    if not os.path.exists(case[file_type]):
                        errors.append(f"{name}配置中的{file_type}文件不存在: {case[file_type]}")
    
    # 检查重复项
    if len(config['names']) != len(set(config['names'])):
        errors.append("姓名列表中存在重复项")
    
    if len(config['groups']) != len(set(config['groups'])):
        errors.append("分组列表中存在重复项")
    
    return errors

def read_config(path):
    """
    读取JSON配置文件的函数
    :param path: 配置文件路径
    """
    print(f"[DEBUG] 开始读取配置文件: {path}")
    global names, groups, config, leave_list, auto_close_enabled
    try:
        with open(path, 'rb') as f:
            rawdata = f.read()
            result = chardet.detect(rawdata)
            encoding = result['encoding']
            with open(path, 'r', encoding=encoding) as f:
                config = json.load(f)
                
                # 执行配置验证
                validation_errors = validate_config(config)
                if validation_errors:
                    error_msg = "配置文件校验失败：\n" + "\n".join(f"• {err}" for err in validation_errors)
                    show_error_popup(error_msg)
                    os._exit(1)
                
                names = config['names']
                names_use = names[:]
                groups = config['groups']
                groups_use = groups[:]

                # 读取自动关闭设置（可选字段，默认值为True）
                if 'auto_close' in config:
                    if isinstance(config['auto_close'], bool):
                        auto_close_enabled = config['auto_close']
                    else:
                        print(f"警告：配置文件中的auto_close字段值无效({config['auto_close']})，使用默认值True")
                else:
                    auto_close_enabled = True  # 默认开启
                    print("配置文件中未找到auto_close字段，使用默认值True")

                # 读取种子重设间隔设置（可选字段，默认值为5分钟）
                global SEED_REFRESH_MINUTES, seed_refresh_interval
                if 'seed_refresh_minutes' in config:
                    if isinstance(config['seed_refresh_minutes'], int) and config['seed_refresh_minutes'] > 0:
                        SEED_REFRESH_MINUTES = config['seed_refresh_minutes']
                        seed_refresh_interval = SEED_REFRESH_MINUTES * 60000  # 转换为毫秒
                        print(f"配置文件中设置种子重设间隔为: {SEED_REFRESH_MINUTES}分钟")
                    else:
                        print(f"警告：配置文件中的seed_refresh_minutes字段值无效({config['seed_refresh_minutes']})，使用默认值5分钟")
                else:
                    print("配置文件中未找到seed_refresh_minutes字段，使用默认值5分钟")

                # 读取抽取模式设置（可选字段，默认值为rotation）
                global personal_mode, group_mode
                if 'personal_mode' in config:
                    if config['personal_mode'] in ['rotation', 'weighted']:
                        personal_mode = config['personal_mode']
                        print(f"配置文件中设置个人抽取模式为: {personal_mode}")
                    else:
                        print(f"警告：配置文件中的personal_mode字段值无效({config['personal_mode']})，使用默认值rotation")
                else:
                    print("配置文件中未找到personal_mode字段，使用默认值rotation")

                if 'group_mode' in config:
                    if config['group_mode'] in ['rotation', 'weighted']:
                        group_mode = config['group_mode']
                        print(f"配置文件中设置小组抽取模式为: {group_mode}")
                    else:
                        print(f"警告：配置文件中的group_mode字段值无效({config['group_mode']})，使用默认值rotation")
                else:
                    print("配置文件中未找到group_mode字段，使用默认值rotation")

                # 初始化权重（在所有配置读取完成后）
                initialize_weights()
                # 延迟显示启动提示，避免阻塞随机种子初始化
                root.after(100, lambda: show_error_popup(
                    f"程序已开始运行，请使用屏幕右下角的方块按钮来抽取！\n当前使用的配置文件：{os.path.abspath(path)}",
                    close_window=False,
                    auto_close=True  # 添加自动关闭参数
                ))

                # 初始化透明度系统
                global last_click_time
                last_click_time = time.time()  # 设置初始点击时间
                print(f"[DEBUG] 透明度系统初始化 - 当前时间: {time.ctime(last_click_time)}")
                print(f"[DEBUG] 正常透明度: {normal_alpha}, 待机透明度: {idle_alpha}")
                print(f"[DEBUG] 程序启动时使用待机透明度: {idle_alpha}")
                root.after(1000, check_transparency_timeout)  # 启动透明度检查定时器
                print("[DEBUG] 透明度检查定时器已启动")

                # 初始化随机种子系统
                print("[DEBUG] 初始化随机种子重新设置系统")

                # 立即执行一次重新做种，让用户能立即看到效果
                print("=== 初始随机种子设置 ===")
                initial_seed = int(time.time() * 1000000)
                random.seed(initial_seed)

                # 生成测试随机数
                initial_test_numbers = [random.randint(1, 100) for _ in range(3)]
                print(f"[DEBUG] 时间戳: {initial_seed}")
                print(f"[DEBUG] 测试随机数: {initial_test_numbers}")
                print("=" * 30)

                # 启动定时器
                root.after(seed_refresh_interval, reseed_random)  # 启动种子重新设置定时器
                print(f"[DEBUG] 种子重新设置定时器已启动 - 间隔: {SEED_REFRESH_MINUTES}分钟")
    except Exception as e:
        handle_config_error(e, path)

def handle_config_error(e, path):
    choice = messagebox.askyesno("随机抽签器 - 配置错误",
                                f"默认配置文件加载失败({e})\n是否要创建示例配置文件？\n(选择'否'将手动选择配置文件)",
                                icon='warning')
    if choice:
        create_sample_config(exit_after=True)
    else:
        select_config_file()

def create_sample_config(parent=None, exit_after=True):
    file_path = filedialog.asksaveasfilename(
        parent=parent,
        title="保存示例配置文件",
        defaultextension=".json",
        initialfile="config.json",
        filetypes=[("JSON文件", "*.json")]
    )
    if file_path:
        try:
            sample_config = {
                "names": ["示例姓名1", "示例姓名2", "示例姓名3"],
                "groups": ["示例分组1", "示例分组2"],
                "auto_close": True,  # 自动关闭功能开关，默认开启
                "egg_cases": [{
                    "name": "示例姓名1",
                    "new_name": "示例姓名1的展示名",
                    "color": "blue",
                    "image": "src/sample_image.png",
                    "voice": "src/sample_bgm.mp3",
                    "s_read_str": "示例姓名1的朗读名"
                }],
                "egg_cases_group": [{
                    "name": "示例分组1",
                    "new_name": "示例分组1的展示名",
                    "color": "blue",
                    "image": "src/sample_image.png",
                    "voice": "src/sample_bgm.mp3",
                    "s_read_str": "示例分组1的朗读名"
                }]
            }
            with open(file_path, 'w') as f:
                json.dump(sample_config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("生成成功", 
                f"示例配置文件已成功生成在：\n{file_path}\n请按需修改后使用", 
                parent=parent)
            if exit_after:
                os._exit(0)
        except Exception as e:
            messagebox.showerror("生成失败", 
                f"配置文件生成失败: {e}", 
                parent=parent)
            if exit_after:
                os._exit(1)
    else:
        messagebox.showinfo("已取消", "已取消配置文件生成", parent=parent)
        if exit_after:
            os._exit(0)

def select_config_file():
    file_path = filedialog.askopenfilename(
        title="选择配置文件",
        defaultextension=".json",
        initialfile="config.json",
        filetypes=[("JSON文件", "*.json")]
    )
    if file_path:
        try:
            read_config(file_path)
        except Exception as e:
            show_error_popup(f"配置文件加载失败: {e}")
            os._exit(1)
    else:
        show_error_popup("已取消配置文件选择")
        os._exit(0)
def create_tray_icon(root, config_path):
    """
    创建系统托盘图标
    :param root: Tkinter根窗口
    """
    # 加载图标
    icon_path = resource_path('favicon.ico')
    
    def exit_action(icon):
        close(root)
    
    def toggle_window(icon):
        if root.winfo_viewable():
            root.withdraw()
        else:
            root.deiconify()
            root.attributes('-topmost', True)
            root.update()
            root.attributes('-topmost', False)
    
    # 添加一个调度器函数，将操作转发到主线程
    def schedule_to_main_thread(func):
        def wrapper(icon):
            root.after(0, func)  # 使用 after 方法在主线程中执行函数
        return wrapper
    
    # 创建独立的系统托盘菜单项
    menu_items = [
        MenuItem('显示/隐藏', toggle_window),
        MenuItem('重置个人', schedule_to_main_thread(reset)),
        MenuItem('重置小组', schedule_to_main_thread(reset_group)),
        # 重读配置函数也需要包装
        MenuItem('重读配置', schedule_to_main_thread(lambda: read_config(config_path))),
        MenuItem('编辑配置', schedule_to_main_thread(edit_config)),
        MenuItem('关于', schedule_to_main_thread(show_about)),
        MenuItem('退出', exit_action)
    ]
    
    # 加载图标图像
    image = Image.open(icon_path)
    
    # 创建托盘图标
    tray_icon = Icon(
        "RandomPicker",
        image,
        "随机抽签器",
        PystrayMenu(*menu_items)
    )
    
    # 在新线程中运行托盘图标
    threading.Thread(target=tray_icon.run, daemon=True).start()
    
    return tray_icon

menu = Menu(root)
menu.add_cascade(label='测试模式', command=test)
menu.add_cascade(label='移动窗口', command=move)  # 初始标签为"移动窗口"
menu.add_cascade(label='重置个人', command=reset)
menu.add_cascade(label='重置小组', command=reset_group)
menu.add_cascade(label='关闭彩蛋', command=egg_set)
menu.add_cascade(label='关闭自动关闭', command=auto_close_set)
menu.add_cascade(label='请假名单', command=set_leave_list)
menu.add_cascade(label='重读配置', command=lambda:read_config(config_path))
menu.add_cascade(label='编辑配置', command=edit_config)
menu.add_cascade(label='关于', command=show_about)
menu.add_cascade(label='退出', command=lambda: close(root))
showPopoutMenu(button_name, menu)
tray_icon = create_tray_icon(root, config_path)
# 添加定期置顶维持（在mainloop之前）
def maintain_topmost():
    root.attributes('-topmost', True)
    root.after(1000, maintain_topmost)
if tray_icon:
        tray_icon.stop()
if __name__ == "__main__":
    # 首先检查是否已有实例在运行
    if not ensure_single_instance():
        print("检测到已有程序实例在运行，当前实例将退出。")
        sys.exit(0)

    print("=== 程序启动 - 准备初始化随机种子系统 ===")
    config_path = "config.json"
    
    # 修改后的主程序入口
    if not os.path.exists(config_path):
        handle_config_error(FileNotFoundError("配置文件不存在"), config_path)
    else:
        try:
            read_config(config_path)
        except Exception as e:
            handle_config_error(e, config_path)
    
    root.after(1000, maintain_topmost)  # 添加定期置顶检查
    root.mainloop()
