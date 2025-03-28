"""
coding: utf-8
©2025 GZYzhy Publish under Apache License 2.0
GitHub: https://github.com/gzyzhy/Name-Random-Picker

随机抽签器 - 配置文件编辑器
"""

import json
import os
import sys
import platform
import traceback
import chardet
from pathlib import Path
from tkinter import (
    Tk, Frame, Label, Button, Entry, Text, Menu, Toplevel, StringVar,
    messagebox, filedialog, ttk, colorchooser, BooleanVar
)
from tkinter.constants import *
import pandas as pd
from functools import partial
import shutil

# 全局变量
CONFIG_TEMPLATE = {
    "names": [],
    "groups": [],
    "egg_cases": [],
    "egg_cases_group": []
}

# 颜色映射表
COLOR_MAP = {
    "黑色": "black",
    "白色": "white",
    "红色": "red",
    "绿色": "green",
    "蓝色": "blue",
    "黄色": "yellow",
    "紫色": "purple"
}
REVERSE_COLOR_MAP = {v: k for k, v in COLOR_MAP.items()}

class ConfigEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("随机抽签器 - 配置文件编辑器")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
    
        # 默认配置文件路径
        self.config_path = os.path.join(os.getcwd(), "config.json")
        self.current_encoding = 'utf-8'  # 默认编码设置
        self.config_data = self._load_default_config()
        self.is_modified = False
        self.last_save_time = None
    
        # 创建主界面
        self._create_menu()
        self._create_main_frame()
    
        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
        # 设置图标
        self._set_app_icon()
    
    def clear_items(self, key):
        """清空指定列表"""
        if key == "names":
            tree = self.names_tree
            message = "确认要清空所有姓名吗？"
            egg_key = "egg_cases"
        else:
            tree = self.groups_tree
            message = "确认要清空所有分组吗？"
            egg_key = "egg_cases_group"
    
        # 确认对话框
        confirm = messagebox.askyesno("确认清空", message)
        if not confirm:
            return
    
        # 清空列表
        self.config_data[key] = []
    
        # 清空相关的彩蛋配置
        if egg_key in self.config_data:
            self.config_data[egg_key] = []
    
        # 刷新显示
        if key == "names":
            self.refresh_names_data()
        else:
            self.refresh_groups_data()
    
        # 刷新彩蛋配置显示
        self.refresh_egg_data(egg_key)
    
        self.is_modified = True
        self.update_status_bar()

    def _set_app_icon(self):
        """设置应用图标"""
        try:
            ico_path = self._resource_path("favicon.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
        except Exception:
            pass  # 如果图标设置失败，不影响程序运行
    
    def _resource_path(self, relative_path):
        """获取资源文件的绝对路径"""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    def _load_default_config(self):
        """加载默认配置文件，自动检测编码"""
        if os.path.exists(self.config_path):
            try:
                # 使用二进制模式读取文件
                with open(self.config_path, 'rb') as f:
                    content = f.read()
                
                    # 使用chardet检测文件编码
                    detect_result = chardet.detect(content)
                    encoding = detect_result['encoding']
                    confidence = detect_result['confidence']
                
                    # 记录当前文件的编码
                    self.current_encoding = encoding
                
                    # 如果置信度太低或者检测失败，尝试常见编码
                    if confidence < 0.7 or not encoding:
                        for enc in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                            try:
                                decoded_content = content.decode(enc)
                                self.current_encoding = enc
                                return json.loads(decoded_content)
                            except (UnicodeDecodeError, json.JSONDecodeError):
                                continue
                    
                        # 如果所有编码都失败
                        messagebox.showerror("错误", f"无法识别配置文件编码，请确保文件格式正确")
                        return dict(CONFIG_TEMPLATE)
                
                    # 使用检测到的编码解析JSON
                    decoded_content = content.decode(encoding)
                    return json.loads(decoded_content)
                
            except json.JSONDecodeError:
                messagebox.showerror("错误", f"配置文件 {self.config_path} 格式错误")
                return dict(CONFIG_TEMPLATE)
            except Exception as e:
                messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
                return dict(CONFIG_TEMPLATE)
        else:
            # 默认使用UTF-8编码
            self.current_encoding = 'utf-8'
            return dict(CONFIG_TEMPLATE)
    
    def _create_menu(self):
        """创建菜单栏"""
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建配置", command=self.new_config)
        file_menu.add_command(label="打开配置", command=self.open_config)
        file_menu.add_command(label="保存配置", command=self.save_config)
        file_menu.add_command(label="另存为...", command=self.save_config_as)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_closing)
        
        # 导入导出菜单
        import_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="导入/导出", menu=import_menu)
        import_menu.add_command(label="从Excel导入名单", command=self.import_from_excel)
        import_menu.add_command(label="导出名单到Excel", command=self.export_to_excel)
        import_menu.add_command(label="生成Excel模板", command=self.create_excel_template)
        
        # 帮助菜单
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
    
    def _create_main_frame(self):
        """创建主框架和标签页"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=5)  # 减少底部padding
    
        # 创建各个标签页
        self.names_frame = self._create_names_tab()
        self.groups_frame = self._create_groups_tab()
        self.egg_personal_frame = self._create_egg_tab("个人")
        self.egg_group_frame = self._create_egg_tab("分组")
    
        # 添加标签页到notebook
        self.notebook.add(self.names_frame, text="姓名管理")
        self.notebook.add(self.groups_frame, text="分组管理")
        self.notebook.add(self.egg_personal_frame, text="个人彩蛋配置")
        self.notebook.add(self.egg_group_frame, text="分组彩蛋配置")
    
        # 添加底部按钮栏
        button_frame = Frame(self.root)
        button_frame.pack(fill=X, padx=10, pady=5)
    
        save_btn = ttk.Button(button_frame, text="保存配置", command=self.save_config)
        save_btn.pack(side=RIGHT, padx=5)
    
        # 创建状态栏
        self.status_bar = Label(self.root, text="就绪", bd=1, relief=SUNKEN, anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X)
    
        # 加载初始数据
        self.refresh_all_data()
    
    def _create_names_tab(self):
        """创建姓名管理标签页"""
        frame = ttk.Frame(self.notebook)
        
        # 工具栏
        toolbar = Frame(frame)
        toolbar.pack(fill=X, padx=5, pady=5)
        
        add_btn = ttk.Button(toolbar, text="添加", command=lambda: self.add_item("names"))
        add_btn.pack(side=LEFT, padx=2)
        
        delete_btn = ttk.Button(toolbar, text="删除", command=lambda: self.delete_item("names"))
        delete_btn.pack(side=LEFT, padx=2)
        
        clear_btn = ttk.Button(toolbar, text="清空", command=lambda: self.clear_items("names"))
        clear_btn.pack(side=LEFT, padx=2)
        
        # 名单列表
        list_frame = ttk.LabelFrame(frame, text="姓名列表")
        list_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # 创建Treeview用于显示姓名列表
        self.names_tree = ttk.Treeview(
            list_frame, 
            columns=('序号', '姓名'),
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        self.names_tree.column('序号', width=50, anchor='center')
        self.names_tree.column('姓名', width=200, anchor='w')
        self.names_tree.heading('序号', text='序号')
        self.names_tree.heading('姓名', text='姓名')
        
        self.names_tree.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.names_tree.yview)
        
        # 绑定双击编辑事件
        self.names_tree.bind('<Double-1>', lambda event: self.edit_item("names"))
        
        return frame
    
    def _create_groups_tab(self):
        """创建分组管理标签页"""
        frame = ttk.Frame(self.notebook)
        
        # 工具栏
        toolbar = Frame(frame)
        toolbar.pack(fill=X, padx=5, pady=5)
        
        add_btn = ttk.Button(toolbar, text="添加", command=lambda: self.add_item("groups"))
        add_btn.pack(side=LEFT, padx=2)
        
        delete_btn = ttk.Button(toolbar, text="删除", command=lambda: self.delete_item("groups"))
        delete_btn.pack(side=LEFT, padx=2)
        
        clear_btn = ttk.Button(toolbar, text="清空", command=lambda: self.clear_items("groups"))
        clear_btn.pack(side=LEFT, padx=2)
        
        # 分组列表
        list_frame = ttk.LabelFrame(frame, text="分组列表")
        list_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # 创建Treeview用于显示分组列表
        self.groups_tree = ttk.Treeview(
            list_frame, 
            columns=('序号', '分组名称'),
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        self.groups_tree.column('序号', width=50, anchor='center')
        self.groups_tree.column('分组名称', width=200, anchor='w')
        self.groups_tree.heading('序号', text='序号')
        self.groups_tree.heading('分组名称', text='分组名称')
        
        self.groups_tree.pack(fill=BOTH, expand=True)
        scrollbar.config(command=self.groups_tree.yview)
        
        # 绑定双击编辑事件
        self.groups_tree.bind('<Double-1>', lambda event: self.edit_item("groups"))
        
        return frame
    
    def _create_egg_tab(self, mode):
        """创建彩蛋配置标签页"""
        frame = ttk.Frame(self.notebook)
        
        # 工具栏
        toolbar = Frame(frame)
        toolbar.pack(fill=X, padx=5, pady=5)
        
        if mode == "个人":
            key = "egg_cases"
            source_list = "names"
        else:
            key = "egg_cases_group"
            source_list = "groups"
            
        add_btn = ttk.Button(toolbar, text="添加", command=lambda: self.add_egg_config(key))
        add_btn.pack(side=LEFT, padx=2)
        
        edit_btn = ttk.Button(toolbar, text="编辑", command=lambda: self.edit_egg_config(key))
        edit_btn.pack(side=LEFT, padx=2)
        
        delete_btn = ttk.Button(toolbar, text="删除", command=lambda: self.delete_egg_config(key))
        delete_btn.pack(side=LEFT, padx=2)
        
        # 彩蛋列表
        list_frame = ttk.LabelFrame(frame, text=f"{mode}彩蛋配置")
        list_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        # 创建Treeview用于显示彩蛋配置
        if mode == "个人":
            self.personal_egg_tree = ttk.Treeview(
                list_frame, 
                columns=('姓名', '显示名', '颜色', '图片', '语音', '朗读'),
                show='headings',
                yscrollcommand=scrollbar.set
            )
            tree = self.personal_egg_tree
        else:
            self.group_egg_tree = ttk.Treeview(
                list_frame, 
                columns=('分组名', '显示名', '颜色', '图片', '语音', '朗读'),
                show='headings',
                yscrollcommand=scrollbar.set
            )
            tree = self.group_egg_tree
        
        # 设置列宽和标题
        name_title = "姓名" if mode == "个人" else "分组名"
        tree.column(name_title, width=100, anchor='w')
        tree.column('显示名', width=100, anchor='w')
        tree.column('颜色', width=60, anchor='center')
        tree.column('图片', width=150, anchor='w')
        tree.column('语音', width=150, anchor='w')
        tree.column('朗读', width=100, anchor='w')
        
        tree.heading(name_title, text=name_title)
        tree.heading('显示名', text='显示名')
        tree.heading('颜色', text='颜色')
        tree.heading('图片', text='图片路径')
        tree.heading('语音', text='语音路径')
        tree.heading('朗读', text='朗读文本')
        
        tree.pack(fill=BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        # 绑定双击编辑事件
        tree.bind('<Double-1>', lambda event: self.edit_egg_config(key))
        
        return frame

    def refresh_all_data(self):
        """刷新所有数据显示"""
        self.refresh_names_data()
        self.refresh_groups_data()
        self.refresh_egg_data("egg_cases")
        self.refresh_egg_data("egg_cases_group")
        self.update_status_bar()
    
    def refresh_names_data(self):
        """刷新姓名列表数据"""
        # 清空原有数据
        for item in self.names_tree.get_children():
            self.names_tree.delete(item)
        
        # 插入新数据
        for i, name in enumerate(self.config_data.get('names', []), start=1):
            self.names_tree.insert('', END, values=(i, name))
    
    def refresh_groups_data(self):
        """刷新分组列表数据"""
        # 清空原有数据
        for item in self.groups_tree.get_children():
            self.groups_tree.delete(item)
        
        # 插入新数据
        for i, group in enumerate(self.config_data.get('groups', []), start=1):
            self.groups_tree.insert('', END, values=(i, group))
    
    def refresh_egg_data(self, key):
        """刷新彩蛋配置数据"""
        tree = self.personal_egg_tree if key == "egg_cases" else self.group_egg_tree
        
        # 清空原有数据
        for item in tree.get_children():
            tree.delete(item)
        
        # 插入新数据
        for item in self.config_data.get(key, []):
            name = item.get('name', '')
            new_name = item.get('new_name', '')
            color = item.get('color', '')
            color_display = REVERSE_COLOR_MAP.get(color, color)
            image = item.get('image', '')
            voice = item.get('voice', '')
            s_read_str = item.get('s_read_str', '')
            
            tree.insert('', END, values=(name, new_name, color_display, image, voice, s_read_str))
    
    def update_status_bar(self):
        """更新状态栏信息"""
        names_count = len(self.config_data.get('names', []))
        groups_count = len(self.config_data.get('groups', []))
        personal_egg_count = len(self.config_data.get('egg_cases', []))
        group_egg_count = len(self.config_data.get('egg_cases_group', []))
    
        status_text = f"当前配置: {os.path.basename(self.config_path)} | 姓名: {names_count} | 分组: {groups_count} | 个人彩蛋: {personal_egg_count} | 分组彩蛋: {group_egg_count} | 编码: {self.current_encoding}"
    
        if self.is_modified:
            status_text += " | [已修改]"
        
        self.status_bar.config(text=status_text)
    
    def add_item(self, key):
        """添加项目到列表"""
        title = "添加姓名" if key == "names" else "添加分组"
        prompt = "请输入姓名:" if key == "names" else "请输入分组名称:"
        
        result = self.show_input_dialog(title, prompt)
        if result and result.strip():
            # 检查重复项
            if result in self.config_data.get(key, []):
                messagebox.showwarning("警告", f"列表中已存在 '{result}'")
                return
            
            # 添加到配置
            if key not in self.config_data:
                self.config_data[key] = []
            self.config_data[key].append(result)
            
            # 刷新显示
            if key == "names":
                self.refresh_names_data()
            else:
                self.refresh_groups_data()
                
            self.is_modified = True
            self.update_status_bar()
    
    def edit_item(self, key):
        """编辑列表中的项目"""
        if key == "names":
            tree = self.names_tree
            title = "编辑姓名"
            prompt = "请输入新的姓名:"
            idx_col = 0
        else:
            tree = self.groups_tree
            title = "编辑分组"
            prompt = "请输入新的分组名称:"
            idx_col = 0
            
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一项")
            return
            
        item = tree.item(selected[0])
        old_value = item['values'][1]  # 获取当前值
        index = int(item['values'][0]) - 1  # 转为0-based索引
        
        # 显示输入对话框
        result = self.show_input_dialog(title, prompt, default=old_value)
        if result is not None and result != old_value:
            # 检查重复
            if result in self.config_data.get(key, []):
                messagebox.showwarning("警告", f"列表中已存在 '{result}'")
                return
                
            # 更新配置
            self.config_data[key][index] = result
            
            # 刷新显示
            if key == "names":
                self.refresh_names_data()
            else:
                self.refresh_groups_data()
                
            # 检查是否需要更新彩蛋配置中的名称
            egg_key = "egg_cases" if key == "names" else "egg_cases_group"
            for egg in self.config_data.get(egg_key, []):
                if egg.get('name') == old_value:
                    egg['name'] = result
            
            # 刷新彩蛋配置显示
            self.refresh_egg_data(egg_key)
                
            self.is_modified = True
            self.update_status_bar()
    
    def delete_item(self, key):
        """删除列表中的项目"""
        if key == "names":
            tree = self.names_tree
            message = "确认要删除选中的姓名吗？"
            egg_key = "egg_cases"
        else:
            tree = self.groups_tree
            message = "确认要删除选中的分组吗？"
            egg_key = "egg_cases_group"
        
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一项")
            return
        
        confirm = messagebox.askyesno("确认删除", message)
        if not confirm:
            return
        
        item = tree.item(selected[0])
        value = item['values'][1]  # 获取当前值
        index = int(item['values'][0]) - 1  # 转为0-based索引
    
        # 从配置中删除
        self.config_data[key].pop(index)
    
        # 同时删除关联的彩蛋配置
        if egg_key in self.config_data:
            # 创建一个新列表，不包含要删除的项
            self.config_data[egg_key] = [
                egg for egg in self.config_data.get(egg_key, [])
                if egg.get('name') != value
            ]
    
        # 刷新显示
        if key == "names":
            self.refresh_names_data()
        else:
            self.refresh_groups_data()
        
        # 刷新彩蛋配置显示
        self.refresh_egg_data(egg_key)
    
        self.is_modified = True
        self.update_status_bar()
    
    def add_egg_config(self, key):
        """添加彩蛋配置"""
        if key == "egg_cases":
            source_key = "names"
            title = "添加个人彩蛋"
            if not self.config_data.get('names'):
                messagebox.showwarning("提示", "请先添加姓名后再配置彩蛋")
                return
        else:
            source_key = "groups"
            title = "添加分组彩蛋"
            if not self.config_data.get('groups'):
                messagebox.showwarning("提示", "请先添加分组后再配置彩蛋")
                return
        
        # 创建彩蛋配置对话框
        self.show_egg_config_dialog(key=key, title=title, source_key=source_key)
    
    def edit_egg_config(self, key):
        """编辑彩蛋配置"""
        tree = self.personal_egg_tree if key == "egg_cases" else self.group_egg_tree
        title = "编辑个人彩蛋" if key == "egg_cases" else "编辑分组彩蛋"
        source_key = "names" if key == "egg_cases" else "groups"
        
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一项")
            return
            
        item = tree.item(selected[0])
        values = item['values']
        name = values[0]
        
        # 查找对应的配置项
        for i, egg in enumerate(self.config_data.get(key, [])):
            if egg.get('name') == name:
                self.show_egg_config_dialog(key=key, title=title, source_key=source_key, egg_data=egg, edit_index=i)
                break
    
    def delete_egg_config(self, key):
        """删除彩蛋配置"""
        if key == "egg_cases":
            tree = self.personal_egg_tree
            message = "确认要删除选中的个人彩蛋配置吗？"
        else:
            tree = self.group_egg_tree
            message = "确认要删除选中的分组彩蛋配置吗？"
    
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择一项")
            return
        
        confirm = messagebox.askyesno("确认删除", message)
        if not confirm:
            return
        
        item = tree.item(selected[0])
        name = item['values'][0]  # 获取姓名或分组名
    
        # 从配置中删除
        self.config_data[key] = [
            egg for egg in self.config_data.get(key, [])
            if egg.get('name') != name
        ]
    
        # 刷新显示
        self.refresh_egg_data(key)
        self.is_modified = True
        self.update_status_bar()
    
    def show_egg_config_dialog(self, key, title, source_key, egg_data=None, edit_index=None):
        """显示彩蛋配置对话框"""
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # 字段列表
        name_label = ttk.Label(main_frame, text="选择姓名:" if key == "egg_cases" else "选择分组:")
        name_label.grid(row=0, column=0, sticky=W, padx=5, pady=5)
        
        # 下拉列表
        name_var = StringVar()
        name_options = self.config_data.get(source_key, [])
        if not name_options:
            name_options = [""]
            
        name_combo = ttk.Combobox(main_frame, textvariable=name_var, values=name_options, width=30)
        name_combo.grid(row=0, column=1, sticky=W, padx=5, pady=5)
        name_combo.state(['readonly'])
        
        # 显示名称
        display_name_label = ttk.Label(main_frame, text="显示名称:")
        display_name_label.grid(row=1, column=0, sticky=W, padx=5, pady=5)
        
        display_name_var = StringVar()
        display_name_entry = ttk.Entry(main_frame, textvariable=display_name_var, width=30)
        display_name_entry.grid(row=1, column=1, sticky=W, padx=5, pady=5)
        
        # 颜色选择
        color_label = ttk.Label(main_frame, text="文字颜色:")
        color_label.grid(row=2, column=0, sticky=W, padx=5, pady=5)
        
        color_var = StringVar()
        color_combo = ttk.Combobox(main_frame, textvariable=color_var, values=list(REVERSE_COLOR_MAP.values()), width=20)
        color_combo.grid(row=2, column=1, sticky=W, padx=5, pady=5)
        
        def choose_color():
            color = colorchooser.askcolor(title="选择颜色")[1]
            if color:
                # 检查是否为支持的颜色
                if color in COLOR_MAP.values():
                    color_var.set(REVERSE_COLOR_MAP.get(color, ""))
                else:
                    messagebox.showwarning("警告", "请选择以下颜色之一: 黑色、白色、红色、绿色、蓝色、黄色、紫色")
        
        color_btn = ttk.Button(main_frame, text="选择颜色", command=choose_color)
        color_btn.grid(row=2, column=2, sticky=W, padx=5, pady=5)
        
        # 图片文件
        image_label = ttk.Label(main_frame, text="图片路径:")
        image_label.grid(row=3, column=0, sticky=W, padx=5, pady=5)
        
        image_var = StringVar()
        image_entry = ttk.Entry(main_frame, textvariable=image_var, width=30)
        image_entry.grid(row=3, column=1, sticky=W+E, padx=5, pady=5)
        
        def browse_image():
            file_path = filedialog.askopenfilename(
                title="选择图片文件",
                filetypes=[
                    ("图片文件", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                    ("所有文件", "*.*")
                ]
            )
            if file_path:
                image_var.set(file_path)
        
        image_btn = ttk.Button(main_frame, text="浏览...", command=browse_image)
        image_btn.grid(row=3, column=2, sticky=W, padx=5, pady=5)
        
        # 语音文件
        voice_label = ttk.Label(main_frame, text="语音路径:")
        voice_label.grid(row=4, column=0, sticky=W, padx=5, pady=5)
        
        voice_var = StringVar()
        voice_entry = ttk.Entry(main_frame, textvariable=voice_var, width=30)
        voice_entry.grid(row=4, column=1, sticky=W+E, padx=5, pady=5)
        
        def browse_voice():
            file_path = filedialog.askopenfilename(
                title="选择语音文件",
                filetypes=[
                    ("音频文件", "*.mp3;*.wav"),
                    ("所有文件", "*.*")
                ]
            )
            if file_path:
                voice_var.set(file_path)
        
        voice_btn = ttk.Button(main_frame, text="浏览...", command=browse_voice)
        voice_btn.grid(row=4, column=2, sticky=W, padx=5, pady=5)
        
        # 朗读文本
        read_label = ttk.Label(main_frame, text="朗读文本:")
        read_label.grid(row=5, column=0, sticky=W, padx=5, pady=5)
        
        read_var = StringVar()
        read_entry = ttk.Entry(main_frame, textvariable=read_var, width=30)
        read_entry.grid(row=5, column=1, sticky=W+E, padx=5, pady=5)
        
        # 帮助信息
        help_frame = ttk.LabelFrame(main_frame, text="配置说明")
        help_frame.grid(row=6, column=0, columnspan=3, sticky=W+E, padx=5, pady=10)
        
        help_text = """
· 显示名称: 抽取时显示的文字，留空则显示原姓名/分组名
· 文字颜色: 支持黑色、白色、红色、绿色、蓝色、黄色、紫色
· 图片路径: 抽取时显示的图片，建议使用PNG格式
· 语音路径: 抽取时播放的背景音频，支持MP3、WAV格式
· 朗读文本: 朗读时使用的文字，留空则读取显示名称
        """
        
        help_label = ttk.Label(help_frame, text=help_text, justify=LEFT)
        help_label.pack(fill=X, padx=5, pady=5)
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        def save_egg_config():
            name = name_var.get()
            if not name:
                messagebox.showwarning("警告", "请选择姓名/分组")
                return
                
            display_name = display_name_var.get()
            color_name = color_var.get()
            color = COLOR_MAP.get(color_name, "")
            image = image_var.get()
            voice = voice_var.get()
            read_text = read_var.get()
            
            # 验证图片和语音文件
            if image and not os.path.exists(image):
                messagebox.showwarning("警告", f"图片文件不存在: {image}")
                return
                
            if voice and not os.path.exists(voice):
                messagebox.showwarning("警告", f"语音文件不存在: {voice}")
                return
            
            # 准备配置数据
            egg_config = {
                "name": name
            }
            
            if display_name:
                egg_config["new_name"] = display_name
            
            if color:
                egg_config["color"] = color
            
            if image:
                egg_config["image"] = image
                
            if voice:
                egg_config["voice"] = voice
                
            if read_text:
                egg_config["s_read_str"] = read_text
            
            # 更新配置
            if edit_index is not None:
                self.config_data[key][edit_index] = egg_config
            else:
                # 检查是否已存在
                for i, egg in enumerate(self.config_data.get(key, [])):
                    if egg.get('name') == name:
                        confirm = messagebox.askyesno("确认替换", f"已存在针对 '{name}' 的彩蛋配置，是否替换？")
                        if confirm:
                            self.config_data[key][i] = egg_config
                        else:
                            return
                        break
                else:
                    # 添加新配置
                    if key not in self.config_data:
                        self.config_data[key] = []
                    self.config_data[key].append(egg_config)
            
            # 刷新显示
            self.refresh_egg_data(key)
            self.is_modified = True
            self.update_status_bar()
            
            dialog.destroy()
        
        save_btn = ttk.Button(button_frame, text="保存", command=save_egg_config)
        save_btn.pack(side=LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="取消", command=dialog.destroy)
        cancel_btn.pack(side=LEFT, padx=5)
        
        # 如果是编辑模式，填充现有数据
        if egg_data:
            name_var.set(egg_data.get('name', ''))
            display_name_var.set(egg_data.get('new_name', ''))
            
            color_code = egg_data.get('color', '')
            color_name = REVERSE_COLOR_MAP.get(color_code, '')
            color_var.set(color_name)
            
            image_var.set(egg_data.get('image', ''))
            voice_var.set(egg_data.get('voice', ''))
            read_var.set(egg_data.get('s_read_str', ''))
            
        # 设置对话框大小和位置
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    def show_input_dialog(self, title, prompt, default=""):
        """显示输入对话框"""
        dialog = Toplevel(self.root)
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=BOTH, expand=True)
        
        # 提示标签
        label = ttk.Label(main_frame, text=prompt)
        label.pack(pady=(0,5))
        
        # 输入框
        entry_var = StringVar(value=default)
        entry = ttk.Entry(main_frame, textvariable=entry_var, width=30)
        entry.pack(fill=X, pady=5)
        entry.select_range(0, END)
        entry.focus_set()
        
        # 结果变量
        result = [None]
        
        # 确定按钮
        def on_ok():
            result[0] = entry_var.get()
            dialog.destroy()
            
        def on_cancel():
            dialog.destroy()
        
        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ok_btn = ttk.Button(button_frame, text="确定", command=on_ok)
        ok_btn.pack(side=LEFT, padx=5)
        
        cancel_btn = ttk.Button(button_frame, text="取消", command=on_cancel)
        cancel_btn.pack(side=LEFT, padx=5)
        
        # 绑定回车键
        entry.bind("<Return>", lambda event: on_ok())
        
        # 设置对话框大小和位置
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'+{x}+{y}')
        
        # 等待对话框关闭
        dialog.wait_window()
        return result[0]
    
    def new_config(self):
        """创建新配置文件"""
        if self.is_modified:
            confirm = messagebox.askyesnocancel("保存更改", "当前配置已修改，是否保存更改？")
            if confirm is None:
                return  # 用户取消操作
            elif confirm:
                self.save_config()
                
        # 重置配置数据
        self.config_data = dict(CONFIG_TEMPLATE)
        self.config_path = os.path.join(os.getcwd(), "config.json")
        self.is_modified = False
        
        # 刷新显示
        self.refresh_all_data()
        self.update_status_bar()
    
    def open_config(self):
        """打开配置文件，自动检测编码"""
        if self.is_modified:
            confirm = messagebox.askyesnocancel("保存更改", "当前配置已修改，是否保存更改？")
            if confirm is None:
                return  # 用户取消操作
            elif confirm:
                self.save_config()
            
        file_path = filedialog.askopenfilename(
            title="打开配置文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
    
        if file_path:
            try:
                # 读取二进制内容
                with open(file_path, 'rb') as f:
                    content = f.read()
                
                # 检测编码
                detect_result = chardet.detect(content)
                encoding = detect_result['encoding']
                confidence = detect_result['confidence']
            
                # 记录当前文件的编码
                self.current_encoding = encoding
            
                # 如果置信度低，尝试不同编码
                if confidence < 0.7 or not encoding:
                    # 尝试多种编码
                    for enc in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                        try:
                            decoded_content = content.decode(enc)
                            config_data = json.loads(decoded_content)
                            self.current_encoding = enc
                            self.config_data = config_data
                            self.config_path = file_path
                            self.is_modified = False
                        
                            # 刷新显示
                            self.refresh_all_data()
                            self.update_status_bar()
                            self.status_bar.config(text=f"{self.status_bar['text']} (编码: {enc})")
                            return
                        except (UnicodeDecodeError, json.JSONDecodeError):
                            continue
                
                    # 所有编码都失败
                    messagebox.showerror("错误", "无法识别文件编码，请确保文件是有效的JSON格式")
                    return
            
                # 使用检测到的编码解析JSON
                decoded_content = content.decode(encoding)
                self.config_data = json.loads(decoded_content)
                self.config_path = file_path
                self.is_modified = False
            
                # 刷新显示
                self.refresh_all_data()
                self.update_status_bar()
                self.status_bar.config(text=f"{self.status_bar['text']} (编码: {encoding})")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置文件失败: {str(e)}")
    
    def save_config(self):
        """保存配置文件，使用原始编码或UTF-8"""
        try:
            # 检查是否有姓名列表
            if not self.config_data.get('names'):
                messagebox.showwarning("警告", "姓名列表不能为空，请添加至少一个姓名后再保存")
                # 自动切换到姓名管理标签页
                self.notebook.select(0)
                return False
        
            # 转换为JSON字符串    
            json_str = json.dumps(self.config_data, indent=4, ensure_ascii=False)
        
            # 使用之前检测到的编码或默认UTF-8
            encoding = getattr(self, 'current_encoding', 'utf-8')
        
            # 如果编码不是常见编码之一，默认使用UTF-8
            if encoding.lower() not in ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']:
                encoding = 'utf-8'
        
            # 以二进制模式写入
            with open(self.config_path, 'wb') as f:
                f.write(json_str.encode(encoding))
            
            self.is_modified = False
            self.update_status_bar()
            self.status_bar.config(text=f"{self.status_bar['text']} (编码: {encoding})")
            messagebox.showinfo("成功", f"配置已保存到 {self.config_path}    如果是正在使用的配置，请右键抽签器主窗口-重读配置以使改动生效！")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存配置文件失败: {str(e)}")
            return False
    
    def save_config_as(self):
        """另存为配置文件，可选择编码"""
        file_path = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
    
        if file_path:
            # 显示编码选择对话框
            encoding_dialog = Toplevel(self.root)
            encoding_dialog.title("选择文件编码")
            encoding_dialog.transient(self.root)
            encoding_dialog.grab_set()
            encoding_dialog.geometry("300x150")
        
            # 主框架
            main_frame = ttk.Frame(encoding_dialog, padding="10")
            main_frame.pack(fill=BOTH, expand=True)
        
            # 编码选择
            ttk.Label(main_frame, text="选择保存编码:").pack(pady=(0,5))
        
            encoding_var = StringVar(value=getattr(self, 'current_encoding', 'utf-8'))
            encodings = ['utf-8', 'gbk', 'gb2312', 'iso-8859-1']
            encoding_combo = ttk.Combobox(main_frame, textvariable=encoding_var, values=encodings)
            encoding_combo.pack(fill=X, pady=5)
        
            # 结果变量
            result = {"path": file_path, "encoding": None}
        
            def on_ok():
                result["encoding"] = encoding_var.get()
                encoding_dialog.destroy()
            
            # 操作按钮
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(pady=10)
        
            ok_btn = ttk.Button(button_frame, text="确定", command=on_ok)
            ok_btn.pack(side=LEFT, padx=5)
        
            cancel_btn = ttk.Button(button_frame, text="取消", command=encoding_dialog.destroy)
            cancel_btn.pack(side=LEFT, padx=5)
        
            # 居中显示对话框
            encoding_dialog.update_idletasks()
            width = encoding_dialog.winfo_width()
            height = encoding_dialog.winfo_height()
            x = (encoding_dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (encoding_dialog.winfo_screenheight() // 2) - (height // 2)
            encoding_dialog.geometry(f'+{x}+{y}')
        
            # 等待对话框关闭
            encoding_dialog.wait_window()
        
            if result["encoding"]:
                self.config_path = result["path"]
                self.current_encoding = result["encoding"]
                return self.save_config()
        
        return False
    
    def import_from_excel(self):
        """从Excel导入名单"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 创建导入对话框
            import_dialog = Toplevel(self.root)
            import_dialog.title("从Excel导入")
            import_dialog.geometry("500x400")
            import_dialog.transient(self.root)
            import_dialog.grab_set()
            
            # 主框架
            main_frame = ttk.Frame(import_dialog, padding="10")
            main_frame.pack(fill=BOTH, expand=True)
            
            # 预览数据
            preview_frame = ttk.LabelFrame(main_frame, text="数据预览")
            preview_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)
            
            # 创建Treeview用于显示Excel内容
            tree_frame = Frame(preview_frame)
            tree_frame.pack(fill=BOTH, expand=True)
            
            # 滚动条
            scrollbar_y = ttk.Scrollbar(tree_frame)
            scrollbar_y.pack(side=RIGHT, fill=Y)
            
            scrollbar_x = ttk.Scrollbar(tree_frame, orient=HORIZONTAL)
            scrollbar_x.pack(side=BOTTOM, fill=X)
            
            # 使用pandas读取Excel
            df = pd.read_excel(file_path)
            columns = list(df.columns)
            
            # 创建Treeview
            excel_tree = ttk.Treeview(
                tree_frame,
                columns=columns,
                show='headings',
                yscrollcommand=scrollbar_y.set,
                xscrollcommand=scrollbar_x.set
            )
            
            # 设置列
            for col in columns:
                excel_tree.column(col, width=100, anchor='center')
                excel_tree.heading(col, text=col)
            
            # 添加数据行
            for i, row in df.iterrows():
                values = [str(row[col]) for col in columns]
                excel_tree.insert('', END, values=values)
            
            excel_tree.pack(fill=BOTH, expand=True)
            scrollbar_y.config(command=excel_tree.yview)
            scrollbar_x.config(command=excel_tree.xview)
            
            # 导入选项框架
            options_frame = ttk.LabelFrame(main_frame, text="导入选项")
            options_frame.pack(fill=X, padx=5, pady=5)
            
            # 选择列
            column_frame = ttk.Frame(options_frame)
            column_frame.pack(fill=X, padx=5, pady=5)
            
            ttk.Label(column_frame, text="姓名列:").grid(row=0, column=0, padx=5, pady=2, sticky=W)
            name_column_var = StringVar()
            name_column_combo = ttk.Combobox(column_frame, textvariable=name_column_var, values=columns, width=15)
            name_column_combo.grid(row=0, column=1, padx=5, pady=2, sticky=W)
            if columns:
                name_column_combo.current(0)
                
            ttk.Label(column_frame, text="分组列:").grid(row=1, column=0, padx=5, pady=2, sticky=W)
            group_column_var = StringVar()
            group_column_combo = ttk.Combobox(column_frame, textvariable=group_column_var, values=[""] + columns, width=15)
            group_column_combo.grid(row=1, column=1, padx=5, pady=2, sticky=W)
            group_column_combo.current(0)  # 默认不选择分组列
            
            # 导入设置
            settings_frame = ttk.Frame(options_frame)
            settings_frame.pack(fill=X, padx=5, pady=5)
            
            merge_var = BooleanVar(value=True)
            merge_check = ttk.Checkbutton(settings_frame, text="合并到现有列表", variable=merge_var)
            merge_check.grid(row=0, column=0, padx=5, pady=2, sticky=W)
            
            # 操作按钮
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(pady=10)
            
            def do_import():
                name_column = name_column_var.get()
                group_column = group_column_var.get()
                
                if not name_column:
                    messagebox.showwarning("警告", "请选择姓名列")
                    return
                
                # 提取姓名和分组数据
                names_to_import = df[name_column].dropna().unique().tolist()
                names_to_import = [str(name).strip() for name in names_to_import if str(name).strip()]
                
                groups_to_import = []
                if group_column:
                    groups_to_import = df[group_column].dropna().unique().tolist()
                    groups_to_import = [str(group).strip() for group in groups_to_import if str(group).strip()]
                
                # 合并或替换数据
                if merge_var.get():
                    # 合并模式 - 添加不存在的项
                    current_names = self.config_data.get('names', [])
                    current_groups = self.config_data.get('groups', [])
                    
                    # 过滤已存在的名称
                    names_to_import = [name for name in names_to_import if name not in current_names]
                    groups_to_import = [group for group in groups_to_import if group not in current_groups]
                    
                    # 添加到现有列表
                    self.config_data['names'] = current_names + names_to_import
                    self.config_data['groups'] = current_groups + groups_to_import
                else:
                    # 替换模式
                    self.config_data['names'] = names_to_import
                    if groups_to_import:
                        self.config_data['groups'] = groups_to_import
                
                # 刷新显示
                self.refresh_names_data()
                self.refresh_groups_data()
                self.is_modified = True
                self.update_status_bar()
                
                messagebox.showinfo("导入成功", f"已导入 {len(names_to_import)} 个姓名" + 
                                    (f" 和 {len(groups_to_import)} 个分组" if groups_to_import else ""))
                import_dialog.destroy()
            
            import_btn = ttk.Button(button_frame, text="导入", command=do_import)
            import_btn.pack(side=LEFT, padx=5)
            
            cancel_btn = ttk.Button(button_frame, text="取消", command=import_dialog.destroy)
            cancel_btn.pack(side=LEFT, padx=5)
            
        except Exception as e:
            messagebox.showerror("错误", f"导入Excel失败: {str(e)}\n{traceback.format_exc()}")
    
    def export_to_excel(self):
        """导出名单到Excel"""
        file_path = filedialog.asksaveasfilename(
            title="导出到Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
    
        if not file_path:
            return
    
        try:
            # 检查是否有数据可以导出
            if not self.config_data.get('names') and not self.config_data.get('groups'):
                messagebox.showwarning("警告", "没有数据可以导出，请先添加姓名或分组")
                return
        
            # 创建Excel写入器
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 写入姓名表 - 单独处理姓名和分组
                names_list = self.config_data.get('names', [])
                if names_list:
                    df_names = pd.DataFrame({"姓名": names_list})
                    df_names.to_excel(writer, sheet_name='姓名列表', index=False)
            
                # 单独写入分组表
                groups_list = self.config_data.get('groups', [])
                if groups_list:
                    df_groups = pd.DataFrame({"分组": groups_list})
                    df_groups.to_excel(writer, sheet_name='分组列表', index=False)
            
                # 写入个人彩蛋表
                egg_data = []
                for egg in self.config_data.get('egg_cases', []):
                    egg_data.append({
                        '姓名': egg.get('name', ''),
                        '显示名': egg.get('new_name', ''),
                        '颜色': REVERSE_COLOR_MAP.get(egg.get('color', ''), ''),
                        '图片路径': egg.get('image', ''),
                        '语音路径': egg.get('voice', ''),
                        '朗读文本': egg.get('s_read_str', '')
                    })
                
                if egg_data:
                    df_egg = pd.DataFrame(egg_data)
                    df_egg.to_excel(writer, sheet_name='个人彩蛋', index=False)
            
                # 写入分组彩蛋表
                group_egg_data = []
                for egg in self.config_data.get('egg_cases_group', []):
                    group_egg_data.append({
                        '分组名': egg.get('name', ''),
                        '显示名': egg.get('new_name', ''),
                        '颜色': REVERSE_COLOR_MAP.get(egg.get('color', ''), ''),
                        '图片路径': egg.get('image', ''),
                        '语音路径': egg.get('voice', ''),
                        '朗读文本': egg.get('s_read_str', '')
                    })
                
                if group_egg_data:
                    df_group_egg = pd.DataFrame(group_egg_data)
                    df_group_egg.to_excel(writer, sheet_name='分组彩蛋', index=False)
            
                # 如果到此处都没有写入任何工作表，添加一个空的默认工作表
                if not (names_list or groups_list or egg_data or group_egg_data):
                    pd.DataFrame().to_excel(writer, sheet_name='空数据', index=False)
        
            messagebox.showinfo("导出成功", f"配置已导出到 {file_path}")
        
        except Exception as e:
            # 显示详细的错误信息以便调试
            error_msg = str(e)
            traceback_info = traceback.format_exc()
            messagebox.showerror("错误", f"导出到Excel失败: {error_msg}\n\n{traceback_info}")

    def create_excel_template(self):
        """生成Excel模板"""
        file_path = filedialog.asksaveasfilename(
            title="保存Excel模板",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # 创建示例数据
            name_data = {
                "姓名": ["张三", "李四", "王五"],
                "分组": ["第一组", "第二组", "第一组"]
            }
            
            personal_egg_data = {
                "姓名": ["张三"],
                "显示名": ["张三同学"],
                "颜色": ["蓝色"],
                "图片路径": ["images/sample.png"],
                "语音路径": ["audio/sample.mp3"],
                "朗读文本": ["张三同学请上台"]
            }
            
            group_egg_data = {
                "分组名": ["第一组"],
                "显示名": ["第一小组"],
                "颜色": ["红色"],
                "图片路径": ["images/group1.png"],
                "语音路径": ["audio/group1.mp3"],
                "朗读文本": ["第一小组上台"]
            }
            
            # 创建Excel写入器
            with pd.ExcelWriter(file_path) as writer:
                df_names = pd.DataFrame(name_data)
                df_names.to_excel(writer, sheet_name='姓名和分组', index=False)
                
                df_personal = pd.DataFrame(personal_egg_data)
                df_personal.to_excel(writer, sheet_name='个人彩蛋', index=False)
                
                df_group = pd.DataFrame(group_egg_data)
                df_group.to_excel(writer, sheet_name='分组彩蛋', index=False)
            
            # 创建说明文档
            instructions_path = os.path.splitext(file_path)[0] + "_说明.txt"
            instructions = """随机抽签器 - Excel模板使用说明

1. 姓名和分组表格：
   - 姓名列必填，每行填写一个人的姓名
   - 分组列选填，用于配置小组抽取功能

2. 个人彩蛋表格：
   - 姓名列必须与"姓名和分组"表中的姓名完全一致
   - 显示名：抽取时显示的名字，留空则显示原姓名
   - 颜色：支持黑色、白色、红色、绿色、蓝色、黄色、紫色
   - 图片路径：抽取时显示的图片文件路径
   - 语音路径：抽取时播放的语音文件路径
   - 朗读文本：朗读时使用的文字，留空则读取显示名称

3. 分组彩蛋表格：
   - 配置方式与个人彩蛋相同，但针对分组
   - 分组名必须与"姓名和分组"表中的分组完全一致

注意事项：
- 所有路径支持相对路径和绝对路径
- 图片推荐使用PNG格式
- 语音支持MP3和WAV格式
- 导入时，会自动检查数据合法性

© 2025 GZYZhy https://github.com/gzyzhy/Name-Random-Picker
"""
            with open(instructions_path, 'w', encoding='utf-8') as f:
                f.write(instructions)
            
            messagebox.showinfo("成功", f"模板已保存到 {file_path}\n使用说明已保存到 {instructions_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成Excel模板失败: {str(e)}")
    
    def show_help(self):
        """显示帮助信息"""
        help_dialog = Toplevel(self.root)
        help_dialog.title("随机抽签器 - 配置编辑器使用说明")
        help_dialog.geometry("700x500")
        help_dialog.transient(self.root)
        
        # 创建滚动文本区域
        text_frame = Frame(help_dialog)
        text_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=RIGHT, fill=Y)
        
        text = Text(text_frame, wrap=WORD, yscrollcommand=scrollbar.set)
        text.pack(fill=BOTH, expand=True)
        scrollbar.config(command=text.yview)
        
        # 帮助内容
        help_content = """随机抽签器 配置文件编辑器 使用说明

【基本操作】
1. 文件操作
   • 新建配置：创建新的空白配置文件
   • 打开配置：打开已有的配置文件
   • 保存配置：保存当前配置到文件
   • 另存为：将配置保存到新文件

2. 姓名管理（第一个标签页）
   • 添加：新增姓名到列表
   • 删除：移除选中的姓名
   • 清空：删除所有姓名
   • 双击姓名可以编辑

3. 分组管理（第二个标签页）
   • 添加：新增分组到列表
   • 删除：移除选中的分组
   • 清空：删除所有分组
   • 双击分组可以编辑

4. 彩蛋配置（第三、四个标签页）
   • 个人彩蛋：设置特定姓名的特殊效果
   • 分组彩蛋：设置特定分组的特殊效果
   • 支持配置显示名称、文字颜色、图片、语音和朗读文本

5. 导入导出
   • 从Excel导入名单：批量导入姓名和分组
   • 导出名单到Excel：将当前配置导出为Excel表格
   • 生成Excel模板：创建标准Excel导入模板

【配置文件说明】
1. 基本结构
   配置文件为JSON格式，包含以下主要部分：
   • names: 姓名列表
   • groups: 分组列表
   • egg_cases: 个人彩蛋配置
   • egg_cases_group: 分组彩蛋配置

2. 彩蛋配置项
   • name: 姓名或分组名（必须与列表中的完全一致）
   • new_name: 显示名称，如不设置则显示原名
   • color: 文字颜色，支持black/white/red/green/blue/yellow/purple
   • image: 图片文件路径
   • voice: 语音文件路径
   • s_read_str: 朗读文本，不设置则读取显示名称

3. 文件路径说明
   • 支持相对路径和绝对路径
   • 相对路径基于程序运行目录
   • 建议将图片和语音文件集中存放在特定文件夹

【注意事项】
• 姓名和分组不允许重复
• 图片建议使用PNG格式，支持透明背景
• 语音文件支持MP3和WAV格式
• 彩蛋设置中的姓名/分组必须已在列表中存在
• 测试彩蛋效果可以在主程序中使用"测试模式"

【技巧提示】
• 使用Excel导入功能可以批量管理大量名单
• 使用"另存为"功能可以为不同场景准备不同配置
• 彩蛋可以只设置部分属性，不必全部填写
• 图片和音频文件路径请确保准确无误
"""
        
        text.insert("1.0", help_content)
        text.config(state=DISABLED)  # 设置为只读
        
        # 添加确定按钮
        button_frame = Frame(help_dialog)
        button_frame.pack(pady=10)
        
        ok_button = ttk.Button(button_frame, text="确定", command=help_dialog.destroy)
        ok_button.pack()
    
    def show_about(self):
        """显示关于信息"""
        about_dialog = Toplevel(self.root)
        about_dialog.title("关于 随机抽签器配置编辑器")
        about_dialog.transient(self.root)
        about_dialog.resizable(False, False)
        
        # 主框架
        main_frame = ttk.Frame(about_dialog, padding=20)
        main_frame.pack()
        
        # 标题
        title_label = ttk.Label(main_frame, text="随机抽签器 配置文件编辑器", font=("", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # 版本信息
        version_label = ttk.Label(main_frame, text="版本: v4.0.0")
        version_label.pack(anchor=W)
        
        # 版权信息
        copyright_label = ttk.Label(main_frame, text="©2024 GZYzhy, Apache License 2.0")
        copyright_label.pack(anchor=W)
        
        # GitHub链接
        def open_github():
            import webbrowser
            webbrowser.open("https://github.com/gzyzhy/Name-Random-Picker")
        
        github_label = ttk.Label(main_frame, text="GitHub: gzyzhy/Name-Random-Picker", foreground="blue", cursor="hand2")
        github_label.pack(anchor=W, pady=(10, 0))
        github_label.bind("<Button-1>", lambda e: open_github())
        
        # 确定按钮
        ttk.Button(main_frame, text="确定", command=about_dialog.destroy).pack(pady=(20, 0))
    
    def on_closing(self):
        """窗口关闭事件处理"""
        if self.is_modified:
            # 检查姓名列表是否为空
            if not self.config_data.get('names'):
                result = messagebox.askyesno("警告", 
                    "姓名列表为空！这可能导致抽签器无法正常工作。\n\n是否继续退出而不保存？")
                if not result:
                    self.notebook.select(0)  # 切换到姓名管理标签
                    return
                self.root.destroy()
                return
            
            # 正常询问是否保存
            confirm = messagebox.askyesnocancel("保存更改", "配置已修改，是否保存更改？")
            if confirm is None:
                return  # 用户取消关闭
            elif confirm:
                saved = self.save_config()
                if not saved:
                    return  # 保存失败，取消关闭
        self.root.destroy()

def main():
    """主函数"""
    root = Tk()
    root.title("随机抽签器 - 配置文件编辑器")
    
    # 适配不同操作系统的外观
    if platform.system() == "Darwin":  # macOS
        # 在macOS上应用系统主题
        pass
    elif platform.system() == "Windows":
        # 在Windows上应用系统主题
        try:
            from ctypes import windll, byref, create_unicode_buffer, create_string_buffer
            windll.shcore.SetProcessDpiAwareness(1)
            
            # 应用系统主题
            import sys
            if sys.getwindowsversion().build >= 10000:
                try:
                    # 启用 DPI 自适应
                    from ctypes import windll
                    windll.shcore.SetProcessDpiAwareness(1)
                except:
                    pass
        except:
            pass
    
    # 设置应用图标
    try:
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except:
        pass
    
    app = ConfigEditor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
