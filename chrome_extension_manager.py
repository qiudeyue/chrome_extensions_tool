import os
import sys
import winreg
import shutil
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import logging
import re
import requests
import time

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                   format='%(asctime)s - %(levelname)s - %(message)s',
                   handlers=[logging.StreamHandler()])

class ChromeExtensionManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Chrome扩展管理工具")
        
        # 设置窗口大小和位置
        window_width = 1000
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置主题样式
        self.style = ttk.Style()
        self.style.theme_use('clam')  # 使用clam主题
        self.style.configure('TButton', padding=5, font=('微软雅黑', 10))
        self.style.configure('TLabel', font=('微软雅黑', 10))
        self.style.configure('Treeview', font=('微软雅黑', 10), rowheight=25)
        self.style.configure('Treeview.Heading', font=('微软雅黑', 10, 'bold'))
        
        # 注册表路径
        self.reg_path = r"Software\Google\Chrome\Extensions"
        self.root_key = winreg.HKEY_CURRENT_USER
        
        # 缓存文件路径
        self.cache_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extension_names.json")
        self.name_cache = self.load_name_cache()
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 创建标题
        title_label = ttk.Label(self.main_frame, text="Chrome扩展管理工具", font=('微软雅黑', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 创建按钮框架
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=1, column=0, columnspan=2, pady=(0, 10), sticky=tk.W)
        
        # 添加按钮
        ttk.Button(self.button_frame, text="添加扩展", command=self.add_extension, style='TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_frame, text="删除选中", command=self.remove_selected, style='TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_frame, text="修改选中", command=self.modify_selected, style='TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_frame, text="刷新列表", command=self.refresh_list, style='TButton').pack(side=tk.LEFT, padx=5)
        
        # 创建Treeview
        self.create_treeview()
        
        # 初始加载扩展列表
        self.refresh_list()

    def load_name_cache(self):
        """加载名称缓存"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载缓存文件失败: {str(e)}")
        return {}

    def save_name_cache(self):
        """保存名称缓存"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.name_cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存缓存文件失败: {str(e)}")

    def create_treeview(self):
        # 创建Treeview框架
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(2, weight=1)
        
        # 创建Treeview和滚动条
        self.tree = ttk.Treeview(self.tree_frame, columns=("名称", "ID", "路径", "版本", "状态"), show="headings")
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        
        # 配置Treeview
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # 设置列标题
        self.tree.heading("名称", text="扩展名称")
        self.tree.heading("ID", text="扩展ID")
        self.tree.heading("路径", text="安装路径")
        self.tree.heading("版本", text="版本号")
        self.tree.heading("状态", text="状态")
        
        # 设置列宽
        self.tree.column("名称", width=150, anchor=tk.W)
        self.tree.column("ID", width=250, anchor=tk.W)
        self.tree.column("路径", width=400, anchor=tk.W)
        self.tree.column("版本", width=150, anchor=tk.W)
        self.tree.column("状态", width=100, anchor=tk.W)
        
        # 设置交替行颜色
        self.tree.tag_configure('oddrow', background='#f0f0f0')
        self.tree.tag_configure('evenrow', background='#ffffff')
        
        # 放置组件
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
        hsb.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 配置grid权重
        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.rowconfigure(0, weight=1)

    def find_extension_id_in_name(self, name):
        """从名称中查找32位连续字母作为扩展ID"""
        matches = re.findall(r'[a-zA-Z]{32}', name)
        return matches[0].lower() if matches else ""

    def get_extension_name_from_store(self, extension_id):
        """从扩展商店获取扩展名称"""
        try:
            # 首先尝试从crxsoso.com获取
            url = f"https://www.crxsoso.com/webstore/detail/{extension_id}"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                # 从页面中提取扩展名称，使用新的HTML结构
                match = re.search(r'<div[^>]*class="name el2"[^>]*>(.*?)<!---->', response.text, re.DOTALL)
                if match:
                    name = match.group(1).strip()
                    if name:
                        print(f"√ 从crxsoso.com获取到名称: {name}")
                        return name
                
                # 如果上面的正则没匹配到，尝试其他可能的模式
                match = re.search(r'<div[^>]*class="name el2"[^>]*>(.*?)</div>', response.text, re.DOTALL)
                if match:
                    name = match.group(1).strip()
                    if name:
                        print(f"√ 从crxsoso.com获取到名称: {name}")
                        return name
            
            # 如果从crxsoso.com获取失败，尝试从Chrome Web Store获取
            # 首先获取重定向后的URL
            url = f"https://chrome.google.com/webstore/detail/{extension_id}"
            response = requests.get(url, headers=headers, timeout=5, allow_redirects=True)
            
            if response.status_code == 200:
                # 从URL中提取扩展名称
                final_url = response.url
                if '/detail/' in final_url:
                    parts = final_url.split('/detail/')
                    if len(parts) > 1:
                        name_part = parts[1].split('/')[0]
                        if name_part and name_part != extension_id:
                            name = name_part.replace('-', ' ').title()
                            print(f"√ 从Chrome Web Store URL获取到名称: {name}")
                            return name
                
                # 如果无法从URL获取，尝试从页面内容获取
                match = re.search(r'<h1 class="e-f-w">(.*?)</h1>', response.text)
                if match:
                    name = match.group(1).strip()
                    if name:
                        print(f"√ 从Chrome Web Store页面获取到名称: {name}")
                        return name
        except Exception as e:
            print(f"从扩展商店获取名称失败: {str(e)}")
        return ""

    def get_extension_name_from_manifest(self, extension_path):
        """从扩展的manifest.json中读取名称（不再使用此方法获取名称）"""
        return ""

    def get_crx_info(self, crx_path):
        """从CRX文件中读取扩展信息"""
        try:
            import zipfile
            import json
            
            # 从文件名中提取ID
            file_name = os.path.basename(crx_path)
            name_without_ext = os.path.splitext(file_name)[0]
            
            print(f"\n检查文件名: {name_without_ext}")
            
            extension_id = self.find_extension_id_in_name(name_without_ext)
            version = ""
            name = ""
            
            # 如果找到了32位连续字母
            if extension_id:
                print(f"√ 从文件名提取到扩展ID: {extension_id}")
                
                # 首先使用缓存中的名称
                cached_name = self.name_cache.get(extension_id)
                if cached_name:
                    name = cached_name
                    print(f"√ 使用缓存中的名称: {cached_name}")
                
                # 尝试从CRX文件中读取版本号
                try:
                    with zipfile.ZipFile(crx_path) as zip_ref:
                        try:
                            manifest_data = zip_ref.read('manifest.json')
                            manifest = json.loads(manifest_data.decode('utf-8'))
                            version = manifest.get('version', '')
                            if version:
                                print(f"√ 从manifest.json读取到版本号: {version}")
                        except Exception as e:
                            print(f"读取manifest.json失败: {str(e)}")
                except Exception as e:
                    print(f"打开CRX文件失败: {str(e)}")
                
                # 如果缓存中没有名称，尝试从扩展商店获取
                if not name:
                    name = self.get_extension_name_from_store(extension_id)
            else:
                print("× 文件名中未找到32位连续字母作为扩展ID")
            
            return extension_id, version, name
            
        except Exception as e:
            print(f"处理CRX文件失败: {str(e)}")
            return "", "", ""

    def get_registry_extensions(self):
        extensions = []
        
        try:
            # 打开Chrome扩展注册表路径
            with winreg.OpenKey(self.root_key, self.reg_path, 0, winreg.KEY_READ) as main_key:
                # 枚举所有子键（扩展ID）
                i = 0
                while True:
                    try:
                        extension_id = winreg.EnumKey(main_key, i)
                        # 打开扩展的子键
                        with winreg.OpenKey(main_key, extension_id) as ext_key:
                            ext_info = {
                                "name": "",  # 初始化为空
                                "id": extension_id,
                                "path": "",
                                "version": "",
                                "status": "正常"
                            }
                            
                            # 读取所有可能的值
                            try:
                                ext_info["path"] = winreg.QueryValueEx(ext_key, "path")[0]
                            except WindowsError:
                                pass
                                
                            try:
                                ext_info["version"] = winreg.QueryValueEx(ext_key, "version")[0]
                            except WindowsError:
                                pass
                            
                            # 检查文件是否存在
                            if ext_info["path"] and os.path.exists(ext_info["path"]):
                                # 首先使用缓存中的名称
                                cached_name = self.name_cache.get(extension_id)
                                if cached_name:
                                    ext_info["name"] = cached_name
                                    print(f"√ 使用缓存中的名称: {cached_name}")
                                else:
                                    # 如果缓存中没有，尝试从扩展商店获取
                                    name = self.get_extension_name_from_store(extension_id)
                                    if name:
                                        ext_info["name"] = name
                                    else:
                                        # 如果无法从扩展商店获取，则从manifest.json中读取
                                        name = self.get_extension_name_from_manifest(ext_info["path"])
                                        if name:
                                            ext_info["name"] = name
                                        else:
                                            # 最后才使用扩展ID作为名称
                                            ext_info["name"] = extension_id
                            else:
                                ext_info["status"] = "文件缺失"
                                # 如果文件不存在，使用缓存中的名称或扩展ID
                                ext_info["name"] = self.name_cache.get(extension_id, extension_id)
                            
                            extensions.append(ext_info)
                        i += 1
                    except WindowsError:
                        break
        except WindowsError as e:
            print(f"访问注册表失败: {str(e)}")
        
        return extensions

    def refresh_list(self):
        # 清空现有项目
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 获取并显示扩展列表
        extensions = self.get_registry_extensions()
        for i, ext in enumerate(extensions):
            tag = 'oddrow' if i % 2 == 0 else 'evenrow'
            self.tree.insert("", tk.END, values=(
                ext["name"],
                ext["id"],
                ext["path"],
                ext["version"],
                ext["status"]
            ), tags=(tag,))

    def add_extension(self):
        # 创建添加扩展对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("添加扩展")
        dialog.transient(self.root)
        
        # 设置对话框位置
        dialog_width = 600
        dialog_height = 250
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 创建输入框框架
        input_frame = ttk.Frame(dialog, padding="20")
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建输入框
        ttk.Label(input_frame, text="扩展名称:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(input_frame, width=50)
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="扩展ID:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        id_entry = ttk.Entry(input_frame, width=50)
        id_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="安装路径:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        path_entry = ttk.Entry(input_frame, width=50)
        path_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        def select_file():
            file_path = filedialog.askopenfilename(filetypes=[("Chrome扩展", "*.crx")])
            if file_path:
                path_entry.delete(0, tk.END)
                path_entry.insert(0, file_path)
                # 自动读取扩展信息
                extension_id, version, name = self.get_crx_info(file_path)
                if extension_id:
                    id_entry.delete(0, tk.END)
                    id_entry.insert(0, extension_id)
                if version:
                    version_entry.delete(0, tk.END)
                    version_entry.insert(0, version)
                if name:
                    name_entry.delete(0, tk.END)
                    name_entry.insert(0, name)
        
        ttk.Button(input_frame, text="浏览", command=select_file).grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Label(input_frame, text="版本号:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        version_entry = ttk.Entry(input_frame, width=50)
        version_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 创建按钮框架
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.grid(row=1, column=0, sticky=tk.E)
        
        def save_extension():
            try:
                # 验证ID格式
                if not re.match(r'^[a-zA-Z]{32}$', id_entry.get()):
                    messagebox.showerror("错误", "扩展ID格式不正确，必须是32位字母")
                    return
                
                # 创建扩展注册表项
                ext_key = winreg.CreateKeyEx(
                    self.root_key,
                    f"{self.reg_path}\\{id_entry.get().lower()}",  # 确保ID是小写
                    0,
                    winreg.KEY_WRITE
                )
                
                # 设置值
                if path_entry.get():
                    winreg.SetValueEx(ext_key, "path", 0, winreg.REG_SZ, path_entry.get())
                if version_entry.get():
                    winreg.SetValueEx(ext_key, "version", 0, winreg.REG_SZ, version_entry.get())
                
                winreg.CloseKey(ext_key)
                
                # 保存名称到缓存
                if name_entry.get():
                    self.name_cache[id_entry.get()] = name_entry.get()
                    self.save_name_cache()
                
                messagebox.showinfo("成功", "扩展已成功添加！")
                dialog.destroy()
                self.refresh_list()
                
            except Exception as e:
                messagebox.showerror("错误", f"添加扩展失败: {str(e)}")
        
        ttk.Button(button_frame, text="保存", command=save_extension).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def modify_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要修改的扩展")
            return
        
        # 获取选中项的值
        values = self.tree.item(selected_items[0])["values"]
        
        # 创建修改对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("修改扩展")
        dialog.transient(self.root)
        
        # 设置对话框位置
        dialog_width = 600
        dialog_height = 250
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog_width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog_height) // 2
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 创建输入框框架
        input_frame = ttk.Frame(dialog, padding="20")
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建输入框
        ttk.Label(input_frame, text="扩展名称:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        name_entry = ttk.Entry(input_frame, width=50)
        name_entry.insert(0, values[0])
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="扩展ID:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        id_entry = ttk.Entry(input_frame, width=50)
        id_entry.insert(0, values[1])
        id_entry.configure(state='readonly')
        id_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(input_frame, text="安装路径:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        path_entry = ttk.Entry(input_frame, width=50)
        path_entry.insert(0, values[2])
        path_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        def select_file():
            file_path = filedialog.askopenfilename(filetypes=[("Chrome扩展", "*.crx")])
            if file_path:
                path_entry.delete(0, tk.END)
                path_entry.insert(0, file_path)
                # 自动读取扩展信息
                _, version, name = self.get_crx_info(file_path)
                if version:
                    version_entry.delete(0, tk.END)
                    version_entry.insert(0, version)
                if name:
                    name_entry.delete(0, tk.END)
                    name_entry.insert(0, name)
        
        ttk.Button(input_frame, text="浏览", command=select_file).grid(row=2, column=2, padx=5, pady=5)
        
        ttk.Label(input_frame, text="版本号:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        version_entry = ttk.Entry(input_frame, width=50)
        version_entry.insert(0, values[3])
        version_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 创建按钮框架
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.grid(row=1, column=0, sticky=tk.E)
        
        def save_changes():
            try:
                # 打开扩展注册表项
                ext_key = winreg.OpenKey(
                    self.root_key,
                    f"{self.reg_path}\\{id_entry.get()}",
                    0,
                    winreg.KEY_WRITE
                )
                
                # 更新值
                if path_entry.get():
                    winreg.SetValueEx(ext_key, "path", 0, winreg.REG_SZ, path_entry.get())
                if version_entry.get():
                    winreg.SetValueEx(ext_key, "version", 0, winreg.REG_SZ, version_entry.get())
                
                winreg.CloseKey(ext_key)
                
                # 更新名称缓存
                new_name = name_entry.get()
                if new_name != values[0]:  # 如果名称有变化
                    self.name_cache[id_entry.get()] = new_name
                    self.save_name_cache()
                
                # 更新Treeview中的显示
                self.tree.item(selected_items[0], values=(
                    new_name,  # 更新名称
                    values[1],  # ID保持不变
                    path_entry.get(),
                    version_entry.get(),
                    values[4]  # 状态保持不变
                ))
                
                messagebox.showinfo("成功", "扩展信息已更新！")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("错误", f"修改扩展失败: {str(e)}")
        
        ttk.Button(button_frame, text="保存", command=save_changes).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def remove_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的扩展")
            return
            
        if not messagebox.askyesno("确认", "确定要删除选中的扩展吗？"):
            return
            
        for item in selected_items:
            values = self.tree.item(item)["values"]
            extension_id = values[1]
            
            try:
                # 删除注册表项
                try:
                    winreg.DeleteKey(self.root_key, f"{self.reg_path}\\{extension_id}")
                    print(f"删除注册表项成功: {extension_id}")
                except WindowsError as e:
                    print(f"删除注册表项失败: {str(e)}")
                    continue
                
                # 从缓存中删除名称
                if extension_id in self.name_cache:
                    del self.name_cache[extension_id]
                    self.save_name_cache()
                
            except Exception as e:
                messagebox.showerror("错误", f"删除扩展 {extension_id} 失败: {str(e)}")
        
        messagebox.showinfo("成功", "已删除选中的扩展")
        self.refresh_list()

def main():
    root = tk.Tk()
    app = ChromeExtensionManager(root)
    root.mainloop()

if __name__ == "__main__":
    main() 