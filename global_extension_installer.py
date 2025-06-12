import os
import sys
import winreg
import shutil
import json
from tkinter import Tk, filedialog
from pathlib import Path

def select_crx_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择Chrome扩展文件",
        filetypes=[("Chrome扩展文件", "*.crx")]
    )
    return file_path if file_path else None

def get_extension_id(crx_path):
    # 从文件名中提取扩展ID（如果文件名包含ID）
    extension_id = os.path.splitext(os.path.basename(crx_path))[0]
    return extension_id

def install_global_extension(crx_path):
    if not crx_path or not os.path.exists(crx_path):
        print("错误：未选择有效的扩展文件")
        return False

    try:
        # 获取扩展ID
        extension_id = get_extension_id(crx_path)
        
        # 创建存储扩展的目录
        app_data = os.environ.get('LOCALAPPDATA')
        extension_dir = os.path.join(app_data, 'Chrome_Extensions')
        os.makedirs(extension_dir, exist_ok=True)
        
        # 复制扩展文件到指定目录
        dest_path = os.path.join(extension_dir, f"{extension_id}.crx")
        shutil.copy2(crx_path, dest_path)
        
        # 修改注册表
        key_path = r"Software\Policies\Google\Chrome\ExtensionInstallForcelist"
        
        try:
            # 尝试创建或打开注册表项
            key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path, 0, 
                                   winreg.KEY_ALL_ACCESS)
            
            # 获取现有值的数量
            try:
                count = 0
                while True:
                    winreg.EnumValue(key, count)
                    count += 1
            except WindowsError:
                pass
            
            # 添加新的扩展
            value_name = str(count + 1)
            value_data = f"{extension_id};file://{dest_path.replace(os.sep, '/')}"
            
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, value_data)
            winreg.CloseKey(key)
            
            print(f"成功：扩展已添加到全局安装列表")
            print(f"扩展ID: {extension_id}")
            print(f"安装路径: {dest_path}")
            return True
            
        except Exception as e:
            print(f"错误：修改注册表失败 - {str(e)}")
            print("请确保以管理员权限运行此程序")
            return False
            
    except Exception as e:
        print(f"错误：安装过程中出现异常 - {str(e)}")
        return False

def main():
    print("Chrome扩展全局安装工具")
    print("请选择要安装的.crx扩展文件...")
    
    crx_path = select_crx_file()
    if crx_path:
        if install_global_extension(crx_path):
            print("\n安装完成！请重启Chrome浏览器以应用更改。")
        else:
            print("\n安装失败！请检查错误信息并重试。")
    else:
        print("已取消操作。")

if __name__ == "__main__":
    main() 