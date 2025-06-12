import os
import sys
import winreg
import shutil
import json
from pathlib import Path

class ChromeExtensionInstaller:
    def __init__(self):
        # 扩展安装的注册表路径
        self.registry_paths = {
            'chrome_extensions': [
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\Extensions"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Wow6432Node\Google\Chrome\Extensions"),
            ],
            'policies': [
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Policies\Google\Chrome\ExtensionInstallForcelist"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Policies\Google\Chrome\ExtensionInstallAllowlist"),
            ],
            'preferences': [
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Policies\Google\Chrome\PreferenceMACs"),
            ]
        }

    def install_extension(self, extension_id, crx_path):
        """安装Chrome扩展的完整流程"""
        try:
            # 1. 复制扩展文件到固定位置
            install_path = self._copy_extension_file(crx_path)
            
            # 2. 在Chrome扩展注册表中注册
            self._register_extension(extension_id, install_path)
            
            # 3. 添加到强制安装列表
            self._add_to_forcelist(extension_id, install_path)
            
            # 4. 添加到允许列表
            self._add_to_allowlist(extension_id)
            
            print(f"扩展安装成功：{extension_id}")
            return True
        except Exception as e:
            print(f"安装失败：{str(e)}")
            return False

    def _copy_extension_file(self, crx_path):
        """复制扩展文件到程序目录"""
        # 获取程序安装目录
        program_dir = os.path.dirname(os.path.abspath(sys.executable))
        extension_dir = os.path.join(program_dir, "Extensions")
        os.makedirs(extension_dir, exist_ok=True)
        
        # 复制文件
        dest_path = os.path.join(extension_dir, os.path.basename(crx_path))
        shutil.copy2(crx_path, dest_path)
        return dest_path

    def _register_extension(self, extension_id, crx_path):
        """在Chrome扩展注册表中注册"""
        for root_key, path in self.registry_paths['chrome_extensions']:
            try:
                key = winreg.CreateKeyEx(root_key, f"{path}\\{extension_id}", 0,
                                       winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
                winreg.SetValueEx(key, "path", 0, winreg.REG_SZ, crx_path)
                winreg.SetValueEx(key, "version", 0, winreg.REG_SZ, "1.0")
                winreg.CloseKey(key)
            except WindowsError:
                continue

    def _add_to_forcelist(self, extension_id, crx_path):
        """添加到强制安装列表"""
        try:
            root_key, path = self.registry_paths['policies'][0]  # ExtensionInstallForcelist
            key = winreg.CreateKeyEx(root_key, path, 0,
                                   winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
            
            # 获取现有项数量
            count = 0
            try:
                while True:
                    winreg.EnumValue(key, count)
                    count += 1
            except WindowsError:
                pass
            
            # 添加新项
            value_name = str(count + 1)
            value_data = f"{extension_id};file://{crx_path.replace(os.sep, '/')}"
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, value_data)
            winreg.CloseKey(key)
        except WindowsError:
            pass

    def _add_to_allowlist(self, extension_id):
        """添加到允许列表"""
        try:
            root_key, path = self.registry_paths['policies'][1]  # ExtensionInstallAllowlist
            key = winreg.CreateKeyEx(root_key, path, 0,
                                   winreg.KEY_ALL_ACCESS | winreg.KEY_WOW64_64KEY)
            
            # 获取现有项数量
            count = 0
            try:
                while True:
                    winreg.EnumValue(key, count)
                    count += 1
            except WindowsError:
                pass
            
            # 添加新项
            value_name = str(count + 1)
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, extension_id)
            winreg.CloseKey(key)
        except WindowsError:
            pass

def main():
    if len(sys.argv) != 3:
        print("用法: chrome_extension_installer.py <扩展ID> <crx文件路径>")
        return
    
    extension_id = sys.argv[1]
    crx_path = sys.argv[2]
    
    installer = ChromeExtensionInstaller()
    installer.install_extension(extension_id, crx_path)

if __name__ == "__main__":
    main() 