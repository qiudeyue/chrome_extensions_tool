import os
import subprocess
import shutil

def build_exe():
    print("开始打包程序...")
    
    # 清理之前的构建文件
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
    
    # 删除spec文件
    if os.path.exists('chrome_extension_manager.spec'):
        os.remove('chrome_extension_manager.spec')
    
    # 构建命令
    cmd = [
        'pyinstaller',
        '--noconfirm',
        '--onefile',
        '--windowed',
        '--uac-admin',  # 请求管理员权限
        '--name', 'Chrome扩展管理工具',
        'chrome_extension_manager.py'
    ]
    
    try:
        # 使用subprocess.Popen获取详细输出
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True
        )
        
        # 实时输出构建信息
        stdout, stderr = process.communicate()
        
        print("\n=== 构建输出 ===")
        if stdout:
            print(stdout)
        if stderr:
            print("错误信息：")
            print(stderr)
            
        if process.returncode == 0:
            print("\n打包完成！")
            print("可执行文件位置：dist/Chrome扩展管理工具.exe")
            
            # 复制README到dist目录
            if os.path.exists('README.md'):
                shutil.copy2('README.md', 'dist/使用说明.md')
                print("使用说明已复制到可执行文件目录")
            return True
        else:
            print(f"\n打包失败：进程返回代码 {process.returncode}")
            return False
            
    except Exception as e:
        print(f"\n发生错误：{str(e)}")
        return False

if __name__ == "__main__":
    build_exe() 