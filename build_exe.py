# -*- coding: utf-8 -*-
"""
打包为可执行文件
"""

import os
import sys
import subprocess
import shutil


def build_exe():
    """打包为可执行文件"""
    
    print("开始打包AIDA64报告解析工具...")
    
    # 检查必要的文件
    required_files = ['main.py', 'parser_core.py', 'templates.py', 'gui.py']
    missing_files = []
    
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"错误：缺少必要的文件: {missing_files}")
        return
    
    # 创建dist目录
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # 使用PyInstaller打包
    try:
        # 检查是否安装了PyInstaller
        import PyInstaller
    except ImportError:
        print("正在安装PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # 打包命令
    cmd = [
        'pyinstaller',
        '--name=AIDA64_Report_Parser',
        '--windowed',
        '--onefile',
        '--icon=icon.ico' if os.path.exists('icon.ico') else '',
        '--add-data=config;config',
        '--add-data=templates;templates',
        'main.py'
    ]
    
    # 移除空参数
    cmd = [arg for arg in cmd if arg]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    # 执行打包
    try:
        subprocess.run(cmd, check=True)
        print("打包完成！")
        print(f"可执行文件位于: dist/AIDA64_Report_Parser.exe")
        
        # 复制配置文件到dist目录
        if os.path.exists('config'):
            dist_config_dir = os.path.join('dist', 'config')
            if os.path.exists(dist_config_dir):
                shutil.rmtree(dist_config_dir)
            shutil.copytree('config', dist_config_dir)
        
        if os.path.exists('templates'):
            dist_templates_dir = os.path.join('dist', 'templates')
            if os.path.exists(dist_templates_dir):
                shutil.rmtree(dist_templates_dir)
            shutil.copytree('templates', dist_templates_dir)
        
        # 创建输出目录
        dist_output_dir = os.path.join('dist', 'output')
        os.makedirs(dist_output_dir, exist_ok=True)
        
        print("已复制配置文件和模板到dist目录")
        
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        return False
    
    return True


def create_zip():
    """创建ZIP发布包"""
    import zipfile
    from datetime import datetime
    
    # 创建发布包
    version = "1.0.0"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"AIDA64_Report_Parser_v{version}_{timestamp}.zip"
    
    print(f"创建发布包: {zip_name}")
    
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 添加可执行文件
        exe_path = 'dist/AIDA64_Report_Parser.exe'
        if os.path.exists(exe_path):
            zipf.write(exe_path, 'AIDA64_Report_Parser.exe')
        
        # 添加配置文件
        if os.path.exists('dist/config'):
            for root, dirs, files in os.walk('dist/config'):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, 'dist')
                    zipf.write(file_path, arcname)
        
        # 添加模板文件
        if os.path.exists('dist/templates'):
            for root, dirs, files in os.walk('dist/templates'):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, 'dist')
                    zipf.write(file_path, arcname)
        
        # 添加使用说明
        readme_content = """# AIDA64报告智能解析工具 v1.0

## 功能特性
1. 批量解析AIDA64生成的系统报告文件(.txt)
2. 支持自定义提取项目
3. 多种输出模板选择
4. 导出为Excel/CSV/JSON格式
5. 可视化操作界面

## 使用方法
1. 运行 `AIDA64_Report_Parser.exe`
2. 点击"选择文件"或"选择文件夹"导入报告
3. 选择解析模板或自定义提取项目
4. 点击"开始解析"
5. 点击"导出Excel"保存结果

## 预置模板
- 标准模板：完整提取所有信息
- 精简模板：仅关键系统信息
- 硬件专用模板：仅硬件信息
- 软件专用模板：仅软件信息

## 系统要求
- Windows 7/8/10/11
- .NET Framework 4.5 或更高版本

## 注意事项
1. 确保AIDA64报告为中文版本
2. 报告文件编码需为UTF-8
3. 导出前请关闭Excel程序
"""
        
        zipf.writestr('README.txt', readme_content)
    
    print(f"发布包创建完成: {zip_name}")
    return zip_name


if __name__ == "__main__":
    print("=" * 60)
    print("AIDA64报告解析工具打包程序")
    print("=" * 60)
    
    # 打包
    if build_exe():
        # 创建ZIP发布包
        create_zip()
        
        print("\n" + "=" * 60)
        print("打包完成！")
        print("请查看 dist/ 目录下的文件")
        print("=" * 60)
        
        # 询问是否打开目录
        if input("\n是否打开dist目录？(y/n): ").lower() == 'y':
            os.startfile('dist')
    else:
        print("打包失败，请检查错误信息。")
