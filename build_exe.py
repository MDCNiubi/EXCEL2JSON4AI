import PyInstaller.__main__
import os

# 当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 图标文件路径（如果有的话）
# icon_path = os.path.join(current_dir, 'icon.ico')

# 打包参数
params = [
    'excel_to_json.py',  # 主程序文件
    '--name=Excel转JSON工具【MDC小助手】',  # 生成的EXE名称
    '--onefile',  # 打包成单个EXE文件
    '--windowed',  # 使用窗口界面（不显示控制台）
    '--clean',  # 清理临时文件
    # f'--icon={icon_path}',  # 如果有图标文件，可以取消注释
    # '--add-data=README.md;.',  # 注释掉这一行，因为文件不存在
]

# 执行打包
PyInstaller.__main__.run(params)