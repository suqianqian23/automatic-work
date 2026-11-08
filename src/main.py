"""
1. 技术栈: openpyxl
2. 项目工具: uv git
3. 项目流程:
    - [2026-11-08 苏浅浅 author 0.1.0]
    - 根据平台不同 自动筛选当前桌面路径 并在桌面新建文件夹 **project-folder**
      自动生成 .xlsx 文件 按 [序号] + [编号] + [名称] +[版本号] 生成名称 其中
      编号 名称 由用户定义 或 按照系统默认自动生成 用户自定义的名称放置 .txt 文
      件中 然后内部的名称按照用户规定的 **sheet-name** 自动生成模板 如果用户还
      添加工序的名称 会在模板的基础上 后方增加工序 **sheet-name**

4. 当前版本: 0.1.0
"""

import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Final

# 配置
FILE_ORDER: int = 1
FILE_VERSION = re.compile(r"")
FILE_SUFFIX: str = ".xlsx"
MAX_SHEET_NAME: Final[int] = 31

DEFAULT_SETTING: dict = {"default_path": "./default/framework.ini"}


def _windows_desktop() -> Path:
    import winreg

    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
    ) as key:
        value, _ = winreg.QueryValueEx(key, "Desktop")
    return Path(value)


def _linus_desktop():
    xdg = os.environ.get("XDG_DESKTOP_DIR")
    if xdg:
        return Path(xdg)

    # 2) 读用户配置文件
    cfg = Path.home() / ".config" / "user-dirs.dirs"
    if cfg.exists():
        for line in cfg.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("XDG_DESKTOP_DIR="):
                val = line.split("=", 1)[1].strip().strip('"')
                # 里面可能是 $HOME/Desktop，要展开变量
                return Path(os.path.expandvars(val))

    # 3) 调 xdg-user-dir 命令（多数桌面发行版都有）
    try:
        out = subprocess.check_output(["xdg-user-dir", "DESKTOP"], text=True).strip()
        if out and out != str(Path.home()):
            return Path(out)
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass

    # 4) 兜底
    return Path.home() / "Desktop"


def _macOS_desktop():
    return Path.home() / "Desktop"


def get_platform_name():

    if sys.platform == "win32":
        return _windows_desktop()

    elif sys.platform == "darwin":
        return _macOS_desktop()

    elif sys.platform.startswith("linux"):
        return _linus_desktop()

    return Path.home() / "Desktop"


def get_project_folder():
    pass


def create_excel_template():
    pass


def get_setting(setting: str | Path):

    target = setting if setting else Path("./default/setting.ini")


def get_user_setting(setting: str | Path):
    pass


def main():

    # 判断平台
    # 选出平台 name
    # 拼接字符串路径
    # 判断当前路径下时候存在 project-folder 的文件夹
    # 存在 返回是否再当前文件下生成
    # 不存在 直接生成
    # 遍历当前文件夹的文件名
    # 判断文件名是否按照业务规则
    # 未按照业务规则 先不编写 等后续再判断怎么编写
    # 按照业务规则 进行下一判断
    # 新建 更新 改名
    # 新建 用户输入 编号 名称 自动生成 .xlsx 文件
    # 自动生成内部框架 由用户在 .txt 内部定义的框架 实例 {sheet1: 封面， sheet2: BOM表 ...}
    # 然后用户是否传递后方的 [] 在框架的后方增加 工序 实例: 压装 电测
    # 现在初版仅进行 sheet-name 的变更
    # 后续增加完整的 excel 模板的控制
    platform_desktop_path = get_platform_name()
    project_path = platform_desktop_path / "project-folder"
    project_path.mkdir(parents=True, exist_ok=True)

    while True:
        if project_path.is_dir():
            cmd = input("退出请输入: exit 继续随意输入字符")

            if cmd.strip().lower() == "exit":
                break

            # 文件夹存在的逻辑
            file_name: str = input("请输入作业指导书名称:______")
            file_code: str = input("请输入作业指导书编号:______")

            xlsx_file_name = str(FILE_ORDER) + file_name + file_code + FILE_SUFFIX
            xlsx_file_path = project_path / xlsx_file_name

            # 读取项目文件夹 判断是否存在文件 如果存在对比名字和编号 一样 增加 后面版本
            # 不一样 新建文件 同时 还要增加用户取消的选项
            default_setting = get_setting()
            user_setting = get_user_setting()
            for sheet_name in get_excel_template(user_setting):
                pass

            from openpyxl import Workbook

            wb = Workbook()
            ws = wb.active

            xlsx_file_path.parent.mkdir(parents=True, exist_ok=True)
            wb.save(xlsx_file_path)

        else:
            pass


if __name__ == "__main__":
    main()
