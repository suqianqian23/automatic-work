import os

import openpyxl as op

# 配置文件
DEFAULT_SETTING: dict = {
    "process": ["技术变更通知单", "封面", "变更履历", "PBOM清单"],
}
USER_SETTING: dict = {
    "process": [],
}


# 文件路径
FOLDER_NAME = ""
FOLDER_PATH = ""


def load_setting():

    SHEET_NAMES = []

    if DEFAULT_SETTING:
        sheet_names = DEFAULT_SETTING.get("process")
        if sheet_names:
            for i in sheet_names:
                return SHEET_NAMES.append(i)


def main():

    # 读取文件夹的内容
    for file_path in os.listdir(FOLDER_PATH):
        path = os.path.join(FOLDER_PATH, file_path)

        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                print(file_path, ":", f.read()[:50])

    chioce_file = input("请输入你想要打开的文件:")
    new_file = input("是否变更名称: 退出 exit ")

    if new_file.strip().lower() != "exit":
        if os.path.isfile(chioce_file):
            os.rename(f"{chioce_file}", f"{new_file}")
        return

    wb = op.load_workbook(chioce_file)

    for sheet_names in wb.sheetnames:
        sheet = wb[sheet_names]

        for order in range(10):
            sheet.title = ""
