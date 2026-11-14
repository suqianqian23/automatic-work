import openpyxl as op

excel_url = r""


def created_excel(sheet_names: list | None, count: int):

    wb = op.Workbook()
    ws = wb.active

    if sheet_names is None:
        print()

        if ws is not None:
            wb.remove(ws)

        for i in range(count):
            size = str(i)
            wb.create_sheet(title=size)

    else:
        for name in sheet_names:
            wb.create_sheet(name)


def load_excel(url: str, mode: str):
    """
    url: 当前文件的 url 位置

    """


def main():
    pass


if __name__ == "__main__":
    pass
