import openpyxl as op


class ExcelOperation:
    """
    - value:
    - row:
    - col:
    - template: boolean 类型 主要判断时候采用模板 一般是默认不采用
    """

    __slots__ = ["col", "default", "path", "row", "workbook", "worksheet"]

    def __init__(self, path: str, default, row=None, col=None, template=False):
        self.workbook = op.load_workbook(path)
        self.worksheet = self.workbook.active
        self.path = path
        self.default = default
        self.row = row if row and template == False else ""
        self.col = col if col and template == False else ""
        # 这里我先不写采用摹本的方法和类 等 excel 操作写完再写

    def add_sheet_order_prefix(self):

        for s in self.workbook:
            pass

    def merge_cell(self):
        """合并指定表格"""
        ws = self.workbook.active

        if ws:
            ws.merge_cells(f"{self.row}:{self.col}")

    def merge_sheet_cell(self):

        if self.worksheet:
            cell = self.worksheet[self.row]

            if self.default == "active":
                cell.value = ""
            else:
                pass


class StyleOperation:
    pass


while True:
    name = "sisis"
    match name:
        case "sisis":
            name = ""
