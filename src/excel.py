import openpyxl as op


class ExcelOperation:
    def __init__(self, value):
        self.value = op.load_workbook(value)
        self.path = value  # 记住原路径，方便保存

    def add_sheet_order_prefix(self, prefix="{i}_", save_path=None, save=True):
        """
        给所有 sheet 按顺序（0 ~ 最后一页）添加前缀
        :param prefix:    前缀模板，{i} 为从 0 开始的序号
                          例: "{i}_"      -> 0_名称 1_名称
                              "{i:02d}-"  -> 00-名称 01-名称
                              "P{i}_"     -> P0_名称 P1_名称
        :param save_path: 保存路径，默认覆盖原文件
        :param save:      是否写回磁盘
        :return:          [(旧名, 新名), ...]
        """
        sheets = self.value.worksheets
        if not sheets:
            return []

        old_titles = [ws.title for ws in sheets]
        new_titles = [prefix.format(i=i) + name for i, name in enumerate(old_titles)]

        # ---- 校验 ----
        invalid = set("[]:*?/\\")
        for t in new_titles:
            if not t:
                raise ValueError("工作表名不能为空")
            if len(t) > 31:
                raise ValueError(f"超过31字符：{t!r}")
            if invalid & set(t):
                raise ValueError(f"含非法字符：{t!r}")
        if len(set(new_titles)) != len(new_titles):
            raise ValueError("重命名后出现重名")

        # ---- 两阶段重命名，避免中间过程名字冲突 ----
        existing = set(self.value.sheetnames)
        tmp_names = []
        for i in range(len(sheets)):
            name = f"__tmp_{i}__"
            while name in existing:
                name = "_" + name
            existing.add(name)
            tmp_names.append(name)

        for ws, tmp in zip(sheets, tmp_names):
            ws.title = tmp
        for ws, new in zip(sheets, new_titles):
            ws.title = new

        if save:
            self.value.save(save_path or self.path)

        return list(zip(old_titles, new_titles))
