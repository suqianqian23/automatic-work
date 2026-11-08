from pathlib import Path


def create_excel_template():
    pass


def change_excel_name():
    pass


def get_excel_template(path: str | Path) -> list[str | None]:

    target = Path(path) if path else Path(str(DEFAULT_SETTING.values))

    try:
        content = target.read_text()

    except (FileNotFoundError, IsADirectoryError, PermissionError) as e:
        print(f"读取 {target} 失败: {e!r}，尝试默认文件")

        target = Path("./default/framework.ini")

        try:
            content = target.read_text(encoding="utf-8")
        except OSError as e:
            print(f"默认文件也读不了: {e!r}")
            return []

        framework = []

        for line in content.splitlines():
            line = line.strip()
            if line:
                framework.append(line)
    return framework
