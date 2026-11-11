# automatic work

> 内部 SOP 文档的**版本管理**与 **Excel 模板自动生成**小工具。

基于 [`openpyxl`](https://openpyxl.readthedocs.io/) 实现，零依赖数据库，所有版本信息以「文件 + 元数据工作表 + 清单文件」的方式落地，方便放进 Git 一起管理。

---

## 目录

- [它解决什么问题](#它解决什么问题)
- [功能特性](#功能特性)
- [目录结构](#目录结构)
- [环境要求](#环境要求)
- [安装](#安装)
- [快速开始](#快速开始)
- [命令说明](#命令说明)
- [SOP 版本控制规范](#sop-版本控制规范)
- [模板生成机制](#模板生成机制)
- [配置文件](#配置文件)
- [常见问题](#常见问题)
- [维护与贡献](#维护与贡献)

---

## 它解决什么问题

内部 SOP 长期存在三个痛点：

1. **版本混乱** —— 文件名里写 `_final`、`_最终版`、`_最终版2`，无法判断哪个是最新。
2. **模板不统一** —— 每个人手搓 Excel，表头、列宽、字体、页眉各不一样。
3. **改版不可追溯** —— 上个季度的 SOP 为什么改、改了哪几条，没人说得清。

本工具的做法很简单：

- **模板集中管理**，所有 SOP 从同一份 `.xlsx` 母版生成；
- **版本号写进文件**，文件名 + 内嵌元数据双重标记；
- **变更留痕**，每次生成/升级都记录到 `CHANGELOG.md` 与元数据表。

---

## 功能特性

| 特性         | 说明                                                                     |
| ------------ | ------------------------------------------------------------------------ |
| 模板驱动生成 | 从 `templates/base_template.xlsx` 读取母版，替换占位符后输出新文件       |
| 占位符替换   | 支持 `{{SOP_ID}}`、`{{VERSION}}`、`{{OWNER}}` 等，文本/表头/页眉均可替换 |
| 元数据内嵌   | 自动写入隐藏工作表 `_meta`，记录 ID、版本、生成时间、模板版本、校验值    |
| 版本清单     | 维护 `sop_versions.yaml`，`list` 一条命令看清所有 SOP 的当前版本         |
| 模板升级     | `upgrade` 用新版母版覆盖旧文件，**保留已填写的数据**，版本号自动 +1      |
| 完整性校验   | `verify` 检查占位符残留、版本号不一致、模板版本过期                      |
| 样式固化     | 统一字体、列宽、冻结窗格、数据验证下拉、打印区域，生成即合规             |

---

## 目录结构

```
automatic-work/
├── README.md
├── requirements.txt
├── .python-version
├── pyproject.toml
├── src/
│   ├── main.py       # 主脚本
│   └── template_map.yaml      # 模板 → SOP 类型 的映射
├── templates/
│   ├── base_template.xlsx      # 通用母版
│   └── inspection_v3.xlsx      # 某类 SOP 的专用母版
├── output/                     # 生成结果（建议 gitignore，或按需提交）
│   └── SOP-QA-001_来料检验_v2.1.0_20250101.xlsx
└── CHANGELOG.md
```

---

## 环境要求

- Python 3.8+
- openpyxl >= 3.1

```bash
# requirements.txt
openpyxl>=3.1.0
PyYAML>=6.0
```

---

## 安装

```bash
git clone <仓库地址> sop-kit
cd sop-kit
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## 快速开始

```bash
# 1. 看看现在都有哪些 SOP、各自什么版本
python sop_tool.py list

# 2. 从母版生成一份新 SOP
python sop_tool.py new \
    --id SOP-QA-001 \
    --title "来料检验作业指导书" \
    --type inspection \
    --owner "张三" \
    --dept "质量部"

# 3. 母版改版后，把存量 SOP 升级到新母版（保留数据）
python sop_tool.py upgrade --id SOP-QA-001 --to-template inspection_v3.xlsx

# 4. 体检：检查占位符残留 / 版本不一致
python sop_tool.py verify
```

输出示例：

```
$ python sop_tool.py list
SOP-ID        TITLE                    VERSION  TEMPLATE         UPDATED
SOP-QA-001    来料检验作业指导书        2.1.0    inspection_v3    2025-01-01
SOP-QA-002    过程巡检作业指导书        1.0.0    inspection_v3    2024-12-18
SOP-PR-001    不合格品处理流程          1.3.2    base_template    2024-11-30
```

---

## 命令说明

### `new` —— 生成新 SOP

| 参数        | 必填 | 说明                                        |
| ----------- | ---- | ------------------------------------------- |
| `--id`      | ✅   | SOP 编号，唯一，格式 `SOP-{部门}-{序号}`    |
| `--title`   | ✅   | 文档标题                                    |
| `--type`    | ✅   | SOP 类型，对应 `template_map.yaml` 中的 key |
| `--owner`   | ✅   | 责任人                                      |
| `--dept`    |      | 所属部门                                    |
| `--version` |      | 初始版本，默认 `1.0.0`                      |
| `--out`     |      | 输出目录，默认 `output/`                    |

### `upgrade` —— 模板升级

用新母版重建文档结构，把旧文件里**已填写的内容按表头/字段名映射回填**，版本号默认 `MINOR +1`。

```bash
python sop_tool.py upgrade --id SOP-QA-001 --to-template inspection_v3.xlsx --bump minor
```

### `list` —— 列出版本

从 `config/sop_versions.yaml` + 输出目录的文件名交叉比对，输出表格；加 `--check` 可同时标出缺失文件或版本不一致的条目。

### `verify` —— 完整性校验

检查三类问题，发现问题返回非 0 退出码，可挂到 CI：

1. 单元格中残留 `{{...}}` 占位符；
2. 文件名版本号 ≠ `_meta` 表中的版本号 ≠ 清单中的版本号；
3. 文件使用的模板版本低于 `template_map.yaml` 中登记的最低要求版本。

### `diff` —— 版本对比（可选）

```bash
python sop_tool.py diff output/SOP-QA-001_..._v2.0.0.xlsx output/SOP-QA-001_..._v2.1.0.xlsx
```

逐单元格对比，输出「新增 / 删除 / 修改」三类差异。

---

## SOP 版本控制规范

### 版本号规则

采用语义化版本 `MAJOR.MINOR.PATCH`：

| 位        | 触发条件                                             | 示例              |
| --------- | ---------------------------------------------------- | ----------------- |
| **MAJOR** | 流程不兼容变更：步骤增删、责任岗位调整、判定标准改变 | 2.1.0 → **3**.0.0 |
| **MINOR** | 向后兼容的新增：新增章节、新增记录字段、补充说明     | 2.1.0 → 2.**2**.0 |
| **PATCH** | 纯文字修正：错别字、排版、措辞优化                   | 2.1.0 → 2.1.**1** |

> 原则：**看版本号就能判断要不要重新培训**。MAJOR 变更必须重新培训并通知相关岗位。

### 文件命名

```
SOP-{部门}-{序号}_{标题}_v{版本号}_{生效日期}.xlsx
SOP-QA-001_来料检验作业指导书_v2.1.0_20250101.xlsx
```

文件名中版本号与 `_meta` 表、`sop_versions.yaml` **三者必须一致**，由 `verify` 强制校验。

### 元数据工作表 `_meta`

每次生成都会写入（默认隐藏）一个 `_meta` 工作表：

| KEY              | VALUE                          |
| ---------------- | ------------------------------ |
| sop_id           | SOP-QA-001                     |
| title            | 来料检验作业指导书             |
| version          | 2.1.0                          |
| template         | inspection_v3.xlsx             |
| template_version | 3.2                            |
| owner            | 张三                           |
| created_at       | 2025-01-01T09:30:00            |
| generated_by     | sop_tool.py v1.4.0             |
| git_commit       | a1b2c3d（可选，CI 环境下写入） |
| checksum         | sha256:...（正文内容摘要）     |

### 变更留痕

- 每次 `new` / `upgrade` 自动向 `CHANGELOG.md` 追加一行；
- 建议 SOP 文件本身与 `config/` 一同提交 Git，提交信息格式：
  `feat(sop): SOP-QA-001 升级至 v2.2.0（新增批次追溯字段）`

### 生命周期状态

`sop_versions.yaml` 中每条记录带 `status` 字段：

| status       | 含义                         |
| ------------ | ---------------------------- |
| `draft`      | 草稿，未生效                 |
| `active`     | 生效中                       |
| `deprecated` | 已废止，仅存档，不建议再使用 |
| `archived`   | 归档，只读                   |

`list` 默认只显示 `draft` + `active`，`--all` 显示全部。

---

## 模板生成机制

### 占位符

母版单元格中写 `{{KEY}}`，脚本遍历全表替换。支持的键：

| 占位符               | 来源                       |
| -------------------- | -------------------------- |
| `{{SOP_ID}}`         | `--id`                     |
| `{{TITLE}}`          | `--title`                  |
| `{{VERSION}}`        | 版本号                     |
| `{{OWNER}}`          | `--owner`                  |
| `{{DEPT}}`           | `--dept`                   |
| `{{DATE}}`           | 生成日期（`YYYY-MM-DD`）   |
| `{{EFFECTIVE_DATE}}` | 生效日期，默认等于生成日期 |

### 生成流程

```
读取 template_map.yaml  →  定位母版 .xlsx
        ↓
openpyxl.load_workbook()  →  遍历 sheet / row / column
        ↓
字符串匹配 {{...}}  →  替换为实际值
        ↓
样式固化（列宽 / 冻结窗格 / 数据验证 / 打印区域 / 保护）
        ↓
写入 _meta 工作表  →  保存到 output/
        ↓
更新 sop_versions.yaml + CHANGELOG.md
```

### 核心实现片段

```python
from openpyxl import load_workbook

PLACEHOLDER_RE = re.compile(r"\{\{(\w+)\}\}")

def render(ws, context: dict):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "{{" in cell.value:
                cell.value = PLACEHOLDER_RE.sub(
                    lambda m: str(context.get(m.group(1), m.group(0))),
                    cell.value,
                )
```

> 注意：合并单元格只有左上角有值，遍历时非左上角返回 `None`，无需特殊处理；但**不要**对合并区域整体赋值，否则会破坏合并。

### 样式固化

生成后统一处理，避免各人手工调整：

- 表头行加粗 + 底色 + 居中 + 自动换行
- 固定列宽（按 `template_map.yaml` 中的 `col_widths`）
- 冻结首行（`ws.freeze_panes = "A2"`）
- 关键字段加数据验证下拉（如「合格/不合格」）
- 打印设置：A4 纵向、缩放至一页宽、重复标题行
- 可选：`ws.protection.sheet = True` 锁定结构，只允许填内容区

---

## 配置文件

### `config/template_map.yaml`

```yaml
inspection:
  file: templates/inspection_v3.xlsx
  template_version: '3.2'
  min_template_version: '3.0' # 低于此版本的存量文件会被 verify 标记
  col_widths: { A: 8, B: 20, C: 30 }
  placeholders: [SOP_ID, TITLE, VERSION, OWNER, DEPT, DATE]

base:
  file: templates/base_template.xlsx
  template_version: '1.5'
  min_template_version: '1.0'
```

### `config/sop_versions.yaml`

```yaml
- id: SOP-QA-001
  title: 来料检验作业指导书
  type: inspection
  version: 2.1.0
  status: active
  owner: 张三
  dept: 质量部
  effective_date: '2025-01-01'
  file: output/SOP-QA-001_来料检验作业指导书_v2.1.0_20250101.xlsx
```

---

## 常见问题

**Q：母版改了，存量 SOP 怎么办？**
跑 `upgrade`。它会用新母版重建结构、按字段名回填旧数据，并把版本号 `MINOR +1`。升级后务必人工抽查一遍，确认字段映射没错位。

**Q：能不能不装 PyYAML？**
可以，把配置改成 JSON 即可，脚本里加个后缀判断。⚠️（按你的实现调整）

**Q：生成的文件要不要提交 Git？**
建议：`templates/`、`config/`、`CHANGELOG.md` 必须提交；`output/` 视情况——若希望留存历史版本，就提交；否则只提交「当前生效版本」，用清单追溯。

**Q：`verify` 报「占位符残留」怎么办？**
说明母版里有未在 `template_map.yaml` 的 `placeholders` 中声明的键，或 `new` 时漏传参数。补上配置或参数重新生成即可。

**Q：中文列宽/字体在老版本 Excel 里显示异常？**
openpyxl 只写样式不负责渲染，确保母版里用的是「微软雅黑」等通用字体，别用冷门字体。

---

## 维护与贡献

- 母版变更请走 MR，并在 `CHANGELOG.md` 记录 `template_version` 的变更原因；
- 修改脚本后请跑一遍 `python sop_tool.py verify`，确保存量文件不受影响；
- 新增 SOP 类型时，在 `template_map.yaml` 注册新 key 并补一份母版。

---

## License

内部使用 ⚠️（按需填写）
