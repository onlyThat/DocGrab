# DocGrab · 文档批量提取

把一个文件夹拖进去，自动提取所有 Word / Excel / PPT / 文本里的**纯文字**和**表格数据**。
不用挨个打开复制。

- 老板发来 100 份 Word 合同要汇总金额？→ **字段汇总** 一键抓出每份合同的金额、编号、日期、甲乙方，生成一张对照表。
- 网上下了一堆 Excel 想挑几列？→ **表格合并** 勾选想要的列，按列名跨文件对齐，拼成一张大表。
- 只想把文档文字导出来？→ **全文提取** 每篇一个 txt，或合并成一个大 txt。

## 下载安装

前往 [**Releases**](https://github.com/onlyThat/DocGrab/releases) 页面下载最新版：

| 文件 | 说明 |
| --- | --- |
| `DocGrab-Setup-<版本>.exe` | 安装包（推荐），双击安装，自动创建开始菜单/桌面快捷方式 |
| `DocGrab.exe` | 免安装单文件版，下载后直接双击即可 |

**系统要求**：Windows 10/11（64 位），无需安装 Python。
处理 `.doc/.xls/.ppt` 等旧格式时需本机装有 Office 或 WPS；新格式（`.docx/.xlsx/.pptx`）无需 Office。

> 程序不联网、不上传任何数据，全文提取与表格合并全程只读原文件。

## 快速开始

### 方式一：用安装包 / exe（推荐给普通用户）

1. 从 Releases 下载并安装（或直接双击 `DocGrab.exe`）
2. 把包含文档的文件夹**拖进窗口**（没装拖拽支持时点「选择文件夹…」）
3. 选页签、点导出

### 方式二：Python 源码运行

### 方式二：Python 源码运行

```bat
pip install -r requirements.txt
python main.py
```

## 三个功能

### 1. 全文提取

列出每个文档的纯文本，可预览。导出方式二选一：
- **每篇导出为 txt**：一份文档一个 `.txt`
- **合并成一个 txt**：所有文档拼成一个大文件，来源用分隔线标出

### 2. 字段汇总（Word/合同为主）

一份文档一行，自动用正则抓取：

| 字段 | 抓什么 |
| --- | --- |
| 金额 | 带 ¥/元/万元/千分位逗号的数字，万元会自动换算成元（如 `18.8万元(≈188000.00元)`）|
| 合同/编号 | 合同编号、单号、`No.xxx` 等 |
| 日期 | `2024-01-05`、`2024年1月5日`、`2024/01/05` |
| 甲乙方 | 甲方/乙方/丙方名称 |

> 裸整数（如日期里的 2024）不会被误当成金额。

导出为 `.xlsx`（带表头样式、冻结首行）或 `.csv`（带 BOM，Excel 打开不乱码）。

### 3. 表格合并（Excel 为主，Word 表格也行）

- 扫描所有文档里的表格（Excel 每个 sheet、Word 里的表格都能识别）
- 列出所有出现过的**列名**，勾选想保留的列
- 按列名**跨文件对齐**合并：列顺序不同也能对上，某文件缺这列就填空
- 结果首列自动加「来源文件」方便溯源
- 先「预览合并结果」确认，再导出 xlsx

## 支持的文件格式

| 格式 | 解析方式 |
| --- | --- |
| `.docx` `.xlsx` `.xlsm` `.pptx` | 直接解析（openpyxl / python-docx / python-pptx）|
| `.txt` `.csv` `.md` | 直接读（自动识别 UTF-8 / GBK 编码）|
| `.doc` `.xls` `.ppt`（旧格式）| 调用本机已装的 **Office / WPS** 转成新格式再解析；没装则跳过并在结果里标注 |

> 旧格式转换不修改原文件，转换在临时目录进行、用完即删。

## 安全说明

- 全程只读原文件，所有导出另存到你指定的位置
- 不联网、不上传任何数据
- 旧格式转换借本机 Office 的 COM 接口，等价于你手动用 Office「另存为」

## 从源码打包 exe

```bat
pip install -r requirements.txt
build.bat
```

产物在 `dist\DocGrab.exe`（单文件，约 15–30 MB，目标机无需装 Python）。
打包会把 `tkinterdnd2` 的原生库一并带上，所以 exe 版支持拖拽。

## 测试

生成示例文档并跑全部测试：

```bat
python tests\make_seed.py
python tests\test_core.py
python tests\test_gui_headless.py
python tests\test_drop.py
python tests\test_legacy.py    REM 需装 Office，否则自动跳过
```

- `test_core`：扫描、四格式提取、字段抓取、万元换算、列合并、导出读回
- `test_gui_headless`：隐藏窗口构造三个页签、喂数据、合并预览
- `test_drop`：拖拽路径解析 + 模拟拖入文件夹 + 后台提取完成
- `test_legacy`：用 Office 生成 `.doc/.xls` 再解析，验证旧格式链路

## 项目结构

```
main.py                 启动入口
docgrab/
├── extract.py          扫描文件夹、识别格式、提取文本与表格
├── fields.py           金额/编号/日期/甲方方的正则抓取
├── export.py           导出 txt / xlsx / csv、表格列合并
├── office_convert.py   旧格式借 Office 的 COM 转换
└── gui.py              tkinter 界面（拖拽 + 三页签）
tests/                  示例文档生成器 + 各类测试
build.bat               PyInstaller 一键打包
```

## 许可证

MIT