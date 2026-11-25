"""DocGrab - 批量提取 Office 文档纯文本与表格数据的核心逻辑。

不依赖 Office 也能处理 .docx/.xlsx/.pptx；.doc/.xls 旧格式在检测到
本机装有 Office 时通过 COM 自动转换后再解析。
"""