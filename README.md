# 项目 i4 · PDF 奇偶拼接打印

把一批 Word / Excel / PDF 转成 PDF 后拼成一份大文件，双面模式自动给奇数页文档补空白页，一键打印不串页。

## 两个模式

- **双面打印**：奇数页文档末尾自动加 1 张空白页凑成偶数 → 每份文档从纸的正面开始，不串页
- **单面打印**：不管奇偶页，直接按顺序拼接，不补空白页

## 支持的输入

- Word（.doc / .docx / .docm / .rtf）
- Excel（.xls / .xlsx / .xlsm）
- PDF（.pdf，直接参与拼接，无需转换）
- 支持拖入整个文件夹，自动提取里面的 Word / Excel / PDF

## 目录结构

```
mac/          macOS 交付包
windows/      Windows 交付包
```

每个平台内部采用四层架构：

```
app/
  paths.py        路径来源
  ui/i4_app.py    界面（统一浅色系 UI）
  engine/i4_core.py  引擎（转换 + 补白 + 拼接）
  shell/          环境自检 / 启动
```

## 怎么跑

| 平台 | 入口 |
|---|---|
| macOS | `mac/启动PDF奇偶拼接打印.command` |
| Windows | 先 `运行前的环境准备.bat`，再 `启动PDF奇偶拼接打印.bat` |

## 说明

- PDF 直接拼不用转；Word/Excel 需本机装 Office 或 LibreOffice
- 原始文件只读，不改动
