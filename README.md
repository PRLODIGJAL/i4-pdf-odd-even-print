# 项目 i4 · PDF 奇偶拼接打印

把一批 Word / Excel / PDF 转成 PDF 后拼成一份大文件，双面模式自动给奇数页文档补空白页，一键打印不串页。

## 两个模式

- **双面打印**：奇数页文档末尾自动加 1 张空白页凑成偶数 → 每份文档从纸的正面开始，不串页
- **单面打印**：不管奇偶页，直接按顺序拼接，不补空白页

## 目录结构

```
PDF奇偶拼接打印1.0_Mac版/        macOS 交付包
  i4_core.py                    核心引擎（转换+补白+拼接）
  i4_app.py                     界面（统一浅色系 UI）
  bootui.py                     启动控制台（各项目逐字节共用）
  setup_env.py                  环境自检 / venv 管理 / 依赖安装
  启动PDF奇偶拼接打印.command     双击入口
  _regress/                     回归自检脚本

PDF奇偶拼接打印1.0_Windows版/    Windows 交付包
  （同上共享文件 + .bat 启动器）
```

## 怎么跑

- **macOS**：双击 `启动PDF奇偶拼接打印.command`，首次自动准备环境
- **Windows**：先双击 `运行前的环境准备.bat`，以后双击 `启动PDF奇偶拼接打印.bat`

## 支持格式

| 格式 | macOS | Windows |
|---|---|---|
| Word（.doc .docx .rtf） | ✓ | ✓ |
| Excel（.xls .xlsx） | ✓ | ✓ |
| PDF（直接拼接） | ✓ | ✓ |

Word / Excel 转 PDF 需要本机装有 Microsoft Office 或 LibreOffice。

## 补白规则（双面模式）

1 页（奇）→ 加 1 张空白 → 2 页
2 页（偶）→ 不补 → 2 页
3 页（奇）→ 加 1 张空白 → 4 页
以此类推。单面模式一律不补白。
