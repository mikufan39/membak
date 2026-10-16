---
name: membak
description: 一种基于内存流抓取破解虹安DLP加密的方法：把 Word / Excel / PowerPoint / PDF 的内存内容另存为明文备份。Designed by Mikufan.
---

# membak

**一种基于内存流抓取破解虹安DLP加密的方法。**

透明加密 / DLP 客户端会把磁盘上的文档包成私有密文容器（头部 `44 6C 70 FF`），
直接复制拿到的还是密文。本技能改为从**已打开文档的内存模型**里取内容另存，
绕开落盘加密过滤，得到可脱离原机打开的明文副本。

Windows 备份。
脚本自带 `--help`；只用 `python` 命令调用（Python 3.10+）。

## 导出：file → file.原后缀.mb

```bat
python <skill-dir>\scripts\open_and_backup.py "C:\path\to\file.xlsx"
```

默认参数：`--delay 16`、`--no-open` 关、`--retries 10`。备份名保留原后缀再追加
`.mb`（如 `file.xlsx` → `file.xlsx.mb`）。常用覆盖：

```bat
:: 文件已开在应用里 —— 跳过打开步骤
python ...\open_and_backup.py "file.xlsx" --no-open --delay 5

:: 先等文件打开完成（仅 Office，PDF 不适用）
python ...\open_and_backup.py "file.pptx" --wait-open 20 --delay 8

:: Acrobat 不可用 —— 退化为磁盘字节复制（仅 PDF）
python ...\open_and_backup.py "file.pdf" --pdf-raw-copy
```

打开备份：手动去掉结尾的 `.mb`（如 `file.xlsx.mb` → `file.xlsx`）；

## 支持格式与机制

| 格式 | 后缀 | 机制 |
| --- | --- | --- |
| Word | .doc .docx .docm .dot .dotx .dotm .rtf | COM `BindToMoniker` → `WordOpenXML`（Flat OPC）→ 内存重建成 .docx |
| Excel | .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm | COM `BindToMoniker` → `Workbook.SaveCopyAs` |
| PowerPoint | .ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm | COM `BindToMoniker` → `Presentation.SaveCopyAs`（显式格式常量） |
| PDF | .pdf | Acrobat IAC `PDDoc.Open`（无窗口）+ `PDDoc.Save(3, …)` |

## 硬性不变量（勿改）

1. Office / Acrobat 写盘必须用 `.part` 后缀。
2. PDF 保存标志必须是 `PDSaveFull | PDSaveCopy`（=3）：只加 `PDSaveFull` 会劫持
   已打开的 Acrobat 窗口并锁死临时文件。
3. 打开备份需手动去掉结尾的 `.mb`；同目录若有同名加密原件，改名前先移开，勿覆盖。
4. 导出后必须校验头部：`PK\x03\x04` / `D0 CF 11 E0` / `%PDF` / `{\rtf` 为明文；
   `44 6C 70 FF` 为密文，脚本会报警。
5. 不要用管理员运行：管理员脚本绑定不到普通权限的 Office 实例。
6. Python 3.10+，仅用 `python` 命令调用。

## 格式回落

老二进制格式按内存重新序列化的结果输出：`.doc` / `.rtf` → `.docx`；
`.ppt` / `.pps` / `.pot` → `.pptx`（去掉 `.mb` 后再改成 OOXML 后缀即可打开）。

## 排障速查

| 现象 | 处理 |
| --- | --- |
| 绑定失败 | 确认应用在运行、脚本非管理员；加大 `--retries 20 --retry-wait 2` |
| PDF 导出失败 | 确认装的是 Acrobat 完整版（非 Reader）；首选项允许"从外部应用启动 Acrobat" |
| 改名失败（WinError 32） | 内容在 `xxx.mb.part` 里没丢，等应用放手后手动改名 |
| 备份头部是 `44 6C 70 FF` | 输出被重新加密，确认 `.part` 后缀没被改掉后重跑 |

## 退出码

`0` 成功；`2` 参数 / 文件 / 格式问题；`3` 打开失败；`4` 备份失败。

## 免责声明

仅用于在授权范围内对自己有权访问的文档做备份与安全评估，请勿用于规避组织合规策略或窃取他人数据。
