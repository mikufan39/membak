# membak

**一种基于内存流抓取破解虹安DLP加密的方法**

在装有透明加密 / DLP 客户端的机器上，磁盘上的 Office、PDF 文件往往是以密文容器落盘的
（常见头部 `44 6C 70 FF`，即 ASCII 的 `Dlp`），只有在本机受信任进程里打开时才被驱动
层透明解密。直接把文件复制走，拿到的还是密文，脱离原机环境就打不开。

`membak` 绕开磁盘字节，改为**让应用自己把内存里的文档重新序列化一份出来**：
脚本先让文件在原应用（Word / Excel / PowerPoint / Acrobat）里正常打开，再通过 COM
自动化从**已打开文档对象的内存模型**里取出内容并另存到同目录，得到一个明文副本。
整个过程不读磁盘字节、不建卷影副本、不需要管理员权限。

## 原理

| 格式 | 机制 |
| --- | --- |
| Word（`.doc` `.docx` `.docm` `.dot` `.dotx` `.dotm` `.rtf`） | COM `BindToMoniker` 绑定到已打开的文档对象 → 取 `Content.WordOpenXML`（Flat OPC 明文 XML）→ 在内存中重新打包成 `.docx` |
| Excel（`.xls` `.xlsx` `.xlsm` `.xlsb` `.xlt` `.xltx` `.xltm`） | COM `BindToMoniker` → `Workbook.SaveCopyAs`，由 Excel 自己按内存模型序列化 |
| PowerPoint（`.ppt` `.pptx` `.pptm` `.pps` `.ppsx` `.ppsm` `.pot` `.potx` `.potm`） | COM `BindToMoniker` → `Presentation.SaveCopyAs`（显式格式常量 + 原后缀回落） |
| PDF（`.pdf`） | Acrobat IAC `AcroExch.PDDoc.Open`（无窗口）+ `PDDoc.Save(3, …)`（`PDSaveFull | PDSaveCopy`） |

关键点：

- 加密驱动的落盘过滤只作用在**文件写入路径**上；应用内存中的文档模型是明文，
  由应用自己另存出来的文件同样是明文。这就是"内存流抓取"绕过 DLP 加密的实质。
- Office 三件套用 `Runtime.InteropServices.Marshal.BindToMoniker` 直接绑定**已经打开**的
  文档实例（比 `GetActiveObject` +遍历集合更稳），拿到的是用户正在看的那份内存状态，
  连未保存的修改也一并带出来。
- Word 没有可用的 `SaveCopyAs`，改走 `WordOpenXML`（Flat OPC）——把内存里的包用
  Python 重新 zip 成 `.docx`，因此老二进制 `.doc` / `.rtf` 的输出统一是 `.docx`。
- PDF 用 Acrobat 的无窗口 IAC 接口，脚本本身不读磁盘字节。保存标志必须是
  `PDSaveFull | PDSaveCopy`（= 3）：只加 `PDSaveFull` 会劫持已打开的 Acrobat 窗口并锁死临时文件。

## 环境要求

- Windows + Python 3.10 以上（只用 `python` 命令调用）
- 对应格式的桌面应用已安装：Microsoft Office（Word / Excel / PowerPoint）、
  Adobe Acrobat 完整版（PDF 用，Reader 不含 IAC 接口）
- **不要用管理员身份运行**：管理员进程绑定不到普通权限下启动的 Office 实例

## 用法

```bat
python scripts\open_and_backup.py "C:\path\to\file.xlsx"
```

默认参数：`--delay 16`、`--no-open` 关、`--retries 10`。
备份名保留原后缀再追加 `.mb`（如 `file.xlsx` → `file.xlsx.mb`）。

```bat
:: 文件已经开在应用里 —— 跳过打开步骤
python scripts\open_and_backup.py "file.xlsx" --no-open --delay 5

:: 先等文件打开完成（仅 Office，PDF 不适用）
python scripts\open_and_backup.py "file.pptx" --wait-open 20 --delay 8

:: Acrobat 不可用 —— 退化为磁盘字节复制（仅 PDF）
python scripts\open_and_backup.py "file.pdf" --pdf-raw-copy

python scripts\open_and_backup.py --help
```

**打开备份**：手动去掉结尾的 `.mb`（`file.xlsx.mb` → `file.xlsx`）。
同目录若存在同名加密原件，改名前先移开，不要覆盖。

输出成功后会打印 `SHA-256` 与头部校验，并明确标注这一份是**内存明文**还是
**又被包成密文**（头部 `44 6C 70 FF` 会告警）。

## 硬性不变量（勿改）

1. Office / Acrobat 写盘必须用 `.part` 后缀，写完再原子改名。
2. PDF 保存标志必须是 `PDSaveFull | PDSaveCopy`（= 3）。
3. 导出后必须校验头部：`PK\x03\x04` / `D0 CF 11 E0` / `%PDF` / `{\rtf` 为明文；
   `44 6C 70 FF` 为密文、重新被加密了。
4. 不要用管理员运行。

## 格式回落

老二进制格式按内存重新序列化的结果输出：`.doc` / `.rtf` → `.docx`；
`.ppt` / `.pps` / `.pot` → `.pptx`（去掉 `.mb` 后把后缀改成对应 OOXML 后缀即可打开）。

## 排障速查

| 现象 | 处理 |
| --- | --- |
| 绑定失败 | 确认应用在运行、脚本非管理员；加大 `--retries 20 --retry-wait 2` |
| PDF 导出失败 | 确认装的是 Acrobat 完整版（非 Reader）；首选项里允许"从外部应用启动 Acrobat" |
| 改名失败（WinError 32） | 内容在 `xxx.mb.part` 里没丢，等应用放手后手动改名 |
| 备份头部是 `44 6C 70 FF` | 输出被重新加密，确认 `.part` 后缀没被改掉后重跑 |

## 退出码

`0` 成功；`2` 参数 / 文件 / 格式问题；`3` 打开失败；`4` 备份失败。

## 目录结构

```
membak/
├── SKILL.md                     # 技能说明（Agent Skills 格式）
├── agents/openai.yaml           # 技能元数据
└── scripts/open_and_backup.py   # 主程序
```

## 免责声明

本项目仅用于**在授权范围内**对自己有权访问的文档做备份、迁移与兼容性研究，
也可用于安全评估中验证透明加密 / DLP 方案的实际防护边界。
请勿用于规避所在组织的合规策略或窃取他人数据；使用者需自行承担由此产生的一切责任。
