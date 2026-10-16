from __future__ import annotations

import argparse
import base64
import hashlib
import io
import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
import zipfile

IS_WINDOWS = os.name == "nt"
CHUNK_SIZE = 1024 * 1024

KIND_WORD = "word"
KIND_EXCEL = "excel"
KIND_POWERPOINT = "powerpoint"
KIND_PDF = "pdf"

WORD_EXTENSIONS = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf",
}
EXCEL_EXTENSIONS = {
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
}
POWERPOINT_EXTENSIONS = {
    ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx", ".potm",
}
PDF_EXTENSIONS = {
    ".pdf",
}

EXT_TO_KIND = {}
for _ext in WORD_EXTENSIONS:
    EXT_TO_KIND[_ext] = KIND_WORD
for _ext in EXCEL_EXTENSIONS:
    EXT_TO_KIND[_ext] = KIND_EXCEL
for _ext in POWERPOINT_EXTENSIONS:
    EXT_TO_KIND[_ext] = KIND_POWERPOINT
for _ext in PDF_EXTENSIONS:
    EXT_TO_KIND[_ext] = KIND_PDF
del _ext

KIND_INFO = {
    KIND_WORD: (
        "Word 文档", "Word", "Word 内存（Flat OPC 重建成 .docx）",
        "Word.Application", "Documents",
    ),
    KIND_EXCEL: (
        "Excel 工作簿", "Excel", "Excel 内存（SaveCopyAs 由 Excel 自己序列化）",
        "Excel.Application", "Workbooks",
    ),
    KIND_POWERPOINT: (
        "PowerPoint 演示文稿", "PowerPoint", "PowerPoint 内存（SaveCopyAs 由 PowerPoint 自己序列化）",
        "PowerPoint.Application", "Presentations",
    ),
    KIND_PDF: (
        "PDF 文件", "Acrobat",
        "Acrobat 内存（Acrobat 无窗口打开后另存，脚本不读磁盘字节）",
        "AcroExch.PDDoc", "",
    ),
}

WAITABLE_KINDS = (KIND_WORD, KIND_EXCEL, KIND_POWERPOINT)

_PS_HEAD = (
    "$ErrorActionPreference='Stop';"
    "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
)

PS51 = os.path.join(
    os.environ.get("SystemRoot", r"C:\Windows"),
    r"System32\WindowsPowerShell\v1.0\powershell.exe",
)

def _fix_console_encoding() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            if not stream.isatty():
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

_fix_console_encoding()

def log(msg: str = "") -> None:
    try:
        print(msg, flush=True)
    except (UnicodeEncodeError, UnicodeDecodeError):
        try:
            sys.stdout.buffer.write((msg + "\n").encode("utf-8", "replace"))
            sys.stdout.buffer.flush()
        except Exception:
            pass

def is_admin() -> bool:
    if not IS_WINDOWS:
        return bool(getattr(os, "geteuid", lambda: 1)() == 0)
    try:
        import ctypes

        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def _powershell_exe() -> str:
    if os.path.exists(PS51):
        return PS51
    return shutil.which("pwsh") or "powershell.exe"

def _run_ps(script: str, timeout: float = 180.0):
    proc = subprocess.run(
        [_powershell_exe(), "-NoProfile", "-NonInteractive",
         "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=timeout,
    )
    return proc.returncode, (proc.stdout or "").strip(), (proc.stderr or "").strip()

def _ps_quote(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"

class DocumentNotOpenError(OSError):
    pass

class PublishError(OSError):
    pass

PUBLISH_RETRIES = 40
PUBLISH_WAIT = 0.25

class PsResult:

    __slots__ = ("rc", "stdout", "stderr", "bound", "ok", "err")

    def __init__(self, rc: int, stdout: str, stderr: str):
        self.rc = rc
        self.stdout = stdout
        self.stderr = stderr
        self.bound = ""
        self.ok = None
        self.err = ""
        for line in stdout.splitlines():
            line = line.strip()
            if line.startswith("ERR|"):
                self.err = line[4:].strip()
            elif line.startswith("DOC|"):
                self.bound = line[4:].strip()
            elif line.startswith("OK|"):
                self.ok = line[3:]

    @property
    def raw_detail(self) -> str:
        return (self.stderr or self.stdout or "").strip()

def _ps_common_prefix(progid: str, collection: str, path: str,
                      out: str | None = None) -> str:
    chunks = [_PS_HEAD, "$p=" + _ps_quote(path) + ";"]
    if out is not None:
        chunks.append("$o=" + _ps_quote(out) + ";")
    chunks.extend([
        "$d=$null;$err='';",
        "try{$d=[Runtime.InteropServices.Marshal]::BindToMoniker($p)}catch{};",
        "if($null -eq $d){try{",
        "$w=[Runtime.InteropServices.Marshal]::GetActiveObject(" + _ps_quote(progid) + ");",
        "$d=(@($w." + collection + ")|Where-Object{$_.FullName -eq $p}|Select-Object -First 1)",
        "}catch{}};",
        "if($null -eq $d){$err='not_open'};",
    ])
    return "".join(chunks)

def _ps_epilogue() -> str:
    return ("if($err -ne ''){Write-Output ('ERR|'+$err)}else{")

def _run_ps_task(script: str, timeout: float = 180.0) -> PsResult:
    rc, out, err = _run_ps(script, timeout=timeout)
    return PsResult(rc, out, err)

def _friendly_error(kind: str, err: str, detail: str) -> str:
    code, _, tail = err.partition(":")
    code = code.strip()
    tail = tail.strip()
    _name, app, _engine, _progid, _coll = KIND_INFO[kind]
    if code == "not_open":
        return (f"没有可用的 {app} 文档：请先在 {app} 中打开这个文件"
                f"（或让 {app} 处于运行状态）。")
    if code == "save_failed":
        return f"{app} 保存副本失败：{tail or '未知原因'}"
    if code == "acrobat_missing":
        return ("没有找到 Adobe Acrobat 的 COM 接口（AcroExch.PDDoc）。"
                "请安装 Adobe Acrobat，或用 --pdf-raw-copy 退化成磁盘字节复制。")
    if code == "acrobat_iac_disabled":
        return ("Acrobat 拒绝了外部自动化调用（IAC 被关闭）。"
                "请在 Acrobat 的首选项 → 一般 里允许“从外部应用启动 Acrobat”，"
                "或用 --pdf-raw-copy 退化成磁盘字节复制。")
    if code == "acrobat_open_failed":
        return ("Acrobat 打不开这个 PDF：文件被占用、已损坏、"
                "加密客户端没给 Acrobat 放行，或者 Acrobat 的 IAC 自动化被关闭了。")
    if code == "acrobat_save_error":
        return f"Acrobat 保存副本失败：{tail or '未知原因'}"
    if code == "acrobat_save_failed":
        return ("Acrobat 拒绝写出副本（文档可能被证书签名保护、"
                "或当前 Acrobat 版本不允许保存）。")
    return tail or detail or "未知错误"

_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
_PART_RE = re.compile(r"<pkg:part\b([^>]*)>(.*?)</pkg:part>", re.S)
_ATTR_RE = re.compile(r'(pkg:[A-Za-z]+)\s*=\s*"([^"]*)"')
_XMLDATA_RE = re.compile(r"<pkg:xmlData\b[^>]*>(.*?)</pkg:xmlData>", re.S)
_BINDATA_RE = re.compile(r"<pkg:binaryData\b[^>]*>(.*?)</pkg:binaryData>", re.S)
_CONTENT_TYPES = "[Content_Types].xml"
_DOCUMENT_PART = "word/document.xml"

def _content_types_xml(parts) -> bytes:
    lines = [
        _XML_DECL.rstrip("\r\n"),
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
    ]
    for name, ctype, _data in parts:
        if name == _CONTENT_TYPES or name.endswith(".rels"):
            continue
        if not ctype:
            ctype = "application/xml"
        lines.append(f'<Override PartName="/{name}" ContentType="{ctype}"/>')
    lines.append("</Types>")
    return ("\r\n".join(lines) + "\r\n").encode("utf-8")

def flat_opc_to_docx(xml_text: str) -> bytes:
    parts = []
    for match in _PART_RE.finditer(xml_text):
        attrs = dict(_ATTR_RE.findall(match.group(1)))
        name = attrs.get("pkg:name", "").lstrip("/")
        ctype = attrs.get("pkg:contentType", "")
        body = match.group(2)
        xml_data = _XMLDATA_RE.search(body)
        bin_data = _BINDATA_RE.search(body)
        if xml_data is not None:
            data = (_XML_DECL + xml_data.group(1)).encode("utf-8")
        elif bin_data is not None:
            data = base64.b64decode(re.sub(r"\s+", "", bin_data.group(1)))
        else:
            continue
        if name and name != _CONTENT_TYPES:
            parts.append((name, ctype, data))

    if not parts:
        raise OSError("内存中取到的文档内容为空，无法重建 .docx。")
    if not any(name == _DOCUMENT_PART for name, _c, _d in parts):
        raise OSError("内存中取到的内容缺少 word/document.xml，无法重建 .docx。")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(_CONTENT_TYPES, _content_types_xml(parts))
        for name, _ctype, data in parts:
            zf.writestr(name, data)
    log(f"    已在内存中重新打包：{len(parts)} 个部件")
    return buf.getvalue()

def word_open_xml_from_memory(path: str, timeout: float = 180.0):
    script = _ps_common_prefix("Word.Application", "Documents", os.path.abspath(path)) + (
        _ps_epilogue()
        + "Write-Output ('DOC|'+$d.FullName);"
        + "$xml=$d.Content.WordOpenXML;"
        + "$b=[System.Text.Encoding]::UTF8.GetBytes($xml);"
        + "Write-Output ('OK|'+[string]$d.Saved+'|'+[Convert]::ToBase64String($b))"
        + "}"
    )
    res = _run_ps_task(script, timeout=timeout)
    if res.ok is None:
        if res.err == "not_open":
            raise DocumentNotOpenError(
                _friendly_error(KIND_WORD, res.err, res.raw_detail))
        raise OSError(_friendly_error(KIND_WORD, res.err, res.raw_detail))
    saved_text, _, b64 = res.ok.partition("|")
    return (
        saved_text.strip().lower() == "true",
        base64.b64decode(b64).decode("utf-8"),
        res.bound,
    )

def _savecopy_script(progid: str, collection: str, path: str, out: str,
                     fmt: int | None = None) -> str:
    call = f"$d.SaveCopyAs($o,{fmt})" if fmt is not None else "$d.SaveCopyAs($o)"
    return _ps_common_prefix(progid, collection, path, out) + (
        _ps_epilogue()
        + "try{$saved=$d.Saved}catch{$saved=$null};"
        + f"try{{{call}}}catch{{"
        + "$err='save_failed: '+(($_.Exception.Message) -replace '\\s+',' ')};"
        + "if($err -ne ''){Write-Output ('ERR|'+$err)}"
        + "else{Write-Output ('DOC|'+$d.FullName);Write-Output ('OK|'+[string]$saved)}"
        + "}"
    )

def _output_missing(path: str) -> bool:
    try:
        return os.path.getsize(path) <= 0
    except OSError:
        return True

PPT_OOXML_FORMAT = {
    ".pptx": 24,
    ".pptm": 25,
    ".potx": 26,
    ".potm": 27,
    ".ppsx": 28,
    ".ppsm": 29,
}

PPT_LEGACY_FALLBACK_FORMAT = 24

def _saved_flag_is_true(text) -> bool:
    return str(text).strip().lower() in ("true", "-1", "1")

class ReadResult:

    __slots__ = ("engine", "notes", "saved", "bound", "disk_copy")

    def __init__(self, engine: str, notes=None, saved=None, bound="",
                 disk_copy: bool = False):
        self.engine = engine
        self.notes = list(notes or [])
        self.saved = saved
        self.bound = bound
        self.disk_copy = disk_copy

def read_word_memory(src: str, dst: str) -> ReadResult:
    saved, xml_text, bound = word_open_xml_from_memory(src)
    data = flat_opc_to_docx(xml_text)
    with open(dst, "wb") as f:
        f.write(data)
    return ReadResult(KIND_INFO[KIND_WORD][2], saved=saved, bound=bound)

def read_office_savecopy(kind: str, src: str, dst: str) -> ReadResult:
    _name, app, engine, progid, collection = KIND_INFO[kind]
    full = os.path.abspath(src)
    notes = []

    if kind == KIND_POWERPOINT:

        fmt = PPT_OOXML_FORMAT.get(
            os.path.splitext(src)[1].lower(), PPT_LEGACY_FALLBACK_FORMAT)
        attempts = [fmt, None]
    else:
        attempts = [None]

    res = None
    for attempt, fmt in enumerate(attempts):
        if attempt:
            notes.append(f"    显式格式常量没写出文件，改让 {app} 按原后缀自己判断…")
        res = _run_ps_task(_savecopy_script(progid, collection, full, dst, fmt))
        if res.ok is None:
            if res.err == "not_open":
                raise DocumentNotOpenError(_friendly_error(kind, res.err, res.raw_detail))
            raise OSError(_friendly_error(kind, res.err, res.raw_detail))
        if not _output_missing(dst):
            break
    else:
        raise OSError(
            f"{app} 报告保存成功，但没有写出任何内容（{os.path.basename(dst)} 不存在或为 0 字节）。"
        )

    if kind == KIND_POWERPOINT:
        out_ext = os.path.splitext(src)[1].lower()
        if out_ext in (".ppt", ".pps", ".pot"):
            notes.append(f"    [!] {out_ext} 是老二进制格式，现代 PowerPoint 不能按原格式另存，"
                         "备份内容是 OOXML(.pptx)；改回 .pptx 后缀即可打开。")
    notes.append(f"    {app} 已把内存中的内容序列化到临时文件")
    return ReadResult(engine, notes=notes, saved=_saved_flag_is_true(res.ok),
                      bound=res.bound)

_PDF_SAVE_FLAGS = 3

def read_pdf_memory(src: str, dst: str, timeout: float = 300.0) -> ReadResult:
    script = (
        _PS_HEAD
        + "$p=" + _ps_quote(os.path.abspath(src)) + ";"
        + "$o=" + _ps_quote(dst) + ";"
        + "$err='';$pages=0;$pd=$null;"
        + "try{$pd=New-Object -ComObject AcroExch.PDDoc}"
        + "catch{$err='acrobat_missing: '+(($_.Exception.Message) -replace '\\s+',' ')};"
        + "if($err -eq ''){"
        + "$ok=$false;"
        + "try{$ok=$pd.Open($p)}"
        + "catch{$err='acrobat_open_failed: '+(($_.Exception.Message) -replace '\\s+',' ')};"
        + "if($err -eq '' -and -not $ok){$err='acrobat_open_failed'}"
        + "};"
        + "if($err -eq ''){try{$pages=$pd.GetNumPages()}catch{$pages=0}};"
        + "if($err -eq ''){"
        + "$saved=$false;"
        + "try{$saved=$pd.Save(" + str(_PDF_SAVE_FLAGS) + ",$o)}"
        + "catch{$err='acrobat_save_failed: '+(($_.Exception.Message) -replace '\\s+',' ')};"
        + "if($err -eq '' -and -not $saved){$err='acrobat_save_failed'}"
        + "};"
        + "if($null -ne $pd){"
        + "try{$pd.Close()|Out-Null}catch{};"
        + "try{[Runtime.InteropServices.Marshal]::ReleaseComObject($pd)|Out-Null}catch{}"
        + "};"
        + "if($err -ne ''){Write-Output ('ERR|'+$err)}"
        + "else{Write-Output ('DOC|'+$p);Write-Output ('OK|'+[string]$pages)}"
    )
    res = _run_ps_task(script, timeout=timeout)
    if res.ok is None:
        err = res.err

        low = err.lower()
        if err.startswith("acrobat_open_failed") and (
                "denied" in low or "拒绝" in err or "0x80070005" in low):
            err = "acrobat_iac_disabled"
        raise OSError(_friendly_error(KIND_PDF, err, res.raw_detail))
    if not os.path.exists(dst):
        raise OSError("Acrobat 报告保存成功，但没有写出文件。")
    note = "    Acrobat 已用无窗口 IAC 打开文档并另存（脚本本身没有读磁盘字节）"
    if res.ok and res.ok.strip().isdigit() and int(res.ok) > 0:
        note += f"：{res.ok.strip()} 页"
    return ReadResult(KIND_INFO[KIND_PDF][2], notes=[note],
                      saved=None, bound=res.bound)

def copy_raw(src: str, dst: str) -> ReadResult:
    shutil.copyfile(src, dst)
    return ReadResult("磁盘字节复制（不是内存读取）", disk_copy=True)

def read_pdf(src: str, dst: str, args) -> ReadResult:
    try:
        return read_pdf_memory(src, dst)
    except OSError as exc:
        if not getattr(args, "pdf_raw_copy", False):
            raise OSError(
                f"{exc}\n"
                "    如果这个 PDF 在磁盘上本来就是明文（文件头是 %PDF），"
                "可以加 --pdf-raw-copy 退化成磁盘字节复制；\n"
                "    但如果它被加密客户端包过（磁盘上是密文），复制出来的备份同样打不开。"
            ) from exc
        log("    [!] Acrobat 内存读取不可用，按 --pdf-raw-copy 退化成磁盘字节复制。")
        log(f"        原因：{exc}")
        return copy_raw(src, dst)

def read_by_kind(kind: str, src: str, dst: str, args) -> ReadResult:
    if kind == KIND_WORD:
        return read_word_memory(src, dst)
    if kind in (KIND_EXCEL, KIND_POWERPOINT):
        return read_office_savecopy(kind, src, dst)
    if kind == KIND_PDF:
        return read_pdf(src, dst, args)
    raise OSError(f"不支持的格式：{kind}")

def app_is_open(kind: str, path: str, timeout: float = 60.0) -> bool:
    if kind not in WAITABLE_KINDS:
        return False
    _name, _app, _engine, progid, collection = KIND_INFO[kind]
    script = (
        _PS_HEAD
        + "$p=" + _ps_quote(os.path.abspath(path)) + ";"
        + "try{$w=[Runtime.InteropServices.Marshal]::GetActiveObject(" + _ps_quote(progid) + ")}"
        + "catch{Write-Output 'NO';exit 0};"
        + "$d=@($w." + collection + ")|Where-Object{$_.FullName -eq $p}|Select-Object -First 1;"
        + "if($null -eq $d){Write-Output 'NO'}else{Write-Output 'YES'}"
    )
    try:
        _rc, out, _err = _run_ps(script, timeout=timeout)
    except Exception:
        return False
    return any(line.strip() == "YES" for line in out.splitlines())

def wait_until_open(kind: str, path: str, timeout: float) -> bool:
    deadline = time.time() + timeout
    while True:
        if app_is_open(kind, path):
            return True
        if time.time() >= deadline:
            return False
        time.sleep(0.5)

def _unlink_quiet(path: str) -> None:
    try:
        os.remove(path)
    except OSError:
        pass

def publish_backup(tmp_path: str, out_path: str) -> None:
    last = None
    for attempt in range(PUBLISH_RETRIES):
        try:
            os.replace(tmp_path, out_path)
            return
        except OSError as exc:
            last = exc
            if attempt == 0:
                log("    备份已写好，正在等应用释放临时文件…")
            time.sleep(PUBLISH_WAIT)
    raise PublishError(str(last)) from last

def read_with_retry(kind: str, src: str, dst: str, retries: int,
                    retry_wait: float, args) -> ReadResult:
    if retries < 1:
        retries = 1
    app = KIND_INFO[kind][1]
    last_exc = None
    for attempt in range(1, retries + 1):
        _unlink_quiet(dst)
        try:
            return read_by_kind(kind, src, dst, args)
        except DocumentNotOpenError as exc:
            last_exc = exc
            if attempt == retries:
                raise
            log(f"    [!] 还没能在 {app} 里绑定到该文档（第 {attempt}/{retries} 次），"
                f"{retry_wait:g}s 后重试…")
            time.sleep(retry_wait)
    raise last_exc

def backup_path_for(path: str) -> str:
    return path + ".mb"

def human_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"

_MAGIC_SIGNATURES = (
    (b"PK\x03\x04", "OOXML / ZIP，.docx .xlsx .pptx 的正常头部"),
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "OLE2 复合文档，.doc .xls .ppt 的正常头部"),
    (b"%PDF", "PDF"),
    (b"{\\rtf", "RTF"),
)

_DLP_MAGIC = b"Dlp\xff"

def describe_head(path: str):
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError:
        return "", False
    if head.startswith(_DLP_MAGIC):
        return '44 6C 70 FF（"Dlp" 私有容器 = 密文）', True
    for sig, desc in _MAGIC_SIGNATURES:
        if head.startswith(sig):
            return f"{sig.hex(' ').upper()}（{desc}）", False
    return head.hex(" ").upper(), False

def _ext_list(exts) -> str:
    return " ".join(sorted(exts))

def _display_width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)

def format_table() -> str:
    exts_by_kind = {
        KIND_WORD: WORD_EXTENSIONS,
        KIND_EXCEL: EXCEL_EXTENSIONS,
        KIND_POWERPOINT: POWERPOINT_EXTENSIONS,
        KIND_PDF: PDF_EXTENSIONS,
    }
    rows = []
    for kind in (KIND_WORD, KIND_EXCEL, KIND_POWERPOINT, KIND_PDF):
        name, app, engine, _progid, _coll = KIND_INFO[kind]
        rows.append((name, _ext_list(exts_by_kind[kind]), engine, app))
    width = max(_display_width(r[0]) for r in rows)
    lines = []
    for name, exts, engine, app in rows:
        pad = " " * (width - _display_width(name))
        lines.append(f"  {name}{pad}  {exts}")
        lines.append(f"  {' ' * width}  → {engine}（{app}）")
    return "\n".join(lines)

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="打开文件（交给对应的 Office / Acrobat），等 N 秒后从应用内存里"
                    "读取文档内容，备份成同目录下的 原名.原后缀.mb（如 1.pdf.mb）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="支持的格式：\n" + format_table()
               + '\n\n示例：python open_and_backup.py "1.xlsx" --delay 16',
    )
    p.add_argument("file", help="要打开并备份的文件（Word / Excel / PowerPoint / PDF）")
    p.add_argument("--delay", type=float, default=16.0,
                   help="打开后等待多少秒再备份（默认 16）")
    p.add_argument("--no-open", action="store_true",
                   help="不调用默认程序打开文件；文件没在应用里时由绑定过程让应用打开")
    p.add_argument("--wait-open", type=float, default=0.0, metavar="SECONDS",
                   help="先等文件真正在应用里打开（最多 N 秒），再开始计时"
                        "（PDF 不适用）")
    p.add_argument("--retries", type=int, default=10,
                   help="文件还没在应用里打开时，最多重试几次（默认 10）")
    p.add_argument("--retry-wait", type=float, default=1.0,
                   help="重试间隔秒数（默认 1）")
    p.add_argument("--pdf-raw-copy", action="store_true",
                   help="PDF 专用：Acrobat 内存读取不可用时，退化成磁盘字节复制"
                        "（磁盘若是密文，备份同样打不开）")
    p.add_argument("--pause", action="store_true", help="结束前按回车（双击运行时用）")
    return p

def _pause_if_needed(args) -> None:
    if args.pause:
        try:
            input("按回车键关闭窗口…")
        except EOFError:
            pass

def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

    path = os.path.abspath(os.path.expanduser(args.file))
    if not os.path.isfile(path):
        log(f"[X] 找不到文件：{path}")
        return 2

    ext = os.path.splitext(path)[1].lower()
    kind = EXT_TO_KIND.get(ext)
    if kind is None:
        log(f"[X] 不支持的格式：{ext or '（无后缀）'}")
        log("    本脚本只从应用内存读取文档内容，目前支持下列格式：")
        log(format_table())
        return 2

    name, app, engine, _progid, _coll = KIND_INFO[kind]

    log("=" * 68)
    log(f"目标文件：{path}")
    log(f"文件类型：{name}（{ext}）")
    log(f"读取方式：{engine}（不读硬盘，不建卷影副本，不需要管理员）")
    if IS_WINDOWS and is_admin():
        log(f"[!] 当前是管理员身份运行：普通权限下打开的 {app} 实例可能绑定不到，")
        log("    建议改用普通权限运行本脚本。")

    if not args.no_open:
        log("[*] 用默认程序打开文件…")
        try:
            os.startfile(path)
        except Exception as exc:
            log(f"[X] 打开失败：{exc}")
            return 3
        if args.wait_open > 0:
            if kind in WAITABLE_KINDS:
                t_wait = time.time()
                if wait_until_open(kind, path, args.wait_open):
                    log(f"[*] 文件已在 {app} 中打开（等了 {time.time() - t_wait:.1f} 秒）")
                else:
                    log(f"[!] 等了 {args.wait_open:g} 秒仍未在 {app} 中打开，按原计划继续。")
            else:
                log("    （PDF 用 Acrobat 无窗口接口读取，不需要等它打开，忽略 --wait-open）")
        log(f"[*] 等待 {args.delay:g} 秒后开始备份…")
        time.sleep(args.delay)
    else:
        log("[*] --no-open：跳过打开步骤，直接从应用内存读取。")
        if kind == KIND_PDF:
            log("    （PDF 不需要事先打开：Acrobat 会用无窗口接口自己打开这份文件）")
        else:
            log(f"    （如果 {app} 里还没打开这份文件，绑定过程会让 {app} 把它打开）")

    out_path = backup_path_for(path)

    tmp_path = out_path + ".part"
    existed = os.path.exists(out_path)
    log(f"[*] 从 {app} 内存读取 → {out_path}")
    if existed:
        log("    （已存在同名备份，将被覆盖）")

    start = time.time()
    try:
        result = read_with_retry(kind, path, tmp_path, args.retries,
                                 args.retry_wait, args)
    except OSError as exc:
        log(f"[X] 备份失败：{exc}")
        log(f"    请确认 {app} 正在运行、并且当前用户能访问这个 {app} 实例；")
        if kind in WAITABLE_KINDS:
            log(f"    以管理员身份运行的脚本绑定不到普通权限下打开的 {app}。")
        _unlink_quiet(tmp_path)
        _pause_if_needed(args)
        return 4

    for line in result.notes:
        log(line)
    if result.bound and os.path.normcase(result.bound) != os.path.normcase(path):
        log(f"[!] 注意：实际绑定到的是另一个文件：{result.bound}")
    if result.saved is not None:
        log("    文档状态：" + ("已保存，没有未保存的修改"
                              if result.saved else "有未保存的修改（内存版本比磁盘新）"))

    try:

        size = os.path.getsize(tmp_path)
        digest = _sha256_file(tmp_path)
    except OSError as exc:
        log(f"[X] 备份失败：读取过程没有产出可用的内容（{exc}）")
        _unlink_quiet(tmp_path)
        _pause_if_needed(args)
        return 4

    try:
        publish_backup(tmp_path, out_path)
    except PublishError as exc:

        log(f"[X] 备份内容读出来了，但最后改名失败：{exc}")
        log(f"    内容完整地保存在：{tmp_path}")
        log(f"    它其实就是一份正常的文档，手动去掉结尾的 .part（改成 {out_path}）即可用。")
        log("    常见原因是应用还占着这个文件；关掉它再改名，或直接重跑一次脚本。")
        _pause_if_needed(args)
        return 4
    cost = time.time() - start

    verify_digest = _sha256_file(out_path)
    head_desc, head_encrypted = describe_head(out_path)

    log("-" * 68)
    log(f"[√] 备份完成：{out_path}")
    log(f"    类型：{name}    读取方式：{result.engine}")
    log(f"    大小：{human_size(size)}（{size} 字节），耗时 {cost:.2f} 秒")
    log(f"    备份 SHA-256：{verify_digest}")
    log(f"    校验：{'一致' if verify_digest == digest else '不一致，请检查！'}")
    if head_desc:
        log(f"    备份头部：{head_desc}")
    if head_encrypted:
        if result.disk_copy:
            log("[!] 备份头部是加密客户端的私有容器（Dlp）：源文件在磁盘上就是密文，")
            log("    复制出来的当然也是密文，这份备份脱离本机环境打不开。")
        else:
            log("[!] 备份头部是加密客户端的私有容器（Dlp），这一份又被包成密文了：")
            log("    脱离本机加密环境打不开，请确认临时文件后缀是 .part 后重跑一次。")
    if result.disk_copy:
        log("    [!] 本次是**磁盘字节复制**，不是内存读取：")
        log("        磁盘上的原文件若被加密客户端包过，这份备份同样是密文。")
    else:
        log(f"    说明：这是从 {app} 内存取出的内容，可以正常打开；")
        log("          字节与磁盘上的原文件不同——磁盘上那份往往是加密容器的密文。")
        log("          内存里未保存的改动也包含在内。")
    log("    提示：备份后缀是 原后缀 + .mb，直接双击打不开；"
        "要打开请去掉结尾的 .mb。")
    log("=" * 68)

    _pause_if_needed(args)
    return 0

def _sha256_file(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(CHUNK_SIZE), b""):
            digest.update(block)
    return digest.hexdigest()

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\n[!] 已取消。")
        sys.exit(130)
