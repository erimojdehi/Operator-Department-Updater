import configparser
import csv
import datetime as dt
import html
import os
import smtplib
import subprocess
import sys
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# -------- Constants --------
FALLBACK_BASE = r"***************local address****************"
ASSETS_DIR = Path(r"***************local address****************")
ASSETS_INPUT = ASSETS_DIR / "UnmatchedDepartment.csv"
ASSETS_ACTIVE = ASSETS_DIR / "Active Operator List.csv"

REQ_OPER = "OPER_oper_no"   # required for XML
REQ_NEW  = "CS_Dept"        # required for XML
OPT_OLD  = "AW_Dept"        # old/current dept for email only

# -------- Small logger --------
class RunContext:
    def __init__(self):
        now_utc = dt.datetime.now(dt.timezone.utc)
        self.when = now_utc.astimezone()
        self.stamp = self.when.strftime("%Y-%m-%d_%H%M%S")
        self.when_str = self.when.strftime("%b %d, %Y %H:%M")
        self.log_path: Optional[Path] = None
        self.lines: List[str] = []
        self.subject_status = "OK"

    def add(self, msg: str):
        self.lines.append(msg)
        try:
            if self.log_path:
                with self.log_path.open("a", encoding="utf-8") as f:
                    f.write(msg + "\n")
        except Exception:
            pass

# -------- Utilities --------
def app_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def get_config_path() -> Path:
    env = os.environ.get("ODU_CONFIG")
    if env and Path(env).exists():
        return Path(env)
    for p in (app_dir() / "config.ini", Path.cwd() / "config.ini", Path(FALLBACK_BASE) / "config.ini"):
        if p.exists():
            return p
    raise FileNotFoundError("config.ini not found (checked app dir, cwd, and base).")

def load_config(cfg_path: Path) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(inline_comment_prefixes=(";", "#"))
    cfg.read(cfg_path, encoding="utf-8")
    for sec in ("PATHS", "SERVER", "UPLOAD", "EMAIL", "RETENTION"):
        if sec not in cfg:
            raise ValueError(f"Missing [{sec}] in config.ini")
    if "OPTIONS" not in cfg:
        cfg["OPTIONS"] = {}
    return cfg

def remove_old_files(folder: Path, pattern: str, retain_days: int) -> int:
    if retain_days <= 0 or not folder.exists():
        return 0
    cutoff = time.time() - retain_days * 86400
    removed = 0
    for p in folder.glob(pattern):
        try:
            if p.stat().st_mtime < cutoff:
                p.unlink()
                removed += 1
        except Exception:
            pass
    return removed

def find_latest_txt_after(folder: Path, epoch: float) -> Optional[Path]:
    if not folder.exists():
        return None
    newest: Optional[Path] = None
    newest_mtime = 0.0
    for p in folder.glob("*.txt"):
        try:
            m = p.stat().st_mtime
            if m >= epoch and m > newest_mtime:
                newest, newest_mtime = p, m
        except Exception:
            pass
    return newest

def append_full_loader_log(log_path: Path, txt_path: Optional[Path]) -> None:
    if not txt_path or not txt_path.exists():
        return
    try:
        with txt_path.open("r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        with log_path.open("a", encoding="utf-8") as out:
            out.write("\n" + "=" * 72 + "\nFA Data Loader Log (raw .txt)\n" + "=" * 72 + "\n")
            out.write(content)
            out.write("\n")
    except Exception:
        pass

def send_email_html(from_addr: str, to_addrs: List[str], subject: str, html_body: str) -> Tuple[bool, Optional[str]]:
    try:
        if not to_addrs:
            return False, "No recipients configured"
        msg = MIMEMultipart()
        msg["From"] = from_addr
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))
        with smtplib.SMTP("**********", 25, timeout=30) as s:
            try:
                s.starttls()
            except Exception:
                pass
            s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

def norm_num(s: str) -> str:
    t = "".join(ch for ch in (s or "") if ch.isdigit())
    if t == "":
        return (s or "").strip()
    return t.lstrip("0") or "0"

# -------- Read UnmatchedDepartment.csv --------
def read_unmatched(csv_path: Path) -> List[Dict[str, str]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if len(rows) < 2:
        return []

    header_low = [h.strip().lower() for h in rows[0]]
    name_map: Dict[str, int] = {}
    for i, h in enumerate(header_low):
        if h in ("oper_oper_no", "oper no", "oper_no", "operator id", "operator", "oper id", "id", "oper"):
            name_map[REQ_OPER] = i
        if h in ("cs_dept", "cs dept", "cd_dept", "dept_code", "dept code", "dept", "department", "new dept", "new department"):
            name_map[REQ_NEW] = i
        if h in ("aw_dept", "aw dept", "aw_dept_code", "aw dept code"):
            name_map[OPT_OLD] = i

    if REQ_OPER not in name_map or REQ_NEW not in name_map:
        name_map[REQ_OPER] = 0
        if len(rows[0]) >= 3:
            name_map[OPT_OLD] = 1
            name_map[REQ_NEW] = 2
        elif len(rows[0]) >= 2:
            name_map[REQ_NEW] = 1

    def cell(r: List[str], i: int) -> str:
        return r[i].strip() if i < len(r) else ""

    out: List[Dict[str, str]] = []
    for r in rows[1:]:
        if not r:
            continue
        oper = cell(r, name_map[REQ_OPER])
        newd = cell(r, name_map[REQ_NEW])
        if not (oper and newd):
            continue
        rec = {REQ_OPER: norm_num(oper), REQ_NEW: norm_num(newd)}
        if OPT_OLD in name_map:
            rec[OPT_OLD] = norm_num(cell(r, name_map[OPT_OLD]))
        out.append(rec)
    return out

# -------- Index Active Operator List by POSITION (0..3) --------
# col0 = Dept number, col1 = Dept name, col2 = Operator name, col3 = Operator number
def index_active_list(csv_path: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
    oper_to_name: Dict[str, str] = {}
    dept_to_name: Dict[str, str] = {}
    if not csv_path.exists():
        return oper_to_name, dept_to_name

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return oper_to_name, dept_to_name

    start_idx = 0
    first = rows[0] if rows else []
    if first:
        c0 = first[0].strip() if len(first) > 0 else ""
        c3 = first[3].strip() if len(first) > 3 else ""
        if not any(ch.isdigit() for ch in c0) or not any(ch.isdigit() for ch in c3):
            start_idx = 1

    for r in rows[start_idx:]:
        if len(r) < 4:
            continue
        dept_code = norm_num(r[0])
        dept_name = r[1].strip()
        oper_name = r[2].strip()
        oper_id   = norm_num(r[3])
        if dept_code:
            dept_to_name.setdefault(dept_code, dept_name)
        if oper_id:
            oper_to_name.setdefault(oper_id, oper_name)
    return oper_to_name, dept_to_name

# -------- Build Excel 2003 XML --------
def build_xml(records: List[Dict[str, str]]) -> str:
    created_utc = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    xml_head = f"""<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Operator Department Updater</Author>
  <LastAuthor>Operator Department Updater</LastAuthor>
  <Created>{created_utc}</Created>
  <Version>16.00</Version>
 </DocumentProperties>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <NumberFormat ss:Format="@"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Work Order Center">
  <Table ss:ExpandedColumnCount="3" x:FullColumns="1" x:FullRows="1">
   <Row>
    <Cell ss:StyleID="s62"><Data ss:Type="Number">2022</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">101:2</Data></Cell>
    <Cell ss:StyleID="s62"><Data ss:Type="String">103:4</Data></Cell>
   </Row>
"""
    rows_xml = []
    for rec in records:
        try:
            oper = int(rec[REQ_OPER])
            newd = int(rec[REQ_NEW])
        except Exception:
            continue
        rows_xml.append(
            "   <Row>\n"
            "    <Cell><Data ss:Type=\"String\">[u:1]</Data></Cell>\n"
            f"    <Cell><Data ss:Type=\"Number\">{oper}</Data></Cell>\n"
            f"    <Cell><Data ss:Type=\"Number\">{newd}</Data></Cell>\n"
            "   </Row>\n"
        )
    xml_tail = """  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
"""
    return xml_head + "".join(rows_xml) + xml_tail

# -------- Main --------
def main():
    ctx = RunContext()
    try:
        cfg = load_config(get_config_path())
        base_dir = Path(cfg.get("PATHS", "base_dir", fallback=FALLBACK_BASE)).resolve()
        input_dir = base_dir / "UnmatchedDepartment_Input"
        dataload_dir = base_dir / "DataLoad_21.1.x"
        emails_dir = base_dir / "emails"
        logs_dir = base_dir / "logs"
        for d in (input_dir, dataload_dir, emails_dir, logs_dir):
            ensure_dir(d)

        ctx.log_path = logs_dir / f"DepartmentUpdate_LOG_{ctx.stamp}.txt"
        with ctx.log_path.open("w", encoding="utf-8") as f:
            f.write(f"Operator Department Updater — Run Report\nDate: {ctx.when_str}\n\n")

        retain_days = int(cfg.get("RETENTION", "days", fallback="30"))
        host = cfg["SERVER"]["host"].strip()
        port = cfg["SERVER"]["port"].strip()
        fa_user = cfg["UPLOAD"]["fadataloader_user"].strip()
        fa_pass = cfg["UPLOAD"]["fadataloader_pass"]
        from_addr = cfg["EMAIL"]["from_address"].strip()
        recips = [x.strip() for x in cfg["EMAIL"]["recipients"].replace(";", ",").split(",") if x.strip()]
        show_loader_window = cfg.get("OPTIONS", "show_loader_window", fallback="0").strip() in ("1", "true", "yes")

        # Step 1: copy input
        dest_csv = None
        if ASSETS_INPUT.exists():
            dest_csv = input_dir / f"UnmatchedDepartment_{ctx.stamp}.csv"
            try:
                dest_csv.write_bytes(ASSETS_INPUT.read_bytes())
                ctx.add("Input CSV copied to local input folder")
            except Exception as e:
                ctx.add(f"Input copy failed: {e}")
                ctx.subject_status = "ISSUE"
        else:
            ctx.add("Input source not found")
            ctx.subject_status = "ISSUE"

        removed_inputs = remove_old_files(input_dir, "UnmatchedDepartment_*.csv", retain_days)
        if removed_inputs:
            ctx.add(f"Removed old input files older than {retain_days} days: {removed_inputs}")

        # Step 2: parse input
        records: List[Dict[str, str]] = []
        if dest_csv and dest_csv.exists():
            try:
                records = read_unmatched(dest_csv)
                ctx.add(f"Parsed {len(records)} record(s) from CSV")
                if not records:
                    ctx.subject_status = "ISSUE"
            except Exception as e:
                ctx.add(f"CSV parse failed: {e}")
                ctx.subject_status = "ISSUE"

        # Step 3: build XML
        xml_generated = False
        xml_name = f"{ctx.when.strftime('%Y-%m-%d')} update operator depts.xml"
        xml_path = dataload_dir / xml_name
        if records:
            try:
                xml_content = build_xml(records)
                xml_path.write_text(xml_content, encoding="utf-8")
                xml_generated = True
                ctx.add(f"XML generated: {xml_name}")
            except Exception as e:
                ctx.add(f"XML generation failed: {e}")
                ctx.subject_status = "ISSUE"
        else:
            ctx.add("No records to write into XML")
            ctx.subject_status = "ISSUE"

        # Step 4: runfile (manual use)
        try:
            runfile = dataload_dir / "runfile.bat"
            bat = (
                "setlocal\r\n"
                "pushd \"%~dp0\"\r\n"
                f"FADATALOADER.EXE -n \"10\" -l \"logs\" -a \"{host}:{port}\" -u \"{fa_user}\" -p \"{fa_pass}\" -i \"{xml_name}\"\r\n"
                "popd\r\n"
                "endlocal\r\n"
            )
            runfile.write_text(bat, encoding="utf-8")
            ctx.add("Data loader runfile prepared")
        except Exception as e:
            ctx.add(f"Runfile write failed: {e}")

        # Step 5: run loader (quiet)
        fa_result = "FAILED (did not run)"
        exe_path = dataload_dir / "FADATALOADER.EXE"
        start_epoch = time.time()
        if xml_generated and exe_path.exists():
            try:
                flags = 0
                if sys.platform.startswith("win") and not show_loader_window:
                    CREATE_NO_WINDOW = 0x08000000
                    flags |= CREATE_NO_WINDOW
                cmd = [
                    str(exe_path), "-n", "10", "-l", "logs",
                    "-a", f"{host}:{port}", "-u", fa_user, "-p", fa_pass, "-i", xml_name
                ]
                result = subprocess.run(cmd, cwd=str(dataload_dir), creationflags=flags, timeout=180)
                if result.returncode == 0:
                    fa_result = "SUCCESS"
                else:
                    fa_result = f"FAILED (exit code {result.returncode})"
            except subprocess.TimeoutExpired:
                fa_result = "FAILED (timeout)"
            except Exception as e:
                fa_result = f"FAILED (exception: {e})"
        elif not xml_generated:
            fa_result = "SKIPPED (no XML)"
        else:
            fa_result = "FAILED (FADATALOADER.EXE not found)"
        ctx.subject_status = "SUCCESS" if fa_result.startswith("SUCCESS") else "ISSUE"
        ctx.add(f"Data loader: {fa_result}")

        # Step 6: names from Active Operator List (position-based)
        oper_to_name, dept_to_name = index_active_list(ASSETS_ACTIVE)

        # Step 7: get FA loader raw .txt content (for email bottom)
        fa_txt_folder = dataload_dir / "logs" / "2022"
        fa_txt = find_latest_txt_after(fa_txt_folder, start_epoch) or (
            max(fa_txt_folder.glob("*.txt"), key=lambda p: p.stat().st_mtime) if fa_txt_folder.exists() else None
        )
        if fa_txt and fa_txt.exists():
            with fa_txt.open("r", encoding="utf-8", errors="ignore") as f:
                fa_log_raw = f.read()
            fa_log_html = html.escape(fa_log_raw)
        else:
            fa_log_html = html.escape(f"(No FA .txt log found in {fa_txt_folder})")

        # Step 8: email — fixed widths, no wrap, “Not Found” for missing names, embed FA log
        ensure_dir(emails_dir)
        def row_line(rec: Dict[str, str]) -> str:
            oper = rec.get(REQ_OPER, "")
            newd = rec.get(REQ_NEW, "")
            oldd = rec.get(OPT_OLD, "")
            op_name = oper_to_name.get(oper) or "Not Found"
            old_name = (dept_to_name.get(oldd) if oldd else None) or "Not Found"
            new_name = (dept_to_name.get(newd) if newd else None) or "Not Found"
            left = f"{oper} ({op_name})"
            right = f"{(oldd or '—')} ({old_name}) &rarr; {(newd or '—')} ({new_name})"
            return f"<tr><td class='c1'>{left}</td><td class='c2'>{right}</td></tr>"

        rows_html = "".join(row_line(r) for r in records[:400])

        email_html = (
            "<!DOCTYPE html><html><head>"
            "<style>"
            "body{font-family:Arial,sans-serif;font-size:14px;margin:24px;background:#111;color:#eee}"
            "table{border-collapse:collapse;width:1100px;table-layout:fixed;margin:0}"
            "th,td{border:1px solid #555;padding:8px 10px;text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}"
            "th{background:#222}"
            ".c1{width:350px}"
            ".c2{width:auto}"
            "h2{margin:0 0 10px 0} h3{margin:20px 0 8px 0}"
            "pre.log{background:#0b0b0b;border:1px solid #444;padding:10px;max-height:500px;overflow:auto;white-space:pre-wrap}"
            "hr{border:0;border-top:1px solid #444;margin:20px 0}"
            "</style></head><body>"
            f"<h2>Operator Department Updater — Run Report</h2>"
            f"<p>Date: {ctx.when_str}</p>"
            f"<table><tr><th class='c1'>Operator</th><th class='c2'>Old Dept → New Dept</th></tr>{rows_html}</table>"
            f"<hr><h3>DataLoader Confirmation</h3><pre class='log'>{fa_log_html}</pre>"
            "</body></html>"
        )

        email_file = emails_dir / f"DepartmentUpdate_EMAIL_{ctx.stamp}.html"
        try:
            email_file.write_text(email_html, encoding="utf-8")
            ctx.add("Email report generated and saved")
        except Exception as e:
            ctx.add(f"Email HTML save failed: {e}")

        sent, err = send_email_html(
            from_addr,
            recips,
            f"[Operator Dept Update] {ctx.when_str} — {ctx.subject_status} — {len(records)} records",
            email_html
        )
        ctx.add("Email sent" if sent else f"Email send failed: {err}")

        # Step 9: retention + append FA raw log to our manager log
        removed_logs = remove_old_files(logs_dir, "*.txt", retain_days)
        if removed_logs:
            ctx.add(f"Removed old logs older than {retain_days} days: {removed_logs}")
        if fa_txt and fa_txt.exists():
            append_full_loader_log(ctx.log_path, fa_txt)

    except Exception as e:
        try:
            with (app_dir() / f"DepartmentUpdate_LOG_{ctx.stamp}.txt").open("a", encoding="utf-8") as f:
                f.write(f"[FATAL] {e}\n")
        except Exception:
            pass

if __name__ == "__main__":
    main()
