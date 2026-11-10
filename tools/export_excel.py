"""把全部卡片内容导出成 Excel 备份，并自动记录每次更新改了什么。

    python3 tools/export_excel.py [--notes "这次更新了什么"] [--date 2026-11-10]

在本仓库根目录运行（或任意目录，脚本按自身位置找文件）。读取 cards.json / catalog.json /
manifest.json，生成 backup/考证闪卡内容备份.xlsx，包含：

    说明      总览数字和使用方法
    科目汇总  90 门科目：主办方、级别、卡片数、免费张数、内购 ID
    全部卡片  每张卡一行，全部字段（名词、原词、英文、科目、释义、要点、例子、分录、考点、相关名词）
    按科目    每门科目 × 每张卡一行，方便按科目筛选
    更新记录  每次内容变化：新增 / 修改 / 删除了哪张卡、哪门科目，改了哪个字段，旧内容和新内容

更新记录怎么来：backup/snapshot.json 存着上一次导出时的内容。再次运行时和当前内容逐张比较，
把差异追加到 backup/changelog.json（更新记录的唯一来源，不要手动删），然后重新生成 Excel。
内容没变就不加记录。需要 openpyxl（pip3 install openpyxl）。
"""
import argparse, datetime, json, os, sys

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("需要 openpyxl：pip3 install openpyxl")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP = os.path.join(ROOT, "backup")
XLSX = os.path.join(BACKUP, "考证闪卡内容备份.xlsx")
SNAPSHOT = os.path.join(BACKUP, "snapshot.json")
CHANGELOG = os.path.join(BACKUP, "changelog.json")

FIELDS = [("t", "名词"), ("o", "原词"), ("e", "英文"), ("s", "科目"), ("d", "释义"), ("p", "要点"),
          ("x", "例子"), ("en_", "分录"), ("tip", "考点提示"), ("r", "相关名词")]
FIELD_NAME = dict(FIELDS)
SUBJ_FIELDS = [("name", "科目名"), ("jp", "日文全称"), ("level", "级别"), ("short", "简称"), ("productID", "内购ID")]

# 首次导出前的制作经过（根据项目记录补录，日期为当时的日期）
HISTORY = [
    ("2026-11-05", "", "制作", "IT 五门名词闪卡：共通16、ITパスポート37、情報セキュリティマネジメント40、基本情報技術者51、AWS Cloud Practitioner 44、G検定52"),
    ("2026-11-05", "", "制作", "日商簿記闪卡：3級32、2級30、1級26"),
    ("2026-11-07", "", "制作", "新增科目第1批 AP/SC/SAA/AIF/AZ-900/CCNA/LinuC1：+226 张，合计 554 张"),
    ("2026-11-07", "", "制作", "第2批 入门与AI（生成AIパスポート、Generative AI Test、DS検定等 11 门）：+153 张，合计 707 张"),
    ("2026-11-07", "", "制作", "第3批 IPA 高度（NW/DB/ES/SM/SA/ST/PM/AU）：+78 张，合计 785 张"),
    ("2026-11-07", "", "制作", "第4-7批 + 补充：云 Associate/Professional、Linux/DB/编程/统计、日商PC/MOS 等，合计 90 门科目、944 张"),
    ("2026-11-07", "2", "发布", "发布卡片内容 v2：90 门科目、944 张卡（首发内容）"),
]


def load(name):
    with open(os.path.join(ROOT, name), encoding="utf-8") as f:
        return json.load(f)


def text(v):
    """字段值转成单元格文字。"""
    if v is None:
        return ""
    if isinstance(v, list):
        if v and isinstance(v[0], list):  # 分录：[借方, 金额, 贷方, 金额]
            return "\n".join(f"（借）{r[0]} {r[1]} ／（贷）{r[2]} {r[3]}" for r in v)
        return "\n".join(str(x) for x in v)
    return str(v)


def subjects_of(catalog):
    out = []
    for fam in catalog["families"]:
        for s in fam["subjects"]:
            out.append((fam, s))
    return out


def free_ids(cards, catalog):
    """和 App 里 Catalog.freeCardIDs 相同的规则。"""
    ratio, mx, mn = catalog.get("freeRatio"), catalog["freeCardsPerSubject"], catalog.get("freeMin", 0)

    def count(total):
        if ratio is None:
            return mx
        n = -(-total * ratio // 1)  # 向上取整
        return int(min(mx, max(mn, n)))

    per, allfree = {}, set()
    for _, s in subjects_of(catalog):
        idx = [(i, c) for i, c in enumerate(cards) if s["id"] in c["s"]]
        picks = sorted(idx, key=lambda ic: (len(ic[1]["s"]), ic[0]))[:count(len(idx))]
        per[s["id"]] = {c["id"] for _, c in picks}
        allfree |= per[s["id"]]
    return per, allfree


def diff(old, cards, catalog, date, version, notes):
    rows = []
    def row(kind, target, name, field="", before="", after=""):
        rows.append({"date": date, "version": version, "type": kind, "target": target, "name": name,
                     "field": field, "before": before, "after": after, "notes": notes or ""})

    old_subj = old.get("subjects", {})
    new_subj = {s["id"]: {**s, "family": f["name"]} for f, s in subjects_of(catalog)}
    for sid, s in new_subj.items():
        if sid not in old_subj:
            row("新增科目", sid, s["name"], after=f'{s["family"]} · {s.get("jp", "")}')
        else:
            for k, label in SUBJ_FIELDS:
                if text(old_subj[sid].get(k)) != text(s.get(k)):
                    row("修改科目", sid, s["name"], label, text(old_subj[sid].get(k)), text(s.get(k)))
    for sid, s in old_subj.items():
        if sid not in new_subj:
            row("删除科目", sid, s["name"], before=f'{s.get("family", "")} · {s.get("jp", "")}')

    old_cards = old.get("cards", {})
    for c in cards:
        o = old_cards.get(c["id"])
        if o is None:
            row("新增卡片", c["id"], c["t"], after="、".join(c["s"]))
            continue
        for k, label in FIELDS:
            if text(o.get(k)) != text(c.get(k)):
                if k == "s":
                    a, b = set(o.get("s", [])), set(c["s"])
                    parts = ([f'+{"、".join(sorted(b - a))}'] if b - a else []) + ([f'-{"、".join(sorted(a - b))}'] if a - b else [])
                    row("修改卡片", c["id"], c["t"], label, "、".join(o.get("s", [])), "、".join(c["s"]) + f'（{" ".join(parts)}）')
                else:
                    row("修改卡片", c["id"], c["t"], label, text(o.get(k)), text(c.get(k)))
    ids = {c["id"] for c in cards}
    for cid, o in old_cards.items():
        if cid not in ids:
            row("删除卡片", cid, o.get("t", ""), before="、".join(o.get("s", [])))
    return rows


HEAD_FILL = PatternFill("solid", fgColor="1F2937")
HEAD_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")


def table(ws, headers, rows, widths):
    ws.append(headers)
    for cell in ws[1]:
        cell.fill, cell.font = HEAD_FILL, HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = WRAP
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build(cards, catalog, manifest, log, generated):
    subj = subjects_of(catalog)
    sname = {s["id"]: s["name"] for _, s in subj}
    per_free, allfree = free_ids(cards, catalog)
    wb = Workbook()

    ws = wb.active
    ws.title = "说明"
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 90
    info = [
        ("考证闪卡 · 卡片内容备份", ""),
        ("生成时间", generated),
        ("内容版本", f'v{manifest.get("version")}　{manifest.get("notes", "")}'),
        ("主办方", len(catalog["families"])),
        ("科目", len(subj)),
        ("卡片", len(cards)),
        ("免费卡", f"{len(allfree)}（至少在一门科目里免费；每门科目的免费张数见「科目汇总」）"),
        ("", ""),
        ("工作表", ""),
        ("科目汇总", "每门科目的主办方、级别、卡片数、免费张数、内购 ID"),
        ("全部卡片", "每张卡一行，含全部字段。一张卡可同时属于多门科目（见「科目」列）"),
        ("按科目", "每门科目 × 每张卡一行，用筛选按钮选一门科目即可看到它的全部卡片"),
        ("更新记录", "每次内容更新改了什么：新增 / 修改 / 删除的卡片和科目，含旧内容和新内容"),
        ("", ""),
        ("怎么更新", "改好内容并发布后，在 JapanITFlashCard-data 仓库运行："),
        ("", 'python3 tools/export_excel.py --notes "这次更新了什么"'),
        ("", "脚本会和上次导出比较，把差异追加到「更新记录」并重新生成这个文件。"),
        ("数据来源", "cards.json / catalog.json / manifest.json（App 远程更新读的同一份）"),
    ]
    for a, b in info:
        ws.append([a, b])
    ws["A1"].font = Font(bold=True, size=16)
    for r in (2, 3, 4, 5, 6, 7, 9, 15, 18):
        ws.cell(r, 1).font = Font(bold=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")

    ws = wb.create_sheet("科目汇总")
    rows = []
    for fam, s in subj:
        mine = [c for c in cards if s["id"] in c["s"]]
        rows.append([fam["name"], fam.get("tag", ""), s["id"], s["name"], s.get("jp", ""), s.get("level", ""),
                     len(mine), sum(1 for c in mine if len(c["s"]) == 1), len(per_free[s["id"]]),
                     s.get("productID") or f'kaozheng.unlock.{s["id"]}'])
    table(ws, ["主办方", "方向", "科目ID", "科目名", "日文全称", "级别", "卡片数", "独有卡片数", "免费张数", "内购ID"],
          rows, [18, 22, 10, 26, 34, 6, 8, 10, 9, 26])

    ws = wb.create_sheet("全部卡片")
    rows = []
    for i, c in enumerate(cards, 1):
        p = list(c.get("p", [])) + ["", "", ""]
        rows.append([i, c["id"], c["t"], c.get("o", ""), c.get("e", ""),
                     "、".join(sname.get(x, x) for x in c["s"]), " ".join(c["s"]), len(c["s"]),
                     "是" if c["id"] in allfree else "",
                     c.get("d", ""), p[0], p[1], p[2], c.get("x", ""), text(c.get("en_")),
                     c.get("tip", ""), "、".join(c.get("r", []))])
    table(ws, ["序号", "卡片ID", "名词", "原词", "英文", "科目", "科目ID", "科目数", "有科目免费",
               "释义", "要点1", "要点2", "要点3", "例子", "分录", "考点提示", "相关名词"],
          rows, [6, 10, 18, 22, 22, 30, 18, 7, 6, 40, 30, 30, 30, 40, 30, 40, 26])

    ws = wb.create_sheet("按科目")
    rows = []
    for fam, s in subj:
        n = 0
        for c in cards:
            if s["id"] in c["s"]:
                n += 1
                rows.append([fam["name"], s["name"], s["id"], n, c["id"], c["t"], c.get("o", ""), c.get("e", ""),
                             "是" if c["id"] in per_free[s["id"]] else "", len(c["s"]) - 1, c.get("d", "")])
    table(ws, ["主办方", "科目名", "科目ID", "序号", "卡片ID", "名词", "原词", "英文", "本科目免费", "共用其他科目数", "释义"],
          rows, [18, 26, 10, 6, 10, 18, 22, 22, 9, 10, 50])

    ws = wb.create_sheet("更新记录")
    rows = [[r["date"], r["version"] and f'v{r["version"]}', r["type"], r["target"], r["name"],
             r["field"], r["before"], r["after"], r["notes"]] for r in log]
    table(ws, ["日期", "内容版本", "类型", "卡片/科目ID", "名词/科目名", "改动字段", "旧内容", "新内容", "备注"],
          rows, [11, 9, 10, 12, 20, 10, 40, 40, 36])
    wb.save(XLSX)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--notes", help="这次更新的说明，写进更新记录（默认用 manifest.json 的 notes）")
    ap.add_argument("--date", help="记录日期，默认今天")
    a = ap.parse_args()

    cards, catalog, manifest = load("cards.json"), load("catalog.json"), load("manifest.json")
    date = a.date or datetime.date.today().isoformat()
    version = str(manifest.get("version", ""))
    notes = a.notes or manifest.get("notes", "")
    os.makedirs(BACKUP, exist_ok=True)

    log = json.load(open(CHANGELOG, encoding="utf-8")) if os.path.exists(CHANGELOG) else []
    if os.path.exists(SNAPSHOT):
        new = diff(json.load(open(SNAPSHOT, encoding="utf-8")), cards, catalog, date, version, notes)
    else:  # 第一次导出：补录制作经过 + 一条初始备份
        new = [{"date": d, "version": v, "type": t, "target": "", "name": "", "field": "", "before": "",
                "after": "", "notes": n} for d, v, t, n in HISTORY]
        new.append({"date": date, "version": version, "type": "初始备份", "target": "", "name": "",
                    "field": "", "before": "",
                    "after": f'{len(subjects_of(catalog))} 门科目、{len(cards)} 张卡',
                    "notes": "建立 Excel 备份，此后每次更新自动记录差异"})
    log += new
    json.dump(log, open(CHANGELOG, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    snap = {"version": version, "date": date,
            "subjects": {s["id"]: {**s, "family": f["name"]} for f, s in subjects_of(catalog)},
            "cards": {c["id"]: c for c in cards}}
    json.dump(snap, open(SNAPSHOT, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))

    build(cards, catalog, manifest, log, datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    print(f"{len(cards)} 张卡、{len(subjects_of(catalog))} 门科目 → {os.path.relpath(XLSX, ROOT)}")
    print(f"更新记录：本次新增 {len(new)} 条，共 {len(log)} 条" if new else "内容和上次导出相同，没有新增更新记录")


if __name__ == "__main__":
    main()
