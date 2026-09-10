# -*- coding: utf-8 -*-
"""銀行支払明細ブックを「202606〜202705の1年分テンプレート」に組み直す。

入力: 元ブック（みずほ銀行 / 武蔵野銀行 / 勘定科目）
出力: 明細一覧（マスタ・月は横軸）＋ 月別シート12枚 ＋ 勘定科目シート ＋ 原本2枚
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

SRC, DST = sys.argv[1], sys.argv[2]

# 元シートの月列 → 月度コード
MIZUHO_COLS = {"202606": 6, "202607": 5, "202608": 4, "202609": 3}   # F,E,D,C
MUSASHINO_COLS = MIZUHO_COLS

MONTHS = ["202606", "202607", "202608", "202609", "202610", "202611",
          "202612", "202701", "202702", "202703", "202704", "202705"]
MONTH_LABEL = {m: f"{int(m[:4])}年{int(m[4:]):d}月" for m in MONTHS}
SPARE = 40                      # 明細一覧の追記用 空行

ACCOUNTS = ["通信費", "水道光熱費", "賃借料", "リース料", "修繕費", "消耗品費",
            "消耗品費（8%）", "車輛燃料費", "旅費交通費", "接待交際費", "会議費",
            "支払手数料", "租税公課", "保険料", "福利厚生費", "広告宣伝費", "雑費",
            "保険積立金", "工具、器具及び備品", "車両運搬具"]

BANKS = ["みずほ銀行", "武蔵野銀行"]

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
NUM = "#,##0_);[Red]\\(#,##0\\)"


def collect(ws, bank, rows, col_map):
    """親行（勘定科目が空＝小計行）を除いた明細行を取り出す。支払先・媒体・備考は親から引き継ぐ。"""
    out = []
    parent = {"payee": "", "media": "", "note": ""}
    for r in rows:
        payee = (ws.cell(r, 1).value or "")
        account = (ws.cell(r, 2).value or "")
        note = (ws.cell(r, 7).value or "")
        media = (ws.cell(r, 8).value or "")
        payee, account = str(payee).strip(), str(account).strip()
        note, media = str(note).strip(), str(media).strip()
        if payee and not account:      # 小計（親）行
            parent = {"payee": payee, "media": media, "note": note}
            continue
        if not account:                # 空行
            continue
        if payee:
            parent = {"payee": payee, "media": media, "note": note}
        else:
            payee = parent["payee"]
            media = media or parent["media"]
            note = note or parent["note"]
        amounts = {}
        for m, c in col_map.items():
            v = ws.cell(r, c).value
            amounts[m] = v if isinstance(v, (int, float)) else None
        out.append({"bank": bank, "payee": payee, "account": account,
                    "note": note, "media": media, "amounts": amounts,
                    "src": (ws.title, r)})
    return out


def main():
    wb = openpyxl.load_workbook(SRC)
    mz, ms = wb["みずほ銀行"], wb["武蔵野銀行"]

    items = collect(mz, "みずほ銀行", range(4, 27), MIZUHO_COLS)
    # 74〜80行は61〜67行と同一内容の重複のため除外
    items += collect(ms, "武蔵野銀行", list(range(2, 74)), MUSASHINO_COLS)

    for it in items:
        if it["account"] not in ACCOUNTS:
            ACCOUNTS.append(it["account"])
    items.sort(key=lambda x: (ACCOUNTS.index(x["account"]),
                              BANKS.index(x["bank"]) if x["bank"] in BANKS else 99,
                              x["src"][1]))

    # 原本を残す
    for name, new in (("みずほ銀行", "原本_みずほ銀行"), ("武蔵野銀行", "原本_武蔵野銀行")):
        wb[name].title = new
    del wb["勘定科目"]

    # ---------------- 明細一覧（マスタ） ----------------
    det = wb.create_sheet("明細一覧", 0)
    headers = ["勘定科目", "銀行", "支払先", "備考", "媒体"] + \
              [f"{m}\n{MONTH_LABEL[m]}" for m in MONTHS] + ["年間合計"]
    for i, h in enumerate(headers, 1):
        c = det.cell(1, i, h)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    first_m = 6                      # F列 = 202606
    last_m = first_m + len(MONTHS) - 1
    for i, it in enumerate(items):
        r = i + 2
        det.cell(r, 1, it["account"])
        det.cell(r, 2, it["bank"])
        det.cell(r, 3, it["payee"])
        det.cell(r, 4, it["note"])
        det.cell(r, 5, it["media"])
        for j, m in enumerate(MONTHS):
            c = det.cell(r, first_m + j, it["amounts"].get(m))
            c.number_format = NUM
        f, l = get_column_letter(first_m), get_column_letter(last_m)
        tot = det.cell(r, last_m + 1,
                       f'=IF(COUNT({f}{r}:{l}{r})=0,"",SUM({f}{r}:{l}{r}))')
        tot.number_format = NUM
        for col in range(1, last_m + 2):
            det.cell(r, col).border = BORDER
    # 追記用の空行（集計式の範囲にあらかじめ含めておく）
    for i in range(SPARE):
        r = len(items) + 2 + i
        for j in range(len(MONTHS)):
            det.cell(r, first_m + j).number_format = NUM
        tot = det.cell(r, last_m + 1,
                       f"=IF(COUNT({get_column_letter(first_m)}{r}:{get_column_letter(last_m)}{r})=0,\"\","
                       f"SUM({get_column_letter(first_m)}{r}:{get_column_letter(last_m)}{r}))")
        tot.number_format = NUM
        for col in range(1, last_m + 2):
            det.cell(r, col).border = BORDER
    nrows = len(items) + SPARE
    tr = nrows + 2
    det.cell(tr, 1, "合計").font = Font(bold=True)
    for col in range(first_m, last_m + 2):
        L = get_column_letter(col)
        c = det.cell(tr, col, f"=SUM({L}2:{L}{nrows + 1})")
        c.number_format, c.font, c.fill = NUM, Font(bold=True), TOTAL_FILL
    for col in range(1, last_m + 2):
        det.cell(tr, col).border = BORDER
        det.cell(tr, col).fill = TOTAL_FILL
    dv_acc = DataValidation(type="list", formula1=f"=勘定科目!$A$6:$A${5 + len(ACCOUNTS)}",
                            allow_blank=True)
    dv_acc.errorTitle, dv_acc.error = "勘定科目が一致しません", "勘定科目シートの一覧から選んでください"
    det.add_data_validation(dv_acc)
    dv_acc.add(f"A2:A{len(items) + SPARE + 1}")
    dv_bank = DataValidation(type="list", formula1='"' + ",".join(BANKS) + '"',
                             allow_blank=True, errorStyle="warning")
    dv_bank.errorTitle, dv_bank.error = "銀行名の確認", "一覧にない銀行名です。月別シートでは「その他銀行」に集計されます。"
    det.add_data_validation(dv_bank)
    dv_bank.add(f"B2:B{len(items) + SPARE + 1}")

    det.freeze_panes = "F2"
    det.auto_filter.ref = f"A1:{get_column_letter(last_m + 1)}{nrows + 1}"
    for col, w in [("A", 20), ("B", 12), ("C", 26), ("D", 30), ("E", 18)]:
        det.column_dimensions[col].width = w
    for col in range(first_m, last_m + 2):
        det.column_dimensions[get_column_letter(col)].width = 13
    det.row_dimensions[1].height = 30

    DET = "明細一覧"
    acc_rng = f"{DET}!$A$2:$A${nrows + 1}"
    bank_rng = f"{DET}!$B$2:$B${nrows + 1}"

    # ---------------- 月別シート（12枚） ----------------
    for j, m in enumerate(MONTHS):
        amt_col = get_column_letter(first_m + j)
        amt_rng = f"{DET}!${amt_col}$2:${amt_col}${nrows + 1}"
        ws = wb.create_sheet(m)
        t = ws.cell(1, 1, f"{m}（{MONTH_LABEL[m]}）  銀行支払 集計・明細")
        t.font = Font(bold=True, size=14)
        ws.cell(2, 1, "※ 金額の入力は「明細一覧」シートで行ってください。このシートは自動集計です。").font = \
            Font(size=9, color="808080")

        ws.cell(4, 1, "■ 勘定科目別 集計").font = Font(bold=True, size=11)
        head = ["勘定科目"] + BANKS + ["その他銀行", "合計", "構成比"]
        for i, h in enumerate(head, 1):
            c = ws.cell(5, i, h)
            c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
            c.alignment = Alignment(horizontal="center")
        r0 = 6
        for k, acc in enumerate(ACCOUNTS):
            r = r0 + k
            ws.cell(r, 1, acc).border = BORDER
            for bi, bank in enumerate(BANKS):
                c = ws.cell(r, 2 + bi,
                            f'=SUMIFS({amt_rng},{acc_rng},$A{r},{bank_rng},"{bank}")')
                c.number_format, c.border = NUM, BORDER
            c = ws.cell(r, 2 + len(BANKS), f"=E{r}-SUM(B{r}:C{r})")
            c.number_format, c.border = NUM, BORDER
            c = ws.cell(r, 3 + len(BANKS), f"=SUMIFS({amt_rng},{acc_rng},$A{r})")
            c.number_format, c.border, c.font = NUM, BORDER, Font(bold=True)
            rt = r0 + len(ACCOUNTS)
            c = ws.cell(r, 4 + len(BANKS), f'=IF($E${rt}=0,"",E{r}/$E${rt})')
            c.number_format, c.border = "0.0%", BORDER
        rt = r0 + len(ACCOUNTS)
        ws.cell(rt, 1, "合計")
        for col in range(2, 6):
            L = get_column_letter(col)
            c = ws.cell(rt, col, f"=SUM({L}{r0}:{L}{rt - 1})")
            c.number_format = NUM
        for col in range(1, 7):
            c = ws.cell(rt, col)
            c.fill, c.font, c.border = TOTAL_FILL, Font(bold=True), BORDER

        d0 = rt + 2
        ws.cell(d0, 1, "■ 明細（この月に金額のある行）").font = Font(bold=True, size=11)
        ws.cell(d0, 3, "※ オートフィルタの「(空白)」を外すと支払のある行だけ表示できます").font = \
            Font(size=9, color="808080")
        dh = d0 + 1
        for i, h in enumerate(["勘定科目", "銀行", "支払先", "金額", "備考", "媒体"], 1):
            c = ws.cell(dh, i, h)
            c.fill, c.font, c.border = SUB_FILL, Font(bold=True), BORDER
            c.alignment = Alignment(horizontal="center")
        for i in range(nrows):
            sr, r = i + 2, dh + 1 + i
            cond = f'{DET}!{amt_col}{sr}=""'
            for ci, sc in enumerate(["A", "B", "C"], 1):
                c = ws.cell(r, ci, f'=IF({cond},"",{DET}!{sc}{sr})')
                c.border = BORDER
            c = ws.cell(r, 4, f'=IF({cond},"",{DET}!{amt_col}{sr})')
            c.number_format, c.border = NUM, BORDER
            for ci, sc in enumerate(["D", "E"], 5):
                c = ws.cell(r, ci, f'=IF({cond},"",{DET}!{sc}{sr})')
                c.border = BORDER
        dt = dh + nrows + 1
        ws.cell(dt, 1, "合計")
        c = ws.cell(dt, 4, f"=SUM(D{dh + 1}:D{dt - 1})")
        c.number_format = NUM
        for col in range(1, 7):
            cc = ws.cell(dt, col)
            cc.fill, cc.font, cc.border = TOTAL_FILL, Font(bold=True), BORDER
        ws.auto_filter.ref = f"A{dh}:F{dt - 1}"
        for col, w in [("A", 20), ("B", 12), ("C", 26), ("D", 14), ("E", 30), ("F", 18)]:
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A6"

    # ---------------- 勘定科目（年間一覧） ----------------
    sm = wb.create_sheet("勘定科目")
    sm.cell(1, 1, "勘定科目別 月次支払合計（202606〜202705）").font = Font(bold=True, size=14)
    sm.cell(2, 1, "※「明細一覧」シートから自動集計されます。").font = Font(size=9, color="808080")
    row = 4
    for title, bank in [("■ 合計（全銀行）", None)] + [(f"■ {b}", b) for b in BANKS]:
        sm.cell(row, 1, title).font = Font(bold=True, size=11)
        hr = row + 1
        c = sm.cell(hr, 1, "勘定科目")
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        for j, m in enumerate(MONTHS):
            c = sm.cell(hr, 2 + j, m)
            c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
            c.alignment = Alignment(horizontal="center")
        c = sm.cell(hr, 2 + len(MONTHS), "年間合計")
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BORDER
        for k, acc in enumerate(ACCOUNTS):
            r = hr + 1 + k
            sm.cell(r, 1, acc).border = BORDER
            for j in range(len(MONTHS)):
                col = first_m + j
                L = get_column_letter(col)
                rng = f"{DET}!${L}$2:${L}${nrows + 1}"
                f = f"=SUMIFS({rng},{acc_rng},$A{r})" if bank is None else \
                    f'=SUMIFS({rng},{acc_rng},$A{r},{bank_rng},"{bank}")'
                c = sm.cell(r, 2 + j, f)
                c.number_format, c.border = NUM, BORDER
            L1, L2 = get_column_letter(2), get_column_letter(1 + len(MONTHS))
            c = sm.cell(r, 2 + len(MONTHS), f"=SUM({L1}{r}:{L2}{r})")
            c.number_format, c.border, c.font = NUM, BORDER, Font(bold=True)
        tr2 = hr + 1 + len(ACCOUNTS)
        sm.cell(tr2, 1, "合計")
        for j in range(len(MONTHS) + 1):
            L = get_column_letter(2 + j)
            c = sm.cell(tr2, 2 + j, f"=SUM({L}{hr + 1}:{L}{tr2 - 1})")
            c.number_format = NUM
        for col in range(1, 3 + len(MONTHS)):
            c = sm.cell(tr2, col)
            c.fill, c.font, c.border = TOTAL_FILL, Font(bold=True), BORDER
        row = tr2 + 3
    sm.column_dimensions["A"].width = 22
    for j in range(len(MONTHS) + 1):
        sm.column_dimensions[get_column_letter(2 + j)].width = 12
    sm.freeze_panes = "B1"

    order = ["明細一覧", "勘定科目"] + MONTHS + ["原本_みずほ銀行", "原本_武蔵野銀行"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    wb.save(DST)
    print("rows:", nrows, "accounts:", len(ACCOUNTS))
    print("sheets:", wb.sheetnames)


main()
