# ═══════════════════════════════════════════════════════════
# НЭХЭМЖЛЭХ (Data Team) — replaces the old build_invoice()
# ═══════════════════════════════════════════════════════════
try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    _HAS_RICH = True
except ImportError:
    _HAS_RICH = False

ACC_MONEY = '"₮"* #,##0.00'          # ₮ left, number right (Эцсийн үнэ column)
TNR = "Times New Roman"

DEFAULT_RECIPIENT_ADDR = ("Монгол улс, Улаанбаатар - 14240,\n"
                          "Сүхбаатар дүүрэг, Чингисийн өргөн чөлөө - 15, “Моннис” цамхаг")
DEFAULT_RECIPIENT_PHONE = "+(976) - (11) - 331880, д/у: 3800 Факс: +(976) - (11) - 331890"
DEFAULT_SELLER_ADDR = ("Новел софт ХХК, 22-р давхар, Хаан банк тауэр, Хан-Уул дүүрэг, "
                       "Улаанбаатар, Монгол улс")


def _parse_date(s):
    s = (s or "").strip()
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y",
                "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def fmt_doc_date(s):
    """'9/29/2026' → '2026.09.29'"""
    d = _parse_date(s)
    return d.strftime("%Y.%m.%d") if d else (s or "")


def fmt_period(s):
    """period_end → '2026-09'"""
    d = _parse_date(s)
    return d.strftime("%Y-%m") if d else (s or "")[:7].replace(".", "-")


def _invoice_desc(name, per, po, dept, leader, cc):
    p1 = f"Tableau expert/{per} сар {name} "
    p2 = f"PO: {po}"
    p3 = f" OT\n{dept} {leader} cost:{cc}"
    if not _HAS_RICH:
        return p1 + p2 + p3
    f = InlineFont(rFont=TNR, sz=9)
    fb = InlineFont(rFont=TNR, sz=9, b=True)
    return CellRichText([TextBlock(f, p1), TextBlock(fb, p2), TextBlock(f, p3)])


def build_invoice(emp_list, pricing, company, doc_number="", doc_date=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Нэхэмжлэх"
    ws.sheet_view.showGridLines = False

    # A Д/д | B:C Барааны нэр | D Тоо | E Нэгж үнэ | F Нийт үнэ | G Эцсийн үнэ
    for col, w in {1: 3.5, 2: 12, 3: 31, 4: 7, 5: 13.5, 6: 13, 7: 19.5}.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top, ws.page_margins.bottom = 0.6, 0.5

    dot = Border(left=_s("dotted"), right=_s("dotted"), top=_s("dotted"), bottom=_s("dotted"))
    blue = Border(bottom=_s("medium", MID_BLUE))
    seller = company.get("seller_name", "Ө.Мөнхжаргал")
    po = emp_list[0].get("po_code", "3106789606") if emp_list else "3106789606"

    # ── Header + logo ──
    if logo_path and os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.width, img.height = 200, 38
        img.anchor = "E1"
        ws.add_image(img)

    ws.row_dimensions[1].height = 16
    mc(ws, 1, 1, 1, 4, "Новел Софт", bold=True, size=11)
    ws.row_dimensions[2].height = 28
    mc(ws, 2, 1, 2, 4, company.get("seller_address", DEFAULT_SELLER_ADDR),
       size=10, wrap=True, v="top")
    ws.row_dimensions[3].height = 16
    mc(ws, 3, 1, 3, 4, "УТАС: (976)-72222828-3; WWW.NOVELSOFT.MN", size=10)

    # ── Title: blue line | НЭХЭМЖЛЭХ | blue line ──
    ws.row_dimensions[4].height = 26
    mc(ws, 4, 1, 4, 4, "", bdr=blue)
    t = mc(ws, 4, 5, 4, 6, "НЭХЭМЖЛЭХ", h="center", v="bottom")
    t.font = Font(name=TNR, size=12, bold=True, italic=True)
    sc(ws, 4, 7, "", bdr=blue)
    ws.row_dimensions[5].height = 12

    # ── Recipient block (dotted) ──
    R = 6
    ws.row_dimensions[R].height = 20
    mc(ws, R, 1, R, 2, "ХЭНД:", size=9, bdr=dot)
    mc(ws, R, 3, R, 5, company.get("invoice_recipient", "Оюу толгой ХХК"),
       bold=True, size=10, bdr=dot)
    sc(ws, R, 6, "ДУГААР:", size=9, bdr=dot)
    sc(ws, R, 7, doc_number, size=10, h="center", bdr=dot)

    R += 1
    ws.row_dimensions[R].height = 34
    mc(ws, R, 1, R, 2, "ХАЯГ:", size=9, bdr=dot)
    mc(ws, R, 3, R, 5, company.get("recipient_address", DEFAULT_RECIPIENT_ADDR),
       size=9, wrap=True, bdr=dot)
    sc(ws, R, 6, "ОГНОО:", size=9, bdr=dot)
    sc(ws, R, 7, fmt_doc_date(doc_date), size=10, h="center", bdr=dot)

    R += 1
    ws.row_dimensions[R].height = 20
    mc(ws, R, 1, R, 2, "Утас:", size=9, bdr=dot)
    mc(ws, R, 3, R, 5, company.get("recipient_phone", DEFAULT_RECIPIENT_PHONE),
       size=9, bdr=dot)
    mc(ws, R, 6, R, 7, f"РО {po}", bold=True, size=10, h="center", bdr=dot)

    # ── Table header ──
    R += 1
    ws.row_dimensions[R].height = 12
    R += 1
    ws.row_dimensions[R].height = 22
    sc(ws, R, 1, "Д/д", size=9, h="center", bdr=_b())
    mc(ws, R, 2, R, 3, "Барааны нэр", size=9, h="center", bdr=_b())
    sc(ws, R, 4, "Тоо", size=9, h="center", bdr=_b())
    sc(ws, R, 5, "Нэгж үнэ", size=9, h="center", bdr=_b())
    sc(ws, R, 6, "Нийт үнэ", size=9, h="center", bdr=_b())
    sc(ws, R, 7, "Эцсийн үнэ", size=9, h="center", bdr=_b())
    ds = R + 1

    # ── Rows ──
    for idx, emp in enumerate(emp_list, 1):
        R += 1
        ws.row_dimensions[R].height = 40
        try:
            hours = float(emp.get("total_hours", 0))
        except Exception:
            hours = 0
        hours = int(hours) if hours.is_integer() else hours

        desc = _invoice_desc(
            name=emp.get("invoice_name") or emp.get("employee_name", ""),
            per=fmt_period(emp.get("period_end")),
            po=emp.get("po_code", po),
            dept=emp.get("department", "Asset management"),
            leader=emp.get("ot_leader_name", "Munkhbayar Mishig"),
            cc=emp.get("cost_code", "49071226"),
        )
        sc(ws, R, 1, idx, size=9, h="center", bdr=_b())
        cell = mc(ws, R, 2, R, 3, "", size=9, h="center", wrap=True, bdr=_b())
        cell.value = desc
        sc(ws, R, 4, hours, size=9, h="center", bdr=_b())
        sc(ws, R, 5, unit_price(emp, pricing), size=9, h="right", bdr=_b(), nf=MONEY)
        sc(ws, R, 6, f"=D{R}*E{R}", size=9, h="right", bdr=_b(), nf=MONEY)
        sc(ws, R, 7, f"=F{R}*1.1", size=9, h="right", bdr=_b(), nf=ACC_MONEY)
    de = R

    # ── Totals ──
    t1, t2, t3 = R + 1, R + 2, R + 3
    totals = [
        (t1, "Нийт төлбөр /НӨАТ ороогүй/", f"=SUM(F{ds}:F{de})", False, ACC_MONEY),
        (t2, "НӨАТ 10%", f"=SUM(G{ds}:G{de})-SUM(F{ds}:F{de})", False, ACC_MONEY),
        (t3, "Нийт төлбөр /НӨАТ орсон/", f"=G{t1}+G{t2}", True, MONEY),
    ]
    for r, label, formula, bold, nf in totals:
        ws.row_dimensions[r].height = 18
        mc(ws, r, 1, r, 6, label, bold=bold, size=9, h="right", bdr=_b())
        sc(ws, r, 7, formula, bold=bold, size=9, h="right", bdr=_b(), nf=nf)
    R = t3

    # ── Signature ──
    R += 2
    ws.row_dimensions[R].height = 34
    mc(ws, R, 2, R, 3, "БАТАЛГААЖУУЛСАН:", size=9, h="right", v="bottom")
    mc(ws, R, 4, R, 6, "", bdr=Border(bottom=_s("thin")))
    R += 1
    sc(ws, R, 4, "/", size=9)
    mc(ws, R, 5, R, 6, seller, size=8, h="center")
    sc(ws, R, 7, "/", size=9)

    # thick black line
    R += 1
    ws.row_dimensions[R].height = 6
    for c in range(1, 8):
        ws.cell(R, c).border = Border(bottom=_s("thick"))

    R += 3
    mc(ws, R, 1, R, 7,
       "Гүйлгээний утга дээр компанийн нэр болон регистрийн дугаарыг заавал бичнэ үү.",
       size=11, h="center")

    # ── Bank info ──
    R += 2
    mc(ws, R, 3, R, 5, "Банкны мэдээлэл", bold=True, size=9, h="center")
    R += 1
    sc(ws, R, 3, company.get("bank_name", "ХХБанк"), size=9, h="center")
    mc(ws, R, 4, R, 5, company.get("bank_account", "10 0004000 435005688 (₮)"), size=9, h="center")
    sc(ws, R, 6, "Борлуулагч:", size=9, h="right")
    sc(ws, R, 7, seller, size=9)

    R += 1
    mc(ws, R, 2, R + 1, 3, "Хүлээн авагчийн нэр:\n     Новел Софт ХХК",
       bold=True, size=9, wrap=True)
    sc(ws, R, 6, "Оффисын утас:", size=9, h="right")
    sc(ws, R, 7, company.get("seller_office_phone", "72222828"), size=9)
    R += 1
    sc(ws, R, 6, "Гар утас:", size=9, h="right")
    sc(ws, R, 7, company.get("seller_mobile", "86635308"), size=9)
    R += 1
    sc(ws, R, 6, "И-мэйл:", size=9, h="right")
    e = sc(ws, R, 7, company.get("seller_email", "munkhjargal.u@novelsoft.mn"), size=9)
    e.font = Font(name=TNR, size=9, color="0563C1", underline="single")

    # dashed bottom line
    R += 2
    for c in range(1, 8):
        ws.cell(R, c).border = Border(bottom=_s("dashed"))

    ws.print_area = f"A1:G{R}"
    return wb_to_pdf_bytes(wb)
