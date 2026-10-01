#!/usr/bin/env python3
"""Build the FieldComm Forms library: fillable PDFs + Excel logs.

    python3 tools/build_forms.py      # writes forms/*.pdf and forms/*.xlsx

Each form is defined below as plain data (header fields, tables, text boxes).
Add a form by adding a definition and a line to FORMS. Requires reportlab and
openpyxl.
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "forms")

NAVY = HexColor("#0D1B2A")
GOLD = HexColor("#C9A84C")
INK = HexColor("#1B2430")
GRID = HexColor("#9AA5B1")
SHADE = HexColor("#EEF1F4")
REV = "Rev 1.0"
FOOTER = ("FieldComm Forms · free to download, use and adapt · "
          "Template only — not legal advice; verify requirements with your contract, AHJ and attorney.")
M = 36  # page margin (0.5 in)


class Page:
    """One PDF page with a cursor that moves down as blocks are added."""

    def __init__(self, c, size, code, title, subtitle):
        self.c, (self.w, self.h) = c, size
        c.setPageSize(size)
        self.code, self.title = code, title
        self.y = self.h - M
        self._n = 0
        self._header(subtitle)

    def name(self, base):
        self._n += 1
        return f"{self.title}_{base}_{self._n}".replace(" ", "_")

    def _header(self, subtitle):
        c, w = self.c, self.w
        c.setFillColor(NAVY)
        c.rect(M, self.y - 46, w - 2 * M, 46, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.rect(M, self.y - 49, w - 2 * M, 3, stroke=0, fill=1)
        c.setFillColor(white)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(M + 12, self.y - 22, self.title.upper())
        c.setFont("Helvetica", 8.5)
        c.drawString(M + 12, self.y - 36, subtitle)
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(w - M - 12, self.y - 20, "FIELDCOMM")
        c.setFillColor(GOLD)
        c.setFont("Helvetica", 8)
        c.drawRightString(w - M - 12, self.y - 33, f"{self.code} · {REV}")
        self.y -= 62
        # footer
        c.setFillColor(GRID)
        c.setFont("Helvetica", 6.5)
        c.drawString(M, M - 14, FOOTER)
        c.drawRightString(w - M, M - 4, f"{self.code} · {self.title}")

    def label(self, x, y, text):
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Bold", 6.8)
        self.c.drawString(x, y, text.upper())

    def field(self, name, x, y, w, h=15, multiline=False, tooltip=""):
        flags = "multiline" if multiline else ""
        self.c.acroForm.textfield(
            name=name, tooltip=tooltip or name, x=x, y=y, width=w, height=h,
            borderWidth=0.6, borderColor=GRID, fillColor=white, textColor=INK,
            fontName="Helvetica", fontSize=0 if multiline else 9, fieldFlags=flags)

    def row(self, items):
        """items: list of (label, width_fraction)."""
        x, avail = M, self.w - 2 * M
        gap = 8
        widths = [f * (avail - gap * (len(items) - 1)) for _, f in items]
        for (lab, _), wd in zip(items, widths):
            self.label(x, self.y - 8, lab)
            self.field(self.name(lab), x, self.y - 27, wd, 16, tooltip=lab)
            x += wd + gap
        self.y -= 36

    def checks(self, label, options):
        self.label(M, self.y - 8, label)
        x = M
        for opt in options:
            self.c.acroForm.checkbox(name=self.name(f"{label}_{opt}"), tooltip=f"{label}: {opt}",
                                     x=x, y=self.y - 26, size=11, borderColor=GRID,
                                     fillColor=white, buttonStyle="check", borderWidth=0.6)
            self.c.setFont("Helvetica", 8.5)
            self.c.setFillColor(INK)
            self.c.drawString(x + 15, self.y - 23, opt)
            x += 15 + self.c.stringWidth(opt, "Helvetica", 8.5) + 16
        self.y -= 34

    def box(self, label, height):
        self.label(M, self.y - 8, label)
        self.field(self.name(label), M, self.y - 12 - height, self.w - 2 * M, height,
                   multiline=True, tooltip=label)
        self.y -= height + 20

    def table(self, title, cols, rows, row_h=17):
        """cols: list of (header, width_fraction)."""
        c = self.c
        avail = self.w - 2 * M
        widths = [f * avail for _, f in cols]
        if title:
            self.label(M, self.y - 8, title)
            self.y -= 12
        # header band
        c.setFillColor(SHADE)
        c.rect(M, self.y - 16, avail, 16, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 6.6)
        x = M
        for (h, _), wd in zip(cols, widths):
            c.drawString(x + 3, self.y - 11, h.upper())
            x += wd
        self.y -= 16
        for r in range(rows):
            x = M
            for (h, _), wd in zip(cols, widths):
                self.field(f"{self.title}_{title or 'row'}_{h}_{r + 1}".replace(" ", "_"), x, self.y - row_h, wd, row_h,
                           tooltip=f"{h} (row {r + 1})")
                x += wd
            self.y -= row_h
        self.y -= 12

    def signature(self, labels):
        x, avail = M, self.w - 2 * M
        wd = (avail - 12 * (len(labels) - 1)) / len(labels)
        for lab in labels:
            self.c.setStrokeColor(INK)
            self.c.setLineWidth(0.7)
            self.c.line(x, self.y - 22, x + wd, self.y - 22)
            self.label(x, self.y - 32, lab)
            x += wd + 12
        self.y -= 40

    def note(self, text):
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Oblique", 8)
        self.c.drawString(M, self.y - 9, text)
        self.y -= 16


# ---------------------------------------------------------------- forms

def daily_report(c):
    p = Page(c, letter, "FC-FORM-001", "Daily Report", "One page per day. Fills the gap between memory and the record.")
    p.row([("Project", .46), ("Report #", .14), ("Date", .18), ("Superintendent", .22)])
    p.row([("Weather AM", .25), ("Weather PM", .25), ("High / Low °F", .2), ("Precip. / Wind", .3)])
    p.checks("Work stopped for weather?", ["No", "Partial", "Yes — full day"])
    p.table("Crew on site", [("Company", .26), ("Trade", .16), ("Workers", .09), ("Hours", .08), ("Work performed / location", .41)], 8)
    p.table("Equipment on site", [("Equipment", .4), ("Owner / rental", .3), ("Hours used", .15), ("Idle?", .15)], 3)
    p.box("Deliveries received (supplier, material, ticket #, short/damaged?)", 40)
    p.box("Visitors & inspections (name, agency, result)", 34)
    p.box("Delays, issues, RFIs or changes raised today", 40)
    p.box("Safety: toolbox talk topic, incidents, near misses", 30)
    p.signature(["Superintendent signature", "Date"])
    p.c.showPage()


def rfi_request(c):
    p = Page(c, letter, "FC-FORM-002", "Request for Information", "Ask one clear question. Reference the sheet and spec. Propose an answer.")
    p.row([("Project", .5), ("RFI #", .15), ("Date sent", .17), ("Response needed by", .18)])
    p.row([("To (design team)", .5), ("From", .5)])
    p.row([("Subject", .6), ("Drawing / sheet ref.", .2), ("Spec section", .2)])
    p.box("Question", 120)
    p.box("Suggested solution (contractor's proposal)", 90)
    p.checks("Potential impact", ["Cost", "Schedule", "None expected", "Unknown"])
    p.box("Response (design team)", 140)
    p.row([("Responded by", .5), ("Date answered", .25), ("Distributed to", .25)])
    p.signature(["Requested by", "Answered by"])
    c.showPage()


def rfi_log(c):
    p = Page(c, landscape(letter), "FC-FORM-002", "RFI Log", "Every RFI, one line. Open items stay visible until they're closed.")
    p.row([("Project", .5), ("Log kept by", .3), ("Page", .2)])
    p.table(None, [("RFI #", .06), ("Date sent", .08), ("Subject", .25), ("Sheet / spec", .1), ("To", .11),
                   ("Needed by", .08), ("Answered", .08), ("Status", .08), ("Cost/Sched impact", .16)], 22, row_h=17)
    c.showPage()


def punch_list(c):
    p = Page(c, landscape(letter), "FC-FORM-003", "Punch List", "Walk it, write it, assign it, verify it. Nothing closes until someone checks it.")
    p.row([("Project", .4), ("Area / building / unit", .3), ("Walk date", .15), ("Walked by", .15)])
    p.table(None, [("#", .04), ("Location", .13), ("Trade", .1), ("Item / deficiency", .31), ("Assigned to", .12),
                   ("Due", .07), ("Done", .07), ("Verified by", .16)], 22, row_h=17)
    c.showPage()


def change_order_log(c):
    p = Page(c, landscape(letter), "FC-FORM-004", "Change Order Log", "Track every change from first mention to signed approval. Cost and days, both.")
    p.row([("Project", .45), ("Original contract $", .2), ("Original completion date", .2), ("Page", .15)])
    p.table(None, [("CO #", .06), ("Ref (RFI / PCO)", .09), ("Description", .29), ("Requested by", .11), ("Date", .08),
                   ("Cost +/-", .1), ("Days +/-", .07), ("Status", .1), ("Approved", .1)], 18, row_h=17)
    p.row([("Total approved $", .25), ("Total approved days", .25), ("Revised contract $", .25), ("Revised completion", .25)])
    c.showPage()


def change_order_request(c):
    p = Page(c, letter, "FC-FORM-004", "Change Order Request", "Describe the change, why it's needed, and what it costs in money and time.")
    p.row([("Project", .5), ("CO request #", .2), ("Date", .3)])
    p.row([("Requested by", .5), ("Related RFI / bulletin / directive", .5)])
    p.checks("Reason", ["Owner request", "Design change", "Unforeseen condition", "Code / AHJ", "Other"])
    p.box("Description of change", 100)
    p.table("Cost breakdown", [("Item", .44), ("Qty", .1), ("Unit", .1), ("Unit cost", .16), ("Amount", .2)], 7)
    p.row([("Subtotal $", .25), ("Overhead & profit $", .25), ("Total change $", .25), ("Schedule impact (days)", .25)])
    p.signature(["Contractor", "Owner / owner's rep", "Date approved"])
    c.showPage()


def toolbox_talk(c):
    p = Page(c, letter, "FC-FORM-005", "Toolbox Talk Sign-In", "Short talk, real topic, every worker signs. The sign-in is your proof it happened.")
    p.row([("Project", .45), ("Date", .18), ("Time", .12), ("Presented by", .25)])
    p.row([("Topic", .7), ("Language(s) used", .3)])
    p.box("Key points covered", 70)
    p.table("Attendance", [("#", .05), ("Print name", .32), ("Company", .25), ("Trade", .14), ("Signature", .24)], 18, row_h=18)
    p.signature(["Presenter signature", "Superintendent"])
    c.showPage()


FORMS = [
    ("FC-FORM-001_Daily-Report", [daily_report]),
    ("FC-FORM-002_RFI-Request-and-Log", [rfi_request, rfi_log]),
    ("FC-FORM-003_Punch-List", [punch_list]),
    ("FC-FORM-004_Change-Order-Request-and-Log", [change_order_request, change_order_log]),
    ("FC-FORM-005_Toolbox-Talk-Sign-In", [toolbox_talk]),
]

# Excel versions of the logs (same columns as the PDFs)
SHEETS = [
    ("FC-FORM-002_RFI-Log", "RFI Log",
     ["RFI #", "Date sent", "Subject", "Sheet / spec", "To", "Needed by", "Answered", "Status", "Cost/Sched impact"],
     [8, 12, 40, 14, 18, 12, 12, 12, 24], 60),
    ("FC-FORM-003_Punch-List", "Punch List",
     ["#", "Location", "Trade", "Item / deficiency", "Assigned to", "Due", "Done", "Verified by"],
     [6, 18, 14, 48, 18, 11, 11, 18], 100),
    ("FC-FORM-004_Change-Order-Log", "Change Order Log",
     ["CO #", "Ref (RFI / PCO)", "Description", "Requested by", "Date", "Cost +/-", "Days +/-", "Status", "Approved"],
     [8, 14, 44, 16, 12, 14, 10, 12, 12], 60),
]


def build_pdf(fname, pages):
    path = os.path.join(OUT, fname + ".pdf")
    c = canvas.Canvas(path, pagesize=letter)
    c.setTitle(fname.replace("_", " ").replace("-", " "))
    c.setAuthor("FieldComm")
    c.setSubject("Free construction field form")
    for page in pages:
        page(c)
    c.save()
    return path


def build_xlsx(fname, title, headers, widths, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    navy = PatternFill("solid", fgColor="0D1B2A")
    thin = Side(style="thin", color="9AA5B1")
    ws["A1"] = f"FIELDCOMM · {title.upper()}"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A2"] = "Project:"
    ws["A2"].font = Font(bold=True)
    for col in range(1, len(headers) + 1):
        ws.cell(row=1, column=col).fill = navy
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=4, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E3A5F")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(5, 5 + rows):
        for col in range(1, len(headers) + 1):
            ws.cell(row=r, column=col).border = Border(top=thin, bottom=thin, left=thin, right=thin)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{4 + rows}"
    ws.cell(row=6 + rows, column=1, value=FOOTER).font = Font(italic=True, size=8, color="667085")
    path = os.path.join(OUT, fname + ".xlsx")
    wb.save(path)
    return path


def main():
    os.makedirs(OUT, exist_ok=True)
    for fname, pages in FORMS:
        print(build_pdf(fname, pages))
    for args in SHEETS:
        print(build_xlsx(*args))


if __name__ == "__main__":
    main()
