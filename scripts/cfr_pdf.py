"""
cfr_pdf.py
Generates one PDF page per order from the CFR orders output DataFrame.
Layout matches the example: dark header, info row, color-coded table, totals footer.
"""

import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Brand colors ──────────────────────────────────────────────
DARK_BLUE   = colors.HexColor("#1B3A6B")
MID_BLUE    = colors.HexColor("#D6E4F0")
GREEN       = colors.HexColor("#27AE60")
GREEN_BG    = colors.HexColor("#E9F7EF")
RED         = colors.HexColor("#E74C3C")
RED_BG      = colors.HexColor("#FDEDEC")
GRAY_HEADER = colors.HexColor("#2C3E6B")
LIGHT_GRAY  = colors.HexColor("#F4F6F8")
WHITE       = colors.white
BLACK       = colors.black
TEXT_MUTED  = colors.HexColor("#7F8C8D")

W, H = A4  # 595 x 842 pt

def make_style(name, **kwargs):
    defaults = dict(fontName="Helvetica", fontSize=10, leading=14,
                    textColor=BLACK, alignment=TA_LEFT)
    defaults.update(kwargs)
    return ParagraphStyle(name, **defaults)

def generate_cfr_pdf(out_df) -> io.BytesIO:
    """
    out_df: DataFrame from transform_for_sap()
    Returns: BytesIO containing the PDF
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )

    story = []

    # Styles
    s_order_title = make_style("order_title",
        fontName="Helvetica-Bold", fontSize=18,
        textColor=WHITE, alignment=TA_LEFT)
    s_order_sub = make_style("order_sub",
        fontName="Helvetica", fontSize=10,
        textColor=colors.HexColor("#BDC3C7"), alignment=TA_LEFT)
    s_info_label = make_style("info_label",
        fontName="Helvetica", fontSize=8,
        textColor=TEXT_MUTED, alignment=TA_LEFT)
    s_info_value = make_style("info_value",
        fontName="Helvetica-Bold", fontSize=13,
        textColor=DARK_BLUE, alignment=TA_LEFT)
    s_col_header = make_style("col_header",
        fontName="Helvetica-Bold", fontSize=9,
        textColor=WHITE, alignment=TA_CENTER)
    s_cell = make_style("cell",
        fontName="Helvetica", fontSize=9,
        textColor=BLACK, alignment=TA_CENTER)
    s_cell_ean = make_style("cell_ean",
        fontName="Helvetica", fontSize=9,
        textColor=BLACK, alignment=TA_LEFT)
    s_cell_red = make_style("cell_red",
        fontName="Helvetica", fontSize=9,
        textColor=RED, alignment=TA_CENTER)
    s_cell_ean_red = make_style("cell_ean_red",
        fontName="Helvetica", fontSize=9,
        textColor=RED, alignment=TA_LEFT)
    s_total_label = make_style("total_label",
        fontName="Helvetica-Bold", fontSize=10,
        textColor=BLACK, alignment=TA_LEFT)
    s_total_green = make_style("total_green",
        fontName="Helvetica-Bold", fontSize=10,
        textColor=GREEN, alignment=TA_LEFT)
    s_total_red = make_style("total_red",
        fontName="Helvetica-Bold", fontSize=10,
        textColor=RED, alignment=TA_LEFT)

    usable_w = W - 30*mm  # 165mm usable

    # Group by order
    for order_ref, group in out_df.groupby("OrderRef", sort=False):
        first = group.iloc[0]

        # ── Order header block ─────────────────────────────
        header_data = [
            [Paragraph(f"ORDER {order_ref}", s_order_title)],
            [Paragraph(
                f"{first['Address']} • {first['SoldTo(Postcode)']} {first['ReceiverCity']}".upper(),
                s_order_sub
            )],
        ]
        header_table = Table(header_data, colWidths=[usable_w])
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), DARK_BLUE),
            ("TOPPADDING",    (0,0), (-1,-1), 10),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
            ("RIGHTPADDING",  (0,0), (-1,-1), 10),
            ("ROWBACKGROUNDS", (0,0), (-1,-1), [DARK_BLUE]),
        ]))
        story.append(header_table)

        # ── Info row (date / postcode / city) ──────────────
        info_data = [[
            Paragraph("ORDER DATE",  s_info_label),
            Paragraph("POSTCODE",    s_info_label),
            Paragraph("CITY",        s_info_label),
        ],[
            Paragraph(first["OrderDate"] or "—",           s_info_value),
            Paragraph(first["SoldTo(Postcode)"] or "—",    s_info_value),
            Paragraph((first["ReceiverCity"] or "—").upper(), s_info_value),
        ]]
        info_table = Table(info_data, colWidths=[usable_w/3]*3)
        info_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), MID_BLUE),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
            ("RIGHTPADDING",  (0,0), (-1,-1), 10),
            ("LINEBELOW",     (0,-1), (-1,-1), 0.5, DARK_BLUE),
        ]))
        story.append(info_table)
        story.append(Spacer(1, 4*mm))

        # ── Data table ─────────────────────────────────────
        col_labels = ["EAN", "Qty", "Unit", "Line Type",
                      "Orig. Qty", "Qty Received", "Qty Rejected"]
        col_widths = [
            usable_w * 0.24,  # EAN
            usable_w * 0.09,  # Qty
            usable_w * 0.08,  # Unit
            usable_w * 0.14,  # Line Type
            usable_w * 0.12,  # Orig Qty
            usable_w * 0.17,  # Qty Received
            usable_w * 0.16,  # Qty Rejected
        ]

        rows = [[Paragraph(h, s_col_header) for h in col_labels]]
        row_styles = []

        for i, (_, row) in enumerate(group.iterrows(), start=1):
            is_rejected = str(row.get("LineType","")).strip().upper() == "REJECTED"
            cs = s_cell_red if is_rejected else s_cell
            ce = s_cell_ean_red if is_rejected else s_cell_ean

            # Line type badge text
            lt_text = f'<font color="{"#E74C3C" if is_rejected else "#27AE60"}"><b>{"REJECTED" if is_rejected else "ACCEPTED"}</b></font>'
            lt_para = Paragraph(lt_text, make_style(f"lt{i}", fontSize=9, alignment=TA_CENTER))

            rows.append([
                Paragraph(str(row.get("EAN","") or ""), ce),
                Paragraph(str(row.get("Qty","") or ""), cs),
                Paragraph(str(row.get("Unit","PC")), cs),
                lt_para,
                Paragraph(str(row.get("OriginalQtyOrdered","") or ""), cs),
                Paragraph(str(row.get("QtyReceived","") or ""), cs),
                Paragraph(str(row.get("QtyRejected","") or ""), cs),
            ])
            if is_rejected:
                row_styles.append(("BACKGROUND", (0,i), (-1,i), RED_BG))
            else:
                bg = LIGHT_GRAY if i % 2 == 0 else WHITE
                row_styles.append(("BACKGROUND", (0,i), (-1,i), bg))

        data_table = Table(rows, colWidths=col_widths, repeatRows=1)
        base_style = [
            ("BACKGROUND",    (0,0), (-1,0), GRAY_HEADER),
            ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#D5D8DC")),
            ("LINEBELOW",     (0,0), (-1,0), 1,   WHITE),
            ("TOPPADDING",    (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
            ("LEFTPADDING",   (0,0), (-1,-1), 6),
            ("RIGHTPADDING",  (0,0), (-1,-1), 6),
            ("ALIGN",         (1,1), (-1,-1), "CENTER"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ]
        data_table.setStyle(TableStyle(base_style + row_styles))
        story.append(data_table)

        # ── Totals footer ───────────────────────────────────
        story.append(Spacer(1, 2*mm))
        total_lines    = len(group)
        accepted_qty   = int(group[group["LineType"]=="ACCEPTED"]["Qty"].sum())
        rejected_qty   = int(group[group["LineType"]=="REJECTED"]["Qty"].sum())

        totals_data = [[
            Paragraph(f"Total lines: <b>{total_lines}</b>", s_total_label),
            Paragraph(f'Accepted qty: <font color="#27AE60"><b>{accepted_qty:,}</b></font>', s_total_green),
            Paragraph(f'Rejected qty: <font color="#E74C3C"><b>{rejected_qty:,}</b></font>', s_total_red),
        ]]
        totals_table = Table(totals_data, colWidths=[usable_w/3]*3)
        totals_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), LIGHT_GRAY),
            ("TOPPADDING",    (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
            ("RIGHTPADDING",  (0,0), (-1,-1), 10),
            ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#BDC3C7")),
        ]))
        story.append(totals_table)
        story.append(PageBreak())

    # Remove trailing PageBreak
    if story and isinstance(story[-1], PageBreak):
        story.pop()

    doc.build(story)
    buf.seek(0)
    return buf
