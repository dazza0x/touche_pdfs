
import io
import pandas as pd
from typing import Optional
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

def _money(x) -> str:
    try:
        if pd.isna(x): return ""
        return f"£{float(x):,.2f}"
    except Exception:
        return ""

def _dt(x) -> str:
    if pd.isna(x) or x is None:
        return ""
    try:
        return pd.to_datetime(x).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(x)

def build_stylist_statement_pdf(
    brand: str,
    stylist: str,
    period_start: str,
    period_end: str,
    services_df: Optional[pd.DataFrame],
    clients_df: Optional[pd.DataFrame],
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{brand} — {stylist}</b>", styles["Title"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"Statement period: {period_start} to {period_end}", styles["Normal"]))
    story.append(Spacer(1, 12))

    story.append(Paragraph(f"<b>{stylist} Services</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))

    if services_df is None or services_df.empty:
        story.append(Paragraph("No service sales data provided.", styles["Italic"]))
        story.append(Spacer(1, 12))
    else:
        sdf = services_df[["Description","Qty","Per Service","Total"]].copy()
        sdf["Qty"] = sdf["Qty"].astype(int)
        sdf["Per Service"] = sdf["Per Service"].apply(_money)
        sdf["Total"] = sdf["Total"].apply(_money)

        data = [["Description","Qty","Per Service","Total"]] + sdf.values.tolist()
        t = Table(data, colWidths=[92*mm, 16*mm, 28*mm, 28*mm])
        t.setStyle(TableStyle([
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ALIGN",(1,1),(1,-1),"RIGHT"),
            ("ALIGN",(2,1),(3,-1),"RIGHT"),
        ]))
        story.append(t)
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>Services total:</b> Qty {int(services_df['Qty'].fillna(0).sum())}, Value {_money(services_df['Total'].fillna(0).sum())}", styles["Normal"]))
        story.append(Spacer(1, 12))

    story.append(PageBreak())
    story.append(Paragraph(f"<b>{stylist} Client Statement</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))

    if clients_df is None or clients_df.empty:
        story.append(Paragraph("No client statement data provided.", styles["Italic"]))
    else:
        cdf = clients_df[["Date","Client","Cash1","Prepaid"]].copy()
        cdf["Date"] = cdf["Date"].apply(_dt)
        cdf["Cash1"] = cdf["Cash1"].apply(_money)
        cdf["Prepaid"] = cdf["Prepaid"].apply(_money)

        data = [["Date/Time","Client","Cash1","Prepaid"]] + cdf.values.tolist()
        t = Table(data, colWidths=[34*mm, 76*mm, 30*mm, 30*mm])
        t.setStyle(TableStyle([
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("ALIGN",(2,1),(3,-1),"RIGHT"),
        ]))
        story.append(t)
        cash1 = clients_df["Cash1"].fillna(0).sum()
        prepaid = clients_df["Prepaid"].fillna(0).sum()
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>Client totals:</b> Cash1 {_money(cash1)} | Prepaid {_money(prepaid)} | Combined {_money(cash1+prepaid)}", styles["Normal"]))

    doc.build(story)
    return buf.getvalue()
