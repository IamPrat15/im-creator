"""
IM Creator Python Backend - PPTX Generator
Version: 7.2.0
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from typing import Dict, List
from models import DESIGN, INDUSTRY_DATA, DOCUMENT_CONFIGS, get_theme_colors
from utils import truncate_text, truncate_description, format_currency, format_date, parse_lines, parse_pipe_separated, calculate_cagr, safe_float, safe_int, adjusted_font, extract_percentage
from ai_layout_engine import analyze_data_for_layout

def hex_to_rgb(hex_color: str) -> RgbColor:
    hex_color = hex_color.lstrip('#')
    return RgbColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def add_slide_header(slide, colors, title, subtitle=None, font_adj=0):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(colors["white"])
    bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.1), Inches(0.85))
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb(colors["secondary"])
    bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.3), Inches(0.15), Inches(9.4), Inches(0.5))
    tb.text_frame.paragraphs[0].text = truncate_text(title, 80)
    tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["title"], font_adj))
    tb.text_frame.paragraphs[0].font.bold = True
    tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
    if subtitle:
        sb = slide.shapes.add_textbox(Inches(0.3), Inches(0.62), Inches(9.4), Inches(0.22))
        sb.text_frame.paragraphs[0].text = subtitle
        sb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["subtitle"], font_adj))
        sb.text_frame.paragraphs[0].font.italic = True
        sb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(0.88), Inches(9.4), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(colors["accent"])
    line.line.fill.background()

def add_slide_footer(slide, colors, page_number):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.2), Inches(10), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    line.line.fill.background()
    cb = slide.shapes.add_textbox(Inches(0.3), Inches(5.28), Inches(3), Inches(0.22))
    cb.text_frame.paragraphs[0].text = "Strictly Private & Confidential"
    cb.text_frame.paragraphs[0].font.size = Pt(9)
    cb.text_frame.paragraphs[0].font.italic = True
    cb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
    nb = slide.shapes.add_textbox(Inches(9.2), Inches(5.28), Inches(0.5), Inches(0.22))
    nb.text_frame.paragraphs[0].text = str(page_number)
    nb.text_frame.paragraphs[0].font.size = Pt(10)
    nb.text_frame.paragraphs[0].font.bold = True
    nb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
    nb.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

def add_section_box(slide, colors, x, y, w, h, title=None, title_bg=None, font_adj=0):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
    box.line.color.rgb = hex_to_rgb(colors["border"])
    if title:
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.36))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = hex_to_rgb(title_bg or colors["primary"])
        hdr.line.fill.background()
        tb = slide.shapes.add_textbox(Inches(x + 0.12), Inches(y + 0.02), Inches(w - 0.24), Inches(0.32))
        tb.text_frame.paragraphs[0].text = truncate_text(title, 45)
        tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["section_header"], font_adj))
        tb.text_frame.paragraphs[0].font.bold = True
        tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])

def add_metric_card(slide, colors, x, y, w, h, value, label, font_adj=0):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
    card.line.color.rgb = hex_to_rgb(colors["border"])
    vb = slide.shapes.add_textbox(Inches(x + 0.08), Inches(y + 0.08), Inches(w - 0.16), Inches(h * 0.55))
    vb.text_frame.paragraphs[0].text = str(value)
    vb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["metric_medium"], font_adj))
    vb.text_frame.paragraphs[0].font.bold = True
    vb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
    lb = slide.shapes.add_textbox(Inches(x + 0.08), Inches(y + h * 0.58), Inches(w - 0.16), Inches(h * 0.38))
    lb.text_frame.paragraphs[0].text = label
    lb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["metric_label"], font_adj))
    lb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])

def add_bar_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), chart_data).chart
    chart.has_legend = False
    series = chart.plots[0].series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    if len(data) >= 2:
        cagr = calculate_cagr(safe_float(data[0].get("value")), safe_float(data[-1].get("value")), len(data) - 1)
        if cagr and cagr > 0:
            cb = slide.shapes.add_textbox(Inches(x + w - 0.9), Inches(y - 0.22), Inches(0.85), Inches(0.2))
            cb.text_frame.paragraphs[0].text = f"CAGR: {cagr}%"
            cb.text_frame.paragraphs[0].font.size = Pt(10)
            cb.text_frame.paragraphs[0].font.bold = True
            cb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["accent"])
            cb.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

def add_pie_chart(slide, colors, x, y, size, data, chart_type="pie", font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "")[:15] for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    ct = XL_CHART_TYPE.DOUGHNUT if chart_type == "donut" else XL_CHART_TYPE.PIE
    chart = slide.shapes.add_chart(ct, Inches(x), Inches(y), Inches(size), Inches(size), chart_data).chart
    chart.has_legend = True

def add_progress_bars(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    bar_h, spacing = 0.18, (h - len(data) * 0.18) / (len(data) + 1)
    for idx, item in enumerate(data[:5]):
        by = y + spacing + idx * (bar_h + spacing)
        pct = min(100, max(0, safe_float(item.get("value", 0))))
        fill_w = (pct / 100) * w
        lb = slide.shapes.add_textbox(Inches(x), Inches(by - 0.22), Inches(w * 0.7), Inches(0.2))
        lb.text_frame.paragraphs[0].text = item.get("label", "")
        lb.text_frame.paragraphs[0].font.size = Pt(11)
        lb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        pb = slide.shapes.add_textbox(Inches(x + w - 0.5), Inches(by - 0.22), Inches(0.5), Inches(0.2))
        pb.text_frame.paragraphs[0].text = f"{int(pct)}%"
        pb.text_frame.paragraphs[0].font.size = Pt(11)
        pb.text_frame.paragraphs[0].font.bold = True
        pb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        pb.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
        bg_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(by), Inches(w), Inches(bar_h))
        bg_bar.fill.solid()
        bg_bar.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
        bg_bar.line.color.rgb = hex_to_rgb(colors["border"])
        if fill_w > 0.05:
            fill_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(by), Inches(fill_w), Inches(bar_h))
            fill_bar.fill.solid()
            fill_bar.fill.fore_color.rgb = hex_to_rgb(colors["chart_colors"][idx % 8])
            fill_bar.line.fill.background()

def add_stacked_bar(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    current_x = x
    for idx, item in enumerate(data):
        pct = extract_percentage(item.get("value")) or 20
        bar_w = (pct / 100) * w
        if bar_w > 0.1:
            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(current_x), Inches(y), Inches(bar_w - 0.02), Inches(h))
            bar.fill.solid()
            bar.fill.fore_color.rgb = hex_to_rgb(colors["chart_colors"][idx % 8])
            bar.line.fill.background()
            if bar_w > 0.8:
                lb = slide.shapes.add_textbox(Inches(current_x + 0.05), Inches(y), Inches(bar_w - 0.1), Inches(h))
                lb.text_frame.paragraphs[0].text = f"{truncate_text(item.get('label', ''), 12)} ({pct}%)"
                lb.text_frame.paragraphs[0].font.size = Pt(11)
                lb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
            current_x += bar_w

# ============================================================================
# SLIDE RENDERERS
# ============================================================================

def render_title_slide(slide, colors, data, doc_config):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(colors["dark_bg"])
    bg.line.fill.background()
    al = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(3.05), Inches(3.5), Inches(0.04))
    al.fill.solid()
    al.fill.fore_color.rgb = hex_to_rgb(colors["secondary"])
    al.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(8), Inches(1.1))
    tb.text_frame.paragraphs[0].text = data.get("project_codename") or data.get("projectCodename") or "Project Phoenix"
    tb.text_frame.paragraphs[0].font.size = Pt(48)
    tb.text_frame.paragraphs[0].font.bold = True
    tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    dt = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(6), Inches(0.45))
    dt.text_frame.paragraphs[0].text = doc_config.get("name", "Management Presentation")
    dt.text_frame.paragraphs[0].font.size = Pt(20)
    dt.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    db = slide.shapes.add_textbox(Inches(0.5), Inches(3.8), Inches(4), Inches(0.35))
    db.text_frame.paragraphs[0].text = format_date(data.get("presentation_date") or data.get("presentationDate"))
    db.text_frame.paragraphs[0].font.size = Pt(14)
    db.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    if data.get("advisor"):
        ab = slide.shapes.add_textbox(Inches(0.5), Inches(4.25), Inches(4), Inches(0.3))
        ab.text_frame.paragraphs[0].text = f"Prepared by {data.get('advisor')}"
        ab.text_frame.paragraphs[0].font.size = Pt(12)
        ab.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    cfb = slide.shapes.add_textbox(Inches(0.5), Inches(4.9), Inches(4), Inches(0.25))
    cfb.text_frame.paragraphs[0].text = "Strictly Private and Confidential"
    cfb.text_frame.paragraphs[0].font.size = Pt(10)
    cfb.text_frame.paragraphs[0].font.italic = True
    cfb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])

def render_disclaimer_slide(slide, colors, data, slide_num, font_adj=0):
    add_slide_header(slide, colors, "Important Notice", None, font_adj)
    advisor = data.get("advisor") or "the Advisor"
    company = data.get("company_name") or data.get("companyName") or "the Company"
    disclaimer = f"This document has been prepared by {advisor} exclusively for the benefit of the party to whom it is directly addressed. This document is strictly confidential and may not be reproduced or redistributed without prior written consent.\n\nThis document does not constitute any offer or inducement to purchase securities, nor shall it form the basis of any contract.\n\nThe information herein has been prepared based on information provided by {company} and from sources believed reliable. No representation or warranty is made as to accuracy or completeness.\n\nNeither {advisor} nor any affiliates shall have liability for any loss or damage arising from the use of this document."
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(1.0), Inches(9.2), Inches(3.95))
    tb.text_frame.word_wrap = True
    tb.text_frame.paragraphs[0].text = disclaimer
    tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
    tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_slide_footer(slide, colors, slide_num)

def render_executive_summary(slide, colors, data, slide_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    company = data.get("company_name") or data.get("companyName") or ""
    add_slide_header(slide, colors, "Executive Summary", company, font_adj)
    ct = 0.95
    add_section_box(slide, colors, 0.3, ct, 4.5, 4.1, "Key Highlights", colors["primary"], font_adj)
    metrics = []
    if data.get("founded_year") or data.get("foundedYear"):
        metrics.append((data.get("founded_year") or data.get("foundedYear"), "Founded"))
    if data.get("headquarters"):
        metrics.append((truncate_text(data.get("headquarters"), 18), "Headquarters"))
    if data.get("employee_count_ft") or data.get("employeeCountFT"):
        metrics.append((f"{data.get('employee_count_ft') or data.get('employeeCountFT')}+", "Employees"))
    if data.get("revenue_fy25") or data.get("revenueFY25"):
        metrics.append((format_currency(data.get("revenue_fy25") or data.get("revenueFY25"), data.get("currency", "INR")), "Revenue FY25"))
    if data.get("ebitda_margin_fy25") or data.get("ebitdaMarginFY25"):
        metrics.append((f"{data.get('ebitda_margin_fy25') or data.get('ebitdaMarginFY25')}%", "EBITDA Margin"))
    if data.get("net_retention") or data.get("netRetention"):
        metrics.append((f"{data.get('net_retention') or data.get('netRetention')}%", "Net Retention"))
    for idx, (value, label) in enumerate(metrics[:6]):
        add_metric_card(slide, colors, 0.42 + (idx % 2) * 2.15, ct + 0.48 + (idx // 2) * 1.15, 2.0, 1.0, value, label, font_adj)
    add_section_box(slide, colors, 5.0, ct, 4.7, 1.5, "About the Company", colors["secondary"], font_adj)
    desc = data.get("company_description") or data.get("companyDescription") or "A leading technology solutions provider."
    db = slide.shapes.add_textbox(Inches(5.12), Inches(ct + 0.45), Inches(4.46), Inches(0.95))
    db.text_frame.word_wrap = True
    db.text_frame.paragraphs[0].text = truncate_description(desc, 180)
    db.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
    db.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_section_box(slide, colors, 5.0, ct + 1.65, 4.7, 2.45, "Revenue Growth", colors["accent"], font_adj)
    clb = slide.shapes.add_textbox(Inches(5.12), Inches(ct + 2.0), Inches(1.2), Inches(0.2))
    clb.text_frame.paragraphs[0].text = f"In {'USD Mn' if data.get('currency') == 'USD' else 'INR Cr'}"
    clb.text_frame.paragraphs[0].font.size = Pt(10)
    clb.text_frame.paragraphs[0].font.italic = True
    clb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
    rev_data = []
    for key, label in [("revenueFY24", "FY24"), ("revenueFY25", "FY25"), ("revenueFY26P", "FY26P"), ("revenueFY27P", "FY27P")]:
        val = data.get(key) or data.get(key.lower().replace("fy", "_fy"))
        if val: rev_data.append({"label": label, "value": val})
    if rev_data:
        add_bar_chart(slide, colors, 5.15, ct + 2.2, 4.4, 1.75, rev_data, font_adj)
    add_slide_footer(slide, colors, slide_num)

def render_services_slide(slide, colors, data, slide_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "donut")
    add_slide_header(slide, colors, "Services & Capabilities", "Core offerings", font_adj)
    ct = 0.95
    services = parse_pipe_separated(data.get("service_lines") or data.get("serviceLines") or "", 6)
    add_section_box(slide, colors, 0.3, ct, 5.4, 2.7, "Service Lines", colors["primary"], font_adj)
    for idx, srv in enumerate(services[:4]):
        col, row = idx % 2, idx // 2
        sx, sy = 0.42 + col * 2.65, ct + 0.48 + row * 1.05
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(sy), Inches(2.52), Inches(0.92))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb(colors["white"])
        card.line.color.rgb = hex_to_rgb(colors["border"])
        nb = slide.shapes.add_textbox(Inches(sx + 0.1), Inches(sy + 0.08), Inches(1.85), Inches(0.35))
        nb.text_frame.paragraphs[0].text = truncate_text(srv[0] if srv else "Service", 28)
        nb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        nb.text_frame.paragraphs[0].font.bold = True
        nb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        if len(srv) > 1 and srv[1]:
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx + 1.98), Inches(sy + 0.1), Inches(0.45), Inches(0.28))
            badge.fill.solid()
            badge.fill.fore_color.rgb = hex_to_rgb(colors["accent"])
            badge.line.fill.background()
            bb = slide.shapes.add_textbox(Inches(sx + 1.98), Inches(sy + 0.1), Inches(0.45), Inches(0.28))
            bb.text_frame.paragraphs[0].text = srv[1]
            bb.text_frame.paragraphs[0].font.size = Pt(10)
            bb.text_frame.paragraphs[0].font.bold = True
            bb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
            bb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        if len(srv) > 2 and srv[2]:
            db = slide.shapes.add_textbox(Inches(sx + 0.1), Inches(sy + 0.48), Inches(2.32), Inches(0.38))
            db.text_frame.paragraphs[0].text = truncate_text(srv[2], 45)
            db.text_frame.paragraphs[0].font.size = Pt(11)
            db.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
    add_section_box(slide, colors, 5.9, ct, 3.8, 2.7, "Revenue Mix", colors["secondary"], font_adj)
    pie_data = [{"label": truncate_text(s[0] if s else "Svc", 12), "value": extract_percentage(s[1] if len(s) > 1 else "25") or 25} for s in services[:4]]
    if pie_data:
        add_pie_chart(slide, colors, 6.1, ct + 0.5, 2.0, pie_data, chart_type, font_adj)
    add_slide_footer(slide, colors, slide_num)

def render_clients_slide(slide, colors, data, slide_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    doc_config = context.get("doc_config", {})
    industry_data = context.get("industry_data", {})
    add_slide_header(slide, colors, "Client Portfolio & Vertical Mix", "Diversified customer base", font_adj)
    ct = 0.95
    add_section_box(slide, colors, 0.3, ct, 3.4, 1.5, "Client Metrics", colors["primary"], font_adj)
    metrics = [("Top 10 Concentration", f"{data.get('top_ten_concentration') or data.get('topTenConcentration') or '58'}%"), ("Net Revenue Retention", f"{data.get('net_retention') or data.get('netRetention') or '120'}%"), ("Primary Vertical", industry_data.get("name", "BFSI"))]
    for idx, (label, value) in enumerate(metrics):
        lb = slide.shapes.add_textbox(Inches(0.42), Inches(ct + 0.5 + idx * 0.34), Inches(1.9), Inches(0.32))
        lb.text_frame.paragraphs[0].text = label
        lb.text_frame.paragraphs[0].font.size = Pt(11)
        lb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
        vb = slide.shapes.add_textbox(Inches(2.35), Inches(ct + 0.5 + idx * 0.34), Inches(1.2), Inches(0.32))
        vb.text_frame.paragraphs[0].text = value
        vb.text_frame.paragraphs[0].font.size = Pt(12)
        vb.text_frame.paragraphs[0].font.bold = True
        vb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        vb.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    add_section_box(slide, colors, 0.3, ct + 1.65, 3.4, 2.4, "Vertical Mix", colors["secondary"], font_adj)
    verticals = [{"label": industry_data.get("name", "BFSI"), "value": safe_int(data.get("primary_vertical_pct") or data.get("primaryVerticalPct")) or 60}]
    other_verts = parse_pipe_separated(data.get("other_verticals") or data.get("otherVerticals") or "", 4)
    for v in other_verts:
        verticals.append({"label": v[0] if v else "Other", "value": extract_percentage(v[1] if len(v) > 1 else "10") or 10})
    if not other_verts:
        verticals.extend([{"label": "FinTech", "value": 15}, {"label": "Healthcare", "value": 15}, {"label": "Retail", "value": 10}])
    add_pie_chart(slide, colors, 0.5, ct + 2.05, 1.6, verticals[:5], "donut", font_adj)
    add_section_box(slide, colors, 3.9, ct, 5.8, 4.05, "Key Clients", colors["accent"], font_adj)
    clients = parse_pipe_separated(data.get("top_clients") or data.get("topClients") or "", 12)
    for idx, client in enumerate(clients[:12]):
        col, row = idx % 3, idx // 3
        cx, cy = 4.02 + col * 1.9, ct + 0.48 + row * 0.88
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(1.8), Inches(0.78))
        card.fill.solid()
        card.fill.fore_color.rgb = hex_to_rgb(colors["white"])
        card.line.color.rgb = hex_to_rgb(colors["border"])
        client_name = client[0] if client else f"Client {idx + 1}"
        if not doc_config.get("include_client_names", True):
            client_name = f"Client {idx + 1}"
        nb = slide.shapes.add_textbox(Inches(cx + 0.08), Inches(cy + 0.12), Inches(1.64), Inches(0.35))
        nb.text_frame.paragraphs[0].text = truncate_text(client_name, 18)
        nb.text_frame.paragraphs[0].font.size = Pt(12)
        nb.text_frame.paragraphs[0].font.bold = True
        nb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        nb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        if len(client) > 1 and client[1]:
            sb = slide.shapes.add_textbox(Inches(cx + 0.08), Inches(cy + 0.48), Inches(1.64), Inches(0.25))
            sb.text_frame.paragraphs[0].text = f"Since {client[1]}"
            sb.text_frame.paragraphs[0].font.size = Pt(10)
            sb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
            sb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_slide_footer(slide, colors, slide_num)

def render_financials_slide(slide, colors, data, slide_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Financial Performance", "Revenue growth and key metrics", font_adj)
    ct = 0.95
    add_section_box(slide, colors, 0.3, ct, 4.8, 2.6, "Revenue Growth", colors["primary"], font_adj)
    currency = data.get("currency", "INR")
    clb = slide.shapes.add_textbox(Inches(0.42), Inches(ct + 0.48), Inches(1.2), Inches(0.2))
    clb.text_frame.paragraphs[0].text = f"In {'USD Mn' if currency == 'USD' else 'INR Cr'}"
    clb.text_frame.paragraphs[0].font.size = Pt(10)
    clb.text_frame.paragraphs[0].font.italic = True
    clb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
    rev_data = []
    for key, label in [("revenueFY24", "FY24"), ("revenueFY25", "FY25"), ("revenueFY26P", "FY26P"), ("revenueFY27P", "FY27P"), ("revenueFY28P", "FY28P")]:
        val = data.get(key) or data.get(key.lower().replace("fy", "_fy").replace("p", "p"))
        if val: rev_data.append({"label": label, "value": val})
    if rev_data:
        add_bar_chart(slide, colors, 0.42, ct + 0.72, 4.5, 1.75, rev_data, font_adj)
    add_section_box(slide, colors, 5.3, ct, 4.4, 2.6, "Key Margins & Metrics", colors["secondary"], font_adj)
    margins = []
    if data.get("ebitda_margin_fy25") or data.get("ebitdaMarginFY25"):
        margins.append({"label": "EBITDA Margin FY25", "value": data.get("ebitda_margin_fy25") or data.get("ebitdaMarginFY25")})
    if data.get("gross_margin") or data.get("grossMargin"):
        margins.append({"label": "Gross Margin", "value": data.get("gross_margin") or data.get("grossMargin")})
    if data.get("net_profit_margin") or data.get("netProfitMargin"):
        margins.append({"label": "Net Profit Margin", "value": data.get("net_profit_margin") or data.get("netProfitMargin")})
    if data.get("net_retention") or data.get("netRetention"):
        margins.append({"label": "Net Revenue Retention", "value": data.get("net_retention") or data.get("netRetention")})
    if margins:
        add_progress_bars(slide, colors, 5.42, ct + 0.65, 4.1, 1.85, margins, font_adj)
    add_section_box(slide, colors, 0.3, ct + 2.75, 9.4, 1.3, "Revenue by Service Line", colors["accent"], font_adj)
    services = parse_pipe_separated(data.get("service_lines") or data.get("serviceLines") or "", 5)
    svc_rev = [{"label": s[0] if s else "Service", "value": s[1] if len(s) > 1 else "20%"} for s in services]
    if svc_rev:
        add_stacked_bar(slide, colors, 0.42, ct + 3.22, 9.15, 0.4, svc_rev, font_adj)
    add_slide_footer(slide, colors, slide_num)

def render_growth_slide(slide, colors, data, slide_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Growth Strategy & Roadmap", "Path to continued expansion", font_adj)
    ct = 0.95
    add_section_box(slide, colors, 0.3, ct, 4.6, 2.0, "Key Growth Drivers", colors["primary"], font_adj)
    drivers = parse_lines(data.get("growth_drivers") or data.get("growthDrivers") or "", 5)
    if not drivers:
        drivers = ["AI adoption accelerating", "Cloud migration growing 25%", "Managed services demand", "Digital transformation mandates", "Geographic expansion"]
    for idx, driver in enumerate(drivers[:5]):
        db = slide.shapes.add_textbox(Inches(0.42), Inches(ct + 0.48 + idx * 0.3), Inches(4.4), Inches(0.28))
        db.text_frame.paragraphs[0].text = f"▸ {truncate_text(driver, 55)}"
        db.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        db.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_section_box(slide, colors, 5.1, ct, 4.6, 2.0, "Strategic Roadmap", colors["secondary"], font_adj)
    stb = slide.shapes.add_textbox(Inches(5.22), Inches(ct + 0.48), Inches(1.5), Inches(0.28))
    stb.text_frame.paragraphs[0].text = "0-12 Months"
    stb.text_frame.paragraphs[0].font.size = Pt(12)
    stb.text_frame.paragraphs[0].font.bold = True
    stb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
    short_goals = parse_lines(data.get("short_term_goals") or data.get("shortTermGoals") or "", 2)
    for idx, goal in enumerate(short_goals[:2]):
        gb = slide.shapes.add_textbox(Inches(5.22), Inches(ct + 0.78 + idx * 0.26), Inches(4.35), Inches(0.24))
        gb.text_frame.paragraphs[0].text = f"• {truncate_text(goal, 45)}"
        gb.text_frame.paragraphs[0].font.size = Pt(11)
        gb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    mtb = slide.shapes.add_textbox(Inches(5.22), Inches(ct + 1.35), Inches(1.5), Inches(0.28))
    mtb.text_frame.paragraphs[0].text = "1-3 Years"
    mtb.text_frame.paragraphs[0].font.size = Pt(12)
    mtb.text_frame.paragraphs[0].font.bold = True
    mtb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
    med_goals = parse_lines(data.get("medium_term_goals") or data.get("mediumTermGoals") or "", 2)
    for idx, goal in enumerate(med_goals[:2]):
        gb = slide.shapes.add_textbox(Inches(5.22), Inches(ct + 1.65 + idx * 0.26), Inches(4.35), Inches(0.24))
        gb.text_frame.paragraphs[0].text = f"• {truncate_text(goal, 45)}"
        gb.text_frame.paragraphs[0].font.size = Pt(11)
        gb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_section_box(slide, colors, 0.3, ct + 2.15, 9.4, 1.9, "Competitive Advantages", colors["accent"], font_adj)
    advantages = parse_lines(data.get("competitive_advantages") or data.get("competitiveAdvantages") or "", 6)
    if not advantages:
        advantages = ["Deep cloud expertise", "Proprietary AI platform", "Strong client relationships", "High retention rates", "Experienced leadership", "Capital-light model"]
    for idx, adv in enumerate(advantages[:6]):
        ax, ay = 0.42 + (idx % 2) * 4.7, ct + 2.58 + (idx // 2) * 0.42
        ab = slide.shapes.add_textbox(Inches(ax), Inches(ay), Inches(4.5), Inches(0.38))
        ab.text_frame.paragraphs[0].text = f"• {truncate_text(adv, 58)}"
        ab.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        ab.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_slide_footer(slide, colors, slide_num)

def render_thank_you_slide(slide, colors, data, doc_config):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(colors["dark_bg"])
    bg.line.fill.background()
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(4.5), Inches(10), Inches(0.04))
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(colors["secondary"])
    line.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0), Inches(1.8), Inches(10), Inches(0.9))
    tb.text_frame.paragraphs[0].text = "Thank You"
    tb.text_frame.paragraphs[0].font.size = Pt(48)
    tb.text_frame.paragraphs[0].font.bold = True
    tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    email = data.get("contact_email") or data.get("contactEmail")
    phone = data.get("contact_phone") or data.get("contactPhone")
    if email or phone:
        ib = slide.shapes.add_textbox(Inches(0), Inches(3.0), Inches(10), Inches(0.35))
        ib.text_frame.paragraphs[0].text = "For Further Information"
        ib.text_frame.paragraphs[0].font.size = Pt(14)
        ib.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
        ib.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        contact = (email or "") + (" | " if email and phone else "") + (phone or "")
        cb = slide.shapes.add_textbox(Inches(0), Inches(3.4), Inches(10), Inches(0.35))
        cb.text_frame.paragraphs[0].text = contact
        cb.text_frame.paragraphs[0].font.size = Pt(12)
        cb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
        cb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    cfb = slide.shapes.add_textbox(Inches(0), Inches(4.7), Inches(10), Inches(0.3))
    cfb.text_frame.paragraphs[0].text = "Strictly Private and Confidential"
    cfb.text_frame.paragraphs[0].font.size = Pt(10)
    cfb.text_frame.paragraphs[0].font.italic = True
    cfb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    cfb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# ============================================================================
# MAIN GENERATOR
# ============================================================================

async def generate_presentation(data: Dict, theme: str = "modern-blue") -> Presentation:
    """Generate complete presentation with AI-powered layouts"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    colors = get_theme_colors(theme)
    doc_type = data.get("document_type") or data.get("documentType") or "management-presentation"
    doc_config = DOCUMENT_CONFIGS.get(doc_type, DOCUMENT_CONFIGS["management-presentation"])
    primary_vertical = data.get("primary_vertical") or data.get("primaryVertical") or "technology"
    industry_data = INDUSTRY_DATA.get(primary_vertical, INDUSTRY_DATA["technology"])
    context = {"doc_config": doc_config, "industry_data": industry_data}
    slide_num = 1
    blank_layout = prs.slide_layouts[6]
    
    # Title
    slide = prs.slides.add_slide(blank_layout)
    render_title_slide(slide, colors, data, doc_config)
    
    # Disclaimer
    slide = prs.slides.add_slide(blank_layout)
    render_disclaimer_slide(slide, colors, data, slide_num)
    slide_num += 1
    
    # Executive Summary
    slide = prs.slides.add_slide(blank_layout)
    layout_rec = await analyze_data_for_layout(data, "executive-summary")
    render_executive_summary(slide, colors, data, slide_num, layout_rec, context)
    slide_num += 1
    
    # Services
    slide = prs.slides.add_slide(blank_layout)
    layout_rec = await analyze_data_for_layout(data, "services")
    render_services_slide(slide, colors, data, slide_num, layout_rec, context)
    slide_num += 1
    
    # Clients
    slide = prs.slides.add_slide(blank_layout)
    layout_rec = await analyze_data_for_layout(data, "clients")
    render_clients_slide(slide, colors, data, slide_num, layout_rec, context)
    slide_num += 1
    
    # Financials (not teaser)
    if doc_config.get("include_financial_detail", True):
        slide = prs.slides.add_slide(blank_layout)
        layout_rec = await analyze_data_for_layout(data, "financials")
        render_financials_slide(slide, colors, data, slide_num, layout_rec, context)
        slide_num += 1
    
    # Growth (not teaser)
    if doc_type != "teaser":
        slide = prs.slides.add_slide(blank_layout)
        layout_rec = await analyze_data_for_layout(data, "growth")
        render_growth_slide(slide, colors, data, slide_num, layout_rec, context)
        slide_num += 1
    
    # Thank You
    slide = prs.slides.add_slide(blank_layout)
    render_thank_you_slide(slide, colors, data, doc_config)
    
    return prs
