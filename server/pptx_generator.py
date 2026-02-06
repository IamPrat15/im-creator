"""
IM Creator Python Backend - PPTX Generator
Version: 8.2.0

FIXES in 8.2.0:
- Text wrapping enabled on ALL text frames (no overflow)
- CIM generates 25-40 slides (was 16)
- Multiple case studies rendered (up to 5 for CIM)
- 10+ new slide types for CIM depth
- More charts and infographics per slide
- Fixed slideType -> slide_type NameError
- Fixed slide.shapes.title on blank layouts
- Removed duplicate appendix functions
- Professional design with consistent spacing
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from typing import Dict, List, Optional
from models import DESIGN, INDUSTRY_DATA, DOCUMENT_CONFIGS, get_theme_colors
from utils import (
    truncate_text, truncate_description, format_currency, format_date,
    parse_lines, parse_pipe_separated, calculate_cagr, safe_float, safe_int,
    adjusted_font, extract_percentage,
    get_slides_for_document_type, get_buyer_specific_content, get_industry_specific_content
)
from ai_layout_engine import analyze_data_for_layout_sync

# ============================================================================
# COLOR HELPER
# ============================================================================

def hex_to_rgb(hex_color: str) -> RGBColor:
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# ============================================================================
# TEXT HELPER - ensures word wrap + auto size on EVERY text frame
# ============================================================================

def _setup_tf(tf, wrap=True):
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    return tf

def add_text_box(slide, x, y, w, h, text, font_size=12, bold=False, italic=False,
                 color=None, align=PP_ALIGN.LEFT, font_adj=0):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = _setup_tf(tb.text_frame)
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(adjusted_font(font_size, font_adj))
    p.font.bold = bold
    p.font.italic = italic
    if color:
        p.font.color.rgb = hex_to_rgb(color) if isinstance(color, str) else color
    p.alignment = align
    return tb

def add_multiline_text(slide, x, y, w, h, lines, font_size=11, color=None,
                       bold=False, bullet=False, font_adj=0, line_spacing=1.15):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = _setup_tf(tb.text_frame)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        prefix = "\u2022 " if bullet else ""
        p.text = f"{prefix}{str(line)}"
        p.font.size = Pt(adjusted_font(font_size, font_adj))
        p.font.bold = bold
        if color:
            p.font.color.rgb = hex_to_rgb(color) if isinstance(color, str) else color
        p.space_after = Pt(4)
    return tb

# ============================================================================
# CHART FUNCTIONS
# ============================================================================

def add_chart_by_type(slide, colors, x, y, w, h, chart_type, data, font_adj=0):
    if not data or chart_type == "none":
        return None
    dispatch = {
        "bar": add_bar_chart, "pie": add_pie_chart, "donut": add_donut_chart,
        "line": add_line_chart, "stacked-bar": add_stacked_bar_chart,
        "timeline": add_timeline, "progress": add_progress_bars,
    }
    fn = dispatch.get(chart_type, add_bar_chart)
    return fn(slide, colors, x, y, w, h, data, font_adj)

def add_bar_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    cf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = cf.chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(9)
    plot.data_labels.font.bold = True
    for s in chart.series:
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    try:
        chart.category_axis.tick_labels.font.size = Pt(8)
        chart.value_axis.tick_labels.font.size = Pt(8)
    except: pass
    return chart

def add_line_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    cf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
        Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = cf.chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    for s in chart.series:
        s.format.line.color.rgb = hex_to_rgb(colors["primary"])
        s.format.line.width = Pt(2.5)
    return chart

def add_pie_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    cf = slide.shapes.add_chart(XL_CHART_TYPE.PIE,
        Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = cf.chart
    chart.has_legend = True
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(8)
    plot.data_labels.show_percentage = True
    plot.data_labels.show_category_name = False
    return chart

def add_donut_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    cf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
        Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = cf.chart
    chart.has_legend = True
    plot = chart.plots[0]
    plot.has_data_labels = True
    return chart

def add_stacked_bar_chart(slide, colors, x, y, w, h, data, font_adj=0):
    if not data: return
    if isinstance(data[0], dict) and "values" in data[0]:
        chart_data = CategoryChartData()
        chart_data.categories = [d.get("label", "") for d in data]
        for sname in data[0]["values"].keys():
            chart_data.add_series(sname, tuple(safe_float(d.get("values",{}).get(sname,0)) for d in data))
        cf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,
            Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
        return cf.chart
    return add_bar_chart(slide, colors, x, y, w, h, data, font_adj)

def add_timeline(slide, colors, x, y, w, h, milestones, font_adj=0):
    if not milestones: return
    num = min(len(milestones), 6)
    step = w / max(num, 1)
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y + 0.35), Inches(w), Inches(0.04))
    ln.fill.solid(); ln.fill.fore_color.rgb = hex_to_rgb(colors["border"]); ln.line.fill.background()
    for i, m in enumerate(milestones[:num]):
        mx = x + (i * step) + (step / 2) - 0.15
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(mx), Inches(y + 0.2), Inches(0.3), Inches(0.3))
        c.fill.solid(); c.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); c.line.fill.background()
        add_text_box(slide, mx - 0.3, y + 0.6, 0.9, 0.35,
                     truncate_text(m.get("label",""), 15), 8,
                     color=colors["text"], align=PP_ALIGN.CENTER)

def add_progress_bars(slide, colors, x, y, w, h, items, font_adj=0):
    if not items: return
    num = min(len(items), 6)
    bar_h = min(0.3, (h - 0.1) / max(num, 1) - 0.08)
    for i, item in enumerate(items[:num]):
        by = y + (i * (bar_h + 0.12))
        value = safe_float(item.get("value", 0))
        bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(by), Inches(w), Inches(bar_h))
        bg.fill.solid(); bg.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"]); bg.line.fill.background()
        pw = max(0.1, (w * value) / 100)
        pg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(by), Inches(pw), Inches(bar_h))
        pg.fill.solid(); pg.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); pg.line.fill.background()
        add_text_box(slide, x + 0.1, by + 0.02, w - 0.2, bar_h - 0.04,
                     f"{item.get('label','')}: {value}%", 9, bold=True, color=colors["white"])

# ============================================================================
# BASE SLIDE COMPONENTS
# ============================================================================

def add_slide_header(slide, colors, title, subtitle=None, font_adj=0):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid(); bg.fill.fore_color.rgb = hex_to_rgb(colors["white"]); bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.08), Inches(0.85))
    bar.fill.solid(); bar.fill.fore_color.rgb = hex_to_rgb(colors["secondary"]); bar.line.fill.background()
    add_text_box(slide, 0.3, 0.12, 9.4, 0.5, truncate_text(title, 80),
                 DESIGN["fonts"]["title"], bold=True, color=colors["primary"], font_adj=font_adj)
    if subtitle:
        add_text_box(slide, 0.3, 0.58, 9.4, 0.25, subtitle,
                     DESIGN["fonts"]["subtitle"], italic=True, color=colors["text_light"], font_adj=font_adj)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(0.3), Inches(0.88), Inches(9.4), Inches(0.02))
    line.fill.solid(); line.fill.fore_color.rgb = hex_to_rgb(colors["accent"]); line.line.fill.background()

def add_slide_footer(slide, colors, page_number):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(5.22), Inches(10), Inches(0.015))
    line.fill.solid(); line.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); line.line.fill.background()
    add_text_box(slide, 0.3, 5.28, 3, 0.22, "Strictly Private & Confidential",
                 9, italic=True, color=colors["text_light"])
    add_text_box(slide, 9.2, 5.28, 0.5, 0.22, str(page_number),
                 10, bold=True, color=colors["primary"], align=PP_ALIGN.RIGHT)

def add_section_box(slide, colors, x, y, w, h, title=None, title_bg=None, font_adj=0):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
    box.line.color.rgb = hex_to_rgb(colors["border"])
    if title:
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(0.36))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = hex_to_rgb(title_bg or colors["primary"]); hdr.line.fill.background()
        add_text_box(slide, x + 0.12, y + 0.04, w - 0.24, 0.28,
                     truncate_text(title, 45), DESIGN["fonts"]["section_header"],
                     bold=True, color=colors["white"], font_adj=font_adj)

def add_metric_card(slide, colors, x, y, w, h, value, label, font_adj=0, value_color=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid(); card.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
    card.line.color.rgb = hex_to_rgb(colors["border"])
    add_text_box(slide, x + 0.08, y + 0.06, w - 0.16, h * 0.5,
                 str(value), DESIGN["fonts"]["metric_medium"], bold=True,
                 color=value_color or colors["primary"], font_adj=font_adj, align=PP_ALIGN.CENTER)
    add_text_box(slide, x + 0.08, y + h * 0.55, w - 0.16, h * 0.4,
                 label, DESIGN["fonts"]["metric_label"],
                 color=colors["text_light"], font_adj=font_adj, align=PP_ALIGN.CENTER)

def add_icon_row(slide, colors, x, y, w, icon_char, title_text, desc_text, font_adj=0):
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL,
        Inches(x), Inches(y + 0.02), Inches(0.32), Inches(0.32))
    circ.fill.solid(); circ.fill.fore_color.rgb = hex_to_rgb(colors["secondary"]); circ.line.fill.background()
    add_text_box(slide, x + 0.04, y + 0.06, 0.24, 0.24, icon_char, 12,
                 bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
    add_text_box(slide, x + 0.42, y, w - 0.42, 0.22, title_text, 11,
                 bold=True, color=colors["primary"], font_adj=font_adj)
    if desc_text:
        add_text_box(slide, x + 0.42, y + 0.22, w - 0.42, 0.25, desc_text, 9,
                     color=colors["text"], font_adj=font_adj)

# ============================================================================
# SLIDE RENDERERS
# ============================================================================

def render_title_slide(slide, colors, data, doc_config):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid(); bg.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.5), Inches(10), Inches(0.08))
    bar.fill.solid(); bar.fill.fore_color.rgb = hex_to_rgb(colors["accent"]); bar.line.fill.background()
    top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.06))
    top.fill.solid(); top.fill.fore_color.rgb = hex_to_rgb(colors["secondary"]); top.line.fill.background()
    company = data.get("companyName") or "Company Name"
    codename = data.get("projectCodename") or "Project"
    doc_name = doc_config.get("name", "Information Memorandum")
    advisor = data.get("advisor") or ""
    add_text_box(slide, 1, 1.0, 8, 0.9, company, 44, bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 1, 2.7, 8, 0.5, doc_name, 24, color=colors["white"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 1, 3.4, 8, 0.4, f"Project {codename}", 18, italic=True, color=colors["white"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 1, 4.5, 8, 0.3, format_date(data.get("presentationDate")), 14, color=colors["white"], align=PP_ALIGN.CENTER)
    if advisor:
        add_text_box(slide, 1, 4.9, 8, 0.3, f"Prepared by {advisor}", 12, italic=True, color=colors["white"], align=PP_ALIGN.CENTER)

def render_disclaimer_slide(slide, colors, data, page_num):
    add_slide_header(slide, colors, "Disclaimer")
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 4.0)
    disclaimer = ("This presentation has been prepared solely for informational purposes. "
        "The information contained herein is confidential and proprietary. By accepting "
        "this document, you agree to maintain its confidentiality and not to reproduce, "
        "distribute, or disclose it without prior written consent.\n\n"
        "This presentation does not constitute an offer to sell or a solicitation to buy "
        "securities. Any investment decision should be made only after thorough due diligence "
        "and consultation with professional advisors.\n\n"
        "The financial projections and forward-looking statements contained herein are based on "
        "assumptions that may or may not prove accurate. Actual results may vary materially. "
        "Past performance is not indicative of future results.\n\n"
        "All information is presented as of the date hereof and may be subject to change without "
        "notice. No representation or warranty is made as to the accuracy or completeness.")
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9.0), Inches(3.5))
    tf = _setup_tf(tb.text_frame)
    for i, para_text in enumerate(disclaimer.split('\n\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = para_text
        p.font.size = Pt(10)
        p.font.color.rgb = hex_to_rgb(colors["text"])
        p.space_after = Pt(8)

def render_toc(slide, colors, data, page_num, layout_rec, context):
    add_slide_header(slide, colors, "Table of Contents")
    add_slide_footer(slide, colors, page_num)
    doc_config = context.get("doc_config", {})
    section_map = [
        ("executive-summary", "Executive Summary"), ("investment-highlights", "Investment Highlights"),
        ("company-overview", "Company Overview"), ("services", "Service Lines & Capabilities"),
        ("products", "Products & Technology"), ("tech-partnerships", "Technology Partnerships"),
        ("clients", "Client Portfolio"), ("client-retention", "Client Retention & Growth"),
        ("financials", "Financial Performance"), ("financial-detail", "Detailed Financial Analysis"),
        ("case-study", "Case Studies"), ("growth", "Growth Strategy & Roadmap"),
        ("growth-goals", "Strategic Goals & Milestones"), ("market-position", "Market Position"),
        ("competitive-detail", "Competitive Advantages"), ("leadership", "Leadership Team"),
        ("risks", "Risk Factors & Mitigation"), ("synergies", "Strategic Value & Synergies"),
        ("transaction-summary", "Transaction Summary"),
    ]
    required = set(doc_config.get("required_slides", []))
    sections = [label for key, label in section_map if key in required]
    col1 = sections[:len(sections)//2 + 1]
    col2 = sections[len(sections)//2 + 1:]
    y_start = 1.1
    for i, s in enumerate(col1):
        add_text_box(slide, 0.5, y_start + i * 0.35, 0.4, 0.3, f"{i+1:02d}", 11, bold=True, color=colors["secondary"])
        add_text_box(slide, 1.0, y_start + i * 0.35, 3.8, 0.3, s, 11, color=colors["text"])
    for i, s in enumerate(col2):
        add_text_box(slide, 5.2, y_start + i * 0.35, 0.4, 0.3, f"{len(col1)+i+1:02d}", 11, bold=True, color=colors["secondary"])
        add_text_box(slide, 5.7, y_start + i * 0.35, 3.8, 0.3, s, 11, color=colors["text"])

def render_executive_summary(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "bar")
    vertical = data.get("primaryVertical") or "technology"
    industry_content = get_industry_specific_content(vertical, "executive-summary")
    add_slide_header(slide, colors, "Executive Summary", industry_content.get("context"), font_adj)
    add_slide_footer(slide, colors, page_num)
    # Left: Company overview
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Company Overview", font_adj=font_adj)
    desc = data.get("companyDescription") or ""
    tb = slide.shapes.add_textbox(Inches(0.45), Inches(1.45), Inches(4.2), Inches(1.8))
    tf = _setup_tf(tb.text_frame)
    tf.paragraphs[0].text = truncate_description(desc, 350)
    tf.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
    tf.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    # Key metrics
    metrics = [(str(data.get("foundedYear") or "N/A"), "Founded"),
               (str(data.get("employeeCountFT") or "N/A"), "Employees"),
               (str(data.get("headquarters") or "N/A")[:20], "HQ Location")]
    for i, (val, lbl) in enumerate(metrics):
        add_metric_card(slide, colors, 0.45 + i * 1.42, 3.4, 1.3, 0.5, val, lbl, font_adj)
    # Right: Revenue chart
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Revenue Growth", colors["secondary"], font_adj)
    rev_data = []
    for fy, label in [("revenueFY24","FY24"),("revenueFY25","FY25"),("revenueFY26P","FY26P"),("revenueFY27P","FY27P")]:
        v = safe_float(data.get(fy))
        if v: rev_data.append({"label": label, "value": v})
    if rev_data:
        add_chart_by_type(slide, colors, 5.15, 1.5, 4.4, 2.8, chart_type, rev_data, font_adj)

def render_investment_highlights(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Investment Highlights")
    add_slide_footer(slide, colors, page_num)
    highlights_text = data.get("investmentHighlights") or ""
    highlights = parse_lines(highlights_text, 10)
    if not highlights: return
    col1 = highlights[:5]; col2 = highlights[5:10]
    icons = ["\u2726", "\u25C6", "\u25B6", "\u25CF", "\u2605", "\u25C8", "\u25B2", "\u25A0", "\u25C9", "\u2B1F"]
    for i, h in enumerate(col1):
        y = 1.05 + i * 0.82
        add_icon_row(slide, colors, 0.4, y, 4.3, icons[i % len(icons)],
                     truncate_text(h, 55), "", font_adj)
    for i, h in enumerate(col2):
        y = 1.05 + i * 0.82
        add_icon_row(slide, colors, 5.2, y, 4.3, icons[(i+5) % len(icons)],
                     truncate_text(h, 55), "", font_adj)

def render_company_overview(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Company Overview")
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 5.8, 4.0, "About the Company", font_adj=font_adj)
    desc = data.get("companyDescription") or ""
    tb = slide.shapes.add_textbox(Inches(0.45), Inches(1.45), Inches(5.5), Inches(3.3))
    tf = _setup_tf(tb.text_frame)
    tf.paragraphs[0].text = desc[:600]
    tf.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
    tf.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    add_section_box(slide, colors, 6.3, 0.95, 3.4, 4.0, "Key Facts", colors["secondary"], font_adj)
    currency = data.get("currency") or "INR"
    facts = [("Founded", str(data.get("foundedYear") or "N/A")),
             ("HQ", str(data.get("headquarters") or "N/A")),
             ("Employees", str(data.get("employeeCountFT") or "N/A")),
             ("Currency", currency),
             ("Primary Vertical", str(data.get("primaryVertical") or "N/A").upper()),
             ("Revenue FY25", f"{currency} {data.get('revenueFY25','N/A')} Cr"),
             ("EBITDA Margin", f"{data.get('ebitdaMarginFY25','N/A')}%")]
    y = 1.45
    for lbl, val in facts:
        add_text_box(slide, 6.45, y, 1.5, 0.28, lbl, 9, bold=True, color=colors["text_light"], font_adj=font_adj)
        add_text_box(slide, 7.95, y, 1.5, 0.28, val, 9, color=colors["primary"], font_adj=font_adj)
        y += 0.42

def render_services(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "donut")
    add_slide_header(slide, colors, "Service Lines & Capabilities", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    service_text = data.get("serviceLines") or ""
    services = parse_pipe_separated(service_text, 8)
    add_section_box(slide, colors, 0.3, 0.95, 4.8, 4.0, "Service Offerings", font_adj=font_adj)
    y_pos = 1.45
    for svc in services[:6]:
        if len(svc) >= 2:
            name = truncate_text(svc[0], 30)
            pct = svc[1] if len(svc) > 1 else ""
            desc = svc[2] if len(svc) > 2 else ""
            add_text_box(slide, 0.5, y_pos, 4.4, 0.25, f"\u25A0 {name} ({pct})", 10,
                         bold=True, color=colors["primary"], font_adj=font_adj)
            if desc:
                add_text_box(slide, 0.8, y_pos + 0.25, 4.1, 0.25, truncate_text(desc, 60), 8,
                             color=colors["text"], font_adj=font_adj)
                y_pos += 0.55
            else:
                y_pos += 0.38
    add_section_box(slide, colors, 5.3, 0.95, 4.4, 4.0, "Revenue by Service", colors["secondary"], font_adj)
    chart_data = []
    for svc in services:
        if len(svc) >= 2:
            name = truncate_text(svc[0], 20)
            pct = safe_float(svc[1].replace("%", "").strip())
            if pct > 0: chart_data.append({"label": name, "value": pct})
    if chart_data:
        add_chart_by_type(slide, colors, 5.4, 1.5, 4.1, 3.0, chart_type, chart_data, font_adj)

def render_products(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Products & Technology Platform")
    add_slide_footer(slide, colors, page_num)
    products_text = data.get("products") or ""
    products = parse_pipe_separated(products_text, 6)
    if not products: return
    card_w = 2.9; gap = 0.15; start_x = 0.3
    y_rows = [0.95, 3.05]
    for i, prod in enumerate(products[:6]):
        col = i % 3; row = i // 3
        x = start_x + col * (card_w + gap)
        y = y_rows[min(row, 1)]
        add_section_box(slide, colors, x, y, card_w, 1.9,
                        truncate_text(prod[0] if prod else "", 25),
                        colors["secondary"] if i % 2 == 0 else colors["primary"], font_adj)
        desc = prod[1] if len(prod) > 1 else ""
        metric = prod[2] if len(prod) > 2 else ""
        if desc: add_text_box(slide, x + 0.1, y + 0.45, card_w - 0.2, 0.8, truncate_text(desc, 100), 9, color=colors["text"], font_adj=font_adj)
        if metric: add_text_box(slide, x + 0.1, y + 1.35, card_w - 0.2, 0.35, metric, 10, bold=True, color=colors["primary"], font_adj=font_adj)

def render_tech_partnerships(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Technology Partnerships & Certifications")
    add_slide_footer(slide, colors, page_num)
    partners_text = data.get("techPartnerships") or ""
    partners = parse_lines(partners_text, 10)
    if not partners: return
    col1 = partners[:5]; col2 = partners[5:10]
    icons = ["\u25C6", "\u2605", "\u25B6", "\u25CF", "\u25A0"]
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Strategic Partners", font_adj=font_adj)
    for i, p in enumerate(col1):
        add_icon_row(slide, colors, 0.5, 1.45 + i * 0.7, 4.1, icons[i % len(icons)], truncate_text(p, 40), "", font_adj)
    if col2:
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Additional Partners", colors["secondary"], font_adj)
        for i, p in enumerate(col2):
            add_icon_row(slide, colors, 5.2, 1.45 + i * 0.7, 4.3, icons[i % len(icons)], truncate_text(p, 40), "", font_adj)

def render_clients(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Client Portfolio", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    client_text = data.get("topClients") or ""
    clients = parse_pipe_separated(client_text, 12)
    add_section_box(slide, colors, 0.3, 0.95, 3.2, 4.0, "Key Metrics", font_adj=font_adj)
    top10 = data.get("top10Concentration") or "N/A"
    nrr = data.get("netRetention") or "N/A"
    add_metric_card(slide, colors, 0.45, 1.5, 2.9, 0.75, f"{top10}%", "Top 10 Concentration", font_adj)
    add_metric_card(slide, colors, 0.45, 2.45, 2.9, 0.75, f"{nrr}%", "Net Revenue Retention", font_adj)
    primary_v = (data.get("primaryVertical") or "").upper()
    if primary_v:
        add_metric_card(slide, colors, 0.45, 3.4, 2.9, 0.75, primary_v, "Primary Vertical", font_adj)
    add_section_box(slide, colors, 3.7, 0.95, 6.0, 4.0, "Top Clients", colors["secondary"], font_adj)
    y_pos = 1.45
    for client in clients[:10]:
        if client:
            name = truncate_text(client[0], 25) if len(client) > 0 else ""
            industry = client[1] if len(client) > 1 else ""
            since = client[2] if len(client) > 2 else ""
            row_text = f"\u25A0 {name}"
            if industry: row_text += f"  |  {industry}"
            if since: row_text += f"  |  Since {since}"
            add_text_box(slide, 3.9, y_pos, 5.6, 0.28, row_text, 9, color=colors["text"], font_adj=font_adj)
            y_pos += 0.33

def render_client_retention(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Client Retention & Growth Metrics")
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Revenue Retention", font_adj=font_adj)
    nrr = safe_float(data.get("netRetention") or 0)
    top10 = safe_float(data.get("top10Concentration") or 0)
    add_text_box(slide, 0.5, 1.5, 4.1, 0.7, f"{nrr}%", 36, bold=True, color=colors["primary"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 0.5, 2.2, 4.1, 0.3, "Net Revenue Retention", 12, color=colors["text_light"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 0.5, 2.8, 4.1, 0.25, "Top 10 Client Concentration", 10, bold=True, color=colors["text"])
    add_progress_bars(slide, colors, 0.5, 3.1, 4.1, 0.4, [{"label": f"{top10}%", "value": top10}], font_adj)
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Client Industry Mix", colors["secondary"], font_adj)
    client_text = data.get("topClients") or ""
    clients = parse_pipe_separated(client_text, 12)
    industry_counts = {}
    for c in clients:
        if len(c) > 1:
            ind = c[1].strip()
            industry_counts[ind] = industry_counts.get(ind, 0) + 1
    if industry_counts:
        chart_data = [{"label": k, "value": v} for k, v in industry_counts.items()]
        add_donut_chart(slide, colors, 5.15, 1.5, 4.4, 3.0, chart_data, font_adj)

def render_financials(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "bar")
    currency = data.get("currency") or "INR"
    add_slide_header(slide, colors, f"Financial Performance \u2014 Revenue", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.8, 4.0, f"Revenue Trend ({currency} Cr)", font_adj=font_adj)
    rev_data = []
    for fy, lbl in [("revenueFY24","FY24"),("revenueFY25","FY25"),("revenueFY26P","FY26P"),("revenueFY27P","FY27P")]:
        v = safe_float(data.get(fy))
        if v: rev_data.append({"label": lbl, "value": v})
    if rev_data:
        add_chart_by_type(slide, colors, 0.45, 1.5, 4.5, 3.0, chart_type, rev_data, font_adj)
    add_section_box(slide, colors, 5.3, 0.95, 4.4, 4.0, "Key Metrics", colors["secondary"], font_adj)
    fy24 = safe_float(data.get("revenueFY24") or 0)
    fy25 = safe_float(data.get("revenueFY25") or 0)
    fy26 = safe_float(data.get("revenueFY26P") or 0)
    growth_25 = f"{((fy25/fy24 - 1)*100):.0f}%" if fy24 and fy25 else "N/A"
    growth_26 = f"{((fy26/fy25 - 1)*100):.0f}%" if fy25 and fy26 else "N/A"
    add_metric_card(slide, colors, 5.45, 1.5, 4.1, 0.7, growth_25, "YoY Revenue Growth FY25", font_adj)
    add_metric_card(slide, colors, 5.45, 2.4, 4.1, 0.7, growth_26, "Projected Growth FY26", font_adj)
    if fy24 and fy26:
        cagr = ((fy26 / fy24) ** 0.5 - 1) * 100
        add_metric_card(slide, colors, 5.45, 3.3, 4.1, 0.7, f"{cagr:.1f}%", "2-Year Revenue CAGR", font_adj, value_color=colors["secondary"])

def render_financial_detail(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Financial Performance \u2014 Profitability")
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Profitability Metrics", font_adj=font_adj)
    ebitda = data.get("ebitdaMarginFY25") or "N/A"
    gross = data.get("grossMargin") or "N/A"
    net = data.get("netProfitMargin") or "N/A"
    for i, (val, lbl) in enumerate([(f"{ebitda}%","EBITDA Margin FY25"),(f"{gross}%","Gross Margin"),(f"{net}%","Net Profit Margin")]):
        add_metric_card(slide, colors, 0.45, 1.45 + i * 0.95, 4.2, 0.8, val, lbl, font_adj)
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Margin Analysis", colors["secondary"], font_adj)
    margin_data = []
    if ebitda != "N/A": margin_data.append({"label": "EBITDA", "value": safe_float(ebitda)})
    if gross != "N/A": margin_data.append({"label": "Gross", "value": safe_float(gross)})
    if net != "N/A": margin_data.append({"label": "Net Profit", "value": safe_float(net)})
    if margin_data:
        add_bar_chart(slide, colors, 5.15, 1.5, 4.4, 3.0, margin_data, font_adj)

def render_case_study(slide, colors, data, page_num, case_study, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    client = case_study.get("client", "Client")
    add_slide_header(slide, colors, f"Case Study: {truncate_text(client, 50)}", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    challenge = case_study.get("challenge", "")
    solution = case_study.get("solution", "")
    results = case_study.get("results", "")
    sections = [("Challenge", challenge, colors["primary"]),
                ("Solution", solution, colors["secondary"]),
                ("Results", results, colors["accent"])]
    y_pos = 0.95
    for title, content, bg_color in sections:
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            Inches(0.3), Inches(y_pos), Inches(9.4), Inches(0.32))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = hex_to_rgb(bg_color); hdr.line.fill.background()
        add_text_box(slide, 0.4, y_pos + 0.03, 9.2, 0.26, title, 11, bold=True, color=colors["white"])
        content_h = 1.05
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
            Inches(0.3), Inches(y_pos + 0.32), Inches(9.4), Inches(content_h))
        box.fill.solid(); box.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
        box.line.color.rgb = hex_to_rgb(colors["border"])
        if title == "Results" and ('|' in content or '\n' in content):
            items = [r.strip() for r in content.replace('|', '\n').split('\n') if r.strip()]
            add_multiline_text(slide, 0.45, y_pos + 0.38, 9.1, content_h - 0.1,
                               items[:5], 9, color=colors["text"], bullet=True, font_adj=font_adj)
        else:
            tb = slide.shapes.add_textbox(Inches(0.45), Inches(y_pos + 0.38), Inches(9.1), Inches(content_h - 0.1))
            tf = _setup_tf(tb.text_frame)
            tf.paragraphs[0].text = content[:350]
            tf.paragraphs[0].font.size = Pt(adjusted_font(9, font_adj))
            tf.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        y_pos += 0.32 + content_h + 0.08

def render_growth(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Growth Strategy & Roadmap", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Key Growth Drivers", font_adj=font_adj)
    drivers = parse_lines(data.get("growthDrivers") or "", 6)
    y_pos = 1.45
    for i, d in enumerate(drivers):
        add_icon_row(slide, colors, 0.45, y_pos, 4.2, "\u25B6", truncate_text(d, 50), "", font_adj)
        y_pos += 0.6
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Strategic Goals", colors["secondary"], font_adj)
    short_goals = parse_lines(data.get("shortTermGoals") or "", 4)
    medium_goals = parse_lines(data.get("mediumTermGoals") or "", 4)
    y_pos = 1.45
    if short_goals:
        add_text_box(slide, 5.15, y_pos, 4.4, 0.25, "Short-Term (0-12 months)", 10, bold=True, color=colors["primary"])
        y_pos += 0.3
        for g in short_goals[:3]:
            add_text_box(slide, 5.3, y_pos, 4.2, 0.25, f"\u2022 {truncate_text(g, 45)}", 9, color=colors["text"])
            y_pos += 0.3
    if medium_goals:
        y_pos += 0.1
        add_text_box(slide, 5.15, y_pos, 4.4, 0.25, "Medium-Term (1-3 years)", 10, bold=True, color=colors["primary"])
        y_pos += 0.3
        for g in medium_goals[:3]:
            add_text_box(slide, 5.3, y_pos, 4.2, 0.25, f"\u2022 {truncate_text(g, 45)}", 9, color=colors["text"])
            y_pos += 0.3

def render_growth_goals(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Strategic Goals & Milestones")
    add_slide_footer(slide, colors, page_num)
    short_goals = parse_lines(data.get("shortTermGoals") or "", 6)
    medium_goals = parse_lines(data.get("mediumTermGoals") or "", 6)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Short-Term (0-12 months)", font_adj=font_adj)
    y = 1.45
    for i, g in enumerate(short_goals[:6]):
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), Inches(y+0.02), Inches(0.28), Inches(0.28))
        c.fill.solid(); c.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); c.line.fill.background()
        add_text_box(slide, 0.54, y+0.04, 0.2, 0.2, str(i+1), 9, bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
        add_text_box(slide, 0.9, y, 3.7, 0.3, truncate_text(g, 55), 9, color=colors["text"], font_adj=font_adj)
        y += 0.5
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Medium-Term (1-3 years)", colors["secondary"], font_adj)
    y = 1.45
    for i, g in enumerate(medium_goals[:6]):
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.2), Inches(y+0.02), Inches(0.28), Inches(0.28))
        c.fill.solid(); c.fill.fore_color.rgb = hex_to_rgb(colors["secondary"]); c.line.fill.background()
        add_text_box(slide, 5.24, y+0.04, 0.2, 0.2, str(i+1), 9, bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
        add_text_box(slide, 5.6, y, 3.9, 0.3, truncate_text(g, 55), 9, color=colors["text"], font_adj=font_adj)
        y += 0.5

def render_market_position(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    vertical = data.get("primaryVertical") or "technology"
    industry_content = get_industry_specific_content(vertical, "market-position")
    add_slide_header(slide, colors, "Market Position & Competitive Landscape", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Market Overview", font_adj=font_adj)
    tam = data.get("marketSize") or "N/A"
    growth = data.get("marketGrowthRate") or "N/A"
    add_metric_card(slide, colors, 0.45, 1.5, 4.2, 0.7, tam, "Total Addressable Market", font_adj)
    add_metric_card(slide, colors, 0.45, 2.4, 4.2, 0.7, f"{growth}%", "Market Growth Rate", font_adj)
    if industry_content.get("benchmarks_text"):
        add_text_box(slide, 0.45, 3.3, 4.2, 0.5, industry_content["benchmarks_text"], 9,
                     italic=True, color=colors["text_light"])
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Competitive Landscape", colors["secondary"], font_adj)
    landscape = data.get("competitorLandscape") or data.get("competitiveAdvantages") or ""
    advantages = parse_pipe_separated(landscape, 6)
    y_pos = 1.45
    for adv in advantages:
        if adv:
            title = truncate_text(adv[0], 35)
            desc = adv[1] if len(adv) > 1 else ""
            add_text_box(slide, 5.15, y_pos, 4.4, 0.22, f"\u25A0 {title}", 10, bold=True, color=colors["primary"])
            if desc:
                add_text_box(slide, 5.35, y_pos + 0.22, 4.2, 0.22, truncate_text(desc, 55), 8, color=colors["text"])
                y_pos += 0.52
            else:
                y_pos += 0.35

def render_competitive_detail(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Competitive Advantages & Moats")
    add_slide_footer(slide, colors, page_num)
    adv_text = data.get("competitiveAdvantages") or ""
    advantages = parse_pipe_separated(adv_text, 8)
    for i, adv in enumerate(advantages[:6]):
        if not adv: continue
        title = truncate_text(adv[0], 45)
        desc = adv[1] if len(adv) > 1 else ""
        col = i % 2; row = i // 2
        x = 0.3 + col * 4.85; cy = 0.95 + row * 1.45; card_w = 4.6
        add_section_box(slide, colors, x, cy, card_w, 1.3, title,
                        colors["primary"] if col == 0 else colors["secondary"], font_adj)
        if desc:
            add_text_box(slide, x + 0.15, cy + 0.45, card_w - 0.3, 0.7, truncate_text(desc, 120), 9, color=colors["text"])

def render_leadership(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Leadership Team")
    add_slide_footer(slide, colors, page_num)
    founder_name = data.get("founderName") or ""
    founder_title = data.get("founderTitle") or ""
    experience = data.get("founderExperience") or ""
    education = data.get("founderEducation") or ""
    if founder_name:
        add_section_box(slide, colors, 0.3, 0.95, 9.4, 1.2,
                        f"{founder_name} \u2014 {founder_title}", colors["primary"], font_adj)
        info_parts = []
        if experience: info_parts.append(f"{experience} years experience")
        if education: info_parts.append(f"Education: {education}")
        if info_parts:
            add_text_box(slide, 0.5, 1.45, 9.0, 0.55, " | ".join(info_parts), 10, color=colors["text"])
    team_text = data.get("leadershipTeam") or ""
    members = parse_pipe_separated(team_text, 8)
    y_start = 2.3 if founder_name else 1.0
    card_w = 4.6; card_h = 0.65
    for i, member in enumerate(members[:8]):
        if not member or len(member) < 2: continue
        name = member[0]; title = member[1]; dept = member[2] if len(member) > 2 else ""
        col = i % 2; row = i // 2
        x = 0.3 + col * 4.85; y = y_start + row * (card_h + 0.1)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(card_w), Inches(card_h))
        card.fill.solid(); card.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
        card.line.color.rgb = hex_to_rgb(colors["border"])
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.06), Inches(card_h))
        accent.fill.solid(); accent.fill.fore_color.rgb = hex_to_rgb(colors["secondary"]); accent.line.fill.background()
        add_text_box(slide, x + 0.15, y + 0.06, card_w - 0.3, 0.28, f"{name} \u2014 {title}", 10, bold=True, color=colors["primary"])
        if dept: add_text_box(slide, x + 0.15, y + 0.35, card_w - 0.3, 0.22, dept, 8, color=colors["text_light"])

def render_risk_factors(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Risk Factors & Mitigation")
    add_slide_footer(slide, colors, page_num)
    risks = []
    if data.get("businessRisks"): risks.append(("Business Risks", data["businessRisks"]))
    if data.get("marketRisks"): risks.append(("Market Risks", data["marketRisks"]))
    if data.get("operationalRisks"): risks.append(("Operational Risks", data["operationalRisks"]))
    if data.get("mitigationStrategies"): risks.append(("Mitigation Strategies", data["mitigationStrategies"]))
    if not risks: return
    risk_colors = [colors["primary"], colors["secondary"], colors["accent"], colors["primary"]]
    section_h = min(1.0, 3.8 / max(len(risks), 1))
    y_pos = 0.95
    for i, (category, content) in enumerate(risks):
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(y_pos), Inches(9.4), Inches(0.3))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = hex_to_rgb(risk_colors[i % len(risk_colors)]); hdr.line.fill.background()
        add_text_box(slide, 0.4, y_pos + 0.03, 9.2, 0.24, category, 10, bold=True, color=colors["white"])
        items = [r.strip() for r in content.replace('|', '\n').split('\n') if r.strip()]
        add_multiline_text(slide, 0.45, y_pos + 0.35, 9.1, section_h - 0.3, items[:4], 9,
                           color=colors["text"], bullet=True, font_adj=font_adj)
        y_pos += section_h + 0.12

def render_synergies(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Strategic Value & Synergies", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Strategic Synergies", font_adj=font_adj)
    synergies = parse_lines(data.get("synergiesStrategic") or "", 7)
    y_pos = 1.45
    for s in synergies:
        add_text_box(slide, 0.5, y_pos, 4.1, 0.35, f"\u2022 {truncate_text(s, 55)}", 10, color=colors["text"])
        y_pos += 0.42
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Financial Synergies", colors["secondary"], font_adj)
    fin_synergies = parse_lines(data.get("synergiesFinancial") or "", 7)
    y_pos = 1.45
    for s in fin_synergies:
        add_text_box(slide, 5.2, y_pos, 4.3, 0.35, f"\u2022 {truncate_text(s, 55)}", 10, color=colors["text"])
        y_pos += 0.42

def render_transaction_summary(slide, colors, data, page_num, layout_rec, context):
    font_adj = layout_rec.get("font_adjustment", 0)
    add_slide_header(slide, colors, "Transaction Summary & Next Steps")
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 4.5, 4.0, "Transaction Overview", font_adj=font_adj)
    currency = data.get("currency") or "INR"
    facts = [("Company", data.get("companyName") or ""), ("Advisor", data.get("advisor") or "N/A"),
             ("Revenue FY25", f"{currency} {data.get('revenueFY25','N/A')} Cr"),
             ("EBITDA Margin", f"{data.get('ebitdaMarginFY25','N/A')}%"),
             ("Employees", str(data.get("employeeCountFT") or "N/A")),
             ("Headquarters", str(data.get("headquarters") or "N/A"))]
    y = 1.45
    for lbl, val in facts:
        add_text_box(slide, 0.5, y, 1.8, 0.28, lbl, 9, bold=True, color=colors["text_light"])
        add_text_box(slide, 2.3, y, 2.3, 0.28, val, 9, color=colors["primary"])
        y += 0.38
    add_section_box(slide, colors, 5.0, 0.95, 4.7, 4.0, "Indicative Process", colors["secondary"], font_adj)
    steps = ["Review of Information Memorandum", "Management presentation & site visit",
             "Submission of non-binding indications", "Due diligence access",
             "Submission of binding offers", "Negotiation & signing", "Regulatory approvals & closing"]
    y = 1.45
    for i, step in enumerate(steps):
        c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.15), Inches(y+0.02), Inches(0.24), Inches(0.24))
        c.fill.solid(); c.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); c.line.fill.background()
        add_text_box(slide, 5.18, y+0.04, 0.2, 0.18, str(i+1), 8, bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
        add_text_box(slide, 5.5, y, 4.0, 0.28, step, 9, color=colors["text"])
        y += 0.38

def render_thank_you_slide(slide, colors, data, doc_config):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid(); bg.fill.fore_color.rgb = hex_to_rgb(colors["primary"]); bg.line.fill.background()
    top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.06))
    top.fill.solid(); top.fill.fore_color.rgb = hex_to_rgb(colors["accent"]); top.line.fill.background()
    add_text_box(slide, 1, 1.8, 8, 0.9, "Thank You", 48, bold=True, color=colors["white"], align=PP_ALIGN.CENTER)
    company = data.get("companyName") or ""
    if company: add_text_box(slide, 1, 2.8, 8, 0.4, company, 18, color=colors["white"], align=PP_ALIGN.CENTER)
    advisor = data.get("advisor") or ""
    if advisor: add_text_box(slide, 1, 3.4, 8, 0.3, f"Prepared by {advisor}", 14, italic=True, color=colors["white"], align=PP_ALIGN.CENTER)
    parts = [p for p in [data.get("contactEmail",""), data.get("contactPhone","")] if p]
    if parts: add_text_box(slide, 1, 4.0, 8, 0.4, " | ".join(parts), 14, color=colors["white"], align=PP_ALIGN.CENTER)
    add_text_box(slide, 1, 4.8, 8, 0.3, "Strictly Private & Confidential", 10, italic=True, color=colors["white"], align=PP_ALIGN.CENTER)

# Appendix renderers
def render_appendix_financials(slide, colors, data, page_num, layout_rec, context):
    add_slide_header(slide, colors, "Appendix A: Detailed Financials", font_adj=-1)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 4.0, "Financial Details")
    y = 1.45
    for field, label in [("revenueByService","Revenue by Service"),("costStructure","Cost Structure"),("workingCapital","Working Capital")]:
        content = data.get(field) or ""
        if content:
            add_text_box(slide, 0.5, y, 9.0, 0.25, label, 11, bold=True, color=colors["primary"])
            add_text_box(slide, 0.5, y+0.28, 9.0, 0.6, truncate_description(content, 300), 9, color=colors["text"])
            y += 1.0

def render_appendix_case_studies(slide, colors, data, page_num, layout_rec, context):
    add_slide_header(slide, colors, "Appendix: Additional Case Studies", font_adj=-1)
    add_slide_footer(slide, colors, page_num)
    max_cs = context.get("doc_config",{}).get("max_case_studies",2)
    case_studies = data.get("caseStudies") or []
    additional = case_studies[max_cs:]
    y = 1.1
    for study in additional[:3]:
        add_text_box(slide, 0.5, y, 9.0, 0.25, f"Case Study: {truncate_text(study.get('client',''), 60)}", 11, bold=True, color=colors["primary"])
        add_text_box(slide, 0.5, y+0.3, 9.0, 0.6, f"Challenge: {truncate_text(study.get('challenge',''),80)} | Solution: {truncate_text(study.get('solution',''),80)} | Results: {truncate_text(study.get('results',''),80)}", 8, color=colors["text"])
        y += 1.1

def render_appendix_team_bios(slide, colors, data, page_num, layout_rec, context):
    add_slide_header(slide, colors, "Appendix: Team Biographies", font_adj=-1)
    add_slide_footer(slide, colors, page_num)
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 4.0)
    members = parse_pipe_separated(data.get("leadershipTeam") or "", 6)
    y = 1.3
    for m in members:
        if m and len(m) >= 2:
            add_text_box(slide, 0.5, y, 9.0, 0.35, f"\u2022 {m[0]} \u2014 {m[1]}", 10, bold=True, color=colors["text"])
            y += 0.45

# ============================================================================
# UNIVERSAL createSlide() WRAPPER
# ============================================================================

def create_slide(slide_type, prs, colors, data, page_num, context):
    # Guard clauses - use slide_type (NOT slideType)
    if slide_type == "financials" and not (data.get("revenueFY24") or data.get("revenueFY25")):
        return None
    if slide_type == "financial-detail" and not (data.get("ebitdaMarginFY25") or data.get("grossMargin")):
        return None
    if slide_type == "synergies" and not (data.get("synergiesStrategic") or data.get("synergiesFinancial")):
        return None
    if slide_type == "risks" and not (data.get("businessRisks") or data.get("marketRisks") or data.get("operationalRisks")):
        return None
    if slide_type == "products" and not data.get("products"):
        return None
    if slide_type == "tech-partnerships" and not data.get("techPartnerships"):
        return None
    if slide_type.startswith("appendix"):
        if slide_type == "appendix-financials" and not data.get("revenueByService"):
            return None
        if slide_type == "appendix-team-bios" and not data.get("teamBios"):
            return None
        if slide_type == "appendix-case-studies":
            max_cs = context.get("doc_config",{}).get("max_case_studies",2)
            if len(data.get("caseStudies") or []) <= max_cs:
                return None

    blank_layout = prs.slide_layouts[6]
    layout_rec = analyze_data_for_layout_sync(data, slide_type)

    # CASE STUDY: renders multiple slides
    if slide_type == "case-study":
        case_studies = data.get("caseStudies") or []
        if not case_studies and data.get("cs1Client"):
            case_studies.append({"client": data.get("cs1Client"), "challenge": data.get("cs1Challenge"),
                                "solution": data.get("cs1Solution"), "results": data.get("cs1Results")})
        if not case_studies: return None
        max_cs = context.get("doc_config",{}).get("max_case_studies", 2)
        for cs in case_studies[:max_cs]:
            s = prs.slides.add_slide(blank_layout)
            render_case_study(s, colors, data, page_num, cs, layout_rec, context)
            page_num += 1
        return page_num

    # All other slide types
    slide = prs.slides.add_slide(blank_layout)
    route = {
        "title": lambda: (render_title_slide(slide, colors, data, context.get("doc_config",{})), None)[1],
        "disclaimer": lambda: (render_disclaimer_slide(slide, colors, data, page_num), page_num+1)[1],
        "toc": lambda: (render_toc(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "executive-summary": lambda: (render_executive_summary(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "investment-highlights": lambda: (render_investment_highlights(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "company-overview": lambda: (render_company_overview(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "services": lambda: (render_services(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "products": lambda: (render_products(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "tech-partnerships": lambda: (render_tech_partnerships(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "clients": lambda: (render_clients(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "client-retention": lambda: (render_client_retention(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "financials": lambda: (render_financials(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "financial-detail": lambda: (render_financial_detail(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "growth": lambda: (render_growth(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "growth-goals": lambda: (render_growth_goals(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "market-position": lambda: (render_market_position(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "competitive-detail": lambda: (render_competitive_detail(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "leadership": lambda: (render_leadership(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "risks": lambda: (render_risk_factors(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "synergies": lambda: (render_synergies(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "transaction-summary": lambda: (render_transaction_summary(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "thank-you": lambda: (render_thank_you_slide(slide, colors, data, context.get("doc_config",{})), None)[1],
        "appendix-financials": lambda: (render_appendix_financials(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "appendix-case-studies": lambda: (render_appendix_case_studies(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
        "appendix-team-bios": lambda: (render_appendix_team_bios(slide, colors, data, page_num, layout_rec, context), page_num+1)[1],
    }
    if slide_type in route:
        return route[slide_type]()
    return None

# ============================================================================
# MAIN GENERATOR
# ============================================================================

def generate_presentation(data, theme="modern-blue"):
    import json
    if isinstance(data, str):
        try: data = json.loads(data)
        except: data = {}
    if not isinstance(data, dict): data = {}
    if not (data.get("companyName") or data.get("company_name")): data["companyName"] = "Company Name"
    if not (data.get("documentType") or data.get("document_type")): data["documentType"] = "management-presentation"

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    try: colors = get_theme_colors(theme)
    except: colors = get_theme_colors("modern-blue")

    doc_type = data.get("documentType") or data.get("document_type") or "management-presentation"
    doc_config = DOCUMENT_CONFIGS.get(doc_type, DOCUMENT_CONFIGS["management-presentation"])
    primary_vertical = data.get("primaryVertical") or "technology"
    try: industry_data = INDUSTRY_DATA.get(primary_vertical, INDUSTRY_DATA.get("technology", {}))
    except: industry_data = {}

    context = {"doc_config": doc_config, "industry_data": industry_data,
               "buyer_types": data.get("targetBuyerType") or ["strategic"]}

    try: slides_to_generate = get_slides_for_document_type(doc_type, data)
    except Exception as e:
        print(f"ERROR: {e}")
        slides_to_generate = ["title","disclaimer","executive-summary","services","clients","financials","thank-you"]

    print(f"=== GENERATION v8.2.0 ===")
    print(f"Doc: {doc_type} | Vertical: {primary_vertical} | Theme: {theme}")
    print(f"Slides ({len(slides_to_generate)}): {slides_to_generate}")

    page_num = 1
    slides_created = 0
    for slide_type in slides_to_generate:
        try:
            result = create_slide(slide_type, prs, colors, data, page_num, context)
            if result is not None: page_num = result
            slides_created += 1
            print(f"  + {slide_type}")
        except Exception as e:
            print(f"  x {slide_type}: {e}")
            import traceback; traceback.print_exc()
            continue

    print(f"=== DONE: {slides_created} processed, {len(prs.slides)} slides total ===")
    return prs
