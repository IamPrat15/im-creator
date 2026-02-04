"""
IM Creator Python Backend - PPTX Generator
Version: 7.1.0

Implements Requirements #15-18:
- Universal createSlide() wrapper
- Dedicated render functions for each slide type
- Chart helper addChartByType()
- AI-powered layout recommendations applied consistently
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor  # Fixed import
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
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
    """Convert hex color to RGBColor"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

# ============================================================================
# REQUIREMENT #17: CHART HELPER - addChartByType()
# ============================================================================

def add_chart_by_type(slide, colors, x, y, w, h, chart_type, data, font_adj=0):
    """
    Universal chart dispatcher - Implements Requirement #17
    Routes to appropriate chart function based on chart_type
    
    Args:
        chart_type: "bar", "pie", "donut", "timeline", "progress", "stacked-bar", "none"
        data: Chart data in appropriate format
    """
    if not data or chart_type == "none":
        return None
    
    if chart_type == "bar":
        return add_bar_chart(slide, colors, x, y, w, h, data, font_adj)
    elif chart_type == "pie":
        return add_pie_chart(slide, colors, x, y, w, h, data, font_adj)
    elif chart_type == "donut":
        return add_donut_chart(slide, colors, x, y, w, h, data, font_adj)
    elif chart_type == "timeline":
        return add_timeline(slide, colors, x, y, w, h, data, font_adj)
    elif chart_type == "progress":
        return add_progress_bars(slide, colors, x, y, w, h, data, font_adj)
    elif chart_type == "stacked-bar":
        return add_stacked_bar_chart(slide, colors, x, y, w, h, data, font_adj)
    else:
        return None

# ============================================================================
# BASE SLIDE COMPONENTS
# ============================================================================

def add_slide_header(slide, colors, title, subtitle=None, font_adj=0):
    """Add slide header with title and optional subtitle"""
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
    """Add slide footer with confidentiality notice and page number"""
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
    """Add a section box with optional title"""
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
    """Add a metric card with value and label"""
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

# ============================================================================
# CHART FUNCTIONS
# ============================================================================

def add_bar_chart(slide, colors, x, y, w, h, data, font_adj=0):
    """Add bar chart"""
    if not data:
        return
    
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, 
        Inches(x), Inches(y), Inches(w), Inches(h), 
        chart_data
    ).chart
    
    chart.has_legend = False
    chart.plots[0].has_data_labels = True
    
    for series in chart.series:
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    
    return chart

def add_pie_chart(slide, colors, x, y, w, h, data, font_adj=0):
    """Add pie chart"""
    if not data:
        return
    
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE, 
        Inches(x), Inches(y), Inches(w), Inches(h), 
        chart_data
    ).chart
    
    chart.has_legend = True
    chart.plots[0].has_data_labels = True
    
    return chart

def add_donut_chart(slide, colors, x, y, w, h, data, font_adj=0):
    """Add donut chart"""
    if not data:
        return
    
    chart_data = CategoryChartData()
    chart_data.categories = [d.get("label", "") for d in data]
    chart_data.add_series("Values", tuple(safe_float(d.get("value", 0)) for d in data))
    
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT, 
        Inches(x), Inches(y), Inches(w), Inches(h), 
        chart_data
    ).chart
    
    chart.has_legend = True
    chart.plots[0].has_data_labels = True
    
    return chart

def add_timeline(slide, colors, x, y, w, h, milestones, font_adj=0):
    """Add timeline visualization"""
    if not milestones:
        return
    
    num = len(milestones)
    step = w / max(num, 1)
    
    for i, milestone in enumerate(milestones):
        mx = x + (i * step)
        my = y + 0.2
        
        # Circle
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(mx), Inches(my), Inches(0.3), Inches(0.3))
        circ.fill.solid()
        circ.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
        circ.line.fill.background()
        
        # Label
        tb = slide.shapes.add_textbox(Inches(mx - 0.2), Inches(my + 0.4), Inches(0.7), Inches(0.3))
        tb.text_frame.text = truncate_text(milestone.get("label", ""), 20)
        tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Line to next
        if i < num - 1:
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(mx + 0.3), Inches(my + 0.13), Inches(step - 0.3), Inches(0.04))
            line.fill.solid()
            line.fill.fore_color.rgb = hex_to_rgb(colors["border"])
            line.line.fill.background()

def add_progress_bars(slide, colors, x, y, w, h, items, font_adj=0):
    """Add progress bars"""
    if not items:
        return
    
    num = len(items)
    bar_height = min(0.25, (h - 0.1) / max(num, 1))
    gap = 0.1
    
    for i, item in enumerate(items):
        by = y + (i * (bar_height + gap))
        label = item.get("label", "")
        value = safe_float(item.get("value", 0))
        
        # Background bar
        bg_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(by), Inches(w), Inches(bar_height))
        bg_bar.fill.solid()
        bg_bar.fill.fore_color.rgb = hex_to_rgb(colors["light_bg"])
        bg_bar.line.fill.background()
        
        # Progress bar
        progress_width = (w * value) / 100
        prog_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(by), Inches(progress_width), Inches(bar_height))
        prog_bar.fill.solid()
        prog_bar.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
        prog_bar.line.fill.background()
        
        # Label
        tb = slide.shapes.add_textbox(Inches(x + 0.1), Inches(by + 0.05), Inches(w - 0.2), Inches(bar_height - 0.1))
        tb.text_frame.text = f"{label}: {value}%"
        tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(11, font_adj))
        tb.text_frame.paragraphs[0].font.bold = True
        tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])

def add_stacked_bar_chart(slide, colors, x, y, w, h, data, font_adj=0):
    """Add stacked bar chart"""
    # Simplified stacked bar - can be enhanced
    return add_bar_chart(slide, colors, x, y, w, h, data, font_adj)

# ============================================================================
# REQUIREMENT #16: DEDICATED RENDER FUNCTIONS
# ============================================================================

def render_executive_summary(slide, colors, data, page_num, layout_rec, context):
    """
    Dedicated render function for Executive Summary slide
    Applies layout_rec.layout, font_adjustment, and chart_type
    """
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "bar")
    layout = layout_rec.get("layout", "two-column")
    
    # Get buyer-specific and industry content
    buyer_types = data.get("targetBuyerType") or data.get("target_buyer_type") or ["strategic"]
    vertical = data.get("primaryVertical") or data.get("primary_vertical") or "technology"
    buyer_content = get_buyer_specific_content(buyer_types, "executive-summary", data)
    industry_content = get_industry_specific_content(vertical, "executive-summary")
    
    add_slide_header(slide, colors, "Executive Summary", industry_content.get("context"), font_adj)
    add_slide_footer(slide, colors, page_num)
    
    if layout == "two-column":
        # Left: Company overview
        add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Company Overview", font_adj=font_adj)
        
        # Company description
        desc = data.get("companyDescription") or data.get("company_description") or ""
        tb = slide.shapes.add_textbox(Inches(0.45), Inches(1.45), Inches(4.2), Inches(1.5))
        tb.text_frame.text = truncate_description(desc, 300)
        tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
        tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        
        # Key metrics
        metrics_y = 3.1
        metrics = [
            (data.get("foundedYear") or "N/A", "Founded"),
            (data.get("employeeCountFT") or "N/A", "Employees"),
            (data.get("headquarters") or "N/A", "Location")
        ]
        
        for i, (value, label) in enumerate(metrics):
            add_metric_card(slide, colors, 0.45 + (i * 1.45), metrics_y, 1.3, 0.5, value, label, font_adj)
        
        # Right: Revenue chart or metrics
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Revenue Growth", colors["secondary"], font_adj)
        
        # Prepare revenue data
        revenue_data = []
        fy24 = safe_float(data.get("revenueFY24"))
        fy25 = safe_float(data.get("revenueFY25"))
        fy26 = safe_float(data.get("revenueFY26P"))
        fy27 = safe_float(data.get("revenueFY27P"))
        
        if fy24:
            revenue_data.append({"label": "FY24", "value": fy24})
        if fy25:
            revenue_data.append({"label": "FY25", "value": fy25})
        if fy26:
            revenue_data.append({"label": "FY26P", "value": fy26})
        if fy27:
            revenue_data.append({"label": "FY27P", "value": fy27})
        
        # Add chart using AI recommendation
        if revenue_data and chart_type != "none":
            add_chart_by_type(slide, colors, 5.2, 1.5, 4.3, 2.8, chart_type, revenue_data, font_adj)
    
    elif layout == "full-width":
        # Single wide section
        add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8, "Executive Summary", font_adj=font_adj)
        
        desc = data.get("companyDescription") or ""
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9.0), Inches(2.8))
        tb.text_frame.text = truncate_description(desc, 600)
        tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_large"], font_adj))
        tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])


def render_services(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Services slide"""
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "donut")
    layout = layout_rec.get("layout", "two-column")
    
    add_slide_header(slide, colors, "Service Lines & Capabilities", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    # Parse service lines
    service_text = data.get("serviceLines") or data.get("service_lines") or ""
    services = parse_pipe_separated(service_text, 8)
    
    if layout == "two-column":
        # Left: Service descriptions
        add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Service Offerings", font_adj=font_adj)
        
        y_pos = 1.5
        for service in services[:6]:
            if len(service) >= 2:
                name = truncate_text(service[0], 30)
                pct = service[1] if len(service) > 1 else ""
                
                tb = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(4.1), Inches(0.35))
                tb.text_frame.text = f"• {name} ({pct})"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.45
        
        # Right: Chart
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Revenue Distribution", colors["secondary"], font_adj)
        
        # Prepare chart data
        chart_data = []
        for service in services:
            if len(service) >= 2:
                name = truncate_text(service[0], 20)
                pct_str = service[1].replace("%", "").strip()
                pct = safe_float(pct_str)
                if pct > 0:
                    chart_data.append({"label": name, "value": pct})
        
        if chart_data and chart_type != "none":
            add_chart_by_type(slide, colors, 5.2, 1.5, 4.3, 2.8, chart_type, chart_data, font_adj)


def render_clients(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Clients slide"""
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "donut")
    layout = layout_rec.get("layout", "two-column-wide-right")
    
    add_slide_header(slide, colors, "Client Portfolio", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    # Parse clients
    client_text = data.get("topClients") or data.get("top_clients") or ""
    clients = parse_pipe_separated(client_text, 12)
    
    if layout == "two-column-wide-right":
        # Left: Metrics
        add_section_box(slide, colors, 0.3, 0.95, 3.2, 3.8, "Key Metrics", font_adj=font_adj)
        
        top10 = data.get("top10Concentration") or data.get("top_10_concentration") or "N/A"
        nrr = data.get("netRetention") or data.get("net_retention") or "N/A"
        
        add_metric_card(slide, colors, 0.45, 1.5, 2.9, 0.75, f"{top10}%", "Top 10 Concentration", font_adj)
        add_metric_card(slide, colors, 0.45, 2.5, 2.9, 0.75, f"{nrr}%", "Net Revenue Retention", font_adj)
        
        # Right: Client list & chart
        add_section_box(slide, colors, 3.7, 0.95, 6.0, 3.8, "Top Clients", colors["secondary"], font_adj)
        
        y_pos = 1.5
        for client in clients[:10]:
            if client:
                name = truncate_text(client[0] if len(client) > 0 else "", 40)
                tb = slide.shapes.add_textbox(Inches(3.9), Inches(y_pos), Inches(5.6), Inches(0.3))
                tb.text_frame.text = f"• {name}"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_small"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.32


def render_financials(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Financials slide"""
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "bar")
    layout = layout_rec.get("layout", "two-column")
    
    # Get buyer-specific emphasis
    buyer_types = data.get("targetBuyerType") or ["strategic"]
    buyer_content = get_buyer_specific_content(buyer_types, "financials", data)
    
    add_slide_header(slide, colors, "Financial Performance", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    if layout == "two-column":
        # Left: Revenue chart
        add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Revenue Trend", font_adj=font_adj)
        
        revenue_data = []
        if data.get("revenueFY24"):
            revenue_data.append({"label": "FY24", "value": safe_float(data.get("revenueFY24"))})
        if data.get("revenueFY25"):
            revenue_data.append({"label": "FY25", "value": safe_float(data.get("revenueFY25"))})
        if data.get("revenueFY26P"):
            revenue_data.append({"label": "FY26P", "value": safe_float(data.get("revenueFY26P"))})
        if data.get("revenueFY27P"):
            revenue_data.append({"label": "FY27P", "value": safe_float(data.get("revenueFY27P"))})
        
        if revenue_data and chart_type != "none":
            add_chart_by_type(slide, colors, 0.5, 1.5, 4.1, 2.8, chart_type, revenue_data, font_adj)
        
        # Right: Margins (prioritized for financial buyers)
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Profitability Metrics", colors["secondary"], font_adj)
        
        ebitda = data.get("ebitdaMarginFY25") or data.get("ebitda_margin_fy25") or "N/A"
        gross = data.get("grossMargin") or data.get("gross_margin") or "N/A"
        net = data.get("netProfitMargin") or data.get("net_profit_margin") or "N/A"
        
        add_metric_card(slide, colors, 5.2, 1.5, 4.3, 0.65, f"{ebitda}%", "EBITDA Margin FY25", font_adj)
        add_metric_card(slide, colors, 5.2, 2.35, 4.3, 0.65, f"{gross}%", "Gross Margin", font_adj)
        add_metric_card(slide, colors, 5.2, 3.2, 4.3, 0.65, f"{net}%", "Net Profit Margin", font_adj)


def render_case_study(slide, colors, data, page_num, case_study, layout_rec, context):
    """Dedicated render function for Case Study slide"""
    font_adj = layout_rec.get("font_adjustment", 0)
    layout = layout_rec.get("layout", "full-width")
    
    client = case_study.get("client", "Client")
    add_slide_header(slide, colors, f"Case Study: {truncate_text(client, 50)}", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    # Full-width layout for case study
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8, font_adj=font_adj)
    
    # Challenge, Solution, Results
    challenge = case_study.get("challenge", "")
    solution = case_study.get("solution", "")
    results = case_study.get("results", "")
    
    y_pos = 1.3
    sections = [
        ("Challenge", challenge),
        ("Solution", solution),
        ("Results", results)
    ]
    
    for title, content in sections:
        # Section title
        tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.25))
        tb_title.text_frame.text = title
        tb_title.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["section_header"], font_adj))
        tb_title.text_frame.paragraphs[0].font.bold = True
        tb_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        
        # Content
        tb_content = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos + 0.3), Inches(9.0), Inches(0.55))
        tb_content.text_frame.text = truncate_description(content, 180)
        tb_content.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
        tb_content.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        
        y_pos += 1.05


def render_growth(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Growth Strategy slide"""
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "timeline")
    layout = layout_rec.get("layout", "two-column")
    
    # Get industry-specific growth drivers
    vertical = data.get("primaryVertical") or "technology"
    industry_content = get_industry_specific_content(vertical, "growth")
    
    add_slide_header(slide, colors, "Growth Strategy & Roadmap", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    if layout == "two-column":
        # Left: Growth drivers
        add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Key Growth Drivers", font_adj=font_adj)
        
        drivers_text = data.get("growthDrivers") or data.get("growth_drivers") or ""
        drivers = parse_lines(drivers_text, 6)
        
        y_pos = 1.5
        for driver in drivers:
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(4.1), Inches(0.4))
            tb.text_frame.text = f"• {truncate_text(driver, 50)}"
            tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
            tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
            y_pos += 0.5
        
        # Right: Goals/Milestones
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Strategic Goals", colors["secondary"], font_adj)
        
        short_goals = parse_lines(data.get("shortTermGoals") or data.get("short_term_goals") or "", 3)
        medium_goals = parse_lines(data.get("mediumTermGoals") or data.get("medium_term_goals") or "", 3)
        
        y_pos = 1.5
        if short_goals:
            tb = slide.shapes.add_textbox(Inches(5.2), Inches(y_pos), Inches(4.3), Inches(0.25))
            tb.text_frame.text = "Short-Term (0-12 months)"
            tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_large"], font_adj))
            tb.text_frame.paragraphs[0].font.bold = True
            tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
            y_pos += 0.35
            
            for goal in short_goals:
                tb = slide.shapes.add_textbox(Inches(5.2), Inches(y_pos), Inches(4.3), Inches(0.3))
                tb.text_frame.text = f"• {truncate_text(goal, 40)}"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_small"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.35


def render_market_position(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Market Position slide (Requirement #4)"""
    font_adj = layout_rec.get("font_adjustment", 0)
    chart_type = layout_rec.get("chart_type", "bar")
    layout = layout_rec.get("layout", "two-column")
    
    # Get industry benchmarks
    vertical = data.get("primaryVertical") or "technology"
    industry_content = get_industry_specific_content(vertical, "market-position")
    
    add_slide_header(slide, colors, "Market Position & Competitive Landscape", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    if layout == "two-column":
        # Left: Market overview
        add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Market Overview", font_adj=font_adj)
        
        tam = data.get("marketSize") or data.get("market_size") or "N/A"
        growth = data.get("marketGrowthRate") or data.get("market_growth_rate") or "N/A"
        
        add_metric_card(slide, colors, 0.45, 1.5, 4.2, 0.65, tam, "Total Addressable Market", font_adj)
        add_metric_card(slide, colors, 0.45, 2.35, 4.2, 0.65, f"{growth}%", "Market Growth Rate", font_adj)
        
        # Industry benchmark
        if industry_content.get("benchmarks_text"):
            tb = slide.shapes.add_textbox(Inches(0.45), Inches(3.2), Inches(4.2), Inches(0.4))
            tb.text_frame.text = industry_content["benchmarks_text"]
            tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_small"], font_adj))
            tb.text_frame.paragraphs[0].font.italic = True
            tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text_light"])
        
        # Right: Competitive advantages
        add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Competitive Advantages", colors["secondary"], font_adj)
        
        advantages_text = data.get("competitiveAdvantages") or data.get("competitive_advantages") or ""
        advantages = parse_pipe_separated(advantages_text, 5)
        
        y_pos = 1.5
        for adv in advantages:
            if adv:
                title = truncate_text(adv[0] if len(adv) > 0 else "", 35)
                tb = slide.shapes.add_textbox(Inches(5.2), Inches(y_pos), Inches(4.3), Inches(0.4))
                tb.text_frame.text = f"• {title}"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.5


def render_synergies(slide, colors, data, page_num, layout_rec, context):
    """Dedicated render function for Synergies slide (Requirement #4)"""
    font_adj = layout_rec.get("font_adjustment", 0)
    layout = layout_rec.get("layout", "two-column")
    
    # Get buyer-specific synergies
    buyer_types = data.get("targetBuyerType") or ["strategic"]
    
    add_slide_header(slide, colors, "Strategic Value & Synergies", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    if layout == "two-column":
        # Left: Strategic synergies (for strategic buyers)
        if "strategic" in buyer_types:
            add_section_box(slide, colors, 0.3, 0.95, 4.5, 3.8, "Strategic Synergies", font_adj=font_adj)
            
            syn_text = data.get("synergiesStrategic") or data.get("synergies_strategic") or ""
            synergies = parse_lines(syn_text, 6)
            
            y_pos = 1.5
            for syn in synergies:
                tb = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(4.1), Inches(0.4))
                tb.text_frame.text = f"• {truncate_text(syn, 50)}"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.5
        
        # Right: Financial synergies (for financial buyers)
        if "financial" in buyer_types:
            add_section_box(slide, colors, 5.0, 0.95, 4.7, 3.8, "Financial Synergies", colors["secondary"], font_adj)
            
            fin_text = data.get("synergiesFinancial") or data.get("synergies_financial") or ""
            fin_synergies = parse_lines(fin_text, 6)
            
            y_pos = 1.5
            for syn in fin_synergies:
                tb = slide.shapes.add_textbox(Inches(5.2), Inches(y_pos), Inches(4.3), Inches(0.4))
                tb.text_frame.text = f"• {truncate_text(syn, 50)}"
                tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(DESIGN["fonts"]["body_small"], font_adj))
                tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
                y_pos += 0.5


# ============================================================================
# TITLE & SPECIAL SLIDES
# ============================================================================

def render_title_slide(slide, colors, data, doc_config):
    """Render title slide"""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    bg.line.fill.background()
    
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.5), Inches(10), Inches(0.1))
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb(colors["accent"])
    bar.line.fill.background()
    
    company = data.get("companyName") or data.get("company_name") or "Company Name"
    codename = data.get("projectCodename") or data.get("project_codename") or "Project"
    doc_name = doc_config.get("name", "Information Memorandum")
    
    title_tb = slide.shapes.add_textbox(Inches(1), Inches(1.3), Inches(8), Inches(0.8))
    title_tb.text_frame.text = company
    title_tb.text_frame.paragraphs[0].font.size = Pt(48)
    title_tb.text_frame.paragraphs[0].font.bold = True
    title_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    title_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    subtitle_tb = slide.shapes.add_textbox(Inches(1), Inches(2.7), Inches(8), Inches(0.5))
    subtitle_tb.text_frame.text = doc_name
    subtitle_tb.text_frame.paragraphs[0].font.size = Pt(24)
    subtitle_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    subtitle_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    codename_tb = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(0.4))
    codename_tb.text_frame.text = f"Project {codename}"
    codename_tb.text_frame.paragraphs[0].font.size = Pt(18)
    codename_tb.text_frame.paragraphs[0].font.italic = True
    codename_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    codename_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    date_tb = slide.shapes.add_textbox(Inches(1), Inches(4.8), Inches(8), Inches(0.3))
    date_tb.text_frame.text = format_date(data.get("presentationDate") or data.get("presentation_date"))
    date_tb.text_frame.paragraphs[0].font.size = Pt(14)
    date_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    date_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


def render_disclaimer_slide(slide, colors, data, page_num):
    """Render disclaimer slide"""
    add_slide_header(slide, colors, "Disclaimer")
    add_slide_footer(slide, colors, page_num)
    
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8)
    
    disclaimer_text = """This presentation has been prepared solely for informational purposes. The information contained herein is confidential and proprietary. By accepting this document, you agree to maintain its confidentiality and not to reproduce, distribute, or disclose it without prior written consent.

This presentation does not constitute an offer to sell or a solicitation to buy securities. Any investment decision should be made only after thorough due diligence and consultation with professional advisors.

The financial projections and forward-looking statements contained herein are based on assumptions that may or may not prove accurate. Actual results may vary materially."""
    
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(9.0), Inches(3.2))
    tb.text_frame.text = disclaimer_text
    tb.text_frame.paragraphs[0].font.size = Pt(11)
    tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
    tb.text_frame.paragraphs[0].line_spacing = 1.3


def render_thank_you_slide(slide, colors, data, doc_config):
    """Render thank you slide"""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(5.625))
    bg.fill.solid()
    bg.fill.fore_color.rgb = hex_to_rgb(colors["primary"])
    bg.line.fill.background()
    
    thank_tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(0.8))
    thank_tb.text_frame.text = "Thank You"
    thank_tb.text_frame.paragraphs[0].font.size = Pt(48)
    thank_tb.text_frame.paragraphs[0].font.bold = True
    thank_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
    thank_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    contact_email = data.get("contactEmail") or data.get("contact_email") or ""
    contact_phone = data.get("contactPhone") or data.get("contact_phone") or ""
    
    if contact_email or contact_phone:
        contact_tb = slide.shapes.add_textbox(Inches(1), Inches(3.2), Inches(8), Inches(0.6))
        contact_text = ""
        if contact_email:
            contact_text += contact_email
        if contact_phone:
            contact_text += f"\n{contact_phone}" if contact_text else contact_phone
        
        contact_tb.text_frame.text = contact_text
        contact_tb.text_frame.paragraphs[0].font.size = Pt(18)
        contact_tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["white"])
        contact_tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# ============================================================================
# REQUIREMENT #15: UNIVERSAL createSlide() WRAPPER
# ============================================================================

def create_slide(slide_type: str, prs: Presentation, colors: dict, data: dict, page_num: int, context: dict) -> Optional[int]:
    """
    Universal slide creation wrapper - Implements Requirement #15
    
    This function:
    1. Calls analyze_data_for_layout() for AI recommendations
    2. Routes to appropriate render function based on slide_type
    3. Applies layout, font adjustments, and chart types from AI
    
    Returns: Updated page number (or None if slide not created)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    
    # Get AI layout recommendations
    layout_rec = analyze_data_for_layout_sync(data, slide_type)
    
    # Route to appropriate render function
    if slide_type == "title":
        render_title_slide(slide, colors, data, context.get("doc_config", {}))
        return None  # Title doesn't have page number
    
    elif slide_type == "disclaimer":
        render_disclaimer_slide(slide, colors, data, page_num)
        return page_num + 1
    
    elif slide_type == "executive-summary":
        render_executive_summary(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "investment-highlights":
        # Simple implementation - can be enhanced
        add_slide_header(slide, colors, "Investment Highlights")
        add_slide_footer(slide, colors, page_num)
        add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8, "Key Highlights")
        
        highlights = parse_lines(data.get("investmentHighlights") or "", 8)
        y_pos = 1.5
        for highlight in highlights:
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.35))
            tb.text_frame.text = f"• {truncate_text(highlight, 80)}"
            tb.text_frame.paragraphs[0].font.size = Pt(12)
            tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
            y_pos += 0.4
        return page_num + 1
    
    elif slide_type == "services":
        render_services(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "clients":
        render_clients(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "financials":
        render_financials(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "case-study":
        # Get case studies
        case_studies = data.get("caseStudies") or []
        if not case_studies:
            # Try legacy format
            if data.get("cs1Client"):
                case_studies.append({
                    "client": data.get("cs1Client"),
                    "challenge": data.get("cs1Challenge"),
                    "solution": data.get("cs1Solution"),
                    "results": data.get("cs1Results")
                })
        
        if case_studies:
            render_case_study(slide, colors, data, page_num, case_studies[0], layout_rec, context)
            return page_num + 1
        return None
    
    elif slide_type == "growth":
        render_growth(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "market-position":
        render_market_position(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "synergies":
        render_synergies(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "appendix-financials":
        render_appendix_financials(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "appendix-case-studies":
        render_appendix_case_studies(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "appendix-team-bios":
        render_appendix_team_bios(slide, colors, data, page_num, layout_rec, context)
        return page_num + 1
    
    elif slide_type == "thank-you":
        render_thank_you_slide(slide, colors, data, context.get("doc_config", {}))
        return None
    
    else:
        # Unknown slide type - skip
        return None


# ============================================================================
# REQUIREMENT #18: MAIN GENERATOR WITH SLIDE ITERATION
# ============================================================================

def generate_presentation(data: Dict, theme: str = "modern-blue") -> Presentation:
    """
    Generate complete presentation with robust error handling
    
    Implements Requirements #1, #18 with comprehensive validation
    """
    # ========================================
    # STEP 1: Input Validation & Normalization
    # ========================================
    
    # Handle string input (in case data comes as JSON string)
    if isinstance(data, str):
        import json
        try:
            data = json.loads(data)
        except json.JSONDecodeError as e:
            print(f"ERROR: Invalid JSON data: {e}")
            data = {}
    
    # Ensure data is a dict
    if not isinstance(data, dict):
        print(f"ERROR: Data must be dict, got {type(data)}")
        data = {}
    
    # ========================================
    # STEP 2: Required Fields Validation
    # ========================================
    
    required_fields = ["companyName", "documentType"]
    missing_fields = []
    
    for field in required_fields:
        # Check both camelCase and snake_case
        camel = field
        snake = ''.join(['_' + c.lower() if c.isupper() else c for c in field]).lstrip('_')
        
        if not data.get(camel) and not data.get(snake):
            missing_fields.append(field)
    
    if missing_fields:
        print(f"WARNING: Missing required fields: {missing_fields}")
        # Set defaults
        if "companyName" in missing_fields:
            data["companyName"] = "Company Name"
        if "documentType" in missing_fields:
            data["documentType"] = "management-presentation"
    
    # ========================================
    # STEP 3: Initialize Presentation
    # ========================================
    
    try:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)
    except Exception as e:
        print(f"ERROR: Failed to create presentation: {e}")
        raise
    
    # ========================================
    # STEP 4: Get Configuration
    # ========================================
    
    try:
        colors = get_theme_colors(theme)
    except Exception as e:
        print(f"ERROR: Invalid theme '{theme}': {e}")
        colors = get_theme_colors("modern-blue")  # Fallback
    
    # Get document type (handle both formats)
    doc_type = (
        data.get("documentType") or 
        data.get("document_type") or 
        "management-presentation"
    ).lower()
    
    # Validate document type
    valid_doc_types = ["management-presentation", "cim", "teaser"]
    if doc_type not in valid_doc_types:
        print(f"WARNING: Invalid document type '{doc_type}', using 'management-presentation'")
        doc_type = "management-presentation"
    
    doc_config = DOCUMENT_CONFIGS.get(doc_type, DOCUMENT_CONFIGS["management-presentation"])
    
    # Get industry data
    primary_vertical = (
        data.get("primaryVertical") or 
        data.get("primary_vertical") or 
        "technology"
    ).lower()
    
    industry_data = INDUSTRY_DATA.get(primary_vertical, INDUSTRY_DATA.get("technology", {}))
    
    # ========================================
    # STEP 5: Build Context
    # ========================================
    
    context = {
        "doc_config": doc_config,
        "industry_data": industry_data,
        "buyer_types": data.get("targetBuyerType") or data.get("target_buyer_type") or ["strategic"]
    }
    
    # ========================================
    # STEP 6: Determine Slides to Generate
    # ========================================
    
    try:
        slides_to_generate = get_slides_for_document_type(doc_type, data)
    except Exception as e:
        print(f"ERROR: Failed to determine slides: {e}")
        import traceback
        traceback.print_exc()
        # Fallback to basic slide list
        slides_to_generate = ["title", "disclaimer", "executive-summary", "services", "clients", "financials", "thank-you"]
    
    print(f"=== GENERATION SUMMARY ===")
    print(f"Document Type: {doc_type}")
    print(f"Industry: {primary_vertical}")
    print(f"Theme: {theme}")
    print(f"Slides to Generate ({len(slides_to_generate)}): {slides_to_generate}")
    print(f"========================")
    
    # ========================================
    # STEP 7: Generate Slides
    # ========================================
    
    page_num = 1
    slides_created = 0
    
    for slide_type in slides_to_generate:
        try:
            result = create_slide(slide_type, prs, colors, data, page_num, context)
            if result is not None:
                page_num = result
            slides_created += 1
            print(f"✓ Created slide: {slide_type}")
        except Exception as e:
            print(f"✗ ERROR creating slide '{slide_type}': {e}")
            import traceback
            traceback.print_exc()
            # Continue with next slide instead of failing completely
            continue
    
    print(f"=== GENERATION COMPLETE ===")
    print(f"Total slides created: {slides_created}/{len(slides_to_generate)}")
    print(f"=========================")
    
    return prs



# ============================================================================
# REQUIREMENT #5: APPENDIX SLIDES
# ============================================================================

def render_appendix_financials(slide, colors, data, page_num, layout_rec, context):
    """
    Render financial appendix with detailed statements
    Implements Requirement #5
    """
    font_adj = layout_rec.get("font_adjustment", -1)  # Smaller font for detailed data
    
    add_slide_header(slide, colors, "Appendix A: Detailed Financial Information", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8, "Financial Details")
    
    y_pos = 1.5
    
    # Revenue breakdown by service
    revenue_breakdown = data.get("revenueByService") or data.get("revenue_by_service") or ""
    if revenue_breakdown:
        tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.25))
        tb_title.text_frame.text = "Revenue Breakdown by Service Line"
        tb_title.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        tb_title.text_frame.paragraphs[0].font.bold = True
        tb_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        
        tb_content = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos + 0.3), Inches(9.0), Inches(0.6))
        tb_content.text_frame.text = truncate_description(revenue_breakdown, 300)
        tb_content.text_frame.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
        tb_content.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        y_pos += 1.0
    
    # Cost structure
    cost_structure = data.get("costStructure") or data.get("cost_structure") or ""
    if cost_structure:
        tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.25))
        tb_title.text_frame.text = "Cost Structure Analysis"
        tb_title.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        tb_title.text_frame.paragraphs[0].font.bold = True
        tb_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        
        tb_content = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos + 0.3), Inches(9.0), Inches(0.6))
        tb_content.text_frame.text = truncate_description(cost_structure, 300)
        tb_content.text_frame.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
        tb_content.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        y_pos += 1.0
    
    # Working capital
    working_capital = data.get("workingCapital") or data.get("working_capital") or ""
    if working_capital:
        tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.25))
        tb_title.text_frame.text = "Working Capital Requirements"
        tb_title.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        tb_title.text_frame.paragraphs[0].font.bold = True
        tb_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        
        tb_content = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos + 0.3), Inches(9.0), Inches(0.6))
        tb_content.text_frame.text = truncate_description(working_capital, 300)
        tb_content.text_frame.paragraphs[0].font.size = Pt(adjusted_font(10, font_adj))
        tb_content.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])


def render_appendix_case_studies(slide, colors, data, page_num, layout_rec, context):
    """
    Render additional case studies in appendix
    Implements Requirement #5
    """
    font_adj = layout_rec.get("font_adjustment", -1)
    
    add_slide_header(slide, colors, "Appendix B: Additional Case Studies", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8)
    
    # Get case studies beyond the first 2 (which are in main presentation)
    case_studies = data.get("caseStudies") or []
    additional_studies = case_studies[2:] if len(case_studies) > 2 else []
    
    y_pos = 1.3
    for i, study in enumerate(additional_studies[:2]):  # Max 2 per slide
        client = study.get("client", "Client")
        challenge = study.get("challenge", "")
        solution = study.get("solution", "")
        results = study.get("results", "")
        
        # Case study title
        tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.25))
        tb_title.text_frame.text = f"Case Study: {truncate_text(client, 60)}"
        tb_title.text_frame.paragraphs[0].font.size = Pt(adjusted_font(12, font_adj))
        tb_title.text_frame.paragraphs[0].font.bold = True
        tb_title.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["primary"])
        y_pos += 0.3
        
        # Compact format
        content = f"Challenge: {truncate_text(challenge, 100)} | Solution: {truncate_text(solution, 100)} | Results: {truncate_text(results, 100)}"
        tb_content = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.6))
        tb_content.text_frame.text = content
        tb_content.text_frame.paragraphs[0].font.size = Pt(adjusted_font(9, font_adj))
        tb_content.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
        y_pos += 0.9


def render_appendix_team_bios(slide, colors, data, page_num, layout_rec, context):
    """
    Render detailed team biographies in appendix
    Implements Requirement #5
    """
    font_adj = layout_rec.get("font_adjustment", -1)
    
    add_slide_header(slide, colors, "Appendix C: Detailed Team Biographies", font_adj=font_adj)
    add_slide_footer(slide, colors, page_num)
    
    add_section_box(slide, colors, 0.3, 0.95, 9.4, 3.8)
    
    # Parse leadership team
    leadership_text = data.get("leadershipTeam") or data.get("leadership_team") or ""
    team_members = parse_pipe_separated(leadership_text, 4)
    
    y_pos = 1.3
    for member in team_members:
        if member and len(member) >= 2:
            name = member[0] if len(member) > 0 else ""
            title = member[1] if len(member) > 1 else ""
            
            # Member info
            tb = slide.shapes.add_textbox(Inches(0.5), Inches(y_pos), Inches(9.0), Inches(0.4))
            tb.text_frame.text = f"• {name} - {title}"
            tb.text_frame.paragraphs[0].font.size = Pt(adjusted_font(11, font_adj))
            tb.text_frame.paragraphs[0].font.bold = True
            tb.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(colors["text"])
            y_pos += 0.5

