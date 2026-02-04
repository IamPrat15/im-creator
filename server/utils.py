"""
IM Creator - Utility Functions
Version: 7.2.0

Contains:
- Text utilities (truncation, condensing, formatting)
- Data parsing helpers
- Color utilities
"""

import re
from typing import List, Tuple, Optional
from datetime import datetime

# ============================================================================
# ABBREVIATIONS FOR TEXT CONDENSING
# ============================================================================
ABBREVIATIONS = {
    "and": "&",
    "with": "w/",
    "without": "w/o",
    "through": "thru",
    "information": "info",
    "technology": "tech",
    "technologies": "tech",
    "management": "mgmt",
    "development": "dev",
    "application": "app",
    "applications": "apps",
    "organization": "org",
    "international": "intl",
    "infrastructure": "infra",
    "implementation": "impl",
    "transformation": "transform",
    "approximately": "~",
    "percentage": "%",
    "percent": "%",
    "number": "#",
    "operations": "ops",
    "operational": "ops",
    "processing": "proc",
    "performance": "perf",
    "specializing": "spec.",
    "enterprise": "enterp.",
    "artificial intelligence": "AI",
    "machine learning": "ML"
}

# ============================================================================
# TEXT UTILITIES
# ============================================================================

def condense_text(text: str) -> str:
    """Apply abbreviations to condense text"""
    if not text:
        return ""
    
    result = text
    for full, abbrev in ABBREVIATIONS.items():
        pattern = re.compile(rf'\b{re.escape(full)}\b', re.IGNORECASE)
        result = pattern.sub(abbrev, result)
    
    # Normalize whitespace
    result = re.sub(r'\s+', ' ', result).strip()
    return result


def truncate_text(text: str, max_length: int, use_ellipsis: bool = True) -> str:
    """Truncate text to max length, trying to break at word boundaries"""
    if not text:
        return ""
    
    if len(text) <= max_length:
        return text
    
    # Try condensing first
    condensed = condense_text(text)
    if len(condensed) <= max_length:
        return condensed
    
    # Try to break at sentence boundary
    sentences = re.findall(r'[^.!?]+[.!?]+', condensed)
    if sentences:
        result = ""
        for sentence in sentences:
            if len(result + sentence) <= max_length:
                result += sentence
            else:
                break
        if result and len(result) >= max_length * 0.6:
            return result.strip()
    
    # Break at word boundary
    cutoff = max_length - (2 if use_ellipsis else 0)
    truncated = condensed[:cutoff]
    last_space = truncated.rfind(' ')
    
    if last_space > cutoff * 0.7:
        return truncated[:last_space].strip() + (".." if use_ellipsis else "")
    
    return truncated.strip() + (".." if use_ellipsis else "")


def truncate_description(text: str, max_length: int) -> str:
    """Truncate description, preferring complete sentences"""
    if not text:
        return ""
    
    if len(text) <= max_length:
        return text
    
    condensed = condense_text(text)
    if len(condensed) <= max_length:
        return condensed
    
    # Find last sentence boundary before max_length
    sentence_end = condensed.rfind('.', 0, max_length - 1)
    if sentence_end > max_length * 0.5:
        return condensed[:sentence_end + 1]
    
    # Fall back to word boundary
    word_end = condensed.rfind(' ', 0, max_length - 2)
    if word_end > max_length * 0.6:
        return condensed[:word_end] + ".."
    
    return condensed[:max_length - 2] + ".."


def format_currency(value: str, currency: str = "INR") -> str:
    """Format currency value"""
    if not value:
        return "N/A"
    
    try:
        num = float(value)
        if currency == "USD":
            return f"${num:.0f}M"
        else:
            return f"₹{num:.0f}Cr"
    except (ValueError, TypeError):
        return str(value)


def format_date(date_str: str = None) -> str:
    """Format date string"""
    if not date_str:
        return datetime.now().strftime("%B %Y")
    
    try:
        date = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return date.strftime("%B %Y")
    except (ValueError, TypeError):
        return str(date_str)


def parse_lines(text: str, max_lines: int = 10) -> List[str]:
    """Parse text into lines"""
    if not text:
        return []
    
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return lines[:max_lines]


def parse_pipe_separated(text: str, max_items: int = 10) -> List[List[str]]:
    """Parse pipe-separated text (e.g., 'Name|30%|Description')"""
    if not text:
        return []
    
    result = []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    for line in lines[:max_items]:
        parts = [part.strip() for part in line.split('|')]
        result.append(parts)
    
    return result


def calculate_cagr(start_value: float, end_value: float, years: int) -> Optional[int]:
    """Calculate Compound Annual Growth Rate"""
    if not start_value or not end_value or start_value <= 0 or years <= 0:
        return None
    
    try:
        cagr = (pow(end_value / start_value, 1 / years) - 1) * 100
        return round(cagr)
    except (ValueError, ZeroDivisionError):
        return None


def adjusted_font(base_size: int, adjustment: int = 0) -> int:
    """Apply font adjustment with minimum size"""
    return max(9, base_size + adjustment)


# ============================================================================
# COLOR UTILITIES
# ============================================================================

def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color to RGB tuple"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB to hex string"""
    return f"{r:02X}{g:02X}{b:02X}"


def lighten_color(hex_color: str, factor: float = 0.3) -> str:
    """Lighten a hex color"""
    r, g, b = hex_to_rgb(hex_color)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return rgb_to_hex(r, g, b)


def darken_color(hex_color: str, factor: float = 0.3) -> str:
    """Darken a hex color"""
    r, g, b = hex_to_rgb(hex_color)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return rgb_to_hex(r, g, b)


# ============================================================================
# DATA PREVIEW FOR AI
# ============================================================================

def build_data_preview(data: dict, slide_type: str) -> dict:
    """Build data preview for AI layout analysis"""
    
    # Safe string splitting
    def count_lines(text):
        if not text:
            return 0
        return len([l for l in str(text).split('\n') if l.strip()])
    
    base = {
        "has_revenue": bool(data.get("revenueFY24") or data.get("revenueFY25")),
        "revenue_years": len([v for v in [
            data.get("revenueFY24"),
            data.get("revenueFY25"),
            data.get("revenueFY26P"),
            data.get("revenueFY27P")
        ] if v]),
        "service_count": count_lines(data.get("serviceLines", "")),
        "client_count": count_lines(data.get("topClients", "")),
        "has_description": bool(data.get("companyDescription") and len(str(data.get("companyDescription", ""))) > 50),
        "description_length": len(str(data.get("companyDescription", ""))),
        "highlight_count": count_lines(data.get("investmentHighlights", "")),
        "has_margins": bool(data.get("ebitdaMarginFY25") or data.get("grossMargin")),
        "case_study_count": len(data.get("caseStudies") or []) or (1 if data.get("cs1Client") else 0) + (1 if data.get("cs2Client") else 0)
    }
    
    # Add slide-specific data
    if slide_type == "executive-summary":
        base["has_founder"] = bool(data.get("founderName"))
        base["has_employees"] = bool(data.get("employeeCountFT"))
    elif slide_type == "services":
        base["has_products"] = bool(data.get("products") and str(data.get("products", "")).strip())
    elif slide_type == "clients":
        base["has_verticals"] = bool(data.get("otherVerticals") and str(data.get("otherVerticals", "")).strip())
        base["has_partners"] = bool(data.get("techPartnerships"))
    elif slide_type == "financials":
        base["has_service_revenue"] = bool(data.get("revenueByService"))
    elif slide_type == "growth":
        base["has_drivers"] = bool(data.get("growthDrivers"))
        base["has_goals"] = bool(data.get("shortTermGoals") or data.get("mediumTermGoals"))
    
    return base


def get_default_layout_recommendation(slide_type: str, data_preview: dict) -> dict:
    """Get default layout recommendation based on slide type"""
    
    defaults = {
        "executive-summary": {
            "chart_type": "bar",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "mixed"
        },
        "investment-highlights": {
            "chart_type": "none",
            "layout": "grid-2x2",
            "font_adjustment": -1 if data_preview.get("highlight_count", 0) > 6 else 0,
            "content_density": "high" if data_preview.get("highlight_count", 0) > 6 else "medium",
            "primary_emphasis": "text"
        },
        "services": {
            "chart_type": "donut" if data_preview.get("service_count", 0) <= 4 else "pie",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "mixed"
        },
        "clients": {
            "chart_type": "donut",
            "layout": "two-column-wide-right",
            "font_adjustment": -1 if data_preview.get("client_count", 0) > 9 else 0,
            "content_density": "high" if data_preview.get("client_count", 0) > 9 else "medium",
            "primary_emphasis": "mixed"
        },
        "financials": {
            "chart_type": "bar",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "chart"
        },
        "case-study": {
            "chart_type": "none",
            "layout": "full-width",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "text"
        },
        "growth": {
            "chart_type": "timeline",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "mixed"
        },
        "market-position": {
            "chart_type": "bar",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "mixed"
        },
        "synergies": {
            "chart_type": "none",
            "layout": "two-column",
            "font_adjustment": 0,
            "content_density": "medium",
            "primary_emphasis": "text"
        }
    }
    
    return defaults.get(slide_type, {
        "chart_type": "none",
        "layout": "two-column",
        "font_adjustment": 0,
        "content_density": "medium",
        "primary_emphasis": "text"
    })


# ============================================================================
# REQUIREMENT 1: DOCUMENT TYPE HELPERS
# ============================================================================

def get_slides_for_document_type(document_type: str, data: dict) -> list:
    """
    Determine which slides to generate based on document type.
    Implements Requirement #1: Document Type differentiation
    """
    from models import DOCUMENT_CONFIGS
    
    config = DOCUMENT_CONFIGS.get(document_type, DOCUMENT_CONFIGS["management-presentation"])
    
    slides = []
    
    # Always add required slides
    slides.extend(config["required_slides"])
    
    # Add optional slides based on available data
    optional = config["optional_slides"]
    
    if "leadership" in optional and data.get("founderName"):
        slides.append("leadership")
    
    if "case-studies" in optional and (data.get("caseStudies") or data.get("cs1Client")):
        slides.append("case-studies")
    
    if "growth" in optional and (data.get("growthDrivers") or data.get("shortTermGoals")):
        slides.append("growth")
    
    if "synergies" in optional and (data.get("synergiesStrategic") or data.get("synergiesFinancial")):
        slides.append("synergies")
    
    if "market-position" in optional and (data.get("marketSize") or data.get("competitivePositioning")):
        slides.append("market-position")
    
    if "risks" in optional and data.get("riskFactors"):
        slides.append("risks")
    
    return slides


# ============================================================================
# REQUIREMENT 2: TARGET BUYER TYPE HELPERS
# ============================================================================

def get_buyer_specific_content(buyer_types: list, slide_type: str, data: dict) -> dict:
    """
    Get buyer-specific content modifications.
    Implements Requirement #2: Target Buyer Type affecting content
    """
    if not buyer_types:
        buyer_types = ["strategic"]  # Default
    
    content_mods = {
        "emphasis": [],
        "highlights": [],
        "metrics_priority": []
    }
    
    # Strategic Buyer Focus
    if "strategic" in buyer_types:
        content_mods["emphasis"].append("synergies")
        content_mods["emphasis"].append("market-expansion")
        content_mods["highlights"].append("Client relationships & retention")
        content_mods["highlights"].append("Market position & competitive moats")
        content_mods["metrics_priority"] = ["revenue_growth", "client_concentration", "market_share"]
    
    # Financial Investor Focus
    if "financial" in buyer_types:
        content_mods["emphasis"].append("returns")
        content_mods["emphasis"].append("profitability")
        content_mods["highlights"].append("EBITDA margins & cash flow")
        content_mods["highlights"].append("Growth potential & scalability")
        content_mods["metrics_priority"] = ["ebitda_margin", "revenue_growth", "profit_margin"]
    
    # International Acquirer Focus
    if "international" in buyer_types:
        content_mods["emphasis"].append("market-entry")
        content_mods["emphasis"].append("local-expertise")
        content_mods["highlights"].append("Local market knowledge & relationships")
        content_mods["highlights"].append("Regulatory compliance & certifications")
        content_mods["metrics_priority"] = ["market_size", "growth_rate", "client_diversity"]
    
    return content_mods


# ============================================================================
# REQUIREMENT 3: INDUSTRY-SPECIFIC CONTENT HELPERS
# ============================================================================

def get_industry_specific_content(vertical: str, slide_type: str) -> dict:
    """
    Get industry-specific terminology and benchmarks.
    Implements Requirement #3: Industry-specific content
    """
    from models import INDUSTRY_DATA
    
    industry_data = INDUSTRY_DATA.get(vertical, INDUSTRY_DATA["technology"])
    
    content = {
        "terminology": industry_data.get("terminology", {}),
        "benchmarks": industry_data.get("benchmarks", {}),
        "context": "",
        "emphasis": []
    }
    
    # Add slide-specific industry context
    if slide_type == "executive-summary":
        content["context"] = f"Leading player in the {industry_data['name']} sector"
        content["emphasis"] = industry_data.get("key_strengths", [])[:3]
    
    elif slide_type == "market-position":
        content["benchmarks_text"] = f"Industry average: {industry_data['benchmarks'].get('growth_rate', 'N/A')}"
        content["emphasis"] = ["market-leadership", "industry-expertise"]
    
    elif slide_type == "growth":
        content["drivers"] = industry_data.get("growth_drivers", [])
        content["emphasis"] = ["industry-trends", "market-opportunity"]
    
    return content
