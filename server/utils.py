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
# V8.0: TEXT OVERFLOW PREVENTION
# ============================================================================

def calculate_text_metrics(text: str, width_inches: float, font_size: int) -> dict:
    """
    Calculate if text will fit in the given space
    Returns metrics and recommended font size
    """
    # Approximate character width in inches at given font size
    # This is empirical: 1pt = ~0.007 inches width per character
    char_width = font_size * 0.007
    
    # Calculate max characters per line
    chars_per_line = int(width_inches / char_width)
    
    # Calculate number of lines needed
    words = text.split()
    lines_needed = 1
    current_line_length = 0
    
    for word in words:
        word_length = len(word) + 1  # +1 for space
        if current_line_length + word_length > chars_per_line:
            lines_needed += 1
            current_line_length = word_length
        else:
            current_line_length += word_length
    
    # Determine if text will fit (assume max 15 lines for readability)
    will_fit = lines_needed <= 15
    
    # Calculate recommended font size if doesn't fit
    if not will_fit:
        # Scale down font to fit
        scale_factor = 15 / lines_needed
        recommended_size = max(9, int(font_size * scale_factor))
    else:
        recommended_size = font_size
    
    return {
        "chars_per_line": chars_per_line,
        "lines_needed": lines_needed,
        "will_fit": will_fit,
        "recommended_size": recommended_size,
        "max_chars": chars_per_line * 15  # For truncation
    }


def get_responsive_font_size(content_type: str, content_length: int, base_size: int = 12) -> int:
    """
    Calculate responsive font size based on content type and length
    
    Args:
        content_type: "title", "subtitle", "section_header", "body", "caption", "metric"
        content_length: number of characters in content
        base_size: base font size (optional, uses defaults per type)
    
    Returns:
        Optimal font size in points
    """
    # Base sizes per content type
    BASE_SIZES = {
        "title": 24,
        "subtitle": 14,
        "section_header": 16,
        "body": 12,
        "body_large": 14,
        "body_small": 10,
        "caption": 10,
        "metric": 32,
        "metric_medium": 24,
        "metric_small": 18
    }
    
    base = BASE_SIZES.get(content_type, base_size)
    
    # Scaling rules based on content length
    if content_type in ["body", "body_large"]:
        if content_length > 500:
            scale = 0.75
        elif content_length > 300:
            scale = 0.85
        elif content_length > 150:
            scale = 0.95
        else:
            scale = 1.0
    
    elif content_type in ["title", "section_header"]:
        if content_length > 60:
            scale = 0.75
        elif content_length > 40:
            scale = 0.85
        else:
            scale = 1.0
    
    else:
        scale = 1.0
    
    # Apply scale
    final_size = int(base * scale)
    
    # Enforce minimum readability (never below 9pt)
    return max(9, final_size)

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
    FIXED v9.0: Strict slide ordering with Thank You ALWAYS last
    Determines which slides to generate based on document type,
    appendix selections, and content variants.
    """
    from models import DOCUMENT_CONFIGS

    # STRICT SLIDE ORDER - Thank You NOT in this list (added at end)
    MAIN_SLIDE_ORDER = [
        "title",
        "disclaimer",
        "toc",
        "executive-summary",
        "investment-highlights",
        "company-overview",
        "services",
        "products",
        "tech-partnerships",
        "clients",
        "client-retention",
        "financials",
        "financial-detail",
        "case-study",
        "growth",
        "growth-goals",
        "market-position",
        "competitive-detail",
        "leadership",
        "synergies",
        "risks",
        "transaction-summary"
    ]

    # Appendix slides order
    APPENDIX_ORDER = [
        "appendix-financials",
        "appendix-case-studies",
        "appendix-team-bios"
    ]

    # Get document configuration
    config = DOCUMENT_CONFIGS.get(document_type, DOCUMENT_CONFIGS["management-presentation"])
    required_slides = set(config["required_slides"])
    optional_slides = set(config["optional_slides"])

    # Build ordered list
    slides = []

    # Add main slides in strict order
    for slide_type in MAIN_SLIDE_ORDER:
        if slide_type in required_slides:
            slides.append(slide_type)
        elif slide_type in optional_slides:
            if should_include_optional_slide(slide_type, data):
                slides.append(slide_type)

    # Add appendix slides in order, but only if selected and data exists
    include_appendix = data.get("includeAppendix", [])
    case_studies = data.get("caseStudies") or []

    if "financial-detail" in include_appendix and data.get("revenueByService"):
        slides.append("appendix-financials")

    if "case-studies-extra" in include_appendix and len(case_studies) > 2:
        slides.append("appendix-case-studies")

    if "team-bios" in include_appendix and data.get("teamBios"):
        slides.append("appendix-team-bios")

    # Handle variants
    variants = data.get("generateVariants", [])
    if "synergy" in variants and (data.get("synergiesStrategic") or data.get("synergiesFinancial")):
        if "synergies" not in slides:
            slides.append("synergies")

    if "market" in variants and data.get("competitorLandscape"):
        if "market-position" not in slides:
            slides.append("market-position")

    # CRITICAL: ALWAYS add thank-you at the END
    slides.append("thank-you")

    print(f"[v9.0] Slide order for {document_type}: {slides}")
    return slides


def should_include_optional_slide(slide_type: str, data: dict) -> bool:
    """Helper to check if optional slide should be included based on data"""
    
    inclusion_rules = {
        "toc": lambda d: d.get("documentType") == "cim",
        "leadership": lambda d: bool(d.get("founderName") or d.get("leadershipTeam")),
        "company-overview": lambda d: bool(d.get("companyDescription")),
        "case-study": lambda d: bool(d.get("caseStudies") or d.get("cs1Client")),
        "growth": lambda d: bool(d.get("growthDrivers") or d.get("shortTermGoals")),
        "growth-goals": lambda d: bool(d.get("shortTermGoals") or d.get("mediumTermGoals")),
        "synergies": lambda d: bool(d.get("synergiesStrategic") or d.get("synergiesFinancial")),
        "market-position": lambda d: bool(d.get("marketSize") or d.get("competitiveAdvantages")),
        "competitive-detail": lambda d: bool(d.get("competitiveAdvantages") or d.get("competitorLandscape")),
        "risks": lambda d: bool(d.get("businessRisks") or d.get("marketRisks") or d.get("operationalRisks")),
        "products": lambda d: bool(d.get("products")),
        "tech-partnerships": lambda d: bool(d.get("techPartnerships")),
        "client-retention": lambda d: bool(d.get("netRetention") or d.get("topClients")),
        "financial-detail": lambda d: bool(d.get("revenueFY24") or d.get("revenueFY25")),
        "investment-highlights": lambda d: bool(d.get("investmentHighlights")),
        "transaction-summary": lambda d: d.get("documentType") == "cim",
        "appendix-financials": lambda d: bool(d.get("includeFinancialAppendix")),
        "appendix-case-studies": lambda d: bool(d.get("includeAdditionalCaseStudies") and len(d.get("caseStudies", [])) > 2),
        "appendix-team-bios": lambda d: bool(d.get("includeTeamBios")),
    }
    
    if slide_type in inclusion_rules:
        return inclusion_rules[slide_type](data)
    
    return True  # Include by default


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
#======================================================================
def safe_float(value, default: float = 0.0) -> float:
    """Safely convert value to float"""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_int(value, default: int = 0) -> int:
    """Safely convert value to int"""
    if value is None:
        return default
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def extract_percentage(text: str) -> Optional[float]:
    """Extract percentage value from text (e.g., '25%' -> 25.0)"""
    if not text:
        return None
    
    import re
    match = re.search(r'(\d+(?:\.\d+)?)\s*%', str(text))
    if match:
        return float(match.group(1))
    
    try:
        return float(text)
    except (ValueError, TypeError):
        return None
