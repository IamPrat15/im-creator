"""
IM Creator - Data Models and Constants
Version: 7.2.0

Contains:
- Version management
- Design constants
- 50 Professional templates
- Industry data
- Buyer content
- Document configurations
"""

from typing import Dict, List, Optional, Any
from pydantic import BaseModel
from datetime import datetime

# ============================================================================
# VERSION MANAGEMENT
# ============================================================================
VERSION = {
    "major": 7,
    "minor": 2,
    "patch": 0,
    "string": "7.2.0",
    "full": "v7.2.0",
    "build_date": "2026-02-03",
    "history": [
        {
            "version": "7.2.0",
            "date": "2026-02-03",
            "type": "major",
            "changes": [
                "Complete Python backend rewrite",
                "python-pptx for reliable chart generation",
                "FastAPI for modern async API",
                "Native pie/donut/bar charts",
                "All AI layout features preserved"
            ]
        },
        {
            "version": "7.1.1",
            "date": "2026-02-03",
            "type": "patch",
            "changes": ["Fixed chart API compatibility", "Better error handling"]
        },
        {
            "version": "7.1.0",
            "date": "2026-02-03",
            "type": "minor",
            "changes": ["Universal createSlide() wrapper", "AI recommendations applied"]
        },
        {
            "version": "7.0.0",
            "date": "2026-02-03",
            "type": "major",
            "changes": ["AI-powered layout engine", "Larger fonts", "Diverse charts"]
        }
    ]
}

# ============================================================================
# DESIGN CONSTANTS
# ============================================================================
DESIGN = {
    "slide_width": 10.0,
    "slide_height": 5.625,
    "margin": {"left": 0.3, "right": 0.3, "top": 0.2, "bottom": 0.3},
    "content_width": 9.4,
    "content_top": 0.95,
    "content_height": 4.2,
    "fonts": {
        "title": 26,
        "subtitle": 14,
        "section_header": 14,
        "body_large": 13,
        "body": 12,
        "body_small": 11,
        "caption": 10,
        "metric": 32,
        "metric_medium": 26,
        "metric_small": 20,
        "metric_label": 11,
        "chart_label": 11,
        "footer": 9
    },
    "spacing": {
        "section_gap": 0.12,
        "item_gap": 0.08,
        "box_padding": 0.1
    }
}

# ============================================================================
# 50 PROFESSIONAL TEMPLATES
# ============================================================================
PROFESSIONAL_TEMPLATES = [
    # Modern & Tech (1-10)
    {"id": "modern-blue", "name": "Modern Blue", "category": "Modern", "primary": "2B579A", "secondary": "86BC25", "accent": "E8463A"},
    {"id": "tech-gradient", "name": "Tech Gradient", "category": "Modern", "primary": "6366F1", "secondary": "8B5CF6", "accent": "06B6D4"},
    {"id": "startup-fresh", "name": "Startup Fresh", "category": "Modern", "primary": "10B981", "secondary": "3B82F6", "accent": "F59E0B"},
    {"id": "digital-dark", "name": "Digital Dark", "category": "Modern", "primary": "1F2937", "secondary": "3B82F6", "accent": "10B981"},
    {"id": "innovation-purple", "name": "Innovation Purple", "category": "Modern", "primary": "7C3AED", "secondary": "EC4899", "accent": "F59E0B"},
    {"id": "cyber-neon", "name": "Cyber Neon", "category": "Modern", "primary": "0EA5E9", "secondary": "22D3EE", "accent": "A855F7"},
    {"id": "minimal-tech", "name": "Minimal Tech", "category": "Modern", "primary": "374151", "secondary": "6B7280", "accent": "3B82F6"},
    {"id": "cloud-sky", "name": "Cloud Sky", "category": "Modern", "primary": "0284C7", "secondary": "38BDF8", "accent": "F472B6"},
    {"id": "ai-future", "name": "AI Future", "category": "Modern", "primary": "4F46E5", "secondary": "818CF8", "accent": "34D399"},
    {"id": "saas-clean", "name": "SaaS Clean", "category": "Modern", "primary": "2563EB", "secondary": "60A5FA", "accent": "FBBF24"},
    
    # Corporate & Professional (11-20)
    {"id": "corporate-navy", "name": "Corporate Navy", "category": "Corporate", "primary": "003366", "secondary": "B8860B", "accent": "C9A227"},
    {"id": "executive-gray", "name": "Executive Gray", "category": "Corporate", "primary": "1F2937", "secondary": "4B5563", "accent": "D97706"},
    {"id": "boardroom-blue", "name": "Boardroom Blue", "category": "Corporate", "primary": "1E3A5F", "secondary": "3B5998", "accent": "E5A823"},
    {"id": "professional-slate", "name": "Professional Slate", "category": "Corporate", "primary": "334155", "secondary": "64748B", "accent": "0EA5E9"},
    {"id": "classic-black", "name": "Classic Black", "category": "Corporate", "primary": "171717", "secondary": "404040", "accent": "DC2626"},
    {"id": "trust-blue", "name": "Trust Blue", "category": "Corporate", "primary": "1D4ED8", "secondary": "3B82F6", "accent": "F59E0B"},
    {"id": "heritage-brown", "name": "Heritage Brown", "category": "Corporate", "primary": "78350F", "secondary": "92400E", "accent": "FBBF24"},
    {"id": "authority-charcoal", "name": "Authority Charcoal", "category": "Corporate", "primary": "27272A", "secondary": "52525B", "accent": "16A34A"},
    {"id": "prestige-gold", "name": "Prestige Gold", "category": "Corporate", "primary": "1C1917", "secondary": "B45309", "accent": "FBBF24"},
    {"id": "institutional-blue", "name": "Institutional Blue", "category": "Corporate", "primary": "1E40AF", "secondary": "3730A3", "accent": "10B981"},
    
    # Elegant & Premium (21-30)
    {"id": "elegant-burgundy", "name": "Elegant Burgundy", "category": "Elegant", "primary": "7C1034", "secondary": "2D3748", "accent": "48BB78"},
    {"id": "luxury-black", "name": "Luxury Black", "category": "Elegant", "primary": "0A0A0A", "secondary": "262626", "accent": "D4AF37"},
    {"id": "royal-purple", "name": "Royal Purple", "category": "Elegant", "primary": "581C87", "secondary": "6B21A8", "accent": "F59E0B"},
    {"id": "champagne-gold", "name": "Champagne Gold", "category": "Elegant", "primary": "451A03", "secondary": "B8860B", "accent": "FFFBEB"},
    {"id": "emerald-elite", "name": "Emerald Elite", "category": "Elegant", "primary": "064E3B", "secondary": "065F46", "accent": "FCD34D"},
    {"id": "sapphire-class", "name": "Sapphire Class", "category": "Elegant", "primary": "1E3A8A", "secondary": "1D4ED8", "accent": "F472B6"},
    {"id": "platinum-gray", "name": "Platinum Gray", "category": "Elegant", "primary": "374151", "secondary": "9CA3AF", "accent": "A78BFA"},
    {"id": "rose-refined", "name": "Rose Refined", "category": "Elegant", "primary": "881337", "secondary": "BE185D", "accent": "FCD34D"},
    {"id": "midnight-luxe", "name": "Midnight Luxe", "category": "Elegant", "primary": "020617", "secondary": "1E293B", "accent": "C084FC"},
    {"id": "ivory-classic", "name": "Ivory Classic", "category": "Elegant", "primary": "292524", "secondary": "78716C", "accent": "B45309"},
    
    # Industry Specific (31-40)
    {"id": "finance-trust", "name": "Finance Trust", "category": "Industry", "primary": "0C4A6E", "secondary": "0369A1", "accent": "16A34A"},
    {"id": "healthcare-care", "name": "Healthcare Care", "category": "Industry", "primary": "0F766E", "secondary": "14B8A6", "accent": "3B82F6"},
    {"id": "legal-authority", "name": "Legal Authority", "category": "Industry", "primary": "1C1917", "secondary": "44403C", "accent": "B91C1C"},
    {"id": "real-estate", "name": "Real Estate Pro", "category": "Industry", "primary": "0F172A", "secondary": "334155", "accent": "EA580C"},
    {"id": "manufacturing-industrial", "name": "Industrial Strong", "category": "Industry", "primary": "1E3A8A", "secondary": "F97316", "accent": "FACC15"},
    {"id": "energy-power", "name": "Energy Power", "category": "Industry", "primary": "164E63", "secondary": "0E7490", "accent": "FDE047"},
    {"id": "consulting-sharp", "name": "Consulting Sharp", "category": "Industry", "primary": "18181B", "secondary": "3F3F46", "accent": "2563EB"},
    {"id": "pharma-health", "name": "Pharma Health", "category": "Industry", "primary": "134E4A", "secondary": "0D9488", "accent": "F97316"},
    {"id": "telecom-connect", "name": "Telecom Connect", "category": "Industry", "primary": "312E81", "secondary": "4338CA", "accent": "06B6D4"},
    {"id": "retail-vibrant", "name": "Retail Vibrant", "category": "Industry", "primary": "BE185D", "secondary": "EC4899", "accent": "FCD34D"},
    
    # Minimalist (41-45)
    {"id": "minimalist-mono", "name": "Minimalist Mono", "category": "Minimalist", "primary": "212121", "secondary": "757575", "accent": "2196F3"},
    {"id": "clean-white", "name": "Clean White", "category": "Minimalist", "primary": "1F2937", "secondary": "E5E7EB", "accent": "3B82F6"},
    {"id": "swiss-design", "name": "Swiss Design", "category": "Minimalist", "primary": "000000", "secondary": "E5E5E5", "accent": "EF4444"},
    {"id": "nordic-light", "name": "Nordic Light", "category": "Minimalist", "primary": "1E293B", "secondary": "F8FAFC", "accent": "0EA5E9"},
    {"id": "zen-simple", "name": "Zen Simple", "category": "Minimalist", "primary": "27272A", "secondary": "FAFAFA", "accent": "84CC16"},
    
    # Bold & Creative (46-50)
    {"id": "bold-impact", "name": "Bold Impact", "category": "Bold", "primary": "DC2626", "secondary": "1F2937", "accent": "FBBF24"},
    {"id": "creative-splash", "name": "Creative Splash", "category": "Bold", "primary": "DB2777", "secondary": "7C3AED", "accent": "06B6D4"},
    {"id": "vibrant-energy", "name": "Vibrant Energy", "category": "Bold", "primary": "F97316", "secondary": "EAB308", "accent": "8B5CF6"},
    {"id": "electric-blue", "name": "Electric Blue", "category": "Bold", "primary": "0066FF", "secondary": "00D4FF", "accent": "FF6B00"},
    {"id": "sunset-warm", "name": "Sunset Warm", "category": "Bold", "primary": "EA580C", "secondary": "F59E0B", "accent": "7C3AED"}
]

def get_theme_colors(theme_id: str) -> Dict[str, str]:
    """Get full theme colors from template ID"""
    template = next((t for t in PROFESSIONAL_TEMPLATES if t["id"] == theme_id), PROFESSIONAL_TEMPLATES[0])
    
    return {
        "primary": template["primary"],
        "secondary": template["secondary"],
        "accent": template["accent"],
        "text": "2D3748",
        "text_light": "718096",
        "white": "FFFFFF",
        "light_bg": "F7FAFC",
        "dark_bg": template["primary"],
        "border": "E2E8F0",
        "success": "38A169",
        "warning": "D69E2E",
        "danger": "E53E3E",
        "chart_colors": [
            template["primary"],
            template["secondary"],
            template["accent"],
            "38A169",
            "E53E3E",
            "805AD5",
            "00B5D8",
            "ED8936"
        ]
    }

# ============================================================================
# INDUSTRY DATA
# ============================================================================
INDUSTRY_DATA = {
    "bfsi": {
        "name": "BFSI",
        "full_name": "Banking, Financial Services & Insurance",
        "benchmarks": {
            "avg_growth_rate": "12-18%",
            "avg_ebitda_margin": "20-35%",
            "avg_deal_multiple": "8-12x EBITDA",
            "market_size": "$150B+ globally"
        },
        "key_metrics": ["AUM Growth", "NIM", "Cost-to-Income Ratio", "NPL Ratio", "CAR"],
        "terminology": {
            "clients": "Financial Institutions",
            "products": "Financial Technology Solutions",
            "market": "Financial Services Sector"
        },
        "key_drivers": ["Digital Banking", "RegTech Solutions", "Open Banking APIs", "AI Risk Mgmt"],
        "acquirer_interests": ["Regulatory Licenses", "Customer Base", "Technology Platform", "Compliance Infra"],
        "regulations": ["RBI Guidelines", "SEBI Compliance", "IRDAI Norms", "PCI-DSS"]
    },
    "healthcare": {
        "name": "Healthcare",
        "full_name": "Healthcare & Life Sciences",
        "benchmarks": {
            "avg_growth_rate": "15-25%",
            "avg_ebitda_margin": "15-25%",
            "avg_deal_multiple": "10-15x EBITDA",
            "market_size": "$200B+ globally"
        },
        "key_metrics": ["Patient Volume", "Bed Occupancy", "ARPOB", "Clinical Outcomes"],
        "terminology": {
            "clients": "Healthcare Providers & Payers",
            "products": "Healthcare Technology Solutions",
            "market": "Healthcare Sector"
        },
        "key_drivers": ["Telemedicine", "AI Diagnostics", "EHR Adoption", "Preventive Care"],
        "acquirer_interests": ["Patient Database", "Clinical Protocols", "Regulatory Approvals"],
        "regulations": ["HIPAA", "FDA Guidelines", "NABH Standards", "HL7/FHIR"]
    },
    "retail": {
        "name": "Retail",
        "full_name": "Retail & Consumer",
        "benchmarks": {
            "avg_growth_rate": "8-15%",
            "avg_ebitda_margin": "8-15%",
            "avg_deal_multiple": "6-10x EBITDA",
            "market_size": "$100B+ globally"
        },
        "key_metrics": ["Same-Store Sales", "Inventory Turnover", "Customer LTV", "Basket Size"],
        "terminology": {
            "clients": "Retail Brands & Chains",
            "products": "Retail Technology Solutions",
            "market": "Retail Sector"
        },
        "key_drivers": ["E-commerce", "Omnichannel", "Supply Chain", "Quick Commerce"],
        "acquirer_interests": ["Brand Portfolio", "Store Network", "Customer Database"],
        "regulations": ["Consumer Protection", "Data Privacy", "FDI Regulations"]
    },
    "manufacturing": {
        "name": "Manufacturing",
        "full_name": "Manufacturing & Industrial",
        "benchmarks": {
            "avg_growth_rate": "6-12%",
            "avg_ebitda_margin": "12-20%",
            "avg_deal_multiple": "6-9x EBITDA",
            "market_size": "$80B+ globally"
        },
        "key_metrics": ["OEE", "Capacity Utilization", "Defect Rate", "Lead Time"],
        "terminology": {
            "clients": "Industrial Enterprises",
            "products": "Industrial Technology Solutions",
            "market": "Manufacturing Sector"
        },
        "key_drivers": ["Industry 4.0", "Smart Mfg", "Sustainability", "Automation"],
        "acquirer_interests": ["Production Capacity", "IP/Patents", "Supplier Relations"],
        "regulations": ["ISO Standards", "Environmental", "Safety Standards"]
    },
    "technology": {
        "name": "Technology",
        "full_name": "Technology & Software",
        "benchmarks": {
            "avg_growth_rate": "20-40%",
            "avg_ebitda_margin": "25-40%",
            "avg_deal_multiple": "10-20x EBITDA",
            "market_size": "$500B+ globally"
        },
        "key_metrics": ["ARR", "Net Revenue Retention", "CAC Payback", "Rule of 40"],
        "terminology": {
            "clients": "Enterprise Customers",
            "products": "Technology Solutions",
            "market": "Technology Sector"
        },
        "key_drivers": ["Cloud Adoption", "AI/ML", "Cybersecurity", "Digital Transform"],
        "acquirer_interests": ["Technology IP", "Engineering Talent", "Customer Base"],
        "regulations": ["Data Privacy", "SOC 2", "ISO 27001"]
    },
    "media": {
        "name": "Media",
        "full_name": "Media, Entertainment & Digital",
        "benchmarks": {
            "avg_growth_rate": "10-20%",
            "avg_ebitda_margin": "15-30%",
            "avg_deal_multiple": "8-14x EBITDA",
            "market_size": "$120B+ globally"
        },
        "key_metrics": ["MAU/DAU", "ARPU", "Content Library Value", "Engagement Time"],
        "terminology": {
            "clients": "Media Companies & Brands",
            "products": "Content & Media Solutions",
            "market": "Media Sector"
        },
        "key_drivers": ["Streaming", "Personalization", "Ad-Tech", "Creator Economy"],
        "acquirer_interests": ["Content Library", "Audience Data", "Distribution Rights"],
        "regulations": ["Copyright Laws", "Content Regulations", "Ad Standards"]
    }
}

# ============================================================================
# BUYER CONTENT
# ============================================================================
BUYER_CONTENT = {
    "strategic": {
        "name": "Strategic Buyer",
        "focus": ["Market expansion", "Technology acquisition", "Talent access"],
        "key_messages": [
            "Complementary capabilities",
            "Established market presence",
            "Skilled workforce ready for integration"
        ],
        "financial_emphasis": ["Revenue synergies", "Cost synergies", "Market share gains"],
        "slide_adjustments": {
            "synergies": {"emphasize": "strategic"},
            "financials": {"show_projections": True}
        }
    },
    "financial": {
        "name": "Financial Investor",
        "focus": ["Growth potential", "Margin expansion", "Exit multiple"],
        "key_messages": [
            "Strong EBITDA margins",
            "Clear path to value creation",
            "Experienced management team"
        ],
        "financial_emphasis": ["EBITDA growth", "Cash conversion", "IRR potential"],
        "slide_adjustments": {
            "synergies": {"emphasize": "financial"},
            "financials": {"show_projections": True, "emphasize_ebitda": True}
        }
    },
    "international": {
        "name": "International Acquirer",
        "focus": ["Market entry", "Local expertise", "Regulatory navigation"],
        "key_messages": [
            "Local market presence",
            "Regulatory understanding",
            "Cost-effective talent base"
        ],
        "financial_emphasis": ["Currency considerations", "Transfer pricing", "Tax efficiency"],
        "slide_adjustments": {
            "synergies": {"emphasize": "international"},
            "financials": {"show_currency_notes": True}
        }
    }
}

# ============================================================================
# DOCUMENT CONFIGURATIONS
# ============================================================================
DOCUMENT_CONFIGS = {
    "management-presentation": {
        "name": "Management Presentation",
        "slide_range": "12-18 slides",
        "min_slides": 12,
        "max_slides": 18,
        "include_financial_detail": True,
        "include_sensitive_data": True,
        "include_client_names": True,
        "max_case_studies": 2,
        "required_slides": ["title", "disclaimer", "executive-summary", "investment-highlights", "services", "clients", "financials", "thank-you"],
        "optional_slides": ["leadership", "case-studies", "growth", "synergies", "market-position"]
    },
    "cim": {
        "name": "Confidential Information Memorandum",
        "slide_range": "20-35 slides",
        "min_slides": 20,
        "max_slides": 35,
        "include_financial_detail": True,
        "include_sensitive_data": True,
        "include_client_names": True,
        "max_case_studies": 5,
        "required_slides": ["title", "disclaimer", "toc", "executive-summary", "investment-highlights", "company-overview", "leadership", "industry", "services", "clients", "financials", "growth", "synergies", "risks", "thank-you"],
        "optional_slides": ["case-studies", "market-position", "team-bios", "financial-detail"]
    },
    "teaser": {
        "name": "Teaser Document",
        "slide_range": "5-8 slides",
        "min_slides": 5,
        "max_slides": 8,
        "include_financial_detail": False,
        "include_sensitive_data": False,
        "include_client_names": False,
        "max_case_studies": 0,
        "required_slides": ["title", "disclaimer", "executive-summary", "services", "thank-you"],
        "optional_slides": ["investment-highlights", "market-position"]
    }
}

# ============================================================================
# PYDANTIC MODELS FOR API
# ============================================================================
class CaseStudy(BaseModel):
    client: Optional[str] = None
    industry: Optional[str] = None
    challenge: Optional[str] = None
    solution: Optional[str] = None
    results: Optional[str] = None

class FormData(BaseModel):
    # Project Info
    projectCodename: Optional[str] = "Project Phoenix"
    documentType: Optional[str] = "management-presentation"
    presentationDate: Optional[str] = None
    advisor: Optional[str] = None
    
    # Company Info
    companyName: Optional[str] = None
    companyDescription: Optional[str] = None
    foundedYear: Optional[str] = None
    headquarters: Optional[str] = None
    employeeCountFT: Optional[str] = None
    employeeCountTotal: Optional[str] = None
    primaryVertical: Optional[str] = "technology"
    
    # Leadership
    founderName: Optional[str] = None
    founderTitle: Optional[str] = None
    founderExperience: Optional[str] = None
    founderEducation: Optional[str] = None
    leadershipTeam: Optional[str] = None
    
    # Services
    serviceLines: Optional[str] = None
    products: Optional[str] = None
    techPartnerships: Optional[str] = None
    
    # Clients
    topClients: Optional[str] = None
    topTenConcentration: Optional[str] = None
    netRetention: Optional[str] = None
    otherVerticals: Optional[str] = None
    primaryVerticalPct: Optional[str] = None
    
    # Financials
    currency: Optional[str] = "INR"
    revenueFY24: Optional[str] = None
    revenueFY25: Optional[str] = None
    revenueFY26P: Optional[str] = None
    revenueFY27P: Optional[str] = None
    revenueFY28P: Optional[str] = None
    ebitdaMarginFY25: Optional[str] = None
    grossMargin: Optional[str] = None
    netProfitMargin: Optional[str] = None
    
    # Growth
    investmentHighlights: Optional[str] = None
    growthDrivers: Optional[str] = None
    shortTermGoals: Optional[str] = None
    mediumTermGoals: Optional[str] = None
    competitiveAdvantages: Optional[str] = None
    
    # Market
    marketSize: Optional[str] = None
    marketGrowthRate: Optional[str] = None
    competitivePositioning: Optional[str] = None
    competitiveAnalysis: Optional[str] = None
    
    # Synergies
    strategicSynergies: Optional[str] = None
    financialSynergies: Optional[str] = None
    
    # Risks
    riskFactors: Optional[str] = None
    
    # Case Studies
    caseStudies: Optional[List[CaseStudy]] = None
    cs1Client: Optional[str] = None
    cs1Industry: Optional[str] = None
    cs1Challenge: Optional[str] = None
    cs1Solution: Optional[str] = None
    cs1Results: Optional[str] = None
    cs2Client: Optional[str] = None
    cs2Industry: Optional[str] = None
    cs2Challenge: Optional[str] = None
    cs2Solution: Optional[str] = None
    cs2Results: Optional[str] = None
    
    # Target Buyers
    targetBuyerTypes: Optional[List[str]] = None
    contentVariants: Optional[List[str]] = None
    
    # Contact
    contactEmail: Optional[str] = None
    contactPhone: Optional[str] = None
    
    class Config:
        extra = "allow"  # Allow additional fields

class GenerateRequest(BaseModel):
    data: FormData
    theme: Optional[str] = "modern-blue"

class LayoutRecommendation(BaseModel):
    chart_type: str = "bar"
    layout: str = "two-column"
    font_adjustment: int = 0
    content_density: str = "medium"
    primary_emphasis: str = "mixed"
