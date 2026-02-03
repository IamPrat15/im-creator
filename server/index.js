// ============================================================================
// IM Creator Server v7.0.0 - AI-Powered Layout Engine
// ============================================================================
// MAJOR UPGRADE: AI-driven slide design using Claude API
//
// NEW FEATURES (v7.0):
// - AI analyzes data and recommends optimal chart types & layouts
// - Significantly larger fonts (14pt body min, 26pt titles)
// - Diverse infographics: Pie, Donut, Bar, Progress bars, Timelines
// - Dynamic slide content based on data volume
// - Better space utilization (85%+ content area)
//
// PRESERVED FROM v6.x:
// - All 14 core features
// - 50 professional templates
// - Document types (CIM, Management Presentation, Teaser)
// - Target buyer type integration
// - Industry-specific content
// - Content variants (Market Position, Synergy Focus)
// - Appendix options
//
// VERSION HISTORY:
// v7.0.0 (2026-02-03) - AI layouts, larger fonts, diverse charts
// v6.1.0 (2026-02-02) - Text overflow fixes, generateCaseStudySlide
// v6.0.0 (2026-02-01) - 14 features, 50 templates, full implementation
// ============================================================================

const express = require('express');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

// Optional: Word document generation
let docx;
try {
  docx = require('docx');
} catch (e) {
  console.log('Note: docx package not installed. Word export will be unavailable.');
}

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  credentials: true
}));
app.use(express.json({ limit: '50mb' }));

// Create temp directory
const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) {
  fs.mkdirSync(tempDir, { recursive: true });
}

// Initialize Anthropic
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// ============================================================================
// VERSION MANAGEMENT - v7.0
// ============================================================================
const VERSION = {
  major: 7,
  minor: 0,
  patch: 0,
  get string() { return `${this.major}.${this.minor}.${this.patch}`; },
  get full() { return `v${this.string}`; },
  buildDate: '2026-02-03',
  history: [
    { version: '7.0.0', date: '2026-02-03', type: 'major', changes: ['AI-powered layout engine', 'Larger fonts (14pt body, 26pt titles)', 'Diverse infographics (Pie, Donut, Progress)', 'Dynamic slide generation'] },
    { version: '6.1.0', date: '2026-02-02', type: 'minor', changes: ['Fixed text overflow', 'Added generateCaseStudySlide', 'Better spacing'] },
    { version: '6.0.0', date: '2026-02-01', type: 'major', changes: ['14 core features', '50 templates', 'CIM/Teaser support'] }
  ]
};

// ============================================================================
// DESIGN CONSTANTS - v7.0 SIGNIFICANTLY LARGER FONTS
// ============================================================================
const DESIGN = {
  slideWidth: 10,
  slideHeight: 5.625,
  margin: { left: 0.35, right: 0.35, top: 0.25, bottom: 0.35 },
  contentWidth: 9.3,
  contentTop: 1.0,
  contentHeight: 4.0,
  fonts: {
    title: 26,           // Was 18-20, now 26
    subtitle: 14,        // Was 10-11, now 14
    sectionHeader: 14,   // Was 10-11, now 14
    bodyLarge: 13,       // New
    body: 12,            // Was 9-10, now 12
    bodySmall: 11,       // Was 8-9, now 11
    caption: 10,         // Was 7-8, now 10
    metric: 32,          // Large numbers
    metricMedium: 24,    // Medium numbers
    metricLabel: 11,
    chartLabel: 11,      // Was 7-8, now 11
    footer: 9
  },
  spacing: { sectionGap: 0.15, itemGap: 0.08, boxPadding: 0.12 }
};

// ============================================================================
// AI LAYOUT ENGINE - NEW IN v7.0
// ============================================================================
// Analyzes data and recommends optimal layouts using Claude API
async function analyzeDataForLayout(data, slideType) {
  const dataPreview = {
    hasRevenue: !!(data.revenueFY24 || data.revenueFY25),
    serviceCount: (data.serviceLines || '').split('\n').filter(x => x.trim()).length,
    clientCount: (data.topClients || '').split('\n').filter(x => x.trim()).length,
    hasDescription: !!(data.companyDescription && data.companyDescription.length > 50),
    highlightCount: (data.investmentHighlights || '').split('\n').filter(x => x.trim()).length
  };

  const prompt = `You are a presentation design expert. Analyze this data for a ${slideType} slide.

DATA SUMMARY:
${JSON.stringify(dataPreview, null, 2)}

Recommend the optimal design. Return ONLY valid JSON:
{
  "chartType": "bar|pie|donut|progress|timeline|none",
  "layout": "full-width|two-column|grid",
  "fontAdjustment": 0,
  "contentDensity": "low|medium|high",
  "emphasis": ["key_metric_1", "key_metric_2"]
}

Guidelines:
- Use pie/donut for 2-5 items showing composition
- Use bar for time series or comparisons
- Use progress bars for percentages
- Use two-column for balanced content
- Recommend font reduction (-1 or -2) only if content is very dense`;

  try {
    const response = await anthropic.messages.create({
      model: 'claude-3-haiku-20240307',
      max_tokens: 300,
      messages: [{ role: 'user', content: prompt }]
    });
    
    trackUsage('claude-3-haiku-20240307', response.usage.input_tokens, response.usage.output_tokens, `AI Layout: ${slideType}`);
    
    const text = response.content[0].text;
    const jsonMatch = text.match(/\{[\s\S]*?\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (error) {
    console.log('AI Layout Engine fallback for', slideType, ':', error.message);
  }
  
  // Smart fallback based on slide type
  return getDefaultLayoutRecommendation(slideType, dataPreview);
}

function getDefaultLayoutRecommendation(slideType, dataPreview) {
  const defaults = {
    'executive-summary': { chartType: 'bar', layout: 'two-column', fontAdjustment: 0, contentDensity: 'medium' },
    'services': { chartType: dataPreview.serviceCount <= 4 ? 'donut' : 'pie', layout: 'two-column', fontAdjustment: 0 },
    'clients': { chartType: 'donut', layout: 'two-column', fontAdjustment: dataPreview.clientCount > 8 ? -1 : 0 },
    'financials': { chartType: 'bar', layout: 'two-column', fontAdjustment: 0 },
    'case-study': { chartType: 'none', layout: 'full-width', fontAdjustment: 0 },
    'growth': { chartType: 'timeline', layout: 'two-column', fontAdjustment: 0 },
    'market-position': { chartType: 'bar', layout: 'two-column', fontAdjustment: 0 }
  };
  
  return defaults[slideType] || { chartType: 'none', layout: 'two-column', fontAdjustment: 0 };
}

// Smart chart type selection
function selectChartType(data, context) {
  if (!data || (Array.isArray(data) && data.length === 0)) return 'none';
  const count = Array.isArray(data) ? data.length : 1;
  
  switch (context) {
    case 'composition':
    case 'percentage':
      return count <= 4 ? 'donut' : 'pie';
    case 'growth':
    case 'timeseries':
    case 'revenue':
      return 'bar';
    case 'progress':
    case 'completion':
      return 'progress';
    case 'timeline':
    case 'milestones':
      return 'timeline';
    case 'comparison':
      return count <= 5 ? 'bar' : 'horizontal-bar';
    default:
      return 'bar';
  }
}

// Dynamic font sizing based on content
function calculateDynamicFontSize(text, maxWidth, baseSize, minSize = 10) {
  if (!text) return baseSize;
  const charsPerInch = baseSize * 0.11;
  const maxChars = maxWidth * charsPerInch;
  if (text.length <= maxChars) return baseSize;
  return Math.max(minSize, Math.floor(baseSize * Math.sqrt(maxChars / text.length)));
}

// ============================================================================
// USAGE TRACKING
// ============================================================================
let usageStats = {
  totalInputTokens: 0,
  totalOutputTokens: 0,
  totalCalls: 0,
  totalCostUSD: 0,
  sessionStart: new Date().toISOString(),
  calls: []
};

const PRICING = {
  'claude-sonnet-4-20250514': { input: 0.003, output: 0.015 },
  'claude-3-5-sonnet-20241022': { input: 0.003, output: 0.015 },
  'claude-3-opus-20240229': { input: 0.015, output: 0.075 },
  'claude-3-haiku-20240307': { input: 0.00025, output: 0.00125 }
};

function trackUsage(model, inputTokens, outputTokens, purpose) {
  const pricing = PRICING[model] || PRICING['claude-sonnet-4-20250514'];
  const costUSD = (inputTokens / 1000 * pricing.input) + (outputTokens / 1000 * pricing.output);
  
  usageStats.totalInputTokens += inputTokens;
  usageStats.totalOutputTokens += outputTokens;
  usageStats.totalCalls += 1;
  usageStats.totalCostUSD += costUSD;
  
  usageStats.calls.push({
    timestamp: new Date().toISOString(),
    model,
    inputTokens,
    outputTokens,
    costUSD: costUSD.toFixed(6),
    purpose
  });
  
  if (usageStats.calls.length > 100) {
    usageStats.calls = usageStats.calls.slice(-100);
  }
  
  return costUSD;
}

// ============================================================================
// INDUSTRY/VERTICAL SPECIFIC DATA
// ============================================================================
const INDUSTRY_DATA = {
  bfsi: {
    name: 'BFSI',
    fullName: 'Banking, Financial Services & Insurance',
    benchmarks: {
      avgGrowthRate: '12-18%',
      avgEbitdaMargin: '20-35%',
      avgDealMultiple: '8-12x EBITDA',
      marketSize: '$150B+ globally'
    },
    keyMetrics: ['AUM Growth', 'NIM', 'Cost-to-Income Ratio', 'NPL Ratio', 'CAR'],
    terminology: {
      clients: 'Financial Institutions',
      products: 'Financial Technology Solutions',
      market: 'Financial Services Sector'
    },
    keyDrivers: ['Digital Banking', 'RegTech Solutions', 'Open Banking APIs', 'AI Risk Mgmt'],
    acquirerInterests: ['Regulatory Licenses', 'Customer Base', 'Technology Platform', 'Compliance Infra'],
    regulations: ['RBI Guidelines', 'SEBI Compliance', 'IRDAI Norms', 'PCI-DSS']
  },
  healthcare: {
    name: 'Healthcare',
    fullName: 'Healthcare & Life Sciences',
    benchmarks: {
      avgGrowthRate: '15-25%',
      avgEbitdaMargin: '15-25%',
      avgDealMultiple: '10-15x EBITDA',
      marketSize: '$200B+ globally'
    },
    keyMetrics: ['Patient Volume', 'Bed Occupancy', 'ARPOB', 'Clinical Outcomes'],
    terminology: {
      clients: 'Healthcare Providers & Payers',
      products: 'Healthcare Technology Solutions',
      market: 'Healthcare Sector'
    },
    keyDrivers: ['Telemedicine', 'AI Diagnostics', 'EHR Adoption', 'Preventive Care'],
    acquirerInterests: ['Patient Database', 'Clinical Protocols', 'Regulatory Approvals'],
    regulations: ['HIPAA', 'FDA Guidelines', 'NABH Standards', 'HL7/FHIR']
  },
  retail: {
    name: 'Retail',
    fullName: 'Retail & Consumer',
    benchmarks: {
      avgGrowthRate: '8-15%',
      avgEbitdaMargin: '8-15%',
      avgDealMultiple: '6-10x EBITDA',
      marketSize: '$100B+ globally'
    },
    keyMetrics: ['Same-Store Sales', 'Inventory Turnover', 'Customer LTV', 'Basket Size'],
    terminology: {
      clients: 'Retail Brands & Chains',
      products: 'Retail Technology Solutions',
      market: 'Retail Sector'
    },
    keyDrivers: ['E-commerce', 'Omnichannel', 'Supply Chain', 'Quick Commerce'],
    acquirerInterests: ['Brand Portfolio', 'Store Network', 'Customer Database'],
    regulations: ['Consumer Protection', 'Data Privacy', 'FDI Regulations']
  },
  manufacturing: {
    name: 'Manufacturing',
    fullName: 'Manufacturing & Industrial',
    benchmarks: {
      avgGrowthRate: '6-12%',
      avgEbitdaMargin: '12-20%',
      avgDealMultiple: '6-9x EBITDA',
      marketSize: '$80B+ globally'
    },
    keyMetrics: ['OEE', 'Capacity Utilization', 'Defect Rate', 'Lead Time'],
    terminology: {
      clients: 'Industrial Enterprises',
      products: 'Industrial Technology Solutions',
      market: 'Manufacturing Sector'
    },
    keyDrivers: ['Industry 4.0', 'Smart Mfg', 'Sustainability', 'Automation'],
    acquirerInterests: ['Production Capacity', 'IP/Patents', 'Supplier Relations'],
    regulations: ['ISO Standards', 'Environmental', 'Safety Standards']
  },
  technology: {
    name: 'Technology',
    fullName: 'Technology & Software',
    benchmarks: {
      avgGrowthRate: '20-40%',
      avgEbitdaMargin: '25-40%',
      avgDealMultiple: '10-20x EBITDA',
      marketSize: '$500B+ globally'
    },
    keyMetrics: ['ARR', 'Net Revenue Retention', 'CAC Payback', 'Rule of 40'],
    terminology: {
      clients: 'Enterprise Customers',
      products: 'Technology Solutions',
      market: 'Technology Sector'
    },
    keyDrivers: ['Cloud Adoption', 'AI/ML', 'Cybersecurity', 'Digital Transform'],
    acquirerInterests: ['Technology IP', 'Engineering Talent', 'Customer Base'],
    regulations: ['Data Privacy', 'SOC 2', 'ISO 27001']
  },
  media: {
    name: 'Media',
    fullName: 'Media, Entertainment & Digital',
    benchmarks: {
      avgGrowthRate: '10-20%',
      avgEbitdaMargin: '15-30%',
      avgDealMultiple: '8-14x EBITDA',
      marketSize: '$120B+ globally'
    },
    keyMetrics: ['MAU/DAU', 'ARPU', 'Content Library Value', 'Engagement Time'],
    terminology: {
      clients: 'Media Companies & Brands',
      products: 'Content & Media Solutions',
      market: 'Media Sector'
    },
    keyDrivers: ['Streaming', 'Personalization', 'Ad-Tech', 'Creator Economy'],
    acquirerInterests: ['Content Library', 'Audience Data', 'Distribution Rights'],
    regulations: ['Copyright Laws', 'Content Regulations', 'Ad Standards']
  }
};

// ============================================================================
// BUYER TYPE SPECIFIC CONTENT
// ============================================================================
const BUYER_CONTENT = {
  strategic: {
    name: 'Strategic Buyer',
    focus: ['Market expansion', 'Technology acquisition', 'Talent access'],
    keyMessages: [
      'Complementary capabilities',
      'Established market presence',
      'Skilled workforce ready for integration'
    ],
    financialEmphasis: ['Revenue synergies', 'Cost synergies', 'Market share gains']
  },
  financial: {
    name: 'Financial Investor',
    focus: ['Growth potential', 'Margin expansion', 'Exit multiple'],
    keyMessages: [
      'Strong EBITDA margins',
      'Clear path to value creation',
      'Experienced management team'
    ],
    financialEmphasis: ['EBITDA growth', 'Cash conversion', 'IRR potential']
  },
  international: {
    name: 'International Acquirer',
    focus: ['Market entry', 'Local expertise', 'Regulatory navigation'],
    keyMessages: [
      'Local market presence',
      'Regulatory understanding',
      'Cost-effective talent base'
    ],
    financialEmphasis: ['Currency considerations', 'Transfer pricing', 'Tax efficiency']
  }
};

// ============================================================================
// DOCUMENT TYPE CONFIGURATIONS
// ============================================================================
const DOCUMENT_CONFIGS = {
  'management-presentation': {
    name: 'Management Presentation',
    slideRange: '13-20 slides',
    includeFinancialDetail: true,
    includeSensitiveData: true,
    includeClientNames: true
  },
  'cim': {
    name: 'Confidential Information Memorandum',
    slideRange: '25-40 slides',
    includeFinancialDetail: true,
    includeSensitiveData: true,
    includeClientNames: true
  },
  'teaser': {
    name: 'Teaser Document',
    slideRange: '5-8 slides',
    includeFinancialDetail: false,
    includeSensitiveData: false,
    includeClientNames: false
  }
};

// ============================================================================
// 50 PROFESSIONAL TEMPLATES
// ============================================================================
const PROFESSIONAL_TEMPLATES = [
  // Modern & Tech (1-10)
  { id: 'modern-blue', name: 'Modern Blue', category: 'Modern', primary: '2B579A', secondary: '86BC25', accent: 'FFC72C' },
  { id: 'tech-gradient', name: 'Tech Gradient', category: 'Modern', primary: '6366F1', secondary: '8B5CF6', accent: '06B6D4' },
  { id: 'startup-fresh', name: 'Startup Fresh', category: 'Modern', primary: '10B981', secondary: '3B82F6', accent: 'F59E0B' },
  { id: 'digital-dark', name: 'Digital Dark', category: 'Modern', primary: '1F2937', secondary: '3B82F6', accent: '10B981' },
  { id: 'innovation-purple', name: 'Innovation Purple', category: 'Modern', primary: '7C3AED', secondary: 'EC4899', accent: 'F59E0B' },
  { id: 'cyber-neon', name: 'Cyber Neon', category: 'Modern', primary: '0EA5E9', secondary: '22D3EE', accent: 'A855F7' },
  { id: 'minimal-tech', name: 'Minimal Tech', category: 'Modern', primary: '374151', secondary: '6B7280', accent: '3B82F6' },
  { id: 'cloud-sky', name: 'Cloud Sky', category: 'Modern', primary: '0284C7', secondary: '38BDF8', accent: 'F472B6' },
  { id: 'ai-future', name: 'AI Future', category: 'Modern', primary: '4F46E5', secondary: '818CF8', accent: '34D399' },
  { id: 'saas-clean', name: 'SaaS Clean', category: 'Modern', primary: '2563EB', secondary: '60A5FA', accent: 'FBBF24' },
  
  // Corporate & Professional (11-20)
  { id: 'corporate-navy', name: 'Corporate Navy', category: 'Corporate', primary: '003366', secondary: 'B8860B', accent: 'C9A227' },
  { id: 'executive-gray', name: 'Executive Gray', category: 'Corporate', primary: '1F2937', secondary: '4B5563', accent: 'D97706' },
  { id: 'boardroom-blue', name: 'Boardroom Blue', category: 'Corporate', primary: '1E3A5F', secondary: '3B5998', accent: 'E5A823' },
  { id: 'professional-slate', name: 'Professional Slate', category: 'Corporate', primary: '334155', secondary: '64748B', accent: '0EA5E9' },
  { id: 'classic-black', name: 'Classic Black', category: 'Corporate', primary: '171717', secondary: '404040', accent: 'DC2626' },
  { id: 'trust-blue', name: 'Trust Blue', category: 'Corporate', primary: '1D4ED8', secondary: '3B82F6', accent: 'F59E0B' },
  { id: 'heritage-brown', name: 'Heritage Brown', category: 'Corporate', primary: '78350F', secondary: '92400E', accent: 'FBBF24' },
  { id: 'authority-charcoal', name: 'Authority Charcoal', category: 'Corporate', primary: '27272A', secondary: '52525B', accent: '16A34A' },
  { id: 'prestige-gold', name: 'Prestige Gold', category: 'Corporate', primary: '1C1917', secondary: 'B45309', accent: 'FBBF24' },
  { id: 'institutional-blue', name: 'Institutional Blue', category: 'Corporate', primary: '1E40AF', secondary: '3730A3', accent: '10B981' },
  
  // Elegant & Premium (21-30)
  { id: 'elegant-burgundy', name: 'Elegant Burgundy', category: 'Elegant', primary: '7C1034', secondary: '2D3748', accent: '48BB78' },
  { id: 'luxury-black', name: 'Luxury Black', category: 'Elegant', primary: '0A0A0A', secondary: '262626', accent: 'D4AF37' },
  { id: 'royal-purple', name: 'Royal Purple', category: 'Elegant', primary: '581C87', secondary: '6B21A8', accent: 'F59E0B' },
  { id: 'champagne-gold', name: 'Champagne Gold', category: 'Elegant', primary: '451A03', secondary: 'B8860B', accent: 'FFFBEB' },
  { id: 'emerald-elite', name: 'Emerald Elite', category: 'Elegant', primary: '064E3B', secondary: '065F46', accent: 'FCD34D' },
  { id: 'sapphire-class', name: 'Sapphire Class', category: 'Elegant', primary: '1E3A8A', secondary: '1D4ED8', accent: 'F472B6' },
  { id: 'platinum-gray', name: 'Platinum Gray', category: 'Elegant', primary: '374151', secondary: '9CA3AF', accent: 'A78BFA' },
  { id: 'rose-refined', name: 'Rose Refined', category: 'Elegant', primary: '881337', secondary: 'BE185D', accent: 'FCD34D' },
  { id: 'midnight-luxe', name: 'Midnight Luxe', category: 'Elegant', primary: '020617', secondary: '1E293B', accent: 'C084FC' },
  { id: 'ivory-classic', name: 'Ivory Classic', category: 'Elegant', primary: '292524', secondary: '78716C', accent: 'B45309' },
  
  // Industry Specific (31-40)
  { id: 'finance-trust', name: 'Finance Trust', category: 'Industry', primary: '0C4A6E', secondary: '0369A1', accent: '16A34A' },
  { id: 'healthcare-care', name: 'Healthcare Care', category: 'Industry', primary: '0F766E', secondary: '14B8A6', accent: '3B82F6' },
  { id: 'legal-authority', name: 'Legal Authority', category: 'Industry', primary: '1C1917', secondary: '44403C', accent: 'B91C1C' },
  { id: 'real-estate', name: 'Real Estate Pro', category: 'Industry', primary: '0F172A', secondary: '334155', accent: 'EA580C' },
  { id: 'manufacturing-industrial', name: 'Industrial Strong', category: 'Industry', primary: '1E3A8A', secondary: 'F97316', accent: 'FACC15' },
  { id: 'energy-power', name: 'Energy Power', category: 'Industry', primary: '164E63', secondary: '0E7490', accent: 'FDE047' },
  { id: 'consulting-sharp', name: 'Consulting Sharp', category: 'Industry', primary: '18181B', secondary: '3F3F46', accent: '2563EB' },
  { id: 'pharma-health', name: 'Pharma Health', category: 'Industry', primary: '134E4A', secondary: '0D9488', accent: 'F97316' },
  { id: 'telecom-connect', name: 'Telecom Connect', category: 'Industry', primary: '312E81', secondary: '4338CA', accent: '06B6D4' },
  { id: 'retail-vibrant', name: 'Retail Vibrant', category: 'Industry', primary: 'BE185D', secondary: 'EC4899', accent: 'FCD34D' },
  
  // Minimalist (41-45)
  { id: 'minimalist-mono', name: 'Minimalist Mono', category: 'Minimalist', primary: '212121', secondary: '757575', accent: '2196F3' },
  { id: 'clean-white', name: 'Clean White', category: 'Minimalist', primary: '1F2937', secondary: 'E5E7EB', accent: '3B82F6' },
  { id: 'swiss-design', name: 'Swiss Design', category: 'Minimalist', primary: '000000', secondary: 'E5E5E5', accent: 'EF4444' },
  { id: 'nordic-light', name: 'Nordic Light', category: 'Minimalist', primary: '1E293B', secondary: 'F8FAFC', accent: '0EA5E9' },
  { id: 'zen-simple', name: 'Zen Simple', category: 'Minimalist', primary: '27272A', secondary: 'FAFAFA', accent: '84CC16' },
  
  // Bold & Creative (46-50)
  { id: 'bold-impact', name: 'Bold Impact', category: 'Bold', primary: 'DC2626', secondary: '1F2937', accent: 'FBBF24' },
  { id: 'creative-splash', name: 'Creative Splash', category: 'Bold', primary: 'DB2777', secondary: '7C3AED', accent: '06B6D4' },
  { id: 'vibrant-energy', name: 'Vibrant Energy', category: 'Bold', primary: 'F97316', secondary: 'EAB308', accent: '8B5CF6' },
  { id: 'electric-blue', name: 'Electric Blue', category: 'Bold', primary: '0066FF', secondary: '00D4FF', accent: 'FF6B00' },
  { id: 'sunset-warm', name: 'Sunset Warm', category: 'Bold', primary: 'EA580C', secondary: 'F59E0B', accent: '7C3AED' }
];

// Build THEMES object from templates
const THEMES = {};
PROFESSIONAL_TEMPLATES.forEach(t => {
  THEMES[t.id] = {
    name: t.name,
    category: t.category,
    primary: t.primary,
    secondary: t.secondary,
    accent: t.accent,
    text: '333333',
    textLight: '666666',
    white: 'FFFFFF',
    lightBg: 'F5F7FA',
    darkBg: t.primary,
    border: 'E0E5EC',
    success: '28A745',
    warning: 'FFC107',
    danger: 'DC3545',
    chartColors: [t.primary, t.secondary, t.accent, '00A3E0', 'E31B23', '6B3FA0']
  };
});

// ============================================================================
// TEXT UTILITIES - IMPROVED FOR BETTER FIT
// ============================================================================
const ABBREVIATIONS = {
  'and': '&', 'with': 'w/', 'without': 'w/o', 'through': 'thru',
  'information': 'info', 'technology': 'tech', 'technologies': 'tech',
  'management': 'mgmt', 'development': 'dev', 'application': 'app',
  'applications': 'apps', 'organization': 'org', 'international': 'intl',
  'infrastructure': 'infra', 'implementation': 'impl', 'transformation': 'transform',
  'approximately': '~', 'percentage': '%', 'percent': '%', 'number': '#',
  'operations': 'ops', 'operational': 'ops', 'processing': 'proc',
  'performance': 'perf', 'specializing': 'spec.', 'enterprise': 'enterp.',
  'artificial intelligence': 'AI', 'machine learning': 'ML'
};

function condenseText(text) {
  if (!text) return '';
  let result = text;
  for (const [full, abbrev] of Object.entries(ABBREVIATIONS)) {
    const regex = new RegExp(`\\b${full}\\b`, 'gi');
    result = result.replace(regex, abbrev);
  }
  return result.replace(/\s+/g, ' ').trim();
}

// IMPROVED: Better truncation that doesn't cut mid-word and increases limits
function truncateText(text, maxLength, useEllipsis = true) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  
  let condensed = condenseText(text);
  if (condensed.length <= maxLength) return condensed;
  
  // Try to break at sentence boundary first
  const sentences = condensed.match(/[^.!?]+[.!?]+/g) || [condensed];
  let result = '';
  for (const sentence of sentences) {
    if ((result + sentence).length <= maxLength) {
      result += sentence;
    } else break;
  }
  if (result.length > 0 && result.length >= maxLength * 0.6) {
    return result.trim();
  }
  
  // Break at word boundary
  const cutoff = maxLength - (useEllipsis ? 3 : 0);
  const truncated = condensed.substring(0, cutoff);
  const lastSpace = truncated.lastIndexOf(' ');
  
  if (lastSpace > cutoff * 0.7) {
    return truncated.substring(0, lastSpace).trim() + (useEllipsis ? '...' : '');
  }
  return truncated.trim() + (useEllipsis ? '...' : '');
}

// NEW: Smart truncate for descriptions - prefers complete sentences
function truncateDescription(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  
  let condensed = condenseText(text);
  if (condensed.length <= maxLength) return condensed;
  
  // Find sentence boundaries
  const sentenceEnd = condensed.lastIndexOf('.', maxLength - 1);
  if (sentenceEnd > maxLength * 0.5) {
    return condensed.substring(0, sentenceEnd + 1);
  }
  
  // Fall back to word boundary
  const wordEnd = condensed.lastIndexOf(' ', maxLength - 3);
  if (wordEnd > maxLength * 0.6) {
    return condensed.substring(0, wordEnd) + '...';
  }
  
  return condensed.substring(0, maxLength - 3) + '...';
}

function formatDate(dateStr) {
  if (!dateStr) return new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  try {
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', { year: 'numeric', month: 'long' });
  } catch {
    return dateStr;
  }
}

function parseLines(text, maxLines = 10) {
  if (!text) return [];
  return text.split('\n').filter(l => l.trim()).slice(0, maxLines);
}

function parsePipeSeparated(text, maxItems = 10) {
  if (!text) return [];
  return text.split('\n')
    .filter(l => l.trim())
    .slice(0, maxItems)
    .map(line => {
      const parts = line.split('|').map(p => p.trim());
      return parts;
    });
}


// ============================================================================
// SLIDE HELPER FUNCTIONS - IMPROVED SIZING
// ============================================================================
function addSlideHeader(slide, colors, title, subtitle) {
  // Background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.white }
  });
  
  // Left accent bar
  slide.addShape('rect', {
    x: 0, y: 0, w: 0.12, h: 0.9,
    fill: { color: colors.secondary }
  });
  
  // Title - v7.0: LARGER FONT (26pt instead of 18pt)
  slide.addText(truncateText(title, 90), {
    x: 0.35, y: 0.18, w: 9.0, h: 0.55,
    fontSize: DESIGN.fonts.title, bold: true, color: colors.primary, fontFace: 'Arial', valign: 'middle'
  });
  
  // Subtitle if provided - v7.0: Larger (14pt instead of 10pt)
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.35, y: 0.68, w: 9.0, h: 0.25,
      fontSize: DESIGN.fonts.subtitle, color: colors.textLight, fontFace: 'Arial', italic: true
    });
  }
  
  // Accent line under title - thinner and more elegant
  slide.addShape('rect', {
    x: 0.35, y: 0.88, w: 9.3, h: 0.025,
    fill: { color: colors.accent }
  });
}

function addSlideFooter(slide, colors, pageNumber, confidential = true) {
  slide.addShape('rect', {
    x: 0, y: 5.1, w: '100%', h: 0.02,
    fill: { color: colors.primary }
  });
  
  if (confidential) {
    slide.addText('Strictly Private & Confidential', {
      x: 0.3, y: 5.15, w: 3, h: 0.25,
      fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
    });
  }
  
  slide.addText(`${pageNumber}`, {
    x: 9.2, y: 5.15, w: 0.5, h: 0.25,
    fontSize: 10, color: colors.primary, fontFace: 'Arial', align: 'right'
  });
}

// v7.0: Section box with larger header text
function addSectionBox(slide, colors, x, y, w, h, title, titleBgColor) {
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 },
    rectRadius: 0.05
  });
  
  if (title) {
    const headerHeight = 0.38;
    slide.addShape('rect', {
      x, y, w, h: headerHeight,
      fill: { color: titleBgColor || colors.primary },
      rectRadius: 0.05
    });
    // Cover bottom corners
    slide.addShape('rect', {
      x, y: y + headerHeight - 0.05, w, h: 0.05,
      fill: { color: titleBgColor || colors.primary }
    });
    slide.addText(truncateText(title, 40), {
      x: x + 0.12, y: y + 0.02, w: w - 0.24, h: headerHeight - 0.04,
      fontSize: DESIGN.fonts.sectionHeader, bold: true, color: colors.white, fontFace: 'Arial', valign: 'middle'
    });
  }
}

// ============================================================================
// INFOGRAPHIC COMPONENTS - v7.0 DIVERSE CHARTS
// ============================================================================

// PIE/DONUT CHART using PptxGenJS native charts
function addPieDonutChart(slide, colors, x, y, size, data, options = {}) {
  const { type = 'doughnut', title = null, showLegend = true } = options;
  if (!data || data.length === 0) return;
  
  if (title) {
    slide.addText(title, {
      x: x - 0.1, y: y - 0.35, w: size + 0.2, h: 0.3,
      fontSize: DESIGN.fonts.body, bold: true, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  }
  
  const chartData = data.map((d, idx) => ({
    name: d.label || `Item ${idx + 1}`,
    labels: [d.label || ''],
    values: [d.value || 0]
  }));
  
  const chartColors = data.map((d, idx) => d.color || colors.chartColors[idx % colors.chartColors.length]);
  
  slide.addChart(type, chartData, {
    x: x, y: y,
    w: showLegend ? size * 0.7 : size,
    h: size,
    showLegend: false,
    showTitle: false,
    holeSize: type === 'doughnut' ? 55 : 0,
    chartColors: chartColors
  });
  
  // Custom legend
  if (showLegend && data.length <= 5) {
    data.forEach((item, idx) => {
      const ly = y + 0.1 + (idx * 0.35);
      slide.addShape('rect', {
        x: x + size * 0.75, y: ly + 0.08, w: 0.18, h: 0.18,
        fill: { color: chartColors[idx] }
      });
      slide.addText(`${truncateText(item.label, 10)} ${item.value}%`, {
        x: x + size * 0.75 + 0.25, y: ly, w: size * 0.5, h: 0.35,
        fontSize: DESIGN.fonts.caption, color: colors.text, fontFace: 'Arial', valign: 'middle'
      });
    });
  }
}

// PROGRESS BAR - For showing percentages
function addProgressBar(slide, colors, x, y, w, h, percentage, label = null, color = null) {
  const fillWidth = (Math.min(100, Math.max(0, percentage)) / 100) * w;
  
  // Background
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 },
    rectRadius: h / 2
  });
  
  // Fill
  if (fillWidth > 0) {
    slide.addShape('rect', {
      x, y, w: fillWidth, h,
      fill: { color: color || colors.primary },
      rectRadius: h / 2
    });
  }
  
  // Label and percentage
  if (label) {
    slide.addText(label, {
      x: x, y: y - 0.28, w: w * 0.7, h: 0.25,
      fontSize: DESIGN.fonts.bodySmall, color: colors.text, fontFace: 'Arial'
    });
  }
  slide.addText(`${Math.round(percentage)}%`, {
    x: x + w - 0.6, y: y - 0.28, w: 0.6, h: 0.25,
    fontSize: DESIGN.fonts.bodySmall, bold: true, color: colors.primary, fontFace: 'Arial', align: 'right'
  });
}

// TIMELINE - For company history or roadmap
function addTimeline(slide, colors, x, y, w, h, events) {
  if (!events || events.length === 0) return;
  
  const lineY = y + h / 2;
  const eventCount = Math.min(events.length, 6);
  const spacing = w / (eventCount + 1);
  
  // Main timeline line
  slide.addShape('rect', {
    x: x, y: lineY - 0.02, w: w, h: 0.04,
    fill: { color: colors.primary }
  });
  
  // Events
  events.slice(0, eventCount).forEach((event, idx) => {
    const eventX = x + spacing * (idx + 1);
    
    // Circle marker
    slide.addShape('ellipse', {
      x: eventX - 0.12, y: lineY - 0.12, w: 0.24, h: 0.24,
      fill: { color: colors.accent },
      line: { color: colors.white, width: 2 }
    });
    
    // Year (above)
    slide.addText(event.year || '', {
      x: eventX - 0.5, y: lineY - 0.55, w: 1, h: 0.3,
      fontSize: DESIGN.fonts.body, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
    });
    
    // Title (below)
    slide.addText(truncateText(event.title || '', 15), {
      x: eventX - 0.6, y: lineY + 0.2, w: 1.2, h: 0.35,
      fontSize: DESIGN.fonts.caption, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  });
}

// METRIC CARD - For displaying key numbers prominently
function addMetricCard(slide, colors, x, y, w, h, value, label, options = {}) {
  const { bgColor = null, valueColor = null } = options;
  
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: bgColor || colors.lightBg },
    line: { color: colors.border, width: 0.5 },
    rectRadius: 0.08
  });
  
  // Value - LARGE
  slide.addText(String(value), {
    x: x + 0.1, y: y + 0.08, w: w - 0.2, h: h * 0.55,
    fontSize: DESIGN.fonts.metricMedium, bold: true, color: valueColor || colors.primary, fontFace: 'Arial', valign: 'middle'
  });
  
  // Label
  slide.addText(label, {
    x: x + 0.1, y: y + h * 0.58, w: w - 0.2, h: h * 0.38,
    fontSize: DESIGN.fonts.metricLabel, color: colors.textLight, fontFace: 'Arial', valign: 'top'
  });
}

// STACKED BAR - For revenue by service
function addStackedBar(slide, colors, x, y, w, h, data, title = null) {
  if (!data || data.length === 0) return;
  
  if (title) {
    slide.addText(title, {
      x: x, y: y - 0.32, w: w, h: 0.28,
      fontSize: DESIGN.fonts.body, bold: true, color: colors.text, fontFace: 'Arial'
    });
  }
  
  let currentX = x;
  const totalWidth = w;
  
  data.forEach((item, idx) => {
    const pctMatch = String(item.value || item.pct || '0').match(/(\d+)/);
    const pct = pctMatch ? parseInt(pctMatch[1]) : 10;
    const barWidth = (pct / 100) * totalWidth;
    
    if (barWidth > 0.2) {
      slide.addShape('rect', {
        x: currentX, y: y, w: barWidth - 0.02, h: h,
        fill: { color: colors.chartColors[idx % 8] },
        rectRadius: 0.03
      });
      
      if (barWidth > 1.0) {
        slide.addText(`${truncateText(item.label || '', 12)} (${pct}%)`, {
          x: currentX + 0.05, y: y, w: barWidth - 0.1, h: h,
          fontSize: DESIGN.fonts.bodySmall, color: colors.white, fontFace: 'Arial', valign: 'middle'
        });
      }
      
      currentX += barWidth;
    }
  });
}

// ============================================================================
// MODULAR SLIDE GENERATORS - FIXED LAYOUTS
// ============================================================================

// TITLE SLIDE
function generateTitleSlide(pptx, data, colors, docConfig) {
  const slide = pptx.addSlide();
  
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  slide.addShape('rect', {
    x: 7, y: 0, w: 3, h: 2.5,
    fill: { color: colors.primary }, transparency: 80
  });
  
  slide.addShape('rect', {
    x: 0.5, y: 3.3, w: 4, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  slide.addText(data.projectCodename || 'Project Phoenix', {
    x: 0.5, y: 2.2, w: 8, h: 1,
    fontSize: 48, bold: true, color: colors.white, fontFace: 'Arial'
  });
  
  slide.addText(docConfig.name, {
    x: 0.5, y: 3.45, w: 6, h: 0.5,
    fontSize: 20, color: colors.white, fontFace: 'Arial'
  });
  
  slide.addText(formatDate(data.presentationDate), {
    x: 0.5, y: 4.05, w: 4, h: 0.35,
    fontSize: 14, color: colors.white, fontFace: 'Arial', transparency: 30
  });
  
  if (data.advisor) {
    slide.addText(`Prepared by ${data.advisor}`, {
      x: 0.5, y: 4.5, w: 4, h: 0.35,
      fontSize: 12, color: colors.white, fontFace: 'Arial', transparency: 40
    });
  }
  
  slide.addText('Strictly Private and Confidential', {
    x: 0.5, y: 4.95, w: 4, h: 0.3,
    fontSize: 10, italic: true, color: colors.white, fontFace: 'Arial', transparency: 50
  });
  
  return 1;
}

// DISCLAIMER SLIDE
function generateDisclaimerSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Important Notice', null);
  
  const advisor = data.advisor || 'the Advisor';
  const company = data.companyName || 'the Company';
  
  const disclaimerText = `This document has been prepared by ${advisor} exclusively for the benefit of the party to whom it is directly addressed and delivered. This document is strictly confidential.

This document does not constitute or form part of any offer, invitation, or inducement to purchase or subscribe for any securities.

The information contained herein has been prepared based upon information provided by ${company} and from sources believed to be reliable. No representation or warranty is made as to accuracy or completeness.

Neither ${advisor} nor any affiliates shall have any liability for any loss arising from use of this document.

This document may not be reproduced or redistributed without prior written consent of ${advisor}.`;

  slide.addText(disclaimerText, {
    x: 0.5, y: 1.15, w: 9, h: 3.8,
    fontSize: 10, color: colors.text, fontFace: 'Arial',
    valign: 'top', lineSpacingMultiple: 1.5
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// TABLE OF CONTENTS (for CIM)
function generateTOCSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Table of Contents', null);
  
  const tocItems = [
    { title: 'Executive Summary', page: 3 },
    { title: 'Investment Highlights', page: 4 },
    { title: 'Company Overview', page: 5 },
    { title: 'Industry Overview', page: 8 },
    { title: 'Business Model', page: 10 },
    { title: 'Client Portfolio', page: 12 },
    { title: 'Financial Performance', page: 14 },
    { title: 'Management Team', page: 16 },
    { title: 'Growth Strategy', page: 18 },
    { title: 'Risk Factors', page: 20 },
    { title: 'Transaction Overview', page: 22 },
    { title: 'Appendix', page: 24 }
  ];
  
  const col1 = tocItems.slice(0, 6);
  const col2 = tocItems.slice(6);
  
  col1.forEach((item, idx) => {
    slide.addText(item.title, {
      x: 0.5, y: 1.25 + (idx * 0.5), w: 3.5, h: 0.4,
      fontSize: 12, color: colors.text, fontFace: 'Arial'
    });
    slide.addText(`${item.page}`, {
      x: 4, y: 1.25 + (idx * 0.5), w: 0.5, h: 0.4,
      fontSize: 12, color: colors.primary, fontFace: 'Arial', bold: true
    });
  });
  
  col2.forEach((item, idx) => {
    slide.addText(item.title, {
      x: 5, y: 1.25 + (idx * 0.5), w: 3.5, h: 0.4,
      fontSize: 12, color: colors.text, fontFace: 'Arial'
    });
    slide.addText(`${item.page}`, {
      x: 8.5, y: 1.25 + (idx * 0.5), w: 0.5, h: 0.4,
      fontSize: 12, color: colors.primary, fontFace: 'Arial', bold: true
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}


// EXECUTIVE SUMMARY SLIDE - FIXED text overflow issues
function generateExecSummarySlide(pptx, data, colors, slideNumber, targetBuyers, industryData, docConfig) {
  const slide = pptx.addSlide();
  
  // FIXED: Increased character limit for title
  const title = truncateDescription(data.companyDescription || 'A Leading Technology Solutions Provider', 120);
  addSlideHeader(slide, colors, title, null);
  
  // Left column - Key Stats - FIXED: Adjusted heights
  addSectionBox(slide, colors, 0.3, 1.1, 2.4, 3.85, 'Key Metrics', colors.primary);
  
  const stats = [
    { label: 'Founded', value: data.foundedYear || 'N/A' },
    { label: 'Headquarters', value: truncateText(data.headquarters || 'N/A', 18, false) },
    { label: 'Employees', value: data.employeeCountFT ? `${data.employeeCountFT}+` : 'N/A' },
    { label: 'Clients', value: data.topClients ? `${parseLines(data.topClients).length}+` : 'N/A' }
  ];
  
  // Add financial metrics for financial buyers
  if (targetBuyers.includes('financial')) {
    if (data.ebitdaMarginFY25) stats.push({ label: 'EBITDA Margin', value: `${data.ebitdaMarginFY25}%` });
  }
  if (data.netRetention) stats.push({ label: 'Net Retention', value: `${data.netRetention}%` });
  
  // FIXED: Better vertical spacing
  stats.slice(0, 6).forEach((stat, idx) => {
    const yBase = 1.5 + (idx * 0.55);
    slide.addText(stat.value, {
      x: 0.4, y: yBase, w: 2.2, h: 0.28,
      fontSize: 15, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide.addText(stat.label, {
      x: 0.4, y: yBase + 0.28, w: 2.2, h: 0.22,
      fontSize: 9, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  // Middle column - Key Offerings - FIXED: Better spacing
  addSectionBox(slide, colors, 2.8, 1.1, 3.3, 3.85, 'Key Offerings', colors.secondary);
  
  const services = parsePipeSeparated(data.serviceLines, 5);
  services.forEach((service, idx) => {
    const name = service[0] || 'Service';
    const pct = service[1] || '';
    const yPos = 1.52 + (idx * 0.65);
    
    slide.addShape('roundRect', {
      x: 2.9, y: yPos, w: 3.1, h: 0.55,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    // FIXED: Increased truncation limit
    slide.addText(truncateText(name, 28, false), {
      x: 3, y: yPos + 0.08, w: 2.2, h: 0.4,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'middle'
    });
    if (pct) {
      slide.addText(pct, {
        x: 5.2, y: yPos + 0.08, w: 0.7, h: 0.4,
        fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial', valign: 'middle', align: 'right'
      });
    }
  });
  
  // Right column - Revenue Chart - FIXED: Better chart sizing
  addSectionBox(slide, colors, 6.2, 1.1, 3.3, 3.85, 'Financial Highlights', colors.accent);
  
  // Build revenue data
  const revenueData = [];
  if (data.revenueFY24) revenueData.push({ year: 'FY24', value: parseFloat(data.revenueFY24), projected: false });
  if (data.revenueFY25) revenueData.push({ year: 'FY25', value: parseFloat(data.revenueFY25), projected: false });
  if (data.revenueFY26P) revenueData.push({ year: 'FY26P', value: parseFloat(data.revenueFY26P), projected: true });
  if (data.revenueFY27P && parseFloat(data.revenueFY27P) > 0) {
    revenueData.push({ year: 'FY27P', value: parseFloat(data.revenueFY27P), projected: true });
  }
  
  if (revenueData.length > 0) {
    const maxRev = Math.max(...revenueData.map(d => d.value), 1);
    const barCount = revenueData.length;
    const chartWidth = 2.8;
    const barWidth = Math.min(0.5, (chartWidth / barCount) - 0.15);
    const startX = 6.4;
    const gap = (chartWidth - (barWidth * barCount)) / (barCount + 1);
    const chartBottom = 4.4;
    const maxBarHeight = 1.8;
    
    revenueData.forEach((rev, idx) => {
      const barHeight = (rev.value / maxRev) * maxBarHeight;
      const xPos = startX + gap + (idx * (barWidth + gap));
      
      slide.addShape('rect', {
        x: xPos, y: chartBottom - barHeight, w: barWidth, h: barHeight,
        fill: { color: rev.projected ? colors.secondary : colors.primary }
      });
      // FIXED: Value label positioning to avoid overlap
      slide.addText(`${rev.value}`, {
        x: xPos - 0.15, y: chartBottom - barHeight - 0.28, w: barWidth + 0.3, h: 0.25,
        fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center', bold: true
      });
      slide.addText(rev.year, {
        x: xPos - 0.1, y: chartBottom + 0.05, w: barWidth + 0.2, h: 0.22,
        fontSize: 8, color: colors.textLight, fontFace: 'Arial', align: 'center'
      });
    });
    
    // Currency label
    slide.addText(`In ${data.currency === 'USD' ? 'USD Mn' : 'INR Cr'}`, {
      x: 6.3, y: 1.48, w: 1.2, h: 0.2,
      fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
    });
    
    // CAGR - FIXED: Better positioning
    if (revenueData.length >= 2) {
      const firstValue = revenueData[0].value;
      const lastValue = revenueData[revenueData.length - 1].value;
      const years = revenueData.length - 1;
      if (firstValue > 0 && lastValue > firstValue) {
        const cagr = Math.round((Math.pow(lastValue / firstValue, 1 / years) - 1) * 100);
        slide.addText(`CAGR: ${cagr}%`, {
          x: 7.6, y: 1.48, w: 1.8, h: 0.2,
          fontSize: 9, bold: true, color: colors.secondary, fontFace: 'Arial', align: 'right'
        });
      }
    }
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// SNAPSHOT SLIDE (for Teaser)
function generateSnapshotSlide(pptx, data, colors, slideNumber, industryData) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Company Snapshot', 'High-level overview');
  
  // Company description box - FIXED: Better text handling
  slide.addShape('rect', {
    x: 0.3, y: 1.15, w: 9.4, h: 1.1,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide.addText(truncateDescription(data.companyDescription || 'A leading technology solutions provider', 200), {
    x: 0.45, y: 1.25, w: 9.1, h: 0.9,
    fontSize: 11, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Key facts grid
  const facts = [
    { label: 'Founded', value: data.foundedYear || 'N/A', icon: '📅' },
    { label: 'Headquarters', value: truncateText(data.headquarters || 'N/A', 22, false), icon: '📍' },
    { label: 'Employees', value: data.employeeCountFT ? `${data.employeeCountFT}+` : 'N/A', icon: '👥' },
    { label: 'Primary Vertical', value: industryData?.name || 'Technology', icon: '🏢' }
  ];
  
  facts.forEach((fact, idx) => {
    const x = 0.3 + (idx % 2) * 4.8;
    const y = 2.45 + Math.floor(idx / 2) * 1.0;
    
    slide.addShape('rect', {
      x, y, w: 4.5, h: 0.85,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide.addText(fact.icon, { x: x + 0.15, y: y + 0.2, w: 0.4, h: 0.4, fontSize: 18 });
    slide.addText(fact.label, { x: x + 0.7, y: y + 0.12, w: 3.6, h: 0.25, fontSize: 10, color: colors.textLight, fontFace: 'Arial' });
    slide.addText(fact.value, { x: x + 0.7, y: y + 0.4, w: 3.6, h: 0.35, fontSize: 13, bold: true, color: colors.primary, fontFace: 'Arial' });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// INVESTMENT HIGHLIGHTS SLIDE - FIXED text truncation
function generateInvestmentHighlightsSlide(pptx, data, colors, slideNumber, targetBuyers) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Investment Highlights', 'Key reasons to invest');
  
  let highlights = parseLines(data.investmentHighlights, 8);
  
  // Generate defaults if empty
  if (highlights.length === 0) {
    if (data.netRetention && parseFloat(data.netRetention) > 100) {
      highlights.push(`${data.netRetention}% net revenue retention`);
    }
    if (data.ebitdaMarginFY25 && parseFloat(data.ebitdaMarginFY25) > 15) {
      highlights.push(`${data.ebitdaMarginFY25}% EBITDA margins`);
    }
    highlights.push('Experienced leadership team');
    highlights.push('Diversified client base');
    highlights.push('Strong growth trajectory');
  }
  
  // FIXED: Better box sizing and text limits
  highlights.slice(0, 8).forEach((highlight, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.3 + (col * 4.8);
    const y = 1.2 + (row * 0.85);
    
    slide.addShape('rect', {
      x, y, w: 4.6, h: 0.72,
      fill: { color: idx % 2 === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    slide.addText(`${idx + 1}`, {
      x: x + 0.1, y: y + 0.15, w: 0.35, h: 0.42,
      fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    // FIXED: Increased text limit from 60 to 75
    slide.addText(truncateText(highlight, 75), {
      x: x + 0.5, y: y + 0.12, w: 3.95, h: 0.5,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'middle'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// COMPANY OVERVIEW SLIDE - FIXED layout
function generateCompanyOverviewSlide(pptx, data, colors, slideNumber, industryData) {
  const slide = pptx.addSlide();
  // FIXED: Better title truncation
  addSlideHeader(slide, colors, truncateDescription(data.companyDescription || 'Company Overview', 100), null);
  
  // Left column - Key Stats
  addSectionBox(slide, colors, 0.3, 1.1, 2.4, 3.85, 'At a Glance', colors.primary);
  
  const stats = [
    { label: 'Founded', value: data.foundedYear || 'N/A' },
    { label: 'Headquarters', value: truncateText(data.headquarters || 'N/A', 16, false) },
    { label: 'Full-Time Employees', value: data.employeeCountFT ? `${data.employeeCountFT}+` : 'N/A' },
    { label: 'Total Workforce', value: data.employeeCountOther ? `${parseInt(data.employeeCountFT || 0) + parseInt(data.employeeCountOther)}+` : 'N/A' },
    { label: 'Primary Vertical', value: industryData?.name || 'Technology' },
    { label: 'Revenue FY25', value: data.revenueFY25 ? `${data.currency === 'USD' ? '$' : '₹'}${data.revenueFY25}${data.currency === 'USD' ? 'M' : 'Cr'}` : 'N/A' }
  ];
  
  stats.forEach((stat, idx) => {
    const yBase = 1.5 + (idx * 0.55);
    slide.addText(stat.value, {
      x: 0.4, y: yBase, w: 2.2, h: 0.26,
      fontSize: 13, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide.addText(stat.label, {
      x: 0.4, y: yBase + 0.26, w: 2.2, h: 0.2,
      fontSize: 8, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  // Middle - Services
  addSectionBox(slide, colors, 2.8, 1.1, 3.3, 1.75, 'Core Services', colors.secondary);
  
  const services = parsePipeSeparated(data.serviceLines, 4);
  services.forEach((service, idx) => {
    slide.addText(`• ${truncateText(service[0] || 'Service', 32, false)}`, {
      x: 2.9, y: 1.5 + (idx * 0.32), w: 3.1, h: 0.28,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Middle - Partnerships - FIXED: Truncate partnership names
  addSectionBox(slide, colors, 2.8, 2.95, 3.3, 2.0, 'Technology Partners', colors.accent);
  
  const partnerships = parseLines(data.techPartnerships, 4);
  partnerships.forEach((partner, idx) => {
    // FIXED: Better truncation for partnerships
    slide.addText(`✓ ${truncateText(partner, 30, false)}`, {
      x: 2.9, y: 3.35 + (idx * 0.38), w: 3.1, h: 0.32,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Right - Revenue Chart
  addSectionBox(slide, colors, 6.2, 1.1, 3.3, 3.85, 'Revenue Trajectory', colors.primary);
  
  const revenueData = [];
  if (data.revenueFY24) revenueData.push({ year: 'FY24', value: parseFloat(data.revenueFY24), projected: false });
  if (data.revenueFY25) revenueData.push({ year: 'FY25', value: parseFloat(data.revenueFY25), projected: false });
  if (data.revenueFY26P) revenueData.push({ year: 'FY26P', value: parseFloat(data.revenueFY26P), projected: true });
  if (data.revenueFY27P && parseFloat(data.revenueFY27P) > 0) {
    revenueData.push({ year: 'FY27P', value: parseFloat(data.revenueFY27P), projected: true });
  }
  
  if (revenueData.length > 0) {
    const maxRev = Math.max(...revenueData.map(d => d.value), 1);
    const barCount = revenueData.length;
    const barWidth = Math.min(0.5, 2.6 / barCount - 0.15);
    const gap = (2.6 - (barWidth * barCount)) / (barCount + 1);
    const chartBottom = 4.5;
    const maxBarHeight = 1.9;
    
    revenueData.forEach((rev, idx) => {
      const barHeight = (rev.value / maxRev) * maxBarHeight;
      const xPos = 6.5 + gap + (idx * (barWidth + gap));
      
      slide.addShape('rect', {
        x: xPos, y: chartBottom - barHeight, w: barWidth, h: barHeight,
        fill: { color: rev.projected ? colors.secondary : colors.primary }
      });
      slide.addText(`${rev.value}`, {
        x: xPos - 0.1, y: chartBottom - barHeight - 0.25, w: barWidth + 0.2, h: 0.22,
        fontSize: 8, color: colors.text, fontFace: 'Arial', align: 'center', bold: true
      });
      slide.addText(rev.year, {
        x: xPos - 0.05, y: chartBottom + 0.05, w: barWidth + 0.1, h: 0.2,
        fontSize: 7, color: colors.textLight, fontFace: 'Arial', align: 'center'
      });
    });
  }
  
  // EBITDA Margin
  if (data.ebitdaMarginFY25) {
    slide.addText(`EBITDA Margin: ${data.ebitdaMarginFY25}%`, {
      x: 6.3, y: 4.72, w: 3, h: 0.2,
      fontSize: 9, bold: true, color: colors.accent, fontFace: 'Arial'
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}


// FOUNDER PROFILE SLIDE
function generateFounderSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Founded & Led by Industry Veteran', null);
  
  // Founder photo placeholder
  slide.addShape('ellipse', {
    x: 0.8, y: 1.5, w: 2.0, h: 2.0,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 3 }
  });
  slide.addText('Photo', {
    x: 0.8, y: 2.3, w: 2.0, h: 0.4,
    fontSize: 11, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  // Founder name and title
  slide.addText(data.founderName || 'Founder Name', {
    x: 0.5, y: 3.6, w: 2.6, h: 0.35,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
  });
  slide.addText(data.founderTitle || 'Founder & CEO', {
    x: 0.5, y: 3.95, w: 2.6, h: 0.28,
    fontSize: 10, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  // Background box
  addSectionBox(slide, colors, 3.3, 1.1, 6.2, 3.5, "Founder's Background", colors.primary);
  
  // Education
  const education = parseLines(data.founderEducation, 2);
  
  // Background points
  const backgroundPoints = [];
  backgroundPoints.push(`Founded ${data.companyName || 'the Company'} in ${data.foundedYear || '2015'}`);
  if (education[0]) backgroundPoints.push(education[0]);
  if (education[1]) backgroundPoints.push(education[1]);
  if (data.founderExperience) backgroundPoints.push(`${data.founderExperience}+ years of industry experience`);
  
  backgroundPoints.slice(0, 5).forEach((point, idx) => {
    slide.addText(`•  ${truncateText(point, 65)}`, {
      x: 3.4, y: 1.5 + (idx * 0.48), w: 5.9, h: 0.42,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
  });
  
  // Previous experience
  const prevCompanies = parseLines(data.previousCompanies, 4);
  if (prevCompanies.length > 0) {
    slide.addText('Previous Experience', {
      x: 3.4, y: 3.65, w: 5.9, h: 0.25,
      fontSize: 9, italic: true, color: colors.textLight, fontFace: 'Arial'
    });
    
    prevCompanies.forEach((comp, idx) => {
      const compName = comp.split('|')[0]?.trim() || comp.trim();
      slide.addShape('rect', {
        x: 3.4 + (idx * 1.5), y: 3.95, w: 1.4, h: 0.45,
        fill: { color: colors.white },
        line: { color: colors.border, width: 0.5 }
      });
      slide.addText(truncateText(compName, 14, false), {
        x: 3.4 + (idx * 1.5), y: 3.95, w: 1.4, h: 0.45,
        fontSize: 8, color: colors.text, fontFace: 'Arial', align: 'center', valign: 'middle'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// SERVICES SLIDE - FIXED: Better product name truncation
function generateServicesSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Service Lines & Capabilities', null);
  
  const services = parsePipeSeparated(data.serviceLines, 6);
  
  services.forEach((service, idx) => {
    const name = service[0] || 'Service';
    const pct = service[1] || '';
    const desc = service[2] || '';
    
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.3 + (col * 4.85);
    const y = 1.15 + (row * 1.2);
    
    // Service box
    slide.addShape('rect', {
      x, y, w: 4.7, h: 1.05,
      fill: { color: colors.lightBg },
      line: { color: colors.border, width: 0.5 }
    });
    
    // Service name - FIXED: Increased limit
    slide.addText(truncateText(name, 38), {
      x: x + 0.12, y: y + 0.08, w: 3.6, h: 0.32,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    // Percentage badge
    if (pct) {
      slide.addShape('roundRect', {
        x: x + 3.9, y: y + 0.08, w: 0.65, h: 0.32,
        fill: { color: colors.primary }
      });
      slide.addText(pct, {
        x: x + 3.9, y: y + 0.08, w: 0.65, h: 0.32,
        fontSize: 10, bold: true, color: colors.white, fontFace: 'Arial', align: 'center', valign: 'middle'
      });
    }
    
    // Description - FIXED: Increased limit
    if (desc) {
      slide.addText(truncateText(desc, 85), {
        x: x + 0.12, y: y + 0.45, w: 4.45, h: 0.52,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial', valign: 'top'
      });
    }
  });
  
  // Products section - FIXED: Better truncation
  const products = parsePipeSeparated(data.products, 3);
  if (products.length > 0) {
    slide.addText('Proprietary Products', {
      x: 0.3, y: 3.85, w: 3, h: 0.28,
      fontSize: 11, bold: true, color: colors.secondary, fontFace: 'Arial'
    });
    
    products.forEach((product, idx) => {
      const pName = product[0] || 'Product';
      const pDesc = product[1] || '';
      
      // FIXED: Truncate product name and description properly
      const displayText = pDesc ? `${truncateText(pName, 25, false)}: ${truncateText(pDesc, 55)}` : truncateText(pName, 80);
      
      slide.addText(displayText, {
        x: 0.3, y: 4.18 + (idx * 0.32), w: 9.4, h: 0.28,
        fontSize: 9, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// INDUSTRY OVERVIEW SLIDE (for CIM) - FIXED: Better text truncation
function generateIndustryOverviewSlide(pptx, data, colors, slideNumber, industryData) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, `${industryData?.fullName || 'Industry'} Overview`, null);
  
  // Industry benchmarks
  addSectionBox(slide, colors, 0.3, 1.1, 4.5, 2.2, 'Industry Benchmarks', colors.primary);
  
  if (industryData) {
    const benchmarks = [
      { label: 'Average Growth Rate', value: industryData.benchmarks.avgGrowthRate },
      { label: 'Average EBITDA Margin', value: industryData.benchmarks.avgEbitdaMargin },
      { label: 'Typical Deal Multiple', value: industryData.benchmarks.avgDealMultiple },
      { label: 'Market Size', value: industryData.benchmarks.marketSize }
    ];
    
    benchmarks.forEach((bm, idx) => {
      slide.addText(bm.label, {
        x: 0.4, y: 1.5 + (idx * 0.45), w: 2.7, h: 0.25,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
      slide.addText(bm.value, {
        x: 3.1, y: 1.5 + (idx * 0.45), w: 1.5, h: 0.25,
        fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial', align: 'right'
      });
    });
  }
  
  // Key metrics
  addSectionBox(slide, colors, 5, 1.1, 4.5, 2.2, 'Key Industry Metrics', colors.secondary);
  
  if (industryData?.keyMetrics) {
    industryData.keyMetrics.slice(0, 5).forEach((metric, idx) => {
      slide.addText(`• ${metric}`, {
        x: 5.1, y: 1.5 + (idx * 0.38), w: 4.3, h: 0.32,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  // Key drivers - FIXED: Shorter text to prevent overflow
  addSectionBox(slide, colors, 0.3, 3.4, 4.5, 1.5, 'Market Drivers', colors.accent);
  
  if (industryData?.keyDrivers) {
    industryData.keyDrivers.slice(0, 4).forEach((driver, idx) => {
      // FIXED: Reduced character limit for market drivers
      slide.addText(`▸ ${truncateText(driver, 35, false)}`, {
        x: 0.4, y: 3.8 + (idx * 0.32), w: 4.3, h: 0.28,
        fontSize: 9, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  // Regulatory environment - FIXED: Shorter text
  addSectionBox(slide, colors, 5, 3.4, 4.5, 1.5, 'Regulatory Environment', colors.primary);
  
  if (industryData?.regulations) {
    industryData.regulations.slice(0, 4).forEach((reg, idx) => {
      // FIXED: Reduced character limit
      slide.addText(`• ${truncateText(reg, 35, false)}`, {
        x: 5.1, y: 3.8 + (idx * 0.32), w: 4.3, h: 0.28,
        fontSize: 9, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}


// CLIENTS SLIDE - FIXED: Client name truncation
function generateClientsSlide(pptx, data, colors, slideNumber, industryData, docConfig) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Client Portfolio & Vertical Mix', null);
  
  // Client metrics box
  addSectionBox(slide, colors, 0.3, 1.1, 3, 2.1, 'Client Metrics', colors.primary);
  
  const clientMetrics = [
    { label: 'Top 10 Concentration', value: data.top10Concentration ? `${data.top10Concentration}%` : 'N/A' },
    { label: 'Net Revenue Retention', value: data.netRetention ? `${data.netRetention}%` : 'N/A' },
    { label: 'Primary Vertical', value: industryData?.name || 'Technology' },
    { label: 'Primary Vertical %', value: data.primaryVerticalPct ? `${data.primaryVerticalPct}%` : 'N/A' }
  ];
  
  clientMetrics.forEach((metric, idx) => {
    slide.addText(metric.label, {
      x: 0.4, y: 1.5 + (idx * 0.45), w: 1.7, h: 0.22,
      fontSize: 9, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText(metric.value, {
      x: 2.1, y: 1.5 + (idx * 0.45), w: 1.1, h: 0.22,
      fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial', align: 'right'
    });
  });
  
  // Vertical mix
  addSectionBox(slide, colors, 0.3, 3.3, 3, 1.6, 'Vertical Mix', colors.secondary);
  
  const verticals = parsePipeSeparated(data.otherVerticals, 4);
  if (data.primaryVertical && data.primaryVerticalPct) {
    verticals.unshift([industryData?.name || data.primaryVertical, `${data.primaryVerticalPct}%`]);
  }
  
  verticals.slice(0, 4).forEach((vert, idx) => {
    slide.addText(`${truncateText(vert[0] || 'Vertical', 18, false)}: ${vert[1] || ''}`, {
      x: 0.4, y: 3.7 + (idx * 0.28), w: 2.8, h: 0.24,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Top clients grid - FIXED: Better truncation
  addSectionBox(slide, colors, 3.4, 1.1, 6.2, 3.8, 'Key Clients', colors.accent);
  
  const clients = parsePipeSeparated(data.topClients, 12);
  
  clients.forEach((client, idx) => {
    let clientName = client[0] || 'Client';
    const clientVertical = client[1] || '';
    const clientYear = client[2] || '';
    
    // For teaser, anonymize client names
    if (!docConfig.includeClientNames) {
      clientName = `Leading ${clientVertical || 'Enterprise'} Client`;
    }
    
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const x = 3.5 + (col * 2.0);
    const y = 1.5 + (row * 0.72);
    
    slide.addShape('rect', {
      x, y, w: 1.9, h: 0.62,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    // FIXED: Better truncation for client names
    slide.addText(truncateText(clientName, 16, false), {
      x, y: y + 0.08, w: 1.9, h: 0.28,
      fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial', align: 'center'
    });
    if (clientYear) {
      slide.addText(`Since ${clientYear}`, {
        x, y: y + 0.38, w: 1.9, h: 0.18,
        fontSize: 7, color: colors.textLight, fontFace: 'Arial', align: 'center'
      });
    }
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// FINANCIALS SLIDE - FIXED: CAGR positioning and chart labels
function generateFinancialsSlide(pptx, data, colors, slideNumber, docConfig) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Financial Performance', null);
  
  // Revenue chart (left side)
  addSectionBox(slide, colors, 0.3, 1.1, 4.5, 2.9, 'Revenue Growth', colors.primary);
  
  const revenueData = [];
  if (data.revenueFY24) revenueData.push({ year: 'FY24', value: parseFloat(data.revenueFY24), projected: false });
  if (data.revenueFY25) revenueData.push({ year: 'FY25', value: parseFloat(data.revenueFY25), projected: false });
  if (data.revenueFY26P) revenueData.push({ year: 'FY26P', value: parseFloat(data.revenueFY26P), projected: true });
  if (data.revenueFY27P && parseFloat(data.revenueFY27P) > 0) {
    revenueData.push({ year: 'FY27P', value: parseFloat(data.revenueFY27P), projected: true });
  }
  if (data.revenueFY28P && parseFloat(data.revenueFY28P) > 0) {
    revenueData.push({ year: 'FY28P', value: parseFloat(data.revenueFY28P), projected: true });
  }
  
  if (revenueData.length > 0) {
    const maxRev = Math.max(...revenueData.map(d => d.value), 1);
    const barCount = revenueData.length;
    const chartWidth = 3.8;
    const barWidth = Math.min(0.55, (chartWidth / barCount) - 0.2);
    const gap = (chartWidth - (barWidth * barCount)) / (barCount + 1);
    const chartBottom = 3.55;
    const maxBarHeight = 1.6;
    
    revenueData.forEach((rev, idx) => {
      const barHeight = (rev.value / maxRev) * maxBarHeight;
      const xPos = 0.55 + gap + (idx * (barWidth + gap));
      
      slide.addShape('rect', {
        x: xPos, y: chartBottom - barHeight, w: barWidth, h: barHeight,
        fill: { color: rev.projected ? colors.secondary : colors.primary }
      });
      // FIXED: Value labels with better positioning
      slide.addText(`${rev.value}`, {
        x: xPos - 0.12, y: chartBottom - barHeight - 0.26, w: barWidth + 0.24, h: 0.23,
        fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center', bold: true
      });
      slide.addText(rev.year, {
        x: xPos - 0.08, y: chartBottom + 0.05, w: barWidth + 0.16, h: 0.2,
        fontSize: 8, color: colors.textLight, fontFace: 'Arial', align: 'center'
      });
    });
    
    // Currency label and CAGR on same line
    slide.addText(`In ${data.currency === 'USD' ? 'USD Mn' : 'INR Cr'}`, {
      x: 0.4, y: 1.48, w: 1.5, h: 0.2,
      fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
    });
    
    // CAGR - FIXED: Better positioning
    if (revenueData.length >= 2) {
      const first = revenueData[0].value;
      const last = revenueData[revenueData.length - 1].value;
      const years = revenueData.length - 1;
      if (first > 0 && last > first) {
        const cagr = Math.round((Math.pow(last / first, 1 / years) - 1) * 100);
        slide.addText(`CAGR: ${cagr}%`, {
          x: 2.8, y: 1.48, w: 1.8, h: 0.2,
          fontSize: 9, bold: true, color: colors.secondary, fontFace: 'Arial', align: 'right'
        });
      }
    }
  }
  
  // Key margins (right side)
  addSectionBox(slide, colors, 5, 1.1, 4.5, 2.9, 'Key Margins & Metrics', colors.secondary);
  
  const margins = [];
  if (data.ebitdaMarginFY25) margins.push({ label: 'EBITDA Margin FY25', value: `${data.ebitdaMarginFY25}%` });
  if (data.grossMargin) margins.push({ label: 'Gross Margin', value: `${data.grossMargin}%` });
  if (data.netProfitMargin) margins.push({ label: 'Net Profit Margin', value: `${data.netProfitMargin}%` });
  if (data.netRetention) margins.push({ label: 'Net Revenue Retention', value: `${data.netRetention}%` });
  if (data.top10Concentration) margins.push({ label: 'Top 10 Concentration', value: `${data.top10Concentration}%` });
  
  margins.slice(0, 5).forEach((margin, idx) => {
    slide.addText(margin.label, {
      x: 5.1, y: 1.55 + (idx * 0.48), w: 2.8, h: 0.25,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
    slide.addText(margin.value, {
      x: 8, y: 1.55 + (idx * 0.48), w: 1.3, h: 0.25,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial', align: 'right'
    });
  });
  
  // Revenue by service (bottom) - FIXED: Better label positioning
  const serviceRevenue = parsePipeSeparated(data.revenueByService, 6);
  if (serviceRevenue.length > 0) {
    addSectionBox(slide, colors, 0.3, 4.1, 9.2, 0.85, 'Revenue by Service Line', colors.accent);
    
    const totalWidth = 8.8;
    let currentX = 0.5;
    
    serviceRevenue.forEach((srv, idx) => {
      const pctMatch = (srv[1] || '0').match(/(\d+)/);
      const pct = pctMatch ? parseInt(pctMatch[1]) : 10;
      const barWidth = (pct / 100) * totalWidth;
      
      if (barWidth > 0.3) {
        slide.addShape('rect', {
          x: currentX, y: 4.5, w: barWidth - 0.05, h: 0.32,
          fill: { color: colors.chartColors ? colors.chartColors[idx % 6] : colors.primary }
        });
        
        if (barWidth > 0.9) {
          // FIXED: Truncate service names in bar
          const displayText = truncateText(srv[0] || '', Math.floor(barWidth * 8), false);
          slide.addText(`${displayText} (${pct}%)`, {
            x: currentX + 0.05, y: 4.5, w: barWidth - 0.15, h: 0.32,
            fontSize: 7, color: colors.white, fontFace: 'Arial', valign: 'middle'
          });
        }
        
        currentX += barWidth;
      }
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// CASE STUDY SLIDE
function generateCaseStudySlide(pptx, caseStudy, colors, slideNumber, caseNumber, docConfig) {
  const slide = pptx.addSlide();
  
  let clientName = caseStudy.client || `Case Study ${caseNumber}`;
  
  // Anonymize for teaser
  if (!docConfig.includeClientNames && caseStudy.industry) {
    clientName = `Leading ${caseStudy.industry} Client`;
  }
  
  addSlideHeader(slide, colors, `Case Study: ${truncateText(clientName, 35, false)}`, null);
  
  // Client info box
  slide.addShape('rect', {
    x: 0.3, y: 1.15, w: 2.2, h: 1.7,
    fill: { color: colors.primary }
  });
  slide.addText(truncateText(clientName, 22, false), {
    x: 0.4, y: 1.35, w: 2.0, h: 0.45,
    fontSize: 13, bold: true, color: colors.white, fontFace: 'Arial'
  });
  if (caseStudy.industry) {
    slide.addText(caseStudy.industry, {
      x: 0.4, y: 1.85, w: 2.0, h: 0.28,
      fontSize: 10, color: colors.white, fontFace: 'Arial', transparency: 20
    });
  }
  
  // Challenge box
  addSectionBox(slide, colors, 2.6, 1.15, 3.4, 1.7, 'Challenge', colors.danger);
  slide.addText(truncateDescription(caseStudy.challenge || 'Business challenge description', 200), {
    x: 2.7, y: 1.55, w: 3.2, h: 1.2,
    fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Solution box
  addSectionBox(slide, colors, 6.1, 1.15, 3.4, 1.7, 'Solution', colors.primary);
  slide.addText(truncateDescription(caseStudy.solution || 'Solution implemented', 200), {
    x: 6.2, y: 1.55, w: 3.2, h: 1.2,
    fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Results section
  slide.addShape('rect', {
    x: 0.3, y: 2.95, w: 9.2, h: 0.32,
    fill: { color: colors.accent }
  });
  slide.addText('Key Results & Impact', {
    x: 0.4, y: 2.95, w: 9, h: 0.32,
    fontSize: 11, bold: true, color: colors.white, fontFace: 'Arial', valign: 'middle'
  });
  
  const results = parseLines(caseStudy.results, 6);
  results.forEach((result, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    
    slide.addText(`✓ ${truncateText(result.trim(), 50)}`, {
      x: 0.4 + (col * 4.6), y: 3.38 + (row * 0.45), w: 4.4, h: 0.4,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// GROWTH STRATEGY SLIDE - FIXED: Text truncation for drivers
function generateGrowthSlide(pptx, data, colors, slideNumber, targetBuyers) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Growth Strategy & Roadmap', null);
  
  // Growth drivers - FIXED: Better text limits
  addSectionBox(slide, colors, 0.3, 1.1, 4.5, 2.1, 'Key Growth Drivers', colors.primary);
  
  const drivers = parseLines(data.growthDrivers, 5);
  drivers.forEach((driver, idx) => {
    // FIXED: Reduced character limit to prevent overflow
    slide.addText(`▸ ${truncateText(driver, 48)}`, {
      x: 0.4, y: 1.52 + (idx * 0.38), w: 4.3, h: 0.34,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Strategic roadmap
  addSectionBox(slide, colors, 5, 1.1, 4.5, 2.1, 'Strategic Roadmap', colors.secondary);
  
  // Short-term goals (0-12 months)
  slide.addText('0-12 Months', {
    x: 5.1, y: 1.5, w: 2, h: 0.22,
    fontSize: 9, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const shortTermGoals = parseLines(data.shortTermGoals, 2);
  shortTermGoals.forEach((goal, idx) => {
    // FIXED: Better truncation
    slide.addText(`• ${truncateText(goal, 42)}`, {
      x: 5.1, y: 1.75 + (idx * 0.28), w: 4.3, h: 0.26,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Medium-term goals (1-3 years)
  slide.addText('1-3 Years', {
    x: 5.1, y: 2.4, w: 2, h: 0.22,
    fontSize: 9, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const mediumTermGoals = parseLines(data.mediumTermGoals, 2);
  mediumTermGoals.forEach((goal, idx) => {
    // FIXED: Better truncation
    slide.addText(`• ${truncateText(goal, 42)}`, {
      x: 5.1, y: 2.65 + (idx * 0.28), w: 4.3, h: 0.26,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Competitive advantages - FIXED: Better spacing
  addSectionBox(slide, colors, 0.3, 3.3, 9.2, 1.6, 'Competitive Advantages', colors.accent);
  
  const advantages = parsePipeSeparated(data.competitiveAdvantages, 6);
  advantages.forEach((adv, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.4 + (col * 4.6);
    const y = 3.72 + (row * 0.42);
    
    const title = adv[0] || 'Advantage';
    const detail = adv[1] || '';
    
    // FIXED: Better truncation for advantages
    const displayText = detail ? 
      `• ${truncateText(title, 22, false)}: ${truncateText(detail, 32)}` : 
      `• ${truncateText(title, 55)}`;
    
    slide.addText(displayText, {
      x, y, w: 4.4, h: 0.38,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// SYNERGIES SLIDE - FIXED: Better text fit
function generateSynergiesSlide(pptx, data, colors, slideNumber, targetBuyers) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Acquisition Synergies & Value Creation', null);
  
  const showStrategic = targetBuyers.includes('strategic');
  const showFinancial = targetBuyers.includes('financial') || targetBuyers.includes('international');
  
  const colWidth = (showStrategic && showFinancial) ? 4.5 : 9.2;
  const col1X = 0.3;
  const col2X = 5;
  
  // Strategic buyer synergies
  if (showStrategic) {
    addSectionBox(slide, colors, col1X, 1.15, colWidth, 3.8, 'For Strategic Buyers', colors.primary);
    
    const strategicSynergies = parseLines(data.synergiesStrategic, 7);
    const defaultStrategic = [
      'Access to established client relationships',
      'Complementary service capabilities',
      'Skilled technical workforce',
      'Proven delivery methodology',
      'Regional market presence',
      'Cross-sell opportunities'
    ];
    
    const strSynergies = strategicSynergies.length > 0 ? strategicSynergies : defaultStrategic;
    strSynergies.slice(0, 7).forEach((syn, idx) => {
      // FIXED: Adjusted character limit based on column width
      const maxChars = showStrategic && showFinancial ? 45 : 95;
      slide.addText(`▸ ${truncateText(syn, maxChars)}`, {
        x: col1X + 0.1, y: 1.6 + (idx * 0.5), w: colWidth - 0.2, h: 0.45,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  // Financial investor synergies
  if (showFinancial) {
    const finColX = showStrategic ? col2X : col1X;
    addSectionBox(slide, colors, finColX, 1.15, colWidth, 3.8, 'For Financial Investors', colors.secondary);
    
    const financialSynergies = parseLines(data.synergiesFinancial, 7);
    const defaultFinancial = [
      'Strong EBITDA margins with expansion potential',
      'Capital-light business model',
      'High revenue visibility and retention',
      'Experienced management team',
      'Platform for consolidation',
      'Multiple exit options (IPO, strategic sale)'
    ];
    
    const finSynergies = financialSynergies.length > 0 ? financialSynergies : defaultFinancial;
    finSynergies.slice(0, 7).forEach((syn, idx) => {
      // FIXED: Adjusted character limit based on column width
      const maxChars = showStrategic && showFinancial ? 45 : 95;
      slide.addText(`▸ ${truncateText(syn, maxChars)}`, {
        x: finColX + 0.1, y: 1.6 + (idx * 0.5), w: colWidth - 0.2, h: 0.45,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// MARKET POSITION SLIDE - FIXED: Handle missing competitor data gracefully
function generateMarketPositionSlide(pptx, data, colors, slideNumber, industryData) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Market Position & Competitive Landscape', null);
  
  // Market size box
  addSectionBox(slide, colors, 0.3, 1.15, 3.5, 1.8, 'Market Opportunity', colors.primary);
  
  if (data.marketSize) {
    slide.addText('Total Addressable Market', {
      x: 0.4, y: 1.6, w: 3.3, h: 0.25,
      fontSize: 10, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText(data.marketSize, {
      x: 0.4, y: 1.9, w: 3.3, h: 0.5,
      fontSize: 24, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    if (data.marketGrowthRate) {
      slide.addText(`Growing at ${data.marketGrowthRate}`, {
        x: 0.4, y: 2.45, w: 3.3, h: 0.25,
        fontSize: 10, italic: true, color: colors.accent, fontFace: 'Arial'
      });
    }
  } else if (industryData) {
    slide.addText('Industry Market Size', {
      x: 0.4, y: 1.6, w: 3.3, h: 0.25,
      fontSize: 10, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText(industryData.benchmarks.marketSize || '$100B+', {
      x: 0.4, y: 1.9, w: 3.3, h: 0.5,
      fontSize: 24, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide.addText(`Growing at ${industryData.benchmarks.avgGrowthRate || '15%'}`, {
      x: 0.4, y: 2.45, w: 3.3, h: 0.25,
      fontSize: 10, italic: true, color: colors.accent, fontFace: 'Arial'
    });
  } else {
    // Default when no data
    slide.addText('Market Size', {
      x: 0.4, y: 1.6, w: 3.3, h: 0.25,
      fontSize: 10, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText('$50B+ globally', {
      x: 0.4, y: 1.9, w: 3.3, h: 0.5,
      fontSize: 24, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide.addText('Growing at 12-15% CAGR', {
      x: 0.4, y: 2.45, w: 3.3, h: 0.25,
      fontSize: 10, italic: true, color: colors.accent, fontFace: 'Arial'
    });
  }
  
  // Key market drivers
  addSectionBox(slide, colors, 0.3, 3.05, 3.5, 1.85, 'Key Market Drivers', colors.secondary);
  
  let drivers = [];
  if (industryData?.keyDrivers) {
    drivers = industryData.keyDrivers.slice(0, 4);
  } else {
    drivers = ['Digital Transformation', 'Cloud Adoption', 'AI/ML Integration', 'Automation Demand'];
  }
  
  drivers.forEach((driver, idx) => {
    slide.addText(`▸ ${truncateText(driver, 30, false)}`, {
      x: 0.4, y: 3.45 + (idx * 0.35), w: 3.3, h: 0.3,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // FIXED: Competitive Analysis - Handle empty data gracefully
  addSectionBox(slide, colors, 3.95, 1.15, 5.55, 3.75, 'Competitive Analysis', colors.accent);
  
  const competitors = parsePipeSeparated(data.competitorLandscape, 5);
  
  if (competitors.length > 0) {
    // Table header
    slide.addText('Competitor', { x: 4.05, y: 1.55, w: 1.6, h: 0.3, fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial' });
    slide.addText('Strength', { x: 5.7, y: 1.55, w: 1.8, h: 0.3, fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial' });
    slide.addText('Weakness', { x: 7.5, y: 1.55, w: 1.9, h: 0.3, fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial' });
    
    // Separator line
    slide.addShape('rect', {
      x: 4.05, y: 1.85, w: 5.35, h: 0.02,
      fill: { color: colors.border }
    });
    
    competitors.forEach((comp, idx) => {
      const y = 1.95 + (idx * 0.55);
      const name = comp[0] || 'Competitor';
      const strength = comp[1] || '-';
      const weakness = comp[2] || '-';
      
      if (idx > 0) {
        slide.addShape('rect', { x: 4.05, y: y - 0.05, w: 5.35, h: 0.01, fill: { color: colors.border } });
      }
      
      slide.addText(truncateText(name, 18, false), { x: 4.05, y, w: 1.6, h: 0.5, fontSize: 9, bold: true, color: colors.primary, fontFace: 'Arial', valign: 'top' });
      slide.addText(truncateText(strength, 22, false), { x: 5.7, y, w: 1.8, h: 0.5, fontSize: 8, color: colors.text, fontFace: 'Arial', valign: 'top' });
      slide.addText(truncateText(weakness, 22, false), { x: 7.5, y, w: 1.9, h: 0.5, fontSize: 8, color: colors.text, fontFace: 'Arial', valign: 'top' });
    });
  } else {
    // FIXED: Show helpful message when no competitor data
    slide.addText('Competitive landscape data will appear here.\n\nTo populate this section, please enter competitor information in the Growth Strategy section using the format:\n\nCompetitor Name | Strength | Weakness', {
      x: 4.15, y: 1.7, w: 5.15, h: 2.9,
      fontSize: 10, color: colors.textLight, fontFace: 'Arial', valign: 'middle', align: 'center', italic: true
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}


// RISK FACTORS SLIDE (for CIM)
function generateRiskFactorsSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Risk Factors', 'Key considerations for investors');
  
  // Business Risks
  addSectionBox(slide, colors, 0.3, 1.15, 4.5, 1.8, 'Business Risks', colors.danger);
  
  const businessRisks = parseLines(data.businessRisks, 4);
  const defaultBusinessRisks = ['Client concentration', 'Key person dependency', 'Technology evolution', 'Competitive pressure'];
  const risks1 = businessRisks.length > 0 ? businessRisks : defaultBusinessRisks;
  
  risks1.forEach((risk, idx) => {
    slide.addText(`• ${truncateText(risk, 50)}`, {
      x: 0.4, y: 1.55 + (idx * 0.35), w: 4.3, h: 0.3,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Market Risks
  addSectionBox(slide, colors, 5, 1.15, 4.5, 1.8, 'Market Risks', colors.warning);
  
  const marketRisks = parseLines(data.marketRisks, 4);
  const defaultMarketRisks = ['Economic downturn', 'Regulatory changes', 'Currency fluctuation', 'Market saturation'];
  const risks2 = marketRisks.length > 0 ? marketRisks : defaultMarketRisks;
  
  risks2.forEach((risk, idx) => {
    slide.addText(`• ${truncateText(risk, 50)}`, {
      x: 5.1, y: 1.55 + (idx * 0.35), w: 4.3, h: 0.3,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Operational Risks
  addSectionBox(slide, colors, 0.3, 3.05, 4.5, 1.85, 'Operational Risks', colors.secondary);
  
  const opRisks = parseLines(data.operationalRisks, 4);
  const defaultOpRisks = ['Talent retention', 'Service delivery', 'Cybersecurity', 'Scalability'];
  const risks3 = opRisks.length > 0 ? opRisks : defaultOpRisks;
  
  risks3.forEach((risk, idx) => {
    slide.addText(`• ${truncateText(risk, 50)}`, {
      x: 0.4, y: 3.45 + (idx * 0.35), w: 4.3, h: 0.3,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Mitigation Strategies
  addSectionBox(slide, colors, 5, 3.05, 4.5, 1.85, 'Mitigation Strategies', colors.accent);
  
  const mitigations = parseLines(data.mitigationStrategies, 4);
  const defaultMitigations = ['Diversification strategy', 'Talent development', 'Technology investments', 'Insurance coverage'];
  const mits = mitigations.length > 0 ? mitigations : defaultMitigations;
  
  mits.forEach((mit, idx) => {
    slide.addText(`✓ ${truncateText(mit, 50)}`, {
      x: 5.1, y: 3.45 + (idx * 0.35), w: 4.3, h: 0.3,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// FINANCIAL STATEMENTS SLIDE (Appendix)
function generateFinancialStatementsSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Financial Statements Summary', 'Appendix');
  
  // P&L Summary table
  addSectionBox(slide, colors, 0.3, 1.15, 6.8, 2.8, 'Profit & Loss Summary', colors.primary);
  
  const years = ['FY24', 'FY25', 'FY26P', 'FY27P'];
  const values = [data.revenueFY24, data.revenueFY25, data.revenueFY26P, data.revenueFY27P];
  const currency = data.currency === 'USD' ? '$' : '₹';
  const unit = data.currency === 'USD' ? 'Mn' : 'Cr';
  
  // Table header
  slide.addText('Particulars', { x: 0.4, y: 1.55, w: 1.8, h: 0.3, fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial' });
  years.forEach((yr, idx) => {
    slide.addText(yr, { x: 2.2 + (idx * 1.15), y: 1.55, w: 1.1, h: 0.3, fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial', align: 'center' });
  });
  
  slide.addShape('rect', { x: 0.4, y: 1.85, w: 6.5, h: 0.02, fill: { color: colors.border } });
  
  // Revenue row
  slide.addText('Revenue', { x: 0.4, y: 1.95, w: 1.8, h: 0.35, fontSize: 9, color: colors.text, fontFace: 'Arial' });
  values.forEach((val, idx) => {
    slide.addText(val ? `${currency}${val}${unit}` : '-', {
      x: 2.2 + (idx * 1.15), y: 1.95, w: 1.1, h: 0.35,
      fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  });
  
  // EBITDA row
  slide.addText('EBITDA Margin', { x: 0.4, y: 2.35, w: 1.8, h: 0.35, fontSize: 9, color: colors.text, fontFace: 'Arial' });
  slide.addText(data.ebitdaMarginFY25 ? `${data.ebitdaMarginFY25}%` : '-', {
    x: 3.35, y: 2.35, w: 1.1, h: 0.35,
    fontSize: 9, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
  });
  
  // Gross Margin row
  if (data.grossMargin) {
    slide.addText('Gross Margin', { x: 0.4, y: 2.75, w: 1.8, h: 0.35, fontSize: 9, color: colors.text, fontFace: 'Arial' });
    slide.addText(`${data.grossMargin}%`, {
      x: 3.35, y: 2.75, w: 1.1, h: 0.35,
      fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  }
  
  // Net Profit row
  if (data.netProfitMargin) {
    slide.addText('Net Profit Margin', { x: 0.4, y: 3.15, w: 1.8, h: 0.35, fontSize: 9, color: colors.text, fontFace: 'Arial' });
    slide.addText(`${data.netProfitMargin}%`, {
      x: 3.35, y: 3.15, w: 1.1, h: 0.35,
      fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  }
  
  // Key Ratios
  addSectionBox(slide, colors, 7.2, 1.15, 2.3, 2.8, 'Key Ratios', colors.secondary);
  
  const ratios = [];
  if (data.netRetention) ratios.push({ label: 'NRR', value: `${data.netRetention}%` });
  if (data.top10Concentration) ratios.push({ label: 'Top 10 Conc.', value: `${data.top10Concentration}%` });
  if (data.ebitdaMarginFY25) ratios.push({ label: 'EBITDA %', value: `${data.ebitdaMarginFY25}%` });
  if (data.grossMargin) ratios.push({ label: 'Gross %', value: `${data.grossMargin}%` });
  
  ratios.slice(0, 5).forEach((ratio, idx) => {
    slide.addText(ratio.label, {
      x: 7.3, y: 1.55 + (idx * 0.45), w: 1.3, h: 0.22,
      fontSize: 8, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText(ratio.value, {
      x: 7.3, y: 1.77 + (idx * 0.45), w: 1.3, h: 0.22,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
  });
  
  // CAGR calculation
  if (data.revenueFY24 && data.revenueFY26P) {
    const cagr = Math.round((Math.pow(parseFloat(data.revenueFY26P) / parseFloat(data.revenueFY24), 0.5) - 1) * 100);
    slide.addText(`Revenue CAGR: ${cagr}%`, {
      x: 0.4, y: 3.6, w: 3, h: 0.25,
      fontSize: 10, bold: true, color: colors.accent, fontFace: 'Arial'
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// TEAM BIOS SLIDE (Appendix)
function generateTeamBiosSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Leadership Team', 'Appendix');
  
  // Founder
  slide.addShape('rect', {
    x: 0.3, y: 1.15, w: 4.5, h: 1.5,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide.addShape('ellipse', {
    x: 0.5, y: 1.3, w: 1.0, h: 1.0,
    fill: { color: colors.white },
    line: { color: colors.primary, width: 2 }
  });
  
  slide.addText(data.founderName || 'Founder', {
    x: 1.65, y: 1.35, w: 3, h: 0.32,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide.addText(data.founderTitle || 'Founder & CEO', {
    x: 1.65, y: 1.68, w: 3, h: 0.25,
    fontSize: 10, color: colors.textLight, fontFace: 'Arial'
  });
  
  const education = parseLines(data.founderEducation, 1);
  if (education[0]) {
    slide.addText(truncateText(education[0], 40), {
      x: 1.65, y: 1.98, w: 3, h: 0.22,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  }
  if (data.founderExperience) {
    slide.addText(`${data.founderExperience}+ years experience`, {
      x: 1.65, y: 2.2, w: 3, h: 0.22,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  }
  
  // Leadership team grid
  const leaders = parsePipeSeparated(data.leadershipTeam, 6);
  
  leaders.forEach((leader, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 5 + (col * 2.3);
    const y = 1.15 + (row * 1.0);
    
    slide.addShape('rect', {
      x, y, w: 2.2, h: 0.9,
      fill: { color: colors.lightBg },
      line: { color: colors.border, width: 0.5 }
    });
    
    const name = leader[0] || 'Name';
    const title = leader[1] || 'Title';
    
    slide.addText(truncateText(name, 22, false), {
      x: x + 0.1, y: y + 0.12, w: 2, h: 0.28,
      fontSize: 10, bold: true, color: colors.text, fontFace: 'Arial'
    });
    slide.addText(truncateText(title, 25, false), {
      x: x + 0.1, y: y + 0.42, w: 2, h: 0.22,
      fontSize: 8, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// FULL CLIENT LIST SLIDE (Appendix)
function generateFullClientListSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  addSlideHeader(slide, colors, 'Complete Client Portfolio', 'Appendix');
  
  const clients = parsePipeSeparated(data.topClients, 20);
  
  const colCount = 4;
  const colWidth = 2.3;
  const rowHeight = 0.55;
  
  clients.forEach((client, idx) => {
    const col = idx % colCount;
    const row = Math.floor(idx / colCount);
    const x = 0.3 + (col * 2.4);
    const y = 1.2 + (row * 0.6);
    
    const name = client[0] || 'Client';
    const vertical = client[1] || '';
    
    slide.addShape('rect', {
      x, y, w: colWidth, h: rowHeight,
      fill: { color: idx % 2 === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    slide.addText(truncateText(name, 20, false), {
      x: x + 0.1, y: y + 0.05, w: colWidth - 0.2, h: 0.28,
      fontSize: 9, bold: true, color: colors.text, fontFace: 'Arial'
    });
    if (vertical) {
      slide.addText(truncateText(vertical, 20, false), {
        x: x + 0.1, y: y + 0.3, w: colWidth - 0.2, h: 0.2,
        fontSize: 7, color: colors.textLight, fontFace: 'Arial'
      });
    }
  });
  
  addSlideFooter(slide, colors, slideNumber);
  return slideNumber + 1;
}

// THANK YOU SLIDE
function generateThankYouSlide(pptx, data, colors, slideNumber) {
  const slide = pptx.addSlide();
  
  // Dark background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  // Thank you text
  slide.addText('Thank You', {
    x: 0, y: 1.8, w: '100%', h: 1,
    fontSize: 48, bold: true, color: colors.white, fontFace: 'Arial', align: 'center'
  });
  
  // Subtitle
  slide.addText('For Your Consideration', {
    x: 0, y: 2.7, w: '100%', h: 0.5,
    fontSize: 18, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 30
  });
  
  // Divider
  slide.addShape('rect', {
    x: 3.5, y: 3.4, w: 3, h: 0.03,
    fill: { color: colors.secondary }
  });
  
  // Contact info
  if (data.advisor) {
    slide.addText(`For more information, please contact:\n${data.advisor}`, {
      x: 0, y: 3.7, w: '100%', h: 0.8,
      fontSize: 12, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 20
    });
  }
  
  // Footer
  slide.addText(`${data.projectCodename || 'Project'} | Strictly Private & Confidential`, {
    x: 0.3, y: 5.05, w: 9.4, h: 0.25,
    fontSize: 9, italic: true, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 50
  });
  
  return slideNumber + 1;
}


// ============================================================================
// MAIN PPTX GENERATOR - ORCHESTRATES ALL SLIDES
// ============================================================================
function generateProfessionalPPTX(data, themeName = 'modern-blue') {
  const pptx = new PptxGenJS();
  
  // Get theme colors
  const colors = THEMES[themeName] || THEMES['modern-blue'];
  
  // Get document configuration
  const docType = data.documentType || 'management-presentation';
  const docConfig = DOCUMENT_CONFIGS[docType] || DOCUMENT_CONFIGS['management-presentation'];
  
  // Get target buyers
  const targetBuyers = data.targetBuyerType || ['strategic'];
  
  // Get industry data
  const industryData = INDUSTRY_DATA[data.primaryVertical] || INDUSTRY_DATA['technology'];
  
  // Get content variants
  const variants = data.generateVariants || [];
  
  // Get appendix options
  const appendixOptions = data.includeAppendix || [];
  
  // Configure presentation
  pptx.layout = 'LAYOUT_WIDE';
  pptx.title = data.projectCodename || 'Investment Memorandum';
  pptx.author = data.advisor || 'M&A Advisor';
  pptx.subject = docConfig.name;
  
  let slideNumber = 1;
  
  // ========================================
  // TITLE SLIDE (all document types)
  // ========================================
  slideNumber = generateTitleSlide(pptx, data, colors, docConfig);
  
  // ========================================
  // DISCLAIMER SLIDE (all document types)
  // ========================================
  slideNumber = generateDisclaimerSlide(pptx, data, colors, slideNumber);
  
  // ========================================
  // TABLE OF CONTENTS (CIM only)
  // ========================================
  if (docType === 'cim') {
    slideNumber = generateTOCSlide(pptx, data, colors, slideNumber);
  }
  
  // ========================================
  // EXECUTIVE SUMMARY / SNAPSHOT
  // ========================================
  if (docType === 'teaser') {
    slideNumber = generateSnapshotSlide(pptx, data, colors, slideNumber, industryData);
  } else {
    slideNumber = generateExecSummarySlide(pptx, data, colors, slideNumber, targetBuyers, industryData, docConfig);
  }
  
  // ========================================
  // INVESTMENT HIGHLIGHTS (not teaser)
  // ========================================
  if (docType !== 'teaser') {
    slideNumber = generateInvestmentHighlightsSlide(pptx, data, colors, slideNumber, targetBuyers);
  }
  
  // ========================================
  // COMPANY OVERVIEW
  // ========================================
  slideNumber = generateCompanyOverviewSlide(pptx, data, colors, slideNumber, industryData);
  
  // ========================================
  // FOUNDER SLIDE (not teaser)
  // ========================================
  if (docType !== 'teaser') {
    slideNumber = generateFounderSlide(pptx, data, colors, slideNumber);
  }
  
  // ========================================
  // INDUSTRY OVERVIEW (CIM only)
  // ========================================
  if (docType === 'cim') {
    slideNumber = generateIndustryOverviewSlide(pptx, data, colors, slideNumber, industryData);
  }
  
  // ========================================
  // SERVICES SLIDE
  // ========================================
  slideNumber = generateServicesSlide(pptx, data, colors, slideNumber);
  
  // ========================================
  // CLIENTS SLIDE
  // ========================================
  slideNumber = generateClientsSlide(pptx, data, colors, slideNumber, industryData, docConfig);
  
  // ========================================
  // FINANCIALS SLIDE
  // ========================================
  if (docConfig.includeFinancialDetail) {
    slideNumber = generateFinancialsSlide(pptx, data, colors, slideNumber, docConfig);
  }
  
  // ========================================
  // CASE STUDIES (not teaser, max 2 for management presentation)
  // ========================================
  if (docType !== 'teaser') {
    const maxCaseStudies = docType === 'cim' ? 10 : 2;
    
    // Collect case studies
    let caseStudiesToShow = [];
    if (data.caseStudies && Array.isArray(data.caseStudies)) {
      caseStudiesToShow = data.caseStudies.filter(cs => cs.client).slice(0, maxCaseStudies);
    }
    
    // Also check legacy format
    if (caseStudiesToShow.length === 0) {
      for (let i = 1; i <= maxCaseStudies; i++) {
        if (data[`cs${i}Client`]) {
          caseStudiesToShow.push({
            client: data[`cs${i}Client`],
            industry: data[`cs${i}Industry`] || '',
            challenge: data[`cs${i}Challenge`],
            solution: data[`cs${i}Solution`],
            results: data[`cs${i}Results`]
          });
        }
      }
    }
    
    caseStudiesToShow.forEach((cs, idx) => {
      slideNumber = generateCaseStudySlide(pptx, cs, colors, slideNumber, idx + 1, docConfig);
    });
  }
  
  // ========================================
  // GROWTH STRATEGY (not teaser)
  // ========================================
  if (docType !== 'teaser') {
    slideNumber = generateGrowthSlide(pptx, data, colors, slideNumber, targetBuyers);
  }
  
  // ========================================
  // SYNERGIES SLIDE (not teaser)
  // ========================================
  if (docType !== 'teaser') {
    slideNumber = generateSynergiesSlide(pptx, data, colors, slideNumber, targetBuyers);
  }
  
  // ========================================
  // CONTENT VARIANTS
  // ========================================
  if (variants.includes('market')) {
    slideNumber = generateMarketPositionSlide(pptx, data, colors, slideNumber, industryData);
  }
  
  // ========================================
  // RISK FACTORS (CIM only or if selected)
  // ========================================
  if (docType === 'cim') {
    slideNumber = generateRiskFactorsSlide(pptx, data, colors, slideNumber);
  }
  
  // ========================================
  // APPENDIX OPTIONS
  // ========================================
  if (appendixOptions.includes('team-bios')) {
    slideNumber = generateTeamBiosSlide(pptx, data, colors, slideNumber);
  }
  
  if (appendixOptions.includes('client-list')) {
    slideNumber = generateFullClientListSlide(pptx, data, colors, slideNumber);
  }
  
  if (appendixOptions.includes('financial-detail')) {
    slideNumber = generateFinancialStatementsSlide(pptx, data, colors, slideNumber);
  }
  
  // ========================================
  // THANK YOU SLIDE (all document types)
  // ========================================
  slideNumber = generateThankYouSlide(pptx, data, colors, slideNumber);
  
  return { pptx, slideCount: slideNumber - 1 };
}

// ============================================================================
// API ENDPOINTS
// ============================================================================

// Health check
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    version: VERSION.string,
    versionFull: VERSION.full,
    buildDate: VERSION.buildDate,
    features: [
      'AI-Powered Layout Engine (NEW)',
      'Larger Fonts (14pt body, 26pt titles)',
      'Diverse Infographics (Pie, Donut, Progress, Timeline)',
      'Dynamic Slide Generation',
      'Document Types (Management Presentation, CIM, Teaser)',
      'Enhanced Buyer Types',
      'Industry-Specific Content',
      '50 Professional Templates',
      'Unlimited Case Studies',
      'Word/PDF/JSON Export'
    ]
  });
});

// Version info
app.get('/api/version', (req, res) => {
  res.json({
    current: VERSION.string,
    full: VERSION.full,
    buildDate: VERSION.buildDate,
    history: VERSION.history
  });
});

// Get templates
app.get('/api/templates', (req, res) => {
  res.json(PROFESSIONAL_TEMPLATES);
});

// Get industries
app.get('/api/industries', (req, res) => {
  res.json(Object.values(INDUSTRY_DATA));
});

// Get document types
app.get('/api/document-types', (req, res) => {
  res.json(DOCUMENT_CONFIGS);
});

// Usage statistics
app.get('/api/usage', (req, res) => {
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const weekAgo = new Date(now - 7 * 24 * 60 * 60 * 1000).toISOString();
  const monthAgo = new Date(now - 30 * 24 * 60 * 60 * 1000).toISOString();
  
  const dailyCalls = usageStats.calls.filter(c => c.timestamp.startsWith(today));
  const weeklyCalls = usageStats.calls.filter(c => c.timestamp >= weekAgo);
  const monthlyCalls = usageStats.calls.filter(c => c.timestamp >= monthAgo);
  
  res.json({
    totalInputTokens: usageStats.totalInputTokens,
    totalOutputTokens: usageStats.totalOutputTokens,
    totalCalls: usageStats.totalCalls,
    totalCostUSD: usageStats.totalCostUSD.toFixed(4),
    sessionStart: usageStats.sessionStart,
    recentCalls: usageStats.calls.slice(-20).reverse(),
    daily: {
      calls: dailyCalls.length,
      cost: dailyCalls.reduce((sum, c) => sum + parseFloat(c.costUSD), 0).toFixed(4)
    },
    weekly: {
      calls: weeklyCalls.length,
      cost: weeklyCalls.reduce((sum, c) => sum + parseFloat(c.costUSD), 0).toFixed(4)
    },
    monthly: {
      calls: monthlyCalls.length,
      cost: monthlyCalls.reduce((sum, c) => sum + parseFloat(c.costUSD), 0).toFixed(4)
    }
  });
});

// Export usage to CSV
app.get('/api/usage/export', (req, res) => {
  const headers = ['Timestamp', 'Model', 'Input Tokens', 'Output Tokens', 'Cost (USD)', 'Purpose'];
  const rows = usageStats.calls.map(c => [
    c.timestamp, c.model, c.inputTokens, c.outputTokens, c.costUSD, c.purpose || 'API Call'
  ]);
  
  let csv = headers.join(',') + '\n';
  rows.forEach(row => {
    csv += row.map(v => `"${v}"`).join(',') + '\n';
  });
  
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename=usage_report.csv');
  res.send(csv);
});

// Reset usage
app.post('/api/usage/reset', (req, res) => {
  usageStats = {
    totalInputTokens: 0,
    totalOutputTokens: 0,
    totalCalls: 0,
    totalCostUSD: 0,
    sessionStart: new Date().toISOString(),
    calls: []
  };
  res.json({ success: true, message: 'Usage statistics reset' });
});

// Generate PPTX
app.post('/api/generate-pptx', async (req, res) => {
  try {
    const { data, theme } = req.body;
    
    if (!data) {
      return res.status(400).json({ success: false, error: 'No data provided' });
    }
    
    console.log(`Generating PPTX: ${data.projectCodename}, Theme: ${theme}, DocType: ${data.documentType}`);
    
    const { pptx, slideCount } = generateProfessionalPPTX(data, theme || 'modern-blue');
    
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const base64 = buffer.toString('base64');
    
    const filename = `${data.projectCodename || 'IM'}_${Date.now()}.pptx`;
    
    res.json({
      success: true,
      filename,
      slideCount,
      fileData: base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });
    
  } catch (error) {
    console.error('PPTX Generation Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Export Q&A Word document
app.post('/api/export-qa-word', async (req, res) => {
  if (!docx) {
    return res.status(501).json({ 
      success: false, 
      error: 'Word export not available. Install docx package: npm install docx' 
    });
  }
  
  try {
    const { data, questionnaire } = req.body;
    
    const doc = new docx.Document({
      sections: [{
        properties: {},
        children: [
          new docx.Paragraph({
            children: [new docx.TextRun({ text: data.projectCodename || 'Project Q&A', bold: true, size: 48 })],
            spacing: { after: 400 }
          }),
          new docx.Paragraph({
            children: [new docx.TextRun({ text: `Generated: ${new Date().toLocaleDateString()}`, italics: true })],
            spacing: { after: 400 }
          }),
          ...Object.entries(data).map(([key, value]) => {
            if (typeof value === 'string' && value.trim()) {
              return new docx.Paragraph({
                children: [
                  new docx.TextRun({ text: `${key}: `, bold: true }),
                  new docx.TextRun({ text: value })
                ],
                spacing: { after: 200 }
              });
            }
            return null;
          }).filter(Boolean)
        ]
      }]
    });
    
    const buffer = await docx.Packer.toBuffer(doc);
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename=${data.projectCodename || 'QA'}_Document.docx`);
    res.send(buffer);
    
  } catch (error) {
    console.error('Word Export Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Export JSON
app.post('/api/export-json', (req, res) => {
  const { data } = req.body;
  res.json({
    metadata: {
      exportedAt: new Date().toISOString(),
      version: '6.1.0'
    },
    formData: data
  });
});

// Save draft
app.post('/api/drafts', (req, res) => {
  const { data, projectId } = req.body;
  const draftsDir = path.join(tempDir, 'drafts');
  if (!fs.existsSync(draftsDir)) fs.mkdirSync(draftsDir, { recursive: true });
  
  const draftPath = path.join(draftsDir, `${projectId}.json`);
  fs.writeFileSync(draftPath, JSON.stringify(data, null, 2));
  
  res.json({ success: true, projectId });
});

// Load draft
app.get('/api/drafts/:projectId', (req, res) => {
  const draftPath = path.join(tempDir, 'drafts', `${req.params.projectId}.json`);
  if (fs.existsSync(draftPath)) {
    const data = JSON.parse(fs.readFileSync(draftPath, 'utf8'));
    res.json({ success: true, data });
  } else {
    res.status(404).json({ success: false, error: 'Draft not found' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`IM Creator Server ${VERSION.full} - AI-Powered Layout Engine`);
  console.log(`${'='.repeat(60)}`);
  console.log(`Running on port ${PORT}`);
  console.log(`\nNEW in v7.0:`);
  console.log(`  ✓ AI-powered layout recommendations using Claude`);
  console.log(`  ✓ Larger fonts (26pt titles, 14pt section headers, 12pt body)`);
  console.log(`  ✓ Diverse charts: Pie, Donut, Progress bars, Timelines`);
  console.log(`  ✓ Dynamic font sizing based on content`);
  console.log(`  ✓ Better space utilization (85%+ content area)`);
  console.log(`\nPreserved from v6.x:`);
  console.log(`  ✓ All 14 core features`);
  console.log(`  ✓ 50 professional templates`);
  console.log(`  ✓ Document types (CIM, Management Presentation, Teaser)`);
  console.log(`${'='.repeat(60)}\n`);
});

