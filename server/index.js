// ============================================================================
// IM Creator Server v7.1.0 - AI Layout Recommendations Applied to All Slides
// ============================================================================
// 
// v7.1.0 UPGRADE: Universal Slide Creation with AI-Powered Layouts
// - createSlide() wrapper function calls AI for each slide type
// - Dedicated render functions with layout, font, and chart recommendations
// - addChartByType() helper handles all chart rendering
// - Full integration of AI recommendations into actual slide rendering
//
// PRESERVED FROM v7.0.0:
// - AI-powered layout engine
// - Larger fonts (14pt body min, 26pt titles)
// - Diverse infographics: Pie, Donut, Bar, Progress bars, Timelines
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
// - Case study support
// - Word/PDF/JSON export
//
// VERSION HISTORY:
// v7.1.0 (2026-02-03) - AI layout recommendations applied to all slide types
// v7.0.0 (2026-02-03) - AI-powered layout engine, larger fonts, diverse charts
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
// VERSION MANAGEMENT - v7.1.0
// ============================================================================
const VERSION = {
  major: 7,
  minor: 1,
  patch: 0,
  get string() { return `${this.major}.${this.minor}.${this.patch}`; },
  get full() { return `v${this.string}`; },
  buildDate: '2026-02-03',
  history: [
    { version: '7.1.0', date: '2026-02-03', type: 'minor', changes: [
      'Universal createSlide() wrapper with AI integration',
      'Dedicated render functions for each slide type',
      'addChartByType() helper for all chart types',
      'AI recommendations now applied to actual rendering',
      'Proper font adjustments based on content density'
    ]},
    { version: '7.0.0', date: '2026-02-03', type: 'major', changes: [
      'AI-powered layout engine',
      'Larger fonts (14pt body, 26pt titles)',
      'Diverse infographics (Pie, Donut, Progress)',
      'Dynamic slide generation'
    ]},
    { version: '6.1.0', date: '2026-02-02', type: 'minor', changes: [
      'Fixed text overflow',
      'Added generateCaseStudySlide',
      'Better spacing'
    ]},
    { version: '6.0.0', date: '2026-02-01', type: 'major', changes: [
      '14 core features',
      '50 templates',
      'CIM/Teaser support'
    ]}
  ]
};

// ============================================================================
// DESIGN CONSTANTS - v7.1 OPTIMIZED FOR AI LAYOUT
// ============================================================================
const DESIGN = {
  // Slide dimensions (widescreen 16:9)
  slideWidth: 10,
  slideHeight: 5.625,
  
  // Margins - optimized for 85%+ content usage
  margin: { left: 0.3, right: 0.3, top: 0.2, bottom: 0.3 },
  
  // Content area
  contentWidth: 9.4,
  contentTop: 0.95,
  contentHeight: 4.2,
  
  // Font sizes - SIGNIFICANTLY LARGER than v6.x
  fonts: {
    title: 26,           // Was 18-20, now 26
    subtitle: 14,        // Was 10-11, now 14
    sectionHeader: 14,   // Was 10-11, now 14
    bodyLarge: 13,       // For emphasis
    body: 12,            // Was 9-10, now 12
    bodySmall: 11,       // Was 8-9, now 11
    caption: 10,         // Was 7-8, now 10
    metric: 32,          // Large key numbers
    metricMedium: 26,    // Medium numbers
    metricSmall: 20,     // Smaller metrics
    metricLabel: 11,     // Labels under metrics
    chartLabel: 11,      // Was 7-8, now 11
    footer: 9
  },
  
  // Spacing
  spacing: {
    sectionGap: 0.12,
    itemGap: 0.08,
    boxPadding: 0.1
  },
  
  // Layout presets for AI recommendations
  layouts: {
    'full-width': { columns: 1, contentWidth: 9.4 },
    'two-column': { columns: 2, leftWidth: 4.55, rightWidth: 4.55, gap: 0.3 },
    'two-column-wide-left': { columns: 2, leftWidth: 5.5, rightWidth: 3.6, gap: 0.3 },
    'two-column-wide-right': { columns: 2, leftWidth: 3.6, rightWidth: 5.5, gap: 0.3 },
    'three-column': { columns: 3, colWidth: 3.0, gap: 0.2 },
    'grid-2x2': { rows: 2, cols: 2, cellWidth: 4.55, cellHeight: 1.9, gap: 0.2 },
    'grid-2x3': { rows: 2, cols: 3, cellWidth: 3.0, cellHeight: 1.9, gap: 0.15 }
  }
};

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
// AI LAYOUT ENGINE - ENHANCED IN v7.1
// ============================================================================

/**
 * Analyzes data and returns AI-powered layout recommendations
 * @param {Object} data - The form data
 * @param {string} slideType - Type of slide to analyze
 * @returns {Object} Layout recommendations
 */
async function analyzeDataForLayout(data, slideType) {
  // Build data preview for AI analysis
  const dataPreview = buildDataPreview(data, slideType);

  const prompt = `You are an expert presentation designer. Analyze this data for a "${slideType}" slide.

DATA SUMMARY:
${JSON.stringify(dataPreview, null, 2)}

Recommend the optimal design. Return ONLY valid JSON (no markdown):
{
  "chartType": "bar|pie|donut|progress|timeline|stacked-bar|none",
  "layout": "full-width|two-column|two-column-wide-left|two-column-wide-right|grid-2x2|grid-2x3",
  "fontAdjustment": 0,
  "contentDensity": "low|medium|high",
  "primaryEmphasis": "metrics|chart|text|mixed",
  "recommendations": ["suggestion1", "suggestion2"]
}

Guidelines:
- Use pie/donut for 2-5 composition items
- Use bar for time series/revenue growth  
- Use progress bars for percentages/margins
- Use timeline for milestones/roadmap
- Use stacked-bar for revenue breakdown
- Use two-column for balanced content
- Use full-width for case studies or text-heavy content
- fontAdjustment: 0 for normal, -1 for dense content, -2 for very dense
- Prioritize readability (12pt body minimum)`;

  try {
    const response = await anthropic.messages.create({
      model: 'claude-3-haiku-20240307',
      max_tokens: 400,
      messages: [{ role: 'user', content: prompt }]
    });
    
    trackUsage('claude-3-haiku-20240307', response.usage.input_tokens, response.usage.output_tokens, `AI Layout: ${slideType}`);
    
    const text = response.content[0].text;
    const jsonMatch = text.match(/\{[\s\S]*?\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      console.log(`AI Layout for ${slideType}:`, parsed);
      return parsed;
    }
  } catch (error) {
    console.log(`AI Layout fallback for ${slideType}:`, error.message);
  }
  
  // Smart fallback based on slide type
  return getDefaultLayoutRecommendation(slideType, dataPreview);
}

/**
 * Builds a data preview object for AI analysis
 */
function buildDataPreview(data, slideType) {
  const base = {
    hasRevenue: !!(data.revenueFY24 || data.revenueFY25),
    revenueYears: [data.revenueFY24, data.revenueFY25, data.revenueFY26P, data.revenueFY27P].filter(Boolean).length,
    serviceCount: (data.serviceLines || '').split('\n').filter(x => x.trim()).length,
    clientCount: (data.topClients || '').split('\n').filter(x => x.trim()).length,
    hasDescription: !!(data.companyDescription && data.companyDescription.length > 50),
    descriptionLength: (data.companyDescription || '').length,
    highlightCount: (data.investmentHighlights || '').split('\n').filter(x => x.trim()).length,
    hasMargins: !!(data.ebitdaMarginFY25 || data.grossMargin),
    caseStudyCount: (data.caseStudies || []).length || (data.cs1Client ? 1 : 0) + (data.cs2Client ? 1 : 0)
  };
  
  // Add slide-specific data
  switch (slideType) {
    case 'executive-summary':
      return { ...base, hasFounder: !!data.founderName, hasEmployees: !!data.employeeCountFT };
    case 'services':
      return { ...base, hasProducts: !!(data.products && data.products.trim()) };
    case 'clients':
      return { ...base, hasVerticals: !!(data.otherVerticals && data.otherVerticals.trim()), hasPartners: !!(data.techPartnerships) };
    case 'financials':
      return { ...base, hasServiceRevenue: !!(data.revenueByService) };
    case 'growth':
      return { ...base, hasDrivers: !!(data.growthDrivers), hasGoals: !!(data.shortTermGoals || data.mediumTermGoals) };
    default:
      return base;
  }
}

/**
 * Returns smart default layout recommendations when AI is unavailable
 */
function getDefaultLayoutRecommendation(slideType, dataPreview) {
  const defaults = {
    'executive-summary': {
      chartType: 'bar',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'mixed'
    },
    'investment-highlights': {
      chartType: 'none',
      layout: 'grid-2x2',
      fontAdjustment: dataPreview.highlightCount > 6 ? -1 : 0,
      contentDensity: dataPreview.highlightCount > 6 ? 'high' : 'medium',
      primaryEmphasis: 'text'
    },
    'services': {
      chartType: dataPreview.serviceCount <= 4 ? 'donut' : 'pie',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'mixed'
    },
    'clients': {
      chartType: 'donut',
      layout: 'two-column-wide-right',
      fontAdjustment: dataPreview.clientCount > 9 ? -1 : 0,
      contentDensity: dataPreview.clientCount > 9 ? 'high' : 'medium',
      primaryEmphasis: 'mixed'
    },
    'financials': {
      chartType: 'bar',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'chart'
    },
    'case-study': {
      chartType: 'none',
      layout: 'full-width',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'text'
    },
    'growth': {
      chartType: 'timeline',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'mixed'
    },
    'market-position': {
      chartType: 'bar',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'mixed'
    },
    'synergies': {
      chartType: 'none',
      layout: 'two-column',
      fontAdjustment: 0,
      contentDensity: 'medium',
      primaryEmphasis: 'text'
    }
  };
  
  return defaults[slideType] || {
    chartType: 'none',
    layout: 'two-column',
    fontAdjustment: 0,
    contentDensity: 'medium',
    primaryEmphasis: 'text'
  };
}

/**
 * Applies font adjustment to base size
 */
function adjustedFont(baseSize, adjustment = 0) {
  return Math.max(9, baseSize + adjustment);
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
    financialEmphasis: ['Revenue synergies', 'Cost synergies', 'Market share gains'],
    slideAdjustments: {
      synergies: { emphasize: 'strategic' },
      financials: { showProjections: true }
    }
  },
  financial: {
    name: 'Financial Investor',
    focus: ['Growth potential', 'Margin expansion', 'Exit multiple'],
    keyMessages: [
      'Strong EBITDA margins',
      'Clear path to value creation',
      'Experienced management team'
    ],
    financialEmphasis: ['EBITDA growth', 'Cash conversion', 'IRR potential'],
    slideAdjustments: {
      synergies: { emphasize: 'financial' },
      financials: { showProjections: true, emphasizeEbitda: true }
    }
  },
  international: {
    name: 'International Acquirer',
    focus: ['Market entry', 'Local expertise', 'Regulatory navigation'],
    keyMessages: [
      'Local market presence',
      'Regulatory understanding',
      'Cost-effective talent base'
    ],
    financialEmphasis: ['Currency considerations', 'Transfer pricing', 'Tax efficiency'],
    slideAdjustments: {
      synergies: { emphasize: 'international' },
      financials: { showCurrencyNotes: true }
    }
  }
};

// ============================================================================
// DOCUMENT TYPE CONFIGURATIONS
// ============================================================================
const DOCUMENT_CONFIGS = {
  'management-presentation': {
    name: 'Management Presentation',
    slideRange: '12-18 slides',
    minSlides: 12,
    maxSlides: 18,
    includeFinancialDetail: true,
    includeSensitiveData: true,
    includeClientNames: true,
    maxCaseStudies: 2,
    requiredSlides: ['title', 'disclaimer', 'executive-summary', 'investment-highlights', 'services', 'clients', 'financials', 'thank-you'],
    optionalSlides: ['leadership', 'case-studies', 'growth', 'synergies', 'market-position']
  },
  'cim': {
    name: 'Confidential Information Memorandum',
    slideRange: '20-35 slides',
    minSlides: 20,
    maxSlides: 35,
    includeFinancialDetail: true,
    includeSensitiveData: true,
    includeClientNames: true,
    maxCaseStudies: 5,
    requiredSlides: ['title', 'disclaimer', 'toc', 'executive-summary', 'investment-highlights', 'company-overview', 'leadership', 'industry', 'services', 'clients', 'financials', 'growth', 'synergies', 'risks', 'thank-you'],
    optionalSlides: ['case-studies', 'market-position', 'team-bios', 'financial-detail']
  },
  'teaser': {
    name: 'Teaser Document',
    slideRange: '5-8 slides',
    minSlides: 5,
    maxSlides: 8,
    includeFinancialDetail: false,
    includeSensitiveData: false,
    includeClientNames: false,
    maxCaseStudies: 0,
    requiredSlides: ['title', 'disclaimer', 'executive-summary', 'services', 'thank-you'],
    optionalSlides: ['investment-highlights', 'market-position']
  }
};

// ============================================================================
// 50 PROFESSIONAL TEMPLATES
// ============================================================================
const PROFESSIONAL_TEMPLATES = [
  // Modern & Tech (1-10)
  { id: 'modern-blue', name: 'Modern Blue', category: 'Modern', primary: '2B579A', secondary: '86BC25', accent: 'E8463A' },
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
    text: '2D3748',
    textLight: '718096',
    white: 'FFFFFF',
    lightBg: 'F7FAFC',
    darkBg: t.primary,
    border: 'E2E8F0',
    success: '38A169',
    warning: 'D69E2E',
    danger: 'E53E3E',
    chartColors: [t.primary, t.secondary, t.accent, '38A169', 'E53E3E', '805AD5', '00B5D8', 'ED8936']
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

function truncateText(text, maxLength, useEllipsis = true) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  
  let condensed = condenseText(text);
  if (condensed.length <= maxLength) return condensed;
  
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
  
  const cutoff = maxLength - (useEllipsis ? 2 : 0);
  const truncated = condensed.substring(0, cutoff);
  const lastSpace = truncated.lastIndexOf(' ');
  
  if (lastSpace > cutoff * 0.7) {
    return truncated.substring(0, lastSpace).trim() + (useEllipsis ? '..' : '');
  }
  return truncated.trim() + (useEllipsis ? '..' : '');
}

function truncateDescription(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  
  let condensed = condenseText(text);
  if (condensed.length <= maxLength) return condensed;
  
  const sentenceEnd = condensed.lastIndexOf('.', maxLength - 1);
  if (sentenceEnd > maxLength * 0.5) {
    return condensed.substring(0, sentenceEnd + 1);
  }
  
  const wordEnd = condensed.lastIndexOf(' ', maxLength - 2);
  if (wordEnd > maxLength * 0.6) {
    return condensed.substring(0, wordEnd) + '..';
  }
  
  return condensed.substring(0, maxLength - 2) + '..';
}

function formatCurrency(value, currency = 'INR') {
  if (!value) return 'N/A';
  const num = parseFloat(value);
  if (isNaN(num)) return value;
  return currency === 'USD' ? `$${num}M` : `₹${num}Cr`;
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
    .map(line => line.split('|').map(p => p.trim()));
}

function calculateCAGR(startValue, endValue, years) {
  if (!startValue || !endValue || startValue <= 0 || years <= 0) return null;
  return Math.round((Math.pow(endValue / startValue, 1 / years) - 1) * 100);
}

// ============================================================================
// SLIDE HELPER FUNCTIONS - v7.1 ENHANCED
// ============================================================================

/**
 * Adds slide header with title and optional subtitle
 */
function addSlideHeader(slide, colors, title, subtitle = null, fontAdj = 0) {
  // Background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.white }
  });
  
  // Left accent bar
  slide.addShape('rect', {
    x: 0, y: 0, w: 0.1, h: 0.85,
    fill: { color: colors.secondary }
  });
  
  // Title - LARGE (26pt base)
  slide.addText(truncateText(title, 80), {
    x: 0.3, y: 0.15, w: 9.4, h: 0.5,
    fontSize: adjustedFont(DESIGN.fonts.title, fontAdj),
    bold: true,
    color: colors.primary,
    fontFace: 'Arial',
    valign: 'middle'
  });
  
  // Subtitle if provided
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.3, y: 0.62, w: 9.4, h: 0.22,
      fontSize: adjustedFont(DESIGN.fonts.subtitle, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial',
      italic: true
    });
  }
  
  // Accent line
  slide.addShape('rect', {
    x: 0.3, y: 0.88, w: 9.4, h: 0.02,
    fill: { color: colors.accent }
  });
}

/**
 * Adds slide footer with page number
 */
function addSlideFooter(slide, colors, pageNumber, confidential = true) {
  slide.addShape('rect', {
    x: 0, y: 5.2, w: '100%', h: 0.015,
    fill: { color: colors.primary }
  });
  
  if (confidential) {
    slide.addText('Strictly Private & Confidential', {
      x: 0.3, y: 5.28, w: 3, h: 0.22,
      fontSize: DESIGN.fonts.footer,
      italic: true,
      color: colors.textLight,
      fontFace: 'Arial'
    });
  }
  
  slide.addText(`${pageNumber}`, {
    x: 9.2, y: 5.28, w: 0.5, h: 0.22,
    fontSize: DESIGN.fonts.caption,
    color: colors.primary,
    fontFace: 'Arial',
    align: 'right',
    bold: true
  });
}

/**
 * Adds a section box with header
 */
function addSectionBox(slide, colors, x, y, w, h, title, titleBgColor = null, fontAdj = 0) {
  // Main box
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 },
    rectRadius: 0.04
  });
  
  // Title bar
  if (title) {
    const headerH = 0.36;
    slide.addShape('rect', {
      x, y, w, h: headerH,
      fill: { color: titleBgColor || colors.primary },
      rectRadius: 0.04
    });
    // Square off bottom corners
    slide.addShape('rect', {
      x, y: y + headerH - 0.04, w, h: 0.04,
      fill: { color: titleBgColor || colors.primary }
    });
    
    slide.addText(truncateText(title, 45), {
      x: x + 0.12, y: y + 0.02, w: w - 0.24, h: headerH - 0.04,
      fontSize: adjustedFont(DESIGN.fonts.sectionHeader, fontAdj),
      bold: true,
      color: colors.white,
      fontFace: 'Arial',
      valign: 'middle'
    });
  }
}

/**
 * Adds a metric card with large value and label
 */
function addMetricCard(slide, colors, x, y, w, h, value, label, fontAdj = 0, options = {}) {
  const { bgColor = null, valueColor = null, valueSize = 'medium' } = options;
  
  // Card background
  slide.addShape('rect', {
    x, y, w, h,
    fill: { color: bgColor || colors.lightBg },
    line: { color: colors.border, width: 0.5 },
    rectRadius: 0.06
  });
  
  // Value size based on option
  const valueFontSize = valueSize === 'large' ? DESIGN.fonts.metric : 
                        valueSize === 'small' ? DESIGN.fonts.metricSmall : 
                        DESIGN.fonts.metricMedium;
  
  // Value
  slide.addText(String(value), {
    x: x + 0.08, y: y + 0.08, w: w - 0.16, h: h * 0.55,
    fontSize: adjustedFont(valueFontSize, fontAdj),
    bold: true,
    color: valueColor || colors.primary,
    fontFace: 'Arial',
    valign: 'middle'
  });
  
  // Label
  slide.addText(label, {
    x: x + 0.08, y: y + h * 0.58, w: w - 0.16, h: h * 0.38,
    fontSize: adjustedFont(DESIGN.fonts.metricLabel, fontAdj),
    color: colors.textLight,
    fontFace: 'Arial',
    valign: 'top'
  });
}

// ============================================================================
// CHART HELPER - addChartByType() - v7.1 UNIFIED CHART RENDERING
// ============================================================================

/**
 * Universal chart rendering function based on AI recommendations
 * @param {Object} slide - PptxGenJS slide object
 * @param {Object} colors - Theme colors
 * @param {string} chartType - Type of chart (bar, pie, donut, progress, timeline, stacked-bar)
 * @param {Object} chartConfig - Configuration including position, size, data
 */
function addChartByType(slide, colors, chartType, chartConfig) {
  const { x, y, w, h, data, options = {} } = chartConfig;
  
  if (!data || (Array.isArray(data) && data.length === 0)) return;
  
  switch (chartType) {
    case 'bar':
      addBarChart(slide, colors, x, y, w, h, data, options);
      break;
    case 'pie':
      addPieDonutChart(slide, colors, x, y, Math.min(w, h), data, { ...options, type: 'pie' });
      break;
    case 'donut':
      addPieDonutChart(slide, colors, x, y, Math.min(w, h), data, { ...options, type: 'doughnut' });
      break;
    case 'progress':
      addProgressBars(slide, colors, x, y, w, h, data, options);
      break;
    case 'timeline':
      addTimeline(slide, colors, x, y, w, h, data, options);
      break;
    case 'stacked-bar':
      addStackedBar(slide, colors, x, y, w, h, data, options);
      break;
    default:
      // No chart
      break;
  }
}

/**
 * Bar chart for revenue growth, comparisons
 */
function addBarChart(slide, colors, x, y, w, h, data, options = {}) {
  const { title = null, showValues = true, showLabels = true, showCAGR = true, fontAdj = 0 } = options;
  
  if (!data || data.length === 0) return;
  
  const maxValue = Math.max(...data.map(d => d.value || 0), 1);
  const barCount = data.length;
  const chartPadding = 0.1;
  const chartWidth = w - (chartPadding * 2);
  const chartHeight = h - 0.55;
  const barGap = 0.1;
  const totalGaps = (barCount - 1) * barGap;
  const barWidth = Math.min(0.65, (chartWidth - totalGaps) / barCount);
  const totalBarsWidth = (barWidth * barCount) + totalGaps;
  const startX = x + chartPadding + (chartWidth - totalBarsWidth) / 2;
  const chartBottom = y + h - 0.3;
  
  // Title
  if (title) {
    slide.addText(title, {
      x: x, y: y, w: w, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  }
  
  // CAGR
  if (showCAGR && data.length >= 2) {
    const cagr = calculateCAGR(data[0].value, data[data.length - 1].value, data.length - 1);
    if (cagr !== null && cagr > 0) {
      slide.addText(`CAGR: ${cagr}%`, {
        x: x + w - 1.0, y: y, w: 0.95, h: 0.25,
        fontSize: adjustedFont(DESIGN.fonts.chartLabel, fontAdj),
        bold: true,
        color: colors.accent,
        fontFace: 'Arial',
        align: 'right'
      });
    }
  }
  
  // Draw bars
  data.forEach((item, idx) => {
    const barHeight = Math.max(0.1, (item.value / maxValue) * chartHeight);
    const barX = startX + (idx * (barWidth + barGap));
    const barY = chartBottom - barHeight;
    
    // Bar
    slide.addShape('rect', {
      x: barX,
      y: barY,
      w: barWidth,
      h: barHeight,
      fill: { color: item.projected ? colors.secondary : colors.primary },
      rectRadius: 0.02
    });
    
    // Value on top
    if (showValues) {
      slide.addText(`${item.value}`, {
        x: barX - 0.1, y: barY - 0.25, w: barWidth + 0.2, h: 0.22,
        fontSize: adjustedFont(DESIGN.fonts.chartLabel, fontAdj),
        bold: true,
        color: colors.text,
        fontFace: 'Arial',
        align: 'center'
      });
    }
    
    // Label below
    if (showLabels) {
      slide.addText(item.label || '', {
        x: barX - 0.1, y: chartBottom + 0.02, w: barWidth + 0.2, h: 0.22,
        fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
        color: colors.textLight,
        fontFace: 'Arial',
        align: 'center'
      });
    }
  });
}

/**
 * Pie/Donut chart for composition
 */
function addPieDonutChart(slide, colors, x, y, size, data, options = {}) {
  const { type = 'doughnut', title = null, showLegend = true, fontAdj = 0 } = options;
  
  if (!data || data.length === 0) return;
  
  // Title
  if (title) {
    slide.addText(title, {
      x: x, y: y - 0.3, w: size + 0.5, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  }
  
  const chartData = data.map((d, idx) => ({
    name: d.label || `Item ${idx + 1}`,
    labels: [d.label || ''],
    values: [d.value || 0]
  }));
  
  const chartColors = data.map((d, idx) => d.color || colors.chartColors[idx % colors.chartColors.length]);
  
  const chartSize = showLegend ? size * 0.7 : size;
  
  slide.addChart(type, chartData, {
    x: x,
    y: y,
    w: chartSize,
    h: chartSize,
    showLegend: false,
    showTitle: false,
    holeSize: type === 'doughnut' ? 55 : 0,
    chartColors: chartColors
  });
  
  // Custom legend
  if (showLegend) {
    data.slice(0, 5).forEach((item, idx) => {
      const ly = y + 0.05 + (idx * 0.32);
      slide.addShape('rect', {
        x: x + chartSize + 0.1, y: ly + 0.06, w: 0.18, h: 0.18,
        fill: { color: chartColors[idx] }
      });
      slide.addText(`${truncateText(item.label, 12)} ${item.value}%`, {
        x: x + chartSize + 0.32, y: ly, w: size * 0.6, h: 0.32,
        fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
        color: colors.text,
        fontFace: 'Arial',
        valign: 'middle'
      });
    });
  }
}

/**
 * Progress bars for percentages/margins
 */
function addProgressBars(slide, colors, x, y, w, h, data, options = {}) {
  const { title = null, fontAdj = 0 } = options;
  
  if (!data || data.length === 0) return;
  
  if (title) {
    slide.addText(title, {
      x: x, y: y - 0.3, w: w, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  }
  
  const barHeight = 0.18;
  const spacing = (h - (data.length * barHeight)) / (data.length + 1);
  
  data.forEach((item, idx) => {
    const barY = y + spacing + (idx * (barHeight + spacing));
    const percentage = Math.min(100, Math.max(0, item.value || 0));
    const fillWidth = (percentage / 100) * w;
    
    // Label and value
    slide.addText(item.label || '', {
      x: x, y: barY - 0.24, w: w * 0.7, h: 0.22,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
    slide.addText(`${Math.round(percentage)}%`, {
      x: x + w - 0.6, y: barY - 0.24, w: 0.6, h: 0.22,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial',
      align: 'right'
    });
    
    // Background bar
    slide.addShape('rect', {
      x: x, y: barY, w: w, h: barHeight,
      fill: { color: colors.lightBg },
      line: { color: colors.border, width: 0.5 },
      rectRadius: barHeight / 2
    });
    
    // Fill bar
    if (fillWidth > 0.05) {
      slide.addShape('rect', {
        x: x, y: barY, w: fillWidth, h: barHeight,
        fill: { color: item.color || colors.chartColors[idx % colors.chartColors.length] },
        rectRadius: barHeight / 2
      });
    }
  });
}

/**
 * Timeline for milestones/roadmap
 */
function addTimeline(slide, colors, x, y, w, h, data, options = {}) {
  const { title = null, fontAdj = 0 } = options;
  
  if (!data || data.length === 0) return;
  
  if (title) {
    slide.addText(title, {
      x: x, y: y - 0.3, w: w, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  }
  
  const lineY = y + h / 2;
  const eventCount = Math.min(data.length, 5);
  const spacing = w / (eventCount + 1);
  
  // Main line
  slide.addShape('rect', {
    x: x, y: lineY - 0.02, w: w, h: 0.04,
    fill: { color: colors.primary }
  });
  
  // Events
  data.slice(0, eventCount).forEach((event, idx) => {
    const eventX = x + spacing * (idx + 1);
    
    // Circle marker
    slide.addShape('ellipse', {
      x: eventX - 0.12, y: lineY - 0.12, w: 0.24, h: 0.24,
      fill: { color: colors.accent },
      line: { color: colors.white, width: 2 }
    });
    
    // Year/date above
    slide.addText(event.year || event.date || '', {
      x: eventX - 0.5, y: lineY - 0.5, w: 1, h: 0.28,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial',
      align: 'center'
    });
    
    // Title below
    slide.addText(truncateText(event.title || '', 18), {
      x: eventX - 0.6, y: lineY + 0.18, w: 1.2, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
      color: colors.text,
      fontFace: 'Arial',
      align: 'center'
    });
  });
}

/**
 * Stacked bar for revenue breakdown
 */
function addStackedBar(slide, colors, x, y, w, h, data, options = {}) {
  const { title = null, fontAdj = 0 } = options;
  
  if (!data || data.length === 0) return;
  
  if (title) {
    slide.addText(title, {
      x: x, y: y - 0.3, w: w, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  }
  
  let currentX = x;
  
  data.forEach((item, idx) => {
    const pctMatch = String(item.value || item.pct || '0').match(/(\d+)/);
    const pct = pctMatch ? parseInt(pctMatch[1]) : 10;
    const barWidth = (pct / 100) * w;
    
    if (barWidth > 0.15) {
      slide.addShape('rect', {
        x: currentX, y: y, w: barWidth - 0.02, h: h,
        fill: { color: colors.chartColors[idx % 8] },
        rectRadius: 0.03
      });
      
      // Label inside if wide enough
      if (barWidth > 0.8) {
        slide.addText(`${truncateText(item.label || '', 15)} (${pct}%)`, {
          x: currentX + 0.05, y: y, w: barWidth - 0.1, h: h,
          fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
          color: colors.white,
          fontFace: 'Arial',
          valign: 'middle'
        });
      }
      
      currentX += barWidth;
    }
  });
}


// ============================================================================
// UNIVERSAL SLIDE CREATION WRAPPER - v7.1 CORE FEATURE
// ============================================================================

/**
 * Universal slide creation wrapper that:
 * 1. Calls analyzeDataForLayout() for AI recommendations
 * 2. Passes recommendations to the appropriate render function
 * 3. Returns the updated slide number
 * 
 * @param {string} slideType - Type of slide to create
 * @param {Object} pptx - PptxGenJS instance
 * @param {Object} colors - Theme colors
 * @param {Object} data - Form data
 * @param {number} slideNumber - Current slide number
 * @param {Object} context - Additional context (industryData, docConfig, etc.)
 * @returns {Promise<number>} Updated slide number
 */
async function createSlide(slideType, pptx, colors, data, slideNumber, context = {}) {
  // Get AI layout recommendations
  const layoutRec = await analyzeDataForLayout(data, slideType);
  
  // Get font adjustment
  const fontAdj = layoutRec.fontAdjustment || 0;
  
  // Create slide
  const slide = pptx.addSlide();
  
  // Switch to appropriate render function
  switch (slideType) {
    case 'title':
      renderTitleSlide(slide, colors, data, context.docConfig, fontAdj);
      break;
      
    case 'disclaimer':
      renderDisclaimerSlide(slide, colors, data, slideNumber, fontAdj);
      break;
      
    case 'toc':
      renderTOCSlide(slide, colors, data, slideNumber, context.slideList || [], fontAdj);
      break;
      
    case 'executive-summary':
      renderExecutiveSummarySlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'investment-highlights':
      renderInvestmentHighlightsSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'company-overview':
      renderCompanyOverviewSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'leadership':
      renderLeadershipSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'industry':
      renderIndustrySlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'services':
      renderServicesSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'clients':
      renderClientsSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'financials':
      renderFinancialsSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'case-study':
      renderCaseStudySlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'growth':
      renderGrowthSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'synergies':
      renderSynergiesSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'market-position':
      renderMarketPositionSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'risks':
      renderRisksSlide(slide, colors, data, slideNumber, layoutRec, context);
      break;
      
    case 'thank-you':
      renderThankYouSlide(slide, colors, data, context.docConfig, fontAdj);
      break;
      
    default:
      // Generic slide
      addSlideHeader(slide, colors, slideType, null, fontAdj);
      addSlideFooter(slide, colors, slideNumber);
      break;
  }
  
  return slideNumber + 1;
}

// ============================================================================
// DEDICATED RENDER FUNCTIONS - Each applies AI layout recommendations
// ============================================================================

/**
 * TITLE SLIDE
 */
function renderTitleSlide(slide, colors, data, docConfig, fontAdj = 0) {
  // Full dark background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  // Decorative shapes
  slide.addShape('rect', {
    x: 7.5, y: 0, w: 2.5, h: 2.2,
    fill: { color: colors.secondary },
    transparency: 75
  });
  slide.addShape('rect', {
    x: 8.5, y: 3.5, w: 1.5, h: 2.125,
    fill: { color: colors.accent },
    transparency: 80
  });
  
  // Accent line
  slide.addShape('rect', {
    x: 0.5, y: 3.05, w: 3.5, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  // Project codename
  slide.addText(data.projectCodename || 'Project Phoenix', {
    x: 0.5, y: 1.7, w: 8, h: 1.1,
    fontSize: 48,
    bold: true,
    color: colors.white,
    fontFace: 'Arial'
  });
  
  // Document type
  slide.addText(docConfig?.name || 'Management Presentation', {
    x: 0.5, y: 3.2, w: 6, h: 0.45,
    fontSize: 20,
    color: colors.white,
    fontFace: 'Arial'
  });
  
  // Date
  slide.addText(formatDate(data.presentationDate), {
    x: 0.5, y: 3.8, w: 4, h: 0.35,
    fontSize: 14,
    color: colors.white,
    fontFace: 'Arial',
    transparency: 25
  });
  
  // Advisor
  if (data.advisor) {
    slide.addText(`Prepared by ${data.advisor}`, {
      x: 0.5, y: 4.25, w: 4, h: 0.3,
      fontSize: 12,
      color: colors.white,
      fontFace: 'Arial',
      transparency: 35
    });
  }
  
  // Confidential notice
  slide.addText('Strictly Private and Confidential', {
    x: 0.5, y: 4.9, w: 4, h: 0.25,
    fontSize: 10,
    italic: true,
    color: colors.white,
    fontFace: 'Arial',
    transparency: 45
  });
}

/**
 * DISCLAIMER SLIDE
 */
function renderDisclaimerSlide(slide, colors, data, slideNumber, fontAdj = 0) {
  addSlideHeader(slide, colors, 'Important Notice', null, fontAdj);
  
  const advisor = data.advisor || 'the Advisor';
  const company = data.companyName || 'the Company';
  
  const disclaimerText = `This document has been prepared by ${advisor} exclusively for the benefit of the party to whom it is directly addressed and delivered. This document is strictly confidential and may not be reproduced, redistributed, or passed on to any other person or published, in whole or in part, for any purpose without the prior written consent of ${advisor}.

This document does not constitute or form part of, and should not be construed as, any offer, invitation, or inducement to purchase or subscribe for any securities, nor shall it or any part of it form the basis of, or be relied upon in connection with, any contract or commitment whatsoever.

The information contained herein has been prepared based upon information provided by ${company} and from sources believed to be reliable. However, no representation or warranty, express or implied, is made as to the accuracy, completeness, or reliability of this information.

Neither ${advisor}, nor any of its affiliates, directors, employees, or agents shall have any liability whatsoever for any direct, indirect, or consequential loss or damage arising from the use of this document or its contents.`;

  slide.addText(disclaimerText, {
    x: 0.4, y: 1.0, w: 9.2, h: 3.95,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.text,
    fontFace: 'Arial',
    valign: 'top',
    lineSpacingMultiple: 1.35
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * TABLE OF CONTENTS (for CIM)
 */
function renderTOCSlide(slide, colors, data, slideNumber, slideList, fontAdj = 0) {
  addSlideHeader(slide, colors, 'Table of Contents', null, fontAdj);
  
  const contentTop = 1.0;
  const itemsPerColumn = 12;
  
  slideList.forEach((item, idx) => {
    const col = Math.floor(idx / itemsPerColumn);
    const row = idx % itemsPerColumn;
    const x = 0.4 + (col * 4.7);
    const y = contentTop + (row * 0.33);
    
    slide.addText(`${idx + 1}.`, {
      x: x, y: y, w: 0.35, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial'
    });
    slide.addText(item.title || item, {
      x: x + 0.35, y: y, w: 4.0, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * EXECUTIVE SUMMARY - AI-POWERED LAYOUT
 */
function renderExecutiveSummarySlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const chartType = layoutRec.chartType || 'bar';
  const layout = layoutRec.layout || 'two-column';
  
  addSlideHeader(slide, colors, 'Executive Summary', data.companyName || '', fontAdj);
  
  const contentTop = 0.95;
  const contentHeight = 4.1;
  
  if (layout === 'two-column' || layout === 'two-column-wide-left') {
    // LEFT COLUMN - Key Metrics
    const leftWidth = layout === 'two-column-wide-left' ? 5.2 : 4.5;
    addSectionBox(slide, colors, 0.3, contentTop, leftWidth, contentHeight, 'Key Highlights', colors.primary, fontAdj);
    
    // Build metrics
    const metrics = [];
    if (data.foundedYear) metrics.push({ value: data.foundedYear, label: 'Founded' });
    if (data.headquarters) metrics.push({ value: truncateText(data.headquarters, 18), label: 'Headquarters' });
    if (data.employeeCountFT) metrics.push({ value: `${data.employeeCountFT}+`, label: 'Employees' });
    if (data.revenueFY25) metrics.push({ value: formatCurrency(data.revenueFY25, data.currency), label: 'Revenue FY25' });
    if (data.ebitdaMarginFY25) metrics.push({ value: `${data.ebitdaMarginFY25}%`, label: 'EBITDA Margin' });
    if (data.netRetention) metrics.push({ value: `${data.netRetention}%`, label: 'Net Retention' });
    
    // Display in 2x3 grid
    metrics.slice(0, 6).forEach((metric, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const mx = 0.42 + (col * 2.15);
      const my = contentTop + 0.48 + (row * 1.15);
      
      addMetricCard(slide, colors, mx, my, 2.0, 1.0, metric.value, metric.label, fontAdj);
    });
    
    // RIGHT COLUMN - Description + Chart
    const rightX = 0.3 + leftWidth + 0.2;
    const rightWidth = 9.4 - leftWidth - 0.2;
    
    // Company Description Box
    addSectionBox(slide, colors, rightX, contentTop, rightWidth, 1.5, 'About the Company', colors.secondary, fontAdj);
    
    const description = data.companyDescription || 'A leading technology solutions provider.';
    slide.addText(truncateDescription(description, 200), {
      x: rightX + 0.12, y: contentTop + 0.45, w: rightWidth - 0.24, h: 0.95,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial',
      valign: 'top'
    });
    
    // Revenue Chart
    addSectionBox(slide, colors, rightX, contentTop + 1.65, rightWidth, 2.45, 'Revenue Growth', colors.accent, fontAdj);
    
    // Build revenue data
    const revenueData = [];
    if (data.revenueFY24) revenueData.push({ label: 'FY24', value: parseFloat(data.revenueFY24), projected: false });
    if (data.revenueFY25) revenueData.push({ label: 'FY25', value: parseFloat(data.revenueFY25), projected: false });
    if (data.revenueFY26P) revenueData.push({ label: 'FY26P', value: parseFloat(data.revenueFY26P), projected: true });
    if (data.revenueFY27P) revenueData.push({ label: 'FY27P', value: parseFloat(data.revenueFY27P), projected: true });
    
    if (revenueData.length > 0) {
      // Currency label
      slide.addText(`In ${data.currency === 'USD' ? 'USD Mn' : 'INR Cr'}`, {
        x: rightX + 0.12, y: contentTop + 2.0, w: 1.2, h: 0.22,
        fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
        italic: true,
        color: colors.textLight,
        fontFace: 'Arial'
      });
      
      addChartByType(slide, colors, chartType, {
        x: rightX + 0.15,
        y: contentTop + 2.2,
        w: rightWidth - 0.3,
        h: 1.8,
        data: revenueData,
        options: { fontAdj, showCAGR: true }
      });
    }
  }
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * INVESTMENT HIGHLIGHTS - AI-POWERED LAYOUT
 */
function renderInvestmentHighlightsSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const layout = layoutRec.layout || 'grid-2x2';
  
  addSlideHeader(slide, colors, 'Investment Highlights', 'Key reasons to invest', fontAdj);
  
  const contentTop = 0.95;
  let highlights = parseLines(data.investmentHighlights, 8);
  
  if (highlights.length === 0) {
    highlights = [
      'Strong revenue growth trajectory with proven scalability',
      'Experienced management team with deep domain expertise',
      'Diversified and loyal client base with high retention',
      'Scalable technology platform with proprietary IP',
      'Market leadership in key segments',
      'Clear path to continued expansion'
    ];
  }
  
  // Display in 2-column grid with numbered cards
  const boxWidth = 4.5;
  const boxHeight = 0.88;
  const colGap = 0.2;
  const rowGap = 0.12;
  
  highlights.slice(0, 8).forEach((highlight, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = 0.3 + (col * (boxWidth + colGap));
    const y = contentTop + (row * (boxHeight + rowGap));
    
    // Box background
    slide.addShape('rect', {
      x, y, w: boxWidth, h: boxHeight,
      fill: { color: col === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 },
      rectRadius: 0.04
    });
    
    // Number badge
    slide.addShape('ellipse', {
      x: x + 0.12, y: y + 0.24, w: 0.4, h: 0.4,
      fill: { color: colors.primary }
    });
    slide.addText(`${idx + 1}`, {
      x: x + 0.12, y: y + 0.24, w: 0.4, h: 0.4,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.white,
      fontFace: 'Arial',
      align: 'center',
      valign: 'middle'
    });
    
    // Highlight text
    slide.addText(truncateText(highlight, 75), {
      x: x + 0.62, y: y + 0.12, w: boxWidth - 0.82, h: boxHeight - 0.24,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial',
      valign: 'middle'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * COMPANY OVERVIEW (for CIM)
 */
function renderCompanyOverviewSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const industryData = context.industryData;
  
  addSlideHeader(slide, colors, 'Company Overview', 'At a Glance', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Company snapshot
  addSectionBox(slide, colors, 0.3, contentTop, 4.2, 2.8, 'Company Profile', colors.primary, fontAdj);
  
  const profileItems = [
    { label: 'Legal Name', value: data.companyName || 'Company Name' },
    { label: 'Founded', value: data.foundedYear || 'N/A' },
    { label: 'Headquarters', value: data.headquarters || 'N/A' },
    { label: 'Employees (FT)', value: data.employeeCountFT || 'N/A' },
    { label: 'Total Workforce', value: data.employeeCountTotal || 'N/A' },
    { label: 'Primary Vertical', value: industryData?.name || 'Technology' }
  ];
  
  profileItems.forEach((item, idx) => {
    const iy = contentTop + 0.48 + (idx * 0.38);
    slide.addText(item.label, {
      x: 0.42, y: iy, w: 1.6, h: 0.34,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
    slide.addText(truncateText(String(item.value), 22), {
      x: 2.05, y: iy, w: 2.3, h: 0.34,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // RIGHT - Key services and partnerships
  addSectionBox(slide, colors, 4.7, contentTop, 4.9, 1.3, 'Core Services', colors.secondary, fontAdj);
  
  const services = parsePipeSeparated(data.serviceLines, 4);
  services.forEach((srv, idx) => {
    slide.addText(`• ${truncateText(srv[0] || '', 40)}`, {
      x: 4.85, y: contentTop + 0.48 + (idx * 0.28), w: 4.6, h: 0.26,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // Technology Partners
  addSectionBox(slide, colors, 4.7, contentTop + 1.45, 4.9, 1.35, 'Technology Partners', colors.accent, fontAdj);
  
  const partners = parseLines(data.techPartnerships, 4);
  partners.forEach((partner, idx) => {
    slide.addText(`✓ ${truncateText(partner, 35)}`, {
      x: 4.85, y: contentTop + 1.93 + (idx * 0.28), w: 4.6, h: 0.26,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // BOTTOM - Description
  addSectionBox(slide, colors, 0.3, contentTop + 2.95, 9.3, 1.15, 'About', colors.primary, fontAdj);
  
  slide.addText(truncateDescription(data.companyDescription || 'Company description.', 300), {
    x: 0.42, y: contentTop + 3.38, w: 9.06, h: 0.65,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.text,
    fontFace: 'Arial',
    valign: 'top'
  });
  
  addSlideFooter(slide, colors, slideNumber);
}


/**
 * LEADERSHIP SLIDE
 */
function renderLeadershipSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  
  addSlideHeader(slide, colors, 'Leadership Team', 'Experienced management driving growth', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Founder Profile
  const founderWidth = 4.4;
  addSectionBox(slide, colors, 0.3, contentTop, founderWidth, 4.1, 'Founder', colors.primary, fontAdj);
  
  // Photo placeholder
  slide.addShape('ellipse', {
    x: 1.55, y: contentTop + 0.55, w: 1.5, h: 1.5,
    fill: { color: colors.white },
    line: { color: colors.primary, width: 3 }
  });
  slide.addText('Photo', {
    x: 1.55, y: contentTop + 1.1, w: 1.5, h: 0.4,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.textLight,
    fontFace: 'Arial',
    align: 'center'
  });
  
  // Founder name and title
  slide.addText(data.founderName || 'Founder Name', {
    x: 0.42, y: contentTop + 2.2, w: founderWidth - 0.24, h: 0.4,
    fontSize: 18,
    bold: true,
    color: colors.primary,
    fontFace: 'Arial',
    align: 'center'
  });
  
  slide.addText(data.founderTitle || 'Founder & CEO', {
    x: 0.42, y: contentTop + 2.6, w: founderWidth - 0.24, h: 0.3,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.textLight,
    fontFace: 'Arial',
    align: 'center'
  });
  
  // Experience and education
  const founderInfo = [];
  if (data.founderExperience) founderInfo.push(`${data.founderExperience}+ years industry experience`);
  parseLines(data.founderEducation, 2).forEach(edu => founderInfo.push(edu));
  
  founderInfo.slice(0, 3).forEach((info, idx) => {
    slide.addText(`• ${truncateText(info, 45)}`, {
      x: 0.5, y: contentTop + 3.0 + (idx * 0.32), w: founderWidth - 0.4, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // RIGHT - Leadership Team Grid
  const rightX = 4.9;
  const rightWidth = 4.8;
  addSectionBox(slide, colors, rightX, contentTop, rightWidth, 4.1, 'Leadership Team', colors.secondary, fontAdj);
  
  const leaders = parsePipeSeparated(data.leadershipTeam, 6);
  
  leaders.forEach((leader, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const lx = rightX + 0.12 + (col * 2.32);
    const ly = contentTop + 0.48 + (row * 1.15);
    
    // Leader card
    slide.addShape('rect', {
      x: lx, y: ly, w: 2.2, h: 1.0,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 },
      rectRadius: 0.04
    });
    
    slide.addText(truncateText(leader[0] || 'Name', 22), {
      x: lx + 0.1, y: ly + 0.12, w: 2.0, h: 0.38,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial'
    });
    
    slide.addText(truncateText(leader[1] || 'Title', 28), {
      x: lx + 0.1, y: ly + 0.52, w: 2.0, h: 0.38,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * INDUSTRY OVERVIEW SLIDE
 */
function renderIndustrySlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const industryData = context.industryData || INDUSTRY_DATA.technology;
  
  addSlideHeader(slide, colors, industryData.fullName || 'Industry Overview', 'Market Analysis', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Industry Benchmarks
  addSectionBox(slide, colors, 0.3, contentTop, 4.6, 2.0, 'Industry Benchmarks', colors.primary, fontAdj);
  
  const benchmarks = [
    { label: 'Average Growth Rate', value: industryData.benchmarks?.avgGrowthRate || '15-25%' },
    { label: 'Average EBITDA Margin', value: industryData.benchmarks?.avgEbitdaMargin || '20-30%' },
    { label: 'Typical Deal Multiple', value: industryData.benchmarks?.avgDealMultiple || '8-12x EBITDA' },
    { label: 'Market Size', value: industryData.benchmarks?.marketSize || '$100B+ globally' }
  ];
  
  benchmarks.forEach((bm, idx) => {
    slide.addText(bm.label, {
      x: 0.42, y: contentTop + 0.48 + (idx * 0.38), w: 2.2, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
    slide.addText(bm.value, {
      x: 2.7, y: contentTop + 0.48 + (idx * 0.38), w: 2.0, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial',
      align: 'right'
    });
  });
  
  // RIGHT - Key Metrics
  addSectionBox(slide, colors, 5.1, contentTop, 4.5, 2.0, 'Key Industry Metrics', colors.secondary, fontAdj);
  
  (industryData.keyMetrics || []).slice(0, 5).forEach((metric, idx) => {
    slide.addText(`• ${metric}`, {
      x: 5.22, y: contentTop + 0.48 + (idx * 0.3), w: 4.2, h: 0.28,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // BOTTOM LEFT - Market Drivers
  addSectionBox(slide, colors, 0.3, contentTop + 2.15, 4.6, 1.9, 'Market Drivers', colors.accent, fontAdj);
  
  (industryData.keyDrivers || []).slice(0, 5).forEach((driver, idx) => {
    slide.addText(`▸ ${truncateText(driver, 45)}`, {
      x: 0.42, y: contentTop + 2.6 + (idx * 0.32), w: 4.4, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // BOTTOM RIGHT - Regulatory Environment
  addSectionBox(slide, colors, 5.1, contentTop + 2.15, 4.5, 1.9, 'Regulatory Environment', colors.primary, fontAdj);
  
  (industryData.regulations || []).slice(0, 5).forEach((reg, idx) => {
    slide.addText(`• ${truncateText(reg, 40)}`, {
      x: 5.22, y: contentTop + 2.6 + (idx * 0.32), w: 4.2, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * SERVICES SLIDE - AI-POWERED with PIE/DONUT CHART
 */
function renderServicesSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const chartType = layoutRec.chartType || 'donut';
  
  addSlideHeader(slide, colors, 'Services & Capabilities', 'Core offerings and solutions', fontAdj);
  
  const contentTop = 0.95;
  const services = parsePipeSeparated(data.serviceLines, 6);
  
  // LEFT - Service Cards
  const leftWidth = 5.4;
  addSectionBox(slide, colors, 0.3, contentTop, leftWidth, 2.7, 'Service Lines', colors.primary, fontAdj);
  
  services.slice(0, 4).forEach((service, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const sx = 0.42 + (col * 2.65);
    const sy = contentTop + 0.48 + (row * 1.05);
    
    const name = service[0] || 'Service';
    const pct = service[1] || '';
    const desc = service[2] || '';
    
    // Service card
    slide.addShape('rect', {
      x: sx, y: sy, w: 2.52, h: 0.92,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 },
      rectRadius: 0.04
    });
    
    // Service name
    slide.addText(truncateText(name, 28), {
      x: sx + 0.1, y: sy + 0.08, w: 1.85, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial'
    });
    
    // Percentage badge
    if (pct) {
      slide.addShape('rect', {
        x: sx + 1.98, y: sy + 0.1, w: 0.45, h: 0.28,
        fill: { color: colors.accent },
        rectRadius: 0.1
      });
      slide.addText(pct, {
        x: sx + 1.98, y: sy + 0.1, w: 0.45, h: 0.28,
        fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
        bold: true,
        color: colors.white,
        fontFace: 'Arial',
        align: 'center',
        valign: 'middle'
      });
    }
    
    // Description
    slide.addText(truncateText(desc, 50), {
      x: sx + 0.1, y: sy + 0.48, w: 2.32, h: 0.38,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial',
      valign: 'top'
    });
  });
  
  // RIGHT - PIE/DONUT CHART for Revenue Mix
  const rightX = 5.9;
  const rightWidth = 3.8;
  addSectionBox(slide, colors, rightX, contentTop, rightWidth, 2.7, 'Revenue Mix', colors.secondary, fontAdj);
  
  const pieData = services.slice(0, 4).map((srv, idx) => {
    const pctMatch = (srv[1] || '25').match(/(\d+)/);
    return {
      label: truncateText(srv[0] || 'Service', 14),
      value: pctMatch ? parseInt(pctMatch[1]) : 25,
      color: colors.chartColors[idx]
    };
  });
  
  if (pieData.length > 0) {
    addChartByType(slide, colors, chartType, {
      x: rightX + 0.15,
      y: contentTop + 0.5,
      w: rightWidth - 0.3,
      h: 2.0,
      data: pieData,
      options: { fontAdj, showLegend: true }
    });
  }
  
  // BOTTOM - Proprietary Products
  const products = parsePipeSeparated(data.products, 3);
  if (products.length > 0) {
    addSectionBox(slide, colors, 0.3, contentTop + 2.85, 9.3, 1.2, 'Proprietary Products', colors.accent, fontAdj);
    
    products.forEach((product, idx) => {
      const px = 0.42 + (idx * 3.08);
      
      slide.addText(truncateText(product[0] || 'Product', 28), {
        x: px, y: contentTop + 3.3, w: 2.95, h: 0.32,
        fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
        bold: true,
        color: colors.text,
        fontFace: 'Arial'
      });
      
      slide.addText(truncateText(product[1] || '', 45), {
        x: px, y: contentTop + 3.62, w: 2.95, h: 0.32,
        fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
        color: colors.textLight,
        fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * CLIENTS SLIDE - AI-POWERED with DONUT CHART
 */
function renderClientsSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const chartType = layoutRec.chartType || 'donut';
  const docConfig = context.docConfig || {};
  
  addSlideHeader(slide, colors, 'Client Portfolio & Vertical Mix', 'Diversified customer base', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Client Metrics + Vertical Mix Donut
  addSectionBox(slide, colors, 0.3, contentTop, 3.4, 1.5, 'Client Metrics', colors.primary, fontAdj);
  
  const clientMetrics = [
    { label: 'Top 10 Concentration', value: data.topTenConcentration ? `${data.topTenConcentration}%` : '58%' },
    { label: 'Net Revenue Retention', value: data.netRetention ? `${data.netRetention}%` : '120%' },
    { label: 'Primary Vertical', value: context.industryData?.name || 'BFSI' }
  ];
  
  clientMetrics.forEach((metric, idx) => {
    slide.addText(metric.label, {
      x: 0.42, y: contentTop + 0.5 + (idx * 0.34), w: 1.9, h: 0.32,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
    slide.addText(metric.value, {
      x: 2.35, y: contentTop + 0.5 + (idx * 0.34), w: 1.2, h: 0.32,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial',
      align: 'right'
    });
  });
  
  // Vertical Mix Donut Chart
  addSectionBox(slide, colors, 0.3, contentTop + 1.65, 3.4, 2.4, 'Vertical Mix', colors.secondary, fontAdj);
  
  // Parse verticals from otherVerticals field
  const verticals = [];
  const primaryVertical = context.industryData?.name || 'BFSI';
  const primaryPct = data.primaryVerticalPct ? parseInt(data.primaryVerticalPct) : 60;
  verticals.push({ label: primaryVertical, value: primaryPct });
  
  if (data.otherVerticals) {
    parsePipeSeparated(data.otherVerticals, 4).forEach(v => {
      const pctMatch = (v[1] || '10').match(/(\d+)/);
      verticals.push({ label: v[0] || 'Other', value: pctMatch ? parseInt(pctMatch[1]) : 10 });
    });
  } else {
    verticals.push({ label: 'FinTech', value: 15 });
    verticals.push({ label: 'Healthcare', value: 15 });
    verticals.push({ label: 'Retail', value: 10 });
  }
  
  addChartByType(slide, colors, 'donut', {
    x: 0.42,
    y: contentTop + 2.05,
    w: 3.1,
    h: 1.9,
    data: verticals.slice(0, 5),
    options: { fontAdj, showLegend: true }
  });
  
  // RIGHT - Key Clients Grid
  const rightX = 3.9;
  const rightWidth = 5.8;
  addSectionBox(slide, colors, rightX, contentTop, rightWidth, 4.05, 'Key Clients', colors.accent, fontAdj);
  
  const clients = parsePipeSeparated(data.topClients, 12);
  
  // Display clients in a 3x4 grid
  clients.slice(0, 12).forEach((client, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const cx = rightX + 0.12 + (col * 1.9);
    const cy = contentTop + 0.48 + (row * 0.88);
    
    // Client card
    slide.addShape('rect', {
      x: cx, y: cy, w: 1.8, h: 0.78,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 },
      rectRadius: 0.04
    });
    
    // Client name (show or anonymize based on docConfig)
    const clientName = docConfig.includeClientNames ? (client[0] || 'Client') : `Client ${idx + 1}`;
    slide.addText(truncateText(clientName, 18), {
      x: cx + 0.08, y: cy + 0.12, w: 1.64, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.text,
      fontFace: 'Arial',
      align: 'center'
    });
    
    // Since year
    if (client[1]) {
      slide.addText(`Since ${client[1]}`, {
        x: cx + 0.08, y: cy + 0.48, w: 1.64, h: 0.25,
        fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
        color: colors.textLight,
        fontFace: 'Arial',
        align: 'center'
      });
    }
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * FINANCIALS SLIDE - AI-POWERED with BAR CHART + PROGRESS BARS
 */
function renderFinancialsSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const chartType = layoutRec.chartType || 'bar';
  
  addSlideHeader(slide, colors, 'Financial Performance', 'Revenue growth and key metrics', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Revenue Growth Bar Chart
  addSectionBox(slide, colors, 0.3, contentTop, 4.8, 2.6, 'Revenue Growth', colors.primary, fontAdj);
  
  // Currency label
  slide.addText(`In ${data.currency === 'USD' ? 'USD Mn' : 'INR Cr'}`, {
    x: 0.42, y: contentTop + 0.48, w: 1.2, h: 0.22,
    fontSize: adjustedFont(DESIGN.fonts.caption, fontAdj),
    italic: true,
    color: colors.textLight,
    fontFace: 'Arial'
  });
  
  // Build revenue data including projections
  const revenueData = [];
  if (data.revenueFY24) revenueData.push({ label: 'FY24', value: parseFloat(data.revenueFY24), projected: false });
  if (data.revenueFY25) revenueData.push({ label: 'FY25', value: parseFloat(data.revenueFY25), projected: false });
  if (data.revenueFY26P) revenueData.push({ label: 'FY26P', value: parseFloat(data.revenueFY26P), projected: true });
  if (data.revenueFY27P) revenueData.push({ label: 'FY27P', value: parseFloat(data.revenueFY27P), projected: true });
  if (data.revenueFY28P) revenueData.push({ label: 'FY28P', value: parseFloat(data.revenueFY28P), projected: true });
  
  addChartByType(slide, colors, 'bar', {
    x: 0.42,
    y: contentTop + 0.72,
    w: 4.55,
    h: 1.75,
    data: revenueData,
    options: { fontAdj, showCAGR: true }
  });
  
  // RIGHT - Key Margins with Progress Bars
  addSectionBox(slide, colors, 5.3, contentTop, 4.4, 2.6, 'Key Margins & Metrics', colors.secondary, fontAdj);
  
  const margins = [];
  if (data.ebitdaMarginFY25) margins.push({ label: 'EBITDA Margin FY25', value: parseFloat(data.ebitdaMarginFY25), color: colors.primary });
  if (data.grossMargin) margins.push({ label: 'Gross Margin', value: parseFloat(data.grossMargin), color: colors.secondary });
  if (data.netProfitMargin) margins.push({ label: 'Net Profit Margin', value: parseFloat(data.netProfitMargin), color: colors.accent });
  if (data.netRetention) margins.push({ label: 'Net Revenue Retention', value: parseFloat(data.netRetention), color: colors.chartColors[3] });
  if (data.topTenConcentration) margins.push({ label: 'Top 10 Concentration', value: parseFloat(data.topTenConcentration), color: colors.chartColors[4] });
  
  if (margins.length > 0) {
    addChartByType(slide, colors, 'progress', {
      x: 5.42,
      y: contentTop + 0.65,
      w: 4.1,
      h: 1.85,
      data: margins.slice(0, 5),
      options: { fontAdj }
    });
  }
  
  // BOTTOM - Revenue by Service Line (Stacked Bar)
  addSectionBox(slide, colors, 0.3, contentTop + 2.75, 9.4, 1.3, 'Revenue by Service Line', colors.accent, fontAdj);
  
  const services = parsePipeSeparated(data.serviceLines, 5);
  const serviceRevenue = services.map(srv => {
    const pctMatch = (srv[1] || '20').match(/(\d+)/);
    return { label: srv[0] || 'Service', value: pctMatch ? parseInt(pctMatch[1]) : 20 };
  });
  
  if (serviceRevenue.length > 0) {
    addChartByType(slide, colors, 'stacked-bar', {
      x: 0.42,
      y: contentTop + 3.22,
      w: 9.15,
      h: 0.4,
      data: serviceRevenue,
      options: { fontAdj }
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
}


/**
 * CASE STUDY SLIDE - FULL WIDTH LAYOUT
 */
function renderCaseStudySlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const caseStudy = context.caseStudy || {};
  
  const clientName = caseStudy.client || 'Client Name';
  const industry = caseStudy.industry || '';
  
  addSlideHeader(slide, colors, `Case Study: ${truncateText(clientName, 40)}`, industry, fontAdj);
  
  const contentTop = 0.95;
  
  // Client Info Box (Left)
  addSectionBox(slide, colors, 0.3, contentTop, 2.6, 2.2, clientName, colors.primary, fontAdj);
  
  slide.addText(industry || 'Industry', {
    x: 0.42, y: contentTop + 0.5, w: 2.36, h: 0.35,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.textLight,
    fontFace: 'Arial'
  });
  
  // Challenge Box
  addSectionBox(slide, colors, 3.1, contentTop, 3.2, 2.2, 'Challenge', colors.accent, fontAdj);
  
  slide.addText(truncateDescription(caseStudy.challenge || 'Client challenge description.', 200), {
    x: 3.22, y: contentTop + 0.48, w: 2.95, h: 1.6,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.text,
    fontFace: 'Arial',
    valign: 'top'
  });
  
  // Solution Box
  addSectionBox(slide, colors, 6.5, contentTop, 3.2, 2.2, 'Solution', colors.secondary, fontAdj);
  
  slide.addText(truncateDescription(caseStudy.solution || 'Solution implemented.', 200), {
    x: 6.62, y: contentTop + 0.48, w: 2.95, h: 1.6,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    color: colors.text,
    fontFace: 'Arial',
    valign: 'top'
  });
  
  // Results Box (Full Width)
  addSectionBox(slide, colors, 0.3, contentTop + 2.35, 9.4, 1.7, 'Key Results & Impact', colors.primary, fontAdj);
  
  const results = parseLines(caseStudy.results || 'Key results achieved', 6);
  
  // Display results in 2 columns
  results.slice(0, 6).forEach((result, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const rx = 0.42 + (col * 4.7);
    const ry = contentTop + 2.82 + (row * 0.38);
    
    slide.addText(`✓ ${truncateText(result, 55)}`, {
      x: rx, y: ry, w: 4.5, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * GROWTH STRATEGY SLIDE - with TIMELINE
 */
function renderGrowthSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const chartType = layoutRec.chartType || 'timeline';
  
  addSlideHeader(slide, colors, 'Growth Strategy & Roadmap', 'Path to continued expansion', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Key Growth Drivers
  addSectionBox(slide, colors, 0.3, contentTop, 4.6, 2.0, 'Key Growth Drivers', colors.primary, fontAdj);
  
  let drivers = parseLines(data.growthDrivers, 5);
  if (drivers.length === 0) {
    drivers = [
      'AI adoption accelerating across enterprise clients',
      'Cloud migration spend growing 25% annually',
      'Increasing demand for managed services',
      'Digital transformation mandates from regulators',
      'Expansion into Middle East and Southeast Asia'
    ];
  }
  
  drivers.slice(0, 5).forEach((driver, idx) => {
    slide.addText(`▸ ${truncateText(driver, 55)}`, {
      x: 0.42, y: contentTop + 0.48 + (idx * 0.3), w: 4.4, h: 0.28,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // RIGHT - Strategic Roadmap
  addSectionBox(slide, colors, 5.1, contentTop, 4.6, 2.0, 'Strategic Roadmap', colors.secondary, fontAdj);
  
  // Short-term goals
  slide.addText('0-12 Months', {
    x: 5.22, y: contentTop + 0.48, w: 1.5, h: 0.28,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    bold: true,
    color: colors.primary,
    fontFace: 'Arial'
  });
  
  parseLines(data.shortTermGoals, 2).forEach((goal, idx) => {
    slide.addText(`• ${truncateText(goal, 45)}`, {
      x: 5.22, y: contentTop + 0.78 + (idx * 0.26), w: 4.35, h: 0.24,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // Medium-term goals
  slide.addText('1-3 Years', {
    x: 5.22, y: contentTop + 1.35, w: 1.5, h: 0.28,
    fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
    bold: true,
    color: colors.primary,
    fontFace: 'Arial'
  });
  
  parseLines(data.mediumTermGoals, 2).forEach((goal, idx) => {
    slide.addText(`• ${truncateText(goal, 45)}`, {
      x: 5.22, y: contentTop + 1.65 + (idx * 0.26), w: 4.35, h: 0.24,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // BOTTOM - Competitive Advantages
  addSectionBox(slide, colors, 0.3, contentTop + 2.15, 9.4, 1.9, 'Competitive Advantages', colors.accent, fontAdj);
  
  let advantages = parseLines(data.competitiveAdvantages, 6);
  if (advantages.length === 0) {
    advantages = [
      'Deep AWS expertise: Only 8 companies in India with Advanced Partner status',
      'Proprietary AI platform: 500+ templates, 60% faster deployment',
      'Strong BFSI relationships: 10+ year partnerships with top banks',
      'High client retention: 137% NRR with zero churn in 3 years',
      'Experienced leadership: 100+ years combined experience',
      'Capital-light model: 25%+ EBITDA with minimal capex'
    ];
  }
  
  advantages.slice(0, 6).forEach((adv, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const ax = 0.42 + (col * 4.7);
    const ay = contentTop + 2.58 + (row * 0.42);
    
    slide.addText(`• ${truncateText(adv, 60)}`, {
      x: ax, y: ay, w: 4.5, h: 0.38,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * SYNERGIES SLIDE
 */
function renderSynergiesSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const buyerContent = context.buyerContent || {};
  
  addSlideHeader(slide, colors, 'Potential Synergies', 'Value creation opportunities', fontAdj);
  
  const contentTop = 0.95;
  
  // Determine which synergies to emphasize based on buyer type
  const emphasize = buyerContent.slideAdjustments?.synergies?.emphasize || 'both';
  const showStrategic = emphasize === 'strategic' || emphasize === 'both';
  const showFinancial = emphasize === 'financial' || emphasize === 'both';
  
  const colWidth = (showStrategic && showFinancial) ? 4.55 : 9.3;
  
  // Strategic Synergies
  if (showStrategic) {
    addSectionBox(slide, colors, 0.3, contentTop, colWidth, 4.05, 'Strategic Synergies', colors.primary, fontAdj);
    
    let strategicSynergies = parseLines(data.strategicSynergies, 8);
    if (strategicSynergies.length === 0) {
      strategicSynergies = [
        'Geographic expansion through established presence',
        'Technology integration and platform consolidation',
        'Cross-sell opportunities with complementary services',
        'Talent acquisition and capability enhancement',
        'Client relationship leverage'
      ];
    }
    
    strategicSynergies.slice(0, 8).forEach((syn, idx) => {
      slide.addText(`▸ ${truncateText(syn, showStrategic && showFinancial ? 50 : 100)}`, {
        x: 0.42, y: contentTop + 0.48 + (idx * 0.42), w: colWidth - 0.24, h: 0.38,
        fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
        color: colors.text,
        fontFace: 'Arial'
      });
    });
  }
  
  // Financial Synergies
  if (showFinancial) {
    const finColX = showStrategic ? 5.05 : 0.3;
    addSectionBox(slide, colors, finColX, contentTop, colWidth, 4.05, 'Financial Synergies', colors.secondary, fontAdj);
    
    let financialSynergies = parseLines(data.financialSynergies, 8);
    if (financialSynergies.length === 0) {
      financialSynergies = [
        'Revenue synergies through cross-selling',
        'Cost optimization through shared services',
        'Procurement savings from combined scale',
        'Operational efficiency improvements',
        'Working capital optimization'
      ];
    }
    
    financialSynergies.slice(0, 8).forEach((syn, idx) => {
      slide.addText(`▸ ${truncateText(syn, showStrategic && showFinancial ? 50 : 100)}`, {
        x: finColX + 0.12, y: contentTop + 0.48 + (idx * 0.42), w: colWidth - 0.24, h: 0.38,
        fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
        color: colors.text,
        fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * MARKET POSITION SLIDE - with competitive analysis
 */
function renderMarketPositionSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  const industryData = context.industryData;
  
  addSlideHeader(slide, colors, 'Market Position & Competitive Landscape', 'Industry positioning', fontAdj);
  
  const contentTop = 0.95;
  
  // LEFT - Market Opportunity
  addSectionBox(slide, colors, 0.3, contentTop, 4.0, 1.8, 'Market Opportunity', colors.primary, fontAdj);
  
  if (data.marketSize) {
    slide.addText('Total Addressable Market', {
      x: 0.42, y: contentTop + 0.5, w: 3.7, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
    slide.addText(data.marketSize, {
      x: 0.42, y: contentTop + 0.78, w: 3.7, h: 0.45,
      fontSize: adjustedFont(DESIGN.fonts.metricMedium, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial'
    });
    if (data.marketGrowthRate) {
      slide.addText(`Growing at ${data.marketGrowthRate}`, {
        x: 0.42, y: contentTop + 1.28, w: 3.7, h: 0.25,
        fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
        italic: true,
        color: colors.accent,
        fontFace: 'Arial'
      });
    }
  } else if (industryData) {
    slide.addText('Industry Market Size', {
      x: 0.42, y: contentTop + 0.5, w: 3.7, h: 0.25,
      fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
      color: colors.textLight,
      fontFace: 'Arial'
    });
    slide.addText(industryData.benchmarks?.marketSize || '$100B+', {
      x: 0.42, y: contentTop + 0.78, w: 3.7, h: 0.45,
      fontSize: adjustedFont(DESIGN.fonts.metricMedium, fontAdj),
      bold: true,
      color: colors.primary,
      fontFace: 'Arial'
    });
  }
  
  // Market Position Box
  addSectionBox(slide, colors, 4.5, contentTop, 5.2, 1.8, 'Competitive Position', colors.secondary, fontAdj);
  
  let positioning = parseLines(data.competitivePositioning, 4);
  if (positioning.length === 0) {
    positioning = [
      'Top 3 player in cloud services for BFSI sector',
      'Only company with all 3 major cloud certifications',
      'Highest customer satisfaction score in segment',
      'Fastest growing in managed services'
    ];
  }
  
  positioning.slice(0, 4).forEach((pos, idx) => {
    slide.addText(`• ${truncateText(pos, 55)}`, {
      x: 4.62, y: contentTop + 0.5 + (idx * 0.32), w: 4.95, h: 0.3,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial'
    });
  });
  
  // Competitive Analysis Table
  addSectionBox(slide, colors, 0.3, contentTop + 1.95, 9.4, 2.1, 'Competitive Analysis', colors.accent, fontAdj);
  
  const competitors = parsePipeSeparated(data.competitiveAnalysis, 4);
  
  if (competitors.length > 0) {
    // Headers
    const headers = ['Company', 'Strengths', 'Weaknesses'];
    headers.forEach((header, idx) => {
      const hx = 0.42 + (idx * 3.1);
      slide.addText(header, {
        x: hx, y: contentTop + 2.38, w: 3.0, h: 0.28,
        fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
        bold: true,
        color: colors.primary,
        fontFace: 'Arial'
      });
    });
    
    // Data rows
    competitors.slice(0, 4).forEach((comp, idx) => {
      const ry = contentTop + 2.7 + (idx * 0.32);
      
      slide.addText(truncateText(comp[0] || '', 22), {
        x: 0.42, y: ry, w: 3.0, h: 0.3,
        fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
        bold: true,
        color: colors.text,
        fontFace: 'Arial'
      });
      
      slide.addText(truncateText(comp[1] || '', 35), {
        x: 3.52, y: ry, w: 3.0, h: 0.3,
        fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
        color: colors.text,
        fontFace: 'Arial'
      });
      
      slide.addText(truncateText(comp[2] || '', 35), {
        x: 6.62, y: ry, w: 3.0, h: 0.3,
        fontSize: adjustedFont(DESIGN.fonts.bodySmall, fontAdj),
        color: colors.text,
        fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * RISKS SLIDE (for CIM)
 */
function renderRisksSlide(slide, colors, data, slideNumber, layoutRec, context) {
  const fontAdj = layoutRec.fontAdjustment || 0;
  
  addSlideHeader(slide, colors, 'Risk Factors', 'Key considerations', fontAdj);
  
  const contentTop = 0.95;
  
  let risks = parseLines(data.riskFactors, 10);
  if (risks.length === 0) {
    risks = [
      'Client concentration risk with top accounts',
      'Dependency on key technology partnerships',
      'Regulatory changes in financial services sector',
      'Talent retention in competitive market',
      'Currency fluctuation exposure',
      'Technology obsolescence risk'
    ];
  }
  
  // Display risks in 2-column format
  addSectionBox(slide, colors, 0.3, contentTop, 9.4, 4.05, 'Key Risk Factors & Mitigants', colors.primary, fontAdj);
  
  risks.slice(0, 10).forEach((risk, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const rx = 0.42 + (col * 4.7);
    const ry = contentTop + 0.5 + (row * 0.72);
    
    slide.addShape('rect', {
      x: rx, y: ry, w: 4.5, h: 0.65,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 },
      rectRadius: 0.04
    });
    
    slide.addText(`${idx + 1}`, {
      x: rx + 0.08, y: ry + 0.12, w: 0.35, h: 0.35,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      bold: true,
      color: colors.accent,
      fontFace: 'Arial'
    });
    
    slide.addText(truncateText(risk, 60), {
      x: rx + 0.48, y: ry + 0.1, w: 3.9, h: 0.45,
      fontSize: adjustedFont(DESIGN.fonts.body, fontAdj),
      color: colors.text,
      fontFace: 'Arial',
      valign: 'middle'
    });
  });
  
  addSlideFooter(slide, colors, slideNumber);
}

/**
 * THANK YOU SLIDE
 */
function renderThankYouSlide(slide, colors, data, docConfig, fontAdj = 0) {
  // Dark background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  // Decorative elements
  slide.addShape('rect', {
    x: 0, y: 4.5, w: '100%', h: 0.04,
    fill: { color: colors.secondary }
  });
  
  // Thank You text
  slide.addText('Thank You', {
    x: 0, y: 1.8, w: '100%', h: 0.9,
    fontSize: 48,
    bold: true,
    color: colors.white,
    fontFace: 'Arial',
    align: 'center'
  });
  
  // Contact info
  if (data.contactEmail || data.contactPhone) {
    slide.addText('For Further Information', {
      x: 0, y: 3.0, w: '100%', h: 0.35,
      fontSize: 14,
      color: colors.white,
      fontFace: 'Arial',
      align: 'center',
      transparency: 25
    });
    
    let contactText = '';
    if (data.contactEmail) contactText += data.contactEmail;
    if (data.contactEmail && data.contactPhone) contactText += ' | ';
    if (data.contactPhone) contactText += data.contactPhone;
    
    slide.addText(contactText, {
      x: 0, y: 3.4, w: '100%', h: 0.35,
      fontSize: 12,
      color: colors.white,
      fontFace: 'Arial',
      align: 'center'
    });
  }
  
  // Confidential notice
  slide.addText('Strictly Private and Confidential', {
    x: 0, y: 4.7, w: '100%', h: 0.3,
    fontSize: 10,
    italic: true,
    color: colors.white,
    fontFace: 'Arial',
    align: 'center',
    transparency: 40
  });
}


// ============================================================================
// MAIN PPTX GENERATOR - v7.1 with Universal createSlide()
// ============================================================================

/**
 * Generates the complete presentation using AI-powered createSlide() wrapper
 */
async function generatePresentationWithAI(data, theme = 'modern-blue') {
  const pptx = new PptxGenJS();
  
  // Set presentation properties
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = data.projectCodename || 'Project Phoenix';
  pptx.author = data.advisor || 'IM Creator';
  pptx.company = data.companyName || 'Company';
  
  // Get theme colors
  const colors = THEMES[theme] || THEMES['modern-blue'];
  
  // Get document configuration
  const docType = data.documentType || 'management-presentation';
  const docConfig = DOCUMENT_CONFIGS[docType] || DOCUMENT_CONFIGS['management-presentation'];
  
  // Get industry data
  const industryData = INDUSTRY_DATA[data.primaryVertical] || INDUSTRY_DATA.technology;
  
  // Get buyer content
  const targetBuyers = Array.isArray(data.targetBuyerTypes) ? data.targetBuyerTypes : [data.targetBuyerTypes || 'strategic'];
  const primaryBuyerType = targetBuyers[0] || 'strategic';
  const buyerContent = BUYER_CONTENT[primaryBuyerType] || BUYER_CONTENT.strategic;
  
  // Context object to pass to render functions
  const context = {
    docConfig,
    industryData,
    buyerContent,
    targetBuyers
  };
  
  let slideNumber = 1;
  
  // =============================================
  // TITLE SLIDE (no number)
  // =============================================
  await createSlide('title', pptx, colors, data, 0, context);
  
  // =============================================
  // DISCLAIMER SLIDE
  // =============================================
  slideNumber = await createSlide('disclaimer', pptx, colors, data, slideNumber, context);
  
  // =============================================
  // TABLE OF CONTENTS (CIM only)
  // =============================================
  if (docType === 'cim') {
    const slideList = [
      'Executive Summary', 'Investment Highlights', 'Company Overview',
      'Leadership Team', 'Industry Overview', 'Services & Capabilities',
      'Client Portfolio', 'Financial Performance', 'Growth Strategy',
      'Synergies', 'Risk Factors'
    ];
    context.slideList = slideList;
    slideNumber = await createSlide('toc', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // EXECUTIVE SUMMARY
  // =============================================
  slideNumber = await createSlide('executive-summary', pptx, colors, data, slideNumber, context);
  
  // =============================================
  // INVESTMENT HIGHLIGHTS
  // =============================================
  slideNumber = await createSlide('investment-highlights', pptx, colors, data, slideNumber, context);
  
  // =============================================
  // COMPANY OVERVIEW (CIM only)
  // =============================================
  if (docType === 'cim') {
    slideNumber = await createSlide('company-overview', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // LEADERSHIP (not teaser)
  // =============================================
  if (docType !== 'teaser') {
    slideNumber = await createSlide('leadership', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // INDUSTRY OVERVIEW (CIM only)
  // =============================================
  if (docType === 'cim') {
    slideNumber = await createSlide('industry', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // SERVICES
  // =============================================
  slideNumber = await createSlide('services', pptx, colors, data, slideNumber, context);
  
  // =============================================
  // CLIENTS
  // =============================================
  slideNumber = await createSlide('clients', pptx, colors, data, slideNumber, context);
  
  // =============================================
  // FINANCIALS (not teaser)
  // =============================================
  if (docConfig.includeFinancialDetail) {
    slideNumber = await createSlide('financials', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // CASE STUDIES (not teaser)
  // =============================================
  if (docType !== 'teaser') {
    const maxCaseStudies = docConfig.maxCaseStudies || 2;
    
    // Collect case studies
    let caseStudiesToShow = [];
    if (data.caseStudies && Array.isArray(data.caseStudies)) {
      caseStudiesToShow = data.caseStudies.filter(cs => cs.client).slice(0, maxCaseStudies);
    }
    
    // Check legacy format
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
    
    // Generate case study slides
    for (const caseStudy of caseStudiesToShow) {
      context.caseStudy = caseStudy;
      slideNumber = await createSlide('case-study', pptx, colors, data, slideNumber, context);
    }
  }
  
  // =============================================
  // GROWTH STRATEGY (not teaser)
  // =============================================
  if (docType !== 'teaser') {
    slideNumber = await createSlide('growth', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // SYNERGIES (not teaser)
  // =============================================
  if (docType !== 'teaser') {
    slideNumber = await createSlide('synergies', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // MARKET POSITION (optional, if content variants selected)
  // =============================================
  const contentVariants = Array.isArray(data.contentVariants) ? data.contentVariants : [];
  if (contentVariants.includes('market-position') || data.marketSize || data.competitiveAnalysis) {
    slideNumber = await createSlide('market-position', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // RISK FACTORS (CIM only)
  // =============================================
  if (docType === 'cim') {
    slideNumber = await createSlide('risks', pptx, colors, data, slideNumber, context);
  }
  
  // =============================================
  // THANK YOU SLIDE
  // =============================================
  await createSlide('thank-you', pptx, colors, data, slideNumber, context);
  
  return pptx;
}

// ============================================================================
// API ENDPOINTS
// ============================================================================

// Health check with version info
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    version: VERSION.string,
    versionFull: VERSION.full,
    buildDate: VERSION.buildDate,
    features: [
      'AI-Powered Layout Engine with createSlide() wrapper',
      'Dedicated render functions for each slide type',
      'addChartByType() helper (Bar, Pie, Donut, Progress, Timeline)',
      'Larger fonts (14pt body, 26pt titles)',
      'Dynamic font adjustment based on content density',
      'Document types (CIM, Management Presentation, Teaser)',
      '50 Professional Templates',
      'Industry-specific content',
      'Target buyer type integration'
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

// ============================================================================
// GENERATE PPTX ENDPOINT - Main endpoint for presentation generation
// ============================================================================
app.post('/api/generate-pptx', async (req, res) => {
  try {
    const { data, theme } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }
    
    console.log(`\n${'='.repeat(50)}`);
    console.log(`Generating PPTX with AI Layout Engine v${VERSION.string}`);
    console.log(`Theme: ${theme || 'modern-blue'}`);
    console.log(`Document Type: ${data.documentType || 'management-presentation'}`);
    console.log(`${'='.repeat(50)}`);
    
    // Generate presentation with AI-powered layouts
    const pptx = await generatePresentationWithAI(data, theme || 'modern-blue');
    
    // Generate unique filename
    const timestamp = Date.now();
    const codename = (data.projectCodename || 'Project').replace(/[^a-zA-Z0-9]/g, '_');
    const filename = `${codename}_${timestamp}.pptx`;
    const filepath = path.join(tempDir, filename);
    
    // Write file
    await pptx.writeFile({ fileName: filepath });
    
    console.log(`Generated: ${filename}`);
    console.log(`AI Layout calls: ${usageStats.calls.filter(c => c.purpose?.startsWith('AI Layout')).length}`);
    
    // Send file
    res.download(filepath, filename, (err) => {
      if (err) {
        console.error('Download error:', err);
      }
      // Clean up
      setTimeout(() => {
        try { fs.unlinkSync(filepath); } catch (e) {}
      }, 60000);
    });
    
  } catch (error) {
    console.error('PPTX generation error:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================================================
// CHAT & AI ENDPOINTS
// ============================================================================
app.post('/api/chat', async (req, res) => {
  try {
    const { message, conversationHistory = [], context = {} } = req.body;
    
    const systemPrompt = `You are an expert M&A advisor helping create professional Information Memorandum presentations. 
    You help users fill out company information for CIM, Management Presentations, and Teasers.
    Be concise, professional, and helpful. Ask clarifying questions when needed.
    Current form context: ${JSON.stringify(context)}`;
    
    const messages = [
      ...conversationHistory.map(msg => ({
        role: msg.role,
        content: msg.content
      })),
      { role: 'user', content: message }
    ];
    
    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1024,
      system: systemPrompt,
      messages: messages
    });
    
    trackUsage('claude-sonnet-4-20250514', response.usage.input_tokens, response.usage.output_tokens, 'Chat');
    
    res.json({
      response: response.content[0].text,
      usage: {
        inputTokens: response.usage.input_tokens,
        outputTokens: response.usage.output_tokens
      }
    });
  } catch (error) {
    console.error('Chat error:', error);
    res.status(500).json({ error: error.message });
  }
});

// AI Suggestions endpoint
app.post('/api/ai-suggest', async (req, res) => {
  try {
    const { field, currentValue, formData } = req.body;
    
    const prompt = `As an M&A advisor, suggest professional content for the "${field}" field.
    Current value: ${currentValue || 'empty'}
    Company: ${formData?.companyName || 'Unknown'}
    Industry: ${formData?.primaryVertical || 'Technology'}
    
    Provide 3 concise, professional suggestions in JSON array format.`;
    
    const response = await anthropic.messages.create({
      model: 'claude-3-haiku-20240307',
      max_tokens: 500,
      messages: [{ role: 'user', content: prompt }]
    });
    
    trackUsage('claude-3-haiku-20240307', response.usage.input_tokens, response.usage.output_tokens, 'AI Suggest');
    
    res.json({
      suggestions: response.content[0].text,
      usage: {
        inputTokens: response.usage.input_tokens,
        outputTokens: response.usage.output_tokens
      }
    });
  } catch (error) {
    console.error('Suggestion error:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================================================
// DRAFT MANAGEMENT
// ============================================================================
app.post('/api/drafts', (req, res) => {
  const { data, projectId } = req.body;
  
  const draftsDir = path.join(tempDir, 'drafts');
  if (!fs.existsSync(draftsDir)) fs.mkdirSync(draftsDir, { recursive: true });
  
  const draftPath = path.join(draftsDir, `${projectId}.json`);
  fs.writeFileSync(draftPath, JSON.stringify(data, null, 2));
  
  res.json({ success: true, projectId });
});

app.get('/api/drafts/:projectId', (req, res) => {
  const draftPath = path.join(tempDir, 'drafts', `${req.params.projectId}.json`);
  if (fs.existsSync(draftPath)) {
    const data = JSON.parse(fs.readFileSync(draftPath, 'utf8'));
    res.json({ success: true, data });
  } else {
    res.status(404).json({ success: false, error: 'Draft not found' });
  }
});

// ============================================================================
// START SERVER
// ============================================================================
app.listen(PORT, () => {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`IM Creator Server ${VERSION.full} - AI-Powered Layout Engine`);
  console.log(`${'='.repeat(60)}`);
  console.log(`Running on port ${PORT}`);
  console.log(`\nNEW in v7.1.0:`);
  console.log(`  ✓ Universal createSlide() wrapper with AI integration`);
  console.log(`  ✓ Dedicated render functions for each slide type`);
  console.log(`  ✓ addChartByType() helper (Bar, Pie, Donut, Progress, Timeline)`);
  console.log(`  ✓ AI recommendations applied to actual rendering`);
  console.log(`  ✓ Dynamic font adjustment based on content density`);
  console.log(`\nPreserved from v7.0.0:`);
  console.log(`  ✓ AI-powered layout engine`);
  console.log(`  ✓ Larger fonts (26pt titles, 14pt headers, 12pt body)`);
  console.log(`  ✓ Diverse charts: Pie, Donut, Progress bars, Timelines`);
  console.log(`\nPreserved from v6.x:`);
  console.log(`  ✓ All 14 core features`);
  console.log(`  ✓ 50 professional templates`);
  console.log(`  ✓ Document types (CIM, Management Presentation, Teaser)`);
  console.log(`${'='.repeat(60)}\n`);
});

