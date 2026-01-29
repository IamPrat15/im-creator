// Enhanced IM Creator Server v3.0 - Complete Fix for All Issues
// Fixes: 
// 1. Consistent slide numbering
// 2. All case studies included
// 3. Proper pie charts with segments
// 4. Text truncation to prevent overflow
// 5. Generic template names (no company names)
// 6. Theme colors properly applied
// 7. Target buyer type affects content
// 8. Content variants and appendix support
// 9. Anthropic usage/cost tracking

const express = require('express');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

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
// ANTHROPIC USAGE TRACKING (Issue #9)
// ============================================================================
let usageStats = {
  totalInputTokens: 0,
  totalOutputTokens: 0,
  totalCalls: 0,
  totalCostUSD: 0,
  sessionStart: new Date().toISOString(),
  calls: []
};

// Pricing (as of 2025 - adjust as needed)
const PRICING = {
  'claude-sonnet-4-20250514': { input: 0.003, output: 0.015 }, // per 1K tokens
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
  
  // Keep only last 100 calls for memory efficiency
  if (usageStats.calls.length > 100) {
    usageStats.calls = usageStats.calls.slice(-100);
  }
  
  return costUSD;
}

// ============================================================================
// PROFESSIONAL COLOR THEMES - Generic Names (Issue #5 & #6)
// ============================================================================
const THEMES = {
  'modern-blue': {
    name: 'Modern Blue',
    primary: '2B579A',
    secondary: '86BC25',
    accent: 'FFC72C',
    text: '333333',
    textLight: '666666',
    white: 'FFFFFF',
    lightBg: 'F5F7FA',
    darkBg: '1A1F36',
    border: 'E0E5EC',
    success: '28A745',
    warning: 'FFC107',
    danger: 'DC3545',
    chartColors: ['2B579A', '86BC25', 'FFC72C', '00A3E0', 'E31B23', '6B3FA0']
  },
  'corporate-navy': {
    name: 'Corporate Navy',
    primary: '003366',
    secondary: 'B8860B',
    accent: 'C9A227',
    text: '333333',
    textLight: '666666',
    white: 'FFFFFF',
    lightBg: 'F8F9FA',
    darkBg: '1A2332',
    border: 'DEE2E6',
    success: '28A745',
    warning: 'FFC107',
    danger: 'DC3545',
    chartColors: ['003366', 'B8860B', '5B9BD5', '70AD47', 'ED7D31', '7030A0']
  },
  'elegant-burgundy': {
    name: 'Elegant Burgundy',
    primary: '7C1034',
    secondary: '2D3748',
    accent: '48BB78',
    text: '1A202C',
    textLight: '718096',
    white: 'FFFFFF',
    lightBg: 'F7FAFC',
    darkBg: '5A0C26',
    border: 'E2E8F0',
    success: '48BB78',
    warning: 'ECC94B',
    danger: 'E53E3E',
    chartColors: ['7C1034', '48BB78', 'ECC94B', '4299E1', '9A1842', '2D3748']
  },
  'minimalist-mono': {
    name: 'Minimalist',
    primary: '212121',
    secondary: '757575',
    accent: '2196F3',
    text: '212121',
    textLight: '757575',
    white: 'FFFFFF',
    lightBg: 'FAFAFA',
    darkBg: '212121',
    border: 'E0E0E0',
    success: '4CAF50',
    warning: 'FF9800',
    danger: 'F44336',
    chartColors: ['212121', '757575', '2196F3', '4CAF50', 'FF9800', '9C27B0']
  },
  'forest-green': {
    name: 'Forest Green',
    primary: '1B5E20',
    secondary: '33691E',
    accent: 'FFC107',
    text: '212121',
    textLight: '616161',
    white: 'FFFFFF',
    lightBg: 'F1F8E9',
    darkBg: '1B5E20',
    border: 'C8E6C9',
    success: '4CAF50',
    warning: 'FF9800',
    danger: 'F44336',
    chartColors: ['1B5E20', '33691E', 'FFC107', '4CAF50', '8BC34A', '689F38']
  }
};

// Map old theme names to new ones for backward compatibility
const THEME_MAP = {
  'modern-tech': 'modern-blue',
  'conservative': 'corporate-navy',
  'minimalist': 'minimalist-mono',
  'acc-brand': 'elegant-burgundy'
};

// ============================================================================
// TEXT UTILITIES (Issue #4 - Prevent overflow)
// ============================================================================
function truncateText(text, maxLength, suffix = '...') {
  if (!text || text.length <= maxLength) return text || '';
  return text.substring(0, maxLength - suffix.length).trim() + suffix;
}

function truncateLines(text, maxLines, maxCharsPerLine = 80) {
  if (!text) return '';
  const lines = text.split('\n').filter(l => l.trim());
  const truncated = lines.slice(0, maxLines).map(l => truncateText(l, maxCharsPerLine));
  return truncated.join('\n');
}

function smartTruncate(text, maxChars, preserveWords = true) {
  if (!text || text.length <= maxChars) return text || '';
  if (!preserveWords) return text.substring(0, maxChars - 3) + '...';
  
  const words = text.split(' ');
  let result = '';
  for (const word of words) {
    if ((result + ' ' + word).trim().length > maxChars - 3) break;
    result = (result + ' ' + word).trim();
  }
  return result + '...';
}

// ============================================================================
// BUYER TYPE CONTENT ADAPTATION (Issue #7)
// ============================================================================
function adaptContentForBuyers(data, targetBuyers = []) {
  const adapted = { ...data };
  
  // Default to all if none selected
  if (!targetBuyers || targetBuyers.length === 0) {
    targetBuyers = ['strategic', 'financial'];
  }
  
  // Emphasis indicators
  adapted._emphasis = {
    financial: targetBuyers.includes('financial'),
    strategic: targetBuyers.includes('strategic'),
    international: targetBuyers.includes('international')
  };
  
  // Generate buyer-specific messaging
  if (targetBuyers.includes('financial')) {
    adapted._financialPitch = [
      'Strong and predictable revenue growth',
      'Healthy EBITDA margins with expansion potential',
      'Asset-light business model',
      'Clear path to value creation',
      'Multiple exit options available'
    ];
  }
  
  if (targetBuyers.includes('strategic')) {
    adapted._strategicPitch = [
      'Complementary technology capabilities',
      'Access to new markets and clients',
      'Skilled talent pool acquisition',
      'Platform for regional expansion',
      'Cross-selling opportunities'
    ];
  }
  
  if (targetBuyers.includes('international')) {
    adapted._internationalPitch = [
      'India market entry/expansion platform',
      'Cost-effective delivery capabilities',
      'English-speaking talent base',
      'Growing Asia-Pacific presence',
      'Global delivery model ready'
    ];
  }
  
  return adapted;
}

// ============================================================================
// CONTENT VARIANTS GENERATOR (Issue #8)
// ============================================================================
function generateVariantContent(data, variants = []) {
  const content = {};
  
  if (variants.includes('financial')) {
    content.financialFocus = {
      title: 'Financial Performance Highlights',
      points: [
        `Revenue CAGR of ~30% over last 3 years`,
        `EBITDA Margin: ${data.ebitdaMarginFY25 || 22}%`,
        `Strong cash flow generation`,
        `Low capital intensity business`,
        `High revenue visibility from recurring contracts`
      ]
    };
  }
  
  if (variants.includes('tech')) {
    content.techFocus = {
      title: 'Technology Capabilities',
      points: [
        'Proprietary AI/ML platforms',
        'Cloud-native architecture expertise',
        'Modern DevOps and automation',
        'Data engineering capabilities',
        'Scalable microservices approach'
      ]
    };
  }
  
  if (variants.includes('market')) {
    content.marketFocus = {
      title: 'Market Position',
      points: [
        `${data.netRetention || 118}% Net Revenue Retention`,
        'Deep domain expertise in BFSI',
        'Long-standing client relationships',
        'Strong brand recognition',
        'Differentiated service offerings'
      ]
    };
  }
  
  if (variants.includes('synergy')) {
    content.synergyFocus = {
      title: 'Synergy Potential',
      points: [
        'Cross-selling opportunities',
        'Geographic expansion platform',
        'Technology capability enhancement',
        'Talent acquisition',
        'Cost optimization potential'
      ]
    };
  }
  
  return content;
}

// ============================================================================
// PROFESSIONAL POWERPOINT GENERATOR - All Fixes Applied
// ============================================================================
async function generateProfessionalPPTX(data, theme = 'modern-blue', options = {}) {
  // Map old theme names
  const mappedTheme = THEME_MAP[theme] || theme;
  const colors = THEMES[mappedTheme] || THEMES['modern-blue'];
  
  // Adapt content for target buyers (Issue #7)
  const targetBuyers = data.targetBuyerType || [];
  const adaptedData = adaptContentForBuyers(data, targetBuyers);
  
  // Generate variant content if specified (Issue #8)
  const variants = data.generateVariants || [];
  const variantContent = generateVariantContent(data, variants);
  
  // Appendix options (Issue #8)
  const appendixOptions = data.includeAppendix || [];
  
  const pptx = new PptxGenJS();
  
  // Presentation metadata
  pptx.author = data.advisor || 'Investment Bank';
  pptx.title = `${data.projectCodename || 'Project'} - Management Presentation`;
  pptx.subject = 'Confidential Information Memorandum';
  pptx.company = data.advisor || 'Investment Bank';
  pptx.layout = 'LAYOUT_16x9';

  let slideNumber = 1;

  // ============================================================================
  // SLIDE 1: COVER PAGE (No slide number)
  // ============================================================================
  const slide1 = pptx.addSlide();
  
  slide1.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { type: 'solid', color: colors.darkBg }
  });
  
  // Decorative elements using theme colors
  slide1.addShape('rect', {
    x: 6.5, y: 0, w: 3.5, h: 2.5,
    fill: { color: colors.primary },
    transparency: 85
  });
  slide1.addShape('rect', {
    x: 7.5, y: 0.5, w: 2.5, h: 2,
    fill: { color: colors.secondary },
    transparency: 80
  });
  
  // Accent line
  slide1.addShape('rect', {
    x: 0.5, y: 3.2, w: 4, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  slide1.addText(data.projectCodename || 'Project Phoenix', {
    x: 0.5, y: 2.2, w: 8, h: 1,
    fontSize: 48, bold: true, color: colors.white,
    fontFace: 'Arial'
  });
  
  slide1.addText('Management Presentation', {
    x: 0.5, y: 3.35, w: 6, h: 0.5,
    fontSize: 22, color: colors.white,
    fontFace: 'Arial'
  });
  
  slide1.addText(formatDate(data.presentationDate), {
    x: 0.5, y: 3.95, w: 4, h: 0.35,
    fontSize: 14, color: colors.white,
    fontFace: 'Arial', transparency: 30
  });
  
  slide1.addText(data.advisor || 'Investment Bank', {
    x: 0.5, y: 4.4, w: 4, h: 0.35,
    fontSize: 12, color: colors.white,
    fontFace: 'Arial', transparency: 20
  });
  
  slide1.addText('Strictly Private and Confidential', {
    x: 0.5, y: 4.85, w: 4, h: 0.3,
    fontSize: 11, italic: true, color: colors.white,
    fontFace: 'Arial', transparency: 40
  });

  // ============================================================================
  // SLIDE 2: DISCLAIMER (Page 2)
  // ============================================================================
  slideNumber++;
  const slide2 = pptx.addSlide();
  addSlideHeader(slide2, colors, 'Important Notice', null);
  
  const disclaimerText = `The information contained in this document has been compiled by ${data.advisor || 'the Advisor'} based on information obtained from public sources. Except in the general context of evaluating the capabilities of ${data.advisor || 'the Advisor'}, no reliance may be placed for any purposes whatsoever on the contents of this document or on its completeness.

This document and its contents are confidential and may not be reproduced, redistributed or passed on, directly or indirectly, to any other person in whole or in part without the prior written consent of ${data.advisor || 'the Advisor'}.

This document does not constitute an offer or agreement between ${data.advisor || 'the Advisor'} and ${data.companyName || 'the Company'}. Furthermore, changes in Company definition of requirements will necessarily affect the proposal set forth herein.`;

  slide2.addText(disclaimerText, {
    x: 0.5, y: 1.3, w: 9, h: 3.5,
    fontSize: 11, color: colors.text, fontFace: 'Arial',
    valign: 'top', lineSpacingMultiple: 1.5
  });
  
  addSlideFooter(slide2, colors, slideNumber);

  // ============================================================================
  // SLIDE 3: EXECUTIVE SUMMARY (Page 3)
  // ============================================================================
  slideNumber++;
  const slide3 = pptx.addSlide();
  
  // Title without section number for cleaner look
  addSlideHeader(slide3, colors, truncateText(data.companyDescription || 'A Leading Digital Transformation Partner', 80), null);
  
  // Left column - Key stats
  const stats = [
    { value: String(data.foundedYear || '2014'), label: 'Founded Year' },
    { value: `${data.employeeCountFT || '350'}+`, label: `Headcount FY${new Date().getFullYear()}` },
    { value: '80+', label: 'Active Clients' },
    { value: '98%', label: 'Domestic Revenue' },
    { value: '300+', label: 'Successful Projects' }
  ];
  
  slide3.addShape('rect', {
    x: 0.3, y: 1.1, w: 2.2, h: 3.9,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 }
  });
  
  stats.forEach((stat, idx) => {
    slide3.addText(stat.value, {
      x: 0.4, y: 1.2 + (idx * 0.75), w: 2, h: 0.35,
      fontSize: 20, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide3.addText(stat.label, {
      x: 0.4, y: 1.55 + (idx * 0.75), w: 2, h: 0.25,
      fontSize: 9, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  // Middle column - Key Offerings
  slide3.addShape('rect', {
    x: 2.6, y: 1.1, w: 3.6, h: 3.9,
    fill: { color: colors.white },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide3.addText('Key Offerings', {
    x: 2.7, y: 1.15, w: 3.4, h: 0.35,
    fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial',
    fill: { color: colors.primary }
  });
  
  // Parse service lines for offerings
  const serviceLines = (data.serviceLines || '').split('\n').filter(s => s.trim()).slice(0, 6);
  const offerings = serviceLines.map(s => {
    const parts = s.split('|');
    return parts[0]?.trim() || s.trim();
  });
  
  // Fill with defaults if needed
  while (offerings.length < 6) {
    offerings.push(['Cloud Services', 'AI Solutions', 'Managed Services', 'Data Analytics', 'Product Engineering', 'Digital Transformation'][offerings.length]);
  }
  
  offerings.slice(0, 6).forEach((offering, idx) => {
    const row = Math.floor(idx / 2);
    const col = idx % 2;
    const offeringColors = [colors.primary, colors.secondary, colors.chartColors[2], colors.chartColors[3], colors.chartColors[4], colors.chartColors[5]];
    
    slide3.addShape('roundRect', {
      x: 2.75 + (col * 1.7), y: 1.6 + (row * 1), w: 1.6, h: 0.85,
      fill: { color: offeringColors[idx] || colors.primary }
    });
    slide3.addText(truncateText(offering, 25), {
      x: 2.75 + (col * 1.7), y: 1.75 + (row * 1), w: 1.6, h: 0.55,
      fontSize: 9, color: colors.white, fontFace: 'Arial',
      align: 'center', valign: 'middle'
    });
  });
  
  // Right column - Financial Highlights
  slide3.addShape('rect', {
    x: 6.3, y: 1.1, w: 3.4, h: 3.9,
    fill: { color: colors.white },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide3.addText('Financial Highlights', {
    x: 6.4, y: 1.15, w: 3.2, h: 0.35,
    fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial',
    fill: { color: colors.primary }
  });
  
  // Revenue bar chart
  const revenueData = [
    { year: 'FY24', value: parseFloat(data.revenueFY24) || 85 },
    { year: 'FY25', value: parseFloat(data.revenueFY25) || 112 },
    { year: 'FY26P', value: parseFloat(data.revenueFY26P) || 150 },
    { year: 'FY27P', value: parseFloat(data.revenueFY27P) || 195 },
    { year: 'FY28P', value: parseFloat(data.revenueFY28P) || 250 }
  ];
  
  const maxRev = Math.max(...revenueData.map(d => d.value));
  revenueData.forEach((rev, idx) => {
    const barHeight = (rev.value / maxRev) * 1.8;
    const xPos = 6.5 + (idx * 0.6);
    const isProjected = rev.year.includes('P');
    
    slide3.addShape('rect', {
      x: xPos, y: 3.7 - barHeight, w: 0.45, h: barHeight,
      fill: { color: isProjected ? colors.secondary : colors.primary }
    });
    slide3.addText(`${rev.value}`, {
      x: xPos - 0.1, y: 3.7 - barHeight - 0.25, w: 0.65, h: 0.25,
      fontSize: 8, color: colors.text, fontFace: 'Arial', align: 'center'
    });
    slide3.addText(rev.year, {
      x: xPos - 0.05, y: 3.75, w: 0.55, h: 0.2,
      fontSize: 7, color: colors.textLight, fontFace: 'Arial', align: 'center'
    });
  });
  
  slide3.addText(`In ${data.currency === 'USD' ? 'USD Mn' : 'INR Cr'}`, {
    x: 6.4, y: 1.55, w: 1, h: 0.2,
    fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  slide3.addText('CAGR: ~30%', {
    x: 8.2, y: 1.55, w: 1.2, h: 0.2,
    fontSize: 9, bold: true, color: colors.secondary, fontFace: 'Arial', align: 'right'
  });
  
  slide3.addText(`EBITDA Margin FY25: ${data.ebitdaMarginFY25 || 22}%`, {
    x: 6.4, y: 4.1, w: 3.2, h: 0.25,
    fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  // Platform capabilities
  slide3.addText('Platform Capabilities', {
    x: 6.4, y: 4.45, w: 3.2, h: 0.25,
    fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const platforms = ['AWS', 'Azure', 'GCP', 'SAP'];
  platforms.forEach((platform, idx) => {
    slide3.addShape('roundRect', {
      x: 6.5 + (idx * 0.75), y: 4.72, w: 0.65, h: 0.25,
      fill: { color: colors.lightBg },
      line: { color: colors.border, width: 0.5 }
    });
    slide3.addText(platform, {
      x: 6.5 + (idx * 0.75), y: 4.72, w: 0.65, h: 0.25,
      fontSize: 7, color: colors.text, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
  });
  
  addSlideFooter(slide3, colors, slideNumber);

  // ============================================================================
  // SLIDE 4: FOUNDER PROFILE (Page 4)
  // ============================================================================
  slideNumber++;
  const slide4 = pptx.addSlide();
  addSlideHeader(slide4, colors, 'Founded & Led by Industry Veteran with Strong Educational Qualification & Industry Experience', null);
  
  // Photo placeholder
  slide4.addShape('ellipse', {
    x: 1.2, y: 1.5, w: 2, h: 2,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 2 }
  });
  slide4.addText('Photo', {
    x: 1.2, y: 2.3, w: 2, h: 0.4,
    fontSize: 12, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  slide4.addText(data.founderName || 'Founder Name', {
    x: 0.7, y: 3.6, w: 3, h: 0.4,
    fontSize: 20, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
  });
  slide4.addText(data.founderTitle || 'Founder & CEO', {
    x: 0.7, y: 4, w: 3, h: 0.3,
    fontSize: 14, color: colors.secondary, fontFace: 'Arial', align: 'center'
  });
  slide4.addText(`~${data.founderExperience || 20} years of total experience`, {
    x: 0.7, y: 4.3, w: 3, h: 0.3,
    fontSize: 11, italic: true, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  // Background info box
  slide4.addShape('rect', {
    x: 4, y: 1.3, w: 5.5, h: 3.2,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide4.addText("Founder's Background", {
    x: 4.1, y: 1.35, w: 5.3, h: 0.35,
    fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial',
    fill: { color: colors.primary }
  });
  
  // Parse education
  const education = (data.founderEducation || 'MBA - IIM Ahmedabad\nB.Tech - IIT').split('\n').filter(e => e.trim()).slice(0, 2);
  
  // Generate background points from input data
  const backgroundPoints = [
    `Founded ${data.companyName || 'the Company'} in ${data.foundedYear || '2015'} and leads its strategic direction`,
    `${education[0] || 'MBA from premier institution'}`,
    `${education[1] || 'Engineering background'}`,
    `${data.founderExperience || 20}+ years in technology and consulting`
  ];
  
  backgroundPoints.forEach((point, idx) => {
    slide4.addText(`•  ${truncateText(point, 70)}`, {
      x: 4.2, y: 1.8 + (idx * 0.5), w: 5.2, h: 0.45,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
  });
  
  // Previous experience
  slide4.addText('Previous Experience', {
    x: 4.1, y: 3.7, w: 5.3, h: 0.25,
    fontSize: 10, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  // Parse previous companies
  const prevCompanies = (data.previousCompanies || '').split('\n').filter(c => c.trim()).slice(0, 4);
  const companies = prevCompanies.map(c => c.split('|')[0]?.trim() || c.trim()).slice(0, 4);
  while (companies.length < 4) {
    companies.push(['Previous Co.', 'Company', 'Firm', 'Corp'][companies.length]);
  }
  
  companies.forEach((company, idx) => {
    slide4.addShape('rect', {
      x: 4.2 + (idx * 1.3), y: 4.0, w: 1.15, h: 0.45,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide4.addText(truncateText(company, 12), {
      x: 4.2 + (idx * 1.3), y: 4.0, w: 1.15, h: 0.45,
      fontSize: 8, color: colors.text, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
  });
  
  addSlideFooter(slide4, colors, slideNumber);

  // ============================================================================
  // SLIDE 5: COMPANY TIMELINE (Page 5)
  // ============================================================================
  slideNumber++;
  const slide5 = pptx.addSlide();
  addSlideHeader(slide5, colors, 'Evolving Continuously from Cloud Solutions to AI Agentic Solutions', null);
  
  slide5.addShape('rect', {
    x: 0.5, y: 2.8, w: 9, h: 0.03,
    fill: { color: colors.primary }
  });
  
  const foundedYear = parseInt(data.foundedYear) || 2015;
  const timeline = [
    { period: `${foundedYear}-${foundedYear + 1}`, points: ['Company founded', 'Initial cloud offerings', 'First office setup'] },
    { period: `${foundedYear + 2}-${foundedYear + 3}`, points: ['Core platform developed', 'Key client acquisition', 'Team expansion'] },
    { period: `${foundedYear + 4}-${foundedYear + 5}`, points: ['Industry recognition', 'New service lines', 'Geographic growth'] },
    { period: `${foundedYear + 6}-${foundedYear + 7}`, points: ['Major partnerships', 'Product innovations', 'Market leadership'] },
    { period: `${foundedYear + 8}-${foundedYear + 9}`, points: ['AI capabilities', 'Enterprise clients', 'Scale operations'] },
    { period: '2025-26', points: ['AI agents launch', 'Enterprise AI', 'Future growth'] }
  ];
  
  timeline.forEach((item, idx) => {
    const xPos = 0.7 + (idx * 1.55);
    
    slide5.addShape('ellipse', {
      x: xPos + 0.55, y: 2.7, w: 0.2, h: 0.2,
      fill: { color: colors.primary }
    });
    
    slide5.addShape('roundRect', {
      x: xPos + 0.1, y: 2.2, w: 1.1, h: 0.35,
      fill: { color: colors.primary }
    });
    slide5.addText(item.period, {
      x: xPos + 0.1, y: 2.2, w: 1.1, h: 0.35,
      fontSize: 10, bold: true, color: colors.white, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
    
    item.points.forEach((point, pIdx) => {
      slide5.addText(`• ${truncateText(point, 20)}`, {
        x: xPos, y: 3.0 + (pIdx * 0.35), w: 1.4, h: 0.35,
        fontSize: 8, color: colors.text, fontFace: 'Arial'
      });
    });
  });
  
  addSlideFooter(slide5, colors, slideNumber);

  // ============================================================================
  // SLIDE 6: SERVICE OFFERINGS (Page 6)
  // ============================================================================
  slideNumber++;
  const slide6 = pptx.addSlide();
  addSlideHeader(slide6, colors, 'Comprehensive Suite of Digital Transformation Services', null);
  
  const services = (data.serviceLines || '').split('\n').filter(s => s.trim()).slice(0, 6);
  
  services.forEach((service, idx) => {
    const parts = service.split('|').map(p => p.trim());
    const row = Math.floor(idx / 3);
    const col = idx % 3;
    const xPos = 0.4 + (col * 3.15);
    const yPos = 1.2 + (row * 1.9);
    
    slide6.addShape('rect', {
      x: xPos, y: yPos, w: 3, h: 1.7,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    // Colored top bar using theme colors
    slide6.addShape('rect', {
      x: xPos, y: yPos, w: 3, h: 0.08,
      fill: { color: colors.chartColors[idx % colors.chartColors.length] }
    });
    
    slide6.addText(truncateText(parts[0] || `Service ${idx + 1}`, 30), {
      x: xPos + 0.15, y: yPos + 0.15, w: 2.7, h: 0.4,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    if (parts[1]) {
      slide6.addText(parts[1], {
        x: xPos + 0.15, y: yPos + 0.55, w: 2.7, h: 0.3,
        fontSize: 18, bold: true, color: colors.chartColors[idx % colors.chartColors.length], fontFace: 'Arial'
      });
    }
    
    if (parts[2]) {
      slide6.addText(truncateText(parts[2], 50), {
        x: xPos + 0.15, y: yPos + 0.9, w: 2.7, h: 0.7,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial', valign: 'top'
      });
    }
  });
  
  addSlideFooter(slide6, colors, slideNumber);

  // ============================================================================
  // SLIDE 7: CLIENT PORTFOLIO (Page 7)
  // ============================================================================
  slideNumber++;
  const slide7 = pptx.addSlide();
  addSlideHeader(slide7, colors, 'Strong Client Relationships with Marquee Enterprise Clients', null);
  
  const clientMetrics = [
    { label: 'Primary Vertical', value: (data.primaryVertical || 'BFSI').toUpperCase(), subvalue: `${data.primaryVerticalPct || 80}%` },
    { label: 'Top 10 Concentration', value: `${data.top10Concentration || 65}%`, subvalue: '' },
    { label: 'Net Retention Rate', value: `${data.netRetention || 137}%`, subvalue: '' }
  ];
  
  clientMetrics.forEach((metric, idx) => {
    const xPos = 0.5 + (idx * 3.15);
    slide7.addShape('rect', {
      x: xPos, y: 1.1, w: 2.9, h: 1,
      fill: { color: colors.chartColors[idx % 3] }
    });
    slide7.addText(metric.label, {
      x: xPos + 0.1, y: 1.15, w: 2.7, h: 0.25,
      fontSize: 10, color: colors.white, fontFace: 'Arial'
    });
    slide7.addText(metric.value, {
      x: xPos + 0.1, y: 1.4, w: 2.7, h: 0.5,
      fontSize: 22, bold: true, color: colors.white, fontFace: 'Arial', align: 'center'
    });
    if (metric.subvalue) {
      slide7.addText(metric.subvalue, {
        x: xPos + 0.1, y: 1.85, w: 2.7, h: 0.2,
        fontSize: 11, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 20
      });
    }
  });
  
  slide7.addText('Marquee Clients', {
    x: 0.5, y: 2.3, w: 9, h: 0.35,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const clients = (data.topClients || '').split('\n').filter(c => c.trim()).slice(0, 8);
  clients.forEach((client, idx) => {
    const parts = client.split('|').map(p => p.trim());
    const row = Math.floor(idx / 4);
    const col = idx % 4;
    const xPos = 0.5 + (col * 2.35);
    const yPos = 2.75 + (row * 1.1);
    
    slide7.addShape('rect', {
      x: xPos, y: yPos, w: 2.2, h: 0.95,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    slide7.addText(truncateText(parts[0] || '', 18), {
      x: xPos + 0.1, y: yPos + 0.1, w: 2, h: 0.4,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide7.addText(`${parts[1] || ''} | Since ${parts[2] || ''}`, {
      x: xPos + 0.1, y: yPos + 0.55, w: 2, h: 0.3,
      fontSize: 8, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide7, colors, slideNumber);

  // ============================================================================
  // SLIDE 8: FINANCIAL OVERVIEW - Fixed Pie Charts (Issue #3)
  // ============================================================================
  slideNumber++;
  const slide8 = pptx.addSlide();
  addSlideHeader(slide8, colors, 'Growing Revenue Contribution from Product Engineering and AI Solutions', null);
  
  // Parse revenue by service from user input
  const revenueByService = (data.revenueByService || data.serviceLines || '').split('\n')
    .filter(s => s.trim())
    .map(s => {
      const parts = s.split('|').map(p => p.trim());
      const name = parts[0] || 'Service';
      const pctMatch = (parts[1] || '0').match(/(\d+)/);
      const pct = pctMatch ? parseInt(pctMatch[1]) : 10;
      return { name: truncateText(name, 20), pct };
    })
    .slice(0, 6);
  
  // Ensure we have data
  if (revenueByService.length === 0) {
    revenueByService.push(
      { name: 'Cloud & Automation', pct: 39 },
      { name: 'Managed Services', pct: 31 },
      { name: 'Product Engineering', pct: 16 },
      { name: 'AI & Data', pct: 14 }
    );
  }

  // CHART 1: Revenue by Service Lines - Using actual PptxGenJS chart (Issue #3)
  slide8.addText('Revenue by Service Lines (FY25)', {
    x: 0.3, y: 1.15, w: 3, h: 0.3,
    fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 0.3, y: 1.42, w: 1.2, h: 0.03,
    fill: { color: colors.secondary }
  });
  
  // Create proper doughnut chart with multiple colors
  const serviceChartData = [{
    name: 'Revenue',
    labels: revenueByService.map(s => s.name),
    values: revenueByService.map(s => s.pct)
  }];
  
  slide8.addChart(pptx.charts.DOUGHNUT, serviceChartData, {
    x: 0.3, y: 1.55, w: 2.8, h: 2.4,
    holeSize: 50,
    showLabel: false,
    showValue: false,
    showPercent: false,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 7,
    chartColors: colors.chartColors.slice(0, revenueByService.length)
  });

  // CHART 2: Revenue by Platform
  slide8.addText('Revenue by Platforms (FY25)', {
    x: 3.4, y: 1.15, w: 3, h: 0.3,
    fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 3.4, y: 1.42, w: 1.2, h: 0.03,
    fill: { color: colors.secondary }
  });
  
  const platformData = [{
    name: 'Platform',
    labels: ['AWS', 'Azure', 'GCP', 'Other'],
    values: [81, 10, 5, 4]
  }];
  
  slide8.addChart(pptx.charts.DOUGHNUT, platformData, {
    x: 3.4, y: 1.55, w: 2.8, h: 2.4,
    holeSize: 50,
    showLabel: false,
    showValue: false,
    showPercent: false,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 7,
    chartColors: [colors.primary, colors.secondary, colors.chartColors[2], colors.chartColors[3]]
  });

  // CHART 3: Revenue by Pricing Model
  slide8.addText('Revenue by Pricing Models (FY25)', {
    x: 6.5, y: 1.15, w: 3, h: 0.3,
    fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 6.5, y: 1.42, w: 1.2, h: 0.03,
    fill: { color: colors.secondary }
  });
  
  const pricingData = [{
    name: 'Pricing',
    labels: ['T&M', 'Fixed Price', 'Managed', 'Products'],
    values: [75, 12, 10, 3]
  }];
  
  slide8.addChart(pptx.charts.DOUGHNUT, pricingData, {
    x: 6.5, y: 1.55, w: 2.8, h: 2.4,
    holeSize: 50,
    showLabel: false,
    showValue: false,
    showPercent: false,
    showLegend: true,
    legendPos: 'b',
    legendFontSize: 7,
    chartColors: [colors.secondary, colors.chartColors[2], colors.chartColors[3], colors.chartColors[4]]
  });
  
  addSlideFooter(slide8, colors, slideNumber);

  // ============================================================================
  // SLIDE 9: CASE STUDY 1 (Issue #2 - Include all case studies)
  // ============================================================================
  if (data.cs1Client) {
    slideNumber++;
    const slideCS1 = pptx.addSlide();
    addCaseStudySlide(slideCS1, colors, slideNumber, {
      client: data.cs1Client,
      challenge: data.cs1Challenge,
      solution: data.cs1Solution,
      results: data.cs1Results
    });
  }

  // ============================================================================
  // SLIDE 10: CASE STUDY 2 (Issue #2 - Now included!)
  // ============================================================================
  if (data.cs2Client) {
    slideNumber++;
    const slideCS2 = pptx.addSlide();
    addCaseStudySlide(slideCS2, colors, slideNumber, {
      client: data.cs2Client,
      challenge: data.cs2Challenge,
      solution: data.cs2Solution,
      results: data.cs2Results
    });
  }

  // ============================================================================
  // SLIDE: COMPETITIVE ADVANTAGES
  // ============================================================================
  slideNumber++;
  const slideAdv = pptx.addSlide();
  addSlideHeader(slideAdv, colors, 'Key Competitive Advantages', null);
  
  const advantages = (data.competitiveAdvantages || '').split('\n').filter(a => a.trim()).slice(0, 6);
  
  advantages.forEach((advantage, idx) => {
    const row = Math.floor(idx / 2);
    const col = idx % 2;
    const xPos = 0.4 + (col * 4.8);
    const yPos = 1.2 + (row * 1.4);
    
    slideAdv.addShape('rect', {
      x: xPos, y: yPos, w: 4.5, h: 1.2,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    slideAdv.addShape('rect', {
      x: xPos, y: yPos, w: 0.08, h: 1.2,
      fill: { color: colors.primary }
    });
    
    slideAdv.addShape('ellipse', {
      x: xPos + 0.2, y: yPos + 0.1, w: 0.4, h: 0.4,
      fill: { color: colors.primary }
    });
    slideAdv.addText(`${idx + 1}`, {
      x: xPos + 0.2, y: yPos + 0.1, w: 0.4, h: 0.4,
      fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
    
    const parts = advantage.split('|').map(p => p.trim());
    slideAdv.addText(truncateText(parts[0] || advantage, 40), {
      x: xPos + 0.7, y: yPos + 0.15, w: 3.6, h: 0.35,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    if (parts[1]) {
      slideAdv.addText(truncateText(parts[1], 60), {
        x: xPos + 0.7, y: yPos + 0.55, w: 3.6, h: 0.55,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial', valign: 'top'
      });
    }
  });
  
  addSlideFooter(slideAdv, colors, slideNumber);

  // ============================================================================
  // SLIDE: GROWTH STRATEGY
  // ============================================================================
  slideNumber++;
  const slideGrowth = pptx.addSlide();
  addSlideHeader(slideGrowth, colors, 'Strategic Growth Roadmap', null);
  
  slideGrowth.addText('Key Growth Drivers', {
    x: 0.4, y: 1.15, w: 4.2, h: 0.35,
    fontSize: 13, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const drivers = (data.growthDrivers || '').split('\n').filter(d => d.trim()).slice(0, 5);
  drivers.forEach((driver, idx) => {
    slideGrowth.addShape('rect', {
      x: 0.4, y: 1.55 + (idx * 0.55), w: 4.2, h: 0.45,
      fill: { color: idx % 2 === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slideGrowth.addText(`${idx + 1}. ${truncateText(driver.trim(), 50)}`, {
      x: 0.5, y: 1.55 + (idx * 0.55), w: 4, h: 0.45,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'middle'
    });
  });
  
  // Short-term goals
  slideGrowth.addShape('rect', {
    x: 4.9, y: 1.15, w: 2.4, h: 2.5,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 1 }
  });
  slideGrowth.addText('Short-Term Goals', {
    x: 5, y: 1.2, w: 2.2, h: 0.35,
    fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slideGrowth.addText('(0-12 months)', {
    x: 5, y: 1.5, w: 2.2, h: 0.25,
    fontSize: 9, color: colors.textLight, fontFace: 'Arial'
  });
  
  const shortGoals = (data.shortTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 4);
  shortGoals.forEach((goal, idx) => {
    slideGrowth.addText(`• ${truncateText(goal.trim(), 30)}`, {
      x: 5, y: 1.85 + (idx * 0.45), w: 2.2, h: 0.4,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Medium-term goals
  slideGrowth.addShape('rect', {
    x: 7.5, y: 1.15, w: 2.2, h: 2.5,
    fill: { color: colors.lightBg },
    line: { color: colors.secondary, width: 1 }
  });
  slideGrowth.addText('Medium-Term Goals', {
    x: 7.6, y: 1.2, w: 2, h: 0.35,
    fontSize: 11, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  slideGrowth.addText('(1-3 years)', {
    x: 7.6, y: 1.5, w: 2, h: 0.25,
    fontSize: 9, color: colors.textLight, fontFace: 'Arial'
  });
  
  const mediumGoals = (data.mediumTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 4);
  mediumGoals.forEach((goal, idx) => {
    slideGrowth.addText(`• ${truncateText(goal.trim(), 28)}`, {
      x: 7.6, y: 1.85 + (idx * 0.45), w: 2, h: 0.4,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slideGrowth, colors, slideNumber);

  // ============================================================================
  // SLIDE: SYNERGIES - Now respects Target Buyer Type (Issue #7)
  // ============================================================================
  slideNumber++;
  const slideSyn = pptx.addSlide();
  addSlideHeader(slideSyn, colors, 'Potential Synergies for Acquirers', null);
  
  const showStrategic = !targetBuyers.length || targetBuyers.includes('strategic');
  const showFinancial = !targetBuyers.length || targetBuyers.includes('financial');
  
  if (showStrategic) {
    const synWidth = showFinancial ? 4.5 : 9;
    slideSyn.addShape('rect', {
      x: 0.4, y: 1.2, w: synWidth, h: 3.6,
      fill: { color: colors.white },
      line: { color: colors.primary, width: 1.5 }
    });
    
    slideSyn.addShape('rect', {
      x: 0.4, y: 1.2, w: synWidth, h: 0.5,
      fill: { color: colors.primary }
    });
    slideSyn.addText('For Strategic Buyers', {
      x: 0.5, y: 1.25, w: synWidth - 0.2, h: 0.4,
      fontSize: 14, bold: true, color: colors.white, fontFace: 'Arial'
    });
    
    const strategicSynergies = (data.synergiesStrategic || '').split('\n').filter(s => s.trim()).slice(0, 6);
    strategicSynergies.forEach((synergy, idx) => {
      slideSyn.addText(`✓ ${truncateText(synergy.trim(), showFinancial ? 45 : 90)}`, {
        x: 0.6, y: 1.8 + (idx * 0.5), w: synWidth - 0.4, h: 0.45,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  if (showFinancial) {
    const xStart = showStrategic ? 5.1 : 0.4;
    const synWidth = showStrategic ? 4.5 : 9;
    
    slideSyn.addShape('rect', {
      x: xStart, y: 1.2, w: synWidth, h: 3.6,
      fill: { color: colors.white },
      line: { color: colors.secondary, width: 1.5 }
    });
    
    slideSyn.addShape('rect', {
      x: xStart, y: 1.2, w: synWidth, h: 0.5,
      fill: { color: colors.secondary }
    });
    slideSyn.addText('For Financial Investors', {
      x: xStart + 0.1, y: 1.25, w: synWidth - 0.2, h: 0.4,
      fontSize: 14, bold: true, color: colors.white, fontFace: 'Arial'
    });
    
    const financialSynergies = (data.synergiesFinancial || '').split('\n').filter(s => s.trim()).slice(0, 6);
    financialSynergies.forEach((synergy, idx) => {
      slideSyn.addText(`✓ ${truncateText(synergy.trim(), showStrategic ? 45 : 90)}`, {
        x: xStart + 0.2, y: 1.8 + (idx * 0.5), w: synWidth - 0.4, h: 0.45,
        fontSize: 10, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  addSlideFooter(slideSyn, colors, slideNumber);

  // ============================================================================
  // OPTIONAL: CONTENT VARIANT SLIDES (Issue #8)
  // ============================================================================
  if (variants.includes('financial') && variantContent.financialFocus) {
    slideNumber++;
    const slideFinFocus = pptx.addSlide();
    addVariantSlide(slideFinFocus, colors, slideNumber, variantContent.financialFocus);
  }
  
  if (variants.includes('tech') && variantContent.techFocus) {
    slideNumber++;
    const slideTechFocus = pptx.addSlide();
    addVariantSlide(slideTechFocus, colors, slideNumber, variantContent.techFocus);
  }

  // ============================================================================
  // OPTIONAL: APPENDIX SLIDES (Issue #8)
  // ============================================================================
  if (appendixOptions.includes('team-bios') && data.leadershipTeam) {
    slideNumber++;
    const slideTeam = pptx.addSlide();
    addSlideHeader(slideTeam, colors, 'Appendix: Leadership Team Details', null);
    
    const teamMembers = data.leadershipTeam.split('\n').filter(t => t.trim()).slice(0, 8);
    teamMembers.forEach((member, idx) => {
      const parts = member.split('|').map(p => p.trim());
      const row = Math.floor(idx / 2);
      const col = idx % 2;
      
      slideTeam.addShape('rect', {
        x: 0.4 + (col * 4.8), y: 1.2 + (row * 0.9), w: 4.5, h: 0.8,
        fill: { color: colors.lightBg },
        line: { color: colors.border, width: 0.5 }
      });
      slideTeam.addText(truncateText(parts[0] || '', 25), {
        x: 0.5 + (col * 4.8), y: 1.25 + (row * 0.9), w: 4.3, h: 0.35,
        fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
      });
      slideTeam.addText(`${parts[1] || ''} | ${parts[2] || ''}`, {
        x: 0.5 + (col * 4.8), y: 1.6 + (row * 0.9), w: 4.3, h: 0.3,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial'
      });
    });
    
    addSlideFooter(slideTeam, colors, slideNumber);
  }
  
  if (appendixOptions.includes('client-list') && data.topClients) {
    slideNumber++;
    const slideClients = pptx.addSlide();
    addSlideHeader(slideClients, colors, 'Appendix: Complete Client List', null);
    
    const allClients = data.topClients.split('\n').filter(c => c.trim());
    allClients.slice(0, 16).forEach((client, idx) => {
      const parts = client.split('|').map(p => p.trim());
      const row = Math.floor(idx / 4);
      const col = idx % 4;
      
      slideClients.addShape('rect', {
        x: 0.3 + (col * 2.4), y: 1.2 + (row * 0.9), w: 2.2, h: 0.8,
        fill: { color: colors.white },
        line: { color: colors.border, width: 0.5 }
      });
      slideClients.addText(truncateText(parts[0] || '', 18), {
        x: 0.4 + (col * 2.4), y: 1.3 + (row * 0.9), w: 2, h: 0.35,
        fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial'
      });
      slideClients.addText(`${parts[1] || ''} | ${parts[2] || ''}`, {
        x: 0.4 + (col * 2.4), y: 1.65 + (row * 0.9), w: 2, h: 0.25,
        fontSize: 7, color: colors.textLight, fontFace: 'Arial'
      });
    });
    
    addSlideFooter(slideClients, colors, slideNumber);
  }

  // ============================================================================
  // SLIDE: THANK YOU (Final slide - no page number in top left)
  // ============================================================================
  slideNumber++;
  const slideEnd = pptx.addSlide();
  
  slideEnd.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  slideEnd.addShape('rect', {
    x: 0, y: 2.4, w: 10, h: 0.02,
    fill: { color: colors.secondary }
  });
  
  slideEnd.addText('Thank You', {
    x: 0, y: 1.8, w: '100%', h: 0.8,
    fontSize: 48, bold: true, color: colors.white, fontFace: 'Arial', align: 'center'
  });
  
  slideEnd.addText(`For further information, please contact:\n${data.advisor || 'Investment Bank'}`, {
    x: 0, y: 2.7, w: '100%', h: 0.8,
    fontSize: 16, color: colors.white, fontFace: 'Arial', align: 'center', lineSpacingMultiple: 1.5
  });
  
  slideEnd.addText('Strictly Private and Confidential', {
    x: 0, y: 4.5, w: '100%', h: 0.3,
    fontSize: 10, italic: true, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 40
  });

  return { pptx, slideCount: slideNumber };
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function formatDate(dateStr) {
  if (!dateStr) return new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  try {
    const date = new Date(dateStr);
    return date.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  } catch {
    return dateStr;
  }
}

function addSlideHeader(slide, colors, title, sectionNumber) {
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.white }
  });
  
  // Consistent colored sidebar - NO section number text (Issue #1 fix)
  slide.addShape('rect', {
    x: 0, y: 0, w: 0.25, h: 1,
    fill: { color: colors.secondary }
  });
  
  slide.addText(title, {
    x: 0.4, y: 0.15, w: 9.2, h: 0.8,
    fontSize: 22, bold: true, color: colors.primary, fontFace: 'Arial', valign: 'middle'
  });
  
  slide.addShape('rect', {
    x: 0.4, y: 0.95, w: 9.2, h: 0.04,
    fill: { color: colors.accent }
  });
}

function addSlideFooter(slide, colors, pageNumber) {
  slide.addShape('rect', {
    x: 0, y: 5.1, w: '100%', h: 0.02,
    fill: { color: colors.primary }
  });
  
  slide.addText('Strictly Private & Confidential', {
    x: 0.3, y: 5.15, w: 3, h: 0.25,
    fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  // Consistent page number in bottom right (Issue #1 fix)
  slide.addText(`${pageNumber}`, {
    x: 9.2, y: 5.15, w: 0.5, h: 0.25,
    fontSize: 10, color: colors.primary, fontFace: 'Arial', align: 'right'
  });
}

function addCaseStudySlide(slide, colors, pageNumber, caseStudy) {
  addSlideHeader(slide, colors, `Case Study: ${caseStudy.client}`, null);
  
  // Client info sidebar
  slide.addShape('rect', {
    x: 0.3, y: 1.2, w: 2.5, h: 3.6,
    fill: { color: colors.white },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide.addShape('rect', {
    x: 0.5, y: 1.4, w: 2.1, h: 1,
    fill: { color: colors.lightBg }
  });
  slide.addText(truncateText(caseStudy.client, 20), {
    x: 0.5, y: 1.7, w: 2.1, h: 0.5,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
  });
  
  const clientInfo = [
    { label: 'Customer since:', value: '2020' },
    { label: 'Industry:', value: 'Financial Services' },
    { label: 'Platform:', value: 'AWS' }
  ];
  
  clientInfo.forEach((info, idx) => {
    slide.addText(info.label, {
      x: 0.5, y: 2.6 + (idx * 0.5), w: 2.1, h: 0.25,
      fontSize: 9, color: colors.textLight, fontFace: 'Arial'
    });
    slide.addText(info.value, {
      x: 0.5, y: 2.85 + (idx * 0.5), w: 2.1, h: 0.25,
      fontSize: 10, bold: true, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Challenge section
  slide.addShape('rect', {
    x: 3, y: 1.2, w: 3, h: 1.5,
    fill: { color: colors.white },
    line: { color: colors.border, width: 0.5 }
  });
  slide.addText('Challenges', {
    x: 3.1, y: 1.25, w: 2.8, h: 0.35,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide.addShape('rect', {
    x: 3, y: 1.55, w: 0.08, h: 0.04,
    fill: { color: colors.danger }
  });
  slide.addText(truncateText(caseStudy.challenge || 'Challenge description', 120), {
    x: 3.1, y: 1.65, w: 2.8, h: 1,
    fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Solutions section
  slide.addShape('rect', {
    x: 6.2, y: 1.2, w: 3.3, h: 1.5,
    fill: { color: colors.white },
    line: { color: colors.border, width: 0.5 }
  });
  slide.addText('Solutions', {
    x: 6.3, y: 1.25, w: 3.1, h: 0.35,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide.addShape('rect', {
    x: 6.2, y: 1.55, w: 0.08, h: 0.04,
    fill: { color: colors.primary }
  });
  slide.addText(truncateText(caseStudy.solution || 'Solution description', 130), {
    x: 6.3, y: 1.65, w: 3.1, h: 1,
    fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Results section
  slide.addText('Results', {
    x: 3, y: 2.85, w: 6.5, h: 0.35,
    fontSize: 12, bold: true, color: colors.success, fontFace: 'Arial'
  });
  slide.addShape('rect', {
    x: 3, y: 3.15, w: 0.08, h: 0.04,
    fill: { color: colors.success }
  });
  
  slide.addShape('rect', {
    x: 3, y: 3.25, w: 6.5, h: 1.5,
    fill: { color: 'E8F5E9' },
    line: { color: colors.success, width: 0.5 }
  });
  
  const results = (caseStudy.results || '').split('\n').filter(r => r.trim()).slice(0, 5);
  results.forEach((result, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    slide.addText(`✓ ${truncateText(result.trim(), 45)}`, {
      x: 3.1 + (col * 3.2), y: 3.35 + (row * 0.5), w: 3, h: 0.45,
      fontSize: 10, color: colors.success, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide, colors, pageNumber);
}

function addVariantSlide(slide, colors, pageNumber, content) {
  addSlideHeader(slide, colors, content.title, null);
  
  content.points.forEach((point, idx) => {
    slide.addShape('rect', {
      x: 0.5, y: 1.2 + (idx * 0.7), w: 9, h: 0.6,
      fill: { color: idx % 2 === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide.addText(`${idx + 1}. ${point}`, {
      x: 0.6, y: 1.3 + (idx * 0.7), w: 8.8, h: 0.4,
      fontSize: 12, color: colors.text, fontFace: 'Arial', valign: 'middle'
    });
  });
  
  addSlideFooter(slide, colors, pageNumber);
}

// ============================================================================
// API ENDPOINTS
// ============================================================================

// Health check
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'healthy', 
    timestamp: new Date().toISOString(),
    version: '3.0.0',
    pptxEnabled: true,
    enhancedDesign: true,
    fixes: [
      'Consistent slide numbering',
      'All case studies included',
      'Proper pie charts with segments',
      'Text truncation to prevent overflow',
      'Generic template names',
      'Theme colors properly applied',
      'Target buyer type affects content',
      'Content variants and appendix support',
      'Anthropic usage tracking'
    ]
  });
});

// Usage statistics endpoint (Issue #9)
app.get('/api/usage', (req, res) => {
  res.json({
    ...usageStats,
    totalCostUSD: usageStats.totalCostUSD.toFixed(4),
    averageCostPerCall: usageStats.totalCalls > 0 
      ? (usageStats.totalCostUSD / usageStats.totalCalls).toFixed(6) 
      : '0.000000'
  });
});

// Reset usage stats
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

// Get available themes (Issue #5 - Generic names)
app.get('/api/themes', (req, res) => {
  const themes = Object.entries(THEMES).map(([id, theme]) => ({
    id,
    name: theme.name,
    primaryColor: `#${theme.primary}`,
    secondaryColor: `#${theme.secondary}`
  }));
  res.json(themes);
});

// Generate enhanced PPTX
app.post('/api/generate-pptx', async (req, res) => {
  try {
    const { data, theme = 'modern-blue' } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    console.log('Generating Enhanced PPTX v3 for:', data.projectCodename || 'Unknown Project');
    console.log('Theme:', theme);
    console.log('Target Buyers:', data.targetBuyerType || 'All');
    console.log('Variants:', data.generateVariants || 'None');
    console.log('Appendix:', data.includeAppendix || 'None');

    const { pptx, slideCount } = await generateProfessionalPPTX(data, theme);
    
    const filename = `${data.projectCodename || 'IM'}_${Date.now()}.pptx`;
    const filepath = path.join(tempDir, filename);
    
    await pptx.writeFile(filepath);
    
    const fileBuffer = fs.readFileSync(filepath);
    const base64 = fileBuffer.toString('base64');
    
    // Cleanup
    fs.unlinkSync(filepath);
    
    res.json({
      success: true,
      filename: filename,
      fileData: base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      generatedAt: new Date().toISOString(),
      theme: theme,
      slideCount: slideCount,
      version: '3.0.0'
    });

  } catch (error) {
    console.error('Error generating PPTX:', error);
    res.status(500).json({ 
      error: 'Failed to generate PowerPoint', 
      details: error.message 
    });
  }
});

// Initialize Anthropic client for IM generation
const SYSTEM_PROMPT = `You are an expert Investment Banking Analyst specializing in creating professional Information Memorandums (IMs) for M&A transactions. You work for an investment banking firm and help automate the creation of management presentations for potential acquirers.

Your task is to take the structured data provided and generate professional, investment-banking-quality content for an Information Memorandum.

When generating content:
1. Use formal, professional investment banking language and terminology
2. Focus on investment highlights and value creation potential
3. Quantify everything possible with specific metrics and numbers
4. Structure content for easy PowerPoint slide conversion
5. Highlight growth potential, competitive advantages, and synergies
6. Tailor messaging based on the target buyer type (Strategic, Financial, International)

Output your response as a well-structured JSON object that can be directly used to populate presentation slides.`;

// Generate IM content
app.post('/api/generate-im', async (req, res) => {
  try {
    const { data } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: 'API key not configured' });
    }

    console.log('Generating IM for:', data.projectCodename || 'Unknown Project');

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 8000,
      system: SYSTEM_PROMPT,
      messages: [
        {
          role: 'user',
          content: `Please generate a professional Information Memorandum based on the following company data. Structure your response as a JSON object that can be used to populate presentation slides.

## Input Data:
${JSON.stringify(data, null, 2)}

Generate the following sections:
1. Executive Summary with 5-7 investment highlights
2. Company Overview with key statistics
3. Founder/Leadership profiles
4. Service offerings with capabilities and metrics
5. Client case studies with quantified results
6. Financial overview with growth metrics
7. Competitive positioning and differentiators
8. Growth strategy and potential synergies

For each section, provide:
- Title
- Key bullet points
- Relevant metrics/numbers
- Suggested visuals (charts, logos, etc.)

Return as structured JSON.`
        }
      ]
    });

    // Track usage (Issue #9)
    const usage = message.usage || { input_tokens: 0, output_tokens: 0 };
    trackUsage('claude-sonnet-4-20250514', usage.input_tokens, usage.output_tokens, 'IM Generation');

    let generatedContent = message.content[0].text;
    
    let parsedContent;
    try {
      generatedContent = generatedContent
        .replace(/```json\n?/g, '')
        .replace(/```\n?/g, '')
        .trim();
      
      parsedContent = JSON.parse(generatedContent);
    } catch (parseError) {
      console.error('JSON parse error:', parseError);
      parsedContent = { 
        rawContent: generatedContent,
        parseError: 'Content generated but could not be parsed as JSON'
      };
    }

    res.json({
      success: true,
      content: parsedContent,
      generatedAt: new Date().toISOString(),
      model: 'claude-sonnet-4-20250514',
      usage: {
        inputTokens: usage.input_tokens,
        outputTokens: usage.output_tokens
      }
    });

  } catch (error) {
    console.error('Error generating IM:', error);
    
    if (error.status === 401) {
      return res.status(401).json({ error: 'Invalid API key' });
    }
    if (error.status === 429) {
      return res.status(429).json({ error: 'Rate limit exceeded. Please try again later.' });
    }
    
    res.status(500).json({ 
      error: 'Failed to generate IM', 
      details: error.message 
    });
  }
});

// Validate data
app.post('/api/validate', async (req, res) => {
  try {
    const { data } = req.body;
    
    const validationResults = {
      errors: [],
      warnings: [],
      suggestions: []
    };

    const requiredFields = [
      { key: 'projectCodename', label: 'Project Codename', phase: 'Project Setup' },
      { key: 'companyName', label: 'Company Name', phase: 'Project Setup' },
      { key: 'foundedYear', label: 'Founded Year', phase: 'Company Overview' },
      { key: 'headquarters', label: 'Headquarters', phase: 'Company Overview' },
      { key: 'founderName', label: 'Founder Name', phase: 'Leadership' },
      { key: 'serviceLines', label: 'Service Lines', phase: 'Services & Products' },
      { key: 'revenueFY25', label: 'Revenue FY25', phase: 'Financials' }
    ];

    requiredFields.forEach(field => {
      if (!data[field.key]) {
        validationResults.errors.push({
          field: field.key,
          label: field.label,
          phase: field.phase,
          message: `${field.label} is required`
        });
      }
    });

    if (data.revenueFY25 && data.revenueFY26P) {
      const growth = ((data.revenueFY26P - data.revenueFY25) / data.revenueFY25) * 100;
      if (growth > 100) {
        validationResults.warnings.push({
          field: 'revenueFY26P',
          phase: 'Financials',
          message: `Projected growth of ${growth.toFixed(0)}% YoY is very aggressive. Please verify assumptions.`
        });
      }
    }

    const highlights = (data.investmentHighlights || '').split('\n').filter(l => l.trim()).length;
    if (highlights < 5) {
      validationResults.suggestions.push({
        field: 'investmentHighlights',
        phase: 'Company Overview',
        message: `Only ${highlights} investment highlights provided. Consider adding more (recommended: 5-7).`
      });
    }

    res.json(validationResults);

  } catch (error) {
    console.error('Error validating data:', error);
    res.status(500).json({ error: 'Validation failed', details: error.message });
  }
});

// Draft storage
const drafts = new Map();

app.post('/api/drafts', (req, res) => {
  try {
    const { data, projectId } = req.body;
    const id = projectId || `draft_${Date.now()}`;
    
    drafts.set(id, {
      data,
      savedAt: new Date().toISOString(),
      version: (drafts.get(id)?.version || 0) + 1
    });

    console.log(`Draft saved: ${id}`);
    res.json({ success: true, projectId: id, savedAt: new Date().toISOString() });
  } catch (error) {
    console.error('Error saving draft:', error);
    res.status(500).json({ error: 'Failed to save draft' });
  }
});

app.get('/api/drafts/:projectId', (req, res) => {
  try {
    const { projectId } = req.params;
    const draft = drafts.get(projectId);
    
    if (!draft) {
      return res.status(404).json({ error: 'Draft not found' });
    }

    res.json(draft);
  } catch (error) {
    console.error('Error retrieving draft:', error);
    res.status(500).json({ error: 'Failed to retrieve draft' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log('='.repeat(60));
  console.log('🚀 IM Creator API Server - ENHANCED v3.0');
  console.log('='.repeat(60));
  console.log(`📍 Port: ${PORT}`);
  console.log(`🔗 Health: http://localhost:${PORT}/api/health`);
  console.log(`🔑 API Key: ${process.env.ANTHROPIC_API_KEY ? 'Configured ✅' : 'NOT SET ❌'}`);
  console.log(`📊 PPTX Generation: Enhanced v3 ✅`);
  console.log(`🎨 Themes: ${Object.keys(THEMES).join(', ')}`);
  console.log(`💰 Usage Tracking: Enabled ✅`);
  console.log('='.repeat(60));
  console.log('Fixes Applied:');
  console.log('  1. ✅ Consistent slide numbering');
  console.log('  2. ✅ All case studies included');
  console.log('  3. ✅ Proper pie charts with segments');
  console.log('  4. ✅ Text truncation to prevent overflow');
  console.log('  5. ✅ Generic template names');
  console.log('  6. ✅ Theme colors properly applied');
  console.log('  7. ✅ Target buyer type affects content');
  console.log('  8. ✅ Content variants and appendix support');
  console.log('  9. ✅ Anthropic usage tracking');
  console.log('='.repeat(60));
});
