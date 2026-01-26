// Enhanced IM Creator Server with Professional PPTX Generation
// Uses html2pptx for high-quality Deloitte-style presentations

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
// PROFESSIONAL COLOR THEMES (Deloitte-inspired)
// ============================================================================
const THEMES = {
  'modern-tech': {
    primary: '2B579A',      // Deep blue (like Deloitte)
    secondary: '86BC25',    // Green accent
    accent: 'FFC72C',       // Gold/Yellow for highlights
    text: '333333',         // Dark gray text
    textLight: '666666',    // Light gray text
    white: 'FFFFFF',
    lightBg: 'F5F7FA',      // Light gray background
    darkBg: '1A1F36',       // Dark background for cover
    border: 'E0E5EC',
    success: '28A745',
    warning: 'FFC107',
    danger: 'DC3545',
    chartColors: ['2B579A', '86BC25', 'FFC72C', '00A3E0', 'E31B23', '6B3FA0']
  },
  'conservative': {
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
  'minimalist': {
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
  }
};

// ============================================================================
// PROFESSIONAL POWERPOINT GENERATOR
// ============================================================================
async function generateProfessionalPPTX(data, theme = 'modern-tech') {
  const pptx = new PptxGenJS();
  const colors = THEMES[theme] || THEMES['modern-tech'];
  
  // Presentation metadata
  pptx.author = data.advisor || 'RMB Securities';
  pptx.title = `${data.projectCodename || 'Project'} - Management Presentation`;
  pptx.subject = 'Confidential Information Memorandum';
  pptx.company = data.advisor || 'RMB Securities';
  pptx.layout = 'LAYOUT_16x9';

  // ============================================================================
  // SLIDE 1: COVER PAGE (Full-bleed dark design)
  // ============================================================================
  const slide1 = pptx.addSlide();
  
  // Dark gradient background
  slide1.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { type: 'solid', color: colors.darkBg }
  });
  
  // Decorative geometric pattern (top right)
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
  
  // Project codename (large)
  slide1.addText(data.projectCodename || 'Project Phoenix', {
    x: 0.5, y: 2.2, w: 8, h: 1,
    fontSize: 48, bold: true, color: colors.white,
    fontFace: 'Arial'
  });
  
  // Management Presentation subtitle
  slide1.addText('Management Presentation', {
    x: 0.5, y: 3.35, w: 6, h: 0.5,
    fontSize: 22, color: colors.white,
    fontFace: 'Arial'
  });
  
  // Date
  slide1.addText(data.presentationDate || new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' }), {
    x: 0.5, y: 3.95, w: 4, h: 0.35,
    fontSize: 14, italic: true, color: colors.white,
    fontFace: 'Arial', transparency: 30
  });
  
  // Confidential notice
  slide1.addText('Strictly Private and Confidential', {
    x: 0.5, y: 4.85, w: 4, h: 0.3,
    fontSize: 11, italic: true, color: colors.white,
    fontFace: 'Arial', transparency: 40
  });
  
  // Advisor logo placeholder (bottom right)
  slide1.addText(data.advisor || 'RMB Securities', {
    x: 7, y: 4.7, w: 2.5, h: 0.4,
    fontSize: 14, bold: true, color: colors.white,
    fontFace: 'Arial', align: 'right'
  });

  // ============================================================================
  // SLIDE 2: DISCLAIMER (Clean white slide)
  // ============================================================================
  const slide2 = pptx.addSlide();
  addSlideHeader(slide2, colors, 'Important Notice', null);
  
  const disclaimerText = `The information contained in this document has been compiled by ${data.advisor || 'RMB Securities'} based on information obtained from public sources. Except in the general context of evaluating the capabilities of ${data.advisor || 'RMB Securities'}, no reliance may be placed for any purposes whatsoever on the contents of this document or on its completeness.

This document and its contents are confidential and may not be reproduced, redistributed or passed on, directly or indirectly, to any other person in whole or in part without the prior written consent of ${data.advisor || 'RMB Securities'}.

This document does not constitute an offer or agreement between ${data.advisor || 'RMB Securities'} and ${data.companyName || 'the Company'}. Furthermore, changes in Company definition of requirements will necessarily affect the proposal set forth herein.`;

  slide2.addText(disclaimerText, {
    x: 0.5, y: 1.3, w: 9, h: 3.5,
    fontSize: 11, color: colors.text, fontFace: 'Arial',
    valign: 'top', lineSpacingMultiple: 1.5
  });
  
  addSlideFooter(slide2, colors, 2);

  // ============================================================================
  // SLIDE 3: EXECUTIVE SUMMARY (Investment Highlights)
  // ============================================================================
  const slide3 = pptx.addSlide();
  addSlideHeader(slide3, colors, 'A Digital Transformation Partner Delivering Cloud, Product Engineering, and AI Solutions', 1);
  
  // Left column - Key stats
  const stats = [
    { value: data.foundedYear || '2014', label: 'Founded Year', icon: '📅' },
    { value: `${data.employeeCountFT || '350'}+`, label: `Headcount as of FY${new Date().getFullYear()}`, icon: '👥' },
    { value: '80+', label: 'Clients in FY2025', icon: '🏢' },
    { value: '98%', label: 'India Revenue in FY2025', icon: '🇮🇳' },
    { value: '300+', label: 'Successful Projects', icon: '✅' }
  ];
  
  // Stats column (left side)
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
  
  // Middle column - Key Offerings (circular diagram representation)
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
  
  const offerings = [
    'Cloud & Automation',
    'AI and Data Solutions',
    'Digitalization & Product Engineering',
    'Cloud Security',
    'Managed Services',
    'Digital Transformation'
  ];
  
  offerings.forEach((offering, idx) => {
    const row = Math.floor(idx / 2);
    const col = idx % 2;
    slide3.addShape('roundRect', {
      x: 2.75 + (col * 1.7), y: 1.6 + (row * 1), w: 1.6, h: 0.85,
      fill: { color: idx < 2 ? colors.primary : idx < 4 ? colors.secondary : colors.accent },
      line: { color: colors.border, width: 0 }
    });
    slide3.addText(offering, {
      x: 2.75 + (col * 1.7), y: 1.75 + (row * 1), w: 1.6, h: 0.55,
      fontSize: 9, color: idx < 4 ? colors.white : colors.text, fontFace: 'Arial',
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
  
  // Revenue bar chart data
  const revenueData = [
    { year: 'FY24', value: parseFloat(data.revenueFY24) || 54 },
    { year: 'FY25', value: parseFloat(data.revenueFY25) || 84 },
    { year: 'FY26P', value: parseFloat(data.revenueFY26P) || 140 },
    { year: 'FY27P', value: parseFloat(data.revenueFY27P) || 182 },
    { year: 'FY28P', value: parseFloat(data.revenueFY28P) || 236 }
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
  
  slide3.addText('In INR Cr', {
    x: 6.4, y: 1.55, w: 1, h: 0.2,
    fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  // CAGR indicator
  slide3.addText(`CAGR: ~30%`, {
    x: 8.2, y: 1.55, w: 1.2, h: 0.2,
    fontSize: 9, bold: true, color: colors.secondary, fontFace: 'Arial', align: 'right'
  });
  
  // EBITDA margin
  slide3.addText(`EBITDA Margin FY25: ${data.ebitdaMarginFY25 || 21}%`, {
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
  
  addSlideFooter(slide3, colors, 3);

  // ============================================================================
  // SLIDE 4: FOUNDER PROFILE
  // ============================================================================
  const slide4 = pptx.addSlide();
  addSlideHeader(slide4, colors, 'Founded & Led by Industry Veteran with Strong Educational Qualification & Industry Experience', null);
  
  // Photo placeholder (circular)
  slide4.addShape('ellipse', {
    x: 1.2, y: 1.5, w: 2, h: 2,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 2 }
  });
  slide4.addText('Photo', {
    x: 1.2, y: 2.3, w: 2, h: 0.4,
    fontSize: 12, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  // Founder name and title
  slide4.addText(data.founderName || 'Founder Name', {
    x: 0.7, y: 3.6, w: 3, h: 0.4,
    fontSize: 20, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
  });
  slide4.addText(data.founderTitle || 'Founder & CEO', {
    x: 0.7, y: 4, w: 3, h: 0.3,
    fontSize: 14, color: colors.secondary, fontFace: 'Arial', align: 'center'
  });
  slide4.addText(`~${data.founderExperience || 21} years of total experience`, {
    x: 0.7, y: 4.3, w: 3, h: 0.3,
    fontSize: 11, italic: true, color: colors.textLight, fontFace: 'Arial', align: 'center'
  });
  
  // Background info box
  slide4.addShape('rect', {
    x: 4, y: 1.3, w: 5.5, h: 3.5,
    fill: { color: colors.lightBg },
    line: { color: colors.border, width: 0.5 }
  });
  
  slide4.addText("Founder's Background", {
    x: 4.1, y: 1.35, w: 5.3, h: 0.35,
    fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial',
    fill: { color: colors.primary }
  });
  
  const backgroundPoints = [
    `Founded ${data.companyName || 'the Company'} in ${data.foundedYear || '2014'} and leads its strategic direction`,
    'Prior experience includes Director of Technology at Reprise Media, Senior Consultant at Wipro, and software engineering roles',
    'Visiting Faculty of Big Data and Analytics at IIT Bombay and Big Data at IIM',
    `Holds an MBA in Marketing from JBIMS and Bachelors in Engineering from VJTI`,
    'Emphasizes grounding technical innovation in real business value, aiming to democratize cloud & AI in organizations'
  ];
  
  backgroundPoints.forEach((point, idx) => {
    slide4.addText(`•  ${point}`, {
      x: 4.2, y: 1.8 + (idx * 0.55), w: 5.2, h: 0.5,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
  });
  
  // Previous experience logos
  slide4.addText('Previous Experience', {
    x: 4.1, y: 4.05, w: 5.3, h: 0.25,
    fontSize: 10, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  const companies = ['Wipro', 'IBM', 'Hexaware', 'Reprise'];
  companies.forEach((company, idx) => {
    slide4.addShape('rect', {
      x: 4.2 + (idx * 1.3), y: 4.35, w: 1.15, h: 0.45,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide4.addText(company, {
      x: 4.2 + (idx * 1.3), y: 4.35, w: 1.15, h: 0.45,
      fontSize: 9, color: colors.text, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
  });
  
  addSlideFooter(slide4, colors, 4);

  // ============================================================================
  // SLIDE 5: COMPANY TIMELINE
  // ============================================================================
  const slide5 = pptx.addSlide();
  addSlideHeader(slide5, colors, 'Evolving Continuously from Cloud Solutions to AI Agentic Solutions, ensuring Alignment with Technology', null);
  
  // Timeline base line
  slide5.addShape('rect', {
    x: 0.5, y: 2.8, w: 9, h: 0.03,
    fill: { color: colors.primary }
  });
  
  const timeline = [
    { period: '2014-16', icon: '🚀', title: 'Foundation', points: ['Incorporated', 'Offered AWS cloud solutions', 'Offices in Thane and Pune'] },
    { period: '2017-18', icon: '📈', title: 'Growth', points: ['Developed loan disbursement platform', 'Introduced continuous client support'] },
    { period: '2019-20', icon: '🏆', title: 'Recognition', points: ['Best BFSI Consulting Partner by AWS', 'Developed Atlas API Management'] },
    { period: '2021-22', icon: '📊', title: 'Expansion', points: ['Launched Big Data Practice', 'Expanded Bangalore operations'] },
    { period: '2023-24', icon: '🏦', title: 'Enterprise', points: ['Secured public sector bank projects', 'Product-first strategy'] },
    { period: '2025-26', icon: '🤖', title: 'AI Era', points: ['Launched AI agents for banking', 'Enterprise AI Practice'] }
  ];
  
  timeline.forEach((item, idx) => {
    const xPos = 0.7 + (idx * 1.55);
    
    // Timeline dot
    slide5.addShape('ellipse', {
      x: xPos + 0.55, y: 2.7, w: 0.2, h: 0.2,
      fill: { color: colors.primary }
    });
    
    // Year badge
    slide5.addShape('roundRect', {
      x: xPos + 0.1, y: 2.2, w: 1.1, h: 0.35,
      fill: { color: colors.primary }
    });
    slide5.addText(item.period, {
      x: xPos + 0.1, y: 2.2, w: 1.1, h: 0.35,
      fontSize: 10, bold: true, color: colors.white, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
    
    // Content below timeline
    item.points.forEach((point, pIdx) => {
      slide5.addText(`• ${point}`, {
        x: xPos, y: 3.0 + (pIdx * 0.35), w: 1.4, h: 0.35,
        fontSize: 8, color: colors.text, fontFace: 'Arial'
      });
    });
  });
  
  addSlideFooter(slide5, colors, 5);

  // ============================================================================
  // SLIDE 6: SERVICE OFFERINGS
  // ============================================================================
  const slide6 = pptx.addSlide();
  addSlideHeader(slide6, colors, 'Comprehensive Suite of Digital Transformation Services', 1);
  
  const services = (data.serviceLines || '').split('\n').filter(s => s.trim()).slice(0, 6);
  const serviceColors = [colors.primary, colors.secondary, colors.accent, '00A3E0', 'E31B23', '6B3FA0'];
  
  services.forEach((service, idx) => {
    const parts = service.split('|').map(p => p.trim());
    const row = Math.floor(idx / 3);
    const col = idx % 3;
    const xPos = 0.4 + (col * 3.15);
    const yPos = 1.2 + (row * 1.9);
    
    // Service card
    slide6.addShape('rect', {
      x: xPos, y: yPos, w: 3, h: 1.7,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 },
      shadow: { type: 'outer', blur: 3, offset: 2, angle: 45, color: '000000', opacity: 0.1 }
    });
    
    // Colored top bar
    slide6.addShape('rect', {
      x: xPos, y: yPos, w: 3, h: 0.08,
      fill: { color: serviceColors[idx] || colors.primary }
    });
    
    // Service name
    slide6.addText(parts[0] || `Service ${idx + 1}`, {
      x: xPos + 0.15, y: yPos + 0.15, w: 2.7, h: 0.4,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    // Revenue percentage
    if (parts[1]) {
      slide6.addText(parts[1], {
        x: xPos + 0.15, y: yPos + 0.55, w: 2.7, h: 0.3,
        fontSize: 18, bold: true, color: serviceColors[idx] || colors.secondary, fontFace: 'Arial'
      });
    }
    
    // Description
    if (parts[2]) {
      slide6.addText(parts[2], {
        x: xPos + 0.15, y: yPos + 0.9, w: 2.7, h: 0.7,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial', valign: 'top'
      });
    }
  });
  
  addSlideFooter(slide6, colors, 6);

  // ============================================================================
  // SLIDE 7: CLIENT PORTFOLIO (with donut chart)
  // ============================================================================
  const slide7 = pptx.addSlide();
  addSlideHeader(slide7, colors, 'Strong Client Relationships with Marquee Enterprise Clients', null);
  
  // Three metric boxes at top
  const clientMetrics = [
    { label: 'Primary Vertical', value: (data.primaryVertical || 'BFSI').toUpperCase(), subvalue: `${data.primaryVerticalPct || 75}%` },
    { label: 'Top 10 Concentration', value: `${data.top10Concentration || 72}%`, subvalue: '' },
    { label: 'Net Retention Rate', value: `${data.netRetention || 118}%`, subvalue: '' }
  ];
  
  clientMetrics.forEach((metric, idx) => {
    const xPos = 0.5 + (idx * 3.15);
    slide7.addShape('rect', {
      x: xPos, y: 1.1, w: 2.9, h: 1,
      fill: { color: idx === 0 ? colors.primary : idx === 1 ? colors.secondary : colors.primary }
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
  
  // Client logos grid
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
    
    slide7.addText(parts[0] || '', {
      x: xPos + 0.1, y: yPos + 0.1, w: 2, h: 0.4,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide7.addText(`${parts[1] || ''} | Since ${parts[2] || ''}`, {
      x: xPos + 0.1, y: yPos + 0.55, w: 2, h: 0.3,
      fontSize: 8, color: colors.textLight, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide7, colors, 7);

  // ============================================================================
  // SLIDE 8: FINANCIAL OVERVIEW (Revenue breakdown)
  // ============================================================================
  const slide8 = pptx.addSlide();
  addSlideHeader(slide8, colors, 'Growing Revenue Contribution from Product Engineering and AI Solutions', null);
  
  // Revenue by Service Lines (Donut chart - simulated with shapes)
  slide8.addText('Revenue by Service Lines (FY25)', {
    x: 0.5, y: 1.15, w: 3, h: 0.3,
    fontSize: 12, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 0.5, y: 1.45, w: 1.5, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  // Donut chart representation
  slide8.addShape('ellipse', {
    x: 0.8, y: 1.8, w: 2.2, h: 2.2,
    fill: { color: colors.primary }
  });
  slide8.addShape('ellipse', {
    x: 1.3, y: 2.3, w: 1.2, h: 1.2,
    fill: { color: colors.white }
  });
  
  // Service breakdown legend
  const serviceBreakdown = [
    { name: 'Cloud & Automation', pct: '39%', color: colors.primary },
    { name: 'Managed Services', pct: '31%', color: colors.secondary },
    { name: 'Digitalization', pct: '16%', color: colors.accent },
    { name: 'AI & Data', pct: '6%', color: '00A3E0' },
    { name: 'Cloud Security', pct: '5%', color: 'E31B23' },
    { name: 'Products', pct: '3%', color: '6B3FA0' }
  ];
  
  serviceBreakdown.forEach((svc, idx) => {
    slide8.addShape('rect', {
      x: 0.5, y: 4.15 + (idx * 0.22), w: 0.15, h: 0.15,
      fill: { color: svc.color }
    });
    slide8.addText(`${svc.name}  ${svc.pct}`, {
      x: 0.7, y: 4.12 + (idx * 0.22), w: 2.5, h: 0.2,
      fontSize: 8, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Revenue by Platform
  slide8.addText('Revenue by Platforms (FY25)', {
    x: 3.5, y: 1.15, w: 3, h: 0.3,
    fontSize: 12, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 3.5, y: 1.45, w: 1.5, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  // Platform donut
  slide8.addShape('ellipse', {
    x: 3.8, y: 1.8, w: 2.2, h: 2.2,
    fill: { color: colors.primary }
  });
  slide8.addShape('ellipse', {
    x: 4.3, y: 2.3, w: 1.2, h: 1.2,
    fill: { color: colors.white }
  });
  slide8.addText('AWS\n81%', {
    x: 4.3, y: 2.5, w: 1.2, h: 0.8,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center', valign: 'middle'
  });
  
  // Revenue by Pricing Model
  slide8.addText('Revenue by Pricing Models (FY25)', {
    x: 6.5, y: 1.15, w: 3, h: 0.3,
    fontSize: 12, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  slide8.addShape('rect', {
    x: 6.5, y: 1.45, w: 1.5, h: 0.04,
    fill: { color: colors.secondary }
  });
  
  // Pricing donut
  slide8.addShape('ellipse', {
    x: 6.8, y: 1.8, w: 2.2, h: 2.2,
    fill: { color: colors.secondary }
  });
  slide8.addShape('ellipse', {
    x: 7.3, y: 2.3, w: 1.2, h: 1.2,
    fill: { color: colors.white }
  });
  slide8.addText('T&M\n75%', {
    x: 7.3, y: 2.5, w: 1.2, h: 0.8,
    fontSize: 12, bold: true, color: colors.secondary, fontFace: 'Arial', align: 'center', valign: 'middle'
  });
  
  addSlideFooter(slide8, colors, 8);

  // ============================================================================
  // SLIDE 9: CASE STUDY
  // ============================================================================
  if (data.cs1Client) {
    const slide9 = pptx.addSlide();
    addSlideHeader(slide9, colors, `Case Study: ${data.cs1Client}`, 'Digitalization & Product Engineering');
    
    // Client info sidebar
    slide9.addShape('rect', {
      x: 0.3, y: 1.2, w: 2.5, h: 3.6,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    // Client logo placeholder
    slide9.addShape('rect', {
      x: 0.5, y: 1.4, w: 2.1, h: 1,
      fill: { color: colors.lightBg }
    });
    slide9.addText(data.cs1Client, {
      x: 0.5, y: 1.7, w: 2.1, h: 0.5,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial', align: 'center'
    });
    
    const clientInfo = [
      { label: 'Customer since:', value: '2020' },
      { label: 'Industry:', value: 'Financial Services' },
      { label: 'Platform:', value: 'AWS' }
    ];
    
    clientInfo.forEach((info, idx) => {
      slide9.addText(info.label, {
        x: 0.5, y: 2.6 + (idx * 0.5), w: 2.1, h: 0.25,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial'
      });
      slide9.addText(info.value, {
        x: 0.5, y: 2.85 + (idx * 0.5), w: 2.1, h: 0.25,
        fontSize: 10, bold: true, color: colors.text, fontFace: 'Arial'
      });
    });
    
    // Challenge section
    slide9.addShape('rect', {
      x: 3, y: 1.2, w: 3, h: 1.5,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide9.addText('Challenges', {
      x: 3.1, y: 1.25, w: 2.8, h: 0.35,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide9.addShape('rect', {
      x: 3, y: 1.55, w: 0.08, h: 0.04,
      fill: { color: colors.danger }
    });
    slide9.addText(data.cs1Challenge || 'Challenge description', {
      x: 3.1, y: 1.65, w: 2.8, h: 1,
      fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
    
    // Solutions section
    slide9.addShape('rect', {
      x: 6.2, y: 1.2, w: 3.3, h: 1.5,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide9.addText('Solutions', {
      x: 6.3, y: 1.25, w: 3.1, h: 0.35,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide9.addShape('rect', {
      x: 6.2, y: 1.55, w: 0.08, h: 0.04,
      fill: { color: colors.primary }
    });
    slide9.addText(data.cs1Solution || 'Solution description', {
      x: 6.3, y: 1.65, w: 3.1, h: 1,
      fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
    
    // Results section
    slide9.addText('Results', {
      x: 3, y: 2.85, w: 6.5, h: 0.35,
      fontSize: 12, bold: true, color: colors.success, fontFace: 'Arial'
    });
    slide9.addShape('rect', {
      x: 3, y: 3.15, w: 0.08, h: 0.04,
      fill: { color: colors.success }
    });
    
    slide9.addShape('rect', {
      x: 3, y: 3.25, w: 6.5, h: 1.5,
      fill: { color: '#E8F5E9' },
      line: { color: colors.success, width: 0.5 }
    });
    
    const results = (data.cs1Results || '').split('\n').filter(r => r.trim()).slice(0, 4);
    results.forEach((result, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      slide9.addText(`✓ ${result.trim()}`, {
        x: 3.1 + (col * 3.2), y: 3.35 + (row * 0.6), w: 3, h: 0.5,
        fontSize: 10, color: colors.success, fontFace: 'Arial'
      });
    });
    
    addSlideFooter(slide9, colors, 9);
  }

  // ============================================================================
  // SLIDE 10: COMPETITIVE ADVANTAGES
  // ============================================================================
  const slide10 = pptx.addSlide();
  addSlideHeader(slide10, colors, 'Key Competitive Advantages', 2);
  
  const advantages = (data.competitiveAdvantages || '').split('\n').filter(a => a.trim()).slice(0, 6);
  
  advantages.forEach((advantage, idx) => {
    const row = Math.floor(idx / 2);
    const col = idx % 2;
    const xPos = 0.4 + (col * 4.8);
    const yPos = 1.2 + (row * 1.4);
    
    // Card with left accent
    slide10.addShape('rect', {
      x: xPos, y: yPos, w: 4.5, h: 1.2,
      fill: { color: colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    
    // Left accent bar
    slide10.addShape('rect', {
      x: xPos, y: yPos, w: 0.08, h: 1.2,
      fill: { color: colors.primary }
    });
    
    // Number badge
    slide10.addShape('ellipse', {
      x: xPos + 0.2, y: yPos + 0.1, w: 0.4, h: 0.4,
      fill: { color: colors.primary }
    });
    slide10.addText(`${idx + 1}`, {
      x: xPos + 0.2, y: yPos + 0.1, w: 0.4, h: 0.4,
      fontSize: 12, bold: true, color: colors.white, fontFace: 'Arial', align: 'center', valign: 'middle'
    });
    
    const parts = advantage.split('|').map(p => p.trim());
    slide10.addText(parts[0] || advantage, {
      x: xPos + 0.7, y: yPos + 0.15, w: 3.6, h: 0.35,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    if (parts[1]) {
      slide10.addText(parts[1], {
        x: xPos + 0.7, y: yPos + 0.55, w: 3.6, h: 0.55,
        fontSize: 9, color: colors.textLight, fontFace: 'Arial', valign: 'top'
      });
    }
  });
  
  addSlideFooter(slide10, colors, 10);

  // ============================================================================
  // SLIDE 11: GROWTH STRATEGY
  // ============================================================================
  const slide11 = pptx.addSlide();
  addSlideHeader(slide11, colors, 'Strategic Growth Roadmap', null);
  
  // Growth drivers
  slide11.addText('Key Growth Drivers', {
    x: 0.4, y: 1.15, w: 4.2, h: 0.35,
    fontSize: 13, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const drivers = (data.growthDrivers || '').split('\n').filter(d => d.trim()).slice(0, 5);
  drivers.forEach((driver, idx) => {
    slide11.addShape('rect', {
      x: 0.4, y: 1.55 + (idx * 0.55), w: 4.2, h: 0.45,
      fill: { color: idx % 2 === 0 ? colors.lightBg : colors.white },
      line: { color: colors.border, width: 0.5 }
    });
    slide11.addText(`${idx + 1}. ${driver.trim()}`, {
      x: 0.5, y: 1.55 + (idx * 0.55), w: 4, h: 0.45,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'middle'
    });
  });
  
  // Short-term goals
  slide11.addShape('rect', {
    x: 4.9, y: 1.15, w: 2.4, h: 2.5,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 1 }
  });
  slide11.addText('Short-Term Goals', {
    x: 5, y: 1.2, w: 2.2, h: 0.35,
    fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  slide11.addText('(0-12 months)', {
    x: 5, y: 1.5, w: 2.2, h: 0.25,
    fontSize: 9, color: colors.textLight, fontFace: 'Arial'
  });
  
  const shortGoals = (data.shortTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 4);
  shortGoals.forEach((goal, idx) => {
    slide11.addText(`• ${goal.trim()}`, {
      x: 5, y: 1.85 + (idx * 0.45), w: 2.2, h: 0.4,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Medium-term goals
  slide11.addShape('rect', {
    x: 7.5, y: 1.15, w: 2.2, h: 2.5,
    fill: { color: colors.lightBg },
    line: { color: colors.secondary, width: 1 }
  });
  slide11.addText('Medium-Term Goals', {
    x: 7.6, y: 1.2, w: 2, h: 0.35,
    fontSize: 11, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  slide11.addText('(1-3 years)', {
    x: 7.6, y: 1.5, w: 2, h: 0.25,
    fontSize: 9, color: colors.textLight, fontFace: 'Arial'
  });
  
  const mediumGoals = (data.mediumTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 4);
  mediumGoals.forEach((goal, idx) => {
    slide11.addText(`• ${goal.trim()}`, {
      x: 7.6, y: 1.85 + (idx * 0.45), w: 2, h: 0.4,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide11, colors, 11);

  // ============================================================================
  // SLIDE 12: SYNERGIES
  // ============================================================================
  const slide12 = pptx.addSlide();
  addSlideHeader(slide12, colors, 'Potential Synergies for Acquirers', null);
  
  // Strategic Buyers column
  slide12.addShape('rect', {
    x: 0.4, y: 1.2, w: 4.5, h: 3.6,
    fill: { color: colors.white },
    line: { color: colors.primary, width: 1.5 }
  });
  
  slide12.addShape('rect', {
    x: 0.4, y: 1.2, w: 4.5, h: 0.5,
    fill: { color: colors.primary }
  });
  slide12.addText('For Strategic Buyers', {
    x: 0.5, y: 1.25, w: 4.3, h: 0.4,
    fontSize: 14, bold: true, color: colors.white, fontFace: 'Arial'
  });
  
  const strategicSynergies = (data.synergiesStrategic || '').split('\n').filter(s => s.trim()).slice(0, 6);
  strategicSynergies.forEach((synergy, idx) => {
    slide12.addText(`✓ ${synergy.trim()}`, {
      x: 0.6, y: 1.8 + (idx * 0.5), w: 4.1, h: 0.45,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Financial Investors column
  slide12.addShape('rect', {
    x: 5.1, y: 1.2, w: 4.5, h: 3.6,
    fill: { color: colors.white },
    line: { color: colors.secondary, width: 1.5 }
  });
  
  slide12.addShape('rect', {
    x: 5.1, y: 1.2, w: 4.5, h: 0.5,
    fill: { color: colors.secondary }
  });
  slide12.addText('For Financial Investors', {
    x: 5.2, y: 1.25, w: 4.3, h: 0.4,
    fontSize: 14, bold: true, color: colors.white, fontFace: 'Arial'
  });
  
  const financialSynergies = (data.synergiesFinancial || '').split('\n').filter(s => s.trim()).slice(0, 6);
  financialSynergies.forEach((synergy, idx) => {
    slide12.addText(`✓ ${synergy.trim()}`, {
      x: 5.3, y: 1.8 + (idx * 0.5), w: 4.1, h: 0.45,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });
  
  addSlideFooter(slide12, colors, 12);

  // ============================================================================
  // SLIDE 13: THANK YOU
  // ============================================================================
  const slide13 = pptx.addSlide();
  
  // Full background
  slide13.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.darkBg }
  });
  
  // Decorative elements
  slide13.addShape('rect', {
    x: 0, y: 2.4, w: 10, h: 0.02,
    fill: { color: colors.secondary }
  });
  
  slide13.addText('Thank You', {
    x: 0, y: 1.8, w: '100%', h: 0.8,
    fontSize: 48, bold: true, color: colors.white, fontFace: 'Arial', align: 'center'
  });
  
  slide13.addText(`For further information, please contact:\n${data.advisor || 'RMB Securities'}`, {
    x: 0, y: 2.7, w: '100%', h: 0.8,
    fontSize: 16, color: colors.white, fontFace: 'Arial', align: 'center', lineSpacingMultiple: 1.5
  });
  
  slide13.addText('Strictly Private and Confidential', {
    x: 0, y: 4.5, w: '100%', h: 0.3,
    fontSize: 10, italic: true, color: colors.white, fontFace: 'Arial', align: 'center', transparency: 40
  });

  return pptx;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function addSlideHeader(slide, colors, title, sectionNumber) {
  // White background
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.white }
  });
  
  // Section number badge (if provided)
  if (sectionNumber !== null && sectionNumber !== undefined) {
    slide.addShape('rect', {
      x: 0, y: 0, w: 0.35, h: 1,
      fill: { color: colors.secondary }
    });
    slide.addText(`${sectionNumber}`, {
      x: 0, y: 0.3, w: 0.35, h: 0.4,
      fontSize: 16, bold: true, color: colors.white, fontFace: 'Arial', align: 'center'
    });
  }
  
  // Title
  const titleX = sectionNumber !== null ? 0.5 : 0.3;
  slide.addText(title, {
    x: titleX, y: 0.15, w: 9.2, h: 0.8,
    fontSize: 22, bold: true, color: colors.primary, fontFace: 'Arial', valign: 'middle'
  });
  
  // Gold underline
  slide.addShape('rect', {
    x: titleX, y: 0.95, w: 9.2, h: 0.04,
    fill: { color: colors.accent }
  });
}

function addSlideFooter(slide, colors, pageNumber) {
  // Footer line
  slide.addShape('rect', {
    x: 0, y: 5.1, w: '100%', h: 0.02,
    fill: { color: colors.primary }
  });
  
  // Confidential text
  slide.addText('Strictly Private & Confidential', {
    x: 0.3, y: 5.15, w: 3, h: 0.25,
    fontSize: 8, italic: true, color: colors.textLight, fontFace: 'Arial'
  });
  
  // Page number
  slide.addText(`${pageNumber}`, {
    x: 9.2, y: 5.15, w: 0.5, h: 0.25,
    fontSize: 10, color: colors.primary, fontFace: 'Arial', align: 'right'
  });
}

// ============================================================================
// API ENDPOINTS
// ============================================================================

// Health check
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'healthy', 
    timestamp: new Date().toISOString(),
    version: '2.0.0',
    pptxEnabled: true,
    enhancedDesign: true
  });
});

// Generate enhanced PPTX
app.post('/api/generate-pptx', async (req, res) => {
  try {
    const { data, theme = 'modern-tech' } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    console.log('Generating Enhanced PPTX for:', data.projectCodename || 'Unknown Project');
    console.log('Theme:', theme);

    const pptx = await generateProfessionalPPTX(data, theme);
    
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
      slideCount: 13
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
const SYSTEM_PROMPT = `You are an expert Investment Banking Analyst specializing in creating professional Information Memorandums (IMs) for M&A transactions. You work for RMB (an investment banking firm) and help automate the creation of management presentations for potential acquirers.

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
      model: 'claude-sonnet-4-20250514'
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
  console.log('🚀 IM Creator API Server - ENHANCED v2.0');
  console.log('='.repeat(60));
  console.log(`📍 Port: ${PORT}`);
  console.log(`🔗 Health: http://localhost:${PORT}/api/health`);
  console.log(`🔑 API Key: ${process.env.ANTHROPIC_API_KEY ? 'Configured ✅' : 'NOT SET ❌'}`);
  console.log(`📊 PPTX Generation: Enhanced Deloitte-Style ✅`);
  console.log(`🎨 Themes: modern-tech, conservative, minimalist`);
  console.log('='.repeat(60));
});
