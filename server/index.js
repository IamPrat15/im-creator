// server/index.js
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

// Create temp directory for generated files
const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) {
  fs.mkdirSync(tempDir, { recursive: true });
}

// Initialize Anthropic client
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// Color themes for presentations
const THEMES = {
  'modern-tech': {
    primary: '7C1034',
    secondary: '2196F3',
    accent: 'F5F5F5',
    text: '333333',
    lightBg: 'FDF2F4',
    darkBg: '7C1034'
  },
  'conservative': {
    primary: '1a237e',
    secondary: 'c9a227',
    accent: 'F5F5F5',
    text: '333333',
    lightBg: 'E8EAF6',
    darkBg: '1a237e'
  },
  'minimalist': {
    primary: '212121',
    secondary: '757575',
    accent: 'FFFFFF',
    text: '212121',
    lightBg: 'FAFAFA',
    darkBg: '212121'
  }
};

// Helper function to create a slide with consistent styling
function createStyledSlide(pptx, theme, title, options = {}) {
  const slide = pptx.addSlide();
  const colors = THEMES[theme] || THEMES['modern-tech'];
  
  // Add header bar
  slide.addShape('rect', {
    x: 0, y: 0, w: '100%', h: 0.8,
    fill: { color: colors.primary }
  });
  
  // Add title
  slide.addText(title, {
    x: 0.5, y: 0.15, w: '90%', h: 0.5,
    fontSize: 24, bold: true, color: 'FFFFFF',
    fontFace: 'Arial'
  });
  
  // Add footer
  slide.addShape('rect', {
    x: 0, y: 5.2, w: '100%', h: 0.3,
    fill: { color: colors.lightBg }
  });
  
  if (options.confidential) {
    slide.addText('STRICTLY PRIVATE AND CONFIDENTIAL', {
      x: 0.5, y: 5.25, w: '50%', h: 0.2,
      fontSize: 8, color: colors.text, fontFace: 'Arial'
    });
  }
  
  return { slide, colors };
}

// Generate PowerPoint presentation
async function generatePowerPoint(data, content, theme = 'modern-tech') {
  const pptx = new PptxGenJS();
  const colors = THEMES[theme] || THEMES['modern-tech'];
  
  // Set presentation properties
  pptx.author = data.advisor || 'RMB Securities';
  pptx.title = `${data.projectCodename} - Management Presentation`;
  pptx.subject = 'Confidential Information Memorandum';
  pptx.company = data.advisor || 'RMB Securities';
  
  // Define master slide layout
  pptx.defineSlideMaster({
    title: 'MASTER_SLIDE',
    background: { color: 'FFFFFF' },
    objects: [
      { rect: { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: colors.primary } } },
      { rect: { x: 0, y: 5.2, w: '100%', h: 0.3, fill: { color: colors.lightBg } } }
    ]
  });

  // ==================== SLIDE 1: COVER ====================
  const slide1 = pptx.addSlide();
  slide1.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.primary }
  });
  
  // Decorative element
  slide1.addShape('rect', {
    x: 0, y: 2.2, w: '100%', h: 0.1,
    fill: { color: colors.secondary }
  });
  
  slide1.addText(data.projectCodename || 'Project Phoenix', {
    x: 0.5, y: 1.5, w: '90%', h: 0.8,
    fontSize: 44, bold: true, color: 'FFFFFF',
    fontFace: 'Arial'
  });
  
  slide1.addText('MANAGEMENT PRESENTATION', {
    x: 0.5, y: 2.4, w: '90%', h: 0.5,
    fontSize: 18, color: 'FFFFFF', fontFace: 'Arial'
  });
  
  slide1.addText(data.presentationDate || new Date().toLocaleDateString(), {
    x: 0.5, y: 3.2, w: '90%', h: 0.4,
    fontSize: 14, color: 'FFFFFF', fontFace: 'Arial'
  });
  
  slide1.addText(`Prepared by: ${data.advisor || 'RMB Securities'}`, {
    x: 0.5, y: 4.5, w: '90%', h: 0.3,
    fontSize: 12, color: 'FFFFFF', fontFace: 'Arial'
  });
  
  slide1.addText('STRICTLY PRIVATE AND CONFIDENTIAL', {
    x: 0.5, y: 5.0, w: '90%', h: 0.3,
    fontSize: 10, bold: true, color: 'FFFFFF', fontFace: 'Arial'
  });

  // ==================== SLIDE 2: DISCLAIMER ====================
  const slide2 = pptx.addSlide();
  slide2.addShape('rect', {
    x: 0, y: 0, w: '100%', h: 0.8,
    fill: { color: colors.primary }
  });
  slide2.addText('Disclaimer', {
    x: 0.5, y: 0.15, w: '90%', h: 0.5,
    fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Arial'
  });
  
  const disclaimerText = `This presentation has been prepared by ${data.advisor || 'RMB Securities'} exclusively for the benefit of the party to whom it is directly addressed and delivered.

This presentation is confidential and may not be reproduced, redistributed or passed on to any other person or published, in whole or in part, for any purpose without the prior written consent of ${data.advisor || 'RMB Securities'}.

The information contained in this presentation has been prepared based on information provided by the management of ${data.companyName || 'the Company'} and has not been independently verified.

This presentation does not constitute an offer or invitation to purchase or subscribe for any securities, and neither this presentation nor anything contained herein shall form the basis of any contract or commitment whatsoever.`;

  slide2.addText(disclaimerText, {
    x: 0.5, y: 1.2, w: 9, h: 3.5,
    fontSize: 11, color: colors.text, fontFace: 'Arial',
    valign: 'top', paraSpaceAfter: 12
  });

  // ==================== SLIDE 3: EXECUTIVE SUMMARY ====================
  const { slide: slide3 } = createStyledSlide(pptx, theme, 'Executive Summary', { confidential: true });
  
  // Company snapshot
  slide3.addText('Company Snapshot', {
    x: 0.5, y: 1.0, w: 4, h: 0.3,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  slide3.addText(data.companyDescription || 'Leading technology services company', {
    x: 0.5, y: 1.4, w: 4.2, h: 1.0,
    fontSize: 11, color: colors.text, fontFace: 'Arial', valign: 'top'
  });
  
  // Key metrics box
  slide3.addShape('rect', {
    x: 5, y: 1.0, w: 4.5, h: 1.8,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 1 }
  });
  
  slide3.addText('Key Metrics', {
    x: 5.2, y: 1.1, w: 4, h: 0.3,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const metrics = [
    `Revenue FY25: ₹${data.revenueFY25 || 'XX'} Cr`,
    `Employees: ${data.employeeCountFT || 'XXX'}+`,
    `EBITDA Margin: ${data.ebitdaMarginFY25 || 'XX'}%`,
    `Founded: ${data.foundedYear || 'XXXX'}`
  ];
  
  slide3.addText(metrics.join('\n'), {
    x: 5.2, y: 1.5, w: 4, h: 1.2,
    fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top',
    paraSpaceAfter: 6
  });
  
  // Investment highlights
  slide3.addText('Investment Highlights', {
    x: 0.5, y: 2.6, w: 9, h: 0.3,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const highlights = (data.investmentHighlights || '').split('\n').filter(h => h.trim()).slice(0, 6);
  highlights.forEach((highlight, idx) => {
    slide3.addText(`• ${highlight.trim()}`, {
      x: 0.5, y: 3.0 + (idx * 0.35), w: 9, h: 0.35,
      fontSize: 11, color: colors.text, fontFace: 'Arial'
    });
  });

  // ==================== SLIDE 4: COMPANY OVERVIEW ====================
  const { slide: slide4 } = createStyledSlide(pptx, theme, 'Company Overview', { confidential: true });
  
  // Company info
  const companyInfo = [
    { label: 'Company Name', value: data.companyName || 'N/A' },
    { label: 'Founded', value: data.foundedYear || 'N/A' },
    { label: 'Headquarters', value: data.headquarters || 'N/A' },
    { label: 'Employees', value: `${data.employeeCountFT || 0} FTE + ${data.employeeCountOther || 0} Contractors` },
    { label: 'Primary Vertical', value: data.primaryVertical?.toUpperCase() || 'N/A' }
  ];
  
  companyInfo.forEach((info, idx) => {
    slide4.addText(info.label, {
      x: 0.5, y: 1.2 + (idx * 0.5), w: 2.5, h: 0.4,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    slide4.addText(info.value, {
      x: 3, y: 1.2 + (idx * 0.5), w: 6, h: 0.4,
      fontSize: 11, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Description box
  slide4.addShape('rect', {
    x: 0.5, y: 3.8, w: 9, h: 1.2,
    fill: { color: colors.lightBg }
  });
  
  slide4.addText('About the Company', {
    x: 0.7, y: 3.9, w: 8.5, h: 0.3,
    fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  slide4.addText(data.companyDescription || '', {
    x: 0.7, y: 4.25, w: 8.5, h: 0.7,
    fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
  });

  // ==================== SLIDE 5: FOUNDER PROFILE ====================
  const { slide: slide5 } = createStyledSlide(pptx, theme, 'Founder & Leadership', { confidential: true });
  
  // Founder info box
  slide5.addShape('rect', {
    x: 0.5, y: 1.0, w: 4.5, h: 2.5,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 1 }
  });
  
  slide5.addText(data.founderName || 'Founder Name', {
    x: 0.7, y: 1.1, w: 4, h: 0.4,
    fontSize: 16, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  slide5.addText(data.founderTitle || 'Founder & CEO', {
    x: 0.7, y: 1.5, w: 4, h: 0.3,
    fontSize: 12, color: colors.text, fontFace: 'Arial'
  });
  
  slide5.addText(`${data.founderExperience || 'XX'}+ Years Experience`, {
    x: 0.7, y: 1.9, w: 4, h: 0.3,
    fontSize: 11, color: colors.text, fontFace: 'Arial'
  });
  
  // Education
  const education = (data.founderEducation || '').split('\n').filter(e => e.trim()).slice(0, 3);
  slide5.addText('Education:', {
    x: 0.7, y: 2.3, w: 4, h: 0.25,
    fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  education.forEach((edu, idx) => {
    slide5.addText(`• ${edu.trim()}`, {
      x: 0.7, y: 2.55 + (idx * 0.25), w: 4, h: 0.25,
      fontSize: 9, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Leadership team
  slide5.addText('Leadership Team', {
    x: 5.2, y: 1.0, w: 4, h: 0.3,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const leaders = (data.leadershipTeam || '').split('\n').filter(l => l.trim()).slice(0, 6);
  leaders.forEach((leader, idx) => {
    const parts = leader.split('|').map(p => p.trim());
    slide5.addText(`• ${parts[0] || ''} - ${parts[1] || ''}`, {
      x: 5.2, y: 1.4 + (idx * 0.35), w: 4.3, h: 0.35,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });

  // ==================== SLIDE 6: SERVICE OFFERINGS ====================
  const { slide: slide6 } = createStyledSlide(pptx, theme, 'Service Offerings', { confidential: true });
  
  const services = (data.serviceLines || '').split('\n').filter(s => s.trim()).slice(0, 4);
  services.forEach((service, idx) => {
    const parts = service.split('|').map(p => p.trim());
    const xPos = idx % 2 === 0 ? 0.5 : 5;
    const yPos = idx < 2 ? 1.2 : 3.0;
    
    // Service box
    slide6.addShape('rect', {
      x: xPos, y: yPos, w: 4.3, h: 1.6,
      fill: { color: colors.lightBg },
      line: { color: colors.primary, width: 1 }
    });
    
    slide6.addText(parts[0] || `Service ${idx + 1}`, {
      x: xPos + 0.2, y: yPos + 0.1, w: 3.9, h: 0.35,
      fontSize: 12, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    slide6.addText(parts[1] ? `${parts[1]} of Revenue` : '', {
      x: xPos + 0.2, y: yPos + 0.45, w: 3.9, h: 0.25,
      fontSize: 10, bold: true, color: colors.secondary, fontFace: 'Arial'
    });
    
    slide6.addText(parts[2] || '', {
      x: xPos + 0.2, y: yPos + 0.75, w: 3.9, h: 0.7,
      fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
  });

  // ==================== SLIDE 7: CLIENT PORTFOLIO ====================
  const { slide: slide7 } = createStyledSlide(pptx, theme, 'Client Portfolio', { confidential: true });
  
  // Client metrics
  slide7.addShape('rect', {
    x: 0.5, y: 1.0, w: 2.8, h: 1.2,
    fill: { color: colors.primary }
  });
  slide7.addText('Primary Vertical', {
    x: 0.6, y: 1.1, w: 2.6, h: 0.3,
    fontSize: 10, color: 'FFFFFF', fontFace: 'Arial'
  });
  slide7.addText(`${data.primaryVertical?.toUpperCase() || 'BFSI'}\n${data.primaryVerticalPct || 'XX'}%`, {
    x: 0.6, y: 1.4, w: 2.6, h: 0.7,
    fontSize: 18, bold: true, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });
  
  slide7.addShape('rect', {
    x: 3.5, y: 1.0, w: 2.8, h: 1.2,
    fill: { color: colors.secondary }
  });
  slide7.addText('Top 10 Concentration', {
    x: 3.6, y: 1.1, w: 2.6, h: 0.3,
    fontSize: 10, color: 'FFFFFF', fontFace: 'Arial'
  });
  slide7.addText(`${data.top10Concentration || 'XX'}%`, {
    x: 3.6, y: 1.5, w: 2.6, h: 0.5,
    fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });
  
  slide7.addShape('rect', {
    x: 6.5, y: 1.0, w: 2.8, h: 1.2,
    fill: { color: colors.primary }
  });
  slide7.addText('Net Retention', {
    x: 6.6, y: 1.1, w: 2.6, h: 0.3,
    fontSize: 10, color: 'FFFFFF', fontFace: 'Arial'
  });
  slide7.addText(`${data.netRetention || 'XXX'}%`, {
    x: 6.6, y: 1.5, w: 2.6, h: 0.5,
    fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });
  
  // Top clients
  slide7.addText('Marquee Clients', {
    x: 0.5, y: 2.4, w: 9, h: 0.3,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const clients = (data.topClients || '').split('\n').filter(c => c.trim()).slice(0, 6);
  clients.forEach((client, idx) => {
    const parts = client.split('|').map(p => p.trim());
    const xPos = idx % 3 === 0 ? 0.5 : idx % 3 === 1 ? 3.5 : 6.5;
    const yPos = idx < 3 ? 2.8 : 3.9;
    
    slide7.addShape('rect', {
      x: xPos, y: yPos, w: 2.8, h: 1.0,
      fill: { color: colors.lightBg },
      line: { color: colors.primary, width: 0.5 }
    });
    
    slide7.addText(parts[0] || '', {
      x: xPos + 0.1, y: yPos + 0.1, w: 2.6, h: 0.35,
      fontSize: 10, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    slide7.addText(`${parts[1] || ''} | Since ${parts[2] || ''}`, {
      x: xPos + 0.1, y: yPos + 0.5, w: 2.6, h: 0.4,
      fontSize: 8, color: colors.text, fontFace: 'Arial'
    });
  });

  // ==================== SLIDE 8: FINANCIAL OVERVIEW ====================
  const { slide: slide8 } = createStyledSlide(pptx, theme, 'Financial Overview', { confidential: true });
  
  slide8.addText(`All figures in INR Crores`, {
    x: 0.5, y: 0.95, w: 9, h: 0.2,
    fontSize: 9, italic: true, color: colors.text, fontFace: 'Arial'
  });
  
  // Revenue data
  const revenueData = [
    { year: 'FY24', value: data.revenueFY24, type: 'actual' },
    { year: 'FY25', value: data.revenueFY25, type: 'actual' },
    { year: 'FY26P', value: data.revenueFY26P, type: 'projected' },
    { year: 'FY27P', value: data.revenueFY27P, type: 'projected' },
    { year: 'FY28P', value: data.revenueFY28P, type: 'projected' }
  ].filter(d => d.value);
  
  // Simple bar chart representation
  const maxRevenue = Math.max(...revenueData.map(d => parseFloat(d.value) || 0));
  
  revenueData.forEach((rev, idx) => {
    const barHeight = (parseFloat(rev.value) / maxRevenue) * 2;
    const xPos = 1 + (idx * 1.6);
    
    // Bar
    slide8.addShape('rect', {
      x: xPos, y: 3.5 - barHeight, w: 1.2, h: barHeight,
      fill: { color: rev.type === 'actual' ? colors.primary : colors.secondary }
    });
    
    // Value
    slide8.addText(`₹${rev.value}`, {
      x: xPos - 0.1, y: 3.5 - barHeight - 0.3, w: 1.4, h: 0.3,
      fontSize: 10, bold: true, color: colors.text, fontFace: 'Arial', align: 'center'
    });
    
    // Year label
    slide8.addText(rev.year, {
      x: xPos - 0.1, y: 3.6, w: 1.4, h: 0.3,
      fontSize: 10, color: colors.text, fontFace: 'Arial', align: 'center'
    });
  });
  
  // EBITDA margin
  slide8.addShape('rect', {
    x: 0.5, y: 4.2, w: 4, h: 0.8,
    fill: { color: colors.lightBg }
  });
  slide8.addText(`EBITDA Margin FY25: ${data.ebitdaMarginFY25 || 'XX'}%`, {
    x: 0.7, y: 4.4, w: 3.6, h: 0.4,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  // Calculate CAGR if possible
  if (data.revenueFY24 && data.revenueFY26P) {
    const cagr = ((Math.pow(data.revenueFY26P / data.revenueFY24, 1/2) - 1) * 100).toFixed(0);
    slide8.addShape('rect', {
      x: 5, y: 4.2, w: 4, h: 0.8,
      fill: { color: colors.lightBg }
    });
    slide8.addText(`Revenue CAGR (FY24-26): ${cagr}%`, {
      x: 5.2, y: 4.4, w: 3.6, h: 0.4,
      fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
    });
  }

  // ==================== SLIDE 9: CASE STUDY ====================
  if (data.cs1Client) {
    const { slide: slide9 } = createStyledSlide(pptx, theme, `Case Study: ${data.cs1Client}`, { confidential: true });
    
    // Challenge
    slide9.addShape('rect', {
      x: 0.5, y: 1.0, w: 4.3, h: 1.8,
      fill: { color: '#FEE2E2' }
    });
    slide9.addText('Challenge', {
      x: 0.7, y: 1.1, w: 3.9, h: 0.3,
      fontSize: 12, bold: true, color: '#991B1B', fontFace: 'Arial'
    });
    slide9.addText(data.cs1Challenge || '', {
      x: 0.7, y: 1.45, w: 3.9, h: 1.2,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
    
    // Solution
    slide9.addShape('rect', {
      x: 5, y: 1.0, w: 4.3, h: 1.8,
      fill: { color: '#DBEAFE' }
    });
    slide9.addText('Solution', {
      x: 5.2, y: 1.1, w: 3.9, h: 0.3,
      fontSize: 12, bold: true, color: '#1E40AF', fontFace: 'Arial'
    });
    slide9.addText(data.cs1Solution || '', {
      x: 5.2, y: 1.45, w: 3.9, h: 1.2,
      fontSize: 10, color: colors.text, fontFace: 'Arial', valign: 'top'
    });
    
    // Results
    slide9.addShape('rect', {
      x: 0.5, y: 3.0, w: 8.8, h: 1.8,
      fill: { color: '#ECFDF5' }
    });
    slide9.addText('Results', {
      x: 0.7, y: 3.1, w: 8.4, h: 0.3,
      fontSize: 12, bold: true, color: '#047857', fontFace: 'Arial'
    });
    
    const results = (data.cs1Results || '').split('\n').filter(r => r.trim()).slice(0, 4);
    results.forEach((result, idx) => {
      const xPos = idx % 2 === 0 ? 0.7 : 4.8;
      const yPos = idx < 2 ? 3.5 : 4.1;
      slide9.addText(`✓ ${result.trim()}`, {
        x: xPos, y: yPos, w: 4, h: 0.5,
        fontSize: 11, color: '#047857', fontFace: 'Arial'
      });
    });
  }

  // ==================== SLIDE 10: COMPETITIVE ADVANTAGES ====================
  const { slide: slide10 } = createStyledSlide(pptx, theme, 'Competitive Advantages', { confidential: true });
  
  const advantages = (data.competitiveAdvantages || '').split('\n').filter(a => a.trim()).slice(0, 6);
  advantages.forEach((advantage, idx) => {
    const xPos = idx % 2 === 0 ? 0.5 : 5;
    const yPos = 1.0 + (Math.floor(idx / 2) * 1.3);
    
    slide10.addShape('rect', {
      x: xPos, y: yPos, w: 4.3, h: 1.1,
      fill: { color: colors.lightBg },
      line: { color: colors.primary, width: 1 }
    });
    
    const parts = advantage.split('|').map(p => p.trim());
    slide10.addText(parts[0] || advantage, {
      x: xPos + 0.2, y: yPos + 0.1, w: 3.9, h: 0.35,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    
    if (parts[1]) {
      slide10.addText(parts[1], {
        x: xPos + 0.2, y: yPos + 0.5, w: 3.9, h: 0.5,
        fontSize: 9, color: colors.text, fontFace: 'Arial', valign: 'top'
      });
    }
  });

  // ==================== SLIDE 11: GROWTH STRATEGY ====================
  const { slide: slide11 } = createStyledSlide(pptx, theme, 'Growth Strategy', { confidential: true });
  
  const growthDrivers = (data.growthDrivers || '').split('\n').filter(g => g.trim()).slice(0, 5);
  
  slide11.addText('Key Growth Drivers', {
    x: 0.5, y: 1.0, w: 9, h: 0.3,
    fontSize: 14, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  growthDrivers.forEach((driver, idx) => {
    slide11.addText(`${idx + 1}. ${driver.trim()}`, {
      x: 0.5, y: 1.4 + (idx * 0.4), w: 9, h: 0.4,
      fontSize: 11, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Short term goals
  const shortTerm = (data.shortTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 3);
  if (shortTerm.length > 0) {
    slide11.addShape('rect', {
      x: 0.5, y: 3.5, w: 4.3, h: 1.5,
      fill: { color: colors.lightBg }
    });
    slide11.addText('Short-Term (0-12 months)', {
      x: 0.7, y: 3.6, w: 3.9, h: 0.3,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    shortTerm.forEach((goal, idx) => {
      slide11.addText(`• ${goal.trim()}`, {
        x: 0.7, y: 3.95 + (idx * 0.3), w: 3.9, h: 0.3,
        fontSize: 9, color: colors.text, fontFace: 'Arial'
      });
    });
  }
  
  // Medium term goals
  const mediumTerm = (data.mediumTermGoals || '').split('\n').filter(g => g.trim()).slice(0, 3);
  if (mediumTerm.length > 0) {
    slide11.addShape('rect', {
      x: 5, y: 3.5, w: 4.3, h: 1.5,
      fill: { color: colors.lightBg }
    });
    slide11.addText('Medium-Term (1-3 years)', {
      x: 5.2, y: 3.6, w: 3.9, h: 0.3,
      fontSize: 11, bold: true, color: colors.primary, fontFace: 'Arial'
    });
    mediumTerm.forEach((goal, idx) => {
      slide11.addText(`• ${goal.trim()}`, {
        x: 5.2, y: 3.95 + (idx * 0.3), w: 3.9, h: 0.3,
        fontSize: 9, color: colors.text, fontFace: 'Arial'
      });
    });
  }

  // ==================== SLIDE 12: SYNERGIES ====================
  const { slide: slide12 } = createStyledSlide(pptx, theme, 'Potential Synergies', { confidential: true });
  
  // Strategic buyer synergies
  slide12.addShape('rect', {
    x: 0.5, y: 1.0, w: 4.3, h: 3.8,
    fill: { color: colors.lightBg },
    line: { color: colors.primary, width: 1 }
  });
  slide12.addText('For Strategic Buyers', {
    x: 0.7, y: 1.1, w: 3.9, h: 0.35,
    fontSize: 13, bold: true, color: colors.primary, fontFace: 'Arial'
  });
  
  const strategicSynergies = (data.synergiesStrategic || '').split('\n').filter(s => s.trim()).slice(0, 5);
  strategicSynergies.forEach((synergy, idx) => {
    slide12.addText(`• ${synergy.trim()}`, {
      x: 0.7, y: 1.55 + (idx * 0.5), w: 3.9, h: 0.5,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });
  
  // Financial investor synergies
  slide12.addShape('rect', {
    x: 5, y: 1.0, w: 4.3, h: 3.8,
    fill: { color: colors.lightBg },
    line: { color: colors.secondary, width: 1 }
  });
  slide12.addText('For Financial Investors', {
    x: 5.2, y: 1.1, w: 3.9, h: 0.35,
    fontSize: 13, bold: true, color: colors.secondary, fontFace: 'Arial'
  });
  
  const financialSynergies = (data.synergiesFinancial || '').split('\n').filter(s => s.trim()).slice(0, 5);
  financialSynergies.forEach((synergy, idx) => {
    slide12.addText(`• ${synergy.trim()}`, {
      x: 5.2, y: 1.55 + (idx * 0.5), w: 3.9, h: 0.5,
      fontSize: 10, color: colors.text, fontFace: 'Arial'
    });
  });

  // ==================== SLIDE 13: THANK YOU ====================
  const slide13 = pptx.addSlide();
  slide13.addShape('rect', {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: colors.primary }
  });
  
  slide13.addText('Thank You', {
    x: 0.5, y: 2, w: '90%', h: 0.8,
    fontSize: 44, bold: true, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });
  
  slide13.addText(`For further information, please contact:\n${data.advisor || 'RMB Securities'}`, {
    x: 0.5, y: 3.2, w: '90%', h: 0.8,
    fontSize: 16, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });
  
  slide13.addText('STRICTLY PRIVATE AND CONFIDENTIAL', {
    x: 0.5, y: 4.5, w: '90%', h: 0.3,
    fontSize: 10, color: 'FFFFFF', fontFace: 'Arial', align: 'center'
  });

  return pptx;
}

// System prompt for IM generation
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

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'healthy', 
    timestamp: new Date().toISOString(),
    version: '1.0.0',
    pptxEnabled: true
  });
});

// Generate IM endpoint (JSON content)
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

// Generate PowerPoint endpoint
app.post('/api/generate-pptx', async (req, res) => {
  try {
    const { data, theme = 'modern-tech' } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    console.log('Generating PPTX for:', data.projectCodename || 'Unknown Project');

    // Generate the PowerPoint
    const pptx = await generatePowerPoint(data, null, theme);
    
    // Generate unique filename
    const filename = `${data.projectCodename || 'IM'}_${Date.now()}.pptx`;
    const filepath = path.join(tempDir, filename);
    
    // Write to file
    await pptx.writeFile(filepath);
    
    // Read file and send as base64
    const fileBuffer = fs.readFileSync(filepath);
    const base64 = fileBuffer.toString('base64');
    
    // Clean up temp file
    fs.unlinkSync(filepath);
    
    res.json({
      success: true,
      filename: filename,
      fileData: base64,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      generatedAt: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error generating PPTX:', error);
    res.status(500).json({ 
      error: 'Failed to generate PowerPoint', 
      details: error.message 
    });
  }
});

// Validate data endpoint
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

// In-memory draft storage
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
  console.log('='.repeat(50));
  console.log('🚀 IM Creator API Server');
  console.log('='.repeat(50));
  console.log(`📍 Port: ${PORT}`);
  console.log(`🔗 Health: http://localhost:${PORT}/api/health`);
  console.log(`🔑 API Key: ${process.env.ANTHROPIC_API_KEY ? 'Configured ✅' : 'NOT SET ❌'}`);
  console.log(`📊 PPTX Generation: Enabled ✅`);
  console.log('='.repeat(50));
});
