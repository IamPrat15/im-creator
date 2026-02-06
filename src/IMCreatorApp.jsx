import React, { useState, useEffect, useCallback, useRef } from 'react';
import { generateIM, generatePPTX, saveDraft, checkHealth, downloadBase64File, getTemplates, getUsageStats, exportUsageCSV, resetUsageStats, exportQAWord, downloadBlob } from './api';

// ============================================================================
// IMCreatorApp v8.1.0 - Complete Production Build
// ============================================================================
// Features:
// - Document Types (Management Presentation, CIM, Teaser)
// - Enhanced Buyer Types (Strategic, Financial, International)
// - Industry-Specific Content (6 industries)
// - 50 Professional Templates
// - Unlimited Dynamic Case Studies
// - Auto-Logout (15 min inactivity)
// - Word/PDF/JSON Export
// - Conditional Mandatory Fields
// ============================================================================

// Auto-logout configuration
const AUTO_LOGOUT_TIMEOUT = 15 * 60 * 1000; // 15 minutes
const AUTO_LOGOUT_WARNING = 14 * 60 * 1000; // Warning at 14 minutes
//const API_BASE = import.meta.env.VITE_API_URL || 'http://localhost:3001';

// ============================================================================
// QUESTIONNAIRE CONFIGURATION
// ============================================================================
const defaultQuestionnaire = {
  phases: [
    {
      id: 'project-setup',
      name: 'Project Setup',
      icon: '📋',
      description: 'Basic project information',
      questions: [
        { id: 'projectCodename', type: 'text', label: 'Project Codename', placeholder: 'e.g., Project Phoenix', required: true, helpText: 'Confidential identifier for the deal', order: 1 },
        { id: 'companyName', type: 'text', label: 'Company Legal Name', required: true, order: 2 },
        { id: 'documentType', type: 'select', label: 'Document Type', required: true, options: [
          { value: 'management-presentation', label: 'Management Presentation (13-20 slides)' },
          { value: 'cim', label: 'Confidential Information Memorandum (25-40 slides)' },
          { value: 'teaser', label: 'Teaser Document (5-8 slides)' }
        ], helpText: 'Document type affects slide count and detail level', order: 3 },
        { id: 'advisor', type: 'text', label: 'Sell-Side Advisor', placeholder: 'e.g., Goldman Sachs, Morgan Stanley', order: 4 },
        { id: 'presentationDate', type: 'date', label: 'Presentation Date', required: true, order: 5 },
        { id: 'targetBuyerType', type: 'multiselect', label: 'Target Buyer Type', required: true, options: [
          { value: 'strategic', label: 'Strategic Buyer - Focus on synergies & market expansion' },
          { value: 'financial', label: 'Financial Investor - Focus on returns & growth potential' },
          { value: 'international', label: 'International Acquirer - Focus on market entry & local expertise' }
        ], helpText: 'Content will be tailored for selected buyers', order: 6 }
      ]
    },
    {
      id: 'company-fundamentals',
      name: 'Company Overview',
      icon: '🏢',
      description: 'Basic company information',
      questions: [
        { id: 'foundedYear', type: 'number', label: 'Founded Year', required: true, validation: { min: 1900, max: 2030 }, order: 1 },
        { id: 'headquarters', type: 'text', label: 'Headquarters', placeholder: 'Mumbai, India', required: true, order: 2 },
        { id: 'companyDescription', type: 'textarea', label: 'Company Description', required: true, helpText: 'Brief description for executive summary (2-3 sentences)', order: 3 },
        { id: 'employeeCountFT', type: 'number', label: 'Full-Time Employees', required: true, order: 4 },
        { id: 'employeeCountOther', type: 'number', label: 'Contractors/Trainees', order: 5 },
        { id: 'investmentHighlights', type: 'textarea', label: 'Investment Highlights', placeholder: 'One highlight per line\ne.g., AWS Advanced Tier Partner\nStrong BFSI client base', helpText: 'Recommend 5-7 compelling highlights', order: 6 }
      ]
    },
    {
      id: 'founder-leadership',
      name: 'Leadership',
      icon: '👥',
      description: 'Founder & management team',
      questions: [
        { id: 'founderName', type: 'text', label: 'Founder Name', required: true, order: 1 },
        { id: 'founderTitle', type: 'text', label: 'Founder Title', placeholder: 'Founder & CEO', required: true, order: 2 },
        { id: 'founderExperience', type: 'number', label: 'Years of Experience', required: true, order: 3 },
        { id: 'founderEducation', type: 'textarea', label: 'Education', placeholder: 'MBA - JBIMS\nB.E. - VJTI', helpText: 'Format: Degree - Institution (one per line)', order: 4 },
        { id: 'previousCompanies', type: 'textarea', label: 'Previous Companies', placeholder: 'Infosys | Senior Architect | 5 years', helpText: 'Enter one company per line. Format: Company | Role | Duration', order: 5 },
        { id: 'leadershipTeam', type: 'textarea', label: 'Leadership Team', placeholder: 'Priya Sharma | CTO | Technology', helpText: 'Enter one leader per line. Format: Name | Title | Department', order: 6 }
      ]
    },
    {
      id: 'services-products',
      name: 'Services & Products',
      icon: '⚙️',
      description: 'Offerings & capabilities',
      questions: [
        { id: 'serviceLines', type: 'textarea', label: 'Service Lines', placeholder: 'Cloud & Automation | 39% | AWS migration, DevOps', required: true, helpText: 'Enter one service per line. Format: Name | Revenue % | Description', order: 1 },
        { id: 'products', type: 'textarea', label: 'Proprietary Products', placeholder: 'NovaCloud Platform | Cloud automation | 500 deployments', helpText: 'Enter one product per line. Format: Name | Description | Key metric', order: 2 },
        { id: 'techPartnerships', type: 'textarea', label: 'Technology Partnerships', placeholder: 'AWS Advanced Tier Partner\nMicrosoft Gold Partner', helpText: 'Enter one partnership per line', order: 3 },
        { id: 'certifications', type: 'textarea', label: 'Certifications & Awards', placeholder: 'AWS Financial Services Competency\nBest BFSI Partner 2024', order: 4 }
      ]
    },
    {
      id: 'clients-verticals',
      name: 'Clients & Verticals',
      icon: '💼',
      description: 'Client portfolio',
      questions: [
        { id: 'primaryVertical', type: 'select', label: 'Primary Vertical', required: true, options: [
          { value: 'bfsi', label: 'BFSI (Banking, Financial Services & Insurance)' },
          { value: 'healthcare', label: 'Healthcare & Life Sciences' },
          { value: 'retail', label: 'Retail & Consumer' },
          { value: 'manufacturing', label: 'Manufacturing & Industrial' },
          { value: 'technology', label: 'Technology & Software' },
          { value: 'media', label: 'Media & Entertainment' }
        ], helpText: 'Industry-specific benchmarks will be included', order: 1 },
        { id: 'primaryVerticalPct', type: 'number', label: 'Primary Vertical Revenue %', required: true, order: 2 },
        { id: 'otherVerticals', type: 'textarea', label: 'Other Verticals', placeholder: 'FinTech | 14%\nMedia | 11%', helpText: 'Enter one vertical per line. Format: Vertical Name | Revenue %', order: 3 },
        { id: 'topClients', type: 'textarea', label: 'Top Clients', placeholder: 'Axis Bank | BFSI | 2015\nHDFC Bank | BFSI | 2018', required: true, helpText: 'Enter one client per line. Format: Client Name | Vertical | Year Started', order: 4 },
        { id: 'top10Concentration', type: 'number', label: 'Top 10 Client Concentration %', required: true, order: 5 },
        { id: 'netRetention', type: 'number', label: 'Net Revenue Retention %', helpText: 'NRR indicates revenue expansion from existing clients', order: 6 }
      ]
    },
    {
      id: 'financials',
      name: 'Financials',
      icon: '📈',
      description: 'Financial performance',
      questions: [
        { id: 'currency', type: 'select', label: 'Currency', options: [{ value: 'INR', label: 'INR (₹ Cr)' }, { value: 'USD', label: 'USD ($ Mn)' }], defaultValue: 'INR', order: 1 },
        { id: 'revenueFY24', type: 'number', label: 'Revenue FY24 (Cr/Mn)', required: true, order: 2 },
        { id: 'revenueFY25', type: 'number', label: 'Revenue FY25 (Cr/Mn)', required: true, order: 3 },
        { id: 'revenueFY26P', type: 'number', label: 'Revenue FY26P (Cr/Mn)', required: true, helpText: 'P = Projected', order: 4 },
        { id: 'revenueFY27P', type: 'number', label: 'Revenue FY27P (Cr/Mn)', helpText: 'Leave blank if not projected', order: 5 },
        { id: 'revenueFY28P', type: 'number', label: 'Revenue FY28P (Cr/Mn)', helpText: 'Leave blank if not projected', order: 6 },
        { id: 'ebitdaMarginFY25', type: 'number', label: 'EBITDA Margin FY25 %', order: 7 },
        { id: 'revenueByService', type: 'textarea', label: 'Revenue by Service', placeholder: 'Cloud Services | 39%\nManaged Services | 31%', helpText: 'Enter one service per line. Format: Service Name | Revenue %', order: 8 },
        { id: 'grossMargin', type: 'number', label: 'Gross Margin %', order: 9 },
        { id: 'netProfitMargin', type: 'number', label: 'Net Profit Margin %', order: 10 }
      ]
    },
    {
      id: 'case-studies',
      name: 'Case Studies',
      icon: '📚',
      description: 'Client success stories',
      isDynamic: true,
      questions: []
    },
    {
      id: 'growth-strategy',
      name: 'Growth Strategy',
      icon: '🎯',
      description: 'Future plans & competitive position',
      questions: [
        { id: 'growthDrivers', type: 'textarea', label: 'Key Growth Drivers', required: true, placeholder: 'AI adoption in enterprise\nCloud migration demand\nExpanding BFSI relationships', order: 1 },
        { id: 'competitiveAdvantages', type: 'textarea', label: 'Competitive Advantages', required: true, helpText: 'Enter one advantage per line. Format: Advantage Title | Supporting Detail (min 5 recommended)', placeholder: 'Deep AWS expertise | Only 8 companies in India with certification\nStrong BFSI relationships | 10+ year partnerships', order: 2 },
        { id: 'shortTermGoals', type: 'textarea', label: 'Short-Term Goals (0-12 months)', placeholder: 'Launch AI Practice\nExpand Bangalore team', order: 3 },
        { id: 'mediumTermGoals', type: 'textarea', label: 'Medium-Term Goals (1-3 years)', placeholder: 'Enter international markets\nAchieve ₹500 Cr revenue', order: 4 },
        { id: 'synergiesStrategic', type: 'textarea', label: 'Synergies for Strategic Buyers', placeholder: 'Access to BFSI client base\nAWS competencies enhancement\nIndian market expansion', order: 5 },
        { id: 'synergiesFinancial', type: 'textarea', label: 'Synergies for Financial Investors', placeholder: 'Strong EBITDA margins\nCapital-light model\nHigh revenue visibility', order: 6 },
        { id: 'marketSize', type: 'text', label: 'Total Addressable Market (TAM)', placeholder: 'e.g., $50B globally', order: 7 },
        { id: 'marketGrowthRate', type: 'text', label: 'Market Growth Rate', placeholder: 'e.g., 15% CAGR', order: 8 },
        { id: 'competitorLandscape', type: 'textarea', label: 'Key Competitors', placeholder: 'TCS | Global reach | Premium pricing\nInfosys | Strong brand | Less agile', helpText: 'Enter one competitor per line. Format: Name | Strength | Weakness. Required for Market Position variant.', order: 9 }
      ]
    },
    {
      id: 'risk-factors',
      name: 'Risk Factors',
      icon: '⚠️',
      description: 'Key risks (Required for CIM)',
      conditionalShow: { field: 'documentType', equals: 'cim' },
      questions: [
        { id: 'businessRisks', type: 'textarea', label: 'Business Risks', placeholder: 'Client concentration risk\nTechnology obsolescence', helpText: 'Key business risks (one per line)', order: 1 },
        { id: 'marketRisks', type: 'textarea', label: 'Market Risks', placeholder: 'Economic downturn\nRegulatory changes', order: 2 },
        { id: 'operationalRisks', type: 'textarea', label: 'Operational Risks', placeholder: 'Key person dependency\nCybersecurity threats', order: 3 },
        { id: 'mitigationStrategies', type: 'textarea', label: 'Mitigation Strategies', placeholder: 'Diversification plan\nInsurance coverage\nBusiness continuity planning', order: 4 }
      ]
    },
    {
      id: 'review-generate',
      name: 'Review & Generate',
      icon: '🏆',
      description: 'Final review and output options',
      questions: [
        { id: 'generateVariants', type: 'multiselect', label: 'Content Variants', options: [
          { value: 'financial', label: 'Financial Focus - Emphasize metrics, margins, growth rates' },
          { value: 'tech', label: 'Technology Focus - Highlight technical capabilities, IP, platforms' },
          { value: 'market', label: 'Market Position - Focus on competitive landscape, market share' },
          { value: 'synergy', label: 'Synergy Focus - Emphasize acquisition benefits for buyers' }
        ], helpText: 'Select which content variants to include in the presentation', order: 1 },
        { id: 'templateStyle', type: 'select', label: 'Presentation Template', required: true, 
          dynamicOptions: 'templates',
          options: [
            { value: 'modern-blue', label: 'Modern Blue' },
            { value: 'corporate-navy', label: 'Corporate Navy' },
            { value: 'elegant-burgundy', label: 'Elegant Burgundy' },
            { value: 'tech-gradient', label: 'Tech Gradient' },
            { value: 'minimalist-mono', label: 'Minimalist Mono' }
          ], 
          defaultValue: 'modern-blue', 
          helpText: '50 professional templates available', order: 2 },
        { id: 'includeAppendix', type: 'multiselect', label: 'Include in Appendix', options: [
          { value: 'team-bios', label: 'Detailed Team Bios' },
          { value: 'client-list', label: 'Full Client List' },
          { value: 'financial-detail', label: 'Detailed Financial Statements (P&L, Balance Sheet)' },
          { value: 'case-studies-extra', label: 'All Case Studies (beyond first two)' }
        ], order: 3 },
        { id: 'exportFormat', type: 'multiselect', label: 'Export Formats', options: [
          { value: 'pptx', label: 'PowerPoint (.pptx) - Editable' },
          { value: 'pdf', label: 'PDF - For distribution' },
          { value: 'json', label: 'JSON Data - For integration' },
          { value: 'docx', label: 'Word Q&A Document (.docx)' }
        ], defaultValue: ['pptx'], order: 4 }
      ]
    }
  ]
};

// ACC Brand Colors
const THEME = {
  primary: '#7C1034',
  primaryDark: '#5A0C26',
  primaryLight: '#9A1842',
  secondary: '#2D3748',
  accent: '#48BB78',
  accentYellow: '#ECC94B',
  accentRed: '#E53E3E',
  accentBlue: '#4299E1',
  background: '#F7FAFC',
  surface: '#FFFFFF',
  text: '#1A202C',
  textLight: '#718096',
  border: '#E2E8F0',
};


export default function IMCreatorApp({ user, onLogout }) {
  // ============================================================================
  // STATE DECLARATIONS
  // ============================================================================
  const [questionnaire, setQuestionnaire] = useState(defaultQuestionnaire);
  const [currentPhase, setCurrentPhase] = useState(0);
  const [formData, setFormData] = useState({ currency: 'INR' });
  const [completedPhases, setCompletedPhases] = useState([]);
  const [errors, setErrors] = useState({});
  
  // UI State
  const [showConfig, setShowConfig] = useState(false);
  const [showUserMenu, setShowUserMenu] = useState(false);
  const [showAddQ, setShowAddQ] = useState(false);
  const [newQ, setNewQ] = useState({ type: 'text', label: '', required: false, placeholder: '' });
  const [showReport, setShowReport] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isGeneratingPPTX, setIsGeneratingPPTX] = useState(false);
  const [generatedContent, setGeneratedContent] = useState(null);
  const [showGeneratedContent, setShowGeneratedContent] = useState(false);
  const [apiStatus, setApiStatus] = useState('checking');
  const [notification, setNotification] = useState(null);
  
  // Usage tracking states
  const [showUsagePanel, setShowUsagePanel] = useState(false);
  const [usageData, setUsageData] = useState(null);
  const [usageLoading, setUsageLoading] = useState(false);
  
  // v6: Case studies state
  const [caseStudies, setCaseStudies] = useState([
    { id: 1, client: '', industry: '', challenge: '', solution: '', results: '' },
    { id: 2, client: '', industry: '', challenge: '', solution: '', results: '' }
  ]);
  
  // v6: Templates state
  const [templates, setTemplates] = useState([]);
  
  // v6: Auto-logout state
  const [showLogoutWarning, setShowLogoutWarning] = useState(false);
  const [logoutCountdown, setLogoutCountdown] = useState(60);
  const logoutTimerRef = useRef(null);
  const warningTimerRef = useRef(null);
  const countdownRef = useRef(null);

  // Computed values
  const visiblePhases = questionnaire.phases.filter(p => shouldShowPhase(p));
  const phase = questionnaire.phases[currentPhase];
  const questions = phase?.questions?.filter(q => !q.isHidden).sort((a, b) => a.order - b.order) || [];
  const progress = Math.round((completedPhases.length / visiblePhases.length) * 100);

  // ============================================================================
  // AUTO-LOGOUT LOGIC (v6)
  // ============================================================================
  const resetLogoutTimer = useCallback(() => {
    if (logoutTimerRef.current) clearTimeout(logoutTimerRef.current);
    if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
    if (countdownRef.current) clearInterval(countdownRef.current);
    
    setShowLogoutWarning(false);
    setLogoutCountdown(60);
    
    // Warning at 14 minutes
    warningTimerRef.current = setTimeout(() => {
      setShowLogoutWarning(true);
      setLogoutCountdown(60);
      countdownRef.current = setInterval(() => {
        setLogoutCountdown(prev => {
          if (prev <= 1) {
            clearInterval(countdownRef.current);
            return 0;
          }
          return prev - 1;
        });
      }, 1000);
    }, AUTO_LOGOUT_WARNING);
    
    // Logout at 15 minutes
    logoutTimerRef.current = setTimeout(() => {
      handleLogout();
    }, AUTO_LOGOUT_TIMEOUT);
  }, []);

  const handleLogout = useCallback(() => {
    if (logoutTimerRef.current) clearTimeout(logoutTimerRef.current);
    if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
    if (countdownRef.current) clearInterval(countdownRef.current);
    onLogout && onLogout();
  }, [onLogout]);

  // ============================================================================
  // INITIALIZATION EFFECTS
  // ============================================================================
  useEffect(() => {
    // Health check
    checkHealth()
      .then(() => setApiStatus('connected'))
      .catch(() => setApiStatus('disconnected'));
    
    // Fetch templates (v6)
    getTemplates()
      .then(data => {
        setTemplates(data);
        // Update questionnaire with dynamic templates
        if (data && data.length > 0) {
          const templateOptions = data.map(t => ({ value: t.id, label: `${t.name} (${t.category})` }));
          setQuestionnaire(prev => {
            const phases = [...prev.phases];
            const reviewPhase = phases.find(p => p.id === 'review-generate');
            if (reviewPhase) {
              const templateQ = reviewPhase.questions.find(q => q.id === 'templateStyle');
              if (templateQ) {
                templateQ.options = templateOptions;
              }
            }
            return { ...prev, phases };
          });
        }
      })
      .catch(err => console.error('Failed to load templates:', err));
    
    // Auto-logout listeners (v6)
    const events = ['mousedown', 'mousemove', 'keypress', 'scroll', 'touchstart', 'click'];
    const handleActivity = () => resetLogoutTimer();
    events.forEach(e => document.addEventListener(e, handleActivity));
    resetLogoutTimer();
    
    return () => {
      events.forEach(e => document.removeEventListener(e, handleActivity));
      if (logoutTimerRef.current) clearTimeout(logoutTimerRef.current);
      if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
      if (countdownRef.current) clearInterval(countdownRef.current);
    };
  }, [resetLogoutTimer]);

  // ============================================================================
  // HELPER FUNCTIONS
  // ============================================================================
  function shouldShowPhase(p) {
    if (!p.conditionalShow) return true;
    const { field, equals } = p.conditionalShow;
    return formData[field] === equals;
  }

  const showNotification = (message, type = 'info') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 5000);
  };

  const updateField = (id, val) => {
    setFormData(p => ({ ...p, [id]: val }));
    if (errors[id]) setErrors(p => { const e = { ...p }; delete e[id]; return e; });
  };

  const validate = () => {
    const errs = {};
    questions.forEach(q => {
      if (q.required && !formData[q.id]) errs[q.id] = 'Required';
      if (q.validation?.min && formData[q.id] < q.validation.min) errs[q.id] = `Min: ${q.validation.min}`;
      if (q.validation?.max && formData[q.id] > q.validation.max) errs[q.id] = `Max: ${q.validation.max}`;
    });
    setErrors(errs);
    return Object.keys(errs).length === 0;
  };

  const fullValidate = () => {
    const report = { errors: [], warnings: [], suggestions: [] };
    questionnaire.phases.forEach(p => {
      if (!shouldShowPhase(p)) return;
      p.questions.forEach(q => {
        if (q.required && !formData[q.id]) report.errors.push({ phase: p.name, field: q.label, msg: 'Required field missing' });
      });
    });
    
    // v8.1.0: Conditional MANDATORY validations per requirements
    if (formData.targetBuyerType?.includes('financial') && !formData.ebitdaMarginFY25) {
      report.errors.push({ phase: 'Financials', field: 'EBITDA Margin FY25', msg: 'Required for financial buyers' });
    }
    if (formData.targetBuyerType?.includes('strategic') && !formData.synergiesStrategic) {
      report.warnings.push({ phase: 'Synergies', msg: 'Strategic synergies recommended when targeting strategic buyers' });
    }
    if (formData.generateVariants?.includes('market') && !formData.competitorLandscape && !formData.competitiveAdvantages) {
      report.errors.push({ phase: 'Growth Strategy', field: 'Competitive Landscape', msg: 'Required for Market Position variant' });
    }
    if (formData.documentType === 'cim' && !formData.businessRisks) {
      report.errors.push({ phase: 'Risk Factors', field: 'Business Risks', msg: 'At least one risk factor required for CIM' });
    }
    if (formData.documentType === 'cim' && !formData.leadershipTeam && !formData.founderName) {
      report.warnings.push({ phase: 'Leadership', msg: 'Leadership team details recommended for CIM documents' });
    }
    
    const fy25 = parseFloat(formData.revenueFY25) || 0;
    const fy26 = parseFloat(formData.revenueFY26P) || 0;
    if (fy26 && fy25 && ((fy26 - fy25) / fy25 * 100) > 100) {
      report.warnings.push({ phase: 'Financials', msg: 'Projected growth exceeds 100% YoY - verify assumptions' });
    }
    const highlights = (formData.investmentHighlights || '').split('\n').filter(l => l.trim()).length;
    if (highlights < 5) report.suggestions.push({ phase: 'Company Overview', msg: `Only ${highlights} investment highlights (recommended: 5-7)` });
    const advantages = (formData.competitiveAdvantages || '').split('\n').filter(l => l.trim()).length;
    if (advantages < 5) report.suggestions.push({ phase: 'Growth Strategy', msg: `Only ${advantages} competitive advantages (recommended: 5+)` });
    return report;
  };

  // ============================================================================
  // NAVIGATION HANDLERS
  // ============================================================================
  const handleNext = () => {
    if (validate()) {
      if (!completedPhases.includes(currentPhase)) setCompletedPhases([...completedPhases, currentPhase]);
      // Find next visible phase
      let next = currentPhase + 1;
      while (next < questionnaire.phases.length && !shouldShowPhase(questionnaire.phases[next])) {
        next++;
      }
      if (next < questionnaire.phases.length) {
        setCurrentPhase(next);
      }
    }
  };

  const handlePrev = () => {
    let prev = currentPhase - 1;
    while (prev >= 0 && !shouldShowPhase(questionnaire.phases[prev])) {
      prev--;
    }
    if (prev >= 0) {
      setCurrentPhase(prev);
    }
  };

  // ============================================================================
  // CASE STUDY HANDLERS (v6)
  // ============================================================================
  const addCaseStudy = () => {
    const newId = Math.max(...caseStudies.map(cs => cs.id), 0) + 1;
    setCaseStudies([...caseStudies, { id: newId, client: '', industry: '', challenge: '', solution: '', results: '' }]);
    showNotification(`Case Study ${caseStudies.length + 1} added`, 'success');
  };

  const removeCaseStudy = (id) => {
    if (caseStudies.length <= 1) {
      showNotification('At least one case study is required', 'warning');
      return;
    }
    setCaseStudies(caseStudies.filter(cs => cs.id !== id));
    showNotification('Case study removed', 'info');
  };

  const updateCaseStudy = (id, field, value) => {
    setCaseStudies(caseStudies.map(cs => cs.id === id ? { ...cs, [field]: value } : cs));
  };

  // ============================================================================
  // QUESTION HANDLERS
  // ============================================================================
  const addQuestion = () => {
    if (!newQ.label.trim()) return;
    const phases = [...questionnaire.phases];
    phases[currentPhase].questions.push({
      id: `custom_${Date.now()}`,
      type: newQ.type,
      label: newQ.label,
      required: newQ.required,
      placeholder: newQ.placeholder,
      isCustom: true,
      order: phase.questions.length + 1
    });
    setQuestionnaire({ ...questionnaire, phases });
    setShowAddQ(false);
    setNewQ({ type: 'text', label: '', required: false, placeholder: '' });
    showNotification('Custom question added successfully!', 'success');
  };

  const toggleHide = (qId) => {
    const phases = [...questionnaire.phases];
    const q = phases[currentPhase].questions.find(x => x.id === qId);
    if (q) q.isHidden = !q.isHidden;
    setQuestionnaire({ ...questionnaire, phases });
  };

  const removeQ = (qId) => {
    const phases = [...questionnaire.phases];
    phases[currentPhase].questions = phases[currentPhase].questions.filter(q => q.id !== qId);
    setQuestionnaire({ ...questionnaire, phases });
    showNotification('Question removed', 'info');
  };

  // ============================================================================
  // USAGE TRACKING
  // ============================================================================
  const fetchUsageData = async () => {
    setUsageLoading(true);
    try {
      const data = await getUsageStats();
      setUsageData(data);
    } catch (error) {
      console.error('Failed to fetch usage data:', error);
      showNotification('Failed to fetch usage data', 'error');
    } finally {
      setUsageLoading(false);
    }
  };

  const exportUsageCSVHandler = async () => {
    try {
      const blob = await exportUsageCSV();
      downloadBlob(blob, `usage_report_${Date.now()}.csv`);
      showNotification('Usage report downloaded', 'success');
    } catch (error) {
      console.error('Failed to export usage:', error);
      showNotification('Failed to export usage data', 'error');
    }
  };

  const resetUsageHandler = async () => {
    if (!window.confirm('Are you sure you want to reset all usage statistics? This cannot be undone.')) return;
    try {
      await resetUsageStats();
      await fetchUsageData();
      showNotification('Usage statistics reset', 'success');
    } catch (error) {
      console.error('Failed to reset usage:', error);
      showNotification('Failed to reset usage data', 'error');
    }
  };

  const openUsagePanel = () => {
    setShowUsagePanel(true);
    fetchUsageData();
  };

  // ============================================================================
  // SAVE & GENERATE HANDLERS
  // ============================================================================
  const handleSaveDraft = async () => {
    try {
      await saveDraft(formData, formData.projectCodename || `draft_${Date.now()}`);
      showNotification('Draft saved successfully!', 'success');
    } catch (error) {
      showNotification('Failed to save draft. Please try again.', 'error');
    }
  };

  const handleGenerateJSON = async () => {
    const r = fullValidate();
    if (r.errors.length) {
      setShowReport(true);
      return;
    }
    if (apiStatus !== 'connected') {
      showNotification('API is not connected. Please check your connection.', 'error');
      return;
    }
    setIsGenerating(true);
    try {
      const result = await generateIM(formData);
      setGeneratedContent(result);
      setShowGeneratedContent(true);
      showNotification('IM content generated successfully!', 'success');
    } catch (error) {
      showNotification('Failed to generate IM. Please try again.', 'error');
      console.error('Generation error:', error);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleGeneratePPTX = async () => {
    const r = fullValidate();
    if (r.errors.length) {
      setShowReport(true);
      return;
    }
    if (apiStatus !== 'connected') {
      showNotification('API is not connected. Please check your connection.', 'error');
      return;
    }
    setIsGeneratingPPTX(true);
    showNotification('Generating Professional PowerPoint... This may take a moment.', 'info');
    
    try {
      // v6: Include case studies in data
      const dataToSend = {
        ...formData,
        caseStudies: caseStudies.filter(cs => cs.client)
      };
      
      const theme = formData.templateStyle || 'modern-blue';
      const result = await generatePPTX(dataToSend, theme);
      
      if (result.success && result.fileData) {
        downloadBase64File(result.fileData, result.filename, result.mimeType);
        showNotification(`PowerPoint downloaded! (${result.slideCount} slides)`, 'success');
      } else {
        throw new Error(result.error || 'Invalid response from server');
      }
    } catch (error) {
      showNotification('Failed to generate PowerPoint. Please try again.', 'error');
      console.error('PPTX Generation error:', error);
    } finally {
      setIsGeneratingPPTX(false);
    }
  };

  // ============================================================================
  // EXPORT HANDLERS (v6)
  // ============================================================================
  const downloadJSON = () => {
    if (generatedContent) {
      const blob = new Blob([JSON.stringify(generatedContent.content, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${formData.projectCodename || 'IM'}_content.json`;
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  const downloadJSONData = () => {
    const exportData = {
      metadata: {
        projectCodename: formData.projectCodename,
        generatedAt: new Date().toISOString(),
        version: '8.1.0'
      },
      formData,
      caseStudies
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${formData.projectCodename || 'IM'}_data.json`;
    a.click();
    URL.revokeObjectURL(url);
    showNotification('JSON data downloaded!', 'success');
  };

  const downloadWordQA = async () => {
    try {
      showNotification('Generating Word document...', 'info');
      
      const blob = await exportQAWord({ ...formData, caseStudies }, questionnaire);
      downloadBlob(blob, `${formData.projectCodename || 'QA'}_Document.docx`);
      
      showNotification('Word document downloaded!', 'success');
    } catch (error) {
      showNotification('Word export failed. Ensure docx package is installed on server.', 'warning');
      console.error('Word export error:', error);
    }
  };


  // ============================================================================
  // RENDER FIELD
  // ============================================================================
  const renderField = (q) => {
    const val = formData[q.id] ?? q.defaultValue ?? (q.type === 'multiselect' ? [] : '');
    const err = errors[q.id];
    const baseInputStyle = {
      width: '100%',
      padding: '12px 16px',
      border: `1px solid ${err ? THEME.accentRed : THEME.border}`,
      borderRadius: '8px',
      fontSize: '14px',
      outline: 'none',
      transition: 'all 0.2s ease',
      backgroundColor: THEME.surface,
      color: THEME.text,
      boxSizing: 'border-box'
    };

    switch (q.type) {
      case 'textarea':
        return (
          <textarea
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            rows={4}
            style={{ ...baseInputStyle, resize: 'vertical', minHeight: '100px' }}
            onFocus={e => e.target.style.borderColor = THEME.primary}
            onBlur={e => e.target.style.borderColor = err ? THEME.accentRed : THEME.border}
          />
        );
      case 'select':
        return (
          <select
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            style={{ ...baseInputStyle, cursor: 'pointer' }}
          >
            <option value="">Select...</option>
            {q.options?.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
          </select>
        );
      case 'multiselect':
        return (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            {q.options?.map(o => (
              <label key={o.value} style={{
                display: 'flex',
                alignItems: 'flex-start',
                gap: '12px',
                padding: '12px 16px',
                backgroundColor: (val || []).includes(o.value) ? `${THEME.primary}10` : THEME.background,
                borderRadius: '8px',
                cursor: 'pointer',
                border: `1px solid ${(val || []).includes(o.value) ? THEME.primary : THEME.border}`,
                transition: 'all 0.2s ease'
              }}>
                <input
                  type="checkbox"
                  checked={(val || []).includes(o.value)}
                  onChange={e => {
                    const arr = val || [];
                    updateField(q.id, e.target.checked ? [...arr, o.value] : arr.filter(v => v !== o.value));
                  }}
                  style={{ marginTop: '2px', width: '18px', height: '18px', accentColor: THEME.primary }}
                />
                <span style={{ fontSize: '14px', color: THEME.text }}>{o.label}</span>
              </label>
            ))}
          </div>
        );
      case 'number':
        return (
          <input
            type="number"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            min={q.validation?.min}
            max={q.validation?.max}
            style={baseInputStyle}
            onFocus={e => e.target.style.borderColor = THEME.primary}
            onBlur={e => e.target.style.borderColor = err ? THEME.accentRed : THEME.border}
          />
        );
      case 'date':
        return (
          <input
            type="date"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            style={{ ...baseInputStyle, cursor: 'pointer' }}
            onFocus={e => e.target.style.borderColor = THEME.primary}
            onBlur={e => e.target.style.borderColor = err ? THEME.accentRed : THEME.border}
          />
        );
      default:
        return (
          <input
            type="text"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            style={baseInputStyle}
            onFocus={e => e.target.style.borderColor = THEME.primary}
            onBlur={e => e.target.style.borderColor = err ? THEME.accentRed : THEME.border}
          />
        );
    }
  };

  // ============================================================================
  // RENDER CASE STUDIES (v6 - Dynamic)
  // ============================================================================
  const renderCaseStudies = () => (
    <div>
      {caseStudies.map((cs, idx) => (
        <div key={cs.id} style={{
          marginBottom: '24px',
          padding: '20px',
          backgroundColor: THEME.background,
          borderRadius: '12px',
          border: `1px solid ${THEME.border}`
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h4 style={{ margin: 0, color: THEME.primary, fontSize: '16px' }}>Case Study {idx + 1}</h4>
            {caseStudies.length > 1 && (
              <button
                onClick={() => removeCaseStudy(cs.id)}
                style={{
                  padding: '6px 12px',
                  backgroundColor: THEME.accentRed,
                  color: 'white',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  fontSize: '12px'
                }}
              >
                Remove
              </button>
            )}
          </div>
          
          <div style={{ display: 'grid', gap: '16px' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px', color: THEME.text }}>
                  Client Name
                </label>
                <input
                  type="text"
                  value={cs.client}
                  onChange={(e) => updateCaseStudy(cs.id, 'client', e.target.value)}
                  placeholder="e.g., HDFC Bank"
                  style={{
                    width: '100%',
                    padding: '12px 16px',
                    border: `1px solid ${THEME.border}`,
                    borderRadius: '8px',
                    fontSize: '14px',
                    outline: 'none',
                    boxSizing: 'border-box'
                  }}
                />
              </div>
              <div>
                <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px', color: THEME.text }}>
                  Industry
                </label>
                <input
                  type="text"
                  value={cs.industry}
                  onChange={(e) => updateCaseStudy(cs.id, 'industry', e.target.value)}
                  placeholder="e.g., Financial Services"
                  style={{
                    width: '100%',
                    padding: '12px 16px',
                    border: `1px solid ${THEME.border}`,
                    borderRadius: '8px',
                    fontSize: '14px',
                    outline: 'none',
                    boxSizing: 'border-box'
                  }}
                />
              </div>
            </div>
            
            <div>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px', color: THEME.text }}>
                Challenge
              </label>
              <textarea
                value={cs.challenge}
                onChange={(e) => updateCaseStudy(cs.id, 'challenge', e.target.value)}
                placeholder="Describe the business challenge faced by the client..."
                rows={3}
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  outline: 'none',
                  boxSizing: 'border-box',
                  resize: 'vertical'
                }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px', color: THEME.text }}>
                Solution
              </label>
              <textarea
                value={cs.solution}
                onChange={(e) => updateCaseStudy(cs.id, 'solution', e.target.value)}
                placeholder="How your company solved the problem..."
                rows={3}
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  outline: 'none',
                  boxSizing: 'border-box',
                  resize: 'vertical'
                }}
              />
            </div>
            
            <div>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px', color: THEME.text }}>
                Results
              </label>
              <textarea
                value={cs.results}
                onChange={(e) => updateCaseStudy(cs.id, 'results', e.target.value)}
                placeholder="40% reduction in processing time&#10;60% cost savings&#10;Improved customer satisfaction"
                rows={3}
                style={{
                  width: '100%',
                  padding: '12px 16px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  outline: 'none',
                  boxSizing: 'border-box',
                  resize: 'vertical'
                }}
              />
              <p style={{ margin: '6px 0 0 0', fontSize: '12px', color: THEME.textLight }}>
                Quantified outcomes (one per line)
              </p>
            </div>
          </div>
        </div>
      ))}
      
      {/* Add Case Study Button */}
      <button
        onClick={addCaseStudy}
        style={{
          width: '100%',
          padding: '16px',
          backgroundColor: THEME.background,
          color: THEME.primary,
          border: `2px dashed ${THEME.primary}`,
          borderRadius: '12px',
          cursor: 'pointer',
          fontSize: '16px',
          fontWeight: '600',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '8px',
          transition: 'all 0.2s ease'
        }}
        onMouseEnter={e => e.target.style.backgroundColor = `${THEME.primary}10`}
        onMouseLeave={e => e.target.style.backgroundColor = THEME.background}
      >
        <span style={{ fontSize: '24px' }}>+</span>
        Add Case Study
      </button>
    </div>
  );


  // ============================================================================
  // MAIN RENDER
  // ============================================================================
  return (
    <div style={{
      minHeight: '100vh',
      backgroundColor: THEME.background,
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
      display: 'flex',
      flexDirection: 'column'
    }}>
      {/* Header */}
      <header style={{
        backgroundColor: THEME.surface,
        borderBottom: `1px solid ${THEME.border}`,
        padding: '12px 24px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        position: 'sticky',
        top: 0,
        zIndex: 100
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <h1 style={{
            margin: 0,
            fontSize: '20px',
            fontWeight: '700',
            color: THEME.primary
          }}>
            IM Creator
          </h1>
          <span style={{
            padding: '4px 10px',
            backgroundColor: `${THEME.accent}20`,
            color: THEME.accent,
            borderRadius: '12px',
            fontSize: '11px',
            fontWeight: '600'
          }}>
            v6.0
          </span>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '6px',
            padding: '6px 12px',
            backgroundColor: apiStatus === 'connected' ? `${THEME.accent}15` : `${THEME.accentRed}15`,
            borderRadius: '16px'
          }}>
            <div style={{
              width: '8px',
              height: '8px',
              borderRadius: '50%',
              backgroundColor: apiStatus === 'connected' ? THEME.accent : THEME.accentRed
            }} />
            <span style={{
              fontSize: '12px',
              color: apiStatus === 'connected' ? THEME.accent : THEME.accentRed,
              fontWeight: '500'
            }}>
              {apiStatus === 'connected' ? 'API Connected' : 'API Disconnected'}
            </span>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          {/* Configure Button */}
          <button
            onClick={() => setShowConfig(!showConfig)}
            style={{
              padding: '8px 16px',
              backgroundColor: showConfig ? THEME.primary : THEME.background,
              color: showConfig ? 'white' : THEME.text,
              border: `1px solid ${showConfig ? THEME.primary : THEME.border}`,
              borderRadius: '8px',
              cursor: 'pointer',
              fontSize: '13px',
              fontWeight: '500',
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              transition: 'all 0.2s ease'
            }}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
            </svg>
            Configure
          </button>

          {/* Usage Button */}
          <button
            onClick={openUsagePanel}
            style={{
              padding: '8px 16px',
              backgroundColor: THEME.background,
              color: THEME.text,
              border: `1px solid ${THEME.border}`,
              borderRadius: '8px',
              cursor: 'pointer',
              fontSize: '13px',
              fontWeight: '500',
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              transition: 'all 0.2s ease'
            }}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
            </svg>
            Usage
          </button>

          {/* User Menu */}
          <div style={{ position: 'relative' }}>
            <button
              onClick={() => setShowUserMenu(!showUserMenu)}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                padding: '6px 12px 6px 6px',
                backgroundColor: THEME.background,
                border: `1px solid ${THEME.border}`,
                borderRadius: '24px',
                cursor: 'pointer',
                transition: 'all 0.2s ease'
              }}
            >
              <div style={{
                width: '32px',
                height: '32px',
                borderRadius: '50%',
                backgroundColor: THEME.primary,
                color: 'white',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: '14px',
                fontWeight: '600'
              }}>
                {(user?.username || 'U').charAt(0).toUpperCase()}
              </div>
              <span style={{ fontSize: '14px', fontWeight: '500', color: THEME.text }}>
                {user?.username || 'User'}
              </span>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={THEME.textLight} strokeWidth="2">
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>

            {showUserMenu && (
              <div style={{
                position: 'absolute',
                top: '100%',
                right: 0,
                marginTop: '8px',
                backgroundColor: THEME.surface,
                borderRadius: '12px',
                boxShadow: '0 4px 20px rgba(0,0,0,0.15)',
                border: `1px solid ${THEME.border}`,
                minWidth: '200px',
                overflow: 'hidden',
                zIndex: 1000
              }}>
                <div style={{ padding: '16px', borderBottom: `1px solid ${THEME.border}` }}>
                  <div style={{ fontSize: '14px', fontWeight: '600', color: THEME.text }}>
                    {user?.username || 'User'}
                  </div>
                  <div style={{ fontSize: '12px', color: THEME.textLight, marginTop: '2px' }}>
                    Logged in via {user?.method === 'office365' ? 'Office 365' : 'Credentials'}
                  </div>
                </div>
                <button
                  onClick={() => { setShowUserMenu(false); handleLogout(); }}
                  style={{
                    width: '100%',
                    padding: '12px 16px',
                    backgroundColor: 'transparent',
                    border: 'none',
                    cursor: 'pointer',
                    fontSize: '14px',
                    color: THEME.accentRed,
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    transition: 'background-color 0.15s ease'
                  }}
                  onMouseEnter={e => e.target.style.backgroundColor = `${THEME.accentRed}10`}
                  onMouseLeave={e => e.target.style.backgroundColor = 'transparent'}
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
                    <polyline points="16 17 21 12 16 7"/>
                    <line x1="21" y1="12" x2="9" y2="12"/>
                  </svg>
                  Sign Out
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        {/* Left Sidebar */}
        <aside style={{
          width: '280px',
          backgroundColor: THEME.surface,
          borderRight: `1px solid ${THEME.border}`,
          display: 'flex',
          flexDirection: 'column',
          flexShrink: 0
        }}>
          <div style={{ padding: '20px', borderBottom: `1px solid ${THEME.border}` }}>
            <span style={{ fontWeight: '600', color: THEME.text, fontSize: '14px' }}>Sections</span>
          </div>

          <nav style={{ flex: 1, overflowY: 'auto', padding: '12px' }}>
            {questionnaire.phases.filter(p => shouldShowPhase(p)).map((p, idx) => {
              const actualIdx = questionnaire.phases.indexOf(p);
              const isActive = actualIdx === currentPhase;
              const isCompleted = completedPhases.includes(actualIdx);
              
              return (
                <button
                  key={p.id}
                  onClick={() => setCurrentPhase(actualIdx)}
                  style={{
                    width: '100%',
                    padding: '14px 16px',
                    marginBottom: '4px',
                    backgroundColor: isActive ? `${THEME.primary}10` : 'transparent',
                    border: 'none',
                    borderRadius: '8px',
                    borderLeft: isActive ? `3px solid ${THEME.primary}` : '3px solid transparent',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '12px',
                    transition: 'all 0.15s ease',
                    textAlign: 'left'
                  }}
                >
                  <span style={{ fontSize: '18px' }}>{p.icon}</span>
                  <div style={{ flex: 1 }}>
                    <div style={{
                      fontSize: '14px',
                      fontWeight: isActive ? '600' : '500',
                      color: isActive ? THEME.primary : THEME.text
                    }}>
                      {p.name}
                    </div>
                    <div style={{ fontSize: '11px', color: THEME.textLight, marginTop: '2px' }}>
                      {p.description}
                    </div>
                  </div>
                  {isCompleted && (
                    <svg width="18" height="18" viewBox="0 0 24 24" fill={THEME.accent}>
                      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                    </svg>
                  )}
                </button>
              );
            })}
          </nav>

          {/* Progress Section */}
          <div style={{ padding: '20px', borderTop: `1px solid ${THEME.border}` }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px' }}>
              <span style={{ fontSize: '12px', fontWeight: '600', color: THEME.text }}>Progress</span>
              <span style={{ fontSize: '12px', fontWeight: '600', color: THEME.primary }}>{progress}%</span>
            </div>
            <div style={{
              height: '6px',
              backgroundColor: THEME.border,
              borderRadius: '3px',
              overflow: 'hidden'
            }}>
              <div style={{
                width: `${progress}%`,
                height: '100%',
                backgroundColor: THEME.primary,
                borderRadius: '3px',
                transition: 'width 0.3s ease'
              }} />
            </div>
            <button
              onClick={handleSaveDraft}
              style={{
                width: '100%',
                marginTop: '16px',
                padding: '10px',
                backgroundColor: THEME.background,
                border: `1px solid ${THEME.border}`,
                borderRadius: '8px',
                cursor: 'pointer',
                fontSize: '13px',
                fontWeight: '500',
                color: THEME.text,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px'
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                <polyline points="17 21 17 13 7 13 7 21"/>
                <polyline points="7 3 7 8 15 8"/>
              </svg>
              Save Draft
            </button>
          </div>
        </aside>


        {/* Main Content */}
        <main style={{ flex: 1, overflow: 'auto', padding: '32px' }}>
          <div style={{ maxWidth: '900px', margin: '0 auto' }}>
            {/* Phase Header */}
            <div style={{ marginBottom: '32px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '8px' }}>
                <span style={{ fontSize: '32px' }}>{phase?.icon}</span>
                <h2 style={{ margin: 0, fontSize: '28px', fontWeight: '700', color: THEME.text }}>
                  {phase?.name}
                </h2>
              </div>
              <p style={{ margin: 0, color: THEME.textLight, fontSize: '15px' }}>
                {phase?.description}
              </p>
            </div>

            {/* Configure Panel */}
            {showConfig && (
              <div style={{
                backgroundColor: THEME.surface,
                borderRadius: '12px',
                padding: '20px',
                marginBottom: '24px',
                border: `1px solid ${THEME.border}`
              }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                  <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '600', color: THEME.text }}>
                    Configure Questions
                  </h3>
                  <button
                    onClick={() => setShowAddQ(true)}
                    style={{
                      padding: '8px 16px',
                      backgroundColor: THEME.primary,
                      color: 'white',
                      border: 'none',
                      borderRadius: '6px',
                      cursor: 'pointer',
                      fontSize: '13px',
                      fontWeight: '500'
                    }}
                  >
                    + Add Question
                  </button>
                </div>
                
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {phase?.questions?.map(q => (
                    <div key={q.id} style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      padding: '10px 14px',
                      backgroundColor: q.isHidden ? `${THEME.textLight}10` : THEME.background,
                      borderRadius: '8px',
                      opacity: q.isHidden ? 0.6 : 1
                    }}>
                      <span style={{ fontSize: '14px', color: THEME.text }}>
                        {q.label} {q.required && <span style={{ color: THEME.accentRed }}>*</span>}
                        {q.isCustom && <span style={{ color: THEME.accentBlue, marginLeft: '8px', fontSize: '11px' }}>(Custom)</span>}
                      </span>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button
                          onClick={() => toggleHide(q.id)}
                          style={{
                            padding: '4px 10px',
                            backgroundColor: 'transparent',
                            border: `1px solid ${THEME.border}`,
                            borderRadius: '4px',
                            cursor: 'pointer',
                            fontSize: '12px',
                            color: THEME.textLight
                          }}
                        >
                          {q.isHidden ? 'Show' : 'Hide'}
                        </button>
                        {q.isCustom && (
                          <button
                            onClick={() => removeQ(q.id)}
                            style={{
                              padding: '4px 10px',
                              backgroundColor: `${THEME.accentRed}10`,
                              border: 'none',
                              borderRadius: '4px',
                              cursor: 'pointer',
                              fontSize: '12px',
                              color: THEME.accentRed
                            }}
                          >
                            Remove
                          </button>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Questions / Dynamic Content */}
            <div style={{
              backgroundColor: THEME.surface,
              borderRadius: '16px',
              padding: '32px',
              boxShadow: '0 1px 3px rgba(0,0,0,0.05)'
            }}>
              {/* Check if this is the dynamic case studies phase */}
              {phase?.isDynamic && phase?.id === 'case-studies' ? (
                renderCaseStudies()
              ) : (
                questions.map(q => (
                  <div key={q.id} style={{ marginBottom: '24px' }}>
                    <label style={{
                      display: 'block',
                      fontSize: '14px',
                      fontWeight: '600',
                      color: THEME.text,
                      marginBottom: '8px'
                    }}>
                      {q.label}
                      {q.required && <span style={{ color: THEME.accentRed, marginLeft: '4px' }}>*</span>}
                    </label>
                    {renderField(q)}
                    {q.helpText && (
                      <p style={{ margin: '6px 0 0 0', fontSize: '12px', color: THEME.textLight }}>
                        {q.helpText}
                      </p>
                    )}
                    {errors[q.id] && (
                      <p style={{ margin: '6px 0 0 0', fontSize: '12px', color: THEME.accentRed }}>
                        {errors[q.id]}
                      </p>
                    )}
                  </div>
                ))
              )}

              {/* Navigation Buttons */}
              <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                marginTop: '32px',
                paddingTop: '24px',
                borderTop: `1px solid ${THEME.border}`
              }}>
                <button
                  onClick={handlePrev}
                  disabled={currentPhase === 0}
                  style={{
                    padding: '12px 24px',
                    backgroundColor: THEME.background,
                    color: currentPhase === 0 ? THEME.textLight : THEME.text,
                    border: `1px solid ${THEME.border}`,
                    borderRadius: '8px',
                    cursor: currentPhase === 0 ? 'not-allowed' : 'pointer',
                    fontSize: '14px',
                    fontWeight: '500',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    opacity: currentPhase === 0 ? 0.5 : 1
                  }}
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <polyline points="15 18 9 12 15 6"/>
                  </svg>
                  Previous
                </button>

                {currentPhase === questionnaire.phases.length - 1 ? (
                  <div style={{ display: 'flex', gap: '12px' }}>
                    <button
                      onClick={() => setShowReport(true)}
                      style={{
                        padding: '12px 24px',
                        backgroundColor: THEME.background,
                        color: THEME.text,
                        border: `1px solid ${THEME.border}`,
                        borderRadius: '8px',
                        cursor: 'pointer',
                        fontSize: '14px',
                        fontWeight: '500'
                      }}
                    >
                      Review
                    </button>
                    <button
                      onClick={handleGeneratePPTX}
                      disabled={isGeneratingPPTX}
                      style={{
                        padding: '12px 24px',
                        backgroundColor: THEME.primary,
                        color: 'white',
                        border: 'none',
                        borderRadius: '8px',
                        cursor: isGeneratingPPTX ? 'wait' : 'pointer',
                        fontSize: '14px',
                        fontWeight: '600',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px'
                      }}
                    >
                      {isGeneratingPPTX ? (
                        <>
                          <div style={{
                            width: '16px',
                            height: '16px',
                            border: '2px solid rgba(255,255,255,0.3)',
                            borderTopColor: 'white',
                            borderRadius: '50%',
                            animation: 'spin 1s linear infinite'
                          }} />
                          Generating...
                        </>
                      ) : (
                        <>
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                          </svg>
                          Generate PowerPoint
                        </>
                      )}
                    </button>
                  </div>
                ) : (
                  <button
                    onClick={handleNext}
                    style={{
                      padding: '12px 24px',
                      backgroundColor: THEME.primary,
                      color: 'white',
                      border: 'none',
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      fontWeight: '600',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px'
                    }}
                  >
                    Next
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="9 18 15 12 9 6"/>
                    </svg>
                  </button>
                )}
              </div>

              {/* Export Buttons (v6) - Show on Review & Generate phase */}
              {phase?.id === 'review-generate' && (
                <div style={{ 
                  display: 'flex', 
                  gap: '12px', 
                  marginTop: '16px',
                  paddingTop: '16px',
                  borderTop: `1px solid ${THEME.border}`
                }}>
                  <button
                    onClick={downloadJSONData}
                    style={{
                      flex: 1,
                      padding: '12px',
                      backgroundColor: THEME.background,
                      color: THEME.text,
                      border: `1px solid ${THEME.border}`,
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '8px'
                    }}
                  >
                    📄 Export JSON Data
                  </button>
                  <button
                    onClick={downloadWordQA}
                    style={{
                      flex: 1,
                      padding: '12px',
                      backgroundColor: THEME.background,
                      color: THEME.text,
                      border: `1px solid ${THEME.border}`,
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '8px'
                    }}
                  >
                    📝 Export Word Q&A
                  </button>
                </div>
              )}
            </div>
          </div>
        </main>
      </div>


      {/* Notification Toast */}
      {notification && (
        <div style={{
          position: 'fixed',
          bottom: '24px',
          right: '24px',
          padding: '16px 24px',
          backgroundColor: notification.type === 'error' ? THEME.accentRed : 
                          notification.type === 'success' ? THEME.accent : 
                          notification.type === 'warning' ? THEME.accentYellow : THEME.accentBlue,
          color: 'white',
          borderRadius: '12px',
          boxShadow: '0 4px 20px rgba(0,0,0,0.2)',
          zIndex: 2000,
          display: 'flex',
          alignItems: 'center',
          gap: '12px',
          maxWidth: '400px'
        }}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
            {notification.type === 'error' ? (
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
            ) : notification.type === 'success' ? (
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
            ) : (
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/>
            )}
          </svg>
          <span style={{ fontSize: '14px', fontWeight: '500' }}>{notification.message}</span>
        </div>
      )}

      {/* Validation Report Modal */}
      {showReport && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1500
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            maxWidth: '600px',
            width: '90%',
            maxHeight: '80vh',
            overflow: 'auto'
          }}>
            <h2 style={{ margin: '0 0 24px 0', color: THEME.text }}>Validation Report</h2>
            
            {(() => {
              const report = fullValidate();
              return (
                <>
                  {report.errors.length > 0 && (
                    <div style={{ marginBottom: '20px' }}>
                      <h3 style={{ color: THEME.accentRed, fontSize: '16px', marginBottom: '12px' }}>
                        ❌ Errors ({report.errors.length})
                      </h3>
                      {report.errors.map((e, i) => (
                        <div key={i} style={{
                          padding: '12px',
                          backgroundColor: `${THEME.accentRed}10`,
                          borderRadius: '8px',
                          marginBottom: '8px',
                          fontSize: '14px'
                        }}>
                          <strong>{e.phase}</strong>: {e.field} - {e.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  
                  {report.warnings.length > 0 && (
                    <div style={{ marginBottom: '20px' }}>
                      <h3 style={{ color: THEME.accentYellow, fontSize: '16px', marginBottom: '12px' }}>
                        ⚠️ Warnings ({report.warnings.length})
                      </h3>
                      {report.warnings.map((w, i) => (
                        <div key={i} style={{
                          padding: '12px',
                          backgroundColor: `${THEME.accentYellow}15`,
                          borderRadius: '8px',
                          marginBottom: '8px',
                          fontSize: '14px'
                        }}>
                          <strong>{w.phase}</strong>: {w.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  
                  {report.suggestions.length > 0 && (
                    <div style={{ marginBottom: '20px' }}>
                      <h3 style={{ color: THEME.accentBlue, fontSize: '16px', marginBottom: '12px' }}>
                        💡 Suggestions ({report.suggestions.length})
                      </h3>
                      {report.suggestions.map((s, i) => (
                        <div key={i} style={{
                          padding: '12px',
                          backgroundColor: `${THEME.accentBlue}10`,
                          borderRadius: '8px',
                          marginBottom: '8px',
                          fontSize: '14px'
                        }}>
                          <strong>{s.phase}</strong>: {s.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  
                  {report.errors.length === 0 && report.warnings.length === 0 && report.suggestions.length === 0 && (
                    <div style={{
                      padding: '24px',
                      backgroundColor: `${THEME.accent}10`,
                      borderRadius: '12px',
                      textAlign: 'center'
                    }}>
                      <svg width="48" height="48" viewBox="0 0 24 24" fill={THEME.accent} style={{ marginBottom: '12px' }}>
                        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                      </svg>
                      <p style={{ margin: 0, fontSize: '16px', color: THEME.accent, fontWeight: '600' }}>
                        All validations passed! Ready to generate.
                      </p>
                    </div>
                  )}
                </>
              );
            })()}
            
            <button
              onClick={() => setShowReport(false)}
              style={{
                marginTop: '24px',
                padding: '12px 24px',
                backgroundColor: THEME.primary,
                color: 'white',
                border: 'none',
                borderRadius: '8px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                width: '100%'
              }}
            >
              Close
            </button>
          </div>
        </div>
      )}

      {/* Add Question Modal */}
      {showAddQ && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1500
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            maxWidth: '500px',
            width: '90%'
          }}>
            <h2 style={{ margin: '0 0 24px 0', color: THEME.text }}>Add Custom Question</h2>
            
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px' }}>Question Label</label>
              <input
                type="text"
                value={newQ.label}
                onChange={e => setNewQ({ ...newQ, label: e.target.value })}
                placeholder="e.g., Additional Notes"
                style={{
                  width: '100%',
                  padding: '12px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  boxSizing: 'border-box'
                }}
              />
            </div>
            
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px' }}>Type</label>
              <select
                value={newQ.type}
                onChange={e => setNewQ({ ...newQ, type: e.target.value })}
                style={{
                  width: '100%',
                  padding: '12px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  boxSizing: 'border-box'
                }}
              >
                <option value="text">Text</option>
                <option value="textarea">Text Area</option>
                <option value="number">Number</option>
                <option value="date">Date</option>
              </select>
            </div>
            
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '14px', fontWeight: '500', marginBottom: '6px' }}>Placeholder</label>
              <input
                type="text"
                value={newQ.placeholder}
                onChange={e => setNewQ({ ...newQ, placeholder: e.target.value })}
                placeholder="Optional placeholder text"
                style={{
                  width: '100%',
                  padding: '12px',
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  fontSize: '14px',
                  boxSizing: 'border-box'
                }}
              />
            </div>
            
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
              <input
                type="checkbox"
                checked={newQ.required}
                onChange={e => setNewQ({ ...newQ, required: e.target.checked })}
                style={{ width: '18px', height: '18px' }}
              />
              <span style={{ fontSize: '14px' }}>Required field</span>
            </label>
            
            <div style={{ display: 'flex', gap: '12px' }}>
              <button
                onClick={() => setShowAddQ(false)}
                style={{
                  flex: 1,
                  padding: '12px',
                  backgroundColor: THEME.background,
                  color: THEME.text,
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '14px'
                }}
              >
                Cancel
              </button>
              <button
                onClick={addQuestion}
                style={{
                  flex: 1,
                  padding: '12px',
                  backgroundColor: THEME.primary,
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '14px',
                  fontWeight: '600'
                }}
              >
                Add Question
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Usage Panel Modal */}
      {showUsagePanel && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1500
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            maxWidth: '800px',
            width: '90%',
            maxHeight: '80vh',
            overflow: 'auto'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
              <h2 style={{ margin: 0, color: THEME.text }}>API Usage Dashboard</h2>
              <button
                onClick={() => setShowUsagePanel(false)}
                style={{
                  padding: '8px',
                  backgroundColor: 'transparent',
                  border: 'none',
                  cursor: 'pointer',
                  borderRadius: '8px'
                }}
              >
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke={THEME.textLight} strokeWidth="2">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>
            
            {usageLoading ? (
              <div style={{ textAlign: 'center', padding: '40px' }}>
                <div style={{
                  width: '40px',
                  height: '40px',
                  border: '3px solid rgba(0,0,0,0.1)',
                  borderTopColor: THEME.primary,
                  borderRadius: '50%',
                  animation: 'spin 1s linear infinite',
                  margin: '0 auto 16px'
                }} />
                <p style={{ color: THEME.textLight }}>Loading usage data...</p>
              </div>
            ) : usageData ? (
              <>
                {/* Summary Cards */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px', marginBottom: '24px' }}>
                  <div style={{
                    padding: '20px',
                    backgroundColor: `${THEME.primary}10`,
                    borderRadius: '12px',
                    textAlign: 'center'
                  }}>
                    <div style={{ fontSize: '28px', fontWeight: '700', color: THEME.primary }}>{usageData.totalCalls}</div>
                    <div style={{ fontSize: '12px', color: THEME.textLight, marginTop: '4px' }}>Total Calls</div>
                  </div>
                  <div style={{
                    padding: '20px',
                    backgroundColor: `${THEME.accent}10`,
                    borderRadius: '12px',
                    textAlign: 'center'
                  }}>
                    <div style={{ fontSize: '28px', fontWeight: '700', color: THEME.accent }}>${usageData.totalCostUSD}</div>
                    <div style={{ fontSize: '12px', color: THEME.textLight, marginTop: '4px' }}>Total Cost</div>
                  </div>
                  <div style={{
                    padding: '20px',
                    backgroundColor: `${THEME.accentBlue}10`,
                    borderRadius: '12px',
                    textAlign: 'center'
                  }}>
                    <div style={{ fontSize: '28px', fontWeight: '700', color: THEME.accentBlue }}>
                      {(usageData.totalInputTokens / 1000).toFixed(1)}K
                    </div>
                    <div style={{ fontSize: '12px', color: THEME.textLight, marginTop: '4px' }}>Input Tokens</div>
                  </div>
                  <div style={{
                    padding: '20px',
                    backgroundColor: `${THEME.accentYellow}10`,
                    borderRadius: '12px',
                    textAlign: 'center'
                  }}>
                    <div style={{ fontSize: '28px', fontWeight: '700', color: THEME.accentYellow }}>
                      {(usageData.totalOutputTokens / 1000).toFixed(1)}K
                    </div>
                    <div style={{ fontSize: '12px', color: THEME.textLight, marginTop: '4px' }}>Output Tokens</div>
                  </div>
                </div>

                {/* Period Summary */}
                {usageData.daily && (
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '16px', marginBottom: '24px' }}>
                    <div style={{ padding: '16px', backgroundColor: THEME.background, borderRadius: '8px' }}>
                      <div style={{ fontSize: '14px', fontWeight: '600', color: THEME.text }}>Daily</div>
                      <div style={{ fontSize: '12px', color: THEME.textLight }}>{usageData.daily.calls} calls · ${usageData.daily.cost}</div>
                    </div>
                    <div style={{ padding: '16px', backgroundColor: THEME.background, borderRadius: '8px' }}>
                      <div style={{ fontSize: '14px', fontWeight: '600', color: THEME.text }}>Weekly</div>
                      <div style={{ fontSize: '12px', color: THEME.textLight }}>{usageData.weekly.calls} calls · ${usageData.weekly.cost}</div>
                    </div>
                    <div style={{ padding: '16px', backgroundColor: THEME.background, borderRadius: '8px' }}>
                      <div style={{ fontSize: '14px', fontWeight: '600', color: THEME.text }}>Monthly</div>
                      <div style={{ fontSize: '12px', color: THEME.textLight }}>{usageData.monthly.calls} calls · ${usageData.monthly.cost}</div>
                    </div>
                  </div>
                )}

                {/* Recent Calls */}
                {usageData.recentCalls && usageData.recentCalls.length > 0 && (
                  <div style={{ marginBottom: '24px' }}>
                    <h3 style={{ fontSize: '16px', fontWeight: '600', color: THEME.text, marginBottom: '12px' }}>Recent Calls</h3>
                    <div style={{ maxHeight: '200px', overflow: 'auto' }}>
                      {usageData.recentCalls.slice(0, 10).map((call, idx) => (
                        <div key={idx} style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          padding: '10px 12px',
                          backgroundColor: idx % 2 === 0 ? THEME.background : 'transparent',
                          borderRadius: '6px',
                          fontSize: '13px'
                        }}>
                          <span style={{ color: THEME.textLight }}>{new Date(call.timestamp).toLocaleString()}</span>
                          <span style={{ color: THEME.text }}>{call.purpose || 'API Call'}</span>
                          <span style={{ color: THEME.accent, fontWeight: '500' }}>${call.costUSD}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Action Buttons */}
                <div style={{ display: 'flex', gap: '12px' }}>
                  <button
                    onClick={exportUsageCSV}
                    style={{
                      flex: 1,
                      padding: '12px',
                      backgroundColor: THEME.primary,
                      color: 'white',
                      border: 'none',
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px',
                      fontWeight: '500'
                    }}
                  >
                    📊 Export CSV
                  </button>
                  <button
                    onClick={fetchUsageData}
                    style={{
                      flex: 1,
                      padding: '12px',
                      backgroundColor: THEME.background,
                      color: THEME.text,
                      border: `1px solid ${THEME.border}`,
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px'
                    }}
                  >
                    🔄 Refresh
                  </button>
                  <button
                    onClick={resetUsageStats}
                    style={{
                      flex: 1,
                      padding: '12px',
                      backgroundColor: `${THEME.accentRed}10`,
                      color: THEME.accentRed,
                      border: 'none',
                      borderRadius: '8px',
                      cursor: 'pointer',
                      fontSize: '14px'
                    }}
                  >
                    🗑️ Reset
                  </button>
                </div>
              </>
            ) : (
              <div style={{ textAlign: 'center', padding: '40px', color: THEME.textLight }}>
                No usage data available
              </div>
            )}
          </div>
        </div>
      )}

      {/* Auto-Logout Warning Modal (v6) */}
      {showLogoutWarning && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.7)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999
        }}>
          <div style={{
            backgroundColor: 'white',
            padding: '32px',
            borderRadius: '16px',
            textAlign: 'center',
            maxWidth: '400px'
          }}>
            <div style={{ fontSize: '48px', marginBottom: '16px' }}>⏰</div>
            <h2 style={{ margin: '0 0 16px 0', color: THEME.text }}>Session Timeout Warning</h2>
            <p style={{ color: THEME.textLight, marginBottom: '24px' }}>
              You will be logged out in <strong style={{ color: THEME.accentRed, fontSize: '24px' }}>{logoutCountdown}</strong> seconds due to inactivity.
            </p>
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
              <button
                onClick={() => {
                  resetLogoutTimer();
                  setShowLogoutWarning(false);
                }}
                style={{
                  padding: '12px 24px',
                  backgroundColor: THEME.primary,
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '16px',
                  fontWeight: '600'
                }}
              >
                Stay Logged In
              </button>
              <button
                onClick={handleLogout}
                style={{
                  padding: '12px 24px',
                  backgroundColor: THEME.background,
                  color: THEME.text,
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '16px'
                }}
              >
                Logout Now
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Generated Content Modal */}
      {showGeneratedContent && generatedContent && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1500
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            maxWidth: '800px',
            width: '90%',
            maxHeight: '80vh',
            overflow: 'auto'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
              <h2 style={{ margin: 0, color: THEME.text }}>Generated Content</h2>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  onClick={downloadJSON}
                  style={{
                    padding: '8px 16px',
                    backgroundColor: THEME.primary,
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    cursor: 'pointer',
                    fontSize: '13px'
                  }}
                >
                  Download JSON
                </button>
                <button
                  onClick={() => setShowGeneratedContent(false)}
                  style={{
                    padding: '8px',
                    backgroundColor: 'transparent',
                    border: 'none',
                    cursor: 'pointer'
                  }}
                >
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke={THEME.textLight} strokeWidth="2">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                  </svg>
                </button>
              </div>
            </div>
            
            {generatedContent.usage && (
              <div style={{
                display: 'flex',
                gap: '16px',
                padding: '16px',
                backgroundColor: THEME.background,
                borderRadius: '8px',
                marginBottom: '20px'
              }}>
                <div>
                  <span style={{ fontSize: '12px', color: THEME.textLight }}>Input Tokens:</span>
                  <span style={{ marginLeft: '8px', fontWeight: '600' }}>{generatedContent.usage.inputTokens}</span>
                </div>
                <div>
                  <span style={{ fontSize: '12px', color: THEME.textLight }}>Output Tokens:</span>
                  <span style={{ marginLeft: '8px', fontWeight: '600' }}>{generatedContent.usage.outputTokens}</span>
                </div>
                <div>
                  <span style={{ fontSize: '12px', color: THEME.textLight }}>Cost:</span>
                  <span style={{ marginLeft: '8px', fontWeight: '600', color: THEME.accent }}>${generatedContent.usage.cost}</span>
                </div>
              </div>
            )}
            
            <pre style={{
              backgroundColor: THEME.background,
              padding: '20px',
              borderRadius: '8px',
              overflow: 'auto',
              fontSize: '13px',
              lineHeight: '1.5',
              maxHeight: '400px'
            }}>
              {JSON.stringify(generatedContent.content, null, 2)}
            </pre>
          </div>
        </div>
      )}

      {/* CSS Animations */}
      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        * {
          box-sizing: border-box;
        }
        html, body {
          margin: 0;
          padding: 0;
        }
        ::-webkit-scrollbar {
          width: 8px;
          height: 8px;
        }
        ::-webkit-scrollbar-track {
          background: ${THEME.background};
        }
        ::-webkit-scrollbar-thumb {
          background: ${THEME.border};
          border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
          background: ${THEME.textLight};
        }
      `}</style>
    </div>
  );
}
