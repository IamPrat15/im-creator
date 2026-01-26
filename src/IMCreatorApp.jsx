import React, { useState, useEffect } from 'react';
import { generateIM, generatePPTX, saveDraft, checkHealth, downloadBase64File } from './api';

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
          { value: 'management-presentation', label: 'Management Presentation' },
          { value: 'cim', label: 'Confidential Information Memorandum' },
          { value: 'teaser', label: 'Teaser Document' }
        ], order: 3 },
        { id: 'advisor', type: 'text', label: 'Sell-Side Advisor', defaultValue: 'RMB Securities', order: 4 },
        { id: 'presentationDate', type: 'date', label: 'Presentation Date', required: true, order: 5 },
        { id: 'targetBuyerType', type: 'multiselect', label: 'Target Buyer Type', required: true, options: [
          { value: 'strategic', label: 'Strategic Buyer' },
          { value: 'financial', label: 'Financial Investor' },
          { value: 'international', label: 'International Acquirer' }
        ], helpText: 'Content will be tailored for selected buyers', order: 6 }
      ]
    },
    {
      id: 'company-fundamentals',
      name: 'Company Overview',
      icon: '🏢',
      description: 'Basic company information',
      questions: [
        { id: 'foundedYear', type: 'number', label: 'Founded Year', required: true, validation: { min: 1900, max: 2026 }, order: 1 },
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
        { id: 'founderEducation', type: 'textarea', label: 'Education', placeholder: 'MBA - JBIMS\nB.E. - VJTI', order: 4 },
        { id: 'previousCompanies', type: 'textarea', label: 'Previous Companies', placeholder: 'Company | Role | Duration', helpText: 'Notable prior experience', order: 5 },
        { id: 'leadershipTeam', type: 'textarea', label: 'Leadership Team', placeholder: 'Name | Title | Department', helpText: 'Key management team members', order: 6 }
      ]
    },
    {
      id: 'services-products',
      name: 'Services & Products',
      icon: '⚙️',
      description: 'Offerings & capabilities',
      questions: [
        { id: 'serviceLines', type: 'textarea', label: 'Service Lines', placeholder: 'Cloud & Automation | 39% | AWS migration, DevOps, Infrastructure', required: true, helpText: 'Format: Name | Revenue % | Description', order: 1 },
        { id: 'products', type: 'textarea', label: 'Proprietary Products', placeholder: 'AI Agent Studio | Platform for AI agents | 500+ templates', helpText: 'Format: Name | Description | Key metric', order: 2 },
        { id: 'techPartnerships', type: 'textarea', label: 'Technology Partnerships', placeholder: 'AWS Advanced Tier Partner\nDatabricks Partner', order: 3 },
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
          { value: 'bfsi', label: 'BFSI' },
          { value: 'healthcare', label: 'Healthcare' },
          { value: 'retail', label: 'Retail' },
          { value: 'manufacturing', label: 'Manufacturing' },
          { value: 'technology', label: 'Technology' },
          { value: 'media', label: 'Media & Entertainment' }
        ], order: 1 },
        { id: 'primaryVerticalPct', type: 'number', label: 'Primary Vertical Revenue %', required: true, order: 2 },
        { id: 'otherVerticals', type: 'textarea', label: 'Other Verticals', placeholder: 'FinTech | 14%\nMedia | 11%', order: 3 },
        { id: 'topClients', type: 'textarea', label: 'Top Clients', placeholder: 'Axis Bank | BFSI | 2015\nHDFC Bank | BFSI | 2018', required: true, helpText: 'Format: Client Name | Vertical | Year Started', order: 4 },
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
        { id: 'currency', type: 'select', label: 'Currency', options: [{ value: 'INR', label: 'INR (₹)' }, { value: 'USD', label: 'USD ($)' }], defaultValue: 'INR', order: 1 },
        { id: 'revenueFY24', type: 'number', label: 'Revenue FY24 (Cr)', required: true, order: 2 },
        { id: 'revenueFY25', type: 'number', label: 'Revenue FY25 (Cr)', required: true, order: 3 },
        { id: 'revenueFY26P', type: 'number', label: 'Revenue FY26P (Cr)', required: true, helpText: 'P = Projected', order: 4 },
        { id: 'revenueFY27P', type: 'number', label: 'Revenue FY27P (Cr)', order: 5 },
        { id: 'revenueFY28P', type: 'number', label: 'Revenue FY28P (Cr)', order: 6 },
        { id: 'ebitdaMarginFY25', type: 'number', label: 'EBITDA Margin FY25 %', required: true, order: 7 },
        { id: 'revenueByService', type: 'textarea', label: 'Revenue by Service', placeholder: 'Cloud & Automation | 39%\nManaged Services | 31%', order: 8 }
      ]
    },
    {
      id: 'case-studies',
      name: 'Case Studies',
      icon: '📚',
      description: 'Client success stories',
      questions: [
        { id: 'cs1Client', type: 'text', label: 'Case Study 1 - Client Name', helpText: 'Featured client success story', order: 1 },
        { id: 'cs1Challenge', type: 'textarea', label: 'Challenge', placeholder: 'Describe the business challenge faced by the client', order: 2 },
        { id: 'cs1Solution', type: 'textarea', label: 'Solution', placeholder: 'How your company solved the problem', order: 3 },
        { id: 'cs1Results', type: 'textarea', label: 'Results', placeholder: '40% reduction in processing time\n60% cost savings\nImproved customer satisfaction', helpText: 'Quantified outcomes (one per line)', order: 4 },
        { id: 'cs2Client', type: 'text', label: 'Case Study 2 - Client Name', order: 5 },
        { id: 'cs2Challenge', type: 'textarea', label: 'Challenge', order: 6 },
        { id: 'cs2Solution', type: 'textarea', label: 'Solution', order: 7 },
        { id: 'cs2Results', type: 'textarea', label: 'Results', order: 8 }
      ]
    },
    {
      id: 'growth-strategy',
      name: 'Growth Strategy',
      icon: '🎯',
      description: 'Future plans & competitive position',
      questions: [
        { id: 'growthDrivers', type: 'textarea', label: 'Key Growth Drivers', required: true, placeholder: 'AI adoption in enterprise\nCloud migration demand\nExpanding BFSI relationships', order: 1 },
        { id: 'competitiveAdvantages', type: 'textarea', label: 'Competitive Advantages', required: true, helpText: 'Minimum 5 advantages (one per line)', placeholder: 'Deep AWS expertise | Only 8 companies in India with this certification\nStrong BFSI relationships | 10+ year partnerships with leading banks', order: 2 },
        { id: 'shortTermGoals', type: 'textarea', label: 'Short-Term Goals (0-12 months)', placeholder: 'Launch AI Practice\nExpand Bangalore team', order: 3 },
        { id: 'mediumTermGoals', type: 'textarea', label: 'Medium-Term Goals (1-3 years)', placeholder: 'Enter international markets\nAchieve ₹500 Cr revenue', order: 4 },
        { id: 'synergiesStrategic', type: 'textarea', label: 'Synergies for Strategic Buyers', placeholder: 'Access to BFSI client base\nAWS competencies enhancement\nIndian market expansion', order: 5 },
        { id: 'synergiesFinancial', type: 'textarea', label: 'Synergies for Financial Investors', placeholder: 'Strong EBITDA margins\nCapital-light model\nHigh revenue visibility', order: 6 }
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
          { value: 'market', label: 'Market Position - Focus on competitive advantages, market share' },
          { value: 'synergy', label: 'Synergy Focus - Emphasize acquisition benefits for buyers' }
        ], helpText: 'Select which content variants to include in the presentation', order: 1 },
        { id: 'templateStyle', type: 'select', label: 'Presentation Template', required: true, options: [
          { value: 'modern-tech', label: 'Modern Tech (Blue/Green)' },
          { value: 'conservative', label: 'Conservative Banking (Navy/Gold)' },
          { value: 'minimalist', label: 'Minimalist (Black/White)' },
          { value: 'acc-brand', label: 'ACC Brand (Burgundy/Maroon)' }
        ], defaultValue: 'modern-tech', helpText: 'Professional template matching your brand', order: 2 },
        { id: 'includeAppendix', type: 'multiselect', label: 'Include in Appendix', options: [
          { value: 'team-bios', label: 'Detailed Team Bios' },
          { value: 'client-list', label: 'Full Client List' },
          { value: 'financial-detail', label: 'Detailed Financial Statements' },
          { value: 'case-studies-extra', label: 'Additional Case Studies' }
        ], order: 3 },
        { id: 'exportFormat', type: 'multiselect', label: 'Export Formats', options: [
          { value: 'pptx', label: 'PowerPoint (.pptx) - Editable' },
          { value: 'pdf', label: 'PDF - For distribution' },
          { value: 'json', label: 'JSON Data - For integration' }
        ], defaultValue: ['pptx'], order: 4 }
      ]
    }
  ]
};

// ACC Brand Colors (from Agentic Underwriting screenshots)
const THEME = {
  primary: '#7C1034',        // Burgundy/Maroon
  primaryDark: '#5A0C26',    // Darker burgundy
  primaryLight: '#9A1842',   // Lighter burgundy
  secondary: '#2D3748',      // Dark gray
  accent: '#48BB78',         // Green (for success states)
  accentYellow: '#ECC94B',   // Yellow/Gold
  accentRed: '#E53E3E',      // Red (for errors)
  accentBlue: '#4299E1',     // Blue
  background: '#F7FAFC',     // Light gray background
  surface: '#FFFFFF',        // White
  text: '#1A202C',           // Dark text
  textLight: '#718096',      // Gray text
  border: '#E2E8F0',         // Border color
};

export default function IMCreatorApp() {
  const [questionnaire, setQuestionnaire] = useState(defaultQuestionnaire);
  const [currentPhase, setCurrentPhase] = useState(0);
  const [formData, setFormData] = useState({});
  const [completedPhases, setCompletedPhases] = useState([]);
  const [errors, setErrors] = useState({});
  const [showConfig, setShowConfig] = useState(false);
  const [showAddQ, setShowAddQ] = useState(false);
  const [newQ, setNewQ] = useState({ type: 'text', label: '', required: false, placeholder: '' });
  const [showReport, setShowReport] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isGeneratingPPTX, setIsGeneratingPPTX] = useState(false);
  const [generatedContent, setGeneratedContent] = useState(null);
  const [showGeneratedContent, setShowGeneratedContent] = useState(false);
  const [apiStatus, setApiStatus] = useState('checking');
  const [notification, setNotification] = useState(null);

  const phase = questionnaire.phases[currentPhase];
  const questions = phase.questions.filter(q => !q.isHidden).sort((a, b) => a.order - b.order);
  const progress = Math.round((completedPhases.length / questionnaire.phases.length) * 100);

  useEffect(() => {
    checkHealth()
      .then(() => setApiStatus('connected'))
      .catch(() => setApiStatus('disconnected'));
  }, []);

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
      p.questions.forEach(q => {
        if (q.required && !formData[q.id]) report.errors.push({ phase: p.name, field: q.label, msg: 'Required field missing' });
      });
    });
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

  const handleNext = () => {
    if (validate()) {
      if (!completedPhases.includes(currentPhase)) setCompletedPhases([...completedPhases, currentPhase]);
      setCurrentPhase(Math.min(currentPhase + 1, questionnaire.phases.length - 1));
    }
  };

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
      const theme = formData.templateStyle || 'modern-tech';
      const result = await generatePPTX(formData, theme);
      
      if (result.success && result.fileData) {
        downloadBase64File(result.fileData, result.filename, result.mimeType);
        showNotification(`PowerPoint downloaded! (${result.slideCount || 13} slides, ${theme} theme)`, 'success');
      } else {
        throw new Error('Invalid response from server');
      }
    } catch (error) {
      showNotification('Failed to generate PowerPoint. Please try again.', 'error');
      console.error('PPTX Generation error:', error);
    } finally {
      setIsGeneratingPPTX(false);
    }
  };

  const downloadJSON = () => {
    if (!generatedContent) return;
    const blob = new Blob([JSON.stringify(generatedContent.content, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${formData.projectCodename || 'IM'}_content.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

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

  const report = fullValidate();

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: THEME.background,
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
      display: 'flex',
      flexDirection: 'column',
      overflow: 'hidden'
    }}>
      {/* Notification Toast */}
      {notification && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '20px',
          padding: '16px 24px',
          borderRadius: '8px',
          backgroundColor: notification.type === 'success' ? THEME.accent : notification.type === 'error' ? THEME.accentRed : THEME.primary,
          color: 'white',
          fontSize: '14px',
          fontWeight: '500',
          boxShadow: '0 4px 20px rgba(0,0,0,0.15)',
          zIndex: 9999,
          display: 'flex',
          alignItems: 'center',
          gap: '10px'
        }}>
          <span>{notification.type === 'success' ? '✓' : notification.type === 'error' ? '✕' : 'ℹ'}</span>
          {notification.message}
        </div>
      )}

      {/* Top Header Bar - ACC Style */}
      <header style={{
        backgroundColor: THEME.surface,
        borderBottom: `1px solid ${THEME.border}`,
        padding: '0 24px',
        height: '64px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexShrink: 0
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          {/* Logo */}
          <div style={{
            width: '40px',
            height: '40px',
            borderRadius: '8px',
            background: `linear-gradient(135deg, ${THEME.primary} 0%, ${THEME.primaryDark} 100%)`,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
              <polyline points="10 9 9 9 8 9"/>
            </svg>
          </div>
          <div>
            <h1 style={{ margin: 0, fontSize: '18px', fontWeight: '600', color: THEME.text }}>
              IM Creator Pro
            </h1>
            <span style={{ fontSize: '12px', color: THEME.textLight }}>
              Professional Information Memorandum Generator
            </span>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
          {/* Progress */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span style={{ fontSize: '13px', color: THEME.textLight, fontWeight: '500' }}>{progress}% Complete</span>
            <div style={{
              width: '150px',
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
          </div>

          {/* API Status Badge */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            padding: '6px 14px',
            backgroundColor: apiStatus === 'connected' ? `${THEME.accent}15` : `${THEME.accentRed}15`,
            borderRadius: '20px',
            border: `1px solid ${apiStatus === 'connected' ? THEME.accent : THEME.accentRed}`
          }}>
            <div style={{
              width: '8px',
              height: '8px',
              borderRadius: '50%',
              backgroundColor: apiStatus === 'connected' ? THEME.accent : THEME.accentRed
            }} />
            <span style={{ fontSize: '13px', fontWeight: '500', color: apiStatus === 'connected' ? THEME.accent : THEME.accentRed }}>
              {apiStatus === 'connected' ? 'API Connected' : 'API Disconnected'}
            </span>
          </div>

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
              <circle cx="12" cy="12" r="3"/>
              <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
            </svg>
            Configure
          </button>
        </div>
      </header>

      {/* Main Content Area - FULL WIDTH */}
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
          <div style={{
            padding: '20px',
            borderBottom: `1px solid ${THEME.border}`
          }}>
            <span style={{ fontWeight: '600', color: THEME.text, fontSize: '14px' }}>Sections</span>
          </div>

          <nav style={{ flex: 1, overflowY: 'auto', padding: '12px' }}>
            {questionnaire.phases.map((p, idx) => {
              const isActive = idx === currentPhase;
              const isCompleted = completedPhases.includes(idx);
              
              return (
                <button
                  key={p.id}
                  onClick={() => setCurrentPhase(idx)}
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
                    textAlign: 'left',
                    transition: 'all 0.15s ease'
                  }}
                >
                  <span style={{
                    width: '32px',
                    height: '32px',
                    borderRadius: '8px',
                    backgroundColor: isCompleted ? THEME.accent : isActive ? THEME.primary : THEME.background,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontSize: '14px',
                    color: isCompleted || isActive ? 'white' : THEME.textLight,
                    flexShrink: 0
                  }}>
                    {isCompleted ? '✓' : p.icon}
                  </span>
                  <div style={{ flex: 1, overflow: 'hidden' }}>
                    <div style={{
                      fontSize: '14px',
                      fontWeight: isActive ? '600' : '500',
                      color: isActive ? THEME.primary : THEME.text,
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }}>
                      {p.name}
                    </div>
                    <div style={{
                      fontSize: '12px',
                      color: THEME.textLight,
                      marginTop: '2px'
                    }}>
                      {p.description}
                    </div>
                  </div>
                </button>
              );
            })}
          </nav>

          {/* Save Draft Button */}
          <div style={{ padding: '16px', borderTop: `1px solid ${THEME.border}` }}>
            <button
              onClick={handleSaveDraft}
              style={{
                width: '100%',
                padding: '12px',
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
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                <polyline points="17 21 17 13 7 13 7 21"/>
                <polyline points="7 3 7 8 15 8"/>
              </svg>
              Save Draft
            </button>
          </div>
        </aside>

        {/* Main Form Area - FILLS REMAINING SPACE */}
        <main style={{ 
          flex: 1, 
          overflow: 'auto', 
          padding: '32px 48px',
          backgroundColor: THEME.background
        }}>
          {/* Phase Header */}
          <div style={{ marginBottom: '32px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '8px' }}>
              <span style={{ 
                fontSize: '32px',
                width: '56px',
                height: '56px',
                backgroundColor: `${THEME.primary}15`,
                borderRadius: '12px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>{phase.icon}</span>
              <div>
                <h2 style={{ margin: 0, fontSize: '28px', fontWeight: '600', color: THEME.text }}>
                  {phase.name}
                </h2>
                <p style={{ margin: '4px 0 0', color: THEME.textLight, fontSize: '15px' }}>
                  {phase.description}
                </p>
              </div>
            </div>

            {/* Config Mode Banner */}
            {showConfig && (
              <div style={{
                marginTop: '20px',
                padding: '16px 20px',
                backgroundColor: `${THEME.accentYellow}20`,
                borderRadius: '8px',
                border: `1px solid ${THEME.accentYellow}`,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between'
              }}>
                <span style={{ color: '#744210', fontWeight: '500', fontSize: '14px' }}>
                  ⚙️ Configuration Mode Active - You can hide/show or add questions
                </span>
                <button
                  onClick={() => setShowAddQ(true)}
                  style={{
                    padding: '8px 16px',
                    backgroundColor: THEME.primary,
                    color: 'white',
                    border: 'none',
                    borderRadius: '6px',
                    cursor: 'pointer',
                    fontWeight: '500',
                    fontSize: '13px'
                  }}
                >
                  + Add Question
                </button>
              </div>
            )}
          </div>

          {/* Questions */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {questions.map(q => (
              <div
                key={q.id}
                style={{
                  backgroundColor: THEME.surface,
                  borderRadius: '12px',
                  padding: '24px',
                  border: `1px solid ${THEME.border}`
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
                  <div>
                    <label style={{
                      display: 'block',
                      fontSize: '14px',
                      fontWeight: '600',
                      color: THEME.text,
                      marginBottom: '4px'
                    }}>
                      {q.label}
                      {q.required && <span style={{ color: THEME.accentRed, marginLeft: '4px' }}>*</span>}
                    </label>
                    {q.helpText && (
                      <p style={{ margin: 0, fontSize: '13px', color: THEME.textLight }}>
                        {q.helpText}
                      </p>
                    )}
                  </div>
                  
                  {showConfig && (
                    <div style={{ display: 'flex', gap: '8px' }}>
                      <button
                        onClick={() => toggleHide(q.id)}
                        style={{
                          padding: '6px 12px',
                          backgroundColor: THEME.background,
                          border: `1px solid ${THEME.border}`,
                          borderRadius: '6px',
                          cursor: 'pointer',
                          fontSize: '12px',
                          color: THEME.textLight
                        }}
                      >
                        Hide
                      </button>
                      {q.isCustom && (
                        <button
                          onClick={() => removeQ(q.id)}
                          style={{
                            padding: '6px 12px',
                            backgroundColor: `${THEME.accentRed}10`,
                            border: `1px solid ${THEME.accentRed}30`,
                            borderRadius: '6px',
                            cursor: 'pointer',
                            fontSize: '12px',
                            color: THEME.accentRed
                          }}
                        >
                          Remove
                        </button>
                      )}
                    </div>
                  )}
                </div>
                
                {renderField(q)}
                
                {errors[q.id] && (
                  <p style={{ margin: '8px 0 0', fontSize: '13px', color: THEME.accentRed, fontWeight: '500' }}>
                    ⚠️ {errors[q.id]}
                  </p>
                )}
              </div>
            ))}
          </div>

          {/* Navigation Buttons */}
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginTop: '40px',
            paddingTop: '24px',
            borderTop: `1px solid ${THEME.border}`
          }}>
            <button
              onClick={() => setCurrentPhase(Math.max(0, currentPhase - 1))}
              disabled={currentPhase === 0}
              style={{
                padding: '12px 24px',
                backgroundColor: THEME.surface,
                color: currentPhase === 0 ? THEME.textLight : THEME.text,
                border: `1px solid ${THEME.border}`,
                borderRadius: '8px',
                cursor: currentPhase === 0 ? 'not-allowed' : 'pointer',
                fontWeight: '500',
                fontSize: '14px',
                opacity: currentPhase === 0 ? 0.5 : 1
              }}
            >
              ← Previous
            </button>

            <div style={{ display: 'flex', gap: '12px' }}>
              {currentPhase === questionnaire.phases.length - 1 ? (
                <>
                  <button
                    onClick={handleGeneratePPTX}
                    disabled={isGeneratingPPTX}
                    style={{
                      padding: '12px 28px',
                      backgroundColor: isGeneratingPPTX ? THEME.textLight : THEME.primary,
                      color: 'white',
                      border: 'none',
                      borderRadius: '8px',
                      cursor: isGeneratingPPTX ? 'not-allowed' : 'pointer',
                      fontWeight: '600',
                      fontSize: '14px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px'
                    }}
                  >
                    {isGeneratingPPTX ? (
                      <>
                        <span style={{
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
                      <>📊 Download PPTX</>
                    )}
                  </button>
                  <button
                    onClick={handleGenerateJSON}
                    disabled={isGenerating}
                    style={{
                      padding: '12px 28px',
                      backgroundColor: isGenerating ? THEME.textLight : THEME.secondary,
                      color: 'white',
                      border: 'none',
                      borderRadius: '8px',
                      cursor: isGenerating ? 'not-allowed' : 'pointer',
                      fontWeight: '600',
                      fontSize: '14px'
                    }}
                  >
                    {isGenerating ? 'Generating...' : '🤖 Generate JSON'}
                  </button>
                </>
              ) : (
                <button
                  onClick={handleNext}
                  style={{
                    padding: '12px 28px',
                    backgroundColor: THEME.primary,
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    cursor: 'pointer',
                    fontWeight: '600',
                    fontSize: '14px'
                  }}
                >
                  Next →
                </button>
              )}
            </div>
          </div>
        </main>

        {/* Right Panel - Validation Report */}
        {showReport && (
          <aside style={{
            width: '350px',
            backgroundColor: THEME.surface,
            borderLeft: `1px solid ${THEME.border}`,
            padding: '24px',
            overflowY: 'auto',
            flexShrink: 0
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '600', color: THEME.text }}>
                Validation Report
              </h3>
              <button
                onClick={() => setShowReport(false)}
                style={{
                  padding: '6px 10px',
                  backgroundColor: THEME.background,
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  color: THEME.textLight
                }}
              >
                ✕
              </button>
            </div>

            {report.errors.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '13px', fontWeight: '600', color: THEME.accentRed }}>
                  Errors ({report.errors.length})
                </h4>
                {report.errors.map((e, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: `${THEME.accentRed}10`,
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: `3px solid ${THEME.accentRed}`
                  }}>
                    <div style={{ fontSize: '11px', color: THEME.accentRed, fontWeight: '600' }}>{e.phase}</div>
                    <div style={{ fontSize: '13px', color: THEME.text }}>{e.field}: {e.msg}</div>
                  </div>
                ))}
              </div>
            )}

            {report.warnings.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '13px', fontWeight: '600', color: THEME.accentYellow }}>
                  Warnings ({report.warnings.length})
                </h4>
                {report.warnings.map((w, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: `${THEME.accentYellow}15`,
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: `3px solid ${THEME.accentYellow}`
                  }}>
                    <div style={{ fontSize: '11px', color: '#744210', fontWeight: '600' }}>{w.phase}</div>
                    <div style={{ fontSize: '13px', color: THEME.text }}>{w.msg}</div>
                  </div>
                ))}
              </div>
            )}

            {report.suggestions.length > 0 && (
              <div>
                <h4 style={{ margin: '0 0 12px', fontSize: '13px', fontWeight: '600', color: THEME.accentBlue }}>
                  Suggestions ({report.suggestions.length})
                </h4>
                {report.suggestions.map((s, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: `${THEME.accentBlue}10`,
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: `3px solid ${THEME.accentBlue}`
                  }}>
                    <div style={{ fontSize: '11px', color: THEME.accentBlue, fontWeight: '600' }}>{s.phase}</div>
                    <div style={{ fontSize: '13px', color: THEME.text }}>{s.msg}</div>
                  </div>
                ))}
              </div>
            )}
          </aside>
        )}
      </div>

      {/* Add Question Modal */}
      {showAddQ && (
        <div style={{
          position: 'fixed',
          inset: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            width: '90%',
            maxWidth: '480px',
            boxShadow: '0 20px 50px rgba(0,0,0,0.2)'
          }}>
            <h3 style={{ margin: '0 0 24px', fontSize: '18px', fontWeight: '600', color: THEME.text }}>
              Add Custom Question
            </h3>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div>
                <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500', color: THEME.text, fontSize: '14px' }}>
                  Question Label
                </label>
                <input
                  type="text"
                  value={newQ.label}
                  onChange={e => setNewQ({ ...newQ, label: e.target.value })}
                  placeholder="Enter question label"
                  style={{
                    width: '100%',
                    padding: '12px 16px',
                    border: `1px solid ${THEME.border}`,
                    borderRadius: '8px',
                    fontSize: '14px',
                    boxSizing: 'border-box'
                  }}
                />
              </div>
              
              <div>
                <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500', color: THEME.text, fontSize: '14px' }}>
                  Field Type
                </label>
                <select
                  value={newQ.type}
                  onChange={e => setNewQ({ ...newQ, type: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '12px 16px',
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
              
              <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={newQ.required}
                  onChange={e => setNewQ({ ...newQ, required: e.target.checked })}
                  style={{ width: '18px', height: '18px', accentColor: THEME.primary }}
                />
                <span style={{ fontWeight: '500', color: THEME.text, fontSize: '14px' }}>Required field</span>
              </label>
            </div>
            
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '32px' }}>
              <button
                onClick={() => setShowAddQ(false)}
                style={{
                  padding: '10px 20px',
                  backgroundColor: THEME.background,
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontWeight: '500',
                  fontSize: '14px'
                }}
              >
                Cancel
              </button>
              <button
                onClick={addQuestion}
                style={{
                  padding: '10px 20px',
                  backgroundColor: THEME.primary,
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontWeight: '500',
                  fontSize: '14px'
                }}
              >
                Add Question
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Generated Content Modal */}
      {showGeneratedContent && generatedContent && (
        <div style={{
          position: 'fixed',
          inset: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000
        }}>
          <div style={{
            backgroundColor: THEME.surface,
            borderRadius: '16px',
            padding: '32px',
            width: '90%',
            maxWidth: '800px',
            maxHeight: '80vh',
            overflow: 'auto',
            boxShadow: '0 20px 50px rgba(0,0,0,0.2)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: '600', color: THEME.text }}>
                Generated IM Content
              </h3>
              <button
                onClick={() => setShowGeneratedContent(false)}
                style={{
                  padding: '8px 12px',
                  backgroundColor: THEME.background,
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer'
                }}
              >
                ✕ Close
              </button>
            </div>
            
            <pre style={{
              backgroundColor: THEME.background,
              padding: '20px',
              borderRadius: '8px',
              overflow: 'auto',
              fontSize: '12px',
              lineHeight: '1.5',
              border: `1px solid ${THEME.border}`
            }}>
              {JSON.stringify(generatedContent.content, null, 2)}
            </pre>
            
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '24px' }}>
              <button
                onClick={downloadJSON}
                style={{
                  padding: '10px 20px',
                  backgroundColor: THEME.primary,
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontWeight: '500'
                }}
              >
                📥 Download JSON
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CSS */}
      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        * {
          box-sizing: border-box;
        }
        html, body, #root {
          margin: 0;
          padding: 0;
          width: 100%;
          height: 100%;
          overflow: hidden;
        }
      `}</style>
    </div>
  );
}
