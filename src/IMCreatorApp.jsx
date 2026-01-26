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
          { value: 'modern-tech', label: 'Modern Tech (Blue/Green - Deloitte Style)' },
          { value: 'conservative', label: 'Conservative Banking (Navy/Gold)' },
          { value: 'minimalist', label: 'Minimalist (Black/White)' }
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
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

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
      padding: '12px 14px',
      border: `2px solid ${err ? '#ef4444' : '#e2e8f0'}`,
      borderRadius: '8px',
      fontSize: '14px',
      outline: 'none',
      transition: 'all 0.2s ease',
      backgroundColor: '#fff',
      color: '#1e293b'
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
            onFocus={e => e.target.style.borderColor = '#3b82f6'}
            onBlur={e => e.target.style.borderColor = err ? '#ef4444' : '#e2e8f0'}
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
                gap: '10px',
                padding: '10px 12px',
                backgroundColor: (val || []).includes(o.value) ? '#eff6ff' : '#f8fafc',
                borderRadius: '8px',
                cursor: 'pointer',
                border: `2px solid ${(val || []).includes(o.value) ? '#3b82f6' : '#e2e8f0'}`,
                transition: 'all 0.2s ease'
              }}>
                <input
                  type="checkbox"
                  checked={(val || []).includes(o.value)}
                  onChange={e => {
                    const arr = val || [];
                    updateField(q.id, e.target.checked ? [...arr, o.value] : arr.filter(v => v !== o.value));
                  }}
                  style={{ marginTop: '2px', width: '18px', height: '18px', accentColor: '#3b82f6' }}
                />
                <span style={{ fontSize: '14px', color: '#1e293b' }}>{o.label}</span>
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
            onFocus={e => e.target.style.borderColor = '#3b82f6'}
            onBlur={e => e.target.style.borderColor = err ? '#ef4444' : '#e2e8f0'}
          />
        );
      case 'date':
        return (
          <input
            type="date"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            style={{ ...baseInputStyle, cursor: 'pointer' }}
            onFocus={e => e.target.style.borderColor = '#3b82f6'}
            onBlur={e => e.target.style.borderColor = err ? '#ef4444' : '#e2e8f0'}
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
            onFocus={e => e.target.style.borderColor = '#3b82f6'}
            onBlur={e => e.target.style.borderColor = err ? '#ef4444' : '#e2e8f0'}
          />
        );
    }
  };

  const report = fullValidate();

  return (
    <div style={{
      minHeight: '100vh',
      width: '100%',
      backgroundColor: '#f1f5f9',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
      display: 'flex',
      flexDirection: 'column'
    }}>
      {/* Notification Toast */}
      {notification && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '20px',
          padding: '16px 24px',
          borderRadius: '12px',
          backgroundColor: notification.type === 'success' ? '#10b981' : notification.type === 'error' ? '#ef4444' : '#3b82f6',
          color: 'white',
          fontSize: '14px',
          fontWeight: '500',
          boxShadow: '0 10px 25px rgba(0,0,0,0.15)',
          zIndex: 9999,
          display: 'flex',
          alignItems: 'center',
          gap: '10px',
          animation: 'slideIn 0.3s ease'
        }}>
          <span>{notification.type === 'success' ? '✓' : notification.type === 'error' ? '✕' : 'ℹ'}</span>
          {notification.message}
        </div>
      )}

      {/* Top Header Bar */}
      <header style={{
        backgroundColor: '#1e293b',
        color: 'white',
        padding: '0 24px',
        height: '64px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
        position: 'sticky',
        top: 0,
        zIndex: 100
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <div style={{
            width: '40px',
            height: '40px',
            borderRadius: '10px',
            background: 'linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '20px'
          }}>
            📊
          </div>
          <div>
            <h1 style={{ margin: 0, fontSize: '18px', fontWeight: '700', letterSpacing: '-0.5px' }}>
              IM Creator Pro
            </h1>
            <span style={{ fontSize: '12px', opacity: 0.7 }}>
              Professional Information Memorandum Generator
            </span>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          {/* Progress Indicator */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span style={{ fontSize: '13px', opacity: 0.8 }}>{progress}% Complete</span>
            <div style={{
              width: '120px',
              height: '8px',
              backgroundColor: 'rgba(255,255,255,0.2)',
              borderRadius: '4px',
              overflow: 'hidden'
            }}>
              <div style={{
                width: `${progress}%`,
                height: '100%',
                background: 'linear-gradient(90deg, #10b981, #34d399)',
                borderRadius: '4px',
                transition: 'width 0.3s ease'
              }} />
            </div>
          </div>

          {/* API Status */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '6px',
            padding: '6px 12px',
            backgroundColor: apiStatus === 'connected' ? 'rgba(16, 185, 129, 0.2)' : 'rgba(239, 68, 68, 0.2)',
            borderRadius: '20px',
            fontSize: '12px'
          }}>
            <div style={{
              width: '8px',
              height: '8px',
              borderRadius: '50%',
              backgroundColor: apiStatus === 'connected' ? '#10b981' : '#ef4444'
            }} />
            {apiStatus === 'connected' ? 'API Connected' : 'API Disconnected'}
          </div>

          {/* Config Button */}
          <button
            onClick={() => setShowConfig(!showConfig)}
            style={{
              padding: '8px 16px',
              backgroundColor: showConfig ? '#3b82f6' : 'rgba(255,255,255,0.1)',
              color: 'white',
              border: 'none',
              borderRadius: '8px',
              cursor: 'pointer',
              fontSize: '13px',
              fontWeight: '500',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            ⚙️ {showConfig ? 'Hide Config' : 'Configure'}
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        {/* Left Sidebar - Phase Navigation */}
        <aside style={{
          width: sidebarCollapsed ? '70px' : '280px',
          backgroundColor: '#fff',
          borderRight: '1px solid #e2e8f0',
          display: 'flex',
          flexDirection: 'column',
          transition: 'width 0.3s ease',
          overflow: 'hidden'
        }}>
          <div style={{
            padding: '16px',
            borderBottom: '1px solid #e2e8f0',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center'
          }}>
            {!sidebarCollapsed && <span style={{ fontWeight: '600', color: '#1e293b' }}>Sections</span>}
            <button
              onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
              style={{
                padding: '8px',
                backgroundColor: '#f1f5f9',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontSize: '14px'
              }}
            >
              {sidebarCollapsed ? '→' : '←'}
            </button>
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
                    padding: sidebarCollapsed ? '12px' : '14px 16px',
                    marginBottom: '6px',
                    backgroundColor: isActive ? '#eff6ff' : 'transparent',
                    border: isActive ? '2px solid #3b82f6' : '2px solid transparent',
                    borderRadius: '10px',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '12px',
                    textAlign: 'left',
                    transition: 'all 0.2s ease',
                    justifyContent: sidebarCollapsed ? 'center' : 'flex-start'
                  }}
                  onMouseEnter={e => { if (!isActive) e.target.style.backgroundColor = '#f8fafc'; }}
                  onMouseLeave={e => { if (!isActive) e.target.style.backgroundColor = 'transparent'; }}
                >
                  <span style={{
                    fontSize: '20px',
                    filter: isCompleted ? 'none' : 'grayscale(50%)'
                  }}>
                    {isCompleted ? '✅' : p.icon}
                  </span>
                  {!sidebarCollapsed && (
                    <div style={{ flex: 1, overflow: 'hidden' }}>
                      <div style={{
                        fontSize: '14px',
                        fontWeight: isActive ? '600' : '500',
                        color: isActive ? '#1e40af' : '#475569',
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis'
                      }}>
                        {p.name}
                      </div>
                      <div style={{
                        fontSize: '11px',
                        color: '#94a3b8',
                        marginTop: '2px'
                      }}>
                        {p.description}
                      </div>
                    </div>
                  )}
                </button>
              );
            })}
          </nav>

          {/* Sidebar Footer Actions */}
          {!sidebarCollapsed && (
            <div style={{ padding: '16px', borderTop: '1px solid #e2e8f0' }}>
              <button
                onClick={handleSaveDraft}
                style={{
                  width: '100%',
                  padding: '12px',
                  backgroundColor: '#f1f5f9',
                  border: '1px solid #e2e8f0',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontSize: '13px',
                  fontWeight: '500',
                  color: '#475569',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '8px'
                }}
              >
                💾 Save Draft
              </button>
            </div>
          )}
        </aside>

        {/* Main Form Area */}
        <main style={{ flex: 1, overflow: 'auto', padding: '32px', backgroundColor: '#f8fafc' }}>
          <div style={{ maxWidth: '900px', margin: '0 auto' }}>
            {/* Phase Header */}
            <div style={{ marginBottom: '32px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '16px', marginBottom: '8px' }}>
                <span style={{ fontSize: '36px' }}>{phase.icon}</span>
                <div>
                  <h2 style={{ margin: 0, fontSize: '28px', fontWeight: '700', color: '#1e293b' }}>
                    {phase.name}
                  </h2>
                  <p style={{ margin: '4px 0 0', color: '#64748b', fontSize: '15px' }}>
                    {phase.description}
                  </p>
                </div>
              </div>

              {/* Config Mode Toggle for current phase */}
              {showConfig && (
                <div style={{
                  marginTop: '16px',
                  padding: '16px',
                  backgroundColor: '#fef3c7',
                  borderRadius: '10px',
                  border: '1px solid #fcd34d',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between'
                }}>
                  <span style={{ color: '#92400e', fontWeight: '500' }}>
                    ⚠️ Configuration Mode Active - You can hide/show or add questions
                  </span>
                  <button
                    onClick={() => setShowAddQ(true)}
                    style={{
                      padding: '8px 16px',
                      backgroundColor: '#f59e0b',
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
            <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              {questions.map(q => (
                <div
                  key={q.id}
                  style={{
                    backgroundColor: '#fff',
                    borderRadius: '12px',
                    padding: '24px',
                    boxShadow: '0 1px 3px rgba(0,0,0,0.08)',
                    border: '1px solid #e2e8f0'
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' }}>
                    <div>
                      <label style={{
                        display: 'block',
                        fontSize: '15px',
                        fontWeight: '600',
                        color: '#1e293b',
                        marginBottom: '4px'
                      }}>
                        {q.label}
                        {q.required && <span style={{ color: '#ef4444', marginLeft: '4px' }}>*</span>}
                      </label>
                      {q.helpText && (
                        <p style={{ margin: 0, fontSize: '13px', color: '#64748b' }}>
                          {q.helpText}
                        </p>
                      )}
                    </div>
                    
                    {showConfig && (
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button
                          onClick={() => toggleHide(q.id)}
                          style={{
                            padding: '4px 10px',
                            backgroundColor: '#f1f5f9',
                            border: '1px solid #e2e8f0',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            fontSize: '12px'
                          }}
                        >
                          👁️ Hide
                        </button>
                        {q.isCustom && (
                          <button
                            onClick={() => removeQ(q.id)}
                            style={{
                              padding: '4px 10px',
                              backgroundColor: '#fef2f2',
                              border: '1px solid #fecaca',
                              borderRadius: '4px',
                              cursor: 'pointer',
                              fontSize: '12px',
                              color: '#dc2626'
                            }}
                          >
                            🗑️
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                  
                  {renderField(q)}
                  
                  {errors[q.id] && (
                    <p style={{ margin: '8px 0 0', fontSize: '13px', color: '#ef4444', fontWeight: '500' }}>
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
              borderTop: '1px solid #e2e8f0'
            }}>
              <button
                onClick={() => setCurrentPhase(Math.max(0, currentPhase - 1))}
                disabled={currentPhase === 0}
                style={{
                  padding: '14px 28px',
                  backgroundColor: currentPhase === 0 ? '#f1f5f9' : '#fff',
                  color: currentPhase === 0 ? '#94a3b8' : '#475569',
                  border: '2px solid #e2e8f0',
                  borderRadius: '10px',
                  cursor: currentPhase === 0 ? 'not-allowed' : 'pointer',
                  fontWeight: '600',
                  fontSize: '14px',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px'
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
                        padding: '14px 32px',
                        background: isGeneratingPPTX ? '#94a3b8' : 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)',
                        color: 'white',
                        border: 'none',
                        borderRadius: '10px',
                        cursor: isGeneratingPPTX ? 'not-allowed' : 'pointer',
                        fontWeight: '600',
                        fontSize: '15px',
                        boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px'
                      }}
                    >
                      {isGeneratingPPTX ? (
                        <>
                          <span style={{
                            width: '18px',
                            height: '18px',
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
                        padding: '14px 32px',
                        background: isGenerating ? '#94a3b8' : 'linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%)',
                        color: 'white',
                        border: 'none',
                        borderRadius: '10px',
                        cursor: isGenerating ? 'not-allowed' : 'pointer',
                        fontWeight: '600',
                        fontSize: '15px',
                        boxShadow: '0 4px 12px rgba(139, 92, 246, 0.3)',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '10px'
                      }}
                    >
                      {isGenerating ? 'Generating...' : '🤖 Generate JSON'}
                    </button>
                  </>
                ) : (
                  <button
                    onClick={handleNext}
                    style={{
                      padding: '14px 32px',
                      background: 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)',
                      color: 'white',
                      border: 'none',
                      borderRadius: '10px',
                      cursor: 'pointer',
                      fontWeight: '600',
                      fontSize: '15px',
                      boxShadow: '0 4px 12px rgba(59, 130, 246, 0.3)',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px'
                    }}
                  >
                    Next →
                  </button>
                )}
              </div>
            </div>
          </div>
        </main>

        {/* Right Panel - Validation Report (shown when report has items) */}
        {showReport && (
          <aside style={{
            width: '350px',
            backgroundColor: '#fff',
            borderLeft: '1px solid #e2e8f0',
            padding: '24px',
            overflowY: 'auto'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: '600', color: '#1e293b' }}>
                Validation Report
              </h3>
              <button
                onClick={() => setShowReport(false)}
                style={{
                  padding: '6px 10px',
                  backgroundColor: '#f1f5f9',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer'
                }}
              >
                ✕
              </button>
            </div>

            {report.errors.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '14px', fontWeight: '600', color: '#dc2626' }}>
                  ❌ Errors ({report.errors.length})
                </h4>
                {report.errors.map((e, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: '#fef2f2',
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: '4px solid #dc2626'
                  }}>
                    <div style={{ fontSize: '12px', color: '#dc2626', fontWeight: '600' }}>{e.phase}</div>
                    <div style={{ fontSize: '13px', color: '#991b1b' }}>{e.field}: {e.msg}</div>
                  </div>
                ))}
              </div>
            )}

            {report.warnings.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 12px', fontSize: '14px', fontWeight: '600', color: '#d97706' }}>
                  ⚠️ Warnings ({report.warnings.length})
                </h4>
                {report.warnings.map((w, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: '#fffbeb',
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: '4px solid #d97706'
                  }}>
                    <div style={{ fontSize: '12px', color: '#d97706', fontWeight: '600' }}>{w.phase}</div>
                    <div style={{ fontSize: '13px', color: '#92400e' }}>{w.msg}</div>
                  </div>
                ))}
              </div>
            )}

            {report.suggestions.length > 0 && (
              <div>
                <h4 style={{ margin: '0 0 12px', fontSize: '14px', fontWeight: '600', color: '#2563eb' }}>
                  💡 Suggestions ({report.suggestions.length})
                </h4>
                {report.suggestions.map((s, i) => (
                  <div key={i} style={{
                    padding: '12px',
                    backgroundColor: '#eff6ff',
                    borderRadius: '8px',
                    marginBottom: '8px',
                    borderLeft: '4px solid #2563eb'
                  }}>
                    <div style={{ fontSize: '12px', color: '#2563eb', fontWeight: '600' }}>{s.phase}</div>
                    <div style={{ fontSize: '13px', color: '#1e40af' }}>{s.msg}</div>
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
            backgroundColor: '#fff',
            borderRadius: '16px',
            padding: '32px',
            width: '90%',
            maxWidth: '500px',
            boxShadow: '0 20px 50px rgba(0,0,0,0.2)'
          }}>
            <h3 style={{ margin: '0 0 24px', fontSize: '20px', fontWeight: '600', color: '#1e293b' }}>
              Add Custom Question
            </h3>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div>
                <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500', color: '#475569' }}>
                  Question Label
                </label>
                <input
                  type="text"
                  value={newQ.label}
                  onChange={e => setNewQ({ ...newQ, label: e.target.value })}
                  placeholder="Enter question label"
                  style={{
                    width: '100%',
                    padding: '12px',
                    border: '2px solid #e2e8f0',
                    borderRadius: '8px',
                    fontSize: '14px'
                  }}
                />
              </div>
              
              <div>
                <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500', color: '#475569' }}>
                  Field Type
                </label>
                <select
                  value={newQ.type}
                  onChange={e => setNewQ({ ...newQ, type: e.target.value })}
                  style={{
                    width: '100%',
                    padding: '12px',
                    border: '2px solid #e2e8f0',
                    borderRadius: '8px',
                    fontSize: '14px'
                  }}
                >
                  <option value="text">Text</option>
                  <option value="textarea">Text Area</option>
                  <option value="number">Number</option>
                  <option value="date">Date</option>
                </select>
              </div>
              
              <div>
                <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500', color: '#475569' }}>
                  Placeholder
                </label>
                <input
                  type="text"
                  value={newQ.placeholder}
                  onChange={e => setNewQ({ ...newQ, placeholder: e.target.value })}
                  placeholder="Optional placeholder text"
                  style={{
                    width: '100%',
                    padding: '12px',
                    border: '2px solid #e2e8f0',
                    borderRadius: '8px',
                    fontSize: '14px'
                  }}
                />
              </div>
              
              <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={newQ.required}
                  onChange={e => setNewQ({ ...newQ, required: e.target.checked })}
                  style={{ width: '18px', height: '18px', accentColor: '#3b82f6' }}
                />
                <span style={{ fontWeight: '500', color: '#475569' }}>Required field</span>
              </label>
            </div>
            
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '32px' }}>
              <button
                onClick={() => setShowAddQ(false)}
                style={{
                  padding: '12px 24px',
                  backgroundColor: '#f1f5f9',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontWeight: '500'
                }}
              >
                Cancel
              </button>
              <button
                onClick={addQuestion}
                style={{
                  padding: '12px 24px',
                  backgroundColor: '#3b82f6',
                  color: 'white',
                  border: 'none',
                  borderRadius: '8px',
                  cursor: 'pointer',
                  fontWeight: '500'
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
            backgroundColor: '#fff',
            borderRadius: '16px',
            padding: '32px',
            width: '90%',
            maxWidth: '800px',
            maxHeight: '80vh',
            overflow: 'auto',
            boxShadow: '0 20px 50px rgba(0,0,0,0.2)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
              <h3 style={{ margin: 0, fontSize: '20px', fontWeight: '600', color: '#1e293b' }}>
                Generated IM Content
              </h3>
              <button
                onClick={() => setShowGeneratedContent(false)}
                style={{
                  padding: '8px 12px',
                  backgroundColor: '#f1f5f9',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer'
                }}
              >
                ✕ Close
              </button>
            </div>
            
            <pre style={{
              backgroundColor: '#f8fafc',
              padding: '20px',
              borderRadius: '8px',
              overflow: 'auto',
              fontSize: '13px',
              lineHeight: '1.5'
            }}>
              {JSON.stringify(generatedContent.content, null, 2)}
            </pre>
            
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '24px' }}>
              <button
                onClick={downloadJSON}
                style={{
                  padding: '12px 24px',
                  backgroundColor: '#3b82f6',
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

      {/* CSS Animation Keyframes */}
      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        @keyframes slideIn {
          from { transform: translateX(100%); opacity: 0; }
          to { transform: translateX(0); opacity: 1; }
        }
        * {
          box-sizing: border-box;
        }
        body {
          margin: 0;
          padding: 0;
        }
      `}</style>
    </div>
  );
}
