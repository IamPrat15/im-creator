import React, { useState, useEffect } from 'react';
import { generateIM, saveDraft, checkHealth } from './api';

const defaultQuestionnaire = {
  phases: [
    {
      id: 'project-setup',
      name: 'Project Setup',
      icon: '📋',
      description: 'Basic project information',
      questions: [
        { id: 'projectCodename', type: 'text', label: 'Project Codename', placeholder: 'e.g., Project Phoenix', required: true, helpText: 'Confidential identifier', order: 1 },
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
        { id: 'companyDescription', type: 'textarea', label: 'Company Description', required: true, helpText: 'Appears in executive summary', order: 3 },
        { id: 'employeeCountFT', type: 'number', label: 'Full-Time Employees', required: true, order: 4 },
        { id: 'employeeCountOther', type: 'number', label: 'Contractors/Trainees', order: 5 },
        { id: 'investmentHighlights', type: 'textarea', label: 'Investment Highlights', placeholder: 'One highlight per line', helpText: 'Recommend 5-7 highlights', order: 6 }
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
        { id: 'previousCompanies', type: 'textarea', label: 'Previous Companies', placeholder: 'Company | Role | Duration', order: 5 },
        { id: 'leadershipTeam', type: 'textarea', label: 'Leadership Team', placeholder: 'Name | Title | Department', order: 6 }
      ]
    },
    {
      id: 'services-products',
      name: 'Services & Products',
      icon: '⚙️',
      description: 'Offerings & capabilities',
      questions: [
        { id: 'serviceLines', type: 'textarea', label: 'Service Lines', placeholder: 'Cloud & Automation | 39% | AWS migration, DevOps', required: true, helpText: 'Name | Revenue % | Description', order: 1 },
        { id: 'products', type: 'textarea', label: 'Proprietary Products', placeholder: 'AI Agent Studio | Platform for AI agents | 500+ templates', order: 2 },
        { id: 'techPartnerships', type: 'textarea', label: 'Technology Partnerships', placeholder: 'AWS Advanced Tier Partner\nDatabricks Partner', order: 3 },
        { id: 'certifications', type: 'textarea', label: 'Certifications', placeholder: 'AWS Financial Services Competency', order: 4 }
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
        { id: 'topClients', type: 'textarea', label: 'Top Clients', placeholder: 'Axis Bank | BFSI | 2015\nHDFC Bank | BFSI | 2018', required: true, order: 4 },
        { id: 'top10Concentration', type: 'number', label: 'Top 10 Client Concentration %', required: true, order: 5 },
        { id: 'netRetention', type: 'number', label: 'Net Retention Rate %', order: 6 }
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
        { id: 'revenueFY26P', type: 'number', label: 'Revenue FY26P (Cr)', required: true, order: 4 },
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
        { id: 'cs1Client', type: 'text', label: 'Case Study 1 - Client', order: 1 },
        { id: 'cs1Challenge', type: 'textarea', label: 'Challenge', order: 2 },
        { id: 'cs1Solution', type: 'textarea', label: 'Solution', order: 3 },
        { id: 'cs1Results', type: 'textarea', label: 'Results', placeholder: '40% reduction in processing time\n60% cost savings', order: 4 },
        { id: 'cs2Client', type: 'text', label: 'Case Study 2 - Client', order: 5 },
        { id: 'cs2Challenge', type: 'textarea', label: 'Challenge', order: 6 },
        { id: 'cs2Solution', type: 'textarea', label: 'Solution', order: 7 },
        { id: 'cs2Results', type: 'textarea', label: 'Results', order: 8 }
      ]
    },
    {
      id: 'growth-strategy',
      name: 'Growth Strategy',
      icon: '🎯',
      description: 'Future plans',
      questions: [
        { id: 'growthDrivers', type: 'textarea', label: 'Key Growth Drivers', required: true, order: 1 },
        { id: 'competitiveAdvantages', type: 'textarea', label: 'Competitive Advantages', required: true, helpText: 'Minimum 5 advantages', order: 2 },
        { id: 'shortTermGoals', type: 'textarea', label: 'Short-Term Goals (0-12 months)', order: 3 },
        { id: 'mediumTermGoals', type: 'textarea', label: 'Medium-Term Goals (1-3 years)', order: 4 },
        { id: 'synergiesStrategic', type: 'textarea', label: 'Synergies for Strategic Buyers', order: 5 },
        { id: 'synergiesFinancial', type: 'textarea', label: 'Synergies for Financial Investors', order: 6 }
      ]
    },
    {
      id: 'review-generate',
      name: 'Review & Generate',
      icon: '🏆',
      description: 'Final options',
      questions: [
        { id: 'generateVariants', type: 'multiselect', label: 'Content Variants', options: [
          { value: 'financial', label: 'Financial Focus Summary' },
          { value: 'tech', label: 'Technology Focus Summary' },
          { value: 'market', label: 'Market Position Summary' }
        ], order: 1 },
        { id: 'templateStyle', type: 'select', label: 'Presentation Template', required: true, options: [
          { value: 'modern-tech', label: 'Modern Tech (Burgundy)' },
          { value: 'conservative', label: 'Conservative Banking (Navy)' },
          { value: 'minimalist', label: 'Minimalist (Black/White)' }
        ], defaultValue: 'modern-tech', order: 2 },
        { id: 'exportFormat', type: 'multiselect', label: 'Export Formats', options: [
          { value: 'pptx', label: 'PowerPoint (.pptx)' },
          { value: 'pdf', label: 'PDF' },
          { value: 'json', label: 'JSON Data' }
        ], defaultValue: ['pptx'], order: 3 }
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
    setTimeout(() => setNotification(null), 4000);
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

  const handleGenerate = async () => {
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
      showNotification('IM generated successfully!', 'success');
    } catch (error) {
      showNotification('Failed to generate IM. Please try again.', 'error');
      console.error('Generation error:', error);
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadJSON = () => {
    if (!generatedContent) return;
    const blob = new Blob([JSON.stringify(generatedContent.content, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${formData.projectCodename || 'im'}_content.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const renderInput = (q) => {
    const val = formData[q.id] || q.defaultValue || '';
    const baseStyle = {
      width: '100%',
      padding: '10px 12px',
      border: `1px solid ${errors[q.id] ? '#EF4444' : '#D1D5DB'}`,
      borderRadius: '8px',
      fontSize: '14px',
      boxSizing: 'border-box',
      backgroundColor: '#FFFFFF',
      color: '#111827'
    };

    switch (q.type) {
      case 'textarea':
        return (
          <textarea
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            rows={4}
            style={{
              ...baseStyle,
              resize: 'vertical',
              fontFamily: 'inherit',
              minHeight: '100px',
              color: '#111827'
            }}
          />
        );
      case 'select':
        return (
          <select
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            style={{ ...baseStyle, color: '#111827' }}
          >
            <option value="">Select...</option>
            {q.options?.map(o => (
              <option key={o.value} value={o.value}>{o.label}</option>
            ))}
          </select>
        );
      case 'multiselect':
        const sel = Array.isArray(val) ? val : [];
        return (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
            {q.options?.map(o => (
              <label
                key={o.value}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '8px 12px',
                  backgroundColor: sel.includes(o.value) ? '#FDF2F4' : '#F9FAFB',
                  border: `1px solid ${sel.includes(o.value) ? '#7C1034' : '#E5E7EB'}`,
                  borderRadius: '6px',
                  fontSize: '13px',
                  cursor: 'pointer',
                  transition: 'all 0.2s',
                  color: '#374151'
                }}
              >
                <input
                  type="checkbox"
                  checked={sel.includes(o.value)}
                  onChange={e => updateField(q.id, e.target.checked ? [...sel, o.value] : sel.filter(v => v !== o.value))}
                  style={{ accentColor: '#7C1034' }}
                />
                {o.label}
              </label>
            ))}
          </div>
        );
      case 'date':
        return (
          <input
            type="date"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            style={{ ...baseStyle, color: '#111827' }}
          />
        );
      case 'number':
        return (
          <input
            type="number"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            style={{ ...baseStyle, color: '#111827' }}
          />
        );
      default:
        return (
          <input
            type="text"
            value={val}
            onChange={e => updateField(q.id, e.target.value)}
            placeholder={q.placeholder}
            style={{ ...baseStyle, color: '#111827' }}
          />
        );
    }
  };

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#F9FAFB', fontFamily: "'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" }}>
      {/* Notification */}
      {notification && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '20px',
          zIndex: 1000,
          padding: '12px 20px',
          borderRadius: '8px',
          boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
          backgroundColor: notification.type === 'success' ? '#10B981' : notification.type === 'error' ? '#EF4444' : '#3B82F6',
          color: '#FFFFFF',
          fontSize: '14px',
          fontWeight: '500'
        }}>
          {notification.type === 'success' && '✅ '}
          {notification.type === 'error' && '❌ '}
          {notification.message}
        </div>
      )}

      {/* Header */}
      <header style={{ backgroundColor: '#FFFFFF', borderBottom: '1px solid #E5E7EB', position: 'sticky', top: 0, zIndex: 50 }}>
        <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '12px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ width: '40px', height: '40px', backgroundColor: '#7C1034', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#FFFFFF', fontSize: '20px' }}>📄</div>
            <div>
              <h1 style={{ fontSize: '18px', fontWeight: '700', color: '#7C1034', margin: 0 }}>IM Creator</h1>
              <p style={{ fontSize: '11px', color: '#6B7280', margin: 0 }}>RMB Investment Banking</p>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{
              padding: '4px 10px',
              borderRadius: '20px',
              fontSize: '11px',
              fontWeight: '500',
              backgroundColor: apiStatus === 'connected' ? '#ECFDF5' : apiStatus === 'disconnected' ? '#FEE2E2' : '#F3F4F6',
              color: apiStatus === 'connected' ? '#047857' : apiStatus === 'disconnected' ? '#DC2626' : '#6B7280'
            }}>
              {apiStatus === 'connected' ? '🟢 API Connected' : apiStatus === 'disconnected' ? '🔴 API Offline' : '⏳ Checking...'}
            </div>
            <button
              onClick={handleSaveDraft}
              style={{
                padding: '8px 16px',
                backgroundColor: '#FFFFFF',
                border: '1px solid #D1D5DB',
                borderRadius: '8px',
                fontSize: '13px',
                cursor: 'pointer',
                fontWeight: '500',
                color: '#374151'
              }}
            >
              💾 Save Draft
            </button>
          </div>
        </div>
      </header>

      {/* Progress Bar */}
      <div style={{ backgroundColor: '#FFFFFF', borderBottom: '1px solid #E5E7EB', padding: '16px 20px' }}>
        <div style={{ maxWidth: '1200px', margin: '0 auto' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#6B7280', marginBottom: '8px' }}>
            <span>Overall Progress</span>
            <span style={{ fontWeight: '600', color: '#7C1034' }}>{progress}% Complete</span>
          </div>
          <div style={{ height: '8px', backgroundColor: '#E5E7EB', borderRadius: '4px', overflow: 'hidden' }}>
            <div style={{ height: '100%', width: `${progress}%`, backgroundColor: '#7C1034', borderRadius: '4px', transition: 'width 0.4s ease' }} />
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '24px 20px', display: 'flex', gap: '24px' }}>
        {/* Sidebar */}
        <div style={{ width: '240px', flexShrink: 0 }}>
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '12px', border: '1px solid #E5E7EB', overflow: 'hidden', position: 'sticky', top: '120px' }}>
            <div style={{ padding: '14px 16px', backgroundColor: '#FDF2F4', borderBottom: '1px solid #E5E7EB' }}>
              <h2 style={{ fontSize: '14px', fontWeight: '600', margin: 0, color: '#7C1034' }}>IM Sections</h2>
            </div>
            <div style={{ padding: '8px' }}>
              {questionnaire.phases.map((p, i) => (
                <button
                  key={p.id}
                  onClick={() => setCurrentPhase(i)}
                  style={{
                    width: '100%',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    padding: '10px 12px',
                    borderRadius: '8px',
                    marginBottom: '4px',
                    border: 'none',
                    cursor: 'pointer',
                    textAlign: 'left',
                    transition: 'all 0.2s',
                    backgroundColor: currentPhase === i ? '#7C1034' : completedPhases.includes(i) ? '#ECFDF5' : 'transparent',
                    color: currentPhase === i ? '#FFFFFF' : completedPhases.includes(i) ? '#047857' : '#4B5563',
                    fontSize: '13px'
                  }}
                >
                  <span style={{ fontSize: '16px' }}>{completedPhases.includes(i) && currentPhase !== i ? '✅' : p.icon}</span>
                  <span style={{ fontWeight: '500' }}>{p.name}</span>
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* Form Area */}
        <div style={{ flex: 1 }}>
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '12px', border: '1px solid #E5E7EB', overflow: 'hidden' }}>
            {/* Phase Header */}
            <div style={{ padding: '16px 20px', backgroundColor: '#FDF2F4', borderBottom: '1px solid #E5E7EB', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h2 style={{ fontSize: '16px', fontWeight: '600', margin: 0, display: 'flex', alignItems: 'center', gap: '8px', color: '#111827' }}>
                  <span>{phase.icon}</span> {phase.name}
                </h2>
                <p style={{ fontSize: '12px', color: '#6B7280', margin: '4px 0 0 0' }}>{phase.description}</p>
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  onClick={() => setShowConfig(!showConfig)}
                  style={{
                    padding: '6px 12px',
                    backgroundColor: '#FFFFFF',
                    border: '1px solid #D1D5DB',
                    borderRadius: '6px',
                    fontSize: '12px',
                    cursor: 'pointer',
                    color: '#374151'
                  }}
                >
                  ⚙️ Configure
                </button>
                <button
                  onClick={() => setShowAddQ(true)}
                  style={{
                    padding: '6px 12px',
                    backgroundColor: '#7C1034',
                    color: '#FFFFFF',
                    border: 'none',
                    borderRadius: '6px',
                    fontSize: '12px',
                    cursor: 'pointer'
                  }}
                >
                  ➕ Add Question
                </button>
              </div>
            </div>

            {/* Config Panel */}
            {showConfig && (
              <div style={{ padding: '16px 20px', backgroundColor: '#F9FAFB', borderBottom: '1px solid #E5E7EB' }}>
                <h4 style={{ fontSize: '13px', fontWeight: '600', margin: '0 0 10px 0', color: '#111827' }}>Toggle Question Visibility:</h4>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                  {phase.questions.map(q => (
                    <label
                      key={q.id}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px',
                        padding: '6px 10px',
                        backgroundColor: q.isHidden ? '#FEE2E2' : '#ECFDF5',
                        borderRadius: '6px',
                        fontSize: '12px',
                        cursor: 'pointer',
                        color: q.isHidden ? '#991B1B' : '#047857'
                      }}
                    >
                      <input type="checkbox" checked={!q.isHidden} onChange={() => toggleHide(q.id)} />
                      {q.label}
                      {q.isCustom && (
                        <button
                          onClick={(e) => { e.preventDefault(); removeQ(q.id); }}
                          style={{ background: 'none', border: 'none', color: '#EF4444', cursor: 'pointer', padding: '0 4px', fontSize: '16px' }}
                        >
                          ×
                        </button>
                      )}
                    </label>
                  ))}
                </div>
              </div>
            )}

            {/* Questions */}
            <div style={{ padding: '24px 20px' }}>
              {questions.map(q => (
                <div key={q.id} style={{ marginBottom: '20px' }}>
                  <label style={{ display: 'block', fontSize: '13px', fontWeight: '500', color: '#374151', marginBottom: '6px' }}>
                    {q.label}
                    {q.required && <span style={{ color: '#EF4444', marginLeft: '4px' }}>*</span>}
                    {q.isCustom && (
                      <span style={{ fontSize: '10px', backgroundColor: '#FDF2F4', color: '#7C1034', padding: '2px 6px', borderRadius: '4px', marginLeft: '8px' }}>
                        Custom
                      </span>
                    )}
                  </label>
                  {renderInput(q)}
                  {q.helpText && <p style={{ fontSize: '11px', color: '#6B7280', margin: '6px 0 0 0' }}>💡 {q.helpText}</p>}
                  {errors[q.id] && <p style={{ fontSize: '11px', color: '#EF4444', margin: '6px 0 0 0' }}>⚠️ {errors[q.id]}</p>}
                </div>
              ))}
            </div>

            {/* Navigation */}
            <div style={{ padding: '16px 20px', backgroundColor: '#F9FAFB', borderTop: '1px solid #E5E7EB', display: 'flex', justifyContent: 'space-between' }}>
              <button
                onClick={() => setCurrentPhase(Math.max(0, currentPhase - 1))}
                disabled={currentPhase === 0}
                style={{
                  padding: '10px 16px',
                  backgroundColor: '#FFFFFF',
                  border: '1px solid #D1D5DB',
                  borderRadius: '8px',
                  fontSize: '13px',
                  cursor: currentPhase === 0 ? 'not-allowed' : 'pointer',
                  opacity: currentPhase === 0 ? 0.5 : 1,
                  color: '#374151'
                }}
              >
                ← Previous
              </button>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button
                  onClick={() => setShowReport(true)}
                  style={{
                    padding: '10px 16px',
                    backgroundColor: '#FFFFFF',
                    border: '1px solid #D1D5DB',
                    borderRadius: '8px',
                    fontSize: '13px',
                    cursor: 'pointer',
                    color: '#374151'
                  }}
                >
                  📋 Validate
                </button>
                {currentPhase === questionnaire.phases.length - 1 ? (
                  <button
                    onClick={handleGenerate}
                    disabled={isGenerating}
                    style={{
                      padding: '10px 24px',
                      backgroundColor: isGenerating ? '#9CA3AF' : '#7C1034',
                      color: '#FFFFFF',
                      border: 'none',
                      borderRadius: '8px',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: isGenerating ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px'
                    }}
                  >
                    {isGenerating ? '⏳ Generating...' : '🚀 Generate IM'}
                  </button>
                ) : (
                  <button
                    onClick={handleNext}
                    style={{
                      padding: '10px 24px',
                      backgroundColor: '#7C1034',
                      color: '#FFFFFF',
                      border: 'none',
                      borderRadius: '8px',
                      fontSize: '13px',
                      fontWeight: '600',
                      cursor: 'pointer'
                    }}
                  >
                    Continue →
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Add Question Modal */}
      {showAddQ && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '12px', padding: '24px', width: '400px', maxWidth: '90vw' }}>
            <h3 style={{ margin: '0 0 20px 0', fontSize: '18px', fontWeight: '600', color: '#111827' }}>➕ Add Custom Question</h3>
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '13px', fontWeight: '500', marginBottom: '6px', color: '#374151' }}>Question Label *</label>
              <input
                type="text"
                value={newQ.label}
                onChange={e => setNewQ({ ...newQ, label: e.target.value })}
                placeholder="Enter question label"
                style={{ width: '100%', padding: '10px', border: '1px solid #D1D5DB', borderRadius: '8px', boxSizing: 'border-box', color: '#111827', backgroundColor: '#FFFFFF' }}
              />
            </div>
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '13px', fontWeight: '500', marginBottom: '6px', color: '#374151' }}>Question Type</label>
              <select
                value={newQ.type}
                onChange={e => setNewQ({ ...newQ, type: e.target.value })}
                style={{ width: '100%', padding: '10px', border: '1px solid #D1D5DB', borderRadius: '8px', color: '#111827', backgroundColor: '#FFFFFF' }}
              >
                <option value="text">Text (Single Line)</option>
                <option value="textarea">Text Area (Multi Line)</option>
                <option value="number">Number</option>
                <option value="date">Date</option>
              </select>
            </div>
            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '13px', fontWeight: '500', marginBottom: '6px', color: '#374151' }}>Placeholder Text</label>
              <input
                type="text"
                value={newQ.placeholder}
                onChange={e => setNewQ({ ...newQ, placeholder: e.target.value })}
                placeholder="Optional hint text"
                style={{ width: '100%', padding: '10px', border: '1px solid #D1D5DB', borderRadius: '8px', boxSizing: 'border-box', color: '#111827', backgroundColor: '#FFFFFF' }}
              />
            </div>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', marginBottom: '20px', cursor: 'pointer', color: '#374151' }}>
              <input type="checkbox" checked={newQ.required} onChange={e => setNewQ({ ...newQ, required: e.target.checked })} style={{ accentColor: '#7C1034' }} />
              Required field
            </label>
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
              <button
                onClick={() => setShowAddQ(false)}
                style={{ padding: '10px 20px', backgroundColor: '#FFFFFF', border: '1px solid #D1D5DB', borderRadius: '8px', fontSize: '13px', cursor: 'pointer', color: '#374151' }}
              >
                Cancel
              </button>
              <button
                onClick={addQuestion}
                style={{ padding: '10px 20px', backgroundColor: '#7C1034', color: '#FFFFFF', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '500', cursor: 'pointer' }}
              >
                Add Question
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Validation Report Modal */}
      {showReport && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '12px', padding: '24px', width: '600px', maxWidth: '90vw', maxHeight: '80vh', overflow: 'auto' }}>
            <h3 style={{ margin: '0 0 20px 0', fontSize: '18px', fontWeight: '600', color: '#111827' }}>📋 Validation Report</h3>
            {(() => {
              const r = fullValidate();
              const hasIssues = r.errors.length || r.warnings.length || r.suggestions.length;
              return (
                <>
                  {r.errors.length > 0 && (
                    <div style={{ marginBottom: '16px' }}>
                      <h4 style={{ fontSize: '14px', color: '#DC2626', margin: '0 0 10px 0' }}>❌ Errors ({r.errors.length})</h4>
                      {r.errors.map((e, i) => (
                        <div key={i} style={{ padding: '10px 14px', backgroundColor: '#FEE2E2', borderRadius: '8px', marginBottom: '6px', fontSize: '13px', borderLeft: '4px solid #DC2626', color: '#991B1B' }}>
                          <strong>{e.phase}</strong>: {e.field} - {e.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  {r.warnings.length > 0 && (
                    <div style={{ marginBottom: '16px' }}>
                      <h4 style={{ fontSize: '14px', color: '#D97706', margin: '0 0 10px 0' }}>⚠️ Warnings ({r.warnings.length})</h4>
                      {r.warnings.map((w, i) => (
                        <div key={i} style={{ padding: '10px 14px', backgroundColor: '#FEF3C7', borderRadius: '8px', marginBottom: '6px', fontSize: '13px', borderLeft: '4px solid #D97706', color: '#92400E' }}>
                          <strong>{w.phase}</strong>: {w.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  {r.suggestions.length > 0 && (
                    <div style={{ marginBottom: '16px' }}>
                      <h4 style={{ fontSize: '14px', color: '#2563EB', margin: '0 0 10px 0' }}>💡 Suggestions ({r.suggestions.length})</h4>
                      {r.suggestions.map((s, i) => (
                        <div key={i} style={{ padding: '10px 14px', backgroundColor: '#DBEAFE', borderRadius: '8px', marginBottom: '6px', fontSize: '13px', borderLeft: '4px solid #2563EB', color: '#1E40AF' }}>
                          <strong>{s.phase}</strong>: {s.msg}
                        </div>
                      ))}
                    </div>
                  )}
                  {!hasIssues && (
                    <div style={{ padding: '20px', backgroundColor: '#ECFDF5', borderRadius: '8px', textAlign: 'center' }}>
                      <span style={{ fontSize: '32px' }}>✅</span>
                      <p style={{ margin: '10px 0 0 0', color: '#047857', fontWeight: '600', fontSize: '16px' }}>All validations passed!</p>
                    </div>
                  )}
                </>
              );
            })()}
            <div style={{ marginTop: '20px', textAlign: 'right' }}>
              <button
                onClick={() => setShowReport(false)}
                style={{ padding: '10px 24px', backgroundColor: '#7C1034', color: '#FFFFFF', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '500', cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Generated Content Modal */}
      {showGeneratedContent && generatedContent && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '12px', padding: '24px', width: '900px', maxWidth: '95vw', maxHeight: '90vh', overflow: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h3 style={{ margin: 0, fontSize: '20px', fontWeight: '600', color: '#111827' }}>🎉 IM Generated Successfully!</h3>
              <span style={{ fontSize: '12px', color: '#6B7280' }}>Generated: {new Date(generatedContent.generatedAt).toLocaleString()}</span>
            </div>
            <div style={{ backgroundColor: '#1F2937', borderRadius: '8px', padding: '16px', marginBottom: '20px' }}>
              <pre style={{ fontSize: '12px', color: '#E5E7EB', overflow: 'auto', maxHeight: '500px', margin: 0, whiteSpace: 'pre-wrap' }}>
                {JSON.stringify(generatedContent.content, null, 2)}
              </pre>
            </div>
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
              <button
                onClick={downloadJSON}
                style={{ padding: '10px 20px', backgroundColor: '#FFFFFF', border: '1px solid #D1D5DB', borderRadius: '8px', fontSize: '13px', cursor: 'pointer', color: '#374151' }}
              >
                📥 Download JSON
              </button>
              <button
                onClick={() => setShowGeneratedContent(false)}
                style={{ padding: '10px 24px', backgroundColor: '#7C1034', color: '#FFFFFF', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '500', cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
