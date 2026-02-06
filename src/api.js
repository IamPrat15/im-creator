// API Service for IM Creator
// Version: 8.1.0
// Handles all backend communication

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

// Helper function to handle API responses
async function handleResponse(response) {
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error || errorData.message || `HTTP error! status: ${response.status}`);
  }
  return response.json();
}

// ============================================================================
// HEALTH & INFO ENDPOINTS
// ============================================================================

// Check API health
export async function checkHealth() {
  const response = await fetch(`${API_BASE_URL}/api/health`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// Get API version
export async function getVersion() {
  const response = await fetch(`${API_BASE_URL}/api/version`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// ============================================================================
// TEMPLATES & CONFIGURATION
// ============================================================================

// Get available templates
export async function getTemplates() {
  const response = await fetch(`${API_BASE_URL}/api/templates`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// Get industry data
export async function getIndustries() {
  const response = await fetch(`${API_BASE_URL}/api/industries`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// Get document type configurations
export async function getDocumentTypes() {
  const response = await fetch(`${API_BASE_URL}/api/document-types`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// ============================================================================
// GENERATION ENDPOINTS
// ============================================================================

// Generate IM content using Claude API
export async function generateIM(data) {
  const response = await fetch(`${API_BASE_URL}/api/generate-im`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data })
  });
  return handleResponse(response);
}

// Generate Professional PowerPoint presentation
export async function generatePPTX(data, theme = 'modern-blue') {
  const response = await fetch(`${API_BASE_URL}/api/generate-pptx`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data, theme })
  });
  return handleResponse(response);
}

// Validate form data
export async function validateData(data) {
  const response = await fetch(`${API_BASE_URL}/api/validate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data })
  });
  return handleResponse(response);
}

// ============================================================================
// EXPORT ENDPOINTS
// ============================================================================

// Export Q&A to Word document
export async function exportQAWord(data, questionnaire = null) {
  const response = await fetch(`${API_BASE_URL}/api/export-qa-word`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data, questionnaire })
  });
  
  if (!response.ok) {
    throw new Error('Failed to export Word document');
  }
  
  return response.blob();
}

// Export to PDF
export async function exportPDF(data) {
  const response = await fetch(`${API_BASE_URL}/api/export-pdf`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data })
  });
  
  if (!response.ok) {
    throw new Error('Failed to export PDF');
  }
  
  return response.blob();
}

// ============================================================================
// DRAFT MANAGEMENT
// ============================================================================

// Save draft
export async function saveDraft(data, projectId) {
  const response = await fetch(`${API_BASE_URL}/api/drafts`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data, project_id: projectId })
  });
  return handleResponse(response);
}

// Load draft
export async function loadDraft(projectId) {
  const response = await fetch(`${API_BASE_URL}/api/drafts/${projectId}`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// List all drafts
export async function listDrafts() {
  const response = await fetch(`${API_BASE_URL}/api/drafts`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// Delete draft
export async function deleteDraft(projectId) {
  const response = await fetch(`${API_BASE_URL}/api/drafts/${projectId}`, {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// ============================================================================
// USAGE TRACKING
// ============================================================================

// Get usage statistics
export async function getUsageStats() {
  const response = await fetch(`${API_BASE_URL}/api/usage`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// Export usage to CSV
export async function exportUsageCSV() {
  const response = await fetch(`${API_BASE_URL}/api/usage/export`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  
  if (!response.ok) {
    throw new Error('Failed to export usage data');
  }
  
  return response.blob();
}

// Reset usage statistics
export async function resetUsageStats() {
  const response = await fetch(`${API_BASE_URL}/api/usage/reset`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

// ============================================================================
// DYNAMIC UPDATES
// ============================================================================

// Update presentation (dynamic slide updates)
export async function updatePresentation(oldData, newData, presentationId) {
  const response = await fetch(`${API_BASE_URL}/api/update-presentation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ 
      old_data: oldData, 
      new_data: newData,
      presentation_id: presentationId
    })
  });
  return handleResponse(response);
}

// ============================================================================
// AI CHAT
// ============================================================================

// Chat with AI assistant
export async function chat(message, history = [], context = {}) {
  const response = await fetch(`${API_BASE_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message,
      conversation_history: history,
      context
    })
  });
  return handleResponse(response);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

// Download base64 file (used for PPTX download)
export function downloadBase64File(base64Data, filename, mimeType) {
  try {
    // Convert base64 to blob
    const byteCharacters = atob(base64Data);
    const byteNumbers = new Array(byteCharacters.length);
    
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: mimeType });
    
    // Create download link
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    
    // Trigger download
    document.body.appendChild(link);
    link.click();
    
    // Cleanup
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    return true;
  } catch (error) {
    console.error('Download error:', error);
    throw new Error('Failed to download file');
  }
}

// Download blob as file
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Format file size for display
export function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Export API_BASE for direct use if needed
export { API_BASE_URL };

// Export all functions as default object
export default {
  // Health & Info
  checkHealth,
  getVersion,
  
  // Templates & Config
  getTemplates,
  getIndustries,
  getDocumentTypes,
  
  // Generation
  generateIM,
  generatePPTX,
  validateData,
  
  // Export
  exportQAWord,
  exportPDF,
  
  // Drafts
  saveDraft,
  loadDraft,
  listDrafts,
  deleteDraft,
  
  // Usage
  getUsageStats,
  exportUsageCSV,
  resetUsageStats,
  
  // Dynamic Updates
  updatePresentation,
  
  // AI Chat
  chat,
  
  // Utilities
  downloadBase64File,
  downloadBlob,
  formatFileSize
};
