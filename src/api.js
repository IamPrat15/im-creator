// API Service for IM Creator
// Handles all backend communication

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

// Helper function to handle API responses
async function handleResponse(response) {
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
  }
  return response.json();
}

// Check API health
export async function checkHealth() {
  const response = await fetch(`${API_BASE_URL}/api/health`, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' }
  });
  return handleResponse(response);
}

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
export async function generatePPTX(data, theme = 'modern-tech') {
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

// Save draft
export async function saveDraft(data, projectId) {
  const response = await fetch(`${API_BASE_URL}/api/drafts`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data, projectId })
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

// Export all functions as default object
export default {
  checkHealth,
  generateIM,
  generatePPTX,
  validateData,
  saveDraft,
  loadDraft,
  downloadBase64File
};
