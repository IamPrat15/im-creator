// src/api.js
import axios from 'axios';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

const api = axios.create({
  baseURL: API_BASE_URL,
  timeout: 120000, // 2 minutes for PPTX generation
  headers: { 'Content-Type': 'application/json' },
});

export const generateIM = async (formData) => {
  try {
    const response = await api.post('/api/generate-im', { data: formData });
    return response.data;
  } catch (error) {
    console.error('Error generating IM:', error);
    throw error;
  }
};

export const generatePPTX = async (formData, theme = 'modern-tech') => {
  try {
    const response = await api.post('/api/generate-pptx', { 
      data: formData,
      theme: theme
    });
    return response.data;
  } catch (error) {
    console.error('Error generating PPTX:', error);
    throw error;
  }
};

export const saveDraft = async (formData, projectId) => {
  try {
    const response = await api.post('/api/drafts', { data: formData, projectId });
    return response.data;
  } catch (error) {
    console.error('Error saving draft:', error);
    throw error;
  }
};

export const getDraft = async (projectId) => {
  try {
    const response = await api.get(`/api/drafts/${projectId}`);
    return response.data;
  } catch (error) {
    console.error('Error getting draft:', error);
    throw error;
  }
};

export const checkHealth = async () => {
  try {
    const response = await api.get('/api/health');
    return response.data;
  } catch (error) {
    console.error('API health check failed:', error);
    throw error;
  }
};

// Helper function to download base64 file
export const downloadBase64File = (base64Data, filename, mimeType) => {
  const byteCharacters = atob(base64Data);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  const blob = new Blob([byteArray], { type: mimeType });
  
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

export default api;
