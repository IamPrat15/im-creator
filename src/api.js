import axios from 'axios';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

const api = axios.create({
  baseURL: API_BASE_URL,
  timeout: 60000,
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

export default api;