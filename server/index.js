// server/index.js
const express = require('express');
const cors = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Initialize Anthropic client
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// System prompt for IM generation
const SYSTEM_PROMPT = `You are an expert Investment Banking Analyst specializing in creating professional Information Memorandums (IMs) for M&A transactions. You work for RMB (an investment banking firm) and help automate the creation of management presentations for potential acquirers.

Your task is to take the structured data provided and generate professional, investment-banking-quality content for an Information Memorandum.

When generating content:
1. Use formal, professional investment banking language
2. Focus on investment highlights and value creation
3. Quantify everything possible with metrics
4. Structure content for easy PowerPoint conversion
5. Highlight growth potential and competitive advantages
6. Tailor messaging based on target buyer type

Output your response as a structured JSON object that can be used to populate presentation slides.`;

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'healthy', timestamp: new Date().toISOString() });
});

// Generate IM endpoint
app.post('/api/generate-im', async (req, res) => {
  try {
    const { data } = req.body;
    
    if (!data) {
      return res.status(400).json({ error: 'No data provided' });
    }

    const message = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 8000,
      system: SYSTEM_PROMPT,
      messages: [
        {
          role: 'user',
          content: `Please generate a professional Information Memorandum based on the following data. Structure your response as JSON that can be used to populate presentation slides.

Input Data:
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

    const generatedContent = message.content[0].text;
    
    // Try to parse as JSON, if it fails return as text
    let parsedContent;
    try {
      // Extract JSON from the response if wrapped in code blocks
      const jsonMatch = generatedContent.match(/```json\n?([\s\S]*?)\n?```/) || 
                        generatedContent.match(/```\n?([\s\S]*?)\n?```/);
      if (jsonMatch) {
        parsedContent = JSON.parse(jsonMatch[1]);
      } else {
        parsedContent = JSON.parse(generatedContent);
      }
    } catch {
      parsedContent = { rawContent: generatedContent };
    }

    res.json({
      success: true,
      content: parsedContent,
      generatedAt: new Date().toISOString()
    });

  } catch (error) {
    console.error('Error generating IM:', error);
    res.status(500).json({ 
      error: 'Failed to generate IM', 
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

    // Check required fields
    const requiredFields = [
      { key: 'projectCodename', label: 'Project Codename' },
      { key: 'companyName', label: 'Company Name' },
      { key: 'foundedYear', label: 'Founded Year' },
      { key: 'headquarters', label: 'Headquarters' },
      { key: 'founderName', label: 'Founder Name' },
      { key: 'revenueFY25', label: 'Revenue FY25' }
    ];

    requiredFields.forEach(field => {
      if (!data[field.key]) {
        validationResults.errors.push({
          field: field.key,
          message: `${field.label} is required`
        });
      }
    });

    // Financial validations
    if (data.revenueFY25 && data.revenueFY26P) {
      const growth = ((data.revenueFY26P - data.revenueFY25) / data.revenueFY25) * 100;
      if (growth > 100) {
        validationResults.warnings.push({
          field: 'revenueFY26P',
          message: `Projected growth of ${growth.toFixed(0)}% is very high. Please verify.`
        });
      }
    }

    // Content suggestions
    const highlights = (data.investmentHighlights || '').split('\n').filter(l => l.trim()).length;
    if (highlights < 5) {
      validationResults.suggestions.push({
        field: 'investmentHighlights',
        message: `Consider adding more investment highlights (current: ${highlights}, recommended: 5-7)`
      });
    }

    res.json(validationResults);

  } catch (error) {
    console.error('Error validating data:', error);
    res.status(500).json({ error: 'Validation failed', details: error.message });
  }
});

// In-memory draft storage (for demo - use database in production)
const drafts = new Map();

// Save draft endpoint
app.post('/api/drafts', (req, res) => {
  try {
    const { data, projectId } = req.body;
    const id = projectId || `draft_${Date.now()}`;
    
    drafts.set(id, {
      data,
      savedAt: new Date().toISOString()
    });

    res.json({ success: true, projectId: id });
  } catch (error) {
    res.status(500).json({ error: 'Failed to save draft' });
  }
});

// Get draft endpoint
app.get('/api/drafts/:projectId', (req, res) => {
  try {
    const { projectId } = req.params;
    const draft = drafts.get(projectId);
    
    if (!draft) {
      return res.status(404).json({ error: 'Draft not found' });
    }

    res.json(draft);
  } catch (error) {
    res.status(500).json({ error: 'Failed to retrieve draft' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/api/health`);
});