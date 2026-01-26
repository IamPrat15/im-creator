import React, { useState } from 'react';

// ============================================================================
// CONFIGURATION - Test User Credentials
// ============================================================================
// These credentials are for testing purposes only.
// In production, use environment variables and proper authentication.
// 
// To change credentials, update these in Render Dashboard:
// Environment Variables:
//   TEST_USERNAME=TestUser
//   TEST_PASSWORD=Password
// ============================================================================

const TEST_CREDENTIALS = {
  username: import.meta.env.VITE_TEST_USERNAME || 'TestUser',
  password: import.meta.env.VITE_TEST_PASSWORD || 'Password'
};

// ============================================================================
// OFFICE 365 / AZURE AD CONFIGURATION (INACTIVE - See instructions below)
// ============================================================================
// To enable Office 365 login:
// 
// 1. Register your app in Azure Portal:
//    - Go to https://portal.azure.com
//    - Navigate to Azure Active Directory → App registrations → New registration
//    - Name: "IM Creator Pro"
//    - Redirect URI: https://your-domain.com/auth/callback (Web)
//    - Copy the Application (client) ID and Directory (tenant) ID
//
// 2. Configure API permissions:
//    - Add Microsoft Graph → Delegated → User.Read
//    - Grant admin consent
//
// 3. Create client secret:
//    - Certificates & secrets → New client secret
//    - Copy the secret value immediately
//
// 4. Add environment variables in Render:
//    VITE_AZURE_CLIENT_ID=your-client-id
//    VITE_AZURE_TENANT_ID=your-tenant-id
//    VITE_AZURE_REDIRECT_URI=https://your-domain.com/auth/callback
//
// 5. Uncomment the Office 365 login section in this file (search for "OFFICE_365_LOGIN")
//
// 6. Install MSAL library:
//    npm install @azure/msal-browser @azure/msal-react
// ============================================================================

/*
// OFFICE_365_CONFIG - Uncomment when ready to use
import { PublicClientApplication } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || '',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || 'common'}`,
    redirectUri: import.meta.env.VITE_AZURE_REDIRECT_URI || window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  }
};

const msalInstance = new PublicClientApplication(msalConfig);

const loginRequest = {
  scopes: ['User.Read']
};
*/

// Theme colors matching ACC/Agentic Underwriting
const THEME = {
  primary: '#7C1034',
  primaryDark: '#5A0C26',
  primaryLight: '#9A1842',
  secondary: '#2D3748',
  accent: '#48BB78',
  background: '#F7FAFC',
  surface: '#FFFFFF',
  text: '#1A202C',
  textLight: '#718096',
  border: '#E2E8F0',
  error: '#E53E3E',
  cardBg: 'linear-gradient(135deg, #7C1034 0%, #5A0C26 100%)',
};

export default function LoginPage({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setError('');
    setIsLoading(true);

    // Simulate network delay
    await new Promise(resolve => setTimeout(resolve, 800));

    if (username === TEST_CREDENTIALS.username && password === TEST_CREDENTIALS.password) {
      // Store auth state
      sessionStorage.setItem('isAuthenticated', 'true');
      sessionStorage.setItem('user', JSON.stringify({ 
        username, 
        loginTime: new Date().toISOString(),
        method: 'credentials'
      }));
      onLogin({ username, method: 'credentials' });
    } else {
      setError('Invalid username or password');
    }
    
    setIsLoading(false);
  };

  // ============================================================================
  // OFFICE_365_LOGIN - Uncomment this function when ready to use
  // ============================================================================
  /*
  const handleOffice365Login = async () => {
    setError('');
    setIsLoading(true);

    try {
      const response = await msalInstance.loginPopup(loginRequest);
      
      if (response.account) {
        sessionStorage.setItem('isAuthenticated', 'true');
        sessionStorage.setItem('user', JSON.stringify({
          username: response.account.username,
          name: response.account.name,
          loginTime: new Date().toISOString(),
          method: 'office365'
        }));
        onLogin({ 
          username: response.account.username, 
          name: response.account.name,
          method: 'office365' 
        });
      }
    } catch (error) {
      console.error('Office 365 login error:', error);
      setError('Office 365 login failed. Please try again.');
    }
    
    setIsLoading(false);
  };
  */

  return (
    <div style={{
      minHeight: '100vh',
      width: '100%',
      backgroundColor: THEME.background,
      display: 'flex',
      flexDirection: 'column',
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif'
    }}>
      {/* Main Content */}
      <div style={{
        flex: 1,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '40px 20px'
      }}>
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: '80px',
          maxWidth: '1100px',
          width: '100%'
        }}>
          {/* Left Side - Branding */}
          <div style={{ flex: 1 }}>
            <h1 style={{
              fontSize: '52px',
              fontWeight: '400',
              color: THEME.text,
              margin: '0 0 8px 0',
              lineHeight: 1.1,
              fontFamily: 'Georgia, "Times New Roman", serif'
            }}>
              Smart Documents.
            </h1>
            <h1 style={{
              fontSize: '52px',
              fontWeight: '400',
              color: THEME.primary,
              margin: '0 0 24px 0',
              lineHeight: 1.1,
              fontFamily: 'Georgia, "Times New Roman", serif'
            }}>
              Simplified.
            </h1>
            <p style={{
              fontSize: '18px',
              color: THEME.textLight,
              margin: '0 0 40px 0',
              maxWidth: '400px'
            }}>
              AI-powered Information Memorandum generator for M&A transactions
            </p>

            {/* Login Card */}
            <div style={{
              backgroundColor: THEME.surface,
              borderRadius: '16px',
              padding: '32px',
              boxShadow: '0 4px 20px rgba(0,0,0,0.08)',
              border: `1px solid ${THEME.border}`,
              maxWidth: '400px'
            }}>
              {/* ============================================================ */}
              {/* OFFICE_365_LOGIN_BUTTON - Uncomment when ready to use        */}
              {/* ============================================================ */}
              {/*
              <button
                onClick={handleOffice365Login}
                disabled={isLoading}
                style={{
                  width: '100%',
                  padding: '14px 20px',
                  backgroundColor: THEME.surface,
                  border: `1px solid ${THEME.border}`,
                  borderRadius: '8px',
                  cursor: isLoading ? 'not-allowed' : 'pointer',
                  fontSize: '15px',
                  fontWeight: '500',
                  color: THEME.text,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '12px',
                  transition: 'all 0.2s ease',
                  marginBottom: '24px'
                }}
                onMouseEnter={e => e.target.style.backgroundColor = THEME.background}
                onMouseLeave={e => e.target.style.backgroundColor = THEME.surface}
              >
                <svg width="20" height="20" viewBox="0 0 21 21">
                  <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
                  <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                  <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
                  <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
                </svg>
                Continue with Office 365
              </button>

              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: '16px',
                marginBottom: '24px'
              }}>
                <div style={{ flex: 1, height: '1px', backgroundColor: THEME.border }} />
                <span style={{ fontSize: '13px', color: THEME.textLight }}>or sign in with credentials</span>
                <div style={{ flex: 1, height: '1px', backgroundColor: THEME.border }} />
              </div>
              */}

              {/* Credential Login Form */}
              <form onSubmit={handleLogin}>
                <div style={{ marginBottom: '20px' }}>
                  <label style={{
                    display: 'block',
                    fontSize: '14px',
                    fontWeight: '500',
                    color: THEME.text,
                    marginBottom: '8px'
                  }}>
                    Username
                  </label>
                  <input
                    type="text"
                    value={username}
                    onChange={e => setUsername(e.target.value)}
                    placeholder="Enter your username"
                    required
                    style={{
                      width: '100%',
                      padding: '12px 16px',
                      border: `1px solid ${error ? THEME.error : THEME.border}`,
                      borderRadius: '8px',
                      fontSize: '15px',
                      outline: 'none',
                      transition: 'border-color 0.2s ease',
                      boxSizing: 'border-box'
                    }}
                    onFocus={e => e.target.style.borderColor = THEME.primary}
                    onBlur={e => e.target.style.borderColor = error ? THEME.error : THEME.border}
                  />
                </div>

                <div style={{ marginBottom: '24px' }}>
                  <label style={{
                    display: 'block',
                    fontSize: '14px',
                    fontWeight: '500',
                    color: THEME.text,
                    marginBottom: '8px'
                  }}>
                    Password
                  </label>
                  <div style={{ position: 'relative' }}>
                    <input
                      type={showPassword ? 'text' : 'password'}
                      value={password}
                      onChange={e => setPassword(e.target.value)}
                      placeholder="Enter your password"
                      required
                      style={{
                        width: '100%',
                        padding: '12px 48px 12px 16px',
                        border: `1px solid ${error ? THEME.error : THEME.border}`,
                        borderRadius: '8px',
                        fontSize: '15px',
                        outline: 'none',
                        transition: 'border-color 0.2s ease',
                        boxSizing: 'border-box'
                      }}
                      onFocus={e => e.target.style.borderColor = THEME.primary}
                      onBlur={e => e.target.style.borderColor = error ? THEME.error : THEME.border}
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      style={{
                        position: 'absolute',
                        right: '12px',
                        top: '50%',
                        transform: 'translateY(-50%)',
                        background: 'none',
                        border: 'none',
                        cursor: 'pointer',
                        color: THEME.textLight,
                        padding: '4px'
                      }}
                    >
                      {showPassword ? (
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
                          <line x1="1" y1="1" x2="23" y2="23"/>
                        </svg>
                      ) : (
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                          <circle cx="12" cy="12" r="3"/>
                        </svg>
                      )}
                    </button>
                  </div>
                </div>

                {error && (
                  <div style={{
                    padding: '12px 16px',
                    backgroundColor: `${THEME.error}10`,
                    border: `1px solid ${THEME.error}30`,
                    borderRadius: '8px',
                    marginBottom: '20px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px'
                  }}>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill={THEME.error}>
                      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
                    </svg>
                    <span style={{ fontSize: '14px', color: THEME.error }}>{error}</span>
                  </div>
                )}

                <button
                  type="submit"
                  disabled={isLoading}
                  style={{
                    width: '100%',
                    padding: '14px 20px',
                    backgroundColor: isLoading ? THEME.textLight : THEME.primary,
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    cursor: isLoading ? 'not-allowed' : 'pointer',
                    fontSize: '15px',
                    fontWeight: '600',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '10px',
                    transition: 'background-color 0.2s ease'
                  }}
                >
                  {isLoading ? (
                    <>
                      <span style={{
                        width: '18px',
                        height: '18px',
                        border: '2px solid rgba(255,255,255,0.3)',
                        borderTopColor: 'white',
                        borderRadius: '50%',
                        animation: 'spin 1s linear infinite'
                      }} />
                      Signing in...
                    </>
                  ) : (
                    'Sign In'
                  )}
                </button>
              </form>

              {/* Features List */}
              <div style={{ marginTop: '28px', paddingTop: '24px', borderTop: `1px solid ${THEME.border}` }}>
                <p style={{ fontSize: '12px', color: THEME.textLight, margin: '0 0 16px 0', textAlign: 'center' }}>
                  Secure enterprise login
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{
                      width: '24px',
                      height: '24px',
                      borderRadius: '50%',
                      backgroundColor: `${THEME.accent}20`,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}>
                      <svg width="12" height="12" viewBox="0 0 24 24" fill={THEME.accent}>
                        <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                      </svg>
                    </div>
                    <span style={{ fontSize: '14px', color: THEME.text }}>Professional IM generation in minutes</span>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{
                      width: '24px',
                      height: '24px',
                      borderRadius: '50%',
                      backgroundColor: `${THEME.primary}15`,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}>
                      <svg width="12" height="12" viewBox="0 0 24 24" fill={THEME.primary}>
                        <path d="M13 10V3L4 14h7v7l9-11h-7z"/>
                      </svg>
                    </div>
                    <span style={{ fontSize: '14px', color: THEME.text }}>AI-powered content generation</span>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{
                      width: '24px',
                      height: '24px',
                      borderRadius: '50%',
                      backgroundColor: `${THEME.secondary}15`,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}>
                      <svg width="12" height="12" viewBox="0 0 24 24" fill={THEME.secondary}>
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8" fill="none" stroke={THEME.secondary} strokeWidth="2"/>
                      </svg>
                    </div>
                    <span style={{ fontSize: '14px', color: THEME.text }}>Export to PowerPoint & PDF</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Right Side - Feature Card */}
          <div style={{
            width: '420px',
            flexShrink: 0
          }}>
            {/* Floating Badge */}
            <div style={{
              display: 'flex',
              justifyContent: 'flex-end',
              marginBottom: '-20px',
              marginRight: '-10px',
              position: 'relative',
              zIndex: 10
            }}>
              <div style={{
                backgroundColor: THEME.surface,
                borderRadius: '24px',
                padding: '10px 20px',
                boxShadow: '0 4px 15px rgba(0,0,0,0.1)',
                display: 'flex',
                alignItems: 'center',
                gap: '10px'
              }}>
                <div style={{
                  width: '24px',
                  height: '24px',
                  borderRadius: '50%',
                  backgroundColor: `${THEME.accent}20`,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill={THEME.accent}>
                    <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                  </svg>
                </div>
                <span style={{ fontSize: '14px', fontWeight: '500', color: THEME.text }}>IM Generated</span>
              </div>
            </div>

            {/* Main Feature Card */}
            <div style={{
              background: THEME.cardBg,
              borderRadius: '24px',
              padding: '40px',
              color: 'white'
            }}>
              <h2 style={{
                fontSize: '28px',
                fontWeight: '600',
                margin: '0 0 16px 0',
                fontStyle: 'italic'
              }}>
                Intelligent Documents
              </h2>
              <p style={{
                fontSize: '15px',
                opacity: 0.9,
                margin: '0 0 32px 0',
                lineHeight: 1.6
              }}>
                Transform your M&A process with AI-powered document analysis, automated content generation, and professional presentations.
              </p>

              {/* Stats Grid */}
              <div style={{
                display: 'grid',
                gridTemplateColumns: '1fr 1fr',
                gap: '20px',
                marginBottom: '32px'
              }}>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>70%</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Faster Processing</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>10+</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>AI Templates</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>50+</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Slide Designs</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>24/7</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Availability</div>
                </div>
              </div>

              {/* Bottom Badge */}
              <div style={{
                backgroundColor: THEME.surface,
                borderRadius: '12px',
                padding: '14px 20px',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '12px'
              }}>
                <div style={{
                  width: '32px',
                  height: '32px',
                  borderRadius: '8px',
                  backgroundColor: THEME.background,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  <svg width="18" height="18" viewBox="0 0 24 24" fill={THEME.primary}>
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8" fill="none" stroke={THEME.primary} strokeWidth="2"/>
                    <line x1="16" y1="13" x2="8" y2="13" stroke="white" strokeWidth="2"/>
                    <line x1="16" y1="17" x2="8" y2="17" stroke="white" strokeWidth="2"/>
                  </svg>
                </div>
                <span style={{ fontSize: '14px', fontWeight: '500', color: THEME.text }}>
                  Document Ready
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Footer */}
      <footer style={{
        textAlign: 'center',
        padding: '24px',
        borderTop: `1px solid ${THEME.border}`
      }}>
        <p style={{ margin: '0 0 8px 0', fontSize: '14px', color: THEME.textLight }}>
          © 2025 Applied Cloud Computing Private Limited · 
          <a href="#" style={{ color: THEME.text, marginLeft: '4px', textDecoration: 'none' }}>Privacy</a> · 
          <a href="#" style={{ color: THEME.text, marginLeft: '4px', textDecoration: 'none' }}>Terms</a>
        </p>
        <p style={{ margin: 0, fontSize: '12px', color: THEME.textLight }}>v1.0.0</p>
      </footer>

      {/* CSS Animation */}
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
        input::placeholder {
          color: ${THEME.textLight};
        }
      `}</style>
    </div>
  );
}
