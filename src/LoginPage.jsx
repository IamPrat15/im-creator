import React, { useState } from 'react';

// ============================================================================
// APP VERSION - Semantic Versioning (Major.Minor.Patch)
// ============================================================================
// Version History:
//   v6.0.0 - Document types, buyer types, industry content, 50 templates,
//            unlimited case studies, auto-logout, Word/PDF/JSON export
//   v5.0.0 - Usage dashboard, dynamic charts, hide empty sections, CSV export
//   v4.0.0 - Smart text handling, format labels, advisor placeholder fix
//   v3.0.0 - Slide numbering, pie charts, case studies, theme colors
//   v2.0.0 - Authentication, Office 365 SSO support
//   v1.0.0 - Initial release
// ============================================================================
const APP_VERSION = 'v6.0.0';

// Test credentials
const TEST_CREDENTIALS = {
  username: import.meta.env.VITE_TEST_USERNAME || 'TestUser',
  password: import.meta.env.VITE_TEST_PASSWORD || 'Password'
};

// Theme colors
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

    await new Promise(resolve => setTimeout(resolve, 800));

    if (username === TEST_CREDENTIALS.username && password === TEST_CREDENTIALS.password) {
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
              <form onSubmit={handleLogin}>
                <h2 style={{
                  fontSize: '20px',
                  fontWeight: '600',
                  color: THEME.text,
                  margin: '0 0 24px 0'
                }}>
                  Welcome Back
                </h2>

                {error && (
                  <div style={{
                    padding: '12px 16px',
                    backgroundColor: '#FED7D7',
                    borderRadius: '8px',
                    marginBottom: '16px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px'
                  }}>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill={THEME.error}>
                      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
                    </svg>
                    <span style={{ fontSize: '14px', color: THEME.error }}>{error}</span>
                  </div>
                )}

                <div style={{ marginBottom: '16px' }}>
                  <label style={{
                    display: 'block',
                    fontSize: '14px',
                    fontWeight: '500',
                    color: THEME.text,
                    marginBottom: '6px'
                  }}>
                    Username
                  </label>
                  <input
                    type="text"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="Enter your username"
                    style={{
                      width: '100%',
                      padding: '12px 16px',
                      border: `1px solid ${THEME.border}`,
                      borderRadius: '8px',
                      fontSize: '14px',
                      outline: 'none',
                      boxSizing: 'border-box',
                      transition: 'border-color 0.2s'
                    }}
                  />
                </div>

                <div style={{ marginBottom: '24px' }}>
                  <label style={{
                    display: 'block',
                    fontSize: '14px',
                    fontWeight: '500',
                    color: THEME.text,
                    marginBottom: '6px'
                  }}>
                    Password
                  </label>
                  <div style={{ position: 'relative' }}>
                    <input
                      type={showPassword ? 'text' : 'password'}
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      placeholder="Enter your password"
                      style={{
                        width: '100%',
                        padding: '12px 40px 12px 16px',
                        border: `1px solid ${THEME.border}`,
                        borderRadius: '8px',
                        fontSize: '14px',
                        outline: 'none',
                        boxSizing: 'border-box'
                      }}
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
                        padding: '4px'
                      }}
                    >
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke={THEME.textLight} strokeWidth="2">
                        {showPassword ? (
                          <>
                            <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/>
                            <line x1="1" y1="1" x2="23" y2="23"/>
                          </>
                        ) : (
                          <>
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                            <circle cx="12" cy="12" r="3"/>
                          </>
                        )}
                      </svg>
                    </button>
                  </div>
                </div>

                <button
                  type="submit"
                  disabled={isLoading}
                  style={{
                    width: '100%',
                    padding: '14px',
                    backgroundColor: isLoading ? THEME.primaryLight : THEME.primary,
                    color: 'white',
                    border: 'none',
                    borderRadius: '8px',
                    fontSize: '16px',
                    fontWeight: '600',
                    cursor: isLoading ? 'wait' : 'pointer',
                    transition: 'all 0.2s',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '8px'
                  }}
                >
                  {isLoading ? (
                    <>
                      <div style={{
                        width: '18px',
                        height: '18px',
                        border: '2px solid rgba(255,255,255,0.3)',
                        borderTopColor: 'white',
                        borderRadius: '50%',
                        animation: 'spin 1s linear infinite'
                      }} />
                      Signing in...
                    </>
                  ) : 'Sign In'}
                </button>
              </form>

              {/* Features list */}
              <div style={{ marginTop: '24px', paddingTop: '24px', borderTop: `1px solid ${THEME.border}` }}>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <div style={{
                      width: '24px',
                      height: '24px',
                      borderRadius: '50%',
                      backgroundColor: `${THEME.accent}15`,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center'
                    }}>
                      <svg width="12" height="12" viewBox="0 0 24 24" fill={THEME.accent}>
                        <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
                      </svg>
                    </div>
                    <span style={{ fontSize: '14px', color: THEME.text }}>3 Document Types (Presentation, CIM, Teaser)</span>
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
                    <span style={{ fontSize: '14px', color: THEME.text }}>50 Professional Templates</span>
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
                      </svg>
                    </div>
                    <span style={{ fontSize: '14px', color: THEME.text }}>Export to PPTX, PDF, Word, JSON</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Right Side - Feature Card */}
          <div style={{ width: '420px', flexShrink: 0 }}>
            {/* Badge */}
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
                <span style={{ fontSize: '14px', fontWeight: '500', color: THEME.text }}>v6.0 Features</span>
              </div>
            </div>

            {/* Main Card */}
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
                Transform your M&A process with AI-powered document generation, industry-specific content, and professional presentations.
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
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>50</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Templates</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>6</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Industries</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>∞</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Case Studies</div>
                </div>
                <div style={{
                  backgroundColor: 'rgba(255,255,255,0.1)',
                  borderRadius: '12px',
                  padding: '20px'
                }}>
                  <div style={{ fontSize: '36px', fontWeight: '700', marginBottom: '4px' }}>4</div>
                  <div style={{ fontSize: '13px', opacity: 0.8 }}>Export Formats</div>
                </div>
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
        <p style={{ margin: 0, fontSize: '12px', color: THEME.textLight }}>{APP_VERSION}</p>
      </footer>

      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        * { box-sizing: border-box; }
        html, body { margin: 0; padding: 0; }
      `}</style>
    </div>
  );
}
