import React, { useState, useEffect } from 'react';
import LoginPage from './LoginPage';
import IMCreatorApp from './IMCreatorApp';

// ============================================================================
// App.jsx - Main Application with Authentication
// ============================================================================
// This component handles routing between Login and IMCreator based on auth state.
// 
// Test Credentials (configured in Render environment):
//   Username: TestUser
//   Password: Password
// ============================================================================

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [user, setUser] = useState(null);
  const [isLoading, setIsLoading] = useState(true);

  // Check for existing session on mount
  useEffect(() => {
    const authState = sessionStorage.getItem('isAuthenticated');
    const userData = sessionStorage.getItem('user');
    
    if (authState === 'true' && userData) {
      setIsAuthenticated(true);
      setUser(JSON.parse(userData));
    }
    
    setIsLoading(false);
  }, []);

  // Handle successful login
  const handleLogin = (userData) => {
    setIsAuthenticated(true);
    setUser(userData);
  };

  // Handle logout
  const handleLogout = () => {
    sessionStorage.removeItem('isAuthenticated');
    sessionStorage.removeItem('user');
    setIsAuthenticated(false);
    setUser(null);
  };

  // Show loading state while checking auth
  if (isLoading) {
    return (
      <div style={{
        minHeight: '100vh',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        backgroundColor: '#F7FAFC',
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
      }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{
            width: '48px',
            height: '48px',
            border: '3px solid #E2E8F0',
            borderTopColor: '#7C1034',
            borderRadius: '50%',
            animation: 'spin 1s linear infinite',
            margin: '0 auto 16px'
          }} />
          <p style={{ color: '#718096', fontSize: '14px' }}>Loading...</p>
        </div>
        <style>{`
          @keyframes spin {
            to { transform: rotate(360deg); }
          }
        `}</style>
      </div>
    );
  }

  // Show login page if not authenticated
  if (!isAuthenticated) {
    return <LoginPage onLogin={handleLogin} />;
  }

  // Show main app with logout capability
  return <IMCreatorApp user={user} onLogout={handleLogout} />;
}
