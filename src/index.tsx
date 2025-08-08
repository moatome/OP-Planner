import React from 'react';
import ReactDOM from 'react-dom/client';
import './tailwind-output.css';
import App from './App';
import { PublicClientApplication, EventType, EventMessage, EventPayload, AuthenticationResult, AccountInfo, LogLevel } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';

// Use environment variables or replace with your actual values
const CLIENT_ID = process.env.REACT_APP_AAD_CLIENT_ID || '06c5c649-973a-49a0-ba36-56ecf11285f1';
const TENANT_ID = process.env.REACT_APP_AAD_TENANT_ID || 'd0c4995a-6bf2-4d26-9281-906c0c59b9cb';

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    // Try using the specific tenant first, fall back to common if needed
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
    // Add these for better error handling
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  system: {
    // Add logging for debugging
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (!containsPii) {
          console.log(`MSAL [${level}]: ${message}`);
        }
      },
      logLevel: LogLevel.Info , // Change to 'Verbose' for more detailed logs
    }
  }
});

// Add account handling with proper type checking
msalInstance.addEventCallback((event: EventMessage) => {
  console.log('MSAL Event:', event.eventType, event);
  
  if (
    (event.eventType === EventType.LOGIN_SUCCESS || 
     event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS) &&
    event.payload
  ) {
    const payload = event.payload as AuthenticationResult;
    if (payload.account) {
      msalInstance.setActiveAccount(payload.account);
      console.log('Active account set:', payload.account);
    }
  }
  
  if (event.eventType === EventType.LOGIN_FAILURE) {
    console.error('Login failed:', event.error);
  }
});

// Initialize MSAL
msalInstance.initialize().then(() => {
  console.log('MSAL initialized successfully');
  
  const root = ReactDOM.createRoot(document.getElementById('root')!);
  root.render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>
  );
}).catch(error => {
  console.error('MSAL initialization failed:', error);
  const root = ReactDOM.createRoot(document.getElementById('root')!);
  root.render(
    <div className="flex items-center justify-center h-screen">
      <div className="text-center text-red-600 max-w-md p-6 bg-white rounded-lg shadow">
        <h2 className="text-xl font-bold mb-4">Authentication System Error</h2>
        <p className="mb-4">Failed to initialize authentication. Please refresh the page or contact support.</p>
        <p className="text-sm text-gray-600">Error: {error.message}</p>
      </div>
    </div>
  );
});