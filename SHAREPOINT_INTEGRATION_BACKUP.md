# SharePoint Integration Backup

This file contains all the SharePoint integration code that was removed from the main application. Keep this for future reference if you want to implement SharePoint integration in another project.

## Files involved:
- `sharepoint-integration.ts` - Main SharePoint service and hooks
- App.tsx authentication logic
- MSAL configuration

## SharePoint Integration Components

### 1. SharePoint Service Class

```typescript
// SharePoint Integration for OR Planner
import { useState, useEffect } from 'react';

// Updated interface for new requirements
interface PersonnelItem {
  id: number;
  name: string;
  gruppe: 'OP-Pflege' | 'Anästhesie Pflege' | 'OP-Praktikant' | 'Anästhesie Praktikant' | 'MFA' | 'ATA Schüler' | 'OTA Schüler';
  department: string;
  kommentar: string;
  verfügbar: 'Bereitschaften (BD)' | 'Rufdienste (RD)' | 'Frühdienste (Früh)' | 'Zwischendienste/Mitteldienste (Mittel)' | 'Spätdienste (Spät)' | 'nicht verfügbar';
  initials: string;
  isActive: boolean;
  isLocallyAdded?: boolean;
  isLocallyModified?: boolean;
  sharePointId?: number;
  // Excel import related fields (handled locally only)
  shiftAssignment?: string;
  availabilityTags?: string[];
  shiftTags?: string[];
  isAvailable?: boolean;
}

class SharePointService {
  accessToken: string;
  baseUrl: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://graph.microsoft.com/v1.0';
  }

  async getPersonnel(siteId: string, listId: string): Promise<PersonnelItem[]> {
    // Return empty array if no access token - will be handled by calling component
    if (!this.accessToken) {
      console.log('No access token available for SharePoint');
      return [];
    }

    try {
      console.log('Fetching personnel from SharePoint...');
      
      // Get the list items with the new schema fields
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Name,Gruppe,Department,Kommentar,Verfügbar)`,
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          }
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('SharePoint API Error:', {
          status: response.status,
          statusText: response.statusText,
          error: errorText
        });
        
        console.log('SharePoint API failed, returning empty array');
        return [];
      }
      
      const data = await response.json();
      console.log('SharePoint data:', data);
      
      if (!data.value || data.value.length === 0) {
        console.log('No data in SharePoint list');
        return [];
      }
      
      const personnel: PersonnelItem[] = [];
      
      // Process each item - now reading directly from the dedicated fields
      for (const item of data.value) {
        try {
          const fields = item.fields || {};
          
          // Read values from the new SharePoint schema
          const displayName = fields.Name || fields.DisplayName || fields.Title || 'Unknown';
          const gruppe = fields.Gruppe || 'OP-Pflege'; // Default to OP-Pflege if not specified
          const department = fields.Department || 'Unknown Department';
          const kommentar = fields.Kommentar || '';
          const verfügbar = fields.Verfügbar || 'nicht verfügbar';
          
          const personnelItem: PersonnelItem = {
            id: parseInt(item.id) || Math.random() * 10000,
            name: displayName,
            gruppe: gruppe as PersonnelItem['gruppe'],
            department: department,
            kommentar: kommentar,
            verfügbar: verfügbar as PersonnelItem['verfügbar'],
            initials: this.generateInitials(displayName),
            isActive: true,
            sharePointId: parseInt(item.id) // Track SharePoint ID
          };
          
          personnel.push(personnelItem);
        } catch (itemError) {
          console.error('Error processing item:', itemError, item);
        }
      }

      console.log('Processed personnel:', personnel);
      return personnel;
      
    } catch (error) {
      console.error('Error fetching personnel from SharePoint:', error);
      console.log('SharePoint fetch failed, returning empty array');
      return [];
    }
  }

  // Additional SharePoint methods...
  generateInitials(name: string): string {
    if (!name) return 'XX';
    const parts = name.trim().split(' ').filter(part => part.length > 0);
    
    if (parts.length === 0) return 'XX';
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  // ... (other SharePoint methods)
}
```

### 2. Authentication Configuration

```typescript
// Updated auth config - make sure this matches your app registration
export const authConfig = {
  auth: {
    clientId: process.env.REACT_APP_AAD_CLIENT_ID || '06c5c649-973a-49a0-ba36-56ecf11285f1',
    authority: process.env.REACT_APP_AAD_AUTHORITY || 'https://login.microsoftonline.com/d0c4995a-6bf2-4d26-9281-906c0c59b9cb',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage' as const,
    storeAuthStateInCookie: false,
  },
  scopes: [
    'Sites.Read.All',
    'Sites.ReadWrite.All',
    'User.Read',
  ]
};
```

### 3. MSAL Authentication Flow

```typescript
// In main App component
const App = () => {
  const { instance, accounts, inProgress } = useMsal();
  const [accessToken, setAccessToken] = React.useState<string | null>(null);
  const [authError, setAuthError] = React.useState<string | null>(null);
  
  const SITE_ID = 'a4cba12d-b1bf-4542-9ca4-7563ad6b7b09';
  const PERSONNEL_LIST_ID = '67d0c026-90c1-4822-b27f-66f73c8139e5';
  
  useEffect(() => {
    const getToken = async () => {
      try {
        if (accounts.length > 0) {
          const response = await instance.acquireTokenSilent({
            scopes: authConfig.scopes,
            account: accounts[0]
          });
          
          setAccessToken(response.accessToken);
          setAuthError(null);
        } else if (inProgress === "none") {
          await instance.loginRedirect({
            scopes: authConfig.scopes
          });
        }
      } catch (error) {
        console.error('Token acquisition failed:', error);
        setAuthError(error instanceof Error ? error.message : 'Authentication failed');
        
        if (error instanceof InteractionRequiredAuthError) {
          try {
            await instance.loginRedirect({
              scopes: authConfig.scopes
            });
          } catch (redirectError) {
            console.error('Redirect failed:', redirectError);
          }
        }
      }
    };

    getToken();
  }, [accounts, inProgress, instance]);

  // ... rest of authentication logic
};
```

### 4. Dependencies to Install for SharePoint Integration

```json
{
  "@azure/msal-browser": "^3.0.0",
  "@azure/msal-react": "^2.0.0"
}
```

### 5. Environment Variables for SharePoint

```env
REACT_APP_AAD_CLIENT_ID=your-client-id
REACT_APP_AAD_AUTHORITY=https://login.microsoftonline.com/your-tenant-id
REACT_APP_SHAREPOINT_SITE_ID=your-site-id
REACT_APP_SHAREPOINT_PERSONNEL_LIST_ID=your-list-id
```

## Key Integration Points

1. **Authentication**: Uses MSAL for Azure AD authentication
2. **Data Sync**: Bidirectional sync between local storage and SharePoint
3. **Permissions**: Requires Sites.Read.All and Sites.ReadWrite.All scopes
4. **List Structure**: SharePoint list should have columns: Name, Gruppe, Department, Kommentar, Verfügbar

## Implementation Notes

- SharePoint serves as backup/collaboration storage
- Local storage remains primary for immediate operations
- Manual sync controls for user-initiated data exchange
- Offline capability with local-first approach
- Error handling for network failures and permission issues
