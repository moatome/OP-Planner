// SharePoint Integration for OR Planner
import { useState, useEffect } from 'react';

// 1. Using Microsoft Graph API to fetch SharePoint list data
class SharePointService {
  accessToken: string;
  baseUrl: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://graph.microsoft.com/v1.0';
  }

  // Fetch personnel from your SharePoint list
  async getPersonnel(siteId: string, listId: string) {
    try {
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items?expand=fields`,
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );
      
      const data = await response.json();
      
      // Transform SharePoint data to your app format
      return data.value.map((item: any) => ({
        id: item.id,
        name: item.fields.DisplayName || item.fields.Title,
        role: item.fields.JobTitle,
        department: item.fields.Department,
        email: item.fields.Email,
        profilePicture: item.fields.Picture,
        initials: this.generateInitials(item.fields.DisplayName),
        isActive: item.fields.IsActive !== false
      })).filter((person: any) => person.isActive);
      
    } catch (error) {
      console.error('Error fetching personnel:', error);
      return [];
    }
  }

  // Clear assignments helper method
  async clearAssignments(siteId: string, listId: string, planDate: string): Promise<void> {
    // Implementation goes here
  }

  // Save assignments back to SharePoint
  async saveAssignments(
    siteId: string, 
    assignmentsListId: string, 
    assignments: Record<string, Array<{id: number, name: string}>>, 
    planDate: string
  ): Promise<boolean> {
    try {
      const assignmentItems = Object.entries(assignments).flatMap(([cellKey, persons]) => 
        persons.map(person => ({
          fields: {
            PersonId: person.id,
            PersonName: person.name,
            CellPosition: cellKey,
            PlanDate: planDate,
            Room: cellKey.split('-')[1],
            Role: cellKey.split('-')[0]
          }
        })));

      await this.clearAssignments(siteId, assignmentsListId, planDate);

      const promises = assignmentItems.map(item =>
        fetch(`${this.baseUrl}/sites/${siteId}/lists/${assignmentsListId}/items`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(item)
        })
      );

      await Promise.all(promises);
      return true;
    } catch (error) {
      console.error('Error saving assignments:', error);
      return false;
    }
  }

  generateInitials(name: string): string {
    if (!name) return '';
    const parts = name.split(' ');
    return parts.length > 1 
      ? (parts[0][0] + parts[parts.length - 1][0]).toUpperCase()
      : parts[0].substring(0, 2).toUpperCase();
  }
}

// 2. React Hook for SharePoint data

export const useSharePointPersonnel = (
  accessToken: string | null, 
  siteId: string, 
  listId: string
) => {
  const [personnel, setPersonnel] = useState<Array<{
    id: number;
    name: string;
    role: string;
    department: string;
    email?: string;
    profilePicture?: string;
    initials: string;
    isActive: boolean;
  }>>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchPersonnel = async () => {
      if (!accessToken) {
        console.log('No access token available');
        return
      }
      
      try {
        console.log('Fetching personnel data...');
        const sharePointService = new SharePointService(accessToken);
        const data = await sharePointService.getPersonnel(siteId, listId);
        console.log('Personnel data received:', data);
        setPersonnel(data);
        setError(null);
      } catch (err: any) {
        console.error('Error fetching personnel:', err);
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    fetchPersonnel();
    const interval = setInterval(fetchPersonnel, 3600000);
    return () => clearInterval(interval);
  }, [accessToken, siteId, listId]);

  return { personnel, loading, error };
};

// 3. Authentication setup (for Teams)
export const authConfig = {
  auth: {
    clientId: 'YOUR_APP_CLIENT_ID',
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID',
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
  scopes: [
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
  ]
};

// 4. Usage in your React component (remove this section - it should be in your App.tsx)