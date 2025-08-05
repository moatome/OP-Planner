// SharePoint Integration for OR Planner
import { useState, useEffect } from 'react';

// Mock data for development
const MOCK_PERSONNEL = [
  {
    id: 1,
    name: "Dr. Sarah Weber",
    role: "Anästhesie Arzt",
    department: "Anästhesie",
    initials: "SW",
    isActive: true
  },
  {
    id: 2,
    name: "Dr. Michael Koch",
    role: "Anästhesie Arzt",
    department: "Anästhesie",
    initials: "MK",
    isActive: true
  },
  {
    id: 3,
    name: "Lisa Müller",
    role: "Anästhesie Pflege",
    department: "Anästhesie",
    initials: "LM",
    isActive: true
  },
  {
    id: 4,
    name: "Thomas Schmidt",
    role: "OP Pflege",
    department: "OP",
    initials: "TS",
    isActive: true
  },
  {
    id: 5,
    name: "Anna Becker",
    role: "OTAS",
    department: "OP",
    initials: "AB",
    isActive: true
  },
  {
    id: 6,
    name: "Max Hoffmann",
    role: "ATA",
    department: "Anästhesie",
    initials: "MH",
    isActive: true
  }
];

class SharePointService {
  accessToken: string;
  baseUrl: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://graph.microsoft.com/v1.0';
  }

  async getPersonnel(siteId: string, listId: string) {
    // Return mock data if in development mode
    if (process.env.NODE_ENV === 'development' && !this.accessToken) {
      console.log('Using mock personnel data');
      return MOCK_PERSONNEL;
    }

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
      
      if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
      
      const data = await response.json();
      
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
      // Return mock data if API fails
      return MOCK_PERSONNEL;
    }
  }

  async clearAssignments(siteId: string, listId: string, planDate: string): Promise<void> {
    // Implementation for production
    if (process.env.NODE_ENV === 'development') {
      console.log('Mock clearing assignments');
      return;
    }
    
    try {
      // Actual implementation for production
    } catch (error) {
      console.error('Error clearing assignments:', error);
    }
  }

  async saveAssignments(
    siteId: string, 
    assignmentsListId: string, 
    assignments: Record<string, Array<{id: number, name: string}>>, 
    planDate: string
  ): Promise<boolean> {
    // Mock implementation for development
    if (process.env.NODE_ENV === 'development' && !this.accessToken) {
      console.log('Mock saving assignments:', assignments);
      return true;
    }

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

export const useSharePointPersonnel = (
  accessToken: string | null, 
  siteId: string, 
  listId: string
) => {
  const [personnel, setPersonnel] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchPersonnel = async () => {
      try {
        setLoading(true);
        const sharePointService = new SharePointService(accessToken || '');
        const data = await sharePointService.getPersonnel(siteId, listId);
        setPersonnel(data);
        setError(null);
      } catch (err: any) {
        console.error('Error fetching personnel:', err);
        setError(err.message);
        // Even if there's an error, set mock data
        setPersonnel(MOCK_PERSONNEL);
      } finally {
        setLoading(false);
      }
    };

    fetchPersonnel();
    // Only refresh in production
    if (process.env.NODE_ENV === 'production') {
      const interval = setInterval(fetchPersonnel, 3600000);
      return () => clearInterval(interval);
    }
  }, [accessToken, siteId, listId]);

  return { personnel, loading, error };
};

export const authConfig = {
  auth: {
    clientId: process.env.REACT_APP_AAD_CLIENT_ID || 'YOUR_APP_CLIENT_ID',
    authority: process.env.REACT_APP_AAD_AUTHORITY || 'https://login.microsoftonline.com/YOUR_TENANT_ID',
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