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
  },
  {
    id: 7,
    name: "Julia Wagner",
    role: "Praktikant",
    department: "OP",
    initials: "JW",
    isActive: true
  },
  {
    id: 8,
    name: "Daniel Richter",
    role: "AA Praktikant",
    department: "Anästhesie",
    initials: "DR",
    isActive: true
  },
  {
    id: 9,
    name: "Dr. Peter Neumann",
    role: "Anästhesie Arzt",
    department: "Anästhesie",
    initials: "PN",
    isActive: true
  },
  {
    id: 10,
    name: "Maria Schulz",
    role: "OP Pflege",
    department: "OP",
    initials: "MS",
    isActive: true
  }
];

interface PersonnelItem {
  id: number;
  name: string;
  role: string;
  department: string;
  initials: string;
  isActive: boolean;
  email?: string;
  profilePicture?: string;
}

class SharePointService {
  accessToken: string;
  baseUrl: string;

  constructor(accessToken: string) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://graph.microsoft.com/v1.0';
  }

  async getPersonnel(siteId: string, listId: string): Promise<PersonnelItem[]> {
    // Return mock data if no access token
    if (!this.accessToken) {
      console.log('Using mock personnel data (no token)');
      return MOCK_PERSONNEL;
    }

    try {
      console.log('Fetching personnel from SharePoint...');
      
      // Get the list items with the direct fields populated by Power Automate
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*,DisplayName,Department,JobTitle,Email)`,
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
        
        console.log('Falling back to mock data due to API error');
        return MOCK_PERSONNEL;
      }
      
      const data = await response.json();
      console.log('SharePoint data:', data);
      
      if (!data.value || data.value.length === 0) {
        console.log('No data in SharePoint list, using mock data');
        return MOCK_PERSONNEL;
      }
      
      const personnel: PersonnelItem[] = [];
      
      // Process each item - now reading directly from the dedicated fields
      for (const item of data.value) {
        try {
          const fields = item.fields || {};
          
          // Read values directly from the dedicated fields populated by Power Automate
          const displayName = fields.DisplayName || fields.Title || 'Unknown';
          const department = fields.Department || 'Unknown Department';
          const jobTitle = fields.JobTitle || 'Unknown Role';
          const email = fields.Email || '';
          
          const personnelItem: PersonnelItem = {
            id: parseInt(item.id) || Math.random() * 10000,
            name: displayName,
            role: jobTitle,
            department: department,
            email: email,
            profilePicture: '', // Profile pictures could be added as another field if needed
            initials: this.generateInitials(displayName),
            isActive: true
          };
          
          personnel.push(personnelItem);
        } catch (itemError) {
          console.error('Error processing item:', itemError, item);
        }
      }

      console.log('Processed personnel:', personnel);
      return personnel.length > 0 ? personnel : MOCK_PERSONNEL;
      
    } catch (error) {
      console.error('Error fetching personnel from SharePoint:', error);
      console.log('Falling back to mock data');
      return MOCK_PERSONNEL;
    }
  }

  async clearAssignments(siteId: string, listId: string, planDate: string): Promise<void> {
    if (!this.accessToken) {
      console.log('No access token available for clearing assignments', planDate);
      return;
    }
    
    try {
      // Get existing assignments for the date
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items?$filter=PlanDate eq '${planDate}'&expand=fields`,
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        
        // Delete each existing assignment
        const deletePromises = data.value.map((item: any) =>
          fetch(`${this.baseUrl}/sites/${siteId}/lists/${listId}/items/${item.id}`, {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${this.accessToken}`,
              'If-Match': '*'
            }
          })
        );

        await Promise.all(deletePromises);
        console.log(`Cleared ${data.value.length} existing assignments for ${planDate}`);
      }
    } catch (error) {
      console.error('Error clearing assignments:', error);
      throw error;
    }
  }

  async saveAssignments(
    siteId: string, 
    assignmentsListId: string, 
    assignments: Record<string, Array<{id: number, name: string}>>, 
    planDate: string
  ): Promise<boolean> {
    if (!this.accessToken) {
      throw new Error('No access token available');
    }

    try {
      // Parse cell keys to extract meaningful data
      const assignmentItems = Object.entries(assignments).flatMap(([cellKey, persons]) => {
        const [tableType, roleIndex, roomIndex] = cellKey.split('-');
        
        return persons.map(person => ({
          fields: {
            Title: `${person.name} - ${cellKey}`,
            PersonId: person.id.toString(),
            PersonName: person.name,
            CellPosition: cellKey,
            PlanDate: planDate,
            TableType: tableType,
            RoleIndex: parseInt(roleIndex) || 0,
            RoomIndex: parseInt(roomIndex) || 0,
            CreatedDate: new Date().toISOString()
          }
        }));
      });

      if (assignmentItems.length === 0) {
        console.log('No assignments to save');
        return true;
      }

      // Clear existing assignments first
      await this.clearAssignments(siteId, assignmentsListId, planDate);

      // Create new assignments
      const promises = assignmentItems.map(item =>
        fetch(`${this.baseUrl}/sites/${siteId}/lists/${assignmentsListId}/items`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(item)
        }).then(response => {
          if (!response.ok) {
            return response.text().then(text => {
              throw new Error(`Failed to save assignment: ${response.status} - ${text}`);
            });
          }
          return response.json();
        })
      );

      await Promise.all(promises);
      console.log(`Successfully saved ${assignmentItems.length} assignments`);
      return true;
    } catch (error) {
      console.error('Error saving assignments:', error);
      return false;
    }
  }

  async loadAssignments(
    siteId: string,
    assignmentsListId: string,
    planDate: string
  ): Promise<Record<string, Array<{id: number, name: string}>>> {
    if (!this.accessToken) {
      throw new Error('No access token available');
    }

    try {
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${assignmentsListId}/items?$filter=PlanDate eq '${planDate}'&expand=fields`,
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const assignments: Record<string, Array<{id: number, name: string}>> = {};

      data.value.forEach((item: any) => {
        const cellKey = item.fields.CellPosition;
        const person = {
          id: parseInt(item.fields.PersonId) || 0,
          name: item.fields.PersonName || 'Unknown'
        };

        if (!assignments[cellKey]) {
          assignments[cellKey] = [];
        }
        assignments[cellKey].push(person);
      });

      console.log('Loaded assignments:', assignments);
      return assignments;
    } catch (error) {
      console.error('Error loading assignments:', error);
      throw error;
    }
  }

  generateInitials(name: string): string {
    if (!name) return 'XX';
    const parts = name.trim().split(' ').filter(part => part.length > 0);
    
    if (parts.length === 0) return 'XX';
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
}

export const useSharePointPersonnel = (
  accessToken: string | null, 
  siteId: string, 
  listId: string
) => {
  const [personnel, setPersonnel] = useState<PersonnelItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchPersonnel = async () => {
      try {
        setLoading(true);
        setError(null);
        
        console.log('Starting personnel fetch...', {
          hasToken: !!accessToken,
          tokenLength: accessToken?.length,
          siteId,
          listId
        });

        if (!accessToken) {
          console.log('No access token, using mock data');
          setPersonnel(MOCK_PERSONNEL);
          return;
        }

        const sharePointService = new SharePointService(accessToken);
        const data = await sharePointService.getPersonnel(siteId, listId);
        
        console.log('Personnel fetched successfully:', data.length, 'items');
        setPersonnel(data);
      } catch (err: any) {
        console.error('Error in useSharePointPersonnel:', err);
        setError(err.message || 'Unknown error occurred');
        // Fallback to mock data on error
        setPersonnel(MOCK_PERSONNEL);
      } finally {
        setLoading(false);
      }
    };

    fetchPersonnel();
    
    // Only set up auto-refresh in production and with valid token
    if (accessToken) {
      const interval = setInterval(fetchPersonnel, 3600000); // 1 hour
      return () => clearInterval(interval);
    }
  }, [accessToken, siteId, listId]);

  return { personnel, loading, error };
};

export const useSharePointAssignments = (
  accessToken: string | null,
  siteId: string,
  assignmentsListId: string
) => {
  const sharePointService = new SharePointService(accessToken || '');

  const saveAssignments = async (
    assignments: Record<string, Array<{id: number, name: string}>>,
    planDate: string
  ): Promise<boolean> => {
    return await sharePointService.saveAssignments(siteId, assignmentsListId, assignments, planDate);
  };

  const loadAssignments = async (
    planDate: string
  ): Promise<Record<string, Array<{id: number, name: string}>>> => {
    return await sharePointService.loadAssignments(siteId, assignmentsListId, planDate);
  };

  return { saveAssignments, loadAssignments };
};

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
 
  // Updated scopes based on your app registration permissions
  scopes: [
    'Sites.Read.All',
    'Sites.ReadWrite.All',
    'User.Read',
    // Add SharePoint specific scopes if needed
    // 'https://graph.microsoft.com/.default' // This will use all consented permissions
  ]
};