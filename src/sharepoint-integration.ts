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
  sharePointId?: number; // Track SharePoint ID for sync operations
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

  // Sync local personnel changes to SharePoint
  async syncPersonnelToSharePoint(
    siteId: string, 
    listId: string, 
    localPersonnel: PersonnelItem[]
  ): Promise<PersonnelItem[]> {
    if (!this.accessToken) {
      console.log('No access token available for sync');
      return localPersonnel;
    }

    console.log('Starting personnel sync to SharePoint...');
    const syncedPersonnel: PersonnelItem[] = [];

    for (const person of localPersonnel) {
      try {
        if (person.isLocallyAdded && !person.sharePointId) {
          // Add new person to SharePoint
          console.log('Adding new person to SharePoint:', person.name);
          const addedPerson = await this.addPersonnel(siteId, listId, {
            name: person.name,
            gruppe: person.gruppe,
            department: person.department,
            kommentar: person.kommentar,
            verfügbar: person.verfügbar,
            initials: person.initials,
            isActive: person.isActive
          });

          if (addedPerson) {
            syncedPersonnel.push({
              ...person,
              sharePointId: addedPerson.id,
              isLocallyAdded: false // Remove local flag after successful sync
            });
          } else {
            // Keep local flag if sync failed
            syncedPersonnel.push(person);
          }
        } else if (person.isLocallyModified && person.sharePointId) {
          // Update existing person in SharePoint
          console.log('Updating person in SharePoint:', person.name);
          const updatedPerson = await this.updatePersonnel(siteId, listId, person.sharePointId, {
            name: person.name,
            gruppe: person.gruppe,
            department: person.department,
            kommentar: person.kommentar
            // Note: verfügbar, initials, isActive are handled locally only
          });

          if (updatedPerson) {
            syncedPersonnel.push({
              ...person,
              isLocallyModified: false // Remove local flag after successful sync
            });
          } else {
            // Keep local flag if sync failed
            syncedPersonnel.push(person);
          }
        } else {
          // No changes needed
          syncedPersonnel.push(person);
        }
      } catch (error) {
        console.error('Error syncing person:', person.name, error);
        // Keep person as-is if sync failed
        syncedPersonnel.push(person);
      }
    }

    console.log('Personnel sync completed');
    return syncedPersonnel;
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

  // Personnel CRUD Operations for Phase 2
  async addPersonnel(siteId: string, listId: string, personnel: Omit<PersonnelItem, 'id'>): Promise<PersonnelItem | null> {
    if (!siteId || !listId) {
      console.error('Missing SharePoint configuration for personnel list');
      return null;
    }

    if (!this.accessToken) {
      console.error('No access token available for adding personnel');
      return null;
    }

    try {
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            fields: {
              Title: personnel.name,
              Name: personnel.name,
              Gruppe: personnel.gruppe,
              Department: personnel.department,
              Kommentar: personnel.kommentar
              // Note: verfügbar, initials, isActive are handled locally only
            }
          })
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Failed to add personnel:', errorText);
        return null;
      }

      const data = await response.json();
      console.log('SharePoint add response:', data);
      
      const newId = parseInt(data.id) || data.id;
      const result = {
        id: newId,
        ...personnel,
        sharePointId: newId
      };
      
      console.log('Returning from addPersonnel:', result);
      return result;
    } catch (error) {
      console.error('Error adding personnel:', error);
      return null;
    }
  }

  async updatePersonnel(siteId: string, listId: string, id: number, updates: Partial<PersonnelItem>): Promise<PersonnelItem | null> {
    if (!siteId || !listId) {
      console.error('Missing SharePoint configuration for personnel list');
      return null;
    }

    if (!this.accessToken) {
      console.error('No access token available for updating personnel');
      return null;
    }

    try {
      const fieldsToUpdate: any = {};
      
      // Only update SharePoint fields that exist in the list
      if (updates.name) {
        fieldsToUpdate.Title = updates.name;
        fieldsToUpdate.Name = updates.name;
      }
      if (updates.gruppe) fieldsToUpdate.Gruppe = updates.gruppe;
      if (updates.department) fieldsToUpdate.Department = updates.department;
      if (updates.kommentar !== undefined) fieldsToUpdate.Kommentar = updates.kommentar;
      // Note: verfügbar, initials, isActive are handled locally only

      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items/${id}`,
        {
          method: 'PATCH',
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'If-Match': '*'
          },
          body: JSON.stringify({
            fields: fieldsToUpdate
          })
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Failed to update personnel:', errorText);
        return null;
      }

      // Return updated item by fetching it again
      return await this.getPersonnelById(siteId, listId, id);
    } catch (error) {
      console.error('Error updating personnel:', error);
      return null;
    }
  }

  async deletePersonnel(siteId: string, listId: string, id: number): Promise<boolean> {
    if (!siteId || !listId) {
      console.error('Missing SharePoint configuration for personnel list');
      return false;
    }

    if (!this.accessToken) {
      console.error('No access token available for deleting personnel');
      return false;
    }

    try {
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items/${id}`,
        {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'If-Match': '*'
          }
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Failed to delete personnel:', errorText);
        return false;
      }

      return true;
    } catch (error) {
      console.error('Error deleting personnel:', error);
      return false;
    }
  }

  async getPersonnelById(siteId: string, listId: string, id: number): Promise<PersonnelItem | null> {
    if (!siteId || !listId) {
      console.error('Missing SharePoint configuration for personnel list');
      return null;
    }

    if (!this.accessToken) {
      console.error('No access token available for fetching personnel');
      return null;
    }

    try {
      const response = await fetch(
        `${this.baseUrl}/sites/${siteId}/lists/${listId}/items/${id}?expand=fields`,
        {
          headers: {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          }
        }
      );

      if (!response.ok) {
        console.error('Failed to fetch personnel by ID:', response.statusText);
        return null;
      }

      const data = await response.json();
      const fields = data.fields || {};

      return {
        id: data.id,
        name: fields.Name || fields.DisplayName || fields.Title || 'Unknown',
        gruppe: fields.Gruppe || 'OP-Pflege',
        department: fields.Department || 'Unknown Department',
        kommentar: fields.Kommentar || '',
        verfügbar: fields.Verfügbar || 'nicht verfügbar',
        initials: fields.Initials || this.generateInitials(fields.Name || fields.DisplayName || fields.Title || 'XX'),
        isActive: fields.IsActive !== undefined ? fields.IsActive : true
      };
    } catch (error) {
      console.error('Error fetching personnel by ID:', error);
      return null;
    }
  }

  async getPersonnelByGroup(siteId: string, listId: string, gruppe: PersonnelItem['gruppe']): Promise<PersonnelItem[]> {
    if (!siteId || !listId) {
      console.error('Missing SharePoint configuration for personnel list');
      return [];
    }

    const allPersonnel = await this.getPersonnel(siteId, listId);
    return allPersonnel.filter(person => person.gruppe === gruppe);
  }
}

export const useSharePointPersonnel = (
  accessToken: string | null, 
  siteId: string, 
  listId: string,
  autoFetch: boolean = false // Only fetch when explicitly requested
) => {
  const [personnel, setPersonnel] = useState<PersonnelItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

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
        console.log('No access token available');
        setPersonnel([]);
        return [];
      }

      const sharePointService = new SharePointService(accessToken);
      const data = await sharePointService.getPersonnel(siteId, listId);
      
      console.log('Personnel fetched successfully:', data.length, 'items');
      setPersonnel(data);
      return data;
    } catch (err: any) {
      console.error('Error in useSharePointPersonnel:', err);
      setError(err.message || 'Unknown error occurred');
      // Return empty array on error
      setPersonnel([]);
      return [];
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    // Only auto-fetch if explicitly enabled
    if (autoFetch && accessToken) {
      fetchPersonnel();
    }
  }, [accessToken, siteId, listId, autoFetch]);

  return { personnel, loading, error, fetchPersonnel };
};

// Primary local storage-based hook with optional SharePoint sync
export const useHybridPersonnelData = (
  accessToken: string | null,
  siteId: string,
  listId: string
) => {
  const { fetchPersonnel: fetchFromSharePoint } = useSharePointPersonnel(accessToken, siteId, listId, false);
  const [personnel, setPersonnel] = useState<PersonnelItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState<Date | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Load initial data from localStorage only
  useEffect(() => {
    const loadLocalData = () => {
      setLoading(true);
      
      const localData = localStorage.getItem('op-planner-personnel');
      if (localData) {
        try {
          const parsedData = JSON.parse(localData);
          console.log('Loaded personnel from localStorage:', parsedData.length, 'items');
          setPersonnel(parsedData);
        } catch (error) {
          console.error('Error parsing local personnel data:', error);
          setPersonnel([]);
        }
      } else {
        console.log('No local personnel data found, starting with empty array');
        setPersonnel([]);
      }
      
      setLoading(false);
    };

    loadLocalData();
  }, []);

  // Save to localStorage whenever personnel changes
  useEffect(() => {
    if (personnel.length > 0) {
      localStorage.setItem('op-planner-personnel', JSON.stringify(personnel));
      console.log('Saved personnel to localStorage:', personnel.length, 'items');
    }
  }, [personnel]);

  // Manual sync TO SharePoint (push local changes)
  const syncToSharePoint = async (): Promise<boolean> => {
    if (!accessToken) {
      console.log('Cannot sync to SharePoint: no access token');
      return false;
    }

    try {
      setLoading(true);
      setError(null);
      const sharePointService = new SharePointService(accessToken);
      
      console.log('Syncing to SharePoint:', personnel.length, 'personnel items');
      
      for (const person of personnel) {
        try {
          if (person.isLocallyAdded && !person.sharePointId) {
            // Add new person to SharePoint
            console.log('Adding new person to SharePoint:', person.name);
            const addedPerson = await sharePointService.addPersonnel(siteId, listId, {
              name: person.name,
              gruppe: person.gruppe,
              department: person.department,
              kommentar: person.kommentar,
              verfügbar: person.verfügbar,
              initials: person.initials,
              isActive: person.isActive
            });

            if (addedPerson) {
              // Update local person with SharePoint ID
              person.sharePointId = addedPerson.id;
              person.isLocallyAdded = false;
            }
          } else if (person.isLocallyModified && person.sharePointId) {
            // Update existing person in SharePoint
            console.log('Updating person in SharePoint:', person.name);
            await sharePointService.updatePersonnel(siteId, listId, person.sharePointId, {
              name: person.name,
              gruppe: person.gruppe,
              department: person.department,
              kommentar: person.kommentar
              // Note: verfügbar, initials, isActive are handled locally only
            });
            
            person.isLocallyModified = false;
          }
        } catch (error) {
          console.error('Error syncing person:', person.name, error);
        }
      }

      // Update personnel state and localStorage
      setPersonnel([...personnel]);
      setLastSyncTime(new Date());
      console.log('Sync to SharePoint completed');
      
      return true;
    } catch (error) {
      console.error('Sync to SharePoint failed:', error);
      setError(error instanceof Error ? error.message : 'Sync failed');
      return false;
    } finally {
      setLoading(false);
    }
  };

  // Manual sync FROM SharePoint (pull SharePoint data and merge)
  const syncFromSharePoint = async (): Promise<boolean> => {
    if (!accessToken) {
      console.log('Cannot sync from SharePoint: no access token');
      return false;
    }

    try {
      setLoading(true);
      setError(null);
      
      console.log('Syncing FROM SharePoint...');
      const sharePointData = await fetchFromSharePoint();
      
      if (sharePointData && sharePointData.length > 0) {
        // Merge SharePoint data with local data
        const merged = new Map<number, PersonnelItem>();
        
        // First add all local personnel
        personnel.forEach(person => {
          merged.set(person.id, person);
        });
        
        // Then merge SharePoint personnel (only for fields that should come from SharePoint)
        sharePointData.forEach(spPerson => {
          const existingLocal = personnel.find(local => 
            local.sharePointId === spPerson.id || 
            (local.name === spPerson.name && local.gruppe === spPerson.gruppe)
          );
          
          if (existingLocal) {
            // Update SharePoint-managed fields only
            merged.set(existingLocal.id, {
              ...existingLocal,
              name: spPerson.name,
              gruppe: spPerson.gruppe,
              department: spPerson.department,
              kommentar: spPerson.kommentar,
              sharePointId: spPerson.id,
              // Keep local-only fields intact
              verfügbar: existingLocal.verfügbar,
              isActive: existingLocal.isActive,
              shiftAssignment: existingLocal.shiftAssignment,
              availabilityTags: existingLocal.availabilityTags,
              shiftTags: existingLocal.shiftTags,
              isAvailable: existingLocal.isAvailable
            });
          } else {
            // Add new person from SharePoint
            merged.set(spPerson.id, {
              ...spPerson,
              sharePointId: spPerson.id
            });
          }
        });
        
        const mergedArray = Array.from(merged.values());
        setPersonnel(mergedArray);
        setLastSyncTime(new Date());
        
        console.log('Sync FROM SharePoint completed:', mergedArray.length, 'items');
        return true;
      } else {
        console.log('No data received from SharePoint');
        return false;
      }
    } catch (error) {
      console.error('Sync from SharePoint failed:', error);
      setError(error instanceof Error ? error.message : 'Sync failed');
      return false;
    } finally {
      setLoading(false);
    }
  };

  // Add personnel function
  const addPersonnel = async (person: Omit<PersonnelItem, 'id'>, syncImmediately = false): Promise<PersonnelItem> => {
    const newPerson: PersonnelItem = {
      ...person,
      id: Date.now() + Math.floor(Math.random() * 1000), // Generate unique local ID
      isLocallyAdded: true
    };

    const updated = [...personnel, newPerson];
    setPersonnel(updated);

    if (syncImmediately && accessToken) {
      console.log('Syncing new person to SharePoint immediately...');
      await syncToSharePoint();
    }

    return newPerson;
  };

  // Update personnel function
  const updatePersonnel = async (id: number, updates: Partial<PersonnelItem>, syncImmediately = false): Promise<boolean> => {
    const updated = personnel.map(person => {
      if (person.id === id) {
        const updatedPerson = { 
          ...person, 
          ...updates,
          isLocallyModified: !person.isLocallyAdded // Don't mark as modified if it's locally added
        };
        return updatedPerson;
      }
      return person;
    });

    setPersonnel(updated);

    if (syncImmediately && accessToken) {
      return await syncToSharePoint();
    }

    return true;
  };

  // Delete personnel function
  const deletePersonnel = async (id: number, syncImmediately = false): Promise<boolean> => {
    const personToDelete = personnel.find(p => p.id === id);
    
    if (personToDelete?.sharePointId && accessToken && syncImmediately) {
      // Delete from SharePoint if it exists there
      try {
        const sharePointService = new SharePointService(accessToken);
        await sharePointService.deletePersonnel(siteId, listId, personToDelete.sharePointId);
      } catch (error) {
        console.error('Failed to delete from SharePoint:', error);
        return false;
      }
    }

    const updated = personnel.filter(person => person.id !== id);
    setPersonnel(updated);

    return true;
  };

  return {
    personnel,
    loading,
    error,
    lastSyncTime,
    syncToSharePoint,
    syncFromSharePoint, // New function for explicit pull from SharePoint
    addPersonnel,
    updatePersonnel,
    deletePersonnel,
    hasUnsyncedChanges: personnel.some(p => p.isLocallyAdded || p.isLocallyModified)
  };
};

export const useSharePointAssignments = (
  accessToken: string | null,
  siteId: string,
  assignmentsListId: string
) => {
  const sharePointService = new SharePointService(accessToken || '');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const saveAssignments = async (
    assignments: Record<string, Array<{id: number, name: string}>>,
    planDate: string
  ): Promise<boolean> => {
    if (!accessToken) {
      console.log('No access token, saving assignments locally only');
      localStorage.setItem(`op-planner-assignments-${planDate}`, JSON.stringify(assignments));
      return true;
    }

    if (!assignmentsListId || assignmentsListId.trim() === '') {
      console.log('No assignments list ID configured, saving locally only');
      localStorage.setItem(`op-planner-assignments-${planDate}`, JSON.stringify(assignments));
      localStorage.setItem('op-planner-assignments', JSON.stringify(assignments));
      return true;
    }

    setLoading(true);
    setError(null);

    try {
      // Save to SharePoint
      const success = await sharePointService.saveAssignments(siteId, assignmentsListId, assignments, planDate);
      
      if (success) {
        // Also save locally as backup
        localStorage.setItem(`op-planner-assignments-${planDate}`, JSON.stringify(assignments));
        localStorage.setItem('op-planner-assignments', JSON.stringify(assignments)); // Current assignments
      }
      
      return success;
    } catch (err: any) {
      console.error('Failed to save assignments to SharePoint:', err);
      setError(err.message || 'Failed to save assignments');
      
      // Fallback to local storage
      localStorage.setItem(`op-planner-assignments-${planDate}`, JSON.stringify(assignments));
      localStorage.setItem('op-planner-assignments', JSON.stringify(assignments));
      
      return true; // Return true since we saved locally
    } finally {
      setLoading(false);
    }
  };

  const loadAssignments = async (
    planDate: string
  ): Promise<Record<string, Array<{id: number, name: string}>>> => {
    if (!accessToken) {
      console.log('No access token, loading assignments from local storage only');
      const local = localStorage.getItem(`op-planner-assignments-${planDate}`);
      return local ? JSON.parse(local) : {};
    }

    if (!assignmentsListId || assignmentsListId.trim() === '') {
      console.log('No assignments list ID configured, using localStorage only');
      const local = localStorage.getItem(`op-planner-assignments-${planDate}`);
      return local ? JSON.parse(local) : {};
    }

    setLoading(true);
    setError(null);

    try {
      // Try to load from SharePoint first
      const assignments = await sharePointService.loadAssignments(siteId, assignmentsListId, planDate);
      
      // Save to local storage as backup
      localStorage.setItem(`op-planner-assignments-${planDate}`, JSON.stringify(assignments));
      localStorage.setItem('op-planner-assignments', JSON.stringify(assignments));
      
      return assignments;
    } catch (err: any) {
      console.error('Failed to load assignments from SharePoint:', err);
      setError(err.message || 'Failed to load assignments');
      
      // Fallback to local storage
      const local = localStorage.getItem(`op-planner-assignments-${planDate}`);
      return local ? JSON.parse(local) : {};
    } finally {
      setLoading(false);
    }
  };

  const syncAssignments = async (
    localAssignments: Record<string, Array<{id: number, name: string}>>,
    planDate: string
  ): Promise<boolean> => {
    if (!assignmentsListId || assignmentsListId.trim() === '') {
      console.log('No assignments list ID configured, sync skipped');
      return true; // Consider it successful since we're not configured for SharePoint assignments
    }
    return await saveAssignments(localAssignments, planDate);
  };

  return { 
    saveAssignments, 
    loadAssignments, 
    syncAssignments,
    loading,
    error
  };
};

// Updated auth config - make sure this matches your app registration
export const authConfig = {
  auth: {
    clientId: import.meta.env.VITE_AAD_CLIENT_ID || '06c5c649-973a-49a0-ba36-56ecf11285f1',
    authority: import.meta.env.VITE_AAD_AUTHORITY || 'https://login.microsoftonline.com/d0c4995a-6bf2-4d26-9281-906c0c59b9cb',
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