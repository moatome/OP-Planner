// Simple Local Storage Personnel Management
import { useState, useEffect } from 'react';

export interface Personnel {
  id: number;
  name: string;
  gruppe: 'OP-Pflege' | 'Anästhesie Pflege' | 'OP-Praktikant' | 'Anästhesie Praktikant' | 'MFA' | 'ATA Schüler' | 'OTA Schüler';
  department: string;
  kommentar: string;
  verfügbar: 'Bereitschaften (BD)' | 'Rufdienste (RD)' | 'Frühdienste (Früh)' | 'Zwischendienste/Mitteldienste (Mittel)' | 'Spätdienste (Spät)' | 'nicht verfügbar';
  initials: string;
  isActive: boolean;
  // Excel import related fields
  shiftAssignment?: string;
  availabilityTags?: string[];
  shiftTags?: string[];
  isAvailable?: boolean;
}

const STORAGE_KEY = 'op-planner-personnel';

export const useLocalPersonnelData = () => {
  const [personnel, setPersonnel] = useState<Personnel[]>([]);
  const [loading, setLoading] = useState(true);

  // Load initial data from localStorage
  useEffect(() => {
    const loadLocalData = () => {
      setLoading(true);
      
      const localData = localStorage.getItem(STORAGE_KEY);
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
    if (!loading) { // Don't save during initial load
      localStorage.setItem(STORAGE_KEY, JSON.stringify(personnel));
      console.log('Saved personnel to localStorage:', personnel.length, 'items');
    }
  }, [personnel, loading]);

  // Add personnel function
  const addPersonnel = async (person: Omit<Personnel, 'id'>): Promise<Personnel> => {
    const newPerson: Personnel = {
      ...person,
      id: Date.now() + Math.floor(Math.random() * 1000), // Generate unique local ID
    };

    setPersonnel(prev => [...prev, newPerson]);
    return newPerson;
  };

  // Update personnel function
  const updatePersonnel = async (id: number, updates: Partial<Personnel>): Promise<boolean> => {
    setPersonnel(prev => prev.map(person => 
      person.id === id ? { ...person, ...updates } : person
    ));
    return true;
  };

  // Delete personnel function
  const deletePersonnel = async (id: number): Promise<boolean> => {
    setPersonnel(prev => prev.filter(person => person.id !== id));
    return true;
  };

  // Generate initials helper
  const generateInitials = (name: string): string => {
    if (!name) return 'XX';
    const parts = name.trim().split(' ').filter(part => part.length > 0);
    
    if (parts.length === 0) return 'XX';
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  };

  // Export to Excel function
  const exportToExcel = () => {
    // Create CSV content
    const headers = ['Name', 'Gruppe', 'Abteilung', 'Verfügbarkeit', 'Schichtzuweisung', 'Kürzel', 'Kommentar'];
    const csvContent = [
      headers.join(','),
      ...personnel.map(person => [
        `"${person.name}"`,
        `"${person.gruppe}"`,
        `"${person.department}"`,
        `"${person.verfügbar}"`,
        `"${person.shiftAssignment || ''}"`,
        `"${person.initials}"`,
        `"${person.kommentar}"`
      ].join(','))
    ].join('\n');

    // Create and download file
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', `personal-export-${new Date().toISOString().split('T')[0]}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  // Clear all data function
  const clearAllData = async (): Promise<boolean> => {
    if (window.confirm('Sind Sie sicher, dass Sie alle Personaldaten löschen möchten? Diese Aktion kann nicht rückgängig gemacht werden.')) {
      localStorage.removeItem(STORAGE_KEY);
      localStorage.removeItem('op-planner-assignments');
      localStorage.removeItem('op-planner-availability-tags');
      setPersonnel([]);
      return true;
    }
    return false;
  };

  return {
    personnel,
    loading,
    addPersonnel,
    updatePersonnel,
    deletePersonnel,
    generateInitials,
    exportToExcel,
    clearAllData
  };
};
