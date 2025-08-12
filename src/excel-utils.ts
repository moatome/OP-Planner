// Excel Processing Utilities for Shift Plan Import
import * as XLSX from 'xlsx';

export interface ShiftAssignment {
  name: string;
  lastName: string;
  firstName: string;
  shiftType: string;
  availability: string;
  originalText?: string;
}

export interface PersonnelAvailabilityTags {
  [personName: string]: string[]; // Array of availability tags for each person
}

export interface ExcelPersonnelData {
  name: string;
  gruppe: string;
  department: string;
  kommentar?: string;
  verfügbar?: string;
  initials?: string;
  sourceSheet?: string;
  isActive: boolean;
}

export interface ExcelParseResult {
  assignments: ShiftAssignment[];
  personnel: ExcelPersonnelData[];
  errors: string[];
  summary: {
    totalAssignments: number;
    assignedPersonnel: number;
    unassignedPersonnel: number;
    sheetsProcessed: string[];
    shiftDate?: string;
  };
}

// Define the expected shift plan columns (shift types)
const SHIFT_COLUMNS = {
  'Bereitschaften (BD)': 'Bereitschaften (BD)',
  'Rufdienste (RD)': 'Rufdienste (RD)', 
  'Frühdienste (Früh)': 'Frühdienste (Früh)',
  'Zwischendienste/Mitteldienste (Mittel)': 'Zwischendienste/Mitteldienste (Mittel)',
  'Spätdienste (Spät)': 'Spätdienste (Spät)'
};

// Mapping for different possible column name variations
const COLUMN_MAPPINGS = {
  'Bereitschaften (BD)': ['bereitschaften', 'bd', 'bereitschaft'],
  'Rufdienste (RD)': ['rufdienste', 'rd', 'rufdienst', 'ruf'],
  'Frühdienste (Früh)': ['frühdienste', 'früh', 'fruh', 'frühdienst', 'early'],
  'Zwischendienste/Mitteldienste (Mittel)': ['zwischendienste', 'mitteldienste', 'mittel', 'zwischen', 'middle'],
  'Spätdienste (Spät)': ['spätdienste', 'spät', 'spaet', 'spätdienst', 'late', 'späte']
};

/**
 * Generate initials from a full name
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function generateInitials(name: string): string {
  if (!name || typeof name !== 'string') return 'XX';
  
  const parts = name.trim().split(' ').filter(part => part.length > 0);
  
  if (parts.length === 0) return 'XX';
  if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
  
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

/**
 * Find shift column indices by matching column headers
 */
function findShiftColumns(headers: string[]): Map<string, number> {
  const shiftColumns = new Map<string, number>();
  
  Object.entries(COLUMN_MAPPINGS).forEach(([shiftType, variations]) => {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i]?.toString().toLowerCase().trim();
      if (variations.some(variation => header.includes(variation.toLowerCase()))) {
        shiftColumns.set(shiftType, i);
        break;
      }
    }
  });
  
  return shiftColumns;
}

/**
 * Parse individual names from a multi-line cell
 * Expected format: "Nachname, Vorname (Department) (Code)"
 */
function parseNamesFromCell(cellContent: string): ShiftAssignment[] {
  if (!cellContent || typeof cellContent !== 'string') return [];
  
  const assignments: ShiftAssignment[] = [];
  
  // Split by newlines and filter empty lines
  const lines = cellContent.split(/\r?\n/).filter(line => line.trim().length > 0);
  
  lines.forEach(line => {
    const trimmedLine = line.trim();
    if (!trimmedLine) return;
    
    try {
      // Parse format: "Nachname, Vorname (Department) (Code)"
      // Example: "Findeisen, Sarah (OP KLD) (VB)"
      
      // First, try to extract the basic name part before any parentheses
      const nameMatch = trimmedLine.match(/^([^(]+)/);
      if (!nameMatch) return;
      
      const namePart = nameMatch[1].trim();
      
      // Split by comma to get last name and first name
      const nameParts = namePart.split(',').map(part => part.trim());
      
      if (nameParts.length >= 2) {
        const lastName = nameParts[0];
        const firstName = nameParts[1];
        const fullName = `${firstName} ${lastName}`;
        
        assignments.push({
          name: fullName,
          lastName,
          firstName,
          shiftType: '', // Will be set by the calling function
          availability: '', // Will be set by the calling function
          originalText: trimmedLine
        });
      } else if (nameParts.length === 1) {
        // Handle cases where there's no comma (single name)
        const parts = nameParts[0].split(' ').filter(p => p.length > 0);
        if (parts.length >= 2) {
          const firstName = parts[0];
          const lastName = parts.slice(1).join(' ');
          const fullName = `${firstName} ${lastName}`;
          
          assignments.push({
            name: fullName,
            lastName,
            firstName,
            shiftType: '',
            availability: '',
            originalText: trimmedLine
          });
        }
      }
    } catch (error) {
      console.warn('Error parsing name from line:', trimmedLine, error);
    }
  });
  
  return assignments;
}

/**
 * Process a single worksheet to extract shift assignments
 */
function processShiftPlan(worksheet: XLSX.WorkSheet, sheetName: string): {
  assignments: ShiftAssignment[];
  errors: string[];
  processedAssignments: number;
} {
  const assignments: ShiftAssignment[] = [];
  const errors: string[] = [];
  let processedAssignments = 0;
  
  try {
    // Convert worksheet to array of arrays
    const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    if (data.length < 2) {
      errors.push(`Sheet "${sheetName}": No data found`);
      return { assignments, errors, processedAssignments };
    }
    
    // Get headers from first row
    const headers = data[0].map(h => h?.toString() || '');
    
    // Find shift column indices
    const shiftColumns = findShiftColumns(headers);
    
    if (shiftColumns.size === 0) {
      errors.push(`Sheet "${sheetName}": No shift columns found. Expected columns like: ${Object.keys(SHIFT_COLUMNS).join(', ')}`);
      return { assignments, errors, processedAssignments };
    }
    
    // Process each shift column
    shiftColumns.forEach((columnIndex, shiftType) => {
      // Look through all rows for this column to find names
      for (let i = 1; i < data.length; i++) {
        const cellContent = data[i][columnIndex];
        if (cellContent && cellContent.toString().trim()) {
          const namesInCell = parseNamesFromCell(cellContent.toString());
          
          // Set the shift type and availability for each parsed name
          // eslint-disable-next-line no-loop-func
          namesInCell.forEach(assignment => {
            assignment.shiftType = shiftType;
            assignment.availability = shiftType; // Use shift type as availability category
            assignments.push(assignment);
            processedAssignments++;
          });
        }
      }
    });
    
  } catch (error) {
    errors.push(`Sheet "${sheetName}": ${error instanceof Error ? error.message : 'Unknown error processing sheet'}`);
  }
  
  return { assignments, errors, processedAssignments };
}

/**
 * Update personnel availability based on shift assignments
 */
export function updatePersonnelAvailability(
  existingPersonnel: any[], 
  shiftAssignments: ShiftAssignment[]
): { updated: any[]; assignments: ShiftAssignment[]; availabilityTags: PersonnelAvailabilityTags } {
  // Create a map of assigned personnel by name with multiple shift types
  const assignmentMap = new Map<string, ShiftAssignment[]>();
  
  shiftAssignments.forEach(assignment => {
    const normalizedName = assignment.name.toLowerCase().trim();
    if (!assignmentMap.has(normalizedName)) {
      assignmentMap.set(normalizedName, []);
    }
    assignmentMap.get(normalizedName)!.push(assignment);
  });
  
  // Debug logging
  console.log('Assignment Map:', Array.from(assignmentMap.entries()).map(([name, assignments]) => ({
    name,
    assignmentCount: assignments.length,
    shiftTypes: assignments.map(a => a.shiftType)
  })));
  
  // Create availability tags map
  const availabilityTags: PersonnelAvailabilityTags = {};
  
  // Update existing personnel
  const updatedPersonnel = existingPersonnel.map(person => {
    const normalizedPersonName = person.name.toLowerCase().trim();
    const assignments = assignmentMap.get(normalizedPersonName);
    
    if (assignments && assignments.length > 0) {
      // Person has one or more shift assignments
      const availabilityTypes = assignments.map(a => a.availability);
      const shiftTypes = assignments.map(a => a.shiftType);
      
      // Store multiple tags for this person using their ID as key
      availabilityTags[person.id] = availabilityTypes;
      
      return {
        ...person,
        verfügbar: availabilityTypes.join(', '), // Join multiple availabilities
        shiftAssignment: shiftTypes.join(', '), // Join multiple shift types
        availabilityTags: availabilityTypes, // Array of all availability tags
        shiftTags: shiftTypes, // Array of all shift tags
        isAvailable: true
      };
    } else {
      // Person is not assigned - mark as not available and clear previous tags
      availabilityTags[person.id] = [];
      
      return {
        ...person,
        verfügbar: 'nicht verfügbar',
        shiftAssignment: undefined,
        availabilityTags: [],
        shiftTags: [],
        isAvailable: false
      };
    }
  });
  
  return { updated: updatedPersonnel, assignments: shiftAssignments, availabilityTags };
}

/**
 * Main function to parse shift plan Excel file
 */
export function parseExcelFile(file: File): Promise<ExcelParseResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const result: ExcelParseResult = {
          assignments: [],
          personnel: [], // Will be populated by updatePersonnelAvailability
          errors: [],
          summary: {
            totalAssignments: 0,
            assignedPersonnel: 0,
            unassignedPersonnel: 0,
            sheetsProcessed: [],
            shiftDate: file.name.includes(new Date().toISOString().split('T')[0]) ? 
              new Date().toISOString().split('T')[0] : undefined
          }
        };
        
        // Process each worksheet
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const { assignments, errors, processedAssignments } = processShiftPlan(worksheet, sheetName);
          
          result.assignments.push(...assignments);
          result.errors.push(...errors);
          result.summary.totalAssignments += processedAssignments;
          result.summary.sheetsProcessed.push(sheetName);
        });
        
        // Calculate summary statistics
        const uniqueAssignedNames = new Set(result.assignments.map(a => a.name.toLowerCase()));
        result.summary.assignedPersonnel = uniqueAssignedNames.size;
        
        resolve(result);
        
      } catch (error) {
        reject(new Error(`Failed to parse Excel file: ${error instanceof Error ? error.message : 'Unknown error'}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };
    
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Validate shift assignments and create personnel updates
 */
export function validatePersonnelData(assignments: ShiftAssignment[]): {
  valid: ShiftAssignment[];
  invalid: { assignment: ShiftAssignment; reason: string }[];
} {
  const valid: ShiftAssignment[] = [];
  const invalid: { assignment: ShiftAssignment; reason: string }[] = [];
  
  // Track combinations of name + shift type to avoid true duplicates
  const seenCombinations = new Set<string>();
  
  assignments.forEach(assignment => {
    // Check for required fields
    if (!assignment.name || assignment.name.length < 2) {
      invalid.push({ assignment, reason: 'Name zu kurz oder fehlt' });
      return;
    }
    
    if (!assignment.shiftType) {
      invalid.push({ assignment, reason: 'Schichttyp fehlt' });
      return;
    }
    
    // Check for exact duplicates (same person + same shift type)
    const normalizedName = assignment.name.toLowerCase().trim();
    const combinationKey = `${normalizedName}|${assignment.shiftType}`;
    if (seenCombinations.has(combinationKey)) {
      invalid.push({ assignment, reason: `Doppelte Zuordnung für ${assignment.name} in ${assignment.shiftType}` });
      return;
    }
    seenCombinations.add(combinationKey);
    
    valid.push(assignment);
  });
  
  return { valid, invalid };
}

/**
 * Convert shift assignments to personnel format for integration with existing personnel list
 */
export function integrateShiftAssignments(
  existingPersonnel: any[], 
  shiftAssignments: ShiftAssignment[]
): any[] {
  return updatePersonnelAvailability(existingPersonnel, shiftAssignments).updated;
}

/**
 * Filter personnel to show only available personnel (for draggable sidebar)
 */
export function getAvailablePersonnel(personnel: any[]): any[] {
  return personnel.filter(person => 
    person.verfügbar && 
    person.verfügbar !== 'nicht verfügbar' && 
    person.isAvailable !== false
  );
}

/**
 * Get shift statistics for summary display
 */
export function getShiftStatistics(assignments: ShiftAssignment[]): {
  totalAssigned: number;
  byShiftType: Record<string, number>;
  assignedNames: string[];
} {
  const byShiftType: Record<string, number> = {};
  const assignedNames = new Set<string>();
  
  assignments.forEach(assignment => {
    if (!byShiftType[assignment.shiftType]) {
      byShiftType[assignment.shiftType] = 0;
    }
    byShiftType[assignment.shiftType]++;
    assignedNames.add(assignment.name);
  });
  
  return {
    totalAssigned: assignedNames.size,
    byShiftType,
    assignedNames: Array.from(assignedNames)
  };
}
