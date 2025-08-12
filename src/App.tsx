import React, { useEffect, useState, useCallback, DragEvent } from 'react';
import { Search, User, Calendar, Settings, FileText, Users, Upload, CheckCircle, RotateCcw, AlertCircle } from 'lucide-react';
import { useDropzone } from 'react-dropzone';
import { parseExcelFile, validatePersonnelData, updatePersonnelAvailability, ExcelParseResult, ShiftAssignment } from './excel-utils';
import { useLocalPersonnelData, Personnel } from './local-storage';

type TableConfigKey = 'main' | 'emergency' | 'weekend';

const ORPlannerApp = () => {
  const [currentPage, setCurrentPage] = useState('planner');
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  
  // Use local storage for personnel management
  const {
    personnel,
    loading: personnelLoading,
    addPersonnel,
    updatePersonnel,
    deletePersonnel,
    generateInitials,
    exportToExcel,
    clearAllData
  } = useLocalPersonnelData();
  
  // Table configurations based on your CSV structure
  const tableConfigs = {
    main: {
      name: "Hauptplan",
      roles: [
        "OP-Pflege", "OP-Pflege (2)", 
        "Anästhesie Pflege", "Anästhesie Pflege (2)",
        "OP-Praktikant", "OP-Praktikant (2)",
        "Anästhesie Praktikant", "Anästhesie Praktikant (2)",
        "MFA", "MFA (2)",
        "ATA Schüler", "ATA Schüler (2)",
        "OTA Schüler", "OTA Schüler (2)"
      ],
      rooms: [
        "A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4",
        "D1", "D2", "D3", "D4", "Kreißsaal", "Derma OP", 
        "Medicum IV", "Medicum V", "Schleuse", "Externer Saal*", 
        "Externer Saal *", "POBE"
      ]
    },
    emergency: {
      name: "Notfallplan",
      roles: [
        "Bereitschaftsarzt", "Notfall-Pflege", "ATA Bereitschaft", "",
        "OP-Koordination", "Springer", ""
      ],
      rooms: ["Notfall-OP", "Schockraum", "Hybrid-OP"]
    },
    weekend: {
      name: "Wochenendplan",
      roles: [
        "Wochenenddienst Arzt", "Wochenenddienst Pflege", "Rufbereitschaft", ""
      ],
      rooms: ["A1", "B1", "D1", "Kreißsaal"]
    }
  };

  const [currentTable, setCurrentTable] = useState<TableConfigKey>('main');
  const [assignments, setAssignments] = useState<Record<string, Personnel[]>>(() => {
    // Try to load assignments from localStorage
    const saved = localStorage.getItem('op-planner-assignments');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (error) {
        console.error('Error parsing saved assignments:', error);
      }
    }
    return {};
  });
  const [searchTerm, setSearchTerm] = useState("");
  const [plannerGroupFilter, setPlannerGroupFilter] = useState<string>('all');
  const [draggedPerson, setDraggedPerson] = useState<Personnel | null>(null);

  // Availability tags state for personnel
  const [availabilityTags, setAvailabilityTags] = useState<Record<number, string[]>>(() => {
    const saved = localStorage.getItem('op-planner-availability-tags');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (error) {
        console.error('Error parsing saved availability tags:', error);
      }
    }
    return {};
  });

  // Save availability tags to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem('op-planner-availability-tags', JSON.stringify(availabilityTags));
  }, [availabilityTags]);

  // Save assignments to localStorage only
  useEffect(() => {
    localStorage.setItem('op-planner-assignments', JSON.stringify(assignments));
  }, [assignments]);

  // Load assignments from localStorage when date changes
  useEffect(() => {
    const loadAssignmentsForDate = async () => {
      // Load assignments from localStorage only
      const saved = localStorage.getItem(`op-planner-assignments-${selectedDate}`);
      if (saved) {
        try {
          setAssignments(JSON.parse(saved));
        } catch (parseError) {
          console.error('Error parsing saved assignments:', parseError);
          setAssignments({});
        }
      } else {
        // Clear assignments for new date if no saved data
        setAssignments({});
      }
    };

    loadAssignmentsForDate();
  }, [selectedDate]);

  // Legend data from your CSV
  const legendData = {
    specialties: {
      "Frühstart": "Früher Beginn geplant",
      "TIVA": "Total intravenöse Anästhesie",
      "Prüfung": "Prüfungspatient",
      "Triggerfrei": "Maligne Hyperthermie Vorsichtsmaßnahmen"
    },
    services: {
      "BD 1": "Bereitschaftsdienst 1",
      "BD 2": "Bereitschaftsdienst 2", 
      "KR": "Kreißsaal",
      "AR": "Aufwachraum",
      "D-R": "Dienst-Rufbereitschaft",
      "L-R": "Leitende Rufbereitschaft",
      "S": "Springer",
      "AS": "Anästhesie Service"
    },
    colors: {
      "KR blau": "Kreißsaal Normalfall",
      "KR grün": "Kreißsaal Notfall",
      "R": "Rufbereitschaft",
      "BD": "Bereitschaftsdienst",
      "SBD": "Spät-Bereitschaftsdienst",
      "PD": "Präsenzdienst",
      "ZD": "Zusatzdienst",
      "PM": "Personalmanagement"
    }
  };

  const NavigationBar = () => (
    <nav className="bg-white border-b border-gray-200 px-6 py-4">
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-6">
          <h1 className="text-2xl font-bold text-green-600">PersonalPlan</h1>
          
          <div className="flex space-x-1">
            {[
              { id: 'planner', name: 'Planer', icon: Calendar },
              { id: 'personnel', name: 'Personal', icon: Users },
              { id: 'legend', name: 'Legende', icon: FileText },
              { id: 'settings', name: 'Einstellungen', icon: Settings }
            ].map(({ id, name, icon: Icon }) => (
              <button
                key={id}
                onClick={() => setCurrentPage(id)}
                className={`flex items-center space-x-2 px-4 py-2 rounded-lg transition-colors ${
                  currentPage === id
                    ? 'bg-green-100 text-green-700'
                    : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
                }`}
              >
                <Icon size={18} />
                <span>{name}</span>
              </button>
            ))}
          </div>
        </div>

      </div>
    </nav>
  );

  const PlannerPage = () => { 
    const config = tableConfigs[currentTable];
    
    const filteredPersonnel = personnel.filter(person => {
      const matchesSearch = person.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
        person.gruppe.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesGroup = plannerGroupFilter === 'all' || person.gruppe === plannerGroupFilter;
      // Only show available personnel in the planning view
      const isAvailable = person.verfügbar && person.verfügbar !== 'nicht verfügbar' && person.isAvailable !== false;
      return matchesSearch && matchesGroup && isAvailable;
    });

    // Don't filter out assigned personnel - allow multiple instances
    const unassignedPersonnel = filteredPersonnel;

    const handleDragStart = useCallback((e: DragEvent<HTMLDivElement>, person: Personnel) => {
      setDraggedPerson(person);
      e.dataTransfer.effectAllowed = 'move';
    }, []);

    const handleDragOver = useCallback((e: DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
    }, []);

    const handleDrop = useCallback((e: DragEvent<HTMLDivElement>, roleIndex: number, roomIndex: number) => {
      e.preventDefault();
      if (!draggedPerson) return;

      const cellKey = `${currentTable}-${roleIndex}-${roomIndex}`;
      setAssignments(prev => {
        const newAssignments = { ...prev };
        
        // Don't remove from other cells - allow multiple instances
        // Just add to the target cell
        if (!newAssignments[cellKey]) {
          newAssignments[cellKey] = [];
        }
        
        // Check if person is already in this specific cell to avoid duplicates in same cell
        const isAlreadyInCell = newAssignments[cellKey].some(p => p.id === draggedPerson.id);
        if (!isAlreadyInCell) {
          newAssignments[cellKey].push({...draggedPerson}); // Create a copy
        }
        
        return newAssignments;
      });
      
      setDraggedPerson(null);
    }, [draggedPerson, currentTable]);

    // Helper function to check if person is in consecutive cells
    const getConsecutiveAssignments = (roleIndex: number, personId: number) => {
      const consecutiveCells = [];
      const config = tableConfigs[currentTable];
      
      for (let roomIndex = 0; roomIndex < config.rooms.length; roomIndex++) {
        const cellKey = `${currentTable}-${roleIndex}-${roomIndex}`;
        const cellPersonnel = assignments[cellKey] || [];
        
        if (cellPersonnel.some(p => p.id === personId)) {
          consecutiveCells.push(roomIndex);
        }
      }
      
      // Group consecutive room indices
      const groups = [];
      let currentGroup = [consecutiveCells[0]];
      
      for (let i = 1; i < consecutiveCells.length; i++) {
        if (consecutiveCells[i] === consecutiveCells[i-1] + 1) {
          currentGroup.push(consecutiveCells[i]);
        } else {
          groups.push(currentGroup);
          currentGroup = [consecutiveCells[i]];
        }
      }
      if (currentGroup.length > 0) {
        groups.push(currentGroup);
      }
      
      return groups;
    };

    // Helper function to get shift type abbreviations
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const getShiftAbbreviations = (personId: number): string => {
      const tags = availabilityTags[personId] || [];
      if (tags.length === 0) return '';
      
      const abbreviations = tags.map(tag => getTagAbbreviation(tag));
      
      return abbreviations.length > 0 ? `(${abbreviations.join('/')})` : '';
    };

    // Helper function to convert a single tag to abbreviation
    const getTagAbbreviation = (tag: string): string => {
      if (tag.includes('Bereitschaft')) return 'BD';
      if (tag.includes('Rufdienst')) return 'RD';
      if (tag.includes('Frühdienst')) return 'Früh';
      if (tag.includes('Spätdienst')) return 'Spät';
      if (tag.includes('Zwischendienst') || tag.includes('Mitteldienst')) return 'Mittel';
      return tag.split(' ')[0]; // Fallback to first word
    };

    const PersonCard = ({ person, isDraggable = true, onRemove = (id: number) => {}, size = "normal" }: {person: Personnel; isDraggable?: boolean; onRemove?: (id: number) => void; size?: "normal" | "small";}) => {
      const isSmall = size === "small";
      return (
        <div
          className={`flex items-center gap-2 p-2 bg-white rounded-lg border border-gray-200 shadow-sm hover:shadow-md transition-all duration-200 ${
            isDraggable ? 'cursor-move hover:bg-gray-50' : ''
          } ${isSmall ? 'text-xs' : 'text-sm'}`}
          draggable={isDraggable}
          onDragStart={isDraggable ? (e) => handleDragStart(e, person) : undefined}
        >
          <div className={`${isSmall ? 'w-6 h-6 text-xs' : 'w-8 h-8 text-sm'} bg-gradient-to-br from-blue-500 to-blue-600 rounded-full flex items-center justify-center text-white font-semibold`}>
            {person.initials}
          </div>
          <div className="flex-1 min-w-0">
            <div className={`font-medium text-gray-900 truncate ${isSmall ? 'text-xs' : 'text-sm'}`}>
              {person.name}
            </div>
            <div className={`text-gray-500 truncate ${isSmall ? 'text-xs' : 'text-xs'}`}>
              {person.gruppe} {getShiftAbbreviations(person.id)}
            </div>
          </div>
          {onRemove && (
            <button
              onClick={() => onRemove(person.id)}
              className="text-red-500 hover:text-red-700 text-xs p-1"
            >
              ×
            </button>
          )}
        </div>
      );
    };

    return (
      <div className="flex h-full">
        <div className="flex-1 flex flex-col overflow-hidden">
          {/* Table Selector and Date Controls */}
          <div className="bg-gray-50 border-b border-gray-200 px-6 py-3">
            <div className="flex items-center justify-between">
              <div className="flex space-x-2">
                {Object.entries(tableConfigs).map(([key, config]) => (
                  <button
                    key={key}
                    onClick={() => setCurrentTable(key as TableConfigKey)}
                    className={`px-4 py-2 rounded-lg transition-colors ${
                      currentTable === key
                        ? 'bg-white text-green-700 shadow-sm border border-green-200'
                        : 'text-gray-600 hover:text-gray-900 hover:bg-white'
                    }`}
                  >
                    {config.name}
                  </button>
                ))}
              </div>
              
              <div className="flex items-center space-x-4">
                <input
                  type="date"
                  value={selectedDate}
                  onChange={(e) => setSelectedDate(e.target.value)}
                  className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                />
                <button
                  onClick={() => setAssignments({})}
                  className="flex items-center space-x-2 px-4 py-2 text-green-600 hover:text-green-700 hover:bg-green-50 rounded-lg transition-colors"
                >
                  <RotateCcw size={18} />
                  <span>Reset</span>
                </button>
              </div>
            </div>
          </div>

          {/* Grid */}
          <div className="flex-1 overflow-auto">
            <div className="inline-block min-w-full">
              <div className="sticky top-0 z-10 bg-white border-b border-gray-200">
                <div className="flex">
                  <div className="w-48 p-3 font-medium text-gray-900 bg-gray-50 border-r border-gray-200">
                    Gruppe
                  </div>
                  {config.rooms.map((room, index) => (
                    <div key={index} className="w-32 p-3 text-center font-medium text-gray-900 border-r border-gray-200 bg-gray-50">
                      {room}
                    </div>
                  ))}
                </div>
              </div>

              <div className="divide-y divide-gray-200">
                {config.roles.map((role, roleIndex) => {
                  // Track which cells have been merged
                  const mergedCells = new Set<number>();
                  
                  return (
                    <div key={roleIndex} className="flex hover:bg-gray-50">
                      <div className="w-48 p-3 font-medium text-gray-700 bg-white border-r border-gray-200 sticky left-0 z-10">
                        {role}
                      </div>
                      
                      {config.rooms.map((room, roomIndex) => {
                        // Skip if this cell is part of a merged span
                        if (mergedCells.has(roomIndex)) {
                          return null;
                        }
                        
                        const cellKey = `${currentTable}-${roleIndex}-${roomIndex}`;
                        const assignedPersons = assignments[cellKey] || [];
                        
                        // Check for consecutive assignments for each person
                        let spanWidth = 1;
                        const allPersonsInConsecutiveCells = new Map<number, number>(); // personId -> span length
                        
                        assignedPersons.forEach(person => {
                          const consecutiveGroups = getConsecutiveAssignments(roleIndex, person.id);
                          consecutiveGroups.forEach(group => {
                            if (group.includes(roomIndex) && group[0] === roomIndex) {
                              // This is the first cell of a consecutive group
                              const span = group.length;
                              if (span > 1) {
                                allPersonsInConsecutiveCells.set(person.id, span);
                                spanWidth = Math.max(spanWidth, span);
                                // Mark subsequent cells as merged
                                for (let i = 1; i < span; i++) {
                                  mergedCells.add(roomIndex + i);
                                }
                              }
                            }
                          });
                        });
                        
                        const cellStyle = spanWidth > 1 ? { width: `${spanWidth * 8}rem` } : {};
                        
                        return (
                          <div
                            key={roomIndex}
                            className={`${spanWidth > 1 ? '' : 'w-32'} min-h-16 p-2 border-r border-gray-200 bg-white hover:bg-blue-50 transition-colors relative`}
                            style={cellStyle}
                            onDragOver={handleDragOver}
                            onDrop={(e) => handleDrop(e, roleIndex, roomIndex)}
                          >
                            <div className="space-y-1">
                              {assignedPersons.map((person) => {
                                const personSpan = allPersonsInConsecutiveCells.get(person.id) || 1;
                                const isSpanned = personSpan > 1;
                                
                                return (
                                  <div key={`${person.id}-${roomIndex}`} className={isSpanned ? 'relative' : ''}>
                                    <PersonCard
                                      person={person}
                                      isDraggable={true}
                                      onRemove={(personId: number) => {
                                        setAssignments(prev => {
                                          const newAssignments = { ...prev };
                                          // Remove only from this specific cell
                                          if (newAssignments[cellKey]) {
                                            newAssignments[cellKey] = newAssignments[cellKey].filter(p => p.id !== personId);
                                            if (newAssignments[cellKey].length === 0) {
                                              delete newAssignments[cellKey];
                                            }
                                          }
                                          return newAssignments;
                                        });
                                      }}
                                      size="small"
                                    />
                                    {isSpanned && (
                                      <div className="absolute top-0 left-0 right-0 bottom-0 border-2 border-blue-300 border-dashed rounded pointer-events-none opacity-50" />
                                    )}
                                  </div>
                                );
                              })}
                            </div>
                            {spanWidth > 1 && (
                              <div className="absolute top-1 right-1 text-xs text-blue-600 bg-blue-100 px-1 rounded">
                                {spanWidth} rooms
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* Personnel Sidebar */}
        <div className="w-80 bg-white border-l border-gray-200 flex flex-col">
          <div className="p-4 border-b border-gray-200">
            <h2 className="text-lg font-semibold text-gray-900 mb-3">Personenauswahl</h2>
            
            <div className="relative mb-3">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" size={16} />
              <input
                type="text"
                placeholder="Suche..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none"
              />
            </div>

            <div className="relative">
              <select
                value={plannerGroupFilter}
                onChange={(e) => setPlannerGroupFilter(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none text-sm"
              >
                <option value="all">Alle Gruppen</option>
                <option value="OP-Pflege">OP-Pflege</option>
                <option value="Anästhesie Pflege">Anästhesie Pflege</option>
                <option value="OP-Praktikant">OP-Praktikant</option>
                <option value="Anästhesie Praktikant">Anästhesie Praktikant</option>
                <option value="MFA">MFA</option>
                <option value="ATA Schüler">ATA Schüler</option>
                <option value="OTA Schüler">OTA Schüler</option>
              </select>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-4">
            <div className="space-y-3">
              {personnel.length === 0 ? (
                <div className="text-center text-gray-500 py-8">
                  <User size={48} className="mx-auto text-gray-300 mb-2" />
                  <p>Keine Personaleinträge verfügbar</p>
                  <p className="text-sm">Bitte SharePoint-Verbindung prüfen</p>
                </div>
              ) : unassignedPersonnel.length === 0 ? (
                <div className="text-center text-gray-500 py-8">
                  <User size={48} className="mx-auto text-gray-300 mb-2" />
                  <p>Alle Personen zugewiesen</p>
                </div>
              ) : (
                unassignedPersonnel.map((person) => (
                  <PersonCard key={person.id} person={person} />
                ))
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const LegendPage = () => (
    <div className="p-6 max-w-6xl mx-auto">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Legende und Besonderheiten</h2>
      
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">Besonderheiten</h3>
          <div className="space-y-3">
            {Object.entries(legendData.specialties).map(([key, description]) => (
              <div key={key} className="flex items-start space-x-3">
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm font-medium">
                  {key}
                </span>
                <span className="text-gray-700 text-sm">{description}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">Dienste</h3>
          <div className="space-y-3">
            {Object.entries(legendData.services).map(([key, description]) => (
              <div key={key} className="flex items-start space-x-3">
                <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded text-sm font-medium">
                  {key}
                </span>
                <span className="text-gray-700 text-sm">{description}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">Farbcodes</h3>
          <div className="space-y-3">
            {Object.entries(legendData.colors).map(([key, description]) => (
              <div key={key} className="flex items-start space-x-3">
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm font-medium">
                  {key}
                </span>
                <span className="text-gray-700 text-sm">{description}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="mt-8 bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">Hinweise</h3>
        <ul className="list-disc list-inside space-y-2 text-gray-700">
          <li>* Von Planenden zu pflegen - Diese Felder müssen manuell aktualisiert werden</li>
          <li>Externe Säle sind mit * markiert und können zusätzliche Vorbereitung erfordern</li>
          <li>Praktikanten benötigen Supervision durch erfahrenes Personal</li>
          <li>Triggerfrei bedeutet maligne Hyperthermie Vorsichtsmaßnahmen</li>
        </ul>
      </div>
    </div>
  );

  const PersonnelPage = () => {
    const [showAddForm, setShowAddForm] = useState(false);
    const [selectedGroup, setSelectedGroup] = useState<string>('all');
    const [editingPerson, setEditingPerson] = useState<Personnel | null>(null);
    const [loading, setLoading] = useState(false);
    const [newPerson, setNewPerson] = useState({
      name: '',
      gruppe: 'OP-Pflege' as Personnel['gruppe'],
      department: '',
      kommentar: '',
      verfügbar: 'nicht verfügbar' as Personnel['verfügbar'],
      initials: ''
    });

    // Excel Import State
    const [showImportModal, setShowImportModal] = useState(false);
    const [importResult, setImportResult] = useState<ExcelParseResult | null>(null);
    const [importStatus, setImportStatus] = useState<'idle' | 'processing' | 'preview' | 'importing' | 'complete' | 'error'>('idle');
    const [importError, setImportError] = useState<string | null>(null);
    const [previewData, setPreviewData] = useState<ShiftAssignment[]>([]);
    const [validationResult, setValidationResult] = useState<{ valid: ShiftAssignment[]; invalid: { assignment: ShiftAssignment; reason: string }[] } | null>(null);

    // Generate initials automatically when name changes
    useEffect(() => {
      if (newPerson.name) {
        const parts = newPerson.name.trim().split(' ').filter(part => part.length > 0);
        let initials = '';
        if (parts.length === 0) initials = 'XX';
        else if (parts.length === 1) initials = parts[0].substring(0, 2).toUpperCase();
        else initials = (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        
        setNewPerson(prev => ({ ...prev, initials }));
      }
    }, [newPerson.name]);

    const groups = [
      'OP-Pflege', 'Anästhesie Pflege', 'OP-Praktikant', 
      'Anästhesie Praktikant', 'MFA', 'ATA Schüler', 'OTA Schüler'
    ];

    const availabilityOptions = [
      'Bereitschaften (BD)', 'Rufdienste (RD)', 'Frühdienste (Früh)', 
      'Zwischendienste/Mitteldienste (Mittel)', 'Spätdienste (Spät)', 'nicht verfügbar'
    ];

    const filteredPersonnel = personnel.filter(person => 
      selectedGroup === 'all' || person.gruppe === selectedGroup
    );

    const handleAddPerson = async () => {
      console.log('Adding person:', newPerson);
      
      try {
        // Create the new person object for local storage
        const personToAdd = {
          name: newPerson.name,
          gruppe: newPerson.gruppe,
          department: newPerson.department,
          kommentar: newPerson.kommentar,
          verfügbar: newPerson.verfügbar,
          initials: newPerson.initials,
          isActive: newPerson.verfügbar !== 'nicht verfügbar'
        };
        
        // Add person using local storage
        addPersonnel(personToAdd);
        
        // Reset the form
        setNewPerson({
          name: '',
          gruppe: 'OP-Pflege',
          department: '',
          kommentar: '',
          verfügbar: 'nicht verfügbar',
          initials: ''
        });
        setShowAddForm(false);
        
      } catch (error) {
        console.error('Error adding person:', error);
        alert('Fehler beim Hinzufügen der Person. Bitte versuchen Sie es erneut.');
      }
    };

    const handleEditPerson = (person: Personnel) => {
      setEditingPerson(person);
      setNewPerson({
        name: person.name,
        gruppe: person.gruppe as Personnel['gruppe'],
        department: person.department,
        kommentar: person.kommentar,
        verfügbar: person.verfügbar as Personnel['verfügbar'],
        initials: person.initials
      });
      setShowAddForm(true);
    };

    const handleUpdatePerson = async () => {
      if (!editingPerson) return;
      
      console.log('Updating person:', editingPerson.id, newPerson);
      
      try {
        const updates = {
          name: newPerson.name,
          gruppe: newPerson.gruppe,
          department: newPerson.department,
          kommentar: newPerson.kommentar,
          verfügbar: newPerson.verfügbar,
          initials: newPerson.initials,
          isActive: newPerson.verfügbar !== 'nicht verfügbar'
        };

        // Update person using local storage
        updatePersonnel(editingPerson.id, updates);
        
        // Reset form
        setNewPerson({
          name: '',
          gruppe: 'OP-Pflege',
          department: '',
          kommentar: '',
          verfügbar: 'nicht verfügbar',
          initials: ''
        });
        setShowAddForm(false);
        setEditingPerson(null);
        
      } catch (error) {
        console.error('Error updating person:', error);
        alert('Fehler beim Aktualisieren der Person. Bitte versuchen Sie es erneut.');
      }
    };

    const handleDeletePerson = async (personId: number) => {
      if (window.confirm('Are you sure you want to delete this person?')) {
        console.log('Deleting person:', personId);
        
        try {
          // Delete person using local storage
          deletePersonnel(personId);
          
          // Also remove from any assignments
          setAssignments(prev => {
            const newAssignments = { ...prev };
            Object.keys(newAssignments).forEach(key => {
              newAssignments[key] = newAssignments[key].filter(p => p.id !== personId);
              if (newAssignments[key].length === 0) {
                delete newAssignments[key];
              }
            });
            return newAssignments;
          });
          
        } catch (error) {
          console.error('Error deleting person:', error);
          alert('Fehler beim Löschen der Person. Bitte versuchen Sie es erneut.');
        }
      }
    };

    // Excel Import Handlers
    const handleFileUpload = async (files: File[]) => {
      if (files.length === 0) return;
      
      const file = files[0];
      
      // Validate file type
      if (!file.name.match(/\.(xlsx|xls)$/i)) {
        setImportError('Bitte wählen Sie eine Excel-Datei (.xlsx oder .xls)');
        setImportStatus('error');
        return;
      }
      
      setImportStatus('processing');
      setImportError(null);
      
      try {
        const result = await parseExcelFile(file);
        setImportResult(result);
        
        if (result.errors.length > 0) {
          console.warn('Excel parsing warnings:', result.errors);
        }
        
        // Validate the parsed assignments
        const validation = validatePersonnelData(result.assignments);
        setValidationResult(validation);
        setPreviewData(validation.valid);
        
        setImportStatus('preview');
        
      } catch (error) {
        console.error('Excel import error:', error);
        setImportError(error instanceof Error ? error.message : 'Fehler beim Verarbeiten der Excel-Datei');
        setImportStatus('error');
      }
    };

    const handleConfirmImport = async () => {
      if (!validationResult || validationResult.valid.length === 0) {
        setImportError('Keine gültigen Daten zum Importieren');
        return;
      }
      
      setImportStatus('importing');
      
      try {
        // Update personnel availability based on shift assignments
        const result = updatePersonnelAvailability(personnel, validationResult.valid);
        
        // Update availability tags if they exist
        if (result.availabilityTags) {
          setAvailabilityTags(result.availabilityTags);
        }
        
        // Update each person's availability directly using local storage
        console.log('Updating personnel availability...');
        
        for (const updatedPerson of result.updated) {
          try {
            // Find the original person to compare
            const originalPerson = personnel.find(p => p.id === updatedPerson.id);
            
            if (originalPerson && (
              originalPerson.verfügbar !== updatedPerson.verfügbar ||
              originalPerson.shiftAssignment !== updatedPerson.shiftAssignment ||
              originalPerson.isAvailable !== updatedPerson.isAvailable
            )) {
              // Update the person using local storage
              updatePersonnel(updatedPerson.id, {
                verfügbar: updatedPerson.verfügbar,
                shiftAssignment: updatedPerson.shiftAssignment,
                availabilityTags: updatedPerson.availabilityTags,
                shiftTags: updatedPerson.shiftTags,
                isAvailable: updatedPerson.isAvailable
              });
            }
          } catch (updateError) {
            console.error('Failed to update person:', updatedPerson.name, updateError);
          }
        }
        
        console.log('Personnel availability updated');
        console.log('Availability tags:', result.availabilityTags);
        console.log('Shift assignments:', result.assignments);
        
        // Simulate API delay
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        setImportStatus('complete');
        
      } catch (error) {
        console.error('Import error:', error);
        setImportError(error instanceof Error ? error.message : 'Fehler beim Importieren der Daten');
        setImportStatus('error');
      }
    };

    const resetImport = () => {
      setShowImportModal(false);
      setImportResult(null);
      setImportStatus('idle');
      setImportError(null);
      setPreviewData([]);
      setValidationResult(null);
    };

    // Dropzone configuration
    const { getRootProps, getInputProps, isDragActive } = useDropzone({
      onDrop: handleFileUpload,
      accept: {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
        'application/vnd.ms-excel': ['.xls']
      },
      maxFiles: 1,
      disabled: importStatus === 'processing' || importStatus === 'importing'
    });

    return (
      <div className="p-6">
        <div className="flex justify-between items-center mb-6">
          <h2 className="text-2xl font-bold text-gray-900">Personalverwaltung</h2>
          <div className="flex gap-3">
            <button
              onClick={() => setShowImportModal(true)}
              className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg transition-colors flex items-center gap-2"
            >
              <Upload className="w-4 h-4" />
              Schichtplan Import
            </button>
            <button
              onClick={() => setShowAddForm(true)}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg transition-colors flex items-center gap-2"
            >
              <User className="w-4 h-4" />
              Person hinzufügen
            </button>
            <button
              onClick={() => exportToExcel()}
              className="bg-gray-600 hover:bg-gray-700 text-white px-3 py-2 rounded-lg transition-colors flex items-center gap-1 text-sm"
              title="Personaldaten als Excel exportieren"
            >
              <Upload className="w-3 h-3" />
              Export
            </button>
          </div>
        </div>

        {/* Group Filter */}
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">Nach Gruppe filtern:</label>
          <select
            value={selectedGroup}
            onChange={(e) => setSelectedGroup(e.target.value)}
            className="border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
          >
            <option value="all">Alle Gruppen</option>
            {groups.map(group => (
              <option key={group} value={group}>{group}</option>
            ))}
          </select>
        </div>

        {/* Add/Edit Form Modal */}
        {showAddForm && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg p-6 w-full max-w-md">
              <h3 className="text-lg font-semibold mb-4">
                {editingPerson ? 'Person bearbeiten' : 'Neue Person hinzufügen'}
              </h3>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Name</label>
                  <input
                    type="text"
                    value={newPerson.name}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, name: e.target.value }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Vor- und Nachname"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Gruppe</label>
                  <select
                    value={newPerson.gruppe}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, gruppe: e.target.value as Personnel['gruppe'] }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    {groups.map(group => (
                      <option key={group} value={group}>{group}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Abteilung</label>
                  <input
                    type="text"
                    value={newPerson.department}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, department: e.target.value }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Abteilung"
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Verfügbarkeit</label>
                  <select
                    value={newPerson.verfügbar}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, verfügbar: e.target.value as Personnel['verfügbar'] }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                  >
                    {availabilityOptions.map(option => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Kürzel</label>
                  <input
                    type="text"
                    value={newPerson.initials}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, initials: e.target.value.toUpperCase() }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Automatisch generiert"
                    maxLength={3}
                  />
                </div>

                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Kommentar</label>
                  <textarea
                    value={newPerson.kommentar}
                    onChange={(e) => setNewPerson(prev => ({ ...prev, kommentar: e.target.value }))}
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                    placeholder="Zusätzliche Informationen"
                    rows={3}
                  />
                </div>
              </div>

              <div className="flex justify-end gap-3 mt-6">
                <button
                  onClick={() => {
                    setShowAddForm(false);
                    setEditingPerson(null);
                    setNewPerson({
                      name: '',
                      gruppe: 'OP-Pflege',
                      department: '',
                      kommentar: '',
                      verfügbar: 'nicht verfügbar',
                      initials: ''
                    });
                  }}
                  className="px-4 py-2 text-gray-600 hover:text-gray-800 transition-colors"
                >
                  Abbrechen
                </button>
                <button
                  onClick={editingPerson ? handleUpdatePerson : handleAddPerson}
                  disabled={!newPerson.name.trim() || !newPerson.department.trim()}
                  className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white px-4 py-2 rounded-lg transition-colors"
                >
                  {editingPerson ? 'Aktualisieren' : 'Hinzufügen'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Excel Import Modal */}
        {showImportModal && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg p-6 w-full max-w-4xl max-h-[90vh] overflow-y-auto">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-lg font-semibold">Schichtplan Import - Tägliche Verfügbarkeit</h3>
                <button
                  onClick={resetImport}
                  className="text-gray-400 hover:text-gray-600"
                >
                  ×
                </button>
              </div>

              {importStatus === 'idle' && (
                <div className="space-y-4">
                  <div className="text-sm text-gray-600 mb-4">
                    <p className="mb-2">Importieren Sie tägliche Schichtpläne um die Personalverfügbarkeit zu aktualisieren.</p>
                    <p className="text-xs text-gray-500">
                      Unterstützte Formate: .xlsx, .xls | Die Excel-Datei sollte Spalten für Schichttypen (BD, RD, Früh, Mittel, Spät) enthalten
                    </p>
                    <p className="text-xs text-yellow-600 mt-1">
                      ⚠️ Hinweis: Nur Personen die in einem Schichtplan zugewiesen sind werden als "verfügbar" markiert
                    </p>
                  </div>
                  
                  <div
                    {...getRootProps()}
                    className={`border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
                      isDragActive
                        ? 'border-blue-400 bg-blue-50'
                        : 'border-gray-300 hover:border-gray-400'
                    }`}
                  >
                    <input {...getInputProps()} />
                    <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
                    {isDragActive ? (
                      <p className="text-blue-600">Datei hier ablegen...</p>
                    ) : (
                      <div>
                        <p className="text-gray-600 mb-2">
                          Excel-Datei hier hinziehen oder klicken zum Auswählen
                        </p>
                        <p className="text-xs text-gray-500">
                          Maximal eine Datei, .xlsx oder .xls Format
                        </p>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {importStatus === 'processing' && (
                <div className="text-center py-8">
                  <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                  <p className="text-gray-600">Excel-Datei wird verarbeitet...</p>
                </div>
              )}

              {importStatus === 'error' && (
                <div className="text-center py-8">
                  <AlertCircle className="w-12 h-12 text-red-500 mx-auto mb-4" />
                  <p className="text-red-600 mb-4">{importError}</p>
                  <button
                    onClick={() => setImportStatus('idle')}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg"
                  >
                    Erneut versuchen
                  </button>
                </div>
              )}

              {importStatus === 'preview' && importResult && validationResult && (
                <div className="space-y-6">
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                    <h4 className="font-medium text-blue-900 mb-2">Schichtplan Import-Zusammenfassung</h4>
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                      <div>
                        <span className="text-blue-700">Arbeitsblätter:</span>
                        <div className="font-medium">{importResult.summary.sheetsProcessed.length}</div>
                      </div>
                      <div>
                        <span className="text-green-700">Gefundene Zuweisungen:</span>
                        <div className="font-medium text-green-800">{validationResult.valid.length}</div>
                      </div>
                      <div>
                        <span className="text-red-700">Fehlerhafte Einträge:</span>
                        <div className="font-medium text-red-800">{validationResult.invalid.length}</div>
                      </div>
                      <div>
                        <span className="text-gray-700">Zugewiesene Personen:</span>
                        <div className="font-medium">{importResult.summary.assignedPersonnel}</div>
                      </div>
                    </div>
                    {importResult.summary.shiftDate && (
                      <div className="mt-2 text-sm text-blue-700">
                        <span>Schichtdatum: </span>
                        <span className="font-medium">{importResult.summary.shiftDate}</span>
                      </div>
                    )}
                  </div>

                  {validationResult.invalid.length > 0 && (
                    <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                      <h4 className="font-medium text-yellow-900 mb-2">
                        Warnungen ({validationResult.invalid.length})
                      </h4>
                      <div className="max-h-32 overflow-y-auto">
                        {validationResult.invalid.slice(0, 5).map((item, index) => (
                          <div key={index} className="text-sm text-yellow-700 mb-1">
                            {item.assignment.name}: {item.reason}
                          </div>
                        ))}
                        {validationResult.invalid.length > 5 && (
                          <div className="text-xs text-yellow-600">
                            ... und {validationResult.invalid.length - 5} weitere
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  {importResult.errors.length > 0 && (
                    <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                      <h4 className="font-medium text-red-900 mb-2">
                        Verarbeitungsfehler ({importResult.errors.length})
                      </h4>
                      <div className="max-h-32 overflow-y-auto">
                        {importResult.errors.slice(0, 3).map((error, index) => (
                          <div key={index} className="text-sm text-red-700 mb-1">
                            {error}
                          </div>
                        ))}
                        {importResult.errors.length > 3 && (
                          <div className="text-xs text-red-600">
                            ... und {importResult.errors.length - 3} weitere
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  {previewData.length > 0 && (
                    <div>
                      <h4 className="font-medium text-gray-900 mb-3">
                        Vorschau der Schichtzuweisungen ({previewData.length})
                      </h4>
                      <div className="border border-gray-200 rounded-lg overflow-hidden">
                        <div className="max-h-64 overflow-y-auto">
                          <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50 sticky top-0">
                              <tr>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500">Name</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500">Schichttyp</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500">Verfügbarkeit</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500">Original Text</th>
                              </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                              {previewData.slice(0, 10).map((assignment, index) => (
                                <tr key={index} className="hover:bg-gray-50">
                                  <td className="px-3 py-2 text-sm text-gray-900">{assignment.name}</td>
                                  <td className="px-3 py-2 text-sm">
                                    <span className="inline-flex px-2 py-1 text-xs font-semibold rounded-full bg-blue-100 text-blue-800">
                                      {assignment.shiftType}
                                    </span>
                                  </td>
                                  <td className="px-3 py-2 text-sm">
                                    <span className="inline-flex px-2 py-1 text-xs font-semibold rounded-full bg-green-100 text-green-800">
                                      {assignment.availability}
                                    </span>
                                  </td>
                                  <td className="px-3 py-2 text-sm text-gray-400 max-w-xs truncate" title={assignment.originalText}>
                                    {assignment.originalText}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                          {previewData.length > 10 && (
                            <div className="px-3 py-2 text-xs text-gray-500 bg-gray-50 text-center">
                              ... und {previewData.length - 10} weitere Zuweisungen
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="flex justify-end gap-3">
                    <button
                      onClick={resetImport}
                      className="px-4 py-2 text-gray-600 hover:text-gray-800 transition-colors"
                    >
                      Abbrechen
                    </button>
                    <button
                      onClick={handleConfirmImport}
                      disabled={validationResult.valid.length === 0}
                      className="bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white px-6 py-2 rounded-lg transition-colors"
                    >
                      {validationResult.valid.length} Personen importieren
                    </button>
                  </div>
                </div>
              )}

              {importStatus === 'importing' && (
                <div className="text-center py-8">
                  <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-green-600 mx-auto mb-4"></div>
                  <p className="text-gray-600">Daten werden importiert...</p>
                </div>
              )}

              {importStatus === 'complete' && (
                <div className="text-center py-8">
                  <CheckCircle className="w-12 h-12 text-green-500 mx-auto mb-4" />
                  <h4 className="text-lg font-medium text-green-900 mb-2">Schichtplan Import erfolgreich!</h4>
                  <p className="text-gray-600 mb-4">
                    {validationResult?.valid.length} Schichtzuweisungen wurden verarbeitet und die Personalverfügbarkeit wurde aktualisiert.
                  </p>
                  <div className="text-sm text-gray-500 mb-4">
                    Verfügbare Personen werden nun nur noch in der Seitenleiste des Planners angezeigt.
                  </div>
                  <button
                    onClick={resetImport}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg"
                  >
                    Schließen
                  </button>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Personnel Table */}
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden">
          <div className="px-6 py-4 border-b border-gray-200">
            <h3 className="text-lg font-medium text-gray-900">
              {selectedGroup === 'all' ? 'Alle Mitarbeiter' : `Gruppe: ${selectedGroup}`} 
              <span className="text-sm text-gray-500 ml-2">({filteredPersonnel.length})</span>
            </h3>
          </div>
          
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Gruppe</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Abteilung</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Verfügbarkeit</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Schichtzuweisung</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Kommentar</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Aktionen</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {filteredPersonnel.map((person) => (
                <tr key={person.id} className="hover:bg-gray-50">
                  <td className="px-6 py-4 whitespace-nowrap">
                    <div className="flex items-center">
                      <div className="w-8 h-8 bg-gradient-to-br from-blue-500 to-blue-600 rounded-full flex items-center justify-center text-white font-semibold text-sm mr-3">
                        {person.initials}
                      </div>
                      <div className="text-sm font-medium text-gray-900">
                        {person.name}
                      </div>
                    </div>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <span className="inline-flex px-2 py-1 text-xs font-semibold rounded-full bg-blue-100 text-blue-800">
                      {person.gruppe}
                    </span>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{person.department}</td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                      person.verfügbar === 'nicht verfügbar' 
                        ? 'bg-red-100 text-red-800' 
                        : 'bg-green-100 text-green-800'
                    }`}>
                      {person.verfügbar}
                    </span>
                  </td>
                  <td className="px-6 py-4 whitespace-nowrap">
                    {person.shiftAssignment ? (
                      <span className="inline-flex px-2 py-1 text-xs font-semibold rounded-full bg-purple-100 text-purple-800">
                        {person.shiftAssignment}
                      </span>
                    ) : (
                      <span className="text-gray-400 text-xs">Keine Zuweisung</span>
                    )}
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500 max-w-xs truncate">{person.kommentar}</td>
                  <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                    <button
                      onClick={() => handleEditPerson(person)}
                      className="text-blue-600 hover:text-blue-900 mr-3"
                    >
                      Bearbeiten
                    </button>
                    <button
                      onClick={() => handleDeletePerson(person.id)}
                      className="text-red-600 hover:text-red-900"
                    >
                      Löschen
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          
          {filteredPersonnel.length === 0 && (
            <div className="px-6 py-8 text-center text-gray-500">
              {selectedGroup === 'all' 
                ? 'Keine Mitarbeiter gefunden' 
                : `Keine Mitarbeiter in der Gruppe "${selectedGroup}" gefunden`
              }
            </div>
          )}
        </div>
      </div>
    );
  };

  const SettingsPage = () => {
    return (
      <div className="p-6 max-w-4xl mx-auto">
        <h2 className="text-2xl font-bold text-gray-900 mb-6">Einstellungen</h2>
        <div className="space-y-6">
          <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">Lokale Daten</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
                <div className="text-2xl font-bold text-blue-600">{personnel.length}</div>
                <div className="text-sm text-blue-800">Gesamt Personal</div>
              </div>
              <div className="bg-gray-50 border border-gray-200 rounded-lg p-4 text-center">
                <div className="text-2xl font-bold text-gray-600">
                  {personnel.filter(p => p.verfügbar !== 'nicht verfügbar').length}
                </div>
                <div className="text-sm text-gray-800">Verfügbares Personal</div>
              </div>
            </div>
            <p className="text-gray-600 mb-4">
              Alle Daten werden lokal im Browser gespeichert und bleiben auch nach dem Neuladen verfügbar.
            </p>
            <div className="flex gap-3">
              <button 
                onClick={() => exportToExcel()}
                className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg transition-colors"
              >
                Personal als Excel exportieren
              </button>
              <button 
                onClick={clearAllData}
                className="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg transition-colors"
              >
                Alle Daten löschen
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderCurrentPage = () => {
    switch (currentPage) {
      case 'planner': return <PlannerPage />;
      case 'personnel': return <PersonnelPage />;
      case 'legend': return <LegendPage />;
      case 'settings': return <SettingsPage />;
      default: return <PlannerPage />;
    }
  };

  return (
    <div className="h-screen flex flex-col bg-gray-50">
      <NavigationBar />
      <div className="flex-1 overflow-auto">
        {renderCurrentPage()}
      </div>
    </div>
  );
};

const App = () => {
  // Return the main OR Planner application
  return (
    <ORPlannerApp />
  );
};

export default App;