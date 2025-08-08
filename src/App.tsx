import React, { useEffect, useState, useCallback, DragEvent } from 'react';
import { Search, RotateCcw, User, Calendar, Settings, FileText, Users } from 'lucide-react';
import { useSharePointPersonnel, authConfig } from './sharepoint-integration';
import { useMsal } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';

type TableConfigKey = 'main' | 'emergency' | 'weekend';

interface Personnel {
  id: number;
  name: string;
  role: string;
  initials: string;
  department: string;
}

const ORPlannerApp = (props: { personnel: any[] }) => {
  const [currentPage, setCurrentPage] = useState('planner');
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  
  // Table configurations based on your CSV structure
  const tableConfigs = {
    main: {
      name: "Hauptplan",
      roles: [
        "Anästhesie Arzt 1", "Anästhesie Arzt 2", "AA Praktikant", "",
        "Anästhesie Pflege", "Anästhesie Pflege", "ATA", "ATA", "Praktikant", "",
        "", "", "",
        "OP Pflege", "OP Pflege", "OTAS", "OTAS", "Praktikant", "Praktikant",
        "", "", "", ""
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
  const [assignments, setAssignments] = useState<Record<string, Personnel[]>>({});
  const [searchTerm, setSearchTerm] = useState("");
  const [draggedPerson, setDraggedPerson] = useState<Personnel | null>(null);

  // Sample personnel data (will be replaced with SharePoint data)
  const [personnel] = useState(
    props.personnel.length > 0 ? props.personnel : [
    { id: 1, name: "Dr. Sarah Weber", role: "Anästhesie Arzt", initials: "SW", department: "Anästhesie" },
    { id: 2, name: "Dr. Michael Koch", role: "Anästhesie Arzt", initials: "MK", department: "Anästhesie" },
    { id: 3, name: "Lisa Müller", role: "Anästhesie Pflege", initials: "LM", department: "Anästhesie" },
    { id: 4, name: "Thomas Schmidt", role: "OP Pflege", initials: "TS", department: "OP" },
    { id: 5, name: "Anna Becker", role: "OTAS", initials: "AB", department: "OP" },
    { id: 6, name: "Max Hoffmann", role: "ATA", initials: "MH", department: "Anästhesie" },
    { id: 7, name: "Julia Wagner", role: "Praktikant", initials: "JW", department: "OP" },
    { id: 8, name: "Daniel Richter", role: "AA Praktikant", initials: "DR", department: "Anästhesie" }
  ]);

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
    </nav>
  );

  const PlannerPage = () => { 
    const config = tableConfigs[currentTable];
    
    const filteredPersonnel = personnel.filter(person =>
      person.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      person.role.toLowerCase().includes(searchTerm.toLowerCase())
    );

    const unassignedPersonnel = filteredPersonnel.filter(person =>
      !Object.values(assignments).flat().some((assigned: Personnel) => assigned.id === person.id)
    );

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
        
        Object.keys(newAssignments).forEach(key => {
          newAssignments[key] = newAssignments[key].filter(p => p.id !== draggedPerson.id);
          if (newAssignments[key].length === 0) {
            delete newAssignments[key];
          }
        });
        
        if (!newAssignments[cellKey]) {
          newAssignments[cellKey] = [];
        }
        newAssignments[cellKey].push(draggedPerson);
        
        return newAssignments;
      });
      
      setDraggedPerson(null);
    }, [draggedPerson, currentTable]);

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
              {person.role}
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
          {/* Table Selector */}
          <div className="bg-gray-50 border-b border-gray-200 px-6 py-3">
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
                {config.roles.map((role, roleIndex) => (
                  <div key={roleIndex} className="flex hover:bg-gray-50">
                    <div className="w-48 p-3 font-medium text-gray-700 bg-white border-r border-gray-200 sticky left-0 z-10">
                      {role}
                    </div>
                    
                    {config.rooms.map((room, roomIndex) => {
                      const cellKey = `${currentTable}-${roleIndex}-${roomIndex}`;
                      const assignedPersons = assignments[cellKey] || [];
                      
                      return (
                        <div
                          key={roomIndex}
                          className="w-32 min-h-16 p-2 border-r border-gray-200 bg-white hover:bg-blue-50 transition-colors"
                          onDragOver={handleDragOver}
                          onDrop={(e) => handleDrop(e, roleIndex, roomIndex)}
                        >
                          <div className="space-y-1">
                            {assignedPersons.map((person) => (
                              <PersonCard
                                key={person.id}
                                person={person}
                                isDraggable={true}
                                onRemove={(personId: number) => {
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
                                }}
                                size="small"
                              />
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Personnel Sidebar */}
        <div className="w-80 bg-white border-l border-gray-200 flex flex-col">
          <div className="p-4 border-b border-gray-200">
            <h2 className="text-lg font-semibold text-gray-900 mb-3">Personenauswahl</h2>
            
            <div className="relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" size={16} />
              <input
                type="text"
                placeholder="Suche..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none"
              />
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-4">
            <div className="space-y-3">
              {unassignedPersonnel.length === 0 ? (
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

  const PersonnelPage = () => (
    <div className="p-6">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Personalübersicht</h2>
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50">
            <tr>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Rolle</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Abteilung</th>
              <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {personnel.map((person) => (
              <tr key={person.id}>
                <td className="px-6 py-4 whitespace-nowrap">
                  <div className="flex items-center">
                    <div className="w-8 h-8 bg-gradient-to-br from-blue-500 to-blue-600 rounded-full flex items-center justify-center text-white font-semibold text-sm mr-3">
                      {person.initials}
                    </div>
                    <div className="text-sm font-medium text-gray-900">{person.name}</div>
                  </div>
                </td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{person.role}</td>
                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{person.department}</td>
                <td className="px-6 py-4 whitespace-nowrap">
                  <span className="inline-flex px-2 py-1 text-xs font-semibold rounded-full bg-green-100 text-green-800">
                    Aktiv
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );

  const SettingsPage = () => (
    <div className="p-6 max-w-4xl mx-auto">
      <h2 className="text-2xl font-bold text-gray-900 mb-6">Einstellungen</h2>
      <div className="space-y-6">
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">SharePoint Integration</h3>
          <p className="text-gray-600 mb-4">Konfiguration der SharePoint-Verbindung für automatische Personalaktualisierung.</p>
          <button className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg transition-colors">
            Verbindung testen
          </button>
        </div>
      </div>
    </div>
  );

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
      <div className="flex-1 overflow-hidden">
        {renderCurrentPage()}
      </div>
    </div>
  );
};

const App = () => {
  const { instance, accounts, inProgress } = useMsal();
  const [accessToken, setAccessToken] = React.useState<string | null>(null);
  const [authError, setAuthError] = React.useState<string | null>(null);
  const [debugInfo, setDebugInfo] = React.useState<any>({});
  
  const SITE_ID = 'a4cba12d-b1bf-4542-9ca4-7563ad6b7b09';
  const PERSONNEL_LIST_ID = '67d0c026-90c1-4822-b27f-66f73c8139e5';
  
  const { personnel, loading, error } = useSharePointPersonnel(
    accessToken, 
    SITE_ID, 
    PERSONNEL_LIST_ID
  );

  useEffect(() => {
    const getToken = async () => {
      try {
        console.log('Auth state:', { 
          accountsCount: accounts.length, 
          inProgress, 
          hasActiveAccount: !!instance.getActiveAccount() 
        });
        
        setDebugInfo({
          accountsCount: accounts.length,
          inProgress,
          hasActiveAccount: !!instance.getActiveAccount(),
          accounts: accounts.map(acc => ({ username: acc.username, name: acc.name }))
        });

        if (accounts.length > 0) {
          console.log('Attempting silent token acquisition...');
          
          const response = await instance.acquireTokenSilent({
            scopes: authConfig.scopes,
            account: accounts[0]
          });
          
          console.log('Token acquired successfully:', {
            tokenLength: response.accessToken.length,
            account: response.account?.username
          });
          
          setAccessToken(response.accessToken);
          setAuthError(null);
        } else if (inProgress === "none") {
          console.log('No accounts found, redirecting to login...');
          await instance.loginRedirect({
            scopes: authConfig.scopes
          });
        }
      } catch (error) {
        console.error('Token acquisition failed:', error);
        setAuthError(error instanceof Error ? error.message : 'Authentication failed');
        
        if (error instanceof InteractionRequiredAuthError) {
          console.log('Interactive login required, redirecting...');
          try {
            await instance.loginRedirect({
              scopes: authConfig.scopes
            });
          } catch (redirectError) {
            console.error('Redirect failed:', redirectError);
            setAuthError('Login redirect failed: ' + (redirectError instanceof Error ? redirectError.message : 'Unknown error'));
          }
        }
      }
    };

    getToken();
  }, [accounts, instance, inProgress]);

  if (inProgress !== "none" || loading) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-green-600 mx-auto mb-4"></div>
          <p className="mb-4">
            {inProgress !== "none" ? 'Authenticating...' : 'Loading personnel...'}
          </p>
          <div className="text-xs text-gray-500 bg-gray-100 p-2 rounded max-w-md">
            <pre>{JSON.stringify(debugInfo, null, 2)}</pre>
          </div>
        </div>
      </div>
    );
  }

  if (authError) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="text-center text-red-600 max-w-md p-6 bg-white rounded-lg shadow">
          <h2 className="text-xl font-bold mb-4">Authentication Error</h2>
          <p className="mb-4 text-sm">{authError}</p>
          <div className="text-xs text-gray-500 bg-gray-100 p-2 rounded mb-4">
            <strong>Debug Info:</strong>
            <pre>{JSON.stringify(debugInfo, null, 2)}</pre>
          </div>
          <div className="space-y-2">
            <button
              className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 block w-full"
              onClick={() => {
                setAuthError(null);
                instance.loginRedirect({ scopes: authConfig.scopes });
              }}
            >
              Sign In Again
            </button>
            <button
              className="bg-gray-600 text-white px-4 py-2 rounded hover:bg-gray-700 block w-full"
              onClick={() => {
                setAuthError(null);
                setAccessToken(null);
                // Continue with mock data
              }}
            >
              Continue with Mock Data
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (!accessToken && accounts.length === 0) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="text-center max-w-md p-6 bg-white rounded-lg shadow">
          <h2 className="text-xl font-bold mb-4">Sign In Required</h2>
          <p className="mb-4 text-gray-600">Please sign in to access SharePoint data.</p>
          <button
            className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700"
            onClick={() => instance.loginRedirect({ scopes: authConfig.scopes })}
          >
            Sign In
          </button>
        </div>
      </div>
    );
  }

  // Return the main OR Planner application
  return <ORPlannerApp personnel={personnel || []} />;
};

export default App;