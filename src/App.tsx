/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { syncStore, subscribeToSyncChanges } from './services/store';
import { PowerPointAddinTaskPane } from './pages/PowerPointAddinTaskPane';
import { DashboardInstrutor } from './pages/DashboardInstrutor';
import { TelaApresentacao } from './pages/TelaApresentacao';
import { TelaParticipante } from './pages/TelaParticipante';
import { Session } from './types';
import { 
  Laptop, 
  Tv, 
  Smartphone, 
  Settings, 
  Info, 
  Check, 
  ChevronRight, 
  ExternalLink,
  ChevronDown,
  ChevronUp,
  Sliders,
  HelpCircle
} from 'lucide-react';

export default function App() {
  // App initialization on boot
  useEffect(() => {
    syncStore.init();
  }, []);

  // Parse initial role & session parameters from URL to support direct bookmark accesses
  const getInitialState = () => {
    const params = new URLSearchParams(window.location.search);
    const urlRole = params.get('role') || 'addin';
    const urlSession = params.get('session') || '8825'; // Default pre-seeded beautiful demo session
    return { urlRole, urlSession };
  };

  const { urlRole, urlSession } = getInitialState();

  // App running states
  const [role, setRole] = useState<string>(urlRole);
  const [activeSessionId, setActiveSessionId] = useState<string>(urlSession);
  const [activeSession, setActiveSession] = useState<Session | null>(null);
  const [refreshTrigger, setRefreshTrigger] = useState<number>(0);
  const [isSimulatorExpanded, setIsSimulatorExpanded] = useState<boolean>(true);

  // Dynamic host URL generation
  const appUrl = `${window.location.protocol}//${window.location.host}${window.location.pathname}`;

  // State loading
  useEffect(() => {
    const s = syncStore.getSession(activeSessionId);
    if (s) {
      setActiveSession(s);
    } else {
      setActiveSession(null);
    }
  }, [activeSessionId, refreshTrigger]);

  // Subscribe to changes (updates from other tabs)
  useEffect(() => {
    const unsubscribe = subscribeToSyncChanges((evt) => {
      setRefreshTrigger(prev => prev + 1);
    });
    return () => unsubscribe();
  }, []);

  const handleRefresh = () => {
    setRefreshTrigger(prev => prev + 1);
  };

  const handleSelectSession = (id: string) => {
    setActiveSessionId(id);
    handleRefresh();
  };

  // Change active simulated role manually
  const updateRole = (newRole: string) => {
    setRole(newRole);
    // Push parameter to browser query bar to simplify bookmarking / reload testing
    const params = new URLSearchParams(window.location.search);
    params.set('role', newRole);
    if (activeSessionId) {
      params.set('session', activeSessionId);
    }
    const newRelativePathQuery = window.location.pathname + '?' + params.toString();
    window.history.pushState(null, '', newRelativePathQuery);
    handleRefresh();
  };

  const executeOpenInNewTab = () => {
    const targetUrl = `${appUrl}?role=${role}&session=${activeSessionId}`;
    window.open(targetUrl, '_blank');
  };

  return (
    <div className="bg-[#f8fafc] min-h-screen text-slate-800 flex flex-col font-sans selection:bg-[#2d7f8d]/20 selection:text-[#2d7f8d]">
      
      {/* CORE ROUTING SHELL */}
      <div className="flex-1 flex flex-col justify-stretch">
        {role === 'addin' && (
          <PowerPointAddinTaskPane
            onSelectSession={handleSelectSession}
            onOpenCreateSession={() => updateRole('instructor')}
            activeSession={activeSession}
            onRefresh={handleRefresh}
            appUrl={appUrl}
          />
        )}
        
        {role === 'instructor' && (
          <DashboardInstrutor
            onSelectSession={handleSelectSession}
            onSelectRole={updateRole}
            activeSession={activeSession}
            onRefresh={handleRefresh}
            appUrl={appUrl}
          />
        )}

        {role === 'tv' && (
          <TelaApresentacao
            sessionId={activeSessionId}
            onSelectSession={handleSelectSession}
            appUrl={appUrl}
          />
        )}

        {role === 'participant' && (
          <TelaParticipante
            initialSessionId={activeSessionId}
            onRefresh={handleRefresh}
          />
        )}
      </div>
    </div>
  );
}
