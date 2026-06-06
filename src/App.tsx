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
    <div className="bg-[#09090b] min-h-screen text-zinc-100 flex flex-col font-sans selection:bg-[#a3e635]/30 selection:text-white">
      
      {/* 🛠️ COMBINED DEV PREVIEW SIMULATOR PANEL */}
      <section className="bg-zinc-950 border-b border-[#27272a] z-35 font-sans text-xs">
        <div className="p-2 px-4 flex items-center justify-between gap-4">
          <div className="flex items-center gap-2.5">
            <Sliders className="w-4 h-4 text-[#a3e635] shrink-0" />
            <div>
              <span className="font-mono text-[10px] uppercase font-bold text-zinc-505 leading-none">Simulado de PowerPoint</span>
              <h2 className="font-black text-zinc-200 text-xs leading-tight">Alternar Entre as Visualizações do Add-In</h2>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setIsSimulatorExpanded(!isSimulatorExpanded)}
              className="p-1 px-2 rounded hover:bg-zinc-900 text-zinc-400 hover:text-zinc-200 flex items-center gap-1 cursor-pointer transition-colors"
            >
              <span className="text-[10px] font-semibold">{isSimulatorExpanded ? 'Recolher Painel' : 'Expandir'}</span>
              {isSimulatorExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
            </button>
            
            <button
              onClick={executeOpenInNewTab}
              className="p-1 px-2 rounded hover:bg-zinc-900 text-orange-400 hover:text-orange-300 flex items-center gap-1 cursor-pointer transition-colors"
              title="Abrir atual visualização em nova aba para testar o tempo real simultâneo!"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span className="text-[10px] font-semibold hidden md:inline">Nova Aba</span>
            </button>
          </div>
        </div>

        {/* Expanded selectors option */}
        {isSimulatorExpanded && (
          <div className="px-4 pb-3.5 pt-0.5 grid grid-cols-2 md:grid-cols-4 gap-2 border-t border-zinc-900 animate-in slide-in-from-top-1 duration-200">
            {/* BUTTON 1: OFFICE TASKPANE */}
            <button
              onClick={() => updateRole('addin')}
              className={`p-2 rounded-xl border text-left cursor-pointer transition-all ${
                role === 'addin'
                  ? 'bg-[#121214] border-[#a3e635] text-zinc-100 shadow-[0_0_12px_rgba(163,230,53,0.1)]'
                  : 'bg-[#121214]/40 border-[#27272a] hover:border-zinc-700 text-zinc-400 hover:text-zinc-200'
              }`}
            >
              <div className="flex items-center gap-2 mb-1">
                <Laptop className={`w-4 h-4 shrink-0 p-0.5 rounded ${role === 'addin' ? 'bg-[#a3e635] text-zinc-950 font-black' : 'bg-zinc-900 text-zinc-400'}`} />
                <span className="font-bold text-[11px]">🔌 Painel PowerPoint</span>
              </div>
              <p className="text-[10px] leading-tight text-zinc-550 line-clamp-2">Taskpane lateral do Add-In com botões de inserção de slide.</p>
            </button>

            {/* BUTTON 2: FULL WEB ADMINISTRACAO */}
            <button
              onClick={() => updateRole('instructor')}
              className={`p-2 rounded-xl border text-left cursor-pointer transition-all ${
                role === 'instructor'
                  ? 'bg-[#121214] border-[#a3e635] text-zinc-100 shadow-[0_0_12px_rgba(163,230,53,0.1)]'
                  : 'bg-[#121214]/40 border-[#27272a] hover:border-zinc-700 text-zinc-400 hover:text-zinc-200'
              }`}
            >
              <div className="flex items-center gap-2 mb-1">
                <Settings className={`w-4 h-4 shrink-0 p-0.5 rounded ${role === 'instructor' ? 'bg-[#a3e635] text-zinc-950 font-black' : 'bg-zinc-900 text-zinc-400'}`} />
                <span className="font-bold text-[11px]">⚙️ Painel Web Completo</span>
              </div>
              <p className="text-[10px] leading-tight text-zinc-550 line-clamp-2">Criação de novas aulas, editor de gabarito e exportação CSV.</p>
            </button>

            {/* BUTTON 3: PROJECTOR DISPLAY */}
            <button
              onClick={() => updateRole('tv')}
              className={`p-2 rounded-xl border text-left cursor-pointer transition-all ${
                role === 'tv'
                  ? 'bg-[#121214] border-[#a3e635] text-zinc-100 shadow-[0_0_12px_rgba(163,230,53,0.1)]'
                  : 'bg-[#121214]/40 border-[#27272a] hover:border-zinc-700 text-zinc-400 hover:text-zinc-200'
              }`}
            >
              <div className="flex items-center gap-2 mb-1">
                <Tv className={`w-4 h-4 shrink-0 p-0.5 rounded ${role === 'tv' ? 'bg-[#a3e635] text-zinc-950 font-black' : 'bg-zinc-900 text-zinc-400'}`} />
                <span className="font-bold text-[11px]">🖥️ Projetor (TV)</span>
              </div>
              <p className="text-[10px] leading-tight text-zinc-550 line-clamp-2">Exibe a pergunta em destaque grande, QR Code e resultados ao vivo.</p>
            </button>

            {/* BUTTON 4: CELLPHONE PARTICIPANT */}
            <button
              onClick={() => updateRole('participant')}
              className={`p-2 rounded-xl border text-left cursor-pointer transition-all ${
                role === 'participant'
                  ? 'bg-[#121214] border-[#a3e635] text-zinc-100 shadow-[0_0_12px_rgba(163,230,53,0.1)]'
                  : 'bg-[#121214]/40 border-[#27272a] hover:border-zinc-700 text-zinc-400 hover:text-zinc-200'
              }`}
            >
              <div className="flex items-center gap-2 mb-1">
                <Smartphone className={`w-4 h-4 shrink-0 p-0.5 rounded ${role === 'participant' ? 'bg-[#a3e635] text-zinc-950 font-black' : 'bg-zinc-900 text-zinc-400'}`} />
                <span className="font-bold text-[11px]">📱 Celular Aluno</span>
              </div>
              <p className="text-[10px] leading-tight text-zinc-550 line-clamp-2">Página mobile que escaneia o QR Code e vota nas alternativas do slide.</p>
            </button>
          </div>
        )}
      </section>

      {/* 🧭 DESKTOP MULTI-WINDOW TESTING TIP */}
      {isSimulatorExpanded && (
        <div className="bg-[#a3e635]/5 border-b border-zinc-900 text-zinc-300 text-[11px] p-2 px-4 flex items-center justify-center gap-1.5 font-medium font-mono">
          <Info className="w-4 h-4 text-[#a3e635] shrink-0 animate-pulse" />
          <span>
            💡 <strong>Dica de Teste em Tempo Real:</strong> Abra este applet em duas abas! Coloque uma aba em <strong>Celular Aluno</strong> e a outra aba em <strong>Projetor</strong>. Vote no Celular e veja o gráfico atualizar instantaneamente!
          </span>
        </div>
      )}

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
