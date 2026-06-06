/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { syncStore } from '../services/store';
import { Session, Question, QuestionBankItem, Vote } from '../types';
import { 
  Plus, 
  Layers, 
  FileText, 
  Trash2, 
  Edit2, 
  Check, 
  Copy, 
  Play, 
  BarChart, 
  Settings, 
  Share2, 
  Download, 
  ExternalLink,
  ChevronRight,
  HelpCircle,
  ToggleLeft,
  ToggleRight,
  Archive,
  ArrowRight,
  User,
  Users,
  Database,
  BarChart2,
  RefreshCw,
  FolderPlus
} from 'lucide-react';
import { QRCodeDisplay } from '../components/QRCodeDisplay';

interface DashboardInstrutorProps {
  onSelectSession: (id: string) => void;
  onSelectRole: (role: string) => void;
  activeSession: Session | null;
  onRefresh: () => void;
  appUrl: string;
}

export const DashboardInstrutor: React.FC<DashboardInstrutorProps> = ({
  onSelectSession,
  onSelectRole,
  activeSession,
  onRefresh,
  appUrl
}) => {
  // Navigation states: 'sessions' | 'questions' | 'bank'
  const [activeTab, setActiveTab] = useState<'sessions' | 'questions' | 'bank'>('sessions');
  const [sessions, setSessions] = useState<Session[]>([]);
  const [questionBank, setQuestionBank] = useState<QuestionBankItem[]>([]);
  
  // Selection
  const [selectedSession, setSelectedSession] = useState<Session | null>(activeSession);

  // Creation State - Session
  const [newSessionName, setNewSessionName] = useState('');
  const [isAnonymous, setIsAnonymous] = useState(false);
  const [isQuizMode, setIsQuizMode] = useState(true);

  // Creation/Edit State - Question
  const [isEditingQuestion, setIsEditingQuestion] = useState<boolean>(false);
  const [editingQuestionId, setEditingQuestionId] = useState<string | null>(null); // null means adding new
  const [questionText, setQuestionText] = useState('');
  const [options, setOptions] = useState<string[]>(['', '', '', '']); // Default 4 options
  const [correctOptionIdx, setCorrectOptionIdx] = useState<number | null>(null); // null means survey question
  
  // Bank Item Creating State
  const [bankCategory, setBankCategory] = useState('Geral');

  const [notification, setNotification] = useState<string | null>(null);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState<string | null>(null);

  // Load from store on boot
  useEffect(() => {
    loadData();
  }, [activeSession]);

  const loadData = () => {
    const list = syncStore.getSessions();
    setSessions(list);
    setQuestionBank(syncStore.getQuestionBank());
    
    // Maintain active select
    if (selectedSession) {
      const refreshedIdx = list.find(s => s.id === selectedSession.id);
      if (refreshedIdx) setSelectedSession(refreshedIdx);
    }
  };

  const showToast = (msg: string) => {
    setNotification(msg);
    setTimeout(() => setNotification(null), 3000);
  };

  // SESSIONS HANDLING
  const handleCreateSession = (e: React.FormEvent) => {
    e.preventDefault();
    const created = syncStore.createSession(newSessionName.trim() || 'Aula Teórica', isAnonymous, isQuizMode);
    setNewSessionName('');
    showToast(`Sessão "${created.name}" criada de forma bem sucedida!`);
    loadData();
    
    // Auto select & focus
    setSelectedSession(created);
    onSelectSession(created.id);
    setActiveTab('questions');
  };

  const handleDeleteSession = (id: string) => {
    syncStore.deleteSession(id);
    if (selectedSession?.id === id) {
      setSelectedSession(null);
      onSelectSession('');
    }
    showToast('Sessão removida.');
    setShowDeleteConfirm(null);
    loadData();
  };

  // QUESTIONS HANDLING
  const handleOpenNewQuestionForm = () => {
    setQuestionText('');
    setOptions(['', '', '', '']); // standard A, B, C, D
    setCorrectOptionIdx(null);
    setEditingQuestionId(null);
    setIsEditingQuestion(true);
  };

  const handleOpenEditQuestionForm = (q: Question) => {
    setQuestionText(q.text);
    setOptions([...q.options]);
    setCorrectOptionIdx(q.correctOptionIndex);
    setEditingQuestionId(q.id);
    setIsEditingQuestion(true);
  };

  const handleCloseQuestionForm = () => {
    setIsEditingQuestion(false);
    setEditingQuestionId(null);
  };

  const handleAddOptionField = () => {
    if (options.length >= 8) {
      showToast('O número limite é de 8 alternativas.');
      return;
    }
    setOptions([...options, '']);
  };

  const handleRemoveOptionField = (idx: number) => {
    if (options.length <= 2) {
      showToast('A pergunta precisa ter pelo menos 2 alternativas.');
      return;
    }
    const copy = [...options];
    copy.splice(idx, 1);
    setOptions(copy);
    if (correctOptionIdx === idx) {
      setCorrectOptionIdx(null);
    } else if (correctOptionIdx !== null && correctOptionIdx > idx) {
      setCorrectOptionIdx(correctOptionIdx - 1);
    }
  };

  const handleOptionTextChange = (text: string, idx: number) => {
    const copy = [...options];
    copy[idx] = text;
    setOptions(copy);
  };

  const handleSaveQuestion = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedSession) return;
    if (questionText.trim() === '') {
      showToast('Por favor, digite o enunciado da pergunta.');
      return;
    }

    const filteredOptions = options.map(o => o.trim()).filter(o => o !== '');
    if (filteredOptions.length < 2) {
      showToast('Preencha pelo menos 2 alternativas preenchidas.');
      return;
    }

    // Adapt corrected answer index if values changed
    let finalCorrectIdx = correctOptionIdx;
    if (finalCorrectIdx !== null && finalCorrectIdx >= filteredOptions.length) {
      finalCorrectIdx = null;
    }

    if (editingQuestionId) {
      // Modify
      syncStore.updateQuestionInSession(selectedSession.id, editingQuestionId, questionText.trim(), filteredOptions, finalCorrectIdx);
      showToast('Pergunta atualizada com sucesso!');
    } else {
      // Add brand new
      syncStore.addQuestionToSession(selectedSession.id, questionText.trim(), filteredOptions, finalCorrectIdx);
      showToast('Pergunta adicionada à sessão!');
    }

    setIsEditingQuestion(false);
    loadData();
    onRefresh();
  };

  const handleDeleteQuestion = (qId: string) => {
    if (!selectedSession) return;
    syncStore.deleteQuestionFromSession(selectedSession.id, qId);
    showToast('Pergunta removida.');
    loadData();
    onRefresh();
  };

  const handleDuplicateQuestion = (q: Question) => {
    if (!selectedSession) return;
    syncStore.addQuestionToSession(selectedSession.id, `(Cópia) ${q.text}`, [...q.options], q.correctOptionIndex);
    showToast('Pergunta duplicada com sucesso!');
    loadData();
    onRefresh();
  };

  // QUESTION BANK INTEGRATION
  const handleSaveToBank = (q: Question) => {
    syncStore.addQuestionToBank(q.text, q.options, q.correctOptionIndex, selectedSession?.name || 'Geral');
    showToast('Pergunta salva no Banco Geral!');
    loadData();
  };

  const handleImportFromBank = (bankItem: QuestionBankItem) => {
    if (!selectedSession) {
      showToast('Selecione ou crie uma sessão primeiro para poder importar perguntas!');
      return;
    }
    syncStore.addQuestionToSession(selectedSession.id, bankItem.text, [...bankItem.options], bankItem.correctOptionIndex);
    showToast('Pergunta importada do Banco de Questões.');
    loadData();
    onRefresh();
  };

  const handleDeleteFromBank = (id: string) => {
    syncStore.deleteQuestionFromBank(id);
    showToast('Item removido do Banco.');
    loadData();
  };

  // EXPORTS METHODS
  const handleDownloadCSV = (sessionId: string) => {
    const csvContent = syncStore.exportToCSV(sessionId);
    if (!csvContent) return;

    // Create blobs for standard local download in UTF-8
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `relatorio-enquete-sessao-${sessionId}.csv`;
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('Download do CSV iniciado!');
  };

  const handleDownloadJSON = (sessionId: string) => {
    const jsonContent = syncStore.exportToJSON(sessionId);
    const blob = new Blob([jsonContent], { type: 'application/json;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `dados-enquete-sessao-${sessionId}.json`;
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('Dados JSON exportados com sucesso!');
  };

  const syncActivePresentationSession = (s: Session) => {
    setSelectedSession(s);
    onSelectSession(s.id);
    showToast(`Controlando sessão ${s.id} agora.`);
  };

  // Simulate multiple participants typing and joining
  const handleAddSimulatedParticipants = () => {
    if (!selectedSession) return;
    syncStore.joinOrCreateParticipant(selectedSession.id, 'Carlos Magno');
    syncStore.joinOrCreateParticipant(selectedSession.id, 'Capitão Silva');
    syncStore.joinOrCreateParticipant(selectedSession.id, 'Letícia Rezende');
    syncStore.joinOrCreateParticipant(selectedSession.id, 'Tenente Moura');
    syncStore.joinOrCreateParticipant(selectedSession.id, 'Professora Marta');
    showToast('+5 Alunos virtuais conectados na sessão!');
    loadData();
  };

  return (
    <div className="bg-slate-950 text-slate-100 min-h-screen py-6 px-4 md:px-8 max-w-7xl mx-auto flex flex-col md:flex-row gap-6">
      {/* Toast Notification */}
      {notification && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-slate-100 border border-emerald-500/40 rounded-xl px-4 py-3 shadow-2xl flex items-center gap-3 animate-in fade-in slide-in-from-bottom-2">
          <Check className="w-5 h-5 text-emerald-400 shrink-0" />
          <span className="text-xs font-semibold">{notification}</span>
        </div>
      )}

      {/* LEFT SIDE PANEL: SESSION MANAGEMENT */}
      <section className="w-full md:w-80 shrink-0 space-y-4">
        {/* Navigation role picker helpful card */}
        <div className="p-4 rounded-xl bg-[#121214]/50 border border-zinc-800 text-xs">
          <h4 className="font-bold text-zinc-300 flex items-center gap-1.5 mb-1.5 text-sm">
            <Settings className="w-4 h-4 text-[#a3e635] animate-spin" />
            Central do Instrutor
          </h4>
          <p className="text-zinc-400 font-medium leading-relaxed">
            Painel principal de gerenciamento. Daqui você monta as enquetes pedagógicas e gera os códigos de interação rápida para PowerPoint.
          </p>
        </div>

        {/* Creation card */}
        <div className="p-4 rounded-xl bg-[#121214] border border-[#27272a] shadow-lg animate-in fade-in zoom-in-95 duration-200">
          <h4 className="font-bold text-white mb-3 text-xs uppercase tracking-wider flex items-center gap-1.5 font-display">
            <Plus className="w-4 h-4 text-[#a3e635]" />
            Nova Sessão
          </h4>
          
          <form onSubmit={handleCreateSession} className="space-y-3.5">
            <div>
              <label className="block text-[11px] font-mono text-zinc-550 mb-1">Nome da Aula / Briefing</label>
              <input
                type="text"
                value={newSessionName}
                onChange={(e) => setNewSessionName(e.target.value)}
                placeholder="Ex: Aula de Marketing Slides"
                className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2.5 focus:border-[#a3e635] focus:outline-none focus:ring-1 focus:ring-[#a3e635]/20 placeholder-zinc-700 transition-colors font-semibold"
                id="input-session-name"
              />
            </div>

            {/* Config Toggles */}
            <div className="space-y-2">
              <div 
                onClick={() => setIsAnonymous(!isAnonymous)}
                className="flex items-center justify-between p-1.5 rounded bg-[#09090b]/80 border border-zinc-800 cursor-pointer select-none"
              >
                <div className="flex flex-col">
                  <span className="text-xs font-bold text-zinc-350">Votos Anônimos</span>
                  <span className="text-[10px] text-zinc-550">Participantes entram sem cadastrar nome</span>
                </div>
                {isAnonymous ? (
                  <ToggleRight className="w-6 h-6 text-[#a3e635] shrink-0" />
                ) : (
                  <ToggleLeft className="w-6 h-6 text-zinc-700 shrink-0" />
                )}
              </div>

              <div 
                onClick={() => setIsQuizMode(!isQuizMode)}
                className="flex items-center justify-between p-1.5 rounded bg-[#09090b]/80 border border-zinc-800 cursor-pointer select-none"
              >
                <div className="flex flex-col">
                  <span className="text-xs font-bold text-zinc-350">Modo Quiz (c/ Notas)</span>
                  <span className="text-[10px] text-zinc-550">Apoia gabarito e ranking de alunos</span>
                </div>
                {isQuizMode ? (
                  <ToggleRight className="w-6 h-6 text-[#a3e635] shrink-0" />
                ) : (
                  <ToggleLeft className="w-6 h-6 text-zinc-700 shrink-0" />
                )}
              </div>
            </div>

            <button
              type="submit"
              className="w-full bg-[#a3e635] hover:bg-[#a3e635]/90 text-zinc-950 font-black text-xs py-2 px-4 rounded-lg flex items-center justify-center gap-1.5 transition-all shadow cursor-pointer uppercase font-mono tracking-wider"
              id="btn-create-session-submit"
            >
              <FolderPlus className="w-4 h-4 text-zinc-950" />
              <span>Criar Sessão</span>
            </button>
          </form>
        </div>

        {/* Existing Session Navigator */}
        <div className="p-4 rounded-xl bg-[#121214]/40 border border-[#27272a] space-y-3">
          <div className="flex items-center justify-between text-xs font-semibold pb-1.5 border-b border-zinc-850">
            <span className="text-zinc-400 font-mono">Sessões Disponíveis</span>
            <span className="bg-zinc-800 text-zinc-450 px-1.5 py-0.5 rounded text-[10px] font-mono">{sessions.length}</span>
          </div>

          <div className="space-y-1.5 max-h-56 overflow-y-auto pr-1">
            {sessions.map((sess) => {
              const isActiveLocal = selectedSession?.id === sess.id;
              return (
                <div
                  key={sess.id}
                  onClick={() => syncActivePresentationSession(sess)}
                  className={`p-2 rounded-lg border text-left cursor-pointer transition-all flex items-center justify-between ${
                    isActiveLocal 
                      ? 'bg-[#a3e635]/10 border-[#a3e635]/65 text-white shadow-sm font-semibold' 
                      : 'bg-[#09090b]/45 border-zinc-800/80 text-zinc-400 hover:bg-[#09090b] hover:border-zinc-700'
                  }`}
                >
                  <div className="truncate pr-1">
                    <span className="text-xs font-bold line-clamp-1">{sess.name}</span>
                    <span className="text-[10px] font-mono text-zinc-500 flex items-center gap-1.5">
                      Code: {sess.id} • {sess.questions?.length || 0} qs
                    </span>
                  </div>
                  <ChevronRight className={`w-3.5 h-3.5 transition-transform ${isActiveLocal ? 'translate-x-0.5 text-[#a3e635]' : 'text-zinc-600'}`} />
                </div>
              );
            })}
          </div>
        </div>
      </section>

      {/* RIGHT SIDE MAIN BODY: ACTIVE TAB PANEL */}
      <section className="flex-1 space-y-4">
        {/* Subnavigation tab headers */}
        <div className="flex items-center gap-1 bg-slate-900/50 p-1.5 rounded-xl border border-slate-900 select-none">
          <button
            onClick={() => setActiveTab('sessions')}
            className={`flex-1 py-1.5 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer transition-all ${
              activeTab === 'sessions' 
                ? 'bg-slate-800 text-slate-100 shadow-md' 
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <BarChart className="w-3.5 h-3.5" />
            <span>Resultados e Monitor</span>
          </button>
          
          <button
            onClick={() => {
              if (!selectedSession) {
                showToast('Selecione ou crie uma sessão primeiro!');
                return;
              }
              setActiveTab('questions');
            }}
            className={`flex-1 py-1.5 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer transition-all ${
              !selectedSession ? 'opacity-40 cursor-not-allowed' : ''
            } ${
              activeTab === 'questions' 
                ? 'bg-slate-800 text-slate-100 shadow-md' 
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <FileText className="w-3.5 h-3.5" />
            <span>Perguntas da Sessão ({selectedSession?.questions?.length || 0})</span>
          </button>

          <button
            onClick={() => setActiveTab('bank')}
            className={`flex-1 py-1.5 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 cursor-pointer transition-all ${
              activeTab === 'bank' 
                ? 'bg-slate-800 text-slate-100 shadow-md' 
                : 'text-slate-400 hover:text-slate-200'
            }`}
          >
            <Database className="w-3.5 h-3.5" />
            <span>Banco de Itens ({questionBank.length})</span>
          </button>
        </div>

        {/* TAB 1: SESSIONS - TOTAL OVERVIEW AND REPORTS */}
        {activeTab === 'sessions' && (
          <div className="space-y-4">
            {!selectedSession ? (
              <div className="text-center py-20 bg-slate-900/20 border border-dashed border-slate-900 rounded-2xl flex flex-col items-center justify-center p-4">
                <Users className="w-10 h-10 text-slate-700 opacity-60 mb-2" />
                <h3 className="font-bold text-slate-350">Selecione ou crie uma sessão ativa</h3>
                <p className="text-xs text-slate-500 max-w-sm mt-1">
                  Use o painel lateral para escolher uma aula ou preencher uma nova palestra de instruções.
                </p>
              </div>
            ) : (
              <div className="p-5 bg-gradient-to-b from-slate-900/60 to-slate-950 border border-slate-900 rounded-2xl space-y-5">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-slate-850 gap-4">
                  <div>
                    <span className="inline-flex items-center gap-1 bg-indigo-900/40 text-indigo-400 text-[10px] font-bold px-2 py-0.5 rounded border border-indigo-800/45 mb-1 text-uppercase">
                      Ativa
                    </span>
                    <h2 className="text-lg font-black text-slate-100 tracking-tight">{selectedSession.name}</h2>
                    <p className="text-xs text-slate-400 mt-1 flex items-center gap-1.5">
                      <span>Criada em: {new Date(selectedSession.createdAt).toLocaleDateString()}</span>
                      <span>•</span>
                      <span>Modo: {selectedSession.isQuizMode ? 'Competitivo (Quiz)' : 'Pesquisa Livre'}</span>
                    </p>
                  </div>

                  {/* Actions Bar */}
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => handleDownloadCSV(selectedSession.id)}
                      className="p-2 bg-slate-800 hover:bg-slate-750 text-slate-200 hover:text-slate-100 rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer transition-colors border border-slate-700"
                    >
                      <Download className="w-3.5 h-3.5 text-slate-400" />
                      <span>CSV (Excel)</span>
                    </button>
                    
                    <button
                      onClick={() => handleDownloadJSON(selectedSession.id)}
                      className="p-2 bg-slate-800 hover:bg-slate-750 text-slate-200 hover:text-slate-100 rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer transition-colors border border-slate-700"
                    >
                      <Database className="w-3.5 h-3.5 text-slate-400" />
                      <span>JSON</span>
                    </button>

                    <button
                      onClick={() => {
                        onSelectRole('tv');
                        onRefresh();
                      }}
                      className="p-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1 cursor-pointer transition-colors shadow"
                    >
                      <Play className="w-3.5 h-3.5" />
                      <span>Projetar Slide</span>
                    </button>
                  </div>
                </div>

                {/* Statistics Highlights Widgets */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                  <div className="p-3 bg-slate-950 border border-slate-900 rounded-xl">
                    <span className="block text-[10px] text-slate-500 uppercase font-mono tracking-wider font-semibold">Código Curto</span>
                    <span className="text-lg font-black text-orange-400 font-mono tracking-wide">{selectedSession.id}</span>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-900 rounded-xl">
                    <span className="block text-[10px] text-slate-500 uppercase font-mono tracking-wider font-semibold">Perguntas</span>
                    <span className="text-lg font-black text-slate-100">{selectedSession.questions.length}</span>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-900 rounded-xl">
                    <span className="block text-[10px] text-slate-500 uppercase font-mono tracking-wider font-semibold">Votantes Ativos</span>
                    <span className="text-lg font-black text-emerald-400 font-mono">
                      {syncStore.getParticipantsForSession(selectedSession.id).length}
                    </span>
                  </div>
                  <div className="p-3 bg-slate-950 border border-slate-900 rounded-xl">
                    <span className="block text-[10px] text-slate-500 uppercase font-mono tracking-wider font-semibold">Votos Registrados</span>
                    <span className="text-lg font-black text-sky-400 font-mono">
                      {syncStore.getVotesForSession(selectedSession.id).length}
                    </span>
                  </div>
                </div>

                {/* QR Code Anchor / Copy URLs info */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 bg-slate-950 p-4 rounded-xl border border-slate-900">
                  <div className="md:col-span-1 flex justify-center">
                    <QRCodeDisplay value={`${appUrl}?role=participant&session=${selectedSession.id}`} size={140} />
                  </div>
                  <div className="md:col-span-2 space-y-2.5 flex flex-col justify-center">
                    <h3 className="font-bold text-slate-350 text-sm flex items-center gap-1.5">
                      <Share2 className="w-4 h-4 text-orange-400" />
                      Acesso dos Alunos
                    </h3>
                    <p className="text-xs text-slate-400 leading-relaxed">
                      Peça para os alunos apontarem o celular para este QR Code, ou envie/escreva o link abaixo para votação no navegador.
                    </p>
                    <div className="bg-slate-900 p-2.5 rounded-lg border border-slate-800 text-xs font-mono select-all overflow-x-auto truncate">
                      {appUrl}?role=participant&session={selectedSession.id}
                    </div>

                    <div className="flex items-center gap-2.5 pt-1">
                      <button
                        onClick={handleAddSimulatedParticipants}
                        className="text-xs text-slate-300 hover:text-slate-100 bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded px-2.5 py-1.5 font-medium flex items-center gap-1 transition-all"
                        id="btn-simulate-participants"
                      >
                        <Users className="w-3.5 h-3.5 text-slate-400" />
                        <span>Mock +5 Alunos Conectando</span>
                      </button>

                      {showDeleteConfirm === selectedSession.id ? (
                        <div className="flex items-center gap-1.5">
                          <span className="text-[10px] text-rose-400 font-bold">Apagar TUDO?</span>
                          <button
                            onClick={() => handleDeleteSession(selectedSession.id)}
                            className="bg-rose-955 text-rose-300 border border-rose-800 text-[10px] px-2 py-1 font-bold rounded hover:bg-rose-900"
                            id="btn-confirm-delete"
                          >
                            Sim
                          </button>
                          <button
                            onClick={() => setShowDeleteConfirm(null)}
                            className="bg-slate-800 text-[10px] px-2 py-1 rounded text-slate-300 hover:bg-slate-755"
                          >
                            Não
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => setShowDeleteConfirm(selectedSession.id)}
                          className="text-xs text-rose-400 hover:text-rose-300 bg-rose-950/20 border border-rose-950 rounded px-2.5 py-1.5 font-medium flex items-center gap-1 transition-all"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          <span>Excluir Sessão</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Participant Roster Table */}
                <div className="space-y-2">
                  <h3 className="font-bold text-slate-300 text-sm">Alunos Conectados nesta Sessão</h3>
                  
                  <div className="bg-slate-950 rounded-xl border border-slate-900 overflow-hidden">
                    <div className="grid grid-cols-3 p-2.5 bg-slate-900/50 border-b border-slate-900 text-[10px] font-mono text-slate-400 font-bold">
                      <span>Nome do Aluno</span>
                      <span className="text-center">Modo</span>
                      <span className="text-right">Acúmulo de Pontos</span>
                    </div>

                    <div className="divide-y divide-slate-900/60 max-h-48 overflow-y-auto">
                      {syncStore.getParticipantsForSession(selectedSession.id).length === 0 ? (
                        <div className="text-center p-4 text-slate-500 text-xs">
                          Nenhum aluno registrado na sessão de votação.
                        </div>
                      ) : (
                        syncStore.getParticipantsForSession(selectedSession.id).map((p, idx) => {
                          const pVotes = syncStore.getVotesForSession(selectedSession.id).filter(v => v.participantId === p.id).length;
                          return (
                            <div key={idx} className="grid grid-cols-3 px-3 py-2 text-xs">
                              <span className="font-medium text-slate-300 truncate">{p.name}</span>
                              <span className="text-center text-[10px] font-mono text-slate-500">
                                {pVotes} {pVotes === 1 ? 'voto lançado' : 'votos lançados'}
                              </span>
                              <span className="text-right font-mono font-bold text-orange-400">
                                {selectedSession.isQuizMode ? `${p.score || 0} XP` : 'Survey'}
                              </span>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 2: QUESTIONS - LIST AND EDITOR FORM */}
        {activeTab === 'questions' && selectedSession && (
          <div className="space-y-4">
            {/* Conditional form rendering */}
            {isEditingQuestion ? (
              <div className="p-5 bg-gradient-to-b from-slate-900/80 to-slate-950 border border-orange-500/10 rounded-2xl space-y-4 shadow-xl">
                <div className="flex items-center justify-between pb-3 border-b border-slate-850">
                  <h3 className="font-bold text-slate-200 text-sm">
                    {editingQuestionId ? '✏️ Editar Pergunta Múltipla Escolha' : '➕ Adicionar Pergunta Múltipla Escolha'}
                  </h3>
                  <button
                    onClick={handleCloseQuestionForm}
                    className="text-xs text-slate-400 hover:text-slate-200 cursor-pointer"
                  >
                    Voltar para lista
                  </button>
                </div>

                <form onSubmit={handleSaveQuestion} className="space-y-4">
                  <div>
                    <label className="block text-xs font-mono text-slate-400 mb-1">Enunciado da Pergunta</label>
                    <textarea
                      value={questionText}
                      onChange={(e) => setQuestionText(e.target.value)}
                      placeholder="Qual a pergunta de foco desta apresentação?"
                      rows={3}
                      className="w-full text-xs font-medium bg-slate-950 border border-slate-800 rounded-lg p-2.5 text-slate-100 placeholder-slate-600 focus:border-orange-500 focus:outline-none focus:ring-1 focus:ring-orange-500/20"
                    />
                  </div>

                  <div className="space-y-2.5">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-mono text-slate-400 block">Alternativas e Gabarito</label>
                      <button
                        type="button"
                        onClick={handleAddOptionField}
                        className="text-[10px] text-orange-400 hover:text-orange-300 font-bold bg-slate-900 px-2.5 py-1 rounded border border-slate-800"
                        id="btn-add-option-field"
                      >
                        + Alternativa
                      </button>
                    </div>

                    <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                      {options.map((opt, oIdx) => {
                        const letter = String.fromCharCode(65 + oIdx);
                        const isCorrectSelection = correctOptionIdx === oIdx;

                        return (
                          <div key={oIdx} className="flex items-center gap-2">
                            {/* Quiz Marker Toggle */}
                            {selectedSession.isQuizMode && (
                              <button
                                type="button"
                                onClick={() => setCorrectOptionIdx(isCorrectSelection ? null : oIdx)}
                                className={`w-6 h-6 rounded font-mono text-[10px] font-bold flex items-center justify-center shrink-0 border select-none transition-colors ${
                                  isCorrectSelection
                                    ? 'bg-emerald-600 border-emerald-500 text-white'
                                    : 'bg-slate-950 border-slate-850 text-slate-500 hover:border-slate-700'
                                }`}
                                title="Marcar como alternativa certa"
                              >
                                {letter}
                              </button>
                            )}

                            <span className="font-mono text-xs text-slate-500">
                              {!selectedSession.isQuizMode ? `${letter})` : ''}
                            </span>

                            <input
                              type="text"
                              value={opt}
                              onChange={(e) => handleOptionTextChange(e.target.value, oIdx)}
                              placeholder={`Alternativa ${letter}...`}
                              className="flex-1 text-xs bg-slate-950 border border-slate-800 rounded-lg p-2 focus:border-slate-650 focus:outline-none text-slate-250 placeholder-slate-650"
                            />

                            <button
                              type="button"
                              onClick={() => handleRemoveOptionField(oIdx)}
                              className="p-2 text-slate-500 hover:text-rose-400 transition-colors"
                              title="Remover alternativa"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {selectedSession.isQuizMode && (
                    <div className="p-3 bg-slate-950 border border-slate-900 rounded-lg text-[11px] text-slate-400 flex items-start gap-1.5 leading-relaxed">
                      <HelpCircle className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
                      <div>
                        Clique na letra à esquerda (<span className="text-emerald-400 font-bold">A, B, C...</span>) para definir qual a alternativa correta. Deixe sem seleção se não for aplicar notas nesta questão.
                      </div>
                    </div>
                  )}

                  <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-850">
                    <button
                      type="button"
                      onClick={handleCloseQuestionForm}
                      className="py-2 px-3 bg-slate-900 hover:bg-slate-800 rounded-lg text-xs font-semibold text-slate-400"
                    >
                      Cancelar
                    </button>
                    <button
                      type="submit"
                      className="py-2 px-4 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-xs font-semibold shadow hover:translate-y-[-1px] transition-transform cursor-pointer"
                      id="btn-save-question"
                    >
                      Salvar Pergunta
                    </button>
                  </div>
                </form>
              </div>
            ) : (
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold text-slate-350 text-sm flex items-center gap-1.5">
                    <FileText className="w-4.5 h-4.5 text-orange-400" />
                    Lista de Perguntas
                  </h3>
                  <button
                    onClick={handleOpenNewQuestionForm}
                    className="bg-orange-600 hover:bg-orange-700 text-white text-xs font-semibold py-1.5 px-3 rounded-lg flex items-center gap-1 transition-all cursor-pointer"
                    id="btn-add-question-to-session"
                  >
                    <Plus className="w-4 h-4" />
                    Adicionar Pergunta
                  </button>
                </div>

                {selectedSession.questions.length === 0 ? (
                  <div className="p-12 text-center rounded-2xl bg-slate-900/20 border border-slate-900 text-slate-450 space-y-3">
                    <Layers className="w-8 h-8 opacity-45 mx-auto" />
                    <p className="text-xs">Nenhuma pergunta nesta sessão.</p>
                    <p className="text-[11px] text-slate-500">Comece a adicionar perguntas acima, ou reutilize itens do Banco Geral.</p>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {selectedSession.questions.map((q, idx) => {
                      const votesCount = syncStore.getVotesForQuestion(selectedSession.id, q.id).length;
                      return (
                        <div
                          key={q.id}
                          className="p-4 rounded-xl bg-slate-900/60 border border-slate-900/80 hover:border-slate-800 flex flex-col md:flex-row md:items-start justify-between gap-4 transition-all"
                        >
                          <div className="space-y-2 flex-1">
                            <div className="flex items-center gap-2">
                              <span className="font-mono text-xs font-black bg-slate-800 border border-slate-700 text-orange-400 px-2 py-0.5 rounded">
                                Pergunta {idx + 1}
                              </span>
                              <span className="text-[10px] font-mono text-slate-500">
                                {votesCount} {votesCount === 1 ? 'voto' : 'votos'} acumulados
                              </span>
                            </div>

                            <p className="font-semibold text-slate-100 text-sm leading-relaxed">{q.text}</p>
                            
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 pt-1">
                              {q.options.map((opt, oIdx) => {
                                const letter = String.fromCharCode(65 + oIdx);
                                const isCorrect = oIdx === q.correctOptionIndex;
                                return (
                                  <div 
                                    key={oIdx} 
                                    className={`p-1.5 px-2 rounded font-medium text-xs leading-none truncate flex items-center gap-1.5 ${
                                      isCorrect && selectedSession.isQuizMode
                                        ? 'bg-emerald-950/20 border border-emerald-900/35 text-emerald-300 text-bold' 
                                        : 'bg-slate-950/50 text-slate-400 text-[11px]'
                                    }`}
                                  >
                                    <span className="font-mono text-[10px] text-slate-500">{letter})</span>
                                    <span className="truncate">{opt}</span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>

                          {/* Controls column */}
                          <div className="flex flex-row md:flex-col items-center md:items-end justify-between md:justify-start gap-2 border-t md:border-t-0 md:border-l border-slate-900 pt-3 md:pt-0 md:pl-3">
                            <div className="flex items-center gap-1">
                              <button
                                onClick={() => handleOpenEditQuestionForm(q)}
                                className="p-1.5 hover:bg-slate-800 text-slate-400 hover:text-slate-200 rounded transition-colors"
                                title="Editar pergunta"
                              >
                                <Edit2 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => handleDuplicateQuestion(q)}
                                className="p-1.5 hover:bg-slate-800 text-slate-400 hover:text-slate-200 rounded transition-colors"
                                title="Duplicar item"
                              >
                                <Copy className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => handleSaveToBank(q)}
                                className="p-1.5 hover:bg-slate-800 text-slate-400 hover:text-slate-200 rounded transition-colors"
                                title="Salvar no Banco"
                              >
                                <Database className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => handleDeleteQuestion(q.id)}
                                className="p-1.5 hover:bg-slate-800 text-slate-500 hover:text-rose-400 rounded transition-colors"
                                title="Excluir pergunta"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB 3: QUESTION BANK */}
        {activeTab === 'bank' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between pb-1.5 border-b border-slate-900">
              <div>
                <h3 className="font-bold text-slate-350 text-sm flex items-center gap-1.5">
                  <Database className="w-4.5 h-4.5 text-orange-400 animate-pulse" />
                  Banco Geral de Itens Reutilizáveis
                </h3>
                <p className="text-[10px] text-slate-500 mt-0.5">Perguntas guardadas ou modelos criados para puxar em qualquer aula posterior.</p>
              </div>

              {selectedSession && (
                <div className="text-xs font-semibold text-slate-400 font-mono">
                  Sessão Destino: <span className="text-orange-400">{selectedSession.id}</span>
                </div>
              )}
            </div>

            {questionBank.length === 0 ? (
              <div className="text-center py-16 rounded-xl bg-slate-900/35 border border-slate-900 p-4 text-slate-500 space-y-2">
                <Database className="w-6 h-6 opacity-30 mx-auto" />
                <p className="text-xs">Nenhum item salvo no Banco!</p>
                <p className="text-[11px] text-slate-600 leading-relaxed">
                  Para guardar perguntas aqui, clique no ícone de banco de dados (<span className="text-indigo-400">Database</span>) ao lado de qualquer pergunta na lista de sessões.
                </p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {questionBank.map((item) => (
                  <div
                    key={item.id}
                    className="p-3.5 rounded-xl bg-slate-900/60 border border-slate-900 hover:border-slate-800 transition-all flex flex-col justify-between space-y-3"
                  >
                    <div className="space-y-1.5">
                      <div className="flex items-center justify-between text-[10px] font-mono text-slate-500">
                        <span className="bg-slate-800 text-slate-300 font-bold px-1.5 py-0.5 rounded text-[8px] uppercase">
                          {item.category || 'Geral'}
                        </span>
                        <span>{new Date(item.createdAt).toLocaleDateString()}</span>
                      </div>
                      
                      <p className="text-slate-100 font-bold text-xs leading-relaxed line-clamp-3">{item.text}</p>

                      <div className="space-y-1 text-[10px] text-slate-450 border-t border-slate-900 pt-2 font-medium">
                        {item.options.slice(0, 3).map((opt, oIdx) => (
                          <div key={oIdx} className="truncate">
                            {String.fromCharCode(65 + oIdx)}) {opt}
                          </div>
                        ))}
                        {item.options.length > 3 && (
                          <div className="text-[9px] text-slate-600">e mais {item.options.length - 3} alternativas...</div>
                        )}
                      </div>
                    </div>

                    <div className="flex items-center justify-between pt-2 border-t border-slate-900/80">
                      <button
                        onClick={() => handleDeleteFromBank(item.id)}
                        className="text-[10px] text-slate-500 hover:text-rose-400 font-medium"
                      >
                        Excluir
                      </button>

                      <button
                        onClick={() => handleImportFromBank(item)}
                        disabled={!selectedSession}
                        className={`py-1 px-2.5 rounded-lg text-[10px] font-bold flex items-center gap-1 transition-colors ${
                          selectedSession 
                            ? 'bg-slate-800 text-slate-200 hover:bg-slate-700' 
                            : 'bg-slate-900/40 text-slate-600 cursor-not-allowed'
                        }`}
                      >
                        <span>Puxar para Sessão</span>
                        <ArrowRight className="w-3 h-3 text-slate-400" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </section>
    </div>
  );
};
