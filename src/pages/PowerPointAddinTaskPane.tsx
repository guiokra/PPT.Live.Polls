/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { syncStore } from '../services/store';
import { Session, Question } from '../types';
import { QRCodeDisplay } from '../components/QRCodeDisplay';
import { 
  FileText, 
  Layers, 
  QrCode, 
  ExternalLink, 
  BarChart2, 
  Eye, 
  EyeOff, 
  Plus, 
  Info,
  Copy,
  Check,
  ChevronLeft,
  ChevronRight,
  Database,
  UserCheck,
  HelpCircle,
  Clock
} from 'lucide-react';
import QRCode from 'qrcode';

// Core OfficeJS helper
declare const Office: any;

interface PowerPointAddinTaskPaneProps {
  onSelectSession: (id: string) => void;
  onOpenCreateSession: () => void;
  activeSession: Session | null;
  onRefresh: () => void;
  appUrl: string;
}

export const PowerPointAddinTaskPane: React.FC<PowerPointAddinTaskPaneProps> = ({
  onSelectSession,
  onOpenCreateSession,
  activeSession,
  onRefresh,
  appUrl
}) => {
  const [sessions, setSessions] = useState<Session[]>([]);
  const [officeStatus, setOfficeStatus] = useState<string>('Buscando ambiente Office...');
  const [isOfficeAvailable, setIsOfficeAvailable] = useState<boolean>(false);
  const [insertSuccess, setInsertSuccess] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);

  // Load active sessions
  useEffect(() => {
    setSessions(syncStore.getSessions().filter(s => s.status === 'active'));
    
    // Detect Microsoft OfficeJS Environment
    if (typeof Office !== 'undefined') {
      Office.onReady((info: any) => {
        if (info.host === Office.HostType.PowerPoint) {
          setIsOfficeAvailable(true);
          setOfficeStatus('Conectado ao PowerPoint');
        } else {
          setOfficeStatus('Office carregado, mas host não é PowerPoint');
        }
      });
    } else {
      setOfficeStatus('Navegador Web (Modo Simulado)');
    }
  }, [activeSession]);

  const handleCopyText = async (text: string, id: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(id);
      setTimeout(() => setCopiedId(null), 2000);
    } catch (_) {}
  };

  const getJoinUrl = (sessionId: string) => {
    return `${appUrl}?role=participant&session=${sessionId}`;
  };

  const getTvUrl = (sessionId: string) => {
    return `${appUrl}?role=tv&session=${sessionId}`;
  };

  // PowerPoint API slide inserting
  const insertTextToSlide = (question: Question, sessionCode: string) => {
    if (!question) return;

    const optLines = question.options.map((opt, idx) => `[ ${String.fromCharCode(65 + idx)} ]  ${opt}`).join('\n');
    const slideText = `${question.text}\n\n${optLines}\n\n👉 Para responder, acesse: ${appUrl}?role=participant&session=${sessionCode}\n📌 Código: ${sessionCode}`;

    if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
      Office.context.document.setSelectedDataAsync(
        slideText,
        { coercionType: Office.CoercionType.Text },
        (asyncResult: any) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            triggerNotification('Erro ao inserir texto no slide: ' + asyncResult.error.message);
          } else {
            triggerNotification('Texto da pergunta inserido com sucesso!');
          }
        }
      );
    } else {
      // Browser Fallback: Copy to clipboard with pretty design
      handleCopyText(slideText, 'slide-content');
      triggerNotification('Texto copiado! Cole (Ctrl+V) no seu slide.');
    }
  };

  // PowerPoint API QR code inserting as an image
  const insertQRCodeToSlide = async (sessionCode: string) => {
    const url = getJoinUrl(sessionCode);
    
    try {
      // Generate standard high res base64 PNG data url
      const qrDataUrl = await QRCode.toDataURL(url, {
        width: 320,
        margin: 1,
        color: {
          dark: '#0f172a', // deep slate
          light: '#ffffff' // white background
        }
      });

      // Strip "data:image/png;base64," header for OfficeJS API requirements
      const base64Str = qrDataUrl.split(',')[1];

      if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
        Office.context.document.setSelectedDataAsync(
          base64Str,
          { coercionType: Office.CoercionType.Image },
          (asyncResult: any) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              triggerNotification('Falha ao inserir QR: ' + asyncResult.error.message);
            } else {
              triggerNotification('QR Code de votação adicionado ao slide!');
            }
          }
        );
      } else {
        // Fallback: Copy Image DataURL or open download link
        const link = document.createElement('a');
        link.href = qrDataUrl;
        link.download = `qrcode-sessao-${sessionCode}.png`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        triggerNotification('Download do QR Code iniciado! Cole a imagem no slide.');
      }
    } catch (e: any) {
      console.error(e);
      triggerNotification('Erro ao processar QR: ' + e.message);
    }
  };

  const triggerNotification = (message: string) => {
    setInsertSuccess(message);
    setTimeout(() => setInsertSuccess(null), 3000);
  };

  // Control slides & responses
  const handlePrevQuestion = () => {
    if (!activeSession) return;
    const newIdx = Math.max(0, activeSession.currentQuestionIndex - 1);
    syncStore.setQuestionIndex(activeSession.id, newIdx);
    onRefresh();
  };

  const handleNextQuestion = () => {
    if (!activeSession) return;
    const newIdx = Math.min(activeSession.questions.length - 1, activeSession.currentQuestionIndex + 1);
    syncStore.setQuestionIndex(activeSession.id, newIdx);
    onRefresh();
  };

  const toggleResultsVisibility = () => {
    if (!activeSession) return;
    syncStore.toggleResults(activeSession.id, !activeSession.showResults);
    onRefresh();
  };

  const toggleRevealCorrect = () => {
    if (!activeSession) return;
    syncStore.toggleRevealAnswer(activeSession.id, !activeSession.revealAnswer);
    onRefresh();
  };

  const simulateIncomingVotes = () => {
    if (!activeSession) return;
    const currentQ = activeSession.questions[activeSession.currentQuestionIndex];
    if (!currentQ) return;
    syncStore.simulateVotes(activeSession.id, currentQ.id, 5); // Add 5 random responses
    onRefresh();
    triggerNotification('+5 votos simulados adicionados!');
  };

  return (
    <div id="ppt-addin-panel" className="bg-[#09090b] text-zinc-100 min-h-screen flex flex-col relative text-sm select-none border-r border-[#27272a] w-full max-w-md mx-auto">
      {/* Header */}
      <header className="p-4 border-b border-[#27272a] bg-[#121214]/65 sticky top-0 z-25 backdrop-blur-md">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-2.5 h-2.5 rounded-full bg-[#a3e635] animate-pulse" />
            <span className="font-bold tracking-tight text-white font-display">PowerPoint Add-In</span>
          </div>
          <div className="text-[10px] uppercase tracking-wider font-mono px-2 py-0.5 rounded bg-zinc-800 text-zinc-400 border border-zinc-700/60">
            {isOfficeAvailable ? 'PowerPoint' : 'Simulador'}
          </div>
        </div>
        <p className="text-[11px] text-zinc-400 mt-1 flex items-center gap-1">
          <Database className="w-3 h-3 text-zinc-500" />
          {officeStatus}
        </p>
      </header>

      {/* Floating alert */}
      {insertSuccess && (
        <div className="absolute top-16 left-4 right-4 z-40 bg-zinc-900/95 text-white border border-[#a3e635]/35 rounded-lg p-3 shadow-xl flex items-start gap-2.5 animate-in fade-in slide-in-from-top-2 duration-300">
          <Check className="w-4 h-4 text-[#a3e635] shrink-0 mt-0.5" />
          <p className="text-xs font-semibold">{insertSuccess}</p>
        </div>
      )}

      {/* Main Container */}
      <main className="flex-1 overflow-y-auto p-4 space-y-4">
        {!activeSession ? (
          /* SECTION: SESSION LIST SELECTOR */
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="font-bold text-zinc-350 font-display">Minhas Sessões</h3>
              <button
                onClick={onOpenCreateSession}
                className="text-xs bg-[#a3e635] hover:bg-[#a3e635]/90 text-zinc-950 font-black px-2.5 py-1.5 rounded-lg flex items-center gap-1 cursor-pointer transition-colors"
                id="btn-new-session"
              >
                <Plus className="w-3.5 h-3.5 text-zinc-950" />
                Criar
              </button>
            </div>

            {sessions.length === 0 ? (
              <div className="text-center py-8 rounded-xl bg-[#121214]/40 border border-[#27272a] text-zinc-400 p-4">
                <Info className="w-6 h-6 mx-auto opacity-30 mb-2 text-zinc-500" />
                <p className="text-xs">Nenhuma sessão ativa encontrada.</p>
                <p className="text-[11px] text-zinc-500 mt-1">Crie uma nova sessão para começar a adicionar enquetes!</p>
              </div>
            ) : (
              <div className="space-y-2.5">
                {sessions.map((sess) => {
                  const hasQuestions = sess.questions.length > 0;
                  return (
                    <div
                      key={sess.id}
                      className="p-3.5 rounded-xl bg-[#121214] border border-[#27272a] transition-all hover:border-zinc-700 flex flex-col justify-between"
                    >
                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <span className="font-bold text-white line-clamp-1">{sess.name}</span>
                          <span className="font-mono text-[10px] font-bold bg-[#a3e635]/10 text-[#a3e635] px-1.5 py-0.5 rounded border border-[#a3e635]/20">
                            ID: {sess.id}
                          </span>
                        </div>
                        <div className="flex items-center gap-2.5 text-zinc-400 text-[11px] font-mono mt-1 w-full">
                          <span className="flex items-center gap-1">
                            <Layers className="w-3 h-3 text-zinc-600" />
                            {sess.questions.length} {sess.questions.length === 1 ? 'pergunta' : 'perguntas'}
                          </span>
                          <span>•</span>
                          <span className="flex items-center gap-1">
                            <UserCheck className="w-3 h-3 text-zinc-600" />
                            {sess.isQuizMode ? 'Modo Quiz' : 'Enquete'}
                          </span>
                        </div>
                      </div>
                      
                      <button
                        onClick={() => onSelectSession(sess.id)}
                        disabled={!hasQuestions}
                        className={`mt-3 w-full py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 cursor-pointer transition-all ${
                          hasQuestions 
                            ? 'bg-zinc-800 hover:bg-zinc-700 text-white border border-zinc-750' 
                            : 'bg-zinc-900/40 text-zinc-650 border border-dashed border-zinc-800 cursor-not-allowed'
                        }`}
                      >
                        {hasQuestions ? (
                          <>
                            <span>Abrir no PowerPoint</span>
                            <ChevronRight className="w-3.5 h-3.5 text-zinc-400" />
                          </>
                        ) : (
                          'Adicione perguntas para abrir'
                        )}
                      </button>
                    </div>
                  );
                })}
              </div>
            )}

            <div className="p-3 rounded-xl bg-[#a3e635]/5 border border-[#a3e635]/15 text-[#a3e635]/90 text-xs flex items-start gap-2">
              <HelpCircle className="w-4 h-4 shrink-0 text-[#a3e635] mt-0.5" />
              <div>
                <p className="font-semibold font-display">Como funciona?</p>
                <p className="text-zinc-400 text-[11px] mt-0.5 leading-relaxed">
                  Crie perguntas, use os botões para injetar enquetes nos slides e abra a Tela do Projetor. Seus alunos respondem escaneando o QR Code!
                </p>
              </div>
            </div>
          </div>
        ) : (
          /* SECTION: CONTROL FLUX PANEL FOR ACTIVE SESSION */
          <div className="space-y-4">
            {/* Session Breadcrumb Info */}
            <div className="flex items-center justify-between p-2 rounded-lg bg-[#121214] border border-[#27272a]">
              <div className="min-w-0 flex-1">
                <p className="text-[10px] uppercase font-mono text-zinc-550 font-bold">Sessão Ativa</p>
                <h4 className="font-bold text-white truncate font-display">{activeSession.name}</h4>
              </div>
              <button
                onClick={() => onSelectSession('')}
                className="text-[10px] font-semibold text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-750 px-2 py-1 rounded cursor-pointer transition-colors"
              >
                Mudar
              </button>
            </div>

            {activeSession.questions.length === 0 ? (
              <div className="text-center py-8">
                <p className="text-zinc-500 text-xs">Esta sessão não possui perguntas cadastrados.</p>
              </div>
            ) : (
              <>
                {/* Active Question Carousel Card */}
                <div className="p-4 rounded-xl bg-[#121214] border border-[#27272a] shadow-md">
                  <div className="flex items-center justify-between pb-2 mb-2 border-b border-zinc-850 text-[11px] font-mono text-zinc-400">
                    <span className="bg-zinc-855 px-2 py-0.5 rounded text-[#a3e635] font-black">
                      Slide {activeSession.currentQuestionIndex + 1} de {activeSession.questions.length}
                    </span>
                    <span className="flex items-center gap-0.5 text-zinc-500">
                      <Clock className="w-3 h-3 text-zinc-650" />
                      {activeSession.isQuizMode ? 'Modo Quiz (c/ gabarito)' : 'Enquete'}
                    </span>
                  </div>

                  <p className="font-bold text-white text-sm leading-snug">
                    {activeSession.questions[activeSession.currentQuestionIndex]?.text}
                  </p>

                  <div className="mt-3 space-y-1.5">
                    {activeSession.questions[activeSession.currentQuestionIndex]?.options.map((opt, oIdx) => {
                      const isCorrect = oIdx === activeSession.questions[activeSession.currentQuestionIndex].correctOptionIndex;
                      return (
                        <div 
                          key={oIdx} 
                          className={`p-2 rounded h-8 flex items-center justify-between text-xs transition-colors ${
                            isCorrect && activeSession.isQuizMode
                              ? 'bg-emerald-950/35 border border-emerald-800/40 text-emerald-300 font-semibold' 
                              : 'bg-[#09090b] border border-zinc-850/70 text-zinc-400'
                          }`}
                        >
                          <span className="truncate">
                            <strong className="font-mono mr-1.5 text-zinc-500">{String.fromCharCode(65 + oIdx)})</strong>
                            {opt}
                          </span>
                          {isCorrect && activeSession.isQuizMode && (
                            <span className="text-[9px] uppercase font-mono font-bold tracking-wider px-1.5 bg-emerald-500/20 rounded">Gabarito</span>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  {/* Carousel Controllers */}
                  <div className="flex items-center justify-between gap-2.5 mt-4 pt-4 border-t border-zinc-850">
                    <button
                      onClick={handlePrevQuestion}
                      disabled={activeSession.currentQuestionIndex === 0}
                      className="flex-1 py-1.5 px-2.5 rounded bg-zinc-800 hover:bg-zinc-700 disabled:opacity-40 disabled:hover:bg-zinc-800 text-xs font-semibold text-white flex items-center justify-center gap-1 cursor-pointer transition-colors"
                    >
                      <ChevronLeft className="w-4 h-4 text-zinc-400" />
                      Anterior
                    </button>
                    <button
                      onClick={handleNextQuestion}
                      disabled={activeSession.currentQuestionIndex === activeSession.questions.length - 1}
                      className="flex-1 py-1.5 px-2.5 rounded bg-zinc-800 hover:bg-zinc-700 disabled:opacity-40 disabled:hover:bg-zinc-800 text-xs font-semibold text-white flex items-center justify-center gap-1 cursor-pointer transition-colors"
                    >
                      Próxima
                      <ChevronRight className="w-4 h-4 text-zinc-400" />
                    </button>
                  </div>
                </div>

                {/* INJECTION POWERPOINT BUTTONS */}
                <div className="space-y-2">
                  <h5 className="font-bold text-[10px] uppercase tracking-widest text-zinc-500 font-mono">Comandos de Slide (PowerPoint)</h5>
                  
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      onClick={() => insertTextToSlide(activeSession.questions[activeSession.currentQuestionIndex], activeSession.id)}
                      className="p-3 rounded-xl bg-[#a3e635] hover:bg-[#a3e635]/90 text-zinc-950 font-bold flex flex-col items-center justify-center gap-2 cursor-pointer shadow transition-all active:scale-95 duration-200"
                    >
                      <FileText className="w-5 h-5 text-zinc-950" />
                      <span className="text-xs text-center font-bold leading-tight uppercase font-mono">Texto do Slide</span>
                    </button>
                    
                    <button
                      onClick={() => insertQRCodeToSlide(activeSession.id)}
                      className="p-3 rounded-xl bg-zinc-900 border border-zinc-800 hover:border-zinc-700 text-zinc-200 font-bold flex flex-col items-center justify-center gap-2 cursor-pointer shadow transition-all active:scale-95 duration-200"
                    >
                      <QrCode className="w-5 h-5 text-[#a3e635]" />
                      <span className="text-xs text-center font-bold leading-tight uppercase font-mono">Inserir QR Code</span>
                    </button>
                  </div>
                </div>

                {/* INTERACTIVE CONTROLS */}
                <div className="space-y-2 pt-1">
                  <h5 className="font-bold text-[10px] uppercase tracking-widest text-zinc-500 font-mono">Controles do Apresentador</h5>
                  
                  <div className="space-y-2">
                    <button
                      onClick={toggleResultsVisibility}
                      className={`w-full py-2 px-3 rounded-xl font-bold flex items-center justify-center gap-2.5 transition-all cursor-pointer ${
                        activeSession.showResults 
                          ? 'bg-indigo-650 hover:bg-zinc-755 text-white shadow-lg' 
                          : 'bg-zinc-800 hover:bg-zinc-750 text-zinc-300'
                      }`}
                    >
                      {activeSession.showResults ? (
                        <>
                          <EyeOff className="w-4 h-4 text-zinc-200" />
                          <span>Ocultar Gráfico no Slide</span>
                        </>
                      ) : (
                        <>
                          <Eye className="w-4 h-4 text-[#a3e635] animate-pulse" />
                          <span>Revelar Resultados</span>
                        </>
                      )}
                    </button>

                    {activeSession.isQuizMode && (
                      <button
                        onClick={toggleRevealCorrect}
                        className={`w-full py-2 px-3 rounded-xl font-semibold flex items-center justify-center gap-2.5 transition-all text-xs cursor-pointer ${
                          activeSession.revealAnswer 
                            ? 'bg-emerald-600/50 hover:bg-emerald-650/50 text-emerald-200 border border-emerald-500/40' 
                            : 'bg-zinc-900 border border-zinc-850 hover:bg-zinc-800 text-zinc-400'
                        }`}
                      >
                        {activeSession.revealAnswer ? (
                          <>
                            <EyeOff className="w-4 h-4 text-emerald-400" />
                            <span>Gabarito Revelado</span>
                          </>
                        ) : (
                          <>
                            <HelpCircle className="w-4 h-4 text-emerald-500" />
                            <span>Revelar Resposta Correta</span>
                          </>
                        )}
                      </button>
                    )}

                    <div className="grid grid-cols-2 gap-2">
                      <button
                        onClick={simulateIncomingVotes}
                        className="py-2 px-3 rounded-xl bg-zinc-900 border border-zinc-850 text-[11px] font-semibold text-zinc-400 hover:bg-zinc-800 hover:text-white transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                        id="btn-simulate-addin-votes"
                      >
                        <Layers className="w-3.5 h-3.5 text-zinc-550" />
                        <span>Simular +5 Alunos</span>
                      </button>

                      <a
                        href={getTvUrl(activeSession.id)}
                        target="_blank"
                        rel="noreferrer"
                        className="py-2 px-3 rounded-xl bg-zinc-900 border border-zinc-850 text-[11px] font-semibold text-[#a3e635] hover:bg-zinc-800 hover:text-[#a3e635]/90 transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                      >
                        <span>Abrir Projetor</span>
                        <ExternalLink className="w-3.5 h-3.5" />
                      </a>
                    </div>
                  </div>
                </div>

                {/* VISUAL MONITOR OF ACTIVE QUESTION STATE */}
                <div className="p-3 bg-zinc-950 rounded-xl border border-zinc-850 space-y-2 mt-2">
                  <div className="flex items-center justify-between text-[11px] text-zinc-500">
                    <span className="font-mono uppercase text-[9px] tracking-wider font-semibold">Votação em tempo real</span>
                    <span className="font-mono bg-[#a3e635]/10 px-2 py-0.5 rounded text-[#a3e635] font-bold">
                      {syncStore.getVotesForQuestion(activeSession.id, activeSession.questions[activeSession.currentQuestionIndex]?.id).length} votos
                    </span>
                  </div>
                  
                  {/* Miniature live ticker */}
                  <div className="h-1.5 w-full bg-zinc-900 rounded-full overflow-hidden flex gap-0.5">
                    {activeSession.questions[activeSession.currentQuestionIndex]?.options.map((_, idx) => {
                      const total = syncStore.getVotesForQuestion(activeSession.id, activeSession.questions[activeSession.currentQuestionIndex]?.id).length;
                      const count = syncStore.getVotesForQuestion(activeSession.id, activeSession.questions[activeSession.currentQuestionIndex]?.id).filter(v => v.selectedOptionIndex === idx).length;
                      const w = total > 0 ? (count / total) * 100 : 0;
                      return (
                        <div 
                          key={idx} 
                          style={{ width: `${w}%` }} 
                          className={`h-full transition-all duration-300 ${
                            idx % 4 === 0 ? 'bg-indigo-500' :
                            idx % 4 === 1 ? 'bg-[#a3e635]' :
                            idx % 4 === 2 ? 'bg-amber-500' : 'bg-[#a3e635]/40'
                          }`} 
                        />
                      );
                    })}
                  </div>
                </div>
              </>
            )}
          </div>
        )}
      </main>

      {/* Footer Instructions / Copy link short keys */}
      <footer className="p-3 border-t border-[#27272a] bg-[#121214]/40 text-[11px] text-center text-zinc-650 mt-auto flex items-center justify-center gap-1 font-mono uppercase tracking-wider">
        <span>Pronto para apresentações no Microsoft PowerPoint.</span>
      </footer>
    </div>
  );
};
