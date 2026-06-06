/**
 * @license
 * SPDX-License-Identifier: Apache-2.5
 */

import React, { useState, useEffect } from 'react';
import { syncStore } from '../services/store';
import { Session, Question, SlideMapping } from '../types';
import { QRCodeDisplay } from '../components/QRCodeDisplay';
import { 
  FileText, 
  Layers, 
  QrCode, 
  Plus, 
  Info,
  Copy,
  Check,
  ChevronLeft,
  ChevronRight,
  Trash2,
  Edit2,
  HelpCircle,
  Eye,
  EyeOff,
  Laptop,
  Sparkles,
  Link,
  RefreshCw,
  FolderPlus
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
  
  // Simulated Slides State (for testing in the browser iframe)
  const [simulatedSlides, setSimulatedSlides] = useState<Array<{ id: string; label: string; active: boolean }>>([
    { id: 'slide-instructions', label: 'Slide 1: QR Code', active: true }
  ]);
  const [activeSimulatedSlideId, setActiveSimulatedSlideId] = useState<string>('slide-instructions');

  // Question Creation Form State (simplifying views)
  const [isAddingQuestion, setIsAddingQuestion] = useState<boolean>(false);
  const [editingQuestionId, setEditingQuestionId] = useState<string | null>(null);
  const [qText, setQText] = useState<string>('');
  const [qType, setQType] = useState<'alternativa' | 'aberta'>('alternativa');
  const [qResultsView, setQResultsView] = useState<'live' | 'results-slide-only'>('results-slide-only');
  const [qOptions, setQOptions] = useState<string[]>(['', '', '', '']);
  const [qCorrectIdx, setQCorrectIdx] = useState<number | null>(null);

  // New session inline creator state
  const [isAddingSession, setIsAddingSession] = useState<boolean>(false);
  const [newSessionName, setNewSessionName] = useState<string>('');
  const [isQuizMode, setIsQuizMode] = useState<boolean>(false);

  // Load and refresh state
  useEffect(() => {
    setSessions(syncStore.getSessions().filter(s => s.status === 'active'));

    // Detect Real Microsoft Office Environment
    if (typeof Office !== 'undefined') {
      Office.onReady((info: any) => {
        if (info.host === Office.HostType.PowerPoint) {
          setIsOfficeAvailable(true);
          setOfficeStatus('Conectado ao PowerPoint');
        } else {
          setOfficeStatus('Modo Web (Office carregado)');
        }
      });
    } else {
      setOfficeStatus('Modo Web (Slides Simulados)');
    }
  }, [activeSession]);

  // Update simulated slides whenever questions change
  useEffect(() => {
    if (!activeSession) return;
    const slides = [{ id: 'slide-instructions', label: 'Slide 1: QR Code de Acesso', active: activeSimulatedSlideId === 'slide-instructions' }];
    
    // Add slide questions & answers for mapping
    activeSession.questions.forEach((q, qIdx) => {
      const qSlideId = `slide-q-${q.id}`;
      const rSlideId = `slide-r-${q.id}`;

      slides.push({
        id: qSlideId,
        label: `Slide ${slides.length + 1}: ❓ Pergunta ${qIdx + 1}`,
        active: activeSimulatedSlideId === qSlideId
      });

      slides.push({
        id: rSlideId,
        label: `Slide ${slides.length + 1}: 📊 Gráfico P${qIdx + 1}`,
        active: activeSimulatedSlideId === rSlideId
      });
    });

    setSimulatedSlides(slides);
  }, [activeSession?.questions, activeSimulatedSlideId]);

  // PowerPoint Event Listener for slide change
  useEffect(() => {
    if (!isOfficeAvailable || !activeSession) return;

    let eventHandler = () => {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const slideRange = result.value;
            if (slideRange && slideRange.slides && slideRange.slides.length > 0) {
              const currentSlideId = slideRange.slides[0].id;
              syncToCurrentSlideId(currentSlideId);
            }
          }
        }
      );
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      eventHandler,
      (asyncResult: any) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("PPT Selection Changed Listener Registered.");
        }
      }
    );

    return () => {
      try {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          { handler: eventHandler }
        );
      } catch (_) {}
    };
  }, [isOfficeAvailable, activeSession?.id, activeSession?.questions]);

  // Dispatch slide movement transitions
  const syncToCurrentSlideId = (slideId: string) => {
    if (!activeSession) return;

    // Fetch mappings for current session
    const session = syncStore.getSession(activeSession.id);
    const mappings = session?.slideMappings || [];
    const matched = mappings.find(m => m.slideId === slideId);

    if (matched) {
      const qIndex = activeSession.questions.findIndex(q => q.id === matched.questionId);
      if (qIndex >= 0) {
        // Set active index
        syncStore.setQuestionIndex(activeSession.id, qIndex);

        // Turn on/off results visibility based on slide style preference
        if (matched.slideType === 'responses') {
          syncStore.toggleResults(activeSession.id, true);
        } else {
          // It's the question slide itself
          const currentQ = activeSession.questions[qIndex];
          const shouldShowLive = currentQ.resultsView === 'live';
          syncStore.toggleResults(activeSession.id, shouldShowLive);
        }
        onRefresh();
      }
    }
  };

  // Simulate slide trigger click in browser pane
  const handleSimulateSlideClick = (slideId: string) => {
    setActiveSimulatedSlideId(slideId);
    
    // Auto sync state inside database
    if (activeSession) {
      if (slideId === 'slide-instructions') {
        // Just holding QR Code slide, no active index shift is mandatory
        return;
      }

      // Check if this slide ID has a simulated mapping, if not create/mock it on the fly!
      const session = syncStore.getSession(activeSession.id);
      const currentMappings = session?.slideMappings || [];
      let matched = currentMappings.find(m => m.slideId === slideId);

      if (!matched) {
        // Map on-the-fly for simulator convenience
        const qIdToMap = slideId.replace('slide-q-', '').replace('slide-r-', '');
        const isResponsesType = slideId.startsWith('slide-r-');
        
        syncStore.saveSlideMapping(
          activeSession.id,
          slideId,
          qIdToMap,
          isResponsesType ? 'responses' : 'question'
        );
      }

      syncToCurrentSlideId(slideId);
    }
  };

  const triggerNotification = (message: string) => {
    setInsertSuccess(message);
    setTimeout(() => setInsertSuccess(null), 3000);
  };

  // Control carousel navigation indices
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

  // Inserir Pergunta no Slide Atual (Writes Text + maps current slide ID)
  const handleInsertQuestionText = () => {
    if (!activeSession) return;
    const currentQ = activeSession.questions[activeSession.currentQuestionIndex];
    if (!currentQ) return;

    // Get alternatives block if multiple choices
    const optionsText = currentQ.type === 'alternativa'
      ? currentQ.options.map((o, idx) => `  ${String.fromCharCode(65 + idx)}) ${o}`).join('\n')
      : '  [ Resposta Aberta / Digite no celular ]';

    const slideText = `❓ PERGUNTA:\n${currentQ.text}\n\n${optionsText}\n\n📱 Escaneie o QR Code inicial da apresentação para responder!`;

    if (isOfficeAvailable) {
      // 1. Get slide ID
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            
            // 2. Save Mapping
            syncStore.saveSlideMapping(activeSession.id, slideId, currentQ.id, 'question');
            
            // 3. Inject text safely
            Office.context.document.setSelectedDataAsync(
              slideText,
              { coercionType: Office.CoercionType.Text },
              (asyncResult: any) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  triggerNotification('Questão e Vínculo inseridos no slide!');
                  onRefresh();
                } else {
                  triggerNotification('Erro ao injetar texto.');
                }
              }
            );
          }
        }
      );
    } else {
      // Simulated Mode mapping
      const simulatedSlideId = `slide-q-${currentQ.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, currentQ.id, 'question');
      triggerNotification('Vínculo criado no (Slide de Pergunta)! Cole o conteúdo no PowerPoint.');
      
      // Copy question helper block
      navigator.clipboard.writeText(slideText).catch(() => {});
      onRefresh();
    }
  };

  // Inserir Slide de Respostas (Writes results link + maps slide ID)
  const handleInsertResponsesSlide = () => {
    if (!activeSession) return;
    const currentQ = activeSession.questions[activeSession.currentQuestionIndex];
    if (!currentQ) return;

    const dummyText = `📊 RESULTADOS AO VIVO:\n${currentQ.text}\n\n[ Os votos de todos os alunos aparecerão aqui ao vivo ]`;

    if (isOfficeAvailable) {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            
            // Save responses mapping
            syncStore.saveSlideMapping(activeSession.id, slideId, currentQ.id, 'responses');
            
            // Inserir marcação
            Office.context.document.setSelectedDataAsync(
              dummyText,
              { coercionType: Office.CoercionType.Text },
              (asyncResult: any) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  triggerNotification('Gráfico e Vínculo vinculados a este slide!');
                  onRefresh();
                }
              }
            );
          }
        }
      );
    } else {
      const simulatedSlideId = `slide-r-${currentQ.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, currentQ.id, 'responses');
      triggerNotification('Vínculo de Resultados criado! (Slide de Respostas)');
      onRefresh();
    }
  };

  // Insert QR Code Image to Slide
  const handleInsertQRCodeImage = async () => {
    if (!activeSession) return;
    const joinUrl = `${appUrl}?role=participant&session=${activeSession.id}`;

    try {
      const qrDataUrl = await QRCode.toDataURL(joinUrl, {
        width: 350,
        margin: 1,
        color: { dark: '#000000', light: '#ffffff' }
      });
      const base64Str = qrDataUrl.split(',')[1];

      if (isOfficeAvailable) {
        Office.context.document.setSelectedDataAsync(
          base64Str,
          { coercionType: Office.CoercionType.Image },
          (asyncResult: any) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              triggerNotification('QR Code inserido com sucesso!');
            } else {
              triggerNotification('Erro ao inserir QR Code.');
            }
          }
        );
      } else {
        // Fallback: Copy URL and trigger file download
        const link = document.createElement('a');
        link.href = qrDataUrl;
        link.download = `qrcode-sala-${activeSession.id}.png`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        triggerNotification('QR Code baixado de forma bem sucedida! Cole-o no PowerPoint.');
      }
    } catch (_) {}
  };

  // Create or update questions inline
  const handleOpenNewQuestion = () => {
    setQText('');
    setQType('alternativa');
    setQResultsView('results-slide-only');
    setQOptions(['', '', '', '']);
    setQCorrectIdx(null);
    setEditingQuestionId(null);
    setIsAddingQuestion(true);
  };

  const handleOpenEditQuestion = (q: Question) => {
    setQText(q.text);
    setQType(q.type || 'alternativa');
    setQResultsView(q.resultsView || 'results-slide-only');
    setQOptions([...(q.options || ['', '', '', ''])]);
    setQCorrectIdx(q.correctOptionIndex);
    setEditingQuestionId(q.id);
    setIsAddingQuestion(true);
  };

  const handleSaveQuestionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeSession) return;
    if (qText.trim() === '') {
      triggerNotification('Preencha o enunciado.');
      return;
    }

    // Process options
    const filteredOptions = qType === 'alternativa'
      ? qOptions.map(o => o.trim()).filter(o => o !== '')
      : [];

    if (qType === 'alternativa' && filteredOptions.length < 2) {
      triggerNotification('Insira pelo menos 2 alternativas.');
      return;
    }

    if (editingQuestionId) {
      // Edit mode
      syncStore.updateQuestionInSession(
        activeSession.id,
        editingQuestionId,
        qText.trim(),
        filteredOptions,
        qType === 'alternativa' ? qCorrectIdx : null
      );

      // Force update question metadata fields separately
      const updatedSessions = syncStore.getSessions();
      const updatedS = updatedSessions.find(s => s.id === activeSession.id);
      if (updatedS) {
        const targetQ = updatedS.questions.find(q => q.id === editingQuestionId);
        if (targetQ) {
          targetQ.type = qType;
          targetQ.resultsView = qResultsView;
        }
        localStorage.setItem('polling_sessions', JSON.stringify(updatedSessions));
      }

      triggerNotification('Pergunta atualizada.');
    } else {
      // Create brand new
      const currentS = syncStore.getSession(activeSession.id);
      if (currentS) {
        const newId = `q-${Date.now()}`;
        const newQuestion: Question = {
          id: newId,
          text: qText.trim(),
          options: filteredOptions,
          correctOptionIndex: qType === 'alternativa' ? qCorrectIdx : null,
          type: qType,
          resultsView: qResultsView
        };
        currentS.questions.push(newQuestion);
        
        // Save back
        const allSessions = syncStore.getSessions();
        const foundIdx = allSessions.findIndex(s => s.id === activeSession.id);
        if (foundIdx >= 0) {
          allSessions[foundIdx] = currentS;
          localStorage.setItem('polling_sessions', JSON.stringify(allSessions));
        }
      }
      triggerNotification('Pergunta criada com sucesso!');
    }

    setIsAddingQuestion(false);
    onRefresh();
  };

  const handleDeleteQuestion = (qId: string) => {
    if (!activeSession) return;
    syncStore.deleteQuestionFromSession(activeSession.id, qId);
    triggerNotification('Pergunta excluída.');
    onRefresh();
  };

  const handleCreateSessionInline = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSessionName.trim()) {
      triggerNotification('Escreva o nome da sessão.');
      return;
    }
    const created = syncStore.createSession(newSessionName.trim(), false, isQuizMode);
    setNewSessionName('');
    setIsQuizMode(false);
    setIsAddingSession(false);
    onSelectSession(created.id);
    triggerNotification(`Sessão "${created.name}" criada.`);
    onRefresh();
  };

  return (
    <div className="bg-[#09090b] text-zinc-100 min-h-screen flex flex-col relative text-xs font-sans h-full overflow-hidden select-none">
      
      {/* HEADER BAR */}
      <header className="p-3 border-b border-[#27272a] bg-[#121214] flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-2 h-2 rounded-full bg-[#a3e635] animate-ping" />
          <span className="font-extrabold uppercase tracking-tight text-white">PowerPoint Polling</span>
        </div>
        <div className="text-[9px] font-mono font-bold bg-[#121214] border border-[#27272a] px-2 py-0.5 rounded text-zinc-400">
          {officeStatus}
        </div>
      </header>

      {/* FLASH SUCCESS DIALOG */}
      {insertSuccess && (
        <div className="absolute top-12 left-3 right-3 z-50 bg-slate-900 border border-[#a3e635]/40 rounded-xl p-2.5 shadow-2xl flex items-center gap-2 animate-in fade-in slide-in-from-top-1">
          <Check className="w-4 h-4 text-[#a3e635]" />
          <p className="text-[11px] font-bold text-zinc-200">{insertSuccess}</p>
        </div>
      )}

      {/* CORE FRAME LAYOUT */}
      <div className="flex-1 overflow-y-auto p-3 space-y-4">
        
        {/* LIST ACTIVE SESSIONS */}
        {!activeSession ? (
          <div className="space-y-3.5">
            {isAddingSession ? (
              <form onSubmit={handleCreateSessionInline} className="p-4 bg-[#121214] rounded-2xl border border-[#27272a] space-y-3">
                <h4 className="font-bold text-white text-xs uppercase tracking-wider font-mono">Nova Sessão de Aula</h4>
                <div>
                  <label className="block text-[10px] text-zinc-550 font-mono mb-1">NOME DA SESSÃO / AULA</label>
                  <input
                    type="text"
                    required
                    value={newSessionName}
                    onChange={(e) => setNewSessionName(e.target.value)}
                    placeholder="Aula Teórica de Hoje"
                    className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2.5 focus:border-[#a3e635] focus:outline-none"
                  />
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    id="addin-cb-quiz"
                    checked={isQuizMode}
                    onChange={(e) => setIsQuizMode(e.target.checked)}
                    className="accent-[#a3e635]"
                  />
                  <label htmlFor="addin-cb-quiz" className="text-[11px] text-zinc-400 cursor-pointer">Ativar Modo Quiz Competitivo (concede XP)</label>
                </div>
                <div className="flex items-center gap-2 pt-2 justify-end">
                  <button
                    type="button"
                    onClick={() => setIsAddingSession(false)}
                    className="py-1.5 px-3 rounded bg-zinc-850 hover:bg-zinc-800 text-zinc-400"
                  >
                    Voltar
                  </button>
                  <button
                    type="submit"
                    className="py-1.5 px-3 rounded bg-[#a3e635] text-zinc-950 font-black flex items-center gap-1"
                  >
                    <FolderPlus className="w-3.5 h-3.5" />
                    <span>CRIAR</span>
                  </button>
                </div>
              </form>
            ) : (
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="font-mono text-[10px] tracking-wider text-zinc-500 uppercase">Escolha a Sessão Ativa:</span>
                  <button
                    onClick={() => setIsAddingSession(true)}
                    className="bg-[#a3e635] text-zinc-950 font-black py-1 px-2.5 rounded-lg text-[10px] flex items-center gap-1 uppercase"
                  >
                    <Plus className="w-3 h-3" />
                    <span>Nova Sessão</span>
                  </button>
                </div>

                {sessions.length === 0 ? (
                  <div className="text-center py-10 bg-[#121214]/50 border border-dashed border-[#27272a] rounded-2xl p-4">
                    <Layers className="w-8 h-8 text-zinc-700 mx-auto opacity-40 mb-2" />
                    <h3 className="font-semibold text-zinc-350">Nenhuma aula criada</h3>
                    <p className="text-[10px] text-zinc-500 mt-1">Clique no botão acima para abrir uma nova aula de enquetes.</p>
                  </div>
                ) : (
                  <div className="space-y-2">
                    {sessions.map((sess) => (
                      <div
                        key={sess.id}
                        onClick={() => onSelectSession(sess.id)}
                        className="p-3 bg-[#121214] border border-[#27272a] hover:border-zinc-700 rounded-xl transition-all cursor-pointer flex items-center justify-between"
                      >
                        <div>
                          <h4 className="font-bold text-white text-xs">{sess.name}</h4>
                          <span className="text-[10px] text-zinc-500 font-mono">ID: {sess.id} • {sess.questions.length} perguntas</span>
                        </div>
                        <ChevronRight className="w-4 h-4 text-[#a3e635]" />
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        ) : (
          
          /* SECTION: INSIDE AN ACTIVE SESSION */
          <div className="space-y-4">
            
            {/* SESSION INFO BREADCRUMB */}
            <div className="p-3 bg-[#121214]/90 border border-[#27272a] rounded-2xl flex items-center justify-between">
              <div>
                <span className="text-[9px] font-mono uppercase text-[#a3e635] font-black">SESSÃO SELECIONADA:</span>
                <h4 className="font-extrabold text-white text-xs">{activeSession.name}</h4>
              </div>
              <button
                onClick={() => {
                  onSelectSession('');
                  setIsAddingQuestion(false);
                }}
                className="text-[10px] text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-750 p-1.5 px-2.5 rounded-lg font-mono font-bold"
              >
                Voltar
              </button>
            </div>

            {/* INTEGRATED OFFICE DESIGN PREVIEW (SIMULATOR FOR CHATS) */}
            {!isOfficeAvailable && (
              <div className="p-2.5 bg-[#121214] border border-orange-500/10 rounded-2xl space-y-2">
                <div className="flex items-center gap-1.5 text-[10px] font-mono text-orange-400">
                  <Laptop className="w-3.5 h-3.5" />
                  <span>SIMULADOR DE APRESENTAÇÃO POWERPOINT:</span>
                </div>
                <p className="text-[10px] text-zinc-400 leading-relaxed">
                  Para testar a sincronização automática no celular de teste do preview, use esta barra de slides do PowerPoint como se estivesse apresentando:
                </p>
                <div className="flex gap-1.5 overflow-x-auto pb-1.5 select-none scrollbar-thin">
                  {simulatedSlides.map((slide) => (
                    <button
                      key={slide.id}
                      onClick={() => handleSimulateSlideClick(slide.id)}
                      className={`text-[9px] font-bold p-1 px-2.5 shrink-0 rounded transition-colors cursor-pointer ${
                        slide.active
                          ? 'bg-[#a3e635] text-zinc-950 font-black shadow'
                          : 'bg-zinc-900 text-zinc-450 border border-zinc-800'
                      }`}
                    >
                      {slide.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* COMPACT BUTTON FOR MAIN CORE QR CODE SLIDE ACTION */}
            <div className="p-3 bg-[#121214] border border-[#27272a] rounded-2xl space-y-2.5">
              <div className="flex items-center justify-between">
                <span className="text-[9px] text-[#a3e635] font-mono font-bold tracking-wider">PASSO 1: CONFIGURAR ACESSO</span>
                <span className="text-[9px] text-zinc-500 font-mono">ID: {activeSession.id}</span>
              </div>
              <button
                onClick={handleInsertQRCodeImage}
                className="w-full bg-[#1e1e24] hover:bg-zinc-800 text-zinc-150 p-3 rounded-xl border border-zinc-800 hover:border-zinc-700 font-bold flex items-center justify-center gap-2 transition-all cursor-pointer shadow-sm text-xs"
              >
                <QrCode className="w-4 h-4 text-[#a3e635] animate-pulse shrink-0" />
                <span className="font-mono uppercase tracking-wide">Inserir QR Code de Votação</span>
              </button>
            </div>

            {/* FORM OR LIST SELECTOR */}
            {isAddingQuestion ? (
              
              /* RENDER QUESTION EDITOR */
              <form onSubmit={handleSaveQuestionSubmit} className="p-4 bg-[#121214] border border-zinc-800 rounded-2xl space-y-4">
                <div className="flex items-center justify-between pb-2 border-b border-zinc-850">
                  <h4 className="font-bold text-white text-xs uppercase tracking-wider font-mono">
                    {editingQuestionId ? 'Editar Pergunta' : 'Nova Pergunta'}
                  </h4>
                  <button
                    type="button"
                    onClick={() => setIsAddingQuestion(false)}
                    className="text-[10px] text-zinc-450 hover:text-zinc-200"
                  >
                    Voltar
                  </button>
                </div>

                {/* TEXT CONTAINER */}
                <div className="space-y-1">
                  <label className="block text-[10px] font-mono text-zinc-400">ENUNCIADO DA QUESTÃO / ENQUETE</label>
                  <textarea
                    required
                    rows={3}
                    value={qText}
                    onChange={(e) => setQText(e.target.value)}
                    placeholder="Ex: Qual dessas opções é correta sobre..."
                    className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-xl p-2.5 text-zinc-100 placeholder-zinc-700 focus:outline-none focus:border-[#a3e635]"
                  />
                </div>

                {/* TYPE SELECTOR */}
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-mono text-zinc-400 mb-1">TIPO DE PERGUNTA</label>
                    <select
                      value={qType}
                      onChange={(e) => setQType(e.target.value as 'alternativa' | 'aberta')}
                      className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2 focus:outline-none focus:border-[#a3e635]"
                    >
                      <option value="alternativa">Múltipla Escolha (Alternativas)</option>
                      <option value="aberta">Aberta (Resposta de Texto)</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-[10px] font-mono text-zinc-400 mb-1">VISUALIZAÇÃO DE RESULTADOS</label>
                    <select
                      value={qResultsView}
                      onChange={(e) => setQResultsView(e.target.value as 'live' | 'results-slide-only')}
                      className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2 focus:outline-none focus:border-[#a3e635]"
                    >
                      <option value="results-slide-only">Mostrar somente no Slide de Resposta</option>
                      <option value="live">Mostrar respostas ao vivo</option>
                    </select>
                  </div>
                </div>

                {/* ALTERNATIVES EDITOR */}
                {qType === 'alternativa' && (
                  <div className="space-y-2">
                    <label className="block text-[10px] font-mono text-zinc-400">ALTERNATIVAS E MARCAÇÃO DE CERTA (GABARITO)</label>
                    <div className="space-y-1.5">
                      {qOptions.map((opt, idx) => {
                        const letChar = String.fromCharCode(65 + idx);
                        const isCorrect = qCorrectIdx === idx;
                        return (
                          <div key={idx} className="flex items-center gap-2">
                            {/* Quiz correctness toggle button */}
                            <button
                              type="button"
                              onClick={() => setQCorrectIdx(isCorrect ? null : idx)}
                              className={`w-6 h-6 rounded flex items-center justify-center font-mono text-[10px] font-bold border transition-colors ${
                                isCorrect
                                  ? 'bg-[#a3e635] text-zinc-950 border-[#a3e635]'
                                  : 'bg-[#09090b] text-zinc-550 border-zinc-800 hover:border-zinc-700'
                              }`}
                              title={isCorrect ? 'Resposta Certa' : 'Marcar como certa'}
                            >
                              {letChar}
                            </button>
                            <input
                              type="text"
                              value={opt}
                              onChange={(e) => {
                                const copy = [...qOptions];
                                copy[idx] = e.target.value;
                                setQOptions(copy);
                              }}
                              placeholder={`Alternativa ${letChar}...`}
                              className="flex-1 text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-1.5 text-zinc-250 placeholder-zinc-700"
                            />
                            {qOptions.length > 2 && (
                              <button
                                type="button"
                                onClick={() => {
                                  if (qOptions.length <= 2) return;
                                  const copy = [...qOptions];
                                  copy.splice(idx, 1);
                                  setQOptions(copy);
                                  if (qCorrectIdx === idx) {
                                    setQCorrectIdx(null);
                                  } else if (qCorrectIdx !== null && qCorrectIdx > idx) {
                                    setQCorrectIdx(qCorrectIdx - 1);
                                  }
                                }}
                                className="text-zinc-600 hover:text-red-400 p-1"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    
                    {qOptions.length < 8 && (
                      <button
                        type="button"
                        onClick={() => setQOptions([...qOptions, ''])}
                        className="text-[10px] font-bold text-orange-400 hover:text-orange-300 flex items-center gap-1 mt-1 bg-zinc-900 border border-zinc-850 px-2 py-1 rounded"
                      >
                        + Adicionar Opção
                      </button>
                    )}
                  </div>
                )}

                <div className="flex justify-end gap-2 pt-2 border-t border-zinc-850">
                  <button
                    type="button"
                    onClick={() => setIsAddingQuestion(false)}
                    className="py-2 px-3 hover:bg-zinc-800 rounded-lg text-zinc-400 font-bold"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    className="py-2 px-4 bg-orange-600 hover:bg-orange-700 text-white font-bold rounded-lg shadow cursor-pointer"
                  >
                    Salvar Pergunta
                  </button>
                </div>
              </form>
            ) : (
              
              /* RENDER LIST OF CREATED QUESTIONS & PRESENTATION NAVIGATION CAROUSEL */
              <div className="space-y-4">
                
                {/* LIST HEADER WITH CREATION TRIGGER */}
                <div className="flex items-center justify-between">
                  <span className="font-mono text-[9px] tracking-wider text-zinc-500 uppercase">Perguntas da Aula ({activeSession.questions.length}):</span>
                  <button
                    onClick={handleOpenNewQuestion}
                    className="bg-[#a3e635]/10 border border-[#a3e635]/35 hover:bg-[#a3e635]/15 text-[#a3e635] text-[10px] font-black p-1 px-2.5 rounded-lg flex items-center gap-1 uppercase tracking-wide"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Adicionar Questão</span>
                  </button>
                </div>

                {/* CAROUSEL CARDS VIEW */}
                {activeSession.questions.length === 0 ? (
                  <div className="text-center p-8 bg-[#121214]/40 border border-[#27272a] rounded-2xl text-zinc-500">
                    Nenhuma pergunta criada nesta aula. Crie uma acima para começar a preencher!
                  </div>
                ) : (
                  <div className="space-y-3.5">
                    
                    {/* ACTIVE QUESTION SLIDES PANEL */}
                    <div className="p-3.5 bg-[#121214] border border-[#27272a] rounded-2xl relative shadow">
                      
                      {/* Active carousel indicator badge */}
                      <div className="flex items-center justify-between pb-2 border-b border-zinc-800/80 mb-2 font-mono text-[10px] text-zinc-500">
                        <span className="bg-[#a3e635]/10 text-[#a3e635] px-2 py-0.5 rounded border border-[#a3e635]/20 font-bold uppercase tracking-wider">
                          Pergunta ativa: Slide {activeSession.currentQuestionIndex + 1}
                        </span>
                        <span>{activeSession.questions[activeSession.currentQuestionIndex]?.type === 'aberta' ? '💬 Aberta' : '❓ Alternativas'}</span>
                      </div>

                      {/* Title block */}
                      <p className="font-bold text-white text-[12px] leading-relaxed">
                        {activeSession.questions[activeSession.currentQuestionIndex]?.text}
                      </p>

                      {/* Live Vote scale inside this slide card component */}
                      <div className="mt-3.5 p-2 bg-[#09090b]/40 rounded-xl border border-zinc-900 flex items-center justify-between font-mono text-[10px] text-zinc-500">
                        <span>Respostas recebidas:</span>
                        <span className="font-black text-[#a3e635]">
                          {syncStore.getVotesForQuestion(activeSession.id, activeSession.questions[activeSession.currentQuestionIndex]?.id).length} votos
                        </span>
                      </div>

                      {/* In slides insertion buttons! */}
                      <div className="space-y-2 pt-3 border-t border-zinc-850 mt-3.5">
                        <span className="text-[9px] text-zinc-500 font-mono tracking-wider block uppercase">PASSO 2: INSERIR NO SLIDE ATUAL DE SUA APRESENTAÇÃO</span>
                        <div className="grid grid-cols-2 gap-2">
                          <button
                            onClick={handleInsertQuestionText}
                            className="p-2.5 rounded-xl bg-zinc-900 border border-zinc-800 hover:border-zinc-700 hover:bg-zinc-850 text-zinc-200 text-center flex flex-col items-center justify-center gap-1.5 transition-all cursor-pointer active:scale-95"
                            title="Insere enunciado e alternativas e vincula a este slide"
                          >
                            <FileText className="w-4 h-4 text-[#a3e635]" />
                            <span className="font-mono text-[9px] uppercase font-bold tracking-tight">1. Slide Pergunta</span>
                          </button>
                          <button
                            onClick={handleInsertResponsesSlide}
                            className="p-2.5 rounded-xl bg-zinc-900 border border-zinc-800 hover:border-zinc-700 hover:bg-zinc-850 text-zinc-200 text-center flex flex-col items-center justify-center gap-1.5 transition-all cursor-pointer active:scale-95"
                            title="Insere layout de gráfico e vincula o próximo slide para abrir respostas ao vivo"
                          >
                            <Layers className="w-4 h-4 text-emerald-400" />
                            <span className="font-mono text-[9px] uppercase font-bold tracking-tight">2. Slide Resposta</span>
                          </button>
                        </div>
                      </div>

                      {/* Carousel Arrow Controllers */}
                      <div className="flex items-center gap-2 mt-4">
                        <button
                          onClick={handlePrevQuestion}
                          disabled={activeSession.currentQuestionIndex === 0}
                          className="flex-1 py-1.5 bg-[#1e1e24] border border-zinc-850 text-zinc-300 disabled:opacity-30 rounded hover:bg-zinc-800 transition-colors text-[10px] font-bold uppercase flex items-center justify-center gap-1"
                        >
                          <ChevronLeft className="w-4 text-zinc-500" />
                          Anterior
                        </button>
                        <button
                          onClick={handleNextQuestion}
                          disabled={activeSession.currentQuestionIndex === activeSession.questions.length - 1}
                          className="flex-1 py-1.5 bg-[#1e1e24] border border-zinc-850 text-zinc-300 disabled:opacity-30 rounded hover:bg-zinc-800 transition-colors text-[10px] font-bold uppercase flex items-center justify-center gap-1"
                        >
                          Próxima
                          <ChevronRight className="w-4 text-zinc-500" />
                        </button>
                      </div>

                      {/* Live state results override toggle switch */}
                      <div className="flex items-center justify-between pt-3 border-t border-zinc-850 mt-3 font-mono text-[10px]">
                        <span className="text-zinc-500">Exibir Gráfico Agora:</span>
                        <button
                          onClick={() => {
                            syncStore.toggleResults(activeSession.id, !activeSession.showResults);
                            onRefresh();
                          }}
                          className={`p-1 px-2.5 rounded text-[9px] font-bold flex items-center gap-1 ${
                            activeSession.showResults ? 'bg-indigo-950/40 border border-indigo-500/35 text-indigo-400' : 'bg-zinc-900 text-zinc-550'
                          }`}
                        >
                          {activeSession.showResults ? (
                            <>
                              <Eye className="w-3.5 h-3.5" />
                              <span>ATIVO</span>
                            </>
                          ) : (
                            <>
                              <EyeOff className="w-3.5 h-3.5" />
                              <span>OCULTO</span>
                            </>
                          )}
                        </button>
                      </div>
                    </div>

                    {/* INDEX LIST OF MINIFIED QUESTION METADATA CELLS */}
                    <div className="space-y-2 pt-2">
                      <span className="text-[10px] text-zinc-500 font-mono tracking-wider block uppercase">ORGANIZAÇÃO E GESTÃO:</span>
                      <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                        {activeSession.questions.map((q, idx) => {
                          const isActiveQ = activeSession.currentQuestionIndex === idx;
                          // Check active mappings
                          const qMappings = (activeSession.slideMappings || []).filter(m => m.questionId === q.id);

                          return (
                            <div
                              key={q.id}
                              className={`p-2.5 rounded-xl border flex items-center justify-between transition-colors ${
                                isActiveQ
                                  ? 'bg-zinc-900 border-zinc-700 text-white font-semibold'
                                  : 'bg-[#121214] border-zinc-900 text-zinc-400 hover:bg-zinc-900/40'
                              }`}
                            >
                              <div className="truncate pr-1.5 flex items-center gap-2">
                                <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${isActiveQ ? 'bg-[#a3e635] text-zinc-950' : 'bg-zinc-800 text-zinc-400'}`}>
                                  P{idx + 1}
                                </span>
                                <div className="truncate text-left leading-none">
                                  <span className="text-[11px] truncate block">{q.text}</span>
                                  {qMappings.length > 0 && (
                                    <span className="text-[9px] text-[#a3e635] font-mono leading-none block mt-0.5">
                                      🔗 {qMappings.length === 2 ? 'Pergunta e Gráfico Mapeados' : '1 Slide Mapeado'}
                                    </span>
                                  )}
                                </div>
                              </div>

                              <div className="flex items-center gap-1.5 shrink-0">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleOpenEditQuestion(q);
                                  }}
                                  className="p-1 hover:bg-zinc-800 rounded text-zinc-450 hover:text-white"
                                  title="Editar"
                                >
                                  <Edit2 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleDeleteQuestion(q.id);
                                  }}
                                  className="p-1 hover:bg-zinc-800 hover:text-red-400 rounded text-zinc-550"
                                  title="Excluir"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {/* FIXED FOOTER BRAND INDICATOR */}
      <footer className="p-2 border-t border-[#27272a] bg-[#121214]/60 text-center text-[9px] text-zinc-650 font-mono tracking-wider uppercase font-medium">
        <span>Microsoft PowerPoint Companion V2</span>
      </footer>
    </div>
  );
};
