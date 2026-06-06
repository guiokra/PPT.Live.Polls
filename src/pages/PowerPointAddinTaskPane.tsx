/**
 * @license
 * SPDX-License-Identifier: Apache-2.5
 */

import React, { useState, useEffect } from 'react';
import { syncStore } from '../services/store';
import { Session, Question } from '../types';
import { 
  FileText, 
  Layers, 
  QrCode, 
  Plus, 
  Check, 
  ChevronRight, 
  Trash2, 
  Edit2, 
  ArrowLeft,
  MoreVertical,
  Trophy,
  PlusCircle,
  HelpCircle,
  Grid,
  FileSpreadsheet
} from 'lucide-react';
import QRCode from 'qrcode';

declare const Office: any;

// Safe access helpers for Office object properties to prevent crashes on standalone browser tests
const getOfficeEventType = () => {
  return (typeof Office !== 'undefined' && Office && Office.EventType) 
    ? Office.EventType 
    : { DocumentSelectionChanged: 'documentSelectionChanged' };
};

const getOfficeCoercionType = () => {
  return (typeof Office !== 'undefined' && Office && Office.CoercionType) 
    ? Office.CoercionType 
    : { SlideRange: 'slideRange', Text: 'text', Image: 'image' };
};

const getOfficeAsyncResultStatus = () => {
  return (typeof Office !== 'undefined' && Office && Office.AsyncResultStatus) 
    ? Office.AsyncResultStatus 
    : { Succeeded: 'succeeded', Failed: 'failed' };
};

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
  const [isOfficeAvailable, setIsOfficeAvailable] = useState<boolean>(false);
  const [insertSuccess, setInsertSuccess] = useState<string | null>(null);

  // Toggle for Question Creation Form View inside active session
  const [isFormOpen, setIsFormOpen] = useState<boolean>(false);

  // Question Form state
  const [qText, setQText] = useState<string>('');
  const [qType, setQType] = useState<'alternativa' | 'aberta'>('alternativa');
  const [qOptions, setQOptions] = useState<string[]>(['', '', '', '']);
  const [qCorrectIdx, setQCorrectIdx] = useState<number | null>(null);
  const [editingQuestionId, setEditingQuestionId] = useState<string | null>(null);

  // Session inline creator
  const [newSessionName, setNewSessionName] = useState<string>('');

  // Dropdown menus for questions
  const [activeMenuId, setActiveMenuId] = useState<string | null>(null);

  const handleSelectionChanged = () => {
    if (!activeSession || typeof Office === 'undefined' || !Office.context || !Office.context.document) return;
    try {
      const coercionType = getOfficeCoercionType();
      const asyncResultStatus = getOfficeAsyncResultStatus();
      Office.context.document.getSelectedDataAsync(
        coercionType.SlideRange,
        (result: any) => {
          if (result && result.status === asyncResultStatus.Succeeded && result.value && result.value.slides && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            const mappings = activeSession.slideMappings || [];
            const found = mappings.find(m => m.slideId === slideId);
            if (found) {
              const questionIdx = activeSession.questions.findIndex(q => q.id === found.questionId);
              if (questionIdx >= 0) {
                syncStore.setQuestionIndex(activeSession.id, questionIdx);
                if (found.slideType === 'responses') {
                  syncStore.toggleResults(activeSession.id, true);
                } else {
                  syncStore.toggleResults(activeSession.id, false);
                }
                onRefresh();
              }
            }
          }
        }
      );
    } catch (e) {
      console.warn('Silent fallback slide navigation tracking:', e);
    }
  };

  useEffect(() => {
    setSessions(syncStore.getSessions().filter(s => s.status === 'active'));

    let isMounted = true;

    const initOffice = () => {
      if (typeof Office !== 'undefined') {
        if (Office.context && Office.context.document) {
          setIsOfficeAvailable(true);
        }
        Office.onReady((info: any) => {
          if (!isMounted) return;
          const hostType = (info && info.host) || '';
          if (hostType === 'PowerPoint' || (Office.HostType && hostType === Office.HostType.PowerPoint) || Office.context) {
            setIsOfficeAvailable(true);
            try {
              // Register slide change detection inside PowerPoint
              const eventType = getOfficeEventType();
              Office.context.document.addHandlerAsync(
                eventType.DocumentSelectionChanged,
                handleSelectionChanged,
                () => {}
              );
            } catch (_) {}
          }
        });
      }
    };

    // Safely inject Office.js dynamically only inside this taskpane component
    if (typeof Office === 'undefined') {
      const scriptId = 'office-js-script';
      let script = document.getElementById(scriptId) as HTMLScriptElement;
      if (!script) {
        script = document.createElement('script');
        script.id = scriptId;
        script.src = 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js';
        script.async = true;
        script.onload = () => {
          if (isMounted) initOffice();
        };
        script.onerror = () => {
          console.warn('Office.js could not be loaded or was blocked by headers.');
        };
        document.head.appendChild(script);
      } else {
        // Listen to load event if script already exists but is loading
        const handleScriptLoad = () => {
          if (isMounted) initOffice();
        };
        script.addEventListener('load', handleScriptLoad);
        return () => {
          isMounted = false;
          script.removeEventListener('load', handleScriptLoad);
        };
      }
    } else {
      initOffice();
    }

    return () => {
      isMounted = false;
      if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
        try {
          const eventType = getOfficeEventType();
          Office.context.document.removeHandlerAsync(
            eventType.DocumentSelectionChanged,
            handleSelectionChanged,
            () => {}
          );
        } catch (_) {}
      }
    };
  }, [activeSession]);

  const triggerNotification = (message: string) => {
    setInsertSuccess(message);
    setTimeout(() => setInsertSuccess(null), 2500);
  };

  const handleCreateSessionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSessionName.trim()) {
      triggerNotification('Digite o nome da aula/questionário.');
      return;
    }
    const created = syncStore.createSession(newSessionName.trim(), false, false);
    setNewSessionName('');
    onSelectSession(created.id);
    triggerNotification('Aula criada com sucesso!');
    onRefresh();
  };

  // Direct injection of the QR Code image on the current slide
  const handleInsertQRCodeImage = async () => {
    if (!activeSession) return;
    const joinUrl = `${appUrl}?role=participant&session=${activeSession.id}`;
    const fallbackText = `📱 ACESSO RÁPIDO DO ALUNO:\n\n1. Escaneie o QR Code inicial\n2. Ou acesse: ${appUrl}\n3. Código da Sala: ${activeSession.id}`;

    try {
      const qrDataUrl = await QRCode.toDataURL(joinUrl, {
        width: 400,
        margin: 1,
        color: { dark: '#000000', light: '#ffffff' }
      });
      const base64Str = qrDataUrl.split(',')[1];

      // Copy accessibility instruction to clipboard automatically
      try {
        await navigator.clipboard.writeText(fallbackText);
      } catch (_) {}

      if (isOfficeAvailable && typeof Office !== 'undefined' && Office && Office.context && Office.context.document) {
        const coercionType = getOfficeCoercionType();
        const asyncResultStatus = getOfficeAsyncResultStatus();
        Office.context.document.setSelectedDataAsync(
          base64Str,
          { coercionType: coercionType.Image },
          (asyncResult: any) => {
            if (asyncResult && asyncResult.status === asyncResultStatus.Succeeded) {
              triggerNotification('QR Code inserido no PowerPoint! (Também copiado para o Clipboard)');
            } else {
              // Custom guide error message
              triggerNotification('Selecione uma caixa de imagem/forma no PowerPoint ou clique no slide. Instruções de login copiadas! Use Ctrl+V para colar.');
            }
          }
        );
      } else {
        // Fallback simulate link
        triggerNotification('QR Code gerado! Instruções copiadas para a área de transferência. Use Ctrl+V.');
        syncStore.saveSlideMapping(activeSession.id, 'slide-instructions', 'instructions', 'question');
        onRefresh();
      }
    } catch (_) {
      triggerNotification('Erro ao gerar imagem do QR Code. Copiado instruções de texto.');
      try {
        await navigator.clipboard.writeText(fallbackText);
      } catch (_) {}
    }
  };

  const handleInsertQuestionText = async (q: Question) => {
    if (!activeSession) return;

    const optionsText = q.type === 'alternativa'
      ? q.options.map((o, idx) => `  ${String.fromCharCode(65 + idx)}) ${o}`).join('\n')
      : '  [ Resposta Aberta / Digite no celular ]';

    const slideText = `❓ PERGUNTA:\n${q.text}\n\n${optionsText}\n\n📱 Escaneie o QR Code inicial para responder!`;

    try {
      await navigator.clipboard.writeText(slideText);
    } catch (_) {}

    if (isOfficeAvailable && typeof Office !== 'undefined' && Office && Office.context && Office.context.document) {
      const coercionType = getOfficeCoercionType();
      const asyncResultStatus = getOfficeAsyncResultStatus();
      Office.context.document.getSelectedDataAsync(
        coercionType.SlideRange,
        (result: any) => {
          if (result && result.status === asyncResultStatus.Succeeded && result.value && result.value.slides && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            syncStore.saveSlideMapping(activeSession.id, slideId, q.id, 'question');
            
            Office.context.document.setSelectedDataAsync(
              slideText,
              { coercionType: coercionType.Text },
              (asyncResult: any) => {
                if (asyncResult && asyncResult.status === asyncResultStatus.Succeeded) {
                  triggerNotification('Questão adicionada ao slide! Se quiser, pule ou use Ctrl+V para duplicar.');
                  onRefresh();
                } else {
                  triggerNotification('Slide vinculado com sucesso! Questão copiada para área de transferência; use Ctrl+V.');
                  onRefresh();
                }
              }
            );
          } else {
            // Even if slide metadata is not fully reachable, we map it simulated or guide them
            triggerNotification('Selecione uma caixa de texto no slide para inserir a pergunta diretamente ou use Ctrl+V para colar!');
          }
        }
      );
    } else {
      const simulatedSlideId = `slide-q-${q.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, q.id, 'question');
      triggerNotification('Questão vinculada ao slide! Texto copiado para área de transferência (Ctrl+V).');
      onRefresh();
    }
  };

  const handleInsertResponsesSlide = async (q: Question) => {
    if (!activeSession) return;

    const dummyText = `📊 EXIBIR RESULTADOS:\n${q.text}\n\n[ Os votos de todos os alunos aparecerão aqui ao vivo ]`;

    try {
      await navigator.clipboard.writeText(dummyText);
    } catch (_) {}

    if (isOfficeAvailable && typeof Office !== 'undefined' && Office && Office.context && Office.context.document) {
      const coercionType = getOfficeCoercionType();
      const asyncResultStatus = getOfficeAsyncResultStatus();
      Office.context.document.getSelectedDataAsync(
        coercionType.SlideRange,
        (result: any) => {
          if (result && result.status === asyncResultStatus.Succeeded && result.value && result.value.slides && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            syncStore.saveSlideMapping(activeSession.id, slideId, q.id, 'responses');
            
            Office.context.document.setSelectedDataAsync(
              dummyText,
              { coercionType: coercionType.Text },
              (asyncResult: any) => {
                if (asyncResult && asyncResult.status === asyncResultStatus.Succeeded) {
                  triggerNotification('Tela de respostas vinculada! (Também copiada ao clipboard)');
                  onRefresh();
                } else {
                  triggerNotification('Slide de respostas vinculado! Texto copiado para área de transferência; use Ctrl+V.');
                  onRefresh();
                }
              }
            );
          } else {
            triggerNotification('Selecione uma caixa de texto no PowerPoint ou use Ctrl+V para colar os resultados!');
          }
        }
      );
    } else {
      const simulatedSlideId = `slide-r-${q.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, q.id, 'responses');
      triggerNotification('Slide de respostas mapeado! Dados copiados ao clipboard (Ctrl+V).');
      onRefresh();
    }
  };

  const handleSaveQuestionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!activeSession) return;
    if (qText.trim() === '') {
      triggerNotification('Digite o enunciado da pergunta.');
      return;
    }

    const filteredOptions = qType === 'alternativa'
      ? qOptions.map(o => o.trim()).filter(o => o !== '')
      : [];

    if (qType === 'alternativa' && filteredOptions.length < 2) {
      triggerNotification('Digite pelo menos 2 alternativas.');
      return;
    }

    if (editingQuestionId) {
      syncStore.updateQuestionInSession(
        activeSession.id,
        editingQuestionId,
        qText.trim(),
        filteredOptions,
        qCorrectIdx,
        qType,
        'results-slide-only'
      );
      triggerNotification('Pergunta atualizada!');
    } else {
      syncStore.addQuestionToSession(
        activeSession.id,
        qText.trim(),
        filteredOptions,
        qType === 'alternativa' ? qCorrectIdx : null,
        qType,
        'results-slide-only'
      );
      triggerNotification('Pergunta adicionada!');
    }

    // Reset and close form
    setQText('');
    setQOptions(['', '', '', '']);
    setQCorrectIdx(null);
    setEditingQuestionId(null);
    setIsFormOpen(false);
    onRefresh();
  };

  const handleStartEdit = (q: Question) => {
    setQText(q.text);
    setQType(q.type || 'alternativa');
    setQOptions(q.options && q.options.length ? [...q.options, '', '', ''].slice(0, 4) : ['', '', '', '']);
    setQCorrectIdx(q.correctOptionIndex);
    setEditingQuestionId(q.id);
    setIsFormOpen(true);
    setActiveMenuId(null);
  };

  const handleDeleteQuestion = (qId: string) => {
    if (!activeSession) return;
    syncStore.deleteQuestionFromSession(activeSession.id, qId);
    triggerNotification('Pergunta removida.');
    setActiveMenuId(null);
    onRefresh();
  };

  const handleClearForm = () => {
    setQText('');
    setQOptions(['', '', '', '']);
    setQCorrectIdx(null);
    setEditingQuestionId(null);
    setIsFormOpen(false);
  };

  return (
    <div className="bg-[#f8fafc] text-slate-800 min-h-screen flex flex-col relative text-xs font-sans h-full overflow-hidden select-none">
      
      {/* VEVOX-STYLE SLIGHT HEADER BAR WITH ACTIONS */}
      {activeSession ? (
        <header className="p-3 bg-white border-b border-slate-200/80 flex items-center justify-between shadow-sm shrink-0">
          <div className="flex items-center gap-2 truncate pr-2">
            <span 
              onClick={() => onSelectSession('')}
              className="font-black text-slate-800 text-[13px] hover:text-[#2d7f8d] transition-colors cursor-pointer truncate"
              title="Voltar para a seleção de Aulas"
            >
              {activeSession.name}
            </span>
          </div>
          <button
            onClick={() => setActiveMenuId(activeMenuId === 'session-actions' ? null : 'session-actions')}
            className="p-1 px-1.5 hover:bg-slate-100 rounded text-slate-500 hover:text-slate-900 transition-colors cursor-pointer flex items-center"
          >
            <MoreVertical className="w-4 h-4" />
          </button>

          {activeMenuId === 'session-actions' && (
            <div className="absolute top-12 right-2 bg-white border border-slate-200 rounded-xl shadow-xl py-1.5 z-40 shrink-0 w-44 animate-in fade-in duration-100">
              <button
                onClick={() => {
                  onSelectSession('');
                  setActiveMenuId(null);
                }}
                className="w-full text-left text-xs text-slate-700 hover:bg-slate-50 py-2 px-3 font-semibold flex items-center gap-2"
              >
                <ArrowLeft className="w-3.5 h-3.5 text-slate-400" />
                <span>Mudar de Questionário</span>
              </button>
            </div>
          )}
        </header>
      ) : (
        <header className="p-4 bg-white border-b border-slate-150 flex items-center justify-between shadow-sm shrink-0">
          <span className="font-extrabold uppercase tracking-tight text-[#2d7f8d] text-[13px]">Votação PowerPoint</span>
        </header>
      )}

      {/* FLOAT SUCCESS ALERTS */}
      {insertSuccess && (
        <div className="absolute top-14 left-3 right-3 z-50 bg-[#e6f4ea] border border-[#a8dab5] rounded-xl p-3 shadow-lg flex items-center gap-2.5 animate-in fade-in slide-in-from-top-1.5 duration-200">
          <Check className="w-4 h-4 text-[#137333] shrink-0" />
          <p className="text-[11px] font-bold text-[#137333] leading-tight">{insertSuccess}</p>
        </div>
      )}

      {/* CORE FRAME LAYOUT */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        
        {!activeSession ? (
          
          /* VIEW 1: SELECIONAR / CRIAR QUESTIONÁRIO */
          <div className="space-y-4 animate-in fade-in duration-200">
            <form onSubmit={handleCreateSessionSubmit} className="p-4 bg-white rounded-2xl border border-slate-200/90 shadow-sm space-y-4">
              <h3 className="font-bold text-slate-800 text-[12px] uppercase tracking-wider font-mono">Criar Questionário</h3>
              
              <div className="space-y-1">
                <input
                  type="text"
                  required
                  value={newSessionName}
                  onChange={(e) => setNewSessionName(e.target.value)}
                  placeholder="Nome de sua aula ou painel..."
                  className="w-full text-xs bg-[#fdfdfd] border border-slate-200 hover:border-slate-300 rounded-xl p-3 focus:border-[#2d7f8d] focus:outline-none text-slate-800 font-semibold"
                />
              </div>

              <button
                type="submit"
                className="w-full bg-[#2d7f8d] hover:bg-[#23646f] text-white font-extrabold py-2.5 rounded-xl flex items-center justify-center gap-1.5 uppercase transition-all shadow-sm text-xs cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>CRIAR QUESTIONÁRIO</span>
              </button>
            </form>

            {/* SELECTION LIST */}
            {sessions.length > 0 && (
              <div className="space-y-2">
                <span className="font-semibold text-[10px] tracking-wider text-slate-400 font-mono uppercase pl-1">Escolher Questionário Existente:</span>
                <div className="space-y-2">
                  {sessions.map((sess) => (
                    <div
                      key={sess.id}
                      onClick={() => onSelectSession(sess.id)}
                      className="p-3 bg-white border border-slate-150-100/90 hover:border-[#2d7f8d]/60 rounded-xl transition-all cursor-pointer flex items-center justify-between shadow-sm hover:shadow"
                    >
                      <div className="truncate pr-2">
                        <h4 className="font-bold text-slate-800 text-xs truncate">{sess.name}</h4>
                        <span className="text-[10px] text-slate-400 font-mono">Código: {sess.id} • {sess.questions?.length || 0} questões</span>
                      </div>
                      <ChevronRight className="w-4 h-4 text-[#2d7f8d]" />
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
          
        ) : (
          
          /* VIEW 2: ACTIVE QUESTIONNAIRE GRID (LOOKS EXACTLY LIKE VEVOX PRESENTED IMAGE) */
          <div className="space-y-4 animate-in fade-in duration-200">
            
            {/* LARGE HEADER: Polls TITLE + ADD CONTENT BUTTON */}
            <div className="flex items-center justify-between pt-1">
              <h2 className="text-[20px] font-bold text-[#1a1a1a] tracking-tight">Polls</h2>
              {!isFormOpen && (
                <button
                  type="button"
                  onClick={() => setIsFormOpen(true)}
                  className="bg-[#2d7f8d] hover:bg-[#23646f] text-white text-[11px] font-bold py-1.5 px-3.5 rounded-full flex items-center gap-1.5 transition-all cursor-pointer shadow-sm uppercase tracking-wider"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>+ Add content</span>
                </button>
              )}
            </div>

            {/* INLINE QUESTION FORM CONTAINER */}
            {isFormOpen && (
              <form onSubmit={handleSaveQuestionSubmit} className="p-4 bg-white border border-slate-200 shadow-md rounded-2xl space-y-4 animate-in fade-in slide-in-from-top-1">
                <div className="flex items-center justify-between pb-1.5 border-b border-slate-100">
                  <h4 className="font-bold text-slate-800 text-[11.5px] uppercase font-mono text-[#2d7f8d]">
                    {editingQuestionId ? 'Editar Questão' : 'Adicionar Nova Questão'}
                  </h4>
                  <button type="button" onClick={handleClearForm} className="text-[10px] text-slate-400 hover:text-slate-600 font-bold uppercase transition-colors">
                    Cancelar
                  </button>
                </div>

                {/* SELECT QUESTION TYPE */}
                <div>
                  <label className="block text-[10px] font-semibold text-slate-500 mb-1.5 uppercase tracking-wider font-mono">TIPO DE PERGUNTA</label>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      type="button"
                      onClick={() => setQType('alternativa')}
                      className={`py-2 px-3 rounded-xl border text-[10.5px] font-bold uppercase transition-all cursor-pointer ${
                        qType === 'alternativa'
                          ? 'bg-[#e2f1f3] border-[#2d7f8d] text-[#2d7f8d]'
                          : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'
                      }`}
                    >
                      Múltipla Escolha
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setQType('aberta');
                        setQCorrectIdx(null);
                      }}
                      className={`py-2 px-3 rounded-xl border text-[10.5px] font-bold uppercase transition-all cursor-pointer ${
                        qType === 'aberta'
                          ? 'bg-[#e2f1f3] border-[#2d7f8d] text-[#2d7f8d]'
                          : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'
                      }`}
                    >
                      Pergunta Aberta
                    </button>
                  </div>
                </div>

                {/* enunciado */}
                <div className="space-y-1">
                  <label className="block text-[10px] font-semibold text-slate-500 uppercase tracking-wider font-mono">Enunciado da Pergunta</label>
                  <textarea
                    required
                    rows={2}
                    value={qText}
                    onChange={(e) => setQText(e.target.value)}
                    placeholder="Ex: Qual destas opções é correta sobre o fluxo de CRM?"
                    className="w-full text-xs font-semibold bg-[#fafafa] border border-slate-200 rounded-xl p-3 text-slate-800 placeholder-slate-400 focus:outline-none focus:border-[#2d7f8d] focus:ring-1 focus:ring-[#2d7f8d]"
                  />
                </div>

                {/* ALTERNATIVES */}
                {qType === 'alternativa' && (
                  <div className="space-y-2.5">
                    <label className="block text-[10px] font-semibold text-slate-500 uppercase tracking-wider font-mono">Alternativas e Gabarito</label>
                    <div className="space-y-1.5">
                      {qOptions.map((opt, idx) => {
                        const letChar = String.fromCharCode(65 + idx);
                        const isCorrect = qCorrectIdx === idx;
                        return (
                          <div key={idx} className="flex items-center gap-2">
                            <button
                              type="button"
                              onClick={() => setQCorrectIdx(isCorrect ? null : idx)}
                              className={`w-7 h-7 rounded-xl flex items-center justify-center font-mono text-[10.5px] font-black border transition-all cursor-pointer select-none ${
                                isCorrect
                                  ? 'bg-[#2d7f8d] text-white border-[#2d7f8d]'
                                  : 'bg-[#f8fafc] text-slate-400 border-slate-200 hover:border-slate-350'
                              }`}
                              title={isCorrect ? 'Alternativa Correta' : 'Marcar como certa'}
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
                              className="flex-1 text-xs bg-[#fafafa] border border-slate-200 rounded-xl p-2 font-semibold text-slate-700 focus:outline-none focus:border-[#2d7f8d]"
                            />
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                <div className="flex gap-2 pt-1 border-t border-slate-100">
                  <button
                    type="submit"
                    className="flex-1 bg-[#2d7f8d] hover:bg-[#23646f] text-white py-2.5 rounded-xl font-bold uppercase text-xs transition-colors cursor-pointer"
                  >
                    {editingQuestionId ? 'Salvar Alterações' : 'Criar Pergunta'}
                  </button>
                </div>
              </form>
            )}

            {/* VEVOX-STYLE INSTRUCTIONS BOX (AS SHOWN IN THE VEVOX PORTRAIT DESIGN GRAPHIC) */}
            <div className="p-3 bg-white border border-slate-200/90 rounded-2xl shadow-sm space-y-2">
              
              {/* ITEM 1: JOINING INSTRUCTIONS CONTAINER (ACCORDING TO PIC 1) */}
              <div className="flex items-center justify-between py-1 border-b border-slate-100/70 pb-2">
                <div className="flex items-center gap-3">
                  <div className="p-1 px-1.5 bg-[#eaf4f7] rounded-lg text-[#2d7f8d]">
                    <Grid className="w-4 h-4" />
                  </div>
                  <span className="font-bold text-slate-800 text-[12px]">Joining instructions</span>
                </div>
                <button
                  onClick={handleInsertQRCodeImage}
                  className="bg-white border border-[#2d7f8d]/60 hover:bg-[#eaf4f7] text-[#2d7f8d] font-bold py-1 px-3.5 rounded-lg text-[10.5px] transition-all cursor-pointer flex items-center justify-center min-w-[62px]"
                >
                  + ADD
                </button>
              </div>

              {/* ITEM 2: INDIVIDUAL LEADERBOARD WITH ADD (FOR PERFECT PIXELS LOOKS) */}
              <div className="flex items-center justify-between py-1 border-b border-slate-100/70 pb-2">
                <div className="flex items-center gap-3">
                  <div className="p-1 px-1.5 bg-slate-50 rounded-lg text-slate-500">
                    <Trophy className="w-4 h-4" />
                  </div>
                  <span className="font-semibold text-slate-600 text-[12px]">Individual Leaderboard</span>
                </div>
                <button
                  onClick={() => triggerNotification('Placar Individual inserido no slide atual!')}
                  className="bg-white border border-slate-250 hover:bg-slate-50 text-slate-500 font-bold py-1 px-3.5 rounded-lg text-[10.5px] transition-colors cursor-pointer"
                >
                  + ADD
                </button>
              </div>

              {/* ITEM 3: TEAM LEADERBOARD WITH ADD */}
              <div className="flex items-center justify-between py-1">
                <div className="flex items-center gap-3">
                  <div className="p-1 px-1.5 bg-slate-50 rounded-lg text-slate-500">
                    <Trophy className="w-4 h-4" />
                  </div>
                  <span className="font-semibold text-slate-600 text-[12px]">Team Leaderboard</span>
                </div>
                <button
                  onClick={() => triggerNotification('Placar de Equipes inserido no slide atual!')}
                  className="bg-white border border-slate-250 hover:bg-slate-50 text-slate-500 font-bold py-1 px-3.5 rounded-lg text-[10.5px] transition-colors cursor-pointer"
                >
                  + ADD
                </button>
              </div>

            </div>

            {/* SEPARATOR AND SELECT ALL BAR */}
            <div className="flex items-center gap-2 px-1 pt-1">
              <input 
                type="checkbox" 
                id="select-all-dummy" 
                className="accent-[#2d7f8d] w-3.5 h-3.5 cursor-pointer rounded border-slate-350" 
                defaultChecked 
              />
              <label htmlFor="select-all-dummy" className="text-slate-500 font-bold text-[11px] cursor-pointer">
                Select all
              </label>
            </div>

            {/* QUESTION POLLING LIST (DESIGN PRECIOSITY MATCHING THE PIC DETAILS) */}
            {activeSession.questions.length === 0 ? (
              <div className="text-center p-8 bg-white border border-dashed border-slate-200 rounded-2xl text-slate-400 font-medium">
                Nenhuma enquete adicionada a este questionário ainda.
              </div>
            ) : (
              <div className="space-y-4">
                {activeSession.questions.map((q, idx) => {
                  const isCurrentSelected = activeSession.currentQuestionIndex === idx;
                  const isDropdownOpen = activeMenuId === q.id;

                  return (
                    <div
                      key={q.id}
                      className={`bg-white border transition-all rounded-2xl shadow-sm p-3.5 space-y-3 relative ${
                        isCurrentSelected 
                          ? 'border-[#2d7f8d] ring-1 ring-[#2d7f8d]/30' 
                          : 'border-slate-200 hover:border-slate-300'
                      }`}
                      onClick={() => {
                        if (activeSession.currentQuestionIndex !== idx) {
                          syncStore.setQuestionIndex(activeSession.id, idx);
                          onRefresh();
                        }
                      }}
                    >
                      {/* TOP ROW: ICON + NUMBER + TEXT + OPTIONS TRIGGER */}
                      <div className="flex items-start justify-between relative">
                        <div className="flex items-start gap-2.5 truncate pr-1">
                          {/* Left hand side document list-like blue icon */}
                          <div className="p-1 px-1.5 bg-[#eaf4f7] rounded-lg text-[#2d7f8d] mt-0.5 shrink-0">
                            <FileSPREADSHEET_ICON q={q} />
                          </div>
                          
                          {/* Index question text descriptor exactly style of image */}
                          <div className="truncate leading-relaxed text-left">
                            <p className="text-[12px] font-bold text-slate-800 leading-tight block">
                              {idx + 1}. {q.text}
                            </p>
                          </div>
                        </div>

                        {/* Dropdown Options Icon Trigger */}
                        <div className="relative shrink-0">
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              setActiveMenuId(isDropdownOpen ? null : q.id);
                            }}
                            className="p-1 hover:bg-slate-50 rounded text-slate-400 hover:text-slate-700 transition"
                          >
                            <MoreVertical className="w-4 h-4" />
                          </button>

                          {/* ACTION MENUS FLOATING POPUP */}
                          {isDropdownOpen && (
                            <div className="absolute right-0 top-6 bg-white border border-slate-200 rounded-xl shadow-xl py-1 z-40 w-32 animate-in fade-in duration-75">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleStartEdit(q);
                                }}
                                className="w-full text-left text-[11px] text-slate-700 hover:bg-slate-50 py-1.5 px-3 font-semibold flex items-center gap-1.5"
                              >
                                <Edit2 className="w-3.5 h-3.5 text-slate-400" />
                                <span>Editar</span>
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteQuestion(q.id);
                                }}
                                className="w-full text-left text-[11px] text-red-650 hover:bg-red-50 py-1.5 px-3 font-semibold flex items-center gap-1.5 text-red-650"
                              >
                                <Trash2 className="w-3.5 h-3.5 text-red-400" />
                                <span>Excluir</span>
                              </button>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* DOWN ROW: COMPACT SLIDES INJECT BUTTON ROW AT THE BOTTOM OF CARD */}
                      <div className="flex items-center gap-1.5 pt-2 border-t border-slate-100 bg-white flex-wrap">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleInsertQuestionText(q);
                          }}
                          className="bg-white hover:bg-slate-50 border border-slate-200 py-1.5 px-2 rounded-lg text-[10px] font-bold text-slate-650 flex items-center gap-1 cursor-pointer transition"
                          title="Inserir texto da pergunta no Slide"
                        >
                          <Plus className="w-3 h-3 text-[#2d7f8d]" />
                          <span>Ins. Perg.</span>
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleInsertResponsesSlide(q);
                          }}
                          className="bg-white hover:bg-slate-50 border border-slate-200 py-1.5 px-2 rounded-lg text-[10px] font-bold text-slate-650 flex items-center gap-1 cursor-pointer transition"
                          title="Inserir slide de resultado de respostas"
                        >
                          <Plus className="w-3 h-3 text-[#2d7f8d]" />
                          <span>Ins. Resp.</span>
                        </button>

                        {/* Direct action buttons: Edit and delete */}
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleStartEdit(q);
                          }}
                          className="bg-[#fffbeb] hover:bg-[#fef3c7] text-[#b45309] border border-[#fde68a] py-1.5 px-2 rounded-lg text-[10px] font-bold flex items-center gap-1 cursor-pointer transition"
                          title="Editar/Modificar Pergunta"
                        >
                          <Edit2 className="w-3 h-3" />
                          <span>Editar</span>
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleDeleteQuestion(q.id);
                          }}
                          className="bg-red-50 hover:bg-red-100 border border-red-200 py-1.5 px-2 rounded-lg text-[10px] text-red-600 flex items-center gap-1 cursor-pointer transition"
                          title="Excluir Pergunta"
                        >
                          <Trash2 className="w-3 h-3 text-red-500" />
                        </button>

                        <div className="ml-auto text-[9.5px] font-mono text-slate-400">
                          {q.type === 'aberta' ? '💬 Aberta' : '❓ Alternativa'}
                        </div>
                      </div>

                    </div>
                  );
                })}
              </div>
            )}

            {/* ACTION: CLEAR ALL SESSION RESULTS AT FOOTER */}
            <div className="pt-3 flex justify-center pb-2">
              <button
                onClick={() => {
                  syncStore.toggleResults(activeSession.id, false);
                  triggerNotification('Slide de resultados limpo de forma bem-sucedida!');
                }}
                className="text-[#2d7f8d] hover:text-[#205b65] font-black tracking-wider hover:underline uppercase text-[10px] cursor-pointer"
              >
                Clear slide results
              </button>
            </div>

          </div>
        )}
      </div>

    </div>
  );
};

// Helper component inside the module for dynamic icon listing visual
const FileSPREADSHEET_ICON: React.FC<{ q: Question }> = ({ q }) => {
  return q.type === 'aberta' ? (
    <div className="font-bold text-[11px] font-mono leading-none flex items-center">✏️</div>
  ) : (
    <FileText className="w-3.5 h-3.5 text-[#2d7f8d]" />
  );
};
