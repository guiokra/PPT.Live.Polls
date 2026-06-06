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
  ArrowLeft
} from 'lucide-react';
import QRCode from 'qrcode';

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
  const [isOfficeAvailable, setIsOfficeAvailable] = useState<boolean>(false);
  const [insertSuccess, setInsertSuccess] = useState<string | null>(null);

  // Question Form State (inline & streamlined)
  const [qText, setQText] = useState<string>('');
  const [qType, setQType] = useState<'alternativa' | 'aberta'>('alternativa');
  const [qOptions, setQOptions] = useState<string[]>(['', '', '', '']);
  const [qCorrectIdx, setQCorrectIdx] = useState<number | null>(null);
  const [editingQuestionId, setEditingQuestionId] = useState<string | null>(null);

  // Session inline creator
  const [newSessionName, setNewSessionName] = useState<string>('');

  useEffect(() => {
    setSessions(syncStore.getSessions().filter(s => s.status === 'active'));

    if (typeof Office !== 'undefined') {
      Office.onReady((info: any) => {
        if (info.host === Office.HostType.PowerPoint) {
          setIsOfficeAvailable(true);
        }
      });
    }
  }, [activeSession]);

  const triggerNotification = (message: string) => {
    setInsertSuccess(message);
    setTimeout(() => setInsertSuccess(null), 3050);
  };

  const handleCreateSessionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSessionName.trim()) {
      triggerNotification('Digite o nome do questionário.');
      return;
    }
    const created = syncStore.createSession(newSessionName.trim(), false, false);
    setNewSessionName('');
    onSelectSession(created.id);
    triggerNotification('Questionário criado com sucesso!');
    onRefresh();
  };

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
              triggerNotification('QR Code de acesso inserido neste slide!');
            } else {
              triggerNotification('Erro ao inserir. Selecione um slide e tente novamente.');
            }
          }
        );
      } else {
        // Fallback simulated mapping (no file download as requested)
        triggerNotification('QR Code adicionado automaticamente ao slide virtual!');
        
        // Save simulated mapping so state is updated
        syncStore.saveSlideMapping(activeSession.id, 'slide-instructions', 'instructions', 'question');
        onRefresh();
      }
    } catch (_) {
      triggerNotification('Erro ao gerar o QR Code.');
    }
  };

  const handleInsertQuestionText = (q: Question) => {
    if (!activeSession) return;

    const optionsText = q.type === 'alternativa'
      ? q.options.map((o, idx) => `  ${String.fromCharCode(65 + idx)}) ${o}`).join('\n')
      : '  [ Resposta Aberta / Digite no celular ]';

    const slideText = `❓ PERGUNTA:\n${q.text}\n\n${optionsText}\n\n📱 Escaneie o QR Code inicial da apresentação para responder!`;

    if (isOfficeAvailable) {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            syncStore.saveSlideMapping(activeSession.id, slideId, q.id, 'question');
            
            Office.context.document.setSelectedDataAsync(
              slideText,
              { coercionType: Office.CoercionType.Text },
              (asyncResult: any) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  triggerNotification('Questão e Vínculo inseridos no slide!');
                  onRefresh();
                }
              }
            );
          }
        }
      );
    } else {
      const simulatedSlideId = `slide-q-${q.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, q.id, 'question');
      triggerNotification('Pergunta vinculada ao slide atual!');
      
      // Copy question helper block to clipboard quietly
      navigator.clipboard.writeText(slideText).catch(() => {});
      onRefresh();
    }
  };

  const handleInsertResponsesSlide = (q: Question) => {
    if (!activeSession) return;

    const dummyText = `📊 RESULTADOS AO VIVO:\n${q.text}\n\n[ Os votos de todos os alunos aparecerão aqui ao vivo ]`;

    if (isOfficeAvailable) {
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value.slides.length > 0) {
            const slideId = result.value.slides[0].id;
            syncStore.saveSlideMapping(activeSession.id, slideId, q.id, 'responses');
            
            Office.context.document.setSelectedDataAsync(
              dummyText,
              { coercionType: Office.CoercionType.Text },
              (asyncResult: any) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  triggerNotification('Respostas vinculadas a este slide!');
                  onRefresh();
                }
              }
            );
          }
        }
      );
    } else {
      const simulatedSlideId = `slide-r-${q.id}`;
      syncStore.saveSlideMapping(activeSession.id, simulatedSlideId, q.id, 'responses');
      triggerNotification('Gráfico de respostas vinculado ao slide atual!');
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
      triggerNotification('Dê pelo menos 2 alternativas.');
      return;
    }

    if (editingQuestionId) {
      syncStore.updateQuestionInSession(
        activeSession.id,
        editingQuestionId,
        qText.trim(),
        filteredOptions,
        qType === 'alternativa' ? qCorrectIdx : null
      );

      // Force type & metadata sync
      const allSessions = syncStore.getSessions();
      const s = allSessions.find(sess => sess.id === activeSession.id);
      if (s) {
        const targetQ = s.questions.find(quest => quest.id === editingQuestionId);
        if (targetQ) {
          targetQ.type = qType;
        }
        localStorage.setItem('polling_sessions', JSON.stringify(allSessions));
      }
      triggerNotification('Pergunta atualizada!');
    } else {
      const currentS = syncStore.getSession(activeSession.id);
      if (currentS) {
        const newId = `q-${Date.now()}`;
        const newQuestion: Question = {
          id: newId,
          text: qText.trim(),
          options: filteredOptions,
          correctOptionIndex: qType === 'alternativa' ? qCorrectIdx : null,
          type: qType,
          resultsView: 'results-slide-only'
        };
        currentS.questions.push(newQuestion);
        
        const allSessions = syncStore.getSessions();
        const foundIdx = allSessions.findIndex(sess => sess.id === activeSession.id);
        if (foundIdx >= 0) {
          allSessions[foundIdx] = currentS;
          localStorage.setItem('polling_sessions', JSON.stringify(allSessions));
        }
      }
      triggerNotification('Pergunta criada!');
    }

    // Reset fields
    setQText('');
    setQOptions(['', '', '', '']);
    setQCorrectIdx(null);
    setEditingQuestionId(null);
    onRefresh();
  };

  const handleStartEdit = (q: Question) => {
    setQText(q.text);
    setQType(q.type || 'alternativa');
    setQOptions(q.options && q.options.length ? [...q.options, '', '', ''].slice(0, 4) : ['', '', '', '']);
    setQCorrectIdx(q.correctOptionIndex);
    setEditingQuestionId(q.id);
  };

  const handleDeleteQuestion = (qId: string) => {
    if (!activeSession) return;
    syncStore.deleteQuestionFromSession(activeSession.id, qId);
    triggerNotification('Pergunta removida.');
    onRefresh();
  };

  const handleClearForm = () => {
    setQText('');
    setQOptions(['', '', '', '']);
    setQCorrectIdx(null);
    setEditingQuestionId(null);
  };

  return (
    <div className="bg-[#09090b] text-zinc-100 min-h-screen flex flex-col relative text-xs font-sans h-full overflow-hidden select-none">
      
      {/* COMPACT CLEAN HEADER */}
      <header className="p-3 border-b border-[#27272a] bg-[#121214] flex items-center justify-between">
        <span className="font-extrabold uppercase tracking-tight text-white text-xs">Votação PowerPoint</span>
        {activeSession && (
          <button
            onClick={() => onSelectSession('')}
            className="text-[10px] text-[#a3e635] hover:underline font-bold flex items-center gap-1"
          >
            <ArrowLeft className="w-3 h-3" /> Alterar Aula
          </button>
        )}
      </header>

      {/* FLASH SUCCESS DIALOG */}
      {insertSuccess && (
        <div className="absolute top-12 left-3 right-3 z-50 bg-slate-900 border border-[#a3e635]/40 rounded-xl p-2.5 shadow-2xl flex items-center gap-2 animate-in fade-in slide-in-from-top-1">
          <Check className="w-4 h-4 text-[#a3e635]" />
          <p className="text-[11px] font-bold text-zinc-200">{insertSuccess}</p>
        </div>
      )}

      <div className="flex-1 overflow-y-auto p-3 space-y-4">
        {!activeSession ? (
          
          /* VIEW 1: CREAR / SELECIONAR QUESTIONÁRIO */
          <div className="space-y-4">
            <form onSubmit={handleCreateSessionSubmit} className="p-3 bg-[#121214] rounded-xl border border-[#27272a] space-y-3">
              <h3 className="font-bold text-white text-xs uppercase tracking-wider font-mono">Criar Questionário</h3>
              
              <div className="space-y-1">
                <input
                  type="text"
                  required
                  value={newSessionName}
                  onChange={(e) => setNewSessionName(e.target.value)}
                  placeholder="Nome da aula / palestra..."
                  className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2.5 focus:border-[#a3e635] focus:outline-none text-zinc-200"
                />
              </div>

              <button
                type="submit"
                className="w-full bg-[#a3e635] hover:bg-[#a3e635]/90 text-zinc-950 font-black py-2 rounded-lg flex items-center justify-center gap-1.5 uppercase transition-all text-xs"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>CRIAR QUESTIONÁRIO</span>
              </button>
            </form>

            {/* SELECTION LIST */}
            {sessions.length > 0 && (
              <div className="space-y-2">
                <span className="font-mono text-[9px] tracking-wider text-zinc-500 uppercase">Selecionar Aula Existente:</span>
                <div className="space-y-1.5">
                  {sessions.map((sess) => (
                    <div
                      key={sess.id}
                      onClick={() => onSelectSession(sess.id)}
                      className="p-2.5 bg-[#121214] border border-[#27272a] hover:border-zinc-700 rounded-lg transition-all cursor-pointer flex items-center justify-between"
                    >
                      <div className="truncate pr-2">
                        <h4 className="font-bold text-white text-xs truncate">{sess.name}</h4>
                        <span className="text-[10px] text-zinc-500 font-mono">ID: {sess.id} • {sess.questions.length} perguntas</span>
                      </div>
                      <ChevronRight className="w-4 h-4 text-[#a3e635]" />
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
          
        ) : (
          
          /* VIEW 2: CONFIGURE ACTIVE QUESTIONNAIRE & QUESTIONS */
          <div className="space-y-4">
            <div className="p-3 bg-[#121214] border border-[#27272a] rounded-xl flex items-center justify-between">
              <div className="truncate pr-2">
                <span className="text-[9px] font-mono text-[#a3e635] uppercase font-bold">Questionário Ativo</span>
                <h4 className="font-extrabold text-white text-xs truncate">{activeSession.name}</h4>
              </div>
              <span className="text-[10.5px] font-bold bg-[#09090b] border border-zinc-800 p-1 px-2.5 rounded font-mono text-zinc-400">
                Código: {activeSession.id}
              </span>
            </div>

            {/* ADD QR CODE BUTTON */}
            <div className="p-3 bg-[#121214] border border-[#27272a] rounded-xl space-y-2">
              <button
                onClick={handleInsertQRCodeImage}
                className="w-full bg-[#1e1e24] hover:bg-zinc-800 text-zinc-150 p-2.5 rounded-lg border border-zinc-800 hover:border-zinc-700 font-bold flex items-center justify-center gap-2 transition-all cursor-pointer text-xs justify-items-center"
              >
                <QrCode className="w-4 h-4 text-[#a3e635] animate-pulse shrink-0" />
                <span className="font-mono uppercase tracking-wide">Adicionar QR Code ao Slide Atual</span>
              </button>
            </div>

            {/* QUESTION CREATION / EDITING CARD */}
            <form onSubmit={handleSaveQuestionSubmit} className="p-3.5 bg-[#121214] border border-zinc-800 rounded-xl space-y-3">
              <div className="flex items-center justify-between pb-1.5 border-b border-zinc-800">
                <h4 className="font-bold text-white text-xs uppercase tracking-wider font-mono">
                  {editingQuestionId ? 'Editar Pergunta' : 'Adicionar Pergunta'}
                </h4>
                {editingQuestionId && (
                  <button type="button" onClick={handleClearForm} className="text-[10px] text-zinc-500 hover:text-zinc-300">
                    Cancelar
                  </button>
                )}
              </div>

              {/* SELECT QUESTION TYPE */}
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 mb-1">TIPO DE PERGUNTA</label>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => setQType('alternativa')}
                    className={`py-2 px-3 rounded-lg border text-[10.5px] font-bold uppercase transition-all cursor-pointer ${
                      qType === 'alternativa'
                        ? 'bg-[#a3e635] border-[#a3e635] text-zinc-950 font-black'
                        : 'bg-[#09090b] border-zinc-800 text-zinc-455 hover:border-zinc-750 text-zinc-400'
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
                    className={`py-2 px-3 rounded-lg border text-[10.5px] font-bold uppercase transition-all cursor-pointer ${
                      qType === 'aberta'
                        ? 'bg-[#a3e635] border-[#a3e635] text-zinc-950 font-black'
                        : 'bg-[#09090b] border-zinc-800 text-zinc-455 hover:border-zinc-750 text-zinc-400'
                    }`}
                  >
                    Pergunta Aberta
                  </button>
                </div>
              </div>

              {/* QUESTION TEXT CHIP */}
              <div className="space-y-1">
                <label className="block text-[10px] font-mono text-zinc-500">ENUNCIADO</label>
                <textarea
                  required
                  rows={2}
                  value={qText}
                  onChange={(e) => setQText(e.target.value)}
                  placeholder="Escreva sua pergunta aqui..."
                  className="w-full text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-2 text-zinc-100 placeholder-zinc-700 focus:outline-[#a3e635] focus:outline-none"
                />
              </div>

              {/* OPTIONS CHIPS IF ALTERNATIVE */}
              {qType === 'alternativa' && (
                <div className="space-y-2">
                  <label className="block text-[10px] font-mono text-zinc-500">ALTERNATIVAS E MARCAÇÃO DE CERTA (GABARITO)</label>
                  <div className="space-y-1.5">
                    {qOptions.map((opt, idx) => {
                      const letChar = String.fromCharCode(65 + idx);
                      const isCorrect = qCorrectIdx === idx;
                      return (
                        <div key={idx} className="flex items-center gap-2">
                          <button
                            type="button"
                            onClick={() => setQCorrectIdx(isCorrect ? null : idx)}
                            className={`w-6 h-6 rounded flex items-center justify-center font-mono text-[10px] font-bold border transition-colors cursor-pointer ${
                              isCorrect
                                ? 'bg-[#a3e635] text-zinc-950 border-[#a3e635]'
                                : 'bg-[#09090b] text-zinc-550 border-zinc-800 hover:border-zinc-700'
                            }`}
                            title={isCorrect ? 'Resposta Certa' : 'Marcar resposta certa'}
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
                            className="flex-1 text-xs bg-[#09090b] border border-zinc-800 rounded-lg p-1.5 text-zinc-200 placeholder-zinc-700 focus:outline-none focus:border-[#a3e635]"
                          />
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <button
                type="submit"
                className="w-full bg-[#a3e635] hover:bg-[#a3e635]/95 text-zinc-950 py-2 rounded-lg font-black uppercase text-xs transition-colors cursor-pointer"
              >
                {editingQuestionId ? 'Salvar Pergunta' : 'Adicionar Pergunta'}
              </button>
            </form>

            {/* LIST OF QUESTIONS */}
            {activeSession.questions.length > 0 && (
              <div className="space-y-2.5">
                <span className="font-mono text-[9px] tracking-wider text-zinc-500 uppercase">Perguntas criadas ({activeSession.questions.length}):</span>
                <div className="space-y-2">
                  {activeSession.questions.map((q, idx) => {
                    const isCurrentActive = activeSession.currentQuestionIndex === idx;
                    return (
                      <div
                        key={q.id}
                        className={`p-3 rounded-xl border transition-all space-y-2.5 cursor-pointer ${
                          isCurrentActive
                            ? 'bg-zinc-900/60 border-zinc-700 text-white'
                            : 'bg-[#121214] border-zinc-900 text-zinc-400 hover:bg-zinc-900/20'
                        }`}
                        onClick={() => {
                          if (activeSession.currentQuestionIndex !== idx) {
                            syncStore.setQuestionIndex(activeSession.id, idx);
                            onRefresh();
                          }
                        }}
                      >
                        <div className="flex items-start justify-between">
                          <div className="flex items-center gap-1.5 truncate">
                            <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-zinc-800 text-[#a3e635]">
                              {idx + 1}
                            </span>
                            <span className="text-[10px] font-mono text-zinc-500">
                              {q.type === 'aberta' ? '💬 Aberta' : '❓ Alternativa'}
                            </span>
                            <span className="text-xs font-bold truncate block text-zinc-100">{q.text}</span>
                          </div>
                          
                          <div className="flex items-center gap-1 md:gap-2 shrink-0">
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleStartEdit(q);
                              }}
                              className="p-1 text-zinc-400 hover:text-white"
                              title="Editar enunciado"
                            >
                              <Edit2 className="w-3 h-3" />
                            </button>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleDeleteQuestion(q.id);
                              }}
                              className="p-1 text-zinc-550 hover:text-red-400"
                              title="Remover pergunta"
                            >
                              <Trash2 className="w-3 h-3" />
                            </button>
                          </div>
                        </div>

                        {/* SLIDES VINCULATION TRIGGERS */}
                        <div className="grid grid-cols-2 gap-1.5 pt-1.5 border-t border-zinc-800">
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleInsertQuestionText(q);
                            }}
                            className="bg-[#09090b] hover:bg-zinc-850 border border-zinc-800 hover:border-zinc-700 p-1.5 rounded text-[10px] font-mono font-bold text-zinc-300 flex items-center justify-center gap-1 transition-all"
                          >
                            <FileText className="w-3 h-3 text-[#a3e635]" />
                            <span>Slide Pergunta</span>
                          </button>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleInsertResponsesSlide(q);
                            }}
                            className="bg-[#09090b] hover:bg-zinc-850 border border-zinc-800 hover:border-zinc-700 p-1.5 rounded text-[10px] font-mono font-bold text-zinc-300 flex items-center justify-center gap-1 transition-all"
                          >
                            <Layers className="w-3 h-3 text-emerald-400" />
                            <span>Slide Respostas</span>
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* DISCREET NO-NOISE FOOTER */}
      <footer className="p-2 border-t border-[#27272a] bg-[#121214]/60 text-center text-[9px] text-zinc-650 font-mono tracking-wider uppercase font-medium">
        <span>Microsoft PowerPoint Polling Companion</span>
      </footer>
    </div>
  );
};
