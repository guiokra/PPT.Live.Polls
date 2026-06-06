/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { syncStore, subscribeToSyncChanges } from '../services/store';
import { Session, Question, Participant, Vote } from '../types';
import { 
  User, 
  Key, 
  Check, 
  Lock, 
  HelpCircle, 
  Loader2, 
  Vote as VoteIcon, 
  TrendingUp, 
  Clock, 
  UserCheck, 
  Moon,
  Hash,
  AlertCircle,
  HelpCircle as HelpIcon,
  ChevronRight
} from 'lucide-react';
import { motion } from 'motion/react';

interface TelaParticipanteProps {
  initialSessionId: string;
  onRefresh: () => void;
}

export const TelaParticipante: React.FC<TelaParticipanteProps> = ({
  initialSessionId,
  onRefresh
}) => {
  // Navigation role stage: 'join_session' | 'voting_panel'
  const [stage, setStage] = useState<'join_session' | 'voting_panel'>('join_session');
  
  // States
  const [sessionCode, setSessionCode] = useState<string>(initialSessionId || '');
  const [participantName, setParticipantName] = useState<string>('');
  const [session, setSession] = useState<Session | null>(null);
  const [me, setMe] = useState<Participant | null>(null);
  const [activeQuestion, setActiveQuestion] = useState<Question | null>(null);
  
  // Voting
  const [selectedIdx, setSelectedIdx] = useState<number | null>(null);
  const [openText, setOpenText] = useState<string>('');
  const [hasVotedCurrent, setHasVotedCurrent] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // Recovery of local participant on load
  useEffect(() => {
    if (initialSessionId) {
      setSessionCode(initialSessionId);
      const activeS = syncStore.getSession(initialSessionId);
      if (activeS) {
        setSession(activeS);
        // Attempt restoring participant from local storage matching this session
        const storedMe = localStorage.getItem(`ppt_myself_${initialSessionId}`);
        if (storedMe) {
          try {
            const parsed = JSON.parse(storedMe);
            setMe(parsed);
            setStage('voting_panel');
          } catch (e) {
            console.error('Error parsing stored student data:', e);
            localStorage.removeItem(`ppt_myself_${initialSessionId}`);
          }
        }
      }
    }
  }, [initialSessionId]);

  // Main active question loader & subscription
  useEffect(() => {
    const activeSessionId = session?.id || initialSessionId;
    if (!activeSessionId) return;

    const fetchSession = () => {
      const activeS = syncStore.getSession(activeSessionId);
      if (activeS) {
        setSession(activeS);
      }
    };

    fetchSession();

    // Subscribe to changes in host (moving question slide, projected states, or showing results)
    const unsubscribe = subscribeToSyncChanges((evt) => {
      fetchSession();
    });

    return () => {
      unsubscribe();
    };
  }, [session?.id, initialSessionId]);

  // Handle question layout swap or projected change
  useEffect(() => {
    if (!session) return;
    loadActiveQuestionState();
  }, [session?.currentQuestionIndex, session?.projectedQuestionId, session?.id]);

  const loadActiveQuestionState = () => {
    if (!session) return;
    
    const activeQ = session.questions[session.currentQuestionIndex] || null;
    setActiveQuestion(activeQ);
    
    if (activeQ && me) {
      // Check if participant has already voted on this question
      const listVotes = syncStore.getVotesForQuestion(session.id, activeQ.id);
      const myVote = listVotes.find(v => v.participantId === me.id);
      
      if (myVote) {
        setSelectedIdx(myVote.selectedOptionIndex);
        setOpenText(myVote.textResponse || '');
        setHasVotedCurrent(true);
      } else {
        setSelectedIdx(null);
        setOpenText('');
        setHasVotedCurrent(false);
      }
    }
  };

  // JOINS FLOW
  const handleJoinSessionSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg(null);
    setLoading(true);

    const cleanCode = sessionCode.trim();
    if (!cleanCode) {
      setErrorMsg('Por favor, digite o código da sessão.');
      setLoading(false);
      return;
    }

    const targetSession = syncStore.getSession(cleanCode);
    if (!targetSession) {
      setErrorMsg('Sessão não econtrada. Verifique se digitou o código de 4 dígitos corretamente.');
      setLoading(false);
      return;
    }

    if (targetSession.status === 'ended') {
      setErrorMsg('Esta sessão já foi encerrada pelo apresentador.');
      setLoading(false);
      return;
    }

    // Name validation if not anonymous
    const cleanName = participantName.trim();
    if (!targetSession.isAnonymous && !cleanName) {
      setErrorMsg('Por favor, digite seu nome ou iniciais para entrar.');
      setLoading(false);
      return;
    }

    // Register
    const participant = syncStore.joinOrCreateParticipant(cleanCode, cleanName);
    
    // Save locally
    localStorage.setItem(`ppt_myself_${cleanCode}`, JSON.stringify(participant));
    
    setMe(participant);
    setSession(targetSession);
    setStage('voting_panel');
    setLoading(false);
  };

  // CASTS VOTE FLOW
  const handleOptionSelect = (idx: number) => {
    if (hasVotedCurrent) return; // Prevent editing after submit
    setSelectedIdx(idx);
  };

  const handleCastVoteSubmit = () => {
    if (!session || !activeQuestion || !me || hasVotedCurrent) return;

    const isAberta = activeQuestion.type === 'aberta';
    if (isAberta && openText.trim() === '') return;
    if (!isAberta && selectedIdx === null) return;

    setLoading(true);
    const success = syncStore.castVote(
      session.id,
      activeQuestion.id,
      me.id,
      me.name,
      isAberta ? -1 : selectedIdx!,
      isAberta ? openText.trim() : undefined
    );

    if (success) {
      setHasVotedCurrent(true);
      // Update local cache XP score
      if (session.isQuizMode && !isAberta && selectedIdx === activeQuestion.correctOptionIndex) {
        const copyMe = { ...me, score: me.score + 100 };
        setMe(copyMe);
        localStorage.setItem(`ppt_myself_${session.id}`, JSON.stringify(copyMe));
      }
    } else {
      setErrorMsg('Não foi possível registrar seu voto.');
    }
    setLoading(false);
  };

  const handleExitRegister = () => {
    if (session) {
      localStorage.removeItem(`ppt_myself_${session.id}`);
    }
    setMe(null);
    setSession(null);
    setStage('join_session');
    setSelectedIdx(null);
    setHasVotedCurrent(false);
  };

  const isQuestionProjected = (() => {
    if (!session || !activeQuestion) return false;
    
    // Check if the session has slide mappings. If so, they are running in PPT mode.
    const hasMappings = session.slideMappings && session.slideMappings.length > 0;
    if (!hasMappings) {
      return true; // No mappings yet, default to projected so it doesn't block demoing without mappings
    }
    
    // In PPT mode, must be explicitly matching the projected question ID
    return session.projectedQuestionId === activeQuestion.id;
  })();

  return (
    <div className="bg-[#09090b] text-zinc-100 min-h-screen py-4 px-4 flex flex-col justify-between items-center w-full select-none max-w-md mx-auto relative overflow-hidden font-sans">
      
      {/* BACKGROUND GRAPHIC FOR MOBILE FEEL */}
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_25%_45%_at_50%_0%,rgba(163,230,53,0.03),rgba(0,0,0,0))]" />

      {/* STAGE 1: GATE JOINS FOR STUDENTS */}
      {stage === 'join_session' ? (
        <div className="w-full my-auto space-y-6 z-10 p-2">
          {/* Brand/Icons info */}
          <div className="text-center space-y-2">
            <div className="w-12 h-12 bg-[#a3e635]/10 border border-[#a3e635]/20 text-[#a3e635] rounded-full flex items-center justify-center mx-auto mb-2.5 animate-bounce">
              <VoteIcon className="w-6 h-6 shrink-0" />
            </div>
            <h2 className="text-lg font-black tracking-tight text-white font-display">POLLSYNC PRO</h2>
            <p className="text-xs text-zinc-400">Entre na sessão para participar e votar pelo celular em tempo real.</p>
          </div>

          <form onSubmit={handleJoinSessionSubmit} className="space-y-4">
            {errorMsg && (
              <div className="p-3 bg-red-950/45 border border-red-900/40 rounded-xl text-xs text-red-350 flex items-start gap-2.5">
                <AlertCircle className="w-4.5 h-4.5 text-red-400 shrink-0 mt-0.5" />
                <span className="font-semibold">{errorMsg}</span>
              </div>
            )}

            <div className="space-y-3">
              <div>
                <label className="block text-[11px] font-mono uppercase tracking-wider text-zinc-500 mb-1">Código da Sessão (ID Pin)</label>
                <div className="relative">
                  <Hash className="w-4 h-4 text-zinc-600 absolute left-3.5 top-3.5" />
                  <input
                    type="tel"
                    maxLength={6}
                    value={sessionCode}
                    onChange={(e) => setSessionCode(e.target.value.replace(/\D/g, ''))}
                    placeholder="Ex: 8825"
                    className="w-full text-sm font-black bg-[#121214] border border-[#27272a] rounded-xl p-3 pl-10 text-[#a3e635] placeholder-zinc-700 focus:border-[#a3e635] focus:outline-none focus:ring-1 focus:ring-[#a3e635]/20 tracking-widest text-[16px]"
                    id="input-voter-id"
                  />
                </div>
              </div>

              {/* Display name field dynamically based on active session parameters if already pre-filled */}
              <div>
                <label className="block text-[11px] font-mono uppercase tracking-wider text-zinc-500 mb-1">Seu Nome / Matrícula</label>
                <div className="relative">
                  <User className="w-4 h-4 text-zinc-600 absolute left-3.5 top-3.5" />
                  <input
                    type="text"
                    maxLength={32}
                    value={participantName}
                    onChange={(e) => setParticipantName(e.target.value)}
                    placeholder="Ex: Alvaro Lima"
                    className="w-full text-sm font-semibold bg-[#121214] border border-[#27272a] rounded-xl p-3 pl-10 text-zinc-100 placeholder-zinc-700 focus:border-[#a3e635] focus:outline-none focus:ring-1 focus:ring-[#a3e635]/20"
                    id="input-voter-name"
                  />
                </div>
                <p className="text-[10px] text-zinc-500 mt-1">Sua identificação é usada para calcular as notas do ranking de alunos.</p>
              </div>
            </div>

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-[#a3e635] hover:bg-[#a3e635]/95 disabled:opacity-50 text-zinc-950 font-black py-3.5 px-4 rounded-xl flex items-center justify-center gap-1.5 transition-all text-sm shadow cursor-pointer"
              id="btn-voter-entry"
            >
              {loading ? (
                <>
                  <Loader2 className="w-4.5 h-4.5 animate-spin shrink-0" />
                  <span>Entrando na sala...</span>
                </>
              ) : (
                <>
                  <span>Participar da Apresentação</span>
                  <ChevronRight className="w-4 h-4 text-zinc-950 shrink-0" />
                </>
              )}
            </button>
          </form>
          
          <div className="text-center p-3 border border-zinc-800/60 rounded-xl bg--pane/40 text-[10px] text-zinc-400 font-mono">
            A participação é instantânea! Não é necessário criar senha ou registrar emails.
          </div>
        </div>
      ) : (
        /* STAGE 2: VOTING INTERFACE */
        <div className="w-full flex-1 flex flex-col justify-between z-10 p-2">
          
          {/* Participant status header Bar */}
          <div className="flex items-center justify-between pb-3.5 border-b border-[#27272a] font-mono text-[11px]">
            <div className="flex items-center gap-1.5">
              <UserCheck className="w-4 h-4 text-[#a3e635]" />
              <span className="font-bold text-zinc-200 truncate max-w-[120px]">{me?.name}</span>
            </div>
            <div className="flex items-center gap-2">
              {session?.isQuizMode && (
                <span className="bg-[#121214] text-[#a3e635] border border-[#27272a] px-2 py-0.5 rounded font-bold">
                  {me?.score || 0} XP
                </span>
              )}
              <button
                onClick={handleExitRegister}
                className="text-[10px] text-zinc-500 hover:text-zinc-300 font-bold"
              >
                Sair
              </button>
            </div>
          </div>

          {!activeQuestion || !isQuestionProjected ? (
            /* HOLDING SCREEN (STAND-BY / NO QUESTION PROJECTED YET) */
            <div className="my-auto text-center space-y-5 py-12 flex flex-col items-center z-10 p-4">
              <div className="relative">
                <div className="w-16 h-16 bg-[#a3e635]/10 border border-[#a3e635]/20 rounded-full flex items-center justify-center animate-[pulse_2s_infinite]">
                  <Moon className="w-7 h-7 text-[#a3e635]" />
                </div>
                <div className="absolute -bottom-1 -right-1 bg-zinc-950 p-1 rounded-full border border-zinc-800">
                  <Loader2 className="w-4 h-4 text-zinc-500 animate-spin" />
                </div>
              </div>
              <div className="space-y-2">
                <h3 className="font-extrabold text-white text-base tracking-tight font-display">Conectado na Sessão • Em Stand-by</h3>
                <p className="text-xs text-zinc-400 max-w-xs mx-auto leading-relaxed">
                  Aguardando o palestrante projetar o slide desta pergunta no PowerPoint.
                </p>
                <div className="mt-4 p-2.5 bg-[#121214] border border-zinc-800 rounded-xl text-[11px] text-zinc-500 font-mono inline-block">
                  A página atualizará automaticamente em tempo real!
                </div>
              </div>
            </div>
          ) : (
            /* ACTIVE VOTING VIEW */
            <div className="my-auto space-y-6 animate-in fade-in zoom-in-95 duration-200">
              <div className="space-y-1">
                <span className="text-[9px] uppercase tracking-wider font-mono text-zinc-500 font-extrabold">Slide Atual da Aula</span>
                <h1 className="text-sm md:text-base font-black text-white leading-snug font-display">{activeQuestion.text}</h1>
              </div>

              {hasVotedCurrent ? (
                /* SCREEN SHOWING SUCCESS VOTE CASTRATION */
                <div className="p-6 rounded-2xl bg-[#121214] border border-[#27272a] space-y-3.5 text-center flex flex-col items-center justify-center animate-in zoom-in-95 duration-400">
                  <div className="w-12 h-12 rounded-full bg-[#a3e635]/10 border border-[#a3e635]/30 flex items-center justify-center text-[#a3e635]">
                    <Check className="w-6 h-6 shrink-0 animate-bounce font-black" />
                  </div>
                  <div>
                    <h3 className="font-bold text-[#a3e635] text-sm font-display">Voto Confirmado com Sucesso!</h3>
                    <p className="text-[11px] text-zinc-400 leading-relaxed mt-1">
                      Sua escolha foi computada nos gráficos em tempo real da apresentação.
                    </p>
                  </div>
                  
                  {session.isQuizMode && activeQuestion.type !== 'aberta' && activeQuestion.correctOptionIndex !== null && (
                    <div className="p-2 w-full bg-[#09090b] rounded-lg text-[10px] font-mono text-zinc-500 mt-1 border border-[#27272a]">
                      {selectedIdx === activeQuestion.correctOptionIndex ? (
                        <span className="text-[#a3e635] font-bold flex items-center justify-center gap-1">
                          🔥 Resposta Correta! Acumulou +100 XP
                        </span>
                      ) : (
                        <span className="text-zinc-500">Verifique a resposta certa na tela do projetor.</span>
                      )}
                    </div>
                  )}

                  <div className="pt-2 flex items-center gap-1.5 font-mono text-[9px] text-[#a3e635] uppercase tracking-widest font-black">
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    <span>Aguardando Próxima Pergunta...</span>
                  </div>
                </div>
              ) : activeQuestion.type === 'aberta' ? (
                /* OPEN ENDED CONTAINER */
                <div className="space-y-4">
                  <span className="text-[10px] text-zinc-400 block font-mono">DIGITE SUA RESPOSTA ABAIXO:</span>
                  <textarea
                    rows={4}
                    value={openText}
                    onChange={(e) => setOpenText(e.target.value)}
                    placeholder="Sua resposta..."
                    className="w-full text-xs font-semibold bg-[#121214] border border-[#27272a] rounded-xl p-3 focus:border-[#a3e635] focus:outline-none focus:ring-1 focus:ring-[#a3e635]/20 text-zinc-100 placeholder-zinc-700"
                    maxLength={150}
                  />
                  <button
                    onClick={handleCastVoteSubmit}
                    disabled={openText.trim() === '' || loading}
                    className="w-full bg-[#a3e635] hover:bg-[#a3e635]/90 text-zinc-950 font-black py-3 rounded-xl flex items-center justify-center gap-1.5 transition-all text-xs cursor-pointer shadow-md uppercase tracking-wider font-mono"
                  >
                    {loading ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      <>
                        <VoteIcon className="w-4 h-4 text-zinc-950" />
                        <span>Enviar Resposta</span>
                      </>
                    )}
                  </button>
                </div>
              ) : (
                /* FORM INPUT OPTIONS FOR VOTING */
                <div className="space-y-2.5">
                  {activeQuestion.options.map((opt, oIdx) => {
                    const letter = String.fromCharCode(65 + oIdx);
                    const isSelected = selectedIdx === oIdx;
                    
                    return (
                      <button
                        key={oIdx}
                        onClick={() => handleOptionSelect(oIdx)}
                        className={`w-full py-3.5 px-4 rounded-xl text-left font-medium text-xs flex items-center justify-between transition-all cursor-pointer ${
                          isSelected 
                            ? 'bg-[#a3e635] text-zinc-950 shadow-lg border border-[#a3e635] translate-x-1 font-bold' 
                            : 'bg-[#121214] border border-[#27272a] hover:bg-zinc-800 text-zinc-300'
                        }`}
                      >
                        <span className="truncate pr-2">
                          <strong className={`font-mono mr-2 text-xs p-1.5 rounded-lg border leading-none ${isSelected ? 'bg-zinc-950 text-[#a3e635] border-[#a3e635]' : 'bg-[#09090b] border-[#27272a] text-zinc-500'}`}>{letter}</strong>
                          {opt}
                        </span>
                        
                        {isSelected && (
                          <Check className="w-4 h-4 text-zinc-950 shrink-0 font-extrabold" />
                        )}
                      </button>
                    );
                  })}

                  <button
                    onClick={handleCastVoteSubmit}
                    disabled={selectedIdx === null || loading}
                    className="w-full mt-4 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:hover:bg-emerald-600 text-white font-bold py-3.5 rounded-xl flex items-center justify-center gap-1.5 transition-all text-xs cursor-pointer shadow-md"
                    id="btn-voter-submit"
                  >
                    {loading ? (
                      <Loader2 className="w-5 h-5 animate-spin" />
                    ) : (
                      <>
                        <VoteIcon className="w-4.5 h-4.5 text-emerald-200" />
                        <span>Confirmar e Enviar Resposta</span>
                      </>
                    )}
                  </button>
                </div>
              )}
            </div>
          )}

          {/* Spacer anchor */}
          <div className="text-center font-mono text-[9px] text-[#a3e635] font-bold uppercase tracking-widest pt-4">
            Código da Sala: {session?.id}
          </div>
        </div>
      )}

      {/* FOOTER BANNER */}
      <footer className="text-center text-[10px] text-zinc-600 font-mono mt-auto pt-4 border-t border-[#27272a] w-full">
        <span>POLLSYNC PRO © 2026</span>
      </footer>
    </div>
  );
};
