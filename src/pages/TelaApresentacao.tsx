/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { syncStore, subscribeToSyncChanges } from '../services/store';
import { Session, Question, Vote, Participant } from '../types';
import { QRCodeDisplay } from '../components/QRCodeDisplay';
import { ResultChart } from '../components/ResultChart';
import { 
  QrCode, 
  Users, 
  HelpCircle, 
  Award, 
  TrendingUp, 
  Monitor, 
  Moon, 
  ChevronRight,
  Sparkles,
  Trophy,
  Activity,
  Heart
} from 'lucide-react';

interface TelaApresentacaoProps {
  sessionId: string;
  onSelectSession: (id: string) => void;
  appUrl: string;
}

export const TelaApresentacao: React.FC<TelaApresentacaoProps> = ({
  sessionId,
  onSelectSession,
  appUrl
}) => {
  const [session, setSession] = useState<Session | null>(null);
  const [votes, setVotes] = useState<Vote[]>([]);
  const [participants, setParticipants] = useState<Participant[]>([]);
  const [leaderboard, setLeaderboard] = useState<Participant[]>([]);
  const [localShowLeaderboard, setLocalShowLeaderboard] = useState<boolean>(false);

  useEffect(() => {
    // Initial load
    fetchState();

    // Subscribe to cross-tab updates (voter casting or instructor change questions!)
    const unsubscribe = subscribeToSyncChanges((event) => {
      fetchState();
    });

    return () => {
      unsubscribe();
    };
  }, [sessionId]);

  const fetchState = () => {
    const s = syncStore.getSession(sessionId);
    if (s) {
      setSession(s);
      const allVotes = syncStore.getVotesForSession(sessionId);
      
      // Filter votes for currently active question in session
      const activeQ = s.questions[s.currentQuestionIndex];
      const qVotes = activeQ ? allVotes.filter(v => v.questionId === activeQ.id) : [];
      setVotes(qVotes);
      
      const pts = syncStore.getParticipantsForSession(sessionId);
      setParticipants(pts);
      setLeaderboard(syncStore.getLeaderboard(sessionId));
    } else {
      setSession(null);
    }
  };

  if (!session) {
    return (
      <div className="bg-slate-950 text-slate-100 min-h-screen flex flex-col items-center justify-center p-4">
        <div className="text-center space-y-4 max-w-sm">
          <Moon className="w-12 h-12 text-slate-700 mx-auto animate-pulse" />
          <h2 className="font-extrabold text-lg text-slate-300">Sala de Apresentação Inativa</h2>
          <p className="text-xs text-slate-500">
            A sessão especificada ID "{sessionId}" não foi localizada ou já foi encerrada pelo instrutor.
          </p>
          <button
            onClick={() => onSelectSession('')}
            className="w-full mt-2 py-2 px-4 bg-slate-900 border border-slate-800 hover:bg-slate-850 rounded-lg text-xs font-semibold text-slate-300 cursor-pointer"
          >
            Voltar para Início
          </button>
        </div>
      </div>
    );
  }

  const activeQuestion: Question | null = session.questions[session.currentQuestionIndex] || null;
  const totalVotes = votes.length;
  const joinUrl = `${appUrl}?role=participant&session=${session.id}`;

  return (
    <div className="bg-[#09090b] text-zinc-100 min-h-screen flex flex-col relative select-none font-sans overflow-hidden">
      {/* Decorative subtle header background grids */}
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_20%_40%_at_50%_0%,rgba(163,230,53,0.03),rgba(0,0,0,0))]" />
      
      {/* SCREEN HEADER */}
      <header className="p-4 md:p-6 border-b border-[#27272a] flex items-center justify-between z-10 sticky top-0 bg-[#09090b]/85 backdrop-blur-md">
        <div className="flex items-center gap-3.5 min-w-0">
          <div className="w-3 h-3 bg-[#a3e635] rounded-full animate-ping shrink-0" />
          <div className="min-w-0">
            <h1 className="text-sm md:text-base font-black text-zinc-100 truncate tracking-tight">{session.name}</h1>
            <p className="text-[10px] md:text-xs text-zinc-500 font-medium font-mono hidden sm:block">
              PowerPoint Live Polling • {participants.length} participantes online
            </p>
          </div>
        </div>

        {/* Action Toggles for Projector view directly */}
        <div className="flex items-center gap-2">
          {session.isQuizMode && (
            <button
              onClick={() => setLocalShowLeaderboard(!localShowLeaderboard)}
              className={`py-1.5 px-3 rounded-lg text-xs font-semibold flex items-center gap-1 border transition-all cursor-pointer ${
                localShowLeaderboard 
                  ? 'bg-[#a3e635] border-[#a3e635] text-zinc-950 shadow' 
                  : 'bg-[#121214] border-[#27272a] text-zinc-400 hover:text-zinc-200'
              }`}
            >
              <Trophy className={`w-3.5 h-3.5 ${localShowLeaderboard ? 'text-zinc-950 animate-bounce' : 'text-zinc-500'}`} />
              <span>{localShowLeaderboard ? 'Exibir Pergunta' : 'Ver Classificação (XP)'}</span>
            </button>
          )}

          <div className="bg-[#121214] border border-[#27272a] px-3 py-1 rounded-lg font-mono text-xs md:text-sm font-black text-[#a3e635] select-all shrink-0">
            CÓDIGO: {session.id}
          </div>
        </div>
      </header>

      {/* CORE BODY GRID */}
      <main className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-6 p-4 md:p-6 items-stretch z-10">
        
        {/* LEADERBOARD TRANSITION OVERLAY */}
        {localShowLeaderboard && session.isQuizMode ? (
          <div className="lg:col-span-9 flex flex-col justify-between p-6 bg-[#121214]/60 border border-[#27272a] rounded-2xl relative shadow-2xl animate-in fade-in zoom-in-95 duration-400 h-full">
            <div className="absolute top-0 right-0 p-4 opacity-5 pointer-events-none">
              <Trophy className="w-64 h-64 text-[#a3e635]" />
            </div>

            <div className="space-y-4">
              <div className="flex items-center gap-2 pb-3 border-b border-[#27272a]">
                <Trophy className="w-6 h-6 text-[#a3e635] animate-bounce shrink-0" />
                <div>
                  <h2 className="text-lg font-black tracking-tight text-white">Leaderboard da Sessão</h2>
                  <p className="text-xs text-zinc-400">Professores e alunos visualizando os acertos em tempo real</p>
                </div>
              </div>

              {leaderboard.length === 0 ? (
                <div className="py-20 text-center text-zinc-500 text-xs font-mono">
                  Aguardando contagem de acertos dos participantes...
                </div>
              ) : (
                <div className="space-y-2.5 max-w-3xl mx-auto w-full pt-2">
                  {leaderboard.map((p, idx) => {
                    // Color code ranks
                    const isGold = idx === 0;
                    const isSilver = idx === 1;
                    const isBronze = idx === 2;
                    let rankBg = 'bg-zinc-800 text-zinc-450 font-bold';
                    let textHighlight = 'text-zinc-300';
                    
                    if (isGold) {
                      rankBg = 'bg-[#a3e635] text-zinc-950 font-black scale-105';
                      textHighlight = 'text-[#a3e635] font-extrabold text-sm';
                    } else if (isSilver) {
                      rankBg = 'bg-zinc-300 text-zinc-950 font-extrabold scale-102';
                      textHighlight = 'text-zinc-100 font-bold';
                    } else if (isBronze) {
                      rankBg = 'bg-amber-700 text-white font-extrabold';
                      textHighlight = 'text-amber-100 font-bold';
                    }

                    return (
                      <div 
                        key={p.id} 
                        className={`flex items-center justify-between p-2.5 rounded-xl border border-[#27272a] bg-[#121214]/90 transition-transform ${
                          isGold ? 'ring-1 ring-[#a3e635]/25 shadow-xl' : ''
                        }`}
                      >
                        <div className="flex items-center gap-3 min-w-0">
                          <span className={`w-6 h-6 rounded-lg font-mono text-[11px] flex items-center justify-center shrink-0 ${rankBg}`}>
                            {idx + 1}
                          </span>
                          <span className={`truncate ${textHighlight}`}>{p.name}</span>
                        </div>
                        <span className="font-mono text-xs font-bold text-[#a3e635] shrink-0 bg-zinc-950 px-2 py-0.5 rounded border border-zinc-880">
                          {p.score || 0} XP
                        </span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="text-center pt-4 border-t border-[#27272a] text-[11px] text-zinc-500 flex items-center justify-center gap-1 font-medium font-mono">
              <Sparkles className="w-3.5 h-3.5 text-[#a3e635]" />
              <span>Guerreie no quiz! Responda rápido para acumular mais XP.</span>
            </div>
          </div>
        ) : (
          /* STANDARD ACTIVE QUESTION VIEW */
          <div className="lg:col-span-9 flex flex-col justify-between p-6 bg-[#121214]/60 border border-[#27272a] rounded-3xl relative shadow-2xl min-h-[480px]">
            {!activeQuestion ? (
              <div className="my-auto text-center py-20 text-zinc-500 text-xs flex flex-col items-center gap-2">
                <HelpCircle className="w-8 h-8 opacity-40 animate-pulse text-[#a3e635]" />
                <p>Nenhuma enquete ativa no momento.</p>
                <p className="text-[11px] text-zinc-650">Aguardando o instrutor iniciar ou selecionar uma alternativa para responder.</p>
              </div>
            ) : (
              <>
                <div className="space-y-4 flex-1 flex flex-col justify-center">
                  <div className="space-y-2">
                    <span className="text-[10px] uppercase font-mono tracking-widest text-[#a3e635] font-bold block">
                      Slide {session.currentQuestionIndex + 1} de {session.questions.length}
                    </span>
                    <h2 className="text-xl md:text-2xl font-extrabold text-[#f4f4f5] leading-tight tracking-tight max-w-4xl font-display">
                      {activeQuestion.text}
                    </h2>
                  </div>

                  {/* REAL-TIME RESULTS CHART OR HOLDING SCREEN */}
                  <div className="pt-6 w-full max-w-3xl">
                    {session.showResults ? (
                      activeQuestion.type === 'aberta' ? (
                        <div className="space-y-4 max-h-[350px] overflow-y-auto pr-2">
                          <span className="text-[10px] text-[#a3e635] font-mono block tracking-wider uppercase">Repostas Enviadas:</span>
                          {votes.length === 0 ? (
                            <div className="text-center p-8 text-zinc-550 border border-[#27272a] rounded-2xl bg-[#09090b]/20 font-mono text-xs">
                              Nenhuma resposta de texto enviada pelos alunos ainda.
                            </div>
                          ) : (
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                              {votes.map((v, vidx) => (
                                <div key={vidx} className="p-3.5 rounded-xl bg-[#121214] border border-[#27272a] animate-in fade-in slide-in-from-bottom-2">
                                  <div className="flex items-center justify-between text-[10px] font-mono text-zinc-550 mb-1.5">
                                    <span className="font-bold text-[#a3e635]">{v.participantName}</span>
                                    <span>{new Date(v.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</span>
                                  </div>
                                  <p className="text-xs text-zinc-150 leading-relaxed font-semibold">"{v.textResponse || 'Resposta enviada'}"</p>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ) : (
                        <div className="p-4 rounded-2xl bg-[#09090b]/40 border border-[#27272a] animate-in fade-in duration-300">
                          <ResultChart
                            options={activeQuestion.options}
                            votes={votes}
                            correctOptionIndex={activeQuestion.correctOptionIndex}
                            revealAnswer={session.revealAnswer}
                          />
                        </div>
                      )
                    ) : (
                      /* BLIND SCREEN (VOTING ACTIVE BUT RESULTS BLOCKED TO REDUCE BIAS) */
                      <div className="p-8 text-center rounded-2xl bg-[#09090b]/60 border border-dashed border-[#27272a] max-w-2xl mx-auto space-y-3 flex flex-col items-center justify-center">
                        <div className="relative w-12 h-12 flex items-center justify-center">
                          <Activity className="w-10 h-10 text-[#a3e635] animate-pulse shrink-0" />
                          <Users className="w-5 h-5 text-zinc-400 absolute shrink-0" />
                        </div>
                        <div>
                          <p className="text-zinc-300 font-bold text-sm">Votação ativa!</p>
                          <p className="text-[11px] text-zinc-550 mt-0.5 leading-relaxed font-mono">
                            Responda pelo celular. Os resultados parciais estão ocultos para dar igualdade de escolha a todos.
                          </p>
                        </div>

                        {/* Participant response speed scale info */}
                        <div className="bg-[#121214]/80 px-3 py-1.5 rounded-lg border border-[#27272a] font-mono text-xs text-[#a3e635] flex items-center gap-1.5">
                          <Users className="w-4 h-4 text-zinc-500" />
                          <span>{totalVotes} {totalVotes === 1 ? 'voto enviado' : 'votos enviados'}</span>
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {/* BOTTOM DISPLAY STATUS */}
                <div className="flex items-center justify-between border-t border-[#27272a] pt-4 mt-6 text-zinc-500 font-medium text-[11px] uppercase tracking-wide font-mono">
                  <span className="flex items-center gap-1.5">
                    <Users className="w-4 h-4 text-zinc-650" />
                    {totalVotes} / {participants.length} alunos inscritos
                  </span>
                  <span>
                    {session.isQuizMode ? 'Modo Quiz Competitivo' : 'Modo Enquete Livre'}
                  </span>
                </div>
              </>
            )}
          </div>
        )}

        {/* RIGHT COLUMN: PERMANENT ACCESS INSTRUCTIONS FOR STUDENTS */}
        <section className="col-span-1 lg:col-span-3 flex flex-col justify-between gap-4">
          {/* Main Join QR code card */}
          <div className="p-4 bg-[#121214]/60 border border-[#27272a] rounded-3xl text-center space-y-4 shadow flex flex-col items-center justify-center h-full">
            <h3 className="font-extrabold text-xs uppercase tracking-wider text-[#a3e635] flex items-center gap-1 bg-[#09090b]/80 border border-[#27272a] rounded px-2.5 py-1">
              <QrCode className="w-3.5 h-3.5 text-[#a3e635] animate-pulse shrink-0" />
              Escaneie e Vote!
            </h3>

            {/* QR Card displaying real joinUrl */}
            <div className="p-2.5 bg-white rounded-2xl">
              <QRCodeDisplay value={joinUrl} size={150} showCopyButton={false} />
            </div>

            <div className="space-y-1 w-full">
              <p className="text-xs text-zinc-400">Acesse no celular:</p>
              <div className="bg-[#09090b]/80 p-2.5 rounded-xl border border-[#27272a] font-mono text-[11px] text-zinc-300 break-all select-all font-semibold">
                {appUrl}
              </div>
              <p className="text-[10px] text-zinc-500 font-bold mt-1.5 font-mono">
                Digite o Código: <span className="text-[#a3e635] font-black">{session.id}</span>
              </p>
            </div>
          </div>
        </section>
      </main>

      {/* FLOAT BRAND FOOTER CREDIT */}
      <footer className="p-3 border-t border-[#27272a] bg-[#121214]/40 text-center text-[10px] text-zinc-600 font-mono tracking-wide z-10 font-medium mt-auto flex items-center justify-center gap-1">
        <span>POLLSYNC PRO • Microsoft PowerPoint Companion</span>
      </footer>
    </div>
  );
};
