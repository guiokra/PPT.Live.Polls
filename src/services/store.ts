/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Question, Session, Vote, Participant, QuestionBankItem } from '../types';

// Key names for localStorage
const SESSIONS_KEY = 'ppt_live_sessions';
const VOTES_KEY = 'ppt_live_votes';
const PARTICIPANTS_KEY = 'ppt_live_participants';
const QUESTION_BANK_KEY = 'ppt_live_question_bank';
const CHANNEL_NAME = 'ppt_polling_broadcast_channel';

// Create BroadcastChannel for cross-tab synchronization
let broadcastChannel: BroadcastChannel | null = null;
try {
  broadcastChannel = new BroadcastChannel(CHANNEL_NAME);
} catch (e) {
  console.warn('BroadcastChannel not supported in this environment.', e);
}

// Initial template questions to populate the question bank and provide instant demo value
const DEFAULT_QUESTION_BANK: QuestionBankItem[] = [
  {
    id: 'qbank-1',
    text: 'Qual o principal fator de sucesso para uma apresentação profissional?',
    options: [
      'Visual limpo com pouco texto (menos é mais)',
      'Ler todos os tópicos do slide com calma',
      'Inserir o máximo de tabelas e dados possíveis',
      'Utilizar transições animadas complexas em cada slide'
    ],
    correctOptionIndex: 0,
    category: 'Apresentações',
    createdAt: new Date().toISOString()
  },
  {
    id: 'qbank-2',
    text: 'Selecione a opção que melhor define a regra de contraste de cores:',
    options: [
      'Texto cinza claro sobre fundo branco',
      'Textos profundos/escuros sobre fundos muito claros, ou vice-versa',
      'Texto azul brilhante sobre fundo verde escuro',
      'Fundo decorado com fotos complexas e texto colorido'
    ],
    correctOptionIndex: 1,
    category: 'Design Gráfico',
    createdAt: new Date().toISOString()
  },
  {
    id: 'qbank-3',
    text: 'No modo QUIZ, o que as enquetes ao vivo promovem primariamente no aprendizado?',
    options: [
      'Ansiedade e competição desnecessária',
      'Gamificação, fixação imediata do conteúdo e feedback rápido',
      'Apenas entretenimento sem valor pedagógico',
      'Substituição completa do instrutor pelo computador'
    ],
    correctOptionIndex: 1,
    category: 'Metodologias Ativas',
    createdAt: new Date().toISOString()
  },
  {
    id: 'qbank-4',
    text: 'A Regra 10-20-30 de Guy Kawasaki para apresentações recomenda quais limites?',
    options: [
      '10 slides, 20 minutos de duração, tamanho de fonte 30 no mínimo',
      '10 minutos, 20 slides, 30 imagens de alta resolução',
      '10 tópicos de foco, 20 slides de cabeçalho, 30 segundos cada',
      '10 cores, 20 transições, 30 minutos de perguntas e respostas'
    ],
    correctOptionIndex: 0,
    category: 'Dicas de Apresentação',
    createdAt: new Date().toISOString()
  }
];

// Seed initial default questions if question bank is empty
function seedQuestionBankIfNeeded(): void {
  try {
    const existing = localStorage.getItem(QUESTION_BANK_KEY);
    if (!existing || JSON.parse(existing).length === 0) {
      localStorage.setItem(QUESTION_BANK_KEY, JSON.stringify(DEFAULT_QUESTION_BANK));
    }
  } catch (e) {
    console.warn('Error reading or parsing question bank from storage, seeding defaults:', e);
    try {
      localStorage.setItem(QUESTION_BANK_KEY, JSON.stringify(DEFAULT_QUESTION_BANK));
    } catch (_) {}
  }
}

// Helper to notify other components/tabs
export function notifySyncChannel(actionType: string, payload?: any): void {
  if (broadcastChannel) {
    broadcastChannel.postMessage({ type: actionType, payload, timestamp: Date.now() });
  }
  // Dispatch local window structural events so current window React triggers re-render
  window.dispatchEvent(new CustomEvent('ppt_local_state_change', { detail: { type: actionType, payload } }));
}

// Subscribe to real-time triggers (from other tabs or same tab)
export function subscribeToSyncChanges(callback: (event: any) => void): () => void {
  const handleMessage = (e: MessageEvent) => {
    callback(e.data);
  };
  
  const handleLocal = (e: Event) => {
    const customE = e as CustomEvent;
    callback(customE.detail);
  };

  if (broadcastChannel) {
    broadcastChannel.addEventListener('message', handleMessage);
  }
  window.addEventListener('ppt_local_state_change', handleLocal);
  window.addEventListener('storage', handleLocal); // listen storage triggers

  return () => {
    if (broadcastChannel) {
      broadcastChannel.removeEventListener('message', handleMessage);
    }
    window.removeEventListener('ppt_local_state_change', handleLocal);
    window.removeEventListener('storage', handleLocal);
  };
}

// Initialise storage if empty
export const syncStore = {
  init(): void {
    seedQuestionBankIfNeeded();
    
    // Core sessions template if nothing exists
    const sessions = this.getSessions();
    if (sessions.length === 0) {
      const demoId = '8825'; // Default memorable room id
      const demoSession: Session = {
        id: demoId,
        name: 'Sessão de Demonstração Interativa',
        createdAt: new Date().toISOString(),
        currentQuestionIndex: 0,
        showResults: false,
        revealAnswer: false,
        isAnonymous: false,
        isQuizMode: true,
        showLeaderboard: false,
        status: 'active',
        questions: [
          {
            id: 'q-demo-1',
            text: 'Como as enquetes ao vivo ajudam no engajamento de turmas?',
            options: [
              'Aumentando a sonolência de quem assiste',
              'Forçando todos a ligarem as câmeras',
              'Estimulando a participação e dando voz a tímidos em tempo real',
              'Substituindo o professor no PowerPoint'
            ],
            correctOptionIndex: 2
          },
          {
            id: 'q-demo-2',
            text: 'Qual a recomendação de contraste de cores para leitura em projetores?',
            options: [
              'Fundo verde brilhante e texto amarelo',
              'Fundo claro ou escuro com alto contraste das fontes (ex: branco e preto)',
              'Fundo cinza claro e texto branco',
              'Nenhum contraste importa para leitura em salas abertas'
            ],
            correctOptionIndex: 1
          },
          {
            id: 'q-demo-3',
            text: 'O que o botão "Exibir Resultados" no painel do instrutor realiza?',
            options: [
              'Encerra a apresentação no PowerPoint imediatamente',
              'Libera o gráfico de barras ao vivo na tela projetada para todos verem',
              'Reinicia as votações dos alunos e zera a lista de presentes',
              'Sorteia um prêmio virtual aleatório'
            ],
            correctOptionIndex: 1
          }
        ]
      };
      
      this.saveSession(demoSession);

      // Pre-add some demo voters to make presentation live and fun straight away!
      const demoVoters: Participant[] = [
        { id: 'v-1', name: 'Alvaro Lima', sessionId: demoId, score: 200 },
        { id: 'v-2', name: 'Sarah Bernardes', sessionId: demoId, score: 300 },
        { id: 'v-3', name: 'Coronel Freitas', sessionId: demoId, score: 100 },
        { id: 'v-4', name: 'Mariana Costa', sessionId: demoId, score: 200 },
        { id: 'v-5', name: 'Lucas Ramos', sessionId: demoId, score: 0 }
      ];
      localStorage.setItem(PARTICIPANTS_KEY, JSON.stringify(demoVoters));

      // Cast initial demo votes for question 1 to populate graph on load
      const demoVotes: Vote[] = [
        { id: 'vt-1', sessionId: demoId, questionId: 'q-demo-1', participantId: 'v-1', participantName: 'Alvaro Lima', selectedOptionIndex: 2, timestamp: new Date().toISOString() },
        { id: 'vt-2', sessionId: demoId, questionId: 'q-demo-1', participantId: 'v-2', participantName: 'Sarah Bernardes', selectedOptionIndex: 2, timestamp: new Date().toISOString() },
        { id: 'vt-3', sessionId: demoId, questionId: 'q-demo-1', participantId: 'v-3', participantName: 'Coronel Freitas', selectedOptionIndex: 1, timestamp: new Date().toISOString() },
        { id: 'vt-4', sessionId: demoId, questionId: 'q-demo-1', participantId: 'v-4', participantName: 'Mariana Costa', selectedOptionIndex: 2, timestamp: new Date().toISOString() },
        { id: 'vt-5', sessionId: demoId, questionId: 'q-demo-1', participantId: 'v-5', participantName: 'Lucas Ramos', selectedOptionIndex: 0, timestamp: new Date().toISOString() },
        // Question 2 votes
        { id: 'vt-6', sessionId: demoId, questionId: 'q-demo-2', participantId: 'v-1', participantName: 'Alvaro Lima', selectedOptionIndex: 1, timestamp: new Date().toISOString() },
        { id: 'vt-7', sessionId: demoId, questionId: 'q-demo-2', participantId: 'v-2', participantName: 'Sarah Bernardes', selectedOptionIndex: 1, timestamp: new Date().toISOString() },
        { id: 'vt-8', sessionId: demoId, questionId: 'q-demo-2', participantId: 'v-3', participantName: 'Coronel Freitas', selectedOptionIndex: 0, timestamp: new Date().toISOString() },
        { id: 'vt-9', sessionId: demoId, questionId: 'q-demo-2', participantId: 'v-4', participantName: 'Mariana Costa', selectedOptionIndex: 1, timestamp: new Date().toISOString() }
      ];
      localStorage.setItem(VOTES_KEY, JSON.stringify(demoVotes));
    }
  },

  // SESSIONS
  getSessions(): Session[] {
    try {
      const raw = localStorage.getItem(SESSIONS_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error('Error parsing sessions from storage:', e);
      return [];
    }
  },

  getSession(id: string): Session | undefined {
    return this.getSessions().find(s => s.id === id);
  },

  saveSession(session: Session): void {
    const sessions = this.getSessions();
    const idx = sessions.findIndex(s => s.id === session.id);
    if (idx >= 0) {
      sessions[idx] = session;
    } else {
      sessions.push(session);
    }
    localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions));
    notifySyncChannel('session_updated', session);
  },

  deleteSession(id: string): void {
    const sessions = this.getSessions().filter(s => s.id !== id);
    localStorage.setItem(SESSIONS_KEY, JSON.stringify(sessions));
    notifySyncChannel('session_deleted', id);
    
    // Clear responses associated with it to save storage memory
    const votes = this.getVotesForSession(id);
    const remainingVotes = this.getAllVotesRaw().filter(v => v.sessionId !== id);
    localStorage.setItem(VOTES_KEY, JSON.stringify(remainingVotes));
    
    const remainingParticipants = this.getParticipantsRaw().filter(p => p.sessionId !== id);
    localStorage.setItem(PARTICIPANTS_KEY, JSON.stringify(remainingParticipants));
  },

  createSession(name: string, isAnonymous: boolean, isQuizMode: boolean): Session {
    // Generate simple readable unique 4-5 digit short code
    const existing = this.getSessions();
    let uniqueId = '';
    while (true) {
      uniqueId = Math.floor(1000 + Math.random() * 9000).toString();
      if (!existing.some(s => s.id === uniqueId)) break;
    }

    const newSession: Session = {
      id: uniqueId,
      name: name || `Aula/Palestra #${uniqueId}`,
      createdAt: new Date().toISOString(),
      questions: [],
      currentQuestionIndex: 0,
      showResults: false,
      revealAnswer: false,
      isAnonymous,
      isQuizMode,
      showLeaderboard: false,
      status: 'active'
    };

    this.saveSession(newSession);
    return newSession;
  },

  // QUESTIONS IN SESSION
  addQuestionToSession(sessionId: string, text: string, options: string[], correctOptionIndex: number | null, type?: 'alternativa' | 'aberta', resultsView?: 'live' | 'results-slide-only'): Question {
    const session = this.getSession(sessionId);
    if (!session) throw new Error('Session not found');

    const newQuestion: Question = {
      id: `q-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      text,
      options: type === 'aberta' ? [] : options.filter(opt => opt.trim() !== ''),
      correctOptionIndex: type === 'aberta' ? null : correctOptionIndex,
      type: type || 'alternativa',
      resultsView: resultsView || 'results-slide-only'
    };

    session.questions.push(newQuestion);
    this.saveSession(session);
    return newQuestion;
  },

  updateQuestionInSession(sessionId: string, questionId: string, text: string, options: string[], correctOptionIndex: number | null, type?: 'alternativa' | 'aberta', resultsView?: 'live' | 'results-slide-only'): void {
    const session = this.getSession(sessionId);
    if (!session) return;

    const idx = session.questions.findIndex(q => q.id === questionId);
    if (idx >= 0) {
      session.questions[idx] = {
        ...session.questions[idx],
        text,
        options: type === 'aberta' ? [] : options.filter(opt => opt.trim() !== ''),
        correctOptionIndex: type === 'aberta' ? null : correctOptionIndex,
        type: type || 'alternativa',
        resultsView: resultsView || 'results-slide-only'
      };
      this.saveSession(session);
    }
  },

  saveSlideMapping(sessionId: string, slideId: string, questionId: string, slideType: 'question' | 'responses'): void {
    const session = this.getSession(sessionId);
    if (!session) return;
    if (!session.slideMappings) session.slideMappings = [];

    // Remove old conflicting mappings for this slide or this specific question configuration
    session.slideMappings = session.slideMappings.filter(
      m => !(m.slideId === slideId || (m.questionId === questionId && m.slideType === slideType))
    );

    session.slideMappings.push({
      slideId,
      questionId,
      slideType
    });

    this.saveSession(session);
  },

  setCurrentSlideAndProjected(sessionId: string, slideId: string | null, questionId: string | null): void {
    const session = this.getSession(sessionId);
    if (session) {
      session.currentSlideId = slideId;
      session.projectedQuestionId = questionId;
      this.saveSession(session);
    }
  },

  deleteQuestionFromSession(sessionId: string, questionId: string): void {
    const session = this.getSession(sessionId);
    if (!session) return;

    // Shift current index if needed to prevent rendering out of array boundary
    const questionIdx = session.questions.findIndex(q => q.id === questionId);
    session.questions = session.questions.filter(q => q.id !== questionId);
    
    if (session.currentQuestionIndex >= session.questions.length && session.questions.length > 0) {
      session.currentQuestionIndex = session.questions.length - 1;
    } else if (session.questions.length === 0) {
      session.currentQuestionIndex = 0;
    }

    this.saveSession(session);

    // Clear votes for this specific deleted question
    const remainingVotes = this.getAllVotesRaw().filter(v => v.questionId !== questionId);
    localStorage.setItem(VOTES_KEY, JSON.stringify(remainingVotes));
  },

  // VOTING ENGINE
  getAllVotesRaw(): Vote[] {
    try {
      const raw = localStorage.getItem(VOTES_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error('Error parsing votes from storage:', e);
      return [];
    }
  },

  getVotesForSession(sessionId: string): Vote[] {
    return this.getAllVotesRaw().filter(v => v.sessionId === sessionId);
  },

  getVotesForQuestion(sessionId: string, questionId: string): Vote[] {
    return this.getAllVotesRaw().filter(v => v.sessionId === sessionId && v.questionId === questionId);
  },

  castVote(sessionId: string, questionId: string, participantId: string, participantName: string, selectedOptionIndex: number, textResponse?: string): boolean {
    const votes = this.getAllVotesRaw();
    
    // Check if participant has already voted on this specific question
    const alreadyVotedIdx = votes.findIndex(v => v.sessionId === sessionId && v.questionId === questionId && v.participantId === participantId);
    if (alreadyVotedIdx >= 0) {
      return false; // Prevent double voting
    }

    const session = this.getSession(sessionId);
    if (!session) return false;
    
    const activeQuestion = session.questions.find(q => q.id === questionId) || session.questions[session.currentQuestionIndex];
    if (!activeQuestion) return false;

    const newVote: Vote = {
      id: `vt-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      sessionId,
      questionId,
      participantId,
      participantName: session.isAnonymous ? 'Anônimo' : (participantName || 'Anônimo'),
      selectedOptionIndex,
      textResponse,
      timestamp: new Date().toISOString()
    };

    votes.push(newVote);
    localStorage.setItem(VOTES_KEY, JSON.stringify(votes));

    // Handle score accrual for quiz mode
    if (session.isQuizMode && activeQuestion && activeQuestion.correctOptionIndex !== null) {
      if (selectedOptionIndex === activeQuestion.correctOptionIndex) {
        this.accrualPoints(sessionId, participantId, 100);
      }
    }

    notifySyncChannel('vote_cast', { sessionId, questionId, vote: newVote });
    return true;
  },

  // PARTICIPANTS
  getParticipantsRaw(): Participant[] {
    try {
      const raw = localStorage.getItem(PARTICIPANTS_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error('Error parsing participants from storage:', e);
      return [];
    }
  },

  getParticipantsForSession(sessionId: string): Participant[] {
    return this.getParticipantsRaw().filter(p => p.sessionId === sessionId);
  },

  joinOrCreateParticipant(sessionId: string, participantName: string): Participant {
    const pts = this.getParticipantsRaw();
    const session = this.getSession(sessionId);
    
    const finalName = session?.isAnonymous ? 'Anônimo' : (participantName.trim() || 'Participante');
    
    // If identified, look for existing participant under the same name in this session to prevent multiple registrations
    if (session && !session.isAnonymous) {
      const match = pts.find(p => p.sessionId === sessionId && p.name.toLowerCase() === finalName.toLowerCase());
      if (match) return match;
    } else {
      // In anonymous mode, we can generate a unique tracker per reload/browser
    }

    const newParticipant: Participant = {
      id: `p-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      name: finalName,
      sessionId,
      score: 0
    };

    pts.push(newParticipant);
    localStorage.setItem(PARTICIPANTS_KEY, JSON.stringify(pts));
    notifySyncChannel('participant_joined', { sessionId, participant: newParticipant });
    return newParticipant;
  },

  accrualPoints(sessionId: string, participantId: string, points: number): void {
    const pts = this.getParticipantsRaw();
    const idx = pts.findIndex(p => p.sessionId === sessionId && p.id === participantId);
    if (idx >= 0) {
      pts[idx].score = (pts[idx].score || 0) + points;
      pts[idx].lastCorrectTime = Date.now();
      localStorage.setItem(PARTICIPANTS_KEY, JSON.stringify(pts));
    }
  },

  getLeaderboard(sessionId: string): Participant[] {
    return this.getParticipantsForSession(sessionId)
      .filter(p => p.name !== 'Anônimo')
      .sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        // Resolve ties by time taken or earlier score update (fallback)
        return (a.lastCorrectTime || 0) - (b.lastCorrectTime || 0);
      })
      .slice(0, 10); // Top 10 in leaderboard
  },

  // QUESTION BANK SERVICES
  getQuestionBank(): QuestionBankItem[] {
    try {
      const raw = localStorage.getItem(QUESTION_BANK_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      console.error('Error parsing question bank from storage:', e);
      return [];
    }
  },

  addQuestionToBank(text: string, options: string[], correctOptionIndex: number | null, category?: string): QuestionBankItem {
    const bank = this.getQuestionBank();
    const newItem: QuestionBankItem = {
      id: `qbank-${Date.now()}`,
      text,
      options: options.filter(opt => opt.trim() !== ''),
      correctOptionIndex,
      category: category || 'Geral',
      createdAt: new Date().toISOString()
    };
    bank.push(newItem);
    localStorage.setItem(QUESTION_BANK_KEY, JSON.stringify(bank));
    return newItem;
  },

  deleteQuestionFromBank(id: string): void {
    const bank = this.getQuestionBank().filter(q => q.id !== id);
    localStorage.setItem(QUESTION_BANK_KEY, JSON.stringify(bank));
  },

  // SESSION CONTROLS
  setQuestionIndex(sessionId: string, index: number): void {
    const session = this.getSession(sessionId);
    if (session && index >= 0 && index < session.questions.length) {
      session.currentQuestionIndex = index;
      session.showResults = false; // default hide results when moving questions
      session.revealAnswer = false; // hide correct reveal
      this.saveSession(session);
    }
  },

  toggleResults(sessionId: string, visible: boolean): void {
    const session = this.getSession(sessionId);
    if (session) {
      session.showResults = visible;
      this.saveSession(session);
    }
  },

  toggleRevealAnswer(sessionId: string, visible: boolean): void {
    const session = this.getSession(sessionId);
    if (session) {
      session.revealAnswer = visible;
      this.saveSession(session);
    }
  },

  endSession(sessionId: string): void {
    const session = this.getSession(sessionId);
    if (session) {
      session.status = 'ended';
      this.saveSession(session);
    }
  },

  // DEMO/TEST ENGINES: SIMULADOR DE VOTOS
  // Simulates random votes cast from virtual participants for quick demo charting.
  simulateVotes(sessionId: string, questionId: string, count: number): void {
    const session = this.getSession(sessionId);
    if (!session) return;

    const question = session.questions.find(q => q.id === questionId);
    if (!question) return;

    // Virtual participant names
    const names = [
      'Capitão Nascimento', 'Sargento Rocha', 'Dr. Gustavo', 'Prof. Angela', 'Eduardo Lima',
      'Camila Santos', 'Roberto Silva', 'Ana Julia', 'Fabiano Marques', 'Tenente Melo',
      'Bruno Gagliasso', 'Patricia Poeta', 'Pedro Bial', 'Rebeca Andrade', 'Vitor Hugo',
      'Aline Barros', 'Felipe Neto', 'Gisele Bündchen', 'Sandra Annenberg', 'Clara Nunes'
    ];

    const currentVoters = this.getParticipantsForSession(sessionId);
    let votesAdded = 0;

    for (let i = 0; i < count; i++) {
      // Pick random name or pick from pool
      const vName = names[Math.floor(Math.random() * names.length)] + ` (${Math.floor(Math.random() * 90 + 10)})`;
      
      // Register voter
      const dummyVoter = this.joinOrCreateParticipant(sessionId, vName);
      
      // Decide a chosen option index
      let selectedOption = Math.floor(Math.random() * question.options.length);
      
      // Add slight bias towards the correct answer to make mock graphs look intelligent!
      if (question.correctOptionIndex !== null && Math.random() < 0.55) {
        selectedOption = question.correctOptionIndex;
      }

      const success = this.castVote(sessionId, questionId, dummyVoter.id, dummyVoter.name, selectedOption);
      if (success) {
        votesAdded++;
      }
    }

    notifySyncChannel('simulated_votes_cast', { sessionId, questionId, count: votesAdded });
  },

  // EXPORT ENGINE
  exportToCSV(sessionId: string): string {
    const session = this.getSession(sessionId);
    if (!session) return '';

    const votes = this.getVotesForSession(sessionId);
    
    let csv = '\ufeff'; // UTF-8 BOM representation for correct Excel character loading
    csv += `Sessão:;"${session.name}"\n`;
    csv += `ID/Código:;"${session.id}"\n`;
    csv += `Data de Criação:;"${new Date(session.createdAt).toLocaleDateString()}"\n`;
    csv += `Status:;"${session.status === 'active' ? 'Ativa' : 'Encerrada'}"\n\n`;

    csv += 'Pergunta №;Pergunta;Alternativas;Total de Votos;Modo;Gabarito\n';
    
    session.questions.forEach((q, idx) => {
      const qVotes = votes.filter(v => v.questionId === q.id);
      const optStr = q.options.map((o, oidx) => `${String.fromCharCode(65 + oidx)}) ${o}`).join(' | ');
      const mode = q.correctOptionIndex !== null ? 'Quiz' : 'Enquete';
      const gabarito = q.correctOptionIndex !== null ? String.fromCharCode(65 + q.correctOptionIndex) : 'N/A';
      csv += `${idx + 1};"${q.text.replace(/"/g, '""')}";"${optStr.replace(/"/g, '""')}";${qVotes.length};"${mode}";"${gabarito}"\n`;
    });

    csv += '\n\nDetalhamento dos Votos de Participantes:\n';
    csv += 'Participante;Pergunta;Alternativa Escolhida;Correta?;Data/Hora\n';
    
    votes.forEach(v => {
      const q = session.questions.find(quest => quest.id === v.questionId);
      const isCorrect = q && q.correctOptionIndex !== null 
        ? (v.selectedOptionIndex === q.correctOptionIndex ? 'Correto' : 'Incorreto')
        : 'N/A';
      const optLabel = q ? `${String.fromCharCode(65 + v.selectedOptionIndex)}: ${q.options[v.selectedOptionIndex]}` : v.selectedOptionIndex;
      
      csv += `"${v.participantName.replace(/"/g, '""')}";"${(q?.text || '').replace(/"/g, '""')}";"${String(optLabel).replace(/"/g, '""')}";"${isCorrect}";"${new Date(v.timestamp).toLocaleString()}"\n`;
    });

    return csv;
  },

  exportToJSON(sessionId: string): string {
    const session = this.getSession(sessionId);
    if (!session) return '{}';

    const votes = this.getVotesForSession(sessionId);
    const participants = this.getParticipantsForSession(sessionId);

    return JSON.stringify({
      session,
      votes,
      participants
    }, null, 2);
  }
};
