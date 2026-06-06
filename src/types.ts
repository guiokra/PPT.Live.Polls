/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface SlideMapping {
  slideId: string;
  slideIndex?: number;
  questionId: string;
  slideType: 'question' | 'responses';
}

export interface Question {
  id: string;
  text: string;
  options: string[]; // e.g. ["Option A", "Option B", ...]
  correctOptionIndex: number | null; // index of correct answer, or null if it's a survey poll
  type?: 'alternativa' | 'aberta';
  resultsView?: 'live' | 'results-slide-only';
}

export interface Session {
  id: string; // unique short code, e.g. "88319" or "POLL-XYZ"
  name: string;
  createdAt: string;
  questions: Question[];
  currentQuestionIndex: number; // index of the active presented question
  showResults: boolean; // whether results are displayed on presentation screen
  revealAnswer: boolean; // whether correct answer is highlighted on presentation screen
  isAnonymous: boolean; // whether participants join without names
  isQuizMode: boolean; // enables correct answers tracking and leaderboard
  showLeaderboard: boolean; // whether to show ranking on screen
  status: 'active' | 'ended';
  slideMappings?: SlideMapping[];
}

export interface Vote {
  id: string;
  sessionId: string;
  questionId: string;
  participantId: string;
  participantName: string;
  selectedOptionIndex: number; // for alternative
  textResponse?: string; // for custom open-ended text answers
  timestamp: string;
}

export interface Participant {
  id: string;
  name: string;
  sessionId: string;
  score: number; // accumulated score for Quiz Mode (+100 XP per correct response)
  lastCorrectTime?: number; // microsecond timing for resolving leaderboard ties
}

export interface QuestionBankItem {
  id: string;
  text: string;
  options: string[];
  correctOptionIndex: number | null;
  category?: string;
  createdAt: string;
}
