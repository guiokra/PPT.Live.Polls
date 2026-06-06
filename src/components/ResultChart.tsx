/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { motion } from 'motion/react';
import { Check, X, Award } from 'lucide-react';

interface ResultChartProps {
  options: string[];
  votes: any[]; // List of Vote objects
  correctOptionIndex: number | null; // null if survey
  revealAnswer: boolean; // if true, highlight correct/incorrect
}

export const ResultChart: React.FC<ResultChartProps> = ({
  options,
  votes,
  correctOptionIndex,
  revealAnswer
}) => {
  const totalVotes = votes.length;

  // Count votes per option
  const voteCounters = options.map((_, idx) => {
    return votes.filter(v => v.selectedOptionIndex === idx).length;
  });

  return (
    <div className="space-y-4 w-full">
      {options.map((option, idx) => {
        const optionVotes = voteCounters[idx];
        const percent = totalVotes > 0 ? Math.round((optionVotes / totalVotes) * 100) : 0;
        const letter = String.fromCharCode(65 + idx); // A, B, C, D...

        // Determine design styling for item inside Sophisticated Dark
        let barColor = 'bg-zinc-800/40 border border-zinc-750/30';
        let badgeColor = 'bg-zinc-800 text-zinc-300 ring-1 ring-zinc-700';
        let textColor = 'text-zinc-200';
        let isCorrect = idx === correctOptionIndex;
        let showIndicatorIcon: 'none' | 'correct' | 'incorrect' = 'none';

        if (revealAnswer && correctOptionIndex !== null) {
          if (isCorrect) {
            barColor = 'bg-[rgba(163,230,53,0.18)] border border-[#a3e635]/50 shadow-[0_0_12px_rgba(163,230,53,0.15)]';
            badgeColor = 'bg-[#a3e635] text-zinc-950 ring-2 ring-[#a3e635]/30 font-bold';
            textColor = 'text-[#a3e635] font-black';
            showIndicatorIcon = 'correct';
          } else if (optionVotes > 0) {
            // Voted on incorrect option
            barColor = 'bg-red-950/20 border border-red-900/30';
            badgeColor = 'bg-red-950 text-red-300 ring-1 ring-red-900';
            textColor = 'text-zinc-500';
            showIndicatorIcon = 'incorrect';
          } else {
            // Correct answers revealed but this incorrect option had 0 votes
            barColor = 'bg-zinc-900/10 border border-dashed border-zinc-800';
            badgeColor = 'bg-zinc-900 text-zinc-600 ring-1 ring-zinc-800';
            textColor = 'text-zinc-650';
          }
        } else {
          // Normal state (before reveal or survey mode)
          if (totalVotes > 0 && optionVotes === Math.max(...voteCounters)) {
            // Leading alternative gets slightly brighter highlight in Sophisticated Dark Theme
            barColor = 'bg-[rgba(163,230,53,0.15)] border border-[#a3e635]/30';
            badgeColor = 'bg-[#a3e635] text-zinc-950 font-bold';
            textColor = 'text-zinc-100 font-semibold';
          } else {
            // Unvoted or alternate options
            barColor = 'bg-zinc-800/50 border border-zinc-700/40';
            badgeColor = 'bg-zinc-900 text-zinc-400 ring-1 ring-zinc-800';
          }
        }

        return (
          <div key={idx} className="relative flex flex-col w-full">
            {/* Top row with option label, text, and numerical result */}
            <div className="flex items-center justify-between z-10 px-3.5 py-3 text-sm">
              <div className="flex items-start gap-3 max-w-[80%]">
                <span className={`inline-flex items-center justify-center w-6 h-6 rounded-md font-mono text-xs font-bold shrink-0 transition-colors duration-300 ${badgeColor}`}>
                  {letter}
                </span>
                <span className={`leading-relaxed pt-0.5 transition-colors duration-300 ${textColor}`}>
                  {option}
                </span>
              </div>
              
              <div className="flex items-center gap-2 font-mono text-xs shrink-0 select-none">
                <span className="text-zinc-500">({optionVotes} {optionVotes === 1 ? 'voto' : 'votos'})</span>
                <span className={`font-bold text-sm ${revealAnswer && isCorrect ? 'text-[#a3e635]' : 'text-zinc-100'}`}>
                  {percent}%
                </span>
                {showIndicatorIcon === 'correct' && (
                  <Check className="w-4.5 h-4.5 text-[#a3e635] animate-bounce shrink-0 font-extrabold" />
                )}
                {showIndicatorIcon === 'incorrect' && (
                  <X className="w-4 h-4 text-red-400 shrink-0" />
                )}
              </div>
            </div>

            {/* Simulated background progress-bar using motion */}
            <div className="absolute inset-0 w-full h-full rounded-xl bg-[#121214]/65 overflow-hidden pointer-events-none border border-zinc-800/30">
              <motion.div
                initial={{ width: 0 }}
                animate={{ width: `${percent}%` }}
                transition={{ duration: 0.8, ease: 'easeOut' }}
                className={`h-full rounded-l-lg transition-colors duration-505 ${barColor}`}
              />
            </div>
          </div>
        );
      })}

      {totalVotes === 0 && (
        <div className="text-center py-8 text-zinc-500 border border-dashed border-zinc-800 rounded-xl bg-zinc-900/10 font-mono text-[11px]">
          <p>Aguardando respostas e votos dos participantes...</p>
        </div>
      )}
    </div>
  );
};
