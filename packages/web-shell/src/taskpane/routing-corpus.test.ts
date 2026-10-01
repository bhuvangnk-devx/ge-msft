import { describe, expect, it } from 'vitest';
import type { Intent, Surface } from '@ge/contracts';
import { inferImplicitIntent, shouldUsePlannerForFreeText } from './components/App.js';

/**
 * A routing regression corpus. Every question must stay in chat on every surface (no planner, no
 * command loop); every action request must leave chat. A starts-with-a-verb question such as
 * "Draft ideas for a subject line?" is deliberately absent: it goes to the planner, whose "chat"
 * verdict then answers it in chat (see controller.test.ts).
 */
const QUESTIONS = [
  'What is the total revenue in this sheet?',
  'Summarize this email',
  'Can you summarize this document?',
  'Can you explain the subscription model?',
  'Could you tell me who is on trial?',
  'I want to know which plan earns the most',
  'I want to understand this formula',
  "I'd like to know the churn rate",
  'I need to understand why MRR dropped',
  'I need help understanding the pivot table',
  'Help me understand this clause',
  'Help me understand the deck',
  'Why did the chart not update?',
  'How do I add a chart?',
  'How can I reply to this email?',
  'Is it possible to add slides from Word?',
  'What should I reply to this customer?',
  'What does this comment mean?',
  'Who wrote the comment on D5?',
  'Do you think I should reply?',
  'Should I change the title?',
  'hey?',
  'hi',
  'thanks!',
  'okay',
  'What changed in this deck?',
  'Tell me about the Enterprise plan',
  'Explain the difference between Pro and Basic',
  'Which customers are inactive?',
  "What's the average spend per plan?",
  'Compare Q3 and Q4 numbers',
  'List the action items in this email',
  'What are the risks in this contract?',
  'Can you check if the totals are correct?',
  'Is the formula in B5 right?',
  'Let me know what you think of this slide',
  'Let me think about it',
  'I want to see the top 5 customers',
  'Can you find the clause about payment terms?',
  'Where is the indemnity section?',
  'Translate this paragraph to Indonesian',
  'Give me a summary of the thread',
  'Show me the trial users',
  'Please explain this to me',
  'So what is the conclusion?',
  'Now tell me the total',
];
const ACTIONS = [
  'I want to reply to this email',
  'Reply saying I will check by Friday',
  'Can you add a chart of revenue by plan?',
  'Create a chart of MRR per plan',
  'I want to add a slide about Q4',
  'Add a slide as the second slide',
  'Please format the header bold',
  'Rename the sheet to Summary',
  'Delete the last slide',
  'Fix the typo in the title',
  'Insert a table with Item and Owner',
  "I'd like you to rewrite this paragraph",
  'Help me update the schedule',
  "Let's change the title to FY26",
];
const SURFACES: Surface[] = ['word', 'excel', 'powerpoint', 'outlook', 'onenote'];
const ALL: Intent[] = [
  'ask',
  'summarize',
  'explain',
  'rewrite',
  'review',
  'draft',
  'notes',
  'visualize',
] as Intent[];
const base = { scope: { kind: 'selection' as const }, mentions: [] };
const route = (s: Surface, raw: string) => {
  const inv = { ...base, raw, instruction: raw };
  const intent = inferImplicitIntent(s, ALL, inv);
  if (intent) return `intent:${intent}`;
  return shouldUsePlannerForFreeText(ALL, inv) ? 'planner' : 'chat';
};
describe('free-text routing corpus', () => {
  it.each(QUESTIONS)('stays in chat on every surface: %s', (q) => {
    for (const surface of SURFACES) expect(route(surface, q), surface).toBe('chat');
  });
  it.each(ACTIONS)('leaves chat on at least one surface: %s', (a) => {
    expect(SURFACES.map((surface) => route(surface, a))).not.toEqual(SURFACES.map(() => 'chat'));
  });
});
