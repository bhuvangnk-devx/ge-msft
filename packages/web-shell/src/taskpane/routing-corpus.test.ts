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
  'What changed in this deck?',
  'Tell me about the Enterprise plan',
  'Explain the difference between Pro and Basic',
  'Which customers are inactive?',
  "What's the average spend per plan?",
  'Compare Q3 and Q4 numbers',
  'What are the risks in this contract?',
  'Can you check if the totals are correct?',
  'Is the formula in B5 right?',
  'Can you find the clause about payment terms?',
  'Where is the indemnity section?',
  'So what is the conclusion?',
  'What is in B5?',
  'Explain the formula in G12',
  'What does the comment on the SLA say?',
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
  'In G12 put a formula that sums G2:G11',
  'show me a bar chart of Total by Product',
  'I want a bar chart of revenue by region',
  'Enter 42 in B5',
  'can you append hello to the selected cell',
  'append hello to the selected cell',
  'could you prepend a dash to each row',
  'can you bold the selected cells',
  'can you merge A1:C1',
  'would you clear the selected range',
  'can you duplicate this slide',
  'can you calculate the total of column D',
  'can you copy row 2 to row 10',
  'can you translate this paragraph into the cell B2',
  'can you swap columns B and C',
  'remove duplicates from this table',
  'could you shorten this paragraph',
  'G12 should be the sum of G2:G11',
  'B5 = 100',
  'Total in G12 please',
  'The value in B5 should be 300',
  'D2:D11 = B2:B11*C2:C11',
  'Reply to the comment on D5 saying I will check',
  'Can you resolve the comment thread on payment terms?',
  'Fix the typo in the title',
  'Insert a table with Item and Owner',
  "I'd like you to rewrite this paragraph",
  'Help me update the schedule',
  "Let's change the title to FY26",
];
/** Indonesian (CIMB users). Same rules as above. */
const ID_QUESTIONS = [
  'Apa isi email ini?',
  'Berapa total pendapatan di sheet ini?',
  'Bagaimana cara menambah grafik?',
  'Kenapa grafiknya tidak muncul?',
  'Mengapa MRR turun bulan ini?',
  'Siapa yang menulis komentar di D5?',
  'Kapan tenggat pembayaran?',
  'Bisa jelaskan perbedaan Pro dan Basic?',
  'Pelanggan mana yang tidak aktif?',
  'Di mana bagian ganti rugi?',
  'Apakah rumus di B5 benar?',
  'Menurutmu saya harus membalas?',
  'Isi dokumen ini tentang apa?',
  'Apa yang harus saya tulis di balasan?',
];
const ID_ACTIONS = [
  'Tolong balas email ini',
  'Balas email ini, bilang saya cek hari Jumat',
  'Saya mau membalas email ini',
  'Buatkan grafik pendapatan per paket',
  'Tolong buat grafik batang dari A1:B8',
  'Tambahkan slide tentang Q4',
  'Tambah slide sebagai slide kedua',
  'Hapus slide terakhir',
  'Ubah judul menjadi FY26',
  'Ganti judul tabel jadi Ringkasan',
  'Tolong perbaiki salah ketik di judul',
  'Format header menjadi tebal',
  'Sisipkan tabel dengan kolom Item dan Pemilik',
  'Bisakah kamu menulis ulang paragraf ini?',
  'Saya ingin menambahkan komentar di sini',
  'Urutkan data berdasarkan tanggal',
  'Isikan kolom D dengan total',
  'Sorot baris yang terlambat',
  'Mohon balas komentar tentang termin pembayaran',
  'Buat draf balasan untuk pelanggan ini',
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
  it.each(ID_QUESTIONS)('stays in chat on every surface (id): %s', (q) => {
    for (const surface of SURFACES) expect(route(surface, q), surface).toBe('chat');
  });
  it.each(ID_ACTIONS)('leaves chat on at least one surface (id): %s', (a) => {
    expect(SURFACES.map((surface) => route(surface, a))).not.toEqual(SURFACES.map(() => 'chat'));
  });
  it.each(QUESTIONS)('stays in chat on every surface: %s', (q) => {
    for (const surface of SURFACES) expect(route(surface, q), surface).toBe('chat');
  });
  it.each(ACTIONS)('leaves chat on at least one surface: %s', (a) => {
    expect(SURFACES.map((surface) => route(surface, a))).not.toEqual(SURFACES.map(() => 'chat'));
  });
});

describe('Indonesian fast paths reach the matching intent', () => {
  it.each([
    ['excel', 'Buatkan grafik pendapatan per paket', 'intent:visualize'],
    ['excel', 'Tolong buat diagram batang dari A1:B8', 'intent:visualize'],
    ['powerpoint', 'Tambahkan slide tentang Q4', 'intent:draft'],
    ['outlook', 'Buat draf balasan untuk pelanggan ini', 'intent:draft'],
    ['outlook', 'Tolong balas email ini', 'planner'],
    ['word', 'Mohon balas komentar tentang termin pembayaran', 'planner'],
  ] as const)('%s: %s → %s', (surface, raw, expected) => {
    expect(route(surface, raw)).toBe(expected);
  });
});

describe('cell references: statements are edits, questions stay in chat (fix E: other text → planner)', () => {
  it.each([
    ['b5 = 100', 'planner'],
    ["'Q3 Sales'!B5 should be 100", 'planner'],
    ['Sheet1!A1:C10 needs a border', 'planner'],
    ['total for g2:g11 in g12', 'planner'],
    ['$B$5 = 10%', 'planner'],
    ['B5 to 100', 'planner'],
    ['Nilai B5 harus 100', 'planner'],
    ['put 5 in A4', 'planner'],
    ['Is B5 correct', 'chat'],
    ['Are the totals in G12 right', 'chat'],
    ['Was G12 changed', 'chat'],
    ['Could you check B5', 'planner'],
    ['Can you tell me what B5 means', 'chat'],
    ['Print this on A4 paper', 'planner'],
    ['MP3 files are attached', 'planner'],
    ['PS5 sales are up', 'planner'],
    ['B2B sales are up', 'planner'],
    ['FY26 plan looks good', 'planner'],
    ['Windows 11 is installed', 'planner'],
    ['H1 results are in', 'planner'],
    ['thanks for B5', 'planner'],
    ['Apa isi B5?', 'chat'],
    ['Jelaskan rumus di G12', 'planner'],
  ] as const)('excel: %s → %s', (raw, expected) => {
    expect(route('excel', raw)).toBe(expected);
  });
});

/**
 * Fix E (docs/COMMAND-RELIABILITY.md): on a surface that can write, text that is not a question goes
 * to the planner, which answers chat-type requests in chat. These non-questions therefore take a
 * planner turn first; recorded so a change to that policy is deliberate.
 */
const NON_QUESTIONS_TO_PLANNER = [
  'Jelaskan model langganan ini',
  'Tolong jelaskan rumus ini',
  'Ringkas email ini',
  'Tolong ringkaskan dokumen ini',
  'Saya mau tahu paket mana yang paling laku',
  'Saya ingin memahami klausul ini',
  'Halo',
  'Terima kasih!',
  'Bagaimana cara membalas komentar?',
  'I want to know which plan earns the most',
  'I want to understand this formula',
  "I'd like to know the churn rate",
  'I need to understand why MRR dropped',
  'I need help understanding the pivot table',
  'Help me understand this clause',
  'Help me understand the deck',
  'hi',
  'thanks!',
  'okay',
  'List the action items in this email',
  'Let me know what you think of this slide',
  'Let me think about it',
  'I want to see the top 5 customers',
  'Translate this paragraph to Indonesian',
  'Give me a summary of the thread',
  'Show me the trial users',
  'Please explain this to me',
  'Now tell me the total',
  'How do I reply to a comment?',
  'Q3 revenue looks low',
  'can you help me understand this',
  'can you check the totals',
  'would you recommend a chart type',
  'hello, how are you',
  'I think the selected cell is wrong',
  'The selected cell looks off',
  'can you be more specific',
  'thanks, that works',
];

describe('non-questions go to the planner (fix E)', () => {
  it.each(NON_QUESTIONS_TO_PLANNER)('%s', (raw) => {
    expect(SURFACES.map((surface) => route(surface, raw))).not.toEqual(SURFACES.map(() => 'chat'));
  });
});
