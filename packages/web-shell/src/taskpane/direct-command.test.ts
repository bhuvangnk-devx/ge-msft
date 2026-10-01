import { describe, expect, it } from 'vitest';
import { extractDirectCommandProgram } from './direct-command.js';

describe('extractDirectCommandProgram', () => {
  it('extracts a pasted schedule command run and ignores prose plus trailing slash chat', () => {
    const program = extractDirectCommandProgram(`
Populate a mock schedule for this please
set 'Daily schedule'!B2 "Time"
set 'Daily schedule'!C2 "Monday"
set 'Daily schedule'!G12 "Wrap Up & Planning"
/summarize @this Summarize the selected range.
`);

    expect(program).toBe(
      [
        `set 'Daily schedule'!B2 "Time"`,
        `set 'Daily schedule'!C2 "Monday"`,
        `set 'Daily schedule'!G12 "Wrap Up & Planning"`,
      ].join('\n'),
    );
  });

  it('does not turn explanatory prose containing a command-looking word into CLI', () => {
    expect(extractDirectCommandProgram('why did set Daily schedule not write?')).toBeUndefined();
    expect(extractDirectCommandProgram('set a reminder for lunch')).toBeUndefined();
  });

  it('treats a sentence that starts with a CLI verb as prose, not a command', () => {
    for (const raw of [
      'Read the last row for me',
      'read the last row for me',
      'Search for the payment clause',
      'List the open risks please',
      'Outline the key points of this deck',
    ]) {
      expect(extractDirectCommandProgram(raw), raw).toBeUndefined();
    }
    for (const cli of [
      'read Sales!A1:B9',
      "read 'Daily schedule'!B2:I10",
      'read selection',
      'search "payment terms"',
      'list shape',
      'chart bar A1:B5',
      'inspect result:abc path=/rows offset=0 limit=20',
    ]) {
      expect(extractDirectCommandProgram(cli), cli).toBe(cli);
    }
  });

  it('treats "find and replace …" as a request, not the workspace find command', () => {
    expect(extractDirectCommandProgram('Find and replace CN-GPT with CNGPT')).toBeUndefined();
    expect(extractDirectCommandProgram('find /workspace *.md')).toBe('find /workspace *.md');
  });

  it('accepts an explicit cmd fence even when it contains one command', () => {
    expect(
      extractDirectCommandProgram(`
\`\`\`cmd
set 'Daily schedule'!B2 "Time"
\`\`\`
`),
    ).toBe(`set 'Daily schedule'!B2 "Time"`);
  });
});
