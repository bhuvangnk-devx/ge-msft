import { contentHash } from '@ge/gemini-client';

/** Snapshot identity must never degrade to a non-cryptographic digest. */
export async function evidenceHash(text: string): Promise<string> {
  const hash = await contentHash(text);
  if (!/^sha256:[a-f0-9]{64}$/.test(hash))
    throw new Error('Evidence snapshots require Web Crypto SHA-256.');
  return hash.slice(7);
}
