/**
 * The pane follows the language the user started the conversation in (English or Bahasa
 * Indonesia). There is no language setting: the first request decides, and a new chat decides again.
 */
export type UiLanguage = 'en' | 'id';

/** Common Bahasa Indonesia words that rarely appear in English requests. */
const BAHASA =
  /\b(?:yang|dan|ini|itu|untuk|dengan|dari|tolong|buat(?:kan)?|tambah(?:kan)?|ubah|ganti|apa|bagaimana|berapa|mana|ringkas(?:kan)?|jelaskan|sebutkan|analisis|semua|setiap|baris|kolom|tabel|grafik|slide nya|surat|balas(?:kan)?|tidak|sudah|belum|saya|kami|kita|tahun|bulan|wilayah|cabang)\b/gi;

/** Bahasa when at least two Bahasa words appear (one can be a shared word such as "data"). */
export function detectLanguage(text: string): UiLanguage {
  return (text.match(BAHASA) ?? []).length >= 2 ? 'id' : 'en';
}

const STRINGS = {
  en: {
    doneOne: 'Done: 1 change',
    doneMany: (n: number) => `Done: ${n} changes`,
    showMe: 'Show me',
    undo: 'Undo',
    undoAll: 'Undo all',
    undone: 'Undone',
    undoInOffice: 'To undo, use Undo (Ctrl+Z) in Office.',
    undoRequest: 'Undo the last change',
    failedTitle: "That didn't finish",
    failedPartial: (made: number) =>
      made === 1
        ? '1 change was made before it stopped.'
        : `${made} changes were made before it stopped.`,
    tryAgain: 'Try again',
    reportSent: (id: string) => `Reported to support · ${id}`,
    reportNotSent: 'Could not send the error report. Use Copy diagnostics in Activity.',
    blockedTitle: 'This request was blocked',
    blockedBody:
      "Your organisation's safety settings stopped this request. Try asking in a different way.",
    resumeTitle: 'Your last request was interrupted',
    resumeBody: (request: string, made: number) =>
      `"${request}": ${made === 0 ? 'nothing was changed yet' : made === 1 ? '1 change was made' : `${made} changes were made`} before the pane closed.`,
    resumeFinish: 'Finish it',
    resumeDismiss: 'Dismiss',
    planTitle: 'Here is what I plan to do',
    planGo: 'Go ahead',
    cancel: 'Cancel',
    planIn: 'In',
    planSelection: 'Using your selection',
    planKeep: 'Leave unchanged:',
    planUsing: 'Using',
    reviewOne: 'Review this change',
    reviewMany: (n: number) => `Review ${n} changes`,
    apply: 'Apply',
    applyNote: 'Nothing changes until you press Apply.',
    applyOrder: 'Changes are applied in order. If one fails, the earlier ones stay.',
    check: 'Check',
    where: 'Where',
    becomes: 'Becomes',
    beforeAfter: 'Before → after',
    noPreview: 'No preview for this change. It runs as shown above.',
  },
  id: {
    doneOne: 'Selesai: 1 perubahan',
    doneMany: (n: number) => `Selesai: ${n} perubahan`,
    showMe: 'Tunjukkan',
    undo: 'Urungkan',
    undoAll: 'Urungkan semua',
    undone: 'Diurungkan',
    undoInOffice: 'Untuk mengurungkan, gunakan Undo (Ctrl+Z) di Office.',
    undoRequest: 'Urungkan perubahan terakhir',
    failedTitle: 'Maaf, belum berhasil',
    failedPartial: (made: number) => `${made} perubahan sudah dibuat sebelum berhenti.`,
    tryAgain: 'Coba lagi',
    reportSent: (id: string) => `Dilaporkan ke tim support · ${id}`,
    reportNotSent: 'Laporan error tidak terkirim. Gunakan Copy diagnostics di Activity.',
    blockedTitle: 'Permintaan ini diblokir',
    blockedBody:
      'Pengaturan keamanan organisasi Anda menghentikan permintaan ini. Coba tanyakan dengan cara lain.',
    resumeTitle: 'Permintaan terakhir Anda terputus',
    resumeBody: (request: string, made: number) =>
      `"${request}": ${made === 0 ? 'belum ada yang diubah' : `${made} perubahan sudah dibuat`} sebelum panel ditutup.`,
    resumeFinish: 'Selesaikan',
    resumeDismiss: 'Tutup',
    planTitle: 'Ini rencana saya',
    planGo: 'Lanjutkan',
    cancel: 'Batal',
    planIn: 'Di',
    planSelection: 'Menggunakan pilihan Anda',
    planKeep: 'Tidak diubah:',
    planUsing: 'Menggunakan',
    reviewOne: 'Periksa perubahan ini',
    reviewMany: (n: number) => `Periksa ${n} perubahan`,
    apply: 'Terapkan',
    applyNote: 'Tidak ada yang berubah sebelum Anda menekan Terapkan.',
    applyOrder: 'Perubahan diterapkan berurutan. Jika satu gagal, yang sebelumnya tetap.',
    check: 'Periksa',
    where: 'Lokasi',
    becomes: 'Menjadi',
    beforeAfter: 'Sebelum → sesudah',
    noPreview: 'Tidak ada pratinjau untuk perubahan ini. Akan dijalankan seperti di atas.',
  },
} as const;

export type UiStrings = (typeof STRINGS)['en'];

export function uiStrings(lang: UiLanguage): UiStrings {
  return STRINGS[lang] as UiStrings;
}

const KIND_LABEL: Record<UiLanguage, Record<string, string>> = {
  en: {
    'write-cells': 'Cells',
    'format-cells': 'Format',
    'create-table': 'Table',
    'insert-chart': 'Chart',
    'format-conditional': 'Conditional format',
    'add-comment': 'Comment',
    'comment-reply': 'Reply to comment',
    'tracked-change': 'Suggested edit',
    'replace-selection': 'Selection text',
    'find-replace': 'Find and replace',
    'insert-text': 'Text',
    'insert-ooxml': 'Content',
    'insert-table': 'Table',
    'apply-style': 'Style',
    'insert-slide': 'New slide',
    'add-table-slide': 'Table slide',
    'set-shape-text': 'Slide text',
    'add-shape': 'Shape',
    'format-shape': 'Shape format',
    'reply-mail': 'Reply draft',
    'create-mail': 'New email',
    'set-recipients': 'Recipients',
    'insert-hyperlink': 'Link',
    'add-attachment': 'Attachment',
    'delete-slide': 'Slide removal',
    'manage-worksheet': 'Sheet',
    'resolve-revisions': 'Tracked changes',
  },
  id: {
    'write-cells': 'Sel',
    'format-cells': 'Format',
    'create-table': 'Tabel',
    'insert-chart': 'Grafik',
    'format-conditional': 'Format bersyarat',
    'add-comment': 'Komentar',
    'comment-reply': 'Balasan komentar',
    'tracked-change': 'Usulan perubahan',
    'replace-selection': 'Teks pilihan',
    'find-replace': 'Cari dan ganti',
    'insert-text': 'Teks',
    'insert-ooxml': 'Konten',
    'insert-table': 'Tabel',
    'apply-style': 'Gaya',
    'insert-slide': 'Slide baru',
    'add-table-slide': 'Slide tabel',
    'set-shape-text': 'Teks slide',
    'add-shape': 'Bentuk',
    'format-shape': 'Format bentuk',
    'reply-mail': 'Draf balasan',
    'create-mail': 'Email baru',
    'set-recipients': 'Penerima',
    'insert-hyperlink': 'Tautan',
    'add-attachment': 'Lampiran',
    'delete-slide': 'Penghapusan slide',
    'manage-worksheet': 'Sheet',
    'resolve-revisions': 'Perubahan terlacak',
  },
};

/** A plain noun for a change kind (read before and after it runs, so never past tense), falling back to the kind with dashes as spaces. */
export function kindLabel(kind: string, lang: UiLanguage): string {
  return KIND_LABEL[lang][kind] ?? kind.replace(/-/g, ' ');
}
