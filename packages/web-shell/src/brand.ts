/**
 * The product names users see in the task pane. The build fills these from
 * brands/<GE_BRAND>/brand.json (see brand-assets.ts); without a build-time brand, as in unit tests,
 * the upstream default applies. Code identifiers (openGemini, GeminiPane) are not brand text.
 */
export interface Brand {
  /** Product name, e.g. in the header and error titles. */
  readonly name: string;
  /** How the assistant is addressed, e.g. "Ask Gemini". */
  readonly assistantName: string;
  /** Long form, e.g. page titles. */
  readonly fullName: string;
  /** Optional line under the name in the header, e.g. "CIMB Niaga AI Assistant". */
  readonly tagline?: string;
}

declare const __GE_BRAND__: Brand | undefined;

/** Mirrors brands/default/brand.json (brand.test.ts keeps the two in step). */
export const DEFAULT_BRAND: Brand = {
  name: 'Gemini Enterprise',
  assistantName: 'Gemini',
  fullName: 'Gemini Enterprise for Microsoft 365',
};

export const brand: Brand = typeof __GE_BRAND__ === 'undefined' ? DEFAULT_BRAND : __GE_BRAND__;
