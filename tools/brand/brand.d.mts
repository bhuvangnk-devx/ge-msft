/** Types for brand.mjs (plain ESM so the release tools can run it without a build step). */
export interface Brand {
  id: string;
  dir: string;
  name: string;
  assistantName: string;
  fullName: string;
  fontStylesheet: string;
  iconsDir: string;
  themePath: string | null;
  assetsDir: string | null;
}
export interface BrandText {
  name: string;
  assistantName: string;
  fullName: string;
}
export const brandsRoot: string;
export const BRAND_ICONS: readonly string[];
export function selectedBrand(): string;
export function loadBrand(id?: string): Brand;
export function brandText(brand: Brand): BrandText;
export function fillBrandTokens(html: string, brand: Brand): string;
export function brandAssets(brand: Brand): Record<string, string>;
export function contentHash(path: string): string;
