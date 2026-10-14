import type { Surface } from '@ge/contracts';

/**
 * Map the Office host to our `Surface`. The web-shell is surface-agnostic; the only host-
 * specific decision it makes is *which bridge to instantiate*, and that hinges on this.
 */
const HOST_TO_SURFACE: Readonly<Record<string, Surface>> = {
  Word: 'word',
  Excel: 'excel',
  PowerPoint: 'powerpoint',
  OneNote: 'onenote',
  Outlook: 'outlook',
};

/** Pure: an `Office.HostType`-like string → Surface (undefined for unknown hosts). */
export function surfaceFromHost(host: string | undefined | null): Surface | undefined {
  if (!host) return undefined;
  return HOST_TO_SURFACE[host];
}

/** Minimal shape of the Office context we read for host detection (keeps this dep-light). */
export interface OfficeContextLike {
  context?: {
    host?: unknown;
    /** Present only in Outlook. */
    mailbox?: unknown;
  };
}

/**
 * Detect the surface at runtime. Outlook exposes `mailbox` rather than a `host`, so we check
 * that first; otherwise we read `Office.context.host`. Falls back to the global `Office`.
 */
export function detectSurface(office?: OfficeContextLike): Surface | undefined {
  const o = office ?? (globalThis as { Office?: OfficeContextLike }).Office;
  const ctx = o?.context;
  if (!ctx) return undefined;
  if (ctx.mailbox) return 'outlook';
  return surfaceFromHost(ctx.host != null ? String(ctx.host) : undefined);
}

/** Where the pane runs, for an error report (`host` in the reports API). Holds no user data. */
export interface ReportHost {
  label: string;
  app: string;
  platform: string;
  /** Office build; desktop only (Office on the web gives `0.0.0.0`). */
  version?: string;
  /** Outlook only: classic, new Outlook on Windows, web, mobile. */
  outlookClient?: string;
  browser: { name: string; version: string };
}

interface OfficeDiagnosticsLike {
  context?: {
    diagnostics?: { host?: unknown; platform?: unknown; version?: unknown };
    mailbox?: { diagnostics?: { hostName?: unknown } };
  };
}

interface NavigatorLike {
  userAgent?: string;
  userAgentData?: {
    brands?: Array<{ brand: string; version: string }>;
    getHighEntropyValues?(hints: string[]): Promise<{
      fullVersionList?: Array<{ brand: string; version: string }>;
    }>;
  };
}

const SURFACE_APP: Readonly<Record<string, string>> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  onenote: 'OneNote',
  outlook: 'Outlook',
  teams: 'Teams',
};

const PLATFORM_WORDS: Readonly<Record<string, string>> = {
  OfficeOnline: 'on the web',
  PC: 'on Windows',
  Mac: 'on Mac',
  iOS: 'on iOS',
  Android: 'on Android',
  Universal: 'on Windows',
};

const OUTLOOK_LABELS: Readonly<Record<string, string>> = {
  newOutlookWindows: 'new Outlook on Windows',
  OutlookWebApp: 'Outlook on the web',
  OutlookIOS: 'Outlook on iOS',
  OutlookAndroid: 'Outlook on Android',
};

const UA_RULES: ReadonlyArray<readonly [string, RegExp]> = [
  ['Microsoft Edge', /Edg(?:A|iOS)?\/([\d.]+)/],
  ['Firefox', /(?:Firefox|FxiOS)\/([\d.]+)/],
  ['Google Chrome', /(?:Chrome|CriOS)\/([\d.]+)/],
  ['Safari', /Version\/([\d.]+).*Safari/],
];

/** Browser (web) or web view (desktop) name and version. Needs no permission. */
export async function readBrowser(
  nav: NavigatorLike | undefined = (globalThis as { navigator?: NavigatorLike }).navigator,
): Promise<{ name: string; version: string }> {
  const uad = nav?.userAgentData;
  if (uad && Array.isArray(uad.brands)) {
    let brands = uad.brands;
    try {
      const high = await uad.getHighEntropyValues?.(['fullVersionList']);
      if (high?.fullVersionList?.length) brands = high.fullVersionList;
    } catch {
      /* The browser may refuse the exact build; the major version is enough. */
    }
    const real = brands.filter((b) => !/not.?a.?brand/i.test(b.brand));
    const pick = real.find((b) => b.brand !== 'Chromium') ?? real[0];
    if (pick) return { name: pick.brand, version: pick.version };
  }
  const ua = nav?.userAgent ?? '';
  for (const [name, re] of UA_RULES) {
    const m = re.exec(ua);
    if (m?.[1]) return { name, version: m[1] };
  }
  // Office on Mac runs the pane in WKWebView: the Safari engine without a `Version/` token.
  const webkit = /AppleWebKit\/([\d.]+)/.exec(ua);
  if (webkit?.[1]) return { name: 'WebKit', version: webkit[1] };
  return { name: 'unknown', version: '' };
}

/** Read the Office app, platform, build and browser for an error report. */
export async function readHost(
  surface: string,
  office: OfficeDiagnosticsLike | undefined = (globalThis as { Office?: OfficeDiagnosticsLike })
    .Office,
  nav?: NavigatorLike,
): Promise<ReportHost> {
  const diag = office?.context?.diagnostics;
  const app =
    typeof diag?.host === 'string' && diag.host ? diag.host : (SURFACE_APP[surface] ?? surface);
  const platform = typeof diag?.platform === 'string' && diag.platform ? diag.platform : 'unknown';
  const version =
    typeof diag?.version === 'string' &&
    /^\d+\.\d+/.test(diag.version) &&
    !/^0(\.0)*$/.test(diag.version)
      ? diag.version
      : undefined;
  const hostName = office?.context?.mailbox?.diagnostics?.hostName;
  const outlookClient = typeof hostName === 'string' && hostName ? hostName : undefined;
  const label =
    (outlookClient && OUTLOOK_LABELS[outlookClient]) ??
    (PLATFORM_WORDS[platform] ? `${app} ${PLATFORM_WORDS[platform]}` : app);
  return {
    label: label.slice(0, 80),
    app,
    platform,
    ...(version ? { version } : {}),
    ...(outlookClient ? { outlookClient } : {}),
    browser: await readBrowser(nav),
  };
}
