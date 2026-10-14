import { describe, it, expect, vi } from 'vitest';
import {
  NaaAuthClient,
  type MsalLike,
  type MsalAccount,
  type MsalAuthResult,
} from './auth-client.js';

const account: MsalAccount = { username: 'v.k@acme', name: 'V K', localAccountId: 'oid-1' };

function fakeMsal(over: Partial<MsalLike> = {}): MsalLike {
  return {
    getActiveAccount: () => account,
    acquireTokenSilent: (req) =>
      Promise.resolve<MsalAuthResult>({
        accessToken: `acc:${req.scopes.join(',')}`,
        idToken: 'id-tok',
        account,
      }),
    ...over,
  };
}

const opts = { idTokenScopes: ['api://ge/.default'], graphScopes: ['Mail.Read'] };

describe('NaaAuthClient', () => {
  it('returns the id token (WIF subject) from silent acquisition', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getIdToken()).toBe('id-tok');
  });

  it('asks MSAL for a newly issued id token when forceRefresh is set, and only then', async () => {
    const silent = vi.fn((req: { forceRefresh?: boolean; scopes: string[] }) =>
      Promise.resolve<MsalAuthResult>({
        accessToken: 'a',
        idToken: req.forceRefresh ? 'fresh' : 'cached',
        account,
      }),
    );
    const auth = new NaaAuthClient(fakeMsal({ acquireTokenSilent: silent }), opts);
    expect(await auth.getIdToken()).toBe('cached');
    expect(await auth.getIdToken({ forceRefresh: true })).toBe('fresh');
    expect(silent.mock.calls[0]![0]).not.toHaveProperty('forceRefresh');
    expect(silent.mock.calls[1]![0]).toMatchObject({ forceRefresh: true, account });
  });

  it('throws when no id token comes back', async () => {
    const auth = new NaaAuthClient(
      fakeMsal({ acquireTokenSilent: () => Promise.resolve({ accessToken: 'a' }) }),
      opts,
    );
    await expect(auth.getIdToken()).rejects.toThrow(/no id token/i);
  });

  it('acquires a Graph token for the requested scopes', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getGraphToken(['Calendars.Read'])).toBe('acc:Calendars.Read');
  });

  it('falls back to the configured graph scopes when none are passed', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getGraphToken([])).toBe('acc:Mail.Read');
  });

  it('falls back to popup when silent acquisition needs interaction', async () => {
    const interaction = Object.assign(new Error('need ui'), {
      name: 'InteractionRequiredAuthError',
    });
    const popup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'popup-acc', idToken: 'popup-id' }),
    );
    const auth = new NaaAuthClient(
      fakeMsal({ acquireTokenSilent: () => Promise.reject(interaction), acquireTokenPopup: popup }),
      opts,
    );
    expect(await auth.getIdToken()).toBe('popup-id');
    expect(popup).toHaveBeenCalledOnce();
  });

  it('derives identity by using Office web loginHint when no account is cached', async () => {
    let remembered: MsalAccount | null = null;
    const ssoSilent = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({
        accessToken: 'sso-acc',
        idToken: 'sso-id',
        account,
      }),
    );
    const msal = fakeMsal({
      getActiveAccount: () => remembered,
      getAllAccounts: () => [],
      setActiveAccount: (next) => {
        remembered = next;
      },
      ssoSilent,
    });

    const auth = new NaaAuthClient(msal, { ...opts, loginHint: 'v.k@acme' });
    await expect(auth.getIdentity()).resolves.toEqual({
      username: 'v.k@acme',
      displayName: 'V K',
      oid: 'oid-1',
    });
    expect(ssoSilent).toHaveBeenCalledWith({
      scopes: opts.idTokenScopes,
      loginHint: 'v.k@acme',
    });
    expect(remembered).toBe(account);
  });

  it('falls back to popup when Office web silent SSO needs interaction', async () => {
    const interaction = Object.assign(new Error('need ui'), {
      name: 'InteractionRequiredAuthError',
    });
    const popup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'popup-acc', idToken: 'popup-id', account }),
    );
    const auth = new NaaAuthClient(
      fakeMsal({
        getActiveAccount: () => null,
        getAllAccounts: () => [],
        ssoSilent: () => Promise.reject(interaction),
        acquireTokenPopup: popup,
      }),
      { ...opts, loginHint: 'v.k@acme' },
    );

    expect(await auth.getIdToken()).toBe('popup-id');
    expect(popup).toHaveBeenCalledWith({
      scopes: opts.idTokenScopes,
      loginHint: 'v.k@acme',
    });
  });

  it('uses popup first for a user-click interactive retry', async () => {
    const ssoSilent = vi.fn();
    const popup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'popup-acc', idToken: 'popup-id', account }),
    );
    const auth = new NaaAuthClient(
      fakeMsal({
        getActiveAccount: () => null,
        getAllAccounts: () => [],
        ssoSilent,
        acquireTokenPopup: popup,
      }),
      { ...opts, loginHint: 'v.k@acme', preferInteractive: true },
    );

    expect(await auth.getIdentity()).toMatchObject({ username: 'v.k@acme' });
    expect(popup).toHaveBeenCalledWith({
      scopes: opts.idTokenScopes,
      loginHint: 'v.k@acme',
      prompt: 'select_account',
    });
    expect(ssoSilent).not.toHaveBeenCalled();
  });

  it('prefers loginPopup with account selection for first interactive bootstrap', async () => {
    const loginPopup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'login-acc', idToken: 'login-id', account }),
    );
    const acquireTokenPopup = vi.fn();
    const auth = new NaaAuthClient(
      fakeMsal({
        getActiveAccount: () => null,
        getAllAccounts: () => [],
        loginPopup,
        acquireTokenPopup,
      }),
      { ...opts, loginHint: 'v.k@acme', preferInteractive: true },
    );

    expect(await auth.getIdentity()).toMatchObject({ username: 'v.k@acme' });
    expect(loginPopup).toHaveBeenCalledWith({
      scopes: opts.idTokenScopes,
      loginHint: 'v.k@acme',
      prompt: 'select_account',
    });
    expect(acquireTokenPopup).not.toHaveBeenCalled();
  });

  it('does not launch a second popup after interactive loginPopup times out', async () => {
    const timeout = Object.assign(new Error('timed_out'), {
      name: 'BrowserAuthError',
      errorCode: 'timed_out',
      subError: 'redirect_bridge_timeout',
    });
    const loginPopup = vi.fn(() => Promise.reject(timeout));
    const acquireTokenPopup = vi.fn();
    const auth = new NaaAuthClient(
      fakeMsal({
        getActiveAccount: () => null,
        getAllAccounts: () => [],
        loginPopup,
        acquireTokenPopup,
      }),
      { ...opts, loginHint: 'v.k@acme', preferInteractive: true },
    );

    await expect(auth.getIdentity()).rejects.toThrow('timed_out');
    expect(loginPopup).toHaveBeenCalledOnce();
    expect(acquireTokenPopup).not.toHaveBeenCalled();
  });

  it('passes MSAL popup override only for explicit interaction recovery', async () => {
    const popup = vi.fn(() =>
      Promise.resolve<MsalAuthResult>({ accessToken: 'popup-acc', idToken: 'popup-id', account }),
    );
    const auth = new NaaAuthClient(
      fakeMsal({
        getActiveAccount: () => null,
        getAllAccounts: () => [],
        acquireTokenPopup: popup,
      }),
      {
        ...opts,
        loginHint: 'v.k@acme',
        preferInteractive: true,
        overrideInteractionInProgress: true,
      },
    );

    expect(await auth.getIdToken()).toBe('popup-id');
    expect(popup).toHaveBeenCalledWith({
      scopes: opts.idTokenScopes,
      loginHint: 'v.k@acme',
      prompt: 'select_account',
      overrideInteractionInProgress: true,
    });
  });

  it('does not pop up for non-interaction errors', async () => {
    const popup = vi.fn();
    const auth = new NaaAuthClient(
      fakeMsal({
        acquireTokenSilent: () => Promise.reject(new Error('network')),
        acquireTokenPopup: popup,
      }),
      opts,
    );
    await expect(auth.getIdToken()).rejects.toThrow('network');
    expect(popup).not.toHaveBeenCalled();
  });

  it('derives identity from the active account', async () => {
    const auth = new NaaAuthClient(fakeMsal(), opts);
    expect(await auth.getIdentity()).toEqual({
      username: 'v.k@acme',
      displayName: 'V K',
      oid: 'oid-1',
    });
  });
});

describe('NaaAuthClient.getSilentAccessToken (error reports)', () => {
  it('asks silently for exactly the given scope and passes forceRefresh', async () => {
    const acquireTokenSilent = vi.fn((req: { scopes: string[] }) =>
      Promise.resolve<MsalAuthResult>({ accessToken: `acc:${req.scopes.join(',')}`, account }),
    );
    const auth = new NaaAuthClient(fakeMsal({ acquireTokenSilent }), opts);
    expect(await auth.getSilentAccessToken(['api://reports/x'], { forceRefresh: true })).toBe(
      'acc:api://reports/x',
    );
    expect(acquireTokenSilent).toHaveBeenCalledWith(
      expect.objectContaining({ scopes: ['api://reports/x'], account, forceRefresh: true }),
    );
  });

  it('never opens a popup when consent or sign-in is needed', async () => {
    const acquireTokenPopup = vi.fn();
    const loginPopup = vi.fn();
    const needsUi = Object.assign(new Error('consent_required'), {
      name: 'InteractionRequiredAuthError',
    });
    const auth = new NaaAuthClient(
      fakeMsal({
        acquireTokenSilent: () => Promise.reject(needsUi),
        acquireTokenPopup,
        loginPopup,
      }),
      { ...opts, preferInteractive: true },
    );
    await expect(auth.getSilentAccessToken(['api://reports/x'])).rejects.toThrow(
      'consent_required',
    );
    const signedOut = new NaaAuthClient(
      fakeMsal({ getActiveAccount: () => null, acquireTokenPopup, loginPopup }),
      opts,
    );
    await expect(signedOut.getSilentAccessToken(['api://reports/x'])).rejects.toThrow();
    expect(acquireTokenPopup).not.toHaveBeenCalled();
    expect(loginPopup).not.toHaveBeenCalled();
  });
});
