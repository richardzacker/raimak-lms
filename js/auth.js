// ============================================================
//  Raimak LMS — Authentication (MSAL)
// ============================================================

const Auth = (() => {

  let msalInstance = null;
  let currentAccount = null;

  // ── Init MSAL ──────────────────────────────────────────────
  function init() {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId:    Config.azure.clientId,
        authority:   `https://login.microsoftonline.com/${Config.azure.tenantId}`,
        redirectUri: Config.azure.redirectUri,
      },
      cache: {
        cacheLocation:          "sessionStorage",
        storeAuthStateInCookie: false,
      },
    });
    return msalInstance.handleRedirectPromise();
  }

  // ── Sign In ────────────────────────────────────────────────
  async function signIn() {
    try {
      await msalInstance.loginRedirect({ scopes: Config.scopes });
    } catch (err) {
      console.error("Sign-in error:", err);
      UI.showToast("Sign-in failed. Please try again.", "error");
    }
  }

  // ── Sign Out ───────────────────────────────────────────────
  function signOut() {
    msalInstance.logoutRedirect({ postLogoutRedirectUri: Config.azure.redirectUri });
  }

  // ── Get Access Token ───────────────────────────────────────
 async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return null;

  currentAccount = accounts[0];

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes:  Config.scopes,
      account: currentAccount,
    });
    return result.accessToken;
  } catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      // Don't throw after redirect — just redirect and let the page reload handle it
      await msalInstance.acquireTokenRedirect({ scopes: Config.scopes });
      return null; // ← this line replaces "throw err"
    }
    throw err;
  }
}
  // ── Current User ───────────────────────────────────────────
  function getUser() {
    const accounts = msalInstance?.getAllAccounts();
    if (!accounts?.length) return null;
    currentAccount = accounts[0];
    return {
      name:  currentAccount.name || currentAccount.username,
      email: currentAccount.username,
    };
  }

  function isSignedIn() {
    return !!(msalInstance?.getAllAccounts()?.length);
  }

  return { init, signIn, signOut, getToken, getUser, isSignedIn };
})();
