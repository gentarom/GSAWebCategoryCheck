
(function () {
  const { msalConfig, graphScopes } = window.__msalConfig__;

  // Create MSAL PublicClientApplication
  const msalInstance = new msal.PublicClientApplication(msalConfig);

  // UI elements
  const signinBtn   = document.getElementById("signin");
  const signoutBtn  = document.getElementById("signout");
  const checkBtn    = document.getElementById("checkBtn");
  const urlInput    = document.getElementById("urlInput");
  const resultPane  = document.getElementById("resultPane");
  const userInfoDiv = document.getElementById("user-info");
  const statusDiv   = document.getElementById("status");

  let account = null;

  // Helper: set UI state
  function setSignedInState(isSignedIn) {
    signinBtn.disabled  = isSignedIn;
    signoutBtn.disabled = !isSignedIn;
    checkBtn.disabled   = !isSignedIn;
    statusDiv.textContent = isSignedIn ? "Signed in. Ready to check web categories." : "Please sign in.";
  }

  // Handle redirect responses (if any)
  msalInstance.handleRedirectPromise().then((response) => {
    if (response && response.account) {
      account = response.account;
      msalInstance.setActiveAccount(account);
      userInfoDiv.textContent = `Signed in as: ${account.username}`;
      setSignedInState(true);
    } else {
      const current = msalInstance.getAllAccounts()[0];
      if (current) {
        account = current;
        msalInstance.setActiveAccount(account);
        userInfoDiv.textContent = `Signed in as: ${account.username}`;
        setSignedInState(true);
      } else {
        setSignedInState(false);
      }
    }
  }).catch((err) => {
    console.error(err);
    resultPane.textContent = JSON.stringify(err, null, 2);
  });

  // Sign-in
  signinBtn.addEventListener("click", async () => {
    try {
      await msalInstance.loginRedirect({
        scopes: graphScopes
      });
    } catch (err) {
      console.error(err);
      resultPane.textContent = JSON.stringify(err, null, 2);
    }
  });

  // Sign-out
  signoutBtn.addEventListener("click", async () => {
    try {
      await msalInstance.logoutRedirect();
    } catch (err) {
      console.error(err);
      resultPane.textContent = JSON.stringify(err, null, 2);
    }
  });

  // Acquire token silently or via redirect if needed
  async function getToken() {
    try {
      const active = msalInstance.getActiveAccount();
      if (!active) throw new Error("No active account. Please sign in.");

      const tokenResponse = await msalInstance.acquireTokenSilent({
        account: active,
        scopes: graphScopes
      });
      return tokenResponse.accessToken;
    } catch (silentErr) {
      // Fallback to interactive
      await msalInstance.acquireTokenRedirect({ scopes: graphScopes });
      // control returns after redirect; no token here
      throw new Error("Redirecting for token...");
    }
  }

  // Call Microsoft Graph (beta) – Network Access categorizeWebUrl action
  async function callCategorizeWebUrl(url) {
    const endpoint = "https://graph.microsoft.com/beta/networkAccess/categorizeWebUrl";

    // If you need to use the function-style GET you mentioned, you could use:
    // const endpoint = `https://graph.microsoft.com/beta/networkaccess/connectivity/microsoft.graph.networkaccess.getWebCategoryByUrl(url='${encodeURIComponent(url)}')`;

    const token = await getToken();

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      body: JSON.stringify({ url })
    });

    if (!response.ok) {
      const errBody = await response.text();
      throw new Error(`Graph error ${response.status}: ${errBody}`);
    }

    return response.json();
  }

  // On click: validate URL and call Graph
  checkBtn.addEventListener("click", async () => {
    const url = urlInput.value.trim();
    if (!url) {
      resultPane.textContent = "Please enter a URL.";
      return;
    }
    resultPane.textContent = "Checking category...";
    try {
      const result = await callCategorizeWebUrl(url);
      // Pretty print JSON result
      resultPane.textContent = JSON.stringify(result, null, 2);

      // Optional: Render a friendlier summary if the response has `categories`
      // if (result?.categories?.length) {
      //   const top = result.categories[0];
      //   resultPane.textContent = `Category: ${top.name} (confidence ${Math.round((top.confidence || 0) * 100)}%)`;
      // }
    } catch (err) {
      console.error(err);
      resultPane.textContent = err.message;
    }
  });
})();
