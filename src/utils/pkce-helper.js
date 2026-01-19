const PKCEHelper = {
  generateCodeVerifier() {
    const array = new Uint8Array(32);
    window.crypto.getRandomValues(array);
    return this.base64UrlEncode(array);
  },

  async generateCodeChallenge(codeVerifier) {
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    const hashBuffer = await window.crypto.subtle.digest('SHA-256', data);
    return this.base64UrlEncode(new Uint8Array(hashBuffer));
  },

  base64UrlEncode(buffer) {
    const base64 = btoa(String.fromCharCode.apply(null, buffer));
    return base64
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=/g, '');
  },

  generateState() {
    const array = new Uint8Array(16);
    window.crypto.getRandomValues(array);
    return this.base64UrlEncode(array);
  },

  storePKCEParams(codeVerifier, state) {
    try {
      sessionStorage.setItem('pkce_code_verifier', codeVerifier);
      sessionStorage.setItem('pkce_state', state);
    } catch (error) {
      throw new Error('Unable to store PKCE parameters securely');
    }
  },

  retrievePKCEParams() {
    try {
      const codeVerifier = sessionStorage.getItem('pkce_code_verifier');
      const state = sessionStorage.getItem('pkce_state');

      if (!codeVerifier || !state) {
        throw new Error('PKCE parameters not found');
      }

      return { codeVerifier, state };
    } catch (error) {
      throw new Error('Unable to retrieve PKCE parameters');
    }
  },

  retrieveCodeVerifier() {
    return sessionStorage.getItem('pkce_code_verifier');
  },

  clearPKCEParams() {
    try {
      sessionStorage.removeItem('pkce_code_verifier');
      sessionStorage.removeItem('pkce_state');
    } catch (error) {
    }
  },

  validateState(receivedState) {
    try {
      const storedState = sessionStorage.getItem('pkce_state');

      if (!storedState) {
        return false;
      }

      if (storedState !== receivedState) {
        return false;
      }

      return true;
    } catch (error) {
      return false;
    }
  }
};

export default PKCEHelper;
