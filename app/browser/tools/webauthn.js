/**
 * WebAuthn override, renderer half. Ported from teams-for-linux
 * app/browser/tools/webauthnOverride.js.
 *
 * Teams patches navigator.credentials straight from its preload, which only
 * works without context isolation. This window isolates its preload, so:
 * - on a sign-in origin, the patch is installed in the page world with
 *   contextBridge.executeInMainWorld and calls back through one bridged
 *   function per ceremony (the main process re-checks the calling frame's
 *   origin against the same allowlist before anything signs),
 * - the relay for sign-in iframes (patched by app/webauthn/index.js, Layer 2)
 *   listens for their postMessage from the isolated world, where ipcRenderer
 *   is, on every page: the iframe's host page may be Outlook itself.
 *
 * Linux-only: elsewhere Chromium handles WebAuthn natively.
 */

const { buildAllowedOrigins } = require("../../webauthn/originAllowlist");

const BRIDGE_KEY = "__oflWebAuthn";

/**
 * Runs in the page world (serialised by executeInMainWorld), so it must be
 * self-contained. Same serialisation and response shapes as the teams-for-linux
 * override, including the real prototypes Microsoft's login page checks with
 * instanceof (teams-for-linux #2719).
 */
function installOverride(bridgeKey) {
  const bridge = globalThis[bridgeKey];
  if (!bridge || !navigator.credentials?.create || !navigator.credentials?.get) return;

  const originalCreate = navigator.credentials.create.bind(navigator.credentials);
  const originalGet = navigator.credentials.get.bind(navigator.credentials);

  const toB64url = (buffer) => {
    if (!buffer) return "";
    const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    let binary = "";
    for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCodePoint(...bytes.subarray(i, i + 8192));
    return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/={1,2}$/, "");
  };
  const fromB64url = (text) => {
    let base64 = text.replaceAll("-", "+").replaceAll("_", "/");
    while (base64.length % 4 !== 0) base64 += "=";
    return Uint8Array.from(atob(base64), (c) => c.codePointAt(0)).buffer;
  };
  const withPrototype = (prototype, properties) => Object.create(prototype, Object.getOwnPropertyDescriptors(properties));
  const credentialList = (list) => (list || []).map((c) => ({ id: toB64url(c.id), type: c.type, transports: c.transports }));
  const timeoutSec = (ms) => (ms ? Math.floor(ms / 1000) : 60);

  const toError = (message) => {
    const msg = (message || "").toLowerCase();
    if (msg.includes("invalid")) return new DOMException(message, "InvalidStateError");
    if (msg.includes("securityerror")) return new DOMException(message, "SecurityError");
    return new DOMException(message || "The operation either timed out or was not allowed.", "NotAllowedError");
  };
  const call = async (operation, data) => {
    const result = await bridge[operation](data);
    if (!result?.success) throw toError(result?.error);
    return result.data;
  };

  navigator.credentials.create = async (options) => {
    if (!options?.publicKey) return originalCreate(options);
    const pk = options.publicKey;
    const r = await call("create", {
      challenge: toB64url(pk.challenge),
      rpId: pk.rp?.id || "",
      rpName: pk.rp?.name || "",
      userId: toB64url(pk.user?.id),
      userName: pk.user?.name || "",
      userDisplayName: pk.user?.displayName || "",
      pubKeyCredParams: pk.pubKeyCredParams,
      timeout: timeoutSec(pk.timeout),
      authenticatorSelection: pk.authenticatorSelection || {},
      attestation: pk.attestation || "none",
      excludeCredentials: credentialList(pk.excludeCredentials),
    });
    return withPrototype(PublicKeyCredential.prototype, {
      id: r.credentialId,
      rawId: fromB64url(r.rawId),
      type: r.type,
      authenticatorAttachment: "cross-platform",
      response: withPrototype(AuthenticatorAttestationResponse.prototype, {
        attestationObject: fromB64url(r.attestationObject),
        clientDataJSON: fromB64url(r.clientDataJson),
        getAuthenticatorData: () => fromB64url(r.authenticatorData),
        getTransports: () => r.transports || ["usb"],
        getPublicKey: () => null,
        getPublicKeyAlgorithm: () => r.publicKeyAlgorithm || -7,
      }),
      getClientExtensionResults: () => ({}),
      toJSON: () => ({
        id: r.credentialId,
        rawId: r.rawId,
        type: r.type,
        response: { attestationObject: r.attestationObject, clientDataJSON: r.clientDataJson },
      }),
    });
  };

  navigator.credentials.get = async (options) => {
    if (!options?.publicKey) return originalGet(options);
    // Passkey autofill probes (conditional mediation) stay native: routing
    // them would ask for the key's PIN on every sign-in page load.
    if (options.mediation === "conditional") return originalGet(options);
    const pk = options.publicKey;
    const r = await call("get", {
      challenge: toB64url(pk.challenge),
      rpId: pk.rpId || "",
      timeout: timeoutSec(pk.timeout),
      userVerification: pk.userVerification || "preferred",
      allowCredentials: credentialList(pk.allowCredentials),
    });
    const authenticatorData = fromB64url(r.authenticatorData);
    return withPrototype(PublicKeyCredential.prototype, {
      id: r.credentialId,
      rawId: fromB64url(r.rawId),
      type: r.type,
      authenticatorAttachment: "cross-platform",
      response: withPrototype(AuthenticatorAssertionResponse.prototype, {
        authenticatorData,
        clientDataJSON: fromB64url(r.clientDataJson),
        signature: fromB64url(r.signature),
        userHandle: r.userHandle ? fromB64url(r.userHandle) : null,
        getAuthenticatorData: () => authenticatorData,
      }),
      getClientExtensionResults: () => ({}),
      toJSON: () => ({
        id: r.credentialId,
        rawId: r.rawId,
        type: r.type,
        authenticatorAttachment: "cross-platform",
        clientExtensionResults: {},
        response: {
          authenticatorData: r.authenticatorData,
          clientDataJSON: r.clientDataJson,
          signature: r.signature,
          userHandle: r.userHandle || null,
        },
      }),
    });
  };
}

// Sign-in iframes patched by the main process post their requests to the top
// frame; relay them to the main process. event.origin is set by the browser,
// so it is the trustworthy record of which origin the key signs (teams-for-linux
// #2828), and the main process holds it to the allowlist again.
function installRelay(allowedOrigins, ipcRenderer) {
  window.addEventListener("message", async (event) => {
    if (event.data?.type !== "webauthn-request") return;
    if (!allowedOrigins.has(event.origin)) {
      console.warn("[WEBAUTHN] Blocked relay: origin not allowed. If this is your federated sign-in page, add it to webauthnExtraOrigins.");
      return;
    }
    const { id, channel, data } = event.data;
    if (channel !== "webauthn:create" && channel !== "webauthn:get") return;
    let reply;
    try {
      const result = await ipcRenderer.invoke(channel, { ...data, frameOrigin: event.origin });
      reply = result.success ? { result: result.data } : { error: result.error };
    } catch (err) {
      reply = { error: err.message };
    }
    event.source?.postMessage({ type: "webauthn-response", id, ...reply }, event.origin);
  });
}

/**
 * @param {{webauthn?: boolean, webauthnExtraOrigins?: string[]}} config
 * @param {Electron.IpcRenderer} ipcRenderer
 * @param {Electron.ContextBridge} contextBridge
 */
function init(config, ipcRenderer, contextBridge) {
  if (process.platform !== "linux" || !config?.webauthn || !contextBridge) return;
  const allowedOrigins = buildAllowedOrigins(config.webauthnExtraOrigins);
  installRelay(allowedOrigins, ipcRenderer);

  // The page-world patch only goes where a ceremony may run, so Outlook's own
  // pages never see the bridge.
  if (!allowedOrigins.has(location.origin)) return;
  contextBridge.exposeInMainWorld(BRIDGE_KEY, {
    create: (data) => ipcRenderer.invoke("webauthn:create", data),
    get: (data) => ipcRenderer.invoke("webauthn:get", data),
  });
  contextBridge.executeInMainWorld({ func: installOverride, args: [BRIDGE_KEY] });
  console.info("[WEBAUTHN] navigator.credentials patched for hardware security keys");
}

module.exports = { init, _internals: { installOverride, BRIDGE_KEY } };
