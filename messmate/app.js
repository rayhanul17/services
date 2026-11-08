(() => {
  "use strict";

  const STORE_KEY = "messmate.v1";
  const APP_SETTINGS_KEY = "messmate.app-settings.v1";
  const DEFAULT_APP_NAME = "MessMate";
  const SHEETS_TAB = "MessMate";
  const SHEETS_CHUNK_SIZE = 30000;
  const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
  const APP_IMAGE_TYPES = ["image/png", "image/jpeg", "image/webp", "image/x-icon", "image/vnd.microsoft.icon"];
  const CURRENCY_SYMBOL = "৳";
  const NAV = [
    { id: "dashboard", label: "Dashboard", icon: "grid" },
    { id: "meals", label: "Meals", icon: "meal" },
    { id: "bazaar", label: "Expenses", icon: "receipt" },
    { id: "deposits", label: "Deposits", icon: "wallet" },
    { id: "members", label: "Members", icon: "users" }
  ];
  const icons = {
    grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/>',
    meal: '<path d="M4 3v7a3 3 0 0 0 6 0V3M7 3v18M17 3v18m0-18c-2 2-3 5-3 8h6c0-3-1-6-3-8Z"/>',
    receipt: '<path d="M5 3h14v18l-3-2-4 2-4-2-3 2V3Z"/><path d="M8 8h8M8 12h8M8 16h3"/>',
    wallet: '<rect x="3" y="6" width="18" height="15" rx="2"/><path d="M3 9h18V5a2 2 0 0 0-2-2H6a3 3 0 0 0-3 3v3Zm12 5h.01"/>',
    users: '<path d="M16 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="10" cy="7" r="4"/><path d="M20 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    fork: '<path d="M6 3v6a3 3 0 0 0 6 0V3M9 3v18M18 3v18m0-18c-2 2-3 5-3 8h6c0-3-1-6-3-8Z"/>',
    coin: '<circle cx="12" cy="12" r="9"/><path d="M15 8.5c-.7-.7-1.6-1-3-1-1.7 0-3 .8-3 2s1.3 1.8 3 2 3 .8 3 2-1.3 2-3 2c-1.3 0-2.3-.4-3-1.1M12 5.5v13"/>',
    chart: '<path d="M3 3v18h18"/><path d="m8 14 4-4 4 3 5-7"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    lock: '<rect x="4" y="10" width="16" height="11" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
    download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5M12 15V3"/>'
  };
  const appSettings = { name: DEFAULT_APP_NAME, icon: "", favicon: "", fontScale: "1", googleClientId: "", spreadsheetId: "", syncRows: {} };
  const state = { mess: null, key: null, page: "dashboard", month: monthKey(new Date()), modal: null, toastTimer: null, skipNextSyncForMess: "" };
  const syncStatus = { kind: "offline", label: "Local only", message: "Changes are saved on this device. Connect Google Sheets to enable online sync." };
  let googleAccessToken = "";
  let googleTokenClient = null;
  let syncRunning = false;
  let syncRequested = false;
  const root = document.getElementById("app");
  const toast = document.getElementById("toast");

  function icon(name, size = 18) {
    return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] || ""}</svg>`;
  }
  function brandMark() {
    return appSettings.icon
      ? `<img src="${appSettings.icon}" alt="" />`
      : icon("fork", 21);
  }
  function applyAppSettings() {
    document.title = appSettings.name;
    document.documentElement.style.fontSize = `${16 * Number(appSettings.fontScale)}px`;
    let favicon = document.getElementById("app-favicon");
    if (!favicon) {
      favicon = document.createElement("link");
      favicon.id = "app-favicon";
      favicon.rel = "icon";
      document.head.append(favicon);
    }
    favicon.href = appSettings.favicon || `data:image/svg+xml,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="8" fill="#277354"/><path d="M10 5v8a4 4 0 0 0 8 0V5m-4 0v22m9-22v22m0-22c-3 3-4 7-4 11h8c0-4-1-8-4-11Z" fill="none" stroke="#fff" stroke-linecap="round" stroke-width="2"/></svg>')}`;
  }
  function readAppSettings() {
    let parsed;
    try { parsed = JSON.parse(localStorage.getItem(APP_SETTINGS_KEY) || "{}"); }
    catch (error) { throw new Error(`App settings are invalid and could not be loaded. ${error.message}`); }
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("App settings have an invalid format.");
    const name = parsed.name === undefined ? DEFAULT_APP_NAME : parsed.name;
    const fontScale = parsed.fontScale === undefined ? "1" : String(parsed.fontScale);
    const googleClientId = parsed.googleClientId === undefined ? "" : parsed.googleClientId;
    const spreadsheetId = parsed.spreadsheetId === undefined ? "" : parsed.spreadsheetId;
    const syncRows = parsed.syncRows === undefined ? {} : parsed.syncRows;
    const imagePattern = /^data:image\/(?:png|jpeg|webp|x-icon|vnd\.microsoft\.icon);base64,[a-z0-9+/]+=*$/i;
    if (typeof name !== "string" || !name.trim() || name.trim().length > 40) throw new Error("App name in settings must be between 1 and 40 characters.");
    if (!validFontScale(fontScale)) throw new Error("Font size setting is invalid.");
    if (typeof googleClientId !== "string" || googleClientId.length > 300 || typeof spreadsheetId !== "string" || spreadsheetId.length > 200) throw new Error("Google Sheets settings are invalid.");
    if (!syncRows || typeof syncRows !== "object" || Array.isArray(syncRows)) throw new Error("Google Sheets sync history is invalid.");
    for (const [messId, syncRow] of Object.entries(syncRows)) {
      if (!messId || !syncRow || !Number.isInteger(syncRow.startRow) || syncRow.startRow < 2 || !Number.isInteger(syncRow.rowCount) || syncRow.rowCount < 1 || typeof syncRow.spreadsheetId !== "string") {
        throw new Error("A saved Google Sheets sync location is invalid.");
      }
    }
    for (const key of ["icon", "favicon"]) {
      if (parsed[key] !== undefined && parsed[key] !== "" && (typeof parsed[key] !== "string" || parsed[key].length > 400000 || !imagePattern.test(parsed[key]))) {
        throw new Error(`The saved ${key} image is invalid.`);
      }
    }
    return { name: name.trim(), icon: parsed.icon || "", favicon: parsed.favicon || "", fontScale, googleClientId, spreadsheetId, syncRows };
  }
  function saveAppSettings(next) {
    const merged = { ...appSettings, ...next };
    try { localStorage.setItem(APP_SETTINGS_KEY, JSON.stringify(merged)); }
    catch (error) { throw new Error(`Could not save app settings. Browser storage may be full or disabled. ${error.message}`); }
    Object.assign(appSettings, merged);
    applyAppSettings();
  }
  function validFontScale(value) {
    const scale = Number(value);
    return Number.isFinite(scale) && scale >= 0.8 && scale <= 1.3 && Math.abs(scale * 20 - Math.round(scale * 20)) < 0.000001;
  }
  function validDate(value) {
    if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const date = new Date(`${value}T00:00:00Z`);
    return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
  }
  function validateMessData(data) {
    if (!data || !Array.isArray(data.members) || !Array.isArray(data.meals) || !Array.isArray(data.expenses) || !Array.isArray(data.deposits)) {
      throw new Error("The saved mess data is incomplete or invalid.");
    }
    const memberIds = new Set();
    data.members.forEach(member => {
      if (!member || typeof member.id !== "string" || !member.id || typeof member.name !== "string" || !member.name.trim() || member.name.trim().length > 60 || memberIds.has(member.id) || (member.active !== undefined && typeof member.active !== "boolean")) {
        throw new Error("A saved member is invalid. Check the imported data or backup file.");
      }
      memberIds.add(member.id);
    });
    const hasMember = item => typeof item.memberId === "string" && memberIds.has(item.memberId);
    const validateRecordIds = records => {
      const ids = records.map(item => item?.id);
      if (ids.some(recordId => typeof recordId !== "string" || !recordId) || new Set(ids).size !== ids.length) throw new Error("Saved records contain a missing or duplicate ID.");
    };
    validateRecordIds(data.meals);
    validateRecordIds(data.expenses);
    validateRecordIds(data.deposits);
    data.meals.forEach(item => {
      const quantity = Number(item?.quantity);
      if (!item || typeof item.id !== "string" || !item.id || !hasMember(item) || !validDate(item.date) || typeof item.type !== "string" || !item.type.trim() || !Number.isFinite(quantity) || quantity <= 0 || quantity > 10000) {
        throw new Error("A saved meal is invalid. Verify dates, member IDs, and meal quantities.");
      }
    });
    data.expenses.forEach(item => {
      const amount = Number(item?.amount);
      if (!item || typeof item.id !== "string" || !item.id || !validDate(item.date) || !["bazaar", "fixed"].includes(item.kind) || !Number.isFinite(amount) || amount <= 0 || amount > 1000000000) {
        throw new Error("A saved expense is invalid. Verify its date, type, and amount.");
      }
    });
    data.deposits.forEach(item => {
      const amount = Number(item?.amount);
      if (!item || typeof item.id !== "string" || !item.id || !hasMember(item) || !validDate(item.date) || !Number.isFinite(amount) || amount === 0 || Math.abs(amount) > 1000000000) {
        throw new Error("A saved deposit is invalid. Verify its date, member, and amount.");
      }
    });
    return data;
  }
  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch]);
  }
  function id() {
    return crypto.randomUUID ? crypto.randomUUID() : `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`;
  }
  function monthKey(date) {
    if (typeof date === "string" && /^\d{4}-\d{2}/.test(date)) return date.slice(0, 7);
    const d = new Date(date);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  function monthLabel(key) {
    const [year, month] = key.split("-").map(Number);
    return new Intl.DateTimeFormat(undefined, { month: "long", year: "numeric" }).format(new Date(year, month - 1, 1));
  }
  function currentDate() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  function entryDate() {
    if (monthKey(new Date()) === state.month) return currentDate();
    const [year, month] = state.month.split("-").map(Number);
    const day = Math.min(new Date().getDate(), new Date(year, month, 0).getDate());
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  function dateLabel(date, options = { month: "short", day: "numeric" }) {
    const d = new Date(`${date}T12:00:00`);
    return Number.isNaN(d.getTime()) ? String(date || "") : new Intl.DateTimeFormat(undefined, options).format(d);
  }
  function money(value) {
    const amount = Number(value || 0);
    const formatted = Math.abs(amount).toLocaleString("en-BD", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return `${amount < 0 ? "−" : ""}${formatted}\u00a0${CURRENCY_SYMBOL}`;
  }
  function sum(values) { return values.reduce((total, value) => total + Number(value || 0), 0); }
  function monthRecords(records) { return records.filter(record => monthKey(record.date) === state.month); }
  function getIndex() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORE_KEY) || "[]");
      if (!Array.isArray(parsed) || parsed.some(item => !item || typeof item.id !== "string" || !item.id || typeof item.name !== "string" || !item.name.trim() || item.name.trim().length > 60 || typeof item.salt !== "string" || typeof item.iv !== "string" || typeof item.payload !== "string") || new Set(parsed.map(item => item.id)).size !== parsed.length || new Set(parsed.map(item => item.name.trim().toLocaleLowerCase())).size !== parsed.length) {
        throw new Error("The local mess list has an invalid format.");
      }
      return parsed;
    } catch (error) {
      throw new Error(`Could not read saved messes from this browser. ${error.message}`);
    }
  }
  function saveIndex(index) {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(index)); }
    catch (error) { throw new Error(`Could not save this mess. Your browser storage may be full or disabled. ${error.message}`); }
  }
  async function deriveKey(password, salt) {
    const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
    return crypto.subtle.deriveKey({ name: "PBKDF2", salt, iterations: 210000, hash: "SHA-256" }, material, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  }
  async function encryptData(data, key) {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const payload = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(JSON.stringify(data)));
    return { iv: toBase64(iv), payload: toBase64(new Uint8Array(payload)) };
  }
  async function decryptData(entry, key) {
    const decoded = await crypto.subtle.decrypt({ name: "AES-GCM", iv: fromBase64(entry.iv) }, key, fromBase64(entry.payload));
    return validateMessData(JSON.parse(new TextDecoder().decode(decoded)));
  }
  function toBase64(bytes) {
    let binary = "";
    bytes.forEach(byte => { binary += String.fromCharCode(byte); });
    return btoa(binary);
  }
  function fromBase64(text) {
    return Uint8Array.from(atob(text), char => char.charCodeAt(0));
  }
  async function persistMess() {
    if (!state.mess || !state.key) throw new Error("No mess is currently unlocked.");
    const index = getIndex();
    const entry = index.find(item => item.id === state.mess.id);
    if (!entry) throw new Error("This mess is no longer in this browser's saved list.");
    const encrypted = await encryptData(state.mess, state.key);
    entry.iv = encrypted.iv;
    entry.payload = encrypted.payload;
    saveIndex(index);
    setSyncStatus("pending", "Syncing changes to Google Sheets…");
    void syncCurrentMess();
  }
  function setSyncStatus(kind, message) {
    syncStatus.kind = kind;
    syncStatus.label = kind === "online" ? "Online" : kind === "pending" ? "Syncing" : kind === "offline" ? "Offline" : "Local only";
    syncStatus.message = message;
    const indicator = document.getElementById("sync-indicator");
    if (indicator) {
      indicator.className = `sync-indicator sync-${kind}`;
      indicator.title = message;
      indicator.setAttribute("aria-label", `Google Sheets sync: ${syncStatus.label}. ${message}`);
      indicator.querySelector(".sync-label").textContent = syncStatus.label;
    }
    const detail = document.getElementById("sync-status-detail");
    if (detail) detail.textContent = message;
  }
  function sheetsApiUrl(range, method) {
    const sheetId = encodeURIComponent(appSettings.spreadsheetId);
    const encodedRange = encodeURIComponent(range);
    const suffix = method === "append" ? ":append?valueInputOption=RAW&insertDataOption=INSERT_ROWS" : "?valueInputOption=RAW";
    return `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodedRange}${suffix}`;
  }
  async function sheetsRequest(url, method, values) {
    const response = await fetch(url, {
      method,
      headers: { Authorization: `Bearer ${googleAccessToken}`, "Content-Type": "application/json" },
      body: values ? JSON.stringify({ majorDimension: "ROWS", values }) : undefined
    });
    let result;
    try { result = await response.json(); }
    catch (error) { throw new Error(`Google Sheets returned an unreadable response. ${error.message}`); }
    if (!response.ok) {
      if (response.status === 401) googleAccessToken = "";
      throw new Error(result?.error?.message || `Google Sheets rejected the update (HTTP ${response.status}).`);
    }
    return result;
  }
  async function ensureGoogleAccessToken() {
    if (googleAccessToken) return;
    if (!navigator.onLine) throw new Error("You are offline. Connect to the internet to access Google Sheets.");
    if (!window.google?.accounts?.oauth2) throw new Error("Google sign-in could not load. Check your internet connection and reload the app.");
    const tokenResponse = await new Promise((resolve, reject) => {
      googleTokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: appSettings.googleClientId,
        scope: SHEETS_SCOPE,
        callback: response => {
          if (response.error) {
            reject(new Error(response.error_description || `Google sign-in failed: ${response.error}`));
            return;
          }
          googleAccessToken = response.access_token;
          resolve(response);
        },
        error_callback: error => reject(new Error(error.message || "Google sign-in was blocked or failed."))
      });
      googleTokenClient.requestAccessToken({ prompt: "" });
    });
    if (!tokenResponse?.access_token) throw new Error("Google did not return an access token.");
  }
  async function readSheetBackups() {
    const range = encodeURIComponent(`${SHEETS_TAB}!A2:H`);
    const response = await sheetsRequest(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(appSettings.spreadsheetId)}/values/${range}`,
      "GET"
    );
    if (response?.values !== undefined && !Array.isArray(response.values)) throw new Error("The Google Sheet backup data has an invalid format.");
    const grouped = new Map();
    (response.values || []).forEach((row, index) => {
      if (!Array.isArray(row) || row.length < 8) return;
      const [messId, chunkNumber, totalChunks, name, updatedAt, salt, iv, payload] = row;
      const chunk = Number(chunkNumber);
      const chunks = Number(totalChunks);
      if (typeof messId !== "string" || !messId || typeof name !== "string" || !name.trim() || name.length > 60 ||
          !Number.isInteger(chunk) || chunk < 1 || !Number.isInteger(chunks) || chunks < 1 || chunks > 10000 || chunk > chunks ||
          typeof updatedAt !== "string" || !Number.isFinite(Date.parse(updatedAt)) ||
          typeof salt !== "string" || typeof iv !== "string" || typeof payload !== "string") return;
      const key = `${messId}|${updatedAt}|${salt}|${iv}|${chunks}`;
      if (!grouped.has(key)) grouped.set(key, { id: messId, name: name.trim(), updatedAt, salt, iv, totalChunks: chunks, parts: new Map(), positions: new Map(), startRow: index + 2 });
      const backup = grouped.get(key);
      if (backup.name !== name.trim() || backup.parts.has(chunk)) return;
      backup.parts.set(chunk, payload);
      backup.positions.set(chunk, index + 2);
    });
    return [...grouped.values()].filter(item => item.parts.size === item.totalChunks &&
      Array.from({ length: item.totalChunks }, (_, index) => item.positions.get(index + 1) === item.startRow + index).every(Boolean))
      .map(item => ({
        id: item.id,
        name: item.name,
        updatedAt: item.updatedAt,
        salt: item.salt,
        iv: item.iv,
        payload: Array.from({ length: item.totalChunks }, (_, index) => item.parts.get(index + 1)).join(""),
        startRow: item.startRow,
        rowCount: item.totalChunks,
        spreadsheetId: appSettings.spreadsheetId
      }))
      .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
  }
  function showRestoreResults(backups) {
    const target = document.getElementById("restore-results");
    if (!target) return;
    if (!backups.length) {
      target.innerHTML = `<p class="sync-form-error">No complete MessMate backups were found in this spreadsheet. Check that the tab is named MessMate and contains the expected rows.</p>`;
      return;
    }
    target.innerHTML = `<h2>Available Backups</h2><p>Select a snapshot to copy into this device. Existing local records will not be overwritten. You will unlock it with the original mess password.</p><div class="restore-list">${backups.map((backup, index) => `<button class="restore-item" type="button" data-action="restore-sheet-backup" data-index="${index}"><strong>${escapeHtml(backup.name)}</strong><span>${escapeHtml(new Date(backup.updatedAt).toLocaleString())}</span><small>Encrypted snapshot · ${backup.rowCount} chunk${backup.rowCount === 1 ? "" : "s"}</small></button>`).join("")}</div>`;
  }
  let availableBackups = [];
  async function findSheetBackups(form) {
    const values = new FormData(form);
    const clientId = String(values.get("google-client-id") || "").trim();
    let spreadsheetId = String(values.get("spreadsheet-id") || "").trim();
    const match = spreadsheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (match) spreadsheetId = match[1];
    if (!/^[\w.-]+\.apps\.googleusercontent\.com$/.test(clientId) || !/^[a-zA-Z0-9_-]{20,}$/.test(spreadsheetId)) {
      throw new Error("Enter a valid Google OAuth client ID and spreadsheet ID or spreadsheet URL.");
    }
    saveAppSettings({ googleClientId: clientId, spreadsheetId });
    await ensureGoogleAccessToken();
    availableBackups = await readSheetBackups();
    showRestoreResults(availableBackups);
  }
  function restoreSheetBackup(indexValue) {
    const backup = availableBackups[Number(indexValue)];
    if (!backup) throw new Error("Choose a backup from the current Google Sheet results.");
    const entry = { id: backup.id, name: backup.name, salt: backup.salt, iv: backup.iv, payload: backup.payload };
    const localIndex = getIndex();
    const sameMess = localIndex.find(item => item.id === entry.id);
    if (sameMess) {
      throw new Error("This mess is already saved on this device. Unlock it instead; restoring will not overwrite local records.");
    }
    if (localIndex.some(item => item.name.trim().toLocaleLowerCase() === entry.name.toLocaleLowerCase())) {
      throw new Error("A different local mess already uses this name. Rename or remove that mess before restoring.");
    }
    localIndex.push(entry);
    saveIndex(localIndex);
    state.skipNextSyncForMess = entry.id;
    availableBackups = [];
    renderAuth("choose");
    showToast(`Encrypted backup for ${entry.name} restored. Unlock it with the original mess password.`);
  }
  async function uploadEncryptedMess(entry) {
    const chunks = [];
    for (let offset = 0; offset < entry.payload.length; offset += SHEETS_CHUNK_SIZE) {
      chunks.push(entry.payload.slice(offset, offset + SHEETS_CHUNK_SIZE));
    }
    if (!chunks.length) chunks.push("");
    const updatedAt = new Date().toISOString();
    const rows = chunks.map((chunk, index) => [
      entry.id, index + 1, chunks.length, entry.name, updatedAt, entry.salt, entry.iv, chunk
    ]);
    const previous = appSettings.syncRows[entry.id];
    if (previous?.spreadsheetId === appSettings.spreadsheetId) {
      const endRow = previous.startRow + Math.max(previous.rowCount, rows.length) - 1;
      while (rows.length < previous.rowCount) rows.push(["", "", "", "", "", "", "", ""]);
      await sheetsRequest(
        sheetsApiUrl(`${SHEETS_TAB}!A${previous.startRow}:H${endRow}`, "update"),
        "PUT",
        rows
      );
      saveAppSettings({
        syncRows: { ...appSettings.syncRows, [entry.id]: { startRow: previous.startRow, rowCount: chunks.length, spreadsheetId: appSettings.spreadsheetId } }
      });
      return;
    }
    const result = await sheetsRequest(sheetsApiUrl(`${SHEETS_TAB}!A:H`, "append"), "POST", rows);
    const updatedRange = result?.updates?.updatedRange || "";
    const match = updatedRange.match(/!A(\d+):H(\d+)$/);
    if (!match) throw new Error("Google Sheets accepted the upload but did not return its saved row range.");
    saveAppSettings({
      syncRows: { ...appSettings.syncRows, [entry.id]: { startRow: Number(match[1]), rowCount: chunks.length, spreadsheetId: appSettings.spreadsheetId } }
    });
  }
  async function syncCurrentMess() {
    if (!state.mess) return;
    if (syncRunning) { syncRequested = true; return; }
    syncRunning = true;
    try {
      do {
        syncRequested = false;
        if (!appSettings.googleClientId || !appSettings.spreadsheetId) {
          setSyncStatus("local", "Changes are saved on this device only. Add your Google settings to enable sync.");
          return;
        }
        if (!navigator.onLine) {
          setSyncStatus("offline", "You are offline. Changes are saved locally and will sync when connected.");
          return;
        }
        if (!googleAccessToken) {
          setSyncStatus("offline", "Google sign-in is required. Reconnect in Mess settings to sync pending changes.");
          return;
        }
        const messId = state.mess.id;
        const entry = getIndex().find(item => item.id === messId);
        if (!entry) throw new Error("This mess is no longer in this browser's saved list.");
        setSyncStatus("pending", `Uploading encrypted data for ${entry.name}…`);
        await uploadEncryptedMess(entry);
        setSyncStatus("online", `Google confirmed the latest save for ${entry.name} at ${new Date().toLocaleTimeString()}.`);
      } while (syncRequested);
    } catch (error) {
      setSyncStatus("offline", `Online sync failed; your local save is safe. ${error.message || "Try reconnecting in Mess settings."}`);
    } finally {
      syncRunning = false;
      if (syncRequested) void syncCurrentMess();
    }
  }
  async function connectGoogleSheets() {
    try {
      if (!appSettings.googleClientId || !appSettings.spreadsheetId) throw new Error("Save a Google OAuth client ID and spreadsheet ID in Mess settings first.");
      await ensureGoogleAccessToken();
      await syncCurrentMess();
      if (syncStatus.kind !== "online") throw new Error(syncStatus.message);
      renderApp();
      showToast("Google Sheets connected. Your encrypted mess data is synced.");
    } catch (error) {
      setSyncStatus("offline", error.message || "Google sign-in or sync failed. Your local save is safe.");
      throw error;
    }
  }
  function saveSyncSettings(form) {
    const clientId = String(new FormData(form).get("google-client-id") || "").trim();
    let spreadsheetId = String(new FormData(form).get("spreadsheet-id") || "").trim();
    const match = spreadsheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (match) spreadsheetId = match[1];
    if ((clientId && !/^[\w.-]+\.apps\.googleusercontent\.com$/.test(clientId)) || (spreadsheetId && !/^[a-zA-Z0-9_-]{20,}$/.test(spreadsheetId))) {
      throw new Error("Enter a valid Google OAuth client ID and spreadsheet ID or spreadsheet URL.");
    }
    if (Boolean(clientId) !== Boolean(spreadsheetId)) throw new Error("Enter both the Google OAuth client ID and spreadsheet ID.");
    const changed = clientId !== appSettings.googleClientId || spreadsheetId !== appSettings.spreadsheetId;
    saveAppSettings({ googleClientId: clientId, spreadsheetId });
    if (changed) googleAccessToken = "";
    setSyncStatus("local", clientId ? "Settings saved. Connect your Google Account to start syncing." : "Google Sheets sync is disabled. Your data remains local.");
    renderApp();
    showToast(clientId ? "Google Sheets settings saved." : "Google Sheets sync disconnected.");
  }
  async function syncNow() {
    if (!googleAccessToken) {
      await connectGoogleSheets();
      return;
    }
    await syncCurrentMess();
    if (syncStatus.kind !== "online") throw new Error(syncStatus.message);
    renderApp();
    showToast("Google Sheets is up to date.");
  }
  function showToast(message) {
    toast.textContent = message;
    toast.classList.add("show");
    clearTimeout(state.toastTimer);
    state.toastTimer = setTimeout(() => toast.classList.remove("show"), 2800);
  }
  function initials(name) {
    return String(name).trim().split(/\s+/).slice(0, 2).map(part => part[0] || "").join("").toUpperCase() || "?";
  }
  function activeMembers() { return state.mess.members.filter(member => member.active !== false); }
  function memberName(memberId) {
    return state.mess.members.find(member => member.id === memberId)?.name || "Removed member";
  }
  function computed() {
    const meals = monthRecords(state.mess.meals);
    const expenses = monthRecords(state.mess.expenses);
    const deposits = monthRecords(state.mess.deposits);
    const bazaar = sum(expenses.filter(item => item.kind === "bazaar").map(item => item.amount));
    const fixed = sum(expenses.filter(item => item.kind === "fixed").map(item => item.amount));
    const mealCount = sum(meals.map(item => item.quantity));
    const depositTotal = sum(deposits.map(item => item.amount));
    const mealRate = mealCount > 0 ? bazaar / mealCount : 0;
    const participants = new Set(activeMembers().map(member => member.id));
    [...meals, ...deposits].forEach(record => participants.add(record.memberId));
    const participantCount = participants.size;
    const balances = state.mess.members.map(member => {
      const memberRecords = meals.filter(item => item.memberId === member.id);
      const memberMeals = sum(memberRecords.map(item => item.quantity));
      const mealDays = new Set(memberRecords.filter(item => Number(item.quantity) > 0).map(item => item.date)).size;
      const memberDeposits = sum(deposits.filter(item => item.memberId === member.id).map(item => item.amount));
      const mealCost = memberMeals * mealRate;
      const fixedShare = participantCount && participants.has(member.id) ? fixed / participantCount : 0;
      const cost = mealCost + fixedShare;
      return { member, meals: memberMeals, mealDays, deposits: memberDeposits, mealCost, fixedShare, cost, net: memberDeposits - cost };
    });
    return {
      meals, expenses, deposits, bazaar, fixed, mealCount, depositTotal, mealRate,
      participants: participantCount, balances, remaining: depositTotal - bazaar - fixed
    };
  }
  function memberBalanceLabel(net) {
    if (Math.abs(net) < .005) return { text: "Settled", className: "neutral" };
    return net > 0 ? { text: "Credit", className: "positive" } : { text: "Debit", className: "negative" };
  }
  function empty(title, detail) {
    return `<div class="empty-state"><strong>${escapeHtml(title)}</strong>${escapeHtml(detail)}</div>`;
  }
  function unallocatedBazaarNotice(data) {
    return data.bazaar > 0 && data.mealCount === 0
      ? `<div class="notice allocation-notice" role="status">Bazaar expenses are included in the cash balance but cannot be assigned to member meal costs until meals are recorded for this month.</div>`
      : "";
  }
  function monthSelect() {
    const now = new Date();
    const keys = new Set([state.month]);
    for (let offset = -11; offset <= 1; offset++) keys.add(monthKey(new Date(now.getFullYear(), now.getMonth() + offset, 1)));
    [...state.mess.meals, ...state.mess.expenses, ...state.mess.deposits].forEach(record => {
      if (record.date) keys.add(monthKey(record.date));
    });
    const options = [...keys].sort().reverse().map(key =>
      `<option value="${key}" ${key === state.month ? "selected" : ""}>${escapeHtml(monthLabel(key))}</option>`
    );
    return `<select class="month-select" id="month-select" aria-label="Select reporting month" aria-describedby="month-filter-hint">${options.join("")}</select>`;
  }
  function header() {
    return `<header class="topbar">
      <div class="breadcrumb"><span>${escapeHtml(appSettings.name)}&nbsp; / &nbsp;</span><strong>${escapeHtml(NAV.find(item => item.id === state.page)?.label || "Settings")}</strong></div>
      <div class="top-actions"><div class="month-filter">${monthSelect()}<small id="month-filter-hint">Month filter only · saved entries stay intact</small></div><span id="sync-indicator" class="sync-indicator sync-${syncStatus.kind}" role="status" title="${escapeHtml(syncStatus.message)}" aria-label="Google Sheets sync: ${escapeHtml(syncStatus.label)}. ${escapeHtml(syncStatus.message)}"><span class="sync-dot"></span><span class="sync-label">${escapeHtml(syncStatus.label)}</span></span><button class="icon-button" data-action="settings" aria-label="Mess Settings" title="Mess Settings">${icon("lock", 16)}</button></div>
    </header>`;
  }
  function sidebar() {
    return `<aside class="sidebar">
      <div class="brand"><div class="brand-mark">${brandMark()}</div><div class="brand-name">${escapeHtml(appSettings.name)}</div></div>
      <button class="mess-switcher" data-action="switch-mess">
        <span class="mess-avatar">${escapeHtml(initials(state.mess.name))}</span>
        <span class="mess-switcher-copy"><small>Your Mess</small><strong>${escapeHtml(state.mess.name)}</strong></span>
        <span aria-hidden="true">⌄</span>
      </button>
      <div class="nav-label">Workspace</div><nav class="nav-list">${navItems()}</nav>
    </aside>`;
  }
  function navItems() {
    return NAV.map(item => `<button class="nav-item ${state.page === item.id ? "active" : ""}" data-page="${item.id}">${icon(item.icon)}<span>${item.label}</span>${item.id === "members" ? `<span class="nav-count">${activeMembers().length}</span>` : ""}</button>`).join("");
  }
  function bottomNav() {
    return `<nav class="bottom-nav" aria-label="Main navigation">${navItems()}</nav>`;
  }
  function pageHeading(title, subtitle, action = "") {
    return `<div class="page-heading"><div><div class="eyebrow">${escapeHtml(monthLabel(state.month))}</div><h1>${title}</h1>${subtitle ? `<p>${subtitle}</p>` : ""}</div>${action}</div>`;
  }
  function stat(label, value, detail, iconName, color = "") {
    return `<article class="card stat-card"><div class="stat-top"><span>${label}</span><span class="stat-icon ${color}">${icon(iconName, 16)}</span></div><div class="stat-value">${value}</div><div class="stat-foot">${detail}</div></article>`;
  }
  function dashboard() {
    const data = computed();
    const recent = [...data.expenses.map(item => ({ ...item, type: item.kind, title: item.description || (item.kind === "fixed" ? "Fixed cost" : "Bazaar expense"), amount: -Number(item.amount) })),
      ...data.deposits.map(item => ({ ...item, type: "deposit", title: `${memberName(item.memberId)} · deposit`, amount: Number(item.amount) }))]
      .sort((a, b) => b.date.localeCompare(a.date)).slice(0, 5);
    const members = data.balances.filter(item => item.member.active !== false || item.meals || item.deposits)
      .sort((a, b) => Math.abs(b.net) - Math.abs(a.net));
    const chart = chartMarkup(data);
    return `${pageHeading("Dashboard", "",
      `<div class="heading-actions"><button class="secondary-button" data-action="print-report">${icon("download", 15)} Download report</button><button class="primary-button" data-action="quick-add"><span class="button-icon">＋</span> Add entry</button></div>`)}
      <section class="grid stats-grid dashboard-stats">
        ${stat("Total deposits", money(data.depositTotal), `${data.deposits.length} deposit${data.deposits.length === 1 ? "" : "s"} this month`, "wallet")}
        ${stat("Bazaar expenses", money(data.bazaar), `${data.mealCount.toLocaleString()} meals recorded`, "receipt", "orange")}
        ${stat("Fixed costs", money(data.fixed), `Shared by ${data.participants} participant${data.participants === 1 ? "" : "s"}`, "coin", "blue")}
        ${stat("Meal rate", money(data.mealRate), "Bazaar cost per meal", "fork", "blue")}
        ${stat("Cash balance", money(data.remaining), data.remaining < 0 ? "More spent than deposited" : "Deposits less all expenses", "chart", data.remaining < 0 ? "red" : "")}
      </section>
      ${unallocatedBazaarNotice(data)}
      <section class="grid dashboard-grid">
        <div class="grid">
          <section class="card panel">
            <div class="panel-header"><div><h2>Member settlements</h2><p>Live totals · meals and deposits in ${escapeHtml(monthLabel(state.month))}</p></div><button class="text-button" data-page="members">View all →</button></div>
            <div class="panel-body">${members.length ? members.slice(0, 6).map(item => {
              const balance = memberBalanceLabel(item.net);
              return `<div class="member-row"><span class="member-avatar">${escapeHtml(initials(item.member.name))}</span><span class="member-info"><strong>${escapeHtml(item.member.name)}${item.member.active === false ? ' <span class="badge badge-archived">Removed</span>' : ""}</strong><small>${item.meals.toLocaleString()} meals · ${item.mealDays} days · Meal ${money(item.mealCost)} + fixed ${money(item.fixedShare)}</small></span><span class="member-balance ${balance.className}">${money(Math.abs(item.net))}<small>${balance.text}</small></span></div>`;
            }).join("") : empty("No members yet", "Add members to start tracking meals and settlements.")}</div>
          </section>
          <section class="card panel section-spacer">
            <div class="panel-header"><div><h2>Recent activity</h2><p>Latest entries for ${escapeHtml(monthLabel(state.month))}</p></div><button class="text-button" data-page="bazaar">See ledger →</button></div>
            <div class="panel-body activity-list">${recent.length ? recent.map(item => `<div class="activity-item"><span class="activity-symbol">${item.type === "deposit" ? "↙" : item.type === "fixed" ? "⌂" : "↗"}</span><span class="activity-copy"><strong>${escapeHtml(item.title)}</strong><small>${escapeHtml(dateLabel(item.date))}${item.note ? ` · ${escapeHtml(item.note)}` : ""}</small></span><span class="activity-money"><small class="${item.amount >= 0 ? "positive" : "negative"}">${item.amount >= 0 ? "Credit" : "Debit"}</small><strong class="${item.amount >= 0 ? "positive" : "negative"}">${money(Math.abs(item.amount))}</strong></span></div>`).join("") : empty("Your ledger starts here", "Add a deposit, meal or expense to see activity.")}</div>
          </section>
        </div>
        <div class="grid">
          <section class="card panel"><div class="panel-header"><div><h2>Spending overview</h2><p>This month's expenses</p></div></div>${chart}</section>
          <section class="card panel section-spacer"><div class="panel-header"><div><h2>Mess snapshot</h2><p>Current month totals</p></div></div><div class="panel-body">
            <div class="member-row"><span class="member-info"><strong>Fixed costs</strong><small>Split between ${data.participants} participant${data.participants === 1 ? "" : "s"}</small></span><span class="member-balance">${money(data.fixed)}</span></div>
            <div class="member-row"><span class="member-info"><strong>Total meals</strong><small>${activeMembers().length} active member${activeMembers().length === 1 ? "" : "s"}</small></span><span class="member-balance">${data.mealCount.toLocaleString()}</span></div>
            <div class="member-row"><span class="member-info"><strong>Expense total</strong><small>Bazaar + fixed costs</small></span><span class="member-balance">${money(data.bazaar + data.fixed)}</span></div>
          </div></section>
        </div>
      </section>`;
  }
  function chartMarkup(data) {
    const expenses = data.expenses;
    const [year, month] = state.month.split("-").map(Number);
    const lastDay = monthKey(new Date()) === state.month ? new Date() : new Date(year, month, 0);
    const daily = Array.from({ length: 7 }, (_, index) => {
      const date = new Date(lastDay);
      date.setDate(date.getDate() - (6 - index));
      const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
      const amount = sum(expenses.filter(item => item.date === key).map(item => item.amount));
      return { label: new Intl.DateTimeFormat(undefined, { weekday: "short" }).format(date).slice(0, 2), amount };
    });
    const max = Math.max(...daily.map(day => day.amount), 1);
    return `<div class="mini-chart"><div class="chart-summary"><strong>${money(data.bazaar + data.fixed)}</strong><span>in ${escapeHtml(monthLabel(state.month))}</span></div><div class="chart-bars">${daily.map((day, index) => `<div class="chart-day"><div class="chart-track"><div class="chart-bar ${index === 6 ? "current" : ""}" style="height:${day.amount ? Math.max(6, day.amount / max * 100) : 3}%"></div></div>${day.label}</div>`).join("")}</div><p class="month-note">Daily expenses · last 7 days</p></div>`;
  }
  function mealsPage() {
    const data = computed();
    const sorted = [...data.meals].sort((a, b) => b.date.localeCompare(a.date));
    return `${pageHeading("Meals", "Track daily meals for every member.", `<button class="primary-button" data-action="add-meal"><span class="button-icon">＋</span> Log meal</button>`)}
      <div class="card panel"><div class="panel-header"><div><h2>Meal log</h2><p>${data.mealCount.toLocaleString()} meals recorded · ${escapeHtml(monthLabel(state.month))}</p></div><span class="badge badge-active">${money(data.mealRate)} / meal</span></div>
      ${sorted.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>Date</th><th>Member</th><th>Meal</th><th>Quantity</th><th></th></tr></thead><tbody>${sorted.map(item => `<tr><td>${escapeHtml(dateLabel(item.date, { month: "short", day: "numeric", year: "numeric" }))}</td><td><strong>${escapeHtml(memberName(item.memberId))}</strong></td><td>${escapeHtml(item.type)}</td><td>${Number(item.quantity).toLocaleString()}</td><td><div class="row-actions"><button data-action="edit-meal" data-id="${item.id}">Edit</button><button class="remove" data-action="delete-meal" data-id="${item.id}">Delete</button></div></td></tr>`).join("")}</tbody></table></div>` : empty("No meals this month", "Log lunch or dinner to calculate your live meal rate and member shares.")}</div>
      <p class="month-note section-spacer">Meal rate = bazaar expenses ÷ total meal quantity. Fixed costs are shared evenly among current participants.</p>`;
  }
  function expensesPage() {
    const data = computed();
    const rows = [...data.expenses].sort((a, b) => b.date.localeCompare(a.date));
    return `${pageHeading("Expenses", "Keep bazaar and fixed costs in one clear ledger.", `<button class="primary-button" data-action="add-expense"><span class="button-icon">＋</span> Add expense</button>`)}
      <section class="grid stats-grid">${stat("Bazaar", money(data.bazaar), "Food and daily shopping", "receipt", "orange")}${stat("Fixed costs", money(data.fixed), "Shared household costs", "chart", "blue")}${stat("Total expenses", money(data.bazaar + data.fixed), "For this month", "coin")}${stat("Meal rate", money(data.mealRate), `${data.mealCount.toLocaleString()} total meals`, "fork")}</section>
      ${unallocatedBazaarNotice(data)}
      <div class="card panel"><div class="panel-header"><div><h2>Expense ledger</h2><p>${rows.length} expense${rows.length === 1 ? "" : "s"} · ${escapeHtml(monthLabel(state.month))}</p></div><div class="filters"><select class="filter-select" id="expense-filter" aria-label="Filter expense type"><option value="all">All expenses</option><option value="bazaar">Bazaar</option><option value="fixed">Fixed cost</option></select></div></div>
      ${rows.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>Date</th><th>Type</th><th>Description</th><th class="number">Amount</th><th></th></tr></thead><tbody>${rows.map(item => `<tr data-expense-kind="${item.kind}"><td>${escapeHtml(dateLabel(item.date, { month: "short", day: "numeric", year: "numeric" }))}</td><td><span class="badge ${item.kind === "fixed" ? "badge-archived" : "badge-active"}">${item.kind === "fixed" ? "Fixed cost" : "Bazaar"}</span></td><td>${escapeHtml(item.description || "—")}</td><td class="number"><strong>${money(item.amount)}</strong></td><td><div class="row-actions"><button data-action="edit-expense" data-id="${item.id}">Edit</button><button class="remove" data-action="delete-expense" data-id="${item.id}">Delete</button></div></td></tr>`).join("")}</tbody></table></div>` : empty("No expenses this month", "Add bazaar or fixed costs to see your live meal rate.")}</div>`;
  }
  function depositsPage() {
    const data = computed();
    const rows = [...data.deposits].sort((a, b) => b.date.localeCompare(a.date));
    return `${pageHeading("Deposits", "Record member contributions and see who's settled up.", `<button class="primary-button" data-action="add-deposit"><span class="button-icon">＋</span> Add deposit</button>`)}
      <section class="grid stats-grid">${stat("Deposits this month", money(data.depositTotal), `${rows.length} contribution${rows.length === 1 ? "" : "s"}`, "wallet")}${stat("Mess expenses", money(data.bazaar + data.fixed), "Bazaar and fixed costs", "receipt", "orange")}${stat("Cash balance", money(data.remaining), "Deposits less all expenses", "chart", data.remaining < 0 ? "red" : "")}${stat("Active members", activeMembers().length, "Currently in this mess", "users", "blue")}</section>
      <div class="card panel"><div class="panel-header"><div><h2>Deposit history</h2><p>${escapeHtml(monthLabel(state.month))}</p></div></div>
      ${rows.length ? `<div class="table-wrap"><table class="data-table"><thead><tr><th>Date</th><th>Member</th><th>Note</th><th>Direction</th><th class="number">Amount</th><th></th></tr></thead><tbody>${rows.map(item => `<tr><td>${escapeHtml(dateLabel(item.date, { month: "short", day: "numeric", year: "numeric" }))}</td><td><strong>${escapeHtml(memberName(item.memberId))}</strong></td><td>${escapeHtml(item.note || "—")}</td><td><span class="badge ${item.amount < 0 ? "badge-archived" : "badge-active"}">${item.amount < 0 ? "Debit" : "Credit"}</span></td><td class="number"><strong>${money(Math.abs(item.amount))}</strong></td><td><div class="row-actions"><button data-action="edit-deposit" data-id="${item.id}">Edit</button><button class="remove" data-action="delete-deposit" data-id="${item.id}">Delete</button></div></td></tr>`).join("")}</tbody></table></div>` : empty("No deposits this month", "Record a member contribution to keep your balance up to date.")}</div>`;
  }
  function membersPage() {
    const data = computed();
    const balances = new Map(data.balances.map(item => [item.member.id, item]));
    const members = [...state.mess.members].sort((a, b) => Number(b.active !== false) - Number(a.active !== false) || a.name.localeCompare(b.name));
    const cards = members.map(member => {
      const stats = balances.get(member.id);
      const balance = memberBalanceLabel(stats.net);
      return `<article class="card settlement-card"><div class="member-row"><span class="member-avatar">${escapeHtml(initials(member.name))}</span><span class="member-info"><strong>${escapeHtml(member.name)}</strong><small>${stats.meals.toLocaleString()} meals · ${money(stats.deposits)} deposited</small></span><span class="badge ${member.active === false ? "badge-archived" : "badge-active"}">${member.active === false ? "Removed" : "Active"}</span></div>
        <div class="settlement-meta"><span>Meal days</span><strong>${stats.mealDays}</strong></div>
        <div class="settlement-meta"><span>Meal cost (${stats.meals.toLocaleString()} × ${money(data.mealRate)})</span><strong>${money(stats.mealCost)}</strong></div>
        <div class="settlement-meta"><span>Fixed-cost share</span><strong>${money(stats.fixedShare)}</strong></div>
        <div class="settlement-meta settlement-total"><span>Total cost</span><strong>${money(stats.cost)}</strong></div>
        <div class="settlement-meta"><span>Deposits</span><strong>${money(stats.deposits)}</strong></div>
        <div class="settlement-meta settlement-balance"><span>Balance</span><strong class="${balance.className}">${money(Math.abs(stats.net))}</strong></div>
        <div class="settlement-meta"><span>Status</span><strong class="${balance.className}">${balance.text}</strong></div>
        <div class="row-actions" style="justify-content:flex-start;margin-top:12px"><button data-action="rename-member" data-id="${member.id}">Rename</button>${member.active === false ? `<button data-action="restore-member" data-id="${member.id}">Restore</button>` : `<button class="remove" data-action="remove-member" data-id="${member.id}">Remove</button>`}</div></article>`;
    });
    return `${pageHeading("Members", "Add, rename, or remove people from your mess.", `<button class="primary-button" data-action="add-member"><span class="button-icon">＋</span> Add member</button>`)}
      <section class="grid stats-grid">${stat("Members", activeMembers().length, "Active in this mess", "users")}${stat("Total meals", data.mealCount.toLocaleString(), `In ${monthLabel(state.month)}`, "fork", "orange")}${stat("Meal rate", money(data.mealRate), "Bazaar cost per meal", "chart", "blue")}${stat("Active participants", data.participants, "Share fixed costs", "coin")}</section>
      <div class="grid card-grid">${cards.length ? cards.join("") : `<div class="card panel">${empty("No members yet", "Add your first member to start tracking your mess.")}</div>`}</div>
      <p class="month-note section-spacer">Removing a member archives them so their past meals and deposits stay in your records. You can restore them anytime.</p>`;
  }
  function openReport() {
    openModal("Download monthly report", `Prepare a printable summary for ${monthLabel(state.month)}.`, `<label class="report-option"><input id="report-details" name="details" type="checkbox"><span><strong>Include daily meal details</strong><small>Show each member's breakfast, lunch, and dinner counts by date.</small></span></label>`, form => {
      const printWindow = window.open("", "_blank");
      if (!printWindow) throw new Error("The report window was blocked. Allow pop-ups for this page and try again.");
      const includeDetails = new FormData(form).has("details");
      printWindow.document.open();
      printWindow.document.write(reportDocument(includeDetails));
      printWindow.document.close();
      closeModal();
      showToast("Report opened. Use Print / Save PDF in the report window.");
    }, { submit: "Open report window" });
  }
  function reportDocument(includeDetails) {
    const data = computed();
    const memberRows = data.balances
      .filter(item => item.member.active !== false || item.meals || item.deposits)
      .map(item => {
        const status = item.net > .005 ? "Credit" : item.net < -.005 ? "Debit" : "Settled";
        return `<tr><td>${escapeHtml(item.member.name)}${item.member.active === false ? " (removed)" : ""}</td><td class="number">${item.mealDays}</td><td class="number">${item.meals.toLocaleString()}</td><td class="number">${money(item.mealCost)}</td><td class="number">${money(item.fixedShare)}</td><td class="number">${money(item.cost)}</td><td class="number">${money(item.deposits)}</td><td class="number">${money(Math.abs(item.net))}</td><td>${status}</td></tr>`;
      }).join("");
    const dailyRows = includeDetails ? dailyMealBreakdown(data.meals) : "";
    const summaryRows = [
      ["Total deposits", money(data.depositTotal)],
      ["Bazaar expenses", money(data.bazaar)],
      ["Fixed costs", money(data.fixed)],
      ["Total expenses", money(data.bazaar + data.fixed)],
      ["Total meals", data.mealCount.toLocaleString()],
      ["Meal rate", money(data.mealRate)],
      ["Remaining cash balance", money(data.remaining)]
    ].map(([label, value]) => `<tr><th>${label}</th><td>${value}</td></tr>`).join("");
    const reportTitle = `${state.mess.name} — ${monthLabel(state.month)} Mess Report`;
    const generatedAt = new Intl.DateTimeFormat(undefined, { dateStyle: "long", timeStyle: "short" }).format(new Date());
    return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(reportTitle)}</title><style>
      *{box-sizing:border-box}body{color:#202b25;font:13px/1.5 Arial,sans-serif;margin:0 auto;max-width:1000px;padding:36px 28px 48px}
      header{border-bottom:3px solid #277354;margin-bottom:24px;padding:0 0 16px;text-align:center}h1{font-size:25px;margin:0 0 4px;text-transform:none}header p{color:#66726a;margin:0}.generated{color:#748078;font-size:11px;margin-top:8px}
      h2{font-size:16px;margin:26px 0 10px;text-transform:capitalize}button,th{ text-transform:capitalize}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dfe5df;padding:8px 10px;text-align:left;vertical-align:top}thead th{background:#edf3ee;font-size:10px;letter-spacing:0}tbody th{font-weight:600}.number,.summary tbody td{text-align:right;white-space:nowrap}.note{color:#69766d;font-size:11px;margin-top:8px}
      .toolbar{display:flex;justify-content:flex-end;margin-bottom:16px}.toolbar button{background:#277354;border:0;border-radius:6px;color:#fff;cursor:pointer;font-size:12px;font-weight:700;padding:10px 14px}
      @page{margin:18mm 14mm 20mm}
      @media print{body{max-width:none;padding:0 0 10mm}.toolbar{display:none}h2{break-after:avoid;page-break-after:avoid}thead{display:table-header-group}tr{break-inside:avoid;page-break-inside:avoid}table{break-inside:auto;page-break-inside:auto}}
      </style></head><body><div class="toolbar"><button onclick="window.print()">Print / Save as PDF</button></div><header><h1>${escapeHtml(state.mess.name)}</h1><p>${escapeHtml(monthLabel(state.month))} Mess Report</p><p class="generated">Generated at ${escapeHtml(generatedAt)}</p></header>
      <section><h2>Mess summary</h2><table class="summary"><thead><tr><th>Summary</th><th class="number">Amount</th></tr></thead><tbody>${summaryRows}</tbody></table><p class="note">Meal rate = bazaar expenses ÷ total meal quantity. Fixed costs are split across ${data.participants} participant${data.participants === 1 ? "" : "s"}${data.bazaar > 0 && data.mealCount === 0 ? ". Bazaar expenses are included in the cash balance but cannot be assigned to member meal costs until meals are recorded" : ""}.</p></section>
      <section><h2>Member-wise summary</h2><table><thead><tr><th>Member</th><th class="number">Meal days</th><th class="number">Meals</th><th class="number">Meal cost</th><th class="number">Fixed share</th><th class="number">Total cost</th><th class="number">Deposits</th><th class="number">Balance</th><th>Status</th></tr></thead><tbody>${memberRows || `<tr><td colspan="9">No members recorded.</td></tr>`}</tbody></table><p class="note">Status is Credit, Debit, or Settled. Balance contains the amount only. Meal days count distinct dates with recorded meals.</p></section>
      ${includeDetails ? `<section><h2>Daily meal details</h2><table><thead><tr><th>Date</th><th>Member</th><th class="number">Breakfast</th><th class="number">Lunch</th><th class="number">Dinner</th><th class="number">Other</th><th class="number">Total</th></tr></thead><tbody>${dailyRows || `<tr><td colspan="7">No meals recorded this month.</td></tr>`}</tbody></table><p class="note">Meal quantities are counted from the saved meal log. Fractional meal quantities are included.</p></section>` : ""}
      </body></html>`;
  }
  function dailyMealBreakdown(meals) {
    const groups = new Map();
    meals.forEach(meal => {
      const key = `${meal.date}|${meal.memberId}`;
      if (!groups.has(key)) groups.set(key, { date: meal.date, memberId: meal.memberId, breakfast: 0, lunch: 0, dinner: 0, other: 0 });
      const group = groups.get(key);
      const type = String(meal.type || "").trim().toLowerCase();
      const category = ["breakfast", "lunch", "dinner"].includes(type) ? type : "other";
      group[category] += Number(meal.quantity || 0);
    });
    return [...groups.values()].sort((a, b) => a.date.localeCompare(b.date) || memberName(a.memberId).localeCompare(memberName(b.memberId)))
      .map(item => `<tr><td>${escapeHtml(dateLabel(item.date, { year: "numeric", month: "short", day: "numeric" }))}</td><td>${escapeHtml(memberName(item.memberId))}</td><td class="number">${item.breakfast.toLocaleString()}</td><td class="number">${item.lunch.toLocaleString()}</td><td class="number">${item.dinner.toLocaleString()}</td><td class="number">${item.other.toLocaleString()}</td><td class="number">${sum([item.breakfast, item.lunch, item.dinner, item.other]).toLocaleString()}</td></tr>`).join("");
  }
  function settingsPage() {
    return `${pageHeading("Mess settings", "Manage this mess and keep a safe copy of your records.")}
      <section class="grid settings-grid">
        <article class="card settings-card"><h2>App Appearance</h2><p>Change the app name, icon, favicon, and text size on this device.</p><button class="secondary-button" data-action="app-appearance">Customize Appearance</button></article>
        <article class="card settings-card"><h2>Google Sheets Sync</h2><p>Keep this device's encrypted mess data backed up to your Google Sheet. Normal sync uploads only; the app reads the sheet only when you explicitly choose Restore from Google Sheets.</p>
          <form id="sync-settings-form" class="sync-settings-form">
            <div class="field"><label for="google-client-id">Google OAuth Client ID</label><input id="google-client-id" name="google-client-id" type="text" value="${escapeHtml(appSettings.googleClientId)}" placeholder="Your Google OAuth client ID" autocomplete="off"></div>
            <div class="field"><label for="spreadsheet-id">Google Spreadsheet ID or URL</label><input id="spreadsheet-id" name="spreadsheet-id" type="text" value="${escapeHtml(appSettings.spreadsheetId)}" placeholder="Paste the Google Sheet URL or ID" autocomplete="off"></div>
            <div class="sync-form-error" id="sync-config-error" role="alert"></div>
            <div class="settings-actions"><button class="secondary-button" type="submit">Save Google Settings</button><button class="primary-button" type="button" data-action="connect-google-sheets">Connect Google Account</button><button class="secondary-button" type="button" data-action="sync-now">Sync Now</button></div>
          </form>
          <p class="sync-status-detail" id="sync-status-detail">${escapeHtml(syncStatus.message)}</p>
          <div class="notice">Setup: Create a Google Cloud project, enable the Google Sheets API, and create an OAuth Client ID for a Web application. Add this app's HTTPS website origin under Authorized JavaScript origins. Create a Google Sheet tab named <strong>MessMate</strong> and put these headers in cells A1:H1, one per column: <code>Mess ID, Chunk, Chunks, Mess Name, Updated At, Salt, IV, Encrypted Payload</code>. Paste the OAuth Client ID and spreadsheet URL above, save, then connect with a Google account that can edit the sheet. Records are encrypted and your mess password is not sent to Google; mess name and sync metadata remain visible. For local development, open through localhost rather than file://.</div>
          <div class="settings-actions section-spacer"><button class="secondary-button" type="button" data-action="restore-from-sheets">Restore from Google Sheets</button></div>
          <div class="notice">A green status means Google confirmed the latest upload. If offline or an upload fails, changes remain saved locally and the status is red until you reconnect and sync. Restoration is a deliberate copy into local storage, not a merge: existing local data is never overwritten. Another browser that has not synced a mess before may upload a separate copy.</div>
        </article>
        <article class="card settings-card"><h2>${escapeHtml(state.mess.name)}</h2><p>This mess is stored separately from every other mess in this browser. Your records are encrypted with the password you set up.</p><div class="settings-actions"><button class="secondary-button" data-action="switch-mess">Switch mess</button><button class="secondary-button" data-action="new-mess">Create another mess</button><button class="secondary-button" data-action="export-backup">${icon("download", 14)} Download data backup (.json)</button><button class="secondary-button" data-action="import-backup">Import data backup (.json)</button></div><div class="notice">To move your mess to another computer, download this encrypted JSON backup and copy it using a USB drive or other file transfer. On the other computer, open ${escapeHtml(appSettings.name)}, choose Import data backup, and then unlock the restored mess with the same password. The backup file is encrypted; keep your password safe.</div></article>
        <article class="card settings-card danger-zone"><h2>Clear this mess</h2><p>Remove all meal, expense, and deposit records from this mess. Your member list will be kept so you can start fresh.</p><button class="danger-button" data-action="clear-data">Clear all activity</button><div class="notice">This action cannot be undone. Export a backup first if you may need these records later.</div></article>
        <article class="card settings-card danger-zone"><h2>Delete this mess</h2><p>Permanently delete this mess and all of its encrypted records from this browser.</p><button class="danger-button" data-action="delete-mess">Delete mess permanently</button></article>
      </section>`;
  }
  function renderApp() {
    const content = state.page === "dashboard" ? dashboard() : state.page === "meals" ? mealsPage() :
      state.page === "bazaar" ? expensesPage() : state.page === "deposits" ? depositsPage() :
        state.page === "members" ? membersPage() : settingsPage();
    root.innerHTML = `<div class="app-shell">${sidebar()}<main class="main">${header()}<section class="content" id="page-content">${content}</section></main>${bottomNav()}<div id="modal-root"></div></div>`;
  }
  function authBrand() {
    return `<div class="brand"><div class="brand-mark">${brandMark()}</div><div class="brand-name">${escapeHtml(appSettings.name)}</div></div>`;
  }
  function renderAuth(mode = "choose", message = "") {
    state.mess = null; state.key = null;
    const index = getIndex();
    if (mode === "choose" && index.length === 0) mode = "create";
    if (mode === "unlock") {
      const chosen = index[0];
      root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>Unlock your mess</h1><form class="auth-form" id="unlock-form"><div class="field"><label for="unlock-name">Mess name</label><select id="unlock-name" required>${index.map(item => `<option value="${escapeHtml(item.id)}" ${chosen.id === item.id ? "selected" : ""}>${escapeHtml(item.name)}</option>`).join("")}</select></div><div class="field"><label for="unlock-password">Password</label><input id="unlock-password" type="password" autocomplete="current-password" required></div><div class="auth-error" id="auth-error">${escapeHtml(message)}</div><button class="primary-button" type="submit">Unlock mess</button></form><button class="auth-link" data-action="new-mess">Create a new mess</button></section></main>`;
    } else if (mode === "restore") {
      root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>Restore a mess</h1><p>Find an encrypted snapshot in your Google Sheet and copy it to this device. You will need the original mess password to unlock it.</p><form class="auth-form" id="restore-search-form"><div class="field"><label for="restore-google-client-id">Google OAuth Client ID</label><input id="restore-google-client-id" name="google-client-id" value="${escapeHtml(appSettings.googleClientId)}" autocomplete="off" required></div><div class="field"><label for="restore-spreadsheet-id">Google Spreadsheet ID or URL</label><input id="restore-spreadsheet-id" name="spreadsheet-id" value="${escapeHtml(appSettings.spreadsheetId)}" autocomplete="off" required></div><div class="auth-error" id="restore-error">${escapeHtml(message)}</div><button class="primary-button" type="submit">Find Google Sheet Backups</button></form><div id="restore-results"></div><button class="auth-link" data-action="choose-mess">← Back</button></section></main>`;
    } else if (mode === "create") {
      root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>Create your mess</h1><p>Each mess has its own password and separate records.</p><form class="auth-form" id="create-form"><div class="field"><label for="mess-name">Mess name</label><input id="mess-name" name="name" maxlength="60" placeholder="e.g. Green House" autocomplete="organization" required></div><div class="field"><label for="initial-members">Starter members <span class="field-hint">(optional, one per line or comma-separated)</span></label><textarea id="initial-members" name="members" placeholder="Fahim, Rayhan, Nishad"></textarea></div><div class="field"><label for="create-password">Create a password</label><input id="create-password" name="password" type="password" minlength="8" autocomplete="new-password" placeholder="At least 8 characters" required><span class="field-hint">Password must be at least 8 characters.</span></div><div class="field"><label for="confirm-password">Confirm password</label><input id="confirm-password" name="confirm" type="password" minlength="8" autocomplete="new-password" required></div><div class="auth-error" id="auth-error">${escapeHtml(message)}</div><button class="primary-button" type="submit">Create mess</button></form><button class="auth-link" data-action="restore-from-sheets">Restore from Google Sheets</button>${index.length ? `<button class="auth-link" data-action="choose-mess">← Back to existing messes</button>` : ""}</section></main>`;
    } else {
      root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>Your messes</h1><p>Select a mess or create another.</p><div class="auth-form">${index.map(item => `<button class="mess-switcher" data-action="pick-mess" data-id="${escapeHtml(item.id)}"><span class="mess-avatar">${escapeHtml(initials(item.name))}</span><span class="mess-switcher-copy"><small>Mess</small><strong>${escapeHtml(item.name)}</strong></span><span aria-hidden="true">→</span></button>`).join("")}<button class="primary-button" data-action="new-mess">＋ Create a new mess</button><button class="auth-link" data-action="restore-from-sheets">Restore from Google Sheets</button></div></section></main>`;
    }
  }
  function showError(idSelector, message) {
    const el = document.getElementById(idSelector);
    if (el) el.textContent = message;
  }
  async function createMess(form) {
    const fd = new FormData(form);
    const name = String(fd.get("name") || "").trim();
    const password = String(fd.get("password") || "");
    if (!name || name.length > 60) throw new Error("Enter a mess name between 1 and 60 characters.");
    if (password.length < 8) throw new Error("Choose a password with at least 8 characters.");
    if (password !== fd.get("confirm")) throw new Error("Your passwords do not match.");
    const index = getIndex();
    if (index.some(item => item.name.trim().toLocaleLowerCase() === name.toLocaleLowerCase())) {
      throw new Error("A mess with this name already exists. Choose a unique name.");
    }
    const membersInput = String(fd.get("members") || "");
    const names = membersInput.split(/[,\n;]/).map(value => value.trim()).filter(Boolean);
    if (names.some(memberName => memberName.length > 60)) throw new Error("Each starter member name must be 60 characters or fewer.");
    if (new Set(names.map(memberName => memberName.toLocaleLowerCase())).size !== names.length) throw new Error("Starter member names must be unique.");
    const mess = {
      id: id(), name, createdAt: new Date().toISOString(),
      members: names.map(memberName => ({ id: id(), name: memberName, active: true, joinedAt: currentDate() })),
      meals: [], expenses: [], deposits: []
    };
    const salt = crypto.getRandomValues(new Uint8Array(16));
    const key = await deriveKey(password, salt);
    const encrypted = await encryptData(mess, key);
    index.push({ id: mess.id, name, salt: toBase64(salt), iv: encrypted.iv, payload: encrypted.payload });
    saveIndex(index);
    state.mess = mess; state.key = key; state.page = "dashboard"; state.month = monthKey(new Date());
    renderApp();
    setSyncStatus("pending", "New mess saved locally. Syncing its encrypted data to Google Sheets…");
    void syncCurrentMess();
  }
  async function unlockMess(form) {
    const index = getIndex();
    const entry = index.find(item => item.id === form.elements.name.value);
    if (!entry) throw new Error("Choose a mess from the list.");
    const password = form.elements.password.value;
    let key;
    try { key = await deriveKey(password, fromBase64(entry.salt)); }
    catch (error) { throw new Error(`Could not read the saved encryption key. ${error.message}`); }
    try { state.mess = await decryptData(entry, key); }
    catch (error) {
      if (error.name === "OperationError") throw new Error("That password didn't unlock this mess. Please try again.");
      throw error;
    }
    state.key = key; state.page = "dashboard"; state.month = monthKey(new Date());
    renderApp();
    if (state.skipNextSyncForMess === entry.id) {
      state.skipNextSyncForMess = "";
      setSyncStatus("online", "Restored snapshot unlocked. It is saved locally and was not uploaded again.");
    } else if (appSettings.googleClientId && appSettings.spreadsheetId) {
      setSyncStatus("pending", "Checking Google Sheets sync for this mess…");
      void syncCurrentMess();
    }
  }
  function openModal(title, description, body, onSubmit, options = {}) {
    const target = document.getElementById("modal-root");
    target.innerHTML = `<div class="modal-backdrop" data-action="backdrop-close"><section class="modal ${options.wide ? "wide" : ""}" role="dialog" aria-modal="true" aria-labelledby="modal-title"><div class="modal-header"><div><h2 id="modal-title">${escapeHtml(title)}</h2><p>${escapeHtml(description)}</p></div><button class="modal-close" type="button" data-action="close-modal" aria-label="Close">×</button></div><form id="modal-form">${body}<div class="modal-error" id="modal-error"></div><div class="modal-actions"><button type="button" class="secondary-button" data-action="close-modal">Cancel</button><button type="submit" class="${options.danger ? "danger-button" : "primary-button"}">${escapeHtml(options.submit || "Save")}</button></div></form></section></div>`;
    state.modal = { onSubmit };
    const first = target.querySelector("input,select,textarea");
    if (first) first.focus();
  }
  function closeModal() {
    if (state.modal?.fontScalePreview && !state.modal.keepFontScalePreview) applyAppSettings();
    state.modal = null;
    const target = document.getElementById("modal-root");
    if (target) target.innerHTML = "";
  }
  function memberOptions(selected = "") {
    const members = state.mess.members.filter(member => member.active !== false || member.id === selected);
    return members.map(member => `<option value="${escapeHtml(member.id)}" ${selected === member.id ? "selected" : ""}>${escapeHtml(member.name)}${member.active === false ? " (removed)" : ""}</option>`).join("");
  }
  function findRecord(collection, recordId) {
    return state.mess[collection].find(record => record.id === recordId);
  }
  function openMeal(recordId = "") {
    const record = recordId ? findRecord("meals", recordId) : null;
    if (!activeMembers().length && !record) {
      showToast("Add a member before logging meals.");
      state.page = "members"; renderApp(); return;
    }
    openModal(record ? "Edit meal" : "Log a meal", "Record a member's meal for this date.", `<div class="form-grid">
      <div class="field"><label for="meal-member">Member</label><select id="meal-member" name="memberId" required>${memberOptions(record?.memberId)}</select></div>
      <div class="field"><label for="meal-type">Meal</label><select id="meal-type" name="type" required><option ${record?.type === "Lunch" ? "selected" : ""}>Lunch</option><option ${record?.type === "Dinner" ? "selected" : ""}>Dinner</option><option ${record?.type === "Breakfast" ? "selected" : ""}>Breakfast</option></select></div>
      <div class="field"><label for="meal-date">Date</label><input id="meal-date" name="date" type="date" value="${escapeHtml(record?.date || entryDate())}" required></div>
      <div class="field"><label for="meal-quantity">Meal quantity</label><input id="meal-quantity" name="quantity" type="number" min="0.01" max="10000" step="0.01" value="${record ? escapeHtml(record.quantity) : "1"}" required></div></div>`, async form => {
        const fd = new FormData(form);
        const data = { memberId: fd.get("memberId"), date: fd.get("date"), type: fd.get("type"), quantity: Number(fd.get("quantity")) };
        if (!data.memberId || !state.mess.members.some(member => member.id === data.memberId && (member.active !== false || member.id === record?.memberId)) || !validDate(data.date) || !["Breakfast", "Lunch", "Dinner"].includes(data.type) || !Number.isFinite(data.quantity) || data.quantity <= 0 || data.quantity > 10000) throw new Error("Enter a valid active member, date, meal type, and quantity between 0.01 and 10,000.");
        if (record) Object.assign(record, data);
        else state.mess.meals.push({ id: id(), ...data, createdAt: new Date().toISOString() });
        await persistMess(); closeModal(); renderApp(); showToast(record ? "Meal updated." : "Meal added.");
      });
  }
  function openExpense(recordId = "") {
    const record = recordId ? findRecord("expenses", recordId) : null;
    openModal(record ? "Edit expense" : "Add an expense", "Bazaar expenses set the meal rate. Fixed costs are shared among members.", `<div class="form-grid">
      <div class="field"><label for="expense-kind">Expense type</label><select id="expense-kind" name="kind" required><option value="bazaar" ${record?.kind !== "fixed" ? "selected" : ""}>Bazaar</option><option value="fixed" ${record?.kind === "fixed" ? "selected" : ""}>Fixed cost</option></select></div>
      <div class="field"><label for="expense-date">Date</label><input id="expense-date" name="date" type="date" value="${escapeHtml(record?.date || entryDate())}" required></div>
      <div class="field"><label for="expense-amount">Amount</label><div class="amount-input"><input id="expense-amount" name="amount" type="number" min="0.01" max="1000000000" step="0.01" value="${record ? escapeHtml(record.amount) : ""}" placeholder="0.00" required><span>${CURRENCY_SYMBOL}</span></div></div>
      <div class="field"><label for="expense-description">Description</label><input id="expense-description" name="description" maxlength="100" value="${escapeHtml(record?.description || "")}" placeholder="e.g. Vegetables, electricity"></div></div>`, async form => {
        const fd = new FormData(form);
        const data = { kind: fd.get("kind"), date: fd.get("date"), amount: Number(fd.get("amount")), description: String(fd.get("description") || "").trim() };
        if (!validDate(data.date) || !["bazaar", "fixed"].includes(data.kind) || !Number.isFinite(data.amount) || data.amount <= 0 || data.amount > 1000000000) throw new Error("Enter a valid date, expense type, and amount up to 1,000,000,000.");
        if (record) Object.assign(record, data);
        else state.mess.expenses.push({ id: id(), ...data, createdAt: new Date().toISOString() });
        await persistMess(); closeModal(); renderApp(); showToast(record ? "Expense updated." : "Expense added.");
      });
  }
  function openDeposit(recordId = "") {
    const record = recordId ? findRecord("deposits", recordId) : null;
    if (!activeMembers().length && !record) {
      showToast("Add a member before recording a deposit.");
      state.page = "members"; renderApp(); return;
    }
    openModal(record ? "Edit deposit" : "Record a deposit", "Add a contribution to a member's account.", `<div class="form-grid">
      <div class="field"><label for="deposit-member">Member</label><select id="deposit-member" name="memberId" required>${memberOptions(record?.memberId)}</select></div>
      <div class="field"><label for="deposit-date">Date</label><input id="deposit-date" name="date" type="date" value="${escapeHtml(record?.date || entryDate())}" required></div>
      <div class="field"><label for="deposit-amount">Amount</label><div class="amount-input"><input id="deposit-amount" name="amount" type="number" min="-1000000000" max="1000000000" step="0.01" value="${record ? escapeHtml(record.amount) : ""}" placeholder="0.00" required><span>${CURRENCY_SYMBOL}</span></div></div>
      <div class="field"><label for="deposit-note">Note (optional)</label><input id="deposit-note" name="note" maxlength="100" value="${escapeHtml(record?.note || "")}" placeholder="e.g. Monthly contribution"></div>
      <div class="field full"><span class="field-hint">Use a negative amount to correct or reverse a previous deposit.</span></div></div>`, async form => {
        const fd = new FormData(form);
        const data = { memberId: fd.get("memberId"), date: fd.get("date"), amount: Number(fd.get("amount")), note: String(fd.get("note") || "").trim() };
        if (!data.memberId || !state.mess.members.some(member => member.id === data.memberId) || !validDate(data.date) || !Number.isFinite(data.amount) || data.amount === 0 || Math.abs(data.amount) > 1000000000) throw new Error("Enter a valid member, date, and non-zero amount up to 1,000,000,000.");
        if (record) Object.assign(record, data);
        else state.mess.deposits.push({ id: id(), ...data, createdAt: new Date().toISOString() });
        await persistMess(); closeModal(); renderApp(); showToast(record ? "Deposit updated." : "Deposit recorded.");
      });
  }
  function openMember(memberId = "") {
    const member = memberId ? state.mess.members.find(item => item.id === memberId) : null;
    openModal(member ? "Rename member" : "Add a member", member ? "Update the name shown in your records." : "New members are available in meal and deposit forms right away.", `<div class="field"><label for="member-name">Member name</label><input id="member-name" name="name" maxlength="60" value="${escapeHtml(member?.name || "")}" placeholder="e.g. Alex" required></div>`, async form => {
      const name = String(new FormData(form).get("name") || "").trim();
      if (!name || name.length > 60) throw new Error("Enter a member name between 1 and 60 characters.");
      if (state.mess.members.some(item => item.id !== member?.id && item.name.toLocaleLowerCase() === name.toLocaleLowerCase())) throw new Error("A member with this name already exists.");
      if (member) member.name = name;
      else state.mess.members.push({ id: id(), name, active: true, joinedAt: currentDate() });
      await persistMess(); closeModal(); renderApp(); showToast(member ? "Member renamed." : "Member added.");
    });
  }
  function readImageFile(file, label, allowedTypes) {
    if (!file) return Promise.resolve("");
    if (!allowedTypes.includes(file.type)) throw new Error(`${label} must be a PNG, JPEG, WebP${allowedTypes.includes("image/x-icon") ? ", or ICO" : ""} image.`);
    if (file.size > 256 * 1024) throw new Error(`${label} must be 256 KB or smaller.`);
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.addEventListener("load", () => resolve(String(reader.result)));
      reader.addEventListener("error", () => reject(new Error(`${label} could not be read.`)));
      reader.readAsDataURL(file);
    });
  }
  function openAppAppearance() {
    const body = `<div class="field"><label for="app-name">App name</label><input id="app-name" name="name" maxlength="40" value="${escapeHtml(appSettings.name)}" required></div>
      <div class="field section-spacer"><label for="app-icon-file">App icon</label><input id="app-icon-file" name="icon" type="file" accept="image/png,image/jpeg,image/webp"><span class="field-hint">PNG, JPEG, or WebP · up to 256 KB</span><label class="check-option"><input name="resetIcon" type="checkbox"> Use the default icon</label></div>
      <div class="field section-spacer"><label for="app-favicon-file">Favicon</label><input id="app-favicon-file" name="favicon" type="file" accept="image/png,image/jpeg,image/webp,image/x-icon,image/vnd.microsoft.icon"><span class="field-hint">PNG, JPEG, WebP, or ICO · up to 256 KB</span><label class="check-option"><input name="resetFavicon" type="checkbox"> Use the default favicon</label></div>
      <div class="field section-spacer"><label for="app-font-scale">Text size</label><div class="font-scale-control"><input id="app-font-scale" name="fontScale" type="range" min="80" max="130" step="5" value="${Math.round(Number(appSettings.fontScale) * 100)}" aria-describedby="font-scale-value"><output id="font-scale-value">${Math.round(Number(appSettings.fontScale) * 100)}%</output></div><span class="field-hint">Drag to preview from 80% to 130%. Changes apply as you move the slider.</span></div>`;
    openModal("App Appearance", "Personalize the app on this device.", body, async form => {
      const formData = new FormData(form);
      const name = String(formData.get("name") || "").trim();
      const fontScale = (Number(formData.get("fontScale")) / 100).toFixed(2);
      if (!name || name.length > 40) throw new Error("Enter an app name between 1 and 40 characters.");
      if (!validFontScale(fontScale)) throw new Error("Choose a text size between 80% and 130%.");
      const iconFile = form.querySelector("#app-icon-file").files[0];
      const faviconFile = form.querySelector("#app-favicon-file").files[0];
      if (formData.has("resetIcon") && iconFile) throw new Error("Choose either a new app icon or the default icon.");
      if (formData.has("resetFavicon") && faviconFile) throw new Error("Choose either a new favicon or the default favicon.");
      const next = {
        name,
        icon: formData.has("resetIcon") ? "" : await readImageFile(iconFile, "App icon", ["image/png", "image/jpeg", "image/webp"]) || appSettings.icon,
        favicon: formData.has("resetFavicon") ? "" : await readImageFile(faviconFile, "Favicon", APP_IMAGE_TYPES) || appSettings.favicon,
        fontScale
      };
      saveAppSettings(next);
      state.modal.keepFontScalePreview = true;
      closeModal();
      renderApp();
      showToast("App appearance updated.");
    }, { wide: true, submit: "Save Appearance" });
    state.modal.fontScalePreview = true;
  }
  function confirmAction(title, description, label, action, options = {}) {
    const hint = options.hint ? `<div class="field section-spacer"><label for="confirm-text">${escapeHtml(options.hint)}</label><input id="confirm-text" name="confirmText" autocomplete="off" required></div>` : "";
    openModal(title, description, hint, async form => {
      if (options.hint && new FormData(form).get("confirmText") !== options.hint) throw new Error(`Type ${options.hint} exactly to confirm.`);
      await action();
    }, { danger: true, submit: label });
  }
  async function deleteRecord(collection, recordId, name) {
    await confirmAction(`Delete ${name}?`, "This record will be permanently removed from the current mess.", "Delete entry", async () => {
      state.mess[collection] = state.mess[collection].filter(item => item.id !== recordId);
      await persistMess(); closeModal(); renderApp(); showToast(`${name} deleted.`);
    });
  }
  function switchToMess(idValue) {
    const index = getIndex();
    const entry = index.find(item => item.id === idValue);
    if (!entry) return renderAuth("choose");
    root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>Unlock ${escapeHtml(entry.name)}</h1><form class="auth-form" id="unlock-form"><input type="hidden" name="name" value="${escapeHtml(entry.id)}"><div class="field"><label for="switch-password">Password</label><input id="switch-password" name="password" type="password" autocomplete="current-password" required></div><div class="auth-error" id="auth-error"></div><button class="primary-button" type="submit">Unlock mess</button></form><button class="auth-link" data-action="choose-mess">← Back to all messes</button></section></main>`;
  }
  async function exportBackup() {
    const file = new Blob([JSON.stringify({ format: "messmate-encrypted-backup-v1", mess: getIndex().find(item => item.id === state.mess.id) }, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(file);
    const link = document.createElement("a");
    link.href = url; link.download = `${state.mess.name.replace(/[^a-z0-9_-]+/gi, "-").toLowerCase()}-messmate-backup.json`;
    link.click(); URL.revokeObjectURL(url);
    showToast("Encrypted backup downloaded.");
  }
  function importBackup() {
    openModal("Import data backup", "Select a Messmate JSON backup to transfer a mess to this device.", `<div class="field"><label for="backup-file">JSON backup file</label><input id="backup-file" type="file" accept=".json,application/json" required></div><p class="field-hint">The imported mess stays encrypted and uses the same password as on the original device. A mess with the same name cannot be overwritten.</p>`, async form => {
      const file = form.querySelector("#backup-file").files[0];
      if (!file) throw new Error("Select a backup file.");
      if (file.size > 10 * 1024 * 1024) throw new Error("Backup files must be 10 MB or smaller.");
      let parsed;
      try { parsed = JSON.parse(await file.text()); } catch { throw new Error("This file is not valid JSON."); }
      if (parsed?.format !== "messmate-encrypted-backup-v1" || !parsed.mess || typeof parsed.mess.id !== "string" || !parsed.mess.id || typeof parsed.mess.name !== "string" || !parsed.mess.name.trim() || parsed.mess.name.length > 60 || typeof parsed.mess.payload !== "string" || typeof parsed.mess.iv !== "string" || typeof parsed.mess.salt !== "string") throw new Error("This is not a valid MessMate encrypted backup.");
      const index = getIndex();
      if (index.some(item => item.name.toLocaleLowerCase() === parsed.mess.name.trim().toLocaleLowerCase())) throw new Error("A mess with this name already exists. Switch to it, or remove it first.");
      index.push(parsed.mess); saveIndex(index); closeModal(); renderAuth("choose"); showToast("Backup restored. Unlock the mess to continue.");
    }, { submit: "Import backup" });
  }
  async function clearAllActivity() {
    await confirmAction("Clear all activity?", "This permanently removes all meals, expenses and deposits for this mess, while keeping your member list.", "Clear activity", async () => {
      state.mess.meals = []; state.mess.expenses = []; state.mess.deposits = [];
      await persistMess(); closeModal(); state.page = "dashboard"; renderApp(); showToast("Activity cleared. Your members are still here.");
    }, { hint: "CLEAR" });
  }
  async function deleteMess() {
    await confirmAction("Delete this mess permanently?", `This removes ${state.mess.name} and all its data from this browser. This cannot be undone.`, "Delete mess", async () => {
      const index = getIndex().filter(item => item.id !== state.mess.id);
      saveIndex(index); closeModal(); renderAuth(index.length ? "choose" : "create");
    }, { hint: "DELETE" });
  }
  async function submitModal(form) {
    if (!state.modal) return;
    const errorEl = document.getElementById("modal-error");
    errorEl.classList.remove("visible");
    const submit = form.querySelector('[type="submit"]');
    submit.disabled = true;
    try { await state.modal.onSubmit(form); }
    catch (error) { errorEl.textContent = error.message || "Could not save this change."; errorEl.classList.add("visible"); }
    finally { if (submit.isConnected) submit.disabled = false; }
  }
  async function handleAction(action, target) {
    const recordId = target.dataset.id;
    switch (action) {
      case "new-mess": renderAuth("create"); break;
      case "choose-mess": renderAuth("choose"); break;
      case "pick-mess": switchToMess(recordId); break;
      case "switch-mess": renderAuth("choose"); break;
      case "settings": state.page = "settings"; renderApp(); break;
      case "connect-google-sheets": await connectGoogleSheets(); break;
      case "sync-now": await syncNow(); break;
      case "restore-from-sheets": availableBackups = []; renderAuth("restore"); break;
      case "restore-sheet-backup": restoreSheetBackup(target.dataset.index); break;
      case "app-appearance": openAppAppearance(); break;
      case "print-report": openReport(); break;
      case "quick-add": openQuickAdd(); break;
      case "add-meal": openMeal(); break;
      case "edit-meal": openMeal(recordId); break;
      case "delete-meal": deleteRecord("meals", recordId, "meal"); break;
      case "add-expense": openExpense(); break;
      case "edit-expense": openExpense(recordId); break;
      case "delete-expense": deleteRecord("expenses", recordId, "expense"); break;
      case "add-deposit": openDeposit(); break;
      case "edit-deposit": openDeposit(recordId); break;
      case "delete-deposit": deleteRecord("deposits", recordId, "deposit"); break;
      case "add-member": openMember(); break;
      case "rename-member": openMember(recordId); break;
      case "remove-member": await setMemberActive(recordId, false); break;
      case "restore-member": await setMemberActive(recordId, true); break;
      case "export-backup": await exportBackup(); break;
      case "import-backup": importBackup(); break;
      case "clear-data": clearAllActivity(); break;
      case "delete-mess": deleteMess(); break;
      case "close-modal": closeModal(); break;
      case "backdrop-close": if (target === document.querySelector(".modal-backdrop")) closeModal(); break;
      default: break;
    }
  }
  function openQuickAdd() {
    openModal("Add a new entry", "Choose what you'd like to add to your mess.", `<div class="grid" style="grid-template-columns:1fr 1fr;gap:8px"><button type="button" class="secondary-button" data-action="quick-meal">${icon("fork", 16)} Log a meal</button><button type="button" class="secondary-button" data-action="quick-expense">${icon("receipt", 16)} Add expense</button><button type="button" class="secondary-button" data-action="quick-deposit">${icon("wallet", 16)} Record deposit</button></div>`, async () => {});
    const form = document.getElementById("modal-form");
    form.querySelector(".modal-actions").remove();
    form.addEventListener("click", event => {
      const button = event.target.closest("[data-action]");
      if (!button) return;
      closeModal();
      if (button.dataset.action === "quick-meal") openMeal();
      if (button.dataset.action === "quick-expense") openExpense();
      if (button.dataset.action === "quick-deposit") openDeposit();
    });
  }
  async function setMemberActive(memberId, active) {
    const member = state.mess.members.find(item => item.id === memberId);
    if (!member) return;
    if (!active) {
      await confirmAction(`Remove ${member.name}?`, "They will no longer appear in new meal or deposit forms. Existing history and balances are preserved.", "Remove member", async () => {
        member.active = false; member.removedAt = currentDate();
        await persistMess(); closeModal(); renderApp(); showToast(`${member.name} removed. Their history is preserved.`);
      });
      return;
    }
    member.active = true; delete member.removedAt;
    await persistMess(); renderApp(); showToast(`${member.name} restored.`);
  }
  root.addEventListener("click", async event => {
    const pageButton = event.target.closest("[data-page]");
    if (pageButton && state.mess) {
      state.page = pageButton.dataset.page; renderApp(); return;
    }
    const actionButton = event.target.closest("[data-action]");
    if (actionButton) {
      if (actionButton.dataset.action === "backdrop-close" && event.target !== actionButton) return;
      if (actionButton.dataset.action === "quick-meal" || actionButton.dataset.action === "quick-expense" || actionButton.dataset.action === "quick-deposit") return;
      try { await handleAction(actionButton.dataset.action, actionButton); }
      catch (error) { showToast(error.message || "Something went wrong."); }
    }
  });
  root.addEventListener("change", event => {
    if (event.target.id === "month-select") {
      state.month = event.target.value; renderApp();
    }
    if (event.target.id === "expense-filter") {
      const selected = event.target.value;
      document.querySelectorAll("[data-expense-kind]").forEach(row => row.classList.toggle("hidden", selected !== "all" && row.dataset.expenseKind !== selected));
    }
  });
  root.addEventListener("input", event => {
    if (event.target.id !== "app-font-scale") return;
    const scale = Number(event.target.value) / 100;
    document.documentElement.style.fontSize = `${16 * scale}px`;
    const output = document.getElementById("font-scale-value");
    if (output) output.textContent = `${event.target.value}%`;
  });
  root.addEventListener("submit", async event => {
    const form = event.target;
    if (form.id === "create-form") {
      event.preventDefault();
      const button = form.querySelector('[type="submit"]');
      button.disabled = true;
      try { await createMess(form); }
      catch (error) { showError("auth-error", error.message || "Could not create your mess."); }
      finally { if (button.isConnected) button.disabled = false; }
    } else if (form.id === "unlock-form") {
      event.preventDefault();
      const button = form.querySelector('[type="submit"]');
      button.disabled = true;
      try { await unlockMess(form); }
      catch (error) { showError("auth-error", error.message || "Could not unlock this mess."); }
      finally { if (button.isConnected) button.disabled = false; }
    } else if (form.id === "modal-form") {
      event.preventDefault(); await submitModal(form);
    } else if (form.id === "sync-settings-form") {
      event.preventDefault();
      const error = document.getElementById("sync-config-error");
      error.textContent = "";
      try { saveSyncSettings(form); }
      catch (saveError) {
        error.textContent = saveError.message || "Could not save Google Sheets settings.";
      }
    } else if (form.id === "restore-search-form") {
      event.preventDefault();
      const error = document.getElementById("restore-error");
      const button = form.querySelector('[type="submit"]');
      error.textContent = "";
      button.disabled = true;
      try { await findSheetBackups(form); }
      catch (searchError) { error.textContent = searchError.message || "Could not read backups from Google Sheets."; }
      finally { if (button.isConnected) button.disabled = false; }
    }
  });
  window.addEventListener("online", () => {
    if (!state.mess) return;
    setSyncStatus("pending", "Internet connection restored. Retrying Google Sheets sync…");
    void syncCurrentMess();
  });
  window.addEventListener("offline", () => {
    if (state.mess) setSyncStatus("offline", "You are offline. Changes are saved locally and will sync when connected.");
  });
  function init() {
    try {
      if (!window.crypto?.subtle) throw new Error("This browser does not support secure local encryption. Please use a current version of Chrome, Edge, Firefox or Safari.");
      Object.assign(appSettings, readAppSettings());
      applyAppSettings();
      if (appSettings.googleClientId && appSettings.spreadsheetId) {
        setSyncStatus("offline", navigator.onLine ? "Sign in to Google to verify online sync." : "You are offline. Changes are saved locally.");
      } else {
        setSyncStatus("local", "Changes are saved on this device only. Connect Google Sheets to enable online sync.");
      }
      const index = getIndex();
      renderAuth(index.length ? "choose" : "create");
    } catch (error) {
      root.innerHTML = `<main class="auth-wrap"><section class="auth-card">${authBrand()}<h1>${escapeHtml(appSettings.name)} needs storage access</h1><p>${escapeHtml(error.message)} Open this app in a modern browser with local storage enabled and try again.</p></section></main>`;
    }
  }
  init();
})();
