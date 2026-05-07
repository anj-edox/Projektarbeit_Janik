/**
 * OrderCheck App - Hauptmodul
 *
 * Architektur:
 * - Frontend: Vanilla JS + ES Modules
 * - Auth: Microsoft Teams SSO via MSAL Browser
 * - Backend: SharePoint Online + Microsoft Graph API
 *
 * Ablauf:
 * 1. User authentifiziert sich via Teams SSO (auth-start.html → auth-end.html)
 * 2. PDFs werden nach SharePoint hochgeladen und mit einer Prüf-ID versehen
 * 3. Ein Power Automate Flow verarbeitet die PDFs und schreibt Ergebnisse in eine SharePoint-Liste
 * 4. Ergebnisse werden per Graph API abgerufen und tabellarisch dargestellt
 */

import { PublicClientApplication } from 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.0/+esm';

const CONFIG = {
    msal: {
        clientId: "9603fe2f-fad8-4a18-9950-14f1eb195745",
        tenantId: "2b71e8a3-7c31-40db-84b0-9c7e50b9ea71",
        scopes: ['User.Read', 'Files.ReadWrite', 'Sites.ReadWrite.All']
    },
    sharePoint: {
        siteId: "edox.sharepoint.com,e68093ae-074a-4851-88dd-31284b07b284,e94104cd-b270-4e09-b29a-1932cc934dbf",
        listId: "ecac4b28-01fb-489a-ad32-f5ee4fd4bd83"
    }
};

const loginBtn = document.getElementById('loginBtn');
const logoutBtn = document.getElementById('logoutBtn');
const authInfo = document.getElementById('authInfo');

const fileInput = document.getElementById('fileInput');
const selectBtn = document.getElementById('selectBtn');
const uploadBtn = document.getElementById('uploadBtn');
const fileList = document.getElementById('fileList');
const statusBox = document.getElementById('status');

const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');
const refreshOrdersBtn = document.getElementById('refreshOrdersBtn');
const resultTable = document.getElementById('resultTable');

let msalInstance = null;
let msalInitPromise = null;
let activeAccount = null;
let selectedFiles = [];
let pruefvorgaenge = [];

await loadSavedConfig();
renderFileList();
updateAuthDisplay();

/**
 * Pollt SharePoint nach Ergebnissen für einen Prüfvorgang.
 * Wiederholt den Abruf alle 5 Sekunden bis max. 10 Mal.
 *
 * @param {string} token - Microsoft Graph Access Token
 * @param {string} pruefId - Eindeutige ID des Prüfvorgangs
 * @returns {Promise<Array>} Array der Ergebnis-Items
 * @throws {Error} Wenn nach max. Versuchen kein Ergebnis vorhanden ist
 */
async function waitForResult(token, pruefId) {
    let tries = 0;
    const maxTries = 10;

    while (tries < maxTries) {
        const result = await getPruefvorgangByPruefId(token, pruefId);
        const items = Array.isArray(result) ? result : [];

        if (items.length > 0) {
            return items;
        }

        await sleep(5000);
        tries++;
    }

    throw new Error("Timeout: Kein Ergebnis vom Flow erhalten");
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}


loginBtn.addEventListener('click', async () => {
    try {
        await window.microsoftTeams.app.initialize();

        await window.microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/auth-start.html",
            width: 600,
            height: 535
        });

        await ensureMsal();

        const accounts = msalInstance.getAllAccounts();

        if (accounts.length > 0) {
            activeAccount = accounts[0];
            msalInstance.setActiveAccount(activeAccount);
            updateAuthDisplay();
            setStatus('Anmeldung erfolgreich.');
        } else {
            setStatus('Anmeldung abgeschlossen. Bitte Seite einmal neu laden.');
        }

    } catch (error) {
        console.error(error);
        setStatus('Fehler bei der Anmeldung:\n' + formatError(error));
    }
});

searchInput.addEventListener("input", () => {
    const search = searchInput.value.toLowerCase().trim();

    if (!search) {
        searchResults.innerHTML = "";
        return;
    }

    const filtered = pruefvorgaenge.filter(item =>
        item.auftragsnummer.toLowerCase().includes(search)
    );

    renderSearchResults(filtered);
});

refreshOrdersBtn.addEventListener("click", async () => {
    try {
        await ensureMsal();
        const token = await getAccessToken();

        searchResults.textContent = "Lade Prüfvorgänge...";
        await loadPruefvorgaenge(token);

    } catch (error) {
        console.error(error);
        searchResults.textContent = "Fehler beim Laden:\n" + formatError(error);
    }
});

window.showScreen = function (name) {
    document.querySelectorAll(".screen").forEach(screen => {
        screen.classList.remove("active");
    });

    const activeScreen = document.getElementById("screen-" + name);
    if (activeScreen) {
        activeScreen.classList.add("active");
    }

    document.querySelectorAll(".nav button").forEach(button => {
        button.classList.remove("active");
    });

    const activeButton = document.querySelector(`.nav button[data-screen="${name}"]`);
    if (activeButton) {
        activeButton.classList.add("active");
    }
};

window.selectPruefvorgang = async function (pruefId) {
    let item = null;

    try {
        await ensureMsal();
        const token = await getAccessToken();

        item = pruefvorgaenge.find(x => x.pruefId === pruefId);

        if (!item) {
            setStatus("Prüfvorgang nicht gefunden.");
            return;
        }

        setStatus(`Lade Analyse für ${item.auftragsnummer}...`);

        const fehlerpositionen = await getFehlerpositionenByPruefId(token, pruefId);
        renderResultTable(fehlerpositionen);
        showScreen("result");

        setStatus(`Analyse geladen: ${item.auftragsnummer}`);

        searchInput.value = item.auftragsnummer;
        searchResults.innerHTML = "";

    } catch (error) {
        console.error(error);
        setStatus("Fehler beim Laden der Analyse:\n" + formatError(error));
    }
};

logoutBtn.addEventListener('click', async () => {
    try {
        await ensureMsal();

        activeAccount = null;
        msalInstance.setActiveAccount(null);

        updateAuthDisplay();
        setStatus('Aus App-Ansicht abgemeldet. Login kann erneut gestartet werden.');

    } catch (error) {
        console.error(error);
        setStatus('Fehler bei der Abmeldung:\n' + formatError(error));
    }
});

selectBtn.addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', (event) => {
    const newFiles = Array.from(event.target.files || []);

    for (const file of newFiles) {
        // doppelte vermeiden
        const exists = selectedFiles.some(item => item.file.name === file.name && item.file.size === file.size);

        if (!exists) {
            selectedFiles.push({ file: file, status: 'bereit' });
        }
    }
    renderFileList();
    setStatus(`${selectedFiles.length} PDF-Datei(en) ausgewählt.`);
});

uploadBtn.addEventListener('click', async () => {
    try {
        await ensureMsal();

        if (selectedFiles.length === 0) {
            setStatus('Bitte zuerst mindestens eine PDF auswählen.');
            return;
        }

        const token = await getAccessToken();
        statusBox.textContent = 'Upload gestartet...\n';

        let hasError = false;
        const pruefId = crypto.randomUUID();
        for (const item of selectedFiles) {
            try {
                item.status = 'uploading';
                renderFileList();

                await uploadSmallFileToSharePoint(item.file, token, pruefId);

                item.status = 'success';
                renderFileList();
            } catch (err) {
                item.status = 'error';
                hasError = true;
                renderFileList();
            }
        }

        if (!hasError) {
            statusBox.textContent += '\nAlle Dateien erfolgreich hochgeladen.';

            const eintraege = await getPruefvorgangByPruefId(token, pruefId);
            console.log("Gefundene Prüfvorgänge:", eintraege);
            const result = await waitForResult(token, pruefId);
            console.log("Finales Ergebnis:", result);

            selectedFiles = [];
            renderFileList();
            fileInput.value = '';
            showScreen("search");
        } else {
            statusBox.textContent += '\nEinige Dateien konnten nicht hochgeladen werden.';
        }
    } catch (error) {
        console.error(error);
        setStatus('Fehler beim Upload:\n' + formatError(error));
    }
});

async function loadSavedConfig() {
    await initializeMsal();
}

/**
 * Aktualisiert die Auth-UI basierend auf dem aktuellen Login-Status.
 * Setzt den authInfo-Farbpunkt und zeigt/versteckt Login/Logout-Buttons.
 */
function updateAuthDisplay() {
    const account = msalInstance?.getActiveAccount() || activeAccount;

    if (account) {
        authInfo.classList.add("logged-in");
        authInfo.title = "Angemeldet";
        loginBtn.style.display = "none";
        logoutBtn.style.display = "inline-block";
    } else {
        authInfo.classList.remove("logged-in");
        authInfo.title = "Nicht angemeldet";
        loginBtn.style.display = "inline-block";
        logoutBtn.style.display = "none";
    }
}

/**
 * Initialisiert die MSAL-Instanz mit der Konfiguration aus dem CONFIG-Objekt.
 * Verarbeitet bestehende Redirect-Ergebnisse und stellt aktive Accounts wieder her.
 */
async function initializeMsal() {
    msalInstance = new PublicClientApplication({
        auth: {
            clientId: CONFIG.msal.clientId,
            authority: `https://login.microsoftonline.com/${CONFIG.msal.tenantId}`,
            redirectUri: window.location.origin + window.location.pathname
        },
        cache: {
            cacheLocation: 'localStorage'
        }
    });

    msalInitPromise = msalInstance.initialize();
    await msalInitPromise;

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
        activeAccount = redirectResult.account;
        msalInstance.setActiveAccount(activeAccount);
    }

    const accounts = msalInstance.getAllAccounts();
    if (!activeAccount && accounts.length > 0) {
        activeAccount = accounts[0];
        msalInstance.setActiveAccount(activeAccount);
    }
    updateAuthDisplay();
}

/**
 * Stellt sicher, dass MSAL initialisiert ist.
 * @throws {Error} Wenn MSAL noch nicht initialisiert wurde
 */
async function ensureMsal() {
    if (!msalInstance) {
        throw new Error('MSAL ist noch nicht initialisiert. Bitte Tenant-ID und Client-ID speichern.');
    }
    if (msalInitPromise) {
        await msalInitPromise;
    }
}

/**
 * Ermittelt ein Access Token für Microsoft Graph.
 * Versucht zunächst silent, bei Fehler wird ein Popup geöffnet.
 *
 * @returns {Promise<string>} Access Token
 * @throws {Error} Wenn kein Account angemeldet ist
 */
async function getAccessToken() {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error('Nicht angemeldet.');
    }

    const response = await msalInstance.acquireTokenSilent({
        scopes: CONFIG.msal.scopes,
        account
    }).catch(async () => {
        return msalInstance.acquireTokenPopup({
            scopes: CONFIG.msal.scopes,
            account
        });
    });

    return response.accessToken;
}

/**
 * Führt einen GET-Request an die SharePoint Graph API aus.
 * Zentrale Stelle für Headers und Error-Handling.
 *
 * @param {string} accessToken - Microsoft Graph Access Token
 * @param {string} endpoint - Vollständige API-Endpoint-URL
 * @returns {Promise<Array>} Array der value-Items aus der API-Response
 */
async function callSharePointAPI(accessToken, endpoint) {
    const response = await fetch(endpoint, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json",
            "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"
        }
    });

    const text = await response.text();

    if (!response.ok) {
        throw new Error(`SharePoint API Fehler: ${response.status}\n${text}`);
    }

    return JSON.parse(text).value || [];
}

/**
 * Lädt Items aus der konfigurierten SharePoint-Liste.
 *
 * @param {string} accessToken - Microsoft Graph Access Token
 * @param {string|null} filter - Optionaler OData-Filter (z.B. "pruefid eq '...'")
 * @returns {Promise<Array>} Array der SharePoint List Items
 */
async function getSharePointListItems(accessToken, filter = null) {
    const endpoint =
        `https://graph.microsoft.com/v1.0/sites/${CONFIG.sharePoint.siteId}` +
        `/lists/${CONFIG.sharePoint.listId}/items` +
        `?$expand=fields` +
        (filter ? `&$filter=fields/${filter}` : '');

    return callSharePointAPI(accessToken, endpoint);
}

/**
 * Lädt eine PDF-Datei nach SharePoint hoch und verknüpft sie mit einer Prüf-ID.
 *
 * @param {File} file - Die hochzuladende PDF-Datei
 * @param {string} accessToken - Microsoft Graph Access Token
 * @param {string} pruefId - Eindeutige Prüf-ID für die Zuordnung
 * @returns {Promise<Object>} Upload-Ergebnis von SharePoint
 */
async function uploadSmallFileToSharePoint(file, accessToken, pruefId) {

    const fileName = sanitizeFileName(file.name);

    const endpoint = `https://graph.microsoft.com/v1.0/sites/${CONFIG.sharePoint.siteId}/drive/root:/${encodePath(fileName)}:/content`;

    const uploadResponse = await fetch(endpoint, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': file.type || 'application/pdf'
        },
        body: file
    });

    if (!uploadResponse.ok) {
        throw new Error('Upload fehlgeschlagen');
    }

    const uploadResult = await uploadResponse.json();
    const driveItemId = uploadResult.id;

    // Metadaten setzen
    const metadataEndpoint = `https://graph.microsoft.com/v1.0/sites/${CONFIG.sharePoint.siteId}/drive/items/${driveItemId}/listItem/fields`;

    await fetch(metadataEndpoint, {
        method: 'PATCH',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            pruef_id: pruefId
        })
    });

    return uploadResult;
}

/**
 * Rendert die Liste der ausgewählten Dateien mit Status-Anzeige.
 */
function renderFileList() {
    fileList.innerHTML = '';

    if (selectedFiles.length === 0) {
        fileList.innerHTML = '<p class="small">Keine Dateien ausgewählt.</p>';
        return;
    }

    selectedFiles.forEach((item, index) => {
        const row = document.createElement('div');
        row.className = 'file-row';

        const statusLabels = {
            bereit: 'Bereit',
            uploading: 'Lädt...',
            success: 'Hochgeladen',
            error: 'Fehler'
        };

        row.innerHTML = `
          <div class="file-left">
            <strong>${index + 1}. ${escapeHtml(item.file.name)}</strong>
            <div class="small">${formatBytes(item.file.size)} • ${item.file.type || 'unbekannter Typ'}</div>
          </div>
          <div class="file-right">
            <span class="file-status ${item.status}">${statusLabels[item.status] || item.status}</span>
            <button class="danger" onclick="removeFile(${index})">Entfernen</button>
          </div>
        `;
        fileList.appendChild(row);
    });
}

/**
 * Sucht einen spezifischen Prüfvorgang nach Prüf-ID.
 *
 * @param {string} accessToken - Microsoft Graph Access Token
 * @param {string} pruefId - Die zu suchende Prüf-ID
 * @returns {Promise<Array>} Array der gefundenen List-Items
 */
async function getPruefvorgangByPruefId(accessToken, pruefId) {
    return getSharePointListItems(accessToken, `pruefid eq '${pruefId}'`);
}

/**
 * Lädt alle Prüfvorgänge aus SharePoint und dedupliziert nach Auftragsnummer.
 * Aktualisiert das globale pruefvorgaenge-Array und rendert die Suchergebnisse.
 *
 * @param {string} accessToken - Microsoft Graph Access Token
 */
async function loadPruefvorgaenge(accessToken) {
    const rawItems = await getSharePointListItems(accessToken);

    const alleEintraege = rawItems.map(item => ({
        id: item.id,
        pruefId: item.fields.pruefid,
        auftragsnummer: item.fields.auftragsnummer || item.fields.Title || "Ohne Auftragsnummer",
        raw: item
    }));

    const uniqueMap = new Map();

    for (const eintrag of alleEintraege) {
        if (!uniqueMap.has(eintrag.auftragsnummer)) {
            uniqueMap.set(eintrag.auftragsnummer, eintrag);
        }
    }

    pruefvorgaenge = Array.from(uniqueMap.values());

    renderSearchResults(pruefvorgaenge);
}

/**
 * Rendert die Suchergebnisse als klickbare Liste.
 *
 * @param {Array} items - Array von Prüfvorgang-Objekten
 */
function renderSearchResults(items) {
    searchResults.innerHTML = "";

    if (!items || items.length === 0) {
        searchResults.textContent = "Keine Prüfvorgänge gefunden.";
        return;
    }

    items.forEach(item => {
        const row = document.createElement("div");
        row.className = "file-row";

        row.innerHTML = `
            <div class="file-left">
                <strong>${escapeHtml(item.auftragsnummer)}</strong>
                <div class="small">Prüf-ID: ${escapeHtml(item.pruefId || "")}</div>
            </div>
            <div class="file-right">
                <button onclick="selectPruefvorgang('${item.pruefId}')">Auswählen</button>
            </div>
        `;

        searchResults.appendChild(row);
    });
}

/**
 * Sucht Fehlerpositionen für einen Prüfvorgang nach Prüf-ID.
 *
 * @param {string} accessToken - Microsoft Graph Access Token
 * @param {string} pruefId - Die Prüf-ID des Vorgangs
 * @returns {Promise<Array>} Array der Fehlerpositionen-Items
 */
async function getFehlerpositionenByPruefId(accessToken, pruefId) {
    return getSharePointListItems(accessToken, `pruefid eq '${pruefId}'`);
}

/**
 * Rendert die Analyseergebnisse als HTML-Tabelle.
 *
 * @param {Array} items - Array von Ergebnis-Items mit fields-Objekt
 */
function renderResultTable(items) {
    if (!items || items.length === 0) {
        resultTable.innerHTML = "<p class='small'>Keine Ergebnisse gefunden.</p>";
        return;
    }

    let html = `
        <table class="result-table">
            <thead>
                <tr>
                    <th>Auftrag</th>
                    <th>Position</th>
                    <th>Pos.-Ampel</th>
                    <th>Cluster</th>
                    <th>Adresse</th>
                    <th>Hinweis Position</th>
                    <th>Hinweis Cluster</th>
                    <th>Hinweis Adresse</th>
                </tr>
            </thead>
            <tbody>
    `;

    for (const item of items) {
        const f = item.fields;

        html += `
            <tr>
                <td>${escapeHtml(f.auftragsnummer || "")}</td>
                <td>${escapeHtml(f.Title || "")}</td>
                <td>${renderAmpel(f.status_ampel)}</td>
                <td>${renderAmpel(f.status_Ampel_Cluster)}</td>
                <td>${renderAmpel(f.status_ampel_Adresse)}</td>
                <td>${escapeHtml(f.Hinweis_pos || "")}</td>
                <td>${escapeHtml(f.Hinweis_Cluster || "")}</td>
                <td>${escapeHtml(f.Hinweis_Adresse || "")}</td>
            </tr>
        `;
    }

    html += `
            </tbody>
        </table>
    `;

    resultTable.innerHTML = html;
}

/**
 * Rendert einen Ampel-Status als farbiges Badge.
 *
 * @param {string} value - Ampel-Wert (grün, gelb, rot oder anderer Wert)
 * @returns {string} HTML-Span-Element mit entsprechender CSS-Klasse
 */
function renderAmpel(value) {
    const text = value || "";
    const lower = text.toLowerCase();

    let cssClass = "default";
    if (lower === "grün" || lower === "gruen") cssClass = "gruen";
    else if (lower === "gelb") cssClass = "gelb";
    else if (lower === "rot") cssClass = "rot";

    return `<span class="ampel ${cssClass}">${escapeHtml(text)}</span>`;
}

/**
 * Zeigt eine Statusmeldung in der Status-Box an.
 * @param {string} text - Die anzuzeigende Meldung
 */
function setStatus(text) {
    statusBox.textContent = text;
}

/**
 * Formatiert Byte-Größen in eine lesbare Darstellung.
 * @param {number} bytes - Größe in Bytes
 * @returns {string} Formatierte Größe (z.B. "1.50 MB")
 */
function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return `${(bytes / Math.pow(1024, i)).toFixed(2)} ${sizes[i]}`;
}

/**
 * Entfernt für SharePoint ungültige Zeichen aus Dateinamen.
 * @param {string} name - Original-Dateiname
 * @returns {string} Bereinigter Dateiname
 */
function sanitizeFileName(name) {
    return name.replace(/["#%*:<>?\\/|]/g, '_');
}

/**
 * Encodiert URL-Pfad-Segmente für SharePoint-Endpunkte.
 * @param {string} path - Dateipfad
 * @returns {string} URL-encodierter Pfad
 */
function encodePath(path) {
    return path.split('/').map(segment => encodeURIComponent(segment)).join('/');
}

/**
 * Escapt HTML-Sonderzeichen zum Schutz vor XSS.
 * @param {string} value - Unescapter String
 * @returns {string} HTML-escapter String
 */
function escapeHtml(value) {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

window.removeFile = function (index) {
    selectedFiles.splice(index, 1);
    renderFileList();
};

function formatError(error) {
    return error?.message || String(error);
}

showScreen("upload");