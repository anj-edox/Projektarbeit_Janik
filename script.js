import { PublicClientApplication } from 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.0/+esm';

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

const siteId = "edox.sharepoint.com,e68093ae-074a-4851-88dd-31284b07b284,e94104cd-b270-4e09-b29a-1932cc934dbf";

const graphScopes = ['User.Read', 'Files.ReadWrite', 'Sites.ReadWrite.All'];

await loadSavedConfig();
renderFileList();
renderAuthInfo();
updateAuthUI();

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
            renderAuthInfo();
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

        renderAuthInfo();
        updateAuthUI();

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
                item.status = '⏳ lädt...';
                renderFileList();

                await uploadSmallFileToSharePoint(item.file, token, pruefId);

                item.status = '✔ hochgeladen';
                renderFileList();
            } catch (err) {
                item.status = '❌ Fehler';
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

function updateAuthUI() {
    const account = msalInstance.getActiveAccount();

    if (account) {
        loginBtn.style.display = "none";
        logoutBtn.style.display = "inline-block";
    } else {
        loginBtn.style.display = "inline-block";
        logoutBtn.style.display = "none";
    }
}

async function initializeMsal() {
    const ClientID = "9603fe2f-fad8-4a18-9950-14f1eb195745";
    const TenantID = "2b71e8a3-7c31-40db-84b0-9c7e50b9ea71";

    msalInstance = new PublicClientApplication({
        auth: {
            clientId: ClientID,
            authority: `https://login.microsoftonline.com/${TenantID}`,
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
    renderAuthInfo();
}

async function ensureMsal() {
    if (!msalInstance) {
        throw new Error('MSAL ist noch nicht initialisiert. Bitte Tenant-ID und Client-ID speichern.');
    }
    if (msalInitPromise) {
        await msalInitPromise;
    }
}

async function getAccessToken() {
    const account = msalInstance.getActiveAccount();
    if (!account) {
        throw new Error('Nicht angemeldet.');
    }

    const response = await msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account
    }).catch(async () => {
        return msalInstance.acquireTokenPopup({
            scopes: graphScopes,
            account
        });
    });

    return response.accessToken;
}

async function uploadSmallFileToSharePoint(file, accessToken, pruefId) {

    const fileName = sanitizeFileName(file.name);

    const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodePath(fileName)}:/content`;

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
    const metadataEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${driveItemId}/listItem/fields`;

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

function renderFileList() {
    fileList.innerHTML = '';

    if (selectedFiles.length === 0) {
        fileList.innerHTML = '<p class="small">Keine Dateien ausgewählt.</p>';
        return;
    }

    selectedFiles.forEach((item, index) => {
        const row = document.createElement('div');
        row.className = 'file-row';
        row.innerHTML = `
          <div class="file-left">
            <strong>${index + 1}. ${escapeHtml(item.file.name)}</strong>
            <div class="small">${formatBytes(item.file.size)} • ${item.file.type || 'unbekannter Typ'}</div>
          </div>
          <div class="file-right">
            <span class="small ready">${item.status}</span>
            <button onclick="removeFile(${index})">❌</button>
          </div>
        `;
        fileList.appendChild(row);
    });
}

function renderAuthInfo() {
    const account = msalInstance?.getActiveAccount() || activeAccount;

    if (!account) {
        authInfo.classList.remove("logged-in");
        authInfo.title = "Nicht angemeldet";
        updateAuthUI();
        return;
    }

    authInfo.classList.add("logged-in");
    authInfo.title = "Angemeldet";
    updateAuthUI();
}

async function getPruefvorgangByPruefId(accessToken, pruefId) {

    const listId = "ecac4b28-01fb-489a-ad32-f5ee4fd4bd83";

    const endpoint =
        `https://graph.microsoft.com/v1.0/sites/${siteId}` +
        `/lists/${listId}/items` +
        `?$expand=fields` +
        `&$filter=fields/pruefid eq '${pruefId}'`;

    const response = await fetch(endpoint, {
        method: "GET",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json",
            "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"
        }
    });

}

async function loadPruefvorgaenge(accessToken) {
    const listId = "ecac4b28-01fb-489a-ad32-f5ee4fd4bd83";

    const endpoint =
        `https://graph.microsoft.com/v1.0/sites/${siteId}` +
        `/lists/${listId}/items?$expand=fields`;

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
        throw new Error(`Prüfvorgänge konnten nicht geladen werden: ${response.status}\n${text}`);
    }

    const json = JSON.parse(text);

    const alleEintraege = json.value.map(item => ({
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

async function getFehlerpositionenByPruefId(accessToken, pruefId) {
    const listId = "ecac4b28-01fb-489a-ad32-f5ee4fd4bd83";

    const endpoint =
        `https://graph.microsoft.com/v1.0/sites/${siteId}` +
        `/lists/${listId}/items` +
        `?$expand=fields` +
        `&$filter=fields/pruefid eq '${pruefId}'`;

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
        throw new Error(`Ergebnisse konnten nicht geladen werden: ${response.status}\n${text}`);
    }

    return JSON.parse(text).value || [];
}

function renderResultTable(items) {
    if (!items || items.length === 0) {
        resultTable.innerHTML = "<p class='small'>Keine Ergebnisse gefunden.</p>";
        return;
    }

    let html = `
        <table style="width:100%; border-collapse: collapse; font-size:14px;">
            <thead>
                <tr>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Auftrag</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Position</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Pos.-Ampel</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Cluster</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Adresse</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Hinweis Position</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Hinweis Cluster</th>
                    <th style="text-align:left; padding:8px; border-bottom:1px solid #ddd;">Hinweis Adresse</th>
                </tr>
            </thead>
            <tbody>
    `;

    for (const item of items) {
        const f = item.fields;

        html += `
            <tr>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${escapeHtml(f.auftragsnummer || "")}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${escapeHtml(f.Title || "")}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${renderAmpel(f.status_ampel)}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${renderAmpel(f.status_Ampel_Cluster)}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${renderAmpel(f.status_ampel_Adresse)}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${escapeHtml(f.Hinweis_pos || "")}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${escapeHtml(f.Hinweis_Cluster || "")}</td>
                <td style="vertical-align:top; padding:8px; border-bottom:1px solid #eee;">${escapeHtml(f.Hinweis_Adresse || "")}</td>
            </tr>
        `;
    }

    html += `
            </tbody>
        </table>
    `;

    resultTable.innerHTML = html;
}

function renderAmpel(value) {
    const text = value || "";

    let color = "#6b7280";
    if (text.toLowerCase() === "grün") color = "#16a34a";
    if (text.toLowerCase() === "gelb") color = "#f59e0b";
    if (text.toLowerCase() === "rot") color = "#dc2626";

    return `
        <span style="
            display:inline-block;
            padding:4px 10px;
            border-radius:999px;
            color:white;
            background:${color};
            font-size:12px;
            font-weight:bold;
        ">
            ${escapeHtml(text)}
        </span>
    `;
}

function setStatus(text) {
    statusBox.textContent = text;
}

function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(1024));
    return `${(bytes / Math.pow(1024, i)).toFixed(2)} ${sizes[i]}`;
}

function sanitizeFileName(name) {
    return name.replace(/["#%*:<>?\\/|]/g, '_');
}

function encodePath(path) {
    return path.split('/').map(segment => encodeURIComponent(segment)).join('/');
}

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