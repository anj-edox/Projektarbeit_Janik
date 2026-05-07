# OrderCheck App

PDF-basierte Bestellprüfung mit SharePoint-Integration und Microsoft Teams SSO.

## Architektur

- **Frontend:** Vanilla JavaScript (ES Modules)
- **Auth:** Microsoft Teams SSO via MSAL Browser (OAuth 2.0 / OpenID Connect)
- **Backend:** SharePoint Online + Microsoft Graph API
- **Verarbeitung:** Power Automate Flow verarbeitet hochgeladene PDFs

## Ablauf

1. **Authentifizierung:** User klickt Login → Teams SSO Popup (`auth-start.html`) → Auth-Callback (`auth-end.html`) → MSAL speichert Account
2. **Upload:** PDFs werden nach SharePoint hochgeladen und mit einer `pruef_id` versehen
3. **Verarbeitung:** Power Automate Flow erkennt neue PDFs, analysiert sie und schreibt Ergebnisse in eine SharePoint-Liste
4. **Ergebnis:** App pollt die SharePoint-Liste nach Ergebnissen und zeigt sie tabellarisch mit Ampel-Status

## Dateistruktur

```
orderCheckApp/
├── index.html          # Hauptanwendung mit 3 Screens (Upload, Suche, Ergebnis)
├── script.js           # App-Logik (Auth, API-Calls, UI-Rendering)
├── style.css           # Styling
├── config.js           # Referenz-Datei für Konfigurationswerte (nicht importiert)
├── auth-start.html     # Teams Auth Popup: startet Login-Redirect
└── auth-end.html       # Teams Auth Callback: verarbeitet Login-Ergebnis
```

## Screens

| Screen    | Funktion |
|-----------|----------|
| **Upload** | PDFs auswählen und nach SharePoint hochladen |
| **Suche**  | Prüfvorgänge nach Auftragsnummer durchsuchen |
| **Ergebnis** | Analyseergebnisse als Tabelle mit Ampel-Status anzeigen |

## Setup

### Voraussetzungen

- Azure AD Tenant mit registrierter App (Public Client / SPA)
- SharePoint Online Site mit konfigurierter Liste
- Power Automate Flow zur PDF-Verarbeitung

### Konfiguration

Alle IDs sind in `config.js` zentralisiert:

```javascript
const CONFIG = {
    msal: {
        clientId: "...",      // Azure AD App Client ID
        tenantId: "...",      // Azure AD Tenant ID
        scopes: [...]         // Microsoft Graph Scopes
    },
    sharePoint: {
        siteId: "...",        // SharePoint Site ID
        listId: "..."         // SharePoint List ID
    }
};
```

### Azure AD App-Registrierung

- **Redirect URIs:** `https://<deine-domain>/auth-end.html`
- **API-Berechtigungen:** `User.Read`, `Files.ReadWrite`, `Sites.ReadWrite.All`
- **Authentifizierungsfluss:** Implicit Grant + Authorization Code (für SPA)

## Technologie-Entscheidungen

| Entscheidung | Begründung |
|--------------|------------|
| **Vanilla JS statt Framework** | Leichtgewichtig, keine Build-Tools nötig, direkt in Teams hostbar |
| **ES Modules** | Native Modul-Unterstützung im Browser, Config-Sharing zwischen Dateien |
| **MSAL Browser** | Offizielle Microsoft-Bibliothek für OAuth/OIDC im Browser |
| **localStorage Cache** | Persistenter Auth-Cache, vermeidet Re-Login bei Seiten-Reload |
| **Polling statt WebHooks** | Einfachere Implementierung; für Produktivbetrieb WebHooks empfehlenswert |

## Bekannte Einschränkungen

- Upload beschränkt auf kleine Dateien (< 4 MB) via Graph API PUT
- Polling-Intervall: 5 Sekunden, max. 10 Versuche (konfigurierbar in `waitForResult`)
- Keine clientseitige Validierung der PDF-Inhalte
- Config-Werte sind client-seitig sichtbar (Public Client - akzeptabel für dieses Szenario)
