// ==UserScript==
// @name        SharePoint Custom Buttons
// @namespace   Sharepoint Scripts
// @match       https://hpwessen.sharepoint.com/**
// @grant       none
// @version     1.4
// @author      Vivian Klein
// @description Fügt client-seitige Buttons zum Aufrufen der Ansichten "Seiten", "Dokumente" und "Websiteinhalte" hinzu
// @run-at      document-start
// @icon   https://upload.wikimedia.org/wikipedia/commons/thumb/e/ec/Microsoft_Office_SharePoint_%282019%E2%80%932025%29.svg/250px-Microsoft_Office_SharePoint_%282019%E2%80%932025%29.svg.png
// ==/UserScript==

(function() {
    'use strict';

    // Version Management
    const CURRENT_VERSION = "1.4";
    const STORAGE_KEY = "sp-custom-buttons-version";

    // Changelog - Add new versions at the top
    const CHANGELOG = {
        "1.4": {
            date: "2025-11-03",
            changes: [
                "Kleinere Verbesserungen für Release 1.4"
            ]
        },
        "1.3": {
            date: "2025-11-03",
            changes: [
                "Versionsmanagement hinzugefügt",
                "Fix: In einigen Fällen wurde der falsche Kontext durch die Buttons geöffnet"
            ]
        },
        "1.2": {
            date: "2025-10-15",
            changes: [
                "Fix: Hook für Buttons wurde in manchen Browsern zur falschen Zeit ausgeführt"
            ]
        }
    };

    let buttonsAdded = false;

    // Version & Changelog Functions
    function getLastVersion() {
        return localStorage.getItem(STORAGE_KEY);
    }

    function setLastVersion(version) {
        localStorage.setItem(STORAGE_KEY, version);
    }

    function compareVersions(v1, v2) {
        const parts1 = v1.split('.').map(Number);
        const parts2 = v2.split('.').map(Number);
        for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
            const num1 = parts1[i] || 0;
            const num2 = parts2[i] || 0;
            if (num1 > num2) return 1;
            if (num1 < num2) return -1;
        }
        return 0;
    }

    function getNewChanges(lastVersion) {
        const changes = [];
        for (const [version, data] of Object.entries(CHANGELOG)) {
            if (!lastVersion || compareVersions(version, lastVersion) > 0) {
                changes.push({ version, ...data });
            }
        }
        return changes.sort((a, b) => compareVersions(b.version, a.version));
    }

    function showChangelog(changes) {
        // Create overlay
        const overlay = document.createElement('div');
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 99999;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        `;

        // Create popup
        const popup = document.createElement('div');
        popup.style.cssText = `
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            max-width: 600px;
            width: 90%;
            max-height: 80vh;
            overflow: hidden;
            display: flex;
            flex-direction: column;
        `;

        // Header
        const header = document.createElement('div');
        header.style.cssText = `
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
            padding: 20px;
            font-size: 20px;
            font-weight: 600;
        `;
        header.textContent = '🎉 SharePoint Custom Buttons - Was ist neu?';

        // Content
        const content = document.createElement('div');
        content.style.cssText = `
            padding: 20px;
            overflow-y: auto;
            flex: 1;
        `;

        changes.forEach(({ version, date, changes: changeList }) => {
            const versionBlock = document.createElement('div');
            versionBlock.style.cssText = `
                margin-bottom: 24px;
            `;

            const versionHeader = document.createElement('div');
            versionHeader.style.cssText = `
                display: flex;
                align-items: center;
                gap: 10px;
                margin-bottom: 10px;
            `;

            const versionBadge = document.createElement('span');
            versionBadge.style.cssText = `
                background: #0078d4;
                color: white;
                padding: 4px 12px;
                border-radius: 12px;
                font-weight: 600;
                font-size: 14px;
            `;
            versionBadge.textContent = `v${version}`;

            const dateBadge = document.createElement('span');
            dateBadge.style.cssText = `
                color: #666;
                font-size: 14px;
            `;
            dateBadge.textContent = date;

            versionHeader.appendChild(versionBadge);
            versionHeader.appendChild(dateBadge);

            const changesList = document.createElement('ul');
            changesList.style.cssText = `
                margin: 0;
                padding-left: 20px;
                color: #333;
            `;

            changeList.forEach(change => {
                const li = document.createElement('li');
                li.style.cssText = `
                    margin: 6px 0;
                    line-height: 1.5;
                `;
                li.textContent = change;
                changesList.appendChild(li);
            });

            versionBlock.appendChild(versionHeader);
            versionBlock.appendChild(changesList);
            content.appendChild(versionBlock);
        });

        // Footer
        const footer = document.createElement('div');
        footer.style.cssText = `
            padding: 20px;
            border-top: 1px solid #e0e0e0;
            display: flex;
            justify-content: flex-end;
        `;

        const okButton = document.createElement('button');
        okButton.style.cssText = `
            background: #0078d4;
            color: white;
            border: none;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
        `;
        okButton.textContent = 'OK';
        okButton.onmouseover = () => okButton.style.background = '#106ebe';
        okButton.onmouseout = () => okButton.style.background = '#0078d4';
        okButton.onclick = () => {
            setLastVersion(CURRENT_VERSION);
            document.body.removeChild(overlay);
        };

        footer.appendChild(okButton);

        // Assemble popup
        popup.appendChild(header);
        popup.appendChild(content);
        popup.appendChild(footer);
        overlay.appendChild(popup);

        // Add to page
        document.body.appendChild(overlay);

        // Close on overlay click
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) {
                setLastVersion(CURRENT_VERSION);
                document.body.removeChild(overlay);
            }
        });
    }

    function checkVersion() {
        const lastVersion = getLastVersion();
        const newChanges = getNewChanges(lastVersion);
        
        if (newChanges.length > 0) {
            console.log('[SharePoint Custom Buttons] New version detected, showing changelog');
            // Wait for page to be fully ready before showing popup
            setTimeout(() => showChangelog(newChanges), 1000);
        } else {
            console.log('[SharePoint Custom Buttons] No new updates');
        }
    }

    function createCustomAnchor(title, idSuffix, relativeUrl, svgPath) {
        const elemToAppend = document.createElement("a");
        elemToAppend.setAttribute("title", title);
        elemToAppend.style.display = "flex";
        elemToAppend.style.justifyContent = "center";
        elemToAppend.style.alignItems = "center";
        elemToAppend.style.color = "rgb(96, 94, 92)";
        elemToAppend.setAttribute("aria-labelledby", `sp-appbar-tooltiplabel-${idSuffix}`);
        elemToAppend.setAttribute("data-interception", "off");
        elemToAppend.setAttribute("id", `sp-appBar-link-${idSuffix}`);
        elemToAppend.setAttribute("role", "button");
        elemToAppend.setAttribute("class", "sp-appBar-link");
        elemToAppend.setAttribute("data-automation-id", `sp-appBar-link-${idSuffix}`);
        elemToAppend.setAttribute("aria-expanded", "false");
        elemToAppend.setAttribute("aria-haspopup", "dialog");

        // Dynamisch URL erzeugen beim Klick (nicht beim Erstellen)
        elemToAppend.addEventListener("click", (e) => {
            e.preventDefault();
            const { origin, pathname } = window.location;
            const pathParts = pathname.split("/").filter(Boolean);
            let href = `${origin}${relativeUrl}`; // fallback
            if (pathParts[0] === "sites" && pathParts.length >= 2) {
                href = `${origin}/sites/${pathParts[1]}${relativeUrl}`;
            }
            window.location.href = href;
        });

        const icon = document.createElement("i");
        icon.setAttribute("aria-hidden", "true");
        icon.style.height = "20px";
        icon.style.width = "20px";
        icon.style.display = "flex";
        icon.style.justifyContent = "center";
        icon.style.alignItems = "center";
        icon.innerHTML = `
    <svg style="width: 100%; height: 100%; max-width: 20px; max-height: 20px;" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
    <path d="${svgPath}" fill="currentColor"></path>
    </svg>
    <div class="sp-appBar-borderLeft"></div>
        `;
        elemToAppend.appendChild(icon);
        return elemToAppend;
    }

    function appendAnchorOnce(classSuffix, anchorElement) {
        const appBar = document.getElementById("sp-appBar");
        if (!appBar || !appBar.children.length) return false;

        const container = appBar.children[0];
        if (!container.querySelector(`a.sp-appBar-link.${classSuffix}`)) {
            anchorElement.classList.add(classSuffix);
            container.append(anchorElement);
            return true;
        }
        return false;
    }

    function addButtons() {
        const appBar = document.getElementById("sp-appBar");

        // Prüfen ob AppBar existiert und mindestens ein Kind hat
        if (!appBar || !appBar.children.length) {
            return false;
        }

        // Prüfen ob bereits hinzugefügt
        if (buttonsAdded || appBar.querySelector('.custom-seiten')) {
            return true;
        }

        console.log('[SharePoint Custom Buttons] Adding buttons...');

        // Button 1 – Beispiel: Seiten
        const anchor1 = createCustomAnchor(
            "Seiten",
            "seiten",
            "/SitePages/Forms/ByAuthor.aspx",
            "m19.5,0h-7c-2.481,0-4.5,2.019-4.5,4.5v7c0,2.481,2.019,4.5,4.5,4.5h7c2.481,0,4.5-2.019,4.5-4.5v-7c0-2.481-2.019-4.5-4.5-4.5Zm1.5,11.5c0,.827-.673,1.5-1.5,1.5h-7c-.827,0-1.5-.673-1.5-1.5v-7c0-.827.673-1.5,1.5-1.5h7c.827,0,1.5.673,1.5,1.5v7Zm-5,8c0,2.481-2.019,4.5-4.5,4.5h-7c-2.481,0-4.5-2.019-4.5-4.5v-7c0-2.481,2.019-4.5,4.5-4.5.828,0,1.5.671,1.5,1.5s-.672,1.5-1.5,1.5-1.5.673-1.5,1.5v7c0,.827.673,1.5,1.5,1.5h7c.827,0,1.5-.673,1.5-1.5s.672-1.5,1.5-1.5,1.5.671,1.5,1.5Z"
        );
        appendAnchorOnce("custom-seiten", anchor1);

        // Button 2 – Beispiel: Dokumente
        const anchor2 = createCustomAnchor(
            "Dokumente",
            "Dokumente",
            "/Freigegebene%20Dokumente/Forms/AllItems.aspx",
            "m6.716 12.216 5.27 5.27c-.166 1.825-.426 2.101-.626 2.301-3.316 3.185-7.244 3.38-8.869 3.138l4.217-4.217c.391-.391.391-1.023 0-1.414s-1.023-.391-1.414 0l-4.22 4.218c-.164-1.537-.012-5.425 3.276-8.733.211-.211.406-.399 2.365-.563zm16.288-7.737c-.032-.133-.334-1.329-1.245-2.238-.909-.91-2.104-1.213-2.237-1.245-.269-.063-.551-.014-.781.139-.12.08-2.989 1.984-6.06 5.054-1.6 1.6-3.129 3.48-4.257 4.906l4.608 4.608c1.395-1.17 3.181-2.785 4.779-4.384 3.07-3.069 4.975-5.938 5.055-6.059.151-.23.201-.512.139-.781z"
        );
        appendAnchorOnce("custom-uebersicht", anchor2);

        // Button 3 – Beispiel: Websiteinhalte
        const anchor3 = createCustomAnchor(
            "Websiteinhalte",
            "Websiteinhalte",
            "/_layouts/15/viewlsts.aspx?view=14",
            "m23.414,16.09l-8.674-3.944c-.03-.014-.06-.025-.09-.036-.728-.248-1.518-.065-2.062.479s-.727,1.334-.442,2.151l3.944,8.675c.159.351.506.579.892.586h.019c.378,0,.725-.214.895-.553l1.851-3.702,3.702-1.851c.345-.173.56-.527.553-.913-.007-.385-.235-.732-.586-.892Zm-10.502,5.869c-.172.015-.345.026-.52.033-.668-.661-2.41-2.518-3.624-4.992h.732c.552,0,1-.447,1-1s-.448-1-1-1h-1.53c-.291-.952-.47-1.96-.47-3s.18-2.047.472-3h8.057c.118.387.218.783.295,1.187.104.543.629.901,1.169.795.542-.103.898-.627.795-1.169-.052-.275-.114-.546-.183-.813h3.393c.013,0,.027,0,.04,0,.299.948.46,1.956.46,3.001,0,.307-.014.61-.041.91-.049.55.356,1.036.906,1.086.552.054,1.037-.356,1.086-.906.032-.359.049-.723.049-1.09C24,5.383,18.617,0,12,0S0,5.383,0,12s5.383,12,11.999,12h0c.367,0,.73-.017,1.088-.049.55-.049.957-.534.908-1.085-.049-.549-.533-.955-1.084-.907Zm-7.02-6.959h-3.392c-.013,0-.026,0-.04,0-.299-.948-.46-1.956-.46-3.001s.161-2.053.46-3.001c.013,0,.026,0,.04,0h3.392c-.245.955-.392,1.959-.392,3s.146,2.045.392,3Zm2.881-8c1.021-2.077,2.415-3.724,3.229-4.588.815.863,2.209,2.507,3.23,4.588h-6.459Zm11.885,0h-3.228c-.791-1.886-1.888-3.488-2.832-4.657,2.579.695,4.749,2.397,6.06,4.657ZM9.4,2.344c-.944,1.169-2.042,2.771-2.832,4.656h-3.226c1.31-2.26,3.479-3.962,6.058-4.656Zm-6.058,14.656h3.226c.79,1.886,1.888,3.487,2.832,4.656-2.578-.695-4.747-2.397-6.058-4.656Z"
        );
        appendAnchorOnce("custom-kontakte", anchor3);

        buttonsAdded = true;
        console.log('[SharePoint Custom Buttons] Buttons added successfully!');
        return true;
    }

    // Observer der auf die AppBar wartet
    const observer = new MutationObserver((mutations) => {
        if (addButtons()) {
            // Buttons erfolgreich hinzugefügt, aber Observer weiterlaufen lassen
            // für den Fall dass SharePoint die AppBar neu lädt (SPA-Navigation)
        }
    });

    // Starte Observer sobald <body> existiert
    function startObserving() {
        if (document.body) {
            observer.observe(document.body, {
                childList: true,
                subtree: true
            });
            console.log('[SharePoint Custom Buttons] Observer started');
            // Ersten Versuch direkt starten
            addButtons();
        } else {
            setTimeout(startObserving, 100);
        }
    }

    // Sofort starten
    startObserving();

    // Zusätzlich: Bei SPA-Navigation Buttons zurücksetzen
    let lastUrl = location.href;
    new MutationObserver(() => {
        const currentUrl = location.href;
        if (currentUrl !== lastUrl) {
            lastUrl = currentUrl;
            buttonsAdded = false;
            console.log('[SharePoint Custom Buttons] URL changed, resetting buttons');
            setTimeout(addButtons, 500);
        }
    }).observe(document, { subtree: true, childList: true });

    // Check for version updates
    checkVersion();

})();
