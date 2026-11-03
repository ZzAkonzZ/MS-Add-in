/*
 * PC Data Bridge — прототип: получить ФИО и почту текущего пользователя по кнопке.
 *
 * Порядок попыток:
 *   1. NAA (MSAL) — только если указан CLIENT_ID (нужна регистрация приложения в Entra ID).
 *      Работает и в Excel, и в PowerPoint.
 *   2. Excel: временный комментарий в ячейке — Excel сам подставляет автора (имя + почту),
 *      мы читаем их и сразу удаляем комментарий. Никаких регистраций не нужно.
 *   3. Ручной ввод (запоминается в этом браузере).
 */

// Оставьте пустым, пока нет регистрации приложения в Entra ID.
// Если зарегистрируете (SPA, redirect URI: brk-multihub://zzakonzz.github.io) — вставьте сюда Client ID.
const CLIENT_ID = "";

const STORAGE_KEY = "pcDataBridge.profile";

let currentHost = null;

Office.onReady((info) => {
    currentHost = info.host;
    const supported = info.host === Office.HostType.Excel || info.host === Office.HostType.PowerPoint;

    document.getElementById("btnCheckProfile").onclick = checkProfile;
    document.getElementById("btnSaveManual").onclick = saveManualProfile;

    if (!supported) {
        showStatus("Надстройка запущена вне Excel/PowerPoint — доступен только ручной ввод.", "warn");
    }
});

async function checkProfile() {
    const card = document.getElementById("outputCard");
    card.classList.remove("hidden");
    document.getElementById("manualForm").classList.add("hidden");

    setText("userPlatform", `${currentHost || "—"} / ${Office.context.platform || "—"} (${(Office.context.diagnostics && Office.context.diagnostics.version) || "—"})`);
    setText("userName", "—");
    setText("userEmail", "—");
    setText("userSource", "—");

    const log = [];
    const attempts = [
        ["NAA (Entra ID)", tryNaa],
        ["Комментарий Excel", tryExcelComment],
        ["Сохранено ранее вручную", tryStoredProfile],
    ];

    for (const [label, fn] of attempts) {
        showStatus(`Пробую: ${label}...`, "info");
        try {
            const profile = await fn();
            if (profile && (profile.name || profile.email)) {
                showProfile(profile, label);
                log.push(`✔ ${label}`);
                setText("diagLog", log.join("\n"));
                showStatus("Профиль определён.", "ok");
                return;
            }
            log.push(`– ${label}: пропущено`);
        } catch (e) {
            log.push(`✖ ${label}: ${(e && (e.code || e.errorCode)) || ""} ${(e && e.message) || e}`);
        }
    }

    setText("diagLog", log.join("\n"));
    document.getElementById("manualForm").classList.remove("hidden");
    showStatus("Автоматически определить не удалось — введите ФИО и почту вручную.", "warn");
}

/* ---------- 1. NAA через MSAL ---------- */
async function tryNaa() {
    if (!CLIENT_ID) return null;
    if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
        throw new Error("NestedAppAuth 1.1 не поддерживается в этом клиенте");
    }
    const pca = await msal.createNestablePublicClientApplication({
        auth: { clientId: CLIENT_ID, authority: "https://login.microsoftonline.com/common" },
    });
    const request = { scopes: ["User.Read"] };
    let result;
    try {
        result = await pca.acquireTokenSilent(request);
    } catch (silentError) {
        result = await pca.acquireTokenPopup(request);
    }
    return { name: result.account.name, email: result.account.username };
}

/* ---------- 2. Excel: автор временного комментария ---------- */
async function tryExcelComment() {
    if (currentHost !== Office.HostType.Excel) return null;
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
        throw new Error("ExcelApi 1.10 (комментарии) не поддерживается");
    }

    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // Последняя ячейка листа — там почти наверняка нет своего комментария.
        const cell = sheet.getRange("XFD1048576");
        const comment = context.workbook.comments.add(cell, "PC Data Bridge: проверка профиля (удаляется автоматически)");
        comment.load("authorName,authorEmail");
        try {
            await context.sync();
            return { name: comment.authorName, email: comment.authorEmail };
        } finally {
            comment.delete();
            await context.sync();
        }
    });
}

/* ---------- 3. Ручной ввод ---------- */
async function tryStoredProfile() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        return raw ? JSON.parse(raw) : null;
    } catch (e) {
        return null;
    }
}

function saveManualProfile() {
    const profile = {
        name: document.getElementById("manualName").value.trim(),
        email: document.getElementById("manualEmail").value.trim(),
    };
    if (!profile.name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(profile.email)) {
        showStatus("Укажите ФИО и корректную почту.", "warn");
        return;
    }
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(profile));
    } catch (e) {
        /* хранилище недоступно — просто покажем без сохранения */
    }
    document.getElementById("manualForm").classList.add("hidden");
    showProfile(profile, "Ручной ввод");
    showStatus("Профиль сохранён.", "ok");
}

/* ---------- UI ---------- */
function showProfile(profile, source) {
    setText("userName", profile.name || "—");
    setText("userEmail", profile.email || "—");
    setText("userSource", source);
}

function setText(id, text) {
    document.getElementById(id).textContent = text;
}

function showStatus(text, kind) {
    const el = document.getElementById("statusMessage");
    el.textContent = text;
    el.className = `status ${kind || ""}`;
}
