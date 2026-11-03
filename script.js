Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btnCheckProfile").onclick = checkProfile;
    } else {
        showStatus("Надстройка запущена вне Excel.");
    }
});

async function checkProfile() {
    const outputCard = document.getElementById("outputCard");
    const userName = document.getElementById("userName");
    const userPlatform = document.getElementById("userPlatform");
    const userToken = document.getElementById("userToken");
    
    showStatus("Опрос контекста Office...");

    outputCard.classList.remove("hidden");
    userPlatform.innerText = `${Office.context.platform} (${Office.context.diagnostics.version})`;

    // Попытка 1: Получение базового контекста сессии
    try {
        if (Office.context.mailbox && Office.context.mailbox.userProfile) {
            userName.innerText = `${Office.context.mailbox.userProfile.displayName} (${Office.context.mailbox.userProfile.emailAddress})`;
        } else {
            userName.innerText = "Excel Online Context Active";
        }
    } catch (e) {
        userName.innerText = "Сессия Excel подключена";
    }

    // Попытка 2: Запрос нативного SSO токена Entra ID
    try {
        if (OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
            showStatus("Запрос токена SSO...");
            const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
            
            // Расшифровка payload JWT токена без библиотек
            const base64Url = token.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = decodeURIComponent(atob(base64).split('').map((c) => {
                return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
            }).join(''));

            const userData = JSON.parse(jsonPayload);
            userName.innerText = `${userData.name || userData.preferred_username || "Авторизован"} (${userData.upn || userData.email || ""})`;
            userToken.innerText = JSON.stringify(userData, null, 2);
            showStatus("SSO профиль успешно определен!");
        } else {
            userToken.innerText = "OfficeRuntime.auth недоступен. Работает в режиме базового веб-контекста.";
            showStatus("Готово к отправке данных через Graph/SharePoint API.");
        }
    } catch (error) {
        userToken.innerText = `SSO код ошибки: ${error.code}\nСообщение: ${error.message}\n(Для полного SSO требуется регистрация App ID в Azure AD. Базовые запросы к SharePoint будут работать от текущего браузера).`;
        showStatus("Базовый режим активен.");
    }
}

function showStatus(text) {
    document.getElementById("statusMessage").innerText = text;
}