# PC Data Bridge — прототип надстройки Excel / PowerPoint (веб)

По кнопке «Сдать отчет» → «Получить профиль» надстройка показывает ФИО и почту текущего пользователя.

## Как определяется пользователь

| Способ | Excel | PowerPoint | Что нужно |
|---|---|---|---|
| Комментарий Excel (временный, сразу удаляется) | ✔ | — | ничего |
| NAA / MSAL (Entra ID) | ✔ | ✔ | регистрация приложения, `CLIENT_ID` в `script.js` |
| Ручной ввод (запоминается в браузере) | ✔ | ✔ | ничего |

## Публикация (GitHub Pages)

1. Залить файлы в ветку `main` репозитория `ZzAkonzZ/MS-Add-in`.
2. Settings → Pages → Source: `Deploy from a branch`, Branch: `main` / `(root)`.
3. Проверить, что открывается `https://zzakonzz.github.io/MS-Add-in/index.html` и `.../icon-32.png`.

## Загрузка в Excel / PowerPoint в браузере

1. Открыть файл в Excel/PowerPoint Online (файл должен лежать в OneDrive/SharePoint).
2. Главная → **Надстройки** → **Дополнительные параметры / Мои надстройки** → **Отправить мою надстройку** (Upload My Add-in).
3. Выбрать `manifest.xml`. Появится вкладка **Project Controls**.

Если пункта «Отправить мою надстройку» нет — загрузку отключил администратор, без ИТ тут не обойтись.

После изменений в `manifest.xml` увеличивайте `<Version>` и загружайте манифест заново.
Проверить манифест: `npx office-addin-manifest validate manifest.xml`.

## (Необязательно) Вход через Entra ID — NAA

Если в организации пользователям разрешено регистрировать приложения
(portal.azure.com → Microsoft Entra ID → App registrations → New registration):

1. Supported account types: «Accounts in this organizational directory only».
2. Redirect URI: платформа **Single-page application (SPA)**, значение `brk-multihub://zzakonzz.github.io`.
3. API permissions: `Microsoft Graph → User.Read` (delegated, обычно уже есть).
4. Скопировать **Application (client) ID** в `CLIENT_ID` в `script.js`.
