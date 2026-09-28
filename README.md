# SMFWorks Outlook Plugin

Hermes Desktop pane for Outlook on the web. Plugin id: `hermes-outlook`.

There is no official Hermes Desktop plugin for Outlook. This repo is the pane: Outlook on the web, inside Hermes, not a separate window.

## What it does

- Sidebar **Outlook** page in Hermes Desktop
- Hosts Outlook on the web (`https://outlook.office.com/mail/`) in the pane
- Sign-in stays in that pane. The session is kept in `persist:hermes-outlook`
- **Pop out** opens the same web app in a browser
- Names the signed-in account from Microsoft Graph via Azure CLI, when that token is present. Mail is not read through Graph

A normal `az login` often cannot read mailbox contents. The pane does not invent an empty inbox. Outlook on the web is the mail client.

Launcher links (`/l/`, `/dl/`) open the desktop app and do nothing inside the pane. Those are sent to the web client instead. Other navigations, including sign-in, are left alone.

## Install

Desktop UI is app-global. The Python API loads only when the plugin is in the **active profile's** `plugins.enabled`. Routes mount at backend start. The web client does not wait on that API.

```text
HERMES_HOME/desktop-plugins/hermes-outlook/plugin.js
HERMES_HOME/plugins/hermes-outlook/          (Python API)
```

The folder name must equal the plugin id.

Windows, from this checkout:

```text
copy desktop\plugin.js  %LOCALAPPDATA%\hermes\desktop-plugins\hermes-outlook\plugin.js
xcopy /E /I .           %LOCALAPPDATA%\hermes\plugins\hermes-outlook\
```

Then:

1. `hermes plugins enable hermes-outlook` with `HERMES_HOME` set to the profile that owns the window
2. Restart that backend if you want the account name in the header
3. Command palette: **Reload desktop plugins**
4. Sidebar → **Outlook**

`install.sh` does the Unix copy and enable.

## Not this plugin

- The Hermes Teams gateway bot
- A Graph mail inbox. Mail stays in Outlook on the web
