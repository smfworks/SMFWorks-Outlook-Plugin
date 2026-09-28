/**
 * Hermes Desktop — Outlook (SMF Works).
 * Folder name must equal id (`hermes-outlook`).
 * Mail: Outlook on the web, in the pane. Graph only names the signed-in account.
 */
import {
  Badge,
  Button,
  Codicon,
  PALETTE_AREA,
  ROUTES_AREA,
  SIDEBAR_NAV_AREA,
  Separator,
  Tip,
  haptic,
  host,
  usePluginI18n,
  useQuery
} from '@hermes/plugin-sdk'
import { useEffect, useRef, useState } from 'react'
import { jsx, jsxs } from 'react/jsx-runtime'

const OUTLOOK_HOME = 'https://outlook.office.com/mail/'
const OUTLOOK_HOSTS = new Set([
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.cloud.microsoft',
  'outlook.live.com',
  'outlook.office365.us',
  'substrate.office.com',
  'login.microsoftonline.com',
  'login.microsoft.com',
  'login.live.com',
  'login.windows.net'
])
const CHROME_UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'

const ID = 'hermes-outlook'
const ROUTE = '/outlook'

let pluginCtx = null

function openUrl(url) {
  if (!url) return
  const opener = pluginCtx?.os?.openExternal
  if (typeof opener === 'function') {
    void opener(url)
    return
  }
  window.open(url, '_blank', 'noopener,noreferrer')
}

function isLauncherPath(pathname) {
  return pathname === '/l' || pathname.startsWith('/l/') || pathname.startsWith('/dl/')
}

function outlookUrl(url) {
  if (!url) return null
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:') return null
    if (!OUTLOOK_HOSTS.has(parsed.hostname)) return null
    return parsed.toString()
  } catch {
    return null
  }
}

function webClientUrl(url) {
  const admitted = outlookUrl(url)
  if (!admitted) return null
  try {
    const parsed = new URL(admitted)
    if (isLauncherPath(parsed.pathname)) return OUTLOOK_HOME
    return admitted
  } catch {
    return null
  }
}

function followUrl(url) {
  if (!url) return null
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:') return null
    if (OUTLOOK_HOSTS.has(parsed.hostname) && isLauncherPath(parsed.pathname)) return OUTLOOK_HOME
    return parsed.toString()
  } catch {
    return null
  }
}

const DISMISS_APP_GATE = `(() => {
  const want = ['use the web app instead', 'continue on the web', 'use outlook on the web', 'continue to outlook on the web', 'open in browser']
  const nodes = Array.from(document.querySelectorAll('button, a'))
  const web = nodes.find((el) => {
    const text = (el.innerText || el.textContent || '').toLowerCase()
    return want.some((phrase) => text.indexOf(phrase) !== -1)
  })
  if (!web) return 'none'
  web.click()
  return 'clicked'
})()`

function dismissAppGate(webview) {
  try {
    void webview.executeJavaScript(DISMISS_APP_GATE, false)
  } catch {
    // Guest not ready yet.
  }
}

async function rest(path) {
  if (!pluginCtx?.rest) {
    const err = new Error('backend-off')
    err.backend = false
    throw err
  }
  return pluginCtx.rest(path)
}

function OutlookEmbed({ url, reloadKey }) {
  const hostRef = useRef(null)
  const viewRef = useRef(null)
  const readyRef = useRef(false)
  const urlRef = useRef(url)
  const reloadSeen = useRef(reloadKey)
  urlRef.current = url

  useEffect(() => {
    const node = hostRef.current
    if (!node) return undefined
    readyRef.current = false
    const webview = document.createElement('webview')
    webview.setAttribute('partition', 'persist:hermes-outlook')
    webview.setAttribute('allowpopups', '')
    webview.setAttribute('webpreferences', 'contextIsolation=yes,nodeIntegration=no,sandbox=no')
    webview.setAttribute('useragent', CHROME_UA)
    webview.setAttribute('src', webClientUrl(urlRef.current) || OUTLOOK_HOME)
    webview.style.border = '0'
    webview.style.display = 'flex'
    webview.style.width = '100%'
    webview.style.height = '100%'
    node.replaceChildren(webview)
    viewRef.current = webview

    const fit = () => {
      const rect = node.getBoundingClientRect()
      webview.style.width = Math.max(1, Math.floor(rect.width)) + 'px'
      webview.style.height = Math.max(1, Math.floor(rect.height)) + 'px'
    }
    fit()
    const observer = new ResizeObserver(fit)
    observer.observe(node)

    const go = (next) => {
      const target = webClientUrl(next) || OUTLOOK_HOME
      if (!readyRef.current) {
        webview.setAttribute('src', target)
        return
      }
      try {
        if (webview.getURL() === target) return
      } catch {
        // Guest not attached yet. src is already the web client.
      }
      try {
        webview.loadURL(target)
      } catch {
        webview.setAttribute('src', target)
      }
    }

    const onWillNavigate = (event) => {
      const next = event.url
      if (!next) return
      const target = webClientUrl(next)
      if (!target || target === next) return
      event.preventDefault()
      go(target)
    }

    const onNewWindow = (event) => {
      event.preventDefault()
      const next = followUrl(event.url)
      if (!next) return
      try {
        webview.loadURL(next)
      } catch {
        webview.setAttribute('src', next)
      }
    }

    const onReady = () => {
      readyRef.current = true
      try {
        webview.focus()
      } catch {
        // Focus is best-effort. The guest still receives clicks once attached.
      }
      try {
        webview.setBackgroundThrottling(false)
      } catch {
        // Older Electron builds omit this. Boot still proceeds.
      }
      dismissAppGate(webview)
      let current = ''
      try {
        current = webview.getURL()
      } catch {
        current = ''
      }
      try {
        if (current && isLauncherPath(new URL(current).pathname)) go(OUTLOOK_HOME)
      } catch {
        // A non-URL guest location is not a launcher. Leave it.
      }
    }

    const onStop = () => dismissAppGate(webview)

    webview.addEventListener('dom-ready', onReady)
    webview.addEventListener('did-stop-loading', onStop)
    webview.addEventListener('will-navigate', onWillNavigate)
    webview.addEventListener('new-window', onNewWindow)
    const dismissTimer = setInterval(() => dismissAppGate(webview), 1500)
    const dismissStop = setTimeout(() => clearInterval(dismissTimer), 20000)
    return () => {
      clearInterval(dismissTimer)
      clearTimeout(dismissStop)
      observer.disconnect()
      webview.removeEventListener('dom-ready', onReady)
      webview.removeEventListener('did-stop-loading', onStop)
      webview.removeEventListener('will-navigate', onWillNavigate)
      webview.removeEventListener('new-window', onNewWindow)
      readyRef.current = false
      viewRef.current = null
      node.replaceChildren()
    }
  }, [])

  useEffect(() => {
    if (reloadSeen.current === reloadKey) return
    reloadSeen.current = reloadKey
    const webview = viewRef.current
    if (!webview || !readyRef.current) return
    try {
      webview.reload()
    } catch {
      webview.setAttribute('src', OUTLOOK_HOME)
    }
  }, [reloadKey])

  return jsx('div', { ref: hostRef, className: 'relative min-h-0 min-w-0 flex-1' })
}

function OutlookPage() {
  const t = usePluginI18n(ID)
  const [reloadKey, setReloadKey] = useState(0)
  const status = useQuery({
    queryKey: [ID, 'status'],
    queryFn: () => rest('/status'),
    staleTime: 30 * 1000,
    retry: false
  })
  const me = status.data?.me || {}
  const who = me.displayName || me.userPrincipalName || t('web')

  return jsxs('div', {
    className: 'flex h-full min-h-0 flex-col',
    children: [
      jsxs('div', {
        className: 'flex items-center gap-2 px-3 py-2',
        children: [
          jsx(Codicon, { name: 'mail', className: 'text-(--ui-text-tertiary)' }),
          jsxs('div', {
            className: 'min-w-0 flex-1',
            children: [
              jsx('div', { className: 'truncate text-sm font-medium', children: t('paneTitle') }),
              jsx('div', {
                className: 'truncate text-[0.6875rem] text-(--ui-text-tertiary)',
                children: who
              })
            ]
          }),
          status.data?.ok
            ? jsx(Badge, { children: t('signedIn') })
            : jsx(Badge, { children: t('web') }),
          jsx(Tip, {
            label: t('popTip'),
            children: jsx(Button, {
              size: 'sm',
              variant: 'ghost',
              onClick: () => {
                haptic('tap')
                openUrl(OUTLOOK_HOME)
              },
              children: t('popOut')
            })
          }),
          jsx(Tip, {
            label: t('refreshTip'),
            children: jsx(Button, {
              size: 'sm',
              variant: 'ghost',
              onClick: () => {
                haptic('tap')
                setReloadKey((n) => n + 1)
                void status.refetch()
              },
              children: t('refresh')
            })
          })
        ]
      }),
      jsx(Separator, {}),
      jsx(OutlookEmbed, { url: OUTLOOK_HOME, reloadKey })
    ]
  })
}

export default {
  id: ID,
  name: 'Outlook',
  defaultEnabled: true,
  register(ctx) {
    pluginCtx = ctx
    ctx.i18n.register({
      en: {
        paneTitle: 'Outlook',
        web: 'Outlook on the web',
        signedIn: 'signed in',
        refresh: 'Refresh',
        refreshTip: 'Reload Outlook',
        popOut: 'Pop out',
        popTip: 'Open Outlook on the web in a browser',
        chipTip: 'Outlook'
      }
    })

    ctx.registerMany([
      {
        id: 'page',
        area: ROUTES_AREA,
        data: { path: ROUTE },
        render: () => jsx(OutlookPage, {})
      },
      {
        id: 'nav',
        area: SIDEBAR_NAV_AREA,
        data: { path: ROUTE, label: 'Outlook', codicon: 'mail' }
      },
      {
        id: 'open',
        area: PALETTE_AREA,
        data: {
          id: 'hermes-outlook.open',
          label: 'Open Outlook',
          keywords: ['outlook', 'mail', 'microsoft', 'o365'],
          run: () => host.navigate(ROUTE)
        }
      }
    ])
  }
}
