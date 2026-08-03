from __future__ import annotations

try:
    from playwright.sync_api import (
        sync_playwright,
        expect,
        TimeoutError as PlaywrightTimeoutError,
        Error as PlaywrightError,
    )
except ImportError:
    sync_playwright = None
    expect = None
    PlaywrightTimeoutError = Exception
    PlaywrightError = Exception
    print("Playwright not installed. Please run: pip install playwright")

from pathlib import Path
from datetime import datetime
import getpass
import json
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
from urllib.parse import parse_qs, urlparse

REQUIRED_CONFIG_KEYS = (
    "d365_url",
    # auth_json_path is optional: defaults to auth_{username}.json next to the exe
    "journal_name",
    "browser_headless",
    "browser_slow_mo_ms",
    "page_load_timeout_ms",
    "page_load_wait_seconds",
)

PLACEHOLDER_TOKENS = (
    "your_tenant",
    "path/to/your",
    "example.com",
    "replace_me",
)

LOGIN_REDIRECT_HOSTS = (
    "login.microsoftonline.com",
    "login.live.com",
    "login.windows.net",
    "autologon.microsoftazuread-sso.com",
)

EMAIL_PATTERN = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

_exe_dir = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
DEFAULT_PLAYWRIGHT_BROWSERS_PATH = _exe_dir / "pw-browsers"
# Per-OS-user auth file so multiple Windows accounts on the same PC don't overwrite each other.
# e.g. auth_Alice.json, auth_Bob.json
DEFAULT_AUTH_JSON_PATH = _exe_dir / f"auth_{getpass.getuser()}.json"
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(DEFAULT_PLAYWRIGHT_BROWSERS_PATH))

DEFAULT_BULK_UPDATE_URL = "https://uat-sobha.docuxray.ai/api/prePost/bulkUpdateReceipt"

BULK_PASTE_BATCH_SIZE = 20

try:
    from clipboard_export import (
        build_paste_clipboard_text,
        col_defs_from_live_headers,
        load_clipboard_col_defs,
        normalize_col_defs,
        normalize_date,
        validate_live_paste_columns,
        SPACER_FIELD_KEY,
    )
except ImportError:
    build_paste_clipboard_text = None
    col_defs_from_live_headers = None
    load_clipboard_col_defs = None
    normalize_col_defs = None
    normalize_date = None
    validate_live_paste_columns = None
    SPACER_FIELD_KEY = "__spacer__"

VISUAL_ENHANCEMENT_SCRIPT = '''
window.addEventListener('DOMContentLoaded', () => {
  const style = document.createElement('style');
  style.innerHTML = `
    .human-cursor {
        position: fixed;
        z-index: 2147483647;
        pointer-events: none;
        transform: translate(-2px, -2px);
        transition: transform 0.08s ease;
        filter: drop-shadow(0 0 2px rgba(0,0,0,0.8));
    }

    @keyframes click-glow {
      0% { width: 0; height: 0; opacity: 1; }
      100% { width: 50px; height: 50px; opacity: 0; }
    }

    .click-glow-effect {
      position: fixed;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(0,0,0,0.6) 0%, rgba(0,0,0,0) 70%);
      box-shadow: 0 0 15px rgba(0, 0, 0, 0.6);
      pointer-events: none;
      z-index: 2147483646;
      transform: translate(-50%, -50%);
      animation: click-glow 0.6s ease-out forwards;
    }

    .element-highlight-rect {
        position: fixed;
        pointer-events: none;
        z-index: 2147483645;
        border: 2px solid rgba(0, 0, 0, 0.6);
        background-color: rgba(0, 0, 0, 0.08);
        transition: all 0.2s ease;
        box-sizing: border-box;
    }
  `;
  document.head.appendChild(style);

  // Black Arrow Cursor
  const cursor = document.createElement('div');
  cursor.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg"
         width="28"
         height="28"
         viewBox="0 0 24 24"
         fill="black">
      <path d="M2 2L2 22L8 16L13 22L16 19L10 13L16 13Z"/>
    </svg>
  `;
  cursor.classList.add('human-cursor');
  document.body.appendChild(cursor);

  window.__automationMouseMoveHandler = (e) => {
    cursor.style.left = e.clientX + 'px';
    cursor.style.top = e.clientY + 'px';
  };
  window.__automationMouseDownHandler = () => {
    cursor.style.transform = 'translate(-2px, -2px) scale(0.85)';
  };
  window.__automationMouseUpHandler = () => {
    cursor.style.transform = 'translate(-2px, -2px) scale(1)';
  };

  document.addEventListener('mousemove', window.__automationMouseMoveHandler);

  document.addEventListener('mousedown', window.__automationMouseDownHandler);

  document.addEventListener('mouseup', window.__automationMouseUpHandler);

  function highlightElement(target) {
     if (!target || !target.getBoundingClientRect) return;
     const rect = target.getBoundingClientRect();
     const overlay = document.createElement('div');
     overlay.classList.add('element-highlight-rect');
     overlay.style.left = rect.left + 'px';
     overlay.style.top = rect.top + 'px';
     overlay.style.width = rect.width + 'px';
     overlay.style.height = rect.height + 'px';
     document.body.appendChild(overlay);
     setTimeout(() => overlay.remove(), 800);
  }

  window.__automationClickHandler = (e) => {
    const glow = document.createElement('div');
    glow.classList.add('click-glow-effect');
    glow.style.left = e.clientX + 'px';
    glow.style.top = e.clientY + 'px';
    document.body.appendChild(glow);
    setTimeout(() => glow.remove(), 600);
    highlightElement(e.target);
  };
  document.addEventListener('click', window.__automationClickHandler, true);

  window.__automationFocusHandler = (e) => highlightElement(e.target);
  window.__automationInputHandler = (e) => highlightElement(e.target);
  document.addEventListener('focus', window.__automationFocusHandler, true);
  document.addEventListener('input', window.__automationInputHandler, true);

  window.__automationDisableVisualEnhancements = () => {
    document.querySelectorAll('.human-cursor, .click-glow-effect, .element-highlight-rect').forEach((el) => el.remove());
    if (window.__automationMouseMoveHandler) {
      document.removeEventListener('mousemove', window.__automationMouseMoveHandler);
      window.__automationMouseMoveHandler = null;
    }
    if (window.__automationMouseDownHandler) {
      document.removeEventListener('mousedown', window.__automationMouseDownHandler);
      window.__automationMouseDownHandler = null;
    }
    if (window.__automationMouseUpHandler) {
      document.removeEventListener('mouseup', window.__automationMouseUpHandler);
      window.__automationMouseUpHandler = null;
    }
    if (window.__automationClickHandler) {
      document.removeEventListener('click', window.__automationClickHandler, true);
      window.__automationClickHandler = null;
    }
    if (window.__automationFocusHandler) {
      document.removeEventListener('focus', window.__automationFocusHandler, true);
      window.__automationFocusHandler = null;
    }
    if (window.__automationInputHandler) {
      document.removeEventListener('input', window.__automationInputHandler, true);
      window.__automationInputHandler = null;
    }
    window.__automationVisualEnhancementsDisabled = true;
  };
});
'''


def _find_config_path() -> Path:
    """Find config.json in common runtime locations."""
    script_dir = Path(__file__).resolve().parent
    exe_dir = Path(sys.executable).resolve().parent
    user_config = Path.home() / ".config" / "sobha-reconciliation" / "config.json"
    system_config = Path("/etc/sobha-reconciliation/config.json")

    candidates = []
    env_path = os.environ.get("SOBHA_CONFIG_PATH")
    if env_path:
        candidates.append(Path(env_path).expanduser())
    candidates.extend([
        Path.cwd() / "config.json",
        exe_dir / "config.json",
        script_dir / "config.json",
        user_config,
        system_config,
    ])

    seen = set()
    unique_candidates = []
    for candidate in candidates:
        candidate_str = str(candidate)
        if candidate_str in seen:
            continue
        seen.add(candidate_str)
        unique_candidates.append(candidate)

    for candidate in unique_candidates:
        if candidate.exists():
            return candidate

    # No config found anywhere: create a default user config so app can start.
    user_config.parent.mkdir(parents=True, exist_ok=True)
    if not user_config.exists():
        default_config = {
            "_comment": "Auto-generated default config. Update d365_url and journal_name.",
            "d365_url": "https://<your-tenant>.sandbox.operations.dynamics.com/?cmp=COMPANY&mi=LedgerJournalTable_CustPaym",
            "auth_json_path": str(DEFAULT_AUTH_JSON_PATH),
            "journal_name": "ARBR Customers Receipt",
            "browser_headless": False,
            "browser_slow_mo_ms": 0,
            "page_load_timeout_ms": 60000,
            "page_load_wait_seconds": 1,
            "post_click_timeout_ms": 300000,
            "manual_login_button_timeout_ms": 1800000,
        }
        with open(user_config, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4)
        print(f"Created default config at {user_config}")
    return user_config


def _load_config() -> tuple[dict, Path]:
    """Load and validate config.json."""
    config_path = _find_config_path()
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    missing = [k for k in REQUIRED_CONFIG_KEYS if k not in config]
    if missing:
        raise KeyError(f"Missing required config keys: {', '.join(missing)}")

    d365_url = config.get("d365_url", "")
    if not isinstance(d365_url, str) or not d365_url.startswith("https://"):
        raise ValueError("config['d365_url'] must be a valid https URL.")

    raw_auth_path = str(config.get("auth_json_path", "")).strip()
    raw_auth_path_l = raw_auth_path.lower()
    if (
        not raw_auth_path
        or "path/to/your" in raw_auth_path_l
        or raw_auth_path_l in {"/path/to/your/auth.json", "replace_me"}
    ):
        auth_json_path = DEFAULT_AUTH_JSON_PATH
        print(f"Using default auth_json_path: {auth_json_path}")
    else:
        auth_json_path = Path(raw_auth_path).expanduser()

    if not auth_json_path.is_absolute():
        # Relative auth path is resolved from config file location, not CWD.
        auth_json_path = (config_path.parent / auth_json_path).resolve()

    # Shared config files may carry machine-specific absolute paths (e.g. /home/other-user/...).
    # If path parent is not writable for this user, fall back to a safe per-user default.
    try:
        auth_json_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError:
        auth_json_path = DEFAULT_AUTH_JSON_PATH
        auth_json_path.parent.mkdir(parents=True, exist_ok=True)
        print(f"Using default auth_json_path (non-writable configured path): {auth_json_path}")

    config["auth_json_path"] = str(auth_json_path)

    return config, config_path


CONFIG, CONFIG_PATH = _load_config()


def d365_company_from_config() -> str:
    query = parse_qs(urlparse(str(CONFIG.get("d365_url", ""))).query)
    return (query.get("cmp") or [""])[0].strip()


def update_user_runtime_config(d365_url: str | None = None, journal_name: str | None = None) -> tuple[bool, str]:
    """Update persisted user config and reload in-memory CONFIG."""
    try:
        config_path = _find_config_path()
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as err:
        return False, f"Failed to load config for update: {err}"

    if d365_url is not None:
        candidate = d365_url.strip()
        if not candidate.startswith("https://"):
            return False, "D365 URL must start with https://"
        config["d365_url"] = candidate

    if journal_name is not None:
        candidate = journal_name.strip()
        if candidate:
            config["journal_name"] = candidate

    # Always normalize to a safe per-user auth storage location.
    raw_auth_path = str(config.get("auth_json_path", "")).strip().lower()
    if not raw_auth_path or "path/to/your" in raw_auth_path or raw_auth_path == "replace_me":
        config["auth_json_path"] = str(DEFAULT_AUTH_JSON_PATH)

    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
    except Exception as err:
        return False, f"Failed to write config: {err}"

    try:
        global CONFIG, CONFIG_PATH
        CONFIG, CONFIG_PATH = _load_config()
    except Exception as err:
        return False, f"Config was written but reload failed: {err}"

    return True, f"Config updated: {config_path}"


def _get_browser_dimensions() -> tuple[int, int]:
    """Return standard 1440x900 viewport dimensions for D365 automation.

    Avoids over-zooming / distorted scaling on macOS Retina displays.
    """
    return 1440, 900


def _create_browser(playwright, *, headless: bool | None = None):
    if headless is None:
        headless = CONFIG["browser_headless"]
    launch_kwargs = {
        "headless": headless,
        "args": ["--start-maximized"],
    }
    slow_mo = int(CONFIG.get("browser_slow_mo_ms", 0) or 0)
    if slow_mo > 0:
        launch_kwargs["slow_mo"] = slow_mo
    browser = playwright.chromium.launch(**launch_kwargs)
    return browser, 0, 0


def _create_context(browser, screen_w=0, screen_h=0, use_storage_state=True):
    context_args = {"no_viewport": True}
    if use_storage_state and Path(CONFIG["auth_json_path"]).exists():
        context_args["storage_state"] = CONFIG["auth_json_path"]
    return browser.new_context(**context_args)


def _persist_storage_state(context):
    auth_path = Path(CONFIG["auth_json_path"])
    auth_path.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(auth_path))
    if not auth_path.exists() or auth_path.stat().st_size == 0:
        raise OSError(f"Auth session file was not created correctly at: {auth_path}")


def clear_saved_session() -> dict:
    """Delete persisted Playwright storage_state file."""
    auth_path = Path(str(CONFIG.get("auth_json_path", ""))).expanduser()
    if not auth_path.exists():
        return {"ok": True, "deleted": False, "path": str(auth_path)}
    try:
        auth_path.unlink()
    except OSError as err:
        return {"ok": False, "deleted": False, "path": str(auth_path), "reason": str(err)}
    return {"ok": True, "deleted": True, "path": str(auth_path)}


def get_config_issues(require_auth_state: bool = False) -> list[str]:
    """Return user-facing config validation issues."""
    issues = []

    d365_url = str(CONFIG.get("d365_url", "")).strip()
    parsed = urlparse(d365_url)
    hostname = (parsed.hostname or "").lower()
    d365_url_l = d365_url.lower()
    if not d365_url.startswith("https://"):
        issues.append("`d365_url` must start with https://")
    if any(token in d365_url_l for token in PLACEHOLDER_TOKENS) or "your_tenant" in hostname:
        issues.append(
            "`d365_url` is still a placeholder. Update it to your real D365 URL "
            "(for example: https://<tenant>.sandbox.operations.dynamics.com/...)."
        )

    auth_path = Path(str(CONFIG.get("auth_json_path", ""))).expanduser()
    auth_path_str = str(auth_path).lower()
    if any(token in auth_path_str for token in ("path/to/your", "replace_me")):
        issues.append(
            "`auth_json_path` is still a placeholder. Set a real path, "
            "for example: ~/.config/sobha-reconciliation/auth.json"
        )

    if require_auth_state and not auth_path.exists():
        issues.append(
            f"Auth state file not found at: {auth_path}\n"
            "Run Login first to create it."
        )

    return issues


def install_playwright_chromium() -> tuple[bool, str]:
    """Install Playwright Chromium browser to a persistent user path."""
    try:
        from playwright._impl._driver import compute_driver_executable
    except Exception as err:
        return False, f"Could not load Playwright driver info: {err}"

    try:
        node_exe, cli_js = compute_driver_executable()
        env = os.environ.copy()
        env.setdefault("PLAYWRIGHT_BROWSERS_PATH", str(DEFAULT_PLAYWRIGHT_BROWSERS_PATH))
        result = subprocess.run(
            [node_exe, cli_js, "install", "chromium"],
            capture_output=True,
            text=True,
            env=env,
            check=False,
        )
    except Exception as err:
        return False, f"Failed to execute Playwright install command: {err}"

    out = (result.stdout or "").strip()
    err = (result.stderr or "").strip()
    combined = "\n".join(part for part in (out, err) if part).strip()
    if result.returncode == 0:
        return True, combined or "Chromium installed successfully."
    return False, combined or f"Playwright install failed with exit code {result.returncode}."


def _find_chromium_executable_on_disk() -> Path | None:
    """Locate the Chromium binary under PLAYWRIGHT_BROWSERS_PATH without spawning the driver.

    Spawning the Playwright Node driver just to ask for the executable path costs
    over a second on startup; a filesystem glob is near-instant and covers the
    layout Playwright itself uses (chromium-<rev>/chrome-win64/chrome.exe, etc).
    """
    browsers_dir = Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH", str(DEFAULT_PLAYWRIGHT_BROWSERS_PATH)))
    if not browsers_dir.is_dir():
        return None
    candidates = (
        "chromium-*/chrome-win64/chrome.exe",
        "chromium-*/chrome-linux/chrome",
        "chromium-*/chrome-mac/Chromium.app/Contents/MacOS/Chromium",
    )
    for pattern in candidates:
        matches = sorted(browsers_dir.glob(pattern), reverse=True)
        if matches:
            return matches[0]
    return None


def is_playwright_chromium_available() -> tuple[bool, str]:
    """Check whether Playwright Chromium executable is available."""
    if not sync_playwright:
        return False, "Playwright Python package is not installed."

    fast_path = _find_chromium_executable_on_disk()
    if fast_path is not None:
        return True, str(fast_path)

    # Fall back to asking Playwright directly (slower: spawns the Node driver)
    # only when the fast filesystem check couldn't confirm an install.
    try:
        with sync_playwright() as playwright:
            exe_path = Path(playwright.chromium.executable_path)
            if exe_path.exists():
                return True, str(exe_path)
            return False, f"Chromium executable not found at: {exe_path}"
    except Exception as err:
        return False, str(err)


def _bulk_update_receipts(records, receipt_numbers) -> tuple[bool, str]:
    """Send UUID -> receipt_number mapping to prePost bulk update endpoint."""
    api_url = str(CONFIG.get("bulk_update_receipt_url", DEFAULT_BULK_UPDATE_URL)).strip()
    bearer = str(
        CONFIG.get("api_bearer_token", os.environ.get("SOBHA_API_TOKEN", ""))
    ).strip()

    if not api_url or not bearer:
        return False, "Skipped PATCH: missing api URL or bearer token."

    # Map receipts from current run to selected records by order.
    run_count = min(len(records), len(receipt_numbers))
    run_receipts = receipt_numbers[-run_count:] if run_count > 0 else []
    payload = []
    for idx in range(run_count):
        rec = records[idx]
        uuid = str(rec.get("uuid", "")).strip()
        if not uuid:
            continue
        payload.append({"uuid": uuid, "receipt_number": str(run_receipts[idx]).strip()})

    if not payload:
        return False, "Skipped PATCH: no UUID mappings available in selected records."

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        api_url,
        data=data,
        method="PATCH",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {bearer}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = resp.read().decode("utf-8", errors="replace")
        return True, body[:1000] if body else "PATCH success."
    except urllib.error.HTTPError as err:
        detail = err.read().decode("utf-8", errors="replace")
        return False, f"PATCH HTTP {err.code}: {detail[:1000]}"
    except urllib.error.URLError as err:
        return False, f"PATCH network error: {err.reason}"
    except Exception as err:
        return False, f"PATCH failed: {err}"


def _d365_date(date_text, fallback: str = "") -> str:
    """Convert any supported date input to dd/mm/yyyy for D365."""
    if normalize_date is None:
        return str(date_text or "").strip() or fallback
    normalized = normalize_date(date_text)
    if normalized:
        return normalized
    raw = str(date_text or "").strip()
    return raw or fallback


class AutomationStoppedByUser(RuntimeError):
    pass


class AutomationController:
    """Thread-safe browser controls for a single automation run."""

    def __init__(self):
        self._paused = threading.Event()
        self._quit = threading.Event()
        self._lock = threading.Lock()
        self._started_at = time.monotonic()
        self._paused_at = None
        self._paused_seconds = 0.0
        self._completed = False
        self._completed_at = None
        self._page = None
        self._bridge_attached = False

    def _status_locked(self) -> dict:
        end_time = self._completed_at or self._paused_at or time.monotonic()
        elapsed = max(0, int(end_time - self._started_at - self._paused_seconds))
        return {
            "paused": self._paused.is_set(),
            "completed": self._completed,
            "quitting": self._quit.is_set(),
            "elapsedSeconds": elapsed,
        }

    def request(self, action: str) -> dict:
        with self._lock:
            if action == "pause" and not self._completed and not self._quit.is_set() and not self._paused.is_set():
                self._paused_at = time.monotonic()
                self._paused.set()
            elif action == "resume" and self._paused.is_set() and not self._quit.is_set():
                self._paused_seconds += time.monotonic() - self._paused_at
                self._paused_at = None
                self._paused.clear()
            elif action == "quit":
                self._quit.set()
            return self._status_locked()

    def complete(self) -> None:
        with self._lock:
            if not self._completed:
                self._completed = True
                self._completed_at = time.monotonic()

    def status(self) -> dict:
        with self._lock:
            return self._status_locked()

    def checkpoint(self) -> None:
        if self._quit.is_set():
            raise AutomationStoppedByUser("Automation stopped by user.")

    @property
    def quit_requested(self) -> bool:
        return self._quit.is_set()

    @property
    def paused(self) -> bool:
        return self._paused.is_set()

    def attach_page(self, page) -> None:
        self._page = page
        if not self._bridge_attached:
            page.expose_function("automationControl", self.request)
            self._bridge_attached = True
        self.show_controls(page)

    def show_controls(self, page=None) -> None:
        page = page or self._page
        if page is None or page.is_closed():
            return
        page.evaluate(
            """
            (initialStatus) => {
                const existing = document.getElementById('automation-run-controls');
                if (existing && existing.__updateAutomationControls) {
                    existing.__updateAutomationControls(initialStatus);
                    return;
                }

                const panel = document.createElement('div');
                panel.id = 'automation-run-controls';
                panel.setAttribute('role', 'group');
                panel.setAttribute('aria-label', 'Automation controls');
                Object.assign(panel.style, {
                    position: 'fixed', right: '24px', bottom: '24px', zIndex: '2147483647',
                    display: 'flex', alignItems: 'center', gap: '8px', padding: '8px', background: 'rgba(255,255,255,0.98)',
                    border: '1px solid rgba(17,24,39,0.16)', borderRadius: '8px',
                    boxShadow: '0 12px 32px rgba(0,0,0,0.22)',
                    fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif'
                });

                const makeButton = (label, background) => {
                    const button = document.createElement('button');
                    button.type = 'button';
                    button.textContent = label;
                    Object.assign(button.style, {
                        padding: '9px 12px', border: '0', borderRadius: '6px', background,
                        color: '#fff', fontWeight: '700', fontSize: '13px', cursor: 'pointer'
                    });
                    return button;
                };

                const pauseButton = makeButton('Force Stop', '#b45309');
                const quitButton = makeButton('Quit', '#dc2626');
                const timer = document.createElement('span');
                const status = document.createElement('span');
                Object.assign(timer.style, { color: '#0f172a', fontSize: '13px', fontVariantNumeric: 'tabular-nums' });
                Object.assign(status.style, { color: '#475569', fontSize: '12px', fontWeight: '700' });
                let elapsedSeconds = 0;
                let runningSince = performance.now();
                let timerRunning = false;

                const renderTimer = () => {
                    const hours = Math.floor(elapsedSeconds / 3600);
                    const minutes = Math.floor((elapsedSeconds % 3600) / 60);
                    const seconds = elapsedSeconds % 60;
                    timer.textContent = [hours, minutes, seconds].map((part) => String(part).padStart(2, '0')).join(':');
                };
                const applyStatus = (next) => {
                    if (!next) return;
                    if (Number.isFinite(next.elapsedSeconds)) elapsedSeconds = next.elapsedSeconds;
                    runningSince = performance.now();
                    timerRunning = !next.paused && !next.completed && !next.quitting;
                    panel.dataset.paused = next.paused ? 'true' : 'false';
                    pauseButton.textContent = next.paused ? 'Resume' : 'Force Stop';
                    pauseButton.style.background = next.paused ? '#16a34a' : '#b45309';
                    pauseButton.hidden = Boolean(next.completed || next.quitting);
                    status.textContent = next.completed ? 'Automation completed' : next.quitting ? 'Quitting...' : next.paused ? 'Paused' : 'Running';
                    quitButton.disabled = Boolean(next.quitting);
                    if (next.quitting) quitButton.textContent = 'Quitting...';
                    renderTimer();
                };
                const send = (action) => {
                    if (typeof window.automationControl !== 'function') return;
                    Promise.resolve(window.automationControl(action)).then(applyStatus).catch(() => {});
                };
                pauseButton.addEventListener('click', () => {
                    const paused = panel.dataset.paused === 'true';
                    applyStatus({ paused: !paused, completed: false, quitting: false, elapsedSeconds });
                    send(paused ? 'resume' : 'pause');
                });
                quitButton.addEventListener('click', () => {
                    applyStatus({ paused: false, completed: false, quitting: true, elapsedSeconds });
                    send('quit');
                });
                panel.__updateAutomationControls = applyStatus;
                setInterval(() => {
                    if (!timerRunning) return;
                    const elapsedNow = Math.floor((performance.now() - runningSince) / 1000);
                    if (elapsedNow) {
                        elapsedSeconds += elapsedNow;
                        runningSince += elapsedNow * 1000;
                        renderTimer();
                    }
                }, 250);
                panel.append(timer, status, pauseButton, quitButton);
                document.body.appendChild(panel);
                applyStatus(initialStatus);
            }
            """,
            self.status(),
        )

    def cleanup(self) -> None:
        page = self._page
        self._page = None
        if page is None or page.is_closed():
            return
        try:
            page.evaluate("document.getElementById('automation-run-controls')?.remove()")
        except PlaywrightError:
            pass


def _automation_checkpoint(page, controller: AutomationController | None) -> None:
    if controller is None:
        return
    while True:
        controller.show_controls(page)
        controller.checkpoint()
        if not controller.paused:
            return
        # Keep Playwright's event loop active so the browser can deliver Resume/Quit.
        _interruptible_wait(page, 200, controller)


def _interruptible_wait(page, ms: int, controller: AutomationController | None = None) -> None:
    """Sleep in short chunks so Quit / Force Stop stay responsive."""
    remaining = max(0, int(ms))
    step = 200
    while remaining > 0:
        if controller is not None:
            controller.checkpoint()
        chunk = min(step, remaining)
        page.wait_for_timeout(chunk)
        remaining -= chunk


def _wait_for_automation_condition(page, expression: str, controller: AutomationController | None) -> None:
    """Poll a page condition so pause and quit work during user-controlled waits."""
    while True:
        _automation_checkpoint(page, controller)
        if page.evaluate(f"() => Boolean({expression})"):
            return
        _interruptible_wait(page, 200, controller)


class SessionExpiredError(RuntimeError):
    pass


def _auth_result(valid: bool, display_name: str | None = None, reason: str | None = None) -> dict:
    return {
        "valid": bool(valid),
        "display_name": display_name.strip() if isinstance(display_name, str) and display_name.strip() else None,
        "reason": reason.strip() if isinstance(reason, str) and reason.strip() else None,
    }


def _is_login_redirect(url: str) -> bool:
    hostname = (urlparse(str(url or "")).hostname or "").lower()
    return any(
        hostname == host or hostname.endswith(f".{host}")
        for host in LOGIN_REDIRECT_HOSTS
    )


def _is_expected_d365_url(url: str) -> bool:
    expected_host = (urlparse(str(CONFIG.get("d365_url", ""))).hostname or "").lower()
    current_host = (urlparse(str(url or "")).hostname or "").lower()
    return bool(expected_host and current_host and current_host == expected_host)


def _ensure_authenticated(page, load_label: str = "page"):
    current_url = str(page.url or "").strip()
    if _is_login_redirect(current_url):
        raise SessionExpiredError("Saved session expired. Click Login and sign in again.")
    if not _is_expected_d365_url(current_url):
        raise SessionExpiredError(
            f"Expected D365 page after loading {load_label}, but got: {current_url or 'unknown page'}"
        )


def _normalize_identity_text(raw_value: str | None) -> str | None:
    value = " ".join(str(raw_value or "").replace("\xa0", " ").split())
    if not value:
        return None

    email_match = EMAIL_PATTERN.search(value)
    if email_match:
        return email_match.group(0)

    lowered = value.lower()
    for prefix in (
        "account manager for ",
        "signed in as ",
        "signed in: ",
        "current account: ",
        "account: ",
        "user: ",
    ):
        if lowered.startswith(prefix):
            value = value[len(prefix):].strip(" :-")
            lowered = value.lower()
            break

    if not value or len(value) < 3 or len(value) > 80:
        return None

    if lowered in {
        "account manager",
        "manage account",
        "user options",
        "my account",
        "profile",
        "settings",
        "sign out",
        "sign in",
        "logout",
        "log out",
        "logged in",
    }:
        return None

    if any(token in lowered for token in ("http://", "https://", "sign out", "logout", "log out")):
        return None

    words = value.split()
    if len(words) > 6:
        return None

    return value


def _extract_signed_in_user(page) -> str | None:
    candidate_script = """
        () => {
            const seen = new Set();
            const results = [];
            const add = (value) => {
                if (typeof value !== 'string') return;
                const normalized = value.replace(/\\s+/g, ' ').trim();
                if (!normalized || seen.has(normalized)) return;
                seen.add(normalized);
                results.push(normalized);
            };
            const selectors = [
                "[data-dyn-controlname='UserOptionsButton']",
                "[data-dyn-title='User options']",
                "[data-dyn-title='Account manager']",
                "#meControl",
                "#mectrl_currentAccount_primary",
                "#mectrl_currentAccount_secondary",
                "#O365_MainLink_Me",
                "[data-testid='mectrl_main_trigger']",
                "[aria-label*='Account manager']",
                "[title*='Account manager']",
                "[aria-label*='@']",
                "[title*='@']"
            ];

            for (const selector of selectors) {
                for (const el of document.querySelectorAll(selector)) {
                    add(el.innerText);
                    add(el.textContent);
                    add(el.getAttribute("aria-label"));
                    add(el.getAttribute("title"));
                    add(el.getAttribute("data-dyn-title"));
                }
            }

            return results;
        }
    """

    for _ in range(4):
        try:
            raw_candidates = page.evaluate(candidate_script)
        except Exception:
            raw_candidates = []

        for candidate in raw_candidates or []:
            cleaned = _normalize_identity_text(candidate)
            if cleaned:
                return cleaned

        try:
            page.wait_for_timeout(750)
        except Exception:
            break

    return None


def _wait_for_d365_ready(page, load_label: str = "page"):
    _ensure_authenticated(page, load_label)
    if load_label == "saved session":
        print("Session checking 3-2-1")
        for i in range(CONFIG["page_load_wait_seconds"], 0, -1):
            print(f"Session checking {i}")
    else:
        print(f"Waiting for {load_label} load...")
        for i in range(CONFIG["page_load_wait_seconds"], 0, -1):
            print(f"Loading... {i}")
            time.sleep(1)
        print(f"{load_label.capitalize()} fully loaded!")
    try:
        page.locator("#ShellBlockingDiv").wait_for(state="hidden", timeout=60000)
    except PlaywrightTimeoutError:
        print("Overlay did not disappear in 60s; continuing.")
    _ensure_authenticated(page, load_label)


def probe_saved_session(headless: bool = True) -> dict:
    issues = get_config_issues(require_auth_state=True)
    if issues:
        return _auth_result(False, reason="; ".join(issues))

    if not sync_playwright:
        return _auth_result(False, reason="Playwright is not installed.")

    with sync_playwright() as playwright:
        browser = None
        context = None
        page = None
        try:
            browser, screen_w, screen_h = _create_browser(playwright, headless=headless)
            context = _create_context(browser, screen_w, screen_h, use_storage_state=True)
            page = context.new_page()
            page.goto(
                CONFIG["d365_url"],
                timeout=CONFIG["page_load_timeout_ms"],
                wait_until="domcontentloaded",
            )
            _wait_for_d365_ready(page, "saved session")
            return _auth_result(True, display_name=_extract_signed_in_user(page))
        except SessionExpiredError as err:
            return _auth_result(False, reason=str(err))
        except Exception as err:
            return _auth_result(False, reason=f"Unable to validate saved session: {err}")
        finally:
            if page is not None and not page.is_closed():
                page.close()
            if context is not None:
                context.close()
            if browser is not None:
                browser.close()


def _open_journal_lines(page):
    try:
        page.get_by_role("button", name=" New").first.click(timeout=15000)
    except PlaywrightError:
        page.get_by_role("button", name=" New").click(timeout=15000)
    page.wait_for_timeout(300)
    open_btn = page.locator("#JournalName_3_0_0").get_by_role("button", name="Open")
    open_btn.wait_for(state="visible", timeout=int(CONFIG.get("page_load_timeout_ms", 60000)))
    open_btn.click()
    page.get_by_role("row", name=CONFIG["journal_name"], exact=True).get_by_label("Name").click()
    page.get_by_role("button", name="Lines", exact=True).click()
    _wait_for_journal_grid_ready(page)


def _extract_voucher_values(page):
    voucher_values = []
    max_retries = 8
    for _ in range(max_retries):
        voucher_values = page.evaluate(
            """
            () => {
                let inputs = document.querySelectorAll("input[id^='LedgerJournalTrans_Voucher_'][id$='_input']");
                if (inputs.length === 0) {
                    inputs = document.querySelectorAll("input[aria-label='Voucher']");
                }
                const values = [];
                inputs.forEach(el => {
                    const val = el.value || el.getAttribute('title') || '';
                    if (val && val.trim() !== '' && val !== '0.00') {
                        values.push(val.trim());
                    }
                });
                return values;
            }
            """
        )
        if voucher_values:
            break
        page.wait_for_timeout(500)
    return voucher_values


_ACCOUNT_INPUT_SEL = 'input[id^="LedgerJournalTrans_AccountNum_"][id$="_input"]'
_ACCOUNT_CONTROL_SEL = (
    "input[id^='LedgerJournalTrans_AccountNum_'],"
    "[id^='LedgerJournalTrans_AccountNum_'],"
    "[data-dyn-controlname='LedgerJournalTrans_AccountNum'],"
    "[data-dyn-controlname='AccountNum']"
)
_JOURNAL_DATA_ROW_SEL = (
    '[role="grid"][aria-label="Journal lines"] [role="row"][id*="-row-"]'
)


def _journal_line_rows(page):
    """Journal data rows from the current D365 grid rendering."""
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    if rows.count():
        return rows
    rows = page.locator("tbody tr").filter(has=page.locator(_ACCOUNT_INPUT_SEL))
    if rows.count():
        return rows
    return page.locator("tbody tr").filter(has=page.locator(_ACCOUNT_CONTROL_SEL))


def _date_journal_line_rows(page):
    """Journal rows addressable by Date, including inactive ARIA-grid rows."""
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    if rows.count():
        return rows
    rows = page.locator("tbody tr").filter(has=page.locator('input[aria-label="Date"]'))
    if rows.count():
        return rows
    return page.locator("tbody tr").filter(has=page.locator('input[aria-label="Value date"]'))


def _resolve_journal_row_index(page, row_index: int) -> int:
    """Pick grid row index: prefer selected/current, else nth line, else last empty account row."""
    resolved = page.evaluate(
        """
        (rowIndex) => {
            const isVisible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                if (st.visibility === 'hidden' || st.display === 'none') return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };
            const accountSel = "input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']";
            const trs = [...document.querySelectorAll('tbody tr')].filter((tr) => {
                const inp = tr.querySelector(accountSel);
                return inp && isVisible(tr);
            });
            if (!trs.length) return 0;

            const selected = document.querySelector("tr[aria-selected='true'], tr[aria-current='true']");
            if (selected && isVisible(selected)) {
                const selectedIdx = trs.indexOf(selected);
                if (selectedIdx >= 0) return selectedIdx;
            }

            for (let i = trs.length - 1; i >= 0; i--) {
                const inp = trs[i].querySelector(accountSel);
                if (!inp || inp.readOnly) continue;
                if (!(inp.value || '').trim()) return i;
            }
            if (rowIndex < trs.length) return rowIndex;
            return trs.length - 1;
        }
        """,
        row_index,
    )
    return int(resolved)


def _click_pasted_journal_account_cell(page, row_index: int) -> bool:
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    row_count = rows.count()
    if row_count:
        if row_count <= row_index:
            return False
        row = rows.nth(row_index)
        account_cell = row.locator(
            '[role="gridcell"][id$="-LedgerJournalTrans_AccountNum"]'
        )
        if account_cell.count():
            cell_id = account_cell.first.get_attribute("id") or "(no id)"
            _scroll_journal_row_into_view(row)
            try:
                account_cell.first.click(timeout=5000)
            except PlaywrightError:
                account_cell.first.click(force=True, timeout=5000)
            print(
                f"Activated journal row {row_index + 1} via Account cell {cell_id}."
            )
            return True
        return False

    return page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden'
                    && rect.width > 0 && rect.height > 0;
            };
            const accountSelector = [
                "input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']",
                "[id^='LedgerJournalTrans_AccountNum_']",
                "[data-dyn-controlname='LedgerJournalTrans_AccountNum']",
                "[data-dyn-controlname='AccountNum']"
            ].join(',');

            const accountRows = [];
            for (const control of document.querySelectorAll(accountSelector)) {
                const row = control.closest('tbody tr,[role="row"]');
                if (row && !row.closest('thead') && visible(row) && !accountRows.includes(row)) {
                    accountRows.push(row);
                }
            }
            const directRow = accountRows[rowIndex];
            if (directRow) {
                const control = [...directRow.querySelectorAll(accountSelector)].find(visible)
                    || directRow.querySelector(accountSelector);
                const cell = control && (control.closest('td,[role="gridcell"]') || control);
                if (cell) {
                    cell.scrollIntoView({ block: 'center', inline: 'center' });
                    cell.click();
                    return true;
                }
            }

            const header = [...document.querySelectorAll('th,[role="columnheader"]')]
                .find((el) => visible(el) && (el.innerText || '').trim() === 'Account');
            const headerRow = header && header.closest('tr,[role="row"]');
            if (!header || !headerRow) return false;

            const columnIndex = Number.isInteger(header.cellIndex)
                ? header.cellIndex
                : [...headerRow.children].indexOf(header);
            let scope = header.parentElement;
            while (scope && !scope.querySelector('tbody tr')) scope = scope.parentElement;
            if (!scope || columnIndex < 0) return false;

            const rows = [...scope.querySelectorAll('tbody tr')].filter(visible);
            const cell = rows[rowIndex] && rows[rowIndex].children[columnIndex];
            if (!cell) return false;
            cell.scrollIntoView({ block: 'center', inline: 'center' });
            cell.click();
            return true;
        }
        """,
        row_index,
    )


def _pasted_account_field_for_row(page, row_index: int):
    rows = _journal_line_rows(page)
    if rows.count() > row_index:
        for sel in (
            f"{_ACCOUNT_INPUT_SEL}:not([readonly])",
            _ACCOUNT_INPUT_SEL,
            "input[id^='LedgerJournalTrans_AccountNum_']",
        ):
            field = rows.nth(row_index).locator(sel)
            if field.count():
                return field.first

    date_rows = page.locator("tbody tr").filter(has=page.locator('input[aria-label="Date"]'))
    if date_rows.count() > row_index:
        row = date_rows.nth(row_index)
        for sel in (
            _ACCOUNT_INPUT_SEL,
            "input[id^='LedgerJournalTrans_AccountNum_']",
            "[data-dyn-controlname='LedgerJournalTrans_AccountNum'] input",
        ):
            field = row.locator(sel)
            if field.count():
                return field.first

    selected = page.locator(
        "tbody tr[aria-selected='true'], tbody tr[aria-current='true']"
    ).locator(_ACCOUNT_INPUT_SEL)
    if selected.count():
        return selected.first

    visible = page.locator(f"{_ACCOUNT_INPUT_SEL}:visible")
    if visible.count() > row_index:
        return visible.nth(row_index)
    if visible.count():
        return visible.first
    return None


def _visible_journal_data_rows(page):
    """Visible journal data rows (not limited to rows with live Account/Date inputs)."""
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    if rows.count():
        return rows
    return page.locator("tbody tr").filter(
        has=page.locator('td, [role="gridcell"]')
    )


def _journal_line_row_at_exact(page, row_index: int):
    date_rows = _date_journal_line_rows(page)
    date_count = date_rows.count()
    if date_count > row_index:
        return date_rows.nth(row_index)

    account_rows = _journal_line_rows(page)
    account_count = account_rows.count()
    if account_count > row_index:
        return account_rows.nth(row_index)

    data_rows = _visible_journal_data_rows(page)
    data_count = data_rows.count()
    if data_count > row_index:
        return data_rows.nth(row_index)

    raise RuntimeError(
        f"Journal grid has {date_count} date line(s), {account_count} account line(s), "
        f"and {data_count} data line(s); row {row_index + 1} is not available."
    )


def _activate_journal_row_for_paste(page, row_index: int) -> None:
    """Click an inactive journal row so D365 renders Date/Account inputs, then focus Date."""
    _dismiss_d365_validation_dialog(page)
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    row_count = rows.count()
    activated = False
    if row_count:
        if row_count <= row_index:
            raise RuntimeError(
                f"Could not activate journal row {row_index + 1} "
                f"(visible data rows: {row_count})."
            )
        row = rows.nth(row_index)
        date_cell = row.locator(
            '[role="gridcell"][id$="-LedgerJournalTrans_TransDate"]'
        )
        if date_cell.count():
            cell_id = date_cell.first.get_attribute("id") or "(no id)"
            _scroll_journal_row_into_view(row)
            try:
                _click_locator_robust(
                    date_cell.first,
                    page,
                    label=f"Date cell {cell_id}",
                )
                print(f"Activated journal row {row_index + 1} via Date cell {cell_id}.")
                page.wait_for_timeout(150)
                activated = True
            except PlaywrightError as err:
                print(
                    f"Warning: Playwright could not click row {row_index + 1} Date cell "
                    f"({err}); trying JS activation."
                )

    if activated:
        return

    clicked = page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                if (st.visibility === 'hidden' || st.display === 'none') return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };

            const dateInputSel = 'input[aria-label="Date"], input[aria-label="Value date"]';
            const accountSel = "input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']";

            const grid = [...document.querySelectorAll('[role="grid"]')]
                .find((el) => (el.getAttribute('aria-label') || '') === 'Journal lines');
            let trs = grid
                ? [...grid.querySelectorAll('[role="row"][id*="-row-"]')].filter(visible)
                : [...document.querySelectorAll('tbody tr')].filter((tr) => {
                    if (!visible(tr) || tr.closest('thead')) return false;
                    return tr.querySelector(dateInputSel) || tr.querySelector(accountSel)
                        || tr.querySelector('td,[role="gridcell"]');
                }).filter((tr) => {
                    const cells = tr.querySelectorAll('td,[role="gridcell"]');
                    return cells.length >= 2 || tr.querySelector(dateInputSel)
                        || tr.querySelector(accountSel);
                });

            const result = (ok, via, reason = '') => ({
                ok,
                via,
                reason,
                count: trs.length,
                rowId: trs[rowIndex]?.id || '',
                ariaRowIndex: trs[rowIndex]?.getAttribute('aria-rowindex') || '',
            });
            const row = trs[rowIndex];
            if (!row) return result(false, '', 'row-missing');

            row.scrollIntoView({ block: 'center', inline: 'nearest' });

            const dateInput = row.querySelector('input[aria-label="Date"]:not([readonly])')
                || row.querySelector('input[aria-label="Date"]')
                || row.querySelector('input[aria-label="Value date"]');
            if (dateInput && visible(dateInput)) {
                dateInput.focus();
                dateInput.click();
                return result(true, 'date-input');
            }

            const dateCell = row.querySelector(
                '[role="gridcell"][id$="-LedgerJournalTrans_TransDate"]'
            );
            if (dateCell && visible(dateCell)) {
                dateCell.scrollIntoView({ block: 'center', inline: 'start' });
                dateCell.click();
                return result(true, 'date-cell');
            }

            // Click cell under Date header if present
            const headers = [...document.querySelectorAll('th,[role="columnheader"]')];
            const dateHeader = headers.find((h) => {
                const t = (h.innerText || '').trim().toLowerCase();
                return t === 'date' || t === 'value date';
            });
            if (dateHeader) {
                const headerRow = dateHeader.closest('tr,[role="row"]');
                const colIndex = headerRow
                    ? [...headerRow.children].indexOf(dateHeader)
                    : -1;
                if (colIndex >= 0 && row.children[colIndex]) {
                    const cell = row.children[colIndex];
                    cell.scrollIntoView({ block: 'center', inline: 'start' });
                    cell.click();
                    return result(true, 'date-header-cell');
                }
            }

            const firstCell = row.querySelector('td,[role="gridcell"]');
            if (firstCell) {
                firstCell.scrollIntoView({ block: 'center', inline: 'start' });
                firstCell.click();
                return result(true, 'first-cell');
            }

            row.click();
            return result(true, 'row-click');
        }
        """,
        row_index,
    )
    if isinstance(clicked, dict):
        if clicked.get("rowId"):
            print(
                f"Requested journal row {row_index + 1} matched D365 row "
                f"{clicked.get('rowId')} (aria-rowindex={clicked.get('ariaRowIndex') or '-'})."
            )
    if not isinstance(clicked, dict) or not clicked.get("ok"):
        count = clicked.get("count", 0) if isinstance(clicked, dict) else 0
        raise RuntimeError(
            f"Could not activate journal row {row_index + 1} "
            f"(visible data rows: {count})."
        )
    print(
        f"Activated journal row {row_index + 1} via {clicked.get('via')} "
        f"({clicked.get('count')} visible rows)."
    )
    page.wait_for_timeout(250)
    try:
        _wait_for_journal_grid_ready(page, timeout_ms=8000)
    except PlaywrightTimeoutError:
        pass

    # Prefer focusing Date after activation
    try:
        _focus_journal_row_at(page, row_index)
    except RuntimeError:
        # Row may still lack Date input; try global Date focus on selected row
        try:
            page.locator('input[aria-label="Date"]:not([readonly])').first.click(timeout=3000)
            page.wait_for_timeout(100)
            page.keyboard.press("Home")
        except (PlaywrightError, PlaywrightTimeoutError):
            pass


def _journal_field_at_row(page, row_index: int, *, aria_label: str | None = None, css: str | None = None):
    row = _journal_line_row_at_exact(page, row_index)
    _scroll_journal_row_into_view(row)
    if css:
        base_sel = css
        editable_sel = f"{css}:not([readonly])"
    else:
        base_sel = f'input[aria-label="{aria_label}"]'
        editable_sel = f'{base_sel}:not([readonly])'
    for sel in (editable_sel, base_sel):
        field = row.locator(sel)
        if field.count():
            return field.first
    raise PlaywrightError(f"No journal field found for row {row_index + 1}.")


def _focus_journal_row_at(page, row_index: int, aria_label: str = "Date") -> None:
    row = _journal_line_row_at_exact(page, row_index)
    _scroll_journal_row_into_view(row)
    if page.locator(_JOURNAL_DATA_ROW_SEL).count():
        field_input = row.locator(f'input[aria-label="{aria_label}"]:not([readonly])').first
        try:
            field_input.wait_for(state="visible", timeout=4000)
        except PlaywrightTimeoutError:
            # Column likely scrolled out of view / virtualized away; reset horizontal
            # scroll to the Date column and give it one more chance before giving up.
            _scroll_journal_grid_to_start(page)
            row = _journal_line_row_at_exact(page, row_index)
            _scroll_journal_row_into_view(row)
            field_input = row.locator(f'input[aria-label="{aria_label}"]:not([readonly])').first
            field_input.wait_for(state="visible", timeout=6000)
        _click_locator_robust(field_input, page, label=f"{aria_label} input row {row_index + 1}")
        page.keyboard.press("Home")
        return

    for sel in (f'input[aria-label="{aria_label}"]:not([readonly])', f'input[aria-label="{aria_label}"]'):
        field_input = row.locator(sel)
        if field_input.count():
            target = field_input.first
            try:
                target.scroll_into_view_if_needed(timeout=5000)
            except PlaywrightError:
                pass
            _click_locator_robust(target, page, label=f"{aria_label} input row {row_index + 1}")
            page.wait_for_timeout(150)
            try:
                page.keyboard.press("Home")
            except PlaywrightError:
                pass
            return
    raise RuntimeError(f"Could not find {aria_label} field for pasted row {row_index + 1}.")


def _journal_grid_row(page, row_index: int):
    """Locator for the journal line row to fill (by index, selection, or empty new line)."""
    rows = _journal_line_rows(page)
    count = rows.count()
    if count == 0:
        return page.locator("tbody tr").first

    pick = _resolve_journal_row_index(page, row_index)
    if pick < count:
        return rows.nth(pick)
    return rows.nth(count - 1)


def _scroll_journal_grid_to_start(page) -> None:
    """Reset the journal grid's horizontal scroll so the Date column (leftmost) is visible.

    D365's journal grid virtualizes/hides off-screen columns, so after interacting with a
    column further to the right (e.g. Reference date), the Date column's input can become
    zero-size / non-visible even though it still exists in the DOM. Scrolling any horizontally
    scrollable ancestor back to the start avoids Playwright visibility timeouts on Date.
    """
    try:
        page.evaluate(
            """
            () => {
                const grid = document.querySelector('[role="grid"][aria-label="Journal lines"]');
                if (!grid) return;
                const candidates = [grid, ...grid.querySelectorAll('*')];
                let parent = grid.parentElement;
                while (parent) {
                    candidates.push(parent);
                    parent = parent.parentElement;
                }
                for (const el of candidates) {
                    const st = window.getComputedStyle(el);
                    const overflowX = st.overflowX;
                    if (
                        (overflowX === 'auto' || overflowX === 'scroll')
                        && el.scrollWidth > el.clientWidth + 4
                        && el.scrollLeft > 0
                    ) {
                        el.scrollLeft = 0;
                    }
                }
            }
            """
        )
    except PlaywrightError:
        pass


def _scroll_journal_row_into_view(row) -> None:
    page = row.page
    try:
        page.evaluate(
            """
            () => {
                const grid = document.querySelector('[role="grid"][aria-label="Journal lines"]');
                if (grid) {
                    grid.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                }
            }
            """
        )
    except PlaywrightError:
        pass
    try:
        row.evaluate(
            """
            (el) => {
                el.scrollIntoView({ block: 'center', inline: 'center' });
                let parent = el.parentElement;
                while (parent) {
                    const st = window.getComputedStyle(parent);
                    const overflowY = st.overflowY;
                    if (
                        (overflowY === 'auto' || overflowY === 'scroll')
                        && parent.scrollHeight > parent.clientHeight + 4
                    ) {
                        const rect = el.getBoundingClientRect();
                        const parentRect = parent.getBoundingClientRect();
                        if (rect.top < parentRect.top) {
                            parent.scrollTop -= (parentRect.top - rect.top + 24);
                        } else if (rect.bottom > parentRect.bottom) {
                            parent.scrollTop += (rect.bottom - parentRect.bottom + 24);
                        }
                    }
                    parent = parent.parentElement;
                }
            }
            """
        )
        page.wait_for_timeout(150)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    _scroll_journal_grid_to_start(page)


def _click_locator_robust(locator, page, *, label: str = "element", timeout: int = 5000) -> None:
    """Click a grid cell even when D365 fixed-table layout leaves it off-screen."""
    last_err: PlaywrightError | None = None
    try:
        locator.scroll_into_view_if_needed(timeout=timeout)
    except PlaywrightError as err:
        last_err = err
    page.wait_for_timeout(100)

    for strategy in ("click", "force", "js"):
        try:
            if strategy == "click":
                locator.click(timeout=timeout)
                return
            if strategy == "force":
                locator.click(force=True, timeout=timeout)
                return
            locator.evaluate(
                """
                (el) => {
                    el.scrollIntoView({ block: 'center', inline: 'center' });
                    const fire = (type) => el.dispatchEvent(
                        new MouseEvent(type, { bubbles: true, cancelable: true, view: window })
                    );
                    fire('mousedown');
                    fire('mouseup');
                    fire('click');
                    if (typeof el.focus === 'function') {
                        el.focus();
                    }
                }
                """
            )
            return
        except PlaywrightError as err:
            last_err = err

    if last_err is not None:
        raise last_err
    raise PlaywrightError(f"Could not click {label}.")


def _journal_field_locator(page, row_index: int, *, aria_label: str | None = None, css: str | None = None):
    """Resolve an editable input scoped to the correct journal grid row."""
    if css:
        base_sel = css
        editable_sel = f"{css}:not([readonly])"
    else:
        base_sel = f'input[aria-label="{aria_label}"]'
        editable_sel = f'{base_sel}:not([readonly])'

    row = _journal_grid_row(page, row_index)
    _scroll_journal_row_into_view(row)
    for sel in (editable_sel, base_sel):
        field = row.locator(sel)
        if field.count() > 0:
            return field.first

    fields = page.locator(editable_sel)
    if fields.count() == 0:
        fields = page.locator(base_sel)
    count = fields.count()
    if count == 0:
        raise PlaywrightError(f"No journal field found for row {row_index + 1}.")
    pick = min(row_index, count - 1)
    return fields.nth(pick)


def _wait_for_journal_grid_ready(page, timeout_ms: int | None = None) -> None:
    """Wait until journal line grid inputs are rendered after opening Lines."""
    if timeout_ms is None:
        timeout_ms = int(CONFIG.get("page_load_timeout_ms", 60000))
    page.locator("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']").first.wait_for(
        state="visible",
        timeout=timeout_ms,
    )
    try:
        page.locator("#ShellBlockingDiv").wait_for(state="hidden", timeout=timeout_ms)
    except PlaywrightTimeoutError:
        pass
    page.wait_for_timeout(300)


def _focus_journal_row_for_paste(page, row_index: int = 0, anchor_label: str = "Date") -> None:
    """Focus the leftmost grid cell (Date column by default) so paste aligns with D365 columns."""
    row = _journal_grid_row(page, row_index)
    _scroll_journal_row_into_view(row)
    resolved_index = _resolve_journal_row_index(page, row_index)

    clicked = page.evaluate(
        """
        (args) => {
            const { rowIndex, anchorLabel } = args;
            const isVisible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                if (st.visibility === 'hidden' || st.display === 'none') return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };
            const accountSel = "input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']";
            const trs = [...document.querySelectorAll('tbody tr')].filter((tr) => {
                const inp = tr.querySelector(accountSel);
                return inp && isVisible(tr);
            });
            const row = trs[rowIndex] || trs[0];
            if (!row) return false;

            const escaped = anchorLabel.replace(/"/g, '\\\\"');
            const anchorInput = row.querySelector(`input[aria-label="${escaped}"]:not([readonly])`)
                || row.querySelector(`input[aria-label="${escaped}"]`);
            if (anchorInput && isVisible(anchorInput)) {
                anchorInput.scrollIntoView({ block: 'center', inline: 'start' });
                anchorInput.focus();
                anchorInput.click();
                return true;
            }

            const inputs = [...row.querySelectorAll('input:not([readonly])')]
                .filter(isVisible)
                .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
            if (!inputs.length) return false;
            inputs[0].scrollIntoView({ block: 'center', inline: 'start' });
            inputs[0].focus();
            inputs[0].click();
            return true;
        }
        """,
        {"rowIndex": resolved_index, "anchorLabel": anchor_label},
    )

    if not clicked:
        for aria_label in (anchor_label, "Date", "Value date"):
            try:
                field = _journal_field_locator(page, row_index, aria_label=aria_label)
                field.click(timeout=5000)
                clicked = True
                break
            except (PlaywrightError, PlaywrightTimeoutError):
                continue

    page.wait_for_timeout(150)
    try:
        page.keyboard.press("Home")
        page.wait_for_timeout(100)
    except PlaywrightError:
        pass


def _focus_new_journal_line(page, row_index: int) -> None:
    """Activate the journal line for row_index so editable inputs are targeted."""
    row = _journal_grid_row(page, row_index)
    _scroll_journal_row_into_view(row)
    account = row.locator(f"{_ACCOUNT_INPUT_SEL}:not([readonly])")
    if account.count() == 0:
        account = row.locator(_ACCOUNT_INPUT_SEL)
    account = account.first
    try:
        account.scroll_into_view_if_needed(timeout=5000)
        try:
            account.click(timeout=5000)
        except PlaywrightError:
            account.click(force=True, timeout=5000)
        page.wait_for_timeout(100)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass


def _fill_text_field(locator, value: str) -> None:
    """Clear and fill a D365 field, then Tab to commit lookup/combobox values."""
    try:
        locator.evaluate(
            "el => { el.scrollIntoView({ block: 'center', inline: 'nearest' }); el.focus(); }"
        )
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    try:
        locator.scroll_into_view_if_needed(timeout=5000)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    try:
        locator.click(timeout=10000)
    except PlaywrightError:
        locator.click(force=True, timeout=10000)
    locator.press("ControlOrMeta+a")
    locator.press("Backspace")
    locator.fill(str(value), timeout=10000)
    locator.press("Tab")


def _fill_text_field_with_retry(locator, value: str) -> None:
    """Fill field and retry once if input_value does not match expected."""
    expected = str(value).strip()
    _fill_text_field(locator, expected)
    try:
        current = (locator.input_value() or "").strip()
        if current and current.casefold() != expected.casefold():
            _fill_text_field(locator, expected)
    except PlaywrightError:
        pass


def _dismiss_d365_validation_dialog(page) -> None:
    """Close D365 'information not valid' modal if it blocks New."""
    for label in ("Close", "OK"):
        try:
            button = page.get_by_role("button", name=label)
            if button.count() == 0:
                continue
            button.first.click(timeout=2000)
            page.wait_for_timeout(200)
            return
        except (PlaywrightTimeoutError, PlaywrightError):
            continue


def _read_d365_validation_issue(page) -> str | None:
    """Return a short D365 validation/blocker message if one is visible, else None."""
    text = page.evaluate(
        """
        () => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                if (st.display === 'none' || st.visibility === 'hidden') return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };

            const blocker = /already used|cannot be saved until you fix|validation issue|payment reference id|information is not valid|fix a validation|fix it/i;
            const linePatterns = [
                /payment reference id .+ already used in the journal[^\\n]*/i,
                /cannot be saved until you fix a validation issue[^\\n]*/i,
                /already used in the journal[^\\n]*/i,
            ];

            const pickText = (raw) => {
                const cleaned = String(raw || '').replace(/\\s+/g, ' ').trim();
                if (!cleaned || cleaned.length < 8) return null;
                for (const pat of linePatterns) {
                    const m = cleaned.match(pat);
                    if (m && m[0]) return m[0].trim();
                }
                if (blocker.test(cleaned)) return cleaned;
                return null;
            };

            const seen = new Set();
            const hits = [];

            const tryPush = (raw) => {
                const hit = pickText(raw);
                if (!hit || seen.has(hit)) return;
                seen.add(hit);
                hits.push(hit);
            };

            const selectors = [
                '[role="alert"]',
                '[role="status"]',
                '[role="dialog"]',
                '.messageBar-message',
                '.messageBar',
                '[class*="messageBar"]',
                '[class*="MessageBar"]',
                '[class*="notification"]',
                '[class*="Notification"]',
                '[class*="infolog"]',
                '[class*="InfoLog"]',
                '[data-dyn-controlname*="Message"]',
                '[id*="MessageBar"]',
                '[id*="messageBar"]',
                '[id*="Infolog"]',
                '[id*="infolog"]',
            ];
            for (const sel of selectors) {
                for (const el of document.querySelectorAll(sel)) {
                    if (!visible(el)) continue;
                    tryPush(el.innerText || el.textContent);
                }
            }

            // Flyout / floating validation panel text
            for (const el of document.querySelectorAll('div,span,p,a')) {
                if (!visible(el)) continue;
                const raw = (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
                if (!raw || raw.length < 20 || raw.length > 500) continue;
                if (!blocker.test(raw)) continue;
                if ((el.children || []).length > 8) continue;
                tryPush(raw);
                if (hits.length >= 5) break;
            }

            if (!hits.length) {
                const bodyText = document.body ? (document.body.innerText || '') : '';
                for (const line of bodyText.split('\\n')) {
                    tryPush(line);
                    if (hits.length) break;
                }
            }

            if (!hits.length) {
                const rowErrors = [...document.querySelectorAll('tbody tr')].filter((tr) => {
                    if (!visible(tr)) return false;
                    return tr.querySelector(
                        '[class*="warning"], [class*="error"], [class*="invalid"], '
                        + '[aria-invalid="true"], [data-dyn-invalid="true"], '
                        + 'img[alt*="warning" i], img[alt*="error" i], '
                        + 'span[title*="error" i], span[title*="warning" i]'
                    );
                });
                if (rowErrors.length) {
                    return 'Validation error on journal row — fix in D365 before continuing';
                }
            }

            return hits[0] || null;
        }
        """
    )
    if not text:
        return None
    cleaned = " ".join(str(text).split())
    if len(cleaned) > 160:
        cleaned = cleaned[:157].rstrip() + "..."
    return cleaned or None


def _wait_until_issue_resolved(page, issue_name: str, controller: AutomationController | None = None) -> None:
    """Block until user clicks Yes on 'is it resolved?'. No re-shows the gate."""
    label = (issue_name or "").strip() or "Validation issue"
    print(f"Waiting for user to resolve issue: {label}")
    while True:
        page.evaluate(
            """
            (issueName) => {
                window.automationIssueResolved = null;
                const old = document.getElementById('automation-issue-resolved-gate');
                if (old) old.remove();

                const wrap = document.createElement('div');
                wrap.id = 'automation-issue-resolved-gate';
                wrap.setAttribute('role', 'dialog');
                Object.assign(wrap.style, {
                    position: 'fixed',
                    right: '24px',
                    bottom: '86px',
                    zIndex: '2147483647',
                    width: 'min(420px, calc(100vw - 24px))',
                    boxSizing: 'border-box',
                    padding: '14px',
                    background: 'rgba(255,255,255,0.98)',
                    border: '1px solid rgba(17,24,39,0.16)',
                    borderRadius: '14px',
                    boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
                    fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                    color: '#0f172a'
                });

                const title = document.createElement('div');
                title.textContent = issueName;
                Object.assign(title.style, {
                    fontSize: '13px',
                    fontWeight: '700',
                    lineHeight: '1.35',
                    wordBreak: 'break-word'
                });
                wrap.appendChild(title);

                const q = document.createElement('div');
                q.textContent = 'Is it resolved?';
                Object.assign(q.style, {
                    marginTop: '8px',
                    fontSize: '13px',
                    fontWeight: '600',
                    lineHeight: '1.35'
                });
                wrap.appendChild(q);

                const actions = document.createElement('div');
                Object.assign(actions.style, {
                    display: 'flex',
                    justifyContent: 'flex-end',
                    gap: '10px',
                    marginTop: '12px'
                });

                const noBtn = document.createElement('button');
                noBtn.type = 'button';
                noBtn.textContent = 'No';
                Object.assign(noBtn.style, {
                    padding: '9px 14px',
                    borderRadius: '12px',
                    border: '1px solid rgba(17,24,39,0.12)',
                    background: '#f8fafc',
                    color: '#0f172a',
                    fontWeight: '800',
                    fontSize: '13px',
                    cursor: 'pointer'
                });
                noBtn.onclick = () => {
                    window.automationIssueResolved = 'no';
                    wrap.remove();
                };

                const yesBtn = document.createElement('button');
                yesBtn.type = 'button';
                yesBtn.textContent = 'Yes';
                Object.assign(yesBtn.style, {
                    padding: '9px 14px',
                    borderRadius: '12px',
                    border: '0',
                    background: '#16a34a',
                    color: '#fff',
                    fontWeight: '800',
                    fontSize: '13px',
                    cursor: 'pointer'
                });
                yesBtn.onclick = () => {
                    window.automationIssueResolved = 'yes';
                    wrap.remove();
                };

                actions.appendChild(noBtn);
                actions.appendChild(yesBtn);
                wrap.appendChild(actions);
                document.body.appendChild(wrap);
            }
            """,
            label,
        )
        _wait_for_automation_condition(page, "window.automationIssueResolved !== null", controller)
        decision = page.evaluate("window.automationIssueResolved")
        if decision == "yes":
            print(f"User confirmed issue resolved: {label}")
            _dismiss_d365_validation_dialog(page)
            page.wait_for_timeout(400)
            return
        print(f"User clicked No — still waiting for resolution: {label}")
        page.wait_for_timeout(400)
        refreshed = _read_d365_validation_issue(page)
        if refreshed:
            label = refreshed


def _gate_if_validation_issue(page, controller: AutomationController | None = None) -> None:
    """If a D365 validation blocker is visible, wait until the user confirms Yes."""
    while True:
        issue = _read_d365_validation_issue(page)
        if not issue:
            return
        print(f"D365 validation issue detected — showing resolution gate: {issue}")
        _wait_until_issue_resolved(page, issue, controller)
        _dismiss_d365_validation_dialog(page)
        page.wait_for_timeout(400)
        if not _read_d365_validation_issue(page):
            return
        print("Validation still visible after Yes — fix remaining issues in D365, then click Yes again.")


def _prepare_journal_grid_for_interaction(page) -> None:
    """Dismiss blockers and scroll the journal grid back into a clickable state."""
    _dismiss_d365_validation_dialog(page)
    try:
        page.evaluate(
            """
            () => {
                document.getElementById('automation-issue-resolved-gate')?.remove();
                const grid = document.querySelector('[role="grid"][aria-label="Journal lines"]');
                if (grid) {
                    grid.scrollIntoView({ block: 'center', inline: 'nearest' });
                }
            }
            """
        )
    except PlaywrightError:
        pass
    _scroll_journal_grid_to_start(page)
    try:
        page.locator("#ShellBlockingDiv").wait_for(state="hidden", timeout=3000)
    except PlaywrightTimeoutError:
        pass
    _wait_for_journal_grid_idle(page)


def _repaste_and_save_row_with_retry(
    page,
    record,
    row_index: int,
    col_defs,
    controller: AutomationController | None = None,
    *,
    use_live_order: bool = False,
    save_label: str | None = None,
) -> None:
    """Paste one journal row and save; if D365 raises a validation issue, wait for the user
    to fix it directly in the grid and just re-save — never re-paste over a manual fix."""
    label = save_label or f"after complete row {row_index + 1} re-paste"
    attempt = 0
    pasted = False
    value_date_filled = False
    grid_reset_attempts = 0
    max_grid_reset_attempts = 6
    while True:
        attempt += 1
        _automation_checkpoint(page, controller)
        _gate_if_validation_issue(page, controller)
        if attempt > 1:
            print(f"Row {row_index + 1}: re-saving after user fix (attempt {attempt})...")
        try:
            _prepare_journal_grid_for_interaction(page)
            if not pasted:
                _repaste_pasted_row(
                    page,
                    record,
                    row_index,
                    col_defs,
                    use_live_order=use_live_order,
                )
                pasted = True
                _dismiss_d365_validation_dialog(page)
            _save_journal_grid(page, label, settle_ms=100)
            _dismiss_d365_validation_dialog(page)
            if not value_date_filled and _value_date_needs_fill(page, record, row_index):
                try:
                    _activate_journal_row_for_paste(page, row_index)
                    _confirm_unsaved_changes_dialog(page, wait_ms=500)
                except (PlaywrightError, PlaywrightTimeoutError, RuntimeError):
                    pass
                _fill_value_date_at_row(page, record, row_index, force=True)
                value_date_filled = True
                _save_journal_grid(page, f"after row {row_index + 1} Value date", settle_ms=100)
                _dismiss_d365_validation_dialog(page)
        except AutomationStoppedByUser:
            raise
        except (PlaywrightError, PlaywrightTimeoutError) as err:
            msg = str(err).casefold()
            transient = (
                "outside of the viewport" in msg
                or "not visible" in msg
                or "to be visible" in msg
                or "intercept" in msg
                or "timeout" in msg
                or "detached" in msg
            )
            grid_reset_attempts += 1
            if transient and grid_reset_attempts <= max_grid_reset_attempts:
                print(
                    f"Row {row_index + 1}: grid interaction failed ({err}); "
                    f"resetting grid view and retrying (reset {grid_reset_attempts}/{max_grid_reset_attempts})..."
                )
                _automation_checkpoint(page, controller)
                _interruptible_wait(page, 600, controller)
                _prepare_journal_grid_for_interaction(page)
                continue
            raise
        grid_reset_attempts = 0
        issue = _read_d365_validation_issue(page)
        if not issue:
            print(f"Row {row_index + 1}: saved successfully.")
            return
        print(
            f"Row {row_index + 1}: validation still present after save — "
            f"waiting for user to fix: {issue}"
        )
        _wait_until_issue_resolved(page, issue, controller)


def _finalize_journal_line_before_new(
    page,
    row_index: int,
    ref_date_loc,
    ref_date: str,
    pay_method: str,
) -> None:
    """Re-apply reference date and method of payment so D365 accepts the line before New."""
    expected_ref = str(ref_date).strip()
    try:
        current_ref = (ref_date_loc.input_value() or "").strip()
        if not current_ref or current_ref != expected_ref:
            _fill_text_field_with_retry(ref_date_loc, expected_ref)
            print(f"Finalized reference date: {expected_ref}")
    except Exception as err:
        print(f"Warning: Could not finalize reference date: {err}")

    try:
        paym_mode_input = _journal_field_locator(
            page,
            row_index,
            css='input[aria-label="Method of payment"]:not([id^="Sel_"])',
        )
        _fill_text_field_with_retry(paym_mode_input, pay_method)
        print(f"Re-filled method of payment at end: {pay_method}")
    except Exception as err:
        print(f"Warning: Could not finalize method of payment: {err}")

    page.wait_for_timeout(250)
    _dismiss_d365_validation_dialog(page)


def _wait_for_post_click(page, *, show_notice=True, controller: AutomationController | None = None):
    page.evaluate(
        """
        ({ showNotice }) => {
            window.postClicked = false;
            const oldNotice = document.getElementById('automation-post-notice');
            if (oldNotice) oldNotice.remove();

            if (showNotice) {
                const notice = document.createElement('div');
                notice.id = 'automation-post-notice';
                notice.setAttribute('role', 'status');
                notice.setAttribute('aria-live', 'polite');
                notice.innerHTML = `
                  <div style="display:flex; align-items:flex-start; gap:10px;">
                    <div style="
                      width:28px; height:28px;
                      display:grid; place-items:center;
                      border-radius:8px;
                      background: rgba(255,255,255,0.18);
                      flex: 0 0 auto;
                      font-size: 16px;
                      line-height: 1;
                    ">i</div>

                    <div style="min-width:0;">
                      <div style="font-weight:800; font-size:14px; line-height:1.2;">
                        Changes saved
                      </div>
                      <div style="font-weight:600; font-size:13px; line-height:1.35; opacity:0.95; margin-top:4px;">
                        Please review all fields, then click <span style="font-weight:900;">Post</span>.
                      </div>
                    </div>

                    <button id="automation-post-close" type="button" aria-label="Dismiss" title="Dismiss" style="
                      margin-left:auto;
                      width:30px; height:30px;
                      display:grid; place-items:center;
                      border:0;
                      border-radius:10px;
                      background: rgba(255,255,255,0.16);
                      color:#fff;
                      cursor:pointer;
                      font-size:16px;
                      line-height:1;
                      flex: 0 0 auto;
                      transition: transform .12s ease, background .12s ease;
                    ">x</button>
                  </div>
                `;

                Object.assign(notice.style, {
                    position: 'fixed',
                    top: '150px',
                    left: '50%',
                    transform: 'translateX(-50%)',
                    zIndex: '2147483647',
                    width: 'min(720px, calc(100vw - 24px))',
                    boxSizing: 'border-box',
                    padding: '12px 14px',
                    background: 'rgba(11, 95, 255, 0.92)',
                    color: '#fff',
                    borderRadius: '14px',
                    border: '1px solid rgba(255,255,255,0.18)',
                    fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                    boxShadow: '0 14px 40px rgba(0,0,0,0.28)',
                    backdropFilter: 'blur(8px)',
                    WebkitBackdropFilter: 'blur(8px)'
                });

                const closeBtn = notice.querySelector('#automation-post-close');
                if (closeBtn) {
                  closeBtn.addEventListener('click', () => notice.remove());
                }

                document.body.appendChild(notice);
            }

            if (window.__postHandler) {
                document.removeEventListener('click', window.__postHandler, true);
            }

            window.__postHandler = function postHandler(e) {
                const el = e.target;
                const txt = (el && (el.innerText || el.textContent)) ? (el.innerText || el.textContent) : '';
                if (txt && txt.trim() === 'Post') {
                    window.postClicked = true;
                    const n = document.getElementById('automation-post-notice');
                    if (n) n.remove();
                }

                const btn = el && el.closest ? el.closest('button') : null;
                if (btn && btn.innerText && btn.innerText.trim() === 'Post') {
                    window.postClicked = true;
                    const n = document.getElementById('automation-post-notice');
                    if (n) n.remove();
                }
            };

            document.addEventListener('click', window.__postHandler, true);
        }
        """,
        {"showNotice": show_notice},
    )
    _wait_for_automation_condition(page, "window.postClicked === true", controller)


def _wait_for_post_confirmation(page, controller: AutomationController | None = None) -> str:
    page.evaluate(
        """
        () => {
          window.automationPostDecision = null;
          const oldGate = document.getElementById('automation-post-continue-gate');
          if (oldGate) oldGate.remove();

          const wrap = document.createElement('div');
          wrap.id = 'automation-post-continue-gate';
          wrap.setAttribute('role', 'dialog');
          wrap.setAttribute('aria-modal', 'false');

          Object.assign(wrap.style, {
            position: 'fixed',
            top: '150px',
            left: '50%',
            transform: 'translateX(-50%)',
            zIndex: '2147483647',
            width: 'min(720px, calc(100vw - 24px))',
            boxSizing: 'border-box',
            padding: '14px',
            background: 'rgba(255,255,255,0.96)',
            border: '1px solid rgba(17,24,39,0.12)',
            borderRadius: '14px',
            boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
            backdropFilter: 'blur(6px)',
            WebkitBackdropFilter: 'blur(6px)',
            fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
            color: '#0f172a'
          });

          wrap.innerHTML = `
            <div style="display:flex; align-items:flex-start; gap:12px;">
              <div style="
                width:34px; height:34px; display:grid; place-items:center;
                border-radius:10px; background: rgba(22,163,74,0.10);
                flex:0 0 auto; font-size:18px; line-height:1;
              ">OK</div>
              <div style="min-width:0;">
                <div style="font-weight:800; font-size:14px; line-height:1.2; margin-top:2px;">
                  Confirm posting is successful
                </div>
                <div style="margin-top:6px; font-size:13px; line-height:1.4; color: rgba(15, 23, 42, 0.88);">
                  Please ensure posting is successful. Click Continue only after successful posting.
                </div>
              </div>
            </div>
          `;

          const actions = document.createElement('div');
          Object.assign(actions.style, {
            display: 'flex',
            justifyContent: 'flex-end',
            gap: '10px',
            marginTop: '12px'
          });

          const waitButton = document.createElement('button');
          waitButton.type = 'button';
          waitButton.textContent = 'Wait';
          Object.assign(waitButton.style, {
            padding: '9px 14px',
            borderRadius: '12px',
            border: '1px solid rgba(17,24,39,0.12)',
            background: '#f8fafc',
            color: '#0f172a',
            fontWeight: '700',
            fontSize: '13px',
            cursor: 'pointer'
          });
          waitButton.addEventListener('click', () => {
            window.automationPostDecision = 'wait';
            wrap.remove();
          });

          const button = document.createElement('button');
          button.type = 'button';
          button.textContent = 'Continue';
          Object.assign(button.style, {
            padding: '9px 14px',
            borderRadius: '12px',
            border: '0',
            background: '#16a34a',
            color: '#fff',
            fontWeight: '800',
            fontSize: '13px',
            cursor: 'pointer',
            boxShadow: '0 10px 26px rgba(22,163,74,0.28)'
          });
          button.addEventListener('click', () => {
            window.automationPostDecision = 'continue';
            wrap.remove();
          });

          actions.appendChild(waitButton);
          actions.appendChild(button);
          wrap.appendChild(actions);
          document.body.appendChild(wrap);
        }
        """
    )
    _wait_for_automation_condition(page, "window.automationPostDecision !== null", controller)
    return page.evaluate("window.automationPostDecision")


def _wait_for_post_and_confirmation(page, *, bulk=False, controller: AutomationController | None = None):
    if bulk:
        _wait_for_bulk_post_click(page, controller=controller)
    else:
        _wait_for_post_click(page, controller=controller)

    while True:
        print("User clicked Post. Waiting for continue confirmation...")
        decision = _wait_for_post_confirmation(page, controller)
        if decision == "continue":
            break
        print("User clicked Wait. Waiting for Post again...")
        if bulk:
            _wait_for_bulk_post_click(page, controller=controller)
        else:
            _wait_for_post_click(page, show_notice=False, controller=controller)


def _wait_for_batch_action(page, *, is_last_sub_batch: bool, current_index: int, total_sub_batches: int, controller: AutomationController | None = None) -> str:
    button_text = "Close Window" if is_last_sub_batch else "Next Batch"
    action_value = "close" if is_last_sub_batch else "next"
    heading = (
        "All sub-batches completed."
        if is_last_sub_batch
        else f"Sub-batch {current_index} of {total_sub_batches} completed. Please Click Ctrl+R to refresh the page before continuing to the next sub-batch."
    )
    body = (
        "Click Close Window to finish."
        if is_last_sub_batch
        else f"Click {button_text} to refresh and continue with sub-batch {current_index + 1}."
    )
    page.evaluate(
        """
        ({buttonText, actionValue, heading, body}) => {
            window.automationBatchAction = null;
            const oldWrap = document.getElementById('automation-batch-action');
            if (oldWrap) oldWrap.remove();

            const wrap = document.createElement('div');
            wrap.id = 'automation-batch-action';
            Object.assign(wrap.style, {
                position: 'fixed',
                top: '150px',
                left: '50%',
                transform: 'translateX(-50%)',
                zIndex: '2147483647',
                width: 'min(720px, calc(100vw - 24px))',
                boxSizing: 'border-box',
                padding: '14px',
                background: 'rgba(255,255,255,0.97)',
                border: '1px solid rgba(17,24,39,0.16)',
                borderRadius: '14px',
                boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
                fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                color: '#0f172a'
            });

            const title = document.createElement('div');
            title.textContent = heading;
            Object.assign(title.style, {
                fontSize: '14px',
                fontWeight: '700',
                lineHeight: '1.35'
            });
            wrap.appendChild(title);

            const text = document.createElement('div');
            text.textContent = body;
            Object.assign(text.style, {
                marginTop: '8px',
                fontSize: '13px',
                lineHeight: '1.35'
            });
            wrap.appendChild(text);

            const actions = document.createElement('div');
            Object.assign(actions.style, {
                display: 'flex',
                justifyContent: 'flex-end',
                gap: '10px',
                marginTop: '12px'
            });

            const button = document.createElement('button');
            button.type = 'button';
            button.textContent = buttonText;
            Object.assign(button.style, {
                padding: '10px 16px',
                borderRadius: '12px',
                border: '0',
                background: actionValue === 'close' ? '#dc2626' : '#2563eb',
                color: '#fff',
                fontWeight: '800',
                fontSize: '13px',
                cursor: 'pointer'
            });
            button.onclick = () => {
                window.automationBatchAction = actionValue;
                wrap.remove();
            };

            actions.appendChild(button);
            wrap.appendChild(actions);
            document.body.appendChild(wrap);
        }
        """,
        {
            "buttonText": button_text,
            "actionValue": action_value,
            "heading": heading,
            "body": body,
        },
    )
    _wait_for_automation_condition(page, "window.automationBatchAction !== null", controller)
    return page.evaluate("window.automationBatchAction")


def _refresh_for_next_batch(page, controller: AutomationController | None = None):
    _automation_checkpoint(page, controller)
    print("Refreshing page for next batch...")
    try:
        page.keyboard.press("ControlOrMeta+r")
        page.wait_for_load_state("domcontentloaded", timeout=CONFIG["page_load_timeout_ms"])
    except Exception as err:
        print(f"Control+R refresh failed ({err}); using page.reload().")
        page.reload(timeout=CONFIG["page_load_timeout_ms"], wait_until="domcontentloaded")
    _wait_for_d365_ready(page, "refreshed page")
    _automation_checkpoint(page, controller)


def _return_to_journal_list(page) -> None:
    """Navigate back to the customer payment journal list (not a posted popout)."""
    d365_url = str(CONFIG.get("d365_url", "")).strip()
    if not d365_url:
        raise ValueError("d365_url is not configured.")
    print("Returning to customer payment journal list...")
    page.goto(
        d365_url,
        timeout=int(CONFIG.get("page_load_timeout_ms", 60000)),
        wait_until="domcontentloaded",
    )
    _wait_for_d365_ready(page, "journal list")


def _journal_lines_grid_visible(page) -> bool:
    try:
        page.locator("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']").first.wait_for(
            state="visible",
            timeout=4000,
        )
        return True
    except PlaywrightTimeoutError:
        return False


def _prepare_journal_lines_for_bulk_paste(page) -> None:
    """Open a fresh journal lines grid for the next bulk paste batch."""
    if _journal_lines_grid_visible(page):
        print("Journal lines grid already visible; returning to list for a new journal.")
    _return_to_journal_list(page)
    _open_journal_lines(page)


def _group_records_by_sub_batch(records):
    grouped = {}
    for record in records:
        sub_batch_id = str(record.get("sub_batch_id", "")).strip()
        batch_id = str(record.get("batch_id", "")).strip()
        key = sub_batch_id or batch_id or "sub_batch_1"
        grouped.setdefault(key, []).append(record)
    return list(grouped.items())


def _chunk_records(records, size=BULK_PASTE_BATCH_SIZE):
    chunks = []
    for idx in range(0, len(records), size):
        chunks.append(records[idx: idx + size])
    return chunks


def _set_page_clipboard(page, text: str) -> None:
    escaped = json.dumps(text)
    page.evaluate(
        f"""
        async () => {{
            await navigator.clipboard.writeText({escaped});
        }}
        """
    )


def _disable_automation_visual_overlays(page) -> None:
    """Remove fake cursor, click glow, and highlight listeners for manual control."""
    page.evaluate(
        """
        () => {
            if (typeof window.__automationDisableVisualEnhancements === 'function') {
                window.__automationDisableVisualEnhancements();
            }
            document.querySelectorAll(
                '.human-cursor, .click-glow-effect, .element-highlight-rect'
            ).forEach((el) => el.remove());
            if (window.__automationClickHandler) {
                document.removeEventListener('click', window.__automationClickHandler, true);
                window.__automationClickHandler = null;
            }
        }
        """
    )


def _show_bulk_paste_toast(page, duration_ms: int = 3000) -> None:
    page.evaluate(
        """
        (durationMs) => {
            const oldNotice = document.getElementById('automation-paste-toast');
            if (oldNotice) oldNotice.remove();

            const notice = document.createElement('div');
            notice.id = 'automation-paste-toast';
            notice.innerHTML = `
              <div style="font-weight:800; font-size:14px; line-height:1.2;">Data pasted</div>
              <div style="font-weight:600; font-size:13px; line-height:1.35; opacity:0.95; margin-top:4px;">
                Review rows, then click <span style="font-weight:900;">Post</span>.
              </div>
            `;
            Object.assign(notice.style, {
                position: 'fixed',
                top: '150px',
                left: '50%',
                transform: 'translateX(-50%)',
                zIndex: '2147483647',
                width: 'min(720px, calc(100vw - 24px))',
                boxSizing: 'border-box',
                padding: '12px 14px',
                background: 'rgba(11, 95, 255, 0.92)',
                color: '#fff',
                borderRadius: '14px',
                fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                boxShadow: '0 14px 40px rgba(0,0,0,0.28)',
                transition: 'opacity 0.35s ease'
            });
            document.body.appendChild(notice);
            setTimeout(() => {
                notice.style.opacity = '0';
                setTimeout(() => notice.remove(), 350);
            }, durationMs);
        }
        """,
        duration_ms,
    )


def _wait_for_bulk_post_click(page, *, controller: AutomationController | None = None):
    page.evaluate(
        """
        () => {
            window.postClicked = false;
            const oldNotice = document.getElementById('automation-post-notice');
            if (oldNotice) oldNotice.remove();

            if (window.__postHandler) {
                document.removeEventListener('click', window.__postHandler, true);
            }
            window.__postHandler = function postHandler(e) {
                const el = e.target;
                const txt = (el && (el.innerText || el.textContent)) ? (el.innerText || el.textContent) : '';
                if (txt && txt.trim() === 'Post') {
                    window.postClicked = true;
                }
                const btn = el && el.closest ? el.closest('button') : null;
                if (btn && btn.innerText && btn.innerText.trim() === 'Post') {
                    window.postClicked = true;
                }
            };
            document.addEventListener('click', window.__postHandler, true);
        }
        """
    )
    _wait_for_automation_condition(page, "window.postClicked === true", controller)


def _show_all_completed_overlay(page, controller: AutomationController | None = None) -> None:
    if controller is not None:
        controller.complete()
        controller.show_controls(page)
    page.evaluate(
        """
        () => {
            window.automationAllCompleted = false;
            const oldWrap = document.getElementById('automation-all-completed');
            if (oldWrap) oldWrap.remove();

            const wrap = document.createElement('div');
            wrap.id = 'automation-all-completed';
            Object.assign(wrap.style, {
                position: 'fixed',
                top: '150px',
                left: '50%',
                transform: 'translateX(-50%)',
                zIndex: '2147483647',
                width: 'min(720px, calc(100vw - 24px))',
                boxSizing: 'border-box',
                padding: '14px',
                background: 'rgba(255,255,255,0.97)',
                border: '1px solid rgba(17,24,39,0.16)',
                borderRadius: '14px',
                boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
                fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                color: '#0f172a'
            });

            const title = document.createElement('div');
            title.textContent = 'All transactions completed';
            Object.assign(title.style, { fontSize: '15px', fontWeight: '800', lineHeight: '1.35' });
            wrap.appendChild(title);

            const text = document.createElement('div');
            text.textContent = 'You can close this browser window when finished.';
            Object.assign(text.style, { marginTop: '8px', fontSize: '13px', lineHeight: '1.35' });
            wrap.appendChild(text);

            const actions = document.createElement('div');
            Object.assign(actions.style, {
                display: 'flex',
                justifyContent: 'flex-end',
                gap: '10px',
                marginTop: '12px'
            });

            const button = document.createElement('button');
            button.type = 'button';
            button.textContent = 'OK';
            Object.assign(button.style, {
                padding: '10px 16px',
                borderRadius: '12px',
                border: '0',
                background: '#16a34a',
                color: '#fff',
                fontWeight: '800',
                fontSize: '13px',
                cursor: 'pointer'
            });
            button.onclick = () => {
                window.automationAllCompleted = true;
                wrap.remove();
            };
            actions.appendChild(button);
            wrap.appendChild(actions);
            document.body.appendChild(wrap);
        }
        """
    )
    _wait_for_automation_condition(page, "window.automationAllCompleted === true", controller)


def _extract_chunk_voucher_values(page, chunk_size: int):
    voucher_values = _extract_voucher_values(page)
    if len(voucher_values) >= chunk_size:
        return voucher_values[-chunk_size:]
    return voucher_values


def _wait_for_journal_grid_idle(page, *, settle_ms: int = 500) -> None:
    timeout_ms = int(CONFIG.get("page_load_timeout_ms", 60000))
    try:
        page.locator("#ShellBlockingDiv").wait_for(
            state="hidden",
            timeout=timeout_ms,
        )
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    try:
        page.get_by_text(
            "Please wait. We're processing your request.",
            exact=False,
        ).last.wait_for(state="hidden", timeout=timeout_ms)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    if settle_ms > 0:
        page.wait_for_timeout(settle_ms)


def _confirm_unsaved_changes_dialog(page, wait_ms: int = 0) -> bool:
    attempts = max(1, wait_ms // 100 + 1)
    for attempt in range(attempts):
        clicked = page.evaluate(
            """
            () => {
                const visible = (el) => {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.display !== 'none' && style.visibility !== 'hidden'
                        && rect.width > 0 && rect.height > 0;
                };
                const prompt = /do you want to save your changes|save your changes/i;
                const matches = [];
                for (const button of document.querySelectorAll('button')) {
                    const text = (button.innerText || '').trim().toLowerCase();
                    if (!visible(button) || !['save', 'yes'].includes(text)) continue;
                    let ancestor = button.parentElement;
                    for (let distance = 1; ancestor && distance <= 8; distance += 1) {
                        if (prompt.test(ancestor.innerText || '')) {
                            matches.push({ button, distance });
                            break;
                        }
                        ancestor = ancestor.parentElement;
                    }
                }
                matches.sort((a, b) => a.distance - b.distance);
                if (!matches.length) return false;
                matches[0].button.click();
                return true;
            }
            """
        )
        if clicked:
            print("Confirmed D365 save-changes prompt.")
            try:
                page.get_by_text(
                    "Do you want to save your changes?",
                    exact=False,
                ).first.wait_for(
                    state="hidden",
                    timeout=int(CONFIG.get("page_load_timeout_ms", 60000)),
                )
            except (PlaywrightTimeoutError, PlaywrightError):
                pass
            _wait_for_journal_grid_idle(page)
            return True
        if attempt + 1 < attempts:
            page.wait_for_timeout(100)
    return False


def _save_journal_grid(page, label: str, *, settle_ms: int = 500) -> None:
    if not _confirm_unsaved_changes_dialog(page):
        clicked = page.evaluate(
            """
            () => {
                const button = document.querySelector(
                    'button[name="SystemDefinedSaveButton"],'
                    + 'button[data-dyn-controlname="SystemDefinedSaveButton"]'
                );
                if (!button || button.disabled) return false;
                button.click();
                return true;
            }
            """
        )
        if not clicked:
            page.get_by_role("button", name=" Save").first.click(
                timeout=5000,
                force=True,
            )
        _confirm_unsaved_changes_dialog(page, wait_ms=5000)
        _wait_for_journal_grid_idle(page, settle_ms=settle_ms)
    print(f"Saved journal grid ({label}).")


def _activate_pasted_account_field(page, row_index: int):
    _wait_for_journal_grid_idle(page)
    for _ in range(3):
        if not _click_pasted_journal_account_cell(page, row_index):
            continue
        if _confirm_unsaved_changes_dialog(page, wait_ms=500):
            _wait_for_journal_grid_idle(page)
            continue
        page.wait_for_timeout(250)
        field = _pasted_account_field_for_row(page, row_index)
        if field is not None and field.is_editable():
            return field
    details = page.evaluate(
        """
        () => ({
            active: {
                tag: document.activeElement?.tagName || '',
                id: document.activeElement?.id || '',
                aria: document.activeElement?.getAttribute('aria-label') || ''
            },
            accountControls: [...document.querySelectorAll(
                "[id*='AccountNum'],[aria-label='Account'],[data-dyn-controlname*='Account']"
            )].slice(0, 12).map((el) => ({
                tag: el.tagName,
                id: el.id || '',
                role: el.getAttribute('role') || '',
                aria: el.getAttribute('aria-label') || '',
                control: el.getAttribute('data-dyn-controlname') || '',
                readonly: Boolean(el.readOnly)
            })),
            headers: [...document.querySelectorAll('th,[role="columnheader"]')]
                .slice(0, 30).map((el) => (el.innerText || '').trim()),
            tbodyRows: [...document.querySelectorAll('tbody')].map(
                (body) => body.querySelectorAll('tr').length
            )
        })
        """
    )
    raise RuntimeError(
        f"Could not activate Account field for pasted row {row_index + 1}. "
        f"Grid diagnostics: {json.dumps(details, ensure_ascii=True)}"
    )


def _fill_account_lookup(locator, value: str) -> None:
    """Fill a D365 segmented account lookup; retry with sequential typing if fill+Tab fails."""
    expected = str(value).strip()
    _fill_text_field_with_retry(locator, expected)
    try:
        current = (locator.input_value() or "").strip()
    except PlaywrightError:
        current = ""
    if expected.casefold() in current.casefold() or current.casefold() in expected.casefold():
        return
    try:
        locator.click(timeout=5000)
    except PlaywrightError:
        locator.click(force=True, timeout=5000)
    locator.press("ControlOrMeta+a")
    locator.press("Backspace")
    locator.press_sequentially(expected, delay=30)
    locator.press("Tab")
    locator.page.wait_for_timeout(500)


def _account_field_at_row(page, row_index: int):
    return _activate_pasted_account_field(page, row_index)


def _read_account_at_row(page, row_index: int) -> str:
    try:
        field = _journal_field_at_row(
            page,
            row_index,
            css='input[id^="LedgerJournalTrans_AccountNum_"][id$="_input"]',
        )
        return _read_pasted_account(field)
    except (PlaywrightError, RuntimeError):
        pass
    if _click_pasted_journal_account_cell(page, row_index):
        page.wait_for_timeout(150)
        field = _pasted_account_field_for_row(page, row_index)
        if field is not None:
            return _read_pasted_account(field)
    return ""


def _wait_for_account_lookup(page, row_index: int, expected: str, timeout_ms: int = 15000) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        _wait_for_journal_grid_idle(page)
        if _account_matches(_read_account_at_row(page, row_index), expected):
            return True
        page.wait_for_timeout(300)
    return False


def _type_pasted_account(page, record, row_index: int) -> bool:
    expected = str(record["account"]).strip()
    account_field = _account_field_at_row(page, row_index)
    account_field.wait_for(state="visible", timeout=int(CONFIG.get("page_load_timeout_ms", 60000)))
    _fill_account_lookup(account_field, expected)
    if _wait_for_account_lookup(page, row_index, expected):
        print(f"Row {row_index + 1}: typed Account into the segmented lookup.")
        return True
    return False


def _read_journal_grid_headers(page) -> list[str]:
    """Return visible journal line column headers left-to-right from the live D365 grid."""
    headers = page.evaluate(
        """
        () => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                if (st.visibility === 'hidden' || st.display === 'none') return false;
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };

            const dateInput = document.querySelector(
                'tbody tr input[aria-label="Date"], tbody tr input[aria-label="Value date"]'
            );
            let scope = dateInput ? dateInput.closest('table') : null;
            if (!scope) {
                const tbody = document.querySelector('tbody');
                scope = tbody ? tbody.closest('table') : null;
            }
            if (!scope) {
                scope = document;
            }

            const headerCells = [...scope.querySelectorAll('thead th, thead [role="columnheader"]')]
                .filter(visible);
            if (headerCells.length) {
                return headerCells
                    .map((cell) => (cell.innerText || cell.textContent || '').trim())
                    .filter(Boolean);
            }

            const headerRow = [...scope.querySelectorAll('tr,[role="row"]')].find((row) => {
                if (!visible(row)) return false;
                return row.querySelector('[role="columnheader"], th');
            });
            if (!headerRow) {
                return [...document.querySelectorAll('th,[role="columnheader"]')]
                    .filter(visible)
                    .map((cell) => (cell.innerText || cell.textContent || '').trim())
                    .filter(Boolean);
            }

            return [...headerRow.children]
                .filter(visible)
                .map((cell) => (cell.innerText || cell.textContent || '').trim())
                .filter(Boolean);
        }
        """
    )
    result = [str(h).strip() for h in (headers or []) if str(h).strip()]
    print(f"Detected D365 journal columns ({len(result)}): {result}")
    return result


def _repaste_pasted_row(page, record, row_index: int, col_defs, *, use_live_order: bool = False) -> None:
    for _ in range(3):
        _activate_journal_row_for_paste(page, row_index)
        if _confirm_unsaved_changes_dialog(page, wait_ms=500):
            _wait_for_journal_grid_idle(page)
            continue
        break
    else:
        raise RuntimeError(f"Could not activate journal row {row_index + 1} for paste.")

    anchor_label = (col_defs[0].get("label") if col_defs else None) or "Date"
    _focus_journal_row_at(page, row_index, aria_label=anchor_label)
    row_date = _read_date_at_row(page, row_index)
    record_for_paste = dict(record, date=row_date) if row_date else dict(record)
    # On an active row D365 mirrors Value date from Date as soon as paste lands,
    # so leave Value date out of the clipboard and fill it directly afterward.
    record_for_paste["value_date"] = ""
    _set_page_clipboard(
        page,
        build_paste_clipboard_text([record_for_paste], col_defs, use_live_order=use_live_order),
    )
    page.keyboard.press("ControlOrMeta+v")
    page.wait_for_timeout(150)
    print(f"Row {row_index + 1}: pasted complete row.")


def _value_date_needs_fill(page, record, row_index: int) -> bool:
    """Return True when the grid Value date does not match the record."""
    expected = str(record.get("value_date", "")).strip()
    if not expected:
        return False
    if normalize_date is None:
        actual = _read_value_date_at_row(page, row_index)
        return (actual or "").strip().casefold() != expected.casefold()

    expected_norm = normalize_date(expected)
    if not expected_norm:
        return True

    actual = _read_value_date_at_row(page, row_index)
    actual_norm = normalize_date(actual) if actual else None
    if actual_norm == expected_norm:
        return False

    # D365 auto-mirrors Value date to Date on active rows — treat that as wrong.
    row_date = _read_date_at_row(page, row_index)
    date_norm = normalize_date(row_date) if row_date else None
    if actual_norm and date_norm and actual_norm == date_norm and expected_norm != date_norm:
        return True

    return not actual_norm or actual_norm != expected_norm


def _fill_value_date_at_row(
    page,
    record,
    row_index: int,
    *,
    force: bool = False,
) -> None:
    """Directly type Value date on the specific journal row (never page-wide .first)."""
    expected = str(record.get("value_date", "")).strip()
    if not expected:
        return
    expected = _d365_date(expected, expected)
    if not force and not _value_date_needs_fill(page, record, row_index):
        return
    rows = page.locator(_JOURNAL_DATA_ROW_SEL)
    if rows.count() <= row_index:
        print(f"Row {row_index + 1}: Warning: Value date row not found for direct fill.")
        return
    row = rows.nth(row_index)
    field = row.locator('input[aria-label="Value date"]:not([readonly])')
    if field.count() == 0:
        field = row.locator('input[aria-label="Value date"]')
    if field.count() == 0:
        print(f"Row {row_index + 1}: Warning: Value date input not found for direct fill.")
        return
    target = field.first
    try:
        target.wait_for(state="visible", timeout=3000)
        _click_locator_robust(target, page, label=f"Value date input row {row_index + 1}")
        target.press("ControlOrMeta+a")
        target.press("Backspace")
        target.fill(str(expected), timeout=10000)
        target.evaluate(
            "el => { el.dispatchEvent(new Event('change', { bubbles: true })); el.blur(); }"
        )
        print(f"Row {row_index + 1}: filled Value date '{expected}' directly.")
    except (PlaywrightError, PlaywrightTimeoutError, RuntimeError) as err:
        print(f"Row {row_index + 1}: Warning: Could not fill Value date '{expected}': {err}")


def _read_pasted_account(account_field) -> str:
    value = account_field.evaluate(
        """
        (field) => {
            const values = [];
            const add = (value) => {
                const text = String(value || '').trim();
                if (text && text.toLowerCase() !== 'account'
                    && !text.startsWith('Segmented Entry control')) {
                    values.push(text);
                }
            };
            add(field.value);
            add(field.getAttribute('title'));
            add(field.getAttribute('aria-valuetext'));

            const cell = field.closest('td,[role="gridcell"]');
            if (cell) {
                for (const input of cell.querySelectorAll('input')) {
                    add(input.value);
                    add(input.getAttribute('title'));
                    add(input.getAttribute('aria-valuetext'));
                }
                for (const line of String(cell.innerText || '').split('\\n')) add(line);
            }
            return values.find((text) => text.length <= 120) || '';
        }
        """
    )
    return " ".join(str(value or "").split())


def _account_matches(actual: str, expected: str) -> bool:
    """Compare segmented-entry text ignoring case, spacing and separators."""
    def strip(text) -> str:
        return re.sub(r"[^a-z0-9]", "", str(text or "").casefold())

    expected_key = strip(expected)
    return bool(expected_key) and expected_key in strip(actual)


def _method_of_payment_matches(actual: str, expected: str) -> bool:
    actual = str(actual or "").strip().casefold()
    expected = str(expected or "").strip().casefold()
    return not expected or bool(actual) and (
        actual == expected.split()[0] or expected.startswith(actual)
    )


def _pasted_row_issues(page, record, row_index: int) -> list[str]:
    issues = []
    expected_account = str(record.get("account", "")).strip()
    actual_account = _read_account_at_row(page, row_index)
    if not _account_matches(actual_account, expected_account):
        issues.append(
            f"Account '{actual_account or '(blank)'}' instead of '{expected_account}'"
        )

    expected_method = str(record.get("method_of_payment", "")).strip()
    actual_method = _read_method_of_payment_at_row(page, row_index)
    if not _method_of_payment_matches(actual_method, expected_method):
        issues.append(
            f"Method of payment '{actual_method or '(blank)'}' instead of "
            f"'{expected_method}'"
        )
    return issues


def _wait_for_pasted_row_issues(
    page,
    record,
    row_index: int,
    timeout_ms: int = 5000,
) -> list[str]:
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        issues = _pasted_row_issues(page, record, row_index)
        if not issues or time.monotonic() >= deadline:
            return issues
        page.wait_for_timeout(500)


def _read_account_at_row_without_click(page, row_index: int) -> str:
    """Read the Account cell without clicking, so it is safe to poll before saving."""
    try:
        field = _journal_field_at_row(
            page,
            row_index,
            css='input[id^="LedgerJournalTrans_AccountNum_"][id$="_input"]',
        )
        return _read_pasted_account(field)
    except (PlaywrightError, RuntimeError):
        return ""


def _read_method_of_payment_at_row_safe(page, row_index: int) -> str:
    try:
        return _read_method_of_payment_at_row(page, row_index)
    except (PlaywrightError, RuntimeError):
        return ""


def _pending_pasted_row_fields(page, record, row_index: int) -> list[str]:
    """Return which of Account / Method of payment still don't match the record, without clicking."""
    pending = []
    expected_account = str(record.get("account", "")).strip()
    if expected_account and not _account_matches(
        _read_account_at_row_without_click(page, row_index), expected_account
    ):
        pending.append("Account")
    expected_method = str(record.get("method_of_payment", "")).strip()
    if expected_method and not _method_of_payment_matches(
        _read_method_of_payment_at_row_safe(page, row_index), expected_method
    ):
        pending.append("Method of payment")
    return pending


def _repair_missing_pasted_accounts(page, records, col_defs, *, use_live_order: bool = False) -> int:
    missing_source_rows = [
        str(index)
        for index, record in enumerate(records, start=1)
        if not str(record.get("account", "")).strip()
    ]
    if missing_source_rows:
        raise ValueError(
            "Missing account in source row(s): " + ", ".join(missing_source_rows)
        )

    repaired_indices = []
    for index, record in enumerate(records):
        expected = str(record["account"]).strip()

        # Dismiss any lingering validation dialogs before reading/repairing
        _dismiss_d365_validation_dialog(page)
        _wait_for_journal_grid_idle(page)

        actual = _read_account_at_row(page, index)
        if _account_matches(actual, expected):
            continue

        print(f"Row {index + 1}: Account was blank after paste; typing it directly.")
        typed_ok = _type_pasted_account(page, record, index)
        if typed_ok:
            repaired_indices.append(index)
            continue

        print(f"Row {index + 1}: typed Account did not stick; re-pasting the row.")
        # Dismiss validation dialogs before re-paste attempt
        _dismiss_d365_validation_dialog(page)
        _wait_for_journal_grid_idle(page)
        _repaste_pasted_row(page, record, index, col_defs, use_live_order=use_live_order)

        repaired_indices.append(index)

    if not repaired_indices:
        return 0

    _wait_for_journal_grid_idle(page)
    _dismiss_d365_validation_dialog(page)
    _save_journal_grid(page, "after Account repairs")
    _wait_for_journal_grid_idle(page)
    for index in repaired_indices:
        expected = str(records[index]["account"]).strip()
        actual = _read_account_at_row(page, index)
        if not _account_matches(actual, expected):
            raise RuntimeError(
                f"D365 kept Account '{actual or '(blank)'}' instead of '{expected}' "
                f"in pasted row {index + 1}."
            )
        print(f"Row {index + 1}: Account confirmed after save.")

    print(f"Repaired {len(repaired_indices)} Account field(s) after bulk paste.")
    return len(repaired_indices)


def _read_method_of_payment_at_row(page, row_index: int) -> str:
    """Read the Method of Payment value for a pasted journal row.

    Uses JS to read the cell text directly from the grid (works even when
    the cell is not active/editable), avoiding stale input element reads.
    """
    value = page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden'
                    && r.width > 0 && r.height > 0;
            };

            const grid = [...document.querySelectorAll('[role="grid"]')]
                .find((el) => (el.getAttribute('aria-label') || '') === 'Journal lines');
            const gridRows = grid
                ? [...grid.querySelectorAll('[role="row"][id*="-row-"]')].filter(visible)
                : [];
            const gridRow = gridRows[rowIndex];
            const gridCell = gridRow && [...gridRow.querySelectorAll('[role="gridcell"]')]
                .find((el) => el.id.endsWith('-LedgerJournalTrans_PaymMode1'));
            if (gridCell) {
                const input = gridCell.querySelector(
                    'input[aria-label="Method of payment"]:not([id^="Sel_"])'
                );
                const value = (input?.value || '').trim();
                return value || (gridCell.innerText || '').trim();
            }

            // Find column index from header
            const headers = [...document.querySelectorAll('th,[role="columnheader"]')];
            const header = headers.find(h => {
                const text = (h.innerText || '').trim();
                return text === 'Method of payment';
            });
            if (!header) return '';

            const headerRow = header.closest('tr,[role="row"]');
            if (!headerRow) return '';
            const colIndex = [...headerRow.children].indexOf(header);
            if (colIndex < 0) return '';

            // Find the grid body scope
            let scope = header.parentElement;
            while (scope && !scope.querySelector('tbody tr')) scope = scope.parentElement;
            if (!scope) return '';

            const rows = [...scope.querySelectorAll('tbody tr')].filter(visible);
            const row = rows[rowIndex];
            if (!row) return '';

            const cell = row.children[colIndex];
            if (!cell) return '';

            // Try input value first (if this row is active)
            const input = cell.querySelector('input[aria-label="Method of payment"]:not([id^="Sel_"])');
            if (input) {
                const v = (input.value || '').trim();
                if (v) return v;
            }

            // Fall back to cell text content (works for inactive rows)
            const text = (cell.innerText || '').trim();
            // Filter out header-like text or placeholder text
            if (text && text !== 'Method of payment') return text;
            return '';
        }
        """,
        row_index,
    )
    return str(value or "").strip()


def _read_date_at_row(page, row_index: int) -> str:
    """Read the Date value D365 auto-filled for a journal row (exact locale format)."""
    value = page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden'
                    && r.width > 0 && r.height > 0;
            };

            const grid = [...document.querySelectorAll('[role="grid"]')]
                .find((el) => (el.getAttribute('aria-label') || '') === 'Journal lines');
            const gridRows = grid
                ? [...grid.querySelectorAll('[role="row"][id*="-row-"]')].filter(visible)
                : [];
            const gridRow = gridRows[rowIndex];
            const gridCell = gridRow && [...gridRow.querySelectorAll('[role="gridcell"]')]
                .find((el) => el.id.endsWith('-LedgerJournalTrans_TransDate'));
            if (gridCell) {
                const input = gridCell.querySelector('input[aria-label="Date"]');
                const value = (input?.value || input?.getAttribute('title') || '').trim();
                return value || (gridCell.innerText || '').trim();
            }

            const headers = [...document.querySelectorAll('th,[role="columnheader"]')];
            const header = headers.find((h) => (h.innerText || '').trim() === 'Date');
            if (!header) return '';
            const headerRow = header.closest('tr,[role="row"]');
            if (!headerRow) return '';
            const colIndex = [...headerRow.children].indexOf(header);
            if (colIndex < 0) return '';

            let scope = header.parentElement;
            while (scope && !scope.querySelector('tbody tr')) scope = scope.parentElement;
            if (!scope) return '';

            const rows = [...scope.querySelectorAll('tbody tr')].filter(visible);
            const row = rows[rowIndex];
            if (!row) return '';
            const cell = row.children[colIndex];
            if (!cell) return '';

            const input = cell.querySelector('input[aria-label="Date"]');
            if (input) {
                const v = (input.value || '').trim();
                if (v) return v;
            }
            const text = (cell.innerText || '').trim();
            return text && text !== 'Date' ? text : '';
        }
        """,
        row_index,
    )
    return str(value or "").strip()


def _read_value_date_at_row(page, row_index: int) -> str:
    """Read the Value date shown for a journal row, without activating/clicking it."""
    value = page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden'
                    && r.width > 0 && r.height > 0;
            };

            const grid = [...document.querySelectorAll('[role="grid"]')]
                .find((el) => (el.getAttribute('aria-label') || '') === 'Journal lines');
            const gridRows = grid
                ? [...grid.querySelectorAll('[role="row"][id*="-row-"]')].filter(visible)
                : [];
            const gridRow = gridRows[rowIndex];
            const gridCell = gridRow && [...gridRow.querySelectorAll('[role="gridcell"]')]
                .find((el) => el.id.endsWith('-LedgerJournalTrans_TransDateValueDate')
                    || el.id.endsWith('-LedgerJournalTrans_ValueDate'));
            if (gridCell) {
                const input = gridCell.querySelector('input[aria-label="Value date"]');
                const value = (input?.value || input?.getAttribute('title') || '').trim();
                return value || (gridCell.innerText || '').trim();
            }

            const headers = [...document.querySelectorAll('th,[role="columnheader"]')];
            const header = headers.find((h) => (h.innerText || '').trim() === 'Value date');
            if (!header) return '';
            const headerRow = header.closest('tr,[role="row"]');
            if (!headerRow) return '';
            const colIndex = [...headerRow.children].indexOf(header);
            if (colIndex < 0) return '';

            let scope = header.parentElement;
            while (scope && !scope.querySelector('tbody tr')) scope = scope.parentElement;
            if (!scope) return '';

            const rows = [...scope.querySelectorAll('tbody tr')].filter(visible);
            const row = rows[rowIndex];
            if (!row) return '';
            const cell = row.children[colIndex];
            if (!cell) return '';

            const input = cell.querySelector('input[aria-label="Value date"]');
            if (input) {
                const v = (input.value || '').trim();
                if (v) return v;
            }
            const text = (cell.innerText || '').trim();
            return text && text !== 'Value date' ? text : '';
        }
        """,
        row_index,
    )
    return str(value or "").strip()


def _ensure_value_date_at_row(page, record, row_index: int) -> bool:
    """Verify Value date matches the record (read-only check); activate + fix only if wrong.

    Returns True if a fix was applied (caller should save afterward).
    """
    if not _value_date_needs_fill(page, record, row_index):
        return False
    expected = str(record.get("value_date", "")).strip()
    actual = _read_value_date_at_row(page, row_index)
    print(f"Row {row_index + 1}: Value date '{actual or '(blank)'}' does not match expected '{expected}'; fixing.")
    try:
        _activate_journal_row_for_paste(page, row_index)
        _confirm_unsaved_changes_dialog(page, wait_ms=500)
        _wait_for_journal_grid_idle(page)
    except (PlaywrightError, PlaywrightTimeoutError, RuntimeError):
        pass
    _fill_value_date_at_row(page, record, row_index, force=True)
    return True


def _records_with_d365_dates(page, records: list, anchor_row_index: int = 0) -> list:
    """Copy records with Date set to D365's auto-filled value (same format D365 uses).

    Pasting an empty Date cell clears D365's auto-fill, so read the grid value after
    focus and echo it back in column 1 for alignment without changing the date.
    """
    d365_date = _read_date_at_row(page, anchor_row_index)
    if not d365_date:
        page.wait_for_timeout(300)
        d365_date = _read_date_at_row(page, anchor_row_index)
    if d365_date:
        print(f"Using D365 auto-filled Date '{d365_date}' for paste alignment.")
        return [dict(record, date=d365_date) for record in records]
    print("Warning: Could not read D365 auto-filled Date; Date column may be blank after paste.")
    return [dict(record) for record in records]


def _read_voucher_at_row(page, row_index: int) -> str:
    """Read the Voucher value for a journal row (works whether the row is active or not).

    D365 assigns the Voucher number only once a row's Save has actually committed
    server-side, which makes it a reliable signal that the row is fully saved —
    unlike Account/Method of payment, which can read blank right up until save commits.
    """
    value = page.evaluate(
        """
        (rowIndex) => {
            const visible = (el) => {
                if (!el) return false;
                const st = window.getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return st.display !== 'none' && st.visibility !== 'hidden'
                    && r.width > 0 && r.height > 0;
            };

            const grid = [...document.querySelectorAll('[role="grid"]')]
                .find((el) => (el.getAttribute('aria-label') || '') === 'Journal lines');
            const gridRows = grid
                ? [...grid.querySelectorAll('[role="row"][id*="-row-"]')].filter(visible)
                : [];
            const gridRow = gridRows[rowIndex];
            const gridCell = gridRow && [...gridRow.querySelectorAll('[role="gridcell"]')]
                .find((el) => el.id.endsWith('-LedgerJournalTrans_Voucher'));
            if (gridCell) {
                const input = gridCell.querySelector('input[aria-label="Voucher"]');
                const value = (input?.value || input?.getAttribute('title') || '').trim();
                return value || (gridCell.innerText || '').trim();
            }

            const headers = [...document.querySelectorAll('th,[role="columnheader"]')];
            const header = headers.find((h) => (h.innerText || '').trim() === 'Voucher');
            if (!header) return '';
            const headerRow = header.closest('tr,[role="row"]');
            if (!headerRow) return '';
            const colIndex = [...headerRow.children].indexOf(header);
            if (colIndex < 0) return '';

            let scope = header.parentElement;
            while (scope && !scope.querySelector('tbody tr')) scope = scope.parentElement;
            if (!scope) return '';

            const rows = [...scope.querySelectorAll('tbody tr')].filter(visible);
            const row = rows[rowIndex];
            if (!row) return '';
            const cell = row.children[colIndex];
            if (!cell) return '';

            const input = cell.querySelector('input[aria-label="Voucher"]');
            if (input) {
                const v = (input.value || '').trim();
                if (v) return v;
            }
            const text = (cell.innerText || '').trim();
            return text && text !== 'Voucher' ? text : '';
        }
        """,
        row_index,
    )
    return str(value or "").strip()


def _wait_for_voucher_at_row(page, row_index: int, timeout_ms: int = 15000) -> bool:
    """Poll a row's Voucher column until D365 has assigned one (i.e. the row's save committed)."""
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        if _read_voucher_at_row(page, row_index):
            return True
        if time.monotonic() >= deadline:
            return False
        page.wait_for_timeout(250)


def _repair_missing_pasted_method_of_payment(page, records) -> int:
    """Re-fill Method of Payment for rows where bulk paste left it blank."""
    # Wait for the grid to be ready — it may have reloaded after account repair saves
    _wait_for_journal_grid_idle(page)
    try:
        _wait_for_journal_grid_ready(page, timeout_ms=15000)
    except (PlaywrightTimeoutError, PlaywrightError):
        print("Warning: Journal grid not ready for Method of payment repair; skipping.")
        return 0

    repaired = 0
    for index, record in enumerate(records):
        expected = str(record.get("method_of_payment", "")).strip()
        if not expected:
            continue

        _dismiss_d365_validation_dialog(page)
        _wait_for_journal_grid_idle(page)

        # Read cell text via JS (works without activating the row)
        actual = _read_method_of_payment_at_row(page, index)
        if _method_of_payment_matches(actual, expected):
            print(f"Row {index + 1}: Method of payment already set to '{actual}'.")
            continue

        if not actual:
            print(f"Row {index + 1}: Method of payment was blank after paste; typing it directly.")
        else:
            print(f"Row {index + 1}: Method of payment '{actual}' does not match expected '{expected}'; retyping.")

        try:
            # D365 renders editable lookup inputs only for the active row.
            _activate_pasted_account_field(page, index)
            _wait_for_journal_grid_idle(page)

            # After activation, only the active row has inputs rendered.
            # Find the Method of Payment input directly on the page.
            mop_sel = 'input[aria-label="Method of payment"]:not([id^="Sel_"]):not([readonly])'
            mop_field = page.locator(mop_sel)
            if mop_field.count() == 0:
                mop_field = page.locator('input[aria-label="Method of payment"]:not([id^="Sel_"])')
            if mop_field.count() == 0:
                print(f"Row {index + 1}: Warning: Method of payment input not found after row activation; skipping.")
                continue
            field = mop_field.first
            _fill_text_field_with_retry(field, expected)
            print(f"Row {index + 1}: typed Method of payment '{expected}'.")
            repaired += 1

        except Exception as err:
            print(f"Row {index + 1}: Warning: Could not fill Method of payment '{expected}': {err}")

    if repaired:
        _wait_for_journal_grid_idle(page)
        _dismiss_d365_validation_dialog(page)
        _save_journal_grid(page, "after Method of payment repairs")
        _wait_for_journal_grid_idle(page)
        print(f"Repaired {repaired} Method of payment field(s) after bulk paste.")
    return repaired


def _paste_bulk_chunk(page, records, col_defs, controller: AutomationController | None = None) -> None:
    _automation_checkpoint(page, controller)
    if build_paste_clipboard_text is None or col_defs_from_live_headers is None:
        raise RuntimeError("clipboard_export module is not available.")

    live_headers = _read_journal_grid_headers(page)
    if not live_headers:
        raise RuntimeError(
            "Could not read D365 journal column headers. "
            "Ensure journal lines grid is visible before bulk paste."
        )

    ordered_cols = col_defs_from_live_headers(live_headers, col_defs)
    validate_live_paste_columns(ordered_cols)
    mapped_labels = [
        col.get("label", "")
        for col in ordered_cols
        if col.get("key") and col.get("key") != SPACER_FIELD_KEY
    ]
    print(f"Bulk paste column map ({len(ordered_cols)} cols): {mapped_labels}")

    # Paste anchors on Date (column 1). Read D365's auto-filled Date after focus and echo
    # it back so paste alignment is preserved without clearing the Date cell.
    _focus_journal_row_for_paste(page, 0)
    page.wait_for_timeout(200)
    records_for_paste = _records_with_d365_dates(page, records, anchor_row_index=0)
    paste_text = build_paste_clipboard_text(records_for_paste, ordered_cols, use_live_order=True)
    col_count = paste_text.split("\r\n")[0].count("\t") + 1 if paste_text else 0
    print(f"Bulk paste: {len(records)} rows x {col_count} columns (starting at Date column).")
    _set_page_clipboard(page, paste_text)
    page.keyboard.press("ControlOrMeta+v")
    page.wait_for_timeout(400)

    _wait_for_journal_grid_idle(page)
    _dismiss_d365_validation_dialog(page)
    _gate_if_validation_issue(page, controller)

    # Save immediately — waiting for Account/Method of payment to resolve beforehand never
    # succeeds (D365 only finalizes those segmented lookups as part of the save commit
    # itself), so any pre-save wait here is pure dead time.
    if not _confirm_unsaved_changes_dialog(page, wait_ms=2000):
        _save_journal_grid(page, "after bulk paste")
    else:
        print("Confirmed save-changes prompt after bulk paste.")
    _wait_for_journal_grid_idle(page)
    try:
        _wait_for_journal_grid_ready(page, timeout_ms=15000)
    except PlaywrightTimeoutError:
        print("Warning: Journal grid not fully ready before complete row re-paste; continuing.")

    _gate_if_validation_issue(page, controller)

    # Voucher is only assigned once a row's save has actually committed server-side, so
    # confirm it here (fast poll, not a fixed delay) before touching rows individually.
    last_index = len(records) - 1
    if not _wait_for_voucher_at_row(page, last_index):
        print(f"Warning: Row {last_index + 1} has no Voucher yet after bulk save; continuing anyway.")

    print("Bulk paste saved; re-pasting rows 2 onward and saving each row before advancing.")

    # Row 1 was bulk-pasted and saved already — fix Value date once here, not inside the loop.
    if _value_date_needs_fill(page, records[0], 0):
        try:
            _activate_journal_row_for_paste(page, 0)
            _confirm_unsaved_changes_dialog(page, wait_ms=500)
        except (PlaywrightError, PlaywrightTimeoutError, RuntimeError):
            pass
        _fill_value_date_at_row(page, records[0], 0, force=True)
        _save_journal_grid(page, "after row 1 Value date", settle_ms=100)
        _dismiss_d365_validation_dialog(page)

    all_issues: list[tuple[int, list[str]]] = []
    for index in range(len(records)):
        if index > 0:
            _repaste_and_save_row_with_retry(
                page,
                records_for_paste[index],
                index,
                ordered_cols,
                controller,
                use_live_order=True,
            )

        issues = _pending_pasted_row_fields(page, records[index], index)
        if not issues:
            print(f"Row {index + 1}: Account and Method of payment confirmed after save.")
            continue

        print(f"Row {index + 1}: {', '.join(issues)} still blank after save; retrying this row once...")
        _repaste_and_save_row_with_retry(
            page,
            records[index],
            index,
            ordered_cols,
            controller,
            use_live_order=True,
            save_label=f"after row {index + 1} immediate retry",
        )
        issues = _wait_for_pasted_row_issues(page, records[index], index, timeout_ms=1500)
        if issues:
            all_issues.append((index, issues))
        else:
            print(f"Row {index + 1}: confirmed after retry.")

    if all_issues:
        details = "; ".join(f"row {index + 1}: {', '.join(issues)}" for index, issues in all_issues)
        raise RuntimeError(f"D365 row(s) remained incomplete after save: {details}")


def _process_bulk_paste_chunks(page, records, col_defs, controller: AutomationController | None = None):
    chunks = _chunk_records(records, BULK_PASTE_BATCH_SIZE)
    if not chunks:
        print("No records to process in bulk paste mode.")
        return []

    all_processed = []
    total_chunks = len(chunks)
    for chunk_index, chunk in enumerate(chunks, start=1):
        _automation_checkpoint(page, controller)
        print(
            f"Starting bulk paste chunk {chunk_index}/{total_chunks} "
            f"({len(chunk)} transactions)"
        )
        _paste_bulk_chunk(page, chunk, col_defs, controller)
        _disable_automation_visual_overlays(page)
        _show_bulk_paste_toast(page, 3000)
        _wait_for_post_and_confirmation(page, bulk=True, controller=controller)
        _automation_checkpoint(page, controller)

        try:
            page.get_by_text("List General Payment fee Bank").click()
        except PlaywrightError as err:
            print(f"Warning: Could not return to voucher list view: {err}")

        voucher_values = _extract_chunk_voucher_values(page, len(chunk))
        print(f"Chunk {chunk_index} vouchers: {voucher_values}")
        if len(voucher_values) != len(chunk):
            print(
                "Warning: Voucher count does not match chunk size "
                f"({len(voucher_values)} vs {len(chunk)}). "
                "Patching will use minimum count by order."
            )
        ok, patch_msg = _bulk_update_receipts(chunk, voucher_values)
        print(f"bulkUpdateReceipt chunk {chunk_index}: {'OK' if ok else 'SKIP/FAIL'}")
        print(patch_msg)
        all_processed.extend(chunk)

        if chunk_index < total_chunks:
            action = _wait_for_batch_action(
                page,
                is_last_sub_batch=False,
                current_index=chunk_index,
                total_sub_batches=total_chunks,
                controller=controller,
            )
            if action == "close":
                print("User closed bulk paste before next chunk.")
                break
            _automation_checkpoint(page, controller)
            _prepare_journal_lines_for_bulk_paste(page)

    _show_all_completed_overlay(page, controller)
    print(f"Bulk paste completed for {len(all_processed)} records.")
    return all_processed


def _process_sub_batch(page, records, controller: AutomationController | None = None):
    iterated_records = []
    reuse_same_row_next = False
    manual_save_done = False

    for idx, record in enumerate(records):
        _automation_checkpoint(page, controller)
        print(f"Processing record {idx + 1}/{len(records)}")
        force_manual_wipe_before_fill = False
        if idx > 0 and not reuse_same_row_next:
            _dismiss_d365_validation_dialog(page)
            try:
                page.get_by_role("button", name=" New").click()
            except PlaywrightError:
                page.get_by_role("button", name=" New").first.click()
            page.wait_for_timeout(150)
            _focus_new_journal_line(page, idx)

        def _verify_row_clear_state():
            try:
                return bool(
                    page.evaluate(
                        """
                        () => {
                            const isVisible = (el) => {
                                if (!el) return false;
                                const st = window.getComputedStyle(el);
                                if (st.visibility === 'hidden' || st.display === 'none') return false;
                                const r = el.getBoundingClientRect();
                                return r.width > 0 && r.height > 0;
                            };
                            const pick = (selector) => {
                                const nodes = [...document.querySelectorAll(selector)];
                                return nodes.find((el) => {
                                    const id = el.id || '';
                                    return isVisible(el) && !id.startsWith('Sel_');
                                }) || null;
                            };
                            const read = (selector) => {
                                const el = pick(selector);
                                return el ? String(el.value || '').trim() : '';
                            };
                            const vals = [
                                read("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']"),
                                read("input[aria-label='Credit']"),
                                read("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']"),
                                read("input[aria-label='Reference date']"),
                                read("input[aria-label='Payment reference']"),
                                read("input[aria-label='Method of payment']")
                            ];
                            return vals.every((v) => v === '');
                        }
                        """
                    )
                )
            except Exception as err:
                print(f"Row clear verification unavailable: {err}")
                return False

        def _clear_current_row_fields_fallback():
            value_date_clear = _journal_field_locator(page, idx, aria_label="Value date")
            credit_clear = _journal_field_locator(page, idx, aria_label="Credit")
            ref_date_clear = _journal_field_locator(page, idx, aria_label="Reference date")
            pay_ref_clear = _journal_field_locator(page, idx, aria_label="Payment reference")
            account_clear = _journal_field_locator(
                page, idx, css='input[id^="LedgerJournalTrans_AccountNum_"][id$="_input"]'
            )
            offset_account_clear = _journal_field_locator(
                page, idx, css='input[id^="LedgerJournalTrans_OffsetAccount_"][id$="_input"]'
            )
            method_clear = _journal_field_locator(
                page,
                idx,
                css='input[aria-label="Method of payment"]:not([id^="Sel_"])',
            )

            def _wipe(locator):
                try:
                    locator.fill("")
                except Exception:
                    try:
                        locator.click()
                        locator.press("ControlOrMeta+a")
                        locator.press("Backspace")
                    except Exception:
                        pass

            _wipe(value_date_clear)
            _wipe(account_clear)
            _wipe(credit_clear)
            _wipe(offset_account_clear)
            _wipe(ref_date_clear)
            _wipe(pay_ref_clear)
            _wipe(method_clear)
            return _verify_row_clear_state()

        def _clear_current_row_fields_fast():
            try:
                fast_clear_ok = page.evaluate(
                    """
                    (rowIndex) => {
                        const isVisible = (el) => {
                            if (!el) return false;
                            const st = window.getComputedStyle(el);
                            if (st.visibility === 'hidden' || st.display === 'none') return false;
                            const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0;
                        };

                        const accountSel = "input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']";
                        const trs = [...document.querySelectorAll('tbody tr')].filter((tr) => {
                            const inp = tr.querySelector(accountSel);
                            return inp && isVisible(tr);
                        });

                        let activeRow = document.querySelector("tr[aria-selected='true'], tr[aria-current='true']");
                        if (!activeRow || !isVisible(activeRow)) {
                            for (let i = trs.length - 1; i >= 0; i--) {
                                const inp = trs[i].querySelector(accountSel);
                                if (!inp || inp.readOnly) continue;
                                if (!(inp.value || '').trim()) {
                                    activeRow = trs[i];
                                    break;
                                }
                            }
                        }
                        if (!activeRow && trs.length) {
                            const pick = rowIndex < trs.length ? rowIndex : trs.length - 1;
                            activeRow = trs[pick];
                        }

                        const pickInput = (selector) => {
                            if (activeRow) {
                                const scoped = [...activeRow.querySelectorAll(selector)].find((el) => {
                                    const id = el.id || '';
                                    return isVisible(el) && !id.startsWith('Sel_') && !el.readOnly;
                                });
                                if (scoped) return scoped;
                            }
                            return null;
                        };

                        const clearValue = (selector) => {
                            const el = pickInput(selector);
                            if (!el) return false;
                            el.value = '';
                            el.dispatchEvent(new Event('input', { bubbles: true }));
                            el.dispatchEvent(new Event('change', { bubbles: true }));
                            return true;
                        };

                        const changed = [
                            clearValue("input[aria-label='Value date']"),
                            clearValue("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']"),
                            clearValue("input[aria-label='Credit']"),
                            clearValue("input[id^='LedgerJournalTrans_OffsetAccount_'][id$='_input']"),
                            clearValue("input[aria-label='Reference date']"),
                            clearValue("input[aria-label='Payment reference']"),
                            clearValue("input[aria-label='Method of payment']")
                        ];
                        return changed.some(Boolean);
                    }
                    """,
                    idx,
                )
                if not fast_clear_ok:
                    print("Fast clear did not target row fields; using fallback clear.")
                    return _clear_current_row_fields_fallback()
                if not _verify_row_clear_state():
                    print("Fast clear verification failed; retrying with fallback clear.")
                    return _clear_current_row_fields_fallback()
                return True
            except Exception as err:
                print(f"Fast clear failed, falling back to control-by-control clear: {err}")
                return _clear_current_row_fields_fallback()

        if reuse_same_row_next:
            print(f"[{time.time():.3f}] Continue path: entering same-row reuse iteration.")
            force_manual_wipe_before_fill = True
            reuse_same_row_next = False

        val_date = _d365_date(record.get("value_date"), "2/17/2026")
        acc_no = str(record.get("account", "")).strip()
        credit_amt = record.get("credit", "25,000")
        offset_acc = str(record.get("offset_account", "")).strip()
        ref_date = _d365_date(record.get("reference_date"), "2/17/2026")
        pay_ref = record.get("payment_reference", "YESBANK")
        pay_method = record.get("method_of_payment", "Wire Wire Transfer")
        if not acc_no:
            raise ValueError(f"Missing account in record {idx + 1}; refusing implicit fallback account.")

        value_date_loc = _journal_field_locator(page, idx, aria_label="Value date")
        credit_loc = _journal_field_locator(page, idx, aria_label="Credit")
        ref_date_loc = _journal_field_locator(page, idx, aria_label="Reference date")
        pay_ref_loc = _journal_field_locator(page, idx, aria_label="Payment reference")
        offset_account_field = _journal_field_locator(
            page, idx, css='input[id^="LedgerJournalTrans_OffsetAccount_"][id$="_input"]'
        )
        paym_mode_input = _journal_field_locator(
            page,
            idx,
            css='input[aria-label="Method of payment"]:not([id^="Sel_"])',
        )

        if force_manual_wipe_before_fill:
            print(f"[{time.time():.3f}] Continue path: starting manual row wipe before fill.")
            for loc in (pay_ref_loc, value_date_loc, credit_loc, ref_date_loc, paym_mode_input):
                try:
                    loc.click()
                    loc.press("ControlOrMeta+a")
                    loc.press("Backspace")
                except Exception:
                    pass
            try:
                offset_account_field.click()
                offset_account_field.press("ControlOrMeta+a")
                offset_account_field.press("Backspace")
            except Exception:
                pass

        account_field = _journal_field_locator(
            page, idx, css='input[id^="LedgerJournalTrans_AccountNum_"][id$="_input"]'
        )
        account_field.wait_for(
            state="visible",
            timeout=int(CONFIG.get("page_load_timeout_ms", 60000)),
        )
        _fill_text_field_with_retry(account_field, acc_no)

        _fill_text_field_with_retry(value_date_loc, val_date)
        _fill_text_field_with_retry(pay_ref_loc, pay_ref)
        _fill_text_field(credit_loc, credit_amt)

        if offset_acc:
            try:
                offset_account_field.wait_for(state="visible", timeout=5000)
                _fill_text_field_with_retry(offset_account_field, offset_acc)
                print(f"Filled offset account: {offset_acc}")
            except Exception as err:
                print(f"Warning: Could not fill offset account '{offset_acc}': {err}")

        _fill_text_field_with_retry(ref_date_loc, ref_date)

        try:
            _fill_text_field_with_retry(paym_mode_input, pay_method)
            print(f"Filled method of payment: {pay_method}")
        except Exception as err:
            print(f"Warning: Could not fill method of payment '{pay_method}': {err}")

        try:
            current_pay_ref = (pay_ref_loc.input_value() or "").strip()
        except Exception:
            current_pay_ref = ""

        if not current_pay_ref:
            print("Payment reference is empty before Save. Refilling and reapplying method.")
            try:
                _fill_text_field_with_retry(pay_ref_loc, pay_ref)
            except PlaywrightError as err:
                print(f"Warning: Could not refill Payment reference '{pay_ref}': {err}")
            try:
                _fill_text_field_with_retry(paym_mode_input, pay_method)
            except Exception as err:
                print(f"Warning: Could not refill method of payment '{pay_method}': {err}")

            try:
                current_pay_ref = (pay_ref_loc.input_value() or "").strip()
            except Exception:
                current_pay_ref = ""

            if not current_pay_ref:
                is_last_record = idx == len(records) - 1
                if is_last_record:
                    print("Payment reference is still empty on final record. Showing info-only prompt.")
                    page.evaluate(
                        """
                        () => {
                            window.automationFinalDuplicateSaveClicked = false;
                            const oldInfo = document.getElementById('automation-final-duplicate-info');
                            if (oldInfo) oldInfo.remove();
                            if (window.__automationFinalDuplicateSaveHandler) {
                                document.removeEventListener('click', window.__automationFinalDuplicateSaveHandler, true);
                            }

                            const wrap = document.createElement('div');
                            wrap.id = 'automation-final-duplicate-info';
                            Object.assign(wrap.style, {
                                position: 'fixed',
                                top: '150px',
                                left: '50%',
                                transform: 'translateX(-50%)',
                                zIndex: '2147483647',
                                width: 'min(760px, calc(100vw - 24px))',
                                boxSizing: 'border-box',
                                padding: '14px',
                                background: 'rgba(255,255,255,0.97)',
                                border: '1px solid rgba(17,24,39,0.16)',
                                borderRadius: '14px',
                                boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
                                fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                                color: '#0f172a'
                            });

                            const msg = document.createElement('div');
                            msg.textContent = 'Last record may already be posted. Please recheck. If needed, delete the row and click Save and Post.';
                            Object.assign(msg.style, {
                                fontSize: '14px',
                                fontWeight: '700',
                                lineHeight: '1.35'
                            });
                            wrap.appendChild(msg);
                            document.body.appendChild(wrap);

                            window.__automationFinalDuplicateSaveHandler = function finalDuplicateSaveHandler(e) {
                                const el = e.target;
                                const txt = (el && (el.innerText || el.textContent)) ? (el.innerText || el.textContent).trim() : '';
                                const btn = el && el.closest ? el.closest('button') : null;
                                const btnTxt = (btn && (btn.innerText || btn.textContent))
                                    ? (btn.innerText || btn.textContent).trim()
                                    : '';
                                if (txt === 'Save' || btnTxt === 'Save') {
                                    window.automationFinalDuplicateSaveClicked = true;
                                }
                            };
                            document.addEventListener('click', window.__automationFinalDuplicateSaveHandler, true);
                        }
                        """
                    )
                    _wait_for_automation_condition(
                        page,
                        "window.automationFinalDuplicateSaveClicked === true",
                        controller,
                    )
                    page.evaluate(
                        """
                        () => {
                            const info = document.getElementById('automation-final-duplicate-info');
                            if (info) info.remove();
                            if (window.__automationFinalDuplicateSaveHandler) {
                                document.removeEventListener('click', window.__automationFinalDuplicateSaveHandler, true);
                                window.__automationFinalDuplicateSaveHandler = null;
                            }
                        }
                        """
                    )
                    manual_save_done = True
                    continue

                print("Payment reference is still empty. Showing already-posted decision dialog.")
                page.evaluate(
                    """
                    () => {
                        window.automationAlreadyPostedDecision = null;
                        const oldDlg = document.getElementById('automation-already-posted-gate');
                        if (oldDlg) oldDlg.remove();

                        const wrap = document.createElement('div');
                        wrap.id = 'automation-already-posted-gate';
                        Object.assign(wrap.style, {
                            position: 'fixed',
                            top: '150px',
                            left: '50%',
                            transform: 'translateX(-50%)',
                            zIndex: '2147483647',
                            width: 'min(720px, calc(100vw - 24px))',
                            boxSizing: 'border-box',
                            padding: '14px',
                            background: 'rgba(255,255,255,0.97)',
                            border: '1px solid rgba(17,24,39,0.16)',
                            borderRadius: '14px',
                            boxShadow: '0 18px 50px rgba(0,0,0,0.22)',
                            fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif',
                            color: '#0f172a'
                        });

                        const msg = document.createElement('div');
                        msg.textContent = 'Current record is already posted. Please click Continue for next iteration.';
                        Object.assign(msg.style, {
                            fontSize: '14px',
                            fontWeight: '700',
                            lineHeight: '1.35'
                        });
                        wrap.appendChild(msg);

                        const actions = document.createElement('div');
                        Object.assign(actions.style, {
                            display: 'flex',
                            justifyContent: 'flex-end',
                            gap: '10px',
                            marginTop: '12px'
                        });

                        const continueBtn = document.createElement('button');
                        continueBtn.type = 'button';
                        continueBtn.textContent = 'Continue';
                        Object.assign(continueBtn.style, {
                            padding: '9px 14px',
                            borderRadius: '12px',
                            border: '0',
                            background: '#16a34a',
                            color: '#fff',
                            fontWeight: '800',
                            fontSize: '13px',
                            cursor: 'pointer'
                        });
                        continueBtn.onclick = () => {
                            window.automationAlreadyPostedDecision = 'continue';
                            wrap.remove();
                        };

                        const closeBtn = document.createElement('button');
                        closeBtn.type = 'button';
                        closeBtn.textContent = 'Close Window';
                        Object.assign(closeBtn.style, {
                            padding: '9px 14px',
                            borderRadius: '12px',
                            border: '0',
                            background: '#dc2626',
                            color: '#fff',
                            fontWeight: '800',
                            fontSize: '13px',
                            cursor: 'pointer'
                        });
                        closeBtn.onclick = () => {
                            window.automationAlreadyPostedDecision = 'close';
                            wrap.remove();
                        };

                        actions.appendChild(continueBtn);
                        actions.appendChild(closeBtn);
                        wrap.appendChild(actions);
                        document.body.appendChild(wrap);
                    }
                    """
                )
                _wait_for_automation_condition(
                    page,
                    "window.automationAlreadyPostedDecision !== null",
                    controller,
                )
                decision = page.evaluate("window.automationAlreadyPostedDecision")
                if decision == "close":
                    raise AutomationStoppedByUser("User closed automation during already-posted handling.")
                print(f"[{time.time():.3f}] Continue decision received; scheduling same-row reuse.")
                reuse_same_row_next = True
                _clear_current_row_fields_fast()
                continue

        _finalize_journal_line_before_new(page, idx, ref_date_loc, ref_date, pay_method)
        iterated_records.append(record)
        print(f"Prepared record {idx + 1}/{len(records)}. Save/Post will run after all rows.")

    if not iterated_records:
        print("No records prepared; skipping Save/Post/PATCH flow.")
        return []

    if manual_save_done:
        print("Manual Save already completed after final duplicate notice. Skipping automatic Save click.")
    else:
        page.get_by_role("button", name=" Save").click()
        print("Saved. Waiting for user to click Post...")

    _wait_for_post_and_confirmation(page, controller=controller)

    processed_records = list(iterated_records)
    print(f"Continue confirmed. Proceeding with {len(processed_records)} records for bulk patch.")
    return processed_records


def test_final8(records=None, *, bulk_paste_mode=True, clipboard_col_defs=None, controller: AutomationController | None = None):
    issues = get_config_issues(require_auth_state=True)
    if issues:
        raise ValueError("Configuration issue(s):\n- " + "\n- ".join(issues))

    if not records and len(sys.argv) > 1:
        try:
            with open(sys.argv[1], "r", encoding="utf-8") as f:
                records = json.load(f)
        except (OSError, json.JSONDecodeError) as err:
            print(f"Failed to load records from file '{sys.argv[1]}': {err}")

    if not records:
        records = [{
            "date": "2/17/2026",
            "value_date": "2/17/2026",
            "account": "-23620",
            "credit": "25,000",
            "offset_account": "Axis Bank Limited A/C No.",
            "reference_date": "2/17/2026",
            "payment_reference": "YESBANK",
            "method_of_payment": "Wire Wire Transfer",
            "batch_id": "DEFAULT_BATCH",
            "sub_batch_id": "DEFAULT_BATCH_1",
        }]

    mode_label = "bulk paste" if bulk_paste_mode else "legacy fill"
    print(f"Starting automation ({mode_label}) with {len(records)} records...")
    if not sync_playwright:
        print("Playwright is not installed. Skipping automation.")
        raise RuntimeError("Playwright is not installed.")

    batch_id = str(records[0].get("batch_id", "")).strip() or "UNASSIGNED"
    keep_browser_open = bulk_paste_mode

    with sync_playwright() as playwright:
        browser = None
        context = None
        page = None
        automation_error = None
        try:
            browser, screen_w, screen_h = _create_browser(playwright)
            try:
                context = _create_context(browser, screen_w, screen_h, use_storage_state=True)
            except (PlaywrightError, OSError, ValueError) as err:
                print(f"Auth state unavailable, starting a fresh context: {err}")
                context = _create_context(browser, screen_w, screen_h, use_storage_state=False)
            if not bulk_paste_mode:
                context.add_init_script(VISUAL_ENHANCEMENT_SCRIPT)
            page = context.new_page()

            print("Navigating to D365...")
            page.goto(
                CONFIG["d365_url"],
                timeout=CONFIG["page_load_timeout_ms"],
                wait_until="domcontentloaded",
            )
            _wait_for_d365_ready(page)
            if controller is not None:
                controller.attach_page(page)
            _automation_checkpoint(page, controller)

            if bulk_paste_mode:
                if normalize_col_defs is None or build_paste_clipboard_text is None:
                    raise RuntimeError("clipboard_export module is required for bulk paste mode.")
                col_defs = normalize_col_defs(clipboard_col_defs or load_clipboard_col_defs())
                try:
                    context.grant_permissions(
                        ["clipboard-read", "clipboard-write"],
                        origin=page.url,
                    )
                except Exception as err:
                    print(f"Warning: Could not grant clipboard permissions: {err}")
                chunk_count = len(_chunk_records(records, BULK_PASTE_BATCH_SIZE))
                print(
                    f"Bulk paste mode: {len(records)} records in {chunk_count} "
                    f"batch(es) of up to {BULK_PASTE_BATCH_SIZE}."
                )
                company = d365_company_from_config()
                if company:
                    for record in records:
                        if not str(record.get("company", "")).strip():
                            record["company"] = company
                    print(f"Bulk paste company: {company}")
                _open_journal_lines(page)
                _process_bulk_paste_chunks(page, records, col_defs, controller)
            else:
                sub_batch_groups = _group_records_by_sub_batch(records)
                print(f"Main batch {batch_id} contains {len(sub_batch_groups)} sub-batches.")
                total_sub_batches = len(sub_batch_groups)
                for sub_batch_index, (sub_batch_id, sub_batch_records) in enumerate(sub_batch_groups, start=1):
                    _automation_checkpoint(page, controller)
                    print(
                        f"Starting sub-batch {sub_batch_index}/{total_sub_batches}: "
                        f"{sub_batch_id} ({len(sub_batch_records)} transactions)"
                    )
                    _open_journal_lines(page)
                    processed_records = _process_sub_batch(page, sub_batch_records, controller)

                    if processed_records:
                        try:
                            page.get_by_text("List General Payment fee Bank").click()
                        except PlaywrightError as err:
                            print(f"Warning: Could not return to voucher list view: {err}")
                        voucher_values = _extract_voucher_values(page)
                        print(voucher_values)
                        if len(voucher_values) != len(processed_records):
                            print(
                                "Warning: Voucher count does not match processed record count "
                                f"({len(voucher_values)} vs {len(processed_records)}). "
                                "Patching will use minimum count by order."
                            )
                        ok, patch_msg = _bulk_update_receipts(processed_records, voucher_values)
                        print(f"bulkUpdateReceipt status for {sub_batch_id}: {'OK' if ok else 'SKIP/FAIL'}")
                        print(patch_msg)
                    else:
                        print(f"Sub-batch {sub_batch_id} produced no processed records; skipping PATCH.")

                    is_last_sub_batch = sub_batch_index == total_sub_batches
                    action = _wait_for_batch_action(
                        page,
                        is_last_sub_batch=is_last_sub_batch,
                        current_index=sub_batch_index,
                        total_sub_batches=total_sub_batches,
                        controller=controller,
                    )
                    if action == "close":
                        break
                    _refresh_for_next_batch(page, controller)

            _persist_storage_state(context)
        except Exception as err:
            automation_error = err
            print(f"Automation failed: {err}")
            if keep_browser_open and browser is not None and not isinstance(err, AutomationStoppedByUser):
                print("Automation failed — browser left open for inspection. Close it when done.")
                try:
                    while browser.is_connected():
                        _automation_checkpoint(page, controller)
                        _interruptible_wait(page, 1000, controller)
                except AutomationStoppedByUser as stop_err:
                    automation_error = stop_err
                except Exception:
                    pass
        else:
            if keep_browser_open and browser is not None:
                print("Waiting for user to close the browser...")
                try:
                    while browser.is_connected():
                        _automation_checkpoint(page, controller)
                        _interruptible_wait(page, 1000, controller)
                except AutomationStoppedByUser as err:
                    automation_error = err
                except Exception:
                    pass
        finally:
            quit_requested = bool(controller and controller.quit_requested)
            if controller is not None:
                controller.cleanup()
            if keep_browser_open and not quit_requested:
                if automation_error is not None:
                    print("Browser closed after automation failure.")
                else:
                    print("Bulk paste completed — browser closed by user.")
            if not keep_browser_open or quit_requested:
                if context is not None:
                    context.close()
                if browser is not None:
                    browser.close()

        if automation_error is not None:
            raise automation_error

def test_loginfunctionality():
    issues = get_config_issues(require_auth_state=False)
    if issues:
        raise ValueError("Configuration issue(s):\n- " + "\n- ".join(issues))

    print("Starting Login Automation...")
    if not sync_playwright:
        print("Playwright is not installed.")
        raise RuntimeError("Playwright is not installed.")

    with sync_playwright() as playwright:
        browser = None
        context = None
        page = None
        try:
            browser, screen_w, screen_h = _create_browser(playwright)
            context = _create_context(browser, screen_w, screen_h, use_storage_state=False)
            context.add_init_script(VISUAL_ENHANCEMENT_SCRIPT)
            page = context.new_page()

            page.goto(
                CONFIG["d365_url"],
                timeout=CONFIG["page_load_timeout_ms"],
                wait_until="domcontentloaded",
            )

            print("Waiting for manual login. Click the 'Login Success' button in the browser when done.")
            login_button_timeout_ms = int(CONFIG.get("manual_login_button_timeout_ms", 1800000))
            login_success_event = threading.Event()

            def _notify_login_success():
                login_success_event.set()

            page.expose_function("notifyLoginSuccess", _notify_login_success)
            start_time = time.time()

            while not login_success_event.is_set():
                elapsed_ms = int((time.time() - start_time) * 1000)
                if elapsed_ms > login_button_timeout_ms:
                    raise TimeoutError(
                        f"Timed out waiting for Login Success click after {login_button_timeout_ms} ms."
                    )

                try:
                    page.evaluate("""
                        () => {
                            const existing = document.getElementById('automation-login-success-btn');
                            if (existing) return;
                            const btn = document.createElement('button');
                            btn.id = 'automation-login-success-btn';
                            btn.textContent = 'Login Success';
                            Object.assign(btn.style, {
                                position: 'fixed',
                                top: '16px',
                                right: '16px',
                                zIndex: '2147483647',
                                padding: '10px 14px',
                                background: '#0b5fff',
                                color: '#ffffff',
                                border: 'none',
                                borderRadius: '8px',
                                fontSize: '14px',
                                fontWeight: '600',
                                cursor: 'pointer',
                                boxShadow: '0 6px 18px rgba(0,0,0,0.25)'
                            });
                            btn.addEventListener('click', () => {
                                btn.textContent = 'Login Confirmed';
                                btn.disabled = true;
                                btn.style.opacity = '0.75';
                                if (typeof window.notifyLoginSuccess === 'function') {
                                    window.notifyLoginSuccess();
                                }
                            });
                            document.body.appendChild(btn);
                        }
                    """)
                except Exception:
                    # Page may be transitioning during redirects; retry until stable.
                    pass

                time.sleep(0.5)

            current_url = page.url.lower()
            if "login.microsoftonline.com" in current_url:
                raise ValueError(
                    "Login is not complete yet. You are still on Microsoft login page. "
                    "Finish sign-in and click 'Login Success' only after landing on D365."
                )

            _wait_for_d365_ready(page, "login flow")
            display_name = _extract_signed_in_user(page)

            page.close()
            _persist_storage_state(context)
            print(f"Login session saved at: {CONFIG['auth_json_path']}")
            return _auth_result(True, display_name=display_name)
        finally:
            if page is not None and not page.is_closed():
                page.close()
            if context is not None:
                context.close()
            if browser is not None:
                browser.close()

if __name__ == "__main__":
    test_final8()
