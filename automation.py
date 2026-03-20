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
from urllib.parse import urlparse

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

  document.addEventListener('mousemove', (e) => {
    cursor.style.left = e.clientX + 'px';
    cursor.style.top = e.clientY + 'px';
  });

  document.addEventListener('mousedown', () => {
    cursor.style.transform = 'translate(-2px, -2px) scale(0.85)';
  });

  document.addEventListener('mouseup', () => {
    cursor.style.transform = 'translate(-2px, -2px) scale(1)';
  });

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

  document.addEventListener('click', (e) => {
    const glow = document.createElement('div');
    glow.classList.add('click-glow-effect');
    glow.style.left = e.clientX + 'px';
    glow.style.top = e.clientY + 'px';
    document.body.appendChild(glow);
    setTimeout(() => glow.remove(), 600);
    highlightElement(e.target);
  }, true);

  document.addEventListener('focus', (e) => highlightElement(e.target), true);
  document.addEventListener('input', (e) => highlightElement(e.target), true);
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
            "browser_slow_mo_ms": 1000,
            "page_load_timeout_ms": 60000,
            "page_load_wait_seconds": 5,
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
    # Use a near-fullscreen viewport to avoid opening in a cramped/half-screen size.
    root = tk.Tk()
    root.withdraw()
    try:
        screen_w = max(1280, int(root.winfo_screenwidth() * 0.96))
        screen_h = max(720, int(root.winfo_screenheight() * 0.9))
    finally:
        root.destroy()
    return screen_w, screen_h


def _create_browser(playwright, *, headless: bool | None = None):
    screen_w, screen_h = _get_browser_dimensions()
    viewport_size = f"{screen_w},{screen_h}"
    if headless is None:
        headless = CONFIG["browser_headless"]
    browser = playwright.chromium.launch(
        headless=headless,
        slow_mo=CONFIG["browser_slow_mo_ms"],
        args=[f"--window-size={viewport_size}"],
    )
    return browser, screen_w, screen_h


def _create_context(browser, screen_w, screen_h, use_storage_state=True):
    context_args = {"viewport": {"width": screen_w, "height": screen_h}}
    if use_storage_state:
        context_args["storage_state"] = CONFIG["auth_json_path"]
    return browser.new_context(**context_args)


def _persist_storage_state(context):
    auth_path = Path(CONFIG["auth_json_path"])
    auth_path.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(auth_path))
    if not auth_path.exists() or auth_path.stat().st_size == 0:
        raise OSError(f"Auth session file was not created correctly at: {auth_path}")


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


def is_playwright_chromium_available() -> tuple[bool, str]:
    """Check whether Playwright Chromium executable is available."""
    if not sync_playwright:
        return False, "Playwright Python package is not installed."
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


def _normalize_reference_date(date_text: str) -> str:
    """Convert to M/D/YYYY for D365 reference date field."""
    raw = str(date_text or "").strip()
    if not raw:
        return raw

    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(raw, fmt)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except ValueError:
            continue
    # Fallback: keep input if parsing fails.
    return raw

class AutomationStoppedByUser(RuntimeError):
    pass


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
        page.get_by_role("button", name=" New").first.click()
    except PlaywrightError:
        page.get_by_role("button", name=" New").click()
    page.locator("#JournalName_3_0_0").get_by_role("button", name="Open").click()
    page.get_by_role("row", name=CONFIG["journal_name"], exact=True).get_by_label("Name").click()
    page.get_by_role("button", name="Lines", exact=True).click()


def _extract_voucher_values(page):
    voucher_values = []
    max_retries = 20
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
        page.wait_for_timeout(2000)
    return voucher_values


def _wait_for_post_click(page):
    page.evaluate(
        """
        () => {
            window.postClicked = false;
            const oldNotice = document.getElementById('automation-post-notice');
            if (oldNotice) oldNotice.remove();

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
        """
    )
    page.wait_for_function("window.postClicked === true", timeout=0)


def _wait_for_post_confirmation(page):
    page.evaluate(
        """
        () => {
          window.automationPostContinueClicked = false;
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
            window.automationPostContinueClicked = true;
            wrap.remove();
          });

          actions.appendChild(button);
          wrap.appendChild(actions);
          document.body.appendChild(wrap);
        }
        """
    )
    page.wait_for_function("window.automationPostContinueClicked === true", timeout=0)


def _wait_for_batch_action(page, *, is_last_sub_batch: bool, current_index: int, total_sub_batches: int) -> str:
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
    page.wait_for_function("window.automationBatchAction !== null", timeout=0)
    return page.evaluate("window.automationBatchAction")


def _refresh_for_next_batch(page):
    print("Refreshing page for next batch...")
    try:
        page.keyboard.press("Control+R")
        page.wait_for_load_state("domcontentloaded", timeout=CONFIG["page_load_timeout_ms"])
    except Exception as err:
        print(f"Control+R refresh failed ({err}); using page.reload().")
        page.reload(timeout=CONFIG["page_load_timeout_ms"], wait_until="domcontentloaded")
    _wait_for_d365_ready(page, "refreshed page")


def _group_records_by_sub_batch(records):
    grouped = {}
    for record in records:
        sub_batch_id = str(record.get("sub_batch_id", "")).strip()
        batch_id = str(record.get("batch_id", "")).strip()
        key = sub_batch_id or batch_id or "sub_batch_1"
        grouped.setdefault(key, []).append(record)
    return list(grouped.items())


def _process_sub_batch(page, records):
    iterated_records = []
    reuse_same_row_next = False
    manual_save_done = False

    for idx, record in enumerate(records):
        print(f"Processing record {idx + 1}/{len(records)}")
        force_manual_wipe_before_fill = False
        if idx > 0 and not reuse_same_row_next:
            try:
                page.get_by_role("button", name=" New").click()
            except PlaywrightError:
                page.get_by_role("button", name=" New").first.click()
            time.sleep(0.5)

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
            value_date_clear = page.get_by_role("combobox", name="Value date")
            credit_clear = page.get_by_role("textbox", name="Credit")
            ref_date_clear = page.get_by_role("combobox", name="Reference date")
            pay_ref_clear = page.get_by_role("textbox", name="Payment reference")
            account_clear = page.locator("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']")
            method_clear = page.get_by_label("Method of payment")
            if idx > 0:
                value_date_clear = value_date_clear.first
                credit_clear = credit_clear.first
                ref_date_clear = ref_date_clear.first
                pay_ref_clear = pay_ref_clear.first
                account_clear = account_clear.first
                method_clear = method_clear.first

            def _wipe(locator):
                try:
                    locator.fill("")
                except Exception:
                    try:
                        locator.click()
                        locator.press("Control+A")
                        locator.press("Backspace")
                    except Exception:
                        pass

            _wipe(value_date_clear)
            _wipe(account_clear)
            _wipe(credit_clear)
            _wipe(ref_date_clear)
            _wipe(pay_ref_clear)
            _wipe(method_clear)
            return _verify_row_clear_state()

        def _clear_current_row_fields_fast():
            try:
                fast_clear_ok = page.evaluate(
                    """
                    () => {
                        const isVisible = (el) => {
                            if (!el) return false;
                            const st = window.getComputedStyle(el);
                            if (st.visibility === 'hidden' || st.display === 'none') return false;
                            const r = el.getBoundingClientRect();
                            return r.width > 0 && r.height > 0;
                        };

                        const accountInput = document.querySelector("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']");
                        const rowCandidates = [
                            document.querySelector("tr[aria-selected='true']"),
                            document.querySelector("tr[aria-current='true']"),
                            accountInput ? accountInput.closest('tr') : null
                        ].filter(Boolean);
                        const activeRow = rowCandidates.find(isVisible) || null;

                        const pickInput = (selector) => {
                            const scoped = activeRow ? [...activeRow.querySelectorAll(selector)] : [];
                            const global = [...document.querySelectorAll(selector)];
                            const all = [...scoped, ...global];
                            return all.find((el) => {
                                const id = el.id || '';
                                return isVisible(el) && !id.startsWith('Sel_');
                            }) || null;
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
                            clearValue("input[aria-label='Reference date']"),
                            clearValue("input[aria-label='Payment reference']"),
                            clearValue("input[aria-label='Method of payment']")
                        ];
                        return changed.some(Boolean);
                    }
                    """
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

        val_date = _normalize_reference_date(record.get("value_date", "2/17/2026"))
        acc_no = str(record.get("account", "")).strip()
        credit_amt = record.get("credit", "25,000")
        ref_date = _normalize_reference_date(record.get("reference_date", "2/17/2026"))
        pay_ref = record.get("payment_reference", "YESBANK")
        pay_method = record.get("method_of_payment", "Wire Wire Transfer")
        if not acc_no:
            raise ValueError(f"Missing account in record {idx + 1}; refusing implicit fallback account.")

        value_date_loc = page.get_by_role("combobox", name="Value date")
        credit_loc = page.get_by_role("textbox", name="Credit")
        ref_date_loc = page.get_by_role("combobox", name="Reference date")
        pay_ref_loc = page.get_by_role("textbox", name="Payment reference")
        if idx > 0:
            value_date_loc = value_date_loc.first
            credit_loc = credit_loc.first
            ref_date_loc = ref_date_loc.first
            pay_ref_loc = pay_ref_loc.first

        if force_manual_wipe_before_fill:
            print(f"[{time.time():.3f}] Continue path: starting manual row wipe before fill.")
            for loc in (pay_ref_loc, value_date_loc, credit_loc, ref_date_loc):
                try:
                    loc.click()
                    loc.press("Control+A")
                    loc.press("Backspace")
                except Exception:
                    pass

        pay_ref_loc.click()
        if force_manual_wipe_before_fill:
            try:
                pay_ref_loc.press("Control+A")
                pay_ref_loc.press("Backspace")
            except Exception:
                pass
        pay_ref_loc.press_sequentially(pay_ref, delay=100)

        value_date_loc.press_sequentially(val_date, delay=200)

        account_field = page.locator("input[id^='LedgerJournalTrans_AccountNum_'][id$='_input']")
        if idx > 0:
            account_field = account_field.first
        account_field.wait_for(state="visible", timeout=200)
        account_field.click()
        account_field.press("Control+A")
        account_field.press("Backspace")
        account_field.press_sequentially(acc_no, delay=20)

        try:
            entered_account = (account_field.input_value() or "").strip()
            if entered_account and entered_account.casefold() != acc_no.casefold():
                account_field.click()
                account_field.press("Control+A")
                account_field.press("Backspace")
                account_field.press_sequentially(acc_no, delay=20)
        except PlaywrightError:
            pass

        credit_loc.click()
        if force_manual_wipe_before_fill:
            try:
                credit_loc.press("Control+A")
                credit_loc.press("Backspace")
            except Exception:
                pass
        credit_loc.press_sequentially(credit_amt, delay=200)

        paym_mode_input = page.get_by_label("Method of payment")
        if idx > 0:
            paym_mode_input = paym_mode_input.first
        if force_manual_wipe_before_fill:
            try:
                paym_mode_input.click(force=True)
                paym_mode_input.press("Control+A")
                paym_mode_input.press("Backspace")
            except Exception:
                pass

        def _select_method_of_payment(force=False):
            try:
                current_method = (paym_mode_input.input_value() or "").strip().lower()
                if not force and current_method == str(pay_method).strip().lower():
                    return
            except Exception:
                pass
            try:
                paym_mode_input.scroll_into_view_if_needed(timeout=5000)
            except (PlaywrightTimeoutError, PlaywrightError):
                pass
            paym_mode_input.click(force=True)
            paym_mode_input.press("Alt+ArrowDown")
            method_selected = False
            method_exact = page.locator(
                f"input[aria-label='Method of payment'][id^='Sel_'][title='{pay_method}']"
            )
            if method_exact.count() > 0:
                method_exact.first.click()
                method_selected = True
            if not method_selected:
                try:
                    page.evaluate(
                        """
                        (method) => {
                            const target = [...document.querySelectorAll(
                                "input[aria-label='Method of payment'][id^='Sel_']"
                            )].find(el => {
                                const txt = (el.getAttribute('title') || el.value || '').trim().toLowerCase();
                                return txt === String(method).trim().toLowerCase();
                            });
                            if (!target) throw new Error(`Method not found in dropdown: ${method}`);
                            target.click();
                        }
                        """,
                        pay_method,
                    )
                    method_selected = True
                except Exception:
                    pass
            if not method_selected:
                try:
                    page.evaluate(
                        """
                        () => {
                            const el = document.querySelector("input[aria-label='Method of payment']");
                            if (el) { el.scrollIntoView({block: 'center'}); el.click(); }
                        }
                        """
                    )
                except Exception as err:
                    print(f"Warning: Method of payment all fallbacks failed for '{pay_method}': {err}")

        _select_method_of_payment(force=False)

        ref_date_loc.click()
        if force_manual_wipe_before_fill:
            try:
                ref_date_loc.press("Control+A")
                ref_date_loc.press("Backspace")
            except Exception:
                pass
        ref_date_loc.press_sequentially(ref_date, delay=200)
        try:
            current_ref_val = ref_date_loc.input_value().strip()
            if not current_ref_val:
                ref_date_loc.click()
                ref_date_loc.press_sequentially(ref_date, delay=200)
        except Exception:
            pass

        try:
            value_date_loc.fill(val_date)
            current_value_date = value_date_loc.input_value().strip()
            if current_value_date != val_date:
                value_date_loc.click()
                value_date_loc.press("Control+A")
                value_date_loc.press("Backspace")
                value_date_loc.press_sequentially(val_date, delay=40)
        except Exception:
            pass

        _select_method_of_payment(force=True)

        try:
            current_pay_ref = (pay_ref_loc.input_value() or "").strip()
        except Exception:
            current_pay_ref = ""

        if not current_pay_ref:
            print("Payment reference is empty before Save. Refilling and reapplying method.")
            try:
                pay_ref_loc.click()
                pay_ref_loc.press("Control+A")
                pay_ref_loc.press("Backspace")
                pay_ref_loc.press_sequentially(pay_ref, delay=120)
            except PlaywrightError as err:
                print(f"Warning: Could not refill Payment reference '{pay_ref}': {err}")
            _select_method_of_payment(force=True)

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
                    page.wait_for_function("window.automationFinalDuplicateSaveClicked === true", timeout=0)
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
                page.wait_for_function("window.automationAlreadyPostedDecision !== null", timeout=0)
                decision = page.evaluate("window.automationAlreadyPostedDecision")
                if decision == "close":
                    raise AutomationStoppedByUser("User closed automation during already-posted handling.")
                print(f"[{time.time():.3f}] Continue decision received; scheduling same-row reuse.")
                reuse_same_row_next = True
                _clear_current_row_fields_fast()
                continue

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

    _wait_for_post_click(page)
    print("User clicked Post. Waiting for continue confirmation...")
    _wait_for_post_confirmation(page)

    processed_records = list(iterated_records)
    print(f"Continue confirmed. Proceeding with {len(processed_records)} records for bulk patch.")
    return processed_records


def test_final8(records=None):
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

    print(f"Starting automation with {len(records)} records...")
    if not sync_playwright:
        print("Playwright is not installed. Skipping automation.")
        raise RuntimeError("Playwright is not installed.")

    batch_id = str(records[0].get("batch_id", "")).strip() or "UNASSIGNED"
    sub_batch_groups = _group_records_by_sub_batch(records)
    print(f"Main batch {batch_id} contains {len(sub_batch_groups)} sub-batches.")

    with sync_playwright() as playwright:
        browser = None
        context = None
        try:
            browser, screen_w, screen_h = _create_browser(playwright)
            try:
                context = _create_context(browser, screen_w, screen_h, use_storage_state=True)
            except (PlaywrightError, OSError, ValueError) as err:
                print(f"Auth state unavailable, starting a fresh context: {err}")
                context = _create_context(browser, screen_w, screen_h, use_storage_state=False)
            context.add_init_script(VISUAL_ENHANCEMENT_SCRIPT)
            page = context.new_page()

            print("Navigating to D365...")
            page.goto(
                CONFIG["d365_url"],
                timeout=CONFIG["page_load_timeout_ms"],
                wait_until="domcontentloaded",
            )
            _wait_for_d365_ready(page)

            total_sub_batches = len(sub_batch_groups)
            for sub_batch_index, (sub_batch_id, sub_batch_records) in enumerate(sub_batch_groups, start=1):
                print(
                    f"Starting sub-batch {sub_batch_index}/{total_sub_batches}: "
                    f"{sub_batch_id} ({len(sub_batch_records)} transactions)"
                )
                _open_journal_lines(page)
                processed_records = _process_sub_batch(page, sub_batch_records)

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
                )
                if action == "close":
                    break
                _refresh_for_next_batch(page)

            _persist_storage_state(context)
        finally:
            if context is not None:
                context.close()
            if browser is not None:
                browser.close()

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
