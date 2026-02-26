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
    "auth_json_path",
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


def _create_browser(playwright):
    screen_w, screen_h = _get_browser_dimensions()
    viewport_size = f"{screen_w},{screen_h}"
    browser = playwright.chromium.launch(
        headless=CONFIG["browser_headless"],
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

def test_final8(records=None):
    issues = get_config_issues(require_auth_state=True)
    if issues:
        raise ValueError("Configuration issue(s):\n- " + "\n- ".join(issues))

    # Fallback default if ran without arguments or None passed
    if not records:
        if len(sys.argv) > 1:
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
            "method_of_payment": "Wire Wire Transfer"
        }]

    print(f"Starting automation with {len(records)} records...")
    
    if not sync_playwright:
        print("Playwright is not installed. Skipping automation.")
        return

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

            print("Waiting for page load...")
            for i in range(CONFIG["page_load_wait_seconds"], 0, -1):
                print(f"Loading... {i}")
                time.sleep(1)
            print("Page fully loaded!")

            # Wait for D365 blocking overlay to disappear
            try:
                page.locator("#ShellBlockingDiv").wait_for(state="hidden", timeout=60000)
            except PlaywrightTimeoutError:
                print("Overlay did not disappear in 60s; continuing.")

            # Click "New" - using .first to avoid strict mode violation
            page.get_by_role("button", name=" New").first.click()
            page.locator("#JournalName_3_0_0").get_by_role("button", name="Open").click()
            page.get_by_role("row", name=CONFIG["journal_name"], exact=True).get_by_label("Name").click()
            page.get_by_role("button", name="Lines", exact=True).click()

            processed_records = []

            # Iterate over records
            for idx, record in enumerate(records):
                print(f"Processing record {idx + 1}/{len(records)}")
                if idx > 0:
                    try:
                        page.get_by_role("button", name=" New").click()
                    except PlaywrightError:
                        page.get_by_role("button", name=" New").first.click()
                    time.sleep(3)

                # Extract data
                val_date = _normalize_reference_date(record.get("value_date", "2/17/2026"))
                acc_no = record.get("account", "-23620")
                credit_amt = record.get("credit", "25,000")
                offset_acc = record.get("offset_account", "Axis Bank Limited A/C No.")
                ref_date = _normalize_reference_date(record.get("reference_date", "2/17/2026"))
                pay_ref = record.get("payment_reference", "YESBANK")
                pay_method = record.get("method_of_payment", "Wire Wire Transfer")

                # Fill form
                value_date_loc = page.get_by_role("combobox", name="Value date")
                credit_loc = page.get_by_role("textbox", name="Credit")
                ref_date_loc = page.get_by_role("combobox", name="Reference date")
                pay_ref_loc = page.get_by_role("textbox", name="Payment reference")
                if idx > 0:
                    value_date_loc = value_date_loc.first
                    credit_loc = credit_loc.first
                    ref_date_loc = ref_date_loc.first
                    pay_ref_loc = pay_ref_loc.first

                value_date_loc.press_sequentially(val_date, delay=200)

                account_open_btn = page.locator(
                    "[id^='LedgerJournalTrans_AccountNum_'][id$='_segmentedEntryLookup']"
                ).get_by_role("button", name="Open")
                if idx > 0:
                    account_open_btn = account_open_btn.first
                account_open_btn.click()
                try:
                    page.get_by_role("gridcell", name=acc_no).get_by_label("Customer account").click()
                except PlaywrightError:
                    print(f"Could not find account {acc_no}")

                credit_loc.click()
                credit_loc.press_sequentially(credit_amt, delay=200)

                def _wait_offset_committed(offset_input_id: str) -> bool:
                    try:
                        page.wait_for_function(
                            """
                            ([inputId, expected]) => {
                                const el = document.getElementById(inputId);
                                if (!el) return false;
                                const title = (el.getAttribute('title') || '').trim();
                                const saved = (el.getAttribute('data-dyn-savedtooltip') || '').trim();
                                const val = (el.value || '').trim();
                                const valid = (el.getAttribute('aria-invalid') || '').toLowerCase() === 'false';
                                return title === expected || saved === expected || (val === expected && valid);
                            }
                            """,
                            arg=[offset_input_id, str(offset_acc).strip()],
                            timeout=1500,
                        )
                        return True
                    except PlaywrightTimeoutError:
                        return False

                page.locator("#LedgerJournalTrans_OffsetAccount_0_segmentedEntryLookup").get_by_role("button", name="Open").click()
                try:
                    page.get_by_title(offset_acc).click()
                except Exception:
                    print(f"Warning: Could not select offset account '{offset_acc}'")

                paym_mode_input = page.get_by_label("Method of payment")
                if idx > 0:
                    paym_mode_input = paym_mode_input.first
                def _select_method_of_payment(force=False):
                    try:
                        current_method = (paym_mode_input.input_value() or "").strip().lower()
                        if not force and current_method == str(pay_method).strip().lower():
                            return
                    except Exception:
                        pass
                    paym_mode_input.click()
                    paym_mode_input.press("Alt+ArrowDown")
                    # Avoid strict-mode collisions by targeting the open dropdown rows only.
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
                        # Last fallback to original approach with .first to avoid strict violation.
                        page.get_by_role("row", name=pay_method).first.get_by_label("Method of payment").first.click()

                _select_method_of_payment(force=False)

                # Fill reference date after method selection because D365 may clear it on row refresh.
                ref_date_loc.click()
                ref_date_loc.press_sequentially(ref_date, delay=200)
                # Guard: if value got cleared, set it again.
                try:
                    current_ref_val = ref_date_loc.input_value().strip()
                    if not current_ref_val:
                        ref_date_loc.click()
                        ref_date_loc.press_sequentially(ref_date, delay=200)
                except Exception:
                    pass

                # Payment reference must be filled last.
                pay_ref_loc.click()
                pay_ref_loc.press_sequentially(pay_ref, delay=200)

                # D365 may auto-reset Value date during other field interactions.
                # Re-apply it immediately before Save and verify.
                try:
                    # Fast path: fill is much faster than sequential key typing.
                    value_date_loc.fill(val_date)
                    current_value_date = value_date_loc.input_value().strip()
                    if current_value_date != val_date:
                        # Fallback path for controls that ignore direct fill.
                        value_date_loc.click()
                        value_date_loc.press("Control+A")
                        value_date_loc.press("Backspace")
                        value_date_loc.press_sequentially(val_date, delay=40)
                except Exception:
                    # Keep flow resilient; Save will still run even if readback is unavailable.
                    pass

                # As requested: after Value date re-apply, fill Method of payment again, then Save.
                _select_method_of_payment(force=True)

                page.get_by_role("button", name=" Save").click()
                print("Saved. Waiting for user to click Post...")

                page.evaluate("""
                    window.postClicked = false;

                    const oldNotice = document.getElementById('automation-post-notice');
                    if (oldNotice) oldNotice.remove();
                    const notice = document.createElement('div');
                    notice.id = 'automation-post-notice';
                    notice.textContent = 'Saved. Please click Post.';
                    Object.assign(notice.style, {
                        position: 'fixed',
                        top: '16px',
                        left: '16px',
                        zIndex: '2147483647',
                        padding: '10px 14px',
                        background: '#0b5fff',
                        color: '#fff',
                        borderRadius: '8px',
                        fontSize: '14px',
                        fontWeight: '600',
                        boxShadow: '0 6px 18px rgba(0,0,0,0.25)'
                    });
                    document.body.appendChild(notice);

                    if (window.__postHandler) {
                        document.removeEventListener('click', window.__postHandler, true);
                    }
                    window.__postHandler = function postHandler(e) {
                        const el = e.target;
                        const txt = el.innerText || el.textContent;
                        if (txt && txt.trim() === 'Post') {
                            window.postClicked = true;
                            const n = document.getElementById('automation-post-notice');
                            if (n) n.remove();
                        }
                        const btn = el.closest('button');
                        if (btn && btn.innerText && btn.innerText.trim() === 'Post') {
                            window.postClicked = true;
                            const n = document.getElementById('automation-post-notice');
                            if (n) n.remove();
                        }
                    };
                    document.addEventListener('click', window.__postHandler, true);
                """)

                # Wait without timeout until user clicks Post.
                page.wait_for_function("window.postClicked === true", timeout=0)

                print(f"User clicked Post for record {idx + 1}. Proceeding...")
                processed_records.append(record)

                # User gate before moving to next record.
                page.evaluate("""
                    () => {
                        window.automationProceedDecision = null;
                        const oldGate = document.getElementById('automation-proceed-gate');
                        if (oldGate) oldGate.remove();
                        const wrap = document.createElement('div');
                        wrap.id = 'automation-proceed-gate';
                        Object.assign(wrap.style, {
                            position: 'fixed',
                            top: '60px',
                            right: '16px',
                            zIndex: '2147483647',
                            padding: '10px',
                            background: '#ffffff',
                            border: '1px solid #d1d5db',
                            borderRadius: '8px',
                            boxShadow: '0 8px 24px rgba(0,0,0,0.2)',
                            fontFamily: 'Arial, sans-serif',
                            minWidth: '220px'
                        });
                        const text = document.createElement('div');
                        text.textContent = 'Posted. Proceed to next?';
                        text.style.marginBottom = '8px';
                        text.style.fontSize = '13px';
                        wrap.appendChild(text);
                        const p = document.createElement('button');
                        p.textContent = 'Proceed';
                        Object.assign(p.style, {
                            marginRight: '8px',
                            padding: '6px 10px',
                            background: '#16a34a',
                            color: '#fff',
                            border: 'none',
                            borderRadius: '6px',
                            cursor: 'pointer'
                        });
                        p.onclick = () => { window.automationProceedDecision = 'proceed'; wrap.remove(); };
                        const c = document.createElement('button');
                        c.textContent = 'Cancel';
                        Object.assign(c.style, {
                            padding: '6px 10px',
                            background: '#ef4444',
                            color: '#fff',
                            border: 'none',
                            borderRadius: '6px',
                            cursor: 'pointer'
                        });
                        c.onclick = () => { window.automationProceedDecision = 'cancel'; wrap.remove(); };
                        wrap.appendChild(p);
                        wrap.appendChild(c);
                        document.body.appendChild(wrap);
                    }
                """)
                page.wait_for_function("window.automationProceedDecision !== null", timeout=0)
                decision = page.evaluate("window.automationProceedDecision")
                if decision == "cancel":
                    print("User chose Cancel. Stopping further records.")
                    break
                time.sleep(1)

            print("All records processed.")
            page.evaluate("alert('All selected records have been processed.')")
            page.get_by_text("List General Payment fee Bank").click()
            page.wait_for_timeout(5000)

            # ---- Extract all voucher IDs ----
            voucher_values = []
            max_retries = 20
            for _ in range(max_retries):
                voucher_values = page.evaluate("""
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
                """)
                if voucher_values:
                    break
                page.wait_for_timeout(2000)
            print(voucher_values)
            ok, patch_msg = _bulk_update_receipts(processed_records, voucher_values)
            print(f"bulkUpdateReceipt status: {'OK' if ok else 'SKIP/FAIL'}")
            print(patch_msg)

            # Final user gate: close automation only after explicit Complete click.
            page.evaluate("""
                () => {
                    window.automationCompleteClicked = false;
                    const oldBtn = document.getElementById('automation-complete-btn');
                    if (oldBtn) oldBtn.remove();
                    const btn = document.createElement('button');
                    btn.id = 'automation-complete-btn';
                    btn.textContent = 'Complete';
                    Object.assign(btn.style, {
                        position: 'fixed',
                        top: '16px',
                        right: '16px',
                        zIndex: '2147483647',
                        padding: '10px 14px',
                        background: '#16a34a',
                        color: '#fff',
                        border: 'none',
                        borderRadius: '8px',
                        fontSize: '14px',
                        fontWeight: '700',
                        cursor: 'pointer',
                        boxShadow: '0 6px 18px rgba(0,0,0,0.25)'
                    });
                    btn.addEventListener('click', () => {
                        window.automationCompleteClicked = true;
                        btn.textContent = 'Completed';
                        btn.disabled = true;
                        btn.style.opacity = '0.75';
                    });
                    document.body.appendChild(btn);
                }
            """)
            print("All steps done. Waiting for user to click Complete...")
            page.wait_for_function("window.automationCompleteClicked === true", timeout=0)

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
        return

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

            # Wait for D365 blocking overlay to disappear (just in case)
            try:
                page.locator("#ShellBlockingDiv").wait_for(state="hidden", timeout=10000)
            except PlaywrightTimeoutError:
                print("Overlay still visible after 10s during login flow.")

            page.close()
            _persist_storage_state(context)
            print(f"Login session saved at: {CONFIG['auth_json_path']}")
        finally:
            if page is not None and not page.is_closed():
                page.close()
            if context is not None:
                context.close()
            if browser is not None:
                browser.close()

if __name__ == "__main__":
    test_final8()
