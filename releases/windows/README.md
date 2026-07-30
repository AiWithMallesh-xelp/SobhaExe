# Windows build outputs

GitHub Actions builds `sobha.exe` and `sobha-setup.msi`, then stores the files in this folder.

## Files

| File | Contents |
|---|---|
| `sobha-setup.msi` | Windows installer (Start Menu shortcut, uninstaller) |
| `sobha-windows-portable.zip` | Full portable package: exe, config, logo, readme, pw-browsers |
| `sobha-app-only.zip` | Small package: exe, config, readme, logo |
| `sobha-pw-browsers.zip` | Playwright Chromium browsers only |
| `build-info.json` | Build metadata from CI |

## Test on Windows (portable ZIP)

1. Download `sobha-windows-portable.zip`.
2. Extract the full folder (do not run from inside the ZIP).
3. Double-click `sobha.exe`.

## Test on Windows (MSI installer)

1. Download `sobha-setup.msi`.
2. Double-click the MSI and follow the install wizard.
3. Open **Sobha Reconciliation** from the Start Menu.

If SmartScreen appears: click **More info** → **Run anyway**.

## Build

Push to the `shobhaexe-msi` branch or run **Build Windows EXE and MSI** from GitHub Actions.
