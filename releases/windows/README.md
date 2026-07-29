# Windows portable build

GitHub Actions builds `sobha.exe` and stores the ZIP files in this folder.

## Files

| File | Contents |
|---|---|
| `sobha-windows-portable.zip` | Full test package: exe, config, readme, pw-browsers |
| `sobha-app-only.zip` | Small package: exe, config, readme |
| `sobha-pw-browsers.zip` | Playwright Chromium browsers only |

## Test on Windows

1. Download `sobha-windows-portable.zip`.
2. Extract the full folder (do not run from inside the ZIP).
3. Double-click `sobha.exe`.

If you downloaded the split ZIPs instead, extract **both** into the same folder.

## Build

Push to the `shobhaexe` branch or run **Build Windows Portable EXE** from GitHub Actions.
