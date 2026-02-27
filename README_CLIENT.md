# Sobha Reconciliation (Portable)

## How to run

1. Extract the ZIP fully — **do not run from inside the ZIP**.
2. Open the extracted folder.
3. Double-click **sobha.exe**.

> If Windows SmartScreen appears: click **More info** → **Run anyway**.

---

## First-time login

1. Click the **Login** button in the app.
2. A browser will open — sign in to your Microsoft / D365 account.
3. Once you are on the D365 page, click the blue **"Login Success"** button in the browser.
4. A file called **auth.json** is created in the same folder as `sobha.exe`.
5. Future runs will reuse this session sobhamatically.

---

## Important

Keep all of these in the **same folder** at all times:

| File / Folder | Required? | Purpose |
|---|---|---|
| `sobha.exe` | ✅ Yes | The application |
| `config.json` | ✅ Yes | App configuration (D365 URL, journal name, etc.) |
| `pw-browsers/` | ✅ Yes | Browser engine — **do not delete or rename** |
| `auth.json` | Created after first login | Stores your login session |

---

## Troubleshooting

| Problem | Fix |
|---|---|
| App won't start | Make sure you extracted the full ZIP (not running inside it) |
| `pw-browsers` error | Do not rename or move the `pw-browsers` folder |
| Browser won't open | Contact support |
| Session expired | Click Login again and sign in |

---

## Support

If you encounter an error, please share:
- A screenshot of the error message
- The folder structure (a screenshot of the folder contents)
