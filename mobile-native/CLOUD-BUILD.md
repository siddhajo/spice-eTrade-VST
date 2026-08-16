# Cloud build — never open Android Studio again

The APK is built + signed automatically by **GitHub Actions**
(`.github/workflows/android-apk.yml`) and published into your own app, so
operators install it from a link. You do a **one-time** setup (~5 min), then it's
a button click whenever you need a new APK — which is rare, because normal app
changes just deploy to Railway as usual and phones pick them up automatically.

---

## One-time setup

### 1. Create a signing key (one terminal command, once)
This key signs every build. **Keep the file + passwords forever** — the same key
is required for every future update.

On your Mac, in Terminal:

```bash
# If `keytool` isn't found, use the one bundled with Android Studio:
#   "/Applications/Android Studio.app/Contents/jbr/Contents/Home/bin/keytool"
keytool -genkeypair -v \
  -keystore spice-release.keystore \
  -alias spice -keyalg RSA -keysize 2048 -validity 10000 \
  -storepass "ChooseAPassword" -keypass "ChooseAPassword" \
  -dname "CN=VSTL Spices, O=VSTL Spices, C=IN"

# Copy its base64 to the clipboard (you'll paste it into a GitHub secret):
base64 -i spice-release.keystore | pbcopy
```

Back up `spice-release.keystore` and the password somewhere safe (password
manager / secure drive). If you lose it, future updates can't install over the
old app.

### 2. Add 4 repo secrets on GitHub
Repo → **Settings → Secrets and variables → Actions → New repository secret**.
Add these four:

| Secret name | Value |
|---|---|
| `ANDROID_KEYSTORE_BASE64` | paste from the clipboard (step 1's base64) |
| `ANDROID_KEYSTORE_PASSWORD` | the password you chose |
| `ANDROID_KEY_ALIAS` | `spice` |
| `ANDROID_KEY_PASSWORD` | the password you chose |

### 3. Push these files to GitHub
Commit + push the new files to your repo (they're already in your working copy):
`.github/workflows/android-apk.yml`, `mobile-native/**`, `public-mobile/install.html`.
Also make sure `mobile-native/capacitor.config.json` → `server.url` is your real
host (`https://REPLACE-ME.up.railway.app/mobile`).

---

## Building the APK (the repeatable part)

1. GitHub repo → **Actions** tab → **Build Android APK** → **Run workflow**.
2. Wait ~5–10 min. It installs, syncs, signs, builds, and **commits the APK into
   `public-mobile/downloads/spice-etrade.apk`** — which triggers a Railway
   deploy.
3. That's it. The APK is now live at:
   **`https://REPLACE-ME.up.railway.app/mobile/install.html`**

(The APK is also attached to the workflow run as a downloadable artifact.)

---

## Rolling out to operators

Send the printing operators this link (open it **on the phone**):

> `https://REPLACE-ME.up.railway.app/mobile/install.html`

The page walks them through: Download → Install → pair the printer → set
**Print method → Bluetooth (app)** → Connect → **🧪 Test print**. Done.

---

## Day-to-day

- **App changes (features/fixes):** deploy to Railway as normal → every phone
  (APK and browser) updates automatically. **No APK rebuild.**
- **Only if the Bluetooth/native layer changes:** click **Run workflow** again to
  publish a fresh APK; operators reopen the install link and reinstall.

## First-run notes
- The very first workflow run creates `public-mobile/downloads/`, so until it
  finishes, the Download button on the install page will 404 — that's expected.
- If the first build errors, open the failed step's log in the Actions tab and
  send it over; CI Android builds occasionally need a one-line tweak (SDK level,
  a permission, etc.).
