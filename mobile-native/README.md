# Spice eTrade — native Android wrapper (APK)

A **thin** native Android app (Capacitor) that loads the existing mobile web app
in a WebView and adds the one thing a browser can't do: **direct Bluetooth
Classic (SPP) ESC/POS printing**. That lets lot receipts print straight to a
basic thermal printer (HOP-HL58, MPT-II, and similar 58mm rolls) — **without
RawBT, and therefore without RawBT's free-version watermark.**

Nothing about the web app changes: the WebView loads your live
`https://<host>/mobile`, so app updates are automatic (no APK rebuild needed for
normal app changes). The wrapper only supplies the Bluetooth bridge
(`cordova-plugin-bluetooth-serial`, exposed to the page as
`window.bluetoothSerial`). The web app detects that bridge and adds a
**"Bluetooth (app)"** print method (see `public-mobile/app.html`).

> This scaffold is ready to build, but it must be **built on a machine with
> Android tooling** (or in CI — see `CLOUD-BUILD.md`, which needs no local
> Android install at all).

## ⚠️ Before the first build — set your host

Edit `capacitor.config.json` → `server.url` and replace the placeholder with
your real hosted mobile URL:

```json
"url": "https://REPLACE-ME.up.railway.app/mobile"
```

This is the only required edit. The app id is already set to
`app.vstl.spiceetrade` — **don't change it after the first release**, or phones
will treat the next build as a different app and refuse to update in place.

## Prerequisites (on the build machine)

- **Node.js** 18+
- **Android Studio** (latest) with an Android SDK installed
- **JDK 17** (bundled with recent Android Studio)

Prefer not to install any of that? Use the GitHub Actions cloud build instead —
see `CLOUD-BUILD.md`.

## Build steps

```bash
cd mobile-native

# 1. Install Capacitor + the Bluetooth-serial plugin
npm install

# 2. Point the app at YOUR hosted mobile URL (see the warning above)
#    Edit capacitor.config.json → server.url

# 3. Create the native Android project and pull in the plugin
npx cap add android
npx cap sync

# 4. Add Bluetooth permissions (see next section), then open in Android Studio
npx cap open android
```

In Android Studio: **Build → Build Bundle(s)/APK(s) → Build APK(s)**, then copy
the generated `app-debug.apk` to the phone and install it (allow "install from
unknown sources"). For distribution, create a signed release build instead
(below).

The `android/` directory is generated, not committed — `npx cap add android`
and `build-release.sh` create it from `capacitor.config.json`, so the package
name always follows the appId.

## Android permissions (required)

Open `android/app/src/main/AndroidManifest.xml` and make sure these are present
inside `<manifest>` (the plugin adds the legacy ones; add the Android-12+ ones
if missing — the CI workflow does this automatically):

```xml
<uses-permission android:name="android.permission.BLUETOOTH" android:maxSdkVersion="30" />
<uses-permission android:name="android.permission.BLUETOOTH_ADMIN" android:maxSdkVersion="30" />
<!-- Android 12 (API 31)+ -->
<uses-permission android:name="android.permission.BLUETOOTH_CONNECT" />
<uses-permission android:name="android.permission.BLUETOOTH_SCAN" />
```

On Android 12+ the app must be granted **Nearby devices / Bluetooth** permission
at runtime — accept the prompt the first time you tap **Connect Bluetooth
printer**. If no prompt appears, enable it manually in Android **Settings → Apps
→ Spice eTrade → Permissions → Nearby devices**.

## Using it (on the phone)

1. **Pair the printer first** in Android **Settings → Bluetooth** (Classic SPP
   printers must be paired at the OS level before an app can open them).
2. Open the installed **Spice eTrade** app and log in.
3. Menu (⋮) → **Print method** → cycle to **Bluetooth (app)**.
4. Menu → **Connect Bluetooth printer** → pick the printer from the in-app list
   (auto-selected if it's the only paired printer; remembered afterwards).
5. Menu → **🧪 Test print** → a small "PRINTER OK" slip should come out. Once
   that works, everything else does.
6. Print any lot / seller / batch — the ESC/POS goes straight to the printer.

The web app handles the fiddly parts so they don't come back as problems:
ESC/POS is sent in **small chunks** (big "all sellers" jobs don't get
truncated), the link **auto-reconnects** if it dropped while idle, and the
printer is picked from a **proper in-app list** (not a typed prompt) and
remembered.

## Receipt width

The thermal slip is rendered server-side at **32 characters per line (58mm)** by
default. The ESC/POS endpoints accept `?width=42` or `?width=48` for wider
rolls. Note this is separate from the PDF path's `lot_receipt_width_mm` setting
(Settings → Lot Entry Defaults), which only affects PDF printing.

If all optional columns (Smp / Gross / Mst%) are enabled they will not fit in 32
characters. Rather than clip a weight — which would print a **wrong number** —
the renderer drops the lowest-value optional column (Mst%, then Gross, then
Smp). Lot# / Bags / Net always print at full precision. Use a wider roll if you
need every column.

## Distributable signed APK

A signed release APK installs cleanly and updates without the "unknown app"
friction of debug builds. One-time setup, then one command per build.

```bash
cd mobile-native

# 1. Generate a signing key (keep the keystore + passwords safe; you need the
#    SAME key for every future update).
keytool -genkeypair -v -keystore android/spice-release.keystore \
  -alias spice -keyalg RSA -keysize 2048 -validity 10000

# 2. Tell Gradle about it
cp signing/keystore.properties.example android/keystore.properties
#    edit android/keystore.properties → set the two passwords you just chose

# 3. Wire the signing config in: add this as the LAST line of
#    android/app/build.gradle
#        apply from: "../../signing/signing.gradle"

# 4. Build
./build-release.sh
#    → mobile-native/android/app/build/outputs/apk/release/app-release.apk
```

`build-release.sh` runs `npm install`, creates the android project if needed,
`cap sync`, and `gradlew assembleRelease`. If you skip the signing setup it
still builds, but the APK is unsigned (fine for a quick test, not for handing
out). Copy the APK to the phone and install it.

## How it fits together

- `capacitor.config.json` → `server.url` = your live `/mobile` (WebView target).
- `cordova-plugin-bluetooth-serial` → injects `window.bluetoothSerial`
  (`list` / `connect` / `isConnected` / `write`).
- `public-mobile/app.html` →
  - `hasNativeBt()` detects the bridge,
  - adds the **`native`** print method (and prefers it inside the wrapper),
  - `nativeBtPrint()` fetches the `…​.escpos` bytes and `write()`s them in chunks,
  - `connectPrinter()` (the menu item) picks/remembers the printer.
- `mobile-bridge.js` → serves the ESC/POS bytes:
  - `GET /api/lots/:id/receipt.escpos`
  - `GET /api/lots/print-seller.escpos`
  - `GET /api/lots/print-batch.escpos`
  - `GET /api/lots/print-all-sellers-escpos/:auctionId`

## Troubleshooting

- **"Bluetooth (app)" method not offered / prints open in browser** → the
  WebView didn't get `window.bluetoothSerial`. Re-run `npx cap sync`, rebuild.
  If a remote `server.url` doesn't surface the bridge on your Capacitor version,
  use the **local-bundle fallback** below.
- **"No paired printers"** → pair the printer in Android Bluetooth settings
  first, and grant the app "Nearby devices".
- **Connects but prints nothing/garbage** → the `write()`/chunking may need a
  small tweak for that printer model; the chunk size lives in
  `_btsWriteChunked()` in `public-mobile/app.html`.
- **Header prints as text, no logo** → expected. The ESC/POS header only rasters
  **JPEG**; this build's bundled logos are PNG. Drop a letterhead banner at
  `public/receipt-header.jpg` to print an image header.

### Local-bundle fallback (only if the remote bridge doesn't inject)

Instead of `server.url`, bundle the web app locally and point it at the API host:

1. Copy the app in: `cp -r ../public-mobile/* www/` (replace the placeholder
   `www/index.html`).
2. In `capacitor.config.json`, **remove the `server` block**.
3. Tell the app where the API lives by injecting a global before it loads — add
   a tiny inline script in `www/app.html`'s `<head>`:
   ```html
   <script>window.SPICE_API_BASE='https://<your-real-host>';</script>
   ```
   (`app.html` honors `window.SPICE_API_BASE`.)
4. `npx cap sync` and rebuild. Downside: you must rebuild the APK whenever the
   web app changes.
