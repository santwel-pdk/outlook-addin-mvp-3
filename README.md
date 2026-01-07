# outlook-addin-mvp-3

A production-ready Outlook Add-in built with React, TypeScript, and Fluent UI, supporting Windows 11 (WebView2) and macOS.

## Features

- ✅ **Modern React 18 + TypeScript** - Type-safe development with latest React features
- ✅ **WebView2 Enforcement** - Ensures Edge Chromium browser on Windows 11 (mandatory)
- ✅ **Cross-Platform Support** - Windows 11 and macOS Outlook compatibility
- ✅ **Fluent UI Components** - Native Office look-and-feel
- ✅ **Email Information Display** - Shows subject, sender, date, and recipients
- ✅ **Error Handling** - Comprehensive error boundaries and user-friendly messages
- ✅ **Platform Detection** - Automatic Windows/macOS detection with browser engine info
- ✅ **Office.js Integration** - Proper initialization and API usage patterns

---

## Prerequisites

### Required Software

- **Node.js:** v16.x or v18.x ([Download](https://nodejs.org/))
- **npm:** 8.x or higher (included with Node.js)
- **Office 365/Microsoft 365:**
  - Windows 11: Office 365 build 16.0.14326+ (WebView2 included)
  - macOS: Outlook for Mac (latest version recommended)

### Development Tools (Recommended)

- **Visual Studio Code** with Office Add-in extension
- **Git** for version control

---

## Installation

### 1. Clone the Repository

```bash
cd /path/to/your/projects
git clone <repository-url> outlook-addin-mvp-3
cd outlook-addin-mvp-3
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Install HTTPS Certificates

**Required for local development** (Office Add-ins require HTTPS):

```bash
npx office-addin-dev-certs install
```

**Note:** You may need to trust the certificate:
- **Windows:** Follow prompts to add certificate to trusted root store
- **macOS:** Certificate will be added to Keychain (may require password)

### 4. Configure Environment (Optional)

```bash
cp .env.example .env
# Edit .env if needed
```

---

## Development

### Start Development Server

```bash
npm start
```

This will:
- Start webpack dev server on `https://localhost:3000`
- Enable hot module replacement (HMR)
- Open browser to taskpane.html

**Verify:**
- Navigate to https://localhost:3000/taskpane.html
- You should see the add-in UI (Office.js won't initialize in browser)

### Sideload in Outlook (Windows 11)

1. Ensure `npm start` is running
2. Open **Outlook for Windows 11**
3. Go to **Home** → **Get Add-ins** → **My Add-ins**
4. Click **Add a custom add-in** → **Add from file**
5. Select `manifest.xml` from project directory
6. Open any email in **Read mode**
7. Click the add-in button in the ribbon
8. Taskpane should open on the right side

**Debugging:**
- Right-click in taskpane → **Inspect** (or F12)
- Check Console for WebView2 detection message
- Verify User Agent contains "Edg/" (Edge Chromium)

### Sideload in Outlook (macOS)

1. Ensure `npm start` is running
2. Open **Outlook for macOS**
3. Go to **Get Add-ins** → **My Add-ins**
4. Click **Add a custom add-in** → **Add from file**
5. Select `manifest.xml` from project directory
6. Open any email in **Read mode**
7. Click the add-in button in toolbar
8. Taskpane should open

**Debugging:**
- **Safari** → **Develop** → **[Your Mac]** → **Outlook**
- Select taskpane context to inspect
- Check Console for any errors

---

## Building for Production

### Create Production Build

```bash
npm run build
```

**Output:** `dist/` directory with optimized bundles

### Verify Build

```bash
# Check bundle sizes
ls -lh dist/*.js

# Expected: Main bundle <500KB (before gzip)
```

### Deploy to Production

1. **Update Manifest URLs:**
   - Edit `manifest.xml`: Replace `https://localhost:3000` with your production URL
   - Edit `extended-manifest.json`: Update URLs to production domain

2. **Host Files:**
   - Upload `dist/` contents to web server (HTTPS required)
   - Upload `extended-manifest.json` to same domain
   - Ensure all assets accessible via HTTPS

3. **Distribution Options:**
   - **Personal:** Share updated `manifest.xml`
   - **Organization:** Deploy via Microsoft 365 admin center
   - **Public:** Submit to Microsoft AppSource

---

## Testing

### Run Linting

```bash
npm run lint
```

Fix auto-fixable issues:

```bash
npm run lint:fix
```

### Run Type Checking

```bash
npx tsc --noEmit
```

### Run Unit Tests (When Implemented)

```bash
npm test

# With coverage
npm test -- --coverage
```

**Expected Coverage:** ≥70% for all metrics (statements, branches, functions, lines)

---

## Project Structure

```
outlook-addin-mvp-3/
├── src/
│   ├── taskpane/
│   │   ├── components/           # React UI components
│   │   │   ├── App.tsx          # Main app (uses hooks)
│   │   │   ├── EmailInfo.tsx    # Email data display
│   │   │   ├── ErrorBoundary.tsx # Error recovery
│   │   │   └── Header.tsx       # Taskpane header
│   │   ├── hooks/               # Custom React hooks
│   │   │   ├── useOfficeContext.ts   # Office.js init state
│   │   │   └── useMailboxItem.ts     # Email data fetching
│   │   ├── services/            # Office.js API wrappers
│   │   │   ├── officeService.ts      # Office.js init
│   │   │   ├── mailService.ts        # Email operations
│   │   │   └── webview2Service.ts    # WebView2 detection
│   │   ├── utils/               # Helper utilities
│   │   │   ├── errorHandler.ts       # Error handling
│   │   │   └── platform.ts           # Platform detection
│   │   ├── types/               # TypeScript types
│   │   │   ├── office.types.ts       # Office.js types
│   │   │   └── app.types.ts          # App types
│   │   ├── index.tsx            # React entry point
│   │   └── taskpane.html        # HTML entry point
│   └── commands/
│       └── commands.ts          # Ribbon commands
├── assets/                      # Icons and resources
├── manifest.xml                 # Add-in manifest
├── extended-manifest.json       # WebView2 runtime config
├── package.json                 # Dependencies and scripts
├── tsconfig.json                # TypeScript config
├── webpack.config.js            # Webpack bundler config
├── .env.example                 # Environment template
├── .gitignore                   # Git ignore patterns
├── PLANNING.md                  # Architecture docs
├── TASK.md                      # Task tracking
└── README.md                    # This file
```

---

## Troubleshooting

### Certificate Errors (HTTPS)

**Symptom:** "Your connection is not private" or SSL errors

**Solution:**
```bash
npx office-addin-dev-certs install
```

If still failing, clear browser cache and restart browser.

---

### WebView2 Not Detected (Windows)

**Symptom:** Warning in console: "WebView2 not detected"

**Cause:** Office version doesn't support WebView2 or using IE11

**Solution:**
1. Check Office version: **File** → **Account** → **About Outlook**
2. Ensure build ≥ 16.0.14326 (required for WebView2)
3. Update Office 365/Microsoft 365 to latest version
4. Restart Outlook after update

---

### Add-in Not Loading

**Symptom:** Ribbon button doesn't appear or taskpane blank

**Debugging Steps:**
1. Verify `npm start` is running on https://localhost:3000
2. Check manifest.xml URLs are correct (localhost:3000 for dev)
3. Clear Office cache:
   - **Windows:** Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - **macOS:** `~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/`
4. Restart Outlook
5. Re-sideload manifest

---

### Office.js Initialization Errors

**Symptom:** Console shows "Office is not defined" or "Office.context is null"

**Solution:**
- Ensure Office.js script is loaded (check `<head>` in taskpane.html)
- Wait for `Office.onReady()` before accessing APIs
- Check you're running in actual Outlook, not just browser

---

### Platform-Specific Issues

**macOS Safari Rendering Differences:**
- Some CSS properties may render differently than WebView2
- Test custom styles on both platforms
- Use Fluent UI components for consistency

**Windows WebView2 Issues:**
- Ensure Windows 11 is up to date
- WebView2 runtime should auto-install with Office updates
- Manual install: [Download WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/)

---

## Code Guidelines

- **File Size Limit:** Maximum 400 lines per file (enforced)
- **TypeScript:** Strict mode enabled, no implicit any
- **Imports:** Use absolute paths with aliases (configured in tsconfig.json)
- **Error Handling:** All Office.js calls wrapped in try-catch
- **Logging:** No sensitive data (email content, addresses) in production logs
- **Components:** Functional components with hooks only (no class components)

---

## Resources

### Documentation
- [Office Add-ins Docs](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/office)
- [Fluent UI React](https://react.fluentui.dev/)
- [PLANNING.md](./PLANNING.md) - Architecture decisions and patterns
- [TASK.md](./TASK.md) - Task tracking and history

### Tutorials
- [Outlook Add-in Quickstart](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart-yo)
- [Sideloading on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
- [Sideloading on macOS](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

### Support
- GitHub Issues: <repository-issues-url>
- Stack Overflow: [office-js tag](https://stackoverflow.com/questions/tagged/office-js)

---

## License

MIT

---

## Contributing

1. Read [PLANNING.md](./PLANNING.md) to understand architecture
2. Check [TASK.md](./TASK.md) for open tasks
3. Follow code guidelines above
4. Write tests for new features
5. Update documentation as needed

---

**Built with ❤️ using Office Add-ins platform, React, TypeScript, and Fluent UI**
