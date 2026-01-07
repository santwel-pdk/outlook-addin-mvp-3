# PLANNING.md - outlook-addin-mvp-3

## Project Overview

**Project Name:** outlook-addin-mvp-3
**Type:** Outlook Add-in (Task Pane)
**Framework:** React 18 + TypeScript
**UI Library:** Fluent UI React v9
**Target Platforms:** Outlook for Windows 11 (WebView2), Outlook for macOS
**Office.js Requirement Set:** Mailbox 1.3+

### Purpose
This add-in demonstrates a production-ready Outlook Add-in architecture with:
- Modern React/TypeScript development
- WebView2 enforcement for Windows 11
- Cross-platform compatibility
- Proper Office.js API integration
- Comprehensive error handling
- Type-safe development

---

## Architecture Decisions

### Why React?
- **Component Reusability:** Modular UI components for maintainability
- **State Management:** Built-in hooks for managing Office.js async state
- **Ecosystem:** Large community and Fluent UI integration
- **Type Safety:** Excellent TypeScript support

### Why WebView2 Enforcement?
**CRITICAL for Windows 11:**
- **Modern JavaScript:** ES2020+ features without transpilation concerns
- **Performance:** Significantly faster than legacy IE11/Trident
- **Security:** Latest browser security features
- **Developer Experience:** Chrome DevTools, React DevTools
- **Consistency:** Same rendering engine as modern browsers

**Implementation:**
1. Manifest.xml includes `<Runtimes>` element
2. Extended-manifest.json specifies WebView2 runtime
3. Runtime detection in webview2Service.ts throws error for IE11/Trident

### Why Fluent UI React v9?
- **Office Native Look:** Consistent with Office UI/UX
- **Accessibility:** WCAG 2.1 compliant components
- **Theming:** Supports Office themes (light/dark)
- **Performance:** Optimized for Office Add-in scenarios

---

## Office.js API Strategy

### Requirement Sets Used
- **Mailbox 1.3:** Minimum version for WebView2 and core mailbox features
- **APIs Used:**
  - `Office.context.mailbox.item` - Current email item
  - `item.subject` - Email subject
  - `item.from` - Sender information
  - `item.dateTimeCreated` - Received date
  - `item.to` - Recipients
  - `item.body.getAsync` - Email body content

### API Patterns
**Async/Await Pattern (Preferred):**
```typescript
// Modern Promise-based approach
const subject = await getCurrentEmailSubject();
```

**Callback Pattern (Avoided):**
```typescript
// Deprecated - NOT used in this project
item.subject.getAsync((result) => { /* ... */ });
```

**Error Handling:**
- All Office.js calls wrapped in try-catch
- User-friendly error messages via errorHandler utility
- No sensitive data logged in production

---

## Cross-Platform Considerations

### Windows 11 vs macOS Differences

| Feature | Windows 11 (WebView2) | macOS (Safari WebKit) |
|---------|----------------------|----------------------|
| JavaScript Engine | V8 (Chromium) | JavaScriptCore |
| CSS Rendering | Blink | WebKit |
| Performance | Excellent | Good |
| DevTools | Edge DevTools (F12) | Safari Web Inspector |
| WebView2 Enforcement | ✅ Enforced | ❌ Not applicable |

### Platform-Specific Code
```typescript
// platform.ts utility
if (isWindows()) {
  // Windows-specific logic
}

if (isMacOS()) {
  // macOS-specific logic
}
```

**Rendering Differences:**
- Custom CSS tested on both platforms
- Flexbox/Grid work consistently
- Font rendering may differ slightly

**API Availability:**
- Some Office.js APIs have platform limitations
- Always check documentation for platform support
- Gracefully degrade features if unavailable

---

## File Structure Explanation

```
src/taskpane/
├── components/           # React UI components
│   ├── App.tsx          # Main app component (uses hooks)
│   ├── EmailInfo.tsx    # Email data display (Fluent UI Card)
│   ├── ErrorBoundary.tsx # Error recovery wrapper
│   └── Header.tsx       # Taskpane header (original)
│
├── hooks/               # Custom React hooks
│   ├── useOfficeContext.ts   # Office.js initialization state
│   └── useMailboxItem.ts     # Current email data fetching
│
├── services/            # Office.js API wrappers
│   ├── officeService.ts      # Office.js initialization
│   ├── mailService.ts        # Email operations (read)
│   └── webview2Service.ts    # WebView2 detection/enforcement
│
├── utils/               # Helper utilities
│   ├── errorHandler.ts       # Centralized error handling
│   └── platform.ts           # Platform detection (Win/Mac)
│
├── types/               # TypeScript type definitions
│   ├── office.types.ts       # Office.js type extensions
│   └── app.types.ts          # Application-specific types
│
├── index.tsx            # React app entry point
└── taskpane.html        # HTML entry point

commands/
└── commands.ts          # Ribbon button command handlers
```

### Design Principles
1. **Separation of Concerns:** Services separate from UI logic
2. **Single Responsibility:** Each module has one clear purpose
3. **Type Safety:** Explicit TypeScript types throughout
4. **Error Resilience:** Graceful error handling at every layer
5. **Testability:** Pure functions, mockable dependencies

---

## Security Considerations

### Data Privacy
- **No Sensitive Logging:** Email subjects, bodies, and addresses NOT logged in production
- **Error Messages:** User-friendly messages without exposing internals
- **Development vs Production:** Detailed errors only in development mode

### API Key Management
- Environment variables for secrets (`.env` file, gitignored)
- `REACT_APP_*` prefix for Create React App variables
- Never commit `.env` file to git

### Content Security Policy (CSP)
- Manifest specifies allowed domains in `<AppDomains>`
- All external resources must use HTTPS
- No inline scripts in production

---

## Performance Optimizations

### Bundle Size
- Code splitting with `React.lazy` for large components
- Tree shaking enabled in webpack production build
- Fluent UI adds ~200KB (unavoidable but worth it)
- **Target:** <500KB gzipped for main bundle

### API Call Optimization
- Parallel Promise.all for independent API calls
- Caching Office context after initialization
- Debouncing user interactions that trigger API calls

### Rendering Performance
- React.memo for expensive components
- useCallback/useMemo for stable references
- Avoid unnecessary re-renders with proper dependencies

---

## Testing Strategy

### Unit Tests (Jest + React Testing Library)
**Services:**
- webview2Service: Detects WebView2, rejects IE11
- officeService: Initializes Office.js correctly
- mailService: Returns email data, handles errors

**Components:**
- EmailInfo: Renders data, loading, error states
- ErrorBoundary: Catches errors, shows fallback

**Utilities:**
- errorHandler: Formats errors correctly
- platform: Detects Windows vs macOS

**Mocking Strategy:**
- Use `office-addin-mock` for Office.js globals
- Mock Office.context for component tests
- Test error paths with rejected promises

### Integration Tests
- Full user flow: Load add-in → Fetch email → Display data
- Test with various email types (read/compose)
- Platform-specific scenarios (Windows WebView2, macOS Safari)

### Manual Testing
- Sideload in Outlook for Windows 11
- Sideload in Outlook for macOS
- Verify WebView2 detection in browser console
- Test error recovery (disconnect network, select no email)

---

## Deployment Considerations

### Development
- `npm start` - Webpack dev server with hot reload
- HTTPS certificates required (npx office-addin-dev-certs install)
- Sideload manifest.xml for testing

### Production Build
- `npm run build` - Optimized production bundle
- Host on HTTPS server (Azure Static Web Apps, AWS S3+CloudFront, etc.)
- Update manifest.xml URLs to production domain
- Update extended-manifest.json URLs

### Distribution Options
1. **Personal Use:** Sideload manually
2. **Organization:** Microsoft 365 admin center deployment
3. **Public:** Microsoft AppSource submission

### AppSource Requirements (if applicable)
- Privacy policy URL
- Support URL
- App icon in all required sizes
- Validation testing on both Windows and Mac
- WebView2 compatibility confirmation

---

## Future Enhancements

### Short Term
- Add unit tests with ≥70% coverage
- Implement compose mode support
- Add attachment viewing/downloading

### Medium Term
- Roaming settings for user preferences
- Multi-language support (i18n)
- Dark mode support (Fluent UI dark theme)
- Email categorization features

### Long Term
- AI-powered email insights
- Integration with external CRM systems
- Calendar integration features
- Mobile support (Outlook Mobile APIs)

---

## Known Limitations

1. **WebView2 Requirement:** Windows 11 users must have Office 365 build 16.0.14326+ (auto-installed)
2. **API Availability:** Some Mailbox APIs not available on macOS (check docs)
3. **Dialog API:** Works differently on Windows vs macOS
4. **Custom Pane:** Not supported on macOS (use dialog instead)
5. **Network Dependency:** Requires internet for loading resources (unless offline-enabled)

---

## References

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://learn.microsoft.com/en-us/javascript/api/office)
- [Fluent UI React](https://react.fluentui.dev/)
- [Yeoman Generator for Office Add-ins](https://github.com/OfficeDev/generator-office)
- [WebView2 in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins)

---

**Document Version:** 1.0
**Last Updated:** 2026-01-06
**Author:** AI Agent (Claude Code)
