# Troubleshooting — especially in enterprise tenants

Getting an Office.js add-in running on a corporate Microsoft 365 tenant can involve several
independent layers of policy and caching, each with its own propagation delay. This guide is
the distilled experience of getting CCG running on a locked-down tenant. If you're stuck,
odds are your symptom is below.

**The golden rule: change ONE setting, then wait (up to hours). Microsoft's config plumbing
is eventually consistent across several services — flipping more switches mid-propagation
just starts new waves and makes the state impossible to reason about.**

## The add-in doesn't appear in Excel at all

1. **You're looking in the wrong dialog.** The Developer tab's "Excel Add-ins" button and
   Tools → Excel Add-ins open the *legacy* XLA/COM dialog (Analysis ToolPak, Solver,
   Browse…). Web add-ins never appear there. Look in **Home → Add-ins** (grid button) or
   **Insert → Add-ins → My Add-ins → Developer Add-ins**.
2. **Excel wasn't fully restarted.** The sideload folder is scanned at launch. ⌘Q, relaunch.
3. **Your workbook is `.xls` (Compatibility Mode).** Office.js add-ins are disabled entirely
   in legacy-format documents. Save as `.xlsx`/`.xlsm` first — check the title bar.
4. **Manifest not sideloaded / dev server down.** `npm run sideload -w @ccg/addin` copies the
   manifest to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`;
   `npm run dev -w @ccg/addin` must be serving — verify `https://localhost:3000/taskpane.html`
   loads in a browser without a certificate warning (`npm run certs` once, first).
5. **Stale add-in cache.** `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/Microsoft/Office/16.0/Wef`
   then relaunch.

## "The add-ins have been disabled. Please contact your IT administrator."

The whole Add-ins pane is dead. This is tenant policy, and there are **two separate places**
it can come from — check both:

1. **Microsoft 365 admin center → Settings → Org settings → Services → "User owned apps and
   services"** → "Let users access the Office Store" must be ticked.
2. **Cloud Policy** (config.office.com → Customization → Policy Management): the four
   **connected experiences** policies. The add-ins platform counts as a connected experience —
   if "Allow the use of connected experiences in Office" or "…additional optional connected
   experiences" is *Disabled*, the pane dies with this exact message regardless of the org
   setting above. In policy exports these show as registry value `2` = Disabled, `1` = Enabled.
   Privacy-hardened tenants very commonly have these disabled.

Side effect worth knowing: while these policies block you, the admin center's
**Integrated apps → Upload custom apps** wizard also greys out the "Office add-in" app type —
if you can only pick "Teams app" there, this is why.

Propagation after fixing: minutes to hours (occasionally the documented "up to 24h").
Signing out of your Microsoft account inside Excel and back in forces a policy refetch.

## Modern functions return #NAME? (LAMBDA, LET, SEQUENCE, TEXTSPLIT)

CCG's "Create LAMBDA name" produces names that need the LAMBDA function. If calling them —
or typing `=LAMBDA(x, x*2)(21)` raw — gives `#NAME?`, walk this ladder:

1. **Fresh blank workbook, same formula.** Still broken → app-level, keep going.
2. **Feature-era ladder:** `=SEQUENCE(3)` (2020), `=LET(a, 2, a*3)` (2021),
   `=TEXTSPLIT("a,b", ",")` (2022+). *All* failing on a current build means the app is
   running with an old feature set — an entitlement/config problem, not missing updates.
3. **Excel → About:** check the version is current and the licence line says
   "Microsoft 365 Subscription". A perpetual licence (2019/2021/LTSC) never gets LAMBDA.
4. **Check the licence is actually assigned:** admin center → Users → your user →
   Licences and apps → you need a plan that includes desktop apps (Microsoft 365 Apps for
   business/enterprise, Business Standard/Premium, E3/E5).
5. **The double restart.** Office fetches feature/flight configuration on one boot and
   *applies it on the next*. After any policy or licence change: sign out, quit ALL Office
   apps, relaunch, sign in, quit again, relaunch. (If the machine sat behind a
   connected-experiences block — see above — its feature-gate cache may never have been
   populated at all; this is the fix once the policy is lifted.)
6. Still stuck → Microsoft's licence-removal tool for Mac, then relaunch, sign in, activate
   fresh, and double-restart once more.

## Things that flip back and forth ("it worked an hour ago!")

Different capabilities are served by different backend services with independent caches:
the add-ins pane, feature flights, and licence entitlements can each be hours apart in
seeing your change. A restart mid-propagation can genuinely pick up the *new* state of one
service and the *old* state of another (e.g. LAMBDA starts working while the add-ins pane
re-disables). Don't chase it — verify the server-side setting is correct (export the policy
CSV and read the values), then stop touching things and re-test every hour or two.

## Useful diagnostics

- Export your Cloud Policy as CSV (config.office.com) — value `1` = enabled, `2` = disabled
  for the connected-experience policies. Diff exports before/after changes.
- Validate the manifest: `./node_modules/.bin/office-addin-manifest validate packages/addin/manifest.xml`
- Task pane webview console: right-click inside the pane → Inspect Element.
- The generated LAMBDA names live in the workbook (Formulas → Name Manager; stored in
  `xl/workbook.xml` → `<definedNames>`), so they keep working even when the add-in pane is
  blocked — only *generating new ones* needs the pane.
