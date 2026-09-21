# Gemini AI Agent Learnings & Rules (GEMINI.md)

This document contains accumulated technical learnings, architecture blueprints, and engineering governance rules derived from real-world development sessions and archived project tasks.

---

## 1. OpenXML & OOXML Document Generation Governance
When generating or modifying Office OpenXML (`.docx`, `.xlsx`, `.pptx`) using C# / `DocumentFormat.OpenXml`:

- **Never Hardcode Relationship IDs (`rId`)**:
  DO NOT hardcode relationship IDs (e.g., `rId1`, `rId2`) in generated XML. ALWAYS use dynamic relationship managers (`package.MainDocumentPart.AddExternalRelationship(...)`) to register external images, hyperlinks, footnotes, or custom parts in `.rels`.
- **Enforce Root Namespaces**:
  Ensure all required XML namespace declarations (`w:`, `a:`, `wp:`, `pic:`, `m:`, `wpg:`, `wps:`, `w15:`) are declared on root elements. Missing namespaces cause Word "File Corrupted" errors upon opening.
- **Memory Overhead & SAX Writing**:
  Avoid loading full `document.xml` trees into DOM memory (`XmlDocument`/`XDocument`) for large documents. Use `OpenXmlWriter` (SAX-style streaming) for $O(1)$ memory footprint.
- **Native Word Element Reality vs AI Hallucinations**:
  - **Tabs (`:::tabs`)**: Word has no native `w15:tabSet` element. Implement using Heading outline levels (`w:outlineLvl`) or Content Controls (`w:sdt`) linked to Custom XML parts + VBA macros.
  - **Online Video (`:::embed`)**: Use `<w15:webVideoPr>` extension inside `<wp:docPr>` containing an `<a:blip>` thumbnail image.
  - **DrawingML Charts (`:::chart`)**: Requires a `ChartPart` backed by an embedded Excel spreadsheet (`EmbeddedPackagePart`) linked via `<c:f>` formulas.
  - **PivotTables / OLE Objects (`:::datagrid`)**: Embed binary `.xlsx` as `EmbeddedObjectPart` with an EMF/PNG preview image linked via `<w:object>`.
  - **Multi-Column Sections (`:::columns`)**: Enclose column blocks in Continuous Section Breaks (`<w:type w:val="continuous"/>` inside `<w:sectPr>`).
  - **Footnotes**: Store in `word/footnotes.xml` and insert `<w:footnoteReference>` inline anchors in document body.
  - **Math & OMML Whitespace**: Preserve literal whitespace inside OMML text blocks using `<m:t xml:space="preserve">`.
  - **Complex HTML Tables**: Translate `colspan` and `rowspan` into native `<w:gridSpan>` and `<w:vMerge>`.
  - **Collapsible Sections**: Map to native Word collapsible headings using `<w15:collapsed w:val="true"/>`.

---

## 2. Security & Anti-Corruption Practices
- **XSS Prevention in Code Fences / Diagrams**:
  When transforming raw code fences (e.g., `<pre><code class="language-mermaid">`) into DOM rendering containers, DO NOT `HtmlDecode` text before insertion. Browsers automatically decode `.textContent`, so pre-decoding exposes XSS execution vectors.
- **Inline Tag Processing**:
  When stripping or transforming inline HTML tags (e.g., `<u>`, `<span>`), preserve adjacent text nodes and trailing whitespace to prevent string truncation bugs.

---

## 3. Markdown AST & Parser State Machines
- **Pipeline Configuration**:
  Always verify that the `MarkdownPipelineBuilder` has explicitly enabled required extension modules (e.g., `.UseDefinitionLists()`) before assuming AST nodes will be present.
- **Parser Leaked Borders & Blockquote Traps**:
  Ensure table parsers handle nested blockquote contexts (`>`) correctly and strip table structural borders (e.g., Pandoc `+------+` lines) rather than leaking them into text runs.

---

## 4. UI/UX Architecture & Dynamic Layouts
- **In-Place Canvas Editor Coordinates**:
  Positioning floating text inputs or overlay controls on interactive canvases must calculate bounds relative to their direct `Canvas` parent wrapper rather than outer Grids or Window containers.
- **Explicit Action Affordances**:
  Provide dual commit mechanisms for inline canvas editors: explicit Commit (✔️) / Cancel (❌) action buttons alongside keyboard accelerators (`Enter` / `Escape`).
- **Dynamic Vector Shape Converters**:
  Use WPF/Avalonia value converters (`ShapeToVisibilityConverter`) on item templates to dynamically switch icon primitives (Actor, Database cylinder, Decision diamond, Root circle) based on model metadata.
- **Workspace Consolidation**:
  Consolidate crowded toolbars into smart dropdown clusters (`Text Style`, `Lists ▼`, `+ Insert ▼`) and utilize custom title bar regions (`AppTitleBar`) to reclaim vertical canvas height.

---

## 5. Multi-Agent & Subagent Orchestration
- **Specialized Role Decomposition**:
  For complex features or refactors, break tasks into specialized subagent personas (e.g., OpenXML Researcher, OpenXML Verifier, DocX Debugger, Layout Debugger, Auditor, Reviewer, Challenger).
- **Empirical Verification Gate**:
  Never declare victory based solely on code compilation. Perform end-to-end verification tests, run automated test suites, and inspect raw output file structures (XML, PDF, DOCX) to confirm fidelity.

---

## 6. Markdown Engine Governance (the syntax contract)
Before modifying the markdown pipeline, the rendering engine, or ANY markdown wrapper syntax,
read **`docs/MD_ENGINE_GOVERNANCE.md`** — it is the architecture map AND the syntax contract:

- **Two pipelines, one contract**: the DOCX/OpenXML path (normalizers → `AdvancedFeaturePipeline`
  feature markers → Markdig AST → native `w:p`/`w:r`/`w:drawing` + native `.glox` SmartArt) and
  the HTML preview path (normalize → Markdig + Mathematics → targeted `HtmlSanitizer` → trusted
  post-inject of mermaid / `:::smartart` SVG / KaTeX / lens / portal / fit-width). A syntax
  change must land in **both** paths or the preview and the exported DOCX disagree.
- **The wrapper catalog**: `:::smartart` / `:::workflow` / `:::tabs` (`=== "Tab"`) / `:::chart` /
  `:::columns` / `:::timeline` / `:::canvas` / `:::shapes`, `$…$` / `$$…$$` / `\(…\)` / `\[…\]`
  math (KaTeX + mhchem), `> [!NOTE]` callouts, special code fences, task lists, footnotes, and
  the INTERNAL `<!-- MARKSMITH_FEATURE:id -->` markers / `<!-- … -->` placeholder comments.
  `{{token}}` and `$$"""…"""` inside `MarkdownHtmlService` are C# raw-interpolated-string
  templating — **not** markdown syntax. Do not invent undocumented wrappers; extend the catalog
  in the same change.
- **Hard rules**: never fork/modify Markdig directly (pre-process or dispatch); output stays
  native OpenXML (raster only as last-resort); after `HtmlSanitizer.Apply` only trusted
  generated markup may be injected (comments survive the targeted sanitizer and are the safe
  placeholder vehicle); ambiguous constructs go through `AmbiguityDetector`/`AmbiguityResolverDialog`
  honoring `AppSettings.AmbiguityPreferences`.
- **Snippet shapes**: `InsertSnippetBuilder` is the single source of truth for what the UI
  inserts — keep it in sync whenever a syntax changes.

---

## 7. AI-Executable 3-Block Cycle + Execution Phase Protocol

### System Directive
“Run a 3‑block cycle. Each block generates 2 new ideas and refines the ideas from previous blocks. After Block 3, refine its 2 ideas again and then execute all 6 refined ideas into code. The execution phase is the 4th stage but not a generation block.”

### Cadence & Structure (15-Minute Blocks / 1-Hour Iteration)
* **Block 1 (0–15 min)**:
  - Generate 2 new ideas.
  - Store them for refinement in Block 2.
* **Block 2 (15–30 min)**:
  - Refine Block 1’s 2 ideas (first refinement pass).
  - Generate 2 new ideas.
  - Store new ideas for refinement in Block 3.
* **Block 3 (30–45 min)**:
  - Refine Block 2’s 2 ideas (first refinement pass).
  - Refine Block 1’s 2 ideas (second refinement pass).
  - Generate 2 new ideas.
  - Store new ideas for refinement in Execution Phase.
* **⚡ Block 4 — Execution Phase (45–60 min)**:
  - *Non-generation block*: Final refinement + production execution.
  - Refine Block 3’s 2 ideas (first refinement pass).
  - Collect all 6 fully refined ideas:
    * Block 1 $\rightarrow$ 2 ideas (fully refined, 2 passes)
    * Block 2 $\rightarrow$ 2 ideas (fully refined, 1 pass)
    * Block 3 $\rightarrow$ 2 ideas (just refined)
  - Execute all 6 refined ideas into code: 100% full implementation, no placeholders, production-ready output, verified with test suite.
  - Carry Block 3’s 2 ideas forward into the next cycle as the first refinement targets.

---

## 8. PowerShell Gallery Publishing & GitHub Actions Governance
Whenever developing, updating, or maintaining PowerShell modules:
- **Dual Destination Publishing**: Modules must be published to both public GitHub repositories AND the PowerShell Gallery (`PSGallery`).
- **Automated GitHub Action Workflow**:
  - Always include `.github/workflows/publish-psgallery.yml` configured to trigger on tagged releases (`v*`) and `workflow_dispatch`.
  - The workflow must run Pester test verification first, set clean NuGet/temp caches (`C:\temp\nuget_cache`), and publish via `Publish-Module -NuGetApiKey ${{ secrets.PSGALLERY_API_KEY }}`.
- **Local `publish.ps1` Standard**:
  - Maintain a standard `publish.ps1` script in the repository root that checks `$env:PSGALLERY_API_KEY` / User / Machine environment scopes, sets clean temp cache paths (`C:\temp`), and executes `Publish-Module -Path $PSScriptRoot -NuGetApiKey $apiKey -Force -Verbose`.
- **NuGet Cache Locking Prevention**:
  - Always override `$env:TEMP`, `$env:TMP`, `$env:NUGET_PACKAGES`, and `$env:NUGET_HTTP_CACHE_PATH` to `C:\temp\...` to avoid OneDrive/AppData file locking during publishing.

---

## 9. Windows Autopilot, Intune & Hybrid Join Governance
When developing or maintaining Autopilot Command Hub, AutopilotFast, IntuneShared, or WingetIntune:

- **Strict OA3 Hardware Hash Validation**:
  - NEVER fabricate or synthesize a 4096-byte hardware hash when WMI (`MDM_DevDetail_Ext01`) cannot be read. Synthetic hashes permanently corrupt Intune tenant records.
  - Genuine OA3 blobs carry the magic header `4F 41 33 00` (`OA3\0`, Base64 prefix `T0EzAA`) and measure between 2,048 and 16,384 bytes. The payload is NOT ASN.1 DER (`0x30`). Block CSV export and Graph upload if the payload does not validate against this binary structure.
- **Runtime Privilege Detection & Elevation Guide**:
  - Automatically detect runtime context via `Get-HubRuntimeContext`: distinguish between `SYSTEM` (OOBE Shift+F10 / defaultuser0), elevated `Administrator` on desktop, and `StandardUser`.
  - If privileges are limited, present guided elevation (`Show-PrivilegeGuide`) with a one-click "Relaunch as Administrator" button that self-copies the active script and `.env` to `%ProgramFiles%\AutopilotCommandHub` and elevates via UAC.
- **WPF STA Runspace Safety**:
  - WPF applications in PowerShell cannot run on a raw `[System.Threading.Thread]` without a Runspace; any MTA host will crash on invoke. Re-run within a dedicated STA Runspace or launch via `powershell -STA`.
- **Hybrid Join & Co-Management Local Diagnostics**:
  - Target failure modes for on-prem domain + Entra hybrid joins via `dsregcmd /status` (PRT token status, tenant IDs, join flags).
  - Decode the ConfigMgr/Intune co-management authority bitmask to determine per-workload ownership (Compliance, Config, Apps, Updates, Office).
  - Gate all remediation actions (`Restart-IntuneExtension`, `Invoke-MdmSync`, `Invoke-BitLockerEscrow`, `Invoke-GpUpdateForce`) behind elevation checks.
- **Non-Blocking Startup Discipline**:
  - The window must render before slow probes finish. Never call `Test-StagedNetwork`, `Get-WindowsLicensingInfo`, `Get-Tpm`, `Get-PhysicalDisk`/`Get-StorageReliabilityInfo`, or `Get-BatteryHealthInfo` synchronously on the startup path - dispatch them through `Start-HubAsyncWork` and apply results in the `-OnComplete` callback (it runs on the UI thread via the 30 ms `HubQueueTimer`). The 7-stage ladder is additionally pre-started in a standalone runspace in Section C so it overlaps the console intro.
  - Prefer cheap local reads over WMI/CIM on the hot path: `[System.IO.DriveInfo]::GetDrives()` over `Win32_Volume`; the `SecureBoot\State` registry value (readable unelevated) over `Confirm-SecureBootUEFI`. Cache `SoftwareLicensingService.OA3xOriginalProductKey` (`$script:CachedOa3Key`).
  - `Get-Tpm` returns nothing without elevation; render that as `NEEDS ELEVATION`, not `NOT READY`/`NOT DETECTED`.
- **Easter Eggs (keep them harmless)**:
  - Any egg (`Show-HubIntro` shootout, `Get-HubQuip` log flavour, Konami -> Outlaw title, 5x title-click -> ASCII horse) must be purely cosmetic, side-effect free, and skippable. `Show-HubIntro` must no-op when `Test-HubConsoleAvailable` is false (no ConsoleHost, redirected output, tiny window), on `-NoIntro`, on `AUTOPILOT_NO_INTRO`, and for the STA relaunch child (`AUTOPILOT_HUB_STA_CHILD=1`). Never let an egg touch provisioning state.
- **Autonomous Multi-Repo Bootstrap Distribution**:
  - Keep `onyachamp.com` in continuous lockstep with `AutopilotCommandHub` via dual automated pipelines:
    1. Push trigger in `AutopilotCommandHub` (`deploy-website.yml`) using repository secret `SITE_DEPLOY_TOKEN`.
    2. Scheduled pull sync in `onyachamp` (`sync-autopilot.yml`) querying GitHub raw every 30 minutes with fallback `workflow_dispatch`.



