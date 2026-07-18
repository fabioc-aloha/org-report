# org-report

A GitHub Copilot CLI plugin that generates polished executive organization briefs from Microsoft 365 signals via WorkIQ.

Given a target executive, the plugin walks the reporting tree in Microsoft Graph, synthesizes an AI profile for every person (leadership chain, direct reports, extended teams, and vendor partners) from accessible mail / calendar / Teams / documents, and produces a Word + PDF report displayed live in a Copilot canvas.

## What you get

- **Word document** — cover page (target, timestamp, FTE/vendor counts, methodology, disclosure), management chain, color-coded team index, individual profiles for chain + target + team FTEs + individual contributors, and a full contractor/vendor appendix.
- **PDF** — rendered from the DOCX for shareable / read-only distribution.
- **Live canvas** — embedded PDF viewer with a download button, refreshed on every regen.

Example output structure for a GM-level org: ~200 profiles, ~120 FTEs, ~75 vendors, single self-contained PDF.

## Installation

### 1. Install the plugin

Copy or clone this folder into your Copilot CLI plugins directory:

```powershell
# Windows
git clone https://github.com/fabioc-aloha/org-report `
  "$env:USERPROFILE\.copilot\installed-plugins\fabioc-aloha\org-report"
```

```bash
# macOS / Linux
git clone https://github.com/fabioc-aloha/org-report \
  ~/.copilot/installed-plugins/fabioc-aloha/org-report
```

Restart Copilot CLI. The `org-report` skill will auto-load.

### 2. Install the canvas extension

The report is served through a canvas extension. Copy it once into your Copilot extensions folder:

```powershell
# Windows
$src = "$env:USERPROFILE\.copilot\installed-plugins\fabioc-aloha\org-report\extensions\org-directory"
$dst = "$env:USERPROFILE\.copilot\extensions\org-directory"
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Copy-Item "$src\*" $dst -Recurse -Force
New-Item -ItemType Directory -Force -Path "$dst\artifacts" | Out-Null
```

```bash
# macOS / Linux
mkdir -p ~/.copilot/extensions/org-directory/artifacts
cp -r ~/.copilot/installed-plugins/fabioc-aloha/org-report/extensions/org-directory/* \
      ~/.copilot/extensions/org-directory/
```

Reload extensions inside Copilot CLI:

```
extensions_reload
```

### 3. Install Python dependencies

```powershell
pip install -r "$env:USERPROFILE\.copilot\installed-plugins\fabioc-aloha\org-report\scripts\requirements.txt"
```

Microsoft Word is required on Windows for the PDF conversion step. On macOS use `docx2pdf` with the Word for Mac; on Linux swap in LibreOffice.

### 4. Sign in to WorkIQ

Ensure the WorkIQ MCP server is configured in your `.mcp.json` and you have signed in with a Microsoft 365 account. The plugin depends on `workiq-ask` being available.

## Usage

In Copilot CLI, ask any of:

- "Generate an org report for Jane Doe"
- "Build an executive org brief for jane.doe@contoso.com"
- "Make an org profile document for Jane"

Copilot will:

1. Confirm the target and reset the state file.
2. Walk Microsoft Graph for the management chain + full report tree.
3. Enrich every person's profile via WorkIQ (in parallel batches of 10–15).
4. Optionally re-validate suspect collaborator networks.
5. Run `scripts/export_org.py` to produce the DOCX + PDF.
6. Reload the extension and open the canvas.

Expect the enrichment phase to take 10–30 minutes for a ~200-person org (WorkIQ latency is 10–60 s per call, parallelized).

## Manual regeneration

If you edit `directory.json` by hand and just want to re-render:

```powershell
python "$env:USERPROFILE\.copilot\installed-plugins\fabioc-aloha\org-report\scripts\export_org.py"
```

CLI flags:

| Flag | Default | Purpose |
|---|---|---|
| `--state PATH` | canvas artifacts `directory.json` | override state file |
| `--docx-out PATH` | `<Target>-Org-Report.docx` next to the script | override DOCX path |
| `--pdf-out PATH` | canvas artifacts `report.pdf` | override PDF path (empty string skips PDF) |

## Repository layout

```
org-report/
├── .plugin/plugin.json          # plugin manifest
├── skills/org-report/SKILL.md   # workflow instructions Copilot follows
├── extensions/org-directory/    # canvas extension (copy to ~/.copilot/extensions/)
│   └── extension.mjs
├── scripts/
│   ├── export_org.py            # DOCX + PDF renderer
│   └── requirements.txt         # python-docx, docx2pdf
└── README.md
```

## Disclosure & data handling

Every generated report includes a cover-page disclosure covering:

1. Data-access limits (information barriers, sensitivity labels, Purview policies)
2. AI accuracy caveat (verify before acting on specific claims)
3. Confidentiality (Microsoft confidential personnel info; do not redistribute)

Adjust these paragraphs for your tenant, audience, or classification by editing the `_method` and `_disclosure` strings near the top of the render section in `scripts/export_org.py`.

## License

MIT

