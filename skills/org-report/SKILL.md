---
name: org-report
description: Generate an executive organization brief (leadership chain, direct reports, extended teams, and vendor partners) for a target person. Produces a Word + PDF report from Microsoft 365 signals via WorkIQ and displays it in a Copilot canvas. Triggers on "generate org report for X", "build org brief for X", "make an org report for X", "org profile document for X".
---

# Org Report

Generate a polished executive organization brief for a target person. The report includes:

- Cover page with target, generation timestamp, FTE / vendor counts, byline, methodology, and disclosure
- Full management chain (up to the top)
- Team index (color-coded)
- Individual profiles: chain, target, each team's FTEs, individual contributors
- Appendix: all contractors and vendors grouped by team

Output is a Word `.docx`, a `.pdf` rendered from it, and a live Copilot canvas view.

## When to use

Trigger on any of:

- "Generate an org report for {Person}"
- "Build an org brief for {Person}"
- "Make an executive org report for {Person}"
- "Create an org profile document for {Person}"
- "Produce an organization brief for {Person}"

If the user asks for a report on themselves, use their profile as the target.

## Prerequisites (verify before starting)

1. **WorkIQ MCP** — the `workiq-ask` tool must be available.
2. **Python 3.11+** with `python-docx` and `docx2pdf` installed:
   ```powershell
   pip install python-docx docx2pdf
   ```
3. **Microsoft Word** — required by `docx2pdf` on Windows.
4. **Canvas extension** — the `org-directory` extension should be installed at
   `~/.copilot/extensions/org-directory/`. This plugin ships a copy in
   `extensions/org-directory/`; copy it once, then `extensions_reload`.

If any prerequisite is missing, stop and tell the user what to install.

## Workflow

### Phase 1 — Confirm the target and reset state

Ask the user to confirm the target (UPN, email, or full name). Then create a
fresh state file at `~/.copilot/extensions/org-directory/artifacts/directory.json`:

```json
{ "orgChart": null, "profiles": {}, "updatedAt": null }
```

### Phase 2 — Resolve the org structure via WorkIQ

Use `workiq-ask` (or `workiq-fetch` on `/users/...` where possible) to build the tree:

1. **Target profile**: get `displayName`, `jobTitle`, `department`, `email`, `officeLocation`.
2. **Management chain**: walk up via `/users/{id}/manager` until you reach the top
   (e.g., Satya Nadella at Microsoft). Save as `managementChain` (ordered top → target's manager).
3. **Direct reports**: recursively fetch `/users/{id}/directReports` for the target
   and every descendant. Attach as nested `reports: [...]`.

Persist to `directory.json`:

```json
{
  "orgChart": {
    "managementChain": [ { "displayName": "...", "jobTitle": "..." }, ... ],
    "target": { "displayName": "...", "jobTitle": "...", "email": "...", "department": "..." },
    "directReports": [
      { "displayName": "...", "jobTitle": "...", "email": "...",
        "reports": [ { "displayName": "...", ... } ] }
    ]
  },
  "profiles": {},
  "updatedAt": "2026-01-01T00:00:00Z"
}
```

**Vendor detection:** in the Microsoft directory, contractors' display names are
formatted `First Last (Company Name)`. The export script uses this convention plus
email domain (see `_is_vendor` in `scripts/export_org.py`).

### Phase 3 — Enrich every person's profile

For every person in the tree (target + all descendants), call `workiq-ask` in
**parallel batches of 10–15** with this exact prompt:

```
For {Full Name} ({title, if known}, in {target}'s org), produce a concise
professional profile grounded in accessible Microsoft 365 signals (email,
meetings, Teams, documents). Return ONLY this JSON in a ```json fence:

{
  "displayName": "{Full Name}",
  "jobTitle": "...",
  "department": "...",
  "email": "...",
  "summary": "2-3 sentence executive summary of their role and focus areas",
  "recentWork": ["3-6 recent projects or themes"],
  "topics": ["4-8 short tags describing what they work on"],
  "expertise": ["3-6 specific skills or domains"],
  "collaborators": ["6-8 people they work most closely with"],
  "officeLocation": "..."
}
```

Rules:

- **Parallel batches only.** WorkIQ latency is 10–60 s per call. Fire 10–15 calls
  in a single response, never sequential loops.
- **Retry MCP failures individually.** `No MCP client found` and timeouts happen —
  retry the specific failed call once.
- **Self-identify.** The `displayName` field lets you match responses back to
  people even in a mixed batch.
- **Merge into state.** Key by `displayName.lower().strip()`. Update `updatedAt`
  on each profile.
- **Empty responses.** If WorkIQ returns nothing usable (common for niche
  contractors), store:
  ```json
  "collaboratorsNote": "Collaborator network not surfaced in accessible signals"
  ```
- **Skip execs above the target.** External executives (e.g., Satya, Judson,
  large-org GMs) have collaborators legitimately outside the org and produce
  noisy profiles. Enrich the target and everyone below; leave upstream chain
  members as chain entries only.

### Phase 4 — (Optional) Validate collaborators

The first-pass collaborator list often includes people from unrelated meetings.
For any profile whose collaborators look off (mostly outside the org, unfamiliar
names), do a second pass with a narrower prompt:

```
Who are the 6-8 people {Name} ({title/company}, in {target}'s org) works most
closely with day-to-day? Return only this exact JSON in a ```json fence:
{"person": "{Name}", "collaborators": ["Name1", "Name2", ...]}
```

Same batch pattern (10–15 parallel). Merge overwrites the previous
`collaborators` list and clears any `collaboratorsNote`. Accept ~5% empties.

### Phase 5 — Generate the DOCX + PDF

Run the export script. It reads `directory.json`, writes the DOCX and PDF, and
copies them into the canvas extension's artifacts folder.

```powershell
python ~/.copilot/installed-plugins/fabioc-aloha/org-report/scripts/export_org.py
```

Optional flags:

- `--state <path>` — override the state file (default: canvas artifacts)
- `--docx-out <path>` — override the DOCX output path
- `--pdf-out <path>` — override the PDF output path (set to empty string to skip PDF)

### Phase 6 — Refresh the canvas

```
extensions_reload
open_canvas({ canvasId: "org-directory", instanceId: "org-report-1" })
```

The canvas embeds the fresh PDF with a "Download PDF" button.

### Phase 7 — Report completion

Tell the user:

- Target name and how many FTE + vendor profiles were enriched
- Where the DOCX lives (path)
- That the canvas is open and refreshable

## Editing the cover copy

The methodology and disclosure paragraphs live near the top of the render
section of `scripts/export_org.py` (search for `_method` and `_disclosure`).
Update them for tenant, audience, or classification differences. The default
disclosure covers three things every exec brief of this kind should say:

1. **Data-access limits** — information barriers, sensitivity labels, Purview
2. **AI accuracy** — summaries may contain errors, verify before acting
3. **Confidentiality** — Microsoft confidential personnel info, do not
   redistribute

## Known limits

- **WorkIQ latency:** 10–60 s per call; some questions run minutes. Always batch
  in parallel.
- **MCP flakiness:** occasional `No MCP client found` / timeouts. Retry the
  specific failed call once.
- **Empty profiles:** ~5–10% of contractors have no surfaceable signals. Mark
  with `collaboratorsNote` and move on.
- **Vendor name format:** relies on the `First Last (Company)` directory
  convention. Extend `_is_vendor` in `export_org.py` for other tenants.
- **Directory vs contacts:** `/users/{id}` is AAD; `/me/contacts` is personal
  Outlook. IDs are not interchangeable.
- **Binary content:** WorkIQ MCP doesn't yet expose files or profile photos as
  bytes — only structured metadata.
- **`docx2pdf` on Windows:** requires Microsoft Word. On macOS/Linux, substitute
  LibreOffice or a headless converter.
