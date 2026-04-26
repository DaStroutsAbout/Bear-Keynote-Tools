# Bear Engineering Inspection Report Tools

A suite of AppleScript and Python tools for formatting and archiving Keynote-based foundation and drainage inspection reports at Bear Engineering. The formatting tools run from Script Editor as plain `.applescript` files and are designed for the Bear Engineering Keynote template. The Report Extractor is a standalone Python tool with its own launcher.

---

## Tools

### Keynote Tools (v2.0)

The primary formatting tool for new inspection reports. Presents a menu at launch with four programs:

**1 — Organize All Slides (with arrows)**
Full pass over the detected slide range. Positions photos into a standardized grid, cleans up text boxes, repositions the body text box based on photo count, and places directional arrows on every photo. If a slide has more photos than available grid slots, continuation slides are created automatically. Start and end slides are auto-detected from the 200×267px property overview photo.

**2 — Fix One Slide (reset photo order, with arrows)**
Same as Program 1 but for a single slide. Photos are reassigned to grid positions sequentially and arrows are placed. Defaults to the slide currently open in Keynote.

**3 — Fix One Slide (keep photo order, no arrows)**
Single-slide photo straightening that snaps each photo to its nearest grid position without reordering. No arrows placed. Use this after manually arranging photos when you want the grid tidied without disturbing your ordering or adding duplicate arrows.

**4 — Condense Slides**
Scans the detected slide range top-to-bottom and merges any slide whose content fits onto the previous slide with a 10pt overlap gap. Useful after content editing has left slides partially filled. Keynote must stay frontmost while this runs.

---

### Re-Inspection Tools (v1.0)

Handles the additional steps specific to re-inspection report preparation. Presents a menu at launch with five programs:

**1 — Expand Slides**
Splits condensed slides back apart, one observation block per slide. Run this before organizing a re-inspection report that was previously condensed.

**2 — Organize Entire Report (Re-Inspection)**
Runs the re-inspection organizer across all detected content slides. Adapted from Keynote Tools but tailored for re-inspection layout and slide title patterns.

**3 — Organize Current Slide (Re-Inspection)**
Same as Program 2 but for the currently visible slide only.

**4 — Snap Photos to Grid**
Snaps photos to the nearest grid position on the current slide. No reordering, no arrows.

**5 — Condense Slides**
Same merge logic as Keynote Tools Program 4.

After organizing, Program 1 also places a Re-Inspection label image on every visible slide in the report range.

---

### Report Extractor (v1.0)

A Python tool that walks a folder tree, finds every Foundation & Drainage Inspection Report Keynote file, and extracts the text content into structured plain-text `.txt` files. One file is produced per report and saved to a timestamped folder on the Desktop.

The primary use case is uploading the resulting `.txt` files into a Claude project so Bear inspectors can query past report language — for example, how the team typically writes up a high water table finding or what language is used for concrete slab cracking within tolerance.

**What it extracts** from each report: the introduction text (slide 2), each observed condition block with its title and body text, and the conclusion. Slides with no body text, hidden slides, the title page, item image labels, page numbers, and appendix slides are excluded. Multi-block condensed slides are split so each condition gets its own section in the output.

**How to run it:** double-click `Extract Reports.command`, select a folder in the dialog, and let it run. The tool searches all subfolders automatically and can handle a top-level archive folder containing hundreds of jobs. When complete, Finder opens to the output folder.

**Runtime:** approximately 16 seconds per report on a MacBook Air with mixed local and cloud-synced files. 100 reports takes roughly 27 minutes; 500 takes about 2 hours.

**Required files** (all three must stay in the same folder):

- `Extract Reports.command` — the launcher
- `extract_reports.py` — main Python script
- `extract_report.applescript` — AppleScript module called once per report

**Installation note:** downloaded `.command` files lose their executable permission on macOS. Before first use, run `chmod +x ~/Desktop/Report\ Extractor/Extract\ Reports.command` in Terminal. Python 3 must also be installed.

---

## Version History

- **v1.0** — Initial release with Programs 1, 2, and 3. Portrait and landscape photo grid layout, text box cleanup and repositioning based on photo count, and arrow placement with three arrow types selected by slide title.

- **v1.1** — Added Program 4 (Condense Slides).

- **v1.2** — Added open file detection across all four programs. Fixed dialog focus and Keynote foreground behavior.

- **v1.3** — Added rebar rusting slide detection with adjusted body text width and arrow type. Reworked Condense Slides to run top-to-bottom. Added auto-detection of start and end pages from property overview photos. Hidden slides excluded from auto-detection. Current slide used as default for Programs 2 and 3.

- **v1.4** — Initial release to Bear team. Program 2 updated to include arrow placement. All page/slide number dialogs removed in favor of auto-detected values. Fixed off-canvas image exclusion across all programs. Added page range auto-detection to Program 4.

- **v1.5** — Fixed wide photo and text box placement. When a wide photo is present, the text box is placed full-width at x=34 in the next row with additional spacing for the caption label.

- **v2.0** — Added overflow handling to Programs 1 and 2. Continuation slides are created when a slide has more photos than grid slots. Labels are created for all photos starting at Item Image 1. Body text moves to the last slide; a text-only continuation slide is created if needed. Program 3 gains label creation for photos missing a label.

- **Re-Inspection Tools v1.0** — Separate tool bundle for re-inspection report preparation, including Expand Slides, re-inspection-specific organizer programs, and Re-Inspection label placement.

- **Report Extractor v1.0** — Python tool for batch extraction of inspection report text from Keynote files. Folder tree search, multi-block slide detection, cloud file pre-open handling, and structured `.txt` output.

---

## Notes for Developers

- All scripts run from Script Editor as plain `.applescript` files. Exporting as compiled apps introduces macOS Accessibility permission issues, especially for programs that use keyboard shortcuts.
- Programs that use clipboard keystrokes (`cmd+A`, `cmd+C`, `cmd+V`) require Keynote to remain frontmost throughout their run.
- `skipped of every slide of theDoc` is fetched as a single Apple Event at the start of each program to avoid per-slide round-trips.
- All coordinate values (grid positions, y thresholds, page detection bounds) are hardcoded to the Bear Engineering Keynote template dimensions. If the template changes, these values will need to be updated.

---

## Team

Developed by David Strout
Bear Engineering, San Francisco Bay Area
Licensed to Domalytx, Inc. under consulting agreement effective 3/20/2026
