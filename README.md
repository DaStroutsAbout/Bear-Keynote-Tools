# Keynote Tools

AppleScript automation for formatting Keynote-based foundation and drainage inspection reports at Bear Engineering. Automates photo grid layout, text box positioning, directional arrow placement, and slide condensing.

---

## Programs

**Program 1 — Organize All Slides (with arrows)**
Full pass over a slide range. Positions all photos into a standardized grid, cleans up text boxes, repositions the body text box based on photo count, and places directional arrows on every photo. Start and end slides are auto-detected from the 200×267px property overview photo.

**Program 2 — Fix One Slide (reset photo order, with arrows)**
Same as Program 1 but for a single slide. Photos are reassigned to grid positions sequentially and arrows are placed. Defaults to the slide currently open in Keynote.

**Program 3 — Fix One Slide (keep photo order, no arrows)**
Single-slide photo straightening that snaps each photo to its nearest grid position without reordering. No arrows placed. Use this after manually arranging photos when you want the grid tidied without disturbing your ordering or adding duplicate arrows. Defaults to the slide currently open in Keynote.

**Program 4 — Condense Slides**
Scans a slide range top-to-bottom and merges any slide whose content fits onto the previous slide with a 10pt overlap gap. Useful after content editing has left slides partially filled. Uses clipboard keystrokes — Keynote must stay frontmost while running.

---

## Version History

- **v1.0** — Initial release with Programs 1, 2, and 3. Portrait and landscape photo grid layout, text box cleanup and repositioning based on photo count, and arrow placement with three arrow types selected by slide title.

- **v1.1** — Added Program 4 (Condense Slides), merging sparse slide content onto the previous slide using clipboard keystrokes.

- **v1.2** — Added open file detection across all four programs, offering to use the frontmost Keynote document rather than requiring manual file selection. Fixed dialog focus so page number prompts appear in front of Keynote. Keynote is brought back to front after dialogs so the run can be watched live.

- **v1.3** — Added rebar rusting slide detection, reducing body text box width to 535pt and switching to the vertical arrow on those slides. Reworked Condense Slides to run top-to-bottom with a -10pt overlap gap. Added auto-detection of start and end pages from 200×267px property overview photos. Hidden slides excluded from page auto-detection. Current slide used as default for Programs 2 and 3.

- **v1.4** — Initial release to BEAR team. Program 2 updated to include arrow placement, making it a full single-slide equivalent of Program 1. All page and slide number dialogs removed in favor of auto-detected values. Fixed off-canvas image exclusion across all three programs so images outside page boundaries are not counted toward text box positioning. Added page range auto-detection to Program 4.

- **v1.5** — Resolved a bug related to wide photo and text box placement. When a wide photo is present, the text box is now always placed full-width (700pt) at x=34 in the next row down with 20pt additional spacing to account for the caption label.

---

## Notes for Developers

- The script is designed to run from Script Editor as a plain `.applescript` file. Exporting as a compiled app introduces macOS Accessibility permission issues, particularly for Program 4 which uses keyboard shortcuts.
- Program 4 uses `cmd+A`, `cmd+C`, `cmd+V` keystrokes and requires Keynote to remain the frontmost application throughout its run.
- `skipped of every slide of theDoc` is fetched as a single Apple Event at the start of each program for performance — this returns the full list of hidden slide states without a per-slide round-trip.
- All coordinate values (grid positions, y thresholds, page number detection bounds) are hardcoded to the Bear Engineering Keynote template dimensions. If the template changes, these values will need to be updated.

---

## Team

Developed by David Strout, E.I.T.  
Bear Engineering, San Francisco Bay Area  
Licensed to Domalytx, Inc. under consulting agreement effective 3/20/2026
