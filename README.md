# Keynote Tools
AppleScript automation suite for formatting Keynote-based foundation and drainage inspection reports at Bear Engineering. Automates photo grid layout, text box positioning, directional arrow placement, slide condensing, and slide splitting.

---

## Files

| File | Description |
|------|-------------|
| `Keynote Tools.applescript` | Main tool — contains Programs 1 through 4 |
| `Keynote_Tools_Guide.docx` | Full user guide and technical reference |

### Required supporting files (not in this repo)

Each job folder must contain an `Arrows/` subfolder alongside the Keynote report file with three PNG arrow images:

```
Your Job Folder/
    Foundation & Drainage Inspection Report_123 Main St.key
    Arrows/
        red_arrow_horiz.png
        red_arrow_vert.png
        blue_arrow_horiz.png
```

The Arrows folder is not stored in this repo because it contains binary image files that don't change. Keep a master copy somewhere on your Mac and duplicate it into each new job folder.

---

## Installation

1. Clone or download this repository to your Mac — a good location is `~/Desktop`
2. Open `Keynote Tools.applescript` in **Script Editor** (search for it in Spotlight)
3. Copy the `Arrows/` folder into each job folder alongside the Keynote report
4. The first time you run Program 4 or `expand_slides`, macOS will prompt for **Accessibility permission** — go to System Settings → Privacy & Security → Accessibility and enable Script Editor

To run: open the script in Script Editor and click the **Run** button (▶). A menu will appear.

---

## Programs

### Master Script (`Keynote Tools.applescript`)

**Program 1 — Organize All Slides (with arrows)**
Full pass over a slide range. Positions all photos into a standardized grid, cleans up text boxes, repositions the body text box based on photo count, and places directional arrows on every photo. Start and end slides are auto-detected from the 200×267px property overview photo.

**Program 2 — Fix One Slide (reset photo order, with arrows)**
Same as Program 1 but for a single slide. Photos are reassigned to grid positions sequentially. Arrows are placed. Defaults to the slide currently open in Keynote.

**Program 3 — Fix One Slide (keep photo order, no arrows)**
Single-slide photo straightening that snaps each photo to its nearest grid position without reordering. No arrows placed. Use this after manually arranging photos when you want the grid tidied without disturbing your ordering or adding duplicate arrows. Defaults to the slide currently open in Keynote.

**Program 4 — Condense Slides**
Scans a slide range top-to-bottom and merges any slide whose content fits onto the previous slide with a 10pt overlap gap. Useful after content editing has left slides partially filled. Uses clipboard keystrokes — Keynote must stay frontmost while running.

---

## How Slides Are Structured

Each content slide in the report template contains:

- A **title/header** text box at y < 50 (e.g. "Observed Condition - Foundation Cracks (small)")
- Up to 9 **site photos** in a portrait or landscape grid
- **Item Image** caption labels below each photo
- A **body text box** (observation details and recommendations) — always the longest text element on the slide
- A **page number** in the bottom-right corner at x > 700, y > 950

---

## Photo Size Reference

| Dimensions | Type | Behavior |
|-----------|------|----------|
| 200 × 267 px | Property overview photo | Marks first and last content slides for auto-detection |
| 333 × 250 px | Landscape site photo | Triggers landscape grid layout for whole slide |
| 434 × 267 px | Wide site photo | Skips column 3 in portrait grid; forces full-width text box |
| All others > 70px | Standard site photo | Placed sequentially in portrait grid |

### Portrait grid positions (x, y)

|  | Col 1 | Col 2 | Col 3 |
|--|-------|-------|-------|
| Row 1 | (50, 57) | (284, 57) | (518, 57) |
| Row 2 | (50, 357) | (284, 357) | (518, 357) |
| Row 3 | (50, 657) | (284, 657) | (518, 657) |

### Landscape grid positions (x, y)

|  | Col 1 | Col 2 |
|--|-------|-------|
| Row 1 | (41, 73) | (394, 73) |
| Row 2 | (41, 373) | (394, 373) |
| Row 3 | (41, 673) | (394, 673) |

---

## Arrow Logic

Arrow type is selected by slide title:

| Slide title contains | Arrow used |
|---------------------|-----------|
| "Drainage Improvements - Exterior grade/slope" | Blue horizontal |
| "Concrete cracking/spalling from steel rebar rusting" | Red vertical |
| Anything else | Red horizontal (default) |

---

## Special Cases

**Rebar rusting slides:** body text box width is reduced to 535px (instead of 700px) to leave room for the vertical arrow.

**Wide photos (434×267):** when a wide photo is present, the body text box is always placed full-width (700px) at x=34 in the next row below the wide photo, with an additional 20pt spacing to account for the image caption label. The wide photo's row is detected by its y position after layout.

---

## Version History

- **v1.0** — Initial release with Programs 1, 2, and 3. Portrait and landscape photo grid layout, text box cleanup and repositioning based on photo count, and arrow placement with three arrow types selected by slide title.

- **v1.1** — Added Program 4 (Condense Slides), merging sparse slide content onto the previous slide using clipboard keystrokes.

- **v1.2** — Added open file detection across all four programs, offering to use the frontmost Keynote document rather than requiring manual file selection. Fixed dialog focus so page number prompts appear in front of Keynote. Keynote is brought back to front after dialogs so the run can be watched live.

- **v1.3** — Added rebar rusting slide detection, reducing body text box width to 535pt and switching to the vertical arrow on those slides. Reworked Condense Slides to run top-to-bottom with a -10pt overlap gap. Added auto-detection of start and end pages from 200×267px property overview photos. Hidden slides excluded from page auto-detection. Current slide used as default for Programs 2 and 3. Condense Slides logic reworked from bottom-to-top to top-to-bottom traversal.

- **v1.4** — Initial release to BEAR team. Program 2 updated to include arrow placement, making it a full single-slide equivalent of Program 1. All page and slide number dialogs removed in favor of auto-detected values. Fixed off-canvas image exclusion across all three programs so images outside page boundaries are not counted toward text box positioning. Added page range auto-detection to Program 4.

- **v1.5** — Resolved a bug related to wide photo and text box placement. Wide photos now correctly consume 2 grid slots, and all prior x-position and remainder slot logic has been replaced with a simplified rule: when a wide photo is present, the text box is always placed full-width (700pt) at x=34 in the next row down with 20pt additional spacing to account for the legend.

---

## Notes for Developers

- The script is designed to run from Script Editor as a plain `.applescript` file. Exporting as a compiled app introduces macOS Accessibility permission issues, particularly for Program 4 and `expand_slides` which use keyboard shortcuts.
- Program 4 and `expand_slides` use `cmd+A`, `cmd+C`, `cmd+V` keystrokes and require Keynote to remain the frontmost application throughout the run.
- `skipped of every slide of theDoc` is fetched as a single Apple Event at the start of each program for performance — this returns the full list of hidden slide states without a per-slide round-trip.
- All coordinate values (grid positions, y thresholds, page number detection bounds) are hardcoded to the Bear Engineering Keynote template dimensions. If the template changes, these values will need to be updated.

---

## Team

Developed by David Strout, E.I.T.  
Bear Engineering, San Francisco Bay Area  
Licensed to Domalytx, Inc. under consulting agreement effective 3/20/2026
