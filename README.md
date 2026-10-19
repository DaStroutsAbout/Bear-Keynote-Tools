# Bear Engineering Inspection Report Tools

A suite of AppleScript and Python tools for formatting and archiving Keynote-based foundation and drainage inspection reports at Bear Engineering. The formatting tools run from Script Editor as plain `.applescript` files and are designed for the Bear Engineering Keynote template. The Report Extractor is a standalone Python tool with its own launcher.

Full user documentation for the three formatting scripts is in **Inspection_Report_Tools_Guide.pdf**. The Report Extractor has its own guide, **Report_Extractor_1_0_Guide.pdf**.

---

## Tools

### Keynote Tools (v2.2)

The primary formatting tool for new inspection reports. Presents a menu at launch with four programs:

**1: Organize All Slides (with arrows)**
Full pass over the detected slide range. Positions photos into a standardized grid, creates Item Image labels, cleans up text boxes, repositions the body text box based on the photos on the slide, and places directional arrows on every photo. If a slide has more photos than available grid slots, continuation slides are created automatically. Start and end slides are auto-detected from the 200×267px property overview photo.

**2: Fix One Slide (reset photo order, with arrows)**
Same as Program 1 but for a single slide. Photos are reassigned to grid positions sequentially and arrows are placed. Defaults to the slide currently open in Keynote.

**3: Fix One Slide (keep photo order, no arrows)**
Single-slide photo straightening that snaps each photo to its nearest grid position without reordering. No arrows placed. Adds labels for any photo missing one and creates a text-only continuation slide if the body text no longer fits. Use this after manually arranging photos when you want the grid tidied without disturbing your ordering or adding duplicate arrows.

**4: Condense Slides**
Scans the detected slide range top-to-bottom and merges any slide whose content fits onto the previous slide. Sections are spaced by `section_offset` (see v2.2 below). Useful after content editing has left slides partially filled. Keynote must stay frontmost while this runs.

---

### Kurt's Version (v2.2)

A variant of Keynote Tools for inspectors who arrange photos by hand before running the cleanup. The menu prompt reads "Select Program - Kurt's Version" so it is easy to tell apart. Uses the same Arrows folder.

**1: Organize All Slides (with arrows)**
Same as Keynote Tools Program 1, except photos are sorted by their current on-slide position (top-to-bottom, then left-to-right) and each is snapped to its nearest free grid slot, so a manual arrangement is preserved.

**2: Fix One Slide (keep photo order, no arrows)**
Same as Keynote Tools Program 3.

**3: Condense Slides**
Same as Keynote Tools Program 4.

The reset-order program (Keynote Tools Program 2) is not included.

---

### Re-Inspection Tools (v1.2)

Handles the additional steps specific to re-inspection report preparation, where original photos are carried forward and new re-inspection photos are added. Presents a menu at launch with five programs:

**1: Expand Slides**
Splits condensed slides back apart, one observation block per slide. Run this before organizing a re-inspection report that was previously condensed. Also adds red "Re-Inspection Date:" and "Report Date:" fields (with "tbd" placeholders) to page 2.

**2: Organize Entire Report (Re-Inspection)**
Runs the re-inspection organizer across all detected content slides. Photos already at grid positions are treated as originals and left in place; new photos fill the remaining slots and get arrows and red Item Image labels. Fresh template slides (red title or bottom label) treat every photo as new.

**3: Organize Current Slide (Re-Inspection)**
Same as Program 2 but for the currently visible slide only.

**4: Snap Photos to Grid**
Snaps photos to the nearest grid position on the current slide. No reordering, no arrows. Also cleans up stray text boxes, adds missing labels, and repositions the body text.

**5: Condense Slides**
Same merge logic as Keynote Tools Program 4. Afterwards, places the Re-Inspection label image (`Arrows/Re-Inspection_Label.png`) on every visible slide in the report range. Run it once per report; each run adds another label.

---

### Report Extractor (v1.0)

A Python tool that walks a folder tree, finds every Foundation & Drainage Inspection Report Keynote file, and extracts the text content into structured plain-text `.txt` files. One file is produced per report and saved to a timestamped folder on the Desktop.

The primary use case is uploading the resulting `.txt` files into a Claude project so Bear inspectors can query past report language, for example how the team typically writes up a high water table finding or what language is used for concrete slab cracking within tolerance.

**What it extracts** from each report: the introduction text (slide 2), each observed condition block with its title and body text, and the conclusion. Slides with no body text, hidden slides, the title page, item image labels, page numbers, and appendix slides are excluded. Multi-block condensed slides are split so each condition gets its own section in the output.

**How to run it:** double-click `Extract Reports.command`, select a folder in the dialog, and let it run. The tool searches all subfolders automatically and can handle a top-level archive folder containing hundreds of jobs. When complete, Finder opens to the output folder.

**Runtime:** approximately 16 seconds per report on a MacBook Air with mixed local and cloud-synced files. 100 reports takes roughly 27 minutes; 500 takes about 2 hours.

**Required files** (all three must stay in the same folder):

- `Extract Reports.command`: the launcher
- `extract_reports.py`: main Python script
- `extract_report.applescript`: AppleScript module called once per report

**Installation note:** downloaded `.command` files lose their executable permission on macOS. Before first use, run `chmod +x ~/Desktop/Report\ Extractor/Extract\ Reports.command` in Terminal. Python 3 must also be installed.

---

## Version History

### Keynote Tools

- **v1.0**: Initial release with Programs 1, 2, and 3. Portrait and landscape photo grid layout, text box cleanup and repositioning based on photo count, and arrow placement with three arrow types selected by slide title.

- **v1.1**: Added Program 4 (Condense Slides).

- **v1.2**: Added open file detection across all four programs. Fixed dialog focus and Keynote foreground behavior.

- **v1.3**: Added rebar rusting slide detection with adjusted body text width and arrow type. Reworked Condense Slides to run top-to-bottom. Added auto-detection of start and end pages from property overview photos. Hidden slides excluded from auto-detection. Current slide used as default for Programs 2 and 3.

- **v1.4**: Initial release to Bear team. Program 2 updated to include arrow placement. All page/slide number dialogs removed in favor of auto-detected values. Fixed off-canvas image exclusion across all programs. Added page range auto-detection to Program 4.

- **v1.5**: Fixed wide photo and text box placement. When a wide photo is present, the text box is placed full-width at x=34 in the next row with additional spacing for the caption label.

- **v2.0**: Added overflow handling to Programs 1 and 2. Continuation slides are created when a slide has more photos than grid slots. Labels are created for all photos starting at Item Image 1. Body text moves to the last slide; a text-only continuation slide is created if needed. Program 3 gains label creation for photos missing a label.

- **v2.2**: Floor elevation map support. The map is now 668×411 (previously 434×267). Programs 1 to 3 place it after all other photos, at grid slot 1 or slot 4, spanning the full grid width and two rows. Its Item Image label goes at x=50, y=map top + 417, numbered in sequence with the other photos, and the body text is placed full width below it. Condense Slides now adds 10pt more space between sections (`section_offset`, previously a 10pt overlap).

### Kurt's Version

- **Initial**: Companion to Keynote Tools v2.0. Program 1 sorts photos by on-slide position and snaps each to its nearest free slot. Reset-order program removed; programs renumbered.

- **v2.2**: Same floor elevation map and Condense spacing changes as Keynote Tools v2.2. Fixed continuation slide label numbering, which assumed every slot on the previous slide was filled.

### Re-Inspection Tools

- **v1.0**: Separate tool bundle for re-inspection report preparation, including Expand Slides, re-inspection-specific organizer programs, and Re-Inspection label placement.

- **v1.2**: Floor elevation map support in Programs 2 to 4, including original maps occupying their full span. Fixed continuation slide numbering on fresh template slides, which repeated the last label number from the previous slide. Condense Slides spacing change, same as Keynote Tools v2.2.

### Report Extractor

- **v1.0**: Python tool for batch extraction of inspection report text from Keynote files. Folder tree search, multi-block slide detection, cloud file pre-open handling, and structured `.txt` output.

---

## Notes for Developers

- All scripts run from Script Editor as plain `.applescript` files. Exporting as compiled apps introduces macOS Accessibility permission issues, especially for programs that use keyboard shortcuts.
- Programs that use clipboard keystrokes (`cmd+A`, `cmd+C`, `cmd+V`) require Keynote to remain frontmost throughout their run. This includes label creation, continuation slides, and Condense.
- `skipped of every slide of theDoc` is fetched as a single Apple Event at the start of each program to avoid per-slide round-trips.
- All coordinate values (grid positions, y thresholds, page detection bounds) are hardcoded to the Bear Engineering Keynote template dimensions. If the template changes, these values will need to be updated.
- Photo types are identified by size. The floor elevation map is detected by the `is_map_size` handler (666 to 670 wide, 409 to 413 tall). Where it can go on each grid is set by `map_valid_slots` and `map_span`, which are defined alongside each grid's coordinate lists.
- Item Image labels are created by copying an existing label on the slide, because Keynote does not allow setting font, size, or color on newly created text items. A slide needs at least one label for new ones to be made.
- Condense spacing is controlled by `section_offset` in each script's Condense program.

---

## Team

Developed by David Strout
Bear Engineering, San Francisco Bay Area
Licensed to Domalytx, Inc. under consulting agreement effective 3/20/2026