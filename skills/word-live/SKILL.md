---
name: word-live
description: Edit the Word document the user has open in Microsoft Word, live, while they watch and keep working - read, fill in forms, replace text, insert images, track changes, and check the layout - without closing Word, popping up windows, or moving their cursor. Works on Windows (agent and Word on the same PC) and on macOS with Word inside a Parallels Windows VM. Use whenever the user asks for changes "in Word", to a document open in Word, or to fill in a .docx form they have open, especially Persian/RTL documents.
compatibility: opencode, claude-code
---

# Live editing of the user's open Word

The scripts in `scripts/` drive the **running** Word over COM, so edits land in
the open document immediately. The user never closes or reopens anything, and
can scroll and type in the document while you work.

`SKILL_DIR` below means the folder that holds this `SKILL.md` (the skill tool
reports it). Call the scripts by absolute path.

## Pick the setup

| Where the agent runs | Where Word runs | Scripts |
|---|---|---|
| Windows | same Windows PC | `word.cmd`, `render.cmd`, `shot.cmd` |
| macOS | Parallels Windows VM on that Mac | `word`, `render`, `shot` |

Both run the same PowerShell snippets, with the same helpers (`prelude.ps1`).

### Windows

Requirements: Windows 10/11, Word open with the document active. Nothing to
install: the scripts use the built-in Windows PowerShell 5.1 (`powershell.exe`,
also when the agent's own shell is PowerShell 7 or Git Bash).

1. Write the snippet to a `.ps1` file with your file-writing tool (UTF-8 is
   fine), e.g. `%TEMP%\wl-job.ps1`. A file is the safe way for Persian text:
   piping text into a program from Windows PowerShell 5.1 turns it into `????`.
2. Run it:

   | Shell | Command |
   |---|---|
   | cmd | `"SKILL_DIR\scripts\word.cmd" C:\path\wl-job.ps1` |
   | PowerShell | `& "SKILL_DIR\scripts\word.cmd" C:\path\wl-job.ps1` |
   | Git Bash | `"SKILL_DIR/scripts/word.cmd" C:/path/wl-job.ps1` (or `cmd //c ...`) |

   It prints the snippet's output; the exit code is the snippet's (1 after an
   exception, with the error on stderr).
3. Layout check: `render.cmd <prefix> job.ps1`, where the job calls
   `Render <from> [<to>]`; it writes `<prefix>-1.png`, `-2.png` … (1600 px wide)
   and prints their paths. Look at the PNGs.
4. What the user sees right now: `shot.cmd [out.png]` (whole screen).

### macOS + Parallels

Requirements: a running Parallels VM with Parallels Tools, Word open there with
the document active. The VM is the first running one, or `$WORD_VM`. No shared
folder is needed for the scripts (they reach Windows through `prlctl`).

```bash
SKILL_DIR/scripts/word <<'PS'
$doc.Name + " | pages=" + $doc.ComputeStatistics(2) + " | read-only=" + $doc.ReadOnly
PS
```

- `word`: snippet on stdin; prints its output; exit code is the snippet's.
- `render <prefix>`: same, for a snippet that calls `Render`; writes
  `<prefix>-1.png` … (uses `swift` / PDFKit, built into macOS).
- `shot [out.png]`: screenshot of the VM screen.

### Quick check (both setups)

```powershell
$doc.Name + " | pages=" + $doc.ComputeStatistics(2) + " | read-only=" + $doc.ReadOnly
```

A read-only document (often one opened from Downloads) can still be edited in
memory, but tell the user it needs Save As to keep the changes. If Word is not
running, the snippet fails with `MK_E_UNAVAILABLE`: ask the user to open the
document.

## What a snippet gets (from `prelude.ps1`)

- `$word`, `$doc` (the active document); `$ErrorActionPreference = 'Stop'`.
- `Find-In $scope $text` — first match inside a range (a cell, a page). Prefer this.
- `Find-Text $text [$after]` — first match in the document. Only for unique anchors.
- `Find-All $text` — every match, collected before you change any of them.
- `Append-Text $para $text` — add text at the end of a paragraph; returns its range.
- `Add-ParagraphAfter $para $text` — new plain (unbulleted) paragraph after
  `$para`, same indent; works on the last paragraph of a table cell.
- `Set-Plain $range [$size]` — not bold, Latin font Times New Roman.
- `Fix-Bidi $range` — fix punctuation order in mixed Persian/English text (below).
- `With-Selection $range { ... }` — for Selection-only calls; restores the user's
  cursor and scroll afterwards.
- `Show $range` — select and scroll there, only if the snippet set `$Watch = $true`.
- `Render $from [$to]` — export those pages to PDF in the background (with `render`).

Page 1 as a range: `$doc.Range(0, $doc.GoTo(1, 1, 2).Start)`.
A table cell holding some label: `(Find-Text 'label').Cells(1).Range`.

## How to work

1. **Turn on Track Changes** at the top of every editing snippet
   (`$doc.TrackRevisions = $true`) so the user can accept or reject each change.
2. **Edit through ranges, not the selection.** The user may be scrolling or
   typing elsewhere; never move their view unless they ask to watch
   (`$Watch = $true`). COM calls rejected while they type are retried for 30 s.
3. **Filled-in answers are plain**: never bold, even next to a bold label; match
   the surrounding size; Latin text in Times New Roman (`Set-Plain`).
4. **After inserting text that mixes Persian and English, call `Fix-Bidi`** on the
   whole paragraph (label and answer), e.g. `Fix-Bidi $para.Range`.
   Text put in through COM leaves spaces and punctuation next to English words
   left-to-right, so in an RTL paragraph `(…)`, `:` and `،` jump to the wrong side
   ("Python 3.13 (سرور)" renders scrambled).
5. **Check the layout with `render`, not screenshots.** It exports only the pages
   you name, in the background, in about 7 s, exactly as Word lays them out
   (tracked changes appear accepted). Use `shot` only to see what the user sees,
   e.g. to spot a dialog box.
6. **Verify every edit landed once.** After any error, count occurrences of the
   new text before retrying anything.
7. **Keep jobs short and scoped.** One logical change per call; guard with checks
   that throw (`if (-not $r) { throw "... not found" }`) instead of guessing.

## Word gotchas (learned the hard way)

- **Find inside tables can jump back** to an earlier match when the search range
  starts mid-table, so `while (Find…)` loops can spin forever. Search a known
  scope with `Find-In`; `Find-All` has a guard.
- **Find's limit is 255 characters.** Search for a shorter, unique part.
- **Replacing with tracking on leaves the old text as a deletion**, which a new
  search finds again. Collect all targets first, then replace from the end.
- **Text found by `Find-Text` may be the first copy, on another page** (page 1
  often repeats the same email or URL). Scope the search.
- **`InsertParagraphAfter` fails on the last paragraph of a cell**, and after a
  break `$para.Next()` can point past the cell. `Add-ParagraphAfter` handles both.
- **Deleting your own tracked insertion removes it outright**; deleting the
  user's text leaves a tracked deletion.
- **Empty range: `$range.Text` is `$null`.** Use `"$($r.Text)"` before trimming.
- **A tall picture that does not fit the page** jumps to the next one and leaves
  a near-empty page. Shrink it until it lands on the caption's page:
  `$doc.Repaginate(); $pic.Range.Information(3)`.
- **Persian digits and dates:** Persian forms write dates as `1405/07/29`
  (year/month/day; it displays with the year on the right). Spaced forms like
  `29 / 07 / 1405` wrap and reorder; use the compact form.
- Revision authors: the user's edits and yours carry the same Word user name,
  so you cannot tell them apart by author. Ask the user before reverting a change
  you did not make.

## PowerShell + COM traps

- **Pass plain strings to Word.** A path from `Join-Path` is a wrapped PSObject, and
  `ExportAsFixedFormat`/`SaveAs2` given one can hang Word's call forever with no error.
  Use a literal (`"C:\Users\Public\x.pdf"`), `[IO.Path]::Combine(…)` or `[string](…)`.
- **Index collections with `.Item(n)`**: `$doc.Paragraphs.Item(5)`, `$doc.Tables.Item(1)`,
  `$s.Headers.Item(1)`. `$doc.Paragraphs(5)` fails ("does not contain a method").
- **ByRef arguments need `[ref]`**: `[ref]$fn = '\\Mac\Home\…docx'; [ref]$ff = 12; $doc.SaveAs2($fn, $ff)`.
- **Some methods fail on late binding** (`TabStop.Clear()`, `TabStops.Add(…)`):
  `[System.__ComObject].InvokeMember('Clear', [Reflection.BindingFlags]::InvokeMethod, $null, $ts, $null)`.
- **Don't rebind `$doc` to `$word.Documents | ? {…}`**: the prelude helpers (`Fix-Bidi`,
  `With-Selection`) then fail with "should be a PSReference". To work on another open
  document, `.Activate()` it (one job); next jobs get it as `$doc`.
- **A read-only template** (e.g. a shared report template): `SaveAs2` it to a
  new path (Downloads) and work on that copy; the original file is never touched. For a
  freshly built copy, Track Changes can stay off — tell the user.
- **Long builds:** keep helper functions in a scratch `helpers.ps1` and run
  `cat helpers.ps1 - <<'PS' | scripts/word`, one section per job.
- **Building from the end:** put text into the last (empty) paragraph, set its style,
  `InsertParagraphAfter()`, then reset the new last paragraph to `Normal` + `Font.Reset()`.
  `Tables.Add($lastPara.Range, rows, cols)` on that empty paragraph.
- **Replace in every story** (body, headers, footers): loop `$doc.StoryRanges` and
  `.NextStoryRange`; `Find` with `MatchCase`, set `$r.Text`, `$r.Collapse(0)` and continue.
- **Deleting a page:** list the paragraphs with their control codes first (12 = page break)
  and delete from the page's first paragraph through the one holding its page break.

## Persian reports (NAAD template)

- **Persian digits:** Word's numeral setting decides how page numbers, list/heading
  numbers and TOC numbers look. Set `$word.Options.ArabicNumeral = 2` (Context) and the
  document language to Persian (`LanguageID = 1065`); digits in Persian context then show
  as ۰–۹, while digits inside English words (`TIER4`) stay Latin. Text you type yourself:
  write Persian digits directly (`[char](0x06F0 + d)`). A PDF keeps them; a .docx sent on
  shows list/page numbers per the recipient's own setting.
- **Table of contents after `Update()`/`UpdatePageNumbers()` breaks in RTL:** entries come
  out left-to-right, kashida-justified, with stray left tab stops (44, 88 …pt), so some
  entries lose the dot leader and the page number sits next to the text. After every
  update, for each `TOC*` paragraph: `ReadingOrder = 0`, `Alignment = 2`,
  `TabStops.ClearAll()`, then add a left tab at ~22pt (after the heading number) and a
  right tab with dot leader at text width − 2 (`PageWidth − LeftMargin − RightMargin`),
  via `InvokeMember('Add', …, @([single]pos, alignment, leader))`. A tab beyond the text
  width wraps the number onto the next line. Render the TOC page to check.
- **Headings ending in a Latin word or `)`:** append U+200F so the heading and its TOC
  entry keep their order.
- **Tables:** the template style is `Grid Table 4 - Accent 5,NAAD`; header cells use the
  paragraph style `سطر اول جداول`, body cells `محتوای جداول`. Set `TableDirection = 0`
  (RTL) and `Rows.Item(1).HeadingFormat = -1`. Those styles are justify-low (8), which
  stretches words with kashida ("ارســـال"); plain justify (3) stretches spaces in mixed
  lines. Set body cells to `Alignment = 2`, `ReadingOrder = 0`. Widths in points:
  `Columns.Item(c).PreferredWidthType = 3; PreferredWidth = …` (text width ≈ 451pt).
  Row height: `Rows.Item(r).HeightRule = 1; Height = …` (at least) or `HeightRule = 0` (auto).
  For questionnaire answer columns, prefer auto height over tall (75pt) rows.
- **A parenthesised English phrase that wraps** inside RTL text scrambles its brackets
  ("( Log Source" / "Disconnection)"). Join the phrase with no-break spaces (U+00A0).
- **No list of tables?** Then skip captions; put a bold label paragraph above a sub-table
  with `KeepWithNext = -1` so it doesn't strand at a page bottom.

## Images

1. **Windows:** write the image anywhere local and use its normal path.
   **macOS + Parallels:** write it in a Mac folder that Parallels shares with
   Windows (e.g. `~/Downloads` → `\\Mac\Home\Downloads\…`, depending on the VM's
   sharing settings). Shared files can take seconds to appear in Windows, and an
   overwritten file may be served stale: **give every version a new file name**,
   and wait until Windows sees it:
   ```bash
   until prlctl exec "$WORD_VM" cmd /c 'if exist "\\Mac\Home\Downloads\img-v2.png" (echo yes)' | grep -q yes; do sleep 1; done
   ```
2. Insert inline: `$doc.InlineShapes.AddPicture('C:\path\img-v2.png', $false, $true, $range)`,
   then `LockAspectRatio = -1` and set `Width`. To swap a picture, read its
   position and width, `Delete()` it and add the new one there.

For an architecture diagram or figure: write it as HTML/SVG (formal, RTL, a
Persian font such as Vazirmatn) and render it with headless Chrome or Edge, A4 at 2×:

```bash
# macOS
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu \
  --hide-scrollbars --force-device-scale-factor=2 --window-size=1240,1754 \
  --virtual-time-budget=8000 --screenshot=out.png file:///path/diagram.html
```

```bat
rem Windows (Edge is always installed)
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=2 --window-size=1240,1754 --virtual-time-budget=8000 --screenshot=C:\path\out.png file:///C:/path/diagram.html
```

Look at the PNG yourself before inserting it (overlapping labels, arrows
crossing boxes).

## How the scripts work (for debugging)

**Windows** (`word.cmd` → `word.ps1`): runs `powershell.exe` 5.1 (PowerShell 7 has no
`Marshal.GetActiveObject`), writes prelude + snippet to `%TEMP%\word-live-<pid>.ps1`
with a BOM and runs it in the same process. `render.cmd` adds `render-win.ps1`,
which turns the exported PDF into PNGs with `Windows.Data.Pdf`; after that,
PowerShell 5.1 crashes on a normal shutdown, so `word.ps1` ends with
`Environment.Exit`. A stuck job is a `powershell.exe` running `word.ps1`; end it
with `taskkill /PID <pid> /F`.

**macOS** (`word`):

- The snippet reaches Windows through `prlctl exec` stdin, written to
  `C:\Users\Public\word-live-<pid>.ps1` as SYSTEM. The two helpers are copied the
  same way, once per version, as `C:\Users\Public\word-live-hidden-<hash>.vbs` and
  `word-live-collect-<hash>.ps1`.
- `hidden.vbs` runs the job in the user's session with no visible window (a plain
  `prlctl exec --current-user powershell` opens a Windows Terminal window), and
  writes `.started` and `.done` markers. `prlctl` does not wait for wscript.
- `collect.ps1` runs as SYSTEM (off screen), waits for `.done` and prints the
  output; the Mac side then deletes the job files.
- `prlctl` sometimes fails a call with "Invalid argument" (exit 255) even though
  it ran; those calls are retried, and the markers keep a retry from running a
  job twice. A stray `PrlJob_GetRetCode: Invalid argument` line on stderr is
  harmless when the job's output and exit code are fine.
- `prlctl exec` arguments longer than about 2,000 characters fail.
- Leftovers after a crash: `prlctl exec "$WORD_VM" cmd /c "dir /b C:\Users\Public\word-live-*"`
  (the `-hidden-`/`-collect-` helpers belong there); a stuck job shows up as a
  `powershell.exe` in session `Console` (`tasklist /fi "imagename eq powershell.exe"`)
  and can be ended with `taskkill /PID <pid> /F`.
- `sips -Z` on a `prlctl capture` PNG produces an all-black image; don't resize.
