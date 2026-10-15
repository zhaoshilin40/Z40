---
name: desktop-control
description: Control the Windows desktop through UI Automation — list windows, inspect control trees, click buttons, type text, read values, take screenshots. Use when a task requires operating a GUI application (clicking menus, filling forms, reading on-screen state, driving a browser or native app) instead of only running shell commands. Prefer this over raw cursor-coordinate simulation, which is unreliable on this machine.
---

# Desktop Control (Windows UI Automation)

Drive the real Windows desktop by **accessibility elements**, not by pixel coordinates.

## Why not coordinate clicking

Measured on this machine: `SetCursorPos` + `mouse_event` produce correct cursor
positions and foreground windows, yet **clicks land nowhere inside Edge**
(no response at all). The same happens in several Chromium hosts. Keyboard
injection (`keybd_event`) does work, and UI Automation works reliably.

Rule: **always locate a control first, then act on it.** Use coordinates only as
a last resort, and verify the effect by reading state back.

## Setup

- Interpreter: `C:\Users\12806\Python314\python.exe` (has `uiautomation` 2.0.29,
  Pillow, pywin32). No admin rights needed — pure ctypes over COM.
- Tool: `<skill-dir>/scripts/gui.py`
- Invoke it as:
  ```
  C:\Users\12806\Python314\python.exe <skill-dir>\scripts\gui.py <command> [options]
  ```
  Add `--json` for machine-readable output (recommended when parsing results).

## Commands

| Command | Purpose |
|---|---|
| `windows` | list top-level windows (pid, exe, size, class, title) |
| `tree --window X` | dump a window's control tree |
| `find --window X --name Y --type Z` | locate controls matching filters |
| `click --window X --name Y` | click a control (tries Invoke → Select → Expand → Legacy → physical click) |
| `type "text" --window X` | type text (supports Chinese) |
| `keys "{Ctrl}s" --window X` | send keystrokes |
| `read --window X [--type Document]` | read text via Value/Text/Legacy pattern |
| `shot -o out.png [--window X]` | screenshot screen or a single window |
| `activate --window X` | bring a window to the foreground |
| `scroll --window X --direction down --times 3` | scroll |
| `focus` | describe the currently focused control |

Common filters: `--name` (substring), `--exact`, `--type`, `--automation-id`,
`--depth`, `--limit`, `--index` (which match), `--visible-only`.

`--window` matches against title, class name, or executable name, case-insensitive.

**Raise `--limit` when you are looking for something specific.** The default is 20
and results are returned in tree order, so window chrome (minimize/maximize/close)
fills the list before any page content appears. Use
`--name "exact text" --limit 50` to jump straight to a known element.

## Standard workflow

1. **Look** — `windows` to see what is open, then `shot -o before.png` and read
   the image to understand the current state.
2. **Locate** — `find` or `tree` to get the exact control. Prefer `--automation-id`
   or an exact `--name` when several controls share a label.
3. **Act** — `click` / `type` / `keys`.
4. **Verify** — `read`, a fresh `shot`, or `find` again. Never assume an action
   landed; always confirm by reading state back.

## Examples

```powershell
$py = "C:\Users\12806\Python314\python.exe"
$g  = "<skill-dir>\scripts\gui.py"

& $py $g windows
& $py $g find --window Notepad --type MenuItem
& $py $g click --window Notepad --name "编辑"
& $py $g type "hello 你好" --window Notepad
& $py $g read --window Notepad --type Document
& $py $g shot --window Notepad -o shot.png
```

## Hard-won details

- **Always set a search timeout.** The script does this globally (3 s). Without it
  a missing control blocks forever.
- **`GetValuePattern()` does not exist.** Use
  `ctrl.GetPattern(auto.PatternId.ValuePattern).Value`, or
  `TextPattern` → `DocumentRange.GetText(-1)`. `gui.py read` already tries
  Value → Text → LegacyIAccessible in that order.
- **`HasKeyboardFocus` is a property, not a method** — calling it raises
  `TypeError: 'bool' object is not callable`.
- **Notepad's editor is a `DocumentControl`, not an `EditControl`.** Modern
  Windows 11 apps use Document/UIA text patterns. Search by `--type Document`
  when `--type Edit` finds nothing.
- **Windows may report negative `top` values** for maximized windows (e.g. -8).
  Clamp to 0 before using a rectangle as a screenshot bounding box; `gui.py`
  does this.
- **Chromium page content IS reachable — but only below the window chrome.**
  Verified on Edge: `find --window msedge --type Hyperlink --limit 30` returns the
  real page links with correct rectangles, and clicking them via
  `InvokePattern` navigates the browser for real. Two traps:
  1. Window-chrome buttons (minimize/maximize/close, back/forward/reload) sit at
     depth ~3 and will fill a default `--limit 20` before page content appears.
     Search by `--name` or raise `--limit` when hunting page elements.
  2. A page control's `Name` may be empty while its text lives in child
     `TextControl`s. If a button label is not found, search `--type Text` or
     `--type Hyperlink` instead.
  If a page still exposes nothing, launch the browser with
  `--force-renderer-accessibility`, or fall back to CDP (below).
- **Menu items open on Invoke but submenus may need a moment.** Wait ~0.5–1 s
  between opening a menu and reading its items.
- **Verify with pixels when the tree is ambiguous.** `shot` + reading the image
  is the ground truth for "did the click actually do something".
- **PowerShell output encoding**: the script re-wraps stdout as UTF-8, so Chinese
  window titles survive the pipe. Do not remove that wrapper.

## CDP fallback for Chromium page content

When you need the real DOM (forms, buttons inside a web page), start the browser
with remote debugging and talk to it over the DevTools Protocol:

```powershell
Start-Process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
  -ArgumentList "--remote-debugging-port=9222","--force-renderer-accessibility","https://example.com"
# then GET http://127.0.0.1:9222/json to find the page target and drive it
# with Runtime.evaluate / Page.captureScreenshot over the WebSocket endpoint.
```

Non-headless Edge is required for sites protected by bot detection; a real window
passes where headless is blocked.

## Safety

- Confirm the target window before typing — keystrokes go to whatever holds focus.
- Prefer `read` over `shot` when you only need text: it is exact and costs no image tokens.
- Take a `shot` before and after any destructive action (deleting, submitting,
  sending) so a mistake is visible and reversible.
