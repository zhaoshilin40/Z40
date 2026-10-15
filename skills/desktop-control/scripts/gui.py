#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
gui.py - Windows desktop control via UI Automation.

Reliable alternative to coordinate-based mouse simulation: locates real UI
elements by name/type and drives them through accessibility patterns.

Usage examples:
  gui.py windows
  gui.py tree --window Notepad --depth 3
  gui.py find --window Notepad --type Button
  gui.py click --window Notepad --name "保存"
  gui.py type "hello world"
  gui.py read --window Notepad
  gui.py shot -o screen.png
  gui.py activate --window Notepad

All commands accept --json for machine-readable output.
"""

import argparse
import io
import json
import sys
import time

# Force UTF-8 on stdout so Chinese window titles survive the PowerShell pipe.
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import uiautomation as auto  # noqa: E402

DEFAULT_TIMEOUT = 3.0


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def _rect(r):
    return {"left": r.left, "top": r.top, "right": r.right, "bottom": r.bottom,
            "width": r.right - r.left, "height": r.bottom - r.top}


def _describe(c, with_rect=True):
    try:
        d = {
            "type": c.ControlTypeName,
            "name": c.Name or "",
            "class": c.ClassName or "",
            "automation_id": c.AutomationId or "",
            "pid": c.ProcessId,
            "enabled": c.IsEnabled,
            "offscreen": c.IsOffscreen,
        }
        if with_rect:
            d["rect"] = _rect(c.BoundingRectangle)
        return d
    except Exception as e:
        return {"type": "?", "name": "", "error": str(e)}


def _top_windows():
    out = []
    for w in auto.GetRootControl().GetChildren():
        try:
            if w.ControlTypeName == "WindowControl":
                out.append(w)
        except Exception:
            continue
    return out


def _proc_name(pid):
    try:
        import ctypes
        from ctypes import wintypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        k = ctypes.windll.kernel32
        h = k.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return ""
        buf = ctypes.create_unicode_buffer(260)
        size = wintypes.DWORD(260)
        ok = k.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
        k.CloseHandle(h)
        return buf.value.split("\\")[-1] if ok else ""
    except Exception:
        return ""


def find_window(spec):
    """Match a top-level window by substring against title, class or exe name."""
    if not spec:
        return None
    s = spec.lower()
    wins = _top_windows()
    # exact title first, then substring
    for w in wins:
        try:
            if (w.Name or "").lower() == s:
                return w
        except Exception:
            pass
    for w in wins:
        try:
            if s in (w.Name or "").lower() or s in (w.ClassName or "").lower():
                return w
        except Exception:
            pass
    for w in wins:
        try:
            if s in _proc_name(w.ProcessId).lower():
                return w
        except Exception:
            pass
    return None


def _walk(ctrl, max_depth, max_nodes, depth=0, counter=None):
    if counter is None:
        counter = [0]
    results = []
    if depth > max_depth or counter[0] >= max_nodes:
        return results
    try:
        children = ctrl.GetChildren()
    except Exception:
        return results
    for ch in children:
        if counter[0] >= max_nodes:
            break
        counter[0] += 1
        results.append((depth, ch))
        results.extend(_walk(ch, max_depth, max_nodes, depth + 1, counter))
    return results


def _matches(ctrl, args):
    try:
        if args.type and args.type.lower() not in ctrl.ControlTypeName.lower():
            return False
        if args.name and args.name.lower() not in (ctrl.Name or "").lower():
            return False
        if args.exact and args.name and (ctrl.Name or "") != args.name:
            return False
        if args.automation_id and args.automation_id != (ctrl.AutomationId or ""):
            return False
        if args.visible_only and ctrl.IsOffscreen:
            return False
        return True
    except Exception:
        return False


def _collect(root, args):
    """Breadth-first collect of descendants matching the filters.

    Adaptive depth: walk progressively deeper and keep going until we either
    fill --limit or run out of levels.  Stopping at the *first* level that
    produced a hit was wrong for Chromium windows: window-chrome buttons live
    at depth ~3 and would shadow the real page controls much deeper down.
    """
    limit = args.limit
    if args.depth is not None:
        found = []
        for depth, c in _walk(root, args.depth, max(limit * 40, 200)):
            if _matches(c, args):
                found.append((depth, c))
        return found[:limit]

    found = []
    seen = set()
    for d in (3, 6, 10, 16, 24):
        for depth, c in _walk(root, d, max(limit * 40, 400)):
            if len(found) >= limit:
                return found
            if depth > (3 if d == 3 else (6 if d == 6 else (10 if d == 10 else (16 if d == 16 else 99)))):
                continue  # already visited at a shallower pass
            try:
                key = (c.ControlType, c.Name, c.BoundingRectangle.left,
                       c.BoundingRectangle.top, c.NativeWindowHandle)
            except Exception:
                key = None
            if key is not None and key in seen:
                continue
            if key is not None:
                seen.add(key)
            if _matches(c, args):
                found.append((depth, c))
    return found[:limit]


def _window_of(ctrl):
    try:
        return ctrl.GetTopLevelControl()
    except Exception:
        return None


def _invoke(ctrl):
    """Try the most direct interaction pattern, fall back to a real click."""
    attempts = []
    # 1. Invoke pattern - no mouse movement, works on buttons/links/menus
    try:
        p = ctrl.GetPattern(auto.PatternId.InvokePattern)
        if p:
            p.Invoke()
            return "InvokePattern"
    except Exception as e:
        attempts.append(f"invoke: {e}")
    # 2. SelectionItem - for tabs, list items
    try:
        p = ctrl.GetPattern(auto.PatternId.SelectionItemPattern)
        if p:
            p.Select()
            return "SelectionItemPattern"
    except Exception as e:
        attempts.append(f"select: {e}")
    # 3. ExpandCollapse - for combos, trees
    try:
        p = ctrl.GetPattern(auto.PatternId.ExpandCollapsePattern)
        if p:
            p.Expand()
            return "ExpandCollapsePattern"
    except Exception as e:
        attempts.append(f"expand: {e}")
    # 4. Legacy IAccessible doDefaultAction - old apps
    try:
        p = ctrl.GetPattern(auto.PatternId.LegacyIAccessiblePattern)
        if p:
            p.DoDefaultAction()
            return "LegacyIAccessiblePattern"
    except Exception as e:
        attempts.append(f"legacy: {e}")
    # 5. Physical click as last resort
    try:
        ctrl.Click(simulateMove=False)
        return "physical-click"
    except Exception as e:
        attempts.append(f"click: {e}")
    raise RuntimeError("; ".join(attempts) or "no interaction pattern available")


# --------------------------------------------------------------------------
# commands
# --------------------------------------------------------------------------

def cmd_windows(args):
    rows = []
    for w in _top_windows():
        d = _describe(w)
        d["exe"] = _proc_name(w.ProcessId)
        rows.append(d)
    rows.sort(key=lambda r: (r["rect"]["top"], r["rect"]["left"]))
    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        for r in rows:
            print(f"[{r['pid']:6}] {r['exe']:24} {r['rect']['width']:4}x{r['rect']['height']:<4} "
                  f"{r['class']:22} {r['name'][:60]}")
    return 0


def cmd_tree(args):
    win = find_window(args.window) if args.window else auto.GetRootControl()
    if win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    nodes = _walk(win, args.depth, args.limit)
    if args.type or args.name:
        nodes = [(d, c) for d, c in nodes if _matches(c, args)]
    rows = []
    for depth, c in nodes:
        d = _describe(c, with_rect=args.rect)
        d["depth"] = depth
        rows.append(d)
    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        for r in rows:
            pad = "  " * r["depth"]
            rect = r.get("rect", {})
            pos = f"[{rect.get('left',0):5},{rect.get('top',0):5}] " if rect else ""
            print(f"{pad}{r['type']:20} {pos}{(r['name'] or '')[:55]}")
    return 0


def cmd_find(args):
    win = find_window(args.window) if args.window else auto.GetRootControl()
    if win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    found = _collect(win, args)
    rows = []
    for depth, c in found:
        d = _describe(c)
        d["depth"] = depth
        rows.append(d)
    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        if not rows:
            print("(no match)")
        for r in rows:
            rect = r["rect"]
            print(f"{r['type']:20} [{rect['left']:5},{rect['top']:5}] {r['name'][:60]}")
    return 0 if rows else 2


def cmd_click(args):
    win = find_window(args.window) if args.window else auto.GetRootControl()
    if win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    found = _collect(win, args)
    if not found:
        print(json.dumps({"error": "no matching control", "window": args.window,
                          "name": args.name, "type": args.type}, ensure_ascii=False))
        return 2
    if args.index >= len(found):
        print(json.dumps({"error": f"index {args.index} out of range ({len(found)} matches)"},
                         ensure_ascii=False))
        return 2
    depth, target = found[args.index]
    info = _describe(target)
    try:
        win.SetActive()
        time.sleep(0.2)
    except Exception:
        pass
    method = _invoke(target)
    time.sleep(args.wait)
    info["clicked_via"] = method
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print(f"clicked {info['type']} {info['name']!r} via {method}")
    return 0


def cmd_type(args):
    win = None
    if args.window:
        win = find_window(args.window)
        if win is None:
            print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
            return 1
        win.SetActive()
        time.sleep(0.3)
    target = win if win else auto.GetFocusedControl()
    if args.clear:
        target.SendKeys("{Ctrl}a", waitTime=0.05)
        time.sleep(0.15)
        target.SendKeys("{Delete}", waitTime=0.05)
        time.sleep(0.15)
    # Escape brace characters so SendKeys treats them literally
    text = args.text.replace("{", "{{").replace("}", "}}")
    target.SendKeys(text, waitTime=args.interval)
    time.sleep(args.wait)
    if args.json:
        print(json.dumps({"typed": args.text, "window": args.window}, ensure_ascii=False))
    else:
        print(f"typed {len(args.text)} chars" + (f" into {args.window!r}" if args.window else ""))
    return 0


def cmd_keys(args):
    win = None
    if args.window:
        win = find_window(args.window)
        if win is None:
            print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
            return 1
        win.SetActive()
        time.sleep(0.3)
    target = win if win else auto.GetFocusedControl()
    target.SendKeys(args.keys, waitTime=0.05)
    time.sleep(args.wait)
    if args.json:
        print(json.dumps({"keys": args.keys}, ensure_ascii=False))
    else:
        print(f"sent keys {args.keys!r}")
    return 0


def cmd_read(args):
    win = find_window(args.window) if args.window else auto.GetFocusedControl()
    if win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    target = win
    if args.name or args.type or args.automation_id:
        found = _collect(win, args)
        if not found:
            print(json.dumps({"error": "no matching control"}, ensure_ascii=False))
            return 2
        target = found[args.index][1]

    text = None
    source = None
    for pattern_name, getter in (
        ("ValuePattern", lambda c: c.GetPattern(auto.PatternId.ValuePattern).Value),
        ("TextPattern", lambda c: c.GetPattern(auto.PatternId.TextPattern).DocumentRange.GetText(-1)),
        ("LegacyIAccessiblePattern", lambda c: c.GetPattern(auto.PatternId.LegacyIAccessiblePattern).Value),
    ):
        try:
            v = getter(target)
            if v is not None:
                text, source = v, pattern_name
                break
        except Exception:
            continue

    if text is None and args.name:
        # window title fallback
        text, source = (target.Name or ""), "Name"

    if args.json:
        print(json.dumps({"text": text, "source": source, "control": _describe(target)},
                         ensure_ascii=False, indent=2))
    else:
        print(text if text is not None else "(no readable text)")
    return 0 if text is not None else 2


def cmd_shot(args):
    try:
        from PIL import ImageGrab
    except ImportError:
        print(json.dumps({"error": "Pillow not installed"}, ensure_ascii=False))
        return 1
    win = find_window(args.window) if args.window else None
    if args.window and win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    if win is not None:
        try:
            win.SetActive()
            time.sleep(0.4)
        except Exception:
            pass
        r = win.BoundingRectangle
        box = (max(r.left, 0), max(r.top, 0), r.right, r.bottom)
        img = ImageGrab.grab(bbox=box)
    else:
        img = ImageGrab.grab(all_screens=args.all_screens)
    if args.scale and args.scale != 1.0:
        img = img.resize((int(img.width * args.scale), int(img.height * args.scale)))
    img.save(args.output)
    info = {"output": args.output, "width": img.width, "height": img.height}
    if args.json:
        print(json.dumps(info, ensure_ascii=False))
    else:
        print(f"saved {args.output} ({img.width}x{img.height})")
    return 0


def cmd_activate(args):
    win = find_window(args.window)
    if win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    try:
        win.SetActive()
    except Exception:
        try:
            win.SetFocus()
        except Exception as e:
            print(json.dumps({"error": str(e)}, ensure_ascii=False))
            return 2
    time.sleep(args.wait)
    info = _describe(win)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print(f"activated {info['name']!r}")
    return 0


def cmd_scroll(args):
    win = find_window(args.window) if args.window else None
    if args.window and win is None:
        print(json.dumps({"error": f"window not found: {args.window}"}, ensure_ascii=False))
        return 1
    target = win if win else auto.GetFocusedControl()
    if args.direction == "down":
        for _ in range(args.times):
            target.WheelDown(wheelTimes=1, waitTime=0.1)
    else:
        for _ in range(args.times):
            target.WheelUp(wheelTimes=1, waitTime=0.1)
    if args.json:
        print(json.dumps({"scrolled": args.direction, "times": args.times}, ensure_ascii=False))
    else:
        print(f"scrolled {args.direction} x{args.times}")
    return 0


def cmd_focus_info(args):
    c = auto.GetFocusedControl()
    info = _describe(c) if c else {"error": "no focused control"}
    if c is not None:
        w = _window_of(c)
        if w is not None:
            info["window"] = {"name": w.Name, "class": w.ClassName, "pid": w.ProcessId}
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    return 0


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _add_json(p):
    """Allow --json after the subcommand too.

    default=SUPPRESS keeps the subparser from overwriting a True already set by
    the top-level parser (gui.py --json windows).
    """
    p.add_argument("--json", action="store_true", default=argparse.SUPPRESS,
                   help="machine-readable output")


def _add_filters(p):
    _add_json(p)
    p.add_argument("--name", help="substring match on control name")
    p.add_argument("--exact", action="store_true", help="exact name match")
    p.add_argument("--type", help="control type substring, e.g. Button, Edit, Document, MenuItem")
    p.add_argument("--automation-id", dest="automation_id", help="exact AutomationId")
    p.add_argument("--depth", type=int, default=None, help="max tree depth (default: adaptive)")
    p.add_argument("--limit", type=int, default=20, help="max matches / nodes")
    p.add_argument("--index", type=int, default=0, help="which match to use")
    p.add_argument("--visible-only", action="store_true", help="skip offscreen controls")


def build_parser():
    p = argparse.ArgumentParser(description="Windows desktop control via UI Automation")
    p.add_argument("--json", action="store_true", help="machine-readable output")
    sub = p.add_subparsers(dest="cmd", required=True)

    w = sub.add_parser("windows", help="list top-level windows")
    _add_json(w)
    w.set_defaults(func=cmd_windows)

    t = sub.add_parser("tree", help="dump the control tree of a window")
    t.add_argument("--window", help="window title / class / exe substring")
    t.add_argument("--rect", action="store_true", help="include rectangles")
    _add_filters(t)
    t.set_defaults(func=cmd_tree)

    f = sub.add_parser("find", help="find controls matching filters")
    f.add_argument("--window", help="window title / class / exe substring")
    _add_filters(f)
    f.set_defaults(func=cmd_find)

    c = sub.add_parser("click", help="click a control by name/type")
    c.add_argument("--window", help="window title / class / exe substring")
    c.add_argument("--wait", type=float, default=0.5, help="seconds to wait after clicking")
    _add_filters(c)
    c.set_defaults(func=cmd_click)

    ty = sub.add_parser("type", help="type text into a window or the focused control")
    _add_json(ty)
    ty.add_argument("text", help="text to type")
    ty.add_argument("--window", help="target window")
    ty.add_argument("--clear", action="store_true", help="select-all + delete first")
    ty.add_argument("--interval", type=float, default=0.01, help="per-char delay")
    ty.add_argument("--wait", type=float, default=0.5, help="seconds to wait after typing")
    ty.set_defaults(func=cmd_type)

    k = sub.add_parser("keys", help="send keystrokes, e.g. '{Ctrl}s' or '{Enter}'")
    _add_json(k)
    k.add_argument("keys", help="SendKeys syntax")
    k.add_argument("--window", help="target window")
    k.add_argument("--wait", type=float, default=0.5)
    k.set_defaults(func=cmd_keys)

    r = sub.add_parser("read", help="read text out of a control")
    r.add_argument("--window", help="window title / class / exe substring")
    _add_filters(r)
    r.set_defaults(func=cmd_read)

    s = sub.add_parser("shot", help="screenshot the screen or a window")
    _add_json(s)
    s.add_argument("-o", "--output", default="screen.png")
    s.add_argument("--window", help="capture just this window")
    s.add_argument("--scale", type=float, default=1.0, help="resize factor")
    s.add_argument("--all-screens", dest="all_screens", action="store_true")
    s.set_defaults(func=cmd_shot)

    a = sub.add_parser("activate", help="bring a window to the foreground")
    _add_json(a)
    a.add_argument("--window", required=True)
    a.add_argument("--wait", type=float, default=0.4)
    a.set_defaults(func=cmd_activate)

    sc = sub.add_parser("scroll", help="scroll a window")
    _add_json(sc)
    sc.add_argument("--window")
    sc.add_argument("--direction", choices=["up", "down"], default="down")
    sc.add_argument("--times", type=int, default=3)
    sc.set_defaults(func=cmd_scroll)

    fi = sub.add_parser("focus", help="describe the currently focused control")
    _add_json(fi)
    fi.set_defaults(func=cmd_focus_info)

    return p


def main():
    args = build_parser().parse_args()
    auto.SetGlobalSearchTimeout(DEFAULT_TIMEOUT)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as e:
        print(json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main())
