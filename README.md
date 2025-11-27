# CreateLVHover (AHK v2)

Explorer‑style **hover** for `SysListView32` that keeps the hot row in sync while the list scrolls under a stationary pointer (mouse wheel or precision touchpad). No overlays. Single shared manager timer. Multi‑ListView ready.

## Why this exists

Ever hovered a ListView and lost the highlight the moment you scroll?  
Ever clicked twice because the “hot row” slipped under your mouse?

**CreateLVHover** fixes that tiny, daily frustration. It keeps the hot row in sync while the list moves beneath a stationary pointer — mouse wheel or precision touchpad. No overlays, no hacks, no flicker. One shared manager timer. Works across multiple ListViews.

---

## Quick start

```ahk
#Requires AutoHotkey v2.0
#Include <CreateLVHover>

gui := Gui()
lv  := gui.Add("ListView", "-Hdr r15 w320", ["Name"]) ; any standard ListView
for v in ["Alpha","Beta","Gamma"]
    lv.Add("", v)

gui.Show()

hover := CreateLVHover(lv, {
    underline: true,    ; underline on hover (Explorer‑style)
    oneClick:  true,    ; one‑click activate + TRACKSELECT
    hoverMs:   1000,     ; time before TRACKSELECT (ms)
    timerMs:   30,      ; poll interval for "scroll under pointer"
    pxJiggle:  1,       ; minimal pointer nudge to trigger real WM_MOUSEMOVE
    scrollWindowMs: 150 ; grace period after scroll where jiggle is allowed
})
```

---

## Installation

Place `CreateLVHover.ahk` in **one** of the standard library locations so `#Include <CreateLVHover>` can find it:

1. **Local library** (next to your script)
   - `.<your script folder>\Lib\CreateLVHover.ahk`
2. **User library** (per‑user)
   - `%A_MyDocuments%\AutoHotkey\Lib\CreateLVHover.ahk`
3. **Standard library** (AutoHotkey install folder)
   - `...\AutoHotkey\Lib\CreateLVHover.ahk` (the `Lib` folder next to `AutoHotkey.exe`)

> Using **angle brackets** with `#Include <CreateLVHover>` tells AutoHotkey to search the library folders above.
>
> **Alternative:** Include by **explicit path** if you prefer:
>
> ```ahk
> #Include "C:\Path\To\CreateLVHover.ahk"
> ```

---

## API

**Constructor**

```ahk
CreateLVHover(target, options := {})
```

- `target` — the ListView to enhance. Accepts:
  - a **ListView control object** (recommended),
  - a **raw HWND** (Integer), or
  - a **WinTitle/ahk_id** string.

 The constructor strictly validates that the resolved window is a **`SysListView32`**. If it isn’t, a friendly message box explains the issue and the script exits. There is no child auto‑discovery to avoid picking the wrong ListView when multiple exist in the same window.

- `options` — all fields optional:

| Option            | Type    | Default | Description                                                                                                  |
| ----------------- | ------- | ------- | ------------------------------------------------------------------------------------------------------------ |
| `underline`       | Boolean | `true`  | Show underline on hot items (`LVS_EX_UNDERLINEHOT`).                                                         |
| `oneClick`        | Boolean | `true`  | One‑click activation + `LVS_EX_TRACKSELECT`. If `false`, behaves like double‑click activation (no auto‑select). |
| `hoverMs`         | Integer | `1000`  | Hover time before TRACKSELECT triggers (only relevant if `oneClick:true`).                                   |
| `timerMs`         | Integer | `30`    | Polling interval used to detect “scroll under pointer.”                                                       |
| `pxJiggle`        | Integer | `1`     | Minimal pointer movement (px) to trigger a real `WM_MOUSEMOVE`.                                              |
| `scrollWindowMs`  | Integer | `150`   | Grace period after a scroll message during which jiggle is allowed.  
| `scrollWindowMsPointer`  | Integer | `220`   | slightly longer grace period for precision touchpad / gesture events.             |.

**Instance methods**

- `UpdateNow()` — Force a single poll of the active instance.
- `IsActive()` — Returns `true` while this instance is the **active** one under the pointer.

**Status helper**

- `CreateLVHover.GetStatusArray()` → returns an array of objects with per‑instance info:

```ahk
[
  { hwnd: 0x123456, active: true/false, visibleEnabled: true/false, lastIdx: N },
  ...
]
```

> Create **one instance per ListView**: `h1 := CreateLVHover(lv1)`, `h2 := CreateLVHover(lv2)`, etc.

---

## How it works (in short)

1. **Explorer styling** — The class applies `SetWindowTheme("Explorer")` and extended styles (e.g., `LVS_EX_UNDERLINEHOT`, `LVS_EX_TRACKSELECT`, `LVS_EX_ONECLICKACTIVATE`) so the ListView behaves like Windows Explorer.
2. **Single manager timer** — A class‑level timer figures out which ListView the pointer is currently over and calls `_Poll()` **only** on that instance.
3. **Scroll hooks (class‑level)** — On scroll/gesture messages, the class resets the hover timer (`LVM_SETHOVERTIME(hoverMs)` + `_RequestHover()`), opens a short **grace window**, and performs a 1‑px “micro‑jiggle” (**unless at top/bottom**) to provoke a real `WM_MOUSEMOVE` without visible cursor stepping.
4. **Hit‑testing** — `_Poll()` uses `LVM_SUBITEMHITTEST` to find the row under the pointer and updates the hot state. When nothing changes (same row, same mouse position, grace window elapsed), it returns early.
5. **No overlays** — The highlight is **native**, coming from the OS theme; the class does not draw custom rectangles.

---

## Notes & limitations

- **Hover colors** come from the OS theme. To customize hover colors per row/cell, use `owner‑draw` or an overlay.
- Works with **all standard ListView views** (report, list, small/large icons) that use the common control class `SysListView32`.
- The class avoids jiggle at the **top/bottom** boundary of the list (no jiggle if a scroll wouldn’t move content).
- Designed to be **lightweight**: one shared timer, fast routing via an internal HWND→instance map, minimal allocations per tick.

---

## Troubleshooting

- **Library not found**

  - Ensure the file name is exactly `CreateLVHover.ahk` and is placed in one of the **library folders** listed under *Installation*.

- **Invalid target**

  - The constructor validates that the target resolves to a `SysListView32`. If not, a friendly diagnostic message is shown and the script stops.

- **No underline / no hover**

  - Make sure `underline: true` and `oneClick: true` (which enables TRACKSELECT). You can tweak `hoverMs` to your liking.

- **Hover doesn’t update during scrolling**
  - Increase `scrollWindowMs` (e.g., 200–300) or reduce `timerMs` (e.g., 20).
  - Ensure the pointer is actually over the ListView while scrolling — the grace window only opens then.

- **Multiple ListViews**
  - Create **one instance per ListView**: `h1 := CreateLVHover(lv1)`, `h2 := CreateLVHover(lv2)`, etc. The class manages a single shared timer internally.

## File layout examples

**Project‑local library (recommended for per‑project isolation):**

```
MyApp\
├─ MyApp.ahk
└─ Lib\
   └─ CreateLVHover.ahk
```

Then in `MyApp.ahk`:

```ahk
#Include <CreateLVHover>
```

**User library (available to all scripts for your user account):**

```
%A_MyDocuments%\AutoHotkey\Lib\CreateLVHover.ahk
```

**Standard library (available to all scripts on the machine):**

```
<AutoHotkey install folder>\Lib\CreateLVHover.ahk
```

**Direct path include (alternative to library folders):**

```ahk
#Include "C:\Path\To\CreateLVHover.ahk"
```

---

## Flowchart 

![Flowdiagram ListView Hover-2025-11-27-151041](https://github.com/user-attachments/assets/c1497842-fec9-4a85-9502-1b6a16e7de96)

---

## License

This project is licensed under the terms described in the accompanying **LICENSE** file.

---

## Credits

- Tested on **Windows 11** with **AutoHotkey v2**.
- Uses only public Windows APIs and standard AHK directives.
- Thanks to the Windows common controls documentation and community examples for the underlying message constants and techniques.

