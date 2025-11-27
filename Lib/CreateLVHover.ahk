; ============================================================================
; Class: CreateLVHover (AHK v2)
; Purpose: Give a SysListView32 the Windows Explorer look + hover underline,
;          and make the underline follow the row under the pointer while the
;          content scrolls (even if the mouse itself is stationary).
; Technique:
;   - SetWindowTheme("Explorer") + extended styles (TRACKSELECT/UNDERLINEHOT)
;   - Lightweight polling timer (~30 ms)
;   - "Micro‑jiggle" (1 px) ONLY when the list is scrolled under a stationary
;     pointer (never during normal mouse movement). Triggers a real WM_MOUSEMOVE.
;   - Cache the ListView rectangle once per poll (avoid duplicate GetWindowRect).
; Notes: No globals. Single class-level hook. Guarded DllCalls.
; ============================================================================

/**
 * CreateLVHover(target [, opts := {}])
 *
 * @description
 * CreateLVHover — Explorer‑style hover for SysListView32 (AHK v2).
 * Adds the Windows Explorer theme to a ListView so items highlight on
 * mouseover and also when the list scrolls under a stationary pointer.
 *
 * @requires AutoHotkey v2, Windows 10/11
 *
 * @param {Gui.Control|Integer|String} target
 * The target ListView to apply the Explorer theme to. Accept either a ListView
 * control object, a raw HWND (Integer), or a WinTitle (String).
 *
 * @param {Boolean} [underline=true]
 * Underlines hot items (LVS_EX_UNDERLINEHOT). Default is true.
 *
 * @param {Boolean} [oneClick=true]
 * When true, you only need to click once to activate the item in the ListView. 
 * Hot track selection is also enabled when true. 
 * Hot track selection means that an item is automatically selected 
 * when the cursor remains over the item for a certain period of time. 
 * The delay can be changed from "hoverMs".
 * When false, double click is enabled, meaning single click to select an item 
 * and double click to activate the item.
 * Default is True.
 *
 * @param {Integer} [hoverMs=1000]
 * Hover timeout (ms) before auto‑select when one‑click is enabled. Default is
 * 1000 ms.
 *
 * @param {Integer} [timerMs=30]
 * Polling interval (ms) while the pointer is over the ListView. Default is
 * 30 ms.
 *
 * @param {Integer} [pxJiggle=1]
 * Distance (px) for the post‑scroll micro‑jiggle to trigger a real
 * WM_MOUSEMOVE. Default is 1 px.
 *
 * @param {Integer} [scrollWindowMs=150]
 * Time window (ms) after scrolling during which a micro‑jiggle is allowed.
 * Default is 150 ms.
 *
 * @methods
 *   Dispose()   → Uninstall message hooks and stop the timer.
 *   UpdateNow() → Force a single poll.
 */
class CreateLVHover {
    ; ---- Class-wide (multi-LV manager) ----
    static _mgrPollCb   := 0
    static _mgrTimerMs  := 30
    static _mgrActive   := 0
    static _hwndMap     := Map()       ; HWND → instance (fast routing)
    static _boundGuis   := Map()       ; GUI.hWnd → true (bind Close kun én gang pr. GUI)
    static _classMsgCb  := 0
    static _classMsgCbMove := 0
    static _classHooksInstalled := false
    static _subclassedInstalled := false
    static _libFile := A_LineFile      ; library file to filter from stack
    static _libDir  := ""              ; computed lazily
    static _mainFile := A_ScriptFullPath

    ; ---- ListView / LVM constants ----
    static LVM_SETEXTENDEDLISTVIEWSTYLE := 0x1036
    static LVM_SETHOVERTIME             := 0x1047
    static LVM_SUBITEMHITTEST           := 0x1039
    static LVM_SETHOTITEM               := 0x103A

    ; ---- Extended styles ----
    static LVS_EX_DOUBLEBUFFER     := 0x00010000
    static LVS_EX_FULLROWSELECT    := 0x00000020
    static LVS_EX_TRACKSELECT      := 0x00000008
    static LVS_EX_UNDERLINEHOT     := 0x00000800
    static LVS_EX_ONECLICKACTIVATE := 0x00000040
    static LVS_EX_TWOCLICKACTIVATE := 0x00000080

    ; ---- Instance fields ----
    hLV                  := 0
    hoverMs              := 1000   ; time before TRACKSELECT (ms)
    oneClick             := true
    underline            := true
    pollMs               := 30     ; timer interval (ms)
    jigglePx             := 1      ; how far to move the cursor (px)
    scrollWindowMs       := 150    ; time window after scroll where jiggle is allowed
    scrollWindowMsPointer:= 220    ; grace window for pointer/gesture scroll (ms)
    _isActive            := false  ; true while this instance is being polled

    ; Jiggle handshake (move-out → wait for WM_MOUSEMOVE → move-back)
    _jiggleArmed := false
    _jiggleOrigX := 0
    _jiggleOrigY := 0
    _hTitle      := ""     ; cached WinTitle string: "ahk_id " . hLV
    _lastIdx := -9999
    ; Mouse/scroll state
    _mx := -1
    _my := -1
    _scrollUntil := 0      ; A_TickCount deadline after a scroll event

    ; Cached rect (updated once per poll)
    _rcL := 0
    _rcT := 0
    _rcR := 0
    _rcB := 0

    ; Reusable buffers (avoid allocations each tick)
    _bufRect := 0
    _bufPt   := 0
    _bufHti  := 0

    __New(target, opts := {}) {
        ; Capture caller file automatically (without wrapper):
        ; take first stack frame outside the library path.
        try {
            throw Error("CreateLVHover capture stack")
        } catch as e {
            frame := this._findCallsiteFrame(e)
            cFile := (ObjHasOwnProp(frame, "File") && frame.File) ? frame.File : ""
            this._callerFile := this._basename(cFile)
        }

        ; Resolve HWND from control/HWND/WinTitle
        hwnd := this._resolveHwnd(target)
        ; Ensure it's a ListView (optionally try child lookup)
        this.hLV := this._ensureListView(hwnd, target)
        this._hTitle := "ahk_id " this.hLV

        ; Options
        this.hoverMs        := (opts.HasOwnProp("hoverMs"))        ? Max(10, opts.hoverMs)          : this.hoverMs
        this.oneClick       := (opts.HasOwnProp("oneClick"))       ? !!opts.oneClick       : this.oneClick
        this.underline      := (opts.HasOwnProp("underline"))      ? !!opts.underline      : this.underline
        this.pollMs         := (opts.HasOwnProp("timerMs"))        ? Max(10, opts.timerMs) : this.pollMs
        this.jigglePx       := (opts.HasOwnProp("pxJiggle"))       ? Max(1,  opts.pxJiggle): this.jigglePx
        this.scrollWindowMs := (opts.HasOwnProp("scrollWindowMs")) ? Max(50, opts.scrollWindowMs) : this.scrollWindowMs

        ; Buffers
        this._bufRect := Buffer(16, 0)
        this._bufPt   := Buffer(8, 0)
        this._bufHti  := Buffer((A_PtrSize = 8) ? 48 : 24, 0)

        ; Apply theme + styles (guarded)
        this._applyTheme()

        ; Map HWND → instance for fast routing
        CreateLVHover._hwndMap[this.hLV] := this

        ; Install class-level hooks (only once per process)
        CreateLVHover._ClassInstallHooks()

        ; ONE Close-handler per GUI
        if (IsObject(target) && target.HasProp("Gui") && IsObject(target.Gui)) {
            g := target.Gui
            this._guiHwnd := g.Hwnd   ; ← save which GUI this instance belongs to
            if !CreateLVHover._boundGuis.Has(g.Hwnd) {
                CreateLVHover._boundGuis[g.Hwnd] := true
                g.OnEvent("Close", (*) => CreateLVHover.DisposeByGui(g))
            }
        }

        ; Register instance and ensure the single class timer is running
        CreateLVHover._ManagerRecomputePeriod()
        CreateLVHover._ManagerEnsureStarted()
    }

    ; Batch-dispose: remove ALL ListViews that belong to a specific GUI
    static DisposeByGui(guiOrHwnd) {
        try {
            guiHwnd := IsObject(guiOrHwnd) ? guiOrHwnd.Hwnd : guiOrHwnd
            if !guiHwnd
                return

            ; Remove bound marker
            try CreateLVHover._boundGuis.Delete(guiHwnd)

            ; Collect and remove all instances whose top-level ancestor == guiHwnd
            removed := 0
            for hwnd, inst in CreateLVHover._hwndMap.Clone() {
                ; Safe: also works when the window is +Owner
                if (inst.HasProp("_guiHwnd") && inst._guiHwnd = guiHwnd) {
                    try CreateLVHover._hwndMap.Delete(hwnd)
                    if (CreateLVHover._mgrActive = inst)
                        CreateLVHover._mgrActive := 0
                    removed++
                }
            }

            ; Stop timers/hooks if none left; otherwise recompute period
            if (CreateLVHover._hwndMap.Count = 0) {
                CreateLVHover._ManagerMaybeStop()
                CreateLVHover._ClassUninstallHooks()
            } else if (removed > 0) {
                CreateLVHover._ManagerRecomputePeriod()
            }
        } catch {
            ; swallow cleanup errors
        }
    }


    ; ---------------- Internal: theme/styles ----------------
    _applyTheme() {
        try DllCall("UxTheme.dll\SetWindowTheme", "ptr", this.hLV, "wstr", "Explorer", "ptr", 0)

        styles := CreateLVHover.LVS_EX_DOUBLEBUFFER
                | CreateLVHover.LVS_EX_FULLROWSELECT
        if (this.underline)
            styles |= CreateLVHover.LVS_EX_UNDERLINEHOT
        styles |= this.oneClick
            ? (CreateLVHover.LVS_EX_ONECLICKACTIVATE | CreateLVHover.LVS_EX_TRACKSELECT)
            :  CreateLVHover.LVS_EX_TWOCLICKACTIVATE

        try SendMessage(CreateLVHover.LVM_SETEXTENDEDLISTVIEWSTYLE, styles, styles,, this._hTitle)
        try SendMessage(CreateLVHover.LVM_SETHOVERTIME, 0, this.hoverMs,, this._hTitle)
    }

    ; ---------------- Class-level message hooks ----------------
    static _ClassInstallHooks() {
        if CreateLVHover._classHooksInstalled
            return
        CreateLVHover._classMsgCb := ObjBindMethod(CreateLVHover, "_ClassOnScrollMsg")
        for msg in [0x020A, 0x020E, 0x024E, 0x024F, 0x0119] ; WM_MOUSEWHEEL, WM_MOUSEHWHEEL, POINTER varianter, GESTURE
            OnMessage(msg, CreateLVHover._classMsgCb)
        ; additionally listen for WM_MOUSEMOVE to complete the jiggle handshake
        CreateLVHover._classMsgCbMove := ObjBindMethod(CreateLVHover, "_ClassOnMouseMove")
        OnMessage(0x0200, CreateLVHover._classMsgCbMove)
        CreateLVHover._classHooksInstalled := true

        ; Keyboard hotkeys only when focus is in a mapped LV
        HotIf(CreateLVHover._FocusIsMappedLV)

        Hotkey("~Up",    (*) => CreateLVHover._OnKeyNav("U"), "On")
        Hotkey("~Down",  (*) => CreateLVHover._OnKeyNav("D"), "On")
        Hotkey("~PgUp",  (*) => CreateLVHover._OnKeyNav("U"), "On")
        Hotkey("~PgDn",  (*) => CreateLVHover._OnKeyNav("D"), "On")
        Hotkey("~Home",  (*) => CreateLVHover._OnKeyNav("U"), "On")
        Hotkey("~End",   (*) => CreateLVHover._OnKeyNav("D"), "On")
        ; Valgfrit:
        Hotkey("~Left",  (*) => CreateLVHover._OnKeyNav("L"), "On")
        Hotkey("~Right", (*) => CreateLVHover._OnKeyNav("R"), "On")

        HotIf() ; reset
    }

    static _ClassUninstallHooks() {
        if !CreateLVHover._classHooksInstalled
            return
        try for msg in [0x020A, 0x020E, 0x024E, 0x024F, 0x0119] ; WM_MOUSEWHEEL, WM_MOUSEHWHEEL, POINTER varianter, GESTURE
            OnMessage(msg, CreateLVHover._classMsgCb, 0)
        try OnMessage(0x0200, CreateLVHover._classMsgCbMove, 0)
        CreateLVHover._classHooksInstalled := false
    }

    static _ClassOnScrollMsg(wParam, lParam, msg, hwnd) {
        ; Route scroll/gesture messages to the instance under the pointer (or directly via hwnd)
        CoordMode "Mouse", "Screen"
        MouseGetPos &sx, &sy
        current := A_TickCount
        isPointer := (msg = 0x024E || msg = 0x024F || msg = 0x0119)

        target := 0
        ; 1) Fast path: direct HWND routing if available
        if (hwnd && CreateLVHover._hwndMap.Has(hwnd)) {
            obj := CreateLVHover._hwndMap[hwnd]
            if IsObject(obj) && obj._IsVisibleEnabled() {
                obj._updateRectCache()
                target := obj
            }
        }
        ; 2) Fallback: find LV under mouse pointer
        if !target {
            for h, obj in CreateLVHover._hwndMap {
                if !obj._IsVisibleEnabled()
                    continue
                obj._updateRectCache()
                if (sx >= obj._rcL && sx < obj._rcR && sy >= obj._rcT && sy < obj._rcB) {
                    target := obj
                    break
                }
            }
        }

        if !target
            return

        obj := target
        dir := obj._ScrollDirection(msg, wParam)
        obj._scrollUntil := current + obj._ScrollWindowMsFor(msg)
        ; Reset TRACKSELECT countdown and re-arm hover; then perform a tiny real cursor jiggle
        try SendMessage(CreateLVHover.LVM_SETHOVERTIME, 0, obj.hoverMs,, obj._hTitle)
        obj._RequestHover()
        if (!dir || !obj._AtScrollBoundary(dir)) {
            obj._jiggleCursorCached(obj.jigglePx)
        }
    }

    ; ---------------- Internal: poll ----------------
    _Poll() {
        now := A_TickCount

        if !this._IsVisibleEnabled()
            return

        ; scroll-cooldown eller jiggle-ack venter? bail out tidligt
        if (now < this._scrollUntil)
            return

        if (this._jiggleArmed) {
            return
        }

        MouseGetPos &sx, &sy
        moved := (sx != this._mx || sy != this._my)
        

        ; Pause while a mouse button is held (drag/resize etc.)
        if GetKeyState("LButton", "P") || GetKeyState("RButton", "P") || GetKeyState("MButton", "P") {
            this._mx := sx, this._my := sy
            return
        }

        ; Outside? clear hot & state
        if !(sx >= this._rcL && sx < this._rcR && sy >= this._rcT && sy < this._rcB) {
            if (this._lastIdx != -1)
                SendMessage(CreateLVHover.LVM_SETHOTITEM, -1, 0,, this._hTitle)
            this._lastIdx := -1
            this._mx := sx, this._my := sy
            return
        }

        ; Hit test under pointer
        idx := this._hitTestIndexFromScreen(sx, sy)

        ; Fast path: nothing changed and not in scroll window
        if (idx = this._lastIdx && (sx = this._mx && sy = this._my) && now > this._scrollUntil) {
            this._mx := sx, this._my := sy
            return
        }

        ; Jiggle only when the item changed and either the mouse did not move
        ; OR we are within the scroll window (recent scroll gesture)
        if (idx != this._lastIdx) {
            if (idx >= 0 && (!moved || now <= this._scrollUntil)) {
                if (!this._jiggleArmed)
                    this._jiggleCursorCached(this.jigglePx)
            }
            this._lastIdx := idx
        }

        this._mx := sx, this._my := sy
    }

    ; ---------------- Helpers: rect/mouse/hittest ----------------
    _updateRectCache() {
        ; Regular per‑tick cache update
        rc := this._bufRect
        DllCall("GetWindowRect", "ptr", this.hLV, "ptr", rc)
        this._rcL := NumGet(rc, 0,  "Int")
        this._rcT := NumGet(rc, 4,  "Int")
        this._rcR := NumGet(rc, 8,  "Int")
        this._rcB := NumGet(rc, 12, "Int")
    }

    _hitTestIndexFromScreen(sx, sy) {
        pt := this._bufPt
        NumPut("Int", sx, pt, 0), NumPut("Int", sy, pt, 4)
        DllCall("ScreenToClient", "ptr", this.hLV, "ptr", pt)
        cx := NumGet(pt, 0, "Int")
        cy := NumGet(pt, 4, "Int")
        hti := this._bufHti
        NumPut("Int", cx, hti, 0), NumPut("Int", cy, hti, 4)
        return SendMessage(CreateLVHover.LVM_SUBITEMHITTEST, 0, hti,, this._hTitle)
    }

    ; -------- Scroll boundary helpers --------
    _SignHiWord(val) {
        delta := (val >> 16) & 0xFFFF
        return (delta >= 0x8000) ? (delta - 0x10000) : delta
    }

    _ScrollDirection(msg, wParam) {
        static WM_MOUSEWHEEL:=0x020A, WM_MOUSEHWHEEL:=0x020E
            ,  WM_POINTERWHEEL:=0x024E, WM_POINTERHWHEEL:=0x024F
            ,  WM_VSCROLL:=0x0115, WM_HSCROLL:=0x0114
            ,  SB_LINEUP:=0, SB_LINEDOWN:=1, SB_PAGEUP:=2, SB_PAGEDOWN:=3
            ,  SB_THUMBPOSITION:=4, SB_THUMBTRACK:=5, SB_TOP:=6, SB_BOTTOM:=7
            ,  SB_LINELEFT:=0, SB_LINERIGHT:=1, SB_PAGELEFT:=2, SB_PAGERIGHT:=3
            ,  SB_LEFT:=6, SB_RIGHT:=7

        ; signed delta fra HIWORD(wParam)
        delta := (wParam >> 16) & 0xFFFF
        if (delta & 0x8000)
            delta -= 0x10000

        ; Wheel / pointer (vertikal)
        if (msg = WM_MOUSEWHEEL || msg = WM_POINTERWHEEL) {
            if GetKeyState("Shift","P")
                return delta > 0 ? "L" : (delta < 0 ? "R" : "")
            return delta > 0 ? "U" : (delta < 0 ? "D" : "")
        }
        ; Wheel / pointer (horisontal)
        if (msg = WM_MOUSEHWHEEL || msg = WM_POINTERHWHEEL) {
            return delta > 0 ? "R" : (delta < 0 ? "L" : "")
        }

        ; V/HSCROLL (tastatur/scrollbar)
        if (msg = WM_VSCROLL) {
            code := wParam & 0xFFFF
            switch code {
                case SB_LINEUP, SB_PAGEUP, SB_TOP:        return "U"
                case SB_LINEDOWN, SB_PAGEDOWN, SB_BOTTOM: return "D"
                case SB_THUMBTRACK, SB_THUMBPOSITION:     return ""   ; drag -> ingen retning
            }
            return ""
        }
        if (msg = WM_HSCROLL) {
            code := wParam & 0xFFFF
            switch code {
                case SB_LINELEFT, SB_PAGELEFT, SB_LEFT:    return "L"
                case SB_LINERIGHT, SB_PAGERIGHT, SB_RIGHT: return "R"
                case SB_THUMBTRACK, SB_THUMBPOSITION:      return ""
            }
            return ""
        }
        return ""
    }
    
    _ScrollWindowMsFor(msg) {
        ; Local constant
        static WM_POINTERWHEEL  := 0x024E
            ,  WM_POINTERHWHEEL := 0x024F
            ,  WM_VSCROLL       := 0x0115
            ,  WM_HSCROLL       := 0x0114

         ; Read safe defaults
        msDefault := (this.HasProp("scrollWindowMs")        && this.scrollWindowMs        > 0) ? this.scrollWindowMs        : 120
        msPointer := (this.HasProp("scrollWindowMsPointer") && this.scrollWindowMsPointer > 0) ? this.scrollWindowMsPointer : 160
        msKeybd   := (this.HasProp("scrollWindowMsKeyboard")&& this.scrollWindowMsKeyboard> 0) ? this.scrollWindowMsKeyboard: msDefault

        ; Pointer/touchpad: use the pointer window
        if (msg = WM_POINTERWHEEL || msg = WM_POINTERHWHEEL)
            return this.scrollWindowMsPointer

        ; Keyboard/scrollbar: use keyboard window if set, otherwise normal
        if (msg = WM_VSCROLL || msg = WM_HSCROLL)
            return this.HasProp("scrollWindowMsKeyboard")
                ? this.scrollWindowMsKeyboard
                : this.scrollWindowMs

        ; Default: regular wheel
        return this.scrollWindowMs
    }

    _AtScrollBoundary(dir) {
    ; Vertical: U/D
        if (dir = "U" || dir = "D") {
            static LVM_GETTOPINDEX := 0x1027
            static LVM_GETCOUNTPERPAGE := 0x1028
            static LVM_GETITEMCOUNT := 0x1004
            top := SendMessage(LVM_GETTOPINDEX, 0, 0,, this._hTitle)
            per := SendMessage(LVM_GETCOUNTPERPAGE, 0, 0,, this._hTitle)
            cnt := SendMessage(LVM_GETITEMCOUNT, 0, 0,, this._hTitle)
            if (per < 1)
                per := 1
            maxTop := Max(0, cnt - per)
            return dir = "U" ? (top <= 0) : (top >= maxTop)
        }

        ; Horizontal: L/R
        if (dir = "L" || dir = "R") {
            static SIF_ALL := 0x17, SB_HORZ := 0

            si := Buffer(28, 0)                    ; sizeof(SCROLLINFO)=28
            NumPut("UInt", 28,     si, 0)          ; cbSize
            NumPut("UInt", SIF_ALL, si, 4)         ; fMask
            DllCall("GetScrollInfo", "ptr", this.hLV, "int", SB_HORZ, "ptr", si)

            nMin  := NumGet(si,  8, "Int")
            nMax  := NumGet(si, 12, "Int")
            nPage := NumGet(si, 16, "UInt")
            nPos  := NumGet(si, 20, "Int")

            edgeMin := (nPos <= nMin)
            edgeMax := (nPos >= (nMax - (nPage ? nPage - 1 : 0)))

            return dir = "L" ? edgeMin : edgeMax
        }

        ; Unknown direction => not boundary
        return false
    }

    _RequestHover() {
        ; Re-arm OS hover tracking so TRACKSELECT counts again even if the mouse is stationary
        TME_HOVER := 0x00000001
        sz := (A_PtrSize = 8) ? 24 : 16
        tme := Buffer(sz, 0)
        NumPut("UInt", sz,  tme, 0)      ; cbSize
        NumPut("UInt", TME_HOVER, tme, 4) ; dwFlags
        NumPut("Ptr",  this.hLV,  tme, 8) ; hwndTrack
        off := (A_PtrSize = 8) ? 16 : 12
        NumPut("UInt", this.hoverMs, tme, off) ; dwHoverTime (ms)
        DllCall("TrackMouseEvent", "ptr", tme)
    }

    static _FocusIsMappedLV() {
        try {
            cn := ControlGetFocus("A")          ; ClassNN in active window
            if (cn = "")
                return false
            fh := ControlGetHwnd(cn, "A")       ; hwnd for the focus control
            return fh && CreateLVHover._hwndMap.Has(fh)
        } catch {
            return false
        }
    }

    static _OnKeyNav(dir) {
        try {
            cn := ControlGetFocus("A")
            if (cn = "")
                return
            fh := ControlGetHwnd(cn, "A")
            if !(fh && CreateLVHover._hwndMap.Has(fh))
                return

            obj := CreateLVHover._hwndMap[fh]
            if !IsObject(obj) || !obj._IsVisibleEnabled()
                return

            current := A_TickCount
            ; use keyboard window if set, otherwise default
            win := obj.HasProp("scrollWindowMsKeyboard") ? obj.scrollWindowMsKeyboard : obj.scrollWindowMs
            obj._scrollUntil := current + win

            ; re-arm hover/TrackSelect immediately
            try SendMessage(CreateLVHover.LVM_SETHOVERTIME, 0, obj.hoverMs,, obj._hTitle)
            obj._RequestHover()

            ; no jiggle here – __poll() can handle one micro-jiggle when the position changes
        }
    }

    ; ---------------- Helpers: jiggle (uses cached rect) ----------------
    _jiggleCursorCached(px := 1) {
        ; Compute a minimal cursor move inside the LV bounds, set it, and let
        ; _ClassOnMouseMove move it back as soon as WM_MOUSEMOVE is delivered.
        if (this._jiggleArmed)
            return ; already waiting for first move ack
        pt := this._bufPt
        DllCall("GetCursorPos", "ptr", pt)
        x := NumGet(pt, 0, "Int")
        y := NumGet(pt, 4, "Int")

        dx := 0, dy := 0
        if (y + px < this._rcB && y + px >= this._rcT)
            dy := px
        else if (y - px >= this._rcT && y - px < this._rcB)
            dy := -px
        else if (x + px < this._rcR && x + px >= this._rcL)
            dx := px
        else if (x - px >= this._rcL && x - px < this._rcR)
            dx := -px
        else
            return  ; No safe jiggle possible

        this._jiggleOrigX := x
        this._jiggleOrigY := y
        this._jiggleArmed := true
        DllCall("SetCursorPos", "int", x + dx, "int", y + dy)
    }

    ; Complete the jiggle once the system has dispatched WM_MOUSEMOVE to the control
    static _ClassOnMouseMove(wParam, lParam, msg, hwnd) {
        if CreateLVHover._hwndMap.Has(hwnd) {
            obj := CreateLVHover._hwndMap[hwnd]
            if obj._jiggleArmed {
                DllCall("SetCursorPos", "int", obj._jiggleOrigX, "int", obj._jiggleOrigY)
                obj._jiggleArmed := false
            }
        }
    }

    ; ---------------- Class manager: single timer, multi-LV ----------------
    static _ManagerEnsureStarted() {
        if !CreateLVHover._mgrPollCb
            CreateLVHover._mgrPollCb := ObjBindMethod(CreateLVHover, "_ManagerTick")
        SetTimer(CreateLVHover._mgrPollCb, CreateLVHover._mgrTimerMs)
    }

    static _ManagerMaybeStop() {
        try if CreateLVHover._mgrPollCb
            SetTimer(CreateLVHover._mgrPollCb, 0)
        catch {
        }
        CreateLVHover._mgrTimerMs := 30 ; reset default period
    }

    static _ManagerTick() {
        CoordMode "Mouse", "Screen"
        MouseGetPos &sx, &sy

        ; If an active instance exists and is enabled & visible, poll it while the pointer stays on it
        inst := CreateLVHover._mgrActive
        if (inst && inst._IsVisibleEnabled()) {
            inst._updateRectCache()
            if (sx >= inst._rcL && sx < inst._rcR && sy >= inst._rcT && sy < inst._rcB) {
                inst._isActive := true
                inst._Poll()
                return
            }
            inst._isActive := false
            CreateLVHover._mgrActive := 0
        }

        ; Find a candidate under the pointer (first match)
        for hwnd, obj in CreateLVHover._hwndMap {
            if (!obj._IsVisibleEnabled())
                continue
            obj._updateRectCache()
            if (sx >= obj._rcL && sx < obj._rcR && sy >= obj._rcT && sy < obj._rcB) {
                CreateLVHover._mgrActive := obj
                obj._mx := sx, obj._my := sy  ; prime instance with current mouse pos
                obj._isActive := true
                obj._Poll()
                return
            }
        }
        ; idle: no LV under pointer → no per‑instance polling
    }

    static _ManagerRecomputePeriod() {
        if (CreateLVHover._hwndMap.Count = 0) {
            CreateLVHover._mgrTimerMs := 30
            return
        }
        minp := 100000
        for h, obj in CreateLVHover._hwndMap
            minp := (obj.pollMs < minp) ? obj.pollMs : minp
        CreateLVHover._mgrTimerMs := minp
        if !CreateLVHover._mgrPollCb
            CreateLVHover._mgrPollCb := ObjBindMethod(CreateLVHover, "_ManagerTick")
        SetTimer(CreateLVHover._mgrPollCb, CreateLVHover._mgrTimerMs)
    }

    ; ---------------- Status API ----------------
    static GetStatusArray() {
        arr := []
        for h, obj in CreateLVHover._hwndMap {
            arr.Push({
                hwnd:           obj.hLV,
                active:         obj._isActive,
                visibleEnabled: obj._IsVisibleEnabled(),
                _lastIdx:        obj._lastIdx
            })
        }
        return arr
    }

    IsActive() => this._isActive

    ; ---------------- Util: resolve HWND & validation ----------------
    _IsVisibleEnabled() {
        vis := DllCall("IsWindowVisible", "ptr", this.hLV, "int")
        gwl := (A_PtrSize = 8) ? "GetWindowLongPtr" : "GetWindowLong"
        style := DllCall(gwl, "ptr", this.hLV, "int", -16, "ptr")  ; GWL_STYLE
        WS_DISABLED := 0x08000000
        return vis && !(style & WS_DISABLED)
    }

    UpdateNow() => this._Poll()
    _resolveHwnd(target) {
        ; Returns 0 if cannot resolve – caller will report a friendly error
        if IsObject(target) {
            try return target.Hwnd
            catch as e {
                return 0
            }
        }
        if (Type(target) = "Integer" && target > 0)
            return target
        if (Type(target) = "String") {
            h := WinExist(target)
            return h ; may be 0
        }
        return 0
    }

    _ensureListView(h, originalTarget := "") {
        ; Strict validation only: no child auto-discovery to avoid ambiguity when multiple LVs exist.
        if this._isListView(h)
            return h
        ; Not a ListView → show rich error and exit
        this._reportInvalidTarget(h, originalTarget)
        ExitApp
    }

    _reportInvalidTarget(h, originalTarget) {
        cls   := this._getClassName(h)
        title := this._safeTitle(h)
        info  := this._formatTargetInfo(originalTarget)

        ; Capture a stack so we can show a friendly Caller: file
        try {
            throw Error("CreateLVHover invalid target (capture stack)")
        } catch as e {
            frame := this._findCallsiteFrame(e)
            ; Prefer exact caller captured by constructor; fall back to stack
            cFile := (this._callerFile != "") ? this._callerFile
                    : (ObjHasOwnProp(frame, "File") && frame.File) ? frame.File : ""
            callerName := this._basename(cFile)
            if (callerName = "")
                callerName := A_ScriptName

            MsgBox Format(
            "CreateLVHover — invalid target.`n`n"
          . "Expected: SysListView32`n"
          . "Got class: {1}`n"
          . "Window title: {2}`n"
          . "Target info: {3}`n`n"
          . "Caller: {4}`n`n"
          . "Fix: Pass a ListView control, its HWND, or a WinTitle that resolves to a SysListView32.",
            cls, (title != "" ? title : "(no title)"), info, callerName)
        }
    }

    _findCallsiteFrame(e) {
        ; Prefer a caller frame *outside* the library directory and WITH a line number.
        ; Fall back gradually if line info is missing in the stack.
        if (CreateLVHover._libDir = "") {
            f := CreateLVHover._libFile
            p := InStr(f, "\", false, -1)
            CreateLVHover._libDir := p ? SubStr(f, 1, p) : ""
        }
        libDir := CreateLVHover._libDir

        st := e.Stack
        if !(IsObject(st) && st.Length)
            return { What: "(script)", File: A_ScriptFullPath, Line: "?" }

        hasLine(f) => (ObjHasOwnProp(f, "Line") && f.Line != "")

        ; 1) First (script) frame outside the library folder and WITH line
        for i, f in st {
            w    := ObjHasOwnProp(f, "What") ? f.What : ""
            file := ObjHasOwnProp(f, "File") ? f.File : ""
            if (w = "(script)" && file != "" && (libDir = "" || InStr(file, libDir) != 1) && hasLine(f))
                return f
        }
        ; 2) Backwards: first frame outside the library and WITH line
        i := st.Length
        while (i >= 1) {
            f    := st[i]
            file := ObjHasOwnProp(f, "File") ? f.File : ""
            if (file != "" && (libDir = "" || InStr(file, libDir) != 1) && hasLine(f))
                return f
            i -= 1
        }
        ; 3) First frame outside the library
        for f in st {
            file := ObjHasOwnProp(f, "File") ? f.File : ""
            if (file != "" && (libDir = "" || InStr(file, libDir) != 1))
                return f
        }
        ; 4) First frame with line
        for f in st {
            if hasLine(f)
                return f
        }
        ; Fallback
        return { What: "(script)", File: A_ScriptFullPath, Line: "?" }
    }

    _isListView(h) {
        return this._getClassName(h) = "SysListView32"
    }

    _getClassName(h) {
        try {
            return WinGetClass("ahk_id " h)
        } catch as e {
            buf := Buffer(256, 0)
            DllCall("GetClassName", "ptr", h, "ptr", buf, "int", 256)
            return StrGet(buf, "UTF-16")
        }
    }

    _safeTitle(h) {
        try {
            return WinGetTitle("ahk_id " h)
        } catch as e {
            return ""
        }
    }

    _formatTargetInfo(target) {
        try {
            if IsObject(target) {
                hasHwnd := this._tryGetProp(target, "Hwnd", &h)
                if (hasHwnd) {
                    cls  := this._getClassName(h)
                    titl := this._safeTitle(h)
                    nameOK := this._tryGetProp(target, "Name", &nm)
                    textOK := this._tryGetProp(target, "Text", &tx)
                    if (StrLen(tx) > 60)
                        tx := SubStr(tx, 1, 57) . "..."
                    return Format(
                        "Object({1}) vName={2} -> HWND=0x{3:X} Class={4} Title={5} Text={6}",
                        Type(target), (nm != "" ? nm : "(none)"), h, cls, (titl != "" ? titl : "(no title)"), (tx != "" ? tx : "(empty)")
                    )
                }
                return "Object(" . Type(target) . ")"
            }
            if (Type(target) = "Integer") {
                h    := target
                cls  := this._getClassName(h)
                titl := this._safeTitle(h)
                return Format("HWND=0x{1:X} Class={2} Title={3}", h, cls, (titl != "" ? titl : "(no title)"))
            }
            if (Type(target) = "String") {
                return "WinTitle: " . (target != "" ? target : "(empty)")
            }
            return Type(target)
        } catch as e {
            return "(unavailable)"
        }
    }

    _tryGetProp(obj, propName, &out) {
        try {
            out := obj.%propName%
            return true
        } catch as e {
            out := ""
            return false
        }
    }

    _basename(path) {
        try {
            if !path
                return ""
            SplitPath path, &fileOnly
            return fileOnly
        } catch as e {
            return ""
        }
    }
}
