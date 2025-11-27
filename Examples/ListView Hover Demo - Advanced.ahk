#Requires AutoHotkey v2.0
#SingleInstance Force
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\CreateLVHover.ahk

; ============================================================================
; CreateLVHover — 5x4 Dynamic Demo with GroupBoxes + Dark/Light mode
; - 20 ListViews (5 columns x 4 rows) created via loops
; - Each row of ListViews is wrapped in a GroupBox
; - Each GroupBox caption describes its row-specific hover settings
; - Row-based hover options (underline / oneClick) per your specification
; - Status labels per LV showing Active/Idle and Visible/Hidden
; - 5 px padding between ListViews/labels and the GroupBox borders
; - Dark/Light mode toggle at the bottom (status labels stay green/red)
; ============================================================================

; --------------------------
; Layout and configuration
; --------------------------
NumColLV  := 5
NumRowLV  := 4
NumOfIcon := 50

M             := 10   ; outer window margin
gap           := 10   ; horizontal gap between ListViews
lvW           := 240
lvH           := 125
labelH        := 16
P             := 5    ; padding inside each GroupBox (all sides)
captionOffset := 10   ; vertical offset from GroupBox top to inner content
gbGap         := 5    ; vertical gap between GroupBoxes

; GroupBox width/height based on inner ListView grid and padding
innerWidth := NumColLV * lvW + (NumColLV - 1) * gap
gbW        := innerWidth + 2 * P
; Height: caption + top padding + LV + spacing + label + bottom padding
gbH        := captionOffset + P + lvH + 2 + labelH + P

; --------------------------
; GUI
; --------------------------
demo := Gui("+Resize", "CreateLVHover — " NumColLV "x" NumRowLV " Dynamic Demo")
demo.MarginX := M
demo.MarginY := M
demo.SetFont("cWhite")
demo.OnEvent("Close", (*) => (SetTimer(__ELVH_UpdateStatusLabels, 0), ExitApp()))

; --------------------------
; Data containers
; --------------------------
LVs    := Map()                         ; key: idx (0-based), value: ListView control
Labels := Map()                         ; key: idx (0-based), value: Text control

global __ELVH_LabelByHwnd      := Map() ; key: LV.Hwnd, value: label control
global __ELVH_TypeByHwnd       := Map() ; key: LV.Hwnd, value: type label text
global __ELVH_StatusLabelHwnd  := Map() ; label.Hwnd -> true (for skipping)

GroupBoxes := Map()                     ; key: row (0-based), value: GroupBox control

; --------------------------
; Helper functions
; --------------------------
LVTypeForCol(col) {
    switch col {
        case 0: return "Icon"
        case 1: return "IconSmall"
        case 2: return "Tile"
        case 3: return "List"
        default: return "Report"
    }
}

LVTypeLabel(vtype) {
    ; Human-friendly labels for each ListView mode
    switch vtype {
        case "Icon":      return "Large icons"
        case "IconSmall": return "Small icons"
        case "Tile":      return "Tile"
        case "List":      return "List"
        case "Report":    return "Details"
        default:          return vtype
    }
}

PosText(col, row, NumCol, NumRow) {
    colPos := (col = 0) ? "Left"
           : (col = NumCol - 1) ? "Right"
           : Format("Middle{}", col)
    rowPos := (row = 0) ? "Top"
           : (row = NumRow - 1) ? "Bottom"
           : Format("Middle{}", row)
    return colPos "-" rowPos
}

RowCaption(row, optsMap) {
    if optsMap.Has(row) {
        o := optsMap[row]
        return Format(
            "underline={} | oneClick={}",
            o.underline ? "true" : "false",
            o.oneClick  ? "true" : "false"
        )
    } else {
        ; Fallback text for any additional rows beyond those explicitly defined
        return Format("default underline=true | oneClick=true")
    }
}

; --------------------------
; Row-based hover options
; --------------------------
; Row 0: underline=true,  oneClick=true
; Row 1: underline=false, oneClick=true
; Row 2: underline=true,  oneClick=false
; Row 3: underline=false, oneClick=false
; Any other rows (if added later): default underline=true, oneClick=true
RowHoverOpts := Map(
    0, { underline: true , oneClick: true ,  hoverMs: 1000, timerMs: 10, pxJiggle: 1 },
    1, { underline: false, oneClick: true ,  hoverMs: 1000, timerMs: 10, pxJiggle: 1 },
    2, { underline: true , oneClick: false,  hoverMs: 1000, timerMs: 10, pxJiggle: 1 },
    3, { underline: false, oneClick: false,  hoverMs: 1000, timerMs: 10, pxJiggle: 1 }
)

; --------------------------
; Create 5×4 ListViews + labels inside GroupBoxes
; --------------------------
Loop NumRowLV {
    row := A_Index - 1

    ; GroupBox position for this entire row of ListViews
    gbX := M
    gbY := M + row * (gbH + gbGap)

    gb := demo.AddGroupBox(
        "x" gbX " y" gbY " w" gbW " h" gbH,
        RowCaption(row, RowHoverOpts)
    )
    GroupBoxes[row] := gb

    ; Now place 5 ListViews + labels inside this GroupBox with 5 px padding
    Loop NumColLV {
        col := A_Index - 1

        idx   := row * NumColLV + col  ; 0-based index across full grid
        x     := gbX + P + col * (lvW + gap)
        y     := gbY + captionOffset + P
        vtype := LVTypeForCol(col)
        header := [ PosText(col, row, NumColLV, NumRowLV) ]

        lv := demo.AddListView(
            vtype " -hdr x" x " y" y " w" lvW " h" lvH,
            header
        )

        typeLabel := LVTypeLabel(vtype)

        lbl := demo.AddText(
            "x" x " y" (y + lvH + 2) " w" lvW " h" labelH " vLbl" (idx + 1),
            typeLabel " | Polling: Idle"
        )

        LVs[idx] := lv
        Labels[idx] := lbl
        __ELVH_LabelByHwnd[lv.Hwnd] := lbl
        __ELVH_TypeByHwnd[lv.Hwnd]  := typeLabel
        __ELVH_StatusLabelHwnd[lbl.Hwnd] := true
    }
}

; --------------------------
; Dark/Light mode toggle (GroupBox under the others)
; --------------------------
modeGbX := M
modeGbY := M + NumRowLV * (gbH + gbGap)  ; keep the same vertical spacing pattern
modeGbW := 185
modeGbH := 35

modeGb := demo.AddGroupBox(
    "x" modeGbX " y" modeGbY " w" modeGbW " h" modeGbH,
    "Color scheme"
)

rY := modeGbY + captionOffset

darkRadio := demo.AddRadio(
    "x" (modeGbX + 15) " y" (rY + 5) " vDarkMode Checked",
    "Dark mode"
)
darkRadio.OnEvent("Click", scheme_update)

lightRadio := demo.AddRadio(
    "x" (modeGbX + 100) " y" (rY + 5) " vLightMode",
    "Light mode"
)
lightRadio.OnEvent("Click", scheme_update)

; --------------------------
; Imagelists (small + large) and icon loading
; --------------------------
; We load the same icon order into both imagelists to keep indices identical.
ilSmall := IL_Create(NumOfIcon)               ; small (~16x16) for Report/List/IconSmall
ilLarge := IL_Create(NumOfIcon, , true)       ; large (3rd param = LargeIcons) for Icon/Tile

IconIdx := []  ; holds indices returned by IL_Add
Loop NumOfIcon {
    i := A_Index
    idxS := IL_Add(ilSmall, "shell32.dll", i)
    idxL := IL_Add(ilLarge, "shell32.dll", i)
    if (idxS = 0 || idxL = 0)
        throw Error("Failed to load icon #" i " from shell32.dll")
    if (idxS != idxL)
        throw Error("Small/Large imagelist index mismatch for icon #" i)
    IconIdx.Push(idxS)
}

; Bind both imagelists to each LV
for _, lv in LVs {
    lv.SetImageList(ilLarge, 0)  ; LVSIL_NORMAL  (used by Icon/Tile views)
    lv.SetImageList(ilSmall, 1)  ; LVSIL_SMALL   (used by Report/List/IconSmall)
}

; --------------------------
; Populate items: 50 rows per LV with distinct icons
; --------------------------
for i, iconIdx in IconIdx {
    text := "Item " i
    for _, lv in LVs
        lv.Add(Format("Icon{}", iconIdx), text)
}

; Auto-size first column header
for _, lv in LVs
    lv.ModifyCol(1, "AutoHdr")

; --------------------------
; Dark/Light mode update function
; --------------------------
scheme_update(radio, info) {
    global LVs, __ELVH_StatusLabelHwnd

    try LockScreenUpdates()
    catch as e {
        MsgBox "Fejl i LockScreenUpdates: " e.Message
        return
    }

    gui := radio.Gui

    ; Decide colors based on which radio was pressed
    switch radio.Name {
        case "LightMode":
            gui.BackColor := 0xE0E0E0          ; light gray background
            text_color := "cBlack"
            lvBk       := 0xFFFFFF             ; white LV background
            lvText     := 0x000000             ; black LV text
        case "DarkMode":
            gui.BackColor := 0x000000          ; black background
            text_color := "cWhite"
            lvBk       := 0x202020             ; dark LV background
            lvText     := 0xFFFFFF             ; white LV text
    }

    ; Remember which control had focus
    prev := gui.FocusedCtrl

    ; loop through every control in the GUI and set font color
    for con in gui {
        ; Skip status labels so they keep green/red
        if __ELVH_StatusLabelHwnd.Has(con.Hwnd)
            continue
        try con.SetFont(text_color)
    }

    ; Now update ListView colors (background + text + text background)
    for _, lv in LVs {
        hwnd := lv.Hwnd
        ; LVM_SETBKCOLOR      = 0x1001
        ; LVM_SETTEXTCOLOR    = 0x1024
        ; LVM_SETTEXTBKCOLOR  = 0x1026
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x1001, "ptr", 0, "ptr", lvBk)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x1024, "ptr", 0, "ptr", lvText)
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x1026, "ptr", 0, "ptr", lvBk)
    }

    ; Force redraw of the GUI window (GroupBoxes, Radios, etc.)
    DllCall("InvalidateRect", "ptr", gui.Hwnd, "ptr", 0, "int", true)
    DllCall("UpdateWindow",   "ptr", gui.Hwnd)

    ; Restore previous focus if possible, otherwise focus back on the radio
    if IsObject(prev) {
        try prev.Focus()
    } else {
        try radio.Focus()
    }

    UnlockScreenUpdates()
}


; Apply initial scheme (Dark mode checked by default)
scheme_update(darkRadio, 0)

; --------------------------
; Show window sized to content
; --------------------------
contentBottom := modeGbY + modeGbH
winW := 2 * M + gbW
winH := contentBottom + M

demo.Show("w" winW " h" winH)

; --------------------------
; Apply row-based hover options
; --------------------------
for idx, lv in LVs {
    row := Floor(idx / NumColLV)
    if RowHoverOpts.Has(row) {
        CreateLVHover(lv, RowHoverOpts[row])
    } else {
        ; Default behavior for any additional rows:
        ; underline=true, oneClick=true, with same timing as row 0
        CreateLVHover(lv)
    }
}

; --------------------------
; Status labels updater
; --------------------------
SetTimer(__ELVH_UpdateStatusLabels, 120)

__ELVH_UpdateStatusLabels() {
    global __ELVH_LabelByHwnd, __ELVH_TypeByHwnd
    statuses := CreateLVHover.GetStatusArray()
    for s in statuses {
        if __ELVH_LabelByHwnd.Has(s.hwnd) {
            lbl := __ELVH_LabelByHwnd[s.hwnd]
            state := s.active ? "Active" : "Idle"
            vis   := s.visibleEnabled ? "Visible" : "Hidden/Disabled"

            typeLabel := __ELVH_TypeByHwnd.Has(s.hwnd)
                ? __ELVH_TypeByHwnd[s.hwnd]
                : "ListView"

            newTxt := Format("{} | Polling: {}  |  {}", typeLabel, state, vis)
            if (lbl.Text != newTxt) {
                lbl.Text := newTxt
                ; green when Active, red when Idle
                lbl.SetFont((state = "Active") ? "c0x008000" : "c0xFF0000")
            }
        }
    }
}


LockScreenUpdates(){
    ;Lock screen update by passing the andle of the desktop window (HWND of 0)
    DllCall("User32.dll\LockWindowUpdate", "Ptr", DllCall("User32.dll\GetDesktopWindow", "Ptr"))
}

UnlockScreenUpdates(){
    DllCall("User32.dll\LockWindowUpdate", "Ptr", 0)
}