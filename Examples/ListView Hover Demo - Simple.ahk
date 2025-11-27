#Requires AutoHotkey v2.0
#SingleInstance Force
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\CreateLVHover.ahk

Demo := Gui()
lv := Demo.AddListView(, "r15 w320", ["Name"]) ; any standard ListView
for v in ["1: one", "2: two", "3: three", "4: four", "5: five",
        "6: six", "7: seven", "8: eight", "9: nine", "10: ten",
        "11: eleven", "12: twelve", "13: thirteen", "14: fourteen", "15: fifteen",
        "16: sixteen", "17: seventeen", "18: eighteen", "19: nineteen", "20: twenty"]
    lv.Add("", v)

lv.ModifyCol(1, "AutoHdr")

Demo.Show()


CreateLVHover(lv, {
            underline: true, ; underline on hover (Explorer‑style)
            oneClick: true, ; one‑click activate + TRACKSELECT
            hoverMs: 1000, ; time before TRACKSELECT (ms)
            timerMs: 30, ; poll interval for "scroll under pointer"
            pxJiggle: 1, ; minimal pointer nudge to trigger real WM_MOUSEMOVE
            scrollWindowMs: 150 ; grace period after scroll where jiggle is allowed
})