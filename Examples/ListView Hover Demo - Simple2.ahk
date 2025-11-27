#Requires AutoHotkey v2.0
#SingleInstance Force
;#include <UIA> ; Uncomment if you have moved UIA.ahk to your main Lib folder
#include ..\Lib\CreateLVHover.ahk

Demo := Gui()

NumArray := ["1: one", "2: two", "3: three", "4: four", "5: five",
        "6: six", "7: seven", "8: eight", "9: nine", "10: ten",
        "11: eleven", "12: twelve", "13: thirteen", "14: fourteen", "15: fifteen",
        "16: sixteen", "17: seventeen", "18: eighteen", "19: nineteen", "20: twenty"]

lv1 := Demo.AddListView("r15 w160", ["Name"]) ; any standard ListView
for v in NumArray
    lv1.Add("", v)

lv1.ModifyCol(1, "AutoHdr")

lv2 := Demo.AddListView("x+15 yp hp wp", ["Name"]) ; any standard ListView
for v in NumArray
    lv2.Add("", v)

lv2.ModifyCol(1, "AutoHdr")

Demo.Show()


CreateLVHover(lv1)
CreateLVHover(lv2)