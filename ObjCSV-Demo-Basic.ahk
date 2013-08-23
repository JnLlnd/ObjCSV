;===============================================
/* ObjCSV Demo Basic v0.1
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By jlalonde on AHK forum
2013-08-22+
Released under GNU GENERAL PUBLIC LICENSE (Version 2, June 1991)
*/

#NoEnv
#SingleInstance
#Include %A_ScriptDir%\JLDev.ahk ; some debugging tools or utilities

Gui, New, +Resize
Gui, Add, Button, x10 Default gButtonLoadFile, Load file
Gui, Add, Button, x+10 yp gButtonRead, Read object
Gui, Add, Button, x+10 yp gButtonChange, Change list
Gui, Add, Button, x+10 yp gButtonSaveFile, Save file
Gui, Add, Button, x+10 yp gGuiClose, Quit
Gui, Add, ListView, x10 r20 w800 -ReadOnly vMyLV
Gui, Show, Autosize
return



ButtonLoadFile:
strListViewName := "MyLV"
strFile := A_ScriptDir .  "\TheBeatles.txt"
strFields := ""
obj := ObjCSV_CSV2Collection(strFile, strFields)
LV_Delete() ; delete all rows of ListView
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all columns of ListView
strFields := "str_Name,str_Album,lng_Track_Number,str_Genre,lng_Total_Time,lng_Size" ; field order in the ListView
ObjCSV_Collection2ListView(obj, , strListViewName, strFields)
LV_ModifyCol(3, "Integer") ; allow numeric sorting
LV_ModifyCol(5, "Integer") ; allow numeric sorting
LV_ModifyCol(6, "Integer") ; allow numeric sorting
return



ButtonRead:
if !obj.MaxIndex() ; object is empty
	return
Gui +OwnDialogs
Random, intRecordNumber, 1, obj.MaxIndex()
MsgBox, , Random Song!, % "Song #" . intRecordNumber . " is """ . obj[intRecordNumber].str_Name . """"
Loop, % obj.MaxIndex() ; loop in each record in the obj object
{
	intRecordNumber := A_Index
	str := "--------------------------------------------------------------------------------`n"
	for strFieldName, strFieldValue in obj[intRecordNumber] ; loop each field in the record
		str := str . "[ " . strFieldName . " ]`t`t " . strFieldValue . "`n"
	str := str . "--------------------------------------------------------------------------------"
	MsgBox, 4, Song #%intRecordNumber%, %str% `n`nContinue?
	IfMsgBox, No
		Break
}
return



ButtonChange:
if !LV_GetCount("") ; ListView is empty
	return
blnAlternate := !blnAlternate
Loop, % LV_GetCount("") ; loop in each row in the ListView
{
	intRowNumber := A_Index
	Loop, % LV_GetCount("Column") ; loop in each column in the row
	{
		LV_GetText(strColData, intRowNumber, A_Index) ;  read this cell
		if blnAlternate
			StringUpper, strColData, strColData
		else
			StringUpper, strColData, strColData, T
		LV_Modify(intRowNumber, "Col" . A_Index, strColData) ;  update this cell
	}
}
MsgBox, , Case, Click "Change list" again to change case again
return



ButtonSaveFile:
if !LV_GetCount("") ; ListView is empty
	return
Gui +OwnDialogs  ; Forces user to dismiss the following dialog before using main window.
FileSelectFile, strFile, S18, %A_ScriptDir%\TheBeatles-Copy.txt, Save CSV file as?
if not strFile ; The user canceled the dialog.
	return
strFields := "str_Name,str_Album,lng_Track_Number,str_Genre,lng_Total_Time,lng_Size" ;  field order in the saved file
obj := ObjCSV_ListView2Collection()
ObjCSV_Collection2CSV(obj, strFile, 1, strFields)
MsgBox, 4, Display file?, File saved:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
return



GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the ListView to match.
GuiControl, -Redraw, MyLV
intWidth := A_GuiWidth + 30
GuiControl, Move, MyLV, % "W" . (A_GuiWidth - 20) . " H" . (A_GuiHeight - 45)
LV_ModifyCol(1, (intWidth * 0.2))
LV_ModifyCol(2, (intWidth * 0.2))
LV_ModifyCol(3, (intWidth * 0.1))
LV_ModifyCol(4, (intWidth * 0.2))
LV_ModifyCol(5, (intWidth * 0.1))
LV_ModifyCol(6, (intWidth * 0.1))
LV_ModifyCol(7, (intWidth * 0.1))
GuiControl, +Redraw, MyLV
return



GuiClose:
ExitApp

