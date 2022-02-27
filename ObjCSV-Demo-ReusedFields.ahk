;===============================================
/*
ObjCSVTest Combine Fields
*/

#NoEnv
#SingleInstance, force

Gosub, LoadFile
Gosub, SaveFile

ExitApp


LoadFile:
strFile := A_ScriptDir .  "\TheBeatles-ReuseLoad.txt"
MsgBox, 4, Display INPUT file?, Input file:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
strFields := ""
obj := ObjCSV_CSV2Collection(strFile, strFields, , , , , , , , , "[]") ; load the CSV file to a collection of objects
return



SaveFile:
strFile := StrReplace(strFile, ".txt", "-OUTPUT.txt")
strFields := "str_Name,str_Album,str_AlbumName,lng_Track_Number,str_Genre,lng_Total_Time,lng_Size,str_TrackAlbumName"
; ObjCSV_Collection2CSV(objCollection, strFilePath, blnHeader := 0, strFieldOrder := "", intProgressType := 0
	; , blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEolReplacement := ""
	; , strProgressText := "", strFileEncoding := "", blnAlwaysEncapsulate := 0, strEol := "", strReuseDelimiters := "")
ObjCSV_Collection2CSV(obj, strFile, 1, strFields, , 1, , , , , , , , "[]") ; save the collection of objects to a CSV file and overwrite this file
MsgBox, 4, Display file?, File saved:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
return

/*
;===============================================
; ObjCSVTest Combine Fields SIMPLE

#NoEnv
#SingleInstance, force

Gosub, LoadFile
Gosub, SaveFile

ExitApp


LoadFile:
; strFile := A_ScriptDir .  "\..\..\CSVBuddy\TEST-Reuse-One-Simple.csv"
; strFile := A_ScriptDir .  "\..\..\CSVBuddy\Reuse-None-Simple.csv"
strFile := A_ScriptDir .  "\..\..\CSVBuddy\Reuse-Double-Simple.csv"
MsgBox, 4, Display INPUT file?, Input file:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
strFields := ""
obj := ObjCSV_CSV2Collection(strFile, strFields, , , , , , , , , "[]") ; load the CSV file to a collection of objects
return



SaveFile:
strFile := StrReplace(strFile, ".csv", "-OUTPUT.csv")
strFields := "F1&2,[[[F1&2][F1&2]][F1&2 Doubled]],F1&2&3,F3,[[[F1&2&3][F1&2&3]][F1&2&3 Doubled]]"
; ObjCSV_Collection2CSV(objCollection, strFilePath, blnHeader := 0, strFieldOrder := "", intProgressType := 0
	; , blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEolReplacement := ""
	; , strProgressText := "", strFileEncoding := "", blnAlwaysEncapsulate := 0, strEol := "", strReuseDelimiters := "")
ObjCSV_Collection2CSV(obj, strFile, 1, strFields, , 1, , , , , , , , "[]") ; save the collection of objects to a CSV file and overwrite this file
MsgBox, 4, Display OUTPUT file?, File saved:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
return

