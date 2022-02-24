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
strFields := ""
obj := ObjCSV_CSV2Collection(strFile, strFields, , , , , , , , , "[]") ; load the CSV file to a collection of objects
strFields := "str_Name,str_Album,lng_Track_Number,str_Genre,lng_Total_Time,lng_Size" ; field order in the ListView
return



SaveFile:
strFile := A_ScriptDir . "\TheBeatles-ReusedFields.txt"
strFields := "lng_Track_Number,str_Genre,""[[str_Genre,lng_Track_Number,str_Name][#[2]: [3], [1]][strNewField]]"""
; ObjCSV_Collection2CSV(objCollection, strFilePath, blnHeader := 0, strFieldOrder := "", intProgressType := 0
	; , blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEolReplacement := ""
	; , strProgressText := "", strFileEncoding := "", blnAlwaysEncapsulate := 0, strEol := "")
ObjCSV_Collection2CSV(obj, strFile, 1, , , 1) ; save the collection of objects to a CSV file and overwrite this file
MsgBox, 4, Display file?, File saved:`n`n%strFile%`n`nDisplay file?
IfMsgBox, Yes
	Run %strFile%
return

