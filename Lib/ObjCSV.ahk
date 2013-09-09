;===============================================
/*ObjCSV Library v0.2.1 (pull strEolReplacement on load from v0.1.3)
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By JnLlnd on AHK forum
2013-09-08
*/

;================================================
/*INSTRUCTIONS
Copy this script in a file named ObjCSV.ahk and save this file in one of these \Lib folders:
  %A_ScriptDir%\Lib\
  %A_MyDocuments%\AutoHotkey\Lib\
  [path to the currently running AutoHotkey_L.exe]\Lib\ (AutoHotKey_L can be downloaded here: http://l.autohotkey.net/ )

You can use the functions in this library by calling ObjCSV_FunctionName (no #Include required)
================================================
*/

;===============================================
/* FUNCTIONS QUICK REFERENCE
AutoHotkey_L (AHK) functions to load from CSV files, sort, display and save collections of records using the Object data type.
Files can be read and saved in any delimited format (CSV, semi-colon, tab delimited, single-line or multi-line, etc.).
Collections can also be displayed, edited and read in GUI ListView objects.
For more info on CSV files, see http://en.wikipedia.org/wiki/Comma-separated_values.

ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = 1, blnMultiline = 1, blnProgress = 0, strFieldDelimiter = ",", strEncapsulator = """", strRecordDelimiter = "`n", strOmitChars = "`r", strEolReplacement = ""])
Transfer the content of a CSV file to a collection of objects. Field names are taken from the first line of the file or from the strFieldNames parameter. Delimiters are configurable.

ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
Transfer the selected fields from a collection of objects to a CSV file. Field names taken from key names are optionnaly included in the CSV file. Delimiters are configurable.

ObjCSV_Collection2Fixed(objCollection, strFilePath, strWidth, [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
Transfer the selected fields from a collection of objects to a fixed-width file. Field names taken from key names are optionnaly included in the file. Delimiters are configurable. Width are determined by the delimited string strWidth.

ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", strSortFields = "", strSortOptions = "", blnProgress = 0])
Transfer the selected fields from a collection of objects to ListView. The collection can be sorted by the function. Field names taken from the objects keys are used as header for the ListView.

ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = 0])
Transfer the selected lines of the selected columns of a ListView to a collection of objects. Lines are transfered in the order they appear in the ListView. Column headers are used as objects keys.

ObjCSV_SortCollection(objCollection, strSortFields [, strSortOptions = "", blnProgress = 0])
Sort the objectfs of a collection using one or multiple fields as sorting key.  Internal AHK Sort options are supported.

See details for each fucntions below.

===============================================
*/



;================================================
ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames, blnHeader := 1, blnMultiline := 1, blnProgress := 0, strFieldDelimiter := ",", strEncapsulator := """", strRecordDelimiter := "`n", strOmitChars := "`r", strEolReplacement := "")
/*
Summary: Transfer the content of a CSV file to a collection of objects. Field names are taken from the first line of the file or from the strFieldNameReplacement parameter. If taken from the file, fields names are returned by the ByRef variable strFieldNames. Delimiters are configurable.

RETURNED VALUES:
This functions returns an object that contains an array of objects. This collection of objects can be viewed as a table in a database. Each object in the collection is like a record (or a line) in a table. These records are, in fact, associative arrays which contain a list key-value pairs. Key names are like field names (or column names) in the table. Key names are taken in the header of the CSV file, if it exists. Keys can be strings or integers, while values can be of any type that can be expressed as text. The records can be read using the syntax obj[1], obj[2] (...). Field values can be read using the syntax obj[1].keyname or, when field names contain spaces, obj[1]["key name"]. The "Loop, Parse" and "For key, value in array" commands allow to easily browse the content of these objects.
If blnHeader is true (or 1), the ByRef variable strFieldNames returns a string containing the field names (object keys) read from the first line of the CSV file, in the format and in the order they appear in the file. If a field name is empty, it is replaced with "Empty_" and its field number.  If a field name is duplicated, the field number is added to the duplicate name.  If blnHeader is false (or 0), the value of strFieldNames is unchanged by the function.

PARAMETERS:
strFilePath:
Path of the file to load, which is assumed to be in A_WorkingDir if an absolute path isn't specified.

ByRef strFieldNames
Input: Names for object keys if blnHeader if false. Names must appear in the same order as they appear in the file, separated by the strFieldDelimiter character (see below). If names are not provided and blnHeader is false, column numbers are used as object keys, starting at 1. Empty by default.
Output: See RETURNED VALUES above.

blnHeader := 1
Optional. If true (or 1), the objects key names are taken from the header of the CSV file (first line of the file). If blnHeader if false (or 0), the first line is considered as data (see strFieldNames). True (or 1) by default.

blnMultiline := 1
Optional. If true (or 1), multi-line fields are supported. Multi-line fields include line breaks (end-of-line characters) which are usualy considered as delimiters for records (lines of data). Multi-line fields must be enclosed by the strEncapsulator character (usualy double-quote, see below). True by default. NOTE-1: If you know that your CSV file does NOT include multi-line fields, turn this option to false (or 0) to allow handling of larger files and improve performance (RegEx experts, help needed! See the function code for details). NOTE-2: If blnMultiline is True, you can use the strEolReplacement parameter to specify a character (or string) that will be converted to line-breaks if found in the CSV file.

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.

strFieldDelimiter := ","
Optional. Field delimiter in the CSV file. One character, usually comma (default value) or tab. According to locale setting of software (e.g. MS Office) or user preferences, delimiter can be semi-colon (;), pipe (|), space, etc. NOTE-1: End of line characters (`n or `r) are prohibited as field separator since they are used as record delimiters. NOTE-2: Using the Trim function, %A_Space% and %A_Tab% (when tab is not a delimiter) are removed from the beginning and end of all field names and data.

strEncapsulator := """"
Optional. Character (usualy double-quote) used in the CSV file to embed fields that include at least one of these special characters: line-breaks, field delimiters or the encapsulator character itself. In this last case, the encapsulator character must be doubled in the string. For example: "one ""quoted"" word". All fields and headers in the CSV file can be encapsulated, if desired by the file creator. Double-quote by default.

strRecordDelimiter := "`n"
Optional. Record delimiter in the CSV file. Rarely modified. Default value is end-of-line (`n).

strOmitChars := "`r"
Optional. List of characters (case sensitive) to exclude from beginning and end of lines. By default `r (carriage return) to support text files with both `r`n (carriage return+linefeed) and `n (linefeed) end-of-lines.

strEolReplacement := ""
Optional. Character (or string) that will be converted to line-breaks if found in the CSV file. Replacements occur only when blnMultiline is True. Empty by default.
*/
{
	objCollection := Object() ; object that will be returned by the function (a collection or array of objects)
	objHeader := Object() ; holds the keys (fields name) of the objects in the collection
	FileRead, strData, %strFilePath%
	if blnMultiline
		chrEolReplacement := Prepare4Multilines(strData, strEncapsulator, blnProgress) ; make sure each record temporarily stands on a single line *** not tested on files with eol other than `n
	strData := Trim(strData, strRecordDelimiter) ; remove empty line (record) at the beginning or end of the string, if present
	if (blnProgress)
	{
		intMaxProgress := StrLen(strData)
		intProgress := 0
		Progress, R0-%intMaxProgress% FS8 A, Loading CSV data..., , , MS Sans Serif
	}
	Loop, Parse, strData, %strRecordDelimiter%, %strOmitChars% ; read each line (record) of the CSV file
	{
		intProgress := intProgress + StrLen(A_LoopField) + 2 ; for Progress bar, augment intProgress of len of line + 2 for cr-lf 
		if (blnProgress AND !Mod(intProgress, 5000))
			Progress, %intProgress% ;  update progress bar only every 5000 chars
		if (A_Index = 1) and (blnHeader) ; we have an header to read
		{
			objHeader := ReturnDSVObjectArray(A_LoopField, strFieldDelimiter, strEncapsulator) ; returns an object array from the first line of the delimited-separated-value file
			strFieldNamesMatchList := ""
			Loop, % objHeader.MaxIndex() ; check if fields names are empty or duplicated
			{
				if !StrLen(objHeader[A_Index]) ; field name is empty
					objHeader[A_Index] := "Empty_" . A_Index ; use field number as field name
				else
					if InStr(strFieldNamesMatchList, objHeader[A_Index]) ; field name is duplicate
						objHeader[A_Index] := objHeader[A_Index] . "_" . A_Index ; add field number to field name
				if !StrLen(strFieldNamesMatchList)
					strFieldNamesMatchList := objHeader[A_Index]
				else
					strFieldNamesMatchList := strFieldNamesMatchList . "," . objHeader[A_Index]
			}
			strFieldNames := ""
			for intIndex, strFieldName in objHeader ; returns the updated field names to the ByRef variable
				strFieldNames := strFieldNames . Format4CSV(strFieldName, strFieldDelimiter, strEncapsulator) . strFieldDelimiter
			StringTrimRight, strFieldNames, strFieldNames, 1 ; remove extra field delimiter
			if !(objHeader.MaxIndex()) ; we don't have an object, something went wrong
			{
				if (blnProgress)
					Progress, Off
				return ; returns no object(more error friendly code could be added here)
			}
		}
		else
		{
			if (A_Index = 1) and StrLen(strFieldNames)
				; If we get here, bnHeader is false so there is no header in the CSV file but we have values in strFieldNames.
				; In this case, we get field names from strFieldNames.
				objHeader := ReturnDSVObjectArray(strFieldNames, strFieldDelimiter, strEncapsulator) ; returns an object array from the delimited-separated-value strFieldNames string
			objData := Object() ;  object of one record in the collection
			for intIndex, strFieldData in ReturnDSVObjectArray(A_LoopField, strFieldDelimiter, strEncapsulator) ; returns an object array from each line of the delimited-separated-value file
			{
				if blnMultiline
				{
					StringReplace, strFieldData, strFieldData, %chrEolReplacement%, %strRecordDelimiter%, 1 ; put back all original end-of-line chararchers in each field, if present
					; Using %strRecordDelimiter% as replacement in the next command, eol are lost when saved using ObjCSV_Collection2CSV and opened in Notepad.
					; However, %strRecordDelimiter% seems to work well in the previous command... Anyway, for safety (at least in the Windows environment), I replaced it with `r`n in the next command.
					; For reference, see http://www.autohotkey.com/board/topic/57364-best-practices-for-handling-newlines-internally-in-ahk/ and http://peterbenjamin.com/seminars/crossplatform/texteol.html
					StringReplace, strFieldData, strFieldData, %strEolReplacement%, `r`n, 1 ; replace all user-supplied replacement character with end-of-line, if present
				}
				if StrLen(objHeader[A_Index])
					objData[objHeader[A_Index]] := strFieldData ; we have field names in objHeader[A_Index]
				else
					objData[A_Index] := strFieldData ; we don't have field names so we use index numbers as key/field names
			}
			objCollection.Insert(objData) ;  add the object (record) to the collection
		}
	}
	objCollection.SetCapacity(0) ; reallocates the object's internal array to fit only its current content
	if (intMaxProgress)
		Progress, Off
	objHeader := ; release object
	return objCollection
}
;================================================



;================================================
ObjCSV_Collection2CSV(objCollection, strFilePath, blnHeader := 0, strFieldOrder := "", blnProgress := 0, blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEndOfLine := "`r`n", strEolReplacement := "")
/*
Summary: Transfer the selected fields from a collection of objects to a CSV file. Field names taken from key names are optionally included in the CSV file. Delimiters are configurable.

RETURNED VALUE:
None.

PARAMETERS:

objCollection
Object containing an array of objects (or collection). Objects in the collection are associative arrays which contain a list key-value pairs. See ObjCSV_CSV2Collection returned value for details.

strFilePath
The name of the CSV file, which is assumed to be in %A_WorkingDir% if an absolute path isn't specified.

blnHeader := 0
Optional. If true, the key names in the collection objects are inserted as header of the CSV file. Fields names are delimited by the strFieldDelimiter character.

strFieldOrder := ""
Optional. List of field to include in the CSV file and the order of these fields in the file. Fields names must be separated by the strFieldDelimiter character and, if required, encapsulated by the strEncapsulator character. If empty, all fields are included. Empty by default.

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.

blnOverwrite := 0
Optional. If true (or 1), overwrite existing files. If false (or 0), content is appended to the existing file. False (or 0) by default. NOTE: If content is appended to an existing file, fields names and order should be the same as in the existing file.

strFieldDelimiter := ","
Optional. Delimiter inserted between fields in the CSV file. Also used as delimiter in the above parameter strFieldOrder. One character, usually comma, tab or semi-colon. You can choose other delimiters like pipe (|), space, etc. Comma by default. NOTE: End of line characters (`n or `r) are prohibited as field separator since they are used as record delimiters.

strEncapsulator := """"
Optional. One character (usualy double-quote) inserted in the CSV file to embed fields that include at least one of these special characters: line-breaks, field delimiters or the encapsulator character itself. In this last case, the encapsulator character is doubled in the string. For example: "one ""quoted"" word". Double-quote by default.

strEndOfLine := "`r`n"
Optional. Character(s) inserted between records at end-of-lines. Can be `r`n (carriage return+linefeed) or `n (linefeed alone) depending on OS text file formats. Carriage return + Linefeed by default.

strEolReplacement := ""
Optional. When empty, multi-line fields are saved unchanged. If not empty, end-of-line in multi-line fields are replaced by the character or string strEolReplacement. Empty by default. NOTE: Strings including replaced end-of-line will still be encapsulated with the strEncapsulator character.
*/
{
	strData := ""
	intMax := objCollection.MaxIndex()
	if (blnProgress)
		Progress, R0-%intMax% FS8 A, Saving data to CSV file..., , , MS Sans Serif
	if (blnHeader) ; put the field names (header) in the first line of the CSV file
	{
		if !StrLen(strFieldOrder) ; we dont have a header, so we take field names from the first record of objCollection, in their natural order 
		{
			for strFieldName, strValue in objCollection[1]
				strFieldOrder := strFieldOrder . Format4CSV(strFieldName, strFieldDelimiter, strEncapsulator) . strFieldDelimiter
			StringTrimRight, strFieldOrder, strFieldOrder, 1 ; remove extra field delimiter
		}
		strData := strFieldOrder . strEndOfLine ; put this header as first line of the file
	}
	Loop, %intMax% ; for each record in the collection
	{
		strRecord := "" ; line to add to the CSV file
		if (blnProgress) and !Mod(%A_index%, 5000)
			Progress, %A_index%
		if StrLen(strFieldOrder) ;  we put only these fields, in this order
		{
			intLineNumber := A_Index
			for intColIndex, strFieldName in ReturnDSVObjectArray(strFieldOrder, strFieldDelimiter, strEncapsulator) ; parse strFieldOrder handling encapsulated field names
			{
				strValue := objCollection[intLineNumber][Trim(strFieldName)]
				if (StrLen(strEolReplacement)) ; multiline field eol replacement
					StringReplace, strValue, strValue, `r`n, %strEolReplacement%, All ; handle multiline data fields
				strRecord := strRecord . Format4CSV(strValue, strFieldDelimiter, strEncapsulator) . strFieldDelimiter
			}
		}
		else ;  we put all fields in the record (I assume the order of fields is the same for each object)
			for strFieldName, strValue in objCollection[A_Index]
			{
				strValue := Format4CSV(strValue, strFieldDelimiter, strEncapsulator)
				if (StrLen(strEolReplacement))
					StringReplace, strValue, strValue, `r`n, %strEolReplacement%, All ; handle multiline data fields
				strRecord := strRecord . Format4CSV(strValue, strFieldDelimiter, strEncapsulator) . strFieldDelimiter
			}
		StringTrimRight, strRecord, strRecord, 1 ; remove extra field delimiter
		strData := strData . strRecord . strEndOfLine
	}
	if (blnOverwrite)
		FileDelete, %strFilePath%
	FileAppend, %strData%, %strFilePath%
	if (blnProgress)
		Progress, Off
	return
}
;================================================



;================================================
ObjCSV_Collection2Fixed(objCollection, strFilePath, strWidth, blnHeader := 0, strFieldOrder := "", blnProgress := 0, blnOverwrite := 0, strFieldDelimiter := ",", strEncapsulator := """", strEndOfLine := "`r`n", strEolReplacement := "")
/*
Summary: Transfer the selected fields from a collection of objects to a fixed-width file. Field names taken from key names are optionnaly included the file. Width are determined by the delimited string strWidth. Field names and data fields shorter than their width are padded with spaces. Field names and data fields longer than their width are truncated at their maximal width.

RETURNED VALUE:
None.

PARAMETERS:

objCollection
Object containing an array of objects (or collection). Objects in the collection are associative arrays which contain a list key-value pairs. See ObjCSV_CSV2Collection returned value for details.

strFilePath
The name of the CSV file, which is assumed to be in %A_WorkingDir% if an absolute path isn't specified.

strWidth
Width for each field. Each numeric values must be in the same order as strFieldOrder and separated by the strFieldDelimiter character.

blnHeader := 0
Optional. If true, the field names in the collection objects are inserted as header of the file, padded or truncated according to the width for each field. NOTE: If field names are longer than their fixed-width they will be truncated as well.

strFieldOrder := ""
Optional. List of field to include in the file and the order of these fields in the file. Fields names must be separated by the strFieldDelimiter character and, if required, encapsulated by the strEncapsulator character. If empty, all fields are included. Empty by default.

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.

blnOverwrite := 0
Optional. If true (or 1), overwrite existing files. If false (or 0), content is appended to the existing file. False (or 0) by default. NOTE: If content is appended to an existing file, fields names and order should be the same as in the existing file.

strFieldDelimiter := ","
Optional. Delimiter inserted between fields names in the strFieldOrder parameter and fields width in the strWidth parameter. This delimiter is NOT used in the file data. One character, usually comma, tab or semi-colon. You can choose other delimiters like pipe (|), space, etc. Comma by default. NOTE: End of line characters (`n or `r) are prohibited as field separator since they are used as record delimiters.

strEncapsulator := """"
Optional. One character (usualy double-quote) inserted in the strFieldOrder parameter to embed field names that include at least one of these special characters: line-breaks, field delimiters or the encapsulator character itself. In this last case, the encapsulator character is doubled in the string. For example: "one ""quoted"" word". Double-quote by default. This delimiter is NOT used in the file data.

strEndOfLine := "`r`n"
Optional. Character(s) inserted between records at end-of-lines. Can be `r`n (carriage return+linefeed) or `n (linefeed alone) depending on OS text file formats. Carriage return + Linefeed by default.

strEolReplacement := ""
Optional. A fixed-width file should not include end-of-line within data. If it does and it a strEolReplacement is provided, end-of-line in multi-line fields are replaced by the string strEolReplacement and this (or these) characters are included in the fixed-width character count. Empty by default.
*/
{
	StringSplit, arrIntWidth, strWidth, %strFieldDelimiter% ; arrIntWidth is a pseudo-array, so %arrIntWidth1% or arrIntWidth%intColIndex%
	strData := ""
	intMax := objCollection.MaxIndex()
	if (blnProgress)
		Progress, R0-%intMax% FS8 A, Saving data to export file..., , , MS Sans Serif
	if (blnHeader) ; put the field names (header) in the first line of the file
	{
		strHeaderFixed := ""
		if StrLen(strFieldOrder) ; convert DSV string to fixed-width
		{
			for intColIndex, strFieldName in ReturnDSVObjectArray(strFieldOrder, strFieldDelimiter, strEncapsulator) ; parse strFieldOrder handling encapsulated field names
			{
				; ###_D("strFieldName: " . strFieldName . "`nintColIndex:" . intColIndex)
				strHeaderFixed := strHeaderFixed . MakeFixedWidth(strFieldName, arrIntWidth%intColIndex%)
				; ###_D("arrIntWidth%intColIndex%: " . arrIntWidth%intColIndex% . "`nstrHeaderFixed: " . strHeaderFixed . "`nStrLen(strHeaderFixed):" . StrLen(strHeaderFixed))
			}
			; ###_D("HEADER FIX AVEC strFieldOrder / strHeaderFixed:`n" . strHeaderFixed, 1)
		}
		else ; we dont have a header, so we take field names from the first record of objCollection, in their natural order 
		{
			intColIndex := 1
			for strFieldName, strValue in objCollection[1]
			{
				strHeaderFixed := strHeaderFixed . MakeFixedWidth(strFieldName, arrIntWidth%intColIndex%)
				intColIndex := intColIndex + 1
			}
			; ###_D("HEADER FIX PAS DE strFieldOrder / strHeaderFixed:`n" . strHeaderFixed, 1)
		}
		strData := strHeaderFixed . strEndOfLine ; put this header as first line of the file
		; ###_D("AVEC HEADER strData:`n" . strData, 1)
	}
	Loop, %intMax% ; for each record in the collection
	{
		strRecord := "" ; line to add to the file
		if (blnProgress) and !Mod(%A_index%, 5000)
			Progress, %A_index%
		if StrLen(strFieldOrder) ; we put only these fields, in this order
		{
			; ###_D("DATA AVEC strFieldOrder", 1)
			intLineNumber := A_Index
			for intColIndex, strFieldName in ReturnDSVObjectArray(strFieldOrder, strFieldDelimiter, strEncapsulator) ; parse strFieldOrder handling encapsulated field names
			{
				strValue := objCollection[intLineNumber][Trim(strFieldName)]
				; ###_D("strFieldName: " . strFieldName . "`nstrValue: " . strValue . "`nintColIndex:" . intColIndex)
				if (StrLen(strEolReplacement)) ; multiline field eol replacement
					StringReplace, strValue, strValue, `r`n, %strEolReplacement%, All ; handle multiline data fields
				strRecord := strRecord . MakeFixedWidth(strValue, arrIntWidth%intColIndex%)
				; ###_D("arrIntWidth%intColIndex%: " . arrIntWidth%intColIndex% . "`nstrRecord: " . strRecord . "`nStrLen(strRecord):" . StrLen(strRecord))
			}
		}
		else ;  we put all fields in the record (I assume the order of fields is the same for each object)
		{
			; ###_D("DATA SANS strFieldOrder")
			intColIndex := 1
			for strFieldName, strValue in objCollection[A_Index]
			{
				; ###_D("strFieldName: " . strFieldName . "`nstrValue: " . strValue . "`nintColIndex:" . intColIndex)
				if (StrLen(strEolReplacement))
					StringReplace, strValue, strValue, `r`n, %strEolReplacement%, All ; handle multiline data fields
				strRecord := strRecord . MakeFixedWidth(strValue, arrIntWidth%intColIndex%)
				; ###_D("arrIntWidth%intColIndex%: " . arrIntWidth%intColIndex% . "`nstrRecord: " . strRecord . "`nStrLen(strRecord):" . StrLen(strRecord))
				intColIndex := intColIndex + 1
			}
		}
		strData := strData . strRecord . strEndOfLine
		; ###_D("strData:`n" . strData, 1)
	}
	if (blnOverwrite)
		FileDelete, %strFilePath%
	FileAppend, %strData%, %strFilePath%
	if (blnProgress)
		Progress, Off
	return
}
;================================================



;================================================
ObjCSV_Collection2ListView(objCollection, strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strEncapsulator := """", strSortFields := "", strSortOptions := "", blnProgress := 0)
/*
Summary: Transfer the selected fields from a collection of objects to ListView. The collection can be sorted by the function. Field names taken from the objects keys are used as header for the ListView. NOTE-1: Due to an AHK limitation, files with more that 200 fields will not be transfered to a ListView. NOTE-2: Although any length of text can be stored in each cell of a ListView, only the first 260 characters are displayed (no lost data).

RETURNED VALUE:
None.

PARAMETERS:

objCollection
Object containing an array of objects (or collection). Objects in the collection are associative arrays which contain a list key-value pairs. See ObjCSV_CSV2Collection returned value for details. NOTE: Multi-line fields can be inserted in a ListView and retreived from a ListView. However, take note that end-of-lines will not be visible in cells with current version of AHK_L (v1.1.09.03).

strGuiID := ""
Optional. Name of the Gui that contains the ListView where the collection will be displayed. If empty, the last default Gui is used. Empty by default. NOTE: If a Gui name is provided, this Gui will remain the default Gui at the termination of the function.

strListViewID := ""
Optional. Name of the target ListView where the collection will be displayed. If empty, the last default ListView is used. The target ListView should be empty or should contain data in the same columns number and order than the data to display. If this is not respected, new columns will be added to the right of existing columns and new rows will be added at the bottom of existing data. Empty by default. NOTE-1: Performance is greatly improved if we provide the ListView ID because we avoid redraw during import. NOTE-2: If a ListView name is provided, this ListView will remain the default at the termination of the function.

strFieldOrder := ""
Optional. List of field to include in the ListView and the order of these columns. Fields names must be separated by the strFieldDelimiter character. If empty, all fields are included. Empty by default.

strFieldDelimiter := ","
Optional. Delimiter of the fields in the strFieldOrder parameter. One character, usually comma, but can also be tab, semi-colon, pipe (|), space, etc. Comma by default.

strEncapsulator := """"
Optional. One character (usualy double-quote) possibly used in the in the strFieldOrder string to embed fields that would include special characters (as described above).

strSortFields := ""
Optional. Field(s) value(s) used to sort the collection before its insertion in the ListView. To sort on more than one field, concatenate field names with the + character (e.g. "LastName+FirstName"). Faster sort can be obtained by manualy clicking on columns headers in the ListView after the collection has been inserted. Empty by default.

strSortOptions := ""
Optional. Sorting options to apply to the sort command above. A string of zero or more of the option letters (in any order, with optional spaces in between). Most frequently used are R (reverse order) and N (numeric sort). All AHK_L sort options are supported. See http://l.autohotkey.net/docs/commands/Sort.htm for more options. Empty by default.

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.
*/
{
	objHeader := Object() ; holds the keys (fields name) of the objects in the collection
	if StrLen(strSortFields)
		objCollection := ObjCSV_SortCollection(objCollection, strSortFields, strSortOptions, blnProgress)
	intMax := objCollection.MaxIndex()
	if (blnProgress)
		Progress, R0-%intMax% FS8 A, Loading data to ListView..., , , MS Sans Serif
	if StrLen(strGuiID)
		Gui, %strGuiID%:Default
	if StrLen(strListViewID)
		GuiControl, -Redraw, %strListViewID% ;  stop drawing the ListView during import
	if StrLen(strListViewID)
		Gui, ListView, %strListViewID% ; sets the default ListView in the default Gui
	if !StrLen(strFieldOrder) ; if we dont have fields restriction or order, take all fields in their natural order in the first records
	{
		for strFieldName, strValue in objCollection[1] ;  use the first record to get the field names
			strFieldOrder := strFieldOrder . strFieldName . strFieldDelimiter
		StringTrimRight, strFieldOrder, strFieldOrder, 1 ; remove extra field delimiter
	}
	objHeader := ReturnDSVObjectArray(strFieldOrder, strFieldDelimiter, strEncapsulator) ; returns an object array from a delimited-separated-value string
	if objHeader.MaxIndex() > 200 ; ListView cannot display more that 200 columns
	{
		if (blnProgress)
			Progress, Off
		return ; displays nothing in the ListView (more error friendly code could be added here)
	}
	for intIndex, strFieldName in objHeader
	{
		LV_GetText(strExistingFieldName, 0, intIndex) ;  line 0 returns column names
		if (Trim(strFieldName) <> strExistingFieldName)
			LV_InsertCol(intIndex, "", Trim(strFieldName))
	}
	loop, %intMax%
	{
		if (blnProgress) and !Mod(%A_index%, 5000)
			Progress, %A_Index%
		intRowNumber := A_Index
		arrFields := Array() ; will contain the values for each cell of a new row
		for intIndex, strFieldName in objHeader
			arrFields[intIndex] := objCollection[intRowNumber][Trim(strFieldName)] ; for each field, in the specified order, add the data to the array
		LV_Add("", arrFields*) ; put each item of the array in cells of a new ListView row
		; "arrFields*" is allowed because LV_Add is a variadic function (see http://www.autohotkey.com/board/topic/92531-lv-add-to-add-an-array/)
	}
	Loop, % arrFields.MaxIndex()
		LV_ModifyCol(A_Index, "AutoHdr") ;  adjust width of each column according to their content
	if StrLen(strListViewID)
		GuiControl, +Redraw, %strListViewID% ; redraw the ListView
	if (blnProgress)
		Progress, Off
	Gui, Show
	objHeader := ; release object
}
;================================================


;================================================
ObjCSV_ListView2Collection(strGuiID := "", strListViewID := "", strFieldOrder := "", strFieldDelimiter := ",", strEncapsulator := """", blnProgress := 0)
/*
Summary: Transfer the selected lines of the selected columns of a ListView to a collection of objects. Lines are transfered in the order they appear in the ListView. Column headers are used as objects keys.

RETURNED VALUE:
This functions returns an object that contains a collection (or array of objects). See ObjCSV_CSV2Collection returned value for details.

PARAMETERS:
strGuiID := ""
Optional. Name of the Gui that contains the ListView where is the data to transfer. If empty, the last default Gui is used. Empty by default. NOTE: If a Gui name is provided, this Gui will remain the default Gui at the termination of the function.

strListViewID := ""
Optional. Name of the target ListView where is the data to transfer. If empty, the last default ListView is used. If one or more rows in the ListView are selected, only these rows will be inserted in the collection. Empty by default. NOTE: If a ListView name is provided, this ListView will remain the default at the termination of the function.

strFieldOrder := ""
Optional. Name of the fields (or ListView columns) to insert in the collection records. Names are separated by the strFieldDelimiter character (see below). If empty, all fields are transfered. Empty by default.

strFieldDelimiter := ","
Optional. Delimiter of the fields in the strFieldOrder parameter. One character, usually comma, but can also be tab, semi-colon, pipe (|), space, etc. Comma by default.

strEncapsulator := """"
Optional. One character (usualy double-quote) possibly used in the in the strFieldOrder string to embed fields data or field names that would include special characters (as described above).

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.
*/
{
	objCollection := Object() ; object that will be returned by the function (a collection or array of objects)
	if StrLen(strGuiID)
		Gui, %strGuiID%:Default
	if StrLen(strListViewID)
		Gui, ListView, %strListViewID%
	intNbCols := LV_GetCount("Column") ; get the number of columns in the ListView
	intNbRows := LV_GetCount() ; get the number of lines in the ListView
	blnSelected := (LV_GetCount("Selected") > 0) ; we will read only selected rows
	objHeaderPositions := Object() ; holds the keys (fields name) of the objects in the collection and their position in the ListView
	; build an object array with field names and their position in the ListView header
	loop, %intNbCols%
	{
		LV_GetText(strFieldHeader, 0, A_Index)
		objHeaderPositions.Insert(strFieldHeader, A_Index)
	}
	if !(StrLen(strFieldOrder)) ; if empty, we build strFieldOrder from the ListView header
	{
		loop, %intNbCols%
		{
			LV_GetText(strFieldHeader, 0, A_Index)
			strFieldOrder := strFieldOrder . Format4CSV(strFieldHeader, strFieldDelimiter, strEncapsulator) . strFieldDelimiter ; handle field named with special characters requiring encapsulation
		}
		StringTrimRight, strFieldOrder, strFieldOrder, 1
	}
	intRowNumber := 0 ; scan each row or selected row of the ListView
	Loop
	{
		if (blnSelected)
			intRowNumber := LV_GetNext(intRowNumber) ; get next selected row number
		else
			intRowNumber := intRowNumber + 1 ; get next row number
		if (not intRowNumber) OR (intRowNumber > intNbRows) ; we passed the last row or the last selected row of the ListView
			break
		objData := Object() ; add row data to a new object in the collection
		for intIndex, strFieldName in ReturnDSVObjectArray(strFieldOrder, strFieldDelimiter, strEncapsulator) ; parse strFieldOrder handling encapsulated fields
		{
			LV_GetText(strFieldData, intRowNumber, objHeaderPositions[Trim(strFieldName)]) ; get data from cell at row number/header position ListView
			objData[strFieldName] := strFieldData ; put data in the appropriate field of the new row
		}
		objCollection.Insert(objData)
	}
	objCollection.SetCapacity(0) ; This reallocates the object's internal array to fit only its current content
	objHeaderPositions := ; release object
	return objCollection
}
;================================================


;================================================
ObjCSV_SortCollection(objCollection, strSortFields, strSortOptions := "", blnProgress := 0)
/*
Summary: Scan a collection of objects, sort the collection on one or more field and return sorted collection. Standard AHK_L sort options are supported.

RETURNED VALUE:
This functions returns an object that contains the array (or collection) of objects of objCollection sorted on strSortFields. See ObjCSV_CSV2Collection returned value for details.

PARAMETERS:
objCollection
Object containing an array of objects (or collection). Objects in the collection are associative arrays which contain a list key-value pairs. See ObjCSV_CSV2Collection returned value for details.

strSortFields
Name(s) of the field(s) to use as sort criteria. To sort on more than one field, concatenate field names with the + character (e.g. "LastName+FirstName").

strSortOptions := ""
Optional. Sorting options to apply to the sort command. A string of zero or more of the option letters (in any order, with optional spaces in between). Most frequently used are R (reverse order) and N (numeric sort). All AHK_L sort options are supported. See http://l.autohotkey.net/docs/commands/Sort.htm for more options. Empty by default.

blnProgress := 0
Optional. If true (or 1), a progress bar is displayed. Should be use only for very large collections. False (or 0) by default.
*/
{
	objCollectionSorted := Object() ; Array (or collection) of sorted objects returned by this function. See ObjCSV_CSV2Collection returned value for details.
	objCollectionSorted.SetCapacity(objCollection.MaxIndex())
	strIndexDelimiter := "|" ;
	intTotalRecords := objCollection.MaxIndex()
	if (blnProgress)
		Progress, R0-%intTotalRecords% FS8 A, Sorting data..., , , MS Sans Serif
	strIndex := ""
	; The variable strIndex is a multi-line string used as an index to sort the collection.
	; Each line of the index contains the sort values and record numbers separated by the pipe (|) character.
	; For example:
	;   value_one|1
	;   value_two|2
	;   value_three|3
	; This string is sorted using the standard AHK_L Sort command:
	;   value_one|1
	;   value_three|3
	;   value_two|2
	; The sorted string is used as an index to sort the records in objCollectionSorted according to the sorting values.
	; In our example, the objects will be added to the sorted collection in this order: 1, 3, 2.
	;
	; Because strIndex can be quite large, we gain performance by splitting the string in substrings of around 300 kb.
	; See discussion on AHK forum http://www.autohotkey.com/board/topic/92832-tip-large-strings-performance-or-divide-to-conquer/
	intOptimalSizeOfSubstrings := 300000 ; found by trial and error - no impact on results if not the optimal size
	strSubstring := ""
	Loop, %intTotalRecords% ; populate index substrings
	{
		intRecordNumber := A_Index
		if (blnProgress) and !Mod(%A_index%, 5000)
			Progress, %A_Index%
		if InStr(strSortFields, "+")
		{
			strSortingValue := ""
			Loop, Parse, strSortFields, +
				strSortingValue := strSortingValue . objCollection[intRecordNumber][A_LoopField] . "+"
		}
		else
			strSortingValue := objCollection[intRecordNumber][strSortFields]
		StringReplace, strSortingValue, strSortingValue, %strIndexDelimiter%, , 1 ; suppress all index delimiters inside sorting values
		StringReplace, strSortingValue, strSortingValue, `n, , 1 ; suppress all end-of-lines characters inside sorting values 
		strSubstring := strSubstring . strSortingValue . strIndexDelimiter . intRecordNumber . "`n"
		if StrLen(strSubstring) > intOptimalSizeOfSubstrings
		{
			strIndex := strIndex . strSubstring ;  add this substring to the final string
			strSubstring := "" ; start a new substring
		}
	}
	strIndex := strIndex . strSubstring ; add the last substring to the final string
	StringTrimRight, strIndex, strIndex, 1
	Sort, strIndex, %strSortOptions%
	Loop, Parse, strIndex, `n
	{
		StringSplit, arrRecordKey, A_LoopField, %strIndexDelimiter% ; get the record numbers in the original collection in the order the have to be inserted in the sorted collection
		objCollectionSorted.Insert(objCollection[arrRecordKey2])
	}
	objCollectionSorted.SetCapacity(0) ; this reallocates the object's internal array to fit only its current content
	if (intProgress)
		Progress, Off
	return objCollectionSorted
}
;================================================



;******************************************************************************************************************** 
; INTERNAL FUNCTIONS
;******************************************************************************************************************** 

Prepare4Multilines(ByRef strCsvData, strFieldEncapsulator := """", blnProgress := 0)
/*
Summary: Replace end-of-line characters (`n) in field data in strCsvData with a replacement character in order to make data rows stand on a single-line before they are processed by the "Loop, Parse" command. A safe replacement character (absent from the strCsvData string) is automatically determined by the function.

RETURNED VALUE:
The replacement character for end-of-lines. Usualy ¡ (inverted exclamation mark, ASCII 161) or the next available safe character: ¢ (ASCII 162), £ (ASCII 163), ¤ (ASCII 164), etc.  The caller of this function *must* put this value in a variable and *must* do the reverse replacement with `n at the appropriate step inside the "Loop, Parse" command.

PARAMETERS:
ByRef strCsvData
Input/Output data string processed by the function and returned to the caller after all end-of-line characters (`n) have been replaced with the safe replacement character.

strFieldEncapsulator
Character used in the strCsvData data to embed fields that include line-breaks. Double-quote by default.

blnProgress
If true (or 1), a progress bar is displayed. Default false (or 0).

CALL-FOR-HELP!
#1 This function uses a very rudimentary algorithm to do the replacements only when the end-of-line charaters are enclosed between double-quotes. I'm confident my code is safe. But there is certainly a more efficient way to accomplish this: RegEx command or another approach? Any help appreciated here :-)
#2 Need help to test it / make sure this work with ASCII files with end-of-line character other than `n (works well on DOS files, need to be tested on Unix or Mac text files -  see http://peterbenjamin.com/seminars/crossplatform/texteol.html)

*/
{
	if (blnProgress)
	{
		intProgress := StrLen(strCsvData)
		Progress, R0-%intProgress% FS8 A, Preparing CSV data for loading..., , , MS Sans Serif
	}
	intEolReplacementAsciiCode := GetFirstUnusedAsciiCode(strCsvData) ; Usualy ¡ (inverted exclamation mark, ASCII 161)
	blnInsideEncapsulators := false
	Loop, Parse, strCsvData ; parsing on a temporary copy of strCsvData -  so we can update the original strCsvData inside the loop
	{
		if (blnProgress AND !Mod(A_Index, 5000))
			Progress, %A_index% ;  update progress bar only every 5000 chars
		if (A_Index = 1)
			strCsvData := ""
		if (blnInsideEncapsulators AND A_Loopfield = "`n")
			strCsvData := strCsvData . Chr(intEolReplacementAsciiCode)
		else
			strCsvData := strCsvData . A_Loopfield
		if (A_Loopfield = strFieldEncapsulator)
			blnInsideEncapsulators := !blnInsideEncapsulators ; beginning or end of encapsulated text
	}
	if (blnProgress)
		Progress, Off
	return Chr(intEolReplacementAsciiCode)
}



GetFirstUnusedAsciiCode(strData, intAscii := 161)
/*
Summary: Returns the ASCII code of the first character absent from the strData string, starting at ASCII code intAscii. By default, ¡ (inverted exclamation mark ASCII 161) or the next available character: ¢ (ASCII 162), £ (ASCII 163), ¤ (ASCII 164), etc.
*/
{
	Loop
		if InStr(strData, Chr(intAscii))
			intAscii := intAscii + 1
		else
			break
	return intAscii
}



MakeFixedWidth(strFixed, intWidth)
{
	while StrLen(strFixed) < intWidth
		strFixed := strFixed . " "
	return SubStr(strFixed, 1, intWidth)
}



Format4CSV(F4C_String, strFieldDelimiter := ",", strEncapsulator := """")
/*
Format4CSV by Rhys (http://www.autohotkey.com/forum/topic27233.html)
Added the strFieldDelimiter parameter to make it work with locales with semi-colon separator
Added the strEncapsulator parameter to make it work with other encapsultors than double-quotes
*/
{
   Reformat:=False ; Assume String is OK
   IfInString, F4C_String,`n ; Check for linefeeds
      Reformat:=True ; String must be bracketed by double-quotes
   IfInString, F4C_String,`r ; Check for linefeeds
      Reformat:=True
   IfInString, F4C_String,%strFieldDelimiter% ; Check for field delimiter
      Reformat:=True
   IfInString, F4C_String, %strEncapsulator% ; Check for encapsulator
   {
      Reformat:=True
      StringReplace, F4C_String, F4C_String, %strEncapsulator%, %strEncapsulator%%strEncapsulator%, All ; The original encapsulator need to be double encapsulator
   }
   If (Reformat)
      F4C_String= %strEncapsulator%%F4C_String%%strEncapsulator% ; If needed, bracket the string in encapsulators
   Return, F4C_String
}


ReturnDSVObjectArray(CurrentDSVLine, Delimiter=",", Encapsulator="""")
/*
Summary: Returns an object array from a delimiter-separated string.
Based on ReturnDSVArray by DerRaphael (thanks for regex hard work).
See Delimiter Seperated Values by DerRaphael (http://www.autohotkey.com/forum/post-203280.html#203280)
*/
{
	objReturnObject := Object()             ; create a local object array that will be returned by the function
	if ((StrLen(Delimiter)!=1)||(StrLen(Encapsulator)!=1))
	{
		return                              ; return empty object indicating an error
	}
	strPreviousFormat := A_FormatInteger    ; save current interger format
	SetFormat,integer,H                     ; needed for escaping the RegExNeedle properly
	d := SubStr(ASC(delimiter)+0,2)         ; used as hex notation in the RegExNeedle
	e := SubStr(ASC(encapsulator)+0,2)      ; used as hex notation in the RegExNeedle
	SetFormat,integer,%strPreviousFormat%   ; restore previous integer format

	p0 := 1                                 ; Start of search at char p0 in DSV Line
	fieldCount := 0                         ; start off with empty fields.
	CurrentDSVLine .= delimiter             ; Add delimiter, otherwise last field
	;                                         won't get recognized
	Loop
	{
		RegExNeedle := "\" d "(?=(?:[^\" e "]*\" e "[^\" e "]*\" e ")*(?![^\" e "]*\" e "))"
		p1 := RegExMatch(CurrentDSVLine,RegExNeedle,tmp,p0)
		; p1 contains now the position of our current delimitor in a 1-based index
		fieldCount++                        ; add count
		field := SubStr(CurrentDSVLine,p0,p1-p0)
		; This is the Line you'll have to change if you want different treatment
		; otherwise your resulting fields from the DSV data Line will be stored in an object array
		if (SubStr(field,1,1)=encapsulator)
		{
			; This is the exception handling for removing any doubled encapsulators and
			; leading/trailing encapsulator chars
			field := RegExReplace(field,"^\" e "|\" e "$")
			StringReplace,field,field,% encapsulator encapsulator,%encapsulator%, All
		}
		objReturnObject.Insert(Trim(field)) ; add an item in the object array and assign our value to it
		;                                     Trim not in the original ReturnDSVArray but added for my script needs
		if (p1=0)
		{                                   ; p1 is 0 when no more delimitor chars have been found
			 objReturnObject.Remove()       ; so remove last item in the object array due to last appended delimitor
			 Break                          ; and exit Loop
		} Else
			p0 := p1 + 1                    ; set the start of our RegEx Search to last result
	}                                       ; added by one
	return objReturnObject                  ; return the object array to the function caller
}
