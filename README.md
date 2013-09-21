ObjCSV Library - Library to load and save CSV files to/from objects and ListView
------------------------------------------------------------------------
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By JnLlnd on AHK forum
2013-08-22+

A set of AutoHotkey_L (AHK) functions to load from CSV files, sort, display and save collections of records using the Object data type. Files can be read and saved in any delimited format (CSV, semi-colon, tab delimited, single-line or multi-line, etc.). Collections can also be displayed, edited and read in Gui ListView objects. For more info on CSV files, see http://en.wikipedia.org/wiki/Comma-separated_values.

Even if you don't know much about AHK objects, simply using these functions will help to:
- Transform a tab or semi-colon delimited CSV file to a straight coma-delimited file (any single character delimited is supported).
- Load a CSV file with multi-line fields (for example, Notes fields in a Google Contact or Outlook tasks export) and save it in a single line CSV file (with the end-of-line replacement character of your choice) that can be imported easily in Excel.

Other usages:
- Load a file to object to run any scripted manipulation on the content of the file with the ease and safety of AHK objets.
- Add/change CSV header names, change the order of fields or remove fields in a CSV file programatically.
- Display the file content in a ListView for further viewing or editing (multiple Gui and ListView controls are supported).
- Sort the data on any field combination before loading to the ListView or saving to a CSV file.
- Save all or selected rows of a ListView to a CSV file.
- Save to a file with or without header, with the fields delimiter and encapsulator of your choice.

The most up-to-date version of this library can be found on GitHub:
https://github.com/JnLlnd/ObjCSV

INSTRUCTIONS

Copy this script in a file named ObjCSV.ahk and save this file in one of these \Lib folders:
  %A_ScriptDir%\Lib\
  %A_MyDocuments%\AutoHotkey\Lib\
  [path to the currently running AutoHotkey_L.exe]\Lib\ (AutoHotKey_L can be downloaded here: http://l.autohotkey.net/ )

You can use the functions in this library by calling ObjCSV_FunctionName. No #Include required!


FUNCTIONS QUICK REFERENCE

ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = true, blnMultiline = 1, intProgress = 0, strFieldDelimiter = ",", strEncapsulator = """", strRecordDelimiter = "`n", strOmitChars = "`r"])
Transfer the content of a CSV file to a collection of objects. Field names are taken from the first line of the file or from the strFieldNames parameter. Delimiters are configurable.

ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0, strFieldOrder = "", blnProgress = 0, blnOverwrite = 0, strFieldDelimiter = ",", strEncapsulator = """", strEndOfLine = "`n", strEolReplacement = ""])
Transfer the selected fields from a collection of objects to a CSV file. Field names taken from key names are included or not in the CSV file. Delimiters are configurable.

ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", strSortFields = "", strSortOptions = "", blnProgress = "0"])
Transfer the selected fields from a collection of objects to ListView. The collection can be sorted by the function. Field names taken from the objects keys are used as header for the ListView.

ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ",", strEncapsulator = """", blnProgress = "0"])
Transfer the selected lines of the selected columns of a ListView to a collection of objects. Lines are transfered in the order they appear in the ListView. Column headers are used as objects keys.

ObjCSV_SortCollection(objCollection, strSortFields [, strSortOptions = "", blnProgress = 0])
Sort the objectfs of a collection using one or multiple fields as sorting key.  Internal AHK Sort options are supported.

DISCUSSION AND TUTORIAL

A discussion on this library and tutorials can be found on the AutoHotkey forum:
http://www.autohotkey.com/board/topic/96618-lib-objcsv-library-v01-library-to-load-and-save-csv-files-tofrom-objects-and-listview/
http://www.autohotkey.com/board/topic/96619-objcsv-library-tutorial-basic/
http://www.autohotkey.com/board/topic/97147-parsing-csv-files-with-multi-line-fields/
