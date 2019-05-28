Attribute VB_Name = "Utils"

'**************************************************************************
'********************* FUNCTIONS PACKAGE ***********************************
'**************************************************************************
'# Version 2.3.7
'# All functions are assumed to be compatible with non .xls formats only

'# Updated on 8/1/2012 - Added new parameter on lastColNum function
'# Updated on 7/12/2012 - lastColNum function - Revised to use a loop as the primary last row detection mechanism instead of lastUsedCol built in function (was causing too many problems)
'# Updated on 7/9/2012: Refactored getMonthNameAndYear function to use built-in Format function instead of a dictionary structure (Thanks to Randal B!)
'                                           Updated move_entries to make the move of the header row as optional
'# Updated on 7/3/2012: Added vbaLookup2 function. Many times faster than vbalookup, and exponentially faster than vlookup. See function for more details.


Public Declare Function GetTickCount Lib "kernel32.dll" () As Long '//

'**************************************************************************
Public Function isxls(ByVal ws As Worksheet) As Boolean
'// Tests to see whether a given worksheet is in the old excel format.
'// Authored by Nathan N on 7/6/2012

If Right(ws.Name, 4) = ".xls" Then isxls = True

End Function
Public Sub trim_data_fields(ByVal ws As Worksheet, ByVal fieldName As String)
'// Authored by Nathan N on 6/21/2012

Dim i As Long

For i = 2 To lastRow(ws)
    ws.Cells(i, FieldColNum(ws, fieldName)) = Trim(ws.Cells(i, FieldColNum(ws, fieldName)))
Next i

End Sub
Public Sub move_all_entries(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
'// Authored by  Nathan N on 6/28/2012
'// Moves all data from one worksheet to another. Strips all data of any previous formatting as well.
'// Assumes that data has no 'pre headers' in the first few rows

wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(lastRow(wsSource), lastColNum(wsSource))) = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow(wsSource), lastColNum(wsSource))).value

End Sub
Public Sub mergeWorksheets(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet, ByVal sourceKey As String, ByVal destKey As String)
'// Authored by  Nathan N on 1/14/2015
'// Merges all fields of a source worksheet into a destination worksheet
'// sourceKey must be located in column 1 (or col "A")

Dim sourceColCount As Integer, destColCount As Integer
Dim destFieldName As String
Dim i As Long, j As Long

sourceColCount = lastColNum(wsSource)
destColCount = lastColNum(wsDest)

If wsSource.Cells(1, 1) <> sourceKey Then
    MsgBox ("SourceKey is not located in column 1 of the Source worksheet! Quitting mergeWorksheets routine!")
    Exit Sub
End If
Application.ScreenUpdating = False

For i = 1 To sourceColCount
    If wsSource.Cells(1, i) = sourceKey Then GoTo NextColumn
    '//destColNum = destColCount + i
    destFieldName = wsSource.Cells(1, i)
    vbaLookup2 wsDest, destFieldName, destKey, wsSource, sourceKey, destFieldName, , , , , , True
NextColumn:
Next i

Application.ScreenUpdating = True


End Sub
Public Sub getMonthNameAndYear(ByVal ws As Worksheet, ByVal dateField As Worksheet, newDateFieldName As String)
'// Inserts a condensed month name and year based on the original date field
'// Added 7/2/2012 by Nathan N
'// Updated 7/9/2012: Refactored function to use built-in Format function instead of a dictionary structure (Thanks to Randal B!)

Dim dateColNum As Integer, monthColNum As Integer
Dim dateColLet As String
Dim dateValue As Date, monthValue As Integer, yearValue As Integer, fullMonthName As String
Dim newDateValue As String
Dim i As Long

Application.ScreenUpdating = False

Set ws = ActiveWorkbook.ActiveSheet
dateColLet = InputBox("Enter the Date Column Letter:")

Application.ScreenUpdating = False

dateColNum = LetToColNum(dateColLet)
monthColNum = dateColNum + 1
ws.Columns(monthColNum).Insert
ws.Columns(monthColNum).NumberFormat = "@"
ws.Cells(1, monthColNum) = "Month"

For i = 2 To lastRow(ws)
    dateValue = ws.Cells(i, dateColNum)
    newDateValue = Format(dateValue, "mmm yyyy")
    ws.Cells(i, monthColNum) = newDateValue
Next i

Application.ScreenUpdating = True

End Sub
Public Sub move_entries(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet, ByVal fieldName As String, ByVal keywordInFieldName As String, Optional ByVal copyHeader As Boolean, Optional ByVal deleteMovedEntries As Boolean)
Dim startRow As Long, endRow As Long, startRowDest As Long, endRowDest As Long
'// Moves larges swathes of entries to a destination sheet quickly and as values (removes all formatting)
'// Authored by Nathan N on 6/28/2012
'// Updated 'itemToMove' to 'keywordInFieldName'
'// Updated on 7/9/2012 to make the copy over of the Header row as optional (header row is assumed to be located on row 1)
Application.ScreenUpdating = False

SortCol wsSource, fieldName
startRow = FieldRowNum(wsSource, keywordInFieldName, FieldColNum(wsSource, fieldName))
endRow = FieldRowNum(wsSource, keywordInFieldName, FieldColNum(wsSource, fieldName), , , , True)

If startRow = 0 Then '// if no matching keywordInFieldName was found then function exits
    Debug.Print "Function move_entries: keywordInFieldName was not found. No entries moved."
    Exit Sub
End If

If copyHeader = True Then wsDest.Rows(1) = wsSource.Rows(1).value

startRowDest = lastRow(wsDest) + 1
endRowDest = endRow - startRow + startRowDest
wsDest.Range(wsDest.Cells(startRowDest, 1), wsDest.Cells(endRowDest, lastColNum(wsSource))) = wsSource.Range(wsSource.Cells(startRow, 1), wsSource.Cells(endRow, lastColNum(wsSource))).value

If deleteMovedEntries = True Then wsSource.Rows(startRow & ":" & endRow).Delete

End Sub
Public Sub convert_string_to_int(ByVal ws As Worksheet, ByVal zipcodeFieldName As String)
'// Authored by Nathan N on 7/5/2012
'// Converts numbers that are in string format to int

Dim i As Long
Dim zipColNum As Integer

On Error Resume Next

zipColNum = FieldColNum(ws, zipcodeFieldName)

For i = 2 To lastRow(ws, , FieldColNum(ws, zipcodeFieldName))
    ws.Cells(i, zipColNum) = Int(ws.Cells(i, zipColNum))
Next i

On Error GoTo 0

End Sub
Public Sub zipcode_county_lookup_manual() '(ByVal ws As Worksheet, ByVal zipcodeFieldName As String, Optional ByVal stateFieldName As String, Optional ByVal singleStateName As String, Optional ByVal listStates As Boolean)
'// Matches zip codes with county names in worksheets with multiple states listed
'// Authored by Nathan N on 6/26/2012

Dim wb As Workbook, ws As Worksheet, wbCounty As Workbook, wsCounty As Worksheet, wsState As Worksheet
Dim state As String, zipCode As Long
Dim colLet As String
Dim zipColLet As String, zipColNum As Integer, countyColNum As Integer, stateColNum As Integer
Dim zipToCountyMappingPath As String
Dim rngCounty As Range
Dim stateFieldName As String, zipcodeFieldName As String, singleStateName As String, listStates As Boolean
Dim i As Long

Application.DisplayAlerts = False
Application.ScreenUpdating = False

zipToCountyMappingPath = file_reference_path() & "\ZipCodeAndCountyMappingDataPub.xlsx"

Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet

If Right(wb.Name, 4) = ".xls" Then
    MsgBox (FirstName() & ", it looks like you're using a workbook that's saved in the old excel format." & vbCrLf & vbCrLf & "Please resave your workbook as an .xlsx format. Then close and reopen.")
    Exit Sub
End If

'// Attempt to auto detect the zip code column
zipColNum = FieldColNum(ws, "zip", , True)
If zipColNum = 0 Then zipColNum = FieldColNum(ws, "zip c", , True)
If zipColNum = 0 Then zipColNum = FieldColNum(ws, "zip_c", , True)
If zipColNum = 0 Then zipColNum = FieldColNum(ws, "zip_", , True)
If zipColNum = 0 Then zipColNum = FieldColNum(ws, "zip", , True)
If zipColNum = 0 Then
    zipColLet = InputBox("Enter the zip code column letter:")
    zipColNum = LetToColNum(zipColLet)
End If

'// This corrects data for which zip codes are listed as Strings on a work sheet instead of int. String data types will cause the vbalookup function to not find the zip codes.
zipcodeFieldName = ws.Cells(1, zipColNum)

convert_string_to_int ws, zipcodeFieldName

countyColNum = zipColNum + 1

If zipColNum = -1 Then
    Debug.Print ("Error: zipcode_county_lookup function. Zip Code Column could not be located. Ended function.")
    Exit Sub
End If

If Not (FieldExists(ws, "County")) Then
    ws.Columns(countyColNum).Insert
    ws.Cells(1, countyColNum) = "County"
    countyColNum = FieldColNum(ws, "County")
End If

If stateFieldName = "" And singleStateName = "" Then
    stateColNum = FieldColNum(ws, "state", , True)
ElseIf stateFieldName <> "" And singleStateName = "" Then
    stateColNum = FieldColNum(ws, stateFieldName, , True)
ElseIf stateFieldName = "" And singleStateName <> "" Then
    state = singleStateName
    stateColNum = 0
End If
'
'If listStates = True Then
'    If FieldExists(ws, "State") Then
'        Debug.Print "State Field Already Exists. Please Delete before continuing."
'        Exit Sub
'    End If
'    ws.Columns(countyColNum + 1).Insert
'    ws.Cells(1, countyColNum + 1) = "State"
'End If
stateColNum = -1
If stateColNum = -1 Then '//If no state field was found or listed, the program assumes that we are to search all states to match counties

    If Not (FieldExists(ws, "State")) Then
        ws.Cells(1, countyColNum + 1) = "State"
        ws.Columns(countyColNum + 1).Insert
    End If
    
    listStates = True
    Set wbCounty = Workbooks.Open(zipToCountyMappingPath, , True)
    Set wsState = wbCounty.Worksheets.Add
    wsState.Move , wbCounty.Worksheets(wbCounty.Worksheets.Count)
    wsState.Cells(1, 1) = "Zip Code"
    wsState.Cells(1, 2) = "City Name"
    wsState.Cells(1, 3) = "State"
    wsState.Cells(1, 4) = "Primary County Name"
    
    For Each wss In wbCounty.Worksheets
        If wss.Name = wsState.Name Then Exit For
        wsState.Range(wsState.Cells(lastRow(wsState) + 1, 1), wsState.Cells(lastRow(wsState) - 1 + lastRow(wss), 4)) = wss.Range(wss.Cells(2, 1), wss.Cells(lastRow(wss), 4)).value
    Next wss
    
    Set wsCounty = wsState
    
    VbaZipLookup ws, "County", zipcodeFieldName, wsCounty, "ZIP Code", "Primary County Name", listStates
    
    '// Check for unmatched items
    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        If rngCounty.value = "" Then rngCounty.value = "0 Not Found"
    Next i
    
    Debug.Print ("Used All State Information because State Column could not originally be located.")
    wbCounty.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox (FirstName() & ", the counties have been added to your data." & vbCrLf & vbcrl & "-Report Robot")
    Exit Sub
    
End If

'// If a valid state field name was provided, then the program limits the scope to only states that are listed
Set wbCounty = Workbooks.Open(zipToCountyMappingPath, , True)
ws.Activate

If singleStateName = "" Then

    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))

        If stateColNum > 0 Then state = ws.Cells(i, stateColNum)
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        On Error GoTo errhandler
        Set wsCounty = wbCounty.Worksheets(state)
        rngCounty.value = wsCounty.Cells(Application.WorksheetFunction.Match(zipCode, wsCounty.Columns(FieldColNum(wsCounty, "ZIP Code")), 0), FieldColNum(wsCounty, "Primary County Name"))
        On Error GoTo 0
NextZipCode:
    Next i
Else
    Set wsCounty = wbCounty.Worksheets(singleStateName)
    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        On Error GoTo ErrHandler2
        rngCounty.value = wsCounty.Cells(Application.WorksheetFunction.Match(zipCode, wsCounty.Columns(FieldColNum(wsCounty, "ZIP Code")), 0), FieldColNum(wsCounty, "Primary County Name"))
        On Error GoTo 0
NextZipCode2:
    Next i
    
End If
wbCounty.Close
Exit Sub

errhandler:
Err.Clear
rngCounty.value = "0 Not Found"
Resume NextZipCode

ErrHandler2:
Err.Clear
rngCounty.value = "0 Not Found"
Resume NextZipCode2

End Sub

Public Sub zipcode_county_lookup(ByVal ws As Worksheet, ByVal zipcodeFieldName As String, Optional ByVal stateFieldName As String, Optional ByVal singleStateName As String, Optional ByVal listStates As Boolean)
'// Matches zip codes with county names in worksheets with multiple states listed
'// Authored by Nathan N on 6/11/2012
'// Updated 6/14/2012 - compiles all state zip code information if no state is listed

Dim wbCounty As Workbook, wsCounty As Worksheet, wsState As Worksheet
Dim wb As Workbook
Dim state As String, zipCode As Long
Dim rngCounty As Range
Dim colLet As String
Dim zipColNum As Integer, countyColNum As Integer, stateColNum As Integer
Dim zipToCountyMappingPath As String

Dim i As Long

Application.DisplayAlerts = False

zipToCountyMappingPath = file_reference_path() & "\ZipCodeAndCountyMappingDataPub.xlsx"

zipColNum = FieldColNum(ws, zipcodeFieldName, , True)
countyColNum = zipColNum + 1

If zipColNum = -1 Then
    Debug.Print ("Error: zipcode_county_lookup function. Zip Code Column could not be located. Ended function.")
    Exit Sub
End If

ws.Columns(countyColNum).Insert
ws.Cells(1, countyColNum) = "County"

If stateFieldName = "" And singleStateName = "" Then
    stateColNum = FieldColNum(ws, "state", , True)
ElseIf stateFieldName <> "" And singleStateName = "" Then
    stateColNum = FieldColNum(ws, stateFieldName, , True)
ElseIf stateFieldName = "" And singleStateName <> "" Then
    state = singleStateName
    stateColNum = 0
End If
'
'If listStates = True Then
'    If FieldExists(ws, "State") Then
'        Debug.Print "State Field Already Exists. Please Delete before continuing."
'        Exit Sub
'    End If
'    ws.Columns(countyColNum + 1).Insert
'    ws.Cells(1, countyColNum + 1) = "State"
'End If

If stateColNum = -1 Then '//If no state field was found or listed, the program assumes that we are to search all states to match counties
    ws.Columns(countyColNum + 1).Insert
    ws.Cells(1, countyColNum + 1) = "State"
    listStates = True
    Set wbCounty = Workbooks.Open(zipToCountyMappingPath, , True)
    Set wsState = wbCounty.Worksheets.Add
    wsState.Move , wbCounty.Worksheets(wbCounty.Worksheets.Count)
    wsState.Cells(1, 1) = "Zip Code"
    wsState.Cells(1, 2) = "City Name"
    wsState.Cells(1, 3) = "State"
    wsState.Cells(1, 4) = "Primary County Name"
    
    For Each wss In wbCounty.Worksheets
        If wss.Name = wsState.Name Then Exit For
        wsState.Range(wsState.Cells(lastRow(wsState) + 1, 1), wsState.Cells(lastRow(wsState) - 1 + lastRow(wss), 4)) = wss.Range(wss.Cells(2, 1), wss.Cells(lastRow(wss), 4)).value
    Next wss
    
    Set wsCounty = wsState
    
    VbaZipLookup ws, "County", zipcodeFieldName, wsCounty, "ZIP Code", "Primary County Name", listStates
    
    '// Check for unmatched items
    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        If rngCounty.value = "" Then rngCounty.value = zipCode & " Not Found"
    Next i
    
    Debug.Print ("Used All State Information because State Column could not originally be located.")
    wbCounty.Close
    Exit Sub
    
End If

'// If a valid state field name was provided, then the program limits the scope to only states that are listed
Set wbCounty = Workbooks.Open(zipToCountyMappingPath, , True)
ws.Activate

If singleStateName = "" Then

    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))

        If stateColNum > 0 Then state = ws.Cells(i, stateColNum)
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        On Error GoTo errhandler
        Set wsCounty = wbCounty.Worksheets(state)
        rngCounty.value = wsCounty.Cells(Application.WorksheetFunction.Match(zipCode, wsCounty.Columns(FieldColNum(wsCounty, "ZIP Code")), 0), FieldColNum(wsCounty, "Primary County Name"))
        On Error GoTo 0
NextZipCode:
    Next i
Else
    Set wsCounty = wbCounty.Worksheets(singleStateName)
    For i = 2 To lastRow(ws)
        Set rngCounty = ws.Cells(i, FieldColNum(ws, "County"))
        zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
        On Error GoTo ErrHandler2
        rngCounty.value = wsCounty.Cells(Application.WorksheetFunction.Match(zipCode, wsCounty.Columns(FieldColNum(wsCounty, "ZIP Code")), 0), FieldColNum(wsCounty, "Primary County Name"))
        On Error GoTo 0
NextZipCode2:
    Next i
    
End If
wbCounty.Close
Exit Sub

errhandler:
Err.Clear
rngCounty.value = "0 Not Found"
Resume NextZipCode

ErrHandler2:
Err.Clear
rngCounty.value = "0 Not Found"
Resume NextZipCode2

End Sub
Public Sub zipcode_state_lookup(ByVal ws As Worksheet, ByVal zipcodeFieldName As String)
'// Authored by Nathan N on 6/11/2012
'// Updated 6/14/2012 - compiles all state zip code information if no state is listed
'// Matches zip codes with county names in worksheets with multiple states listed
Dim wbCounty As Workbook, wsCounty As Worksheet, wsState As Worksheet
Dim wb As Workbook
Dim state As String, zipCode As Long
Dim rngCounty As Range
Dim colLet As String
Dim stateColNum As Integer, zipColNum As Integer
Dim i As Long

zipColNum = FieldColNum(ws, zipcodeFieldName, , True)
stateColNum = zipColNum + 1

If zipColNum = -1 Then
    Debug.Print ("Error: zipcode_county_lookup function. Zip Code Column could not be located. Ended function.")
    Exit Sub
End If

ws.Columns(stateColNum).Insert
ws.Cells(1, stateColNum) = "State"

Set wbCounty = Workbooks.Open(file_reference_path() & "\ZipCodeAndCountyMappingDataPub.xlsx", , True)
Set wsState = wbCounty.Worksheets.Add
wsState.Move , wbCounty.Worksheets(wbCounty.Worksheets.Count)
wsState.Cells(1, 1) = "Zip Code"
wsState.Cells(1, 2) = "City Name"
wsState.Cells(1, 3) = "State"
wsState.Cells(1, 4) = "Primary County Name"

For Each wss In wbCounty.Worksheets
    If wss.Name = wsState.Name Then Exit For
    wsState.Range(wsState.Cells(lastRow(wsState) + 1, 1), wsState.Cells(lastRow(wsState) - 1 + lastRow(wss), 4)) = wss.Range(wss.Cells(2, 1), wss.Cells(lastRow(wss), 4)).value
Next wss

Set wsCounty = wsState

vbaLookup ws, "State", zipcodeFieldName, wsCounty, "ZIP Code", "State"

'// Check for unmatched items
For i = 2 To lastRow(ws)
    Set rngCounty = ws.Cells(i, FieldColNum(ws, "State"))
    zipCode = ws.Cells(i, FieldColNum(ws, zipcodeFieldName))
    If rngCounty.value = "" Then rngCounty.value = zipCode & " Not Found"
Next i

wbCounty.Close

End Sub
Public Function sigPath() As String
'// Authored by Nathan N on 6/6/2012
'// Returns the email signature path for the user

If windows7os() Then
    sigPath = user_root(True) & "\AppData\Roaming\Microsoft\Signatures"
Else
    sigPath = user_root(True) & "\Application Data\Microsoft\Signatures"
End If

End Function
Public Function user_root(Optional ByVal excludeMyDocuments As Boolean) As String
Dim docPath As String
'* Authored by Nathan N on 6/4/2012
'* Modified 6/6/2012
'* Returns the Path to windows 7 user's "My Documents" folder
'* If excludeMyDocuments = true, the function returns the path only up to the UserName folder
If excludeMyDocuments = False Then docPath = "\My Documents"

If windows7os() = True Then
    If docPath <> "" Then
        docPath = "\" & Right(docPath, Len(docPath) - 4)
        user_root = "C:\Users\" & loginName() & docPath
    Else
        user_root = "C:\Users\" & loginName()
    End If
Else
    user_root = "C:\Documents and Settings\" & loginName() & docPath
End If

End Function
Public Function windows7os() As Boolean
'* Authored by Nathan N on 6/4/2012
'* Determines if user is using windows 7

If Left(Application.OperatingSystem, 21) = "Windows (32-bit) NT 5" Then
    windows7os = False
Else
    windows7os = True
End If

End Function
Public Function hub_path() As String
'// Created by Nathan N on 7/9/2012
hub_path = "\\hqclienthub\Client Hub"
End Function
Public Function file_reference_path() As String
'// Created by Nathan N on 7/9/2012
'// Helper function used avoid using long absolute file paths for common file resources
'// You must specify the path within the quotations below.
file_reference_path = ""
End Function

Public Sub delete_empty_entries(ByVal ws As Worksheet, ByVal fieldName As String)
'Authored by Nathan N on 4/9/2012

Dim i As Long

For i = 2 To lastRow(ws)
    
    If (ws.Cells(i, FieldColNum(ws, fieldName)) = "") Then
        ws.Rows(i).Delete Shift:=xlUp
        i = i - 1
        If (i = lastRow(ws)) Then
            Exit For
        End If
    End If
    
Next i

End Sub
Public Sub AggregateServices(ByVal ws As Worksheet, ByVal WOFieldName As String, ByVal serviceFieldName As String, ByVal newFieldName As String, Optional ByVal sortBeforehand As Boolean)
'*** Created on 3/20/2012 by Nathan N ***
'*** Combines services on seperate row entries into one line for the same work order number***
'*** User needs to remove duplicate work orders seperately ***

Dim currentWO As String
Dim aggServiceType As String, newServiceType As String
Dim WOCol As Integer, currentServiceCol As Integer, newServiceCol As Integer
Dim i As Long, j As Long

WOCol = FieldColNum(ws, WOFieldName)
currentServiceCol = FieldColNum(ws, serviceFieldName)
newServiceCol = lastColNum(ws) + 1
ws.Cells(1, newServiceCol) = newFieldName

If sortBeforehand = True Then SortMultiCol ws, ws.Cells(1, WOCol), ws.Cells(1, currentServiceCol)

For i = 2 To lastRow(ws)

    currentWO = ws.Cells(i, WOCol)
    aggServiceType = ""
    
    For j = i To lastRow(ws)

        If currentWO = ws.Cells(j, WOCol) Then
            newServiceType = ws.Cells(j, currentServiceCol)
            aggServiceType = aggServiceType & ", " & newServiceType
        End If
        
        If currentWO <> ws.Cells(j + 1, WOCol) Then
            ws.Cells(i, newServiceCol) = aggServiceType
            If j >= lastRow(ws) Then
                ws.Cells(i, newServiceCol) = aggServiceType
                '// Last Row error correction
                ws.Cells(lastRow(ws), newServiceCol) = Right(ws.Cells(lastRow(ws), newServiceCol), Len(ws.Cells(lastRow(ws), newServiceCol)) - 2)
                Exit Sub
            End If
            If (Len(ws.Cells(i, newServiceCol)) > 0) Then
                ws.Cells(i, newServiceCol) = Right(ws.Cells(i, newServiceCol), Len(aggServiceType) - 2)
            End If
            i = j
            GoTo nextWO
        End If
        
    Next j

nextWO:

Next i

End Sub
Public Function InArray(ByRef arrayName() As Variant, ByVal elementValue As Variant) As Boolean
'Created on 3/13/2012 by Nathan N

Dim i As Long

For i = 0 To UBound(arrayName())
    If (elementValue = arrayName(i)) Then 'If the elementValue is found at any point in the array, then TRUE
        InArray = True
        Exit Function
    End If
Next i


InArray = False 'Otherwise, InArray is false

End Function

Public Function DateStamp(Optional AMPM As Boolean, Optional dateOnly As Boolean, Optional AMPMDateStamp As Boolean, Optional AMPMTimeStamp As Boolean) As String
'***************************************************
'**** created by Nathan N on 1/25/2012 ****
'Updated on 3/9/2012 to correct for the extra space in the AMPMTimeStamp

Dim timeStamp As String, timeStampAMPM As String, AMPMAndDateStamp As String

If dateOnly = True Then
    DateStamp = Date
    DateStamp = Replace(DateStamp, "/", " ")
    Exit Function
End If

If AMPM = True Then
    timeStamp = Time()
    timeStamp = Replace(timeStamp, ":", " ")
    timeStamp = Left(Trim(timeStamp), 2)
    timeStampAMPM = Time()
    timeStampAMPM = Right(Trim(Now()), 2)
    DateStamp = timeStampAMPM
    Exit Function
End If

If AMPMTimeStamp = True Then
    DateStamp = Date
    DateStamp = Replace(DateStamp, "/", " ")
    timeStamp = Time()
    timeStamp = Replace(timeStamp, ":", " ")
    timeStamp = Replace(Left(Trim(timeStamp), 2), " ", "")
    timeStampAMPM = Time()
    timeStampAMPM = Right(Trim(Now()), 2)
    AMPMAndDateStamp = timeStampAMPM & " - " & DateStamp
    timeAndDateStamp = timeStamp & " " & timeStampAMPM & " - " & DateStamp
    DateStamp = timeAndDateStamp
    Exit Function
End If

If AMPMDateStamp = True Then
    DateStamp = Date
    DateStamp = Replace(DateStamp, "/", " ")
    timeStamp = Time()
    timeStamp = Replace(timeStamp, ":", " ")
    timeStamp = Left(Trim(timeStamp), 2)
    timeStampAMPM = Time()
    timeStampAMPM = Right(Trim(Now()), 2)
    AMPMAndDateStamp = timeStampAMPM & " - " & DateStamp
    DateStamp = AMPMAndDateStamp
    Exit Function
End If

End Function
Public Sub CleanServiceName(ByVal ws As Worksheet, ByVal serviceColNum As Integer)
'*** 3/6/2012 ***
Dim i As Long
Dim serviceName As String

For i = 2 To lastRow(ws)

    serviceName = LCase(ws.Cells(i, serviceColNum))
    
    If (serviceName Like "*recut*") Or (serviceName Like "*lawn*") Then
        ws.Cells(i, serviceColNum) = "Ongoing Recut"
    ElseIf serviceName Like "*maid*" Then
        ws.Cells(i, serviceColNum) = "Ongoing Maid"
    ElseIf serviceName Like "*pool*" Then
        ws.Cells(i, serviceColNum) = "Ongoing Pool"
    ElseIf serviceName Like "*snow*" Then
        ws.Cells(i, serviceColNum) = "Ongoing Snow"
    End If
    
Next i

End Sub
Public Sub CopyMultiValRows(sourceSheet As Worksheet, colLet As String, destSheet As Worksheet, copyValues As String, Optional copyHeader As Boolean)
Dim inverseRow As Long
Dim lastRowCurrent As Long
Dim lastRowDest As Long
Dim i As Long


If copyHeader <> False Then
    copyHeader = True
End If

lastRowCurrent = sourceSheet.Range("A" & Rows.Count).End(xlUp).Row
origRowCount = sourceSheet.Range("A" & Rows.Count).End(xlUp).Row

If copyHeader = True Then
    destSheet.Rows(1) = sourceSheet.Rows(1).value
    lastRowDest = destSheet.Range("A" & Rows.Count).End(xlUp).Row
End If

For i = 2 To lastRowCurrent
    inverseRowNum = origRowCount - i + 2
    currentCellValue = sourceSheet.Range(colLet & inverseRowNum).value

    If InStr(1, copyValues, currentCellValue) >= 1 Then
        destSheet.Rows(lastRowDest + 1) = sourceSheet.Rows(inverseRowNum).value
        lastRowDest = destSheet.Range("A" & Rows.Count).End(xlUp).Row
    End If
    
Next i

End Sub

Sub CutMultiValRows(ByVal sourceSheet As Worksheet, ByVal colLet As String, ByVal destSheet As Worksheet, ByVal cutValues As String, Optional copyHeader As Boolean)
Dim inverseRow As Long
Dim lastRow As Long
Dim lastRowDest As Long
Dim i As Long

If copyHeader = "" Then
    copyHeader = True
End If

lastRow = sourceSheet.Range(colLet & Rows.Count).End(xlUp).Row
origRowCount = sourceSheet.Range(colLet & Rows.Count).End(xlUp).Row

If copyHeader = True Then
    destSheet.Rows(1) = sourceSheet.Rows(1).value
    lastRowDest = destSheet.Range("A" & Rows.Count).End(xlUp).Row
End If

For i = 2 To lastRow
    inverseRowNum = origRowCount - i + 2
    currentCellValue = sourceSheet.Range(colLet & inverseRowNum).value

    If InStr(1, cutValues, currentCellValue) >= 1 Then
        destSheet.Rows(lastRowDest + 1) = sourceSheet.Rows(inverseRowNum).value
        lastRowDest = destSheet.Range("A" & Rows.Count).End(xlUp).Row
        sourceSheet.Rows(inverseRowNum).Delete Shift:=xlUp
    End If
    
Next i

End Sub
Public Sub SortCol(ByVal ws As Worksheet, ByVal fieldName As String, Optional ByVal headerRowNum As Integer)
'*********************************************************
'*** Authored by Nathan N, on 1/18/2012
'***
'*** Updateed on 2/1/2012 at 11:19AM - Switched from function to sub
'*** Updated on 2/3/2012 at 1:00PM - Switched out sortColNum parameter with fieldname:string
'***                                                         - Removed the application.visible property and implemented a Wait method

'*********************************************************
Dim sortColNum As Integer

sortColNum = FieldColNum(ws, fieldName)

If headerRowNum = 0 Then
    headerRowNum = 1
End If

ws.Activate
ws.Sort.SortFields.Clear
ws.Sort.SortFields.Add ws.Cells(headerRowNum, sortColNum)
ws.Sort.Header = xlNo
ws.Sort.MatchCase = False
ws.Sort.SetRange ws.Range(Cells(headerRowNum + 1, 1), Cells(lastRow(ws), lastColNum(ws)))
'Application.Visible = True
Application.Wait Now + 0.000008

ws.Sort.Apply
'Application.Visible = False

        
End Sub

Public Sub SortMultiCol(ByVal ws As Worksheet, ByVal fieldName1 As String, Optional ByVal fieldName2 As String, Optional ByVal fieldName3 As String, Optional ByVal headerRowNum As Integer)
'*********************************************************
'*** Authored by Nathan N, on 2/27/2012
'*** Updated 2/28/2012 8:17AM   - Added functionality for 3 columns to be sorted
'*********************************************************
Dim sortColNum1 As Integer, sortColNum2 As Integer

sortColNum1 = FieldColNum(ws, fieldName1)

If (FieldExists(ws, fieldName2)) And (fieldName2 <> "") Then
    sortColNum2 = FieldColNum(ws, fieldName2)
End If
If FieldExists(ws, fieldName3) And (fieldName3 <> "") Then
    sortColNum3 = FieldColNum(ws, fieldName3)
End If

If headerRowNum = 0 Then headerRowNum = 1

ws.Activate
ws.Sort.SortFields.Clear
ws.Sort.SortFields.Add ws.Cells(headerRowNum, sortColNum1)

If FieldExists(ws, fieldName2) And (fieldName2 <> "") Then ws.Sort.SortFields.Add ws.Cells(headerRowNum, sortColNum2)

If FieldExists(ws, fieldName3) And (fieldName3 <> "") Then ws.Sort.SortFields.Add ws.Cells(headerRowNum, sortColNum3)

ws.Sort.Header = xlNo
ws.Sort.MatchCase = False
ws.Sort.SetRange ws.Range(Cells(headerRowNum + 1, 1), Cells(lastRow(ws), lastColNum(ws)))
'Application.Visible = True
Application.Wait Now + 0.000008 '// This is needed in order for the function to work successfully
ws.Sort.Apply
'Application.Visible = False

End Sub
Public Sub remove_dupes(ByVal ws As Worksheet, ByVal fieldName As String)
'// Created on 5/3/2012 by Nathan N

Dim colNum As Integer

colNum = FieldColNum(ws, fieldName)

ws.Range(ws.Cells(1, 1), ws.Cells(lastRow(ws), lastColNum(ws))).RemoveDuplicates Columns:=colNum, Header:=xlYes

End Sub
Public Function lastRow(ByVal ws As Worksheet, Optional ByVal colLet As String, Optional ByVal colNum As Integer, Optional ByVal startRow As Long) As Long
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'*** Updated 3/2/2012    - Added startRow functionality to find last rows of irregular data
'*** Updated 6/13/2012 - Added conditional block that took into account if a ColNum greater than 0 was used.
'*** Optional [colLet] parameter takes precedence over Optional [colNum]
'*** Detects the last row with data in a specified column letter
'*** PreCondition: 1. At least one worksheet row must have data
'***                        2. Optional: A specific column letter can be declared.
'***                            If no column letter is declared, function will default
'***                            to the "A" column
'*** PostCondition: Returns a number (Long type)
'*** Updated 3/9/2012 - included IF statement for colNum = 1 and startRow = 1
'*** Updated 4/3/2012 - Corrected bug that ended up adding one to the last row if the startRow>0
'*********************************************************

If colLet = "" And colNum = 0 Then colLet = "A"

If colLet <> "" And startRow = 0 Then
    lastRow = ws.Range(colLet & Rows.Count).End(xlUp).Row
    Exit Function
End If

If colLet = "" And colNum > 0 And startRow = 0 Then
    lastRow = ws.Cells(Rows.Count, colNum).End(xlUp).Row
    Exit Function
End If

If (colNum = 0) Then colNum = 1
If (startRow = 0) Then startRow = 1

If colNum = 1 And startRow = 1 Then
    lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    Exit Function
End If

If startRow > 0 Then

    For i = startRow To 500000
    
        If ws.Cells(i, colNum) = "" Then
            lastRow = i - 1
            Exit Function
        End If
        
    Next i
    
End If

End Function
Public Function lastColNum(ByVal ws As Worksheet, Optional ByVal rowNum As Long = 1, Optional ByVal ignoreBreaks As Boolean = False) As Integer
'*********************************************************
'*** Authored byNathan N, on 1/13/2012
'*** Updated 6/13/2012 - Added Optional [rowNum] parameter to detect lastCol of a specific row
'*** This detects the last column number
'*** PreCondition: At least one column needs to be used on a worksheet
'*** PostCondition: Returns an integer of the last most column number
'*** that has data
'// Updated 7/12/2012 - Revised to use a loop as the primary last row detection mechanism instead of lastUsedCol built in function (was causing too many problems)
'// Updated 8/1/2012 - included optional ignoreBreaks parameter
'*********************************************************
Dim i As Integer

If ignoreBreaks Then
    lastColNum = ws.UsedRange.Columns.Count
    Exit Function
End If

For i = 1 To ws.UsedRange.Columns.Count + 1
    If ws.Cells(rowNum, i) = "" Then
        lastColNum = i - 1
        Exit For
    End If
Next i

If lastColNum = 0 Then lastColNum = ws.UsedRange.Columns.Count

End Function
Public Function LastColtoLet(ByVal currentSheet As Worksheet) As String
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** This returns the letter of an automatically detected last column number
'*** PreCondition: 1. A specific worksheet name must be declared
'***                         2. At least one column must have data within the column
'***                             range of A to Z
'***
'*** PostCondition: Returns the letter (String Type) that was associated with
'***                           the detected last column
'*********************************************************
Dim letArray() As Variant
Dim maxArrayCount As Long
Dim lastCol As Integer
ReDim letArray(0 To 100)

'--- More letters can be added to the letArray(0 to 25) to extend functionality, however
'--- the letArray must be redimensioned with more elements to match the
'--- quantity of letters listed. For instance, if "AA", "AB","AC" were added to letArray,
'--- we must redimension the letArray with the following: ReDim letArray(0 to 28)

letArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF")

maxArrayCount = (Application.WorksheetFunction.CountA(letArray))
ReDim Preserve letArray(0 To maxArrayCount - 1)

lastCol = currentSheet.UsedRange.Columns.Count

LastColtoLet = letArray(lastCol - 1)

End Function
Public Function LetToColNum(ByVal s As String) As Integer
'*********************************************************
'*** Authored by Nathan N, on 3/20/2012
'***
'*** This returns the Col Num of a specified Col Let, up to BF

'*********************************************************
Dim letArray() As Variant
Dim maxArrayCount As Long
ReDim letArray(0 To 100)

letArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF")
maxArrayCount = (Application.WorksheetFunction.CountA(letArray))
ReDim Preserve letArray(0 To maxArrayCount - 1)

For i = 0 To maxArrayCount
    If UCase(s) = letArray(i) Then
        LetToColNum = i + 1
        Exit Function
    End If
Next i

End Function
Public Function ColNumToLet(ByVal currentSheet As Worksheet, ByVal colNum As Integer) As String
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** This returns the letter of an automatically detected last column number
'*** PreCondition: 1. A specific worksheet name must be declared
'***                         2. At least one column must have data within the column
'***                             range of A to Z
'***
'*** PostCondition: Returns the letter (String Type) that was associated with
'***                           the detected last column
'*********************************************************
Dim letArray() As Variant
Dim maxArrayCount As Long

ReDim letArray(0 To 100)
'--- More letters can be added to the letArray(0 to 25) to extend functionality, however
'--- the letArray must be redimensioned with more elements to match the
'--- quantity of letters listed. For instance, if "AA", "AB","AC" were added to letArray,
'--- we must redimension the letArray with the following: ReDim letArray(0 to 28)

letArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF")
maxArrayCount = (Application.WorksheetFunction.CountA(letArray))
ReDim Preserve letArray(0 To maxArrayCount - 1)

'colNum = currentSheet.UsedRange.Columns.Count

ColNumToLet = letArray(colNum - 1)

End Function
Public Function FieldColLet(ByVal currentSheet As Worksheet, ByVal fieldName As String, Optional ByVal rowNum As Integer) As String
'*********************************************************
'*** Authored by Nathan N, on 1/18/2012
'***
'*** This returns the column letter of a specified field name
'*** PreCondition: 1. A specific worksheet name must be declared
'***                         2. At least one column must have data within the column
'***                             range of A to AZ
'***                         3. if rowNum is not specified, then default is 1
'***
'*** PostCondition: Returns the col letter (String Type) that was associated with
'***                           the fieldName.
'***
'*********************************************************
Dim letArray() As Variant
Dim FieldColNum As Integer
Dim maxArrayCount As Long

ReDim letArray(0 To 100)
'--- More letters can be added to the letArray(0 to 25) to extend functionality, however
'--- the letArray must be redimensioned with more elements to match the
'--- quantity of letters listed. For instance, if "AA", "AB","AC" were added to letArray,
'--- we must redimension the letArray with the following: ReDim letArray(0 to 28)
If rowNum = 0 Then rowNum = 1

letArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH")
maxArrayCount = (Application.WorksheetFunction.CountA(letArray))
ReDim Preserve letArray(0 To maxArrayCount - 1)

FieldColNum = Application.WorksheetFunction.Match(fieldName, currentSheet.Rows(rowNum), 0)
FieldColLet = letArray(FieldColNum - 1)

End Function
Public Function FieldColNum(ByVal ws As Worksheet, ByVal fieldName As String, Optional ByVal rowNum As Long, Optional approxMatch As Boolean) As Integer
'*********************************************************
'*** Authored by Nathan N, on 1/18/2012
'***Updated on 1/31/2012 - Added For i loop that addresses field names with spaces
'***Updated on 3/5/2012 - Changed rowNum type to Long from Ingeter
'***Updated on 6/12/2012 - Added optional approxMatch boolean and functionality
'***                                        - Returns -1 upon failure
'***
'*** This returns the column letter of a specified field name
'*** PreCondition: 1. A specific worksheet name must be declared
'***                         2. A valid fieldName must be declared and located on the
'***                            specified rowNum. If no rowNum is specified, then default
'***                            is 1.
'*** PostCondition: Returns a column number
'***
'***
'*********************************************************
Dim i As Long

If rowNum = 0 Then rowNum = 1

If approxMatch = True Then

    For i = 1 To lastColNum(ws)
        If LCase(ws.Cells(rowNum, i)) Like "*" & LCase(fieldName) & "*" Then
            FieldColNum = i
            Exit Function
        End If
    Next i
    
End If

If InStr(1, " ", fieldName) Then 'This takes into account the two different types of data spreadsheet formats (One that has spaces inbetween field names, and one that has no spaces)

    For i = 1 To lastColNum(ws)
        If ws.Cells(rowNum, i) = fieldName Then
            FieldColNum = i
            Exit Function
        End If
    Next i

Else
    On Error Resume Next
    FieldColNum = Application.WorksheetFunction.Match(fieldName, ws.Rows(rowNum), 0)
    If (FieldColNum = 0) Then FieldColNum = Application.WorksheetFunction.Match(CInt(fieldName), ws.Rows(rowNum), 0)
    If (FieldColNum = 0) Then FieldColNum = Application.WorksheetFunction.Match(CDec(fieldName), ws.Rows(rowNum), 0)
    If (FieldColNum = 0) Then '// If the function fails, return -1
        FieldColNum = -1
        Exit Function
    End If
    On Error GoTo 0
End If

End Function
Public Function outline_cells(ByVal rng As Range)
'Authored by Nathan N on 4/6/2012

    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Function
Public Function FieldRowNum(ByVal ws As Worksheet, ByVal fieldName As String, Optional ByVal colNum As Integer, Optional ByVal colLet As String, Optional startRowNum As Long, Optional ByVal fromTop As Boolean, Optional ByVal fromBottom As Boolean, Optional exactMatch As Boolean) As Long

'*** Authored by Nathan N on 2/1/2012 ***
'*** Finds the row number for the first found instance of the specified fieldName:string, from the top of the data or the bottom
'*** Updated 2/21/2012 at 1:30PM - Added Exact Match parameter
'*** Updated 2/27/2012 at 3:49PM    - Fixed the "From Bottom" functionality
'*** Updated 3/2/2012 at 1:35PM       - Defaulted to fromTop = true if neither fromTop or fromBottom were selected
'*** Updated 3/15/2012 at 1:44PM      - Corrected issue whereby the search would start from the top, even if specified to start from the bottom
'                                                                       when the colNum was 0 or empty
Dim fieldFound As Boolean
Dim rowNum As Long
Dim i As Long, inverseRowNum As Long

If (fromTop = False And fromBottom = False) Then
    fromTop = True
End If

If (exactMatch = False) Then
    If (fromTop = True) Or (fromTop = False And fromBottom = False) Then
        If (colNum <> 0) Then
            If startRowNum = 0 Then
                startRowNum = 1
            End If
            
            For i = startRowNum To lastRow(ws, , colNum)
            
                If (InStr(1, LCase(ws.Cells(i, colNum)), LCase(fieldName)) > 0) Then
                    FieldRowNum = i
                    'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & i & ", column number " & colNum & "."
                    Exit Function
                End If
                
            Next i
            
            'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
            Exit Function
        End If
        
        If colNum = 0 Then
            If startRowNum = 0 Then
                startRowNum = 1
            End If
            For i = startRowNum To lastRow(ws)
            
                If (InStr(1, LCase(ws.Cells(i, 1)), LCase(fieldName)) > 0) Then
                    FieldRowNum = i
                    'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & i & ", column number " & colNum & "."
                    Exit Function
                End If
                
            Next i
            'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & rowNum & ", column number " & colNum & "."
            Exit Function
        End If
        
        'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
        Exit Function
    End If
    
If (fromBottom = True) Then

    If (colNum = 0) Then colNum = 1
    
    If startRowNum = 0 Then startRowNum = lastRow(ws)
        
    origRowCount = lastRow(ws)
    
    For i = 1 To startRowNum
        inverseRowNum = startRowNum - i + 1
        
        If (InStr(1, LCase(ws.Cells(inverseRowNum, colNum)), LCase(fieldName)) > 0) Then 'This currently only supports approx match
            FieldRowNum = inverseRowNum
            'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & inverseRowNum & ", column number " & colNum & "."
            Exit Function
        End If
        
    Next i
    
    'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
    Exit Function
    
    If colNum = 0 Then
        On Error Resume Next
        colLet = "A"
        rowNum = Application.WorksheetFunction.Match(fieldName, ws.Columns(colLet & ":" & colLet), 0)
        fieldFound = True
        FieldRowNum = rowNum
        On Error GoTo 0
        'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & rowNum & ", column number " & colNum & "."
        Exit Function
    End If
    
   ' Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
    Exit Function
End If
    
Else

    If (fromTop = True) Or (fromTop = False And fromBottom = False) Then
        If (colNum <> 0) Then
            If startRowNum = 0 Then
                startRowNum = 1
            End If
            
            For i = startRowNum To lastRow(ws)
            
                If (ws.Cells(i, colNum) = fieldName) Then
                    FieldRowNum = i
                    'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & i & ", column number " & colNum & "."
                    Exit Function
                End If
                
            Next i
            
            'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
            Exit Function
        End If
        
        If colNum = 0 Then
            On Error Resume Next
            colLet = "A"
            rowNum = Application.WorksheetFunction.Match(fieldName, ws.Columns(colLet & ":" & colLet), 0)
            fieldFound = True
            FieldRowNum = rowNum
            On Error GoTo 0
            'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & rowNum & ", column number " & colNum & "."
            Exit Function
        End If
        
        'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
        Exit Function
    End If
    
    If (fromBottom = True) Then
        If (colNum <> 0) Then
            If startRowNum = 0 Then
                startRowNum = lastRow(ws)
            End If
            
            origRowCount = lastRow(ws)
            
            For i = 1 To startRowNum
                inverseRowNum = startRowNum - i + 1
                
                If (ws.Cells(inverseRowNum, colNum) = fieldName) Then
                    FieldRowNum = inverseRowNum
                    'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & inverseRowNum & ", column number " & colNum & "."
                    Exit Function
                End If
            
        Next i
            
            'Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
            Exit Function
        End If
    
        If colNum = 0 Then
            On Error Resume Next
            colLet = "A"
            rowNum = Application.WorksheetFunction.Match(fieldName, ws.Columns(colLet & ":" & colLet), 0)
            fieldFound = True
            FieldRowNum = rowNum
            On Error GoTo 0
            'Debug.Print "SUCCESS: FieldRowNum function found the string """ & fieldName & """ in row " & rowNum & ", column number " & colNum & "."
            Exit Function
        End If
        
       ' Debug.Print "WARNING: FieldRowNum function did not find the string """ & fieldName & """ in column number " & colNum & "."
        Exit Function
    End If
End If

End Function
'Sub testt()
'Dim wsDest As Worksheet, wsSource As Worksheet
'
'Set wsDest = Workbooks("OneWestNonRecurringAlerts.csv").Worksheets("OneWestNonRecurringAlerts")
'Set wsSource = Workbooks("ClientReferenceList.xlsx").Worksheets("ClientUmbrella")
'
'vbaLookup2 wsDest, "ClientNameCondensed", "ClientName", wsSource, "ClientName", "UmbrellaName"
'
'End Sub
Public Sub vbaLookup2(ByVal wsDest As Worksheet, ByVal destField As String, ByVal destKeyField As String, _
    ByVal wsSource As Worksheet, ByVal sourceKeyField As String, Optional ByVal sourceValueField As String, _
    Optional insertNewColumn As Boolean, Optional keyFoundValue As String, Optional keyNotFoundValue As String, _
    Optional ByVal destStartRow As Long, Optional ByVal sourceStartRow As Long, Optional ByVal createNewLastColumn As Boolean)
'// Authored by Nathan N on 7/2/2012
'// Updated 1/16/2015
'// Improved version of the vbaLookup. Uses the a dictionary object (built into other languages) that fills with key-value pairs.
 '  --Measured exponential speed savings as data set gets larger (very large data sets took longer than 10 minutes with the old vbaLookup, but with vbaLookup2 it took 40 seconds)
 '  --Option for new column to be inserted into the destination workbook which takes on the destField as the column title
 '  --Can return optional custom value (instead of matched value)blank or optional value if Key is not found in dictionary
 '      If keyFoundValue is used, then sourceValueField is not required, and in fact will be ignored.
 
 '// Future Improvements: Create multiple dictionaries, or 'buckets,' with the quantity of buckets increasing with proportion to the most
 '// common first few characters
 
Dim destKeyCol As Integer, destFieldCol As Integer, sourceKeyCol As Integer, sourceValueCol As Integer
Dim rngDest As Range
Dim i As Long

'// Variables for the Dictionary Object
Dim d As Object
Dim key As String, value As String, keyToMatch As String

'// Set option below (vbTextCompare ignores upper/lowercase, while vbBinaryCompare is faster but requires exact matches
Set d = CreateObject("Scripting.Dictionary")
d.CompareMode = vbTextCompare
'd.CompareMode = vbBinaryCompare
'd.CompareMode = vbDatabaseCompare
'd.CompareMode = vbUseCompareOption (THIS IS WHATEVER THE IDE HAS BEEN DEFAULTED TO. DO A GOOGLE SEARCH TO FIND OUT HOW TO SET THIS)

destKeyCol = FieldColNum(wsDest, destKeyField)

'// Insert Optional New Column
If insertNewColumn = True Then
    If destStartRow > 0 Then
        If FieldExists(wsDest, destField, destStartRow, True) Then Debug.Print "vbaLookup2 Warning: Destination Field Name Already Exists"
    Else
        If FieldExists(wsDest, destField, , True) Then Debug.Print "vbaLookup2 Warning: Destination Field Name Already Exists"
    End If
    destFieldCol = destKeyCol + 1
    wsDest.Columns(destFieldCol).Insert Shift:=xlRight
    wsDest.Cells(1, destFieldCol) = destField
ElseIf createNewLastColumn = True Then
    destFieldCol = lastColNum(wsDest) + 1
    wsDest.Cells(1, destFieldCol) = destField
Else
    If FieldExists(wsDest, destField) Then
        destFieldCol = FieldColNum(wsDest, destField)
    Else
        Debug.Print "vbaLookUp2 Error: destField = """ & destField & """ does not exist"
        Exit Sub
    End If
End If
sourceKeyCol = FieldColNum(wsSource, sourceKeyField)
sourceValueCol = FieldColNum(wsSource, sourceValueField)

'// Create Dictionary with key-value pairs from the source worksheet
If keyFoundValue = "" Then
    For i = 2 To lastRow(wsSource)
        key = wsSource.Cells(i, sourceKeyCol)
        value = wsSource.Cells(i, sourceValueCol)
        If Not d.Exists(key) Then d.Add key, value
    Next i
Else
    For i = 2 To lastRow(wsSource)
        key = wsSource.Cells(i, sourceKeyCol)
        value = keyFoundValue
        If Not d.Exists(key) Then d.Add key, value
    Next i
End If
'// Match destKey (keyToMatch) with corresponding values
If keyNotFoundValue = "" Then
    '// Algorithm for the default blank value for keys that were not found in the dictionary
    For i = 2 To lastRow(wsDest, , destKeyCol)
        Set rngDest = wsDest.Cells(i, destFieldCol)
        keyToMatch = wsDest.Cells(i, destKeyCol)
        rngDest.value = d(keyToMatch)
    Next i
ElseIf keyFoundValue = "" And keyNotFoundValue <> "" Then
    '// Algorithm for user specified value for keys that were not found in the dictionary
    For i = 2 To lastRow(wsDest, , destKeyCol)
        Set rngDest = wsDest.Cells(i, destFieldCol)
        keyToMatch = wsDest.Cells(i, destKeyCol)
        If d.Exists(keyToMatch) Then
            rngDest.value = d(keyToMatch)
        Else
            rngDest.value = keyNotFoundValue
        End If
    Next i
ElseIf keyFoundValue <> "" And keyNotFoundValue <> "" Then
    For i = 2 To lastRow(wsDest, , destKeyCol)
        Set rngDest = wsDest.Cells(i, destFieldCol)
        keyToMatch = wsDest.Cells(i, destKeyCol)
        If d.Exists(keyToMatch) Then
            rngDest.value = keyFoundValue
        Else
            rngDest.value = keyNotFoundValue
        End If
    Next i
End If

Set d = Nothing

End Sub

Public Sub vbaLookup(ByVal currentSheet As Worksheet, ByVal destField As String, ByVal matchField As String, ByVal sourceSheet As Worksheet, ByVal sourceMatchField As String, ByVal sourceMatchDataField As String, Optional ByVal destMatchFieldRowNum As Long, Optional ByVal sourceFieldRowNum As Long)
'**** 1/31/2012 ****
'Updated 3/1/2012 at 5:24PM     - Added Error notifications for unknown clients to be added to the client list
'Updated 3/2/2012 at 1:16PM    - Added destFieldRowNum Option and sourceFieldRowNum Option
'                                                       - Corrected error nofication to send only on a client error trigger
Dim destFieldColNum  As Integer, matchFieldColNum  As Integer, sourceMatchFieldColNum  As Integer, sourceMatchDataFieldColNum As Integer
Dim i As Long

Application.ScreenUpdating = False

If (destMatchFieldRowNum = 0) Then
    destMatchFieldRowNum = 1
End If
If (sourceFieldRowNum = 0) Then
    sourceFieldRowNum = 1
End If

destFieldColNum = Application.WorksheetFunction.Match(destField, currentSheet.Rows(destMatchFieldRowNum), 0)
matchFieldColNum = Application.WorksheetFunction.Match(matchField, currentSheet.Rows(destMatchFieldRowNum), 0)
sourceMatchFieldColNum = Application.WorksheetFunction.Match(sourceMatchField, sourceSheet.Rows(sourceFieldRowNum), 0)
sourceMatchDataFieldColNum = Application.WorksheetFunction.Match(sourceMatchDataField, sourceSheet.Rows(sourceFieldRowNum), 0)

For i = destMatchFieldRowNum + 1 To lastRow(currentSheet, , , destMatchFieldRowNum)
    'Debug.Print lastRow(currentSheet, , , destMatchFieldRowNum)
    On Error GoTo ErrHandlerClient
    sourceMatchRow = Application.WorksheetFunction.Match(currentSheet.Cells(i, matchFieldColNum).value, sourceSheet.Columns(sourceMatchFieldColNum), 0)
    currentSheet.Cells(i, destFieldColNum).value = sourceSheet.Cells(sourceMatchRow, sourceMatchDataFieldColNum).value
    
NextItem:
Next i

On Error GoTo 0

Application.ScreenUpdating = False

Exit Sub

ErrHandlerClient:
    Err.Clear
    If (InStr(1, LCase(currentSheet.Cells(1, matchFieldColNum)), "client") > 0) Then
        EmailWhoever "user@domain.com", "", "Unknown client """ & currentSheet.Cells(i, matchFieldColNum) & """ listed in workbook """ & currentSheet.Parent.Name & """ in worksheet " & currentSheet.Name
        currentSheet.Cells(i, matchFieldColNum + 1) = "[UNKNOWN CLIENT]"
    Else
        'EmailWhoever "user@domain.com", "", "There was no match found for """ & currentSheet.Cells(i, matchFieldColNum) & """ listed in workbook """ & currentSheet.Parent.Name & """ in worksheet " & currentSheet.Name
        'currentSheet.Cells(i, matchFieldColNum + 1) = "[UNKNOWN CLIENT]"
    End If
    Resume NextItem

End Sub
Public Sub VbaZipLookup(ByVal currentSheet As Worksheet, ByVal destField As String, ByVal matchField As String, ByVal sourceSheet As Worksheet, ByVal sourceMatchField As String, ByVal sourceMatchDataField As String, Optional ByVal listStates As Boolean, Optional ByVal matchFieldRowNum As Long, Optional ByVal sourceFieldRowNum As Long)
'*** Copied from VbaLookup function - specifically created for the CountyMatching programs - Matches both county and state info to a zip code list***
'*** Edited by Nathan N on 6/15/2012 ***
'// Updated 6/26/2012: Only added '0 NOT FOUND' to state column if it was generated by the caller (to the right of the zip code column)

Dim destFieldColNum  As Integer, matchFieldColNum  As Integer, sourceMatchFieldColNum  As Integer, sourceMatchDataFieldColNum As Integer
Dim i As Long

If (matchFieldRowNum = 0) Then
    matchFieldRowNum = 1
End If
If (sourceFieldRowNum = 0) Then
    sourceFieldRowNum = 1
End If

destFieldColNum = Application.WorksheetFunction.Match(destField, currentSheet.Rows(matchFieldRowNum), 0)
If listStates = True Then stateFieldColNum = Application.WorksheetFunction.Match("State", currentSheet.Rows(matchFieldRowNum), 0)
matchFieldColNum = Application.WorksheetFunction.Match(matchField, currentSheet.Rows(matchFieldRowNum), 0)
sourceMatchFieldColNum = Application.WorksheetFunction.Match(sourceMatchField, sourceSheet.Rows(sourceFieldRowNum), 0)
sourceMatchDataFieldColNum = Application.WorksheetFunction.Match(sourceMatchDataField, sourceSheet.Rows(sourceFieldRowNum), 0)
If listStates = True Then sourceMatchStateDataFieldColNum = Application.WorksheetFunction.Match("State", sourceSheet.Rows(sourceFieldRowNum), 0)

If listStates = True Then stateMatchFieldColNum = Application.WorksheetFunction.Match("State", currentSheet.Rows(matchFieldRowNum), 0)

For i = matchFieldRowNum + 1 To lastRow(currentSheet, , , matchFieldRowNum)
    On Error GoTo ErrHandlerClient
    sourceMatchRow = Application.WorksheetFunction.Match(currentSheet.Cells(i, matchFieldColNum).value, sourceSheet.Columns(sourceMatchFieldColNum), 0)
    currentSheet.Cells(i, destFieldColNum).value = sourceSheet.Cells(sourceMatchRow, sourceMatchDataFieldColNum).value
    If listStates = True Then
        currentSheet.Cells(i, stateFieldColNum).value = sourceSheet.Cells(sourceMatchRow, sourceMatchStateDataFieldColNum).value
    End If
NextItem:
Next i

On Error GoTo 0

Exit Sub

ErrHandlerClient:
    Err.Clear
    If (InStr(1, LCase(currentSheet.Cells(1, matchFieldColNum)), "client") > 0) Then
        EmailWhoever "user@domain.com", "", "Unknown client """ & currentSheet.Cells(i, matchFieldColNum) & """ listed in workbook """ & currentSheet.Parent.Name & """ in worksheet " & currentSheet.Name
        currentSheet.Cells(i, matchFieldColNum + 1) = "[UNKNOWN CLIENT]"
    Else
        currentSheet.Cells(i, matchFieldColNum + 1) = "0 NOT FOUND"
        If listStates = True And currentSheet.Cells(1, matchFieldColNum + 2) = "State" Then currentSheet.Cells(i, matchFieldColNum + 2) = "0 NOT FOUND" '// this may not work properly if the data state column happened to be in the same place
        'EmailWhoever "user@domain.com", "", "There was no match found for """ & currentSheet.Cells(i, matchFieldColNum) & """ listed in workbook """ & currentSheet.Parent.Name & """ in worksheet " & currentSheet.Name
        'currentSheet.Cells(i, matchFieldColNum + 1) = "[UNKNOWN CLIENT]"
    End If
    Resume NextItem

End Sub
Public Function lastEntryInRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
'Created by Nathan N on 4/19/2012
Dim i As Long

For i = 1 To 1000
    
    If (Len(ws.Cells(rowNum, i)) <= 0) Then
        lastEntryInRow = i - 1
        Exit Function
    End If
    
Next i
    
End Function
Public Sub AddColorToCol(ByVal currentSheet As Worksheet, startRow As Long, Optional ByVal colNum As Integer, Optional ByVal colLet As String)
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** Designed to be used with a pivot table, this adds color to a column
'*** within a pivot table, excluding the footer
'*** PreCondition: A starting row number (startRow) must be declared (in case
'***  there are multiple pivot tables in the same column space).
'*********************************************************
Dim lastRow As Long

If colNum > 0 Then
    colLet = ColNumToLet(currentSheet, colNum)
Else
    colLet = ColNumToLet(currentSheet, lastColNum(currentSheet))
End If

    lastRow = currentSheet.Range(colLet & Rows.Count).End(xlUp).Row - 1
    
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions.Delete
    
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions.AddColorScale ColorScaleType:=3
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions.Count).SetFirstPriority
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    With currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 13011546
        .TintAndShade = 0
    End With
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(2).value = 50
    With currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    With currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With

End Sub
Public Sub AddColorToColLimit(ByVal currentSheet As Worksheet, startRow As Long, Optional ByVal colNum As Integer, Optional ByVal colLet As String)
'*********************************************************
'*** Authored by Nathan N, on 2/6/2012
'***
'*** Designed to be used with a pivot table, this adds color to a column
'*** within a pivot table, excluding the footer, according to specified values
'*** PreCondition: A starting row number (startRow) must be declared (in case
'*** there are multiple pivot tables in the same column space).
'*********************************************************

Dim lastRow As Long
Dim cs As ColorScale

If colNum > 0 Then colLet = ColNumToLet(currentSheet, colNum)

lastRow = currentSheet.Range(colLet & Rows.Count).End(xlUp).Row - 1

currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions.Delete

Set cs = currentSheet.Range(colLet & startRow & ":" & colLet & lastRow).FormatConditions.AddColorScale(3)

cs.ColorScaleCriteria(1).Type = xlConditionValueNumber
cs.ColorScaleCriteria(1).value = 0.95
cs.ColorScaleCriteria(1).FormatColor.Color = 7039480

cs.ColorScaleCriteria(2).Type = xlConditionValueNumber
cs.ColorScaleCriteria(2).value = 0.8
cs.ColorScaleCriteria(2).FormatColor.Color = 8711167

cs.ColorScaleCriteria(3).Type = xlConditionValueNumber
cs.ColorScaleCriteria(3).value = 0.7
cs.ColorScaleCriteria(3).FormatColor.Color = 13011546

End Sub
Public Function createPivot(ByVal sourceSheet As Worksheet, ByVal destSheet As Worksheet, ByVal destCell As String, ByVal pivotName As String, _
                                                ByVal countItem As String, Optional ByVal titleName As String, Optional ByVal rowItem1 As String, Optional ByVal colItem1 As String, Optional ByVal rowItem2 As String, _
                                                Optional ByVal rowItem3 As String, Optional ByVal colItem2 As String, Optional ByVal colItem3 As String, Optional lastCol As Integer) As PivotCaches
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** Updated 3/6/2012   - Added optional lastCol parameter
'*** Updated 3/8/2012   - Improved lastCol auto detection (error in excel gets the last col incorrect sometimes)
'***
'*** This creates an aesthetically pleasing pivot table with more ease than
'*** the built-in excel functions. Supports up to three row categories and
'*** three column categories. Note that this pivot table only supports Count;
'*** it does not support Sum.
'***
'*** PreCondition: Field names on the source worksheet must be on the first row.
'***                         Additionally, data must be entered in one or more rows underneath
'***                         the field names
'***
'***                         User must declare the following:
'***                         1. Source workbook name
'***                         2. Source worksheet name
'***                         3. A destination worksheet
'***                         4. A destination cell, ex: "A3"
'***                         5. A Pivot Name
'***                         6. The Field for which the pivot table should Count
'***                         7. Option variables are used to title the pivot table, which
'***                              is different from the pivotName (which is the name that
'***                              is used by excel to manipulate data), as well as add various
'***                              fields to the pivot table
'***
'*** PostCondition: A new pivot table object is created and named pivotName
'*********************************************************

'Dim lastRowP As Long
Dim sourceRange As String
Dim i As Long

'lastRowP = lastRow(sourceSheet)

If (lastCol = 0) Then
    For i = 1 To sourceSheet.UsedRange.Columns.Count + 1
        If (sourceSheet.Cells(1, i) = "") Then
            lastCol = i - 1
            Exit For
        End If
    Next i
End If

sourceRange = sourceSheet.Name & "!R1C1:R" & lastRow(sourceSheet) & "C" & lastCol
'Debug.Print countItem

    sourceSheet.Parent.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange, Version:=xlPivotTableVersion12).CreatePivotTable TableDestination:=destSheet.Range(destCell), TableName:=pivotName, DefaultVersion:=xlPivotTableVersion12
    
    destSheet.PivotTables(pivotName).AddDataField destSheet.PivotTables(pivotName).PivotFields(countItem), "Count of " & countItem, xlCount

'**** If the user has entered the optional parameters below, then the Function adds these to the pivot table dimensions ****
    If colItem1 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(colItem1)
            .Orientation = xlColumnField
            .Position = 1
        End With
    End If
    If rowItem1 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(rowItem1)
            .Orientation = xlRowField
            .Position = 1
        End With
    End If
    If rowItem2 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(rowItem2)
            .Orientation = xlRowField
            .Position = 2
        End With
    End If
    If rowItem3 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(rowItem3)
            .Orientation = xlRowField
            .Position = 3
        End With
    End If
    If colItem2 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(colItem2)
            .Orientation = xlColumnField
            .Position = 2
        End With
    End If
    If colItem3 <> "" Then
        With destSheet.PivotTables(pivotName).PivotFields(colItem3)
            .Orientation = xlColumnField
            .Position = 3
        End With
    End If
    If titleName <> "" Then
        destSheet.PivotTables(pivotName).CompactLayoutRowHeader = titleName
    End If
End Function
Public Sub CreatePivotFromExisting(ByVal sourceSheet As Worksheet, ByVal sourcePivName As String, ByVal destSheet As Worksheet, ByVal destCell As String, _
                                                ByVal destPivotName As String, ByVal countItem As String, Optional ByVal titleName As String, Optional ByVal rowItem1 As String, Optional ByVal colItem1 As String, Optional ByVal rowItem2 As String, _
                                                Optional ByVal rowItem3 As String, Optional ByVal colItem2 As String, Optional ByVal colItem3 As String)
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'*** This creates an aesthetically pleasing pivot table with more ease than
'*** the built-in excel functions. Supports up to three row categories and
'*** three column categories. Note that this pivot table only supports Count;
'*** it does not support Sum.
'*********************************************************

Dim pc As PivotCache, pt As PivotTable

'Set pc = wb.PivotCache(sourcePivotName)
'Set pt = wb.CreatePivotTable destSheet.Range(destCell), destPivotName, True
Set pc = sourceSheet.PivotTables(1).PivotCache
Set pt = pc.CreatePivotTable(destSheet.Range(destCell), destPivotName, True)
    'sourcewb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange, Version:=xlPivotTableVersion12).CreatePivotTable TableDestination:=destSheet.Range(destCell), TableName:=pivotName, DefaultVersion:=xlPivotTableVersion12
pt.AddDataField pt.PivotFields(countItem), "Count of " & countItem, xlCount
   
'pt.AddFields (pt.PivotFields("pivotFieldName").Name)
'**** If the user has entered the optional parameters below, then the Function adds these to the pivot table dimensions ****
    If colItem1 <> "" Then
        With pt.PivotFields(colItem1)
            .Orientation = xlColumnField
            .Position = 1
        End With
    End If
    If rowItem1 <> "" Then
        With pt.PivotFields(rowItem1)
            .Orientation = xlRowField
            .Position = 1
        End With
    End If
    If rowItem2 <> "" Then
        With destSheet.PivotTables(destPivotName).PivotFields(rowItem2)
            .Orientation = xlRowField
            .Position = 2
        End With
    End If
    If rowItem3 <> "" Then
        With destSheet.PivotTables(destPivotName).PivotFields(rowItem3)
            .Orientation = xlRowField
            .Position = 3
        End With
    End If
    If colItem2 <> "" Then
        With destSheet.PivotTables(destPivotName).PivotFields(colItem2)
            .Orientation = xlColumnField
            .Position = 2
        End With
    End If
    If colItem3 <> "" Then
        With destSheet.PivotTables(destPivotName).PivotFields(colItem3)
            .Orientation = xlColumnField
            .Position = 3
        End With
    End If
    
    If titleName <> "" Then destSheet.PivotTables(destPivotName).CompactLayoutRowHeader = titleName
    
End Sub
Public Sub MoveSheets(sheetToMove As Worksheet, sheetAnchor As Worksheet, beforeOrAfter As String)
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** Moves worksheets around a workbook for aesthetic reorganization
'*** the sheetAnchor variable is the worksheet for which the other worksheet
'*** will move around (before or after)
'***
'*** PreCondition: sheetToMove & sheetAnchor are declared.
'***                         beforeOrAfter uses the values "before" or "after"
'*** PostCondition: Does not return a value. beforeOrAfter is set to "".
'*********************************************************

    If beforeOrAfter = "before" Then sheetToMove.Move before:=sheetAnchor

    If beforeOrAfter = "after" Then sheetToMove.Move after:=sheetAnchor
    
    beforeOrAfter = ""
    
End Sub
Public Sub FormatRowToDefault(ByVal ws As Worksheet, ByVal rowNum As Integer, Optional ByVal startCol As Integer, Optional ByVal endCol As Integer, Optional startLet As String, Optional endLet As String)
'*********************************************************
'*** Authored by Nathan N, on 1/13/2012
'***
'*** Updated on 6/8/2012 by  Nathan N: Added default startCol/endCol to equal 1 if no optional argument was provided
'*** Formats extra textual row additions to a pivot table to look as if it is a part of the
'*** default pivot table, as opposed to added text
'***
'*** PreCondition: ws must be declared. startLet and endLet specify
'***                         the range for which the formatting should apply.
'***                         lastRowFormatting is either True or False, and indicates whether the
'***                         formatting should apply to the last row or not. If set to True, then
'***                         the optional topRowNum is not needed.
'***
'*** PostCondition: Does not return a value.
'***
'*********************************************************
Dim rng As Range

If startCol = 0 Then startCol = 1
If endCol = 0 Then endCol = 1

If startCol <> 0 And endCol <> 0 And (startLet = "" Or endLet = "") Then
    Set rng = ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, endCol))
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    rng.Font.Bold = True
    Exit Sub
End If

If startLet <> "" And endLet <> "" Then
    rng = ws.Range(startLet & topRowNum & ":" & endLet & topRowNum)
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ws.Range(rng).Font.Bold = True
    Exit Sub
End If

End Sub
Public Sub FormatColToDefault(ByVal ws As Worksheet, ByVal startCol As Integer, Optional ByVal endCol As Integer)
'*********************************************************
'*** Authored by Nathan N, on 1/27/2012
'***
'*** Formats extra col textual additions to a pivot table to look as if it is a part of the
'*** default pivot table, as opposed to added text
'***
'*** PreCondition: ws must be declared.
'***
'*** PostCondition: Does not return a value.
'***
'*********************************************************
Dim i As Integer

If endCol = 0 Then
    With ws.Columns(startCol)
        .ColumnWidth = 14.57
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
End If

'If endCol <> 0 Then
'    ws.Columns(CStr(startCol) & ":" & CStr(endCol)).ColumnWidth = 14.57
'    ws.Columns(CStr(startCol) & ":" & CStr(endCol)).HorizontalAlignment = xlCenter
'    ws.Columns(CStr(startCol) & ":" & CStr(endCol)).VerticalAlignment = xlCenter
'    ws.Columns(CStr(startCol) & ":" & CStr(endCol)).WrapText = True
'End If

If endCol <> 0 Then
    For i = startCol To endCol
        With ws.Columns(i)
            .ColumnWidth = 14.57
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
End If

End Sub
Public Function FieldExists(ByVal currentSheet As Worksheet, ByVal fieldName As String, Optional ByVal rowNum As Integer, Optional ByVal exactMatch As Boolean) As Boolean
'*** Authored by Nathan N on 1/18/2012 ***
'*** Updated on 2/3/2012 at 11:40AM - Added default rowNum = 1
'*** Updated on 5/2/2012 at 10:22 AM - Included exactMatch functionality

Dim colNum As Integer, fieldList As String, fieldNameLoc As Integer

'** Set Default Row Number (1) if rowNum = 0 **
If rowNum = 0 Then rowNum = 1

fieldList = ""
fieldNameLoc = 0

If (exactMatch = False) Then

    For i = 1 To lastColNum(currentSheet)
        fieldList = currentSheet.Cells(rowNum, i).value & "," & fieldList
    Next i
    
    fieldNameLoc = InStr(1, fieldList, fieldName)
    
    If fieldNameLoc > 0 Then
        FieldExists = True
    Else
        FieldExists = False
    End If
    
    Exit Function
    
Else

    On Error Resume Next
    fieldNameLoc = Application.WorksheetFunction.Match(fieldName, currentSheet.Rows(rowNum), 0)
    
    If fieldNameLoc > 0 Then
        FieldExists = True
    Else
        FieldExists = False
    End If

    On Error GoTo 0
    Exit Function
End If

End Function
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
'*** Authored by Nathan N on 2/3/2012***
On Error Resume Next

    For Each ws In wb.Worksheets
    
        If ws.Name = sheetName Then
        'Debug.Print ws.Name
            SheetExists = True
            On Error GoTo 0
            Exit Function
        End If
        
    Next ws
    
On Error GoTo 0

End Function

Public Sub EmailWhoever(recipients As String, Optional BCC_Recipients As String, Optional emailSubject As String, Optional signatureName As String, Optional attachmentFile As String, Optional txtBody As String)

'*********************************************************
'*** Authored by Nathan N, on 5/1/2009
'*** Note that "signLoc" may be different on your system, and should be adjusted accordingly.
'*** Updated on 5/22/2012 - added option for text in body (txtBody). Also removed conditionals that were not needed.
'*** Updated on 2/1/2012 - replaced loginName strings with loginName() function
'***
'*** Emails any number of recipients with one optional attachment. If more than
'*** one, a semicolon must be used, ex: "1@field.com; 2@woohoo.com".
'*** BCC recipients can also be specified. Does not support a CC list.
'***
'*** PreCondition: At least one recipient must be specified. Everything else is
'***                        optional.
'***
'*** PostCondition: Does not return a value.
'***
'*********************************************************

Dim olApp As Outlook.Application
Dim olMail As MailItem
Dim signLoc As String
Dim signature As String

If Len(signatureName) > 1 Then

    signLoc = "C:\Documents and Settings\" & loginName() & "\Application Data\Microsoft\Signatures\" & signatureName & ".htm"
    
    If Dir(signLoc) <> "" Then
        signature = GetBoiler(signLoc)
    Else
        signature = ""
    End If
    
Else
    signature = ""
End If

'--begin email process--
Set olApp = New Outlook.Application
Set olMail = olApp.CreateItem(olMailItem)

With olMail
    .To = recipients
    .BCC = BCC_Recipients
    .Subject = emailSubject
    .HTMLBody = txtBody & vbCrLf & signature
    If attachmentFile <> "" Then
        .Attachments.Add attachmentFile
    End If
   ' .Display
    .Send
End With

Set olMail = Nothing
Set olApp = Nothing

End Sub
Public Function findAndReplace(textToFind, containedInFile, replaceWith) As String
Dim transferText As Long

transferText = FreeFile

Open containedInFile For Input As #transferText
textContent = Input$(LOF(transferText), transferText)
Close #transferText
newText = Replace(textContent, textToFind, replaceWith)

Open containedInFile For Output As transferText
Print #transferText, newText
Close #transferText

findAndReplace = newText

End Function
Public Sub replaceFileContents(containedInFile, replaceWith)
Dim transferText As Long
Dim transferText2 As Long
transferText = FreeFile
transferText2 = FreeFile

Open containedInFile For Input As #transferText
textContent = Input$(LOF(transferText), transferText)
Close #transferText

'newText = Replace(textContent, textToFind, replaceWith)

Open replaceWith For Input As #transferText2
textContent2 = Input$(LOF(transferText2), transferText2)
Close #transferText2

Open containedInFile For Output As transferText
Print #transferText, textContent2
Close #transferText

End Sub
Public Function FileExists(ByVal fileLocation As String) As Boolean
'// Changed function name to FileExists from FileThere because it made more sense
'// Authored by Nathan N on 6/5/2012
'// Checks to see if a certain file

    FileExists = (Dir(fileLocation) > "")
End Function
Public Function FileThere(ByVal fileLocation As String) As Boolean
'// DEPRECIATED - 6/5/2012
    FileThere = (Dir(fileLocation) > "")
End Function
Public Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function
Public Function DivideText(ByVal txt As String, ByVal target1 As String, ByVal target2 As String) As String ', ByRef before As _
    String, ByRef between As String, ByRef after As String)

'Copied over from VendorProfilePriceSearch project on 5/21/2012

    Dim pos As Long
    pos = InStr(txt, target1)
    If pos = 0 Then Exit Function
    
    before = Left$(txt, pos - 1)
        ' Remove up to target1 from the string.
    txt = Mid$(txt, pos + Len(target1))
    
    pos = InStr(txt, target2)
    'pos = 5
    between = Left(txt, pos - 1)
        ' Remove up to target2 from the string.
    txt = Mid(txt, pos + Len(target2))
        ' Set between.
        
    ' Return what remains.
    after = txt
    'Debug.Print between
    'Debug.Print txt
DivideText = between
Exit Function

errhandler:
DivideText = ""
End Function
Public Sub waitForIE(ByVal IEObject As Object)
Do While IEObject.readyState <> 4 Or IEObject.busy = True
Loop
End Sub
Public Function RegKeyRead(ByVal i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  Set myWS = CreateObject("WScript.Shell")
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function
Public Function RegKeyExists(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  Set myWS = CreateObject("WScript.Shell")
  myWS.RegRead i_RegKey
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  RegKeyExists = False
End Function
Public Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
      Optional i_Type As String = "REG_SZ")
Dim myWS As Object

  Set myWS = CreateObject("WScript.Shell")
  myWS.RegWrite i_RegKey, i_Value, i_Type

End Sub
Public Function RegKeyDelete(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  Set myWS = CreateObject("WScript.Shell")
  myWS.RegDelete i_RegKey
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  RegKeyDelete = False
End Function
Private Function IsTime(sTime As String) As Boolean

'http://www.freevbcode.com/ShowCode.Asp?ID=1321
'by Phil Fresle

    If Left(Trim(sTime), 1) Like "#" Then
        IsTime = IsDate(Date & " " & sTime)
    End If
End Function
Public Sub DeleteLineItem(ByVal ws As Worksheet, ByVal fieldName As String, Optional ByVal serviceStr1 As String, Optional ByVal serviceStr2 As String, Optional ByVal serviceStr3 As String, Optional ByVal deleteAnyNonBlank As Boolean)
'// Deletes entire entries that contain keywords specified by the user. Up to three unique keywords are supported.

'*** Authored by Nathan N on 1/20/2012 ***
'*** Update by Nathan N on 2/2/2012 at 11:25AM - Added multiple string capability
'***                                            - Added fieldName capability
'// Updated 6/28/2012 by Nathan N: Changed structure for efficiency. Also added clause that exited the function more efficiently as well. Also added ability to delete blank entries or Non Blank Entries.

Dim inverseRowNum As Long, origRowCount As Long, serviceColNum As Integer, delCount As Long
Dim currentService As String
Dim startRow As Long, endRow As Long
Dim i As Long

Application.ScreenUpdating = False

serviceStr1 = LCase(serviceStr1)
serviceStr2 = LCase(serviceStr2)
serviceStr3 = LCase(serviceStr3)
fieldName = LCase(fieldName)

origRowCount = lastRow(ws)
serviceColNum = FieldColNum(ws, fieldName)

If deleteAnyNonBlank = True Then
    SortCol ws, fieldName
    If ws.Cells(2, FieldColNum(ws, fieldName)) = "" Then
        Debug.Print "Function DeleteLineItem: No Non-Blank Entries found"
        Exit Sub
    End If
    startRow = 2
    endRow = lastRow(ws, , FieldColNum(ws, fieldName))
    ws.Rows(startRow & ":" & endRow).Delete
    Exit Sub
End If

'// Delete Blank Entries
If serviceStr1 = "" Then
    SortCol ws, fieldName
    startRow = lastRow(ws, , FieldColNum(ws, fieldName)) + 1
    endRow = lastRow(ws)
    
    If startRow = endRow + 1 Then '// if no blank entry was found then exit function
        Debug.Print "Function DeleteLineItem: No Blank Entries detected"
        Exit Sub
    End If
    
    ws.Rows(startRow & ":" & endRow).Delete
    Exit Sub
End If

'// One service to be deleted
If (serviceStr1 <> "" And serviceStr2 = "" And serviceStr3 = "") Then

    For i = 2 To origRowCount
        If i > origRowCount - delCount Then Exit Sub '// This ensures that the function exits as soon as it reaches the NEW last row after all the deletions
        currentService = LCase(ws.Cells(i, serviceColNum))
        If currentService Like "*" & serviceStr1 & "*" Then
            ws.Rows(i).Delete Shift:=xlUp
            i = i - 1
            delCount = delCount + 1
            GoTo NextEntry1
        End If
NextEntry1:
    Next i

'// Two Services to be Deleted
ElseIf (serviceStr1 <> "" And serviceStr2 <> "" And serviceStr3 = "") Then

    For i = 2 To origRowCount
        If i > origRowCount - delCount Then Exit Sub
        currentService = LCase(ws.Cells(i, serviceColNum))
        If (currentService Like "*" & serviceStr1 & "*") _
        Or (currentService Like "*" & serviceStr2 & "*") Then
            ws.Rows(i).Delete Shift:=xlUp
            i = i - 1
            delCount = delCount + 1
            GoTo NextEntry2
        End If
NextEntry2:
    Next i

'// Three Services to be Deleted
ElseIf (serviceStr1 <> "" And serviceStr2 <> "" And serviceStr3 <> "") Then

    For i = 2 To origRowCount
        If i > origRowCount - delCount Then Exit Sub
        currentService = LCase(ws.Cells(i, serviceColNum))
        If currentService Like "*" & serviceStr1 & "*" _
        Or currentService Like "*" & serviceStr2 & "*" _
        Or currentService Like "*" & serviceStr3 & "*" Then
            ws.Rows(i).Delete Shift:=xlUp
            i = i - 1
            delCount = delCount + 1
            GoTo NextEntry3
        End If
NextEntry3:
    Next i
    
End If

End Sub
Sub DeleteMultiLineItem(ByVal ws As Worksheet, ByVal fieldName1 As String, ByVal fieldName2 As String, ByVal entryStr1 As String, ByVal entryStr2 As String)
'*** authored by Nathan N on 2/13/2012***

Dim inverseRowNum As Long, origRowCount As Long, serviceColNum As Integer
Dim currentEntry1 As String, currentEntry2 As String
Dim i As Long

entryStr1 = LCase(entryStr1)
entryStr2 = LCase(entryStr2)

fieldName1 = LCase(fieldName1)
fieldName2 = LCase(fieldName2)

origRowCount = ws.Range("A" & Rows.Count).End(xlUp).Row
entryColNumA = FieldColNum(ws, fieldName1)
entryColNumB = FieldColNum(ws, fieldName2)

For i = 2 To lastRow(ws)

    inverseRowNum = origRowCount - i + 2
    'Debug.Print inverseRowNum
    currentEntry1 = LCase(ws.Cells(inverseRowNum, entryColNumA))
    currentEntry2 = LCase(ws.Cells(inverseRowNum, entryColNumB))
    
    If (entryStr1 <> "" And entryStr2 <> "") Then
        If (currentEntry1 Like "*" & entryStr1 & "*") _
        And (currentEntry2 Like "*" & entryStr2 & "*") Then
            ws.Rows(inverseRowNum).Delete Shift:=xlUp
        End If
    End If
    
Next i
End Sub


Public Function loginName() As String
'*** Authored by Hogan K in 2010 ***
'*** Updated on 1/31/2012 by Nathan N - Changed main variables and removed parameter "i_RegKey" in order to include in the main function
'*** UPdated on 4/26/2012 by Nathan N - Simplified everything.

loginName = Environ("username")

End Function
Public Sub DeleteOtherRecords(ByVal ws As Worksheet, ByVal fieldName As String, ByVal entryToKeep1 As String, Optional ByVal entryToKeep2 As String, Optional ByVal entryToKeep3 As String)
'Created on 3/14/2012 by Nathan N
'*** Deletes all other records except for the those containing the entryToKeep ***
'*** Updated on 5/16/2012 by Nathan N: added "If endRow = -1" conditional to exit sub if desired entry to keep is not found.
'***                                                                          added "If (startRow > endRow) Then Exit Sub" due to a miscalc that deleted the last entry if the entryToKeep was not found
Dim i As Long
Dim startRow As Long, endRow As Long

SortCol ws, fieldName

If (InStr(1, ws.Cells(2, FieldColNum(ws, fieldName)), entryToKeep1) <= 0) Then
    startRow = 2
    endRow = FieldRowNum(ws, entryToKeep1, FieldColNum(ws, fieldName), , , True) - 1
    If endRow = -1 Then Debug.Print ("Function 'DeleteOtherRecord' did not find entryToKeep1: '" & entryToKeep1 & "'")
    If endRow = -1 Then Exit Sub
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).Delete Shift:=xlUp
Else
    startRow = FieldRowNum(ws, entryToKeep1, FieldColNum(ws, fieldName), , , , True) + 1
    endRow = lastRow(ws)
    If (startRow > endRow) Then Exit Sub
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).Delete Shift:=xlUp
    Exit Sub
End If

If (InStr(1, LCase(ws.Cells(lastRow(ws), FieldColNum(ws, fieldName))), LCase(entryToKeep1)) <= 0) Then
    startRow = FieldRowNum(ws, entryToKeep1, FieldColNum(ws, fieldName), , , , True) + 1
    endRow = lastRow(ws)
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).Delete Shift:=xlUp
End If

End Sub
Public Function DirExists(ByVal directoryPath As String) As Boolean
'**Authored by Nathan N 2/16/2012 **
    If Len(Dir(directoryPath, vbDirectory)) = 0 Then
        DirExists = False
    Else
        DirExists = True
    End If
End Function
Public Sub CreateDir(ByVal directoryPath As String)
'// Created by Nathan N on 4/25/2012
'--- Create a new directory if none was present---

If (DirExists(directoryPath) = False) Then MkDir (directoryPath)

End Sub
Public Sub emsg(ByVal moduleName As String, ByVal functionName As String, Optional ByVal addDebugErrMsg As String, Optional ByVal progName As String, Optional emailAdmin As Boolean, Optional ByVal emailUser As Boolean, Optional ByVal adminEmail As String, Optional ByVal quitProgram As Boolean, Optional ByVal wbQuit1 As Workbook, Optional ByVal wbQuit2 As Workbook, Optional ByVal wbQuit3 As Workbook, Optional ByVal wsQuit As Worksheet)
'// Created by Nathan N on 6/13/2012
'// Provides a way to notify people of errors and quit program in a more controlled manner, instead of suddenly 'breaking'
Dim adminFirstName As String
Dim adminErrMsg As String, userErrMsg As String

Debug.Print ("There was an error in the program workbook '" & ThisWorkbook.Name & "', module '" & moduleName & "', function '" & functionName & "'.")
If addDebugErrMsg <> "" Then Debug.Print (addDebugErrMsg)
If quitProgram = True Then Debug.Print ("Program Quit.")

If progName = "" Then
    programName = ThisWorkbook.Name & " program"
Else
    progName = progName & " program"
End If
If adminEmail = "" Then
    adminEmail = "someone who can help"
    adminFirstName = "Hello"
Else
    adminFirstName = FirstName(Left(adminEmail, InStr(1, adminEmail, "@")))
End If
'----------------------------------------------------------------------------------------------------------------
'// Set message template to admin and / or user
If (quitProgram = True) Then
    adminErrMsg = "<h4>" & adminFirstName & ", <br>User " & loginName() & " experienced an error in " & functionName & " function in the " & moduleName & " module while running the " & progName & " within the " & ThisWorkbook.Name & " workbook. I've quit the program to avoid additional errors. <br><br> Please address at your earliest convenience.<br><br> Your Favorite,<br>Report Robot"
    userErrMsg = "<h4>Hello " & FirstName() & ",<br><br>I experienced a problem with the " & progName & " you're trying to run. I've quit the program to avoid additional errors. <br><br> Please contact " & adminEmail & " at your earliest convenience.<br><br> Your Favorite,<br>Report Robot"
ElseIf (quitProgram = False) Then
    adminErrMsg = "<h4>" & adminFirstName & ", <br>User " & loginName() & " experienced an error in " & functionName & " function in the " & moduleName & " module while running the " & progName & " within the " & ThisWorkbook.Name & " workbook. I tried to keep running the program, so it may or may not work out.<br><br>Please address at your earliest convenience.<br><br> Your Favorite,<br>Report Robot"
    userErrMsg = "<h4>Hello " & FirstName() & ",<br><br>I experienced a problem with the " & progName & " you're trying to run. I tried to keep running the program, so it may or may not work out. <br><br> Please contact " & adminEmail & " at your earliest convenience.<br><br> Your Favorite,<br>Report Robot"
End If

'// Send message to admin and / or user
If emailAdmin = True Then EmailWhoever adminEmail, , "Report Robot Error: " & ThisWorkbook.Name & " Workbook", , , adminErrMsg
If emailUser = True Then EmailWhoever loginName() & "@domains.com", , "Report Robot Had an Error: " & progName, , , userErrMsg

If quitProgram = True Then
    Application.DisplayAlerts = False
    If Not (wbQuit1 Is Nothing) Then wbQuit1.Close
    If Not (wbQuit2 Is Nothing) Then wbQuit2.Close
    If Not (wbQuit3 Is Nothing) Then wbQuit3.Close
    If Not (wsQuit Is Nothing) Then wsQuit.Delete
    End '// Ends subroutine
End If

End Sub
Public Sub splitzips(ByVal ws As Worksheet, ByVal zipFieldName As String, Optional ByVal deleteExtraFields As Boolean)
'// Transfers comma seperated zip codes into their own cell. This is needed because an automated 'text-to-columns' typically does not
'// support the number of columns required to do this process
'// Authored by Nathan N on 6/14/2012
Dim i As Long, j As Long
Dim zipCol As Integer
Dim zipList() As String
Dim zipCount As Long
Dim rng As Range

'// Deletes fields I thought were not needed for my purposes
If deleteExtraFields = True Then
    If FieldExists(ws, "Insurance") Then ws.Columns(FieldColNum(ws, "Insurance", , True)).Delete
    If FieldExists(ws, "Services") Then ws.Columns(FieldColNum(ws, "Services", , True)).Delete
    If FieldExists(ws, "Mailing") Then ws.Columns(FieldColNum(ws, "Mailing", , True)).Delete
    If FieldExists(ws, "Physical") Then ws.Columns(FieldColNum(ws, "Physical", , True)).Delete
End If

zipCol = FieldColNum(ws, "Coverage", , True)

For i = 2 To lastRow(ws)
    If ws.Cells(i, zipCol) = "" Then GoTo NextItem '// prevents strange errors from occuring
    If InStr(1, ws.Cells(i, zipCol), ",") < 1 Then GoTo NextItem  '// prevents strange errors from occuring
    zipList = Split(ws.Cells(i, zipCol), ", ")
    zipCount = Application.WorksheetFunction.CountA(zipList)
    ws.Range(ws.Cells(i, zipCol), ws.Cells(i, zipCount)) = Split(ws.Cells(i, zipCol), ",")
    Set rng = ws.Range(ws.Cells(i, zipCol), ws.Cells(i, lastColNum(ws, i)))
    rng.value = rng.value
    Set rng = Nothing
    ReDim zipList(0)
NextItem:
Next i

End Sub
Public Sub deleteColRange(ByVal startCol As Integer, ByVal endCol As Integer)
'// Enables the deletion by range numbers, otherwise columns have to be deleted one-by-one,
'// or using letters, which is rather inconvenient.
'// Authored by Nathan N on 4.18.2014

Dim colArray As Variant
Dim colRange As Range
Dim rangeIter As Integer
Dim i As Long

colArray = ""
rangeIter = endCol - startCol

For i = 0 To rangeIter
    colArray = colArray & Str(startCol + i) & ","
Next i

colArray = Left(colArray, Len(colArray) - 1)
colArray = Split(colArray, ",")

Set colRange = Columns(colArray(0))
For i = 1 To UBound(colArray)
    Set colRange = Union(colRange, Columns(colArray(i)))
Next i

colRange.Delete Shift:=xlToLeft

End Sub
Public Sub listifyAndJsonify()
'// Turns two rows of data into a List (array) and JSON formatted data
'// i.e. [{"City":"San Jose, "Lat":0.3,"Lon":1.4},{"City":"Austin","Lat":0.5, "Lon":11.1}]
'/ Output will be on the first row of the third column: cell(1,3)
'// This function must be run  from the VBA IDE.

Dim wb As Workbook
Dim ws As Worksheet
Dim wbName As String, wsName As String
Dim key0 As String, key1 As String
Dim value0 As String, value1 As String
Dim s As String
Dim i As Long



key = ""
s = ""

While wbName = ""
workbookname:
    wbName = ""
    wbName = InputBox("Workbook name?")
Wend

On Error GoTo workbookname
Set wb = Workbooks(wbName)
On Error GoTo 0

While wsName = ""
    wsName = ""
    wsName = InputBox("Sheet Name?")
    If Not SheetExists(wb, wsName) Then
        wsName = ""
    End If
Wend
While key0 = ""
    key0 = InputBox("Key0 name?")
Wend
While key1 = ""
    key1 = InputBox("Key1 name?")
Wend




Set ws = wb.Worksheets(wsName)


If VarType(ws.Cells(2, 2)) = vbDouble Then
    quoteString = ""
Else
    quoteString = """"
End If

For i = 2 To lastRow(ws)
    value0 = ws.Cells(i, 1)
    value1 = ws.Cells(i, 2)
    s = s & "{" & """" & key0 & """" & ":" & """" & value0 & """" & "," & """" & key1 & """" & ":" & quoteString & value1 & quoteString & "},"
Next i

'// Remove the extra comma at the end of the string
s = Left(s, Len(s) - 1)
ws.Cells(1, 3) = "[" & s & "]"

End Sub
Function Get_Variable_Type(myVar)

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' ---------------------------------------------------------------

If VarType(myVar) = vbNull Then
MsgBox "Null (no valid data) "
ElseIf VarType(myVar) = vbInteger Then
MsgBox "Integer "
ElseIf VarType(myVar) = vbLong Then
MsgBox "Long integer "
ElseIf VarType(myVar) = vbSingle Then
MsgBox "Single-precision floating-point number "
ElseIf VarType(myVar) = vbDouble Then
MsgBox "Double-precision floating-point number "
ElseIf VarType(myVar) = vbCurrency Then
MsgBox "Currency value "
ElseIf VarType(myVar) = vbDate Then
MsgBox "Date value "
ElseIf VarType(myVar) = vbString Then
MsgBox "String "
ElseIf VarType(myVar) = vbObject Then
MsgBox "Object "
ElseIf VarType(myVar) = vbError Then
MsgBox "Error value "
ElseIf VarType(myVar) = vbBoolean Then
MsgBox "Boolean value "
ElseIf VarType(myVar) = vbVariant Then
MsgBox "Variant (used only with arrays of variants) "
ElseIf VarType(myVar) = vbDataObject Then
MsgBox "A data access object "
ElseIf VarType(myVar) = vbDecimal Then
MsgBox "Decimal value "
ElseIf VarType(myVar) = vbByte Then
MsgBox "Byte value "
ElseIf VarType(myVar) = vbUserDefinedType Then
MsgBox "Variants that contain user-defined types "
ElseIf VarType(myVar) = vbArray Then
MsgBox "Array "
Else
MsgBox VarType(myVar)
End If

End Function
