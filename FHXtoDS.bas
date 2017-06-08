Attribute VB_Name = "FHXtoDS"
Sub fhx()
Dim Report As Worksheet, bReport As Workbook, Report2 As Worksheet, _
Report3 As Worksheet, Report4 As Worksheet, Report5 As Worksheet, Report6 As Worksheet, _
Report7 As Worksheet, Report8 As Worksheet, Report9 As Worksheet, Report10 As Worksheet  ' Create your worksheet and workbook variables.
Dim i As Integer, k As Integer, j As Integer, m As Integer, l As Integer 'Create some variables for counting.
Dim iCount As Integer, c As Integer 'This variable will hold the index of the array of "Text1" instances.
Dim myDate As String, Text3 As String, ParamValue As String, ParamName As String, _
OutputParamValue As String, OutputParamName As String, _
PhaseParamValue As String, PhaseParamName As String, _
DynRefParamValue As String, DynRefParamName As String, _
Text1 As String, Data_Test As String, Data2 As String, _
Data3 As String, Data4 As String, Data5 As String, Data6 As String, _
Data7 As String, Data8 As String, Data9 As String, Data10 As String 'Create some string variables to hold your data.
Dim rText1() As Integer 'Create an array to store the row numbers we'll reference later.
Dim r As Range, rEnd As Range 'Create a range variable to hold the range we need.

'==============================================================================================================================
' This assigns our worksheet and workbook variables.
'==============================================================================================================================
Data_Test = "fhx" 'Assign the name of our "Data_Test" worksheet.

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set bReport = Excel.ActiveWorkbook 'Set your current workbook to our workbook variable.
Set Report = bReport.Worksheets(Data_Test) 'Set the Data_Test worksheet to our first worksheet variable.

On Error GoTo 0 'Reset the error-catcher to default.
'Operations for processing of raw data
If Report.Cells(1, 2).Value = 1 Then GoTo FirstTab 'this line checks if data already has been processed and if yes, skips that part of code
l = Report.UsedRange.Columns.Count
For i = l To 2 Step -1
    For j = 1 To Report.UsedRange.Rows.Count
    If Not IsEmpty(Report.Cells(j, i)) Then
     Report.Cells(j, 1).Value = Report.Cells(j, i).Value
    End If
    Report.Cells(j, 8).FormulaR1C1 = "=TRIM(RC[-7])"
    Next j
Next i

Report.Columns("H").Copy
Report.Columns("A").PasteSpecial Paste:=xlPasteValues
Report.Columns("B:H").ClearContents

Report.Cells(1, 2).Value = "1"

FirstTab:
'Input Parameters tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Enter the names of your two worksheets below
Data2 = "InputParam" 'Assign the name of our "Data2" worksheet.

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report2 = bReport.Worksheets(Data2) 'Set the Data2 worksheet to our second worksheet variable.
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report2.Cells(1, 1) = "Parameter"
Report2.Cells(1, 2) = "Type"
Report2.Cells(1, 3) = "Default Value"
Report2.Cells(1, 4) = "Low Limit"
Report2.Cells(1, 5) = "High Limit"
Report2.Cells(1, 6) = "Units"

Text1 = "ATTRIBUTE_INSTANCE NAME=""R_" 'Assign the text we want to search for to our Text1 variable.


'==============================================================================================================================
' This gets an array of row numbers for our text.
'==============================================================================================================================
iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
iCount = iCount / 2
If iCount = 0 Then GoTo noInputParam 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.

'==============================================================================================================================
' This loops through the array and creates input parameters tab
'==============================================================================================================================
For c = 1 To iCount 'Loop through the array.

    ParamName = Report.Cells(rText1(c), 1).Value 'Subtract the date row by six, and store the "Text2"/[city, state, zip] value in our Text2 variable.
    InputParamValue = Report.Cells(rText1(c) + 2, 1).Value
    m = Report2.Cells(Report2.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report2.Cells(m, 1).Value = ParamName
    Report2.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report2.Cells(m, 7).Value = InputParamValue

    Report2.Cells(m, 1).Replace what:="ATTRIBUTE_INSTANCE NAME=""", Replacement:=""
    Report2.Cells(m, 1).Replace what:="""", Replacement:=""


    ParamName = """" & Report2.Cells(m, 1) & """"
    Set r = Report.Columns("A").Find(ParamName, SearchDirection:=xlPrevious)
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "TYPE") Then
            Report2.Cells(m, 2) = r.Cells.Value
        End If
    Loop While InStr(r.Cells.Value, "TYPE") > 0

            If InStr(Report2.Cells(m, 2).Value, "FLOAT") Then
            Report2.Cells(m, 2) = "REAL"
            End If
            If InStr(Report2.Cells(m, 2).Value, "STRING") Then
            Report2.Cells(m, 2) = "STRING"
            End If
            If InStr(Report2.Cells(m, 2).Value, "INTEGER") Then
            Report2.Cells(m, 2) = "INTEGER"
            End If
    Report2.Cells(m, 2).Font.Color = r.Font.Color

    Report2.Cells(m, 3).FormulaR1C1 = _
        "=MID(RC[4],FIND(""CV"",RC[4])+3,FIND(""UNITS"",RC[4])-1-FIND(""CV"",RC[4])-3)"
    Report2.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

        Report2.Cells(m, 4).FormulaR1C1 = _
        "=MID(RC[3],FIND(""LOW"",RC[3])+4,FIND(""SCALABLE"",RC[3])-1-FIND(""LOW"",RC[3])-4)"
    Report2.Cells(m, 4).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

        Report2.Cells(m, 5).FormulaR1C1 = _
        "=MID(RC[2],FIND(""HIGH"",RC[2])+5,FIND(""LOW"",RC[2])-1-FIND(""HIGH"",RC[2])-5)"
    Report2.Cells(m, 5).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

        Report2.Cells(m, 6).FormulaR1C1 = _
        "=MID(RC[1],FIND(""UNITS"",RC[1])+7,FIND(""}"",RC[1])-2-FIND(""UNITS"",RC[1])-7)"
    Report2.Cells(m, 6).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

Next c 'Go to the next array-item.

    Report2.Columns("C:F").Copy
    Report2.Columns("C:F").PasteSpecial Paste:=xlPasteValues
    Report2.Columns("G:G").Clear


'Output Parameters tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noInputParam:
Text1 = "ATTRIBUTE_INSTANCE NAME=""L_" 'Assign the text we want to search for to our Text1 variable.

'Enter the names of your two worksheets below
Data3 = "OutputParam"

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report3 = bReport.Worksheets(Data3)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report3.Cells(1, 1) = "Parameter"
Report3.Cells(1, 2) = "Type"
Report3.Cells(1, 3) = "Units"
Report3.Cells(1, 4) = "Comments"


'==============================================================================================================================
' This gets an array of row numbers for our text.
'==============================================================================================================================
iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
iCount = iCount / 2
If iCount = 0 Then GoTo noOutputParam 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.

For c = 1 To iCount 'Loop through the array.

    ParamName = Report.Cells(rText1(c), 1).Value 'Subtract the date row by six, and store the "Text2"/[city, state, zip] value in our Text2 variable.
    ParamValue = Report.Cells(rText1(c) + 2, 1).Value
    m = Report3.Cells(Report3.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report3.Cells(m, 1).Value = ParamName 'Paste the value of the city,state,zip into the first available cell in column "A"
    Report3.Cells(m, 7).Value = ParamValue
    Report3.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report3.Cells(m, 1).Replace what:="ATTRIBUTE_INSTANCE NAME=""", Replacement:=""
    Report3.Cells(m, 1).Replace what:="""", Replacement:=""

    ParamName = """" & Report3.Cells(m, 1) & """"
    Set r = Report.Columns("A").Find(ParamName, SearchDirection:=xlPrevious)
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "TYPE") Then
            Report3.Cells(m, 2) = r.Cells.Value
        End If
    Loop While InStr(r.Cells.Value, "TYPE") > 0
       Report3.Cells(m, 2).Font.Color = r.Font.Color

            If InStr(Report3.Cells(m, 2).Value, "FLOAT") Then
            Report3.Cells(m, 2) = "FLOAT"
            End If
            If InStr(Report3.Cells(m, 2).Value, "STRING") Then
            Report3.Cells(m, 2) = "STRING"
            End If
            If InStr(Report3.Cells(m, 2).Value, "INTEGER") Then
            Report3.Cells(m, 2) = "INTEGER"
            End If


        Report3.Cells(m, 3).FormulaR1C1 = _
        "=MID(RC[4],FIND(""UNITS"",RC[4])+7,FIND(""}"",RC[4])-2-FIND(""UNITS"",RC[4])-7)"
       Report3.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

    If InStr(Report3.Cells(m, 7).Value, "UNITS") = 0 Then
        Report3.Cells(m, 3).Value = "N/A"
    End If

Next c 'Go to the next array-item.

    Report3.Columns("C").Copy
    Report3.Columns("C").PasteSpecial Paste:=xlPasteValues
    Report3.Columns("G").Clear


'Failure Conditions tab(Not working yet)+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noOutputParam:
Text1 = "Failure Condition" 'Assign the text we want to search for to our Text1 variable.

'Enter the names of your two worksheets below
Data9 = "Fail_Cond"

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report9 = bReport.Worksheets(Data9)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report9.Cells(1, 1) = "CND No"
Report9.Cells(1, 2) = "Condition"
Report9.Cells(1, 3) = "Description"
Report9.Cells(1, 4) = "Delay"


'''==============================================================================================================================
''' This gets an array of row numbers for our text.
'''==============================================================================================================================
''iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
''iCount = iCount / 2
''If iCount = 0 Then GoTo noFailCond 'If no instances were found.
''ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.
''
''i = 1 'Assign a temp variable for this next snippet.
''For c = 1 To iCount 'Loop through the items in the array.
''    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
''    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
''    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
''Next c 'Go to the next array item.
''
''For c = 1 To iCount 'Loop through the array.
''
''    ParamName = Report.Cells(rText1(c), 1).Value 'Subtract the date row by six, and store the "Text2"/[city, state, zip] value in our Text2 variable.
''    ParamValue = Report.Cells(rText1(c) + 2, 1).Value
''    m = Report9.Cells(Report9.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.
''
''    Report9.Cells(m, 1).Value = ParamName 'Paste the value of the city,state,zip into the first available cell in column "A"
''    Report9.Cells(m, 7).Value = ParamValue
''    Report9.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
''    Report9.Cells(m, 1).Replace what:="ATTRIBUTE_INSTANCE NAME=""", Replacement:=""
''    Report9.Cells(m, 1).Replace what:="""", Replacement:=""
''
''    ParamName = """" & Report9.Cells(m, 1) & """"
''    Set r = Report.Columns("A").Find(ParamName, SearchDirection:=xlPrevious)
''    Do
''        Set r = Report.Columns("A").FindPrevious(after:=r)
''        If InStr(r.Cells.Value, "TYPE") Then
''            Report9.Cells(m, 2) = r.Cells.Value
''        End If
''    Loop While InStr(r.Cells.Value, "TYPE") > 0
''       Report9.Cells(m, 2).Font.Color = r.Font.Color
''
''            If InStr(Report9.Cells(m, 2).Value, "FLOAT") Then
''            Report9.Cells(m, 2) = "FLOAT"
''            End If
''            If InStr(Report9.Cells(m, 2).Value, "STRING") Then
''            Report9.Cells(m, 2) = "STRING"
''            End If
''            If InStr(Report9.Cells(m, 2).Value, "INTEGER") Then
''            Report9.Cells(m, 2) = "INTEGER"
''            End If
''
''
''        Report9.Cells(m, 3).FormulaR1C1 = _
''        "=MID(RC[4],FIND(""UNITS"",RC[4])+7,FIND(""}"",RC[4])-2-FIND(""UNITS"",RC[4])-7)"
''       Report9.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color
''
''    If InStr(Report9.Cells(m, 7).Value, "UNITS") = 0 Then
''        Report9.Cells(m, 3).Value = "N/A"
''    End If
''
''Next c 'Go to the next array-item.
''
''    Report9.Columns("C").Copy
''    Report9.Columns("C").PasteSpecial Paste:=xlPasteValues
''    Report9.Columns("G").Clear


'Phase Parameters tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noFailCond:
'Enter your "Text1" value below (e.g., "Housing Counseling Agencies")
Text1 = "ATTRIBUTE_INSTANCE NAME=""P_" 'Assign the text we want to search for to our Text1 variable.

'Enter the names of your two worksheets below
Data4 = "PhaseParam"

'==============================================================================================================================
' This assigns our worksheet variables.
'==============================================================================================================================
On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report4 = bReport.Worksheets(Data4)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report4.Cells(1, 1) = "Parameter"
Report4.Cells(1, 2) = "Type"
Report4.Cells(1, 3) = "Default Value"

iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
If iCount = 0 Then GoTo noPhaseParam 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.


For c = 1 To iCount 'Loop through the array.

    ParamName = Report.Cells(rText1(c), 1).Value
    ParamValue = Report.Cells(rText1(c) + 2, 1).Value
    m = Report4.Cells(Report4.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report4.Cells(m, 1).Value = ParamName 'Paste the value of the city,state,zip into the first available cell in column "A"
    Report4.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report4.Cells(m, 7).Value = ParamValue

    Report4.Cells(m, 1).Replace what:="ATTRIBUTE_INSTANCE NAME=""", Replacement:=""
    Report4.Cells(m, 1).Replace what:="""", Replacement:=""

    ParamName = """" & Report4.Cells(m, 1) & """"
    Set r = Report.Columns("A").Find(ParamName, SearchDirection:=xlPrevious)
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "TYPE=") Then
            Report4.Cells(m, 2) = r.Cells.Value
        End If
    Loop While InStr(r.Cells.Value, "TYPE=") = 0

       Report4.Cells(m, 2).Font.Color = r.Font.Color

            If InStr(Report4.Cells(m, 2).Value, "FLOAT") Then
            Report4.Cells(m, 2) = "Floating Point"
            End If
            If InStr(Report4.Cells(m, 2).Value, "STRING") Then
            Report4.Cells(m, 2) = "String"
            End If
            If InStr(Report4.Cells(m, 2).Value, "UINT8") Then
            Report4.Cells(m, 2) = "8 bit unsigned integer"
            End If
            If InStr(Report4.Cells(m, 2).Value, "BOOLEAN") Then
            Report4.Cells(m, 2) = "Boolean"
            End If
            If InStr(Report4.Cells(m, 2).Value, "ENUMERATION") Then
            Report4.Cells(m, 2) = "Named Set"
            End If

     Report4.Cells(m, 3).FormulaR1C1 = _
        "=MID(RC[4],FIND(""CV"",RC[4])+3,FIND(""}"",RC[4])-1-FIND(""CV"",RC[4])-3)"
      Report4.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 2, 1).Font.Color

Next c 'Go to the next array-item.

    Report4.Cells.Replace what:="""""", Replacement:="Null"

    Report4.Columns("C").Copy
    Report4.Columns("C").PasteSpecial Paste:=xlPasteValues
    Report4.Columns("G").Clear

'Dynamic References tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noPhaseParam:
'Enter your "Text1" value below (e.g., "Housing Counseling Agencies")
Text1 = "ATTRIBUTE NAME=""D_" 'Assign the text we want to search for to our Text1 variable.

'Enter the names of your two worksheets below
Data5 = "DynRefParam"

'==============================================================================================================================
' This assigns our worksheet variables.
'==============================================================================================================================
On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report5 = bReport.Worksheets(Data5)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report5.Cells(1, 1) = "Parameter"
Report5.Cells(1, 2) = "Type"
Report5.Cells(1, 3) = "Default Value"

'==============================================================================================================================
' This gets an array of row numbers for our text.
'==============================================================================================================================
iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
If iCount = 0 Then GoTo noDynRefParam 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.


For c = 1 To iCount 'Loop through the array.

    ParamName = Report.Cells(rText1(c), 1).Value
    m = Report5.Cells(Report5.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report5.Cells(m, 1).Value = ParamName
    Report5.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report5.Cells(m, 2).Value = "Dynamic Reference"
    Report5.Cells(m, 3).Value = "N/A"

    Report5.Cells(m, 1).Replace what:="ATTRIBUTE NAME=""", Replacement:=""
    Report5.Cells(m, 1).Replace what:="TYPE=DYNAMIC_REFERENCE", Replacement:=""
    Report5.Cells(m, 1).Replace what:="""", Replacement:=""
    Report5.Cells(m, 1).Replace what:=" ", Replacement:=""

Next c 'Go to the next array-item.
        
'Composites tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noDynRefParam:
'Enter your "Text1" value below (e.g., "Housing Counseling Agencies")
Text1 = "DEFINITION=""BMS_" 'Assign the text we want to search for to our Text1 variable.

'Enter the names of your two worksheets below
Data6 = "Composites"

'==============================================================================================================================
' This assigns our worksheet variables.
'==============================================================================================================================
On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report6 = bReport.Worksheets(Data6)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report6.Cells(1, 1) = "Composite Instance Name"
Report6.Cells(1, 2) = "Composite Class Name"
Report6.Cells(1, 3) = "Common Composite/Area Specific"
Report6.Cells(1, 4) = "Logic"

'==============================================================================================================================
' This gets an array of row numbers for our text.
'==============================================================================================================================
iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
If iCount = 0 Then GoTo noComposites 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.


For c = 1 To iCount 'Loop through the array.

    ParamValue = Report.Cells(rText1(c), 1).Value 'Subtract the date row by six, and store the "Text2"/[city, state, zip] value in our Text2 variable.
    m = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report6.Cells(m, 7).Value = ParamValue

    Report6.Cells(m, 1).FormulaR1C1 = _
        "=MID(RC[6],FIND(""NAME="",RC[6])+6,FIND(""DEFINITION"",RC[6])-2-FIND(""NAME"",RC[6])-6)"

    Report6.Cells(m, 2).FormulaR1C1 = _
        "=MID(RC[5],FIND(""DEFINITION="",RC[5])+12,LEN(RC[5])-FIND(""DEFINITION"",RC[5])-12)"

    Report6.Cells(m, 3).Value = "Common Composite"
    Report6.Range("A" & m & ":C" & m).Font.Color = Report.Cells(rText1(c), 1).Font.Color
Next c 'Go to the next array-item.

    Report6.Columns("A:B").Copy
    Report6.Columns("A:B").PasteSpecial Paste:=xlPasteValues
    Report6.Columns("G").Clear

'Hold Logic tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noComposites:
'Enter the names of your two worksheets below
Data7 = "Hold"

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report7 = bReport.Worksheets(Data7)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report7.Cells(1, 1) = "Step / Transition"
Report7.Cells(1, 2) = "Step Description"
Report7.Cells(1, 3) = "Action No"
Report7.Cells(1, 4) = "Action/Condition"
Report7.Cells(1, 5) = "Action Type"
Report7.Cells(1, 6) = "Expression"
Report7.Cells(1, 7) = "Delay"
Report7.Cells(1, 8) = "Confirm Expression"

Dim StartRange As Range, EndRange As Range, EndOfStep As Range, StartSubrange As Range, EndOfAction As Range
Dim StepCount As Integer, StepsArray() As Integer
If Report.Columns("A").Find("Hold Logic") Is Nothing Then GoTo noHold

    Set r = Report.Columns("A").Find("Phase Class:")
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "Hold Logic") Then
        Set StartRange = r
        End If
    Loop While InStr(r.Cells.Value, "Hold Logic") = 0

Set EndRange = Report.Columns("A").Find("STEP_TRANSITION_CONNECTION", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "STEP NAME=" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report7.Cells(Report7.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME=")
        Report7.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report7.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION=")
        Report7.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=""") + 2, Len(r.Value) - InStr(r.Value, "=""") - 2)
        Report7.Cells(l, 2).Font.Color = r.Font.Color

        If InStr(Report.Cells(StartRange.Row + 1, 1), "S99") > 0 Then
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("INITIAL_STEP", after:=r)
        Else
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME", after:=r)
        End If

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "ACTION NAME" & "*")

        Report7.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report7.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("ACTION NAME").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report7.Cells(Report7.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report7.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, "=") - 2)
            Report7.Cells(m, 3).Font.Color = r.Font.Color
            Report7.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 2, 1).Value, InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c) + 2, 1).Value) - InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") - 2)
            Report7.Cells(m, 4).Font.Color = r.Font.Color
            Report7.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 4, 1).Value, InStr(Report.Cells(rText1(c) + 4, 1).Value, "=") + 1, 1)
            Report7.Cells(m, 5).Font.Color = r.Font.Color

               Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("DELAY_")
            Set EndOfAction = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("}")
            If r Is Nothing Then
               Set r = EndOfAction
            End If
            'Write action expression
            For k = (rText1(c) + 5) To r.Row - 1
              If Report7.Cells(m, 6).Value = "" Then
              Report7.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report7.Cells(m, 6).Value = Report7.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report7.Cells(m, 6).Font.Color = r.Font.Color
              Report7.Cells(m, 6).Replace what:="EXPRESSION=""", Replacement:=""
              Report7.Cells(m, 6).Replace what:="""", Replacement:=""
              Report7.Cells(m, 6).Replace what:="'", Replacement:=""
              Report7.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report7.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report7.Cells(m, 6).Replace what:=";", Replacement:=""
              Report7.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report7.Cells(m, 6).Replace what:="//", Replacement:=""
              Report7.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k

        'Write Delay Expression
            If r Is Nothing Then
            Report7.Cells(m, 7).Value = "N/A"
            Else
            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_")


            For k = StartSubrange.Row To r.Row - 1
              If Report7.Cells(m, 7).Value = "" Then
              Report7.Cells(m, 7).Value = Report.Cells(k, 1)
              Else
              Report7.Cells(m, 7).Value = Report7.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)

              End If
                Report7.Cells(m, 7).Font.Color = r.Font.Color
                Report7.Cells(m, 7).Replace what:="DELAY_EXPRESSION=""", Replacement:=""
                Report7.Cells(m, 7).Replace what:="DELAY_TIME=", Replacement:=""
                Report7.Cells(m, 7).Replace what:="""", Replacement:=""
                Report7.Cells(m, 7).Replace what:="'", Replacement:=""
                Report7.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report7.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report7.Cells(m, 7).Replace what:=";", Replacement:=""
                Report7.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report7.Cells(m, 7).Replace what:="//", Replacement:=""
                Report7.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
           End If

          'line of code that cleanes up delay expressions if there is need
            If InStr(Report7.Cells(m, 7).Value, "}") Then
                 Report7.Cells(m, 7).Value = Left(Report7.Cells(m, 7).Value, InStr(Report7.Cells(m, 7).Value, "}") - 1)
            End If

                Set StartSubrange = r
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_TIME_OUT")

            'Write Comfirm Expression
            For k = StartSubrange.Row To r.Row - 1
              If Report7.Cells(m, 8).Value = "" Then
              Report7.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report7.Cells(m, 8).Value = Report7.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report7.Cells(m, 8).Font.Color = r.Font.Color
                Report7.Cells(m, 8).Replace what:="CONFIRM_EXPRESSION=""", Replacement:=""
                Report7.Cells(m, 8).Replace what:="""", Replacement:=""
                Report7.Cells(m, 8).Replace what:="'", Replacement:=""
                Report7.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report7.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report7.Cells(m, 8).Replace what:=";", Replacement:=""
                Report7.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report7.Cells(m, 8).Replace what:="//", Replacement:=""
                Report7.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k


        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report7.Cells(l, 1), 4)

            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("T" & ParamValue)
            Do
                m = m + 1
                Report7.Cells(m, 1) = Mid(r.Value, InStr(r.Value, "=") + 2, Len(r.Value) - InStr(r.Value, "=") - 2)
                Report7.Cells(m, 1).Font.Color = r.Font.Color
                Report7.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report7.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report7.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report7.Cells(m, 4) = Mid(Report.Cells(r.Row + 2, 1).Value, InStr(Report.Cells(r.Row + 2, 1), "=""") + 2, _
                Len(Report.Cells(r.Row + 2, 1).Value) - InStr(Report.Cells(r.Row + 2, 1).Value, "=""") - 2)
                Report7.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 5 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report7.Cells(m, 6).Value = "" Then
                    Report7.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report7.Cells(m, 6).Value = Report7.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report7.Cells(m, 6).Replace what:=" EXPRESSION=""", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="""", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report7.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "}") = 0

                 Report7.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("TRANSITION NAME=", after:=r)
            Loop While InStr(r.Value, ParamValue) > 0

    Next j
    
'Abort Logic tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noHold:
'Enter the names of your two worksheets below
Data8 = "Abort"

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report8 = bReport.Worksheets(Data8)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report8.Cells(1, 1) = "Step / Transition"
Report8.Cells(1, 2) = "Step Description"
Report8.Cells(1, 3) = "Action No"
Report8.Cells(1, 4) = "Action/Condition"
Report8.Cells(1, 5) = "Action Type"
Report8.Cells(1, 6) = "Expression"
Report8.Cells(1, 7) = "Delay"
Report8.Cells(1, 8) = "Confirm Expression"

If Report.Columns("A").Find("Abort Logic") Is Nothing Then GoTo noAbort
    Set r = Report.Columns("A").Find("Phase Class:")
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "Abort Logic") Then
        Set StartRange = r
        End If
    Loop While InStr(r.Cells.Value, "Abort Logic") = 0

Set EndRange = Report.Columns("A").Find("STEP_TRANSITION_CONNECTION", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "STEP NAME=" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report8.Cells(Report8.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME=")
        Report8.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report8.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION=")
        Report8.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=""") + 2, Len(r.Value) - InStr(r.Value, "=""") - 2)
        Report8.Cells(l, 2).Font.Color = r.Font.Color

        If InStr(Report.Cells(StartRange.Row + 1, 1), "S99") > 0 Then
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("INITIAL_STEP", after:=r)
        Else
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME", after:=r)
        End If

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "ACTION NAME" & "*")

        Report8.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report8.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("ACTION NAME").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report8.Cells(Report8.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report8.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, "=") - 2)
            Report8.Cells(m, 3).Font.Color = r.Font.Color
            Report8.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 2, 1).Value, InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c) + 2, 1).Value) - InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") - 2)
            Report8.Cells(m, 4).Font.Color = r.Font.Color
            Report8.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 4, 1).Value, InStr(Report.Cells(rText1(c) + 4, 1).Value, "=") + 1, 1)
            Report8.Cells(m, 5).Font.Color = r.Font.Color

               Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("DELAY_")
            Set EndOfAction = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("}")
            If r Is Nothing Then
               Set r = EndOfAction
            End If
            'Write action expression
            For k = (rText1(c) + 5) To r.Row - 1
              If Report8.Cells(m, 6).Value = "" Then
              Report8.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report8.Cells(m, 6).Value = Report8.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report8.Cells(m, 6).Font.Color = r.Font.Color
              Report8.Cells(m, 6).Replace what:="EXPRESSION=""", Replacement:=""
              Report8.Cells(m, 6).Replace what:="""", Replacement:=""
              Report8.Cells(m, 6).Replace what:="'", Replacement:=""
              Report8.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report8.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report8.Cells(m, 6).Replace what:=";", Replacement:=""
              Report8.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report8.Cells(m, 6).Replace what:="//", Replacement:=""
              Report8.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k
        'Write Delay Expression
            If r Is Nothing Then
            Report8.Cells(m, 7).Value = "N/A"
            Else
            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_")


            For k = StartSubrange.Row To r.Row - 1
              If Report8.Cells(m, 7).Value = "" Then
              Report8.Cells(m, 7).Value = Report.Cells(k, 1)
              Else
              Report8.Cells(m, 7).Value = Report8.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)

              End If
                Report8.Cells(m, 7).Font.Color = r.Font.Color
                Report8.Cells(m, 7).Replace what:="DELAY_EXPRESSION=""", Replacement:=""
                Report8.Cells(m, 7).Replace what:="DELAY_TIME=", Replacement:=""
                Report8.Cells(m, 7).Replace what:="""", Replacement:=""
                Report8.Cells(m, 7).Replace what:="'", Replacement:=""
                Report8.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report8.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report8.Cells(m, 7).Replace what:=";", Replacement:=""
                Report8.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report8.Cells(m, 7).Replace what:="//", Replacement:=""
                Report8.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
           End If
          'line of code that cleanes up delay expressions if there is need
            If InStr(Report8.Cells(m, 7).Value, "}") Then
                 Report8.Cells(m, 7).Value = Left(Report8.Cells(m, 7).Value, InStr(Report8.Cells(m, 7).Value, "}") - 1)
            End If

                Set StartSubrange = r
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_TIME_OUT")

            'Write Comfirm Expression
            For k = StartSubrange.Row To r.Row - 1
              If Report8.Cells(m, 8).Value = "" Then
              Report8.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report8.Cells(m, 8).Value = Report8.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report8.Cells(m, 8).Font.Color = r.Font.Color
                Report8.Cells(m, 8).Replace what:="CONFIRM_EXPRESSION=""", Replacement:=""
                Report8.Cells(m, 8).Replace what:="""", Replacement:=""
                Report8.Cells(m, 8).Replace what:="'", Replacement:=""
                Report8.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report8.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report8.Cells(m, 8).Replace what:=";", Replacement:=""
                Report8.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report8.Cells(m, 8).Replace what:="//", Replacement:=""
                Report8.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k


        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report8.Cells(l, 1), 4)

            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("T" & ParamValue)
            Do
                m = m + 1
                Report8.Cells(m, 1) = Mid(r.Value, InStr(r.Value, "=") + 2, Len(r.Value) - InStr(r.Value, "=") - 2)
                Report8.Cells(m, 1).Font.Color = r.Font.Color
                Report8.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report8.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report8.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report8.Cells(m, 4) = Mid(Report.Cells(r.Row + 2, 1).Value, InStr(Report.Cells(r.Row + 2, 1), "=""") + 2, _
                Len(Report.Cells(r.Row + 2, 1).Value) - InStr(Report.Cells(r.Row + 2, 1).Value, "=""") - 2)
                Report8.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 5 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report8.Cells(m, 6).Value = "" Then
                    Report8.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report8.Cells(m, 6).Value = Report8.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report8.Cells(m, 6).Replace what:=" EXPRESSION=""", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="""", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report8.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "}") = 0

                 Report8.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("TRANSITION NAME=", after:=r)
            Loop While InStr(r.Value, ParamValue) > 0

    Next j
    
'Run Logic tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noAbort:
'Enter the names of your two worksheets below
Data6 = "Run"

On Error GoTo wksheetError 'Set an error-catcher in case the worksheets aren't found.
Set Report6 = bReport.Worksheets(Data6)
On Error GoTo 0 'Reset the error-catcher to default.

'Sheet Header
Report6.Cells(1, 1) = "Step / Transition"
Report6.Cells(1, 2) = "Step Description"
Report6.Cells(1, 3) = "Action No"
Report6.Cells(1, 4) = "Action/Condition"
Report6.Cells(1, 5) = "Action Type"
Report6.Cells(1, 6) = "Expression"
Report6.Cells(1, 7) = "Delay"
Report6.Cells(1, 8) = "Confirm Expression"


    If Report.Columns("A").Find("Run Logic") Is Nothing Then GoTo noRun
    Set r = Report.Columns("A").Find("Phase Class:")
    Do
        Set r = Report.Columns("A").FindPrevious(after:=r)
        If InStr(r.Cells.Value, "Run Logic") Then
        Set StartRange = r
        End If
    Loop While InStr(r.Cells.Value, "Run Logic") = 0

Set EndRange = Report.Columns("A").Find("STEP_TRANSITION_CONNECTION", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "STEP NAME=" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME=")
        Report6.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report6.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION=")
        Report6.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=""") + 2, Len(r.Value) - InStr(r.Value, "=""") - 2)
        Report6.Cells(l, 2).Font.Color = r.Font.Color

        If InStr(Report.Cells(StartRange.Row + 1, 1), "S99") > 0 Then
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("INITIAL_STEP", after:=r)
        Else
        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("STEP NAME", after:=r)
        End If

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "ACTION NAME" & "*")

        Report6.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report6.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("ACTION NAME").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report6.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, "=") - 2)
            Report6.Cells(m, 3).Font.Color = r.Font.Color
            Report6.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 2, 1).Value, InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") + 2, _
            Len(Report.Cells(rText1(c) + 2, 1).Value) - InStr(Report.Cells(rText1(c) + 2, 1).Value, "=") - 2)
            Report6.Cells(m, 4).Font.Color = r.Font.Color
            Report6.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 4, 1).Value, InStr(Report.Cells(rText1(c) + 4, 1).Value, "=") + 1, 1)
            Report6.Cells(m, 5).Font.Color = r.Font.Color

            Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("DELAY_")
            Set EndOfAction = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("}")
            If r Is Nothing Then
               Set r = EndOfAction
            End If

            'Write action expression
            For k = (rText1(c) + 5) To r.Row - 1
              If Report6.Cells(m, 6).Value = "" Then
              Report6.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report6.Cells(m, 6).Value = Report6.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report6.Cells(m, 6).Font.Color = r.Font.Color
              Report6.Cells(m, 6).Replace what:="EXPRESSION=""", Replacement:=""
              Report6.Cells(m, 6).Value = Left(Report6.Cells(m, 6).Value, Len(Report6.Cells(m, 6).Value) - 1)
              Report6.Cells(m, 6).Replace what:=" """, Replacement:=""
              Report6.Cells(m, 6).Replace what:=""" ", Replacement:=""
              Report6.Cells(m, 6).Replace what:="'", Replacement:=""
              Report6.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report6.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report6.Cells(m, 6).Replace what:=";", Replacement:=""
              Report6.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report6.Cells(m, 6).Replace what:="//", Replacement:=""
              Report6.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k

            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_")

        'Write Delay Expression
        If r Is Nothing Then
        Report6.Cells(m, 7).Value = "N/A"
        Else
            For k = StartSubrange.Row To r.Row - 1
              If Report6.Cells(m, 7).Value = "" Then
              Report6.Cells(m, 7).Value = Report.Cells(k, 1)
              Else
              Report6.Cells(m, 7).Value = Report6.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)

              End If
                Report6.Cells(m, 7).Font.Color = r.Font.Color
                Report6.Cells(m, 7).Replace what:="DELAY_EXPRESSION=""", Replacement:=""
                Report6.Cells(m, 7).Replace what:="DELAY_TIME=", Replacement:=""
                Report6.Cells(m, 7).Replace what:="""", Replacement:=""
                Report6.Cells(m, 7).Replace what:="'", Replacement:=""
                Report6.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report6.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report6.Cells(m, 7).Replace what:=";", Replacement:=""
                Report6.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report6.Cells(m, 7).Replace what:="//", Replacement:=""
                Report6.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
          End If

          'line of code that cleanes up delay expressions if there is need
            If InStr(Report6.Cells(m, 7).Value, "}") Then
                 Report6.Cells(m, 7).Value = Left(Report6.Cells(m, 7).Value, InStr(Report6.Cells(m, 7).Value, "}") - 1)
            End If

           'Write Comfirm Expression
            If r Is Nothing Then
            Report6.Cells(m, 8).Value = "N/A"

            Else
            Set StartSubrange = r
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("CONFIRM_TIME_OUT")

            For k = StartSubrange.Row To r.Row - 1
              If Report6.Cells(m, 8).Value = "" Then
              Report6.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report6.Cells(m, 8).Value = Report6.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report6.Cells(m, 8).Font.Color = r.Font.Color
                Report6.Cells(m, 8).Replace what:="CONFIRM_EXPRESSION=""", Replacement:=""
                Report6.Cells(m, 8).Replace what:="""", Replacement:=""
                Report6.Cells(m, 8).Replace what:="'", Replacement:=""
                Report6.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report6.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report6.Cells(m, 8).Replace what:=";", Replacement:=""
                Report6.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report6.Cells(m, 8).Replace what:="//", Replacement:=""
                Report6.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k
            End If

        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report6.Cells(l, 1), 4)

            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("T" & ParamValue)
            Do
                m = m + 1
                Report6.Cells(m, 1) = Mid(r.Value, InStr(r.Value, "=") + 2, Len(r.Value) - InStr(r.Value, "=") - 2)
                Report6.Cells(m, 1).Font.Color = r.Font.Color
                Report6.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report6.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report6.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report6.Cells(m, 4) = Mid(Report.Cells(r.Row + 2, 1).Value, InStr(Report.Cells(r.Row + 2, 1), "=""") + 2, _
                Len(Report.Cells(r.Row + 2, 1).Value) - InStr(Report.Cells(r.Row + 2, 1).Value, "=""") - 2)
                Report6.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 5 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report6.Cells(m, 6).Value = "" Then
                    Report6.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report6.Cells(m, 6).Value = Report6.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report6.Cells(m, 6).Replace what:=" EXPRESSION=""", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="""", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report6.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "}") = 0

                 Report6.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("TRANSITION NAME=", after:=r)
            Loop While InStr(r.Value, ParamValue) > 0

    Next j
   
Dim sht As Worksheet
    For Each sht In ThisWorkbook.Sheets
    sht.Select
    If Not sht.Name = "fhx" Then
        sht.UsedRange.Rows(1).Select
        With Selection
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(0, 0, 204)
        End With
        
        sht.UsedRange.Select
        With Selection
            .WrapText = True
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlLeft
            .Font.Name = "Arial"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
        End With
        If Not (sht.Name = "Run" Or sht.Name = "Abort" Or sht.Name = "Hold") Then
            Range("A5:A" & sht.UsedRange.Row).Select
    sht.Sort.SortFields.Clear
    sht.Sort.SortFields.Add Key:=Range("A2")
    Dim LastCol
    With sht.Sort
        .SetRange Range(sht.Cells(2, 1), sht.Cells(sht.UsedRange.Rows.Count, sht.UsedRange.Columns.Count))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        End If
    End If
Next sht

Exit Sub
wksheetError:
    MsgBox ("The worksheet was not found.")
    Exit Sub

noRun:
   
    Exit Sub


End Sub

   

