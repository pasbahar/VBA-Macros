Attribute VB_Name = "VCATtoDS"
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
    If Not Report.Cells(j, 1).Find("%CH%") Is Nothing Then
    Report.Cells(j, 1).Replace what:="%CH%", Replacement:=""
    Report.Cells(j, 1).Font.Color = RGB(204, 0, 0)
    End If
    If Not Report.Cells(j, 1).Find("%ADD%") Is Nothing Then
    Report.Cells(j, 1).Replace what:="%ADD%", Replacement:=""
    Report.Cells(j, 1).Font.Color = RGB(0, 128, 0)
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

Text1 = "BEGIN Batch Parameter ""R_" 'Assign the text we want to search for to our Text1 variable.


'==============================================================================================================================
' This gets an array of row numbers for our text.
'==============================================================================================================================
iCount = Application.WorksheetFunction.CountIf(Report.Columns("A"), "=*" & Text1 & "*") 'Get the total number of instances of our text.
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

    m = Report2.Cells(Report2.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report2.Cells(m, 1).Value = Report.Cells(rText1(c), 1).Value
    Report2.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report2.Cells(m, 1).Replace what:="Begin Batch Parameter", Replacement:=""
    Report2.Cells(m, 1).Replace what:="""", Replacement:=""

    Report2.Cells(m, 2).Value = Report.Cells(rText1(c) + 12, 1).Value
    Report2.Cells(m, 2).Font.Color = Report.Cells(rText1(c) + 12, 1).Font.Color
    Report2.Cells(m, 2).Replace what:="Attr Type = ", Replacement:=""
    
    Report2.Cells(m, 3).Value = Report.Cells(rText1(c) + 15, 1).Value
    Report2.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 15, 1).Font.Color
    Report2.Cells(m, 3).Replace what:="CV = ", Replacement:=""
    Report2.Cells(m, 3).Replace what:=".000", Replacement:=""
    
    Report2.Cells(m, 4).Value = Report.Cells(rText1(c) + 14, 1).Value
    Report2.Cells(m, 4).Font.Color = Report.Cells(rText1(c) + 14, 1).Font.Color
    Report2.Cells(m, 4).Replace what:="EU0 = ", Replacement:=""
    Report2.Cells(m, 4).Replace what:=".000", Replacement:=""
    
    Report2.Cells(m, 5).Value = Report.Cells(rText1(c) + 13, 1).Value
    Report2.Cells(m, 5).Font.Color = Report.Cells(rText1(c) + 13, 1).Font.Color
    Report2.Cells(m, 5).Replace what:="EU100 = ", Replacement:=""
    Report2.Cells(m, 5).Replace what:=".000", Replacement:=""
    
    Report2.Cells(m, 6).Value = Report.Cells(rText1(c) + 17, 1).Value
    Report2.Cells(m, 6).Font.Color = Report.Cells(rText1(c) + 17, 1).Font.Color
    Report2.Cells(m, 6).Replace what:="UNITS = ", Replacement:=""

Next c 'Go to the next array-item.


'Output Parameters tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noInputParam:
Text1 = "Begin Batch Parameter ""L_" 'Assign the text we want to search for to our Text1 variable.

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
iCount = iCount
If iCount = 0 Then GoTo noOutputParam 'If no instances were found.
ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

i = 1 'Assign a temp variable for this next snippet.
For c = 1 To iCount 'Loop through the items in the array.
    Set r = Report.Range("A" & i & ":A" & Report.UsedRange.Rows.Count + 1) 'Get the range starting with the row after the last instance of Text1.
    rText1(c) = r.Find(Text1).Row 'Find the specified text you want to search for and store its row number in our array.
    i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
Next c 'Go to the next array item.

For c = 1 To iCount 'Loop through the array.

    m = Report3.Cells(Report3.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report3.Cells(m, 1).Value = Report.Cells(rText1(c), 1).Value
    Report3.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report3.Cells(m, 1).Replace what:="Begin Batch Parameter", Replacement:=""
    Report3.Cells(m, 1).Replace what:="""", Replacement:=""

    Report3.Cells(m, 2).Value = Report.Cells(rText1(c) + 12, 1).Value
    Report3.Cells(m, 2).Font.Color = Report.Cells(rText1(c) + 12, 1).Font.Color
    Report3.Cells(m, 2).Replace what:="Attr Type = ", Replacement:=""
    
    Report3.Cells(m, 3).Value = Report.Cells(rText1(c) + 14, 1).Value
    Report3.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 14, 1).Font.Color
    Report3.Cells(m, 3).Replace what:="UNITS = ", Replacement:=""


Next c 'Go to the next array-item.



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
Text1 = "Begin Parameter ""P_" 'Assign the text we want to search for to our Text1 variable.

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

    m = Report4.Cells(Report4.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.
    
    Report4.Cells(m, 1).Value = Report.Cells(rText1(c), 1).Value
    Report4.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report4.Cells(m, 1).Replace what:="Begin Batch Parameter", Replacement:=""
    Report4.Cells(m, 1).Replace what:="""", Replacement:=""

    Report4.Cells(m, 2).Value = Report.Cells(rText1(c) + 12, 1).Value
    Report4.Cells(m, 2).Font.Color = Report.Cells(rText1(c) + 12, 1).Font.Color
    Report4.Cells(m, 2).Replace what:="Attr Type = ", Replacement:=""
    
    Report4.Cells(m, 3).Value = Report.Cells(rText1(c) + 13, 1).Value
    Report4.Cells(m, 3).Font.Color = Report.Cells(rText1(c) + 13, 1).Font.Color
    Report4.Cells(m, 3).Replace what:="CV = ", Replacement:=""
    Report4.Cells(m, 3).Replace what:=".000", Replacement:=""

Next c 'Go to the next array-item.



'Dynamic References tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noPhaseParam:
'Enter your "Text1" value below (e.g., "Housing Counseling Agencies")
Text1 = "Begin Parameter ""D_" 'Assign the text we want to search for to our Text1 variable.

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

    m = Report5.Cells(Report5.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report5.Cells(m, 1).Value = Report.Cells(rText1(c), 1).Value
    Report5.Cells(m, 1).Font.Color = Report.Cells(rText1(c), 1).Font.Color
    Report5.Cells(m, 1).Replace what:="Begin Batch Parameter", Replacement:=""
    Report5.Cells(m, 1).Replace what:="""", Replacement:=""

    Report5.Cells(m, 2).Value = Report.Cells(rText1(c) + 11, 1).Value
    Report5.Cells(m, 2).Font.Color = Report.Cells(rText1(c) + 11, 1).Font.Color
    Report5.Cells(m, 2).Replace what:="Attr Type = ", Replacement:=""
    
    Report5.Cells(m, 3).Value = "N/A"

Next c 'Go to the next array-item.
        
'Composites tab+++++++++++++++++++++++++++++++++++++++++++++++++++++++
noDynRefParam:
'Enter your "Text1" value below (e.g., "Housing Counseling Agencies")
Text1 = "Definition = BMS_" 'Assign the text we want to search for to our Text1 variable.

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

      m = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 1).End(xlUp).Row + 1 'Get the row after the last cell in column "A" that contains a value.

    Report6.Cells(m, 1).Value = Report.Cells(rText1(c) - 1, 1).Value
    Report6.Cells(m, 1).Font.Color = Report.Cells(rText1(c) - 1, 1).Font.Color
    Report6.Cells(m, 1).Replace what:="Begin Usage", Replacement:=""
    Report6.Cells(m, 1).Replace what:="""", Replacement:=""

    Report6.Cells(m, 2).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, "BMS_"), _
    InStr(Report.Cells(rText1(c), 1).Value, " (") - InStr(Report.Cells(rText1(c), 1).Value, "BMS_"))
    Report6.Cells(m, 2).Font.Color = Report.Cells(rText1(c) + 11, 1).Font.Color
    
    Report6.Cells(m, 3).Value = "Common Composite"
Next c 'Go to the next array-item.



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

Set EndRange = Report.Columns("A").Find("End Usage", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "BEGIN Step" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report7.Cells(Report7.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("BEGIN Step")
        Report7.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report7.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION =")
        Report7.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=") + 1, Len(r.Value) - InStr(r.Value, "="))
        Report7.Cells(l, 2).Font.Color = r.Font.Color


        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("END Step", after:=r)

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "Begin Action" & "*")

        Report7.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report7.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("Begin Action").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report7.Cells(Report7.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report7.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, """") + 1, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, """") - 1)
            Report7.Cells(m, 3).Font.Color = r.Font.Color
            Report7.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 1, 1).Value, InStr(Report.Cells(rText1(c) + 1, 1).Value, "=") + 1, _
            Len(Report.Cells(rText1(c) + 1, 1).Value) - InStr(Report.Cells(rText1(c) + 1, 1).Value, "="))
            Report7.Cells(m, 4).Font.Color = r.Font.Color
            Report7.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 3, 1).Value, InStr(Report.Cells(rText1(c) + 3, 1).Value, "=") + 2, 1)
            Report7.Cells(m, 5).Font.Color = r.Font.Color

               Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Action Delay")

            'Write action expression
            For k = (rText1(c) + 4) To r.Row - 1
              If Report7.Cells(m, 6).Value = "" Then
              Report7.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report7.Cells(m, 6).Value = Report7.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report7.Cells(m, 6).Font.Color = r.Font.Color
              Report7.Cells(m, 6).Replace what:="Action Text =", Replacement:=""
              Report7.Cells(m, 6).Replace what:="'", Replacement:=""
              Report7.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report7.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report7.Cells(m, 6).Replace what:=";", Replacement:=""
              Report7.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report7.Cells(m, 6).Replace what:="//", Replacement:=""
              Report7.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k

            'Write Comfirm Expression
                Set StartSubrange = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Confirm Expression")
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Expression")


            For k = StartSubrange.Row To r.Row - 1
              If Report7.Cells(m, 8).Value = "" Then
              Report7.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report7.Cells(m, 8).Value = Report7.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report7.Cells(m, 8).Font.Color = r.Font.Color
                Report7.Cells(m, 8).Replace what:="CONFIRM EXPRESSION =""", Replacement:=""
                Report7.Cells(m, 8).Replace what:="'", Replacement:=""
                Report7.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report7.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report7.Cells(m, 8).Replace what:=";", Replacement:=""
                Report7.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report7.Cells(m, 8).Replace what:="//", Replacement:=""
                Report7.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k
        
        
        'Write Delay Expression

            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Time")

         If StartSubrange.Value = "Delay Expression" Then
            Report7.Cells(m, 7).Value = Report.Cells(StartSubrange.Row + 1, 1).Value
            Report7.Cells(m, 7).Value = Mid(Report7.Cells(m, 7).Value, InStr(Report7.Cells(m, 7).Value, "=") + 2, _
            InStr(Report7.Cells(m, 7).Value, ".") - InStr(Report7.Cells(m, 7).Value, "=") - 2)
            Else
           
            For k = StartSubrange.Row To r.Row - 1
                  If Report7.Cells(m, 7).Value = "" Then
                  Report7.Cells(m, 7).Value = Report.Cells(k, 1)
                  Else
                  Report7.Cells(m, 7).Value = Report7.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)
    
                  End If
                Report7.Cells(m, 7).Font.Color = r.Font.Color
                Report7.Cells(m, 7).Replace what:="DELAY EXPRESSION =""", Replacement:="Delay:"
                Report7.Cells(m, 7).Replace what:="'", Replacement:=""
                Report7.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report7.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report7.Cells(m, 7).Replace what:=";", Replacement:=""
                Report7.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report7.Cells(m, 7).Replace what:="//", Replacement:=""
                Report7.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
          
         End If



        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report7.Cells(l, 1), 4)
            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("T" & ParamValue)
            
            Do
                m = m + 1
                Report7.Cells(m, 1) = Mid(r.Value, InStr(r.Value, """T") + 1, Len(r.Value) - InStr(r.Value, """T") - 1)
                Report7.Cells(m, 1).Font.Color = r.Font.Color
                Report7.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report7.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report7.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report7.Cells(m, 4) = Mid(Report.Cells(r.Row + 1, 1).Value, InStr(Report.Cells(r.Row + 1, 1), "=") + 2, _
                Len(Report.Cells(r.Row + 1, 1).Value) - InStr(Report.Cells(r.Row + 1, 1).Value, "=") - 1)
                Report7.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 3 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report7.Cells(m, 6).Value = "" Then
                    Report7.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report7.Cells(m, 6).Value = Report7.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report7.Cells(m, 6).Replace what:="Transition Condition =", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report7.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report7.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report7.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "END Transition") = 0

                 Report7.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("Begin Transition", after:=r)
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

Set EndRange = Report.Columns("A").Find("End Usage", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "BEGIN Step" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report8.Cells(Report8.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("BEGIN Step")
        Report8.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report8.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION =")
        Report8.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=") + 1, Len(r.Value) - InStr(r.Value, "="))
        Report8.Cells(l, 2).Font.Color = r.Font.Color


        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("END Step", after:=r)

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "Begin Action" & "*")

        Report8.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report8.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("Begin Action").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report8.Cells(Report8.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report8.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, """") + 1, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, """") - 1)
            Report8.Cells(m, 3).Font.Color = r.Font.Color
            Report8.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 1, 1).Value, InStr(Report.Cells(rText1(c) + 1, 1).Value, "=") + 1, _
            Len(Report.Cells(rText1(c) + 1, 1).Value) - InStr(Report.Cells(rText1(c) + 1, 1).Value, "="))
            Report8.Cells(m, 4).Font.Color = r.Font.Color
            Report8.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 3, 1).Value, InStr(Report.Cells(rText1(c) + 3, 1).Value, "=") + 2, 1)
            Report8.Cells(m, 5).Font.Color = r.Font.Color

               Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Action Delay")

            'Write action expression
            For k = (rText1(c) + 4) To r.Row - 1
              If Report8.Cells(m, 6).Value = "" Then
              Report8.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report8.Cells(m, 6).Value = Report8.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report8.Cells(m, 6).Font.Color = r.Font.Color
              Report8.Cells(m, 6).Replace what:="Action Text =", Replacement:=""
              Report8.Cells(m, 6).Replace what:="'", Replacement:=""
              Report8.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report8.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report8.Cells(m, 6).Replace what:=";", Replacement:=""
              Report8.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report8.Cells(m, 6).Replace what:="//", Replacement:=""
              Report8.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k

            'Write Comfirm Expression
                Set StartSubrange = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Confirm Expression")
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Expression")


            For k = StartSubrange.Row To r.Row - 1
              If Report8.Cells(m, 8).Value = "" Then
              Report8.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report8.Cells(m, 8).Value = Report8.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report8.Cells(m, 8).Font.Color = r.Font.Color
                Report8.Cells(m, 8).Replace what:="CONFIRM EXPRESSION =""", Replacement:=""
                Report8.Cells(m, 8).Replace what:="'", Replacement:=""
                Report8.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report8.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report8.Cells(m, 8).Replace what:=";", Replacement:=""
                Report8.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report8.Cells(m, 8).Replace what:="//", Replacement:=""
                Report8.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k
        
        
        'Write Delay Expression

            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Time")

         If StartSubrange.Value = "Delay Expression" Then
            Report8.Cells(m, 7).Value = Report.Cells(StartSubrange.Row + 1, 1).Value
            Report8.Cells(m, 7).Value = Mid(Report8.Cells(m, 7).Value, InStr(Report8.Cells(m, 7).Value, "=") + 2, _
            InStr(Report8.Cells(m, 7).Value, ".") - InStr(Report8.Cells(m, 7).Value, "=") - 2)
            Else
           
            For k = StartSubrange.Row To r.Row - 1
                  If Report8.Cells(m, 7).Value = "" Then
                  Report8.Cells(m, 7).Value = Report.Cells(k, 1)
                  Else
                  Report8.Cells(m, 7).Value = Report8.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)
    
                  End If
                Report8.Cells(m, 7).Font.Color = r.Font.Color
                Report8.Cells(m, 7).Replace what:="DELAY EXPRESSION =""", Replacement:="Delay:"
                Report8.Cells(m, 7).Replace what:="'", Replacement:=""
                Report8.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report8.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report8.Cells(m, 7).Replace what:=";", Replacement:=""
                Report8.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report8.Cells(m, 7).Replace what:="//", Replacement:=""
                Report8.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
          
         End If



        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report8.Cells(l, 1), 4)
            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("T" & ParamValue)
            
            Do
                m = m + 1
                Report8.Cells(m, 1) = Mid(r.Value, InStr(r.Value, """T") + 1, Len(r.Value) - InStr(r.Value, """T") - 1)
                Report8.Cells(m, 1).Font.Color = r.Font.Color
                Report8.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report8.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report8.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report8.Cells(m, 4) = Mid(Report.Cells(r.Row + 1, 1).Value, InStr(Report.Cells(r.Row + 1, 1), "=") + 2, _
                Len(Report.Cells(r.Row + 1, 1).Value) - InStr(Report.Cells(r.Row + 1, 1).Value, "=") - 1)
                Report8.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 3 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report8.Cells(m, 6).Value = "" Then
                    Report8.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report8.Cells(m, 6).Value = Report8.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report8.Cells(m, 6).Replace what:="Transition Condition =", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report8.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report8.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report8.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "END Transition") = 0

                 Report8.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("Begin Transition", after:=r)
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

Set EndRange = Report.Columns("A").Find("End Usage", after:=r)


StepCount = Application.WorksheetFunction.CountIf(Report.Range(StartRange.Address, EndRange.Address), "=*" & "BEGIN Step" & "*")
ReDim StepsArray(1 To StepCount) 'Redefine the boundaries of the array.

    For j = 1 To StepCount

        l = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "C" that contains a value.
        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("BEGIN Step")
        Report6.Cells(l, 1) = Mid(r.Value, InStr(r.Value, """S") + 1, Len(r.Value) - InStr(r.Value, """S") - 1)
        Report6.Cells(l, 1).Font.Color = r.Font.Color

        Set r = Report.Range(StartRange.Address, EndRange.Address).Find("DESCRIPTION =")
        Report6.Cells(l, 2) = Mid(r.Value, InStr(r.Value, "=") + 1, Len(r.Value) - InStr(r.Value, "="))
        Report6.Cells(l, 2).Font.Color = r.Font.Color


        Set EndOfStep = Report.Range(StartRange.Address, EndRange.Address).Find("END Step", after:=r)

        Set StartRange = Report.Range("A" & EndOfStep.Row - 1, EndOfStep.Address)

        iCount = Application.WorksheetFunction.CountIf(Report.Range(r.Address, EndOfStep.Address), "=*" & "Begin Action" & "*")

        Report6.Range("A" & l & ":A" & l + iCount - 1).Merge
        Report6.Range("B" & l & ":B" & l + iCount - 1).Merge
        ReDim rText1(1 To iCount) 'Redefine the boundaries of the array.

        i = r.Row 'Assign a temp variable for this next snippet.
        For c = 1 To iCount 'Loop through the items in the array.
            Set r = Report.Range("A" & i, EndOfStep.Address)
            rText1(c) = r.Find("Begin Action").Row 'Find the specified text you want to search for and store its row number in our array.
            Set r = Report.Range("A" & r.Row + 1, EndOfStep.Address) 'Get the range starting with the row after the last instance of Text1.
            i = rText1(c) + 1 'Re-assign the temp variable to equal the row after the last instance of Text1.
        Next c 'Go to the next array item.


        For c = 1 To iCount 'Loop through the array and write action name, Action description and Action Type.

            m = Report6.Cells(Report6.UsedRange.Rows.Count + 1, 4).End(xlUp).Row + 1 'Get the row after the last cell in column "D" that contains a value.
            Report6.Cells(m, 3).Value = Mid(Report.Cells(rText1(c), 1).Value, InStr(Report.Cells(rText1(c), 1).Value, """") + 1, _
            Len(Report.Cells(rText1(c), 1).Value) - InStr(Report.Cells(rText1(c), 1).Value, """") - 1)
            Report6.Cells(m, 3).Font.Color = r.Font.Color
            Report6.Cells(m, 4).Value = Mid(Report.Cells(rText1(c) + 1, 1).Value, InStr(Report.Cells(rText1(c) + 1, 1).Value, "=") + 1, _
            Len(Report.Cells(rText1(c) + 1, 1).Value) - InStr(Report.Cells(rText1(c) + 1, 1).Value, "="))
            Report6.Cells(m, 4).Font.Color = r.Font.Color
            Report6.Cells(m, 5).Value = Mid(Report.Cells(rText1(c) + 3, 1).Value, InStr(Report.Cells(rText1(c) + 3, 1).Value, "=") + 2, 1)
            Report6.Cells(m, 5).Font.Color = r.Font.Color

               Set r = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Action Delay")

            'Write action expression
            For k = (rText1(c) + 4) To r.Row - 1
              If Report6.Cells(m, 6).Value = "" Then
              Report6.Cells(m, 6).Value = Report.Cells(k, 1)
              Else
              Report6.Cells(m, 6).Value = Report6.Cells(m, 6).Value & vbCrLf & Report.Cells(k, 1)
              End If
              Report6.Cells(m, 6).Font.Color = r.Font.Color
              Report6.Cells(m, 6).Replace what:="Action Text =", Replacement:=""
              Report6.Cells(m, 6).Replace what:="'", Replacement:=""
              Report6.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
              Report6.Cells(m, 6).Replace what:="^/", Replacement:=""
              Report6.Cells(m, 6).Replace what:=";", Replacement:=""
              Report6.Cells(m, 6).Replace what:=".CV", Replacement:=""
              Report6.Cells(m, 6).Replace what:="//", Replacement:=""
              Report6.Cells(m, 6).Replace what:="--", Replacement:=""
            Next k

            'Write Comfirm Expression
                Set StartSubrange = Report.Range("A" & (rText1(c) + 4), EndOfStep.Address).Find("Confirm Expression")
                Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Expression")


            For k = StartSubrange.Row To r.Row - 1
              If Report6.Cells(m, 8).Value = "" Then
              Report6.Cells(m, 8).Value = Report.Cells(k, 1)
              Else
              Report6.Cells(m, 8).Value = Report6.Cells(m, 8).Value & vbCrLf & Report.Cells(k, 1)
              End If
                Report6.Cells(m, 8).Font.Color = r.Font.Color
                Report6.Cells(m, 8).Replace what:="CONFIRM EXPRESSION =""", Replacement:=""
                Report6.Cells(m, 8).Replace what:="'", Replacement:=""
                Report6.Cells(m, 8).Replace what:=".CV :=", Replacement:=" ="
                Report6.Cells(m, 8).Replace what:="^/", Replacement:=""
                Report6.Cells(m, 8).Replace what:=";", Replacement:=""
                Report6.Cells(m, 8).Replace what:=".CV", Replacement:=""
                Report6.Cells(m, 8).Replace what:="//", Replacement:=""
                Report6.Cells(m, 8).Replace what:="--", Replacement:=""
            Next k
        
        
        'Write Delay Expression

            Set StartSubrange = r
            Set r = Report.Range(r.Address, EndOfStep.Address).Find("Delay Time")

         If StartSubrange.Value = "Delay Expression" Then
            Report6.Cells(m, 7).Value = Report.Cells(StartSubrange.Row + 1, 1).Value
            Report6.Cells(m, 7).Value = Mid(Report6.Cells(m, 7).Value, InStr(Report6.Cells(m, 7).Value, "=") + 2, _
            InStr(Report6.Cells(m, 7).Value, ".") - InStr(Report6.Cells(m, 7).Value, "=") - 2)
            Else
           
            For k = StartSubrange.Row To r.Row - 1
                  If Report6.Cells(m, 7).Value = "" Then
                  Report6.Cells(m, 7).Value = Report.Cells(k, 1)
                  Else
                  Report6.Cells(m, 7).Value = Report6.Cells(m, 7).Value & vbCrLf & Report.Cells(k, 1)
    
                  End If
                Report6.Cells(m, 7).Font.Color = r.Font.Color
                Report6.Cells(m, 7).Replace what:="DELAY EXPRESSION =", Replacement:="Delay:"
                Report6.Cells(m, 7).Replace what:="'", Replacement:=""
                Report6.Cells(m, 7).Replace what:=".CV :=", Replacement:=" ="
                Report6.Cells(m, 7).Replace what:="^/", Replacement:=""
                Report6.Cells(m, 7).Replace what:=";", Replacement:=""
                Report6.Cells(m, 7).Replace what:=".CV", Replacement:=""
                Report6.Cells(m, 7).Replace what:="//", Replacement:=""
                Report6.Cells(m, 7).Replace what:="--", Replacement:=""

            Next k
          
         End If



        Next c 'Go to the next array-item.

        'Write Transitions
            ParamValue = Right(Report6.Cells(l, 1), 4)
            Set r = Report.Range(StartRange.Address, EndRange.Address).Find("Begin Transition ""T" & ParamValue)
            
            Do
                m = m + 1
                Report6.Cells(m, 1) = Mid(r.Value, InStr(r.Value, """T") + 1, Len(r.Value) - InStr(r.Value, """T") - 1)
                Report6.Cells(m, 1).Font.Color = r.Font.Color
                Report6.Range("B" & m & ":C" & m).Interior.Color = RGB(191, 191, 191)
                Report6.Cells(m, 5).Interior.Color = RGB(191, 191, 191)
                Report6.Range("G" & m & ":H" & m).Interior.Color = RGB(191, 191, 191)
                Report6.Cells(m, 4) = Mid(Report.Cells(r.Row + 1, 1).Value, InStr(Report.Cells(r.Row + 1, 1), "=") + 2, _
                Len(Report.Cells(r.Row + 1, 1).Value) - InStr(Report.Cells(r.Row + 1, 1).Value, "=") - 1)
                Report6.Cells(m, 4).Font.Color = r.Font.Color
                'Write Transition Expression
                k = 3 'Transition Expression starts 5 lines below of thansition name
                 Do
                    If Report6.Cells(m, 6).Value = "" Then
                    Report6.Cells(m, 6).Value = Report.Cells(r.Row + k, 1)
                    Else
                    Report6.Cells(m, 6).Value = Report6.Cells(m, 6).Value & vbCrLf & Report.Cells(r.Row + k, 1)
                    End If
                    k = k + 1
                    Report6.Cells(m, 6).Replace what:="Transition Condition =", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="'", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=".CV :=", Replacement:=" ="
                    Report6.Cells(m, 6).Replace what:="^/", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=";", Replacement:=""
                    Report6.Cells(m, 6).Replace what:=".CV", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="//", Replacement:=""
                    Report6.Cells(m, 6).Replace what:="--", Replacement:=""
                 Loop While InStr(Report.Cells(r.Row + k, 1).Value, "END Transition") = 0

                 Report6.Cells(m, 6).Font.Color = r.Font.Color

                Set r = Report.Columns("A").Find("Begin Transition", after:=r)
            Loop While InStr(r.Value, ParamValue) > 0

    Next j
    'sort transitions and steps if new Setps/Transitions were added

Dim StartOfCut As Range

   While Not Report6.Cells(Report6.UsedRange.Rows.Count, 1) = "T9900"
    ParamValue = Mid(Report6.Cells(Report6.UsedRange.Rows.Count, 1), InStr(Report6.Cells(Report6.UsedRange.Rows.Count, 1), "T") + 1, 4)
    Set StartOfCut = Report6.Columns("A").Find("S" & ParamValue)
    Set r = Report6.Columns("A").Find("T")
    For i = 2 To Report6.UsedRange.Rows.Count
        Set r = Report6.Columns("A").FindNext(after:=r)
        If Report6.Cells(Report6.UsedRange.Rows.Count, 1).Value > Report6.Cells(r.Row, 1).Value And _
        Report6.Cells(Report6.UsedRange.Rows.Count, 1).Value < Report6.Cells(Report6.Columns("A").FindNext(after:=r).Row, 1).Value Then
            Report6.Rows(StartOfCut.Row & ":" & Report6.UsedRange.Rows.Count).Cut
            Report6.Rows(r.Row + 1).Insert Shift:=xlDown
            Exit For
        End If
    Next i
   Wend

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

   



