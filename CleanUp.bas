Attribute VB_Name = "CleanUp"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Sheets
    If Not sht.Name = "fhx" Then
        sht.Select
    sht.UsedRange.Select
    Selection.Clear
    End If
    Next sht

   
End Sub

