Attribute VB_Name = "FormulaBackupToHelperSheet"

' This macro backs up all formulas from the currently active sheet into a "HelperSheet" as text.
'
' SETUP & PREREQS:
' 1. Your workbook must allow VBA macros (enable macros in security settings).
' 2. Ensure you're on the sheet you want to back up before running this macro.
' 3. If a sheet named "HelperSheet" doesn't exist, this script will create it automatically.
' 4. The used range of the active sheet is copied to the HelperSheet. Existing data on HelperSheet is cleared.

Sub StoreFormulasAsText()
    Dim ws As Worksheet  ' The source (active) worksheet
    Dim helper As Worksheet  ' The destination ("HelperSheet")
    Dim r As Range       ' Each cell in the active sheet's used range

    ' Capture the active worksheet
    Set ws = ActiveSheet

    ' Attempt to set the HelperSheet; create one if it doesn't exist
    On Error Resume Next
    Set helper = Worksheets("HelperSheet")
    If helper Is Nothing Then
        Set helper = Worksheets.Add
        helper.Name = "HelperSheet"
    End If
    On Error GoTo 0

    ' Clear out old data from HelperSheet
    helper.Cells.Clear

    ' Loop through each cell in the used range of the active sheet
    For Each r In ws.UsedRange
        If r.HasFormula Then
            ' Save the formula as text (leading apostrophe prevents evaluation)
            helper.Cells(r.Row, r.Column).Value = "'" & r.Formula
        Else
            ' If no formula, just copy the value
            helper.Cells(r.Row, r.Column).Value = r.Value
        End If
    Next r

    ' Confirm success
    MsgBox "Formulas stored as text in HelperSheet.", vbInformation
End Sub

