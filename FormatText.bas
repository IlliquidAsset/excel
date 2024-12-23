Attribute VB_Name = "FormatText"

' This macro re-applies the existing cell styles throughout all sheets.
' It's useful if you've set up a "master" sheet with custom styles, then
' imported or changed data in other sheets, and need to ensure consistent styling.
'
' SETUP & PREREQS:
' 1. Enable macros in Excel.
' 2. Define or import your custom styles in at least one sheet (the "master style" sheet).
' 3. Run this macro to refresh/re-apply those styles across every sheet.

Sub UpdateCellStyles()
    Dim ws As Worksheet       ' Looping variable for each worksheet in the workbook
    Dim cell As Range         ' Looping variable for each cell in the used range
    Dim styleName As String   ' Temporary store for the cell's current style
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' For each cell in the used range of the current sheet
        For Each cell In ws.UsedRange
            ' Capture the current style
            styleName = cell.Style
            
            ' Reapply the same style if it exists
            If styleName <> "" Then
                cell.Style = styleName
            End If
        Next cell
    Next ws
End Sub
