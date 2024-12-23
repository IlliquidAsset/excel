' ****Excel Macro to hide all rows based on a Hide/Show helper column ***
'
' SETUP & PREREQS:
' 1. Enable macros in Excel.
' 2. Decide which sheet you want to toggle rows on, or use ActiveSheet as in this code.
' 3. Make sure column S (or your chosen helper column) contains "Hide" or "Show" (case-insensitive).
'    - For example, put "Hide" in cells for any row you want to toggle, "Show" for rows you always want visible.
' 4. Adjust the range if your data starts on a different row or a different column.

Sub ToggleHide()
    Dim ws As Worksheet          ' Target worksheet
    Dim helperColumn As Range    ' Helper column containing "Hide" / "Show"
    Dim cell As Range            ' Individual cell in helper column
    Dim shouldHide As Boolean    ' State to determine whether to hide or unhide

    ' Set the target worksheet (currently active sheet)
    Set ws = ActiveSheet

    ' Define the helper column range, starting at row 4
    Set helperColumn = ws.Range("S4:S" & ws.Cells(ws.Rows.Count, "S").End(xlUp).Row)

    ' Determine the current state by checking the first "Hide" we find
    For Each cell In helperColumn
        If Not IsEmpty(cell) And VarType(cell.Value) = vbString Then
            If LCase(cell.Value) = "hide" Then
                ' If a "Hide" row is visible, we hide all "Hide" rows, otherwise we show them
                shouldHide = Not cell.EntireRow.Hidden
                Exit For
            End If
        End If
    Next cell

    ' Toggle rows based on "Hide" or "Show"
    For Each cell In helperColumn
        If Not IsEmpty(cell) And VarType(cell.Value) = vbString Then
            If LCase(cell.Value) = "show" Then
                cell.EntireRow.Hidden = False
            ElseIf LCase(cell.Value) = "hide" Then
                cell.EntireRow.Hidden = shouldHide
            End If
        End If
    Next cell

    ' Optional feedback
    If shouldHide Then
        MsgBox "Rows labeled 'Hide' are now HIDDEN."
    Else
        MsgBox "Rows labeled 'Hide' are now VISIBLE."
    End If
End Sub
