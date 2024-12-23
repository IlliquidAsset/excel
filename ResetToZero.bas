' This macro clears the contents of a user-defined range. 
' Assign your desired range to the "TargetRange" variable.
' Attach it to a button for easy resets.

Sub ResetToZero()
    Dim TargetRange As Range

    ' >>> CHANGE THIS TO THE RANGE YOU WANT TO CLEAR <<<
    Set TargetRange = Range("V5:V361")

    ' Clear the contents of the specified range
    TargetRange.ClearContents

End Sub
