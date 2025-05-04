Attribute VB_Name = "Module_ConvertToUpperCase"
Sub ConvertToUpperCase()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select Range", "Range Selection", Type:=8)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        rng.Value = Evaluate("INDEX(UPPER(" & rng.Address & "),)")
    End If
End Sub
