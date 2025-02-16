Function MyFunction(param1)
  On Error Resume Next
  ' Check for empty string
  If IsEmpty(param1) Or param1 = "" Then
    ' Handle the empty string case gracefully
    MyFunction = "Parameter is empty"
  Else
    ' Process param1 as appropriate
    MyFunction = UCase(param1) ' Example processing
  End If
  If Err.Number <> 0 Then
    ' Handle other unexpected errors
    MsgBox "An error occurred: " & Err.Description
    Err.Clear
  End If
End Function