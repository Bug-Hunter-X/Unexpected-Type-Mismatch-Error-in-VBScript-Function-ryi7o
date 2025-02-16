Function MyFunction(param1)
  ' Some code here that may cause an error
  If param1 = "" Then
    Err.Raise 13, , "Type mismatch"
  End If
End Function