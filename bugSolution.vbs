Function MyFunction(param1, param2)
  'Explicitly check for empty strings
  If isEmpty(param1) Then
    Err.Raise vbError, , "Param1 cannot be empty"
  End If
  'Handle other conditions
End Function

Function isEmpty(str)
  isEmpty = (Len(Trim(str)) = 0)
End Function