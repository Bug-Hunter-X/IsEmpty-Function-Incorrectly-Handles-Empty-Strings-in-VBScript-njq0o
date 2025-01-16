Function f(a, b)
  If Len(a) = 0 Then
    MsgBox "'a' is empty", vbExclamation
  ElseIf IsEmpty(a) Then
    MsgBox "'a' is empty", vbExclamation
  End If
  If Len(b) = 0 Then
    MsgBox "'b' is empty", vbExclamation
  ElseIf IsEmpty(b) Then
    MsgBox "'b' is empty", vbExclamation
  End If
  f = a + b
End Function