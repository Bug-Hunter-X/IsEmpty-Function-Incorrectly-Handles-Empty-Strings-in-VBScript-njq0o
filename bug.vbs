Function f(a,b)
  If IsEmpty(a) Then
    MsgBox "'a' is empty", vbExclamation
  End If
  If IsEmpty(b) Then
    MsgBox "'b' is empty", vbExclamation
  End If
  f = a+b
End Function