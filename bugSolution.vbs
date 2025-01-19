Function f(a, b)
  If IsEmpty(a) Or a = "" Then
    a = 0
  End If
  If IsEmpty(b) Or b = "" Then
    b = 0
  End If
  f = a + b
End Function

MsgBox f(1, "")
MsgBox f(Null,Null) 