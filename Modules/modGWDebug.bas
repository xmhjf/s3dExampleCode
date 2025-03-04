Attribute VB_Name = "modGWDebug"
Option Explicit

' Debug-functions:
Public Function strVector(p As DVector) As String
strVector = "( " & f(p.x) & ", " & f(p.y) & ", " & f(p.z) & " )"
End Function
Public Function strPosition(p As AutoMath.DPosition) As String
strPosition = "( " & f(p.x) & ", " & f(p.y) & ", " & f(p.z) & " )"
End Function

Public Function strBx(s As String, Bx() As DVector) As String
Dim i As Long
For i = LBound(Bx) To UBound(Bx)
    strBx = strBx & vbCrLf & s & " " & i & "= " & strVector(Bx(i))
Next i
End Function

Private Function f(d As Double) As String
f = Format(d, "###0.000")
End Function

