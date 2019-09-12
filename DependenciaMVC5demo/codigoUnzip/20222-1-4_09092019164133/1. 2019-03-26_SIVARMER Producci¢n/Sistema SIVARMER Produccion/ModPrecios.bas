Attribute VB_Name = "ModPrecios"
Option Explicit

Function DevPosicionDesglosada(ByRef mata() As Variant, ByRef MPrecio() As Double)
Dim n As Integer
Dim i As Integer
Dim j As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'de la posicion obtenida y del vector de precios, obtenidos
'se calcula el valor de la posición
'y se desglosa en una matriz el valor de la posición

n = UBound(mata, 1)
ReDim matp(1 To 20, 1 To 1) As Double
For i = 1 To n
For j = 1 To 28
If mata(i, 2) = j Then
matp(j) = matp(j) + MPrecio(i, 1) * mata(i, 5)
Exit For
End If
Next j
Next i

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function
