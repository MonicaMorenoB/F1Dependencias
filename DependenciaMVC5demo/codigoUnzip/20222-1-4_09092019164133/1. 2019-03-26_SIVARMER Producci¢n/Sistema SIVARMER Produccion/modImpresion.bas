Attribute VB_Name = "ModImpresion"
Option Explicit

Sub ImprimirCuadroObjeto(objeto1, objeto2, ByVal posx As Integer, ByVal posy As Integer, ByRef matfilas() As Integer, ByRef matcolumnas() As Integer, ByRef filasimp() As Integer, ByRef colsimp() As Integer)
Dim nofilas As Integer
Dim nocolumnas As Integer
Dim i As Integer
Dim j As Integer
Dim texto As String
Dim largo As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'objeto1 puede ser una matriz o una objeto tabla
'objeto2 es la salida a impresora
'posx es la posicion x incial
'posy es la posicion y incial
'matfilas es la matriz con la informacion de lineas de filas
'matcolumnas es la matriz con la informacion de lineas de columnas
'filasimp indica si la fila se va a imprimir
'colsimp indica si la columna se va a imprimir


'se imprime el cuadro en en objeto 2
'de acuerdo a ciertos parametros

nofilas = UBound(filasimp, 1)
nocolumnas = UBound(colsimp, 1)
For i = 1 To nofilas
 For j = 1 To nocolumnas
 'este es el texto a imprimir
  If IsArray(objeto1) Then
  texto = objeto1(i, j)
  Else
  texto = objeto1.TextMatrix(filasimp(i) - 1, colsimp(j) - 1)
  End If
  largo = objeto2.TextWidth(texto)
  If i = 1 Then
   objeto2.CurrentY = posy + matfilas(i) / 2 - 30
  Else
   objeto2.CurrentY = posy + matfilas(i - 1) + (matfilas(i) - matfilas(i - 1)) / 2 - 30
  End If
  If j = 1 Then
   objeto2.CurrentX = posx + matcolumnas(j) - largo - 100
  Else
   objeto2.CurrentX = posx + matcolumnas(j) - largo - 100
  End If
  objeto2.Print texto
 Next j
Next i

'se imprimen las lineas

 For i = 1 To nofilas
  objeto2.Line (posx, matfilas(i) + posy)-(matcolumnas(nocolumnas) + posx, matfilas(i) + posy)
 Next i
 For i = 1 To nocolumnas
  objeto2.Line (posx + matcolumnas(i), posy)-(matcolumnas(i) + posx, matfilas(nofilas) + posy)
 Next i
 'se imprime la caja
 objeto2.Line (posx, posy)-(matcolumnas(nocolumnas) + posx, matfilas(nofilas) + posy), , B

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ImprimirLineaReporte(objeto1, objeto2, ByRef matfilas() As Integer, ByRef matcolumnas() As Integer, ByRef filasimp() As Integer, ByRef colsimp() As Integer, ByVal indice As Integer)
Dim nocolumnas As Integer
Dim i As Integer
Dim j As Integer
Dim texto As String
Dim largo As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'objeto1 puede ser una matriz o una objeto tabla
'objeto2 es la salida a impresora
'posx es la posicion x incial
'posy es la posicion y incial
'matfilas es la matriz con la informacion de lineas de filas
'matcolumnas es la matriz con la informacion de lineas de columnas
'filasimp indica si la fila se va a imprimir
'colsimp indica si la columna se va a imprimir
'aling indica la alineacion del texto


'se imprime el cuadro en en objeto 2
'de acuerdo a ciertos parametros

nocolumnas = UBound(colsimp, 1)
For i = indice To indice
 For j = 1 To nocolumnas - 1
 'este es el texto a imprimir
  If IsArray(objeto1) Then   'si es matriz
   texto = objeto1(i, j)
  Else
   texto = objeto1.TextMatrix(filasimp(i) - 1, colsimp(j) - 1)
  End If
  largo = objeto2.TextWidth(texto)
  objeto2.CurrentY = matfilas(i) + (matfilas(i + 1) - matfilas(i)) / 2 - 30
  If j = 1 Then
   objeto2.CurrentX = matcolumnas(j + 1) - largo - 100
  Else
   objeto2.CurrentX = matcolumnas(j + 1) - largo - 100
  End If
  objeto2.Print texto
 Next j
Next i

'se imprimen las lineas

 For i = indice To indice
  'objeto2.Line (posx, matfilas(i) + posy)-(matcolumnas(nocolumnas) + posx, matfilas(i) + posy)
 Next i
 For i = 1 To nocolumnas
  objeto2.Line (matcolumnas(i), matfilas(indice))-(matcolumnas(i), matfilas(indice + 1))
 Next i
 'se imprime la caja
 objeto2.Line (matcolumnas(1), matfilas(indice))-(matcolumnas(nocolumnas), matfilas(indice + 1)), , B

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "ImprimirLineaReporte " & error(Err())
On Error GoTo 0
End Sub


