Attribute VB_Name = "ModGraficos"
Option Explicit

Sub GraficarHistograma(ByRef a() As Variant, ByRef B As MSChart, ByVal c As String)
Dim n As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se grafica en el objeto a los datos b
'y se le agrega el titulo c
n = UBound(a, 1)
B.chartType = 1
B.ColumnCount = 1
B.RowCount = n - 2
For i = 1 To n - 2
B.Column = 1
B.row = i
B.Data = a(i, 3)
B.RowLabel = a(i, 1)
Next i
B.Title = Left("HISTOGRAMA - " & c, 80)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub GraficarHistograma2(ByRef a() As Variant, ByRef B As Form, ByVal c As String)
Dim n As Integer
Dim i As Integer
Dim mm As Double
Dim ds As Double
Dim nd As Double
Dim vmin As Double
Dim vmax As Double
Dim fmax As Double
Dim limitel As Integer
Dim limiter As Integer
Dim limiteu As Integer
Dim limited As Integer
Dim Ancho As Integer
Dim Altura As Integer
Dim alto As Integer
Dim EscalaX As Integer
Dim EscalaY As Integer
Dim x As Integer
Dim xa As Integer
Dim Y As Integer
Dim ya As Integer
Dim valx As Integer
Dim valx1 As Integer
Dim NoSeccion As Integer

' se hace uso de una rutina manual
'para graficar el histograma sin el uso de
'el control mschart
'hace uso de los datos del histograma
'se hace uso de los datos a, para graficarlos
'en el objeto b y se agrega el titulo c
n = UBound(a, 1)
mm = a(n - 1, 1)
ds = a(n - 1, 2)
nd = a(n - 1, 3)
vmin = a(n - 1, 4)
vmax = a(n, 1)
fmax = a(n, 2)
B.Cls
B.ForeColor = QBColor(5)
limitel = 40
limiter = B.ScaleWidth - 40
limiteu = 100
limited = B.ScaleHeight - 40
B.Line (limitel, limiteu)-(limiter, limited), , B
Ancho = limiter - limitel
Altura = Maximo(DNormal(mm, mm, ds, 0), fmax)
alto = limited - limiteu
For i = 1 To (n - 2)
valx = (i - 1) * Ancho / (n - 2)
valx1 = (i) * Ancho / (n - 2)
'se grafican las barras del histograma

B.Line (limitel + valx, limited)-(limitel + valx1, limited - a(i, 3) * alto / Altura), QBColor(7), BF
B.Line (limitel + valx, limited)-(limitel + valx1, limited - a(i, 3) * alto / Altura), QBColor(1), B
Next i
EscalaX = (limiter - limitel) / (vmax - vmin)
EscalaY = (limited - limiteu) / Altura
'se grafica la curva norma sobre el histograma
NoSeccion = 100
x = vmin
For i = 1 To NoSeccion
    xa = x + (vmax - vmin) / NoSeccion
    Y = DNormal(x, mm, ds, 0)
    ya = DNormal(x + (vmax - vmin) / NoSeccion, mm, ds, 0)
    B.Line (limitel + EscalaX * (x - vmin), limited - EscalaY * (Y))-(limitel + EscalaX * (xa - vmin), limited - EscalaY * (ya))
    x = x + (vmax - vmin) / NoSeccion
Next i

B.FontSize = 10
B.CurrentY = 20
B.CurrentX = (B.ScaleWidth - B.TextWidth(Left("HISTOGRAMA - " & c, 80))) / 2
B.Print Left("HISTOGRAMA - " & c, 80)
B.CurrentX = (B.ScaleWidth - B.TextWidth("(Muestra de " & nd & " días)")) / 2
B.Print "(Muestra de " & nd & " días)"
B.FontSize = 6
For i = 1 To Int(Altura) + 1
B.CurrentY = limited - (i - 1) * (limited - limiteu) / Int(Altura)
B.CurrentX = 10
B.Print i - 1
Next i

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Sub GraficarSimulacion(ByRef a() As Variant, ByRef B As MSChart, ByVal c As String)
Dim i As Integer
Dim n As Integer
Dim m As Integer
Dim j As Integer
Dim noseries As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'grafico de la simulacion de precios de una posicion
n = UBound(a, 1)
m = UBound(a, 2)
If m Mod 2 = 0 Then
noseries = Int(m / 2)
B.chartType = VtChChartType2dXY
B.ColumnCount = m
B.RowCount = n
'se grafican los limites
For i = 1 To n
B.row = i
For j = 1 To noseries
B.Column = 2 * j - 1
B.Data = a(i, 2 * j - 1) * 115
B.Column = 2 * j
B.Data = a(i, 2 * j)
Next j
Next i
B.Title = c
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TitulosHistograma(ByRef a As MSFlexGrid)
a.Rows = 2
a.Cols = 5
a.row = 0
a.col = 0
a.Text = "No de intervalo"
a.col = 1
a.Text = "Inicio intervalo"
a.col = 2
a.Text = "Fin intervalo"
a.col = 3
a.Text = "Frecuencia"
a.col = 4
a.Text = "Distribución normal"
a.RowHeight(0) = 800
a.ColWidth(0) = 900
a.ColWidth(1) = 900
a.ColWidth(2) = 900
a.ColWidth(3) = 900
End Sub

Sub GraficoXY(ByRef a() As Variant, ByRef objeto As Form, ByVal c As String)
Dim m As Integer
Dim n As Integer
Dim m1 As Integer
Dim i As Integer
Dim j As Integer
Dim xa1 As Integer
Dim xmin1 As Integer
Dim xmax1 As Integer
Dim ya1 As Integer
Dim ymin1 As Integer
Dim ymax1 As Integer
Dim xa2 As Integer
Dim ya2 As Integer
Dim xmin As Integer
Dim xmax As Integer
Dim cont As Integer
Dim ymin As Integer
Dim ymax As Integer
Dim ccolor As String

'En funcion de la matriz a se grafica datos
'de tipo xy en el objeto objeto. Con el Titulo de
'grafico "C"
'¿Que formato debe de tener a?
'el no de columnas de a debe de ser par

n = UBound(a, 1)
m = UBound(a, 2)
If m Mod 2 = 0 And n > 1 Then
m1 = Int(m / 2)
'se cuenta el no de datos por serie
ReDim nodat(1 To m1) As Long
For j = 1 To m1
cont = 0
For i = 1 To n
If Len(Trim(a(i, 2 * j - 1))) Then
cont = cont + 1
Else
Exit For
End If
Next i
nodat(j) = cont
Next j
'se procede a encontrar el valor maximo y minimo
' de las x y de las y con el fin de que
'todos los datos aparezcan en la grafica
xmin = a(1, 1)
xmax = a(1, 1)
For j = 1 To m1
For i = 1 To nodat(j)
xmin = Minimo(xmin, a(i, 2 * j - 1))
xmax = Maximo(xmax, a(i, 2 * j - 1))
Next i
Next j
ymin = a(1, 2)
ymax = a(1, 2)
For j = 1 To m1
    For i = 1 To nodat(j)
        ymin = Minimo(ymin, a(i, 2 * j))
        ymax = Maximo(ymax, a(i, 2 * j))
    Next i
Next j
xmin1 = 50
xmax1 = objeto.ScaleWidth - 50
ymin1 = 100
ymax1 = objeto.ScaleHeight - 50
'se dibuja el cuadro de la gráfica
objeto.Line (xmin1, ymin1)-(xmax1, ymax1), , B

'se procede a graficar todas las series x,y
For j = 1 To m1
'se elige un color al azar para esta serie
    ccolor = RGB(Int(255 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
    For i = 1 To nodat(j) - 1
        xa1 = xmin1 + (a(i, 2 * j - 1) - xmin) * (xmax1 - xmin1) / (xmax - xmin)
        ya1 = ymax1 - (a(i, 2 * j) - ymin) * (ymax1 - ymin1) / (ymax - ymin)
        xa2 = xmin1 + (a(i + 1, 2 * j - 1) - xmin) * (xmax1 - xmin1) / (xmax - xmin)
        ya2 = ymax1 - (a(i + 1, 2 * j) - ymin) * (ymax1 - ymin1) / (ymax - ymin)
        objeto.Line (xa1, ya1)-(xa2, ya2), ccolor
    Next i
Next j
'se agregan los titulos
objeto.FontSize = 14
objeto.CurrentY = 20
objeto.CurrentX = (objeto.ScaleWidth - objeto.TextWidth(c)) / 2
objeto.Print c
Else
 MsgBox "Los datos no estan en el formato esperado"
End If
End Sub


