Attribute VB_Name = "Funciones"
Option Explicit

Function tanh(ByVal x As Double) As Double
    tanh = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function


Function Logarit(ByVal x As Double) As Double
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 Logarit = Log(Abs(x))
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function Exponen(ByVal x As Double) As Double
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion acota los valores que puede tomar la funcion exponencial
 Exponen = Exp(Minimo(x, 109))
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function SuperOrden(a, ByVal ind As Integer, ByRef exito As Boolean)
Dim nodatos As Long
Dim nocols As Integer
Dim i As Long
Dim n As Long
Dim j As Long
Dim kk As Long
Dim nopasos As Long
Dim nogrupos As Long
Dim pivot As Variant, pivot1 As Long
Dim Maximo1 As Variant, Maximo2 As Variant
Dim Minimo1 As Variant, Minimo2 As Variant
Dim IndMin1 As Integer, IndMin2 As Integer, IndMax1 As Integer
Dim Contadorfinal As Long
Dim indfinal As Long
Dim imin As Long


'algoritmo de ordenacion basado en el metodo
'divide y venceras
exito = True
nodatos = UBound(a, 1)
nocols = UBound(a, 2)
ReDim mata(1 To nodatos, 1 To 2) As Variant
'se ponen los datos originales en una matriz mata
For i = 1 To nodatos
mata(i, 1) = i
mata(i, 2) = a(i, ind)
Next i
nopasos = 500

nogrupos = Maximo(-Int(-nodatos / nopasos), 1)
ReDim matmin(1 To nogrupos) As Long
ReDim matmax(1 To nogrupos) As Long
ReDim MatElementos(1 To nogrupos) As Long
'se determinan los elementos de cada grupo
For i = 1 To nogrupos
matmin(i) = (i - 1) * nopasos                    'elemento incial
matmax(i) = Minimo(i * nopasos + 1, nodatos + 1) 'elemento final
MatElementos(i) = matmax(i) - matmin(i) - 1      'no de elementos
Next i
ReDim Matdos(1 To nodatos, 1 To 2) As Variant
ReDim matpivot(1 To nodatos, 1 To nocols) As Variant
'INTRODUCIMOS MATRICES PARA ReaLIZAR UNA PARTICION DEL CONJUNTO ORIGINAL
'Y REALIZAR UNA ORDENACION EN CADA SUBCONJUNTO DE DATOS

Dim Indices(1 To 10000, 1 To 2) As Long, Valores(1 To 10000, 1 To 2) As Variant
Dim CONTI(1 To 10000) As Integer, MiniFinal As Variant

'PRIMERO SE ORDENA CADA SUBCONJUNTO POR SEPARADO
'se hace con una variante del metodo de la
'burbuja, mientras que se busca el elemento
'minimo, tambien se busca el elemento maximo

For n = 1 To nogrupos
For i = 1 To Int(MatElementos(n) / 2)
'para empezar ponemos el primer elemento y
'el ultimo elemento de cada tabla
Indices(n, 1) = matmin(n) + i
Valores(n, 1) = mata(Indices(n, 1), 2)
Indices(n, 2) = matmin(n) + i
Valores(n, 2) = mata(Indices(n, 2), 2)
For j = matmin(n) + i To matmax(n) - i
'busco el minimo
If Valores(n, 1) > mata(j, 2) Then
Indices(n, 1) = j
Valores(n, 1) = mata(j, 2)
End If
'busco el maximo
If Valores(n, 2) < mata(j, 2) Then
Indices(n, 2) = j
Valores(n, 2) = mata(j, 2)
End If
Next j

'se aplica el método de la burbuja en cada columna de
'en el principio y fin de cada tabla
'primero el elemento minimo
pivot = mata(matmin(n) + i, 2)
pivot1 = mata(matmin(n) + i, 1)
mata(matmin(n) + i, 2) = mata(Indices(n, 1), 2)
mata(matmin(n) + i, 1) = mata(Indices(n, 1), 1)
mata(Indices(n, 1), 2) = pivot
mata(Indices(n, 1), 1) = pivot1

If Indices(n, 2) = matmin(n) + i Then
Indices(n, 2) = Indices(n, 1)
End If
'luego el elemento maximo
pivot = mata(matmax(n) - i, 2)
pivot1 = mata(matmax(n) - i, 1)
mata(matmax(n) - i, 2) = mata(Indices(n, 2), 2)
mata(matmax(n) - i, 1) = mata(Indices(n, 2), 1)
mata(Indices(n, 2), 2) = pivot
mata(Indices(n, 2), 1) = pivot1

Next i
Next n

'DESPUES SE PROCEDE A REVISAR CADA LISTA
'CONSECUTIVAMENTE PARA OBTENER LOS
'ELEMENTOS MINIMOS Y MAXIMOS GENERALES

Contadorfinal = nodatos
For i = 1 To nogrupos
CONTI(i) = 1
Next i

Do While Contadorfinal > 0
'se empieza tomando cualquier elemento
'de la lista de datos

For kk = 1 To nogrupos
If CONTI(kk) <= MatElementos(kk) Then
MiniFinal = mata(matmin(kk) + CONTI(kk), 2)
indfinal = mata(matmin(kk) + CONTI(kk), 1)
imin = kk
Exit For
End If
Next kk

For i = 1 To nogrupos
If CONTI(i) <= MatElementos(i) Then
If MiniFinal > mata(matmin(i) + CONTI(i), 2) Then
MiniFinal = mata(matmin(i) + CONTI(i), 2)
indfinal = mata(matmin(i) + CONTI(i), 1)
imin = i
End If
End If
Next i

If CONTI(imin) <= MatElementos(imin) Then
CONTI(imin) = CONTI(imin) + 1
Contadorfinal = Contadorfinal - 1
Matdos(nodatos - Contadorfinal, 2) = MiniFinal
Matdos(nodatos - Contadorfinal, 1) = indfinal
End If
DoEvents
Loop

'Se verifica ahora que todos
'los datos esten ordenados

For i = 1 To nodatos
If i <> 1 Then
If Matdos(i - 1, 2) > Matdos(i, 2) Then
exito = False
End If
End If
Next i

For i = 1 To nodatos
For j = 1 To nocols
matpivot(i, j) = a(Matdos(i, 1), j)
Next j
Next i
SuperOrden = matpivot
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function DNormal(ByVal x As Double, ByVal med As Double, ByVal Desv As Double, ByVal Acum As Boolean) As Double
Dim Cnst As Double
Dim w As Double
Dim t As Double
Dim s As Double
Dim i As Integer

'Esta funcion de distribucion normal
'esta construida a partir de 2 formulas de aproximación
'a la distribucion normal
Cnst = Sqr(2 * Pi)
'primero se normaliza la variable
If Desv <> 0 Then
w = (CDbl(x) - med) / Desv

If Acum Then
  If Abs(w) <= 5 Then
     t = 0#
     s = Abs(w)
     For i = 1 To 40
     t = t + s / (2# * CDbl(i - 1) + 1#)
     s = s * (-1) * w * w / (2 * CDbl(i))
     Next i
     t = t / Cnst
     t = 0.5 + t * Sgn(w)
  Else
     t = 0#
     s = 1# / Abs(CDbl(w))
     For i = 0 To 27
     t = t + s
     s = s * (-1) * (2# * CDbl(i) + 1#) / (w * w)
     Next i
     t = (Exponen(-w * w / 2#) / Cnst) * t
     If w >= 0 Then t = 1# - t
  End If
  DNormal = t
Else
  DNormal = 1 / (Cnst * Desv) * Exponen(-w ^ 2 / 2)
End If
Else
    DNormal = 0
End If

End Function

Function NormalInv(ByVal x As Double) As Double
Dim s As Double
Dim t As Double
Dim difer As Double

'solo que si x es menor que .0001 o mayor que .0000 este metodo no sirve
If x > 0.000001 And x <= 0.99999 Then
difer = 0.0000001
s = 0
Do
t = s
s = s - (DNormal(s, 0, 1, 1) - x) / DNormal(s, 0, 1, 0)
Loop Until Abs(t - s) < difer
NormalInv = s
ElseIf x <= 0.00001 Then
 NormalInv = -10
ElseIf x > 0.99999 Then
 NormalInv = 10
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenerarIndFactRiesgo(ByRef matfr() As Variant, ByRef matdc() As propNodosFRiesgo) As Long()

Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

'compara las matrices matfactor con MatSensib y cuando encuentra una
'coincidencia indica la columna donde se dio esta.
n = UBound(matfr, 1)           'estos son los factores a los cuales es sensible el portafolio
m = UBound(matdc, 1)           'estos son todos los factores de riesgo
ReDim matc(1 To n) As Long
For i = 1 To n                 'matsensib
    For j = 1 To m                 'MatFactRiesgo total
        If matfr(i, 1) = matdc(j).indFactor Then  'ambos coinciden entonces
           matc(i) = j                    'indica correctamente el indice en MatFactRiesgo
           Exit For
        End If
    Next j
'indica la columna en la matriz de factores de riesgo donde es sensible el factor
    AvanceProc = i / n
    MensajeProc = "Creando el vector de indices " & Format(AvanceProc, "##0.00 %")
Next i
GenerarIndFactRiesgo = matc
End Function

Function GenMedias(ByRef a() As Double, ByVal tvol As Integer, ByVal lambda As Double) As Double()
Dim n As Long
Dim i As Long
Dim n0() As Double
Dim matr() As Double

'se genera el vector con los promedios
'para una matriz cuadrada con series en columnas
'si opc=0 son segun el modelo normal
'si opc=1 son segun el modelo exponencial
n = UBound(a, 1)
If tvol = 0 Then
'se generan las medias segun modelo
'normal
   ReDim n0(1 To n, 1 To 1) As Double
   For i = 1 To n
       n0(i, 1) = 1 / n
   Next i
   matr = MMult(MTranD(a), n0)
   GenMedias = matr
ElseIf tvol = 1 Then
'se generan medias modelo exponencial
   ReDim n0(1 To n, 1 To 1) As Double
   For i = 1 To n
       n0(i, 1) = lambda ^ (n - i) * (1 - lambda) / (1 - lambda ^ n)
   Next i
   matr = MMult(MTranD(a), n0)
   GenMedias = matr
End If
End Function


Function CMedia(ByRef vec() As Double) As Double
Dim n As Long
Dim i As Long
Dim s As Double

'calcula la media para un solo vector de datos
'de dimension unitaria
n = UBound(vec, 1)
s = 0
For i = 1 To n
s = s + vec(i)
Next i
CMedia = s / CDbl(n)
End Function

Function CVarianza(ByRef vec() As Double) As Double
Dim n As Long
Dim i As Long
Dim s As Double

'calcula la varianza para un vector de
'datos de 1 dimension
n = UBound(vec, 1)
For i = 1 To n
s = s + vec(i) ^ 2
Next i
CVarianza = (s - n * (CMedia(vec)) ^ 2) / (n - 1)

End Function

Function CMedia2(ByRef mat() As Double, ByVal B As Integer, ByVal opc As String) As Double
Dim i As Long
Dim n As Long
Dim s As Double
'calcula la media para un solo vector de datos
'de una matriz de 2 dimensiones
'opc=f en filas
'opc=c en columnas
 s = 0
 
 If opc = "f" Then
    n = UBound(mat, 2)
    For i = 1 To n
    s = s + mat(B, i)
    Next i
 ElseIf opc = "c" Then
    n = UBound(mat, 1)
    For i = 1 To n
    s = s + Val(mat(i, B))
    Next i
    End If
 CMedia2 = s / CDbl(n)
End Function

Function CVarianza2(ByRef mat() As Double, ByVal ncol As Integer, ByVal opc As String) As Double
Dim vmedia As Double
Dim s As Double
Dim n As Integer
Dim i As Integer

'calcula la varianza para un vector de
'datos en una matriz de 2 dimensiones
'opc=f en filas
'opc=c en columnas
 vmedia = CMedia2(mat, ncol, opc)
 s = 0
 If opc = "f" Then
 n = UBound(mat, 2)
 For i = 1 To n
 s = s + (mat(ncol, i) - vmedia) ^ 2
 Next i
 ElseIf opc = "c" Then
  n = UBound(mat, 1)
  For i = 1 To n
  s = s + (Val(mat(i, ncol)) - vmedia) ^ 2
  Next i
 End If
 CVarianza2 = s / (n - 1)
End Function

Function CVarianzap(ByRef mat() As Double, ByVal ncol As Integer, ByVal opc As String) As Double
Dim vmedia As Double
Dim s As Double
Dim n As Integer
Dim i As Integer

'calcula la varianza para un vector de
'datos en una matriz de 2 dimensiones
'opc=f en filas
'opc=c en columnas
 vmedia = CMedia2(mat, ncol, opc)
 s = 0
 If opc = "f" Then
 n = UBound(mat, 2)
 For i = 1 To n
 s = s + (mat(ncol, i) - vmedia) ^ 2
 Next i
 ElseIf opc = "c" Then
  n = UBound(mat, 1)
  For i = 1 To n
  s = s + (Val(mat(i, ncol)) - vmedia) ^ 2
  Next i
 End If
 CVarianzap = s / n
End Function


Function CVarLam(ByRef mat() As Double, ByVal ncol As Integer, ByVal fcol As String, ByVal lambda As Double)
Dim vmedia As Double
Dim noreg As Integer
Dim suma As Double
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se calcula el coeficiente de volatilidad exponencial para un vector
'en la columna/renglon ncol
If fcol = "f" Then
CVarLam = 0
ElseIf fcol = "c" Then
vmedia = 0
noreg = UBound(mat, 1)
For i = 1 To noreg
vmedia = vmedia + mat(i, ncol) * lambda ^ (noreg - i)
Next i
vmedia = vmedia * (1 - lambda) / (1 - lambda ^ noreg)
suma = 0
For i = 1 To noreg
 suma = suma + (mat(i, ncol) - vmedia) ^ 2 * lambda ^ (noreg - i)
Next i
CVarLam = suma * (1 - lambda) / (1 - lambda ^ noreg)
End If

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CVarExp(ByRef vec1() As Double, ByRef vec2() As Double, ByVal lambda As Double)
Dim vmedia1 As Double
Dim vmedia2 As Double
Dim noreg As Integer
Dim i As Integer
Dim suma As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se calcula el coeficiente de volatilidad exponencial para 2 vectores
noreg = UBound(vec1, 1)
vmedia1 = 0: vmedia2 = 0
For i = 1 To noreg
 vmedia1 = vmedia1 + vec1(i, 1) * lambda ^ (noreg - i)
 vmedia2 = vmedia2 + vec1(i, 1) * lambda ^ (noreg - i)
Next i
vmedia1 = vmedia1 * (1 - lambda) / (1 - lambda ^ noreg)
vmedia2 = vmedia2 * (1 - lambda) / (1 - lambda ^ noreg)
suma = 0
For i = 1 To noreg
 suma = suma + (vec1(i, 1) - vmedia1) * (vec2(i, 1) - vmedia2) * lambda ^ (noreg - i)
Next i
CVarExp = suma * (1 - lambda) / (1 - lambda ^ noreg)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CMatVaRExp(ByRef mata() As Double, ByVal lambda As Double)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
'calcula la matriz de covarianzas exponenciales

noreg = UBound(mata, 2)
ReDim matb(1 To noreg, 1 To noreg) As Variant
For i = 1 To noreg
For j = 1 To noreg
matb(i, j) = CVarExp(ExtVecMatD(mata, i, 0), ExtVecMatD(mata, j, 0), lambda)
Next j
Next i
CMatVaRExp = matb
End Function

Function CCurtosis2(ByRef mat() As Double, ByVal B As Integer, ByVal opc As Integer)
Dim i As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim vm As Double
Dim vds As Double
Dim s As Double
Dim s2 As Double
Dim s3 As Double
Dim s4 As Double
Dim n As Integer

'calcula la curtosis para un vector de
'especifico de una matriz de 2 dimensiones
vm = CMedia2(mat, B, opc)
vds = Sqr(CVarianza2(mat, B, opc))

s4 = 0
s3 = 0
s2 = 0
If opc = 0 Then
n = UBound(mat, 2)
For i = 1 To n
  s4 = s4 + mat(i, B) ^ 4
  s3 = s3 + mat(i, B) ^ 3
  s2 = s2 + mat(i, B) ^ 2
Next i
ElseIf opc = 1 Then
n = UBound(mat, 1)
For i = 1 To n
  s4 = s4 + mat(i, B) ^ 4
  s3 = s3 + mat(i, B) ^ 3
  s2 = s2 + mat(i, B) ^ 2
Next i
End If
s = s4 - 4 * vm * s3 + 6 * vm ^ 2 * s2 - 3 * n * vm ^ 4
If n >= 4 And vds <> 0 Then
CCurtosis2 = n * (n + 1) * s / (vds ^ 4) / ((Dias - 1) * (Dias - 2) * (Dias - 3)) - 3 * ((n - 1) ^ 2) / ((n - 1) * (n - 3))
Else
CCurtosis2 = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CCAsimetria2(ByRef mat() As Double, ByVal B As Integer, ByVal opc As Integer)
Dim i As Long
Dim vm As Double
Dim vds As Double
Dim n As Long
Dim s3 As Double
Dim s2 As Double
Dim s As Double


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se calcula el coeficiente de asimetria
vm = CMedia2(mat, B, opc)
vds = (CVarianza2(mat, B, opc)) ^ 0.5
If opc = 0 Then
n = UBound(mat, 2)
For i = 1 To n
s3 = s3 + (mat(B, i)) ^ 3
s2 = s2 + (mat(B, i)) ^ 2
Next i
ElseIf opc = 1 Then
n = UBound(mat, 1)
For i = 1 To n
s3 = s3 + (mat(i, B)) ^ 3
s2 = s2 + (mat(i, B)) ^ 2
Next i
End If
s = s3 - 3 * s2 * vm - 2 * n * (vm) ^ 3
If n >= 3 And vds <> 0 Then
CCAsimetria2 = n * s / (vds) ^ 3 / ((n - 1) * (n - 2))
Else
CCAsimetria2 = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function VMaximo(ByRef mat() As Double, ByVal B As Integer, ByVal opc As Integer)
Dim i As Long
Dim s As Double
Dim n As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'calcula el valor maximo en la matriz mat
'ya sea por filas o por renglones
'opc=0 se busca por renglones
'opc=1 se busca por columnas

'en la fila/columna b

If opc = 0 Then
 n = UBound(mat, 2)
 s = mat(B, 1)
 For i = 1 To n
 s = Maximo(s, mat(B, i))
 Next i
ElseIf opc = 1 Then
 n = UBound(mat, 1)
 s = mat(1, B)
 For i = 1 To n
 s = Maximo(s, mat(i, B))
 Next i
End If
VMaximo = s
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function VMinimo(ByRef mat() As Double, ByVal B As Integer, ByVal opc As Integer)
Dim i As Long
Dim n As Long
Dim s As Double
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'calcula el valor minimo en una tabla
'ya sea por filas o por renglones
'opc=0 se busca por renglones
'opc=1 se busca por columnas
' en la fila/columna b
If opc = 0 Then
 n = UBound(mat, 2)
 s = mat(B, 1)
 For i = 1 To n
 s = Minimo(s, mat(B, i))
 Next i
ElseIf opc = 1 Then
 n = UBound(mat, 1)
 s = mat(1, B)
 For i = 1 To n
 s = Minimo(s, mat(i, B))
 Next i
End If
VMinimo = s
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function VMaximo2(mat, ByVal B As Integer, ByVal c As Integer, ByVal opc As Integer)
Dim i As Long
Dim n As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es una modificacion mat la funcion vmaximo
'su objetivo es que al buscar el maximo en una
'matriz mat en la columna b, devuelve el valor maximo de b y el
'valor asociado en la columna c
'si opc = 0 se busca el maximo en filas
'si opc = 1 se busca el maximo en columnas
ReDim valor(1 To 2) As Variant
If opc = 0 Then
 n = UBound(mat, 2)
 valor(1) = mat(B, 1)
 For i = 1 To n
  If valor(1) < mat(B, i) Then
   valor(1) = mat(B, i)
   valor(2) = mat(c, i)
  End If
 Next i
ElseIf opc = 1 Then
 n = UBound(mat, 1)
 valor(1) = mat(1, B)
 For i = 1 To n
  If valor(1) < mat(i, B) Then
   valor(1) = mat(i, B)
   valor(2) = mat(i, c)
  End If
 Next i
End If
VMaximo2 = valor
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function VMinimo2(mat, ByVal B As Integer, ByVal c As Integer, ByVal opc As Integer)
Dim i As Long
Dim n As Long
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es una modificacion mat la funcion vminimo
'su objetivo es que al buscar el minimo en una
'matriz mat en la columna b, devuelve el valor minimo y el
'valor asociado de la columna c
'si opc = 0 se busca el minimo en columnas
'si opc = 1 se busca el minimo en filas
ReDim valor(1 To 2) As Variant
If opc = 0 Then
 n = UBound(mat, 2)
 valor(1) = mat(B, 1)
 valor(2) = mat(c, 1)
 For i = 1 To n
  If valor(1) > mat(B, i) Then
   valor(1) = mat(B, i)
   valor(2) = mat(c, i)
  End If
 Next i
ElseIf opc = 1 Then
 n = UBound(mat, 1)
 valor(1) = mat(1, B)
 valor(2) = mat(1, c)
 For i = 1 To n
  If valor(1) > mat(i, B) Then
   valor(1) = mat(i, B)
   valor(2) = mat(i, c)
  End If
 Next i
End If
VMinimo2 = valor
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function CMedia3(mat, ByVal nocf As Integer, ByVal inicio As Integer, ByVal final As Integer, ByVal opc As String)
Dim i As Long
Dim n As Long
Dim s As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'calcula la media en un vector de una
'matriz cuadrada no en funcion de los limites
'inicio o final y en funcion de si es una
'columna (1) o una fila (0)
'mat   matriz de datos
'nocf   columna o fila donde se va mat hacer la suma
'inicio   inicio de la suma
'final   fin de la suma
'opc por columnas o filas

'es una media estandar
n = final - inicio + 1
If opc = "f" Then
s = 0
For i = inicio To final
s = s + mat(nocf, i)
Next i
CMedia3 = s / n
ElseIf opc = "c" Then
s = 0
For i = inicio To final
s = s + mat(i, nocf)
Next i
CMedia3 = s / n
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CVarianza3(a, ByVal nocf As Integer, ByVal inicio As Integer, ByVal final As Integer, ByVal opc As String)
Dim i As Long
Dim n As Long
Dim s As Long
'
n = final - inicio + 1
If n > 1 Then
s = 0
If opc = "f" Then
For i = inicio To final
s = s + a(nocf, i) ^ 2
Next i
ElseIf opc = "c" Then
For i = inicio To final
s = s + a(i, nocf) ^ 2
Next i
End If
CVarianza3 = (s - n * (CMedia3(a, nocf, inicio, final, opc)) ^ 2) / (n - 1)
Else
CVarianza3 = 0
End If
End Function

Function TFutura(ByRef mat() As propCurva, ByVal p1 As Integer, ByVal pf As Integer, ByVal tinterpol As Integer) As Double
Dim tc As Double
Dim tl As Double
'esta funcion permite obtener las tasas
'a futuro
If pf <> 0 Then
tl = CalculaTasa(mat, p1 + pf, tinterpol)
tc = CalculaTasa(mat, p1, tinterpol)
 TFutura = ((1 + tl * (p1 + pf) / 360) / (1 + tc * p1 / 360) - 1) * 360 / pf
Else
 TFutura = 0
End If
End Function

Function TFuturaT(ByRef mat() As Variant, ByVal p1 As Integer, ByVal pf As Integer) As Double
Dim tc As Double
Dim tl As Double
'esta funcion permite obtener las tasas
'a futuro
If pf <> 0 Then
tl = BuscarTasaC(mat, p1 + pf)
tc = BuscarTasaC(mat, p1)
TFuturaT = ((1 + tl * (p1 + pf) / 360) / (1 + tc * p1 / 360) - 1) * 360 / pf
Else
TFuturaT = 0
End If
End Function

Function DefinirCurva(ByRef mat() As Variant, ByRef plazo() As Long, ByVal ind As Integer) As Variant()
Dim i As Long
Dim n As Long
'se crea la curva en funcion del indice i
n = UBound(plazo, 1)
ReDim mata(1 To n, 1 To 2) As Variant
For i = 1 To n
    mata(i, 1) = mat(ind + i, 1)
    mata(i, 2) = plazo(i)
Next i
DefinirCurva = mata
End Function

Function CrearCurvaNodos(ByVal texto As String, ByVal indice As Integer, ByRef mfriesgo() As Variant)
Dim indice0 As Integer
Dim j As Integer

'texto es el nombre de la curva o factor
'indice indica la fecha donde se empieza a leer si es cero es un vector
'mfriesgo1 es el vector donde se alojan los valores
'matp contiene el plazo de cada nodo
'da como resultado una matriz de 2 dimensiones con los siguientes valores
indice0 = BuscarValorArray(texto, MatResFRiesgo1, 1)
If indice0 <> 0 Then
   ReDim mata(1 To MatResFRiesgo(indice0, 2)) As New propCurva
   For j = 1 To MatResFRiesgo(indice0, 2)
       If indice0 <> 1 Then
          mata(j).valor = mfriesgo(indice, MatResFRiesgo(indice0 - 1).nonodosacum + j + 1)   'el valor del nodo
          mata(j).plazo = MatCaracFRiesgo(MatResFRiesgo(indice0 - 1).nonodosacum + j).plazo
       Else
          mata(j).valor = mfriesgo(indice, j + 1)
          mata(j).plazo = MatCaracFRiesgo(j).plazo
       End If
       mata(j).tfactor = MatResFRiesgo(indice0).tfactor                                      'tipo de factor
   Next j
Else
   ReDim mata(0 To 0) As New propCurva
   MsgBox "Falta el factor de riesgo " & texto
End If
CrearCurvaNodos = mata
End Function

Function CrearCurvaNodos1(ByVal texto As String, ByRef mfriesgo1() As Double)
Dim indice0 As Integer
Dim j As Integer

'texto es el nombre de la curva o factor
'indice indica la fecha donde se empieza a leer si es cero es un vector
'mfriesgo1 es el vector donde se alojan los valores
'matp contiene el plazo de cada nodo
'da como resultado una matriz de 2 dimensiones con los siguientes valores
indice0 = BuscarValorArray(texto, MatResFRiesgo1, 1)
If indice0 <> 0 Then
   ReDim mata(1 To MatResFRiesgo(indice0).nonodos) As New propCurva
   For j = 1 To MatResFRiesgo(indice0).nonodos
       If indice0 <> 1 Then
          mata(j).valor = mfriesgo1(MatResFRiesgo(indice0 - 1).nonodosacum + j, 1)      'el valor del nodo
          mata(j).plazo = MatCaracFRiesgo(MatResFRiesgo(indice0 - 1).nonodosacum + j).plazo
       Else
          mata(j).valor = mfriesgo1(j, 1)                                               'el valor del nodo
          mata(j).plazo = MatCaracFRiesgo(j).plazo
       End If
       mata(j).tfactor = MatResFRiesgo(indice0).tfactor                                 'tipo de factor
   Next j
Else
   ReDim mata(0 To 0) As New propCurva
   MsgBox "Falta el factor de riesgo " & texto
End If
CrearCurvaNodos1 = mata
End Function

Function ObtieneFRiesgo(ByVal texto As String, ByRef mfriesgo1() As Double)
Dim indice0 As Long
Dim j As Integer
'texto es el nombre de la curva o factor
'indice indica la fecha donde se empieza a leer si es cero es un vector
'mfriesgo1 es el vector donde se alojan los valores
'matp contiene el plazo de cada nodo

'da como resultado una matriz de 2 dimensiones con los siguientes valores
indice0 = BuscarValorArray(texto, MatResFRiesgo1, 1)
If indice0 <> 0 Then
   ObtieneFRiesgo = mfriesgo1(MatResFRiesgo(indice0 - 1).nonodosacum + 1, 1)      'el valor del nodo
Else
   ObtieneFRiesgo = 0
End If
End Function

Function LeerCurvaC(ByVal fecha As Date, ByRef txtcurva As String)
Dim indice As Integer
Dim i As Integer
Dim noreg As Integer
Dim idcurva As Integer
Dim txtcadena As String
Dim txtfecha As String
Dim matc() As String
Dim mata() As New propCurva
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
'esta rutina tiene como objetivo obtener la curva completa para valuar
'la posicion
indice = 0
For i = 1 To UBound(MatCatCurvas, 1)
    If txtcurva = MatCatCurvas(i, 2) Then
       idcurva = MatCatCurvas(i, 1)
       Exit For
    End If
Next i
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaCurvas & " WHERE FECHA = " & txtfecha & " AND IDCURVA = " & idcurva
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   txtcadena = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)
   matc = EncontrarSubCadenas(txtcadena, ",")
   rmesa.Close
Else
   ReDim matc(0 To 0)
End If

'si se encontro en el las curvas de pip se carga
'se crea la curva con los factores de la columna indicada
If UBound(matc, 1) = 12000 Then
   ReDim mata(1 To 12000)
    For i = 1 To 12000
        mata(i).valor = CDbl(matc(i))   'valor factor
        mata(i).plazo = i                               'plazo
    Next i
    LeerCurvaC = mata
Else
    ReDim mata(0 To 0)
End If
LeerCurvaC = mata
End Function

Function CrearCurvaCompleta(ByVal fecha As Date, ByRef txtcurva As String, matcurvas)
Dim indice As Integer
Dim i As Integer
Dim idcurva As Integer
Dim j As Long
'esta rutina tiene como objetivo obtener la curva completa para valuar
'la posicion
indice = 0
For i = 1 To UBound(MatCatCurvas, 1)
    If txtcurva = MatCatCurvas(i, 2) Then
       idcurva = MatCatCurvas(i, 1)
       Exit For
    End If
Next i
For i = 1 To UBound(matcurvas, 2)
    If idcurva = matcurvas(1, i) Then
       indice = i
       Exit For
    End If
Next i
'si se encontro en el las curvas de pip se carga
'se crea la curva con los factores de la columna indicada
If UBound(matcurvas, 1) = 12001 And indice <> 0 Then
   ReDim mata(1 To 12000) As New propCurva
    For j = 1 To 12000
        mata(j).valor = Val(matcurvas(j + 1, indice))   'valor factor
        mata(j).plazo = j                               'plazo
    Next j
    CrearCurvaCompleta = mata
Else
ReDim mata(0 To 0) As New propCurva
End If
CrearCurvaCompleta = mata
End Function

Function CrearCurva(ByVal fecha As Date, ByVal txtcurva As String, ByRef matcurvas() As Variant, ByRef mfriesgo1() As Double, ByVal vexacta As Boolean)
'esta rutina tiene como objetivo obtener la curva completa para valuar
'la posicion
If vexacta Then
   CrearCurva = CrearCurvaCompleta(fecha, txtcurva, matcurvas)
   If EsArrayVacio(CrearCurva) Then CrearCurva = CrearCurvaNodos1(txtcurva, mfriesgo1)
Else
   CrearCurva = CrearCurvaNodos1(txtcurva, mfriesgo1)
End If
End Function

Function CalculaTasa(ByRef matcurva() As propCurva, ByVal plazo As Integer, ByVal tinterpol As Integer) As Double
If ActivarControlErrores Then
On Error GoTo hayerror
End If
'en la columna 2 se encuentra el plazo
If tinterpol > 3 Or tinterpol < 1 Then
   MensajeProc = "No se ha definido el tipo de interpolación"
   CalculaTasa = 0
   Exit Function
ElseIf tinterpol = 1 Then
   CalculaTasa = CalcTasaIntL(matcurva, plazo)
ElseIf tinterpol = 2 Then
   CalculaTasa = CalcTasaAlamb(matcurva, plazo)
ElseIf tinterpol = 3 Then
   CalculaTasa = CalcTasaEsc(matcurva, plazo)
End If
On Error GoTo 0
Exit Function
hayerror:
MsgBox "Calculatasa " & error(Err())
End Function

Function CalcTasaIntL(ByRef matcurva() As propCurva, ByVal plazo As Integer) As Double
If ActivarControlErrores Then
 On Error GoTo hayerror
End If

Dim noreg As Long
Dim indice As Long

If Not EsArrayVacio(matcurva) Then
'  matcurva = tasa
  noreg = UBound(matcurva, 1)
  If noreg > 1 Then
    If plazo = 0 And noreg > 1 Then   'los indices tienen plazo 0 y noreg=1
       CalcTasaIntL = 0
       Exit Function
    ElseIf noreg = 12000 Then
      CalcTasaIntL = matcurva(plazo).valor
      Exit Function
    End If
    If plazo >= matcurva(noreg).plazo Then
       CalcTasaIntL = matcurva(noreg).valor   'lineal, fija la ultima tasa
       Exit Function
    ElseIf plazo < matcurva(1).plazo Then
      CalcTasaIntL = matcurva(1).valor
      Exit Function
    Else
      indice = BuscarValorInt(plazo, matcurva, 2)
      If indice = 0 Then
        CalcTasaIntL = 0
      Else
        If plazo >= matcurva(indice).plazo And plazo < matcurva(indice + 1).plazo Then
           CalcTasaIntL = FInterpol(matcurva(indice).valor, matcurva(indice).plazo, matcurva(indice + 1).valor, matcurva(indice + 1).plazo, plazo)
           Exit Function
        End If
      End If
    End If
  Else     '   no es una curva, es un indice
    CalcTasaIntL = matcurva(1).valor
 End If
Else   'no es una matriz con datos
   CalcTasaIntL = 0
End If
Exit Function
hayerror:
MsgBox "calctasaIntL " & error(Err())
End Function

Function CalcTasaAlamb(ByRef matcurva() As propCurva, ByVal plazo As Integer) As Double
If ActivarControlErrores Then
 On Error GoTo hayerror
End If

Dim noreg As Long
Dim indice As Long

If Not EsArrayVacio(matcurva) Then
'  matcurva = tasa
  noreg = UBound(matcurva, 1)
  If noreg > 1 Then
    If plazo = 0 And noreg > 1 Then   'los indices tienen plazo 0 y noreg=1
      CalcTasaAlamb = 0
      Exit Function
    ElseIf noreg = 12000 Then
      CalcTasaAlamb = matcurva(plazo).valor
      Exit Function
    End If
    If plazo >= matcurva(noreg).plazo Then
      CalcTasaAlamb = matcurva(noreg).valor   'lineal, fija la ultima tasa
      Exit Function
    ElseIf plazo < matcurva(1).plazo Then
      CalcTasaAlamb = matcurva(1).valor
      Exit Function
    Else
      indice = BuscarValorInt(plazo, matcurva, 2)
      If indice = 0 Then
        CalcTasaAlamb = 0
      Else
        If plazo >= matcurva(indice).plazo And plazo < matcurva(indice + 1).plazo Then
         CalcTasaAlamb = FInterpol(matcurva(indice).valor, matcurva(indice).plazo, matcurva(indice + 1).valor, matcurva(indice + 1).plazo, plazo)
         Exit Function
        End If
      End If
    End If
  Else     '   no es una curva, es un indice
    CalcTasaAlamb = matcurva(1).valor
 End If
Else   'no es una matriz con datos
   CalcTasaAlamb = 0
End If

On Error GoTo 0
Exit Function
hayerror:
MsgBox "calctasaalamb  " & error(Err())

End Function

Function CalcTasaEsc(ByRef matcurva() As propCurva, ByVal plazo As Integer) As Double

If ActivarControlErrores Then
 On Error GoTo hayerror
End If

Dim noreg As Long
Dim indice As Long
If Not EsArrayVacio(matcurva) Then
  noreg = UBound(matcurva, 1)
  If noreg > 1 Then
    If plazo = 0 And noreg > 1 Then   'los indices tienen plazo 0 y noreg=1
      CalcTasaEsc = 0
      Exit Function
    ElseIf noreg = 12000 Then
      CalcTasaEsc = matcurva(plazo).valor
      Exit Function
    End If
    If plazo >= matcurva(noreg).plazo Then
      CalcTasaEsc = matcurva(noreg).valor   'lineal, fija la ultima tasa
      Exit Function
    ElseIf plazo < matcurva(1).plazo Then
      CalcTasaEsc = matcurva(1).valor
      Exit Function
    Else
      indice = BuscarValorInt(plazo, matcurva, 2)
      If indice = 0 Then
         CalcTasaEsc = 0
      Else
        If plazo >= matcurva(indice).plazo And plazo < matcurva(indice + 1).plazo Then
           CalcTasaEsc = matcurva(indice).valor   'escalon, la tasa a menor plazo
           Exit Function
        End If
      End If
    End If
  Else     '   no es una curva, es un indice
    CalcTasaEsc = matcurva(1).valor
 End If
Else   'no es una matriz con datos
   CalcTasaEsc = 0
End If


On Error GoTo 0
Exit Function
hayerror:
MsgBox "CalcTasaEsc   " & error(Err())



End Function


Function BuscarTasaC(ByRef tasa() As Variant, ByVal p As Integer) As Double
Dim indice As Long

'se supone que la curva es completa,no se hace interpolacion
'calcula la tasa en funcion de una serie de puntos
'por medio de interpolacion
'0 interpolacion lineal
'1 tasa alambrada
'2 tasa a la izquierda
If IsArray(tasa) Then
indice = BuscarValorArray(p, tasa, 2)
 If indice <> 0 Then
    BuscarTasaC = tasa(indice, 1)
 Else
    BuscarTasaC = 0
 End If
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function DCalcTasa(ByRef curva() As propCurva, ByVal p As Integer, ByVal tinterpol As Integer) As Variant()
Dim n As Long
Dim i As Long
Dim Tasa1 As Double
Dim Tasa2 As Double



'que hace esta función?
'calcula que fraccion de la sensibilidad por nodo le corresponde a un nodo de
'una curva
'
If IsArray(curva) Then
n = UBound(curva, 1)
If n > 1 Then
 For i = 2 To n
 If curva(i, 1) = 0 Then curva(i, 1) = TEquiv(curva(i - 1, 1), curva(i - 1, 2), curva(i, 2))
 Next i
For i = 2 To n
   If p > curva(i - 1, 2) And p < curva(i, 2) Then
      Tasa1 = CalculaTasa(curva, curva(i - 1, 2), tinterpol)
      Tasa2 = CalculaTasa(curva, curva(i, 2), tinterpol)
      If tinterpol = 1 Then    'interpolacion lineal
       ReDim mate(1 To 2, 1 To 2) As Variant
       mate(1, 1) = (1 - (p - curva(i - 1, 2)) / (curva(i, 2) - curva(i - 1, 2))) * Tasa1
       mate(1, 2) = curva(i - 1, 2)
       mate(2, 1) = (p - curva(i - 1, 2)) / (curva(i, 2) - curva(i - 1, 2)) * Tasa2
       mate(2, 2) = curva(i, 2)
       DCalcTasa = mate
       Exit Function
     ElseIf tinterpol = 3 Then
       ReDim mate(1 To 1, 1 To 2) As Variant
       mate(1, 1) = Tasa1
       mate(1, 2) = curva(i - 1, 2)
       DCalcTasa = mate
       End If
       Exit Function
   ElseIf p = curva(i - 1, 2) Then
      Tasa1 = CalculaTasa(curva, curva(i - 1, 2), tinterpol)
      ReDim mate(1 To 1, 1 To 2) As Variant
      mate(1, 1) = Tasa1
      mate(1, 2) = curva(i - 1, 2)
      DCalcTasa = mate
      Exit Function
   ElseIf p = curva(i, 2) Then
      Tasa1 = CalculaTasa(curva, curva(i, 2), tinterpol)
      ReDim mate(1 To 1, 1 To 2) As Variant
      mate(1, 1) = Tasa1
      mate(1, 2) = curva(i, 2)
      DCalcTasa = mate
      Exit Function
   End If
Next i
If p > curva(n, 2) Then
      If tinterpol = 1 Then     'interpolacion lineal
       Tasa1 = CalculaTasa(curva, curva(n, 2), tinterpol)
       ReDim mate(1 To 1, 1 To 2) As Variant
       mate(1, 1) = Tasa1
       mate(1, 2) = curva(n, 2)
       DCalcTasa = mate
      ElseIf tinterpol = 3 Then
       Tasa1 = CalculaTasa(curva, curva(n, 2), tinterpol)
       ReDim mate(1 To 1, 1 To 2) As Variant
       mate(1, 1) = Tasa1
       mate(1, 2) = curva(n, 2)
       DCalcTasa = mate
      End If
      Exit Function
ElseIf p < curva(1, 2) Then
      Tasa1 = CalculaTasa(curva, curva(1, 2), tinterpol)
      ReDim mate(1 To 1, 1 To 2) As Variant
      mate(1, 1) = Tasa1
      mate(1, 2) = curva(1, 2)
      DCalcTasa = mate
      Exit Function
End If
Else
'solo hay un nodo en la curva de tasas
    Tasa1 = CalculaTasa(curva, curva(1, 2), tinterpol)
    ReDim mate(1 To 1, 1 To 2) As Variant
    mate(1, 1) = Tasa1
    mate(1, 2) = curva(1, 2)
    DCalcTasa = mate
    Exit Function
End If
End If
End Function

Function ConvertirVectorMatriz(a)
Dim i As Long
Dim n As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion convierte un vector de n x 1
'en una matriz de n x n con los elementos
'de a en la diagonal
n = UBound(a, 1)
ReDim B(1 To n, 1 To n) As Variant
For i = 1 To n
B(i, i) = a(i, 1)
Next i
ConvertirVectorMatriz = B
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub TablaHistograma(a, B)
Dim n As Integer
Dim i As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se listan los datos de la matriz a en
'la tabla b
'los datos de a tienen formato de
'de histograma,

n = UBound(a, 1)
B.Rows = n
B.Cols = 5
For i = 1 To n - 2
B.TextMatrix(i, 1) = Format(a(i, 1), "###,###,###,###,##0.0000")
B.TextMatrix(i, 2) = Format(a(i, 2), "###,###,###,###,##0.0000")
B.TextMatrix(i, 3) = Format(a(i, 3), "###,###,###,###,##0.0000")
B.TextMatrix(i, 4) = Format(a(i, 4), "###,###,###,###,##0.0000")
Next i
B.TextMatrix(n - 1, 3) = a(n - 1, 3)

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function GenerarFactoresMontecarlo(ByRef mvalor0() As Double, ByRef mmedias() As Double, ByRef mcholes() As Double, ByVal dxv As Integer, ByRef valeatorio() As Double, ByVal trends As Integer, ByVal tmedia As Integer)
Dim n As Long
Dim matd() As Double
Dim i As Long
'en funcion del vector mvalor0, la matriz de medias mmedias
'y la matriz de Choleski mcholes se genera una simulación
'de las tasas y precios
Dim vdesv As Double
n = UBound(mmedias, 1)
ReDim mats(1 To n, 1 To 1) As Double
'1 se genera una muestra normal

'se multiplica la matriz de choleski por
'el vector de no aleatorios
matd = MMult(MTranD(mcholes), valeatorio)
'3 se genera la simulación de factores de riesgo
For i = 1 To n
    vdesv = matd(i, 1)
    If trends = 0 Then
       If tmedia = 1 Then
          mats(i, 1) = (1 + mmedias(i, 1) * dxv + vdesv * Sqr(dxv)) * mvalor0(i, 1)
       Else
          mats(i, 1) = (1 + vdesv * Sqr(dxv)) * mvalor0(i, 1)
       End If
    ElseIf trends = 1 Then
       If tmedia = 1 Then
          mats(i, 1) = Exponen(mmedias(i, 1) * dxv + vdesv * Sqr(dxv)) * mvalor0(i, 1)
       Else
          mats(i, 1) = Exponen(vdesv * Sqr(dxv)) * mvalor0(i, 1)
       End If
    End If
Next i
GenerarFactoresMontecarlo = mats
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function TAlamb(tc, pc, tl, pl, p0)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
TAlamb = ((1 + tc * pc / 360) * ((1 + tl * pl / 360) / (1 + tc * pc / 360)) ^ ((p0 - pc) / (pl - pc)) - 1) * 360 / p0
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function FInterpol(ByVal tc As Double, ByVal pc As Long, ByVal tl As Double, ByVal pl As Long, ByVal p0 As Long) As Double
 FInterpol = tc + (tl - tc) / (pl - pc) * (p0 - pc)
End Function

Function TEquiv(ByVal T1 As Double, ByVal p1 As Long, ByVal p2 As Long) As Double
If p1 <> 0 And p2 <> 0 Then
TEquiv = ((1 + T1 * p1 / 360) ^ (p2 / p1) - 1) * 360 / p2
Else
TEquiv = 0
End If
End Function

Function Minimo(x, Y)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Minimo = (x + Y - Abs(x - Y)) / 2
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function Maximo(x, Y)
If x < Y Then
  Maximo = Y
Else
  Maximo = x
End If
End Function

Function DevLimitesVaR(ByVal fecha As Date, ByRef mat() As Variant, ByVal txtcampo As String) As Double
Dim n As Integer
Dim naño As Integer
Dim nmes As Integer
Dim i As Integer

'esta rutina devuelve el capital tabla de para
'una fecha corespondiente
'4 es capital tabla
'5 es capital neto
n = UBound(mat, 1)
DevLimitesVaR = 0
naño = Year(fecha)
nmes = Month(fecha)
For i = 1 To n
If (mat(i, 2) <= fecha And fecha < mat(i, 3)) And mat(i, 4) = txtcampo Then
 DevLimitesVaR = mat(i, 5)
 Exit Function
End If
Next i
If DevLimitesVaR = 0 Then
'MsgBox "No hay limites definidos para esta fecha para " & txtcampo
'MsgBox "No hay " & txtcampo & " definido para esta fecha"
End If
End Function

Function DevFechaLimite(ByVal fecha As Date, ByRef mat() As Variant, ByVal txtcampo As String) As Double
Dim n As Integer
Dim naño As Integer
Dim nmes As Integer
Dim i As Integer
'esta rutina devuelve el capital tabla de para
'una fecha corespondiente
'4 es capital tabla
'5 es capital neto
n = UBound(mat, 1)
DevFechaLimite = 0
naño = Year(fecha)
nmes = Month(fecha)
For i = 1 To n
    If (mat(i, 2) <= fecha And fecha < mat(i, 3)) And mat(i, 4) = txtcampo Then
       DevFechaLimite = mat(i, 1)
       Exit Function
    End If
Next i
If DevFechaLimite = 0 Then
   MsgBox "No hay limites definidos para esta fecha para " & txtcampo
   MsgBox "No hay " & txtcampo & " definido para esta fecha"
End If
End Function

Function IncFechaMes(ByVal fecha As Date, ByVal n As Integer)
Dim naño As Integer
Dim nmes As Integer
Dim caño As Integer
Dim cmes As Integer
Dim tfecha As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'incrementa la fecha en varios meses
naño = Year(fecha)
nmes = Month(fecha)
caño = naño + Int((nmes + n - 1) / 12)
If nmes + n - 1 >= 0 Then
 cmes = (nmes + n - 1) Mod 12 + 1
Else
 cmes = 12 - ((nmes + n - 1) Mod 12 + 3)
End If
tfecha = "01/" & cmes & "/" & caño
IncFechaMes = CDate(tfecha)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function QuitarCaracter(ByVal txtcadena As String, ByVal caracter As String) As String
Dim texto As String
Dim largo As Long
Dim i As Long

'se devuelve la misma cadena sin que se presente
'"caracter" en esta
texto = ""
largo = Len(txtcadena)
For i = 1 To largo
    If UCase(Mid(txtcadena, i, 1)) <> UCase(Left(caracter, 1)) Then
       texto = texto & Mid(txtcadena, i, 1)
    End If
Next i
QuitarCaracter = texto
End Function

Function LunesSemana(ByVal fechai As Date)
'esta funcion me debe de dar el primer dia Lunes de la
'semana
LunesSemana = fechai + 2 - Weekday(fechai)
End Function

Function ViernesSemana(ByVal fechai As Date)
'esta funcionme debe de dar el ultimo dia laborable
'de la semana inglesa
ViernesSemana = fechai + 6 - Weekday(fechai)
End Function

Function ReemplazaCadenaTexto(ByVal texto As String, ByVal texto1 As String, ByVal texto2 As String) As String
Dim cont As Long
Dim largo As Long
Dim largo1 As Long
Dim largo2 As Long
Dim cadenal As String
Dim cadenar As String
Dim textos As String

'busca el texto1 en texto y lo reemplaza por texto2
cont = 1
textos = texto
largo = Len(textos)
largo1 = Len(texto1)
largo2 = Len(texto2)
Do While cont <= largo
If Mid(textos, cont, largo1) = texto1 Then
   If cont > 1 Then
      cadenal = Mid(textos, 1, cont - 1) & texto2
   Else
      cadenal = texto2
   End If
   cadenar = Mid(textos, cont + largo1, largo - largo1 - cont + 1)
   textos = cadenal & cadenar
   cont = Len(cadenal)
   largo = Len(textos)
End If
cont = cont + 1
Loop
ReemplazaCadenaTexto = textos
End Function

Function TextBold(ByVal txt As String) As String
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
TextBold = "<B>" & txt & "</B>"
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function DIPAB(ByVal vn As Double, ByVal tc0 As Double, ByVal tc As Double, ByVal st As Double, ByVal dxv As Integer, ByVal pc As Integer, ByVal tasast As Integer) As Double
Dim nc As Integer
Dim dvc As Integer
Dim fac0 As Double
Dim df As Double
Dim fac1 As Double
Dim fac2 As Double
Dim fac3 As Double
Dim fac4 As Double
Dim fac5 As Double

'funcion para el cálculo de la DERIVADA de un IPAB
'OJO DERIVADA, NO CONFUNDIR CON DURACION
'revisar esta formula
'en funcion de opc dara la duracion por
'tasa o sobretasa
'opc=0 duracion por tasa
'opc=1 duracion por sobretasa
If tc + st <> 0 And pc <> 0 Then
 nc = CupVen(dxv, pc)
 dvc = pc * nc - dxv
 fac0 = 1 + (tc + st) * pc / 360
 df = pc / 360
  fac1 = (nc - dxv / pc - 1) * tc0 * pc / 360 * fac0 ^ (nc - dxv / pc - 2)
  fac2 = tc * (tc + st) ^ (-1) * ((nc - dxv / pc - 1) * fac0 ^ (nc - dxv / pc - 2) + dxv / pc * fac0 ^ (-dxv / pc - 1))
  fac3 = -dxv / pc * fac0 ^ (-dxv / pc - 1)
  fac4 = st * (tc + st) ^ (-2) * (fac0 ^ (nc - dxv / pc - 1) - fac0 ^ (-dxv / pc))
  fac5 = -tc * (tc + st) ^ (-2) * (fac0 ^ (nc - dxv / pc - 1) - fac0 ^ (-dxv / pc))
 If tasast = 0 Then
  DIPAB = vn * ((fac1 + fac2 + fac3) * df + fac4) * pc / 360
 ElseIf tasast = 1 Then
  DIPAB = vn * ((fac1 + fac2 + fac3) * df + fac5) * pc / 360
 End If
Else
DIPAB = 0
End If
End Function

Function DurBonoCurva(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal tc As Double, ByVal pc As Integer, ByRef curva() As propCurva, ByVal tinterpol As Integer) As Double
Dim precio As Double
Dim i As Integer
Dim indice As Integer
Dim dvc As Integer
Dim dtc As Integer
Dim tdesc As Double
Dim valor As Double


'la duracion en terminos practicos es como el centro de gravedad
'en este caso se busca el plazo ponderado promedio de los bonos
precio = PBonoCurva(fecha, tc, pc, 0, flujos, curva, tinterpol)
For i = 1 To UBound(flujos, 1)
  If flujos(i).finicio <= fecha And fecha < flujos(i).ffin Then
    indice = i
    Exit For
  End If
Next i
If indice <> 0 Then
  valor = 0
  For i = indice To UBound(flujos, 1)
     dvc = flujos(i).ffin - fecha
     dtc = fecha - flujos(i).finicio
     tdesc = CalculaTasa(curva, dvc, tinterpol)
     valor = valor + dvc / 360 * (flujos(i).saldo * tc * pc / 360 + flujos(i).amort) / (1 + tdesc * dvc / 360)
  Next i
  If precio <> 0 Then
    valor = valor / precio
    DurBonoCurva = valor
  Else
    DurBonoCurva = 0
  End If
Else
    DurBonoCurva = 0
End If
End Function

Function DurBonoY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc As Double, ByVal pc As Integer, ByVal yield As Double) As Double
Dim precio As Double
Dim i As Integer
Dim indice As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim valor As Double

'la duracion en terminos practicos es como el centro de gravedad
'en este caso se busca el plazo ponderado promedio de los bonos
precio = PBonoYield(fecha, matfl, tc, yield, pc, 0, "", 1, 0)
If fecha < matfl(1).finicio Then
   indice = 1
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
  For i = indice To UBound(matfl, 1)
     dvc = matfl(i).ffin - fecha
     dtc = fecha - matfl(i).finicio
     If pc <> 0 Then
        valor = valor + dvc / 360 * (matfl(i).saldo * tc * pc / 360 + matfl(i).amort) / (1 + yield * pc / 360) ^ (dvc / pc)
     Else
        valor = valor + 0
     End If
  Next i
  If precio <> 0 Then
    valor = valor / precio
    DurBonoY = valor
  Else
    DurBonoY = 0
  End If
Else
    DurBonoY = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function PIPABV1(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal yield As Double, ByVal st As Double, ByVal pc0 As Integer) As Double
Dim i As Integer
Dim indice As Integer
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim saldo0 As Double
Dim amort0 As Double
Dim saldo As Double
Dim amort As Double
Dim suma As Double

'If yield = 0 Then MsgBox "PIPABV1. La tasa de descuento es nula"
'saldo valor nominal
'tc0 tasa cupon actual
'yield tasa cupon vigente
'st sobretasa para valuar el IPAB
'dxv dias de vencimiento del IPAB
'pc periodo cupon del IPAB
'yield = Round(yield, decim)
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   If (yield + st) <> 0 Then
      fact0 = 1 + (yield + st) * pc0 / 360
      If indice <> 0 Then
         dvc = matfl(indice).ffin - fecha
         dtc = fecha - matfl(indice).finicio
         pc = matfl(indice).ffin - matfl(indice).finicio
         saldo0 = matfl(indice).saldo
         amort0 = matfl(indice).amort
         suma = (saldo0 * tc0 * pc / 360 + amort0) * fact0 ^ (-dvc / pc0)  'el primer cupon
      End If
      For i = indice + 1 To UBound(matfl, 1)
          dvc = matfl(i).ffin - fecha
          pc = matfl(i).ffin - matfl(i).finicio
          saldo = matfl(i).saldo
          amort = matfl(i).amort
          suma = suma + (saldo * yield * pc / 360 + amort) * fact0 ^ (-dvc / pc0)
      Next i
      PIPABV1 = suma
  End If
Else
  PIPABV1 = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
PIPABV1 = 0
'MsgBox Error(Err())
On Error GoTo 0
End Function

Function PIPABYield(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal yield As Double, ByVal pc0 As Integer) As Double
Dim i As Integer
Dim suma As Double
Dim dvc As Integer
Dim dtc As Integer
Dim pc As Integer
Dim vn0 As Double
Dim amort0 As Double
Dim vn As Double
Dim amort As Double
Dim fact0 As Double
Dim indice As Integer


If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If

If indice <= UBound(matfl, 1) Then
   If yield <> 0 Then
      fact0 = 1 + yield * pc0 / 360
      If indice <> 0 Then
         dvc = matfl(indice).ffin - fecha
         dtc = fecha - matfl(indice).finicio
         pc = matfl(indice).ffin - matfl(indice).finicio
         vn0 = matfl(indice).saldo
         amort0 = matfl(indice).amort
         suma = (vn0 * tc0 * pc / 360 + amort0) * fact0 ^ (-dvc / pc0) 'el primer cupon
      End If
      For i = indice + 1 To UBound(matfl, 1)
          dvc = matfl(i).ffin - fecha
          pc = matfl(i).ffin - matfl(i).finicio
          vn = matfl(i).saldo
          amort = matfl(i).amort
          suma = suma + (vn * tc * pc / 360 + amort) * fact0 ^ (-dvc / pc0)
      Next i
      PIPABYield = suma
   End If
Else
   PIPABYield = 0
End If

End Function

Function DurIPAB(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal st As Double, ByVal pc0 As Integer) As Double
Dim precio As Double
Dim i As Integer
Dim indice As Integer
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim vn0 As Double
Dim amort0 As Double
Dim vn As Double
Dim amort As Double
Dim suma As Double

'If tc = 0 Then MsgBox "PIPABV1. La tasa de descuento es nula"
'vn valor nominal
'tc0 tasa cupon actual
'tc tasa cupon vigente
'st sobretasa para valuar el IPAB
'dxv dias de vencimiento del IPAB
'pc periodo cupon del IPAB
'tc = Round(tc, decim)
precio = PIPABV1(fecha, matfl, tc0, tc, st, pc0)
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   If (tc + st) <> 0 Then
      fact0 = 1 + (tc + st) * pc0 / 360
      If indice <> 0 Then
         dvc = matfl(indice).ffin - fecha
         dtc = fecha - matfl(indice).finicio
         pc = matfl(indice).ffin - matfl(indice).finicio
         vn0 = matfl(indice).saldo
         amort0 = matfl(indice).amort
         suma = dvc / 360 * (vn0 * tc0 * pc / 360 + amort0) * fact0 ^ (-dvc / pc0) 'el primer cupon
      End If
      For i = indice + 1 To UBound(matfl, 1)
          dvc = matfl(i).ffin - fecha
          pc = matfl(i).ffin - matfl(i).finicio
          vn = matfl(i).saldo
          amort = matfl(i).amort
          suma = suma + dvc / 360 * (vn * tc * pc / 360 + amort) * fact0 ^ (-dvc / pc0)
      Next i
   DurIPAB = suma / precio
  End If
 Else
 'si pc es 0 se debe de obtener el limite cuando pc->0
    DurIPAB = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
DurIPAB = 0
'MsgBox Error(Err())
On Error GoTo 0
End Function

Function DurIPABY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal yield As Double, ByVal pc0 As Integer) As Double
Dim i As Integer
Dim precio As Double
Dim indice As Integer
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim vn0 As Double
Dim amort0 As Double
Dim vn As Double
Dim amort As Double
Dim suma As Double

precio = PIPABYield(fecha, matfl, tc0, tc, yield, pc0)
For i = 1 To UBound(matfl, 1)
 If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
  indice = i
  Exit For
 End If
Next i
If indice <> 0 Then
   fact0 = 1 + yield * pc0 / 360
   dvc = matfl(indice).ffin - fecha
   dtc = fecha - matfl(indice).finicio
   pc = matfl(indice).ffin - matfl(indice).finicio
   vn0 = matfl(indice).saldo
   amort0 = matfl(indice).amort
   suma = dvc / 360 * (vn0 * tc0 * pc / 360 + amort0) * fact0 ^ (-dvc / pc0) 'el primer cupon
   For i = indice + 1 To UBound(matfl, 1)
       dvc = matfl(i).ffin - fecha
       pc = matfl(i).ffin - matfl(i).finicio
       vn = matfl(i).saldo
       amort = matfl(i).amort
       suma = suma + dvc / 360 * (vn * tc * pc / 360 + amort) * fact0 ^ (-dvc / pc0)
   Next i
   If precio <> 0 Then
   DurIPABY = suma / precio
   Else
   DurIPABY = 0
   End If
 Else
 'si pc es 0 se debe de obtener el limite cuando pc->0
    DurIPABY = 0
End If
End Function

Function IDevIPAB(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal pc0 As Integer) As Double
Dim indice As Integer
Dim i As Integer
Dim dtc As Integer

indice = 0
For i = 1 To UBound(matfl, 1)
 If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
   indice = i
   dtc = fecha - matfl(i).finicio
   IDevIPAB = matfl(i).saldo * tc0 * dtc / 360
  Exit Function
 End If
Next i
If indice = 0 Then
   IDevIPAB = 0
End If
End Function

Function IDevIPABY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal pc0 As Integer)
Dim i As Long
Dim indice As Long
Dim dtc As Integer

'calculo de los intereses devengados en el periodo
'If tref = 0 Then MsgBox "PIPABV1. La tasa de descuento es nula"
'vn valor nominal
'tc0 tasa cupon actual
'tref tasa cupon vigente
'st sobretasa para valuar el IPAB
'dxv dias de vencimiento del IPAB
'pc periodo cupon del IPAB
indice = 0
For i = 1 To UBound(matfl, 1)
 If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
   indice = i
   dtc = fecha - matfl(i).finicio
   IDevIPABY = matfl(i).saldo * tc0 * dtc / 360
  Exit Function
 End If
Next i
If indice = 0 Then IDevIPABY = 0
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function PIPABV2(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal yield As Double, ByVal st As Double, ByVal pc0 As Integer) As Double
Dim i As Integer
Dim indice As Integer
Dim fact0 As Double
Dim dvc As Integer
Dim dtc As Integer
Dim pc As Integer
Dim vn0 As Double
Dim amort0 As Double
Dim suma As Double
Dim vn As Double
Dim amort As Double

If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
 If (yield + st) <> 0 Then
     fact0 = 1 + (yield + st) * pc0 / 360
    If indice <> 0 Then
       dvc = matfl(indice).ffin - fecha
       dtc = fecha - matfl(indice).finicio
       pc = matfl(indice).ffin - matfl(indice).finicio
       vn0 = matfl(indice).saldo
       amort0 = matfl(indice).amort
       suma = (vn0 * tc0 * pc / 360 + amort0) * fact0 ^ (-dvc / pc0) 'el primer cupon
   End If
   For i = indice + 1 To UBound(matfl, 1)
    dvc = matfl(i).ffin - fecha
    pc = matfl(i).ffin - matfl(i).finicio
    vn = matfl(i).saldo
    amort = matfl(i).amort
    suma = suma + (vn * yield * pc / 360 + amort) * fact0 ^ (-dvc / pc0)
   Next i
   PIPABV2 = suma
  End If
Else
 'si pc es 0 se debe de obtener el limite cuando pc->0
   MsgBox "esta fuera del rango"
   PIPABV2 = 0
End If
End Function

Function PBonoStCY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal stc As Double, ByVal yield As Double, ByVal pc0 As Integer, ByVal sicp1 As Boolean, ByVal calif As Integer, ByVal escala As String, ByVal recupera As Double, ByVal tfuncion As Integer)
Dim i As Long
Dim indice As Long
Dim tdt As Double
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim saldo As Double
Dim amort As Double
Dim suma As Double
Dim valor As Integer
Dim factor As Double

'funcion de pagares con redondeo
'vn valor nominal
'tc0 tasa cupon actual
'tc tasa cupon vigente
'stc sobretasa para valuar el bono
'dxv dias de vencimiento del bono
'pc periodo cupon del IPAB

If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   fact0 = 1 + yield * pc0 / 360
   If indice <> 0 Then
      dvc = matfl(indice).ffin - fecha
      factor = CalcFactorPE(dvc, calif, escala, recupera, tfuncion)
      dtc = fecha - matfl(indice).finicio
      pc = matfl(indice).ffin - matfl(indice).finicio
      saldo = matfl(indice).saldo
      amort = matfl(indice).amort
      suma = (amort + saldo * tc0 * pc / 360) * fact0 ^ (-dvc / pc0) * factor 'el primer cupon
   End If
   
   If Not sicp1 Then
      valor = 1
   Else
      valor = 0
   End If
   For i = indice + 1 To UBound(matfl, 1)
       dvc = matfl(i).ffin - fecha
       factor = CalcFactorPE(dvc, calif, escala, recupera, tfuncion)
       pc = matfl(i).ffin - matfl(i).finicio
       saldo = matfl(i).saldo
       amort = matfl(i).amort
       suma = suma + (amort * valor + saldo * (tc * valor + stc) * pc / 360) * fact0 ^ (-dvc / pc0) * factor
   Next i
   PBonoStCY = suma
 Else
 'si pc es 0 se debe de obtener el limite cuando pc->0
  PBonoStCY = 0
End If
End Function

Function PBonoSt(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal stc As Double, ByRef curva() As propCurva, ByVal pc0 As Integer, ByVal sicp1 As Boolean)
Dim i As Long
Dim indice As Long
Dim tdt As Double
Dim yield As Double
Dim tasa As Double
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim saldo As Double
Dim amort As Double
Dim suma As Double
Dim valor As Integer
'funcion de pagares con redondeo
'vn valor nominal
'tc0 tasa cupon actual
'tc tasa cupon vigente
'stc sobretasa para valuar el bono
'dxv dias de vencimiento del bono
'pc periodo cupon del IPAB

If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   If indice <> 0 Then
      dvc = matfl(indice).ffin - fecha
      dtc = fecha - matfl(indice).finicio
      pc = matfl(indice).ffin - matfl(indice).finicio
      tasa = CalculaTasa(curva, dvc, 1)
      yield = ((1 + tasa * dvc / 360) ^ (pc0 / dvc) - 1) * 360 / pc0
      fact0 = 1 + (yield + stc) * pc0 / 360
      saldo = matfl(indice).saldo
      amort = matfl(indice).amort
      suma = (amort + saldo * tc0 * pc / 360) * fact0 ^ (-dvc / pc0) 'el primer cupon
   End If
  
   If Not sicp1 Then
      valor = 1
   Else
      valor = 0
   End If
   For i = indice + 1 To UBound(matfl, 1)
       dvc = matfl(i).ffin - fecha
       pc = matfl(i).ffin - matfl(i).finicio
       tasa = CalculaTasa(curva, dvc, 1)
       yield = ((1 + tasa * dvc / 360) ^ (pc0 / dvc) - 1) * 360 / pc0
       fact0 = 1 + (yield + stc) * pc0 / 360
       saldo = matfl(i).saldo
       amort = matfl(i).amort
       suma = suma + (amort * valor + saldo * (tc * valor + stc) * pc / 360) * fact0 ^ (-dvc / pc0)
   Next i
   PBonoSt = suma
Else
    PBonoSt = 0
End If
End Function



Function PBonoStCY2(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal stc As Double, ByVal td As Double, ByVal pc0 As Integer)
Dim i As Integer
Dim indice As Integer
Dim pc As Integer
Dim dvc As Integer
Dim dtc As Integer
Dim fact0 As Double
Dim saldo As Double
Dim amort As Double
Dim suma As Double

'funcion de pagares con redondeo
'tc0 tasa cupon actual
'tc tasa cupon vigente
'st sobretasa para valuar el IPAB
'dxv dias de vencimiento del IPAB
'pc periodo cupon del IPAB
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
 If td <> 0 Then
   fact0 = 1 + td * pc0 / 360
   If indice <> 0 Then
      dvc = matfl(indice).ffin - fecha
      dtc = fecha - matfl(indice).finicio
      pc = matfl(indice).ffin - matfl(indice).finicio
      saldo = matfl(indice).saldo
      amort = matfl(indice).amort
      suma = (amort + saldo * tc0 * pc / 360) * fact0 ^ (-dvc / pc0) 'el primer cupon
   End If
   For i = indice + 1 To UBound(matfl, 1)
       dvc = matfl(i).ffin - fecha
       pc = matfl(i).ffin - matfl(i).finicio
       saldo = matfl(i).saldo
       amort = matfl(i).amort
       suma = suma + (amort + saldo * (tc + stc) * pc / 360) * fact0 ^ (-dvc / pc0)
   Next i
   PBonoStCY2 = suma
  End If
 Else
  PBonoStCY2 = 0
End If
End Function

Function DurBonoStCY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal yield As Double, ByVal stc As Double, ByVal td As Double, ByVal pc0 As Integer)
Dim i As Long
Dim indice As Long
Dim precio As Double
Dim fact0 As Double
Dim dtc As Integer
Dim dvc As Integer
Dim pc As Integer
Dim saldo As Double
Dim amort As Double
Dim suma As Double

'funcion de pagares con redondeo
'vn valor nominal
'tc0 tasa cupon actual
'yield tasa cupon vigente
'stc sobretasa para valuar el bono
'dxv dias de vencimiento del bono
'pc periodo cupon del IPAB
precio = PBonoStCY(fecha, matfl, tc0, yield, stc, td, pc0, False, 0, "", 0, 0)
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   fact0 = 1 + td * pc0 / 360
   If indice <> 0 Then
      dvc = matfl(indice).ffin - fecha
      dtc = fecha - matfl(indice).finicio
      pc = matfl(indice).ffin - matfl(indice).finicio
      saldo = matfl(indice).saldo
      amort = matfl(indice).amort
      suma = dvc / 360 * (amort + saldo * tc0 * pc / 360) * fact0 ^ (-dvc / pc0) 'el primer cupon
   End If
   For i = indice + 1 To UBound(matfl, 1)
    dvc = matfl(i).ffin - fecha
    pc = matfl(i).ffin - matfl(i).finicio
    saldo = matfl(i).saldo
    amort = matfl(i).amort
    suma = suma + dvc / 360 * (amort + saldo * yield * pc / 360) * fact0 ^ (-dvc / pc0)
   Next i
   DurBonoStCY = suma / precio
 Else
 'si pc es 0 se debe de obtener el limite cuando pc->0
  DurBonoStCY = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
DurBonoStCY = 0
'MsgBox Error(Err())
On Error GoTo 0
End Function

Function PBonoCurva(ByVal fecha As Date, ByVal tc As Double, ByVal pc As Integer, ByVal pf As Integer, ByRef matfl() As estFlujosMD, ByRef curva() As propCurva, ByVal tinterpol As Integer)
Dim i As Integer
Dim indice As Integer
Dim dvc As Integer
Dim dtc As Integer
Dim valor As Double
Dim vn As Double
Dim amort As Double
Dim tcorta As Double
Dim tlarga As Double
Dim tdesc As Double

'esta funcion valua el precio de un bono a tasa fija en
'funcion de la curva del proveedor
'tprecio=1 precio limpio
If fecha < matfl(1).finicio Then
   indice = 1
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
valor = 0
If indice <= UBound(matfl, 1) Then
  For i = indice To UBound(matfl, 1)
      dvc = matfl(i).ffin - fecha
      pc = matfl(i).ffin - matfl(i).finicio
      vn = matfl(i).saldo
      amort = matfl(i).amort
      tlarga = CalculaTasa(curva, dvc + pf, tinterpol)
      tcorta = CalculaTasa(curva, pf, tinterpol)
      If dvc <> 0 Then
         tdesc = ((1 + tlarga * (dvc + pf) / 360) / (1 + tcorta * pf / 360) - 1) * 360 / dvc
      Else
         tdesc = 0
      End If
      valor = valor + (vn * tc * pc / 360 + amort) / (1 + tdesc * dvc / 360)
  Next i
Else
   valor = 0
End If
PBonoCurva = valor
End Function

Sub DNBonoCurva(ByVal fecha As Date, ByVal indice As Integer, ByVal pf As Integer, ByRef matfl() As estFlujosMD, ByVal txtcurva As String)
Dim i As Integer
Dim contar As Integer
Dim indicex As Integer
Dim dvc As Integer

'esta funcion valua el precio de un bono a tasa fija en
'funcion de la curva del proveedor
'tprecio=1 precio limpio
    contar = UBound(MatNodosFREx, 2)
    For i = 1 To UBound(matfl, 1)
        If fecha >= matfl(i, 2) And fecha < matfl(i, 3) Then
           indicex = i
           Exit For
        End If
    Next i
    If pf <> 0 Then
       contar = contar + 1
       ReDim Preserve MatNodosFREx(0 To contar)
       MatNodosFREx(contar).nomFactor = txtcurva & " " & pf
       MatNodosFREx(contar).descFactor = txtcurva
       MatNodosFREx(contar).plazo = pf
       'MatNodosFREx(contar) = txtcurva & " " & Format(pf, "000000")
       'MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva & ","
       'MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & pf & ","

    End If
    If indicex <> 0 Then
       For i = indicex To UBound(matfl, 1)
           dvc = matfl(i, 3) - fecha
           contar = contar + 1
           ReDim Preserve MatNodosFREx(0 To contar)
           MatNodosFREx(contar).nomFactor = txtcurva & " " & dvc + pf
           MatNodosFREx(contar).descFactor = txtcurva
           MatNodosFREx(contar).plazo = dvc + pf
           MatNodosFREx(contar) = txtcurva & " " & Format(dvc + pf, "000000")
           'MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva & ","
           'MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & (dvc + pf) & ","
       Next i
    End If
End Sub

Function IDevBonoCurva(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc As Double) As Double
Dim i As Integer
Dim indice As Integer
Dim dtc As Integer

'esta funcion valua el precio de un bono a tasa fija en
'funcion de la curva del proveedor
'tprecio=1 precio limpio
indice = 0
For i = 1 To UBound(matfl, 1)
If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
   indice = i
   dtc = fecha - matfl(i).finicio
   IDevBonoCurva = matfl(i).saldo * tc * dtc / 360
   Exit Function
End If
Next i
If indice = 0 Then IDevBonoCurva = 0
End Function

Function PBonoYield(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc As Double, ByVal tdesc As Double, ByVal pc As Integer, ByVal calif As Integer, ByVal escala As String, ByVal recupera As Double, ByVal tfuncion As Integer) As Double
Dim i As Integer
Dim indice As Integer
Dim dvc As Integer
Dim vn As Double
Dim amort As Double
Dim pc1 As Integer
Dim valor As Double
Dim td1 As Double
Dim factor As Double

'esta funcion valua el precio de un bono a tasa fija en
'funcion de la curva del proveedor
If fecha < matfl(1).finicio Then
   indice = 1
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
      If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
         indice = i
         Exit For
      End If
   Next i
End If
valor = 0
If indice <= UBound(matfl, 1) Then
   For i = indice To UBound(matfl, 1)
       dvc = matfl(i).ffin - fecha
       vn = matfl(i).saldo
       amort = matfl(i).amort
       pc1 = matfl(i).ffin - matfl(i).finicio
       factor = CalcFactorPE(dvc, calif, escala, recupera, tfuncion)
     'la tasa de descuento es
       If 1 + tdesc * pc / 360 > 0 Then
          If pc <> 0 Then
             td1 = ((1 + tdesc * pc / 360) ^ (dvc / pc) - 1) * 360 / dvc
          Else
             td1 = tdesc
          End If
       Else
          td1 = 0
       End If
       valor = valor + (amort + vn * tc * pc1 / 360) / (1 + td1 * dvc / 360) * factor
   Next i
  PBonoYield = valor
Else
  PBonoYield = 0
End If
End Function

Function PBonoYieldPCFijo(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc As Double, ByVal yield As Double, ByVal pc As Integer, ByVal calif As Integer, ByVal escala As String, ByVal recupera As Double, ByVal tfuncion As Integer) As Double
Dim i As Integer
Dim indice As Integer
Dim valor As Double
Dim dvc As Integer
Dim dvc1 As Integer
Dim vn As Double
Dim amort As Double
Dim diasc As Integer
Dim factor As Double

'en esta funcion se considera un periodo cupon fijo aunque los periodos efectivos sean calendarios

If fecha < matfl(1).finicio Then
   indice = 1
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
      If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
         indice = i
         Exit For
      End If
   Next i
End If
valor = 0
If indice <= UBound(matfl, 1) Then
     diasc = diasTransPCFijo(matfl(indice).finicio, fecha, 30)
     dvc1 = pc - diasc
     dvc = dvc1
     For i = indice To UBound(matfl, 1)
         vn = matfl(i).saldo
         amort = matfl(i).amort
         factor = CalcFactorPE(dvc, calif, escala, recupera, tfuncion)
     'la tasa de descuento es
         valor = valor + (amort + vn * tc * pc / 360) * 1 / (1 + yield * pc / 360) ^ (dvc / pc) * factor
         dvc = dvc + pc
     Next i
  PBonoYieldPCFijo = valor
Else
  PBonoYieldPCFijo = 0
End If
End Function

Function diasTransPCFijo(ByVal fecha1 As Date, ByVal fecha As Date, ByVal pcfijo As Integer)
Dim i As Integer
Dim noreg As Integer
Dim mes1 As Integer
Dim mes2 As Integer
Dim mesx As Integer
Dim fechax As Date
Dim suma As Integer
Dim contar As Integer
Dim fechaa As Date
Dim fechab As Date
Dim pcaplicar As Integer

If Month(fecha1) = Month(fecha) And Year(fecha1) = Year(fecha) And fecha1 <= fecha Then
   diasTransPCFijo = Minimo(fecha - fecha1, DateSerial(Year(fecha), Month(fecha), pcfijo) - fecha1)
Else
   
   contar = 0
   suma = 0
   Do While True
      fechaa = DateSerial(Year(fecha1), Month(fecha1) + contar, 1)
      fechab = DateSerial(Year(fecha1), Month(fecha1) + contar + 1, 1) - 1
      If fecha >= fechaa And fecha <= fechab Then
         suma = suma + Minimo(fecha - fechaa + 1, pcfijo)
         Exit Do
      Else
         If fecha1 >= fechaa And fecha1 <= fechab Then
              suma = suma + pcfijo - Minimo(fecha1 - fechaa + 1, pcfijo)
         Else
           suma = suma + pcfijo
         End If
      End If
      contar = contar + 1
    Loop
    diasTransPCFijo = suma
End If

End Function


Function PBondesDV1(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal intcr As Double, ByVal td As Double, ByVal st As Double, ByVal pc0 As Integer)
Dim i As Integer
Dim suma As Double
Dim indice As Double
Dim pc As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim trt As Double
Dim yield As Double
Dim r1 As Double
Dim c1 As Double
Dim r2 As Double
Dim c2 As Double

'esta funcion trunca los valores con mas decimales de precision
suma = 0
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   yield = Round((1 + (td + st) / 360) ^ pc0 - 1, 8)            'yield
   If indice <> 0 Then
      pc = matfl(indice).ffin - matfl(indice).finicio
      dtc = fecha - matfl(indice).finicio
      dvc = matfl(indice).ffin - fecha
      If dtc <> 0 Then
         trt = Round(intcr * 360 / (dtc * 100), 4)
      Else
         trt = 0
      End If
'tasa del periodo que esta corriendo
      r1 = Round(((1 + trt * dtc / 360) * (1 + td / 360) ^ (dvc) - 1) * 360 / pc, 8)
      c1 = matfl(indice).saldo * r1 * pc / 360 'el primer cupon
      suma = 0
      suma = suma + (c1 + matfl(indice).amort) * (1 + yield) ^ (-dvc / pc0)
   End If
   For i = indice + 1 To UBound(matfl, 1)
       pc = matfl(i).ffin - matfl(i).finicio
       r2 = Round(((1 + td / 360) ^ pc - 1) * 360 / pc, 8)              'tasa anual esperada para el pago de intereses 2,3,...,kk
       c2 = matfl(i).saldo * r2 * pc / 360                              'segundo cupon
       dvc = matfl(i).ffin - fecha
       suma = suma + (c2 + matfl(i).amort) * (1 + yield) ^ (-dvc / pc0)
   Next i
Else
   suma = 0
End If
PBondesDV1 = Round(suma, 8)
End Function

Function DurBondesD(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tr As Double, ByVal tm As Double, ByVal st As Double, ByVal pc0 As Integer)
Dim i As Integer
Dim precio As Double
Dim suma As Double
Dim indice As Integer
Dim pc As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim r1 As Double
Dim c1 As Double
Dim r2 As Double
Dim c2 As Double
Dim trt As Double
Dim vy As Double


'esta funcion trunca los valores con mas decimales de precision
precio = PBondesDV1(fecha, matfl, tr, tm, st, pc0)
suma = 0
If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   If indice <> 0 Then
      pc = matfl(indice).ffin - matfl(indice).finicio
     dtc = fecha - matfl(indice).finicio
     'If dtc <> 0 And tr = 0 Then MsgBox "Los intereses corridos no pueden ser cero"
     dvc = matfl(indice).ffin - fecha
     trt = tr
     vy = Round((1 + (tm + st) / 360) ^ pc0 - 1, 8) 'tasa de descuento
   'tasa del periodo que esta corriendo
     r1 = Round(((1 + trt * dtc / 360) * (1 + tm / 360) ^ (dvc) - 1) * 360 / pc, 8)
     c1 = Round(matfl(indice).saldo * r1 * pc / 360, 6) 'el primer cupon
     suma = 0
     suma = suma + dvc / 360 * (c1 + matfl(indice).amort) * Round((1 + vy) ^ (-dvc / pc0), 11)
  End If
     For i = indice + 1 To UBound(matfl, 1)
        pc = matfl(i).ffin - matfl(i).finicio
        r2 = Round(((1 + tm / 360) ^ pc - 1) * 360 / pc, 8)          'tasa anual esperada para el pago de intereses 2,3,...,kk
        c2 = Round(matfl(i).saldo * r2 * pc / 360, 6)                'segundo cupon
        dvc = matfl(i).ffin - fecha
        suma = suma + dvc / 360 * (c2 + matfl(i).amort) * Round((1 + vy) ^ (-dvc / pc0), 11)
     Next i
     suma = suma / precio
  Else
    suma = 0
 End If
 DurBondesD = Round(suma, 8)
End Function

Function IDevBondesD(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tr As Double, ByVal tm As Double, ByVal st As Double, ByVal pc0 As Integer)
'esta funcion trunca los valores con mas decimales de precision
Dim suma As Double
Dim indice As Integer
Dim i As Integer
Dim dtc As Integer

suma = 0
indice = 0
For i = 1 To UBound(matfl, 1)
   If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
      indice = i
      Exit For
   End If
Next i
If indice <> 0 Then
   dtc = fecha - matfl(indice).finicio
   IDevBondesD = matfl(indice).saldo * tr * dtc / 360
Else
    IDevBondesD = 0
End If
End Function

Function PBondesDV2(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal intcr As Double, ByVal tm As Double, ByVal st As Double, ByVal pc0 As Integer) As Double
Dim i As Integer
Dim suma As Double
Dim indice As Integer
Dim pc As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim trt As Double
Dim vy As Double
Dim r1 As Double
Dim c1 As Double
Dim r2 As Double
Dim c2 As Double

suma = 0

If fecha < matfl(1).finicio Then
   indice = 0
ElseIf fecha >= matfl(UBound(matfl, 1)).ffin Then
   indice = UBound(matfl, 1) + 1
Else
   For i = 1 To UBound(matfl, 1)
       If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
          indice = i
          Exit For
       End If
   Next i
End If
If indice <= UBound(matfl, 1) Then
   vy = (1 + (tm + st) / 360) ^ pc0 - 1  'tasa de descuento
   If indice <> 0 Then
   pc = matfl(indice).ffin - matfl(indice).finicio
   dtc = fecha - matfl(indice).finicio
   dvc = matfl(indice).ffin - fecha
   If dtc <> 0 Then
      trt = Round(intcr * 360 / (dtc * 100), 4)
   Else
      trt = 0
   End If
'tasa del periodo que esta corriendo
   r1 = ((1 + trt * dtc / 360) * (1 + tm / 360) ^ (dvc) - 1) * 360 / pc
   c1 = matfl(indice).saldo * r1 * pc / 360 'el primer cupon
   suma = 0
   suma = suma + (c1 + matfl(indice).amort) * (1 + vy) ^ (-dvc / pc0)
   End If
   For i = indice + 1 To UBound(matfl, 1)
       pc = matfl(i).ffin - matfl(i).finicio
       r2 = ((1 + tm / 360) ^ pc - 1) * 360 / pc   'tasa anual esperada para el pago de intereses 2,3,...,kk
       c2 = matfl(i).saldo * r2 * pc / 360                  'segundo cupon
       dvc = matfl(i).ffin - fecha
       suma = suma + (c2 + matfl(i).amort) * (1 + vy) ^ (-dvc / pc0)
   Next i
Else
   suma = 0
End If
PBondesDV2 = suma
On Error GoTo 0
Exit Function
ControlErrores:
PBondesDV2 = 0
MsgBox error(Err())
On Error GoTo 0
End Function

Function CalcProbDAcum(ByVal dxv As Integer, ByVal calif As Integer, ByVal escala As String)
Dim contar As Integer
If dxv > 0 And calif > 0 Then
   contar = 1
   Do While True
      If -Int(-dxv / 360) > contar - 1 And -Int(-dxv / 360) <= contar Then
         If escala = "N" Then
            CalcProbDAcum = CalcProbDefAcum(calif, mTransicionN, contar)
         ElseIf escala = "I" Then
            CalcProbDAcum = CalcProbDefAcum(calif, mTransicionI, contar)
         Else
            CalcProbDAcum = 0
         End If
         Exit Function
      End If
      contar = contar + 1
   Loop
Else
   CalcProbDAcum = 0
End If
End Function

Function CalcProbDefAcum(ByVal calif As Integer, ByRef mtran() As Double, ByVal noper As Integer) As Double
Dim mata() As Double
Dim noreg As Long
Dim i As Long
If noper > 0 And calif > 0 Then
   noreg = UBound(mtran, 1)             'no de calificaciones
   mata = MIdentidad(noreg)
   ReDim matper(1 To noper) As Double
   For i = 1 To noper
       If i = 1 Then
          mata = mtran
       Else
          mata = MMult(mata, mtran)
       End If
       matper(i) = mata(calif, noreg)
   Next i
   CalcProbDefAcum = matper(noper)
Else
   CalcProbDefAcum = 0
End If
End Function


Function PBonoC0(ByVal fecha As Date, ByVal vn As Double, ByRef curva() As propCurva, ByVal st As Double, ByVal dxv As Integer, ByVal pf As Integer, ByVal tinterpol As Integer, ByVal calif As Integer, ByVal escala As String, ByVal recupera As Double, ByVal tfuncion As Integer) As Double
Dim tcorta As Double
Dim tlarga As Double
Dim tforward As Double
Dim factor As Double
'se usa para encontrar el precio de
'un bono cupon 0 por el ajuste de
'duracion convexidad

  tlarga = CalculaTasa(curva, dxv + pf, tinterpol)
  tcorta = CalculaTasa(curva, pf, tinterpol)
  If dxv > 0 Then
     factor = CalcFactorPE(dxv, calif, escala, recupera, tfuncion)
     tforward = ((1 + tlarga * (dxv + pf) / 360) / (1 + tcorta * (pf) / 360) - 1) * 360 / dxv
     PBonoC0 = vn / (1 + (tforward + st) * dxv / 360) * factor
  Else
   PBonoC0 = vn
  End If
End Function

Function CalcFactorPE(ByVal dxv As Long, ByVal calif As Integer, ByVal escala As String, ByVal recupera As Double, ByVal tfuncion As Integer)
If tfuncion = 1 Then
    CalcFactorPE = (1 - recupera) * CalcProbDAcum(dxv, calif, escala)
Else
   CalcFactorPE = 1
End If
End Function

Function PBonoC0Y(ByVal vn As Double, ByVal yield As Double, ByVal st As Double, ByVal dxv As Integer) As Double
'se usa para encontrar el precio de
'un bono cupon 0 por el ajuste de
'duracion convexidad
  If dxv <> 0 Then
   PBonoC0Y = vn / (1 + (yield + st) * dxv / 360)
  Else
   PBonoC0Y = vn
  End If
End Function


Function DV01BonoStCY(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByVal tc0 As Double, ByVal tc As Double, ByVal stc As Double, ByVal yield As Double, ByVal pc0 As Integer) As Double
Dim precio0 As Double
Dim precio1 As Double
'funcion de pagares con redondeo
'vn valor nominal
'tc0 tasa cupon actual
'tc tasa cupon vigente
'stc sobretasa para valuar el bono
'dxv dias de vencimiento del bono
'pc periodo cupon del IPAB
precio0 = PBonoStCY(fecha, matfl, tc0, tc, stc, yield, pc0, True, 0, "", 0, 0)
precio1 = PBonoStCY(fecha, matfl, tc0, tc, stc, yield + 0.0001, pc0, True, 0, "", 0, 0)
DV01BonoStCY = precio1 - precio0
End Function

Function DV01BonoC0(ByVal vn As Double, ByRef curva() As propCurva, ByVal dxv As Integer, ByVal pf As Integer, ByVal tinterpol As Integer)
Dim tcorta As Double
Dim tlarga As Double
Dim tforward0 As Double
Dim tforward1 As Double
Dim precio0 As Double
Dim precio1 As Double
'se usa para encontrar el precio de
'un bono cupon 0 por el ajuste de
'duracion convexidad
  tlarga = CalculaTasa(curva, dxv + pf, tinterpol)
  tcorta = CalculaTasa(curva, pf, tinterpol)
  If dxv <> 0 Then
     tforward0 = ((1 + tlarga * (dxv + pf) / 360) / (1 + tcorta * (pf) / 360) - 1) * 360 / dxv
     tforward1 = tforward0 + 0.0001
     precio0 = vn / (1 + tforward0 * dxv / 360)
     precio1 = vn / (1 + tforward1 * dxv / 360)
     DV01BonoC0 = precio1 - precio0
   Else
     DV01BonoC0 = 0
  End If
End Function

Function ConvexCetes(ByVal tr As Double, ByVal dxv As Integer) As Double
ConvexCetes = 2 * (dxv / 360) ^ 2 / (1 + tr * dxv / 360) ^ 2
End Function

Function DurCetes(tr, dxv, DurMod)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If DurMod = 0 Then
DurCetes = dxv / 360
ElseIf DurMod = 1 Then
DurCetes = dxv / (360 * (1 + tr * dxv / 360))
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function PBinomial(ByVal precio As Double, ByVal valK As Double, ByVal valU As Double, ByVal valD As Double, ByVal valR As Double, ByVal fdesc As Double, ByVal nnodos As Integer, ByVal estiloOpc As String, ByVal toper As String) As Double
Dim palto As Double
Dim pbajo As Double
Dim p1 As Double
Dim p2 As Double
Dim valP As Double
Dim valQ As Double
Dim i As Integer
Dim suma As Double

'se obiene el precio de una opcion en funcion del
'precio del subyacente por medio del metodo binomial
If UCase(estiloOpc) = "A" Then
  palto = precio * valU
  pbajo = precio * valD
 If nnodos > 1 Then
  p1 = PBinomial(palto, valK, valU, valD, valR, fdesc, nnodos - 1, estiloOpc, toper)
  p2 = PBinomial(pbajo, valK, valU, valD, valR, fdesc, nnodos - 1, estiloOpc, toper)
 Else
  If UCase(toper) = "C" Then
   p1 = Maximo(palto - valK, 0)
   p2 = Maximo(pbajo - valK, 0)
  ElseIf UCase(toper) = "P" Then
   p1 = Maximo(valK - palto, 0)
   p2 = Maximo(valK - pbajo, 0)
  End If
 End If
  valP = (valR - valD) / (valU - valD)
  valQ = 1 - valP
  PBinomial = (p1 * valP + p2 * valQ) / fdesc
ElseIf UCase(estiloOpc) = "E" Then
 valP = (valR - valD) / (valU - valD)
 valQ = 1 - valP
 ReDim mata(0 To nnodos) As Double, matb(0 To nnodos) As Double
 For i = 0 To nnodos
     mata(i) = precio * valU ^ i * valD ^ (nnodos - i)
     If UCase(toper) = "C" Then
        mata(i) = Maximo(mata(i) - valK, 0)
     ElseIf UCase(toper) = "P" Then
        mata(i) = Maximo(valK - mata(i), 0)
     End If
 Next i
 suma = 0
 For i = 0 To nnodos
 matb(i) = mata(i) * Fact(CLng(nnodos)) / (Fact(CLng(i)) * Fact(nnodos - i)) * valP ^ i * valQ ^ (nnodos - i)
 suma = suma + matb(i)
 Next i
 PBinomial = suma / fdesc ^ nnodos
Else
 PBinomial = 0
End If
End Function

Function IntCorrBondesD(ByVal dxv As Integer, ByVal pc As Integer, ByRef matbrm() As Variant) As Double
Dim noreg As Integer
Dim nc As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim i As Integer
Dim producto As Double

'con la matriz matbrm se debe de calcular los intereses corridos
'del titulo
 noreg = UBound(matbrm, 1)
 nc = CupVen(dxv, pc)
 dtc = nc * pc - dxv
 dvc = pc - dtc
'esta rutina se simplifica con la multiplicacion simple de los factores de matbrm
producto = 1
For i = 1 To dtc
producto = producto * (1 + matbrm(noreg - dtc + i - 1, 2) / 360) 'se le resta 1 por ser del dia anterior
Next i
If dtc <> 0 Then
 IntCorrBondesD = (producto - 1) * 360 / dtc
Else
 IntCorrBondesD = 0
End If
End Function

Function IntCorrBondesD2(ByVal fecha As Date, ByRef matfl() As estFlujosMD, ByRef matbrm() As Variant) As Double
Dim i As Integer
Dim indice As Integer
Dim producto As Double
Dim dtc As Integer
Dim dvc As Integer

'con la matriz matbrm se debe de calcular los intereses corridos
'del titulo
'esta rutina se simplifica con la multiplicacion simple de los factores de matbrm
For i = 1 To UBound(matfl, 1)
 If fecha >= matfl(i, 1) And fecha < matfl(i, 2) Then
    dtc = fecha - matfl(i, 1)
    dvc = matfl(i, 2) - fecha
    Exit For
 End If
Next i
producto = 1
indice = BuscarValorArray(fecha, matbrm, 1)
If indice <> 0 Then
For i = 1 To dtc
    producto = producto * (1 + matbrm(indice - dtc + i - 1, 2) / 360) 'se le resta 1 por ser del dia anterior
Next i
If dtc <> 0 Then
 IntCorrBondesD2 = (producto - 1) * 100
Else
 IntCorrBondesD2 = 0
End If
Else
'MsgBox "Falta la TPFB para la fecha " & fecha
End If

End Function

Function MConvDouble(mata) As Double()
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Long
'esta funcion convierte los datos en datos de precision doble
 n = UBound(mata, 1)
 m = UBound(mata, 2)
 ReDim matb(1 To n, 1 To m) As Double
 For i = 1 To n
  For j = 1 To m
   matb(i, j) = CDbl(mata(i, j))
  Next j
 Next i
 MConvDouble = matb
End Function

Function ValFwdInd(ByVal vn As Double, ByVal pspot As Double, ByVal x As Double, ByVal dxv As Integer, ByRef curva1() As propCurva, ByRef curva2() As propCurva, ByVal tinterpol As Integer) As Double
Dim valor1 As Double
Dim valor2 As Double
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 valor1 = CalculaTasa(curva1, dxv, tinterpol)
 valor2 = CalculaTasa(curva2, dxv, tinterpol)
 ValFwdInd = vn * (pspot * (1 + valor1 * dxv / 360) - x) / (1 + valor2 * dxv / 360)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ValFwdDiv(ByVal x As Double, ByVal dxv As Integer, ByVal pf As Integer, ByRef curva1() As propCurva, ByRef curva2() As propCurva, ByRef curva3() As propCurva, tc As Double, ByVal tinterpol As Integer, ByRef resval() As Variant) As Double
ReDim resval(1 To 5) As Variant
Dim Tasa1 As Double
Dim Tasa2 As Double
Dim tdesc As Double
  Tasa1 = TasaFwdCurva(pf, dxv + pf, curva1, tinterpol)
  Tasa2 = TasaFwdCurva(pf, dxv + pf, curva2, tinterpol)
  tdesc = TasaFwdCurva(pf, dxv + pf, curva3, tinterpol)
  If dxv > 0 Then
     resval(1) = tc * (1 + Tasa1 * dxv / 360) / (1 + Tasa2 * dxv / 360) / (1 + tdesc * dxv / 360)
     resval(2) = x / (1 + tdesc * dxv / 360)
     resval(3) = Tasa1
     resval(4) = Tasa2
     resval(5) = tc
     ValFwdDiv = resval(1) - resval(2)
  Else
     ValFwdDiv = 0
  End If
  If SiAgregarDatosFwd Then Call AgregarDatosFwd(resval, MatParamFwds)
End Function

Sub AgregarDatosFwd(ByRef mrvalflujo() As Variant, ByRef mata() As Variant)
Dim contar As Long
contar = UBound(mata, 2)
contar = contar + 1
ReDim Preserve mata(1 To 5, 1 To contar) As Variant
mata(1, contar) = mrvalflujo(1)
mata(2, contar) = mrvalflujo(2)
mata(3, contar) = mrvalflujo(3)
mata(4, contar) = mrvalflujo(4)
mata(5, contar) = mrvalflujo(5)
End Sub

Function ValFwdTasa(ByVal vn As Double, ByVal tp As Double, ByVal dxv As Integer, ByVal pf As Integer, ByRef curva1() As propCurva, ByRef curva2() As propCurva, ByVal tinterpol As Integer) As Double
Dim pc As Integer
Dim pl As Integer
Dim tc As Double
Dim tl As Double
Dim tf As Double
Dim tdesc As Double

'valor de ejercicio
pc = dxv - pf
pl = dxv
tc = CalculaTasa(curva1, pc, tinterpol)
tl = CalculaTasa(curva1, pl, tinterpol)
tf = ((1 + tl * pl / 360) / (1 + tc * pc / 360) - 1) * 360 / (pl - pc)
tdesc = CalculaTasa(curva2, pl, tinterpol)    'tasa de descuento
ValFwdTasa = vn * (tp - tf) * (pf / 360) / (1 + tdesc * pl / 360)
End Function


Function BuscarTasaCupon(ByVal fecha As Date, ByVal dxv As Integer, ByVal pc As Integer, ByVal pc1 As Integer, ByVal nomcurva As String, ByVal tinterpol As Integer)
Dim mfriesgo() As Double
Dim curva() As propCurva
Dim nc As Integer
Dim dtc As Integer
Dim fechax As Date
Dim indice As Integer
Dim f_val As Date
Dim desf As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'busca el valor de la tasa en una curva
 nc = CupVen(dxv, pc)
 dtc = nc * pc - dxv
 fechax = fecha - dtc
 indice = 0
 f_val = fechax
 desf = 0
 Do While indice = 0
    indice = BuscarValorArray(f_val, MatFactRiesgo, 1)
    If indice <> 0 Then Exit Do
    f_val = f_val - 1
    desf = desf + 1
 Loop
 If indice <> 0 Then
    mfriesgo = ExtFRMatFR(indice, MatFactRiesgo)
    curva = CrearCurvaNodos1(nomcurva, mfriesgo)
    BuscarTasaCupon = TFutura(curva, desf, pc1, tinterpol)
 Else
   BuscarTasaCupon = 0
 End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function BSDOne(ByVal s As Double, ByVal x As Double, ByVal r As Double, ByVal Q As Double, ByVal tyr As Double, ByVal sigma As Double)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'   Returns the Black-Scholes d1 value
    BSDOne = (Logarit(s / x) + (r - Q + 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function
    
Function BSDTwo(ByVal s As Double, x, r, Q, tyr, sigma)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'   Returns the Black-Scholes d2 value
    BSDTwo = (Logarit(s / x) + (r - Q - 0.5 * sigma ^ 2) * tyr) / (sigma * Sqr(tyr))
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function PPNormInv(Z, n)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'   Returns the Peizer-Pratt Inversion
'   Only defined for n odd
'   Used in LR Binomial Option Valuation
    Dim c1
    n = SigImpar(n)
    c1 = Exponen(-((Z / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2) * (n + 1 / 6))
    PPNormInv = 0.5 + Sgn(Z) * Sqr((0.25 * (1 - c1)))
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function FSigno(x)
If x >= 0 Then
   FSigno = 1
ElseIf x < 0 Then
   FSigno = -1
End If
End Function

Function SigImpar(valx)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
SigImpar = -2 * Int(-(valx + 1) / 2) - 1
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub MostrarMensajeSistema(ByVal txtmsg As String, objeto, ByVal tiempo As Double, ByVal fecha As Date, ByVal hora As Date, ByVal usuario As String)
Dim tiempo0 As Double
Dim tiempo1 As Double
If TypeName(objeto) = "Label" Then
 objeto.Caption = txtmsg
 objeto.Refresh
  tiempo0 = CDbl(Date) + CDbl(Time)
' tiempo1 = CDbl(Date) + CDbl(Time)
 If tiempo <> 0 Then
 Do While CDbl(tiempo1 - tiempo0) < tiempo / 86400
  tiempo1 = CDbl(Date) + CDbl(Time)
 Loop
 End If
ElseIf TypeName(objeto) = "IPanel" Then
 objeto.Text = MensajeProc
 tiempo0 = CDbl(Date) + CDbl(Time)
 tiempo1 = CDbl(Date) + CDbl(Time)
 If tiempo <> 0 Then
 Do While CDbl(tiempo1 - tiempo0) < tiempo / 86400
  tiempo1 = CDbl(Date) + CDbl(Time)
  Loop
 End If
End If
End Sub

Sub CalcVolMuestraO(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtfr As String, ByVal ndias As Integer, ByVal txtconc As String, ByVal dxv As Integer)
Dim matvec() As Variant
Dim matvolatil1() As Double
Dim matrends1() As Double
Dim fecha As Date
Dim txtfecha As String
Dim txtcadena As String
Dim ivol As Long
Dim volatil As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'se identifica cual es el subyacente para el calculo de la volatilidad
'se va a calcular las volatilidades y a guardar en la tabla de datos
'como un factor mas
'se carga la informacion de este factor

 matvec = Leer1FactorR(fecha1 - 500, fecha2, txtfr, 0)
 If UBound(matvec, 1) > 0 Then
 fecha = fecha1
 Do While fecha <= fecha2
  ivol = 0
  ivol = BuscarValorArray(fecha, matvec, 1)
  If ivol <> 0 Then
   matvolatil1 = ExtSerieFR(matvec, 2, ivol, ndias)
   matrends1 = CalculaRendimientoColumna(matvolatil1, 1)
   volatil = (CVarianza2(matrends1, 1, "c") * 251) ^ 0.5
  Else
   volatil = 0
  End If
  If volatil <> 0 Then
     txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     ConAdo.Execute "DELETE FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtconc & "' AND FECHA = " & txtfecha & " AND PLAZO = " & dxv
     txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
     txtcadena = txtcadena & txtfecha & ",'"
     txtcadena = txtcadena & txtconc & "',"
     txtcadena = txtcadena & dxv & ","
     txtcadena = txtcadena & volatil & ","
     txtcadena = txtcadena & "'" & CLng(fecha) & Trim("txtconc") & Trim(Format(dxv, "00000000")) & "')"
   ConAdo.Execute txtcadena
   MensajeProc = "Actualizando registros de la tabla " & TablaFRiesgoO & " " & txtconc
  End If
  fecha = fecha + 1
 Loop
 Else
  MensajeProc = "No se calculo la volatilidad de " & txtfr
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function FactorConv(ByVal vn As Double, ByVal tc As Double, ByVal tr As Double, ByVal dxv As Integer, ByVal pc As Integer)
Dim nc As Integer
Dim dtc As Integer
Dim dvc As Integer
Dim rend As Double
Dim cup As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

' factor de conversion para los futuros del bono a tasa
' fija
nc = CupVen(dxv, pc)
dtc = nc * pc - dxv
dvc = pc - dtc
rend = tr * pc / 360
cup = vn * tc * pc / 360
FactorConv = ((cup + cup * (1 / rend - 1 / (rend * (1 + rend) ^ (nc - 1))) + vn / (1 + rend) ^ (nc - 1)) / (1 + rend) ^ (1 - dtc / pc) - cup * dtc / pc) / 100
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub Bootstrapp(ByRef matdatos() As Variant, ByRef matsal() As Variant, ByVal pc As Integer, ByVal pconv As Integer, ByVal nodat As Integer, ByVal tinterpol As Integer)
Dim noreg As Integer
Dim i As Integer
Dim tasa As Double


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es una variante del algoritmo de bootstrapping
'de hecho este algoritmo es mas flexible ya que los plazos no tienen que
'estar repartidas de manera uniforme en la curva
noreg = UBound(matdatos, 1)
ReDim pilatasas(0 To 0) As propCurva    'los resultados se ponen transpuestos
For i = 1 To noreg
 tasa = CalculaTasaObjetivo(100, matdatos(i, 1), matdatos(i, 2), pc, pconv, pilatasas)
 ReDim Preserve pilatasas(0 To i) As propCurva
 pilatasas(i).valor = tasa
 pilatasas(i).plazo = matdatos(i, 2)
 MensajeProc = "" & Format(AvanceProc, "#,##0.00 %")
Next i
ReDim matsal(1 To nodat, 1 To 2) As Variant
For i = 1 To nodat
 matsal(i, 1) = CalculaTasa(pilatasas, i, tinterpol)
 matsal(i, 2) = i
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub Bootstrapp1(ByVal fecha As Date, ByRef matdatos() As Variant, ByRef matsal() As propCurva, ByVal pc As Integer, ByVal nodat As Integer, ByVal tinterpol As Integer)
Dim noreg As Integer
Dim i As Integer
Dim tasa As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es una variante del algoritmo de bootstrapping
'de hecho este algoritmo es mas flexible ya que las emisiones no tienen que
'se de plazos similares o multiplos
noreg = UBound(matdatos, 1)
ReDim pilatasas(0 To 0) As propCurva     'los resultados se ponen transpuestos
For i = 1 To nodat
 tasa = CalculaTasaObjetivo(100, matdatos(i, 1), matdatos(i, 2), pc, 360, pilatasas)
 ReDim Preserve pilatasas(0 To i) As propCurva
 pilatasas(i).valor = tasa
 pilatasas(i).plazo = matdatos(i, 2)
 MensajeProc = "" & Format(AvanceProc, "#,##0.00 %")
Next i
matsal = pilatasas
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function CalculaPBono(ByVal vn As Double, ByVal tc As Double, ByVal dxv As Integer, ByVal pc As Integer, ByRef pilatasas() As propCurva, ByVal tdesc As Double, ByVal pact As Integer) As Double
Dim noreg As Integer
Dim i As Integer
Dim nc As Integer
Dim precio As Double
Dim dvc As Integer
Dim tasa As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(pilatasas, 2)

If noreg > 0 Then
ReDim mtasas(1 To noreg + 1) As propCurva
 For i = 1 To noreg
  mtasas(i).valor = pilatasas(i).valor
  mtasas(i).plazo = pilatasas(i).plazo
 Next i
 mtasas(noreg + 1).valor = tdesc
 mtasas(noreg + 1).plazo = dxv
Else
ReDim mtasas(1 To noreg + 1) As propCurva
 mtasas(noreg + 1).valor = tdesc
 mtasas(noreg + 1).plazo = dxv
End If
nc = CupVen(dxv, pc)
precio = 0
For i = 1 To nc
 dvc = dxv - pc * (nc - i)
 tasa = CalculaTasa(mtasas, dvc, 1)
 If i = 1 Then
    precio = precio + vn * tc * dvc / pact / (1 + tasa * dvc / pact)
 Else
    precio = precio + vn * tc * pc / pact / (1 + tasa * dvc / pact)
 End If
Next i
precio = precio + vn / (1 + tasa * dvc / pact)
CalculaPBono = precio
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CalculaTasaObjetivo(ByVal vn As Double, ByVal tc As Double, ByVal dxv As Integer, ByVal pc As Integer, ByVal pact As Integer, ByRef pilatasas() As propCurva) As Double
Dim td0 As Double
Dim tnueva As Double
Dim inc As Double
Dim precio As Double
Dim precio1 As Double
Dim derivada As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se utiliza el algoritmo de newton rapson para encontrar la tasa objetivo

td0 = tc
tnueva = tc
inc = 0.0000001
Do
td0 = tnueva
 precio = CalculaPBono(100, tc, dxv, pc, pilatasas, td0, pact)
 precio1 = CalculaPBono(100, tc, dxv, pc, pilatasas, td0 + inc, pact)
 derivada = (precio1 - precio) / inc
 tnueva = td0 - (precio - vn) / derivada
Loop Until Abs(tnueva - td0) < 0.000000001
CalculaTasaObjetivo = tnueva
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function PBonoEsp5(ByVal vn As Double, ByVal tc As Double, ByVal pc As Integer, ByRef tasas() As Variant) As Double
Dim noreg As Integer
Dim valor As Double
Dim i As Integer
Dim plazo As Integer
Dim tcupon As Integer


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'precio del bono en funcion de la tasa cupon
'se espera que la tasa cupon este al plazo pc
noreg = UBound(tasas, 1)
valor = 0
For i = 1 To noreg
 If i <> 1 Then
  plazo = tasas(i, 2) - tasas(i - 1, 2)
 Else
  plazo = tasas(i, 2)
 End If
 tcupon = ((1 + tc * pc / 360) ^ (plazo / pc) - 1) * 360 / plazo
 valor = valor + vn * tcupon * plazo / 360 / (1 + tasas(i, 1) * tasas(i, 2) / 360)
Next i
 valor = valor + vn / (1 + tasas(noreg, 1) * tasas(noreg, 2) / 360)
PBonoEsp5 = valor
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CrearTablaAmortiza(ByVal fecha As Date, ByVal indice As Long, ByVal fflujo As Long, ByVal intercf As String, ByVal spread As Double, ByRef matflujos() As estFlujosDeuda)
Dim i As Long

    ReDim mrvalflujo(1 To fflujo - indice + 1) As New resValFlujo
       For i = 1 To fflujo - indice + 1
           mrvalflujo(i).c_operacion = matflujos(i + indice - 1).coperacion            'clave ikos
           mrvalflujo(i).t_pata = matflujos(i + indice - 1).tpata                      'pata
           mrvalflujo(i).fecha_ini = matflujos(i + indice - 1).finicio                 'FECHA INICIO
           mrvalflujo(i).fecha_fin = matflujos(i + indice - 1).ffin                    'fecha FINAL
           mrvalflujo(i).fecha_desc = matflujos(i + indice - 1).fpago                  'fecha PAGO INT
           mrvalflujo(i).si_paga_int = matflujos(i + indice - 1).pago_int              'pag intereses
           mrvalflujo(i).int_s_saldo = matflujos(i + indice - 1).int_t_saldo           'saldo*intereses
           mrvalflujo(i).saldo_periodo = matflujos(i + indice - 1).saldo               'saldo
           If intercf = "S" Or ValEficiencia Then
              If mrvalflujo(i).fecha_desc > fecha Then
                 mrvalflujo(i).amortizacion = matflujos(i + indice - 1).amort          'amortizacion
              Else
                 mrvalflujo(i).amortizacion = 0                                        'amortizacion
              End If
           Else
              mrvalflujo(i).amortizacion = 0                                           'amortizacion
           End If
           mrvalflujo(i).t_cupon_per = matflujos(i + indice - 1).t_cupon               'tasa cupon periodo
           mrvalflujo(i).sobretasa = spread                                            'spread
           mrvalflujo(i).dxv1 = matflujos(i + indice - 1).finicio - fecha              'dias inicio cupon
           mrvalflujo(i).dxv2 = matflujos(i + indice - 1).ffin - fecha                 'dias fin cupon
           mrvalflujo(i).dxv3 = matflujos(i + indice - 1).fpago - fecha                'dias para descontar flujo
           mrvalflujo(i).p_cupon = mrvalflujo(i).dxv2 - mrvalflujo(i).dxv1             'periodo cupon
     
'se quedan 2 columna de la matriz sin datos con el fin de homologar los modulos
       Next i
    CrearTablaAmortiza = mrvalflujo
End Function

Function CrearTablaIntDev(ByVal fecha As Date, ByVal indice As Long, ByVal fflujo As Long, ByVal intercf As String, ByVal spread As Double, ByRef matflujos() As estFlujosDeuda)
Dim contar1 As Long
Dim i As Long
Dim j As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim convInt As String

 ReDim mrvalflujo(1 To fflujo - indice + 1) As New resValFlujo
 contar1 = 0
       For i = 1 To fflujo - indice + 1
           mrvalflujo(i).c_operacion = matflujos(i + indice - 1).coperacion            'clave ikos
           mrvalflujo(i).t_pata = matflujos(i + indice - 1).tpata                      'pata
           mrvalflujo(i).fecha_ini = matflujos(i + indice - 1).finicio                 'FECHA INICIO
           If fecha < mrvalflujo(i).fecha_ini Then Exit For
           contar1 = contar1 + 1
           mrvalflujo(i).fecha_fin = matflujos(i + indice - 1).ffin                    'fecha FINAL
           mrvalflujo(i).fecha_desc = matflujos(i + indice - 1).fpago                  'fecha descuento
           mrvalflujo(i).si_paga_int = matflujos(i + indice - 1).pago_int              'pag intereses
           mrvalflujo(i).int_s_saldo = matflujos(i + indice - 1).int_t_saldo           'saldo * intereses
           mrvalflujo(i).saldo_periodo = matflujos(i + indice - 1).saldo               'saldo
           If intercf = "S" Or ValEficiencia Then
              If mrvalflujo(i).fecha_fin > fecha Then
                 mrvalflujo(i).amortizacion = matflujos(i + indice - 1).amort     'amortizacion
              Else
                 mrvalflujo(i).amortizacion = 0                                'amortizacion
              End If
           Else
              mrvalflujo(i).amortizacion = 0                                   'amortizacion
           End If
           mrvalflujo(i).t_cupon_per = matflujos(i + indice - 1).t_cupon           'tasa cupon periodo
           mrvalflujo(i).sobretasa = spread                                    'spread
           mrvalflujo(i).dxv1 = matflujos(i + indice - 1).finicio - fecha           'dias inicio cupon
           mrvalflujo(i).dxv2 = matflujos(i + indice - 1).ffin - fecha           'dias fin cupon
           mrvalflujo(i).dxv3 = matflujos(i + indice - 1).fpago - fecha           'dias pago cupon
           mrvalflujo(i).p_cupon = mrvalflujo(i).dxv2 - mrvalflujo(i).dxv1

'se quedan 2 columna de la matriz sin datos con el fin de homologar los modulos
       Next i
       
  If contar1 <> 0 Then
    ReDim mrvalflujo1(1 To contar1) As New resValFlujo
    For i = 1 To contar1
        Set mrvalflujo1(i) = mrvalflujo(i)
    Next i
  
 Else
   ReDim mrvalflujo1(0 To 0) As New resValFlujo
 End If
 CrearTablaIntDev = mrvalflujo1
End Function

Function EstIndTAmortiza(ByVal fecha As Date, ByVal iflujo As Long, ByVal fflujo As Long, ByRef matflujos() As estFlujosDeuda) As Long
Dim i As Long
Dim indice As Long
Dim contar As Long

      If fecha < matflujos(iflujo).finicio Then
       indice = iflujo
      End If
      For i = iflujo To fflujo
       If matflujos(i).ffin = matflujos(i).fpago Then
          If fecha >= matflujos(i).finicio And fecha < matflujos(i).ffin Then
             indice = i
             Exit For
         End If
       ElseIf matflujos(i).ffin < matflujos(i).fpago Then
          If fecha >= matflujos(i).finicio And fecha < matflujos(i).fpago Then
             indice = i
             Exit For
         End If
       ElseIf matflujos(i).ffin > matflujos(i).fpago Then
          If fecha >= matflujos(i).finicio And fecha < matflujos(i).ffin Then
             If fecha < matflujos(i).fpago Then
                indice = i
                Exit For
             Else
                indice = i + 1
                Exit For
             End If
         End If
       End If
      Next i
      contar = indice
      For i = indice - 1 To iflujo Step -1
          If matflujos(i).pago_int = "S" Then
            Exit For
         Else
            contar = contar - 1
         End If
      Next i
      indice = contar
      If fecha >= matflujos(fflujo).ffin Then 'el valor de la pata es cero
         indice = -1
     End If


EstIndTAmortiza = indice
End Function

Function ValDeudaFija(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curva1() As propCurva, ByRef parval As paramValFlujo, ByRef mrvalflujo() As resValFlujo)
Dim pf As Integer
Dim i As Long
Dim indice As Long
Dim valor As Double
Dim fecha1 As Date
Dim fecha2 As Date


'estas es la definicion de matpar

'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento

'primero determinamos el flujos

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  parval.sobret y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata
indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then
   ReDim mrvalflujo(0 To 0) As resValFlujo
   ValDeudaFija = 0
   Exit Function
End If
mrvalflujo = CrearTablaAmortiza(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1
    If parval.pcref <> 0 Then
       mrvalflujo(i).f_tiempo_aplicar = parval.pcref / 360                                                      'fraccion de año a aplicar
    Else
       mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)                    'fraccion de año a aplicar
    End If
    mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
'determinacion del saldo a aplicar intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+parval.sobret)*parval.pcref/360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-interses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar=amortizacion+intereses pagados
   'If sive Then
   '   Prob = CalcProbDf(dxv, calif, mattran)
   '   mrvalflujo(i).pago_total = (mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo) * Prob * (1 - recupera)
   'Else
      mrvalflujo(i).pago_total = mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo
   'End If
 'tasa de descuento
   mrvalflujo(i).t_desc = TasaFwdCurva(parval.perfwd, mrvalflujo(i).dxv3 + parval.perfwd, curva1, parval.modint1)
 'factor descuento
   mrvalflujo(i).factor_desc = 1 / (1 + mrvalflujo(i).t_desc * mrvalflujo(i).dxv3 / 360)
  'valor presente =(amortizacion+intereses)*factor descuento
   mrvalflujo(i).valor_presente = mrvalflujo(i).pago_total * mrvalflujo(i).factor_desc
 'vp acumulado
   valor = valor + mrvalflujo(i).valor_presente
Next i
ValDeudaFija = valor
End Function

Function CalcProbDf(ByVal dxv As Integer, ByVal calif As Integer, ByRef mattran() As Double)
Dim i As Integer
For i = 2 To 50
    If -Int(-dxv / 360) >= i - 1 And -Int(-dxv / 360) < i Then
       CalcProbDf = CalcPDAcum(calif, mattran, i, i)
       Exit Function
   End If
Next i
End Function

Function CalcPDAcum(ByVal calif As Integer, ByRef mattran() As Double, ByVal noper As Integer, ByVal orden As Integer) As Double
Dim mata() As Double
Dim matb() As Double
Dim noreg As Long
Dim i As Long
Dim j As Integer
Dim mtran2() As Double
Dim mattran1() As Double

mattran1 = mattran
noreg = UBound(mattran1, 1)             'no de calificaciones
ReDim mtran2(1 To noreg, 1 To noreg)
For i = 1 To noreg
    For j = 1 To noreg
        mtran2(i, j) = mattran1(i, j)
    Next j
Next i
mata = MIdentidad(noreg)
matb = MIdentidad(noreg)
ReDim matper(1 To noper) As Double
For i = 1 To noper
    mata = matb
    If i = 1 Then
       matb = mtran2
    Else
       matb = MMult(mata, mtran2)
    End If
    matper(i) = matb(calif, noreg)
Next i
CalcPDAcum = matper(orden)
End Function



Function ValDeudaFijaYield(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByVal yield As Double, ByRef parval As paramValFlujo, ByRef mrvalflujo() As resValFlujo)
Dim i As Long
Dim indice As Long
Dim valor As Double
Dim fecha1 As Date
Dim fecha2 As Date
'estas es la definicion de matpar
'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento
'primero determinamos el flujos

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata
indice = EstIndTAmortiza(fecha + parval.perfwd, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then
   ReDim mrvalflujo1(0 To 0) As resValFlujo
   ValDeudaFijaYield = 0
   Exit Function
End If
mrvalflujo = CrearTablaAmortiza(fecha + parval.perfwd, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1

    If parval.pcref <> 0 Then
       mrvalflujo(i).f_tiempo_aplicar = parval.pcref / 360                                                       'fraccion de año a aplicar
    Else
       mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)                'fraccion de año a aplicar
    End If
    mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
'determinacion del saldo a aplicar intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+parval.sobret)*parval.pcref/360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-interses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar=amortizacion+intereses pagados
   mrvalflujo(i).pago_total = mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo
 'tasa de descuento
   mrvalflujo(i).t_desc = yield
 'factor descuento
   mrvalflujo(i).factor_desc = 1 / (1 + yield * mrvalflujo(i).p_cupon / 360) ^ (mrvalflujo(i).dxv3 / mrvalflujo(i).p_cupon)
  'valor presente =(amortizacion+intereses)*factor descuento
   mrvalflujo(i).valor_presente = mrvalflujo(i).pago_total * mrvalflujo(i).factor_desc
 'vp acumulado
   valor = valor + mrvalflujo(i).valor_presente
Next i
ValDeudaFijaYield = valor
End Function

Function DetYieldDeudaFija(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curva() As propCurva, ByRef parval As paramValFlujo)
Dim precio0 As Double
Dim precio1 As Double
Dim precio2 As Double
Dim precio3 As Double
Dim yield As Double
Dim yield2 As Double
Dim deriv As Double
Dim noiter As Long
Dim mrvalflujo() As resValFlujo
Dim inc As Double

precio0 = ValDeudaFija(fecha, matflujos, curva, parval, mrvalflujo)
yield = 0.05
precio1 = ValDeudaFijaYield(fecha, matflujos, yield, parval, mrvalflujo)
noiter = 0
Do While Abs(precio1 - precio0) > 0.00001 And noiter < 9900
   inc = 0.000001
   precio2 = ValDeudaFijaYield(fecha, matflujos, yield + inc, parval, mrvalflujo)
   deriv = (precio2 - precio1) / inc
   yield2 = yield - (precio1 - precio0) / deriv
   yield = yield2
   precio3 = ValDeudaFijaYield(fecha, matflujos, yield2, parval, mrvalflujo)
   noiter = noiter + 1
   precio1 = precio3
Loop
If Abs(precio1 - precio0) < 0.001 Then
   DetYieldDeudaFija = yield
Else
   DetYieldDeudaFija = 0
End If
End Function

Function DurDFijaYield(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByVal yield As Double, ByRef parval As paramValFlujo)
Dim i As Long
Dim indice As Long
Dim valor As Double
Dim mrvalflujo() As resValFlujo
Dim fecha1 As Date
Dim fecha2 As Date
'estas es la definicion de matpar


'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento

'primero determinamos el flujos

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata
indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then
   DurDFijaYield = 0
   Exit Function
End If
mrvalflujo = CrearTablaAmortiza(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1
    If parval.pcref <> 0 Then
       mrvalflujo(i).f_tiempo_aplicar = parval.pcref / 360
    Else
       mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)                'fraccion de año a aplicar
    End If
    mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
'determinacion del saldo a aplicar intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+spread)*parval.pcref/360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-intereses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total ajustador por plazo = (amortizacion+intereses pagados) * plazo
   mrvalflujo(i).pago_total = (mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo) * mrvalflujo(i).dxv3
 'tasa de descuento
 'factor descuento
   mrvalflujo(i).factor_desc = 1 / (1 + yield * mrvalflujo(i).p_cupon / 360) ^ (mrvalflujo(i).dxv3 / mrvalflujo(i).p_cupon)
  'valor presente =(amortizacion+intereses)*factor descuento
   mrvalflujo(i).valor_presente = mrvalflujo(i).pago_total * mrvalflujo(i).factor_desc
 'vp acumulado
   valor = valor + mrvalflujo(i).valor_presente
Next i
DurDFijaYield = valor
End Function

Function DurDeudaFija(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curva1() As propCurva, ByRef parval As paramValFlujo)
Dim i As Long
Dim indice As Long
Dim valor As Double
Dim mrvalflujo() As resValFlujo
Dim fecha1 As Date
Dim fecha2 As Date


'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento

'primero determinamos el flujos

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata
indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then
   DurDeudaFija = 0
   Exit Function
End If
mrvalflujo = CrearTablaAmortiza(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1
    If parval.pcref <> 0 Then
       mrvalflujo(i).f_tiempo_aplicar = parval.pcref / 360
    Else
       mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)                'fraccion de año a aplicar
    End If
    mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
'determinacion del saldo a aplicar intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+parval.sobret)*parval.pcref/360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-interses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar afectado por plazo=(amortizacion+intereses pagados)*dias por vencer
   mrvalflujo(i).pago_total = (mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo) * mrvalflujo(i).dxv3
 'tasa de descuento
   mrvalflujo(i).t_desc = TasaFwdCurva(parval.perfwd, mrvalflujo(i).dxv3 + parval.perfwd, curva1, parval.modint1)
 'factor descuento
   mrvalflujo(i).factor_desc = 1 / (1 + mrvalflujo(i).t_desc * mrvalflujo(i).dxv3 / 360)
  'valor presente =(amortizacion+intereses)*factor descuento
   mrvalflujo(i).valor_presente = mrvalflujo(i).pago_total * mrvalflujo(i).factor_desc
 'vp acumulado
   valor = valor + mrvalflujo(i).valor_presente
Next i
DurDeudaFija = valor
End Function

Function ValDeudaVariable(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvadesc() As propCurva, ByRef curvacp() As propCurva, ByRef parval As paramValFlujo, ByRef mrvalflujo() As resValFlujo)
'mrvalflujo es una matriz de salida para la vista de los resultados en la pantalla
Dim i As Long
Dim indice As Long
Dim valor As Double

'25  valor presente
'se determina la estructura de la tabla de amortizacion
indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then 'no se debe de valuar si ya finalizo
   ReDim mrvalflujo(0 To 0) As resValFlujo
   ValDeudaVariable = 0
   Exit Function
End If
'se dimensiona una matriz donde se colocan el desglose de los resultados
mrvalflujo = CrearTablaAmortiza(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1
    mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)
    If fecha < mrvalflujo(i).fecha_ini Then                                  'tiene que ser una tasa forward
       mrvalflujo(i).tc_aplicar = TasaFwdCurva(mrvalflujo(i).dxv1 + parval.perfwd, mrvalflujo(i).dxv1 + parval.pcref + parval.perfwd, curvacp, parval.modint2)
    Else
      If Not SiIncTasaCVig Then
         mrvalflujo(i).tc_aplicar = TasaFwdCurva(parval.perfwd, parval.pcref + parval.perfwd, curvacp, parval.modint2)
      Else
         mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
      End If
    End If
'saldo al que se le aplicaran los intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+parval.sobret)*parval.pcref/360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-interses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar=amortizacion+intereses pagados
   mrvalflujo(i).pago_total = mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo
 'tasa de descuento
   mrvalflujo(i).t_desc = TasaFwdCurva(parval.perfwd, mrvalflujo(i).dxv3 + parval.perfwd, curvadesc, parval.modint1)
 'factor descuento
   mrvalflujo(i).factor_desc = 1 / (1 + mrvalflujo(i).t_desc * mrvalflujo(i).dxv3 / 360)
  'valor presente =(amortizacion+intereses)*factor descuento
   mrvalflujo(i).valor_presente = mrvalflujo(i).pago_total * mrvalflujo(i).factor_desc
 'vp acumulado
   valor = valor + mrvalflujo(i).valor_presente
Next i
ValDeudaVariable = valor
End Function

Function DurDeudaVariable(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvadesc() As propCurva, ByRef curvacp() As propCurva, ByRef parval As paramValFlujo)
Dim indice As Long
Dim mrvalflujo() As resValFlujo
Dim valor As Double
Dim i As Long

'mrvalflujo es una matriz de salida para la vista de los resultados en la pantalla
'para el calculo del valor presente de una deuda de tasa variable
'con amortizaciones

'curvadesc para el descuento
'curvacp para los cupones

'10  tasa a aplicar
'11

'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento
'24  factor descuento
'25  valor presente

'se determina la estructura de la tabla de amortizacion

indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then 'no se debe de valuar si ya finalizo
   DurDeudaVariable = 0
   Exit Function
End If
'se dimensiona una matriz donde se colocan el desglose de los resultados
mrvalflujo = CrearTablaAmortiza(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To parval.finFLujo - indice + 1
    mrvalflujo(i, 16) = mrvalflujo(i, 15)                              'periodo cupon a aplicar
    If fecha < mrvalflujo(i, 3) Then                                  'tiene que ser una tasa forward
       mrvalflujo(i, 10) = TasaFwdCurva(mrvalflujo(i, 12) + parval.perfwd, mrvalflujo(i, 12) + parval.pcref + parval.perfwd, curvacp, parval.modint2)
    Else
       If mrvalflujo(i, 10) = 0 Then mrvalflujo(i, 10) = TasaFwdCurva(parval.perfwd, parval.pcref + parval.perfwd, curvacp, parval.modint2)
    End If
'saldo al que se le aplicaran los intereses
    If mrvalflujo(i, 7) Then
       mrvalflujo(i, 17) = mrvalflujo(i, 8) * (mrvalflujo(i, 10) + mrvalflujo(i, 11))
    Else
       If i <> 1 Then
          mrvalflujo(i, 17) = mrvalflujo(i - 1, 17) - (mrvalflujo(i - 1, 8) - mrvalflujo(i, 8)) * (mrvalflujo(i, 10) + mrvalflujo(i, 11))
       Else
          mrvalflujo(i, 17) = mrvalflujo(i, 8) * (mrvalflujo(i, 10) + mrvalflujo(i, 11))
       End If
    End If
'intereses generados periodo = (saldo+intereses per ant) * (tasa+spread) * parval.pcref /360
 If i <> 1 Then
    If parval.acumInt Then
       mrvalflujo(i, 18) = (mrvalflujo(i, 17) + mrvalflujo(i - 1, 21) * (mrvalflujo(i, 10) + mrvalflujo(i, 11))) * mrvalflujo(i, 16) / 360
    Else
       mrvalflujo(i, 18) = mrvalflujo(i, 17) * mrvalflujo(i, 16) / 360
    End If
 Else
    mrvalflujo(i, 18) = mrvalflujo(i, 17) * mrvalflujo(i, 16) / 360
 End If
 'intereses acumulados periodo
 If i = 1 Then
    mrvalflujo(i, 19) = mrvalflujo(i, 18)
 Else
    mrvalflujo(i, 19) = mrvalflujo(i, 18) + mrvalflujo(i - 1, 21)
 End If
 'intereses pagados en el periodo= hasta el total de intereses acumulados
 If mrvalflujo(i, 6) Then
    mrvalflujo(i, 20) = mrvalflujo(i, 19)
 Else
    mrvalflujo(i, 20) = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-intereses pagados
 mrvalflujo(i, 21) = mrvalflujo(i, 19) - mrvalflujo(i, 20)
 'pago total  por plazo=(amortizacion+intereses pagados)*plazo
 mrvalflujo(i, 22) = (mrvalflujo(i, 9) + mrvalflujo(i, 20)) * mrvalflujo(i, 14)
 'tasa de descuento
 mrvalflujo(i, 23) = TasaFwdCurva(parval.perfwd, mrvalflujo(i, 14) + parval.perfwd, curvadesc, parval.modint1)
 'factor descuento
 mrvalflujo(i, 24) = 1 / (1 + mrvalflujo(i, 23) * mrvalflujo(i, 14) / 360)
  'valor presente =(amortizacion+intereses)*factor descuento
 mrvalflujo(i, 25) = mrvalflujo(i, 22) * mrvalflujo(i, 24)
 'vp acumulado
 valor = valor + mrvalflujo(i, 25)
Next i
DurDeudaVariable = valor
End Function

Function TasaFwdCurva(ByVal plazoc As Long, ByVal plazol As Long, ByRef curva() As propCurva, ByVal tinterpol As Integer)
Dim tasac As Double
Dim tasal As Double

If plazoc > 0 Then
   If plazol > 0 And plazol > plazoc Then
      tasal = CalculaTasa(curva, plazol, tinterpol)
      tasac = CalculaTasa(curva, plazoc, tinterpol)
      TasaFwdCurva = ((1 + tasal * plazol / 360) / (1 + tasac * plazoc / 360) - 1) * 360 / (plazol - plazoc)
   Else
       TasaFwdCurva = 0
   End If
Else
   If plazol > 0 Then
       TasaFwdCurva = CalculaTasa(curva, plazol, tinterpol)
   Else
       TasaFwdCurva = 0
   End If
End If
End Function

Function IDevDeudaVariable(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvacp() As propCurva, ByRef parval As paramValFlujo)
Dim indice As Long
Dim i As Long
Dim mrvalflujo() As resValFlujo
Dim valor As Double

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata


'para el calculo del valor presente de una deuda de tasa variable
'con amortizaciones
indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then 'no se debe de valuar si ya finalizo
   IDevDeudaVariable = 0
   Exit Function
End If
'se dimensiona una matriz donde se colocan el desglose de los resultados
mrvalflujo = CrearTablaIntDev(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
'se debe de determinar el plazo que hay entre cada corte de intereses
'para evitarnos problemas por ahora, supondremos que fecha>=matflujos(1,2)
valor = 0
For i = 1 To UBound(mrvalflujo, 1)
    mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)
    mrvalflujo(i).f_tiempo_aplicar = Minimo(mrvalflujo(i).f_tiempo_aplicar, (fecha - mrvalflujo(i).fecha_ini) / 360)
    If fecha < mrvalflujo(i).fecha_ini Then                                  'tiene que ser una tasa forward
       mrvalflujo(i).tc_aplicar = TasaFwdCurva(mrvalflujo(i).dxv1 + parval.perfwd, mrvalflujo(i).dxv1 + parval.pcref + parval.perfwd, curvacp, parval.modint2)
    Else
      If Not SiIncTasaCVig Then
         mrvalflujo(i).tc_aplicar = TasaFwdCurva(parval.perfwd, parval.pcref + parval.perfwd, curvacp, parval.modint2)
      Else
         mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
      End If
    End If
'saldo al que se le aplicaran los intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo = (saldo+intereses per ant) * (tasa+spread) * parval.pcref /360
 If i <> 1 Then
    If parval.acumInt = "S" Then
       mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 Else
    mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
 End If
 'intereses acumulados periodo
   If i = 1 Then
      mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
   Else
      mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
   End If
 'intereses pagados en el periodo= hasta el total de intereses acumulados
   If mrvalflujo(i).si_paga_int = "S" Then
      mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
   Else
     mrvalflujo(i).int_pag_periodo = 0
  End If
 'intereses acumulados sig periodo=intereses acumulados-intereses pagados
  mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar=amortizacion+intereses pagados
  mrvalflujo(i).pago_total = mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo
  valor = valor + mrvalflujo(i).int_pag_periodo
Next i
IDevDeudaVariable = valor

End Function

Function DeudaIntLineales(ByVal fecha, ByRef matflujos() As estFlujosDeuda, ByRef curvadesc() As propCurva, ByRef curvacp() As propCurva, ByVal pc As Integer, ByVal interci As Boolean, ByVal intercf As Boolean, ByVal tinterpol As Integer, ByRef mate() As Variant)
Dim noreg As Long
Dim i As Long
Dim indice As Long
Dim valor As Double
Dim plazo As Integer
Dim rate1 As Double
Dim rate2 As Double
Dim spread As Double
Dim amortiz As Double

If ActivarControlErrores Then
 On Error GoTo ControlErrores
End If
'para el calculo del valor presente de una deuda de tasa variable
'con amortizaciones
'curvadesc para el descuento
'curvacp para los cupones
'matflujos debe de tener la siguiente estructura
'1  fecha inicio pago
'2  fecha fin pago
'3  tasa cupon
'4  monto nocional
'5  spread
'6  subyacente
noreg = UBound(matflujos, 1)
'se determina en que renglon cae la fecha de valuacion
'para determinar que flujos se deben de incluir y cuales no

If fecha < matflujos(1, 3) Then  'no se debe de valuar si no han iniciado
 DeudaIntLineales = 0
 Exit Function
End If

For i = 1 To noreg
 If fecha < matflujos(i, 5) And fecha >= matflujos(i, 4) Then
  indice = i
  Exit For
 End If
Next i

If fecha < matflujos(1, 4) Then  'no se debe de valuar si no han iniciado
 indice = 1
End If
If fecha >= matflujos(noreg, 5) Then
 DeudaIntLineales = 0
 Exit Function
End If

ReDim mate(0 To noreg - indice + 1, 1 To 10) As Variant
'1 dias inicio cupon
'2 dias fin cupon
'3 periodo cupon
'4 saldo insoluto
'5 amortizacion
'6 tasa cupon vigente
'7 spread
'8 tasa de descuento
'9 intereses generados
'10 pago a valor presente

'se debe de determinar el plazo que hay entre cada corte de intereses
'para evitarnos problemas por ahora, supondremos que fecha>=matflujos(1,2)

For i = 1 To noreg - indice + 1
 mate(i, 1) = matflujos(i + indice - 1, 4) - fecha                        'dias por vencer i cupon
 mate(i, 2) = matflujos(i + indice - 1, 5) - fecha                        'dias por vencer f cupon
 mate(i, 3) = matflujos(i + indice - 1, 5) - matflujos(i + indice - 1, 4) 'plazo cupon
 mate(i, 4) = matflujos(i + indice - 1, 6)                                'saldo insoluto
 If intercf Then
  mate(i, 5) = matflujos(i + indice - 1, 7)                               'amortizacion de capital
 Else
  mate(i, 5) = 0                                                          'amortizacion de capital
 End If
 mate(i, 6) = matflujos(i + indice - 1, 8)                                'tasa cupon vigente
 mate(i, 7) = matflujos(i + indice - 1, 9)                                'spread
Next i
'1

ReDim tfut(1 To noreg - indice + 1) As Double

valor = 0

For i = 1 To noreg - indice + 1
'se calcula la tasa futura
If mate(i, 1) > 0 Then
 If pc <> 0 Then
  plazo = pc
 Else
  plazo = mate(i, 3)
 End If
 rate1 = CalculaTasa(curvacp, mate(i, 1), tinterpol)
 rate2 = CalculaTasa(curvacp, mate(i, 1) + plazo, tinterpol)
 'tasa forward
 mate(i, 6) = ((1 + rate2 * (mate(i, 1) + plazo) / 360) / (1 + rate1 * (mate(i, 1)) / 360) - 1) * 360 / plazo

Else
If i = 1 Then
'mate(i, 6) = calculartasa(curvacp, pc, tinterpol)
End If
End If
'la tasa de descuento
 mate(i, 8) = CalculaTasa(curvadesc, mate(i, 2), tinterpol)
 spread = mate(i, 7)
 amortiz = mate(i, 5)
' If mate(i, 8) = 0 Then MsgBox "Una tasa de descuento es nula"
' If mate(i, 6) = 0 Then MsgBox "La tasa cupon del flujo " & i & " es nula"
' If mate(i, 6) = 0 Then Call MostrarMensajeSistema("La tasa cupon del " & matflujos(i + indice - 1, 4) & " es nula", frmprogreso.label2, 0.5, Date, Time, NomUsuario)
' If mate(i, 8) = 0 Then Call MostrarMensajeSistema("Una tasa de descuento es nula", frmprogreso.label2, 0.5, Date, Time, NomUsuario)
 'se calcula el pago total
If i <> 1 Then
If mate(i, 5) = 0 And mate(i - 1, 5) <> 0 Then
 mate(i, 9) = mate(i, 4) * (mate(i, 6) + spread) * mate(i, 3) / 360
Else
 mate(i, 9) = mate(i - 1, 9) + mate(i, 4) * (mate(i, 6) + spread) * mate(i, 3) / 360
End If
Else
  mate(i, 9) = mate(i, 4) * (mate(i, 6) + spread) * mate(i, 3) / 360
End If
If amortiz <> 0 Then
 mate(i, 10) = (amortiz + mate(i, 9)) / (1 + mate(i, 8) * mate(i, 2) / 360)
Else
 mate(i, 10) = (amortiz) / (1 + mate(i, 8) * mate(i, 2) / 360)
End If
 
' MsgBox mate(i, 9)
 valor = valor + mate(i, 10)
Next i
DeudaIntLineales = valor
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function IDevDeudaFija(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef parval As paramValFlujo) As Double
Dim indice As Long
Dim mrvalflujo() As New resValFlujo
Dim valor As Double
Dim i As Long

'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata


'primero determinamos el flujo

indice = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indice <= 0 Then 'no se debe de valuar si ya finalizo
   IDevDeudaFija = 0
   
   Exit Function
End If
'se dimensiona una matriz donde se colocan el desglose de los resultados
mrvalflujo = CrearTablaIntDev(fecha, indice, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
valor = 0
For i = 1 To UBound(mrvalflujo, 1)
   If parval.pcref <> 0 Then
       mrvalflujo(i).f_tiempo_aplicar = parval.pcref / 360                                                       'fraccion de año a aplicar
    Else
       mrvalflujo(i).f_tiempo_aplicar = DefPlazo(mrvalflujo(i).fecha_ini, mrvalflujo(i).fecha_fin, parval.convInt)                'fraccion de año a aplicar
    End If
    mrvalflujo(i).f_tiempo_aplicar = Minimo(mrvalflujo(i).f_tiempo_aplicar, (fecha - mrvalflujo(i).fecha_ini) / 360)
    mrvalflujo(i).tc_aplicar = mrvalflujo(i).t_cupon_per
'determinacion del saldo a aplicar intereses
    If mrvalflujo(i).int_s_saldo = "S" Then
       mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
    Else
       If i <> 1 Then
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i - 1).saldo_efectivo_aplicar - (mrvalflujo(i - 1).saldo_periodo - mrvalflujo(i).saldo_periodo) * (mrvalflujo(i, 10) + mrvalflujo(i).sobretasa)
       Else
          mrvalflujo(i).saldo_efectivo_aplicar = mrvalflujo(i).saldo_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)
       End If
    End If
'intereses generados periodo=(saldo+intereses per ant)*(tasa+spread)*parval.pcref /360
    If i <> 1 Then
       If parval.acumInt = "S" Then
          mrvalflujo(i).int_gen_periodo = (mrvalflujo(i).saldo_efectivo_aplicar + mrvalflujo(i - 1).int_acum_sig_periodo * (mrvalflujo(i).tc_aplicar + mrvalflujo(i).sobretasa)) * mrvalflujo(i).f_tiempo_aplicar
       Else
          mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
       End If
    Else
       mrvalflujo(i).int_gen_periodo = mrvalflujo(i).saldo_efectivo_aplicar * mrvalflujo(i).f_tiempo_aplicar
    End If
 'intereses acumulados periodo=intereses generados+intereses periodo anterior
    If i = 1 Then
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo
    Else
       mrvalflujo(i).int_acum_periodo = mrvalflujo(i).int_gen_periodo + mrvalflujo(i - 1).int_acum_sig_periodo
    End If
 'intereses pagados en el periodo <= hasta el total de intereses acumulados
 If mrvalflujo(i).si_paga_int = "S" Then
    mrvalflujo(i).int_pag_periodo = mrvalflujo(i).int_acum_periodo
 Else
    mrvalflujo(i).int_pag_periodo = 0
 End If
 'intereses acumulados sig periodo=intereses acumulados-interses pagados
   mrvalflujo(i).int_acum_sig_periodo = mrvalflujo(i).int_acum_periodo - mrvalflujo(i).int_pag_periodo
 'pago total sin descontar=amortizacion+intereses pagados
   mrvalflujo(i).pago_total = mrvalflujo(i).amortizacion + mrvalflujo(i).int_pag_periodo
 'vp acumulado
   valor = valor + mrvalflujo(i).int_pag_periodo
Next i
IDevDeudaFija = valor

End Function

Function ValSwap(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvad1() As propCurva, ByRef curvad2() As propCurva, ByRef curvap1() As propCurva, ByRef curvap2() As propCurva, ByRef parval1 As paramValFlujo, ByRef parval2 As paramValFlujo, ByRef matval() As Double, ByRef mrvalflujo() As resValFlujo)
Dim mrvalflujo1() As New resValFlujo
Dim mrvalflujo2() As New resValFlujo

ReDim matval(1 To 2) As Double
matval(1) = 0
matval(2) = 0
  If parval1.modint1 <> 0 And parval1.modint2 = 0 Then
     matval(1) = parval1.tCambio * ValDeudaFija(fecha, matflujos, curvad1, parval1, mrvalflujo1)
  End If
  If parval1.modint1 <> 0 And parval1.modint2 <> 0 Then
     matval(1) = parval1.tCambio * ValDeudaVariable(fecha, matflujos, curvad1, curvap1, parval1, mrvalflujo1)
  End If
   If parval2.modint1 <> 0 And parval2.modint2 = 0 Then
      matval(2) = parval2.tCambio * ValDeudaFija(fecha, matflujos, curvad2, parval2, mrvalflujo2)
   End If
   If parval2.modint1 <> 0 And parval2.modint2 <> 0 Then
      matval(2) = parval2.tCambio * ValDeudaVariable(fecha, matflujos, curvad2, curvap2, parval2, mrvalflujo2)
   End If
'se unen las matrices con el desglose de las valuaciones

   If Not EsArrayVacio(mrvalflujo1) And Not EsArrayVacio(mrvalflujo2) Then
      mrvalflujo = unirResVFlujos(mrvalflujo1, mrvalflujo2)
   End If
'se devuelve la marca a mercado
ValSwap = matval(1) - matval(2)
End Function

Function unirResVFlujos(a, B)
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
noreg1 = UBound(a, 1)
noreg2 = UBound(B, 1)
ReDim matf(1 To noreg1 + noreg2) As New resValFlujo
For i = 1 To noreg1
  Set matf(i) = a(i)
Next i
For i = 1 To noreg2
Set matf(i + noreg1) = B(i)
Next i
unirResVFlujos = matf
End Function

Function IDevSwap(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvad1() As propCurva, ByRef curvad2() As propCurva, ByRef curvap1() As propCurva, ByRef curvap2() As propCurva, ByRef parval1 As paramValFlujo, ByRef parval2 As paramValFlujo)
Dim mrvalflujo1() As Variant
Dim mrvalflujo2() As Variant
ReDim matval(1 To 2) As Double
matval(1) = 0
matval(2) = 0
'If Val(parval1(8)) <> 0 And Val(parval1(9)) <> 0 Then     'posicion activa
   If parval1.modint1 <> 0 And parval1.modint2 = 0 Then
      matval(1) = parval1.tCambio * IDevDeudaFija(fecha, matflujos, parval1)
   End If
   If parval1.modint1 <> 0 And parval1.modint2 <> 0 Then
      matval(1) = parval1.tCambio * IDevDeudaVariable(fecha, matflujos, curvap1, parval1)
   End If
'End If
'If Val(parval2(8)) <> 0 And Val(parval2(9)) <> 0 Then     'posicion pasiva
  If parval2.modint1 <> 0 And parval2.modint2 = 0 Then
     matval(2) = parval2.tCambio * IDevDeudaFija(fecha, matflujos, parval2)
  End If
  If parval2.modint1 <> 0 And parval2.modint2 <> 0 Then
     matval(2) = parval2.tCambio * IDevDeudaVariable(fecha, matflujos, curvap2, parval2)
  End If
'End If

'se devuelve la marca a mercado
IDevSwap = matval
End Function

Function CDurSwap(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvad1() As propCurva, ByRef curvad2() As propCurva, ByRef curvap1() As propCurva, ByRef curvap2() As propCurva, ByRef parval1 As paramValFlujo, ByRef parval2 As paramValFlujo)
Dim yield As Double
Dim valsucio As Double
'parval1.intInFlujos = "S"
'parval1.intFinFlujos = "S"
'parval2.intInFlujos = "S"
'parval2.intFinFlujos = "S"
ReDim matval(1 To 2) As Double
ReDim matval1(1 To 2) As Double
ReDim matval2(1 To 2) As Double
Dim mrvalflujo() As New resValFlujo

matval1(1) = 0
matval1(2) = 0
  If parval1.modint1 <> 0 And parval1.modint2 = 0 Then
      yield = DetYieldDeudaFija(fecha, matflujos, curvad1, parval1)
      matval1(1) = parval1.tCambio * DurDFijaYield(fecha, matflujos, yield, parval1)
  End If
  If parval1.modint1 <> 0 And parval1.modint2 <> 0 Then
      matval1(1) = 0
  End If
   If parval2.modint1 <> 0 And parval2.modint2 = 0 Then
      yield = DetYieldDeudaFija(fecha, matflujos, curvad2, parval2)
      matval1(2) = parval2.tCambio * DurDFijaYield(fecha, matflujos, yield, parval2)
   End If
  If parval2.modint1 <> 0 And parval2.modint2 <> 0 Then
     matval1(2) = 0
  End If
  valsucio = ValSwap(fecha, matflujos, curvad1, curvad2, curvap1, curvap2, parval1, parval2, matval2, mrvalflujo)
  If matval2(1) <> 0 Then
     matval(1) = matval1(1) / matval2(1)
  Else
     matval(1) = 0
  End If
  If matval2(2) <> 0 Then
     matval(2) = matval1(2) / matval2(2)
  Else
     matval(2) = 0
  End If
  CDurSwap = matval
End Function

Function ValDeuda(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvad1() As propCurva, ByRef curvap1() As propCurva, ByRef parval1 As paramValFlujo, ByRef mrvalflujo() As resValFlujo)
Dim valor As Double

If parval1.modint1 <> 0 And parval1.modint2 = 0 Then
   valor = parval1.tCambio * ValDeudaFija(fecha, matflujos, curvad1, parval1, mrvalflujo)
End If
If parval1.modint1 <> 0 And parval1.modint2 <> 0 Then
   valor = parval1.tCambio * ValDeudaVariable(fecha, matflujos, curvad1, curvap1, parval1, mrvalflujo)
End If
'se devuelve la marca a mercado
ValDeuda = valor
End Function

Function IDevDeuda(ByVal fecha As Date, ByRef matflujos() As estFlujosDeuda, ByRef curvad1() As propCurva, ByRef curvap1() As propCurva, ByRef parval1 As paramValFlujo) As Double
Dim valor As Double
  If parval1.modint1 <> 0 And parval1.modint2 = 0 Then
      valor = parval1.tCambio * IDevDeudaFija(fecha, matflujos, parval1)
  End If
  If parval1.modint1 <> 0 And parval1.modint2 <> 0 Then
      valor = parval1.tCambio * IDevDeudaVariable(fecha, matflujos, curvap1, parval1)
  End If

'se devuelve la marca a mercado
IDevDeuda = valor
End Function

Function ContVarVacio(var)
If Len(Trim(var)) = 0 Or IsNull(var) Then
 ContVarVacio = True
Else
 ContVarVacio = False
End If
End Function

Function ValFIndice(ByVal vn As Double, ByVal dxv As Integer, ByVal desf As Integer, ByVal pasigna As Double, ByVal valindice As Double, ByRef curva1() As propCurva, ByRef cdesc() As propCurva, ByVal tmarca As String)
Dim tcorta As Double
Dim tdesc As Double


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
tcorta = CalculaTasa(curva1, dxv, 1)
tdesc = CalculaTasa(cdesc, desf, 1)
If tmarca = "VF" Then
    ValFIndice = vn * valindice * (1 + tcorta * dxv / 360)
   ElseIf tmarca = "VP" Then
    ValFIndice = vn * (valindice * (1 + tcorta * dxv / 360) - pasigna) / (1 + tdesc * desf / 360)
   End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function Moda(ByRef mat() As Variant, ByVal ind As Integer)
Dim noreg As Long

If ActivarControlErrores Then
   On Error GoTo ControlErrores
End If
'se determina la moda de una muestra en la matriz mat
   noreg = UBound(mat, 1)
   On Error GoTo 0
   Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CupVen(ByVal x As Integer, ByVal Y As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If Y <> 0 Then
CupVen = -Int(-x / Y)
Else
CupVen = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CND(ByVal x As Double, ByVal vm As Double, ByVal ds As Double) As Double
Dim r As Double
Dim kk As Double


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'la distribucion normal

    Const a1 = 0.31938153:  Const a2 = -0.356563782: Const a3 = 1.781477937:
    Const a4 = -1.821255978:  Const a5 = 1.330274429
    
    r = Abs((x - vm) / ds)
    kk = 1 / (1 + 0.2316419 * r)
    CND = 1 - 1 / Sqr(2 * Pi) * Exp(-r ^ 2 / 2) * (a1 * kk + a2 * kk ^ 2 + a3 * kk ^ 3 + a4 * kk ^ 4 + a5 * kk ^ 5)
    
    If x < 0 Then
        CND = 1 - CND
    End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Private Sub pDisplayError(ByVal sError As String)

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim sMsg As String

Screen.MousePointer = vbDefault
If Trim$(sError) = "" Then
    sMsg = Err.Description
Else
    sMsg = sError & "  " & Err.Description
End If
If Err.Number = 0 Then
    Call MostrarMensajeSistema(sMsg, frmProgreso.Label2, 2, Date, Time, NomUsuario)
Else
    Call MostrarMensajeSistema(sMsg & " (" & CStr(Err.Number) & ")", frmProgreso.Label2, 2, Date, Time, NomUsuario)
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Public Function PbonoY(ByVal vn As Double, ByVal tc As Double, ByVal tdesc As Double, ByVal nper As Long)
Dim suma As Double
Dim i As Long

suma = 0
For i = 1 To nper
    suma = suma + (vn * tc * 182 / 360) / (1 + tdesc / 2) ^ (i)
Next i
suma = suma + vn / (1 + tdesc / 2) ^ nper
PbonoY = suma
End Function

Function DeterminaTObj(ByVal pobj As Double, ByVal vn As Double, ByVal tc As Double, ByVal nper As Long, ByVal verror As Double) As Double
Dim t0 As Double
Dim inc As Double
Dim precio1 As Double
Dim precio As Double
Dim dprecio As Double

t0 = 0.005
precio = PbonoY(vn, tc, t0, nper)
inc = 0.0000001
Do While Abs(precio - pobj) > verror
 precio1 = PbonoY(vn, tc, t0 + inc, nper)
 dprecio = (precio1 - precio) / inc
 If dprecio <> 0 Then
    t0 = t0 - (precio - pobj) / dprecio
 Else
    MsgBox "la derivada se anulo"
    Exit Function
 End If
 precio = PbonoY(vn, tc, t0, nper)
Loop
DeterminaTObj = t0
End Function

Function ConvertirTextoFecha(ByVal txtfecha As String, ByVal tfecha As Integer) As Date
If Len(Trim(txtfecha)) = 8 Then
 If tfecha = 0 Then
    ConvertirTextoFecha = CDate(Mid(txtfecha, 7, 2) & "/" & Mid(txtfecha, 5, 2) & "/" & Mid(txtfecha, 1, 4))
 Else
    ConvertirTextoFecha = CDate(Mid(txtfecha, 5, 2) & "/" & Mid(txtfecha, 7, 2) & "/" & Mid(txtfecha, 1, 4))
 End If
Else
   ConvertirTextoFecha = CDate(txtfecha)
End If
End Function


Function DetFechasEscEfic()
Dim matfechas(1 To 10) As Date
matfechas(1) = #3/11/2015#
matfechas(2) = #9/2/2015#
matfechas(3) = #9/15/2016#
matfechas(4) = #9/12/2016#
matfechas(5) = #4/18/2016#

matfechas(6) = #2/4/2016#
matfechas(7) = #8/27/2014#
matfechas(8) = #9/10/2013#
matfechas(9) = #9/11/2012#
matfechas(10) = #4/27/2012#
DetFechasEscEfic = matfechas
End Function

Function EstresFREsc(ByVal fecha As Date, ByRef mfriesgo0() As Double)
Dim noesc As Integer
Dim i As Long
Dim mfriesgo() As Variant
Dim mfriesgo2() As Double
Dim exito As Boolean

    Call CrearMatFRiesgo2(fecha - 10, fecha, mfriesgo, "", exito)
    noesc = UBound(mfriesgo, 1)
    ReDim mfriesgo2(1 To NoFactores, 1 To 1)
    For i = 1 To NoFactores
        If Val(mfriesgo(noesc - 1, i + 1) * mfriesgo(noesc, i + 1)) <> 0 Then
           mfriesgo2(i, 1) = mfriesgo0(i, 1) * mfriesgo(noesc, i + 1) / mfriesgo(noesc - 1, i + 1)
        Else
           mfriesgo2(i, 1) = mfriesgo0(i, 1)
        End If
    Next i
EstresFREsc = mfriesgo2
End Function

Sub CalculaEficProsFWD(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
On Error GoTo hayerror
Dim mfriesgo() As Variant
Dim fecha0 As Date
Dim fecha1 As Date
Dim txtcurva1 As String
Dim txtcurva2 As String
Dim txtcurva3 As String
Dim txttcambio As String
Dim curva1() As New propCurva
Dim curva2() As New propCurva
Dim curva3() As New propCurva
Dim valcam() As Variant
Dim mattxt() As String
Dim finicio As Date
Dim hinicio As Date
Dim mata() As Variant
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim noreg As Long
Dim noescen As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim dxv0 As Long
Dim dxv1 As Long
Dim indice As Long
Dim fcompra As Date
Dim fvenc As Date
Dim tcstrike As Double
Dim mnocional As Double
Dim eficpros As Double
Dim txtcadena As String
Dim vtasa1 As Double
Dim vtasa2 As Double
Dim tdesc As Double
Dim tc0 As Double
Dim tc As Double
Dim vfwd0 As Double
Dim vfwd10 As Double
Dim tcstrike0 As Double
Dim contar As Long
Dim nodim1 As Long
Dim kk As Long
Dim fechax As Date
Dim vfwd1 As Double
Dim vfwd2 As Double
Dim txtsalida As String
Dim txtmsg2 As String
Dim exito2 As Boolean
Dim txtmsg0 As String
Dim horareg As String
Dim noffut  As Long
Dim matffut() As Date
Dim matfechas() As Date
Dim MDblTasas1() As Double
Dim pfwd As Long

'esta rutina calcula la eficiencia de los forward de tipo de cambio,
'se tienen que cargar varios insumos para esta tarea
'la historia del tipo de cambio
'los escenarios de estres
finicio = Date
hinicio = Time

matfechas = DetFechasEscEfic
noescen = UBound(matfechas, 1) 'no de escenarios sobre los que se va a calcular la efectividad
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)

noreg = UBound(matpos, 1) 'el no de forwards a analizar
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      matffut = DefFechasEfCob(fecha, matposfwd(1).FVencFwd)
      noffut = UBound(matffut, 1)

      ReDim matres(1 To noreg, 1 To 20) As Variant
     'se cargan las tasas del escenario
      fecha0 = PBD1(fecha, 1, "MX")
      MatFactR1 = CargaFR1Dia(fecha0, exito)

      For i = 1 To noreg  'el total de fwds a analizar
          indice = matpos(i).IndPosicion
          fcompra = matposfwd(indice).FCompraFwd
          fvenc = matposfwd(indice).FVencFwd
          tcstrike = matposfwd(indice).PAsignadoFwd
          dxv0 = fvenc - fecha
          txtcurva1 = matposfwd(indice).FRiesgo1Fwd
          txtcurva3 = matposfwd(indice).FRiesgo2Fwd
          txtcurva2 = matposfwd(indice).FRiesgo3Fwd
          txttcambio = matposfwd(indice).TCambioFwd
          mnocional = matposfwd(indice).MontoNocFwd                 'monto nocional
          'ReDim matefic(1 To dxv0, 1 To 2) As Variant
          ReDim matefic(1 To noffut - 1, 1 To 4) As Variant
          For j = 1 To noffut - 1
              matefic(j, 1) = matffut(j)
              matefic(j, 2) = matffut(j + 1)
              matefic(j, 3) = noescen
          Next j
'se dimensiona la matriz con los resultados
          eficpros = 0
          If dxv0 <> 0 Then

             For j = 1 To noescen
                 fecha1 = matfechas(j)
                 ReDim matres1(1 To dxv0 + 1, 1 To 18) As Variant
                 ReDim matres1(1 To noffut + 1, 1 To 18) As Variant
   
   'se calcula el precio del forward en la fecha de compra
                 MatFactR1 = CargaFR1Dia(fecha0, exito)
                 curva1 = CrearCurvaNodos1(txtcurva1, MatFactR1)
                 curva2 = CrearCurvaNodos1(txtcurva2, MatFactR1)
                 curva3 = CrearCurvaNodos1(txtcurva3, MatFactR1)
                 vtasa1 = CalculaTasa(curva1, dxv0, 1)    'tasa local
                 vtasa2 = CalculaTasa(curva2, dxv0, 1)    'tasa extranjera
                 tdesc = CalculaTasa(curva3, dxv0, 1)     'tasa descuento
                 tc0 = ObtieneFRiesgo(txttcambio, MatFactR1)
                 vfwd0 = mnocional * (tc0 / (1 + vtasa2 * dxv0 / 360) - tcstrike / (1 + vtasa1 * dxv0 / 360))
                 tcstrike0 = tc0 * (1 + vtasa1 * dxv0 / 360) / (1 + vtasa2 * dxv0 / 360)
                 vfwd10 = 0
                 matres1(1, 1) = 0                  'clave de la operacion
                 matres1(1, 2) = contar             'orden
                 matres1(1, 3) = fecha              'fecha de analisis
                 matres1(1, 4) = vfwd0              '
                 matres1(1, 5) = 0                  '
    
   'se estresan las curvas de valuacion
                 MDblTasas1 = EstresFREsc(fecha1, MatFactR1)
                 For k = 2 To noffut
                    contar = contar + 1
                    matres1(k, 1) = matpos(1).c_operacion         'clave de la operacion
                    matres1(k, 2) = k                             'orden
                    matres1(k, 3) = matffut(k)                    'fecha de analisis
                    dxv1 = fvenc - matffut(k)
                    pfwd = matffut(k) - fecha
                    curva1 = CrearCurvaNodos1(txtcurva1, MDblTasas1)
                    curva2 = CrearCurvaNodos1(txtcurva2, MDblTasas1)
                    curva3 = CrearCurvaNodos1(txtcurva3, MDblTasas1)
                    vtasa1 = TasaFwdCurva(pfwd, dxv1 + pfwd, curva1, 1)  'tasa local
                    vtasa2 = TasaFwdCurva(pfwd, dxv1 + pfwd, curva2, 1) 'tasa extranjera
                    tdesc = TasaFwdCurva(pfwd, dxv1 + pfwd, curva3, 1)  'tasa descuento
                 
                    tc = ObtieneFRiesgo(txttcambio, MDblTasas1)
   'suponemos tipo de cambio spot t= tipo de cambio fwd t
                    vfwd1 = mnocional * (tc / (1 + vtasa2 * dxv1 / 360) - tcstrike / (1 + vtasa1 * dxv1 / 360))
                    matres1(k, 4) = vfwd1      '
                    vfwd2 = mnocional * (tc / (1 + vtasa2 * dxv1 / 360) - tcstrike0 / (1 + vtasa1 * dxv1 / 360))
                    matres1(k, 5) = vfwd2
                    matres1(k, 6) = matres1(k, 4) - matres1(k - 1, 4)  'DIFERENCIAS VALOR RAZONABLE
                    matres1(k, 7) = matres1(k, 5) - matres1(k - 1, 5)  'diferencias valor razonable
                    If matres1(k, 7) <> 0 Then
                       matres1(k, 8) = matres1(k, 6) / matres1(k, 7)
                    Else
                       matres1(k, 8) = 0
                    End If
                    If matres1(k, 8) >= 0.8 And matres1(k, 8) <= 1.25 Then
                       matres1(k, 9) = 1
                    Else
                       matres1(k, 9) = 0
                    End If
                    matres1(k, 10) = matres1(k, 10) + matres1(k, 9)
                    matefic(k - 1, 4) = matefic(k - 1, 4) + matres1(k, 9)
                    'fechax = fechax + 1
                    txtsalida = ""
                    txtsalida = txtsalida & matres1(k, 1) & Chr(9) & matres1(k, 2) & Chr(9)
                    'Print #1, txtsalida
                 Next k
                 'Print #1, ""
                 For kk = 2 To noffut
                     eficpros = eficpros + matres1(kk, 9)
                 Next kk
             Next j
          End If
          If ((noffut - 1) * noescen) <> 0 Then eficpros = eficpros / ((noffut - 1) * noescen)
          matres(i, 1) = matpos(1).c_operacion
          matres(i, 20) = eficpros
          If eficpros >= 0.95 Then
             exito = True
             txtmsg = "La efectividad de la operacion " & matpos(1).c_operacion & " es del " & Format(eficpros, "##0.00 %")
          Else
             exito = False
             txtmsg = "La operación " & matpos(1).c_operacion & " no es eficiente prospectivamente " & Format(eficpros, "##0.00 %")
          End If
     
      Next i
      horareg = matpos(1).HoraRegOp
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, matpos(1).c_operacion, eficpros, conAdo2)
      Call GuardarResEfectPros(fecha, matpos(1).c_operacion, matefic)
      conAdo2.Close
      ReDim mata(1 To noreg, 1 To 5) As Variant
      Call ValidarOperacion3(matpos(1).c_operacion, matpos(1).HoraRegOp, finicio, hinicio, Date, Time)
 
   Else
      exito = False
      txtmsg = "no se parametrizo la operacion"
   End If
Else
   txtmsg = "La posicion no tiene registros validos"
   exito = False
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Sub EficienciaRetroFwd(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim indice As Integer
Dim mfriesgo1() As Double
Dim mfriesgo2() As Double
Dim curva1() As New propCurva
Dim curva2() As New propCurva
Dim curva3() As New propCurva
Dim valcam() As Variant
Dim resval() As Variant
Dim txtcoper As String
Dim fcompra As Date
Dim fven As Date
Dim pstrike As Double
Dim txtfriesgo1 As String
Dim txtfriesgo2 As String
Dim txtfriesgo3 As String
Dim txttcambio As String
Dim vn As Double
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim dxv0 As Long
Dim dxv1 As Long
Dim dxv2 As Long
Dim dxv As Long
Dim vtasa1 As Double
Dim vtasa2 As Double
Dim tdesc As Double
Dim tc0 As Double
Dim pstriket0 As Double
Dim puntosfwd0 As Double
Dim valfwd0 As Double
Dim tc As Double
Dim valfwd1 As Double
Dim valfwdt1 As Double
Dim valfwd2 As Double
Dim valfwdt2 As Double
Dim difvfwd As Double
Dim difvfwdt As Double
Dim txtcadena As String
Dim exitofr As Boolean
Dim txtmsg0 As String
Dim fechax As Date


'se obtiene las caracterisricas de la operacion a procesar
mattxt = CrearFiltroPosPort(fecha2, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg, exito)
   ReDim matres(1 To noreg, 1 To 19) As Variant
   MatFactRiesgo1 = CargaFR1Dia(fecha1, exito)
   fechax = PBD1(fecha2, 1, "MX")
   MatFactRiesgo2 = CargaFR1Dia(fechax, exito)
   indice = matpos(1).IndPosicion
   fcompra = matposfwd(indice).FCompraFwd
   fven = matposfwd(indice).FVencFwd
   pstrike = matposfwd(1).PAsignadoFwd
   dxv0 = fven - fcompra
'se calcula el precio del forward con la fecha de compra
   vn = matposfwd(indice).MontoNocFwd
   txtfriesgo1 = matposfwd(indice).FRiesgo1Fwd
   txtfriesgo2 = matposfwd(indice).FRiesgo2Fwd
   txtfriesgo3 = matposfwd(indice).FRiesgo3Fwd
   txttcambio = matposfwd(indice).TCambioFwd
   MatFactR1 = CargaFR1Dia(CDate(Minimo(fcompra, fecha2)), exito)
   curva1 = CrearCurvaNodos1(txtfriesgo1, MatFactR1)
   curva2 = CrearCurvaNodos1(txtfriesgo3, MatFactR1)
   curva3 = CrearCurvaNodos1(txtfriesgo2, MatFactR1)
   vtasa1 = CalculaTasa(curva1, dxv0, 1)        'tasa local
   vtasa2 = CalculaTasa(curva2, dxv0, 1)        'tasa extranjera
   tdesc = CalculaTasa(curva3, dxv0, 1)         'tasa descuento
   tc0 = ObtieneFRiesgo(txttcambio, MatFactR1) 'tipo de cambio en la fecha de negociacion
   pstriket0 = tc0 * (1 + vtasa1 * dxv0 / 360) / (1 + vtasa2 * dxv0 / 360)   'strike pactado teorico
   If fecha2 <= fven Then
'    se calcula el strike del derivado hipotetico
'9   valor del forward al inicio de la operacion
'10  valor del forward en fecha
'11  puntos fwd a la fecha
'12  valor del fwd en t0
'13  efectividad 1
'14  devengo
'15  fwd -devengo
'16  puntos fwds
'19  valor del fwd
      matres(1, 1) = 1
      matres(1, 2) = matposfwd(1).c_operacion        'clave de la operacion
      matres(1, 3) = matposfwd(1).ClaveProdFwd
      matres(1, 4) = fven
      dxv1 = fven - fecha1
      matres(1, 5) = dxv1
      matres(1, 6) = pstrike
     
'se cargan las curvas para esta fecha
      curva1 = CrearCurvaNodos1(txtfriesgo1, MatFactRiesgo1)
      curva2 = CrearCurvaNodos1(txtfriesgo3, MatFactRiesgo1)
      curva3 = CrearCurvaNodos1(txtfriesgo2, MatFactRiesgo1)
      tc = ObtieneFRiesgo(txttcambio, MatFactRiesgo1)
      valfwd1 = ValFwdDiv(pstrike, dxv1, 0, curva1, curva2, curva3, tc, 1, resval)
      matres(1, 7) = valfwd1
      valfwdt1 = ValFwdDiv(pstriket0, dxv1, 0, curva1, curva2, curva3, tc, 1, resval)
      matres(1, 8) = valfwdt1
      curva1 = CrearCurvaNodos1(txtfriesgo1, MatFactRiesgo2)
      curva2 = CrearCurvaNodos1(txtfriesgo3, MatFactRiesgo2)
      curva3 = CrearCurvaNodos1(txtfriesgo2, MatFactRiesgo2)
      tc = ObtieneFRiesgo(txttcambio, MatFactRiesgo2)
      dxv2 = fven - fecha2
      matres(1, 9) = valfwdt1
      valfwd2 = ValFwdDiv(pstrike, dxv2, 0, curva1, curva2, curva3, tc, 1, resval)
      matres(1, 10) = valfwd2
      valfwdt2 = ValFwdDiv(pstriket0, dxv2, 0, curva1, curva2, curva3, tc, 1, resval)
      matres(1, 11) = valfwdt2
      difvfwd = valfwd2 - valfwd1
      difvfwdt = valfwdt2 - valfwdt1
      matres(1, 12) = difvfwd
      matres(1, 13) = difvfwdt
      matres(1, 14) = difvfwd / difvfwdt
      Call GuardarResEfRetroFwd(fecha2, matpos, matposfwd, 1, 4, matres(1, 14))
      Call GuardaResEficRetro(fecha2, matpos(1).c_operacion, matres(1, 14), ConAdo)
      If matres(1, 14) < 0.8 Or matres(1, 14) > 1.25 Then
        txtmsg = "La operacion " & matpos(1).c_operacion & " no es efectiva " & Format(matres(1, 12), "#,##0.00 %")
        final = True
        exito = False
      Else
        txtmsg = "El proceso finalizo correctamente"
        final = True
        exito = True
      End If
      
   End If
Else
   final = True
   exito = False
End If
End Sub

Function ObTasaCupVigMD(ByVal fecha As Date, ByRef matflujos() As estFlujosMD, ByVal txtfactor As String)
Dim fechai As Date
Dim noreg As Long
Dim infecha As Long
Dim indice As Long
Dim mata() As Variant
Dim i As Long

'primero se obtiene el factor de riesgo
fechai = #1/1/2006#
mata = Leer1FactorR(fechai, fecha, txtfactor, 0)
noreg = UBound(matflujos, 1)
infecha = 0
indice = 0
For i = 1 To noreg
If matflujos(i, 4) <= fecha And fecha < matflujos(i, 5) Then
 infecha = matflujos(i, 4)
Exit For
End If
Next i
If infecha <> 0 Then
indice = BuscarValorArray(infecha, mata, 1)
If indice <> 0 Then
ObTasaCupVigMD = mata(indice, 2)
Else
ObTasaCupVigMD = 0
End If
Else
ObTasaCupVigMD = 0
End If
End Function

Function RutinaOrden(ByRef mata() As Variant, ByVal ncol As Integer, ByVal oprutina As Integer) As Variant()
Dim exito As Boolean
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Integer

'esta rutina funciona mejor si hay pocos elementos repetidos
n = UBound(mata, 1)
m = UBound(mata, 2)
ReDim matb(1 To n, 1 To 2) As Variant
For i = 1 To n
    matb(i, 1) = i
    matb(i, 2) = mata(i, ncol)
Next i
If oprutina = 1 Then         'rutina quicksort
   'Call OrdenQuickSort3(matb, 2, 1, n, 1, False, exito)
   Call QuickSort4(matb, 2, 1, n)
ElseIf oprutina = 2 Then    'primera rutina de ordenacion de diseño propio
   matb = SuperOrden(matb, 2, exito)
ElseIf oprutina = 3 Then    'algoritmo clasico de la burbuja
   matb = OrdenarMat(matb, 2, exito)
End If
ReDim matc(1 To n, 1 To m) As Variant
For i = 1 To UBound(matb, 1) - 1
   If matb(i, 2) > matb(i + 1, 2) Then MsgBox "no se ordenaron los datos"
Next i
For i = 1 To n
 For j = 1 To m
    matc(i, j) = mata(matb(i, 1), j)
 Next j
Next i
RutinaOrden = matc
End Function

Function ROrdenDbl(ByRef mata() As Double, ByVal ncol As Integer) As Double()
Dim exito As Boolean
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Integer

'esta rutina funciona mejor si hay pocos elementos repetidos
n = UBound(mata, 1)
m = UBound(mata, 2)
ReDim matb(1 To n, 1 To 2) As Variant
For i = 1 To n
    matb(i, 1) = i
    matb(i, 2) = mata(i, ncol)
Next i
   Call QuickSort4(matb, 2, 1, n)
ReDim matc(1 To n, 1 To m) As Double
For i = 1 To UBound(matb, 1) - 1
   If matb(i, 2) > matb(i + 1, 2) Then MsgBox "no se ordenaron los datos"
Next i
For i = 1 To n
 For j = 1 To m
    matc(i, j) = mata(matb(i, 1), j)
 Next j
Next i
ROrdenDbl = matc
End Function

Function ROrdenF(ByRef mata() As Date, ByVal ncol As Integer) As Date()
Dim exito As Boolean
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Integer

'esta rutina funciona mejor si hay pocos elementos repetidos
n = UBound(mata, 1)
m = UBound(mata, 2)
ReDim matb(1 To n, 1 To 2) As Variant
For i = 1 To n
    matb(i, 1) = i
    matb(i, 2) = mata(i, ncol)
Next i
   Call QuickSort4(matb, 2, 1, n)
ReDim matc(1 To n, 1 To m) As Date
For i = 1 To UBound(matb, 1) - 1
   If matb(i, 2) > matb(i + 1, 2) Then MsgBox "no se ordenaron los datos"
Next i
For i = 1 To n
 For j = 1 To m
    matc(i, j) = mata(matb(i, 1), j)
 Next j
Next i
ROrdenF = matc
End Function

Function DefMatVariant(n1, n2)
 ReDim matc(1 To n1, 1 To n2) As Variant
 DefMatVariant = matc
End Function

Function DefMatDate(n1, n2)
 ReDim matc(1 To n1, 1 To n2) As Date
 DefMatDate = matc
End Function

Function DefMatDouble(n1, n2)
 ReDim matc(1 To n1, 1 To n2) As Double
 DefMatDouble = matc
End Function

Sub OrdenQuickSort3(ByRef mata() As Variant, ByVal ncol As Integer, ByVal inicio As Integer, ByVal final As Integer, ByRef orden As Integer, ByRef salir As Boolean, ByRef exito As Boolean)
Dim pivote As Integer

'modificacion al algoritmo quicksort
If orden > 2700 Then
   salir = True
   exito = False
   Exit Sub
ElseIf salir Then
    exito = False
    Exit Sub
End If
If inicio < final Then
   pivote = OrdenaQuickS(mata, ncol, inicio, final)
   Call OrdenQuickSort3(mata, ncol, inicio, pivote, orden + 1, salir, exito)
   Call OrdenQuickSort3(mata, ncol, pivote + 1, final, orden + 1, salir, exito)
   DoEvents
End If
End Sub

Function OrdenaQuickS(ByRef mata() As Variant, ByVal ncol As Integer, ByVal inicio As Integer, ByVal final As Integer) As Integer
Dim orden As Integer
Dim indice As Integer
Dim ind1 As Integer
Dim ind2 As Integer
Dim valpiv As Variant

orden = UBound(mata, 2) 'no de columnas
indice = inicio + Int((final - inicio) * Rnd)
Call PermRengV(mata, indice, final)
ind1 = inicio
ind2 = final - 1
valpiv = mata(final, ncol)
 Do While ind1 < ind2
  Do While mata(ind1, ncol) <= valpiv And ind1 < ind2
   ind1 = ind1 + 1
  Loop
  Do While mata(ind2, ncol) > valpiv And ind1 < ind2
   ind2 = ind2 - 1
  Loop
  If mata(ind1, ncol) > mata(ind2, ncol) Then
   Call PermRengV(mata, ind1, ind2)
  End If
 Loop
If mata(ind1, ncol) > mata(final, ncol) Then
 Call PermRengV(mata, ind1, final)
End If
OrdenaQuickS = ind1
End Function

Sub QuickSort4(ByRef mata() As Variant, ByVal ncol As Integer, ByVal first As Long, ByVal last As Long)
    Dim Low As Long, High As Long
    Dim MidValue As Variant
    Low = first
    High = last
    MidValue = mata((first + last) \ 2, ncol)
    Do
        While mata(Low, ncol) < MidValue
            Low = Low + 1
        Wend
        
        While mata(High, ncol) > MidValue
            High = High - 1
        Wend
        If Low <= High Then
            Call PermRengV(mata, Low, High)
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    If first < High Then Call QuickSort4(mata, ncol, first, High)
    If Low < last Then Call QuickSort4(mata, ncol, Low, last)
End Sub

Function VerifAccesoArch(ByVal txtarchivo As String) As Boolean
On Error GoTo Control
If Dir(txtarchivo) <> "" And Len(Trim(txtarchivo)) <> 0 Then
If IsFileOpen(txtarchivo) Then
   VerifAccesoArch = False
Else
   VerifAccesoArch = True
End If
Else
 VerifAccesoArch = False
'  MsgBox "No existe el archivo " & txtarchivo
End If
Exit Function
Control:
' MsgBox error(Err())
If Err() = 52 Then
 VerifAccesoArch = False
 MensajeProc = "No hay acceso al archivo " & txtarchivo
End If
On Error GoTo 0
End Function

Function IsFileOpen(ByVal FileName As String) As Boolean
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function

Function EncontrarSubCadenas(ByVal texto As String, ByVal txtclave As String) As String()
Dim largo As Long
Dim contar As Long
Dim i As Long
Dim inicio As Long

largo = Len(texto)
contar = 0
For i = 1 To largo
If Mid(texto, i, 1) = txtclave Or i = largo Then
contar = contar + 1
End If
Next i
ReDim mata(1 To contar) As String
contar = 0
inicio = 0
For i = 1 To largo
If Mid(texto, i, 1) = txtclave Then
 contar = contar + 1
 mata(contar) = Mid(texto, inicio + 1, i - inicio - 1)
 inicio = i
ElseIf i = largo Then
 contar = contar + 1
 mata(contar) = Mid(texto, inicio + 1, i - inicio)
 inicio = i
End If
Next i
EncontrarSubCadenas = mata
End Function

Function DeterminaPerfilVal(txtperfil)

Dim parval As New ParamValPos
Select Case txtperfil
Case "VALUACION"
     parval.siPLimpio = True
     parval.indpos = 0
     parval.siValExc = ValExacta         'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Valuacion de la posicion"  'Mensaje a mostrar en la pantalla
     parval.mVBondesD = 1                 'forma de valuacion del BondesD
     parval.perfwd = 0                 'periodo fwd
     parval.sicalcdur = True       'calcular duracion
     parval.sicalcPE = True
     
     parval.sicalcdv01 = True             'calcular dv01
     parval.siTCambio = True             'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
Case "SENSIBILIDADES"
     parval.siPLimpio = False
     parval.indpos = 0
     parval.siValExc = False             'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Calculando sensibilidades"   'MensajeProc a mostrar en la pantalla
     parval.mVBondesD = 2                 'forma de valuacion del BondesD
     parval.perfwd = 0                 'periodo fwd
     parval.sicalcdur = False             'calcular duracion
     parval.sicalcdv01 = False            'calcular dv01
     parval.siTCambio = True             'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "HISTORICO"
     parval.siPLimpio = False
     parval.indpos = 0
     parval.siValExc = False              'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Calculando CVaR Histórico"   'MensajeProc a mostrar en la pantalla
     parval.mVBondesD = 1                 'forma de valuacion del BondesD
     parval.perfwd = 0                    'periodo fwd
     parval.sicalcdur = False             'calcular duracion
     parval.sicalcdv01 = False            'calcular dv01
     parval.siTCambio = True              'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "MONTECARLO"
     parval.siPLimpio = False
     parval.indpos = 0
     parval.siValExc = False             'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Calculando Montecarlo"   'Mensaje a mostrar en la pantalla
     parval.mVBondesD = 2                 'forma de valuacion del BondesD 1 truncado, 2 exacto
     parval.perfwd = 0                 'periodo fwd
     parval.sicalcdur = False             'calcular duracion
     parval.sicalcdv01 = False            'calcular dv01
     parval.siTCambio = True             'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "ESTRES"
     parval.siPLimpio = False
     parval.indpos = 0
     parval.siValExc = False             'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Calculando escenarios estres"  'Mensaje a mostrar en la pantalla
     parval.mVBondesD = 1                 'forma de valuacion del BondesD 1 truncado, exacto
     parval.perfwd = 0             'periodo fwd
     parval.sicalcdur = False             'calcular duracion
     parval.sicalcdv01 = False            'calcular dv01
     parval.siTCambio = True              'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "BACKTESTING"
     parval.siPLimpio = True
     parval.indpos = 0
     parval.siValExc = True                          'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Valuacion de la posicion"     'Mensaje a mostrar en la pantalla
     parval.mVBondesD = 1                            'forma de valuacion del BondesD
     parval.perfwd = 0                               'periodo fwd
     parval.sicalcdur = False                        'calcular duracion
     parval.sicalcdv01 = False                       'calcular dv01
     parval.siTCambio = True                         'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "LCONTRAPARTE"
     parval.siPLimpio = False
     parval.indpos = 0
     parval.siValExc = False                       'valuacion exacta o con alguna aproximacion
     parval.msgSist = "Valuacion de la posicion"   'MensajeProc a mostrar en la pantalla
     parval.mVBondesD = 1                          'forma de valuacion del BondesD
     parval.perfwd = 0                             'periodo fwd
     parval.sicalcdur = False                      'calcular duracion
     parval.sicalcdv01 = False                     'calcular dv01
     parval.siTCambio = False                      'considerar tipo de cambio en los calculos
     parval.si_int_flujos = False
     parval.sicalcPE = False
Case "EFECTIVIDAD"
     parval.siPLimpio = True
     parval.indpos = 0
     parval.siValExc = False                       'valuacion exacta o con alguna aproximacion
     parval.msgSist = "calculo de efectividad"     'MensajeProc a mostrar en la pantalla
     parval.mVBondesD = 1                          'forma de valuacion del BondesD
     parval.perfwd = 0                             'periodo fwd
     parval.sicalcdur = False                      'calcular duracion
     parval.sicalcdv01 = False                     'calcular dv01
     parval.siTCambio = True                       'considerar tipo de cambio en los calculos
     parval.si_int_flujos = True
     parval.sicalcPE = False
Case Else
   MsgBox "No se ha definido parametros de valuación para el perfil " & txtperfil
End Select
 Set DeterminaPerfilVal = parval
End Function

Function GenMatRendRiesgo(ByRef matfr() As Variant, ByVal indice As Integer, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim n As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim r As Long
Dim estaln As Boolean
'se generan una matriz solo con los factores
'de riesgo necesarios para un var markowitz de portafolio
'matfr matriz de factores de riesgo
'indice - renglon donde se encuentra la fecha objetivo
'dias  -  dias de busqueda de informacion
n = UBound(matfr, 1)
ReDim matd(1 To noesc, 1 To n) As Double
ReDim mattr(1 To n) As Variant
For i = 1 To n
    For j = 1 To NoFactores
 'se incluye la historia del factor i en el calculo
        If matfr(i, 1) = MatCaracFRiesgo(j).indFactor Then
           For p = 1 To noesc
               r = indice - noesc + p
               estaln = Esblacklistfr(MatFactRiesgo(r, 1), MatCaracFRiesgo(j).indFactor)
               If estaln Then
                  matd(p, i) = 0
               Else
                  matd(p, i) = CalcRend2(MatFactRiesgo(r - htiempo, j + 1), MatFactRiesgo(r, j + 1), MatCaracFRiesgo(j + 1).tfactor)
               End If
           Next p
           Exit For
        End If
    Next j
Next i
GenMatRendRiesgo = matd
End Function

Function CalcRend2(ByVal x As Double, ByVal Y As Double, ByVal tfactor As String) As Double
Dim umbral As Double
umbral = DetUmbralF(tfactor)
If Abs(x) <= umbral Or Abs(Y) <= umbral Then
   If x <> 0 And Y <> 0 Then
      CalcRend2 = Y - x
   Else
      CalcRend2 = 0
   End If
ElseIf x > 0 Then
   CalcRend2 = Y / x - 1
ElseIf x < 0 Then
  CalcRend2 = -(Y / x - 1)
End If
End Function

Function CalcRend5(ByVal x0 As Double, ByVal x As Double, ByVal Y As Double, ByVal tfactor As String) As Double
Dim umbral As Double
umbral = DetUmbralF(tfactor)
If Abs(x) <= umbral Or Abs(Y) <= umbral Then
   If x <> 0 And Y <> 0 Then
      If x0 <> 0 Then
         CalcRend5 = (Y - x) / x0
      Else
         CalcRend5 = Y - x
      End If
   Else
      CalcRend5 = 0
   End If
ElseIf x > 0 Then
   CalcRend5 = Y / x - 1
ElseIf x < 0 Then
  CalcRend5 = -(Y / x - 1)
End If
End Function

Function DetUmbralF(ByVal tfactor As String)
If tfactor = "T CAMBIO" Or tfactor = "INDICE" Or tfactor = "UDI" Then
   DetUmbralF = 0.5
Else
   DetUmbralF = 0.005
End If
End Function

Function IncrementoCurva(ByRef curva() As propCurva, ByVal lugar As Long, ByVal inc As Double)
Dim n As Long
Dim i As Long
Dim curva1() As New propCurva

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
curva1 = curva
n = UBound(curva, 1)
If lugar > n Then
IncrementoCurva = curva
Else
 ReDim mata(1 To n, 1 To 3)
 For i = 1 To n
 If i = lugar Then
  mata(i, 1) = curva1(i, 1) + inc
 Else
  mata(i, 1) = curva1(i, 1)
 End If
 mata(i, 2) = curva1(i, 2)
 mata(i, 3) = curva1(i, 3)
 Next i
 IncrementoCurva = mata
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub RutinaControlError(error)
If error = 12 Then

Else

End If
End Sub

Sub AnexarFlujosSwaps2(ByRef mata() As resValFlujoExt, ByRef mate() As resValFlujo, ByVal fecha As Date, ByVal intenc As String, ByVal inter1 As String, ByVal inter2 As String, ByVal tCambio1 As Double, ByVal tCambio2 As Double, ByVal txtmoneda1 As String, ByVal txtmoneda2 As String)
Dim clave1 As Integer
Dim clave2 As Integer
Dim noreg As Long
Dim nr As Long
Dim nort As Long
Dim i As Long
Dim contar As Long
Dim vardum() As New resValFlujoExt
'fecha del calculo de flujos
'mata es la matriz donde se acumulan los flujos de los swaps
'mate es el resultado que se obtuvo para el swap actual
'emision es la clave de la ension
'pata es el tipo de pata del swap

If EsVariableVacia(txtmoneda1) Then
  clave1 = 1
ElseIf txtmoneda1 = "UDI" Then
  clave1 = 2
ElseIf txtmoneda1 = "DOLAR PIP FIX" Then
  clave1 = 10
ElseIf txtmoneda1 = "YEN BM" Then
  clave1 = 7
End If

If EsVariableVacia(txtmoneda2) Then
  clave2 = 1
ElseIf txtmoneda2 = "UDI" Then
  clave2 = 2
ElseIf txtmoneda2 = "DOLAR PIP FIX" Then
  clave2 = 10
ElseIf txtmoneda2 = "YEN BM" Then
  clave2 = 7
End If

If IsArray(mate) Then
   noreg = UBound(mate, 1)
   If noreg <> 0 Then
      nr = UBound(mata, 1)
      nort = nr + noreg
      contar = nr
      ReDim Preserve mata(1 To nort) As resValFlujoExt
      ReDim vardum(1 To noreg)
      For i = 1 To noreg
          vardum(i).f_val = fecha                                        'fecha de valuacion
          vardum(i).c_operacion = mate(i).c_operacion                     'Clave de operación
          vardum(i).t_pata = mate(i).t_pata                               'pata
          vardum(i).intencion = intenc                                    'intencion
          vardum(i).fecha_ini = mate(i).fecha_ini                         'fecha de inicio del flujo
          vardum(i).fecha_fin = mate(i).fecha_fin                         'fecha de vencimiento del flujo
          vardum(i).fecha_desc = mate(i).fecha_desc                       'fecha de descuento del flujo
          vardum(i).p_cupon = mate(i).p_cupon                             'periodo cupon
          vardum(i).saldo_periodo = mate(i).saldo_periodo                 'saldo
          vardum(i).amortizacion = mate(i).amortizacion                   'amortizacion
          vardum(i).tc_aplicar = mate(i).tc_aplicar                       'tasa a pagar en el plazo
          vardum(i).sobretasa = mate(i).sobretasa                         'sobretasa
          vardum(i).int_ini = inter1                                      'intercambio inicial
          vardum(i).int_fin = inter2                                      'intercambio final
          vardum(i).si_paga_int = mate(i).si_paga_int                     'pago de intereses en el periodo
          vardum(i).int_s_saldo = mate(i).int_s_saldo                     'intereses sobre todo el saldo
          vardum(i).int_gen_periodo = mate(i).int_gen_periodo             'intereses generados
          vardum(i).int_acum_periodo = mate(i).int_acum_periodo           'intereses acumulados
          vardum(i).int_pag_periodo = mate(i).int_pag_periodo             'intereses pagados
          vardum(i).int_acum_sig_periodo = mate(i).int_acum_sig_periodo   'intereses acumulados sig periodo
          vardum(i).pago_total = mate(i).pago_total                       'pago total
          vardum(i).t_desc = mate(i).t_desc                               'tasa de descuento
          vardum(i).factor_desc = mate(i).factor_desc                     'factor de descuento
          vardum(i).valor_presente = mate(i).valor_presente               'vardum presente del flujo
          If mate(i).t_pata = "B" Then
             vardum(i).t_cambio = tCambio1                                't cambio
             vardum(i).moneda = clave1                                    'moneda
          Else
             vardum(i).t_cambio = tCambio2                                't cambio
             vardum(i).moneda = clave2                                    'moneda
          End If
      Set mata(nr + i) = vardum(i)
 Next i
End If
End If

End Sub


Function LargoCadenaCoindice(ByVal a As String, ByVal B As String) As Long
Dim largo1 As Long
Dim i As Long

largo1 = Minimo(Len(a), Len(B))
For i = 1 To largo1
If Mid(a, i, 1) <> Mid(B, i, 1) Then
 LargoCadenaCoindice = i - 1
 Exit Function
End If
Next i
LargoCadenaCoindice = largo1
End Function

Function TamañoArch(nomarch)
Dim fso As Object
Dim objFile
Set fso = CreateObject("Scripting.FileSystemObject")
Set objFile = fso.GetFile(nomarch)
TamañoArch = objFile.Size
Set objFile = Nothing
Set fso = Nothing
End Function

Function CalculaEigenv(ByRef matcov() As Double, ByRef matvalp() As Double, ByRef matvec() As Double)
Dim mate() As Double
Dim matd() As Double
Dim matf() As Double
Dim matl() As Double
Dim matg() As Double
Dim s As Double
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim suma As Double
Dim mfriesgo() As Double

noreg = UBound(matcov, 1)
matvalp = HouseHolder(matcov) 'se obtienen los eigenvalores de la matriz de cov
ReDim matc(1 To noreg, 1 To noreg) As Double
ReDim matvec(1 To noreg, 1 To noreg) As Double
For i = 1 To noreg
 For j = 1 To noreg
  matc(j, j) = matvalp(i)
 Next j
 matd = MResta(matcov, matc)
'en teoria matd debe ser singular por lo que al hacerla triangular
'se debe de anular la ultima ecuacion
 mate = MTriangular(matd)
 matf = ExtraeSubMatD(mate, 1, noreg - 1, 1, noreg - 1)
 matg = ExtraeSubMatD(mate, noreg, noreg, 1, noreg - 1)
 ReDim matI(1 To noreg - 1, 1 To 1) As Double
 For p = 1 To noreg - 1
  matI(p, 1) = -matg(p, 1)
 Next p
 matl = MInversa(matf, s)
 mfriesgo = MMult(matl, matI)
 suma = 0
 For j = 1 To noreg - 1
  suma = suma + mfriesgo(j, 1) ^ 2
 Next j
suma = (suma + 1) ^ 0.5
For j = 1 To noreg - 1
matvec(j, i) = mfriesgo(j, 1) / suma
Next j
matvec(noreg, i) = 1 / suma
MensajeProc = "Obteniendo el eigenvector " & i
DoEvents
Next i
ReDim Matroot(1 To noreg, 1 To noreg) As Double
For i = 1 To noreg
If matvalp(i) > 0 Then
 Matroot(i, i) = (matvalp(i)) ^ 0.5
Else
 Matroot(i, i) = 0
End If
Next i
CalculaEigenv = MMult(Matroot, MTranD(matvec))
End Function

Function Truncar(x, dec)
 Truncar = Int(x * 10 ^ dec) / (10 ^ dec)
End Function

Function DV01BonoC(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal tc As Double, ByVal pc As Integer, ByRef curva() As propCurva, ByVal tinterpol As Integer)
Dim precio0 As Double
Dim precio1 As Double
Dim yield As Double

 precio0 = PBonoCurva(fecha, tc, pc, 0, flujos, curva, tinterpol)
 yield = ObtenerYield(fecha, tc, pc, flujos, curva, tinterpol)
 precio1 = PBonoYield(fecha, flujos, tc, yield + 0.0001, pc, 0, "", 0, 0)
 DV01BonoC = precio1 - precio0
End Function

Function DV01BonoY(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal tc As Double, ByVal pc As Integer, ByVal yield As Double) As Double
Dim precio0 As Double
Dim precio1 As Double

   precio0 = PBonoYield(fecha, flujos, tc, yield, pc, 0, "", 0, 0)
   precio1 = PBonoYield(fecha, flujos, tc, yield + 0.0001, pc, 0, "", 0, 0)
   DV01BonoY = precio1 - precio0
End Function

Function DV01IPAB(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal tc0 As Double, ByVal tref As Double, ByVal st As Double, ByVal pc As Integer)
Dim precio0 As Double
Dim precio1 As Double

   precio0 = PIPABV1(fecha, flujos, tc0, tref, st, pc)
   precio1 = PIPABV1(fecha, flujos, tc0, tref, st + 0.0001, pc)
   DV01IPAB = precio1 - precio0
End Function

Function DV01IPABY(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal tc0 As Double, ByVal tr As Double, ByVal yield As Double, ByVal pc As Integer)
Dim precio0 As Double
Dim precio1 As Double

   precio0 = PIPABYield(fecha, flujos, tc0, tr, yield, pc)
   precio1 = PIPABYield(fecha, flujos, tc0, tr, yield + 0.0001, pc)
   DV01IPABY = precio1 - precio0
End Function


Function DV01BondesD(ByVal fecha As Date, ByRef flujos() As estFlujosMD, ByVal intdev As Double, ByVal tr As Double, ByVal st As Double, ByVal pc As Integer)
Dim precio0 As Double
Dim precio1 As Double
   precio0 = PBondesDV1(fecha, flujos, intdev, tr, st, pc)
   precio1 = PBondesDV1(fecha, flujos, intdev, tr, st + 0.0001, pc)
   DV01BondesD = precio1 - precio0
End Function

Function ObtenerYield(ByVal fecha As Date, ByVal tc As Double, ByVal pc As Integer, ByRef flujos() As estFlujosMD, ByRef curva() As propCurva, ByVal tinterpol As Integer)
Dim yield As Double
Dim yield1 As Double
Dim precio0 As Double
Dim precio1 As Double
Dim precio2 As Double
Dim precio3 As Double
Dim inc As Double
Dim dprecio As Double

yield = 0.05
precio0 = PBonoCurva(fecha, tc, pc, 0, flujos, curva, tinterpol)
precio1 = PBonoYield(fecha, flujos, tc, yield, pc, 0, "", 0, 0)
inc = 0.000001
Do While Abs(precio0 - precio1) > 0.000001
   precio2 = PBonoYield(fecha, flujos, tc, yield + inc, pc, 0, "", 0, 0)
   dprecio = (precio2 - precio1) / inc
   yield1 = yield - (precio1 - precio0) / dprecio
   yield = yield1
   precio3 = PBonoYield(fecha, flujos, tc, yield, pc, 0, "", 0, 0)
   precio1 = precio3
Loop
ObtenerYield = yield
End Function

Function TransfFRSplit(ByVal fecha As Date, ByRef matfr() As Variant) As Variant()
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
Dim j As Long
Dim realizar As Boolean
Dim matf() As Variant

noreg1 = UBound(matfr, 1)
noreg2 = UBound(matfr, 2)
'se obtienen las fechas de la matriz
ReDim matfr1(1 To noreg1, 1 To noreg2) As Variant
For i = 1 To UBound(matfr, 1)
  matfr1(i, 1) = matfr(i, 1)
Next i
'primero se determina para que factor de riesgo se va a afectar por un split
For i = 1 To UBound(MatCaracFRiesgo, 1)
    realizar = BuscarFRSplit(MatCaracFRiesgo(i).indFactor)
    If realizar Then
       matf = ConvertFRSplit(fecha, i, matfr)
       For j = 1 To UBound(matfr, 1)
           matfr1(j, i + 1) = matf(j)
       Next j
    Else
        For j = 1 To UBound(matfr, 1)
            matfr1(j, i + 1) = matfr(j, i + 1)
         Next j
    End If
Next i
TransfFRSplit = matfr1
End Function

Function BuscarFRSplit(ByVal clave As String)
Dim encontro As Boolean
Dim i As Long

   encontro = False
   For i = 1 To UBound(MatFRSplit, 1)
   If clave = MatFRSplit(i, 1) Then
      encontro = True
      Exit For
   End If
   Next i
   BuscarFRSplit = encontro
End Function

Function ConvertFRSplit(ByVal fecha As Date, ByVal indice As Long, ByRef matfr() As Variant) As Variant()
Dim contar As Long
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim indice2 As Long

ReDim matf1(1 To 3, 0 To 0) As Variant

contar = 0
For i = 1 To UBound(MatFRSplit, 1)
    If MatCaracFRiesgo(indice, 1) = MatFRSplit(i, 1) Then
       contar = contar + 1
       ReDim matf1(1 To 3, 0 To contar) As Variant
       matf1(1, contar) = MatFRSplit(i, 1)
       matf1(2, contar) = MatFRSplit(i, 2)
       matf1(3, contar) = MatFRSplit(i, 3)
    End If
Next i
matf1 = MTranV(matf1)
noreg = UBound(matfr, 1)
'en matf2 se colocan los escalares que afectan a el vector de datos
ReDim matf2(1 To noreg, 1 To 2) As Variant
For i = 1 To noreg
    matf2(i, 1) = 1
    matf2(i, 2) = 1
Next i
For i = 1 To contar
    For j = 1 To noreg
        If matf1(i, 2) = matfr(j, 1) Then
           matf2(j, 1) = matf1(i, 3)
        End If
    Next j
Next i
indice2 = 0
For i = 1 To noreg
    If fecha = matfr(i, 1) Then
       indice2 = i
       Exit For
    End If
Next i
For i = indice2 To 2 Step -1
    matf2(i - 1, 2) = matf2(i, 2) / matf2(i, 1)
Next i
For i = indice2 + 1 To noreg
    matf2(i, 2) = matf2(i - 1, 2) * matf2(i, 1)
Next i

ReDim matx(1 To noreg) As Variant
For i = 1 To noreg
    matx(i) = matfr(i, indice + 1) * matf2(i, 2)
Next i
ConvertFRSplit = matx
End Function

Sub RutinaVaRHistórico1(ByVal f_factor As Date, _
                        ByVal fechaval As Date, _
                        ByVal noesc As Integer, _
                        ByVal htiempo As Integer, _
                        ByRef matpos() As propPosRiesgo, _
                        ByRef matposmd() As propPosMD, _
                        ByRef matposdiv() As propPosDiv, _
                        ByRef matposswaps() As propPosSwaps, _
                        ByRef matposfwd() As propPosFwd, _
                        ByRef matflswap() As estFlujosDeuda, _
                        ByRef matposdeuda() As propPosDeuda, _
                        ByRef matfldeuda() As estFlujosDeuda, _
                        ByRef matvald0() As resValIns, ByRef matpygd() As Double)

Dim continuar As Boolean

SiValFR = False
 If Not EsArrayVacio(MatFactRiesgo) Then
    Call ValidarCalcVaR(f_factor, noesc, MatFactRiesgo, continuar)
    If continuar Then
         'corre el proceso de calculo de escenarios
         Call CalcEscHist(fechaval, f_factor, htiempo, noesc, 0, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
    End If
 Else
    MensajeProc = "No se han cargado los factores de riesgo"
 End If
End Sub

Sub GuardarEscHist(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef matpos() As propPosRiesgo, ByRef mvalt0() As resValIns, ByRef matpl() As Double, ByRef exito As Boolean)
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtborra As String
Dim txtfiltro As String
Dim i As Integer, j As Integer
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim noreg2 As Integer
Dim txttexto As String
If Not EsArrayVacio(matpl) Then
txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha1
txtborra = txtborra & " AND F_FACTORES = " & txtfecha2
txtborra = txtborra & " AND F_VALUACION = " & txtfecha3
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND ESC_FACTORES = '" & txtescfr & "'"
txtborra = txtborra & " AND NOESC = " & noesc
txtborra = txtborra & " AND HTIEMPO = " & htiempo
ConAdo.Execute txtborra
For i = 1 To UBound(matpos, 1)
    txtcadena = ""
    For j = 1 To UBound(matpl, 1) - 1
        txtcadena = txtcadena & matpl(j, i) & ","
    Next j
    txtcadena = txtcadena & matpl(UBound(matpl, 1), i)
    RGuardarPL.AddNew
    RGuardarPL.Fields("F_POSICION") = CLng(f_pos)                'la f_pos de proceso
    RGuardarPL.Fields("F_FACTORES") = CLng(f_factor)             'la f_pos de proceso
    RGuardarPL.Fields("F_VALUACION") = CLng(f_val)               'la f_pos de proceso
    RGuardarPL.Fields("PORTAFOLIO") = txtport                    'el portafolio
    RGuardarPL.Fields("ESC_FACTORES") = txtescfr                              'el escenario de factores de riesgo
    RGuardarPL.Fields("TIPOPOS") = matpos(i).tipopos             'Tipo de posicion
    RGuardarPL.Fields("FREGISTRO") = CLng(matpos(i).fechareg)     'f_pos de registro
    RGuardarPL.Fields("NOMPOS") = matpos(i).nompos               'nombre de la posicion
    RGuardarPL.Fields("HORAREG") = matpos(i).HoraRegOp
    RGuardarPL.Fields("CPOSICION") = matpos(i).C_Posicion             'clave de la posicion
    RGuardarPL.Fields("COPERACION") = matpos(i).c_operacion           'clave de operacion
    RGuardarPL.Fields("NOESC") = noesc                                'no de escenarios
    RGuardarPL.Fields("HTIEMPO") = htiempo                            'horizonte de tiempo
    RGuardarPL.Fields("VALT0") = mvalt0(i).mtm_sucio                  'valuacion de escenario base
    Call GuardarElementoClob(txtcadena, RGuardarPL, "DATOS")
    RGuardarPL.Update
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando los escenarios historicos " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
 exito = True
Else
 exito = False
End If
End Sub

Sub GuardarEscHist2(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef matpos() As propPosRiesgo, ByRef mvalt0() As resValIns, ByRef matpl() As Double, ByRef exito As Boolean)
Dim txtcadena As String
Dim txtfiltro As String
Dim i As Integer, j As Integer
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim noreg2 As Integer
Dim txttexto As String
If Not EsArrayVacio(matpl) Then

For i = 1 To UBound(matpos, 1)
    txtcadena = ""
    For j = 1 To UBound(matpl, 1) - 1
        txtcadena = txtcadena & matpl(j, i) & ","
    Next j
    txtcadena = txtcadena & matpl(UBound(matpl, 1), i)
    RGuardarPL.AddNew
    RGuardarPL.Fields("F_POSICION") = CLng(f_pos)                 'la fecha de proceso
    RGuardarPL.Fields("F_FACTORES") = CLng(f_factor)               'la fecha de proceso
    RGuardarPL.Fields("F_VALUACION") = CLng(f_val)                 'la fecha de valuacion
    RGuardarPL.Fields("PORTAFOLIO") = txtport                     'el portafolio
    RGuardarPL.Fields("ESC_FACTORES") = txtescfr                  'el escenario de factores de riesgo
    RGuardarPL.Fields("TIPOPOS") = matpos(i).tipopos              'Tipo de posicion
    RGuardarPL.Fields("FREGISTRO") = CLng(matpos(i).fechareg)    'fecha de registro
    RGuardarPL.Fields("NOMPOS") = matpos(i).nompos               'nombre de la posicion
    RGuardarPL.Fields("HORAREG") = matpos(i).HoraRegOp           'nombre de la posicion
    RGuardarPL.Fields("CPOSICION") = matpos(i).C_Posicion        'clave de la posicion
    RGuardarPL.Fields("COPERACION") = matpos(i).c_operacion      'clave de operacion
    RGuardarPL.Fields("NOESC") = noesc                            'no de escenarios
    RGuardarPL.Fields("HTIEMPO") = htiempo                        'horizonte de tiempo
    RGuardarPL.Fields("VALT0") = mvalt0(i).mtm_sucio              'valuacion de escenario base
    Call GuardarElementoClob(txtcadena, RGuardarPL, "DATOS")
    RGuardarPL.Update
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando los escenarios historicos " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
 exito = True
Else
 exito = False
End If
End Sub

Sub GuardarEscHistVR(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal pfwd As Integer, ByRef matpos() As propPosRiesgo, ByRef mvalt0() As resValIns, ByRef mvalt1() As resValIns, ByRef matpl() As Double)
Dim txtcadena As String
Dim txtborra As String
Dim txtfiltro As String
Dim i As Integer, j As Integer
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim noreg2 As Integer
Dim txttexto As String
Dim RInterfIKOS As New ADODB.recordset

txtfiltro = "SELECT * FROM " & TablaPLHistOperVR
RInterfIKOS.Open txtfiltro, ConAdo, 1, 3
For i = 1 To UBound(matpos, 1)
    txtcadena = ""
    For j = 1 To UBound(matpl, 1) - 1
        txtcadena = txtcadena & matpl(j, i) & ","
    Next j
    txtcadena = txtcadena & matpl(UBound(matpl, 1), i)
    RInterfIKOS.AddNew
    RInterfIKOS.Fields(0) = CLng(fecha)                    'la fecha de proceso
    RInterfIKOS.Fields(1) = txtport                        'el portafolio
    RInterfIKOS.Fields(2) = txtescfr                       'el escenario de factores de riesgo
    RInterfIKOS.Fields(3) = matpos(i).tipopos              'clave de la posicion
    RInterfIKOS.Fields(4) = matpos(i).nompos               'clave de la posicion
    RInterfIKOS.Fields(5) = CLng(matpos(i).fechareg)       'fecha de registro
    RInterfIKOS.Fields(6) = matpos(i).HoraRegOp            'clave de la posicion
    RInterfIKOS.Fields(7) = matpos(i).C_Posicion           'clave de la posicion
    RInterfIKOS.Fields(8) = matpos(i).c_operacion          'clave de operacion
    RInterfIKOS.Fields(9) = noesc                          'no de escenarios
    RInterfIKOS.Fields(10) = htiempo                       'horizonte de tiempo
    RInterfIKOS.Fields(11) = pfwd                          'periodo forward
    RInterfIKOS.Fields(12) = mvalt0(i).mtm_sucio           'valuacion en t0
    RInterfIKOS.Fields(13) = mvalt1(i).mtm_sucio          'valuacion en t+1
    Call GuardarElementoClob(txtcadena, RInterfIKOS, "DATOS")
    RInterfIKOS.Update
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando los escenarios historicos " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
RInterfIKOS.Close
End Sub

Sub GuardarEscMont(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByRef matpos() As propPosRiesgo, ByRef mvalt0() As resValIns, ByRef matpl() As Double)
Dim txtcadena As String
Dim i As Integer, j As Integer
Dim txttexto As String
For i = 1 To UBound(matpos, 1)
    txtcadena = ""
    For j = 1 To UBound(matpl, 1) - 1
        txtcadena = txtcadena & matpl(j, i) & ","
    Next j
    txtcadena = txtcadena & matpl(UBound(matpl, 1), i)
    RGuardarPLMont.AddNew
    RGuardarPLMont.Fields("FECHA") = CLng(fecha)              'la fecha de proceso
    RGuardarPLMont.Fields(1) = txtport                        'el portafolio
    RGuardarPLMont.Fields(2) = txtescfr                       'el escenario de factores de riesgo
    RGuardarPLMont.Fields(3) = noesc                          'no de escenarios vol
    RGuardarPLMont.Fields(4) = nosim                          'no de simulaciones
    RGuardarPLMont.Fields(5) = htiempo                        'horizonte de tiempo
    RGuardarPLMont.Fields(6) = matpos(i).C_Posicion           'clave de la posicion
    RGuardarPLMont.Fields(7) = CLng(matpos(i).fechareg)       'fecha de registro
    RGuardarPLMont.Fields(8) = matpos(i).c_operacion          'clave de operacion
    RGuardarPLMont.Fields(9) = mvalt0(i).mtm_sucio                   'valuacion de escenario base
    Call GuardarElementoClob(txtcadena, RGuardarPLMont, "DATOS")
    RGuardarPLMont.Update
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando los escenarios montecarlo " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i

End Sub

Sub ValidarCalcVaR(ByVal fecha As Date, ByVal dvol As Long, ByRef matfriesgo() As Variant, ByRef continuar As Boolean)
Dim indice As Long
indice = BuscarValorArray(fecha, matfriesgo, 1)
If indice = 0 Or indice < dvol Then
MsgBox "no se puede calcular un VaR con los datos disponibles"
   continuar = False
Else
   continuar = True
End If
End Sub

Function DefPortDerivados()
ReDim mata(1 To 5) As Variant
mata(1) = "DERIVADOS DE COBERTURA"
mata(2) = "DERIVADOS DE NEGOCIACION"
mata(3) = "DERIVADOS ESTRUCTURALES"
mata(4) = "DERIVADOS NEGOCIACION RECLASIFICACION"
DefPortDerivados = mata
End Function

 Function Codificar(valor As Variant, accion As Integer) As Variant
     Dim i As Long
     Select Case accion
         Case 0
             For i = 1 To Len(Trim(valor))
                 Mid(valor, i, 1) = Chr(Asc(Mid(valor, i, 1)) - 1)
             Next i
         Case 1
             For i = 1 To Len(Trim(valor))
                 Mid(valor, i, 1) = Chr(Asc(Mid(valor, i, 1)) + 1)
             Next i
     End Select
     Codificar = valor
 End Function

Function DeterminaContraparte(ByVal id_contrap As Integer)
Dim i As Integer
For i = 1 To UBound(MatContrapartes, 1)
    If id_contrap = MatContrapartes(i, 1) Then
       DeterminaContraparte = MatContrapartes(i, 3)
       Exit Function
    End If
Next i
DeterminaContraparte = ""
End Function

Function TiempoProc(ByVal fecha1 As Date, ByVal hora1 As Date, ByVal fecha2 As Date, ByVal hora2 As Date)
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
valor1 = CDbl(fecha1) + CDbl(hora1) - Int(CDbl(hora1))
valor2 = CDbl(fecha2) + CDbl(hora2) - Int(CDbl(hora2))
TiempoProc = valor2 - valor1
End Function

Function EsVariableVacia(x)
On Error GoTo hayerror:
If Not IsNull(x) Then
 If Len(Trim(x)) <> 0 Then
  EsVariableVacia = False
 Else
  EsVariableVacia = True
 End If
Else
 EsVariableVacia = True
End If
Exit Function
hayerror:
EsVariableVacia = True
End Function

Function DetFValxSwapAsociado(ByVal tipopos As Integer, ByVal fechareg As Date, ByVal coperacion As String, ByVal horareg As String, ByVal TOPERACION As Integer) As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim indice As Long
Dim fvalua As String
Dim txttoper As String
Dim rmesa As New ADODB.recordset

     txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE COPERACION = '" & coperacion & "' AND TIPOPOS = " & tipopos
     txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha
     txtfiltro2 = txtfiltro2 & " AND HORAREG = '" & horareg & "'"
     txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
     rmesa.Open txtfiltro1, ConAdo
     noreg = rmesa.Fields(0)
     rmesa.Close
     If noreg <> 0 Then
        rmesa.Open txtfiltro2, ConAdo
        txttoper = rmesa.Fields("FVALUACION")   'tipo de operacion
        rmesa.Close
        indice = BuscarValorArray(txttoper, MatRelSwapsDeuda, 1)
        If indice <> 0 Then
           If TOPERACION = 1 Then
              fvalua = MatRelSwapsDeuda(indice, 3)
           Else
              fvalua = MatRelSwapsDeuda(indice, 2)
           End If
           DetFValxSwapAsociado = fvalua
        Else
           DetFValxSwapAsociado = ""
        End If
     Else
        DetFValxSwapAsociado = ""
     End If
End Function

Function DetCoper(ByVal txtcoper As String) As String
If Left(txtcoper, 8) = "PRIMARIA" And Right(txtcoper, 1) <> "A" And Right(txtcoper, 1) <> "P" Then
   DetCoper = Right(txtcoper, Len(txtcoper) - 9)
ElseIf Left(txtcoper, 8) = "PRIMARIA" And (Right(txtcoper, 1) = "A" Or Right(txtcoper, 1) = "P") Then
   DetCoper = Right(Left(txtcoper, Len(txtcoper) - 2), Len(txtcoper) - 11)
Else
   DetCoper = ""
End If
End Function


Function ConvValor(a) As Double
On Error GoTo error
ConvValor = Val(a)
Exit Function
error:
ConvValor = 0
End Function

Sub ActUHoraUsuario()
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim txtfecha As String
Dim txthora As String
Dim txtcadena As String
   txtfecha = "TO_DATE('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txthora = "TO_DATE('" & Format(Time, "HH:MM:SS") & "','HH24:MI:SS')"
   txtcadena = "UPDATE " & TablaUsuarios & " SET FUREPORTE = " & txtfecha & ", HUREPORTE = " & txthora & " WHERE USUARIO = '" & NomUsuario & "'"
   ConAdo.Execute txtcadena
   txtcadena = "UPDATE " & TablaSesiones & " SET F_ACTIVIDAD = " & txtfecha & ", H_ACTIVIDAD = " & txthora & " WHERE ID_SESION = '" & Id_Sesion & "'"
   ConAdo.Execute txtcadena
Exit Sub
hayerror:
Call TratamientoErrores(Err())
End Sub

Function ExpresarMillPesos(ByVal valor As Double)
If valor >= 1000000 Then
   ExpresarMillPesos = Format(valor / 1000000, "$###,###,###,###,##0.00") & " mdp"
ElseIf valor < 1000000 And valor >= 1000 Then
   ExpresarMillPesos = Format(valor / 1000, "$###,###,###,###,##0.00") & " mil pesos"
 ElseIf valor < 1000 And valor >= 0 Then
   ExpresarMillPesos = Format(valor, "$###,###,###,###,##0.00") & " pesos"
End If
End Function

Function ExpresarMillUSD(ByVal valor As Double)
If valor >= 1000000 Then
   ExpresarMillUSD = Format(valor / 1000000, "$###,###,###,###,##0.00") & " mdd"
ElseIf valor < 1000000 And valor >= 1000 Then
   ExpresarMillUSD = Format(valor / 1000, "$###,###,###,###,##0.00") & " mil dólares"
 ElseIf valor < 1000 And valor >= 0 Then
   ExpresarMillUSD = Format(valor, "$###,###,###,###,##0.00") & " dólares"
End If
End Function

Function GenerarComentarioRE(ByVal fecha As Date, ByVal poscam As Double)
Dim txtcadena As String
Dim cvarm() As Double
Dim cvarmp() As Double
Dim porc() As Double
Dim lim() As Double
Dim consum() As Double
Dim vdif() As Double
Dim i As Integer
Dim fecha1 As Date
Dim noesc As Integer
Dim txtport As String
Dim nconf As Double
ReDim cvarm(1 To 6) As Double
ReDim cvarmp(1 To 6) As Double
ReDim vdif(1 To 6) As Double
ReDim porc(1 To 6) As Double
ReDim lim(1 To 6) As Double
ReDim consum(1 To 6) As Double
noesc = 500
nconf = 0.03
txtport = "TOTAL"

fecha1 = PBD1(fecha, 1, "MX")
CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
cvarmp(1) = Abs(LeerCVaRHist(fecha1, txtport, "MERCADO DE DINERO", nconf, noesc, 1))
cvarmp(2) = Abs(LeerCVaRHist(fecha1, txtport, "MESA DE CAMBIOS", nconf, noesc, 1))
cvarmp(3) = Abs(LeerCVaRHist(fecha1, txtport, "DERIVADOS DE NEGOCIACION", nconf, noesc, 1))
cvarmp(4) = Abs(LeerCVaRHist(fecha1, txtport, "DERIVADOS ESTRUCTURALES", nconf, noesc, 1))
cvarmp(5) = Abs(LeerCVaRHist(fecha1, txtport, "DERIVADOS NEGOCIACION RECLASIFICACION", nconf, noesc, 1))
cvarmp(6) = Abs(LeerCVaRHist(fecha1, txtport, "CONSOLIDADO", nconf, noesc, 1))

cvarm(1) = Abs(LeerCVaRHist(fecha, txtport, "MERCADO DE DINERO", nconf, noesc, 1))
cvarm(2) = Abs(LeerCVaRHist(fecha, txtport, "MESA DE CAMBIOS", nconf, noesc, 1))
cvarm(3) = Abs(LeerCVaRHist(fecha, txtport, "DERIVADOS DE NEGOCIACION", nconf, noesc, 1))
cvarm(4) = Abs(LeerCVaRHist(fecha, txtport, "DERIVADOS ESTRUCTURALES", nconf, noesc, 1))
cvarm(5) = Abs(LeerCVaRHist(fecha, txtport, "DERIVADOS NEGOCIACION RECLASIFICACION", nconf, noesc, 1))
cvarm(6) = Abs(LeerCVaRHist(fecha, txtport, "CONSOLIDADO", nconf, noesc, 1))
For i = 1 To 6
    vdif(i) = cvarm(i) - cvarmp(i)
Next i

porc(1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MD")
porc(2) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MC")
porc(3) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
porc(4) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
porc(5) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
porc(6) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR CON")
For i = 1 To 6
    If porc(i) <> 0 Then
       lim(i) = CapitalNeto * porc(i)
       If lim(i) <> 0 Then consum(i) = cvarm(i) / lim(i)
    Else
       MsgBox "No se encontro limites de CVaR para la fecha " & fecha
    End If
Next i
txtcadena = CompararPosicionesMD2(fecha1, fecha) & " "
If vdif(1) >= 1000000 Then
   txtcadena = txtcadena & "El CVaR de la posición subió a " & ExpresarMillPesos(cvarm(1)) & ", con un consumo de límite de CVaR de " & Format(consum(1), "##0.00%") & "." & vbCrLf & vbCrLf
ElseIf vdif(1) <= -1000000 Then
   txtcadena = txtcadena & "El CVaR de la posición bajó a " & ExpresarMillPesos(cvarm(1)) & ", con un consumo de límite de CVaR de " & Format(consum(1), "##0.00%") & "." & vbCrLf & vbCrLf
Else
   txtcadena = txtcadena & "El CVaR de la posición es de " & ExpresarMillPesos(cvarm(1)) & ", con un consumo de límite de CVaR de " & Format(consum(1), "##0.00%") & "." & vbCrLf & vbCrLf
End If

If Abs(poscam) >= 70000000 * 0.9 Then
   txtcadena = txtcadena & "La posición cambiaria dolarizada es de " & ExpresarMillUSD(poscam) & ", esta posición representa el " & Format(Abs(poscam) / 70000000, "##0.0%") & " de consumo de límite de CAIR. "
   
Else
   txtcadena = txtcadena & "La posición cambiaria dolarizada es de " & ExpresarMillUSD(poscam) & ". "
End If

If vdif(2) >= 1000000 Then
   txtcadena = txtcadena & "El cálculo de CVaR asociado a dicha posición se ubicó en " & ExpresarMillPesos(cvarm(2)) & ", lo que representó un consumo de límite de " & Format(consum(2), "##0.00%") & "." & vbCrLf & vbCrLf
ElseIf vdif(2) <= -1000000 Then
   txtcadena = txtcadena & "El cálculo de CVaR asociado a dicha posición se ubicó en " & ExpresarMillPesos(cvarm(2)) & ", lo que representó un consumo de límite de " & Format(consum(2), "##0.00%") & "." & vbCrLf & vbCrLf
Else
   txtcadena = txtcadena & "El cálculo de CVaR asociado a dicha posición se ubicó en " & ExpresarMillPesos(cvarm(2)) & ", lo que representó un consumo de límite de " & Format(consum(2), "##0.00%") & "." & vbCrLf & vbCrLf
End If
If vdif(3) >= 1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados de Negociación subió a " & ExpresarMillPesos(cvarm(3)) & ", con un consumo de límite de CVaR de " & Format(consum(3), "##0.00%") & "." & vbCrLf & vbCrLf
ElseIf vdif(3) <= -1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados de Negociación bajó a " & ExpresarMillPesos(cvarm(3)) & ", con un consumo de límite de CVaR de " & Format(consum(3), "##0.00%") & "." & vbCrLf & vbCrLf
Else
   txtcadena = txtcadena & "El portafolio de Derivados de Negociación tiene un CVaR de " & ExpresarMillPesos(cvarm(3)) & ", con un consumo de límite de CVaR de " & Format(consum(3), "##0.00%") & "." & vbCrLf & vbCrLf
End If
If vdif(4) >= 1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados Estructurales subió a " & ExpresarMillPesos(cvarm(4)) & ", con un consumo de límite de CVaR de " & Format(consum(4), "##0.00%") & "." & vbCrLf & vbCrLf
ElseIf vdif(4) <= -1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados Estructurales bajó a " & ExpresarMillPesos(cvarm(4)) & ", con un consumo de límite de CVaR de " & Format(consum(4), "##0.00%") & "." & vbCrLf & vbCrLf
Else
   txtcadena = txtcadena & "El portafolio de Derivados Estructurales tiene un CVaR de " & ExpresarMillPesos(cvarm(4)) & ", con un consumo de límite de CVaR de " & Format(consum(4), "##0.00%") & "." & vbCrLf & vbCrLf
End If
If vdif(5) >= 1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados de Negociación por reclasificación subió a " & ExpresarMillPesos(cvarm(5)) & ", con un consumo de límite de CVaR de " & Format(consum(5), "##0.00%") & "." & vbCrLf & vbCrLf
ElseIf vdif(5) <= -1000000 Then
   txtcadena = txtcadena & "El CVaR del portafolio de Derivados de Negociación por reclasificación bajó a " & ExpresarMillPesos(cvarm(5)) & ", con un consumo de límite de CVaR de " & Format(consum(5), "##0.00%") & "." & vbCrLf & vbCrLf
Else
  If cvarm(5) <> 0 Then
     txtcadena = txtcadena & "El portafolio de Derivados de Negociación por reclasificación tiene un CVaR de " & ExpresarMillPesos(cvarm(5)) & ", con un consumo de límite de CVaR de " & Format(consum(5), "##0.00%") & "." & vbCrLf & vbCrLf
  Else
     txtcadena = txtcadena & "El portafolio de Derivados de Negociación por reclasificación no tiene consumo de límite de CVaR." & vbCrLf & vbCrLf
  End If
End If
If vdif(6) >= 1000000 Then
   txtcadena = txtcadena & "El CVaR de la posición Consolidada subió a " & ExpresarMillPesos(cvarm(6)) & ", con un consumo de límite de CVaR de " & Format(consum(6), "##0.00%") & "."
ElseIf vdif(6) <= -1000000 Then
   txtcadena = txtcadena & "El CVaR de la posición Consolidada bajó a " & ExpresarMillPesos(cvarm(6)) & ", con un consumo de límite de CVaR de " & Format(consum(6), "##0.00%") & "."
Else
   txtcadena = txtcadena & "El CVaR de la posición Consolidada es de " & ExpresarMillPesos(cvarm(6)) & ", con un consumo de límite de CVaR de " & Format(consum(6), "##0.00%") & "."
End If
GenerarComentarioRE = txtcadena
End Function

Function CompararPosicionesMD2(ByRef fecha1 As Date, ByRef fecha2 As Date)
Dim mata1() As propPosMD
Dim mata2() As propPosMD
Dim matunion() As New propPosMD
Dim noreg As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim indice As Integer
Dim textpos1 As String
Dim textpos2 As String
Dim textneg1 As String
Dim textneg2 As String
Dim texto1 As String
Dim texto2 As String
Dim nogrp As Integer
Dim notit As String
Dim montonoc As String

'esta es la clave para el filtrado: 2 mesa+teso+pidv
Call LeerYUnirPosMD(fecha1, fecha2, mata1, mata2, matunion, "MERCADO DE DINERO")

noreg = UBound(matunion, 1)
noreg1 = UBound(mata1, 1)  'mesa1
noreg2 = UBound(mata2, 1)  'mesa2

ReDim matresumen(1 To noreg, 1 To 15) As Variant
For i = 1 To noreg
    matresumen(i, 1) = matunion(i).cEmisionMD
    matresumen(i, 2) = matunion(i).tValorMD
    matresumen(i, 3) = matunion(i).emisionMD
    matresumen(i, 4) = matunion(i).serieMD
Next i
'leer el vector de precios de la fecha2
For i = 1 To noreg
    matresumen(i, 5) = 0
    matresumen(i, 8) = 0
    matresumen(i, 11) = 0
    For j = 1 To noreg1 'fecha 1
        If mata1(j).cEmisionMD = matresumen(i, 1) And (mata1(j).Tipo_Mov = 1 Or mata1(j).Tipo_Mov = 6) Then
           matresumen(i, 5) = matresumen(i, 5) + mata1(j).noTitulosMD    'compra en directo
        End If
        If mata1(j).cEmisionMD = matresumen(i, 1) And (mata1(j).Tipo_Mov = 4 Or mata1(j).Tipo_Mov = 7) Then
           matresumen(i, 5) = matresumen(i, 5) - mata1(j).noTitulosMD    'venta en directo
        End If
        If mata1(j).cEmisionMD = matresumen(i, 1) And mata1(j).Tipo_Mov = 2 Then
           matresumen(i, 8) = matresumen(i, 8) + mata1(j).noTitulosMD    'compra en reporto
        End If
       If mata1(j).cEmisionMD = matresumen(i, 1) And mata1(j).Tipo_Mov = 3 Then
           matresumen(i, 11) = matresumen(i, 11) + mata1(j).noTitulosMD    'venta en reporto
       End If
    Next j
    matresumen(i, 6) = 0
    matresumen(i, 9) = 0
    matresumen(i, 12) = 0
    For j = 1 To noreg2   'fecha 2
        If mata2(j).cEmisionMD = matresumen(i, 1) And (mata2(j).Tipo_Mov = 1 Or mata2(j).Tipo_Mov = 6) Then
           matresumen(i, 6) = matresumen(i, 6) + mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And (mata2(j).Tipo_Mov = 4 Or mata2(j).Tipo_Mov = 7) Then
           matresumen(i, 6) = matresumen(i, 6) - mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And mata2(j).Tipo_Mov = 2 Then
           matresumen(i, 9) = matresumen(i, 9) + mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And mata2(j).Tipo_Mov = 3 Then
           matresumen(i, 12) = matresumen(i, 12) + mata2(j).noTitulosMD
        End If
    Next j
matresumen(i, 7) = matresumen(i, 6) - matresumen(i, 5)
matresumen(i, 10) = matresumen(i, 9) - matresumen(i, 8)
matresumen(i, 13) = matresumen(i, 12) - matresumen(i, 11)

Next i
'SE realiza el analisis de la posicion
matvp = LeerVPrecios(fecha2, mindvp)
For i = 1 To noreg
If matresumen(i, 4) <> 0 Then
    indice = BuscarValorArray(matresumen(i, 1), mindvp, 1)
    If indice <> 0 Then
       matresumen(i, 14) = matvp(mindvp(indice, 2)).psucio 'precio sucio
       matresumen(i, 15) = matresumen(i, 7) * matresumen(i, 14) 'marca a mercado
    End If
End If
Next i

nogrp = 12
ReDim Matinst(1 To nogrp, 1 To 5) As Variant
Matinst(1, 1) = "Certificados bursátiles"
Matinst(2, 1) = "Papel CFE"
Matinst(3, 1) = "Papel PEMEX"
Matinst(4, 1) = "bonos BPAG 28"
Matinst(5, 1) = "bonos BPAG 91"
Matinst(6, 1) = "bonos IPAB con cupón semestral"
Matinst(7, 1) = "Bondes D"
Matinst(8, 1) = "bonos M"
Matinst(9, 1) = "Udibonos"
Matinst(10, 1) = "Cetes"
Matinst(11, 1) = "PRLVs"
Matinst(12, 1) = "Bonos USD"

For i = 1 To noreg
    indice = DetermInstGrupo(matresumen(i, 2), matresumen(i, 3), matresumen(i, 4))
    If indice <> 0 Then
       Matinst(indice, 4) = Matinst(indice, 4) + matresumen(i, 7)
       Matinst(indice, 5) = Matinst(indice, 5) + matresumen(i, 15)
    End If
Next i

textpos1 = ""
textpos2 = ""
textneg1 = ""
textneg2 = ""
For i = 1 To nogrp
    If Truncar(Abs(Matinst(i, 4)) / 1000000, 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)) / 1000000, 2), "#,###,###,##0.00") & " millones de"
    ElseIf Truncar(Abs(Matinst(i, 4)) / 1000000, 2) < 1 And Truncar(Abs(Matinst(i, 4)) / 1000, 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)) / 1000, 2), "###,##0.00") & " mil"
    ElseIf Truncar(Abs(Matinst(i, 4)) / 1000, 2) < 1 And Truncar(Abs(Matinst(i, 4)), 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)), 2), "###,##0.00")
    End If
    If Truncar(Abs(Matinst(i, 5)) / 1000000, 2) > 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)) / 1000000, 2), "###,##0.00") & " mdp"
    ElseIf Truncar(Abs(Matinst(i, 5)) / 1000000, 2) < 1 And Truncar(Abs(Matinst(i, 5)) / 1000, 2) >= 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)) / 1000, 2), "###,##0.00") & " mil pesos"
    ElseIf Truncar(Abs(Matinst(i, 5)) / 1000, 2) < 1 And Truncar(Abs(Matinst(i, 5)), 2) >= 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)), 2), "###,##0.00") & " pesos"
    End If
   
    If Matinst(i, 4) > 0 Then
       textpos1 = textpos1 & notit & " títulos ($" & montonoc & ") en " & Matinst(i, 1) & ", "
       textpos2 = textpos2 & "$" & montonoc & " en " & Matinst(i, 1) & ", "
    End If
    If Matinst(i, 4) < 0 Then
       textneg1 = textneg1 & notit & " títulos ($" & montonoc & ") en " & Matinst(i, 1) & ", "
       textneg2 = textneg2 & "$" & montonoc & " en " & Matinst(i, 1) & ", "
    End If
Next i
textpos1 = Trim(textpos1)
If Right(textpos1, 1) = "," Then textpos1 = Left(textpos1, Len(textpos1) - 1)
textneg1 = Trim(textneg1)
If Right(textneg1, 1) = "," Then textneg1 = Left(textneg1, Len(textneg1) - 1)

If Len(textpos1) <> 0 And Len(textneg1) <> 0 Then
   texto1 = "La posición de Mercado de Dinero aumentó " & textpos1 & " y disminuyó " & textneg1 & "."
ElseIf Len(textpos1) <> 0 And Len(textneg1) = 0 Then
   texto1 = "La posición de Mercado de Dinero aumentó " & textpos1 & "."
ElseIf Len(textpos1) = 0 And Len(textneg1) <> 0 Then
   texto1 = "La posición de Mercado de Dinero disminuyó " & textneg1 & "."
End If
textpos2 = Trim(textpos2)
If Right(textpos2, 1) = "," Then textpos2 = Left(textpos2, Len(textpos2) - 1)
textneg2 = Trim(textneg2)
If Right(textneg2, 1) = "," Then textneg2 = Left(textneg2, Len(textneg2) - 1)
If Len(textpos2) <> 0 And Len(textneg2) <> 0 Then
   texto2 = "La posición de Mercado de Dinero aumentó " & textpos2 & " y disminuyó " & textneg2 & "."
ElseIf Len(textpos2) <> 0 And Len(textneg2) = 0 Then
   texto2 = "La posición de Mercado de Dinero aumentó " & textpos2 & "."
ElseIf Len(textpos2) = 0 And Len(textneg2) <> 0 Then
   texto2 = "La posición de Mercado de Dinero disminuyó " & textneg2 & "."
End If
texto2 = Trim(texto2)
If Right(texto2, 1) = "," Then texto2 = Left(texto2, Len(texto2) - 1) & "."
CompararPosicionesMD2 = texto2

End Function


Function DetermInstGrupo(ByVal tv As String, ByVal emision As String, ByVal serie As String)
If (tv = "2U" Or tv = "F" Or tv = "CD" Or tv = "90" Or tv = "91" Or tv = "92" Or tv = "93" Or tv = "94" Or tv = "95" Or tv = "JE") And (emision <> "CFE" And emision <> "PEMEX") Then
   DetermInstGrupo = 1
End If
If emision = "CFE" Then
   DetermInstGrupo = 2
End If
If emision = "PEMEX" Then
   DetermInstGrupo = 3
End If
If tv = "IM" Then
   DetermInstGrupo = 4
End If
If tv = "IQ" Then
   DetermInstGrupo = 5
End If
If tv = "IS" Then
   DetermInstGrupo = 6
End If
If tv = "LD" Then
   DetermInstGrupo = 7
End If
If tv = "M" Then
   DetermInstGrupo = 8
End If
If tv = "S" Then
   DetermInstGrupo = 9
End If
If tv = "BI" Then
   DetermInstGrupo = 10
End If
If tv = "I" Then
   DetermInstGrupo = 11
End If
If tv = "D1" Or tv = "D2" Then
   DetermInstGrupo = 12
End If
End Function

Function DefinirFiltroPosMD(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String)
Dim mata(1 To 3) As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim i As Integer
Dim txtcad1 As String
Dim txtcad2 As String
Dim txtcad As String

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To UBound(MatSQLPort, 1)
    If txtport = MatSQLPort(i, 2) Then
       mata(1) = "SELECT * FROM " & TablaPosMD & " WHERE (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (" & TraducirCadenaSQL(MatSQLPort(i, 3), txtfecha1, 1) & ")"
       mata(2) = "SELECT * FROM " & TablaPosMD & " WHERE (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (" & TraducirCadenaSQL(MatSQLPort(i, 3), txtfecha2, 1) & ")"
       txtcad1 = "SELECT C_EMISION,TV,EMISION,SERIE FROM " & TablaPosMD & " WHERE (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (" & TraducirCadenaSQL(MatSQLPort(i, 3), txtfecha1, 1) & ")"
       txtcad2 = "SELECT C_EMISION,TV,EMISION,SERIE FROM " & TablaPosMD & " WHERE (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (" & TraducirCadenaSQL(MatSQLPort(i, 3), txtfecha2, 1) & ")"
       txtcad = txtcad1 & " UNION " & txtcad2
       mata(3) = "SELECT C_EMISION,TV,EMISION,SERIE FROM " & TablaPosMD & " WHERE (C_EMISION,TV,EMISION,SERIE) IN "
       mata(3) = mata(3) & "(" & txtcad & ")"
       mata(3) = mata(3) & " GROUP BY C_EMISION,TV,EMISION,SERIE ORDER BY C_EMISION,TV,EMISION,SERIE"
       Exit For
    End If
Next i

DefinirFiltroPosMD = mata
End Function

Sub LeerYUnirPosMD(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef mata1() As propPosMD, ByRef mata2() As propPosMD, ByRef matunion() As propPosMD, ByVal txtport As String)
Dim txtfiltro As String
Dim noreg As Long
Dim i As Long
Dim mata() As String
Dim rmesa As New ADODB.recordset

 mata = DefinirFiltroPosMD(fecha1, fecha2, txtport)
 mata1 = LeerBaseMD(mata(1))
 mata2 = LeerBaseMD(mata(2))
 txtfiltro = "SELECT COUNT(*) FROM (" & mata(3) & ")"
 rmesa.Open txtfiltro, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
    rmesa.Open mata(3), ConAdo
    ReDim matun(1 To noreg) As New propPosMD
    For i = 1 To noreg
        With matun(i)
             .cEmisionMD = rmesa.Fields(0)
             .tValorMD = rmesa.Fields(1)
             .emisionMD = rmesa.Fields(2)
             .serieMD = rmesa.Fields(3)
        End With
        rmesa.MoveNext
    Next i
    rmesa.Close
    matunion = matun
 End If
End Sub

Function DefinirCalifEmFP(ByVal fecha As Date, ByVal txtemision As String, ByVal indice As Long, ByRef matv() As propVecPrecios)

Dim valor1 As String
Dim valor2 As String
Dim valor3 As String
Dim valor4 As String
Dim valors1 As Double
Dim valors2 As Double
Dim valors3 As Double
Dim valors4 As Double
Dim escalan As Integer
Dim escala As String
Dim tval As String

If indice <> 0 Then
   tval = matv(indice).tv
   If tval = "BI" Or tval = "IS" Or tval = "LD" Or tval = "M" Or tval = "S" Or tval = "IQ" Or tval = "IM" Or tval = "PI" Or tval = "2U" Then
     escala = "GF"
   Else
     valor1 = matv(indice).calif_sp       'sp
     valor2 = matv(indice).calif_fitch    'fitch
     valor3 = matv(indice).calif_moodys   'moodys
     valor4 = matv(indice).calif_hr       'hr
     If EsCalifCP(valor1, valor2, valor3, valor4) Then
        escala = "CP"
     Else
        valors1 = ConvCalSrt2Num(TradCalifEscSP(TradCalifEscSPGL(valor1)))
        valors2 = ConvCalSrt2Num(TradCalifEscFitch(TradCalifEscFitchGL(valor2)))
        valors3 = ConvCalSrt2Num(TradCalifEscMdy(TradCalifEscMdyGL(valor3)))
        valors4 = ConvCalSrt2Num(TradCalifEscHR(TradCalifEscHRGL(valor4)))
        escala = ConvCalNum2Str(DefinEscMin(valors1, valors2, valors3, valors4, 0))
        escala = ConvEscFP(escala)
        If escala = "ND" Then MensajeProc = txtemision & " sin calificacion  al " & fecha
     End If
   End If
Else
   escala = "ND"
End If
DefinirCalifEmFP = escala
End Function

Function ConvEscFP(escala)
If escala = "AAA" Then
   ConvEscFP = "AAA"
ElseIf escala = "AA+" Then
   ConvEscFP = "AA"
ElseIf escala = "AA" Then
   ConvEscFP = "AA"
ElseIf escala = "AA-" Then
   ConvEscFP = "AA"
ElseIf escala = "A+" Then
   ConvEscFP = "A"
ElseIf escala = "A" Then
   ConvEscFP = "A"
ElseIf escala = "A-" Then
   ConvEscFP = "A"
ElseIf escala = "BBB+" Then
   ConvEscFP = "BBB+"
Else
  ConvEscFP = escala
End If
End Function

Function EsCalifCP(ByVal valor1 As String, ByVal valor2 As String, ByVal valor3 As String, ByVal valor4 As String) As Boolean
If valor1 = "mxA-1" Or valor1 = "mxA-1+" Or valor1 = "mxA-2" Or valor1 = "mxA-3" Then
   EsCalifCP = True
   Exit Function
End If
If valor2 = "F1(mex)" Or valor2 = "F1+(mex)" Or valor2 = "F2(mex)" Or valor2 = "F3(mex)" Then
   EsCalifCP = True
   Exit Function
End If
If valor3 = "MX-1" Or valor3 = "MX-2" Or valor3 = "MX-3" Or valor3 = "MX-4" Then
   EsCalifCP = True
   Exit Function
End If
If valor4 = "HR1" Or valor4 = "HR1+" Or valor4 = "HR2" Or valor4 = "HR2-" Or valor4 = "HR3" Or valor4 = "HR4" Or valor4 = "HR5" Then
   EsCalifCP = True
   Exit Function
End If
EsCalifCP = False
End Function

Function DefinEscMin(ByVal valor1 As Double, ByVal valor2 As Double, ByVal valor3 As Double, ByVal valor4 As Double, opcion As Integer)
DefinEscMin = Maximo(valor1, Maximo(valor2, Maximo(valor3, valor4)))
End Function

Function ConvEscalaNum(ByVal valor As String)
Select Case valor
Case "AAA"
    ConvEscalaNum = 1
Case "AA+"
    ConvEscalaNum = 2
Case "AA"
    ConvEscalaNum = 3
Case "AA-"
    ConvEscalaNum = 4
Case "A+"
    ConvEscalaNum = 5
Case "A"
    ConvEscalaNum = 6
Case "A-"
    ConvEscalaNum = 7
Case "BBB+"
    ConvEscalaNum = 8
Case "BBB"
    ConvEscalaNum = 9
Case "BBB-"
    ConvEscalaNum = 10
Case "BB+"
    ConvEscalaNum = 11
Case "BB"
    ConvEscalaNum = 12
Case "BB-"
    ConvEscalaNum = 13
Case "B+"
    ConvEscalaNum = 14
Case "B"
    ConvEscalaNum = 15
Case "B-"
    ConvEscalaNum = 16
Case "CCC"
    ConvEscalaNum = 17
Case "D"
    ConvEscalaNum = 18
Case Else
    ConvEscalaNum = 0
End Select


End Function

Function ConvNumEscala(ByVal valor As Double)
Select Case valor
Case 1
ConvNumEscala = "AAA"
Case 2
ConvNumEscala = "AA+"
Case 3
ConvNumEscala = "AA"
Case 4
ConvNumEscala = "AA-"
Case 5
ConvNumEscala = "A+"
Case 6
ConvNumEscala = "A"
Case 7
ConvNumEscala = "A-"
Case 8
ConvNumEscala = "BBB+"
Case 9
ConvNumEscala = "BBB"
Case 10
ConvNumEscala = "BBB-"
Case 11
ConvNumEscala = "BB+"
Case 12
ConvNumEscala = "BB"
Case 13
ConvNumEscala = "BB-"
Case 14
ConvNumEscala = "B+"
Case 15
ConvNumEscala = "B"
Case 16
ConvNumEscala = "B-"
Case 17
ConvNumEscala = "CCC"
Case 18
ConvNumEscala = "D"
Case Else
   ConvNumEscala = "ND"
End Select
End Function

Function TradCalifEscSP(ByVal calif As String)
Dim califs As String
'traduce si es escala local a una escala univa
Select Case calif
Case "mxAAA"
   TradCalifEscSP = "AAA"
Case "mxAA+"
   TradCalifEscSP = "AA+"
Case "mxAA"
   TradCalifEscSP = "AA"
Case "mxAA-"
   TradCalifEscSP = "AA-"
Case "mxA+"
   TradCalifEscSP = "A+"
Case "mxA"
   TradCalifEscSP = "A"
Case "mxA-"
   TradCalifEscSP = "A-"
Case "mxBBB+"
   TradCalifEscSP = "BBB+"
Case "mxBBB"
   TradCalifEscSP = "BBB"
Case "mxBBB-"
   TradCalifEscSP = "BBB-"
Case "mxBB+"
   TradCalifEscSP = "BB+"
Case "mxBB"
   TradCalifEscSP = "BB"
Case "mxBB-"
   TradCalifEscSP = "BB-"
Case "mxB+"
   TradCalifEscSP = "B+"
Case "mxB"
   TradCalifEscSP = "B"
Case "mxB-"
   TradCalifEscSP = "B-"
Case "mxCCC"
   TradCalifEscSP = "CCC"
Case "mxCC"
   TradCalifEscSP = "CC"
Case "mxC"
   TradCalifEscSP = "C"
Case "mxD"
   TradCalifEscSP = "D"
Case "ND"
   TradCalifEscSP = "ND"
Case Else
   TradCalifEscSP = calif
End Select
End Function

Function TradCalifEscSPG(ByVal calif As String)
Dim califs As String
'traduce si es escala local a una escala univa
Select Case calif
Case "AAA"
   TradCalifEscSPG = "AAA"
Case "AA+"
   TradCalifEscSPG = "AA+"
Case "AA"
   TradCalifEscSPG = "AA"
Case "AA-"
   TradCalifEscSPG = "AA-"
Case "A+"
   TradCalifEscSPG = "A+"
Case "A"
   TradCalifEscSPG = "A"
Case "A-"
   TradCalifEscSPG = "A-"
Case "BBB+"
   TradCalifEscSPG = "BBB+"
Case "BBB"
   TradCalifEscSPG = "BBB"
Case "BBB-"
   TradCalifEscSPG = "BBB-"
Case "BB+"
   TradCalifEscSPG = "BB+"
Case "BB"
   TradCalifEscSPG = "BB"
Case "BB-"
   TradCalifEscSPG = "BB-"
Case "B+"
   TradCalifEscSPG = "B+"
Case "B"
   TradCalifEscSPG = "B"
Case "B-"
   TradCalifEscSPG = "B-"
Case "CCC"
   TradCalifEscSPG = "CCC"
Case "CC"
   TradCalifEscSPG = "CC"
Case "C"
   TradCalifEscSPG = "C"
Case "D"
   TradCalifEscSPG = "D"
Case "ND"
   TradCalifEscSPG = "ND"
Case Else
   TradCalifEscSPG = calif
End Select
End Function

Function TradCalifEscSPGL(ByVal calif As String)
Dim califs As String
'traduce directamente a escala nacional

Select Case calif
Case "AAA", "AA+", "AA", "AA-"
   califs = "AAA"
Case "A+", "A", "A-"
   califs = "AAA"
Case "BBB+"
   califs = "AA+"
Case "BBB"
   califs = "AA"
Case "BBB-"
   califs = "AA-"
Case "BB+"
   califs = "A+"
Case "BB"
   califs = "A"
Case "BB-"
   califs = "A-"
Case "B+"
   califs = "BB+"
Case "B"
   califs = "BB"
Case "B-"
   califs = "BB-"
Case "CCC"
   califs = "B+"
Case "CC"
    califs = "B"
Case "C"
    califs = "B-"
Case Else
   califs = calif
End Select
TradCalifEscSPGL = califs
End Function


Function TradCalifEscFitch(ByVal calif As String)
Dim califs As String
Select Case calif
Case "AAA(mex)"
   TradCalifEscFitch = "AAA"
Case "AA+(mex)"
   TradCalifEscFitch = "AA+"
Case "AA(mex)"
   TradCalifEscFitch = "AA"
Case "AA-(mex)"
   TradCalifEscFitch = "AA-"
Case "A+(mex)"
   TradCalifEscFitch = "A+"
Case "A(mex)"
   TradCalifEscFitch = "A"
Case "A-(mex)"
   TradCalifEscFitch = "A-"
Case "BBB+(mex)"
   TradCalifEscFitch = "BBB+"
Case "BBB(mex)"
   TradCalifEscFitch = "BBB"
Case "BBB-(mex)"
   TradCalifEscFitch = "BBB-"
Case "BB+(mex)"
   TradCalifEscFitch = "BB+"
Case "BB(mex)"
   TradCalifEscFitch = "BB"
Case "BB-(mex)"
   TradCalifEscFitch = "BB-"
Case "B+(mex)"
   TradCalifEscFitch = "B+"
Case "B(mex)"
   TradCalifEscFitch = "B"
Case "B-(mex)"
   TradCalifEscFitch = "B-"
Case "CCC(mex)"
   TradCalifEscFitch = "CCC"
Case "CC(mex)"
   TradCalifEscFitch = "CC"
Case "C(mex)"
   TradCalifEscFitch = "C"
Case "D(mex)"
   TradCalifEscFitch = "D"
Case Else
   TradCalifEscFitch = "ND"
End Select
End Function

Function TradCalifEscFitchG(ByVal calif As String)
Dim califs As String
Select Case calif
Case "AAA"
   TradCalifEscFitchG = "AAA"
Case "AA+"
   TradCalifEscFitchG = "AA+"
Case "AA"
   TradCalifEscFitchG = "AA"
Case "AA-"
   TradCalifEscFitchG = "AA-"
Case "A+"
   TradCalifEscFitchG = "A+"
Case "A"
   TradCalifEscFitchG = "A"
Case "A-"
   TradCalifEscFitchG = "A-"
Case "BBB+"
   TradCalifEscFitchG = "BBB+"
Case "BBB"
   TradCalifEscFitchG = "BBB"
Case "BBB-"
   TradCalifEscFitchG = "BBB-"
Case "BB+"
   TradCalifEscFitchG = "BB+"
Case "BB"
   TradCalifEscFitchG = "BB"
Case "BB-"
   TradCalifEscFitchG = "BB-"
Case "B+"
   TradCalifEscFitchG = "B+"
Case "B"
   TradCalifEscFitchG = "B"
Case "B-"
   TradCalifEscFitchG = "B-"
Case "CCC"
   TradCalifEscFitchG = "CCC"
Case "CC"
   TradCalifEscFitchG = "CC"
Case "C"
   TradCalifEscFitchG = "C"
Case "D"
   TradCalifEscFitchG = "D"
Case Else
   TradCalifEscFitchG = "ND"
End Select
End Function

Function TradCalifEscFitchGL(ByVal calif As String)
Dim califs As String
Select Case calif
Case "AAA", "AA+", "AA", "AA-"
   califs = "AAA"
Case "A+", "A", "A-"
   califs = "AAA"
Case "BBB+"
   califs = "AA+"
Case "BBB"
   califs = "AA"
Case "BBB-"
   califs = "AA-"
Case "BB+"
   califs = "A+"
Case "BB"
   califs = "A"
Case "BB-"
   califs = "A-"
Case "B+"
   califs = "BB+"
Case "B"
   califs = "BB"
Case "B-"
   califs = "BB-"
Case "CCC"
   califs = "B+"
Case "CC"
    califs = "B"
Case "C"
    califs = "B-"
Case Else
   califs = calif
End Select
TradCalifEscFitchGL = califs
End Function


Function TradCalifEscMdy(ByVal calif As String)
Select Case calif
Case "Aaa.mx"
   TradCalifEscMdy = "AAA"
Case "Aa1.mx"
   TradCalifEscMdy = "AA+"
Case "Aa2.mx"
   TradCalifEscMdy = "AA"
Case "Aa3.mx"
   TradCalifEscMdy = "AA-"
Case "A1.mx"
   TradCalifEscMdy = "A+"
Case "A2.mx"
   TradCalifEscMdy = "A"
Case "A3.mx"
   TradCalifEscMdy = "A-"
Case "Baa1.mx"
   TradCalifEscMdy = "BBB+"
Case "Baa2.mx"
   TradCalifEscMdy = "BBB"
Case "Baa3.mx"
   TradCalifEscMdy = "BBB-"
Case "Ba1.mx"
   TradCalifEscMdy = "BB+"
Case "Ba2.mx"
   TradCalifEscMdy = "BB"
Case "Ba3.mx"
   TradCalifEscMdy = "BB-"
Case "B1.mx"
   TradCalifEscMdy = "B+"
Case "B2.mx"
   TradCalifEscMdy = "B"
Case "B3.mx"
   TradCalifEscMdy = "B-"
Case "Caa1.mx"
   TradCalifEscMdy = "CCC"
Case "Caa2.mx"
   TradCalifEscMdy = "CC"
Case "Caa3.mx"
   TradCalifEscMdy = "C"
Case "D.mx"
   TradCalifEscMdy = "D"
Case Else
   TradCalifEscMdy = "ND"
End Select
End Function

Function TradCalifEscMdyG(ByVal calif As String)
Select Case calif
Case "Aaa"
   TradCalifEscMdyG = "AAA"
Case "Aa1"
   TradCalifEscMdyG = "AA+"
Case "Aa2"
   TradCalifEscMdyG = "AA"
Case "Aa3"
   TradCalifEscMdyG = "AA-"
Case "A1"
   TradCalifEscMdyG = "A+"
Case "A2"
   TradCalifEscMdyG = "A"
Case "A3"
   TradCalifEscMdyG = "A-"
Case "Baa1"
   TradCalifEscMdyG = "BBB+"
Case "Baa2"
   TradCalifEscMdyG = "BBB"
Case "Baa3"
   TradCalifEscMdyG = "BBB-"
Case "Ba1"
   TradCalifEscMdyG = "BB+"
Case "Ba2"
   TradCalifEscMdyG = "BB"
Case "Ba3"
   TradCalifEscMdyG = "BB-"
Case "B1"
   TradCalifEscMdyG = "B+"
Case "B2"
   TradCalifEscMdyG = "B"
Case "B3"
   TradCalifEscMdyG = "B-"
Case "Caa1"
   TradCalifEscMdyG = "CCC"
Case "Caa2"
   TradCalifEscMdyG = "CC"
Case "Caa3"
   TradCalifEscMdyG = "C"
Case "D"
   TradCalifEscMdyG = "D"
Case Else
   TradCalifEscMdyG = "ND"
End Select
End Function

Function TradCalifEscMdyGL(ByVal calif As String)
Dim califs As String
Select Case calif
Case "Aaa", "Aa1", "Aa2", "Aa3"
   califs = "AAA"
Case "A1", "A2", "A3"
   califs = "AAA"
Case "Baa1"
   califs = "AA+"
Case "Baa2"
   califs = "AA"
Case "Baa3"
   califs = "AA-"
Case "Ba1"
   califs = "A+"
Case "Ba2"
   califs = "A"
Case "Ba3"
   califs = "A-"
Case "B1"
   califs = "BB+"
Case "B2"
   califs = "BB"
Case "B3"
   califs = "BB-"
Case "Caa"
   califs = "B+"
Case "Ca"
    califs = "B"
Case "C"
    califs = "B-"
Case Else
   califs = calif
End Select
TradCalifEscMdyGL = califs
End Function


Function TradCalifEscHR(ByVal calif As String)
Dim califs As String
Select Case calif
Case "HR AAA"
   TradCalifEscHR = "AAA"
Case "HR AA+"
   TradCalifEscHR = "AA+"
Case "HR AA"
   TradCalifEscHR = "AA"
Case "HR AA-"
   TradCalifEscHR = "AA-"
Case "HR A+"
   TradCalifEscHR = "A+"
Case "HR A"
   TradCalifEscHR = "A"
Case "HR A-"
   TradCalifEscHR = "A-"
Case "HR BBB+"
   TradCalifEscHR = "BBB+"
Case "HR BBB"
   TradCalifEscHR = "BBB"
Case "HR BBB-"
   TradCalifEscHR = "BBB-"
Case "HR BB+"
   TradCalifEscHR = "BB+"
Case "HR BB"
   TradCalifEscHR = "BB"
Case "HR BB-"
   TradCalifEscHR = "BB-"
Case "HR B+"
   TradCalifEscHR = "B+"
Case "HR B"
   TradCalifEscHR = "B"
Case "HR B-"
   TradCalifEscHR = "B-"
Case "HR C+"
   TradCalifEscHR = "CCC"
Case "HR C"
   TradCalifEscHR = "CC"
Case "HR C-"
   TradCalifEscHR = "C"
Case "HR D"
   TradCalifEscHR = "D"
Case Else
   TradCalifEscHR = "ND"
End Select
End Function

Function TradCalifEscHRG(ByVal calif As String)
Dim califs As String
Select Case calif
Case "HR AAA"
   TradCalifEscHRG = "AAA"
Case "HR AA+"
   TradCalifEscHRG = "AA+"
Case "HR AA"
   TradCalifEscHRG = "AA"
Case "HR AA-"
   TradCalifEscHRG = "AA-"
Case "HR A+"
   TradCalifEscHRG = "A+"
Case "HR A"
   TradCalifEscHRG = "A"
Case "HR A-"
   TradCalifEscHRG = "A-"
Case "HR BBB+"
   TradCalifEscHRG = "BBB+"
Case "HR BBB"
   TradCalifEscHRG = "BBB"
Case "HR BBB-"
   TradCalifEscHRG = "BBB-"
Case "HR BB+"
   TradCalifEscHRG = "BB+"
Case "HR BB"
   TradCalifEscHRG = "BB"
Case "HR BB-"
   TradCalifEscHRG = "BB-"
Case "HR B+"
   TradCalifEscHRG = "B+"
Case "HR B"
   TradCalifEscHRG = "B"
Case "HR B-"
   TradCalifEscHRG = "B-"
Case "HR C+"
   TradCalifEscHRG = "CCC"
Case "HR C"
   TradCalifEscHRG = "CC"
Case "HR C-"
   TradCalifEscHRG = "C"
Case "HR D"
   TradCalifEscHRG = "D"
Case Else
   TradCalifEscHRG = "ND"
End Select
End Function

Function TradCalifEscHRGL(ByVal calif As String)
Dim califs As String
Select Case calif
Case "HR AAA (G)", "HR AA + (G)", "HR AA (G)", "HR AA - (G)"
   califs = "AAA"
Case "HR A + (G)", "HR A (G)", "HR A - (G)"
   califs = "AAA"
Case "HR BBB + (G)"
   califs = "AA+"
Case "HR BBB (G)"
   califs = "AA"
Case "HR BBB - (G)"
   califs = "AA-"
Case "HR BB + (G)"
   califs = "A+"
Case "HR BB (G)"
   califs = "A"
Case "HR BB - (G)"
   califs = "A-"
Case "HR B + (G)"
   califs = "BB+"
Case "HR B (G)"
   califs = "BB"
Case "HR B - (G)"
   califs = "BB-"
Case "HR C + (G)"
   califs = "B+"
Case "HR C (G)"
    califs = "B"
Case "HR C - (G)"
    califs = "B-"
Case Else
   califs = calif
End Select
TradCalifEscHRGL = califs
End Function

Function TradCalifEscHRGL2(ByVal calif As String)
Dim califs As String
Select Case calif
Case "HR AAA (G)", "HR AA + (G)", "HR AA (G)", "HR AA - (G)"
   califs = "AAA"
Case "HR A + (G)", "HR A (G)", "HR A - (G)"
   califs = "AAA"
Case "HR BBB + (G)"
   califs = "AA+"
Case "HR BBB (G)"
   califs = "AA"
Case "HR BBB - (G)"
   califs = "AA-"
Case "HR BB + (G)"
   califs = "A+"
Case "HR BB (G)"
   califs = "A"
Case "HR BB - (G)"
   califs = "A-"
Case "HR B + (G)"
   califs = "BB+"
Case "HR B (G)"
   califs = "BB"
Case "HR B - (G)"
   califs = "BB-"
Case "HR C + (G)"
   califs = "B+"
Case "HR C (G)"
    califs = "B"
Case "HR C - (G)"
    califs = "B-"
Case Else
   califs = calif
End Select
TradCalifEscHRGL2 = califs
End Function


Function ObtPyGTasaMRef(ByVal fecha As Date, ByVal coperacion As String, ByVal nconf As Double)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg0 As Integer
Dim noreg As Integer
Dim nofilas As Long
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim matc() As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT COPERACION,FECHA_F from " & TablaPLEscMW & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "' GROUP BY COPERACION, FECHA_F ORDER BY COPERACION,FECHA_F"
txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg0 = rmesa.Fields(0)
rmesa.Close
If noreg0 <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matoper(1 To noreg0, 1 To 2) As Variant
   For i = 1 To noreg0
       matoper(i, 1) = rmesa.Fields("COPERACION")
       matoper(i, 2) = rmesa.Fields("FECHA_F")
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
ReDim mata(1 To noreg0, 1 To 4) As Variant
Dim matd() As Double
For i = 1 To noreg0
     txtfecha1 = "TO_DATE('" & Format$(matoper(i, 2), "DD/MM/YYYY") & "','DD/MM/YYYY')"
     txtfiltro2 = "SELECT * from " & TablaPLEscMW & " WHERE FECHA = " & txtfecha
     txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & matoper(i, 1) & "' AND FECHA_F = " & txtfecha1 & " ORDER BY GRUPO"
     txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
     rmesa.Open txtfiltro1, ConAdo
     noreg = rmesa.Fields(0)
     rmesa.Close
     rmesa.Open txtfiltro2, ConAdo
     txtcadena = ""
     For j = 1 To noreg
         txtcadena = txtcadena & rmesa.Fields("VECTORPYG").GetChunk(rmesa.Fields("VECTORPYG").ActualSize) & ","
         rmesa.MoveNext
     Next j
     rmesa.Close
     matc = EncontrarSubCadenas(txtcadena, ",")
     ReDim matd(1 To UBound(matc, 1), 1 To 1) As Double
     For j = 1 To UBound(matc, 1)
         matd(j, 1) = matc(j)
     Next j
     mata(i, 4) = CPercentil(nconf, matd, 0, 0)
     mata(i, 1) = matoper(i, 1)          'CLAVE operacion
     mata(i, 2) = matoper(i, 2)          'fecha forward
     mata(i, 3) = mata(i, 2) - fecha     'dias fwd
     MensajeProc = "Obteniendo los percentiles "
     DoEvents
  Next i
ObtPyGTasaMRef = mata
End Function

Function ObtPyGTasaMRef2(ByVal fecha As Date, ByVal coperacion As String)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg0 As Integer
Dim noreg As Integer
Dim nofilas As Long
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim matc() As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT COPERACION,FECHA_F from " & TablaPLEscMW & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "' GROUP BY COPERACION, FECHA_F ORDER BY COPERACION,FECHA_F"
txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg0 = rmesa.Fields(0)
rmesa.Close
If noreg0 <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matoper(1 To noreg0, 1 To 3) As Variant
   For i = 1 To noreg0
       matoper(i, 1) = rmesa.Fields("COPERACION")
       matoper(i, 2) = rmesa.Fields("FECHA_F")
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
For i = 1 To noreg0
     txtfecha1 = "TO_DATE('" & Format$(matoper(i, 2), "DD/MM/YYYY") & "','DD/MM/YYYY')"
     txtfiltro2 = "SELECT * from " & TablaPLEscMW & " WHERE FECHA = " & txtfecha
     txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & matoper(i, 1) & "' AND FECHA_F = " & txtfecha1 & " ORDER BY GRUPO"
     txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
     rmesa.Open txtfiltro1, ConAdo
     noreg = rmesa.Fields(0)
     rmesa.Close
     rmesa.Open txtfiltro2, ConAdo
     txtcadena = ""
     For j = 1 To noreg
         txtcadena = txtcadena & rmesa.Fields("VECTORPYG").GetChunk(rmesa.Fields("VECTORPYG").ActualSize) & ","
         rmesa.MoveNext
     Next j
     matoper(i, 3) = txtcadena
     rmesa.Close
  Next i
ObtPyGTasaMRef2 = matoper
End Function

Function DetermTasaEquilibrio(ByVal fecha As Date, ByVal tipopos As Integer, ByVal fechar As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal op_tasa As Integer)
Dim MatFactR1() As Double
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim mattxt() As String
Dim tasa As Double
Dim tnueva As Double
Dim inc As Double
Dim valor As Double
Dim parval As ParamValPos
Dim MatCurvasT() As Variant
Dim mrvalflujo() As New resValFlujo
Dim deriv As Double
Dim mprecio1() As New resValIns
Dim mprecio2() As New resValIns
Dim exito As Boolean
Dim exito1 As Boolean
Dim txtmsg As String
Dim exito3 As Boolean
Dim txtmsg3 As String
Dim txtmsg0 As String

mattxt = CrearFiltroPosOperPort(tipopos, fechar, txtnompos, horareg, cposicion, coperacion)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
If UBound(matpos, 1) <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg, exito)
   MatFactR1 = CargaFR1Dia(fecha, exito1)
   MatCurvasT = LeerCurvaCompleta(fecha, exito)
   ValExacta = True
   Set parval = DeterminaPerfilVal("VALUACION")
   parval.sicalcdur = False
   'Se carga la estructura de tasas para ese día de la matriz vector tasas
   tasa = 0.08
   inc = 0.000001
   valor = 100000
   Do While Abs(valor) > 0.0001
      If op_tasa = 1 Then
         matposswaps(1).STActiva = tasa
         mprecio1 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
         matposswaps(1).STActiva = tasa + inc
      Else
         matposswaps(1).STPasiva = tasa
         mprecio1 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
         matposswaps(1).STPasiva = tasa + inc
      End If
      mprecio2 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
      valor = mprecio1(1).mtm_sucio
      deriv = (mprecio2(1).mtm_sucio - mprecio1(1).mtm_sucio) / inc
      tnueva = tasa - valor / deriv
      tasa = tnueva
   Loop
End If
DetermTasaEquilibrio = tasa
End Function

