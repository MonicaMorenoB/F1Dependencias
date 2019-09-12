Attribute VB_Name = "Estadisticas"
Option Explicit

Function CalculaRendimientoColumna(ByRef mata() As Double, ByVal ncol As Integer) As Double()
Dim n As Long
Dim i As Long
Dim c() As Double

'esta rutina funciona mejor como funcion,
'para la matriz mata en la columna ncol, calcula
' sus rendimiento y los devuelve en la matriz c
'por el momento solo funciona por columnas
'devuelve como resultado un vector de (n-1) x 1
'se supone que las series de datos estan por columnas

n = UBound(mata, 1)
ReDim c(1 To n - 1, 1 To 1) As Double
For i = 1 To n - 1
    If mata(i, ncol) > 0 Then
       c(i, 1) = (mata(i + 1, ncol) / mata(i, ncol)) - 1
    ElseIf mata(i, ncol) < 0 Then
       c(i, 1) = -((mata(i + 1, ncol) / mata(i, ncol)) - 1)
    ElseIf mata(i, ncol) = 0 Then
       If mata(i + 1, ncol) <> 0 Then
          c(i, 1) = FSigno(mata(i + 1, ncol))
       Else
         c(i, 1) = 0
       End If
    End If
Next i
CalculaRendimientoColumna = c
End Function

Function Fact(n As Long)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'es la funcion factorial, obtenida como una
'formula recursiva
If n = 0 Then
Fact = 1
Else
Fact = n * Fact(n - 1)
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenMuestraNormal(ByVal n As Long) As Double()
Dim i As Long
Dim j As Long
Dim a() As Double
Dim s As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'en esta rutina se genera una muestra normal
' de tamaño n, con la ventaja de que el tiempo
'para el calculo de las v.a. normales es menor
Randomize Timer
ReDim a(1 To n, 1 To 1) As Double
For i = 1 To n
s = 0
For j = 1 To 27
s = s + Rnd
Next j
a(i, 1) = 18 * (s / 27 - 0.5)
Next i
GenMuestraNormal = a
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenMuestraNormal1(ByVal n As Long) As Double()
Dim i As Long
Dim nopares As Long
Dim a() As Double
Dim B() As Double
Dim alea1 As Double
Dim alea2 As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se genera una muestra normal de tamaño n con
'un algoritmo nuevo en el que la muestra normal
'se obtiene a partir de una muestra uniforme
'de tamaño 2
Randomize Timer
If n Mod 2 = 0 Then
nopares = Int(n / 2)
Else
nopares = Int(n / 2) + 1
End If
ReDim a(1 To 2 * nopares) As Double
ReDim B(1 To n, 1 To 1) As Double
For i = 1 To nopares
alea1 = Rnd
alea2 = Rnd
a(2 * i - 1) = Sqr(-2 * Logarit(alea1)) * Cos(2 * Pi * alea2)
a(2 * i) = Sqr(-2 * Logarit(alea1)) * Sin(2 * Pi * alea2)
Next i
For i = 1 To n
 B(i, 1) = a(i)
Next i
GenMuestraNormal1 = B
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenHistograma(ByRef a() As Double, ByVal B As Double, Optional ni As Long)
Dim n As Long
Dim mm As Double
Dim ds As Double
Dim valmax As Double
Dim valmin As Double
Dim i As Long
Dim j As Long
Dim fmaximo As Double
Dim nd As Double


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'a    datos con los que se genera el histograma
'b    factor de escala, tiene efecto en el no
'     de intervalos
'como condicion se pide que a sea una matriz de
'2 dimensiones con solo una columna

n = UBound(a, 1)
'se calcula la media y la desviacion estandar
mm = CMedia2(a, 1, "c")
ds = (CVarianza2(a, 1, "c")) ^ 0.5
'calcula los valores maximo y minimo de la muestra
valmax = a(1, 1)
valmin = a(1, 1)
For i = 1 To n
valmax = Maximo(valmax, a(i, 1))
valmin = Minimo(valmin, a(i, 1))
Next i
'se calcula el no de intervalos ni
If IsNull(ni) Or ni = 0 Then ni = CInt(Minimo(Maximo((valmax - valmin) * n * B, 6), 500))
'se genera los deciles para el histograma
'1  limite inferior
'2  limite superior
'3  valor del decil
'4  valor de la distribucion normal,
'   se dimensiona la matriz con 2 renglones de mas
'   para datos anexos
ReDim mata(1 To ni + 2, 1 To 4) As Variant
fmaximo = 0
nd = 0
For i = 1 To ni
mata(i, 1) = valmin + CDbl(i - 1) * (valmax - valmin) / ni
mata(i, 2) = valmin + CDbl(i) * (valmax - valmin) / ni
mata(i, 3) = 0
Next i

For i = 1 To n
For j = 1 To ni
If mata(j, 1) <= a(i, 1) And a(i, 1) < mata(j, 2) Then
mata(j, 3) = mata(j, 3) + 1
Exit For
End If
Next j
Next i

For i = 1 To ni
    nd = nd + mata(i, 3)
    fmaximo = Maximo(fmaximo, mata(i, 3))
    mata(i, 4) = DNormal(mata(i, 2), mm, ds, 0)
Next i

'se colocan en el 1er renglon adicional, la media
'y la desviación estandar el no de datos y el valor minimo,
'en el primer lugar del segundo renglon
mata(ni + 1, 1) = mm
mata(ni + 1, 2) = ds
mata(ni + 1, 3) = nd
mata(ni + 1, 4) = valmin
mata(ni + 2, 1) = valmax
mata(ni + 2, 2) = fmaximo

GenHistograma = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CPercentil(ByVal x As Double, ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double) As Double
Dim n As Long
Dim i As Long
Dim matempirica() As Variant
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'
'en funcion de una muestra sample se tiene
'que encontrar el valor que w que le
'corresponde al nivel de probabilidad x
'sample tiene que ser una matriz de n x 1
CPercentil = 0
'se genera la matriz con la distribucion empirica
matempirica = GenerarEmpirica(sample, opc, lambda)
n = UBound(matempirica, 1)
'aqui entra la siguiente modificacion
'se busca los valores entre los cuales cae el valor x
'cuando se encuentran se calcula el CPercentil por interpolacion lineal
If x <= matempirica(1, 2) Then
   CPercentil = matempirica(1, 1)
   Exit Function
Else
   For i = 2 To n
      If x <= matempirica(i, 2) And x >= matempirica(i - 1, 2) Then
         CPercentil = matempirica(i - 1, 1) + (matempirica(i, 1) - matempirica(i - 1, 1)) / (matempirica(i, 2) - matempirica(i - 1, 2)) * (x - matempirica(i - 1, 2))
         Exit Function
      End If
   Next i
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CPercentil2(ByVal nconf As Double, ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double, ByVal siorden As Boolean) As Double
Dim n As Long
Dim i As Long
Dim matempirica() As Double
'opc  es el tipo de percentil a calcular
'en funcion de una muestra sample se tiene
'que encontrar el valor que w que le
'corresponde al nivel de probabilidad nconf
'sample tiene que ser una matriz de n nconf 1
CPercentil2 = 0
'se genera la matriz con la distribucion empirica
matempirica = GenerarEmpirica2(sample, opc, lambda, siorden)
n = UBound(matempirica, 1)

If nconf <= matempirica(1, 2) Then
   CPercentil2 = matempirica(1, 1)
   Exit Function
ElseIf nconf >= matempirica(n, 2) Then
   CPercentil2 = matempirica(1, 2)
Else
   For i = 2 To n
   If nconf <= matempirica(i, 2) And nconf > matempirica(i - 1, 2) Then
      CPercentil2 = matempirica(i, 1)
      Exit Function
   End If
Next i
End If
End Function

Function CPercentilCVaR(ByVal nconf As Double, ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double, ByVal siorden As Boolean) As Double
Dim n As Long
Dim i As Long
Dim indice As Long
Dim suma As Double
Dim matempirica() As Double

'calcula el cvar de la distribucion
CPercentilCVaR = 0
If UBound(sample, 1) <> 0 Then
   'se genera la matriz con la distribucion empirica
   matempirica = GenerarEmpirica2(sample, opc, lambda, siorden)
   n = UBound(matempirica, 1)
   If nconf <= matempirica(1, 2) Then
      CPercentilCVaR = matempirica(1, 1)
      Exit Function
   Else
      For i = 2 To n
         If nconf > matempirica(i - 1, 2) And nconf <= matempirica(i, 2) Then
            indice = i
            Exit For
         End If
      Next i
      suma = 0
      If nconf <= 0.5 Then
         For i = 1 To indice
         suma = suma + matempirica(i, 1)
         Next i
         CPercentilCVaR = suma / indice
      Else
         For i = indice To n
         suma = suma + matempirica(i, 1)
         Next i
         CPercentilCVaR = suma / (n - indice + 1)
      End If
   End If
Else
   CPercentilCVaR = 0
End If
End Function

Function GenerarEmpirica(ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double) As Variant()
Dim n As Long
Dim i As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'en funcion de una muestra sample se genera
'la distribución empírica en una matriz de
'n x 2, en la primera columna debe de venir
'
'la muestra debe de venir en una matriz
'unidimensional de n x 1

n = UBound(sample, 1)
ReDim matpivot(1 To n, 1 To 3) As Variant
For i = 1 To n
    matpivot(i, 1) = sample(i, 1)
    If opc = 0 Then
       matpivot(i, 2) = 1 / (n - 1)
    ElseIf opc = 1 Then
       matpivot(i, 2) = lambda ^ (n - 1 - i) * (1 - lambda) / (1 - lambda ^ (n - 1))
    End If
Next i
'se ordenan los datos
matpivot = RutinaOrden(matpivot, 1, SRutOrden)
'a continuación se va acumulando la función
For i = 1 To n
 If i = 1 Then
  matpivot(i, 3) = 0
 Else
  matpivot(i, 3) = matpivot(i, 2) + matpivot(i - 1, 3)
 End If
Next i
contador = n

'se redimensiona una matriz de datos
'donde se van a colocar los datos ya depurados
'se debe de redimensionar la matriz en funcion de los
'puntos distintos que encontro para la distribución empirica
ReDim matempirica(1 To n, 1 To 2) As Variant
For i = 1 To n
 matempirica(i, 1) = matpivot(i, 1)
 matempirica(i, 2) = matpivot(i, 3)
Next i
GenerarEmpirica = matempirica
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenerarEmpirica1(ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double, ByVal siorden As Boolean) As Double()
Dim n As Long
Dim i As Long

'en funcion de una muestra sample se genera
'la distribución empírica en una matriz de
'n x 2, en la primera columna debe de venir
'
'la muestra debe de venir en una matriz
'unidimensional de n x 1
'esta rutina tiene algunas modificaciones
'no se distribuye el peso de los percentiles en n intervalor sino en n-1
n = UBound(sample, 1)
ReDim matpivot(1 To n, 1 To 3) As Variant
For i = 1 To n
matpivot(i, 1) = sample(i, 1)
If opc = 0 Then
   matpivot(i, 2) = 1 / (n - 1)
ElseIf opc = 1 Then
   matpivot(i, 2) = lambda ^ (n - 1 - i) * (1 - lambda) / (1 - lambda ^ (n - 1))
End If
Next i
'se ordenan los datos
If siorden Then matpivot = RutinaOrden(matpivot, 1, SRutOrden)
'a continuación se va acumulando la función
For i = 1 To n
 If i = 1 Then
  matpivot(i, 3) = 0
 Else
  matpivot(i, 3) = matpivot(i, 2) + matpivot(i - 1, 3)
 End If
Next i
contador = n

'se redimensiona una matriz de datos
'donde se van a colocar los datos ya depurados
'se debe de redimensionar la matriz en funcion de los
'puntos distintos que encontro para la distribución empirica
ReDim matempirica(1 To n, 1 To 2) As Double
For i = 1 To n
 matempirica(i, 1) = matpivot(i, 1)
 matempirica(i, 2) = matpivot(i, 3)
Next i
GenerarEmpirica1 = matempirica
End Function

Function GenerarEmpirica2(ByRef sample() As Double, ByVal opc As Integer, ByVal lambda As Double, ByVal siorden As Boolean) As Double()
Dim n As Long
Dim i As Long

'en funcion de una muestra sample se genera
'la distribución empírica en una matriz de
'n x 2, en la primera columna debe de venir
'
'la muestra debe de venir en una matriz
'unidimensional de n x 1
'esta rutina tiene algunas modificaciones
'no se distribuye el peso de los percentiles en n intervalor sino en n-1
n = UBound(sample, 1)
If n > 0 Then
ReDim matpivot(1 To n, 1 To 3) As Variant
For i = 1 To n
    matpivot(i, 1) = sample(i, 1)
    If opc = 0 Then
       matpivot(i, 2) = 1 / n
    ElseIf opc = 1 Then
       matpivot(i, 2) = lambda ^ (n - i) * (1 - lambda) / (1 - lambda ^ (n))
    End If
Next i
'se ordenan los datos
If siorden Then matpivot = RutinaOrden(matpivot, 1, SRutOrden)
'a continuación se va acumulando la función
For i = 1 To n
       matpivot(i, 3) = i / n
Next i
contador = n

'se redimensiona una matriz de datos
'donde se van a colocar los datos ya depurados
'se debe de redimensionar la matriz en funcion de los
'puntos distintos que encontro para la distribución empirica
ReDim matempirica(1 To n, 1 To 2) As Double
For i = 1 To n
 matempirica(i, 1) = matpivot(i, 1)   'valor de la muestra
 matempirica(i, 2) = matpivot(i, 3)   'percentil
Next i
Else
ReDim matempirica(0 To 0, 0 To 0) As Double
End If
GenerarEmpirica2 = matempirica
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ErrVar(ByVal fvol As Double, ByVal ndias1 As Integer, ByVal novolatil As Integer, ByVal indice1 As Integer, ByVal indice2 As Integer, ByVal opvol As Integer, ByVal lambda As Double)
Dim i As Long
Dim matvolatil1() As Double
Dim matvolatil2() As Double
Dim matrends1() As Double
Dim matrends2() As Double
Dim rejilla As MSFlexGrid
Dim ivol As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
' se realiza el calculo de las volatilidades
'esa rutina se debe de dejar aparte como un modulo
'para esto se debe de eliminar toda referencia de
'cualquier objeto
If indice1 = 0 Or indice2 = 0 Then
 Call MostrarMensajeSistema("No se puede realizar el calculo", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Exit Function
End If

Set rejilla = frmVolatilidades.MSFlexGrid4
ivol = BuscarValorArray(fvol, MatFactRiesgo, 1)

'ahora se revisa si se tiene suficiente datos
'hacia atras para realizar los calculos
If ivol = 0 Then
 Call MostrarMensajeSistema("Falta la fecha en la tabla de datos", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Exit Function
End If
If ivol < ndias1 + novolatil + 1 Then
 Call MostrarMensajeSistema("No ha suficientes datos para realizar los calculos", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Exit Function
End If
'como si hay suficientes datos se leen los datos para el calculo
'de volatilidades
 matvolatil1 = ExtSerieFR(MatFactRiesgo, ivol, indice1, ndias1 + novolatil)
 matvolatil2 = ExtSerieFR(MatFactRiesgo, ivol, indice2, ndias1 + novolatil)
'SE CALCULAN los rendimientos, ya sean logaritmicos o
'aritmeticos

 matrends1 = CalculaRendimientoColumna(matvolatil1, 2)
 matrends2 = CalculaRendimientoColumna(matvolatil2, 2)

'Ahora si se procede a calcular medias y volatilidades
'aqui se usan 2 tecnicas: una rutina que obtiene la
'submatriz de la cual se va a obtener la media y la desviacion estandar
'y las funciones para obtener medias y deviaciones estandar
'de un vector que ya estan bien definidas, estos resultados
'a su vez se ponen en un vector llamado MatA
ErrorVarianza = 0
ReDim mata(1 To ndias1, 1 To 2) As Variant
For i = 1 To UBound(matrends1, 1) - novolatil + 1
    mata(i, 1) = GenMedias(ExtSerieAD(matrends1, 1, i, i + novolatil - 1), opvol, lambda)
    mata(i, 2) = GenCovar(ExtSerieAD(matrends1, 1, i, i + novolatil - 1), ExtSerieAD(matrends2, 1, i, i + novolatil - 1), opvol, lambda)
If i > 1 Then ErrorVarianza = ErrorVarianza + (matrends1(i, 1) * matrends2(i, 1) - mata(i - 1, 2) ^ 2) ^ 2
AvanceProc = i / ndias1
Call MostrarMensajeSistema("Calculando Medias y Volatilidades: " & Format(i / ndias1, "###,##0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)

Next i

ErrVar = (ErrorVarianza / ndias1) ^ 0.5


On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function DErrVar(ByVal fvol As Double, ByVal ndias1 As Integer, ByVal novolatil As Integer, ByVal indice1 As Integer, ByVal indice2 As Integer, ByVal opvol As Integer, ByVal lambda As Double)
Dim inch As Double
Dim valor1 As Double
Dim valor2 As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
inch = 0.0000001
valor1 = ErrVar(fvol, ndias1, novolatil, indice1, indice2, opvol, lambda)
valor2 = ErrVar(fvol, ndias1, novolatil, indice1, indice2, opvol, lambda + inch)
DErrVar = (valor2 - valor1) / inch
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub AnalisisKolmogorov(ByRef mata() As Double)
Dim n As Long
Dim i As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
' es en esta rutina donde se calcula el estadistico de kolmogorov
'asi como las distribuciones teoricas y empiricas de la muestra
matorden = GenerarEmpirica(mata, 0, 0.95)
n = UBound(matorden, 1)
For i = 1 To n
matorden(i, 3) = DNormal(matorden(i, 1), media, desvest, 1)
Call MostrarMensajeSistema(media & " " & desvest & " " & matorden(i, 1) & " " & matorden(i, 3), frmProgreso.Label2, 0#, Date, Time, NomUsuario)
Next i
'se procede a calcular el valor del Estadistico de Kolmogorov-Smirnov
Kolmogorov = 0
For i = 1 To n
Kolmogorov = Maximo(Kolmogorov, Abs(matorden(i, 2) - matorden(i, 3)))
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

