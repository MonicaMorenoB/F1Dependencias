Attribute VB_Name = "FuncionesMatriciales"
Option Explicit

Function MSuma(ByRef a() As Double, ByRef B() As Double) As Double()
'objetivo de la funcion: multiplicar 2 arrays y obtener un resultado
'condicion de los datos de entrada: las dos matrices deben de ser de tipo doble
'y el numero de columnas de la segunda matriz debe de ser
'igual al numero de filas de la segunda matriz
'y deben de empezar en 1 en ambas dimensiones
'en caso de no cumplirse la condicion, la rutina devuelve un array de cero dimensiones
'la salida es una matriz que empieza sus dimensiones en el valor 1

Dim ndima1 As Integer
Dim ndima2 As Integer
Dim ndimb1 As Integer
Dim ndimb2 As Integer
Dim c() As Double
Dim i As Integer
Dim j As Integer

'para la suma de matrices
'como condición,
ndima1 = UBound(a, 1)
ndima2 = UBound(a, 2)
ndimb1 = UBound(B, 1)
ndimb2 = UBound(B, 2)
If ndima1 = ndimb1 And ndima2 = ndimb2 Then
 ReDim c(1 To ndima1, 1 To ndima2) As Double
 For i = 1 To ndima1
 For j = 1 To ndima2
 c(i, j) = a(i, j) + B(i, j)
 Next j
 Next i
Else
  ReDim c(0 To 0, 0 To 0) As Double
  MsgBox "no se puede realizar la suma de matrices"
End If
MSuma = c
End Function

Function MResta(ByRef a() As Double, ByRef B() As Double) As Double()
'objetivo de la funcion
Dim ndima1 As Long
Dim ndima2 As Long
Dim ndimb1 As Long
Dim ndimb2 As Long
Dim i As Long
Dim j As Long

Dim c() As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'para la resta de matrices
ndima1 = UBound(a, 1)
ndima2 = UBound(a, 2)
ndimb1 = UBound(B, 1)
ndimb2 = UBound(B, 2)
If ndima1 = ndimb1 And ndima2 = ndimb2 Then
   ReDim c(1 To ndima1, 1 To ndima2) As Double
   For i = 1 To ndima1
   For j = 1 To ndima2
       c(i, j) = a(i, j) - B(i, j)
   Next j
   Next i
Else
 ReDim c(0 To 0, 0 To 0) As Double

End If
  MResta = c
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function MMult(ByRef a() As Double, ByRef B() As Double) As Double()
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim d1 As Long
Dim d2 As Long
Dim c() As Double

'Programa para multiplicacion de matrices
'pide como condicion que l=ll
If IsArray(a) And IsArray(B) Then
 n = UBound(a, 1)
 d1 = UBound(a, 2)
 d2 = UBound(B, 1)
 m = UBound(B, 2)
 If d1 = d2 Then
    ReDim c(1 To n, 1 To m) As Double
    For i = 1 To n
        For j = 1 To m
            For k = 1 To d1
                c(i, j) = c(i, j) + a(i, k) * B(k, j)
            Next k
        Next j
    Next i
 Else
   ReDim c(0 To 0, 0 To 0) As Double
   MsgBox "no se puede realizar el producto"
 End If
Else
   ReDim c(0 To 0, 0 To 0) As Double
End If
 MMult = c
End Function

Function MTranF(ByRef a() As Date) As Date()
Dim n1 As Long
Dim n2 As Long
Dim i As Long
Dim j As Long
'¿Adivinan? es la matriz transpuesta
n1 = UBound(a, 1)
n2 = UBound(a, 2)
ReDim B(1 To n2, 1 To n1) As Date
For i = 1 To n1
   For j = 1 To n2
      B(j, i) = a(i, j)
   Next j
Next i
MTranF = B
End Function

Function MTranV(ByRef a() As Variant) As Variant()
Dim n1 As Long
Dim n2 As Long
Dim B() As Variant
Dim i As Long
Dim j As Long

'¿Adivinan? es la matriz transpuesta
n1 = UBound(a, 1)
n2 = UBound(a, 2)
ReDim B(1 To n2, 1 To n1) As Variant
For i = 1 To n1
   For j = 1 To n2
      B(j, i) = a(i, j)
   Next j
Next i
MTranV = B
End Function

Function MTranD(ByRef a() As Double) As Double()
Dim n1 As Long
Dim n2 As Long
Dim B() As Double
Dim i As Long
Dim j As Long

'¿Adivinan? es la matriz transpuesta
n1 = UBound(a, 1)
n2 = UBound(a, 2)
ReDim B(1 To n2, 1 To n1) As Double
For i = 1 To n1
   For j = 1 To n2
      B(j, i) = a(i, j)
   Next j
Next i
MTranD = B
End Function

Function Choleski(ByRef a() As Double, ByRef m_ind() As Variant, ByRef txtmsg As String, ByRef exito As Boolean) As Double()
Dim n As Integer
Dim m As Integer
Dim x() As Double
Dim i As Integer
Dim j As Integer
Dim kk As Integer
Dim s As Double
Dim no_var_dep As Integer

' este algoritmo devuelve como valor
' una matriz triangular
' superior (EN TEORIA)
' como condicion nos pide que no haya
' ceros en la diagonal
txtmsg = ""
exito = True
ReDim m_ind(1 To 1, 1 To 1)
no_var_dep = 0
n = UBound(a, 1)
m = UBound(a, 2)
If n = m Then
   ReDim x(1 To n, 1 To n) As Double
   x(1, 1) = Sqr(a(1, 1))
   If x(1, 1) <> 0 Then
      For i = 2 To n
          x(i, 1) = a(i, 1) / x(1, 1)
      Next i
   Else
      txtmsg = txtmsg & "el primer  valor de la diagonal es cero"
      exito = False
   End If
   For i = 2 To n
     For j = 2 To i - 1
         s = 0
         For kk = 1 To i - 1
             s = s - x(i, kk) * x(j, kk)
         Next kk
         If x(j, j) <> 0 Then
            x(i, j) = (a(i, j) + s) / x(j, j)
         Else
            x(i, j) = 0
            txtmsg = txtmsg & "un valor de la diagonal es nulo"
            exito = False
         End If
     Next j
     s = 0
     For kk = 1 To i - 1
         s = s - x(i, kk) * x(i, kk)
     Next kk
     If a(i, i) + s > 0 Then
        x(i, i) = Sqr(a(i, i) + s)
     Else
        x(i, i) = 1
        no_var_dep = no_var_dep + 1
        ReDim Preserve m_ind(1 To 1, 1 To no_var_dep)
        m_ind(1, no_var_dep) = i
        txtmsg = txtmsg & " el factor " & i & " esta correlacionado"
        exito = False
     End If
     DoEvents
 Next i
 If no_var_dep <> 0 Then m_ind = MTranV(m_ind)
 If exito Then
    ReDim m_ind(0 To 0, 0 To 0)
    txtmsg = "El proceso finalizo correctamente"
 End If
 'una vez calculada se devuelve el resultado
Else
  ReDim x(0 To 0, 0 To 0) As Double
  exito = False
  txtmsg = "la matriz no es cuadrada"
End If
Choleski = x
End Function

Function DistanciaMat(a, B) As Double
'obtiene un valor de distancia entre 2 matrices numericas
'pueden ser de cualquier tipo
'la condicion es que ambas matrices tengan el mismo numeros de dimensiones
'la salida es un valor que es la raiz de la suma del cuadrado de las diferencias
'entre valores del array

Dim n1 As Integer
Dim m1 As Integer
Dim n2 As Integer
Dim m2 As Integer
Dim s As Double
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'la funcion distancia matricial es una medida de la
'diferencias entre la matriz a y b
If IsArray(a) And IsArray(B) Then
n1 = UBound(a, 1)
m1 = UBound(a, 2)
n2 = UBound(B, 1)
m2 = UBound(B, 2)

 If n1 = n2 And m1 = m2 Then
    s = 0
    For i = 1 To n1
        For j = 1 To m1
            s = s + (a(i, j) - B(i, j)) ^ 2
        Next j
    Next i
    DistanciaMat = Sqr(s)
 Else
    DistanciaMat = -1
 End If
Else
 DistanciaMat = 0
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function UnirVecLng(ByRef a() As Long, ByRef B() As Long) As Long()
Dim n1 As Long
Dim n2 As Long
Dim c() As Long
Dim i As Long

n1 = UBound(a, 1)
n2 = UBound(B, 1)
ReDim c(1 To n1 + n2)
For i = 1 To n1
c(i) = a(i)
Next i
For i = 1 To n2
c(i + n1) = B(i)
Next i
UnirVecLng = c
End Function

Function UnirMatDbl(ByRef a() As Double, ByRef B() As Double, ByVal opc As Integer) As Double()
Dim n11 As Integer
Dim n12 As Integer
Dim n21 As Integer
Dim n22 As Integer
Dim m11 As Integer
Dim m12 As Integer
Dim m21 As Integer
Dim m22 As Integer
Dim alto As Integer
Dim largo As Integer
Dim alto1 As Integer
Dim largo1 As Integer
Dim alto2 As Integer
Dim largo2 As Integer
Dim i As Integer
Dim j As Integer


'los datos de la matriz b se agregan a la derecha
'a la matriz a, para ello se dimensiona una
'nueva matriz
'opc=0 se unen a lo ancho
'opc=1 se unen a lo alto
If IsArray(a) And IsArray(B) Then
 n11 = LBound(a, 1)   'las n's a la primer matriz
 n12 = LBound(a, 2)
 n21 = UBound(a, 1)
 n22 = UBound(a, 2)
 
 m11 = LBound(B, 1)   'las m's a la segunda matriz
 m12 = LBound(B, 2)
 m21 = UBound(B, 1)
 m22 = UBound(B, 2)
 alto1 = n21 - n11 + 1
 largo1 = n22 - n12 + 1
 alto2 = m21 - m11 + 1
 largo2 = m22 - m12 + 1
 
If opc = 0 Then
   alto = Maximo(alto1, alto2)
   largo = largo1 + largo2
   ReDim c(1 To alto, 1 To largo) As Double
   For i = n11 To n21
   For j = n12 To n22
    c(i - n11 + 1, j - n12 + 1) = a(i, j)
   Next j
   Next i
   For i = m11 To m21
   For j = m12 To m22
   c(i - m11 + 1, j + largo1) = B(i, j)
   Next j
   Next i
ElseIf opc = 1 Then
  alto = alto1 + alto2
  largo = Maximo(largo1, largo2)
  ReDim c(1 To alto, 1 To largo) As Double
  For i = n11 To n21
      For j = n12 To n22
          c(i - n11 + 1, j - n12 + 1) = a(i, j)
      Next j
  Next i
  For i = m11 To m21
  For j = m12 To m22
   c(i + alto1, j - m12 + 1) = B(i, j)
  Next j
  Next i
End If
UnirMatDbl = c
End If
End Function


Function UnirMatrices(a, B, ByVal opc As Integer)
'crea una matriz cuyos valores con los valores de las matrices a y b que se unieron
'en una sola matriz en funcion de opc
'si opc = 0 une las matrices a lo ancho
'si opc = 1 une las matrices a lo largo
'el rango de union de la nueva matriz sera la suma de los rangos de a y b
'el segundo rango sera el maximo de los rangos de a y b que no son de union

Dim n11 As Integer
Dim n12 As Integer
Dim n21 As Integer
Dim n22 As Integer
Dim m11 As Integer
Dim m12 As Integer
Dim m21 As Integer
Dim m22 As Integer
Dim alto As Integer
Dim largo As Integer
Dim alto1 As Integer
Dim largo1 As Integer
Dim alto2 As Integer
Dim largo2 As Integer
Dim i As Integer
Dim j As Integer

'los datos de la matriz b se agregan a la derecha
'a la matriz a, para ello se dimensiona una
'nueva matriz
'opc=0 se unen a lo ancho
'opc=1 se unen a lo alto
If IsArray(a) And IsArray(B) Then
 n11 = LBound(a, 1)   'las n's a la primer matriz
 n12 = LBound(a, 2)
 n21 = UBound(a, 1)
 n22 = UBound(a, 2)
 
 m11 = LBound(B, 1)   'las m's a la segunda matriz
 m12 = LBound(B, 2)
 m21 = UBound(B, 1)
 m22 = UBound(B, 2)
 alto1 = n21 - n11 + 1
 largo1 = n22 - n12 + 1
 alto2 = m21 - m11 + 1
 largo2 = m22 - m12 + 1
 
If opc = 0 Then
  alto = Maximo(alto1, alto2)
  largo = largo1 + largo2
 ReDim c(1 To alto, 1 To largo) As Variant
  For i = n11 To n21
  For j = n12 To n22
   c(i - n11 + 1, j - n12 + 1) = a(i, j)
  Next j
  Next i
  For i = m11 To m21
  For j = m12 To m22
  c(i - m11 + 1, j + largo1) = B(i, j)
  Next j
  Next i
ElseIf opc = 1 Then
  alto = alto1 + alto2
  largo = Maximo(largo1, largo2)
  ReDim c(1 To alto, 1 To largo) As Variant
  For i = n11 To n21
      For j = n12 To n22
          c(i - n11 + 1, j - n12 + 1) = a(i, j)
      Next j
  Next i
  
  For i = m11 To m21
  For j = m12 To m22
   c(i + alto1, j - m12 + 1) = B(i, j)
  Next j
  Next i
Else
  ReDim c(0 To 0, 0 To 0) As Variant
End If
UnirMatrices = c
ElseIf Not IsArray(a) And IsArray(B) Then
 UnirMatrices = B
ElseIf IsArray(a) And Not IsArray(B) Then
 UnirMatrices = a
ElseIf Not IsArray(a) And Not IsArray(B) Then
 UnirMatrices = a
End If
End Function

Function UnirTablas(ByRef a() As Variant, ByRef B() As Variant, ByVal opc As Integer) As Variant()
'unirtablas es parecida a la funcion unirmatrices
'solo que esta funcion solo une matrices de tipo variant
'las dimensiones de las matrices de entrada a y b no estan restringidas
'opc = 0  - indica que las matrices se deben unir a lo ancho del array
'opc = 1  - indica que las matrices se deben unir a lo alto del array
'el resultado es una matriz que es la union de las matrices a y b, pero esta matriz
'si restringe sus cotas inferiores en 1



Dim n11 As Integer
Dim n12 As Integer
Dim n21 As Integer
Dim n22 As Integer
Dim m11 As Integer
Dim m12 As Integer
Dim m21 As Integer
Dim m22 As Integer
Dim alto As Integer
Dim largo As Integer
Dim alto1 As Integer
Dim largo1 As Integer
Dim alto2 As Integer
Dim largo2 As Integer
Dim i As Integer
Dim j As Integer
Dim tipov As Integer
Dim tipov2 As Integer


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
tipov = VarType(a)
tipov2 = VarType(B)

'la principal diferencia entre esta funcion y unirmatrices es que
'las dimensiones inferiores de las matrices no pueden ser 0
'los datos de la matriz b se agregan a la derecha
'a la matriz a, para ello se dimensiona una
'nueva matriz
'opc=0 se unen a lo ancho
'opc=1 se unen a lo alto

If EsArrayValAnexar(a) = True And EsArrayValAnexar(B) Then
 n11 = LBound(a, 1)   'las n's a la primer matriz
 n12 = LBound(a, 2)
 n21 = UBound(a, 1)
 n22 = UBound(a, 2)
 
 m11 = LBound(B, 1)   'las m's a la segunda matriz
 m12 = LBound(B, 2)
 m21 = UBound(B, 1)
 m22 = UBound(B, 2)
 alto1 = n21 - n11 + 1
 largo1 = n22 - n12 + 1
 alto2 = m21 - m11 + 1
 largo2 = m22 - m12 + 1
 
If opc = 0 Then
   alto = Maximo(alto1, alto2)
   largo = largo1 + largo2
  ReDim c(1 To largo, 1 To alto) As Variant
  For i = n11 To n21
      For j = n12 To n22
          c(i - n11 + 1, j - n12 + 1) = a(i, j)
      Next j
  Next i
  For i = m11 To m21
  For j = m12 To m22
  c(i - m11 + 1, j + largo1) = B(i, j)
  Next j
  Next i
ElseIf opc = 1 Then
      alto = alto1 + alto2
      largo = Maximo(largo1, largo2)
      ReDim c(1 To alto, 1 To largo) As Variant
      For i = n11 To n21
      For j = n12 To n22
      c(i - n11 + 1, j - n12 + 1) = a(i, j)
  Next j
  Next i
  
  For i = m11 To m21
  For j = m12 To m22
   c(i + alto1, j - m12 + 1) = B(i, j)
  Next j
  Next i
End If
UnirTablas = c
ElseIf EsArrayValAnexar(a) And Not EsArrayValAnexar(B) Then
  UnirTablas = a
ElseIf Not EsArrayValAnexar(a) And EsArrayValAnexar(B) Then
 UnirTablas = B
ElseIf Not EsArrayValAnexar(a) And Not EsArrayValAnexar(B) Then
 UnirTablas = a
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function ObtFactUnicos(ByRef vec() As Variant, ByVal ind As Integer) As Variant()
'obtiene los elementos unidos de la columna ind de una matriz vec() y los devuelve en
'una matriz de n x 1 dimensiones
'primero extrae la columna a analizar en un array de nodatos x 1 y luego procede
'a ordenarlos, despues busca los datos unicos y los agrega en una lista de elementos nuevos
'la condicion para vec() es que empieze su rango de filas en 1
'la condicion para ind es que este entre el rango de las columnas de la matriz vec()


Dim noreg As Long
Dim i As Long
Dim indice As Long
Dim contar As Long

'se obtiene los factores unicos en la columna ind de una matriz y se
'ordenan
noreg = UBound(vec, 1)
ReDim matc(1 To noreg, 1 To 1) As Variant
For i = 1 To noreg
 matc(i, 1) = vec(i, ind)
Next i
matc = RutinaOrden(matc, 1, SRutOrden)
noreg = UBound(matc, 1)
If noreg > 0 Then
   contar = 0
   indice = 1
   ReDim matb(1 To 1, 1 To indice) As Variant
   If noreg > 1 Then
       matb(1, indice) = matc(1, 1)
       For i = indice + 1 To noreg
          If matc(i, 1) <> matb(1, indice) Then
             indice = indice + 1
             ReDim Preserve matb(1 To 1, 1 To indice) As Variant
             matb(1, indice) = matc(i, 1)
          End If
       Next i
   Else
      matb(1, 1) = matc(1, 1)
   End If
   matb = MTranV(matb)
   ReDim matd(1 To UBound(matb, 1), 1 To 1) As Variant
   For i = 1 To UBound(matb, 1)
       matd(i, 1) = matb(i, 1)
   Next i
   ObtFactUnicos = matd
Else
 ReDim matb(0 To 0, 0 To 0) As Variant
 ObtFactUnicos = matb
End If
End Function

Function ObtFechasU(ByRef vec() As Date, ByVal ind As Integer) As Date()
Dim noreg As Long
Dim i As Integer
Dim contar As Integer
Dim indice As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'se obtiene los factores unicos en la columna ind de una matriz y se
'ordenan
noreg = UBound(vec, 1)
ReDim matc(1 To noreg, 1 To 1) As Date
For i = 1 To noreg
 matc(i, 1) = vec(i, ind)
Next i
matc = ROrdenF(matc, 1)
noreg = UBound(matc, 1)
If noreg > 0 Then
   contar = 0
   indice = 1
   ReDim matb(1 To 1, 1 To indice) As Date
   If noreg > 1 Then
       matb(1, indice) = matc(1, 1)
       For i = indice + 1 To noreg
          If matc(i, 1) <> matb(1, indice) Then
             indice = indice + 1
             ReDim Preserve matb(1 To 1, 1 To indice) As Date
             matb(1, indice) = matc(i, 1)
          End If
       Next i
   Else
      matb(1, 1) = matc(1, 1)
   End If
   matb = MTranF(matb)
   ReDim matd(1 To UBound(matb, 1), 1 To 1) As Date
   For i = 1 To UBound(matb, 1)
       matd(i, 1) = matb(i, 1)
   Next i
   ObtFechasU = matd
Else
   ReDim matb(0 To 0, 0 To 0) As Date
   ObtFechasU = matb
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function AgruparFactores(ByRef vec() As Variant, ByVal ind As Integer)
Dim mata() As Variant
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim nocol As Integer
Dim indice As Integer

'se eliminan los renglones repetidos en una matriz de n x m
mata = vec
mata = RutinaOrden(mata, ind, SRutOrden)
noreg = UBound(mata, 1)
nocol = UBound(mata, 2)
If noreg > 0 Then
indice = 1
ReDim matb(1 To nocol, 1 To indice) As Variant
If noreg > 1 Then
For j = 1 To nocol
 matb(j, indice) = mata(1, j)
Next j
 For i = indice + 1 To noreg
 If mata(i, ind) <> matb(ind, indice) Then
  indice = indice + 1
  ReDim Preserve matb(1 To nocol, 1 To indice) As Variant
  For j = 1 To nocol
  matb(j, indice) = mata(i, j)
  Next j
 End If
 Next i
AgruparFactores = MTranV(matb)
Else
 AgruparFactores = mata
End If
Else
 ReDim matb(0 To 0, 0 To 0) As Variant
 AgruparFactores = matb
End If

End Function

Function ConvAVtF(ByRef mata() As Variant) As Date()
Dim n1 As Integer
Dim n2 As Integer
Dim i As Integer
Dim j As Integer

n1 = UBound(mata, 1)
n2 = UBound(mata, 2)
ReDim matb(1 To n1, 1 To n2) As Date
For i = 1 To n1
    For j = 1 To n2
        matb(i, j) = mata(i, j)
    Next j
Next i
ConvAVtF = matb
End Function

Function ConvADtV(ByRef mata() As Double) As Variant()
Dim n1 As Integer
Dim n2 As Integer
Dim i As Integer
Dim j As Integer

n1 = UBound(mata, 1)
n2 = UBound(mata, 2)
ReDim matb(1 To n1, 1 To n2) As Variant
For i = 1 To n1
    For j = 1 To n2
        matb(i, j) = mata(i, j)
    Next j
Next i
ConvADtV = matb
End Function

Function ExtVecMatV(ByRef mata() As Variant, ByVal indice As Integer, ByVal tvect As Integer) As Variant()
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

'de una matriz a obtiene un vector a partir
'de la columna/fila indice en funcion si es un
'tvect=0 vector columna
'tvect=1 vector fila
If tvect = 0 Then
n = UBound(mata, 1)
ReDim Matdos(1 To n, 1 To 1) As Variant
 For i = 1 To n
  Matdos(i, 1) = mata(i, indice)
 Next i
ElseIf tvect = 1 Then
m = UBound(mata, 2)
ReDim Matdos(1 To 1, 1 To m) As Variant
 For i = 1 To m
     Matdos(1, i) = mata(indice, i)
 Next i
End If
ExtVecMatV = Matdos
End Function

Function ExtVecMatD(ByRef mata() As Double, ByVal indice As Integer, ByVal tvect As Integer) As Double()
Dim n As Integer
Dim m As Integer
Dim i As Integer

'de una matriz a obtiene un vector a partir
'de la columna/fila indice en funcion si es un
'tvect=0 vector columna
'tvect=1 vector fila
If tvect = 0 Then
n = UBound(mata, 1)
ReDim Matdos(1 To n, 1 To 1) As Double
 For i = 1 To n
  Matdos(i, 1) = mata(i, indice)
 Next i
ElseIf tvect = 1 Then
m = UBound(mata, 2)
ReDim Matdos(1 To 1, 1 To m) As Double
 For i = 1 To m
     Matdos(1, i) = mata(indice, i)
 Next i
End If
ExtVecMatD = Matdos
End Function

Function ExtraerCapa(mata, ind, limsup)
Dim m As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'Extrae una capa de 2 dimensiones de una
'matriz de 3 dimensiones
'de la matriz MatA, extrae la capa del
'nivel Ind y ademas restringida por LimSup
'OJO EXTRAE UNA CAPA formada por la primera y segunda dimensiones

m = UBound(mata, 2)
ReDim mats(1 To limsup, 1 To m) As Variant
For i = 1 To limsup
For j = 1 To m
mats(i, j) = mata(i, j, ind)
Next j
Next i
ExtraerCapa = mats
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenCorr(ByRef a() As Double, ByVal opc As Integer, ByVal lambda As Double)
'genera la matriz de correlaciones de una serie de datos que se introducen en a
'los datos de entrada son las matriz de datos a, el tipo de covarianza que se quiere calcular
'opc =0  - covarianzas normales
'opc = 1 - covarianzas ponderadas exponencialmente
'lambda  - factor de decaimiento para covarianzas ponderadas exponencialmente

Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim matc() As Double
Dim mats() As Double
Dim matres() As Double
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se genera la matriz de correlaciones de los datos
n = UBound(a, 1)
'se genera la matriz de covarianzas
matc = GenCovar(a, a, opc, lambda)
'se obtiene la matriz s con las desviaciones
'estandar en la diagonal
m = UBound(matc, 1)
ReDim mats(1 To m, 1 To m) As Double
For i = 1 To m
    If matc(i, i) <> 0 Then
       mats(i, i) = 1 / (matc(i, i)) ^ 0.5
    Else
      mats(i, i) = 0
    End If
Next i
matres = MMult(MMult(mats, matc), mats)
GenCorr = matres
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenCovar(ByRef mat1() As Double, ByRef mat2() As Double, ByVal tvol As Integer, ByVal lambda As Double) As Double()
'calcula las covarianzas de dos conjuntos de datos
'mata1 y mata2 deben de ser de datos de doble precision y deben de tener el mismo numero de filas
'las dimensiones de ambas matrices debes empezar en 1
'las filas son el numero de datos historicos
'las columnas son el numero de variables a analizar
'tvol indica el tipo de volatilidad a calcular
'tvol = 0 - se calculan volatilidades ponderadas de forma normal
'tvol = 1 - se calculan volatilidades ponderadas exponencialmente
'lambda   - es el factor de decaimiento para el calculo de volatilidades ponderadas exponencialmente
'lambda>0 y lambda<1
'el resultado es una matriz de covarianzas cuadrada

Dim valor() As Double
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim m0() As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se genera la matriz de covarianzas
'mat1 es la matriz de rendimientos
'si tvol=0 se generas covarianzas del modelo normal
'si tvol=1 se generan las covarianzas del modelo exponencial
'lambda es el factor de decaimiento
noreg1 = UBound(mat1, 1)
noreg2 = UBound(mat2, 1)
If noreg2 = noreg1 Then
   If tvol = 0 Then
  'se generan covarianzas segun modelo normal
      ReDim m0(1 To noreg2, 1 To noreg2) As Double
      For i = 1 To noreg2
          m0(i, i) = 1#
          For j = 1 To noreg2
              m0(i, j) = (m0(i, j) - 1# / noreg2) / (noreg2 - 1)
          Next j
      Next i
      valor = MMult(MMult(MTranD(mat1), m0), mat2)
      If IsArray(valor) Then
         GenCovar = valor
      Else
         ReDim valor(0 To 0, 0 To 0) As Double
         GenCovar = valor
      End If
   ElseIf tvol = 1 Then
 'se generan covarianzas segun modelo exponencial
 'se genera la matriz M0
     ReDim m0(1 To noreg2, 1 To noreg2) As Double
     For i = 1 To noreg2
         m0(i, i) = 1
         For j = 1 To noreg2
             m0(i, j) = m0(i, j) - (1 - lambda) / (1 - lambda ^ noreg2) * lambda ^ (noreg2 - j)
         Next j
     Next i
  'se genera  la matriz lambda
     ReDim MatLambda(1 To noreg2, 1 To noreg2) As Double
     For i = 1 To noreg2
         MatLambda(i, i) = (1 - lambda) / (1 - lambda ^ noreg2) * lambda ^ (noreg2 - i)
     Next i
  'se realiza el cuadruple producto de matrices
     valor = MMult(MMult(MMult(MMult(MTranD(mat1), MTranD(m0)), MatLambda), m0), mat2)
   End If
Else
   ReDim valor(0 To 0, 0 To 0) As Double
   MsgBox "No se pueden calcular las covarianzas"
End If
GenCovar = valor

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function GenRends(ByRef mata() As Double, ByVal htiempo As Integer, fechas() As Date) As Double()
'se genera la matriz de rendimientos en funcion de la matriz
'de factores de riesgo y la matriz de fechas()
'fechas() indica las fechas para las cuales se debe de omitir el calculo de redimiento en
'funcion de una lista negra de escenarios cuyos rendimientos son valores muy grandes

Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer
Dim matb() As Double
Dim estaln As Boolean
'genera rendimientos
  n = UBound(mata, 1)
  m = UBound(mata, 2)
  ReDim matb(1 To n - htiempo, 1 To m) As Double
  For i = 1 To n - htiempo
      For j = 1 To m
          estaln = Esblacklistfr(fechas(i, 1), MatCaracFRiesgo(j).indFactor)
          If estaln Then
             matb(i, j) = 0
          Else
             matb(i, j) = CalcRend2(mata(i, j), mata(i + htiempo, j), MatCaracFRiesgo(j).tfactor)
          End If
      Next j
  Next i
  GenRends = matb
End Function

Sub GenRends3(ByRef mata() As Double, ByVal htiempo As Integer, fechas() As Date, ByRef matr() As Double, ByRef matb() As Integer)
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

Dim estaln As Boolean
'genera rendimientos
  n = UBound(mata, 1)
  m = UBound(mata, 2)
  ReDim matr(1 To n - htiempo, 1 To m) As Double
  ReDim matb(1 To n - htiempo, 1 To m) As Integer
  
  For i = 1 To n - htiempo
      For j = 1 To m
          estaln = Esblacklistfr(fechas(i, 1), MatCaracFRiesgo(j).indFactor)
          matb(i, j) = DetTRendFR(mata(i, j), mata(i + htiempo, j), MatCaracFRiesgo(j).tfactor)
          If estaln Then
             matr(i, j) = 0
          Else
             matr(i, j) = CalcRend2(mata(i, j), mata(i + htiempo, j), MatCaracFRiesgo(j).tfactor)
          End If
      Next j
  Next i
  
End Sub

Function DetTRendFR(ByVal x As Double, ByVal Y As Double, ByVal tfactor As String)
Dim umbral As Double
umbral = DetUmbralF(tfactor)
If Abs(x) <= umbral Or Abs(Y) <= umbral Then
   DetTRendFR = 0
ElseIf x > 0 And Abs(Y) > umbral Then
   DetTRendFR = 1
ElseIf x < 0 And Abs(Y) > umbral Then
  DetTRendFR = 1
End If
End Function

Function GenRends2(ByRef mata() As Double, ByVal htiempo As Integer) As Double()
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer
Dim matb() As Double
Dim estaln As Boolean
'genera rendimientos
  n = UBound(mata, 1)
  m = UBound(mata, 2)
  ReDim matb(1 To n - htiempo, 1 To m) As Double
  For i = 1 To n - htiempo
      For j = 1 To m
           matb(i, j) = CalcRend2(mata(i, j), mata(i + htiempo, j), MatCaracFRiesgo(j).tfactor)
      Next j
  Next i
  GenRends2 = matb
End Function

Function CalcVol5(ByVal x As Double, ByRef mata() As Double, ByVal htiempo As Integer, ByVal tfactor As String) As Double
Dim matrends() As Double
matrends = GenRends4(x, mata, htiempo, tfactor)
CalcVol5 = Abs((CVarianza2(matrends, 1, "c")) ^ 0.5 * x)
End Function

Function GenRends4(ByVal x As Double, ByRef mata() As Double, ByVal htiempo As Integer, ByVal tfactor As String) As Double()
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer
Dim matb() As Double
Dim estaln As Boolean
'genera rendimientos
  n = UBound(mata, 1)
  m = UBound(mata, 2)
  ReDim matb(1 To n - htiempo, 1 To m) As Double
  For i = 1 To n - htiempo
      For j = 1 To m
           matb(i, j) = CalcRend5(x, mata(i, j), mata(i + htiempo, j), tfactor)
      Next j
  Next i
  GenRends4 = matb
End Function


Function MTriangular(ByRef mat() As Double) As Double()
Dim n As Integer
Dim m As Integer
Dim a() As Double
Dim B() As Double
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim kk As Integer
Dim F As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'obtiene la matriz triangular a partir de una matriz rectangular
n = UBound(mat, 1)
m = UBound(mat, 2)
'
ReDim a(1 To n, 1 To m) As Double
ReDim B(1 To m), c(1 To n, 1 To n) As Double
a = mat

For i = 1 To n
j = i 'se empieza en la fila i
25
If a(j, i) <> 0! Then GoTo 30
If j = n Then
 MTriangular = a
 Exit Function
End If
j = j + 1
GoTo 25
30
If i <> j Then a = MPermuta(a, i, j, "f")   'se hace el intercambio de las filas i y j
For p = i + 1 To n
 F = a(p, i)
 For kk = i To m
  a(p, kk) = a(p, kk) - a(i, kk) * F / a(i, i)
 Next kk
Next p
Next i
MTriangular = a
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function MInversa(ByRef mat() As Double, ByRef s As Double) As Double()
Dim m As Integer
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim kk As Integer
Dim a() As Double
Dim B() As Double
Dim F As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'INVERSION DE UNA MATRIZ BASADA EN EL METODO DE GAUSS-JORDAN
n = UBound(mat, 1)
'se crean las matrices auxiliares
ReDim a(1 To n, 1 To 2 * n) As Double
ReDim B(1 To 2 * n), c(1 To n, 1 To n) As Double
'b matriz pivote
For i = 1 To n
    a(i, i + n) = 1#
    For j = 1 To n
        a(i, j) = mat(i, j)
    Next j
Next i
'el proceso se hace con contadores
i = 1
20
j = i 'se empieza en la fila i
25
If a(j, i) <> 0! Then GoTo 30
If j = n Then GoTo 50   'la matriz no tiene inversa
j = j + 1
GoTo 25
30

For kk = 1 To 2 * n    'se hace la transposicion de las filas i y j
    B(kk) = a(i, kk)
    a(i, kk) = a(j, kk)
    a(j, kk) = B(kk)
Next kk
For p = 1 To n
If i = p Then GoTo 45
F = a(p, i)
For m = i To 2 * n
a(p, m) = a(p, m) - a(i, m) * F / a(i, i)
Next m
45  Next p
If i = n Then GoTo 55
i = i + 1: GoTo 20
50
 MensajeProc = "NO EXISTE LA INVERSA DE LA MATRIZ..."
 s = 0
GoTo 80
55
For i = 1 To n
    For j = n + 1 To 2 * n
        a(i, j) = a(i, j) / a(i, i)
    Next j
Next i
'para obtener el valor del determinante
s = 1#
For i = 1 To n
    s = s * a(i, i)
Next i
For i = 1 To n
    For j = 1 To n
        c(i, j) = a(i, j + n)
    Next j
Next i
MInversa = c
'aqui es donde se ponen los datos de la matriz inversa
80

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function MIdentidad(ByVal i As Integer) As Double()
Dim kk As Integer
Dim j As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
' en funcion del valor i se devuelve una matriz identidad de ixi
ReDim mat(1 To i, 1 To i) As Double

For kk = 1 To i
For j = 1 To i
mat(kk, j) = 0#
Next j
mat(kk, kk) = 1#
Next kk
MIdentidad = mat
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function CargaFR1Dia(fecha As Date, ByRef exito As Boolean) As Double()
Dim matfr() As Variant
Dim j As Integer
Dim indice As Integer
Dim SiValoresCero As Boolean

SiValoresCero = False
If Not EsArrayVacio(MatFechasFR) Then
   indice = BuscarValorArray(fecha, MatFechasFR, 1)
   If indice <> 0 Then
      Call CrearMatFRiesgo2(fecha, fecha, matfr, "", exito)
      ReDim mt(1 To NoFactores, 1 To 1) As Double
      For j = 1 To NoFactores
          If matfr(1, j + 1) = 0 Or IsNull(matfr(1, j + 1)) Then
             SiValoresCero = True
          Else
             mt(j, 1) = matfr(1, j + 1)
          End If
      Next j
      CargaFR1Dia = mt
      exito = True
   Else
      ReDim matc(0 To 0, 0 To 0) As Double
      CargaFR1Dia = matc
      MsgBox "no hay factores de riesgo para esta fecha"
      exito = False
   End If
End If
End Function

Function CargaValExtDia(fecha As Date) As Double()
Dim exito As Boolean
Dim mata() As Variant
Dim n As Integer
Dim j As Integer
Dim SiValoresCero As Boolean

'se cargan las tasas a una matriz pivote
'esta rutina debe de tener acceso a la tabla de datos
'esta rutina esta muy ligada a la estructura de datos
SiValoresCero = False
 Call CrearMatFRiesgo2(fecha, fecha, mata, "", exito)
n = UBound(mata, 2)
If n > 1 Then
   ReDim mt(1 To n - 1, 1 To 1) As Double
   For j = 1 To NoFactores
   If mata(1, j + 1) = 0 Or IsNull(mata(1, j + 1)) Then
      SiValoresCero = True
   Else
      mt(j, 1) = mata(1, j + 1)
   End If
Next j

Else
 ReDim mt(0 To 0, 0 To 0) As Double
End If
CargaValExtDia = mt
'If SiValoresCero Then Call MostrarMensajeSistema("Hay valores Cero en el Vector de Tasas", frmprogreso.label2, 0.0001)
End Function

Function ExtFRMatFR(ByVal ind As Integer, ByRef mata() As Variant) As Double()
Dim n As Integer
Dim j As Integer
Dim SiValoresCero As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan las tasas a una matriz pivote
'sin la fecha
n = UBound(mata, 2)
If n > 1 Then
ReDim mt(1 To n - 1, 1 To 1) As Double
For j = 1 To NoFactores
If mata(ind, j + 1) = 0 Or IsNull(mata(ind, j + 1)) Then
   SiValoresCero = True
Else
   mt(j, 1) = CDbl(mata(ind, j + 1))
End If
Next j

Else
   ReDim mt(0 To 0, 0 To 0) As Double
End If
ExtFRMatFR = mt
'If SiValoresCero Then Call MostrarMensajeSistema("Hay valores Cero en el Vector de Tasas", frmprogreso.label2, 0.0001)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaTasasPIP(ByVal i As Integer)
Dim SiValoresCero As Boolean
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan las tasas a una matriz pivote
SiValoresCero = False
If i <> 0 Then
ReDim mt(1 To NoFactores, 1 To 1) As Double
For j = 1 To NoFactores
If MatCurvaPIP(i, j + 1) = 0 Then SiValoresCero = True
mt(j, 1) = MatCurvaPIP(i, j + 1)
Next j
CargaTasasPIP = mt
Else
CargaTasasPIP = 0
End If
'If SiValoresCero Then Call MostrarMensajeSistema( "Hay valores Cero en el Vector de Tasas"
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ExtraerSMatFR(ByVal ind As Integer, ByVal noreg As Integer, ByRef mata() As Variant, ByVal cfecha As Boolean, ByRef sifr() As Boolean) As Variant()
Dim n As Integer
Dim i As Integer
Dim j As Integer
Dim SiValorCero As Boolean

'se carga en memoria los datos para su uso
'en funcion de la columna, el no de datos y la fecha
'en esta matriz solo se cargan los factores de
'riesgo basicos
'ind        es el indice donde se ubica la fecha ind analizar
'noreg      es el no de datos ind extraer hacia atras
'mata       es la matriz de datos
'cfecha     indica si a la matriz se le anexo la fecha
'sifr       indica que factores son factores de riesgo
n = UBound(mata, 2)
SiValorCero = False
If cfecha Then
ReDim matb(1 To noreg, 1 To n) As Variant
   For i = 1 To noreg
       matb(i, 1) = mata(ind - noreg + i, 1)     'la fecha
       For j = 1 To n - 1
           If sifr(j) Then
              If Not IsNull(mata(ind - noreg + i, j + 1)) Then
                 matb(i, j + 1) = mata(ind - noreg + i, j + 1)
              Else
                 matb(i, j + 1) = 0
              End If
              If matb(i, j + 1) = 0 Then SiValorCero = True
           Else
              matb(i, j + 1) = 0
           End If
       Next j
   Next i
ElseIf Not cfecha Then
'por alguna extraña razon tiene 4 escenarios en la tabla
ReDim matb(1 To noreg, 1 To n - 1) As Variant
   For i = 1 To noreg
       For j = 1 To n - 1
           If sifr(j) Then
              If Not EsVariableVacia(mata(ind - noreg - 1 + i, j + 1)) Then
                 matb(i, j) = mata(ind - noreg + i, j + 1)
               Else
                  matb(i, j) = 0
                  SiValorCero = True
               End If
            Else
               matb(i, j) = 0
            End If
        Next j
    Next i
End If
ExtraerSMatFR = matb
End Function

Function ExtraeRangoMat(ByVal a As Integer, ByVal nodat As Integer, ByRef vect() As Variant)
Dim n As Integer
Dim SiValorCero As Boolean
Dim i As Integer
Dim j As Integer


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'de la matriz vect se extrae otro rango
n = UBound(vect, 2)
SiValorCero = False
ReDim mata(1 To nodat, 1 To n) As Variant
For i = 1 To nodat
 For j = 1 To n
  If Not IsNull(vect(a - nodat + i, j)) Then
   mata(i, j) = vect(a - nodat + i, j)
  Else
   mata(i, j) = 0
  End If
 Next j
Next i
ExtraeRangoMat = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function NormAleato(ByRef mata() As Double) As Double()
Dim n As Integer
Dim i As Integer
Dim sumat As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'normaliza una muestra aleatoria de numeros para que no se aleje
'mucho de valores estimados
 n = UBound(mata, 1)
 ReDim matb(1 To n, 1 To 1) As Double
 sumat = 0
 For i = 1 To n
     sumat = sumat + mata(i, 1)
 Next i
 If Abs(sumat) > 1 Then
    For i = 1 To n
        matb(i, 1) = mata(i, 1) / Abs(sumat)
    Next i
    NormAleato = matb
 Else
    NormAleato = mata
 End If
 
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CalculaLambdaOptima(ByVal fecha As Date, ByVal indice As Integer, ByVal ndias1 As Integer, ByVal nodvol As Integer) As Double()
Dim i As Integer
Dim ivol As Integer
Dim matrends1() As Double
Dim matrends2() As Double
Dim matvolatil1() As Double
Dim nconf As Double
Dim NivelCritico As Double
Dim indice1 As Integer
Dim novolatil As Integer
Dim OpcionVol As Integer
Dim lambda As Double
Dim opvol As Integer

'esta funcion trata de calcular la lambda opt
NivelCritico = NormalInv(nconf)
If indice1 = 0 Then
 Call MostrarMensajeSistema("no se puede realizar el calculo", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Exit Function
End If

ivol = BuscarValorArray(fecha, MatFactRiesgo, 1)

'ahora se revisa si se tiene suficiente datos
'hacia atras para realizar los calculos
If ivol = 0 Then
 Call MostrarMensajeSistema("Falta la fecha en la tabla de datos, se hara con la ultima fecha de la tabla", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 ivol = UBound(MatFactRiesgo, 1)
End If
frmVolatilidades.Combo4.Text = MatFactRiesgo(ivol, 1)

If ivol < ndias1 + novolatil + 1 Then
 Call MostrarMensajeSistema("No ha suficientes datos para realizar los calculos", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Exit Function
End If
'como si hay suficientes datos se leen los datos para el calculo
'de volatilidades
 matvolatil1 = ExtSerieFR(MatFactRiesgo, ivol, indice1, ndias1 + novolatil)
 
'SE CALCULAN los rendimientos, ya sean logaritmicos o
'aritmeticos
 matrends1 = CalculaRendimientoColumna(matvolatil1, 2)

'Ahora si se procede a calcular medias y volatilidades
'aqui se usan 2 tecnicas: una rutina que obtiene la
'submatriz de la cual se va a obtener la media y la desviacion estandar
'y las funciones para obtener medias y deviaciones estandar
'de un vector que ya estan bien definidas, estos resultados
'a su vez se ponen en un vector llamado MatA
ErrorVarianza = 0
ReDim mata(1 To ndias1, 1 To 2) As Variant
For i = 1 To UBound(matrends1, 1) - novolatil + 1
 mata(i, 1) = GenMedias(ExtSerieAD(matrends1, 1, i, i + novolatil - 1), OpcionVol, lambda)
 mata(i, 2) = GenCovar(ExtSerieAD(matrends1, 1, i, i + novolatil - 1), ExtSerieAD(matrends2, 1, i, i + novolatil - 1), opvol, lambda)
If i > 1 Then ErrorVarianza = ErrorVarianza + (matrends1(i, 1) * matrends2(i, 1) - mata(i - 1, 2) ^ 2) ^ 2
AvanceProc = i / ndias1
Call MostrarMensajeSistema("Calculando Medias y Volatilidades: " & Format(i / ndias1, "###,##0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)

Next i

ErrorVarianza = (ErrorVarianza / ndias1) ^ 0.5
'Se muestran los rendimientos mas las volatilidades en pantalla

'en funcion de las volatilidades se
'calculan los limites
NoAciertos = 0
'Call MostrarMensajeSistema( NoAciertos / (ndias1 - 1) * 100
'se grafican las volatilidades en un gráfico de columnas

End Function

Sub BuscarRendExt(mata)
Dim n1 As Integer
Dim n2 As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
n1 = UBound(mata, 1)
n2 = UBound(mata, 2)
For i = 1 To n1
For j = 1 To n2
If Abs(mata(i, j)) > 1 Then
 Call MostrarMensajeSistema("Incrementos excesivos: " & Format(mata(i, j), "#0.000000 %") & " Factor: " & MatCaracFRiesgo(j).indFactor, frmProgreso.Label2, 0, Date, Time, NomUsuario)
End If
Next j
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function MPermFilaCol(ByRef mata() As Double, ByVal v1 As Integer, ByVal v2 As Integer)
Dim n As Integer
Dim i As Integer
Dim matb() As Double

'en una matriz cuadrada de n x n transpone las filas/columnas
'i y j
n = UBound(mata, 1)
matb = mata
ReDim matpivot(1 To n) As Variant
For i = 1 To n
 matpivot(i) = matb(i, v2)
 matb(i, v2) = matb(i, v1)
 matb(i, v1) = matpivot(i)
Next i
For i = 1 To n
 matpivot(i) = matb(v2, i)
 matb(v2, i) = matb(v1, i)
 matb(v1, i) = matpivot(i)
Next i
MPermFilaCol = matb
End Function

Function MPermuta(ByRef mata() As Double, ByVal v1 As Integer, ByVal v2 As Integer, ByVal cf As String) As Double()
Dim n As Integer
Dim i As Integer
Dim matb() As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'en una matriz cuadrada de n x n transpone las filas/columnas
'i y j
n = UBound(mata, 1)
matb = mata
If cf = "c" Then
ReDim matpivot(1 To n) As Double
   For i = 1 To n
       matpivot(i) = matb(i, v2)
       matb(i, v2) = matb(i, v1)
       matb(i, v1) = matpivot(i)
   Next i
ElseIf cf = "f" Then
ReDim matpivot(1 To n) As Double
   For i = 1 To n
       matpivot(i) = matb(v2, i)
       matb(v2, i) = matb(v1, i)
       matb(v1, i) = matpivot(i)
   Next i
End If
MPermuta = matb
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function MatDesChol(ByRef mata() As Double, ByRef matv() As Double, ByRef orden As Integer) As Double()
Dim n As Integer
Dim mindent() As Double
Dim i As Integer
Dim j As Integer
Dim matb() As Double
Dim matc() As Double
Dim txtmsg As String
Dim exito As Boolean
Dim m_ind() As Variant
Dim m_ind2() As Variant
Dim vdet As Double
Dim minv() As Double
Dim contar As Long
Dim contar1 As Long
Dim indice As Long
Dim mats() As Double
Dim matt() As Double
Dim dist As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'esta rutina tiene como objetivo obtener la descomposicion de
'una matriz simetrica en 2 partes
'1 matriz a la que se le pueda aplicar el algoritmo de
'choleski y otra que es combinacion linal de esta
'
matb = mata   'se usa una copia del original
n = UBound(mata, 1)
mindent = MIdentidad(n)
matv = MIdentidad(n)
matc = Choleski(matb, m_ind, txtmsg, exito)
If exito Then
   MatDesChol = matc
   orden = n
Else
   orden = n - UBound(m_ind, 1)
   contar = 0
   contar1 = 1
   Do While contar1 <= UBound(m_ind, 1)
       indice = n - contar
       If m_ind(contar1) < indice Then
          If Not esfilanula(matb, indice) Then
             matb = MPermFilaCol(matb, m_ind(contar1), indice)
             matv = MMult(matv, MPermuta(mindent, m_ind(contar1), indice, "f"))
             contar1 = contar1 + 1
          End If
       Else
          contar1 = contar1 + 1
       End If
       contar = contar + 1
   Loop
   orden = UBound(m_ind, 1)
   ReDim matc(1 To n - UBound(m_ind, 1), 1 To n - UBound(m_ind, 1)) As Double
   For i = 1 To n - UBound(m_ind, 1)
       For j = 1 To n - UBound(m_ind, 1)
           matc(i, j) = matb(i, j)
       Next j
   Next i
   mats = Choleski(matc, m_ind2, txtmsg, exito)
   MatDesChol = mats
   txtmsg = "Se obtuvo una matriz no singular de orden " & orden
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "MatDesChol: " & error(Err())
On Error GoTo 0
End Function

Function esfilanula(ByRef mata() As Double, ByVal col As Long)
Dim suma As Double
Dim i As Long

suma = 0
For i = 1 To UBound(mata, 1)
suma = suma + Abs(mata(i, col))
Next i
If suma <> 0 Then
   esfilanula = False
Else
   esfilanula = True
End If
End Function

Function MDescSimTriang(ByRef mata() As Double) As Double()
Dim n As Integer
Dim orden As Integer
Dim nofil1 As Integer
Dim nofil2 As Integer

Dim vdet As Double
Dim matb() As Double
Dim matc() As Double
Dim matd() As Double
Dim mats() As Double
Dim matx() As Double
Dim matv() As Double
Dim matind1() As Variant
Dim matind2() As Long

Dim matww() As Double
Dim orden1 As Integer
Dim match() As Double
Dim match1() As Double
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim matr() As Double
Dim nocomp As Integer
Dim txtcadena As String
Dim dist As Double

'descomposicion de una matriz simetrica en triangular
n = UBound(mata, 1)
'se obtiene la matriz b a la que se le puede aplicar el algoritmo de choleski

'de esta matriz obten   emos matc
match = MatDesChol(matb, matv, nofil2)

orden = n - (nofil1 + nofil2)
If orden < n Then
   matc = MMult(MMult(MTranD(matv), matb), matv)
   ReDim matd(1 To orden, 1 To orden) As Double
   For i = 1 To orden
       For j = 1 To orden
           matd(i, j) = matc(i, j)
       Next j
   Next i
'y la matriz e
   ReDim mate(1 To orden, 1 To n - orden) As Double
   For i = 1 To orden
       For j = 1 To n - orden
           mate(i, j) = matc(i, j + orden)
       Next j
   Next i
'estas son las ecuaciones que sustentan el funcionamiento esta rutina
'|c   d|
'|dt  e|
'd = c * x
'x=inv(c)*d
'e = dt * x
'c=m_chol_t * m_chol
'm_chol_t =[m_chol, m_chol* inv(c)* d]* matv
   matx = MMult(MInversa(matc, vdet), matd)
   matr = MMult(MMult(MTranD(match), UnirMatDbl(MIdentidad(orden), matx, 0)), matv)
    
   MDescSimTriang = matr
Else
   MDescSimTriang = MTranD(match)
End If

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub ObtMatCholeski(ByRef mata() As Double, ByRef matr() As Double)
Dim i As Long
Dim j As Long
Dim k As Long
Dim noreg As Long
Dim noreg1 As Long
Dim dist As Double
Dim contar As Long
Dim contar1 As Long
Dim contar2 As Long
Dim matb() As Double
Dim matc() As Double
Dim matd() As Double
Dim mate() As Double
Dim matx() As Double
Dim mcol() As Variant
Dim mcol2() As Variant
Dim match() As Double
Dim matv() As Double
Dim mindent() As Double
Dim suma As Double
Dim txtmsg As String
Dim exito As Boolean
Dim val_s As Double
Dim indice As Long
'se buscan las columnas/filas que esten repetidas en la matriz de covarianzas
contar = 0
matb = mata
noreg = UBound(matb, 1)
mindent = MIdentidad(noreg)
matv = MIdentidad(noreg)
ReDim mcol(1 To 1, 1 To 1)
For i = 1 To UBound(matb, 1)
    For j = i To UBound(matb, 1)
        If i <> j Then
           dist = 0
           For k = 1 To UBound(matb, 2)
               dist = dist + Abs(matb(k, i) - matb(k, j))
           Next k
           If dist < 0.0000001 Then
              contar = contar + 1
              ReDim Preserve mcol(1 To 1, 1 To contar)
              mcol(1, contar) = j
           End If
        End If
    Next j
Next i
mcol = MTranV(mcol)
'se corre el proceso de choleski 1 vez, lo cual da tambien las columnas/filas que son dependientes
match = Choleski(matb, mcol2, txtmsg, exito)
mcol = UnirMatrices(mcol, mcol2, 1)
mcol = ObtFactUnicos(mcol, 1)
contar = 1
contar1 = 1
contar2 = 0
Do While contar1 <= UBound(mcol, 1)
       indice = noreg - contar + 1
       If mcol(contar1, 1) < indice Then
          If Not esfilanula(matb, indice) And Not sivalorenVec(mcol, indice) Then
             matb = MPermFilaCol(matb, mcol(contar1, 1), indice)
             matv = MMult(matv, MPermuta(mindent, mcol(contar1, 1), indice, "f"))
             contar1 = contar1 + 1
             contar2 = contar2 + 1
          End If
       Else
         contar1 = contar1 + 1
       End If
       contar = contar + 1
Loop
'lo que ahora queda claro es que la descomposicion se debe de hacer sobre
'el total de filas dependientes
noreg1 = noreg - UBound(mcol, 1)
ReDim matc(1 To noreg1, 1 To noreg1)
For i = 1 To noreg1
    For j = 1 To noreg1
        matc(i, j) = matb(i, j)
    Next j
Next i

ReDim matd(1 To noreg1, 1 To UBound(mcol, 1))
For i = 1 To noreg1
    For j = 1 To contar2
        matd(i, j) = matb(i, j + noreg1)
    Next j
Next i
match = Choleski(matc, mcol2, txtmsg, exito)
matx = MMult(MInversa(matc, val_s), matd)
matr = MMult(MMult(MTranD(match), UnirMatDbl(MIdentidad(noreg1), matx, 0)), matv)

End Sub
Function sivalorenVec(ByRef mata() As Variant, valor)
Dim i As Long
sivalorenVec = False
For i = 1 To UBound(mata, 1)
    If mata(i, 1) = valor Then
       sivalorenVec = True
       Exit Function
    End If
Next i
End Function


Sub PermRengV(ByRef mata() As Variant, ByVal ind1 As Long, ByVal ind2 As Long)
Dim Temp As Variant
Dim i As Long
Dim nocols As Long
'permuta los renglones ind1 e ind2 de la matriz mata
nocols = UBound(mata, 2)
For i = 1 To nocols
    Temp = mata(ind1, i)
    mata(ind1, i) = mata(ind2, i)
    mata(ind2, i) = Temp
    DoEvents
Next i
End Sub

Sub PermRengDb(ByRef mata() As Double, ByVal ind1 As Long, ByVal ind2 As Long)
Dim Temp As Double
Dim i As Long
Dim nocols As Long
'permuta los renglones ind1 e ind2 de la matriz mata
nocols = UBound(mata, 2)
For i = 1 To nocols
    Temp = mata(ind1, i)
    mata(ind1, i) = mata(ind2, i)
    mata(ind2, i) = Temp
    DoEvents
Next i
End Sub

Function ExtraeSubMatrizV(ByRef mata() As Variant, ByVal inix As Integer, ByVal finx As Integer, ByVal iniy As Integer, ByVal finy As Integer) As Variant()
Dim i As Integer
Dim j As Integer

'extrae una submatriz en funcion de los parametros inix, finx, iniy, finy
ReDim matx(1 To finy - iniy + 1, 1 To finx - inix + 1) As Variant
For i = 1 To finy - iniy + 1
    For j = 1 To finx - inix + 1
        matx(i, j) = mata(iniy + i - 1, inix + j - 1)
    Next j
Next i
ExtraeSubMatrizV = matx
End Function

Function ExtraeSubMatD(ByRef mata() As Double, ByVal inix As Integer, ByVal finx As Integer, ByVal iniy As Integer, ByVal finy As Integer) As Double()
Dim i As Integer
Dim j As Integer
'extrae una submatriz en funcion de los parametros inix, finx, iniy, finy
ReDim matx(1 To finy - iniy + 1, 1 To finx - inix + 1) As Double
For i = 1 To finy - iniy + 1
    For j = 1 To finx - inix + 1
        matx(i, j) = mata(iniy + i - 1, inix + j - 1)
    Next j
Next i
ExtraeSubMatD = matx
End Function

Function HouseHolder(ByRef mata() As Double) As Double()
Dim ndim As Integer
Dim mata1() As Double
Dim m As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim NoIteraciones As Integer
Dim matq() As Double
Dim suma As Double
Dim suma1 As Double
Dim Scala As Integer
Dim SDif As Double

mata1 = mata
ndim = UBound(mata1, 1)
ReDim vect1(1 To ndim, 1 To 1) As Double
ReDim vect2(1 To ndim, 1 To 1) As Double
ReDim vect3(1 To ndim, 1 To 1) As Double
ReDim eigenv(1 To ndim) As Double

'DE ENTRADA SE REALIZAN 1000 ITERACIONES DEL METODO
NoIteraciones = 1000
For m = 1 To NoIteraciones
 matq = MIdentidad(ndim)
 For i = 1 To ndim
  suma = 0
  For j = i To ndim
  suma = suma + mata1(j, i) ^ 2
  Next j
  suma = Sqr(suma)
  If suma <> Abs(mata1(i, i)) Then
  Scala = 1 / Sqr(suma * (Abs(mata1(i, i)) + suma))
  If mata1(i, i) < 0 Then suma = -suma
  For j = 1 To ndim
   If j <= i Then
    vect1(j, 1) = 0
   Else
    vect1(j, 1) = Scala * mata1(j, i)
   End If
  Next j
  vect1(i, 1) = Scala * (mata1(i, i) + suma)
'Se procede a obtener la matriz Hu y la matriz Q
 For k = 1 To ndim
  vect2(k, 1) = 0
  vect3(k, 1) = 0
  suma1 = 0
  For l = 1 To ndim
   vect2(k, 1) = vect2(k, 1) + matq(k, l) * vect1(l, 1)
   suma1 = suma1 + vect1(l, 1) * mata1(l, k)
  Next l
  For l = 1 To ndim
   matq(k, l) = matq(k, l) - vect2(k, 1) * vect1(l, 1)
   mata1(l, k) = mata1(l, k) - suma1 * vect1(l, 1)
  Next l
 Next k
 End If
 Next i
 'SE REALIZA EL PRODUCTO DE LA MATRIZ R POR LA MATRIZ Q
 'Y SE VUELVE A COLOCAR EN LA MATRIZ A
 mata1 = MMult(mata1, matq)
 SDif = 0
 For i = 1 To ndim
 SDif = SDif + (mata1(i, i) - eigenv(i)) ^ 2
 eigenv(i) = mata1(i, i)
 Next i
 SDif = Sqr(SDif)
If SDif < 0.000000001 * m Then Exit For
Call MostrarMensajeSistema("Iteración no " & m & " para el calculo de los eigenvalores " & Format(SDif, "0.000e+0"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
DoEvents
Next m
HouseHolder = eigenv
End Function

Function MTranDt(ByRef a() As Date)
    Dim i As Integer
    Dim j As Integer
    Dim n1 As Long
    Dim n2 As Long

    '¿Adivinan? es la matriz transpuesta
    n1 = UBound(a, 1)
    n2 = UBound(a, 2)
    ReDim B(1 To n2, 1 To n1) As Date

    For i = 1 To n1
        For j = 1 To n2
            B(j, i) = a(i, j)
        Next j
    Next i

    MTranDt = B
End Function

Function ConvArVtDbl(ByRef mata() As Variant) As Double()
    Dim m As Integer, n As Integer
    Dim i      As Integer, j As Integer
    Dim matb() As Double

    m = UBound(mata, 1)
    n = UBound(mata, 2)
    ReDim matb(1 To m, 1 To n) As Double

    For i = 1 To m
        For j = 1 To n
            matb(i, j) = CDbl(mata(i, j))
        Next j
    Next i

    ConvArVtDbl = matb
End Function

Function ConvArVtDT(ByRef mata() As Variant) As Date()
    Dim m As Integer, n As Integer
    Dim i      As Integer, j As Integer
    Dim matb() As Date

    m = UBound(mata, 1)
    n = UBound(mata, 2)
    ReDim matb(1 To m, 1 To n) As Date

    For i = 1 To m
        For j = 1 To n
            matb(i, j) = CDate(mata(i, j))
        Next j
    Next i

    ConvArVtDT = matb
End Function

Function ExtraeSubMatV(ByRef mata() As Variant, _
                       ByVal inix As Integer, _
                       ByVal finx As Integer, _
                       ByVal iniy As Integer, _
                       ByVal finy As Integer) As Variant()

    'extrae una submatriz en funcion de los parametros inix, finx, iniy, finy
    Dim i As Integer
    Dim j As Integer

    ReDim matx(1 To finy - iniy + 1, 1 To finx - inix + 1) As Variant

    For i = 1 To finy - iniy + 1
        For j = 1 To finx - inix + 1
            matx(i, j) = mata(iniy + i - 1, inix + j - 1)
        Next j
    Next i

    ExtraeSubMatV = matx
End Function


