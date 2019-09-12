Attribute VB_Name = "FBusqueda"
Option Explicit

Sub CompararCVPrecios(ByVal fechai As Date, mat, etiqueta2)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim n As Long
Dim i As Long
Dim dven As Integer
Dim indicep As Long


'en esta rutina se verifican inconsistencias entre
'la posicion de la mesa y el vector de precios
If noprecios = 0 Then Exit Sub
n = UBound(mat, 1)
For i = 1 To n
'dias de vencimiento del titulo
'primero se calculan por formula
dven = Maximo(mat(i).fVencMD - fechai, 0)
mat(i, CDVencimiento) = dven
indicep = 0
If Not EsVariableVacia(mat(i).cEmisionMD) Then indicep = BuscarValorArray(mat(i).cEmisionMD, VectorPrecios, 5)
'se verifica que sea el titulo que se esta
'buscando
'se verifica una vez mas que coincida el nombre en la posicion
'y en el vector de precios
If indicep <> 0 Then
If mat(i).cEmisionMD = VectorPrecios(indicep, 5) Or "I" & Right(mat(i).cEmisionMD, Len(mat(i).cEmisionMD) - 1) = VectorPrecios(indicep, 5) Then
 'si el valor nominal es cero en la posicion se usa el
 'del vector de precios
 If mat(i).vNominalMD = 0 And VectorPrecios(indicep, 11) <> 0 Then mat(i).vNominalMD = VectorPrecios(indicep, 11)
 'dias de vencimiento
 If Not IsDate(mat(i).fVencMD) Or mat(i).fVencMD <> VectorPrecios(indicep, 16) Then
  mat(i).fVencMD = VectorPrecios(indicep, 16)
 End If
 If VectorPrecios(indicep, 12) <> 0 Then
  dven = VectorPrecios(indicep, 12)
  mat(i, CDVencimiento) = VectorPrecios(indicep, 12)
 End If
 If VectorPrecios(indicep, 7) <> 0 Then mat(i).tCuponVigenteMD = VectorPrecios(indicep, 7)       'tasa de interes cupon vigente
 If VectorPrecios(indicep, 6) <> 0 Then mat(i).PCuponActSwap = VectorPrecios(indicep, 6)
 
 
End If
End If
If indicep = 0 Then
 If mat(i).vNominalMD <> 0 Or Not IsNull(mat(i).vNominalMD) Then
  mat(i).valLimpioPIP = mat(i).vNominalMD
  mat(i).valSucioPIP = mat(i).vNominalMD
 Else
  mat(i).valLimpioPIP = 0
  mat(i).valSucioPIP = 0
 End If
If mat(i).fVencMD = fechai Then
   Call MostrarMensajeSistema("Atencion: Falta el papel " & mat(i).cEmisionMD & " en el vector de precios" & mat(i, CDVencimiento), etiqueta2, 0.5, Date, Time, NomUsuario)
   mat(i, CDVencimiento) = 0
Else
   mat(i, CDVencimiento) = 0
End If
End If
 AvanceProc = i / n
 MensajeProc = "Buscando precio en el Vector Precios " & Format(AvanceProc, "#,##0.00 %")
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub BuscarPrVPrecios(ByVal fechai As Date, ByRef mat() As Variant)
Dim n As Long
Dim i As Long
Dim indicep As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'en esta rutina solo se anexa el precio
n = UBound(mat, 1)
For i = 1 To n
indicep = 0
If Len(Trim(mat(i).cEmisionMD)) <> 0 Then indicep = BuscarValorArray(mat(i).cEmisionMD, VectorPrecios, 5)
'se verifica que sea el titulo que se esta
'buscando
'se verifica una vez mas que coincida el nombre en la posicion
'y en el vector de precios
If indicep <> 0 Then
If mat(i).cEmisionMD = VectorPrecios(indicep, 5) Or "I" & Right(mat(i).cEmisionMD, Len(mat(i).cEmisionMD) - 1) = VectorPrecios(indicep, 5) Then

  If Val(VectorPrecios(indicep, 8)) <> 0 Then mat(i).valLimpioPIP = Val(VectorPrecios(indicep, 8))  'PRECIO LIMPIO
 If Val(VectorPrecios(indicep, 9)) <> 0 Then mat(i).valSucioPIP = Val(VectorPrecios(indicep, 9))   'PRECIO SUCIO  volver aqui
' If VectorPrecios(indicep, 7) <> 0 Then mat(i).tCuponVigenteMD = VectorPrecios(indicep, 7)       'tasa de interes cupon vigente

End If
End If
If indicep = 0 Then
 If mat(i).vNominalMD <> 0 Or Not IsNull(mat(i).vNominalMD) Then
  mat(i).valLimpioPIP = mat(i).vNominalMD
  mat(i).valSucioPIP = mat(i).vNominalMD
 Else
  mat(i).valLimpioPIP = 0
  mat(i).valSucioPIP = 0
 End If
If mat(i).fVencMD = fechai Then
  MensajeProc = "Atencion: Falta el papel " & mat(i).cEmisionMD & " en el vector de precios" & mat(i, CDVencimiento)
  mat(i, CDVencimiento) = 0
Else
  mat(i, CDVencimiento) = 0
End If
End If
 AvanceProc = i / n
 
 Call MostrarMensajeSistema("Buscando precio en el Vector Precios " & Format(i / n, "#,##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function BuscarDatosVPrecios(ByVal fechai As Date, ByRef mat() As Variant, ByRef matvp() As Variant)
Dim i As Long
Dim indicep As Long
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'compara algunos parametros del vector de precios con la posicion
If noprecios = 0 Then Exit Function
noreg = UBound(mat, 1)
For i = 1 To noreg
indicep = 0
If Len(Trim(mat(i).cEmisionMD)) <> 0 Then indicep = BuscarValorArray(mat(i).cEmisionMD, matvp, 5)
'se verifica que sea el titulo que se esta
'buscando
'se verifica una vez mas que coincida el nombre en la posicion
'y en el vector de precios
If indicep <> 0 Then
If mat(i).cEmisionMD = matvp(indicep, 5) Or "I" & Right(mat(i).cEmisionMD, Len(mat(i).cEmisionMD) - 1) = matvp(indicep, 5) Then
  If Val(matvp(indicep, 8)) <> 0 Then mat(i).valLimpioPIP = Val(matvp(indicep, 8))  'PRECIO LIMPIO
  If Val(matvp(indicep, 9)) <> 0 Then mat(i).valSucioPIP = Val(matvp(indicep, 9))    'PRECIO SUCIO  volver aqui
  If matvp(indicep, 7) <> 0 Then mat(i).tCuponVigenteMD = matvp(indicep, 7)       'tasa de interes cupon vigente
  If matvp(indicep, 16) <> 0 Then mat(i).fVencMD = matvp(indicep, 16)       'tasa de interes cupon vigente
  If matvp(indicep, 11) <> 0 Then mat(i).vNominalMD = matvp(indicep, 11)       'valor nominal
  If matvp(indicep, 6) <> 0 Then mat(i).PCuponActSwap = matvp(indicep, 6)       'periodo cupon
End If
Else
MsgBox "no se econtraron datos de " & mat(i).cEmisionMD & " en el vector de precios"
End If
 AvanceProc = i / noreg
 
 Call MostrarMensajeSistema("Buscando precio en el Vector Precios " & Format(AvanceProc, "#,##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
Next i
BuscarDatosVPrecios = mat
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ExtSerieFR(ByRef mata() As Variant, ByVal ncol As Integer, ByVal ifinal As Integer, ByVal noesc As Integer) As Double()
Dim i As Long

'se leen los datos de la tabla de datos
'b    indice correspondiente al vector a cargar
'c    no de datos a cargar
'     devuelve como resultado una matriz de c por 2
'     16 de julio del 2001
ReDim matb(1 To noesc, 1 To 1) As Double
'la lectura de datos se hace hacia atras
For i = 1 To noesc
 matb(i, 1) = mata(ifinal - noesc + i, ncol)
Next i
ExtSerieFR = matb
End Function

Function ExtSerieAD(ByRef mata() As Double, ByVal ncol As Integer, ByVal ifinal As Integer, ByVal noesc As Integer) As Double()
Dim i As Integer

'se leen los datos de la tabla de datos
'b    indice correspondiente al vector a cargar
'c    no de datos a cargar
'     devuelve como resultado una matriz de c por 2
'     16 de julio del 2001
ReDim matb(1 To noesc, 1 To 1) As Double
'la lectura de datos se hace hacia atras
For i = 1 To noesc
 matb(i, 1) = mata(ifinal - noesc + i, ncol)
Next i
ExtSerieAD = matb
End Function

Function ExtraerSerieFRyF(ByRef mata() As Variant, ByVal ncol As Integer, ByVal ifinal As Integer, ByVal ndias As Integer)
Dim i As Integer
'se leen los datos de la tabla de datos
'b    indice correspondiente al vector a cargar
'c    no de datos a cargar
'     devuelve como resultado una matriz de c por 2
'     16 de julio del 2001

ReDim matb(1 To ndias, 1 To 2) As Variant
'la lectura de datos se hace hacia atras
For i = 1 To ndias
 matb(i, 1) = mata(ifinal - ndias + i, 1)
 matb(i, 2) = mata(ifinal - ndias + i, ncol)
Next i
ExtraerSerieFRyF = matb
End Function


Function BuscarValorArray(ByVal a As Variant, mat, ByVal j As Integer) As Long
Dim n As Long
Dim rangominimo As Long
Dim rangomaximo As Long
Dim rangomedio As Long
Dim sibuscar As Boolean

'esta es una rutina de en la que se prueba
'un nuevo metodo para la busqueda de datos
'se pone como condicion en la columna de busqueda los
'datos estan ordenados
'se busca el valor a en la matriz mat en la columna j
'y devuelve el indice del renglon donde lo encuentra
BuscarValorArray = 0
If Not IsEmpty(mat) Then
n = UBound(mat, 1)
rangominimo = 1
rangomaximo = n
sibuscar = True
Do While sibuscar
rangomedio = Int((rangominimo + rangomaximo) / 2)
If rangomedio = 0 Then
   BuscarValorArray = 0
   Exit Function
End If
If mat(rangomedio, j) < a And rangominimo <> rangomedio Then
   rangominimo = rangomedio
ElseIf mat(rangomedio, j) > a And rangomaximo <> rangomedio Then
   rangomaximo = rangomedio
ElseIf mat(rangomedio, j) = a Then
   BuscarValorArray = rangomedio
   Exit Do
ElseIf mat(rangominimo, j) = a Then
   BuscarValorArray = rangominimo
   sibuscar = False
ElseIf mat(rangomaximo, j) = a Then
   BuscarValorArray = rangomaximo
   sibuscar = False
ElseIf mat(rangomedio, j) <> a And ((rangominimo = rangomedio) Or (rangomaximo = rangomedio)) Then
   sibuscar = False
End If
Loop
Else
  MensajeProc = "No hay datos en la matriz donde se busca el valor"
End If
End Function

Function BuscarValorInt(ByVal a As Variant, mat, ByVal j As Integer) As Long
On Error GoTo hayerror

Dim n As Long
Dim rangominimo As Long
Dim rangomaximo As Long
Dim rangomedio As Long
Dim sibuscar As Boolean

'modificacion de la rutina buscarvalorint encuentra
BuscarValorInt = 0
n = UBound(mat, 1)
rangominimo = 1
rangomaximo = n
sibuscar = True
Do While sibuscar
   rangomedio = Int((rangominimo + rangomaximo) / 2)   'punto medio
If mat(rangomedio).plazo < a And rangominimo <> rangomedio Then
   rangominimo = rangomedio                            'este es el nuevo punto minimo
ElseIf mat(rangomedio).plazo > a And rangomaximo <> rangomedio Then
   rangomaximo = rangomedio                            'este es el nuevo punto maximo
ElseIf mat(rangomedio).plazo = a Then
   BuscarValorInt = rangomedio
   Exit Do
ElseIf mat(rangominimo).plazo = a Then
   BuscarValorInt = rangominimo
   sibuscar = False
ElseIf mat(rangomaximo).plazo = a Then
   BuscarValorInt = rangomaximo
   sibuscar = False
ElseIf mat(rangomedio).plazo <> a And (rangominimo + 1 = rangomaximo) Then
   BuscarValorInt = rangominimo
   Exit Do
ElseIf mat(rangominimo).plazo = a And (rangominimo + 1 = rangomaximo) Then
   BuscarValorInt = rangominimo
   Exit Do
End If
Loop
On Error GoTo 0
Exit Function
hayerror:
MsgBox "BuscarvalorInt " & error(Err())
End Function

Function OrdenarMat(ByRef a() As Variant, ByVal ind As Integer, ByRef exito As Boolean)
Dim n As Long
Dim m As Long
Dim i As Long
Dim j As Long
Dim kk As Long
Dim mtemp() As Variant
Dim pivot As Variant

'rutina para ordenacion de datos.
'esta es una rutina de ordenacion estandar
'1 si es una matriz
'algoritmo basado en el método de la burbuja
mtemp = a
n = UBound(mtemp, 1)
m = UBound(mtemp, 2)
 For i = 1 To n
  For j = 1 To n - i
   If mtemp(j, ind) > mtemp(j + 1, ind) Then
      For kk = 1 To m
          pivot = mtemp(j + 1, kk)
          mtemp(j + 1, kk) = mtemp(j, kk)
          mtemp(j, kk) = pivot
      Next kk
  End If
 Next j
 DoEvents
Next i
exito = True
OrdenarMat = mtemp
End Function


Function BuscarValVDT(ByVal a As Date, ByRef mat() As Date, ByVal j As Integer)

    Dim n As Long
    Dim rangomedio As Long
    Dim rangominimo As Long
    Dim rangomaximo As Long

    Dim sibuscar As Boolean

    BuscarValVDT = 0

    If Not IsEmpty(mat) Then
        n = UBound(mat, 1)
        rangominimo = 1
        rangomaximo = n
        sibuscar = True

        Do While sibuscar
            rangomedio = Int((rangominimo + rangomaximo) / 2)

            If rangomedio = 0 Then
                BuscarValVDT = 0

                Exit Function

            End If

            If mat(rangomedio, j) < a And rangominimo <> rangomedio Then
                rangominimo = rangomedio
            ElseIf mat(rangomedio, j) > a And rangomaximo <> rangomedio Then
                rangomaximo = rangomedio
            ElseIf mat(rangomedio, j) = a Then
                BuscarValVDT = rangomedio

                Exit Do

            ElseIf mat(rangominimo, j) = a Then
                BuscarValVDT = rangominimo
                sibuscar = False
            ElseIf mat(rangomaximo, j) = a Then
                BuscarValVDT = rangomaximo
                sibuscar = False
            ElseIf mat(rangomedio, j) <> a And ((rangominimo = rangomedio) Or (rangomaximo = rangomedio)) Then
                sibuscar = False
            End If

        Loop

    Else
        MensajeProc = "No hay datos en la matriz donde se busca el valor"
    End If

End Function

