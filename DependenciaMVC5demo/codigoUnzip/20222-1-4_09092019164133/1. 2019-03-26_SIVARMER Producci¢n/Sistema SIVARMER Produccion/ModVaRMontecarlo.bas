Attribute VB_Name = "ModVaRMontecarlo"
Option Explicit

Sub CalculoVaRMontecarlo(ByVal f_val As Date, _
                         ByRef matpos() As propPosRiesgo, _
                         ByRef matposmd() As propPosMD, _
                         ByRef matposdiv() As propPosDiv, _
                         ByRef matposswaps() As propPosSwaps, _
                         ByRef matposfwd() As propPosFwd, _
                         ByRef matflswap() As estFlujosDeuda, _
                         ByRef matposdeuda() As propPosDeuda, _
                         ByRef matfldeuda() As estFlujosDeuda, _
                         ByRef matorden2() As Variant, ByVal nosim As Long, ByVal htiempo As Integer, ByRef medias1() As Double, ByRef matrizch1() As Double, ByRef mat_num_a() As String, ByVal tmedia As Integer, ByRef matv0() As resValIns, ByRef matpyg() As Double)
                         
Dim matFRiesgo0() As Double
Dim parval As ParamValPos
Dim mrvalflujo() As resValFlujo
Dim matv1() As New resValIns
Dim indice As Integer
Dim nocont As Integer
Dim nofactind As Integer
Dim noregpos As Integer
Dim matindfriesgo() As Long
Dim i As Integer
Dim j As Integer
Dim w As Integer

Dim matfactsim() As Double
Dim numalea1() As Double
Dim txtmsg As String
Dim exito As Boolean

If ActivarControlErrores Then
 On Error GoTo ControlErrores
End If
'matindfriesgo1 indica los factores de riesgo para los cuales es sensible el portafolio
'en las simulaciones del var montecarlo
nofactind = UBound(matrizch1, 1) 'la dimension vertical de la matriz de choleski
'se generan los indice de los factores de riesgo
matindfriesgo = GenerarIndFactRiesgo(matorden2, MatCaracFRiesgo)
nocont = UBound(matindfriesgo, 1)
noregpos = UBound(matpos, 1)
'primero se calculan los precios para hoy
'hace una primera valuacion
indice = BuscarValorArray(f_val, MatFactRiesgo, 1)
matFRiesgo0 = ExtFRMatFR(indice, MatFactRiesgo)
Set parval = DeterminaPerfilVal("MONTECARLO")
matv0 = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matFRiesgo0, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
'se procede a poner los valores iniciales en cada simulacion
ReDim matfactsim0(1 To nocont, 1 To 1) As Double
For w = 1 To nocont
    matfactsim0(w, 1) = matFRiesgo0(matindfriesgo(w), 1)
Next w
ReDim matpyg(1 To nosim, 1 To noregpos) As Double
For i = 1 To nosim
    'If matpos(1).c_operacion = "43" And i = 9253 Then Stop
    numalea1 = LeerMuestraNormal1(nofactind, i, mat_num_a)
    matfactsim = GenerarFactoresMontecarlo(matfactsim0, medias1, matrizch1, htiempo, numalea1, 0, tmedia)
    'MsgBox matfactsim(1, 1) & Chr(9) & matfactsim(2, 1) & Chr(9) & matfactsim(3, 1) & Chr(9) & matfactsim(4, 1) & Chr(9) & matfactsim(5, 1) & Chr(9) & matfactsim(6, 1) & Chr(9) & matfactsim(7, 1)
  'se acumulan los resultados
 'este es el vector donde estan las tasas simuladas, de este vector y con valua la posicion suponiendo que permanecio
 'constante para el día siguiente
    ReDim MatFactR1(1 To NoFactores, 1 To 1) As Double
    For w = 1 To NoFactores '
        MatFactR1(w, 1) = matFRiesgo0(w, 1)
    Next w
    For w = 1 To nocont 'aqui esta la parte importante
        MatFactR1(matindfriesgo(w), 1) = matfactsim(w, 1)
    Next w
    matv1 = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
    For j = 1 To noregpos
        matpyg(i, j) = matv1(j).mtm_sucio - matv0(j).mtm_sucio
    Next j
    AvanceProc = i / nosim
    MensajeProc = "Avance de la simulación Montecarlo: " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i

'se exportan los escenarios simulados a un archivo de texto
On Error GoTo 0
Exit Sub
ControlErrores:
  MsgBox error(Err())
On Error GoTo 0
End Sub

Function LeerMuestraNormal1(nofact, i, mata)
Dim a() As String
Dim B() As Double
Dim j As Long
a = EncontrarSubCadenas(mata(i, 1), ",")
ReDim B(1 To nofact, 1 To 1) As Double
For j = 1 To nofact
   B(j, 1) = CDbl(a(j))
Next j
LeerMuestraNormal1 = B
End Function


Sub ImprimirInsumosMontecarlo(ByRef matfsensib() As Variant, ByRef matcov() As Double, ByRef chol() As Double)
Dim nomarch As String
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim n As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim exitoarch As Boolean

'se imprimen en un archivo de texto los insumos para el
'var montecarlo
nomarch = DirResVaR & "\SimMonte.txt"
Call VerificarSalidaArchivo(nomarch, 5, exitoarch)
If exitoarch Then
Print #5, "La posicion es sensible a los factores"
p = UBound(matfsensib, 1)
For i = 1 To p
Print #5, matfsensib(i, 1) & Chr(9);
Next i
Print #5, ""

n = UBound(matcov, 1)
Print #5, "La matriz de covarianzas"
For i = 1 To n
 For j = 1 To n
  Print #5, matcov(i, j) & Chr(9);
 Next j
 Print #5, ""
Next i
n1 = UBound(chol, 1)
n2 = UBound(chol, 2)
Print #5, "La matriz de choleski"
For i = 1 To n1
For j = 1 To n2
Print #5, chol(i, j) & Chr(9);
Next j
Print #5, ""
Next i
Close #5
End If
End Sub


Sub CalculoMatCholeski(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Long, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
'para el calculo de la matriz de choleski a utilizar en el calculo de escenarios de la simulacion
'montecarlo
'fecha     - es la fecha de calculo, de los factores de riesgo y de la posicion
'txtport   - es el nombre del portafolio de posicion
'txtportfr - es el nombre del portafolio de factores de riesgo de la posicion
'noesc     - es el numero de escenarios para calcular rendimientos
'htiempo   - es el numeros de dias para calcular un rendimiento
'txtmsg    - devuelve un mensaje con el resultados de los calculos
'exito     - indica si el proceso tuvo exito o no


Dim i As Long
Dim j As Long
Dim indice As Long
Dim noreg As Long
Dim matsensib() As Variant
Dim matrend() As Double
Dim m_orden As String
Dim m_medias As String
Dim m_covar As String
Dim m_choleski As String
Dim matr() As Double
Dim dif As Double
Dim matorden2() As Variant
Dim mmedias() As Double
Dim matcov() As Double
Dim matrizch() As Double
Dim txtfecha As String
Dim txtborra As String
Dim exitofr As Boolean
Dim rmesa As New ADODB.recordset

'se busca la fecha en el vector de factores de riesgo
Call VerifCargaFR2(fecha, noesc + htiempo, exitofr)
indice = BuscarValorArray(fecha, MatFactRiesgo, 1)
If indice > 0 Then
' se cargan las sesnbilidades calculadas para el portafolio a analizar
   matsensib = LeerSensibPort(fecha, txtport, txtportfr, txtport)
   noreg = UBound(matsensib, 1)
   If noreg <> 0 Then
   ReDim matnfsen(1 To UBound(matsensib, 1), 1 To 1) As Variant
      For i = 1 To UBound(matsensib, 1)
          matnfsen(i, 1) = matsensib(i, 1)
      Next i
      MensajeProc = "La posicion es sensible en " & UBound(matsensib, 1) & " factores"
'se obtiene la matriz de rendimientos entre factores de riesgo
'se calculan la matriz de medias y de covarianzas
      matrend = GenMatRendRiesgo(matnfsen, indice, noesc, htiempo)
      mmedias = GenMedias(matrend, 0, 0)
      matcov = GenCovar(matrend, matrend, 0, 0)
      ReDim matorden1(1 To UBound(matsensib, 1), 1 To 2) As Variant
      For i = 1 To UBound(matsensib, 1)
          matorden1(i, 1) = matsensib(i, 1)
          matorden1(i, 2) = Abs(matsensib(i, 3) * matsensib(i, 2)) * (matcov(i, i)) ^ 0.5
      Next i
  'se reordena matorden1 en funcion de la columna 2
   matorden1 = RutinaOrden(matorden1, 2, SRutOrden)
   ReDim matorden2(1 To UBound(matsensib, 1), 1 To 1) As Variant
   For i = 1 To UBound(matsensib, 1)
       matorden2(i, 1) = matorden1(UBound(matsensib, 1) - i + 1, 1)
   Next i
'se vuelve a repetir el proceso de seleccion
   matrend = GenMatRendRiesgo(matorden2, indice, noesc, htiempo)
   mmedias = GenMedias(matrend, 0, 0)
   matcov = GenCovar(matrend, matrend, 0, 0)
'Se calcula la matriz de Choleski a partir de la matriz de covarianzas
   Call ObtMatCholeski(matcov, matrizch)
   'se encuentra la distancia matricial entre el producto de la matriz de choleski y la original
   matr = MMult(MTranD(matrizch), matrizch)
   dif = DistanciaMat(matr, matcov)
   'si la distancia es menor en un cierto nivel, se acepta el calculo y se guarda para utilizarlo
   'en el calculo de var montecarlo
   If dif < 0.0001 Then
      m_orden = ""
      For i = 1 To UBound(matorden2, 1)
          m_orden = m_orden & matorden2(i, 1) & ","
      Next i
      m_medias = ""
      For i = 1 To UBound(mmedias, 1)
          m_medias = m_medias & mmedias(i, 1) & ","
      Next i
      m_covar = ""
      For i = 1 To UBound(matcov, 1)
          For j = 1 To UBound(matcov, 2)
              m_covar = m_covar & matcov(i, j) & ","
          Next j
      Next i
      m_choleski = ""
      For i = 1 To UBound(matrizch, 1)
          For j = 1 To UBound(matrizch, 2)
              m_choleski = m_choleski & matrizch(i, j) & ","
          Next j
      Next i
      txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
      txtborra = "DELETE FROM " & TablaFactChol & " WHERE FECHA = " & txtfecha
      ConAdo.Execute txtborra
      rmesa.Open "SELECT * FROM " & TablaFactChol, ConAdo, 1, 3
      rmesa.AddNew
      rmesa.Fields("FECHA") = fecha
      rmesa.Fields("PORTAFOLIO") = txtport
      rmesa.Fields("ESC_FACTORES") = txtportfr
      rmesa.Fields("NOESC") = noesc
      rmesa.Fields("HTIEMPO") = htiempo
      Call GuardarElementoClob(m_orden, rmesa, "M_ORDEN")
      Call GuardarElementoClob(m_medias, rmesa, "M_MEDIAS")
      Call GuardarElementoClob(m_covar, rmesa, "MATCOV")
      Call GuardarElementoClob(m_choleski, rmesa, "M_CHOLESKI")
      rmesa.Fields("N1") = UBound(matrizch, 1)
      rmesa.Fields("N2") = UBound(matrizch, 2)
      rmesa.Update
      rmesa.Close
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "El proceso no obtuvo una matriz consistente"
      exito = False
   End If
 Else
   exito = False
   txtmsg = "No hay datos para el portafolio " & txtport
 End If
End If
End Sub

Function LeerMuestraNormal(ByVal fecha As Date, ByVal nofact As Integer, ByVal nosim As Long)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaNumDistNormal & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND NOSIM =" & nosim & " AND NOFACT = " & nofact
txtfiltro2 = txtfiltro2 & " ORDER BY ORDEN"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 1) As String
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("VALORES")
       rmesa.MoveNext
   Next i
   rmesa.Close
   AvanceProc = i / noreg
   MensajeProc = "Leyendo muestra de aleatorios"
Else
   ReDim mata(0 To 0, 0 To 0) As String
End If
LeerMuestraNormal = mata
End Function


Sub GenNormMont(ByVal fecha As Date, ByVal nofact As Integer, ByVal nosim As Long)
Dim i As Long
Dim j As Integer
Dim txtfecha As String
Dim txtcadena As String
Dim txtborra As String
Dim txtcad1 As String
Dim numalea1() As Double
Dim regmont As New ADODB.recordset
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & TablaNumDistNormal & " WHERE FECHA = " & txtfecha & " AND nosim = " & nosim
ConAdo.Execute txtborra
regmont.Open "SELECT * FROM " & TablaNumDistNormal & "", ConAdo, 1, 3
For i = 1 To nosim
    numalea1 = GenMuestraNormal(nofact)
    txtcad1 = ""
    For j = 1 To UBound(numalea1, 1)
        txtcad1 = txtcad1 & numalea1(j, 1) & ","
    Next j
    regmont.AddNew
    regmont.Fields("FECHA") = CDbl(fecha)
    regmont.Fields("NOSIM") = nosim
    regmont.Fields("NOFACT") = nofact
    regmont.Fields("ORDEN") = i
    Call GuardarElementoClob(txtcad1, regmont, "VALORES")
    regmont.Update
    AvanceProc = i / nosim
    DoEvents
Next i
regmont.Close
End Sub
