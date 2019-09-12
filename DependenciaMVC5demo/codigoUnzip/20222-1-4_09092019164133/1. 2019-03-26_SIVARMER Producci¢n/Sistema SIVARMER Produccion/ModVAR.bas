Attribute VB_Name = "ModVAR"
Option Explicit

Sub VerCaracPosicion(rejilla As MSFlexGrid)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

 rejilla.TextMatrix(1, 1) = UBound(MatPosRiesgo, 1)
 rejilla.TextMatrix(2, 1) = 0
 rejilla.TextMatrix(3, 1) = Format(TTitCompra, "###,###,###,###,###,###,###,##0")
 rejilla.TextMatrix(4, 1) = Format(ComprasDirecto, "###,###,###,##0.00")
 rejilla.TextMatrix(5, 1) = Format(ComprasReporto, "###,###,###,##0.00")
 rejilla.TextMatrix(6, 1) = NoVenReporto
 rejilla.ColWidth(0) = 3000
 rejilla.ColWidth(1) = 1300
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub CalcSensibNuevo(ByVal fecha As Date, _
                    ByVal txtport As String, _
                    ByVal txtportfr As String, _
                    ByRef matpos() As propPosRiesgo, _
                    ByRef matposmd() As propPosMD, _
                    ByRef matposdiv() As propPosDiv, _
                    ByRef matposswaps() As propPosSwaps, _
                    ByRef matposfwd() As propPosFwd, _
                    ByRef matflswap() As estFlujosDeuda, _
                    ByRef matposdeuda() As propPosDeuda, _
                    ByRef matfldeuda() As estFlujosDeuda, _
                    ByRef matfr() As Double)

Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim mfactr12() As Double
Dim i As Integer
Dim j As Integer
Dim contar As Integer
Dim incre As Double
Dim nocomp As Integer
Dim deriv As Double
Dim txtfecha As String
Dim txtfechar As String
Dim txtborra As String
Dim txtcadena As String
Dim matpr1() As New resValIns
Dim matpr2() As New resValIns
Dim exito3 As Boolean
Dim txtmsg3 As String

'en esta rutina se usan otras Sens que coincidan
contar = 0
ReDim matrep(1 To 1, 0 To contar) As Variant
incre = 0.00001
nocomp = 0
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaSensibN & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtborra
Set parval = DeterminaPerfilVal("SENSIBILIDADES")
For i = 1 To UBound(matpos, 1)   'toda la posicion
    For j = 1 To NoFactores
       deriv = 0
'se comprueba si el instrumento i en la posicion es sensible al factor jj
       If EsSensibleaFactor(MatCaracFRiesgo(j).nomFactor, matpos(i).C_Posicion, matpos(i).c_operacion, matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda) Then
          mfactr12 = matfr
'se agrega el incremento al factor jj
          mfactr12(j, 1) = matfr(j, 1) + incre
          parval.indpos = i    'se pide que solo se calcule la sensibilidad del instrumento i
          matpr1 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matfr, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
          matpr2 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mfactr12, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
'se obtiene la derivada numerica del factor
          deriv = (matpr2(i).mtm_sucio - matpr1(i).mtm_sucio) / incre
          If deriv <> 0 Then
             txtcadena = "INSERT INTO " & TablaSensibN & " VALUES("
             txtcadena = txtcadena & txtfecha & ","
             txtcadena = txtcadena & "'" & txtport & "',"
             txtcadena = txtcadena & "'" & txtportfr & "',"
             txtcadena = txtcadena & matpos(i).C_Posicion & ","
             txtfechar = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
             txtcadena = txtcadena & txtfechar & ","
             txtcadena = txtcadena & "'" & matpos(i).c_operacion & "',"
             txtcadena = txtcadena & "'" & MatCaracFRiesgo(j).indFactor & "',"
             txtcadena = txtcadena & "'" & MatCaracFRiesgo(j).descFactor & "',"
             txtcadena = txtcadena & "'" & MatCaracFRiesgo(j).nomFactor & "',"
             txtcadena = txtcadena & MatCaracFRiesgo(j).plazo & ","
             txtcadena = txtcadena & "'" & MatCaracFRiesgo(j).tfactor & "',"
             txtcadena = txtcadena & matfr(j, 1) & ","
             txtcadena = txtcadena & deriv & ")"
             ConAdo.Execute txtcadena
           End If
       End If
    Next j
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Calc. sensibilidades de primer orden de " & matpos(i).C_Posicion & " " & matpos(i).c_operacion & "  " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
End Sub

Sub CalcSensibPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtborra As String
Dim i As Integer, j As Integer, noreg As Integer, contar As Integer
Dim txtconcepto As String
Dim txtdescripcion As String
Dim txtcurva As String
Dim vplazo As Integer
Dim tvalor As String
Dim vfactor As Double
Dim vderiv As Double
Dim siencontro As Boolean
Dim txtinserta As String
Dim mata() As New ResCalcSens
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaSensibPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND PORT_FR = '" & txtportfr & "' AND SUBPORT = '" & txtsubport & "'"
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaSensibN & " WHERE "
txtfiltro2 = txtfiltro2 & "FECHA = " & txtfecha & " AND "
txtfiltro2 = txtfiltro2 & "PORTAFOLIO = '" & txtport & "' AND "
txtfiltro2 = txtfiltro2 & "PORT_FR = '" & txtportfr & "' AND "
txtfiltro2 = txtfiltro2 & "(ID_POSICION,FECHA_REG,ID_OPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPortPosicion & " "
txtfiltro2 = txtfiltro2 & "WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To 1)
   contar = 0
   For i = 1 To noreg
       txtconcepto = rmesa.Fields("FACTOR")
       txtdescripcion = rmesa.Fields("DESCRIPCION")
       txtcurva = rmesa.Fields("CURVA")
       vplazo = rmesa.Fields("PLAZO")
       tvalor = rmesa.Fields("TVALOR")
       vfactor = rmesa.Fields("VALOR")
       vderiv = rmesa.Fields("DERIVADA")
       siencontro = False
       For j = 1 To contar
           If txtconcepto = mata(j).factor Then
              mata(j).deriv = mata(j).deriv + vderiv
              siencontro = True
           End If
       Next j
       If j = contar + 1 And Not siencontro Then
          contar = contar + 1
          ReDim Preserve mata(1 To contar)
          mata(contar).factor = txtconcepto
          mata(contar).descrip = txtdescripcion
          mata(contar).curva = txtcurva
          mata(contar).plazo = vplazo
          mata(contar).tfactor = tvalor
          mata(contar).valor = vfactor
          mata(contar).deriv = vderiv
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To contar
       txtinserta = "INSERT INTO " & TablaSensibPort & " VALUES("
       txtinserta = txtinserta & txtfecha & ","
       txtinserta = txtinserta & "'" & txtport & "',"
       txtinserta = txtinserta & "'" & txtportfr & "',"
       txtinserta = txtinserta & "'" & txtsubport & "',"
       txtinserta = txtinserta & "'" & mata(i).factor & "',"
       txtinserta = txtinserta & "'" & mata(i).descrip & "',"
       txtinserta = txtinserta & "'" & mata(i).curva & "',"
       txtinserta = txtinserta & mata(i).plazo & ","
       txtinserta = txtinserta & "'" & mata(i).tfactor & "',"
       txtinserta = txtinserta & mata(i).valor & ","
       txtinserta = txtinserta & mata(i).deriv & ","
       txtinserta = txtinserta & "0)"                                   'volatilidad
       ConAdo.Execute txtinserta
       AvanceProc = i / contar
       MensajeProc = "Generando sensibilidades del portafolio " & txtport & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No hay datos"
   exito = True
End If

End Sub

Sub CalcSensibPort2(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtborra As String
Dim i As Integer, j As Integer, noreg As Integer, contar As Integer
Dim txtconcepto As String
Dim txtdescripcion As String
Dim txtcurva As String
Dim vplazo As Integer
Dim tvalor As String
Dim vfactor As Double
Dim vderiv As Double
Dim siencontro As Boolean
Dim txtinserta As String
Dim mata() As New ResCalcSens
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaSensibPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND PORT_FR = '" & txtportfr & "' AND SUBPORT = '" & txtsubport & "'"
ConAdo.Execute txtborra

txtfiltro2 = "SELECT * FROM " & TablaSensibN & " WHERE "
txtfiltro2 = txtfiltro2 & "FECHA = " & txtfecha & " AND "
txtfiltro2 = txtfiltro2 & "PORTAFOLIO = '" & txtport & "' AND "
txtfiltro2 = txtfiltro2 & "PORT_FR = '" & txtportfr & "' AND "
txtfiltro2 = txtfiltro2 & "(ID_POSICION,FECHA_REG,ID_OPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPortPosicion & " "
txtfiltro2 = txtfiltro2 & "WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To 1)
   contar = 0
   For i = 1 To noreg
       txtconcepto = rmesa.Fields("FACTOR")
       txtdescripcion = rmesa.Fields("DESCRIPCION")
       txtcurva = rmesa.Fields("CURVA")
       vplazo = rmesa.Fields("PLAZO")
       tvalor = rmesa.Fields("TVALOR")
       vfactor = rmesa.Fields("VALOR")
       vderiv = rmesa.Fields("DERIVADA")
       siencontro = False
       For j = 1 To contar
           If txtconcepto = mata(j).factor Then
              mata(j).deriv = mata(j).deriv + vderiv
              siencontro = True
           End If
       Next j
       If j = contar + 1 And Not siencontro Then
          contar = contar + 1
          ReDim Preserve mata(1 To contar)
          mata(contar).factor = txtconcepto
          mata(contar).descrip = txtdescripcion
          mata(contar).curva = txtcurva
          mata(contar).plazo = vplazo
          mata(contar).tfactor = tvalor
          mata(contar).valor = vfactor
          mata(contar).deriv = vderiv
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtinserta = "DELETE FROM " & TablaSensibPort & " WHERE FECHA = " & txtfecha
   txtinserta = txtinserta & " AND PORTAFOLIO = '" & txtsubport & "'"
   ConAdo.Execute txtinserta
   For i = 1 To contar
       txtinserta = "INSERT INTO " & TablaSensibPort & " VALUES("
       txtinserta = txtinserta & txtfecha & ","
       txtinserta = txtinserta & "'" & txtsubport & "',"
       txtinserta = txtinserta & "'" & txtportfr & "',"
       txtinserta = txtinserta & "'" & txtsubport & "',"
       txtinserta = txtinserta & "'" & mata(i).factor & "',"
       txtinserta = txtinserta & "'" & mata(i).descrip & "',"
       txtinserta = txtinserta & "'" & mata(i).curva & "',"
       txtinserta = txtinserta & mata(i).plazo & ","
       txtinserta = txtinserta & "'" & mata(i).tfactor & "',"
       txtinserta = txtinserta & mata(i).valor & ","
       txtinserta = txtinserta & mata(i).deriv & ","
       txtinserta = txtinserta & "0)"                                   'volatilidad
       ConAdo.Execute txtinserta
       AvanceProc = i / contar
       MensajeProc = "Generando sensibilidades del portafolio " & txtport & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No hay datos"
   exito = True
End If

End Sub



Sub CalVaRMark(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txttvar As String
Dim noreg As Integer
Dim i As Integer
Dim indice As Integer
Dim rends() As Double
Dim mmedias1() As Double
Dim cov1() As Double
Dim valfa1 As Double
Dim valfa2 As Double
Dim val1() As Double
Dim val2() As Double
Dim valvar1 As Double
Dim valvar2 As Double
Dim txtinserta As String
Dim txtact As String
Dim deltas() As Double
Dim txtborra As String
Dim mata() As Variant

txttvar = "VARMark"
mata = LeerSensibPort(fecha, txtport, txtportfr, txtsubport)
noreg = UBound(mata, 1)
If noreg <> 0 Then
   indice = BuscarValorArray(fecha, MatFactRiesgo, 1)
   rends = GenMatRendRiesgo(mata, indice, noesc, htiempo)
   mmedias1 = GenMedias(rends, 0, 0)      'medias peso normal
   cov1 = GenCovar(rends, rends, 0, 0)    'covarianzas peso normal
' se obtiene el var con el a % de nivel de confianza
   ReDim deltas(1 To 1, 1 To noreg) As Double
   For i = 1 To noreg
       deltas(1, i) = mata(i, 2) * mata(i, 3)
   Next i
   valfa1 = NormalInv(nconf)
   valfa2 = NormalInv(1 - nconf)
   val1 = MMult(deltas, mmedias1)
   val2 = MMult(MMult(deltas, cov1), MTranD(deltas))
   valvar1 = val1(1, 1) * htiempo + valfa1 * Sqr(htiempo * val2(1, 1))
   valvar2 = val1(1, 1) * htiempo + valfa2 * Sqr(htiempo * val2(1, 1))
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO ='" & txtport & "'"
   txtborra = txtborra & " AND SUBPORT = '" & txtsubport & "'"
   txtborra = txtborra & " AND TVAR = '" & txttvar & "'"
   ConAdo.Execute txtborra
   txtinserta = "INSERT INTO " & TablaResVaR & " VALUES("
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & "'" & txtport & "',"
   txtinserta = txtinserta & "'" & txtsubport & "',"
   txtinserta = txtinserta & "'" & txtportfr & "',"
   txtinserta = txtinserta & "'" & txttvar & "',"
   txtinserta = txtinserta & nconf & ","
   txtinserta = txtinserta & noesc & ","
   txtinserta = txtinserta & "0,"
   txtinserta = txtinserta & htiempo & ","
   txtinserta = txtinserta & "0,"
   txtinserta = txtinserta & valvar1 & ")"
   ConAdo.Execute txtinserta
   txtinserta = "INSERT INTO " & TablaResVaR & " VALUES("
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & "'" & txtport & "',"
   txtinserta = txtinserta & "'" & txtsubport & "',"
   txtinserta = txtinserta & "'" & txtportfr & "',"
   txtinserta = txtinserta & "'" & txttvar & "',"
   txtinserta = txtinserta & 1 - nconf & ","
   txtinserta = txtinserta & noesc & ","
   txtinserta = txtinserta & "0,"
   txtinserta = txtinserta & htiempo & ","
   txtinserta = txtinserta & "0,"
   txtinserta = txtinserta & valvar2 & ")"
   ConAdo.Execute txtinserta
   For i = 1 To noreg
       txtact = "UPDATE " & TablaSensibPort & " SET VOLATIL = " & Sqr(Abs(cov1(i, i)))
       txtact = txtact & " WHERE FECHA = " & txtfecha
       txtact = txtact & " AND PORTAFOLIO = '" & txtport & "' "
       txtact = txtact & " AND PORT_FR = '" & txtportfr & "' "
       txtact = txtact & " AND SUBPORT = '" & txtsubport & "' "
       txtact = txtact & " AND FACTOR = '" & mata(i, 1) & "'"
       ConAdo.Execute txtact
   Next i
   exito = True
Else
   exito = False
End If

End Sub

Function LeerSensibPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaSensibPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND PORT_FR = '" & txtportfr & "' AND SUBPORT = '" & txtsubport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("FACTOR")    'nombre del factor
       mata(i, 2) = rmesa.Fields("VALOR")     'valor
       mata(i, 3) = rmesa.Fields("DERIVADA")  'derivada
       mata(i, 4) = rmesa.Fields("VOLATIL")   'volatilidad
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerSensibPort = mata
End Function

Function EsSensibleaFactor(ByVal txtfr As String, _
                           ByVal cposicion As Integer, _
                           ByVal coperacion As String, _
                           ByRef matpos() As propPosRiesgo, _
                           ByRef matposmd() As propPosMD, _
                           ByRef matposdiv() As propPosDiv, _
                           ByRef matposswaps() As propPosSwaps, _
                           ByRef matposfwds() As propPosFwd, _
                           ByRef matposdeuda() As propPosDeuda)

Dim matfr() As String
Dim i As Integer
Dim j As Integer
Dim indice0 As Integer

EsSensibleaFactor = False
For i = 1 To UBound(matpos, 1)
    If matpos(i).C_Posicion = cposicion And matpos(i).c_operacion = coperacion Then
       indice0 = matpos(i).IndPosicion
       If matpos(i).No_tabla = 1 Then   'md
          ReDim matfr(1 To 5) As String
          matfr(1) = matposmd(indice0).fRiesgo1MD
          matfr(2) = matposmd(indice0).fRiesgo2MD
          matfr(3) = matposmd(indice0).fRiesgo3MD
          matfr(4) = matposmd(indice0).fRiesgo4MD
          If Not EsVariableVacia(matposmd(indice0).tCambioMD) Then
             matfr(5) = matposmd(indice0).tCambioMD
          Else
             matfr(5) = ""
          End If
       ElseIf matpos(i).No_tabla = 2 Then      'divisas
          ReDim matfr(1 To 1) As String
          matfr(1) = matposdiv(indice0).TCambioDiv
       ElseIf matpos(i).No_tabla = 3 Then      'swaps
          ReDim matfr(1 To 6) As String
          matfr(1) = matposswaps(indice0).FRiesgo1Swap
          matfr(2) = matposswaps(indice0).FRiesgo2Swap
          If Not EsVariableVacia(matposswaps(indice0).FRiesgo3Swap) Then
             matfr(3) = matposswaps(indice0).FRiesgo3Swap
          Else
             matfr(3) = ""
          End If
          If Not EsVariableVacia(matposswaps(indice0).FRiesgo4Swap) Then
             matfr(4) = matposswaps(indice0).FRiesgo4Swap
          Else
             matfr(4) = ""
          End If
          matfr(5) = matposswaps(indice0).TCambio1Swap
          matfr(6) = matposswaps(indice0).TCambio2Swap
       ElseIf matpos(i).No_tabla = 4 Then        'forwards
          ReDim matfr(1 To 4) As String
          matfr(1) = matposfwds(indice0).FRiesgo1Fwd
          matfr(2) = matposfwds(indice0).FRiesgo2Fwd
          If Not EsVariableVacia(matposfwds(indice0).FRiesgo3Fwd) Then
             matfr(3) = matposfwds(indice0).FRiesgo3Fwd
          Else
             matfr(3) = ""
          End If
          matfr(4) = matposfwds(indice0).TCambioFwd
       ElseIf matpos(i).No_tabla = 5 Then        'deuda
          ReDim matfr(1 To 3) As String
          matfr(1) = matposdeuda(indice0).FRiesgo1Deuda
          matfr(2) = matposdeuda(indice0).FRiesgo2Deuda
          If Not EsVariableVacia(matposdeuda(indice0).TCambioDeuda) Then
             matfr(3) = matposdeuda(indice0).TCambioDeuda
          Else
             matfr(3) = ""
          End If
       End If
       For j = 1 To UBound(matfr, 1)
           If matfr(j) = txtfr Then
              EsSensibleaFactor = True
              Exit Function
           End If
       Next j
    End If
Next i
End Function

Function DetSiFRenMat(ByVal txtf1 As String, ByVal txtf2 As Double, ByRef matf1() As Variant, ByRef matf2() As Variant)
Dim i As Integer

For i = 1 To UBound(matf1, 1)
If txtf1 = matf1(i) And txtf2 = Val(matf2(i)) Then
   DetSiFRenMat = True
   Exit Function
End If
Next i
DetSiFRenMat = False
End Function

Sub DefinirTasasSimulacion(ByRef matpos() As Variant)
Dim nreg As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
nreg = UBound(matpos, 1)
ReDim VecTasasSimulacion(1 To 1, 1 To nreg + nofriesgo) As Double
For i = 1 To NoFactores
VecTasasSimulacion(1, i) = MatFactR1(i, 1)
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub SecuenciaBack(ByVal fecha As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim indice As Integer
Dim i As Integer
Dim j As Integer
Dim fecha1 As Date
Dim txtmsg1 As String
Dim txtmsg2 As String

txtmsg = ""
exito = False
exito3 = True
ValExacta = True
indice = BuscarValorArray(fecha, MatFechasVaR, 1)
If indice <> 0 Then
   'determinar la fecha previa
   fecha1 = CDate(MatFechasVaR(indice - 1, 1))
   Call CalcularBacktesting(fecha1, fecha1, txtport, txtmsg1, exito1)
   If exito1 Then
      MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
      If UBound(MatGruposPortPos, 1) <> 0 Then
         For j = 1 To UBound(MatGruposPortPos, 1)
             Call CalcResBackPortPos(fecha1, txtport, MatGruposPortPos(j, 3))
             DoEvents
             AvanceProc = j / UBound(MatGruposPortPos, 1)
         Next j
         Call CompararPyGvsCVaR(fecha1, MatGruposPortPos, txtmsg2, exito2)
         exito3 = exito3 And exito2
         If Not exito2 Then txtmsg = txtmsg & " " & txtmsg2
         If exito3 Then txtmsg = "El proceso finalizo correctamente"
         exito = True
      Else
         txtmsg = "El portafolio no esta definido"
         exito = False
      End If
   End If
Else
   txtmsg = "No se pudo realizar el calculo del backtesting"
   exito = False
End If
End Sub

Sub CompararPyGvsCVaR(ByVal fecha As Date, ByRef matgrp() As Variant, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim ValorVAR As Double
Dim valorpl As Double
Dim i As Integer
Dim noesc As Integer
Dim rmesa As New ADODB.recordset

noesc = 500
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To UBound(matgrp, 1)
    ValorVAR = 0: valorpl = 0
    txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = 'TOTAL'"
    txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matgrp(i, 3) & "'"
    txtfiltro2 = txtfiltro2 & " AND NCONF = .03 AND NOESC = " & noesc & " AND HTIEMPO = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       ValorVAR = rmesa.Fields("VALOR")
       rmesa.Close
    End If
    txtfiltro2 = "SELECT * FROM " & TablaBackPort & " WHERE FECHA = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = 'NEGOCIACION + INVERSION'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matgrp(i, 3) & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       valorpl = rmesa.Fields("DIFERENCIA")
       rmesa.Close
    End If
    If ValorVAR <> 0 And valorpl <> 0 Then
       If ValorVAR > valorpl Then
          txtmsg = "El backtesting se violo para el portafolio " & matgrp(i, 3)
          exito = False
          Exit Sub
       End If
    End If
Next i
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub

Sub SubprocCalculoPyGPosicion(ByVal f_pos As Date, ByVal f_val As Date, ByVal f_factor As Date, ByVal txtnompos As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
        Dim exito1 As Boolean
        Dim exito2 As Boolean
        Dim txtmsg0 As String
        Dim txtmsg2 As String
        Dim fecha1 As Date
        Dim matvalt0() As Double
        Dim matpygd() As Double
        Dim matpos() As New propPosRiesgo
        Dim matposmd() As New propPosMD
        Dim matposdiv() As New propPosDiv
        Dim matposswaps() As New propPosSwaps
        Dim matposfwd() As New propPosFwd
        Dim matposdeuda() As New propPosDeuda
        Dim matflswap() As New estFlujosDeuda
        Dim matfldeuda() As New estFlujosDeuda
        Dim indice As Integer
        Dim mattxt() As String
        Dim exitofr As Boolean
        
        
        exito = False
        Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
        If txtnompos <> txtNomPosRiesgo Or EsArrayVacio(matpos) Then
           mattxt = CrearFiltroPosSim(txtnompos)
           Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
           FechaPosRiesgo = f_pos
           txtNomPosRiesgo = txtnompos
        End If
        If exito1 Then
           Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
           If exito2 Then
              Call RutinaCargaFR(f_factor, exito)
              Call AnexarDatosVPrecios(f_val, matposmd)
              Call RutinaVaRHistórico1(f_factor, f_val, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
              Call GuardarEscHist(f_pos, f_val, f_factor, txtnompos, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito2)
              exito = True
              txtmsg = "El proceso finalizo correctamente"
           Else
              exito = False
              txtmsg = txtmsg2
           End If
        Else
          txtmsg = txtmsg2
          exito = False
        End If
End Sub

Sub SubprocCalculoPyGPort(ByVal f_pos As Date, ByVal f_val As Date, ByVal f_factor As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
        Dim exito1 As Boolean
        Dim exito2 As Boolean
        Dim fecha1 As Date
        Dim matpos() As New propPosRiesgo
        Dim matposmd() As New propPosMD
        Dim matposdiv() As New propPosDiv
        Dim matposswaps() As New propPosSwaps
        Dim matposfwd() As New propPosFwd
        Dim matposdeuda() As New propPosDeuda
        Dim matflswap() As New estFlujosDeuda
        Dim matfldeuda() As New estFlujosDeuda
        Dim matvalt0() As Double
        Dim matpygd() As Double
        Dim indice As Integer
        Dim mattxt() As String
        Dim txtmsg0 As String
        Dim txtmsg2 As String
        Dim exitofr As Boolean

        exito = False
        Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
        If f_pos <> FechaPosRiesgo Or txtport <> txtNomPosRiesgo Or EsArrayVacio(matpos) Then
           mattxt = CrearFiltroPosPort(f_pos, txtport)
           Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
           FechaPosRiesgo = f_pos
           txtNomPosRiesgo = txtport
        Else
           exito1 = True
        End If
        If exito1 Then
           Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
           If exito2 Then
              Call RutinaCargaFR(f_factor, exito)
              Call AnexarDatosVPrecios(f_val, matposmd)
              Call RutinaVaRHistórico1(f_factor, f_val, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
              Call GuardarEscHist(f_pos, f_factor, f_val, txtport, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito2)
              exito = True
              txtmsg = "El proceso finalizo correctamente"
           Else
              exito = False
              txtmsg = txtmsg2
           End If
         Else
           txtmsg = "No hay datos en la posicion"
           exito = False
        End If
End Sub

Sub SubprocCalculoPyGPos(ByVal f_pos As Date, ByVal f_val As Date, ByVal f_factor As Date, ByVal txtnompos As String, ByVal txtescfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
        Dim exito1 As Boolean
        Dim exito2 As Boolean
        Dim fecha1 As Date
        Dim matpos() As New propPosRiesgo
        Dim matposmd() As New propPosMD
        Dim matposdiv() As New propPosDiv
        Dim matposswaps() As New propPosSwaps
        Dim matposfwd() As New propPosFwd
        Dim matposdeuda() As New propPosDeuda
        Dim matflswap() As New estFlujosDeuda
        Dim matfldeuda() As New estFlujosDeuda
        Dim matvalt0() As Double
        Dim matpygd() As Double
        Dim indice As Integer
        Dim mattxt() As String
        Dim txtmsg0 As String
        Dim txtmsg2 As String
        Dim exitofr As Boolean

        exito = False
        Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
        If f_pos <> FechaPosRiesgo Or txtnompos <> txtNomPosRiesgo Or EsArrayVacio(matpos) Then
           mattxt = CrearFiltroPosSim(txtnompos)
           Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
           FechaPosRiesgo = f_pos
           txtNomPosRiesgo = txtnompos
        Else
           exito1 = True
        End If
        If exito1 Then
           Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
           If exito2 Then
              Call RutinaCargaFR(f_factor, exito)
              Call AnexarDatosVPrecios(f_val, matposmd)
              Call RutinaVaRHistórico1(f_factor, f_val, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
              Call GuardarEscHist(f_pos, f_factor, f_val, txtnompos, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito2)
              exito = True
              txtmsg = "El proceso finalizo correctamente"
           Else
              exito = False
              txtmsg = txtmsg2
           End If
        Else
          txtmsg = "No hay datos en la posicion"
          exito = False
        End If
End Sub

Sub CalcPyG1Oper(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
        Dim exito1 As Boolean
        Dim exito2 As Boolean
        Dim fecha1 As Date
        Dim matpos() As New propPosRiesgo
        Dim matposmd() As New propPosMD
        Dim matposdiv() As New propPosDiv
        Dim matposswaps() As New propPosSwaps
        Dim matposfwd() As New propPosFwd
        Dim matposdeuda() As New propPosDeuda
        Dim matflswap() As New estFlujosDeuda
        Dim matfldeuda() As New estFlujosDeuda
        Dim matvalt0() As Double
        Dim matpygd() As Double
        Dim indice As Integer
        Dim mattxt() As String
        Dim matfechassh() As Date
        Dim txtmsg0 As String
        Dim txtmsg2 As String
        Dim exitofr As Boolean
        final = False
        exito = False
        Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
        mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
        Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
        If exito1 And UBound(matpos, 1) <> 0 Then
           Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
           If exito2 Then
              Call AnexarDatosVPrecios(f_val, matposmd)
              Call CalcEscHist(f_val, f_factor, htiempo, noesc, 0, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
              Call GuardarEscHist2(f_pos, f_factor, f_val, txtport, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito2)
              exito = True
              txtmsg = "El proceso finalizo correctamente"
           Else
              exito = False
              txtmsg = txtmsg2
           End If
        Else
           txtmsg = txtmsg0
           exito = False
        End If
        final = True
End Sub

Sub CalcPyG1OperVR(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal pfwd As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
        Dim exito1 As Boolean
        Dim exito2 As Boolean
        Dim fecha1 As Date
        Dim matpos() As New propPosRiesgo
        Dim matposmd() As New propPosMD
        Dim matposdiv() As New propPosDiv
        Dim matposswaps() As New propPosSwaps
        Dim matposfwd() As New propPosFwd
        Dim matposdeuda() As New propPosDeuda
        Dim matflswap() As New estFlujosDeuda
        Dim matfldeuda() As New estFlujosDeuda
        Dim matvalt0() As resValIns
        Dim matvalt1() As resValIns
        Dim matpygd() As Double
        Dim indice As Integer
        Dim mattxt() As String
        Dim matfechassh() As Date
        Dim txtmsg0 As String
        Dim txtmsg2 As String
        Dim mvalt1() As Double
        Dim exitofr As Boolean
        
        exito = False
        SiIncTasaCVig = False
        Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
        mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
        Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
        If exito1 And UBound(matpos, 1) <> 0 Then
           Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
           If exito2 Then
              Call AnexarDatosVPrecios(f_val, matposmd)
              Call CalcEscHistVR(f_val, f_factor, htiempo, noesc, pfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matvalt0, matvalt1, MatPyGT)
              Call GuardarEscHistVR(f_pos, txtport, txtescfr, noesc, htiempo, pfwd, matpos, matvalt0, matvalt1, MatPyGT)
              exito = True
              txtmsg = "El proceso finalizo correctamente"
           Else
              exito = False
              txtmsg = txtmsg2
           End If
        Else
           txtmsg = "No hay datos en la posicion"
           exito = False
        End If
        SiIncTasaCVig = True
End Sub


Sub VerifCargaFR(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim fechax As Date
Dim exito As Boolean
fechax = PBD1(fecha2, 1, "MX")
    If EsArrayVacio(MatFactRiesgo) Then
       Call CrearMatFRiesgo2(fecha1, fecha2, MatFactRiesgo, "", exito)
    Else
        If fecha1 < MatFactRiesgo(1, 1) Or fechax > MatFactRiesgo(UBound(MatFactRiesgo, 1), 1) Then
           Call CrearMatFRiesgo2(fecha1, fecha2, MatFactRiesgo, "", exito)
        End If
    End If
End Sub

Sub VerifCargaFR3(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim fechax As Date
Dim exito As Boolean
fechax = PBD1(fecha2, 1, "MX")
    If EsArrayVacio(MatFactRiesgo) Then
       Call CrearMatFRiesgo(fecha1, fecha2, MatFactRiesgo, "", exito)
    Else
        If fecha1 < MatFactRiesgo(1, 1) Or fechax > MatFactRiesgo(UBound(MatFactRiesgo, 1), 1) Then
           Call CrearMatFRiesgo(fecha1, fecha2, MatFactRiesgo, "", exito)
        End If
    End If
End Sub

Sub VerifCargaFR2(ByVal fecha As Date, ByVal noreg As Long, ByRef exito As Boolean)
Dim siesfv As Boolean
Dim fecha0 As Date
    siesfv = EsFechaVaR(fecha)
    If siesfv Then
       If EsArrayVacio(MatFactRiesgo) Then
          fecha0 = DetFechaFNoEsc(fecha, noreg)
          Call CrearMatFRiesgo2(fecha0, fecha, MatFactRiesgo, "", exito)
       Else
          fecha0 = DetFechaFNoEsc(fecha, noreg)
          If fecha0 < MatFactRiesgo(1, 1) Or fecha > MatFactRiesgo(UBound(MatFactRiesgo, 1), 1) Then
             Call CrearMatFRiesgo2(fecha0, fecha, MatFactRiesgo, "", exito)
          End If
       End If
    Else
       MsgBox "La fecha " & fecha & " no se encontro en la base dias habiles"
    End If
End Sub

Sub CalcEscEstresOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByRef txtmsg As String, ByRef exito As Boolean)
         Dim exito1 As Boolean
         Dim exito2 As Boolean
         Dim fecha1 As Date
         Dim mattxt() As String
         Dim matpos() As New propPosRiesgo
         Dim matposmd() As New propPosMD
         Dim matposdiv() As New propPosDiv
         Dim matposswaps() As New propPosSwaps
         Dim matposfwd() As New propPosFwd
         Dim matposdeuda() As New propPosDeuda
         Dim matflswap() As New estFlujosDeuda
         Dim matfldeuda() As New estFlujosDeuda
         Dim txtmsg0 As String
         Dim txtmsg2 As String
         Dim noesc As Integer
         Dim exitofr As Boolean
         noesc = 501
         exito = False
        'se carga la posición, los factores de riesgo y se realiza la valuacion
         Call VerifCargaFR2(fecha, noesc, exitofr)
         mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
         Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
         Call AnexarDatosVPrecios(fecha, matposmd)
         If UBound(matpos, 1) <> 0 Then
            Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
            If exito2 Then
               If FechaMatFactR1 <> fecha Then
                  MatFactR1 = CargaFR1Dia(fecha, exito1)
                  FechaMatFactR1 = fecha
               End If
               Call Rutina2EscEstres(fecha, txtport, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, exito1)
               txtmsg = "El proceso finalizo correctamente"
               exito = True
           Else
               txtmsg = txtmsg2
               exito = False
           End If
         Else
            MensajeProc = "No hay datos en la posicion del " & fecha
            txtmsg = MensajeProc
            exito = False
         End If


End Sub

Sub CalcEscEstresSubPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByVal txtesc As String, ByRef matres() As Variant, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim i As Long
Dim exito1 As Boolean
Dim txtmsg1 As String
    exito = True
    MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
    If UBound(MatGruposPortPos, 1) <> 0 Then
       For i = 1 To UBound(MatGruposPortPos, 1)
           Call GenerarEscEstresPort(fecha, txtesc, txtport, MatGruposPortPos(i, 3), matres, txtmsg1, exito1)
           exito = exito And exito1
           If Not exito1 Then
               txtmsg = txtmsg & "," & txtmsg1
           End If
       Next i
       If exito Then
          txtmsg = "El proceso finalizo correctamente"
       Else
          txtmsg = "No se calculo un subportafolio"
       End If
    Else
       txtmsg = "No esta definido el portafolio"
       exito = False
    End If
    final = True
End Sub

Sub SubSecVaRMontecarlo(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim indice As Integer
Dim fecha1 As Date
Dim exito1 As Boolean
Dim mattxt() As String
Dim txtfecha As String
Dim txtborra As String
Dim matorden2() As Variant
Dim mmedias() As Double
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim exito2 As Boolean
Dim matnuma() As String
Dim exitofr As Boolean
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda

Call VerifCargaFR2(fecha, noesc + htiempo, exitofr)
If exito1 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtborra = "DELETE FROM " & TablaPyGMontOper & " WHERE FECHA = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND ESC_FACTORES = '" & txtportfr & "'"
   txtborra = txtborra & " AND NOESC = " & noesc
   txtborra = txtborra & " AND NOSIM  = " & nosim
   txtborra = txtborra & " AND HTIEMPO = " & htiempo
   ConAdo.Execute txtborra
   exito = False
   'se carga la posición, los factores de riesgo y se realiza la valuacion
   If fecha <> FechaPosRiesgo Or txtport <> txtNomPosRiesgo Or EsArrayVacio(matpos) Then
      mattxt = CrearFiltroPosPort(fecha, txtport)
      Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
      FechaPosRiesgo = fecha
      txtNomPosRiesgo = txtport
   End If
   If UBound(matpos, 1) <> 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
      Call RutinaCargaFR(fecha, exito)
      Call AnexarDatosVPrecios(fecha, matposmd)
     'el var montecarlo
      matnuma = LeerMuestraNormal(#1/1/2018#, 700, nosim)
      Call CalculoVaRMontecarlo(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matordenMont, nosim, htiempo, mmediasMont, MatCholeski, matnuma, 1, MatVal0T, MatPyGT)
      Call GuardarEscMont(fecha, txtport, txtportfr, noesc, htiempo, nosim, matpos, MatVal0T, MatPyGT)
      exito = True
      txtmsg = "El proceso finalizo correctamente"
   Else
      exito = False
   End If
Else
   exito = False
   txtmsg = "No hay sensibilidades"
End If
End Sub

Sub CalcPyGMontOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal tipopos As Integer, ByVal fechareg As Date, txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim indice As Integer
Dim fecha1 As Date
Dim exito1 As Boolean
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim txtfecha As String
Dim txtborra As String
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim exito2 As Boolean
Dim exitofr As Boolean
final = False
Call VerifCargaFR2(fecha, 1, exitofr)
mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
If nosim <> NoSimMont Or EsArrayVacio(MatNumaMont) Then
   MatNumaMont = LeerMuestraNormal(#1/1/2018#, 700, nosim)
   NoSimMont = nosim
End If
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
If EsArrayVacio(MatCholeski) Or EsArrayVacio(matordenMont) Or EsArrayVacio(mmediasMont) Then
   Call LeerMatCholeski(fecha, txtport, matordenMont, mmediasMont, MatCholeski)
End If
If UBound(matpos, 1) <> 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
      Call RutinaCargaFR(fecha, exito)
      Call AnexarDatosVPrecios(fecha, matposmd)
     'el var montecarlo
      Call CalculoVaRMontecarlo(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matordenMont, nosim, htiempo, mmediasMont, MatCholeski, MatNumaMont, 1, MatVal0T, MatPyGT)
      Call GuardarEscMont(fecha, txtport, txtportfr, noesc, htiempo, nosim, matpos, MatVal0T, MatPyGT)
      exito = True
      txtmsg = "El proceso finalizo correctamente"
Else
   exito = False
End If
final = True
End Sub

Sub LeerMatCholeski(ByVal fecha As Date, ByVal txtport As String, ByRef matorden() As Variant, ByRef mmedias() As Double, ByRef match() As Double)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim m_orden As String
Dim m_medias As String
Dim m_choleski As String
Dim mata() As String
Dim matb() As String
Dim matc() As String
Dim i As Long
Dim j As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaFactChol & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   noreg1 = rmesa.Fields("N1")
   noreg2 = rmesa.Fields("N2")
   m_orden = rmesa.Fields("M_ORDEN")
   m_medias = rmesa.Fields("M_MEDIAS")
   m_choleski = rmesa.Fields("M_CHOLESKI")
   ReDim match(1 To noreg1, 1 To noreg2) As Double
   mata = EncontrarSubCadenas(m_orden, ",")
   matb = EncontrarSubCadenas(m_medias, ",")
   matc = EncontrarSubCadenas(m_choleski, ",")
   ReDim matorden(1 To UBound(mata, 1), 1 To 1) As Variant
   For i = 1 To UBound(matb, 1)
       matorden(i, 1) = CStr(mata(i))
   Next i
   ReDim mmedias(1 To UBound(matb, 1), 1 To 1) As Double
   For i = 1 To UBound(matb, 1)
       mmedias(i, 1) = CDbl(matb(i))
   Next i
   For i = 1 To noreg1
       For j = 1 To noreg2
           match(i, j) = CDbl(matc((i - 1) * noreg2 + j))
       Next j
   Next i
   rmesa.Close
End If
End Sub

Sub GeneraResCVaRPos(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String

Dim valor As Double
Dim txtborra As String
Dim txttvar As String
Dim i As Integer
Dim j As Integer
Dim noport As Integer
Dim noreg As Long
Dim exito1 As Boolean
Dim exito2 As Boolean

txttvar = "CVARH"
exito = False
txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha1
txtborra = txtborra & " AND F_FACTORES = " & txtfecha2
txtborra = txtborra & " AND F_VALUACION = " & txtfecha3
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
ConAdo.Execute txtborra, noreg
txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha1
txtborra = txtborra & " AND F_FACTORES = " & txtfecha2
txtborra = txtborra & " AND F_VALUACION = " & txtfecha3
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
ConAdo.Execute txtborra, noreg
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For j = 1 To UBound(MatGruposPortPos, 1)
       valor = CalcularCVaRPyG(f_pos, f_factor, f_val, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(f_pos, f_factor, f_val, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(f_pos, f_factor, f_val, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(f_pos, f_factor, f_val, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
   Next j
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "El portafolio no esta definido"
   exito = False
End If
End Sub

Sub GeneraResCVaRPrevPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal noesc2 As Integer, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim valor As Double
Dim txtborra As String
Dim txtportfr As String
Dim txttvar As String
Dim j As Integer

exito = False
txtportfr = "Normal"
txttvar = "CVARPrev"

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "' AND NOESC = " & noesc & " AND TVAR ='" & txttvar & "'"
ConAdo.Execute txtborra
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For j = 1 To UBound(MatGruposPortPos, 1)
       valor = CalcularCVaRPrev(fecha, txtport, MatGruposPortPos(j, 3), noesc, htiempo, noesc2, 1 - nconf, exito)
       Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc2, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPrev(fecha, txtport, MatGruposPortPos(j, 3), noesc, htiempo, noesc2, nconf, exito)
       Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, "CVARPrev", noesc2, htiempo, 0, nconf, 0, valor)
   Next j
   exito = True
   txtmsg = "El proceso finalizo correctamente"
Else
   exito = False
   txtmsg = "No esta definido el portafolio"
End If
End Sub

Sub GeneraResCVaRExpPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal lambda As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim valor As Double
Dim txtborra As String
Dim txttvar As String
Dim i As Integer
Dim j As Integer

txttvar = "VARExp"
exito = False
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND LAMBDA = " & lambda & " AND TVAR= '" & txttvar & "'"
ConAdo.Execute txtborra
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For j = 1 To UBound(MatGruposPortPos, 1)
       valor = CalcularVaRExp(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, 1 - nconf, lambda, exito)
       Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, lambda, valor)
       valor = CalcularVaRExp(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, nconf, lambda, exito)
       Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, 0, nconf, lambda, valor)
   Next j
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No esta definido el portafolio"
   exito = False
End If
End Sub

Sub GeneraResVaRMontPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim valor As Double
Dim txtborra As String
Dim txttvar As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim i As Integer
Dim j As Integer

txttvar = "VARMont"
   
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND NOSIM = " & nosim & " AND TVAR= '" & txttvar & "'"
   ConAdo.Execute txtborra
   MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
   If UBound(MatGruposPortPos, 1) <> 0 Then
      For j = 1 To UBound(MatGruposPortPos, 1)
          valor = CalcularVaRMont(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, nosim, 1 - nconf, exito1)
          If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, nosim, 1 - nconf, 0, valor)
          valor = CalcularVaRMont(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, nosim, nconf, exito2)
          If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, MatGruposPortPos(j, 3), txtportfr, txttvar, noesc, htiempo, nosim, nconf, 0, valor)
      Next j
      exito = True
      txtmsg = "El proceso finalizo correctamente"
   Else
      exito = False
      txtmsg = "No esta definido el portafolio"
   End If
End Sub

Sub InsertaRegVaR(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal txtescfr As String, ByVal txttvar As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByVal nconf As Double, ByVal lambda As Double, ByVal valor As Double)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String
Dim txtborra As String
    txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha1
    txtborra = txtborra & " AND F_FACTORES = " & txtfecha2
    txtborra = txtborra & " AND F_VALUACION = " & txtfecha3
    txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
    txtborra = txtborra & " AND SUBPORT = '" & txtsubport & "'"
    txtborra = txtborra & " AND ESC_FACTORES = '" & txtescfr & "'"
    txtborra = txtborra & " AND TVAR = '" & txttvar & "'"
    txtborra = txtborra & " AND NOESC = " & noesc
    txtborra = txtborra & " AND HTIEMPO = " & htiempo
    txtborra = txtborra & " AND NCONF = " & nconf
    txtborra = txtborra & " AND LAMBDA = " & lambda
    ConAdo.Execute txtborra


    txtcadena = "INSERT INTO " & TablaResVaR & " VALUES("
    txtcadena = txtcadena & txtfecha1 & ","
    txtcadena = txtcadena & txtfecha2 & ","
    txtcadena = txtcadena & txtfecha3 & ","
    txtcadena = txtcadena & "'" & txtport & "',"
    txtcadena = txtcadena & "'" & txtsubport & "',"
    txtcadena = txtcadena & "'" & txtescfr & "',"
    txtcadena = txtcadena & "'" & txttvar & "',"
    txtcadena = txtcadena & nconf & ","
    txtcadena = txtcadena & noesc & ","
    txtcadena = txtcadena & nosim & ","
    txtcadena = txtcadena & htiempo & ","
    txtcadena = txtcadena & lambda & ","
    txtcadena = txtcadena & valor & ")"
    ConAdo.Execute txtcadena
End Sub


Sub CalcValReemplazo(ByVal dtfecha As Date, ByVal nconf As Double, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal pfwd As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim mrvalflujo() As resValFlujo
    Dim txtsubport As String
    Dim txtfecha As String
    Dim txtborra As String
    Dim txtinserta As String
    Dim txttfondeo As String
    Dim tfondeo As Double
    Dim bl_exito As Boolean
    Dim txtport As String
    Dim noreg As Integer
    Dim sicargarvp As Boolean
    Dim indice0  As Integer
    Dim nocontrap As Integer
    Dim fechaid As Date
    Dim matresvar() As Double, matresvart() As Double
    Dim i As Integer
    Dim j As Integer
    Dim clave As String
    Dim contar As Integer
    Dim parval As ParamValPos
    Dim matval0() As New resValIns
    Dim matval() As New resValIns
    Dim valor As Double
    Dim mattxt() As String
    Dim id_contrap As String
    Dim CSector As String
    Dim matvald0() As Double
    Dim matpygd() As Double
    Dim fecha1 As Date
    Dim mtm0 As Double
    Dim mtm1 As Double
    Dim cvar1 As Double
    Dim cvar2 As Double
    Dim exito2 As Boolean
    Dim exito3 As Boolean
    Dim txtmsg0 As String
    Dim txtmsg2 As String
    Dim txtmsg3 As String
    Dim exitofr As Boolean

    
    'primero se va a realizar un calculo de VaR este calculo se va a desglosar
    'operacion por operacion
    'una vez definido el desglose del VaR se procede a calcular el VaR
    ValExacta = True
    txttfondeo = InputBox("Dame la tasa de fondeo", , 0)
    tfondeo = Val(txttfondeo) / 100
    txtport = "DERIVADOS"
    Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
    Call VerifCargaFR2(dtfecha, noesc + htiempo, exitofr)
    mattxt = CrearFiltroPosPort(dtfecha, txtport)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, bl_exito)
    noreg = UBound(matpos, 1)
    If noreg <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       MatFactR1 = CargaFR1Dia(dtfecha, bl_exito)
       'se carga todo el vector completo de curvas para ese dia, el
       If dtfecha <> FechaArchCurvas Or EsArrayVacio(MatCurvasT) Then
          FechaArchCurvas = dtfecha
          MatCurvasT = LeerCurvaCompleta(dtfecha, bl_exito)
       End If
       Set parval = DeterminaPerfilVal("VALUACION")
        'primero la valuacion en t0
       parval.perfwd = 0
       matval0 = CalcValuacion(dtfecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
        'despues la valuacion a dias forward
       parval.perfwd = pfwd
       matval = CalcValuacion(dtfecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
        'se captura la valuacion del sistema, mt
        'corre el motor de calculo que realiza el calculo de los escenarios historicos
       MatFactRiesgo = MatFactRiesgo
       Call CalcEscHist(dtfecha, dtfecha, htiempo, noesc, pfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
       Call GuardarEscHistVR(dtfecha, txtport, "V_R", noesc, htiempo, pfwd, matpos, matval0, matval, MatPyGT)
       nocontrap = UBound(MatContrapartes, 1)
       ReDim matres(1 To nocontrap + 3, 1 To 8) As Variant
       For i = 1 To UBound(MatContrapartes, 1)
           txtsubport = "Deriv Contrap " & MatContrapartes(i, 1)
           matres(i, 1) = MatContrapartes(i, 1)                   'id contraparte
           matres(i, 2) = MatContrapartes(i, 3)                   'descripcion
           matres(i, 3) = MatContrapartes(i, 6)                   'sector
           Call GenPyGPortVR(dtfecha, txtport, "V_R", txtsubport, noesc, htiempo, exito)
           Call CalcularCVaR_VR(dtfecha, txtport, "V_R", txtsubport, noesc, htiempo, 1 - nconf, exito, mtm0, mtm1, cvar1)
           Call CalcularCVaR_VR(dtfecha, txtport, "V_R", txtsubport, noesc, htiempo, nconf, exito, mtm0, mtm1, cvar2)
           matres(i, 4) = mtm0
           matres(i, 5) = mtm1
           matres(i, 6) = cvar1
           matres(i, 7) = cvar2
           If matres(i, 5) > 0 Then
              matres(i, 8) = matres(i, 5) - matres(i, 6)
           Else
              matres(i, 8) = matres(i, 5) - matres(i, 7)
           End If
       Next i
       ReDim matb(1 To 3) As String
       matb(1) = "DERIV SECT FINANCIERO"
       matb(2) = "DERIV SECT NO FINANCIERO"
       matb(3) = txtport
       For i = 1 To 3
           matres(i + nocontrap, 1) = matb(i)
           matres(i + nocontrap, 2) = matb(i)
           matres(i + nocontrap, 3) = ""
           Call GenPyGPortVR(dtfecha, txtport, "V_R", matb(i), noesc, htiempo, exito)
           Call CalcularCVaR_VR(dtfecha, txtport, "V_R", matb(i), noesc, htiempo, 1 - nconf, exito, mtm0, mtm1, cvar1)
           Call CalcularCVaR_VR(dtfecha, txtport, "V_R", matb(i), noesc, htiempo, nconf, exito, mtm0, mtm1, cvar2)
           matres(i + nocontrap, 4) = mtm0
           matres(i + nocontrap, 5) = mtm1
           matres(i + nocontrap, 6) = cvar1
           matres(i + nocontrap, 7) = cvar2
           If matres(i + nocontrap, 5) > 0 Then
              matres(i + nocontrap, 8) = matres(i + nocontrap, 5) - matres(i + nocontrap, 6)
           Else
              matres(i + nocontrap, 8) = matres(i + nocontrap, 5) - matres(i + nocontrap, 7)
           End If
       Next i
       txtfecha = "to_date('" & Format(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtborra = "DELETE FROM " & TablaResVReemplazo & " WHERE FECHA = " & txtfecha
       ConAdo.Execute txtborra
       For i = 1 To nocontrap + 3
           If matres(i, 8) <> 0 Then
              txtinserta = "INSERT INTO " & TablaResVReemplazo & " VALUES("
              txtinserta = txtinserta & txtfecha & ","                  'fecha
              txtinserta = txtinserta & Val(matres(i, 1)) & ","         'clave de contraparte
              txtinserta = txtinserta & "'" & matres(i, 2) & "',"       'descripcion
              txtinserta = txtinserta & "'" & matres(i, 3) & "',"       'sector
              txtinserta = txtinserta & matres(i, 4) & ","              'mtm t
              txtinserta = txtinserta & matres(i, 5) & ","              'mtm t+1
              txtinserta = txtinserta & matres(i, 6) & ","              'cvar t
              txtinserta = txtinserta & matres(i, 7) & ","              'cvar t+1
              txtinserta = txtinserta & matres(i, 8) & ")"              'valor de reemplazo
              ConAdo.Execute txtinserta
           End If
       Next i
       ValExacta = False
    End If

End Sub

Sub CalcCostoReemplazo(ByVal fecha As Date, ByVal tfondeo As Double, ByVal noesc As Integer, ByVal nconf As Double)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim suma As Double, suma1 As Double, suma2 As Double
Dim matpg() As String
Dim cvar1 As Double
Dim cvar2 As Double
Dim txtborra As String
Dim txtinserta As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM  " & TablaPLHistOperVR & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT count(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mv0(1 To noreg) As Double
   ReDim mv1(1 To noreg) As Double
   ReDim mvec(1 To noreg) As String
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mv0(i) = rmesa.Fields("VALT0")
       mv1(i) = rmesa.Fields("VALT1")
       mvec(i) = ReemplazaVacioValor(rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize), "")
       rmesa.MoveNext
   Next i
   rmesa.Close
   suma = 0
   suma1 = 0
   suma2 = 0
   ReDim matres2(1 To noreg, 1 To 5) As Double
   For i = 1 To noreg
       matres2(i, 1) = mv0(i)                         'marca a mercado en t
       matres2(i, 2) = mv1(i)                         'marca a mercado en t+1
       If matres2(i, 1) > 0 Then
          matres2(i, 3) = matres2(i, 2)           'marca a mercado en t+1 de las mtm positivas en t
       End If
       If matres2(i, 1) < 0 Then
          matres2(i, 4) = matres2(i, 2)        'marca a mercado en t+1 de las mtm no positivas en t
          matres2(i, 5) = matres2(i, 1)        'marca a mercado en t de las mtm no positivas en t
       End If
       suma1 = suma1 + matres2(i, 3)
       suma2 = suma2 + matres2(i, 4)
       suma = suma + matres2(i, 5)
   Next i
   ReDim matres3(1 To noesc, 1 To 2) As Variant
   For j = 1 To noreg
       matpg = EncontrarSubCadenas(mvec(j), ",")
       For i = 1 To noesc
           If matres2(j, 3) > 0 Then
              matres3(i, 1) = matres3(i, 1) + CDbl(matpg(i))
           Else
              matres3(i, 2) = matres3(i, 2) + CDbl(matpg(i))
           End If
       Next i
   Next j

   Dim matv1() As Variant
   Dim matv2() As Variant
   Dim vr1 As Double, vr2 As Double
   Dim cr1 As Double, cr2 As Double

   matv1 = ExtVecMatV(matres3, 1, 0)
   matv2 = ExtVecMatV(matres3, 2, 0)
   cvar1 = CPercentilCVaR(1 - nconf, ConvArVtDbl(matv1), 0, 0, True)         'var
   cvar2 = CPercentilCVaR(nconf, ConvArVtDbl(matv2), 0, 0, True)             'antivar
   vr1 = suma1 - cvar1
   vr2 = suma2 - cvar2

   cr1 = vr1 * tfondeo / 12
   cr2 = (vr2 - suma) * tfondeo / 12
   txtborra = "DELETE FROM " & TablaResCalcVReemplazo & " WHERE FECHA = " & txtfecha
   ConAdo.Execute txtborra
   txtinserta = "INSERT INTO " & TablaResCalcVReemplazo & " VALUES("
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & tfondeo & ","
   txtinserta = txtinserta & suma1 & ","
   txtinserta = txtinserta & suma2 & ","
   txtinserta = txtinserta & suma & ","
   txtinserta = txtinserta & cvar1 & ","
   txtinserta = txtinserta & cvar2 & ","
   txtinserta = txtinserta & vr1 & ","
   txtinserta = txtinserta & vr2 & ","
   txtinserta = txtinserta & cr1 & ","
   txtinserta = txtinserta & cr2 & ")"
   ConAdo.Execute txtinserta
End If
End Sub


Sub CalcPyGMontSubport(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
On Error GoTo hayerror
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim valor As Variant
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim matc() As String
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim txttexto As String
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPyGMontOper & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc & " AND NOSIM = " & nosim
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & " AND (FREGISTRO, CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT FECHAREG, CPOSICION, COPERACION FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matpl(1 To noesc) As Double
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma = 0
   For i = 1 To noreg
       valt01 = rmesa.Fields("VALT0")
       valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       suma = suma + valt01
       For j = 1 To noesc
           matpl(j) = matpl(j) + CDbl(matc(j))
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Generando las p y g del portafolio " & txtsubport & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   rmesa.Close
   txtborra = "DELETE FROM " & TablaPyGMontPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtsubport & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc & " AND NOSIM = " & nosim & " AND HTIEMPO = " & htiempo
   ConAdo.Execute txtborra
   txtfiltro = "SELECT * FROM " & TablaPyGMontPort
   RInterfIKOS.Open txtfiltro, ConAdo, 1, 3
   txtcadena = ""
   For j = 1 To UBound(matpl, 1) - 1
       txtcadena = txtcadena & matpl(j) & ","
   Next j
   txtcadena = txtcadena & matpl(UBound(matpl, 1))
   RInterfIKOS.AddNew
   RInterfIKOS.Fields(0) = CLng(fecha)                'la fecha de proceso
   RInterfIKOS.Fields(1) = txtport                    'el portafolio
   RInterfIKOS.Fields(2) = txtsubport                   'el subportafolio
   RInterfIKOS.Fields(3) = txtescfr                   'el escenario de factores de riesgo
   RInterfIKOS.Fields(4) = noesc                      'no de escenarios PARA VOL
   RInterfIKOS.Fields(5) = nosim                      'no de escenarios SIM
   RInterfIKOS.Fields(6) = htiempo                    'horizonte de tiempo
   RInterfIKOS.Fields(7) = suma                       'valuacion de escenario base
   Call GuardarElementoClob(txtcadena, RInterfIKOS, "DATOS")
   RInterfIKOS.Update
   RInterfIKOS.Close
End If
exito = True
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Function CalcularCVaRPyG(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef exito As Boolean) As Double
Dim matv() As Double

       matv = LeerPyGHistSubport(f_pos, f_factor, f_val, txtport, txtsubport, txtportfr, noesc, htiempo, 0)
       If UBound(matv, 1) <> 0 Then
          CalcularCVaRPyG = CPercentilCVaR(nconf, matv, 0, 0, True)
          exito = True
       Else
          CalcularCVaRPyG = 0
          exito = False
       End If

End Function

Function CalcularCVaRPyG2(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc1 As Integer, ByVal htiempo As Integer, ByVal noesc2 As Long, ByVal nconf As Double, ByRef exito As Boolean) As Double
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim l As Integer
Dim valt01 As Double
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

       txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc1 & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          ReDim matv(1 To noesc2, 1 To 1) As Double
          valt01 = rmesa.Fields(6)
          valor = rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To noesc2
              matv(l, 1) = CDbl(matc(noesc1 - noesc2 + l))
          Next l
          rmesa.Close
          CalcularCVaRPyG2 = CPercentilCVaR(nconf, matv, 0, 0, True)
          exito = True
       Else
          CalcularCVaRPyG2 = 0
          exito = False
       End If

End Function


Sub CalcularCVaR_VR(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef exito As Boolean, ByRef mtm0 As Double, ByRef mtm1 As Double, ByRef var1 As Double)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim l As Integer
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPyGHistPVR & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          mtm0 = rmesa.Fields("VALT0")
          mtm1 = rmesa.Fields("VALT1")
          ReDim matv(1 To noesc, 1 To 1) As Double
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              matv(l, 1) = CDbl(matc(l))
          Next l
          rmesa.Close
          var1 = CPercentilCVaR(nconf, matv, 0, 0, True)
          exito = True
       Else
          mtm0 = 0
          mtm1 = 0
          var1 = 0
          exito = False
       End If

End Sub




Function CalcularVaRMont(ByVal fecha As Date, _
                         ByVal txtport As String, _
                         ByVal txtportfr As String, _
                         ByVal txtsubport As String, _
                         ByVal noesc As Integer, _
                         ByVal htiempo As Integer, _
                         ByVal nosim As Integer, _
                         ByVal nconf As Double, _
                         ByRef exito As Boolean)
                         
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim l As Integer
Dim noreg As Integer
Dim valt01 As Double
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPyGMontPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "'"
       txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro2 = txtfiltro2 & " AND NOSIM  = " & nosim
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          ReDim matv(1 To noesc, 1 To 1) As Double
          valt01 = rmesa.Fields(7)
          valor = rmesa.Fields(8).GetChunk(rmesa.Fields(8).ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              matv(l, 1) = CDbl(matc(l))
          Next l
          rmesa.Close
          CalcularVaRMont = CPercentil2(nconf, matv, 0, 0, True)
          exito = True
       Else
          CalcularVaRMont = 0
          exito = False
       End If

End Function

Function CalcularVaRExp(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal lambda As Double, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim l As Integer
Dim valt01 As Double
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

       txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          ReDim matv(1 To noesc, 1 To 1) As Double
          valt01 = rmesa.Fields("VALT0")
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              matv(l, 1) = CDbl(matc(l))
          Next l
          rmesa.Close
          CalcularVaRExp = CalcVaRMark(matv, 1, nconf, lambda)
          exito = True
       Else
          CalcularVaRExp = 0
          exito = False
       End If

End Function

Function CalcularVaRExp2(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal noesc2 As Long, ByVal nconf As Double, ByVal lambda As Double, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim l As Integer
Dim valt01 As Double
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

       txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          ReDim matv(1 To noesc2, 1 To 1) As Double
          valt01 = rmesa.Fields(6)
          valor = rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To noesc2
              matv(l, 1) = CDbl(matc(noesc - noesc2 + l))
          Next l
          rmesa.Close
          CalcularVaRExp2 = CalcVaRMark(matv, 1, nconf, lambda)
          exito = True
       Else
          CalcularVaRExp2 = 0
          exito = False
       End If

End Function


Function CalcularCVaRPrev(ByVal f_pos As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal noesc1 As Integer, ByVal nconf As Double, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim l As Integer
Dim noesc3 As Integer
Dim txtportfr As String
Dim matval() As Double
Dim valt01 As Double
Dim valor As String
Dim matc() As String
Dim suma As Double
Dim rmesa As New ADODB.recordset

txtportfr = "Normal"
       txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          ReDim matv(1 To noesc, 1 To 1) As Double
          valt01 = rmesa.Fields("VALT0")
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          noreg1 = UBound(matc, 1)
          For l = 1 To noreg1
              matv(l, 1) = CDbl(matc(l))
          Next l
          rmesa.Close
          ReDim matval(1 To noesc1, 1 To 1) As Double
          For i = 1 To noesc1
             matval(i, 1) = matv(noreg1 - noesc1 + i, 1)
          Next i
          matval = ROrdenDbl(matval, 1)
          suma = 0
          noesc3 = 2
          If nconf < 0.5 Then
             For i = 1 To noesc3
                 suma = suma + Minimo(matval(i, 1), matval(1, 1) * 0.8)
             Next i
             suma = suma / noesc3
          Else
             For i = 1 To noesc3
                 suma = suma + Maximo(matval(noesc1 - noesc3 + i, 1), matval(noesc3, 1) * 0.8)
             Next i
             suma = suma / noesc3
          End If
          CalcularCVaRPrev = suma
          exito = True
       Else
          CalcularCVaRPrev = 0
          exito = False
       End If

End Function

Sub GenSubprocCalcPyGSubport(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_proc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim noport As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Long
Dim txtcadena As String
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"

txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,FECHAREG,COPERACION) NOT IN"
txtfiltro2 = txtfiltro2 & " (SELECT CPOSICION,FREGISTRO,COPERACION"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtescfr & "'"
txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg = 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
   If UBound(MatGruposPortPos, 1) <> 0 Then
      For j = 1 To UBound(MatGruposPortPos, 1)
          contar = contar + 1
          txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Calculo de pyg de subport", txtport, txtescfr, MatGruposPortPos(j, 3), noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
          ConAdo.Execute txtcadena
          DoEvents
      Next j
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "El portafolio no existe"
      exito = True
   End If
Else
    txtmsg = "No hay suficientes datos para realizar el proceso"
    exito = False
End If
End Sub

Sub GenSubprocCalcPyGMontSubport(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Long, ByVal id_proc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim noport As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Long
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
contar = DeterminaMaxRegSubproc(id_tabla)
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,FECHAREG,COPERACION) NOT IN"
txtfiltro2 = txtfiltro2 & " (SELECT CPOSICION,FREGISTRO,COPERACION"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPyGMontOper & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtescfr & "'"
txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg = 0 Then
   MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
   If UBound(MatGruposPortPos, 1) <> 0 Then
      For j = 1 To UBound(MatGruposPortPos, 1)
          contar = contar + 1
          txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Calculo de pyg Mont por subport", txtport, txtescfr, MatGruposPortPos(j, 3), noesc, htiempo, nosim, "", "", "", "", "", "", id_tabla)
          ConAdo.Execute txtcadena
          DoEvents
      Next j
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "El portafolio no existe"
      exito = True
   End If
Else
    txtmsg = "No hay suficientes datos para realizar el proceso"
    exito = False
End If
End Sub

Sub ConsolidaPyGSubport(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupo As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
On Error GoTo hayerror
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtborra As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim txttexto As String
Dim valor As String
Dim valt01 As Double
Dim matc() As String
Dim suma As Double
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha1
txtborra = txtborra & " AND F_FACTORES = " & txtfecha2
txtborra = txtborra & " AND F_VALUACION = " & txtfecha3
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtgrupo & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha3
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,FREGISTRO,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,FECHAREG, COPERACION FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha1 & " AND PORTAFOLIO = '" & txtgrupo & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matpl(1 To noesc) As Double
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma = 0
   For j = 1 To noesc
       matpl(j) = 0
   Next j
   For i = 1 To noreg
       valt01 = rmesa.Fields("VALT0")
       valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       suma = suma + valt01
       For j = 1 To noesc
           matpl(j) = matpl(j) + CDbl(matc(j))
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Generando las p y g del portafolio " & txtgrupo & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   rmesa.Close
   txtfiltro = "SELECT * FROM " & TablaPLEscHistPort
   RInterfIKOS.Open txtfiltro, ConAdo, 1, 3
   txtcadena = ""
   For j = 1 To UBound(matpl, 1) - 1
       txtcadena = txtcadena & matpl(j) & ","
   Next j
   txtcadena = txtcadena & matpl(UBound(matpl, 1))
   RInterfIKOS.AddNew
   RInterfIKOS.Fields("F_POSICION") = CLng(f_pos)                'la fecha de proceso
   RInterfIKOS.Fields("F_FACTORES") = CLng(f_factor)             'la fecha de LOS FACTORES
   RInterfIKOS.Fields("F_VALUACION") = CLng(f_val)               'la fecha de LOS VALUACION
   RInterfIKOS.Fields(3) = txtport                               'el portafolio
   RInterfIKOS.Fields(4) = txtgrupo                              'el subportafolio
   RInterfIKOS.Fields(5) = txtescfr                              'el escenario de factores de riesgo
   RInterfIKOS.Fields(6) = noesc                                 'no de escenarios
   RInterfIKOS.Fields(7) = htiempo                               'horizonte de tiempo
   RInterfIKOS.Fields(8) = suma                                  'valuacion de escenario base
   Call GuardarElementoClob(txtcadena, RInterfIKOS, "DATOS")
   RInterfIKOS.Update
   RInterfIKOS.Close
   DoEvents
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "no hay pyg para este subportafolio"
   exito = True
End If

On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Function LeerPyG1Oper(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim i As Long
Dim valt01 As Double
Dim valor As String
Dim matc() As String
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matpl(1 To noesc, 1 To 1) As Double
   rmesa.Open txtfiltro2, ConAdo
   valt01 = rmesa.Fields("VALT0")
   valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
   matc = EncontrarSubCadenas(valor, ",")
   For i = 1 To noesc
       matpl(i, 1) = CDbl(matc(i))
   Next i
   rmesa.Close
Else
   ReDim matpl(0 To 0, 0 To 0) As Double
End If
LeerPyG1Oper = matpl
End Function

Sub GenPyGPortVR(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupo As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef exito As Boolean)

'On Error GoTo hayerror
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim txttexto As String
Dim valor As String
Dim valt0 As Double
Dim valt1 As Double
Dim matc() As String
Dim suma0 As Double
Dim suma1 As Double
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPyGHistPVR & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtgrupo & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
ConAdo.Execute txtborra

txtfiltro2 = "SELECT * FROM " & TablaPLHistOperVR & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtgrupo & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0).value
rmesa.Close
If noreg <> 0 Then
   ReDim matpl(1 To noesc) As Double
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma0 = 0
   suma1 = 0
   For j = 1 To noesc
       matpl(j) = 0
   Next j
   For i = 1 To noreg
       valt0 = rmesa.Fields("VALT0")
       valt1 = rmesa.Fields("VALT1")
       valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       suma0 = suma0 + valt0
       suma1 = suma1 + valt1
       For j = 1 To noesc
           matpl(j) = matpl(j) + CDbl(matc(j))
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Generando las p y g del portafolio " & txtgrupo & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   rmesa.Close
   txtfiltro = "SELECT * FROM " & TablaPyGHistPVR
   RInterfIKOS.Open txtfiltro, ConAdo, 1, 3
   txtcadena = ""
   For j = 1 To UBound(matpl, 1) - 1
       txtcadena = txtcadena & matpl(j) & ","
   Next j
   txtcadena = txtcadena & matpl(UBound(matpl, 1))
   RInterfIKOS.AddNew
   RInterfIKOS.Fields("FECHA") = CLng(fecha)                       'la fecha de proceso
   RInterfIKOS.Fields("PORTAFOLIO") = txtport                      'el portafolio
   RInterfIKOS.Fields("ESC_FACTORES") = txtescfr                   'el portafolio
   RInterfIKOS.Fields("SUBPORT") = txtgrupo                        'el subportafolio
   RInterfIKOS.Fields("NOESC") = noesc                             'no de escenarios
   RInterfIKOS.Fields("HTIEMPO") = htiempo                         'horizonte de tiempo
   RInterfIKOS.Fields("P_FWD") = 30
   RInterfIKOS.Fields("VALT0") = suma0                        'valuacion de escenario base
   RInterfIKOS.Fields("VALT1") = suma1                      'valuacion de escenario en t+1
   Call GuardarElementoClob(txtcadena, RInterfIKOS, "DATOS")
   RInterfIKOS.Update
   RInterfIKOS.Close
   DoEvents
   exito = True
Else
   exito = False
End If

On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Sub GenProcCVaR(ByVal dtfecha As Date, ByVal id_proc As Integer, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_tabla As Integer)
    Dim noport As Integer
    Dim bl_exito As Boolean, bl_exito1 As Boolean
    Dim lambda As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim jj As Integer, indice As Integer
    Dim p As Integer
    Dim contar As Long
    Dim dtfechaf, dtfechav As Date
    Dim MatPrecios() As Double
    Dim txtnomarch As String, txtsalida As String, txtfecha1 As String, txthorap As String, txtcadena As String
    Dim txtfecha2 As String
    Dim noesce As Integer
    Dim matesce() As Variant
    Dim matfechas() As Date
    Dim txtfiltro As String
    Dim txttabla As String
    
    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "Cálculo de CVaR", dtfecha, dtfecha, dtfecha, txtport, txtportfr, noesc, htiempo, "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub
  
  Sub GenSubprocGPort(ByVal dtfecha As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
    Dim noport As Integer
    Dim bl_exito As Boolean, bl_exito1 As Boolean
    Dim lambda As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim jj As Integer, indice As Integer
    Dim p As Integer
    Dim contar As Long
    Dim dtfechaf, dtfechav As Date
    Dim MatPrecios() As Double
    Dim txtnomarch As String, txtsalida As String, txtfecha1 As String, txthorap As String, txtcadena As String
    Dim txtfecha2 As String
    Dim noesce As Integer
    Dim matesce() As Variant
    Dim matfechas() As Date
    Dim txtfiltro As String
    Dim txttabla As String
    
    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "Generacion de portafolios", "", "", "", "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub


Sub GenProcCPYG(ByVal dtfecha As Date, ByVal id_proc As Integer, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_tabla As Integer)
    Dim noport As Integer
    Dim bl_exito As Boolean, bl_exito1 As Boolean
    Dim lambda As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim jj As Integer, indice As Integer
    Dim p As Integer
    Dim contar As Long
    Dim dtfechaf, dtfechav As Date
    Dim MatPrecios() As Double
    Dim txtnomarch As String, txtsalida As String, txtfecha1 As String, txthorap As String, txtcadena As String
    Dim txtfecha2 As String
    Dim noesce As Integer
    Dim matesce() As Variant
    Dim matfechas() As Date
    Dim txtfiltro As String
    Dim txttabla As String
    
    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "Cálculo de CVaR", dtfecha, dtfecha, dtfecha, txtport, txtportfr, noesc, htiempo, "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub
  
Function ValidarMtMMercadoDinero(ByVal fecha As Date, ByVal txtport As String)

Dim noreg As Integer
Dim i As Integer
Dim suma As Double
Dim matv() As Double
noreg = 23
ReDim txtsubport(1 To noreg) As Variant

txtsubport(1) = "MDT DIRECTO BONDES D"
txtsubport(2) = "MDT DIRECTO BONOS M"
txtsubport(3) = "MDT DIRECTO BONOS USD"
txtsubport(4) = "MDT DIRECTO CBICS"
txtsubport(5) = "MDT DIRECTO CERTIFICADOS BURSATILES"
txtsubport(6) = "MDT DIRECTO CETES"
txtsubport(7) = "MDT DIRECTO IPAB IM"
txtsubport(8) = "MDT DIRECTO IPAB IQ"
txtsubport(9) = "MDT DIRECTO IPAB IS"
txtsubport(10) = "MDT DIRECTO UDIBONOS"
txtsubport(11) = "MDT DIRECTO PRLV"
txtsubport(12) = "MDT REPORTO BONDES D"
txtsubport(13) = "MDT REPORTO BONOS M"
txtsubport(14) = "MDT REPORTO BONOS USD"
txtsubport(15) = "MDT REPORTO CBICS"
txtsubport(16) = "MDT REPORTO CERTIFICADOS BURSATILES"
txtsubport(17) = "MDT REPORTO CETES"
txtsubport(18) = "MDT REPORTO IPAB IM"
txtsubport(19) = "MDT REPORTO IPAB IQ"
txtsubport(20) = "MDT REPORTO IPAB IS"
txtsubport(21) = "MDT REPORTO UDIBONOS"
txtsubport(22) = "MDT REPORTO PRLV"
txtsubport(23) = "MERCADO DE DINERO"
suma = 0
For i = 1 To noreg - 1
    matv = LeerResValPort(fecha, txtport, txtsubport(i), 1)
    If UBound(matv, 1) > 0 Then
       suma = suma + matv(1)
    End If
Next i
matv = LeerResValPort(fecha, txtport, txtsubport(noreg), 1)
If UBound(matv, 1) > 0 Then
ValidarMtMMercadoDinero = matv(1) - suma
Else
  ValidarMtMMercadoDinero = suma
End If
End Function
  

Function ValidarResVaR(ByVal fecha As Date, ByVal txtport As String, ByVal noesc As Long, ByVal htiempo As Integer)
'rutina para validar resultados de CVaR por subportafolio y el CVaR global
Dim txtsubport(1 To 6) As Variant
Dim matpl() As Double
Dim matplt() As Double
Dim matpls() As Double
Dim matdif() As Double
Dim i As Long
Dim j As Long
Dim suma As Double

txtsubport(1) = "CONSOLIDADO"
txtsubport(2) = "MERCADO DE DINERO"
txtsubport(3) = "MESA DE CAMBIOS"
txtsubport(4) = "DERIVADOS DE NEGOCIACION"
txtsubport(5) = "DERIVADOS ESTRUCTURALES"
txtsubport(6) = "DERIVADOS NEGOCIACION RECLASIFICACION"
matplt = LeerPyGHistSubport(fecha, fecha, fecha, txtport, txtsubport(1), "Normal", noesc, htiempo, 0)
ReDim matpls(1 To noesc) As Double
For i = 2 To 6
    matpl = LeerPyGHistSubport(fecha, fecha, fecha, txtport, txtsubport(i), "Normal", noesc, htiempo, 0)
    If UBound(matpl, 1) > 0 Then
       For j = 1 To noesc
           matpls(j) = matpls(j) + matpl(j, 1)
       Next j
    End If
Next i
suma = 0
For i = 1 To noesc
    suma = suma + Abs(matpls(i) - matplt(i, 1))
Next i

End Function

