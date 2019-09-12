Attribute VB_Name = "ModVaRMarkowitz"
Option Explicit

Sub GenVaRMark(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim noport As Integer
Dim i As Integer
Dim j As Integer
Dim exitofr As Boolean
exito = True
Call VerifCargaFR2(fecha, noesc + htiempo, exitofr)
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For j = 1 To UBound(MatGruposPortPos, 1)
       Call CalVaRMark(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, nconf, exito3)
       exito = exito And exito3
   Next j
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No esta definido el portafolio"
   exito = False
End If
End Sub

Sub CalcSensibOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim noport As Integer
Dim indice As Integer
Dim fecha1 As Date
Dim i As Integer
Dim j As Integer
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As New ParamValPos
Dim deriv As Double
Dim incre As Double
Dim mfactr12() As Double
Dim matfr() As Double
Dim mrvalflujo() As New resValFlujo
Dim matpr1() As New resValIns
Dim matpr2() As New resValIns
Dim txtfecha As String
Dim txtfechar As String
Dim txtcadena As String
Dim txtmsg1 As String
Dim txtmsg2 As String

exito = False
final = False
incre = 0.0000001
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg1, exito1)
If UBound(matpos, 1) <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call RutinaCargaFR(fecha, exito3)
      Call AnexarDatosVPrecios(fecha, matposmd)
      Set parval = DeterminaPerfilVal("SENSIBILIDADES")
      For i = 1 To UBound(matpos, 1)   'toda la posicion
          For j = 1 To NoFactores
              deriv = 0
'se comprueba si el instrumento i en la posicion es sensible al factor jj
              If EsSensibleaFactor(MatCaracFRiesgo(j).nomFactor, matpos(i).C_Posicion, matpos(i).c_operacion, matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda) Then
                 mfactr12 = MatFactR1
'se agrega el incremento al factor jj
                 mfactr12(j, 1) = MatFactR1(j, 1) + incre
                 parval.indpos = i    'se pide que solo se calcule la sensibilidad del instrumento i
                 matpr1 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg1, exito1)
                 matpr2 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mfactr12, MatCurvasT, parval, mrvalflujo, txtmsg2, exito2)
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
                    txtcadena = txtcadena & MatFactR1(j, 1) & ","
                    txtcadena = txtcadena & deriv & ")"
                    ConAdo.Execute txtcadena
                 End If
              End If
          Next j
          AvanceProc = i / UBound(matpos, 1)
          MensajeProc = "Calc. sensibilidades de primer orden de " & matpos(i).C_Posicion & " " & matpos(i).c_operacion & "  " & Format(AvanceProc, "##0.00 %")
          DoEvents
      Next i
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = txtmsg2
      exito = False
   End If
Else
   exito = False
   txtmsg = "No hay registros de la posicion"
End If
final = True
End Sub

Sub GenSubprocCalcSensibSubPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim noport As Integer
Dim indice As Integer
Dim fecha1 As Date
Dim j As Integer
Dim contar As Long
Dim mattxt() As String
Dim txttabla As String

txttabla = DetermTablaSubproc(id_tabla)
exito = False
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & txttabla & " WHERE ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND FECHAP = " & txtfecha
ConAdo.Execute txtborra
contar = DeterminaMaxRegSubproc(id_tabla)
contar = contar + 1
txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calc sensib subportafolio", txtport, txtportfr, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
ConAdo.Execute txtcadena
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
For j = 1 To UBound(MatGruposPortPos, 1)
    contar = contar + 1
    txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calc sensib subportafolio", txtport, txtportfr, MatGruposPortPos(j, 3), "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
Next j
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub

Sub GenSubprocPyGMontOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Long, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim exito1 As Boolean
Dim indice As Integer
Dim fecha1 As Date
Dim noreg As Long
Dim j As Integer
Dim contar As Long
Dim mattxt() As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)
exito = False
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & txttabla & " WHERE ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND FECHAP = " & txtfecha
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & TablaPyGMontOper & " WHERE FECHA = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND ESC_FACTORES = '" & txtportfr & "'"
txtborra = txtborra & " AND NOESC = " & noesc
txtborra = txtborra & " AND NOSIM  = " & nosim
txtborra = txtborra & " AND HTIEMPO = " & htiempo
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For j = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calc pyg mont x oper", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, noesc, htiempo, nosim, "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next j
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No se generaron los subprocesos"
   exito = False
End If
End Sub


Sub GenVaRMark2(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByRef matport() As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim noport As Integer
Dim indice As Integer
Dim fecha1 As Date
Dim i As Integer
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
Dim exitofr As Boolean
exito = False

Call VerifCargaFR2(fecha, noesc + htiempo, exitofr)
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
If UBound(matpos, 1) <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call RutinaCargaFR(fecha, exito3)
      Call AnexarDatosVPrecios(fecha, matposmd)
      Call CalcSensibNuevo(fecha, txtport, txtportfr, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1)
      Call CalcSensibPort(fecha, txtport, txtportfr, txtport, txtmsg, exito2)
      Call CalVaRMark(fecha, txtport, txtportfr, txtport, noesc, htiempo, nconf, exito3)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   End If
Else
   exito = False
   txtmsg = "No hay registros de la posicion"
End If

End Sub


