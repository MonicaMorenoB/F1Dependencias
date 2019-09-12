Attribute VB_Name = "ModuloValuacion"
Option Explicit

Sub GenerarFlujosSwapsVFD2(ByVal fecha As Date, _
                           ByVal tipopos As Integer, _
                           ByVal fechareg As Date, _
                           ByVal txtnompos As String, _
                           ByVal horareg As String, _
                           ByVal cposicion As Integer, _
                           ByVal coperacion As String, _
                           ByVal id_val As Integer, _
                           ByRef txtmsg As String, _
                           ByRef exito As Boolean)
Dim bl_exito As Boolean
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposswaps() As New propPosSwaps
Dim matpr() As New resValIns
Dim i As Long
Dim indice As Long
Dim txtcadena As String
Dim txtport As String
Dim fecha1 As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim txtinserta As String

    SiAnexarFlujosSwaps = True
'primero se procede a leer la interfase de
    Call RutinaValOper(fecha, fecha, fecha, matpos, matposmd, matposswaps, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, matpr, id_val, txtmsg, exito)
    If UBound(matpos, 1) <> 0 Then
   'se procede a actualizar una tabla con estos valores y despues se realiza un filtro
       txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       For i = 2 To UBound(MatValFlujosD, 1)
           txtcadena1 = "INSERT INTO " & TablaFlujosRichard & " VALUES("
           txtcadena1 = txtcadena1 & txtfecha1 & ","                                      'fecha de la posicion
           txtcadena1 = txtcadena1 & "'" & MatValFlujosD(i).c_operacion & "',"            'Clave de operación
           txtcadena1 = txtcadena1 & "'" & MatValFlujosD(i).t_pata & "',"                 'PATA
           txtcadena1 = txtcadena1 & "'" & MatValFlujosD(i).intencion & "',"              'INTENCION
           txtfecha = "to_date('" & Format(MatValFlujosD(i).fecha_ini, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena1 = txtcadena1 & txtfecha & ","                                       'FECHA INICIO DEL FLUJO
           txtfecha = "to_date('" & Format(MatValFlujosD(i).fecha_fin, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena1 = txtcadena1 & txtfecha & ","                                       'FECHA FIN DEL FLUJO
           txtfecha = "to_date('" & Format(MatValFlujosD(i).fecha_desc, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena1 = txtcadena1 & txtfecha & ","                                       'FECHA DE PAGO DEL FLUJO
           txtcadena1 = txtcadena1 & MatValFlujosD(i).saldo_periodo & ","                 'SALDO
           txtcadena1 = txtcadena1 & MatValFlujosD(i).amortizacion & ","                  'AMORTIZACION
           txtcadena1 = txtcadena1 & "'" & Format(MatValFlujosD(i).tc_aplicar, "###0.0000000") & "',"   'TASA txtgrupo
           txtcadena1 = txtcadena1 & MatValFlujosD(i).sobretasa & ","                     'SPREAD
           txtcadena1 = txtcadena1 & MatValFlujosD(i).p_cupon & ","                       'P CUPON
           txtcadena1 = txtcadena1 & "'" & MatValFlujosD(i).int_ini & "',"                'INTER INICIAL
           txtcadena1 = txtcadena1 & "'" & MatValFlujosD(i).int_fin & "',"                'INTER FINAL
           txtcadena1 = txtcadena1 & ReemplazaVacioValor(Format(MatValFlujosD(i).pago_total), 0) & ","        'VALOR FLUJO SIN DESCONTAR
           txtcadena1 = txtcadena1 & ReemplazaVacioValor(MatValFlujosD(i).moneda, 0) & ")"                      'moneda
           ConAdo.Execute txtcadena1
           AvanceProc = i / UBound(MatValFlujosD, 1)
           MensajeProc = "Guardando los flujos de la posición de swaps " & Format(AvanceProc, "##0.00 %")
           DoEvents
       Next i
       For i = 1 To UBound(matpos, 1)
           indice = matpos(i).IndPosicion
           txtcadena1 = "UPDATE " & TablaFlujosRichard & " SET TASA = '" & matposswaps(indice).TCActivaSwap & "' WHERE FECHA_CORTE = " & txtfecha1 & " AND EMISION = '" & matpos(i).c_operacion & "' AND TPATA = 'B'"
           txtcadena2 = "UPDATE " & TablaFlujosRichard & " SET TASA = '" & matposswaps(indice).TCPasivaSwap & "' WHERE FECHA_CORTE = " & txtfecha1 & " AND EMISION = '" & matpos(i).c_operacion & "' AND TPATA = 'C'"
           ConAdo.Execute txtcadena1
           ConAdo.Execute txtcadena2
           AvanceProc = i / UBound(matpos, 1)
           MensajeProc = "Actualizando la tasa cupon de flujos " & Format(AvanceProc, "##0.00 %")
           DoEvents
       Next i
       exito = True
       txtmsg = "Proceso finalizado correctamente"
    Else
       exito = False
       txtmsg = "No ha posicion para valuar"
    End If
    SiAnexarFlujosSwaps = False
End Sub

Sub GenerarDatosFwds(ByVal fecha As Date, ByRef matpos() As propPosRiesgo, _
                     ByRef matposfwd() As propPosFwd, ByRef txtmsg As String, ByRef exito As Boolean)
Dim bl_exito As Boolean
Dim i As Long
Dim indice As Long
Dim txtcadena As String
Dim txtport As String
Dim fecha1 As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim txtinserta As String
Dim tipopos As Integer
Dim mata() As Variant
    SiAgregarDatosFwd = True
    ValExacta = True
    txtport = "FORWARDS"
    fecha1 = fecha - 10
'primero se procede a leer la interfase de
    Call CrearMatFRiesgo2(fecha1, fecha, MatFactRiesgo, "", exito)
    tipopos = 1
    Call CalculaValPos(fecha, fecha, fecha, txtport, 2, bl_exito)  'procesando la informacion de la fecha
'se procede a actualizar una tabla con estos valores y despues se realiza un filtro
    Open "d:\resultados.txt" For Output As #1
    txtcadena = "Clave de operación" & Chr(9) & "Tipo de operacion" & Chr(9) & "Fecha de inicio" & Chr(9)
    txtcadena = txtcadena & "Fecha de vencimiento" & Chr(9) & "Monto nocional" & Chr(9) & "Activa/pasiva" & Chr(9)
    txtcadena = txtcadena & "Strike" & Chr(9) & "Tasa local" & Chr(9) & "Tasa ext" & Chr(9) & "Tipo de cambio " & Chr(9) & "Val activa" & Chr(9) & "Val pasiva"
    Print #1, txtcadena
    For i = 1 To UBound(matpos, 1)
        txtcadena = matpos(i).c_operacion & Chr(9)
        txtcadena = txtcadena & matposfwd(i).ClaveProdFwd & Chr(9)
        txtcadena = txtcadena & matposfwd(i).FCompraFwd & Chr(9)
        txtcadena = txtcadena & matposfwd(i).FVencFwd & Chr(9)
        txtcadena = txtcadena & matposfwd(i).MontoNocFwd & Chr(9)
        txtcadena = txtcadena & matposfwd(i).Tipo_Mov & Chr(9)
        txtcadena = txtcadena & matposfwd(i).PAsignadoFwd & Chr(9)
        txtcadena = txtcadena & MatParamFwds(3, i + 1) & Chr(9)
        txtcadena = txtcadena & MatParamFwds(4, i + 1) & Chr(9)
        txtcadena = txtcadena & MatParamFwds(5, i + 1) & Chr(9)
        txtcadena = txtcadena & MatParamFwds(1, i + 1) * matposfwd(i).MontoNocFwd & Chr(9)
        txtcadena = txtcadena & MatParamFwds(2, i + 1) * matposfwd(i).MontoNocFwd & Chr(9)
        Print #1, txtcadena
    Next i
    Close #1
   SiAgregarDatosFwd = False
End Sub

Sub ProcValOper(ByVal f_pos As Date, _
                ByVal f_factor As Date, _
                ByVal f_val As Date, _
                ByVal txtport As String, _
                ByVal txtportfr As String, _
                ByVal tipopos As Integer, _
                ByVal fechareg As Date, _
                ByVal txtnompos As String, _
                ByVal horareg As String, _
                ByVal cposicion As Integer, _
                ByVal coperacion As String, _
                ByVal id_val As Integer, _
                ByRef txtmsg As String, _
                ByRef final As Boolean, _
                ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim contar1 As Integer
Dim contar2 As Integer
Dim contar3 As Long
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matpr() As New resValIns
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim txtmsg1 As String
Dim txtmsg2 As String
final = False
   Call RutinaValOper(f_pos, f_factor, f_val, matpos, matposmd, matposswaps, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, matpr, id_val, txtmsg1, exito1)
   If exito1 Then
      Call GuardarValOper(f_pos, f_factor, f_val, txtport, txtportfr, matpos, matposmd, matpr, tipopos, fechareg, txtnompos, cposicion, coperacion, id_val, exito2)
   End If
   exito = exito1 And exito2
   If exito Then
     txtmsg = "El proceso finalizo correctamente"
   Else
     txtmsg = txtmsg1 & txtmsg2
   End If
 final = True
End Sub


Sub ProcValPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim contar1 As Long
Dim contar2 As Long
Dim contar3 As Long
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim txtmsg1 As String
Dim alerta As String

   Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, 1, txtmsg1, exito1)
   Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, 1, exito2)
   Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, 2, txtmsg1, exito1)
   Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, 2, exito2)
  
   mata = LeerResValDeriv(fecha, txtportCalc1, 1, contar1)
   matb = LeerResValDeriv(fecha, txtportCalc1, 2, contar2)
   Call LeerValPosMD(fecha, matc, contar3, alerta)
   exito = exito1 And exito2
   If (contar1 = 0 Or contar2 = 0) And contar3 = 0 Then
     If exito Then
        txtmsg = "El proceso finalizo correctamente"
     Else
        txtmsg = "Algun proceso de valuacion no se realizo"
     End If
   Else
      txtmsg = "Hay diferencias en la valuación"
      exito = False
   End If
End Sub

Sub GenSubprocValContrap(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_subproc As Integer, ByVal opcion As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito3 As Boolean
Dim exito4 As Boolean
   Call GenValPortContrap(fecha, txtportfr, txtport, "CCS Contrap", id_subproc, opcion, exito4)
   Call GenValPortContrap(fecha, txtportfr, txtport, "Fwds Contrap", id_subproc, opcion, exito4)
   Call GenValPortContrap(fecha, txtportfr, txtport, "IRS Contrap", id_subproc, opcion, exito3)
   Call GenValPortContrap(fecha, txtportfr, txtport, "Deriv Contrap", id_subproc, opcion, exito4)
   txtmsg = "El proceso finalizo correctamente"
   exito = True
End Sub

Sub RutinaValOper(ByVal f_pos As Date, _
                  ByVal f_factor As Date, _
                  ByVal f_val As Date, _
                  ByRef matpos() As propPosRiesgo, _
                  ByRef matposmd() As propPosMD, _
                  ByRef matposswaps() As propPosSwaps, _
                  ByVal tipopos As Integer, _
                  ByVal fechareg As Date, _
                  ByVal txtnompos As String, _
                  ByVal horareg As String, _
                  ByVal cposicion As Integer, _
                  ByVal coperacion As String, _
                  ByRef matpr() As resValIns, _
                  ByVal tval As Integer, _
                  ByRef txtmsg As String, _
                  ByRef exito As Boolean)

Dim mattxt() As String
Dim matposdiv() As New propPosDiv
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim i As Long
Dim indice As Long
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim exito4 As Boolean
Dim exito5 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg4 As String

   exito = False
   mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
   Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
   If UBound(matpos, 1) <> 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, tval, txtmsg2, exito2)
      If exito2 Then
         Call RutinaCargaFR(f_factor, exito3)
         If f_factor <> FechaArchCurvas Or EsArrayVacio(MatCurvasT) Then
            FechaArchCurvas = f_factor
            MatCurvasT = LeerCurvaCompleta(f_factor, exito5)
         End If
         If f_val <> fechaMatTrans Then
            mTransicionN = CargarMatTrans(f_val, "N")
            mTransicionI = CargarMatTrans(f_val, "I")
            fechaMatTrans = f_val
         End If
         
         Call AnexarDatosVPrecios(f_val, matposmd)
         If f_val <> fechavalIKOS Or EsArrayVacio(matvalIKOS) Then
            matvalIKOS = LeerValDerivIKOS(f_val)
            fechavalIKOS = f_val
         End If
         For i = 1 To UBound(matpos, 1)
             If matpos(i).C_Posicion = ClavePosDeriv Then
                indice = BuscarValorArray(matpos(i).c_operacion, matvalIKOS, 2)
                If indice <> 0 Then
                   matpos(i).ValActivaIKOS = matvalIKOS(indice, 3)
                   matpos(i).ValPasivaIKOS = matvalIKOS(indice, 4)
                   matpos(i).MtmIKOS = matvalIKOS(indice, 5)
                Else
                   txtmsg = "No estoy encontrando la valuacion en las tablas de ikos " & matpos(i).c_operacion
                End If
             End If
         Next i
         ValExacta = True
         Set parval = DeterminaPerfilVal("VALUACION")
         matpr = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, MatResValFlujo, txtmsg4, exito4)
         exito = True
         txtmsg = "El proceso finalizo correctamente"
      Else
         exito = False
         txtmsg = txtmsg2
      End If
   Else
       exito = False
       txtmsg = "No hay registros en la posicion"
   End If
End Sub

Sub RutinaValPort(ByVal f_pos As Date, _
                  ByVal f_factor As Date, _
                  ByVal f_val As Date, _
                  ByVal txtport As String, _
                  ByRef matpos() As propPosRiesgo, _
                  ByRef matposmd() As propPosMD, _
                  ByVal tval As Integer, _
                  ByRef txtmsg As String, _
                  ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim mattxt() As String
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String
    
    ValExacta = True
    exito = False
    mattxt = CrearFiltroPosPort(f_pos, txtport)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
    If UBound(matpos, 1) <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, tval, txtmsg2, exito2)
       If exito2 Then
          Call RutinaCargaFR(f_factor, exito)
          Call AnexarDatosVPrecios(f_val, matposmd)
          Set parval = DeterminaPerfilVal("VALUACION")
          MatPrecios = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, MatResValFlujo, txtmsg3, exito3)
          exito = True
       Else
         exito = False
       End If
    Else
      exito = False
    End If
End Sub

Sub RutinaValPos(ByVal f_pos As Date, _
                 ByVal f_factor As Date, _
                 ByVal f_val As Date, _
                 ByRef matpos() As propPosRiesgo, _
                 ByVal txtnompos As String, _
                 ByVal tval As Integer, _
                 ByRef matpr() As resValIns, _
                 ByRef txtmsg As String, _
                 ByRef exito As Boolean)

Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim mattxt() As String
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String
  
    ValExacta = True
    exito = False
    mattxt = CrearFiltroPosSim(txtnompos)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
    If UBound(matpos, 1) <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, tval, txtmsg2, exito2)
       If exito2 Then
          Call RutinaCargaFR(f_factor, exito)
          Call AnexarDatosVPrecios(f_val, matposmd)
          Set parval = DeterminaPerfilVal("VALUACION")
          matpr = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
          exito = True
          txtmsg = "El proceso finalizo correctamente"
       Else
          exito = False
          txtmsg = txtmsg2
       End If
     Else
       exito = False
     End If
End Sub



Sub ProcValPos2(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal tval As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtmsg1 As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD

    Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, tval, txtmsg1, exito1)
    Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, tval, exito2)
    exito = exito1 And exito2
End Sub

Function YieldPIPAB(ByVal fecha As Date, ByRef matf() As estFlujosMD, ByVal precio As Double, ByVal tc0 As Double, ByVal tr As Double, ByVal pcupon As Integer)
Dim yield As Double
Dim noitera As Integer
Dim precio0 As Double
Dim precio1 As Double
Dim inc As Double
Dim deriv As Double


    yield = 0.05
    noitera = 0
    precio0 = PIPABYield(fecha, matf, tc0, tr, yield, pcupon)
    Do While Abs(precio - precio0) > 0.0000001 And noitera < 100000
       inc = 0.0000001
       precio0 = PIPABYield(fecha, matf, tc0, tr, yield, pcupon)
       precio1 = PIPABYield(fecha, matf, tc0, tr, yield + inc, pcupon)
       deriv = (precio1 - precio0) / inc
       yield = yield - (precio0 - precio) / deriv
       noitera = noitera + 1
    Loop
    If Abs(precio - precio0) > 0.0000001 And noitera >= 100000 Then
       YieldPIPAB = 0
    Else
       YieldPIPAB = yield
    End If
End Function

Sub LeerValPosPension(ByVal fecha As Date, ByVal cposicion As Integer, ByRef mata() As Variant, ByRef nodif As Long)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

nodif = 0
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT A.CPOSICION,A.COPERACION,A.P_SUCIO,A.VAL_PIP_S,A.NO_TITULOS_,C.TV,C.EMISION,C.SERIE from " & TablaValPos & " A"
txtfiltro2 = txtfiltro2 & " JOIN (SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
txtfiltro2 = txtfiltro2 & ") C ON A.COPERACION = C.COPERACION  WHERE A.FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND A.ID_VALUACION = 1"
txtfiltro2 = txtfiltro2 & " AND A.VAL_PIP_S <> 0"
txtfiltro2 = txtfiltro2 & " AND A.CPOSICION = " & cposicion
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   ReDim mata(1 To noreg, 1 To 9) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("CPOSICION")
       mata(i, 2) = rmesa.Fields("COPERACION")
       mata(i, 3) = rmesa.Fields("TV")
       mata(i, 4) = rmesa.Fields("EMISION")
       mata(i, 5) = rmesa.Fields("SERIE")
       mata(i, 6) = rmesa.Fields("NO_TITULOS_")
       mata(i, 7) = rmesa.Fields("P_SUCIO")
       mata(i, 8) = rmesa.Fields("VAL_PIP_S")
       mata(i, 9) = Abs(mata(i, 7) - mata(i, 8))
       If mata(i, 9) > 0.001 Then
       nodif = nodif + 1
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
   nodif = 0
End If

End Sub


Sub LeerValPosMD(ByVal fecha As Date, ByRef mata() As Variant, ByRef nodif As Long, ByRef alerta As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset
nodif = 0
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT C.CPOSICION,C.COPERACION,A.P_SUCIO,A.VAL_PIP_S,A.NO_TITULOS_,C.TV,C.EMISION,C.SERIE from " & TablaValPos & " A"
txtfiltro2 = txtfiltro2 & " JOIN (SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & ") C ON A.COPERACION = C.COPERACION  WHERE A.FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND A.ID_VALUACION = 1"
txtfiltro2 = txtfiltro2 & " AND A.VAL_PIP_S <> 0"
txtfiltro2 = txtfiltro2 & " AND (A.CPOSICION = " & ClavePosMD
txtfiltro2 = txtfiltro2 & " OR A.CPOSICION = " & ClavePosTeso
txtfiltro2 = txtfiltro2 & " OR A.CPOSICION = " & ClavePosPIDV
txtfiltro2 = txtfiltro2 & " OR A.CPOSICION = " & ClavePosPICV & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
alerta = ""
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   ReDim mata(1 To noreg, 1 To 9) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("CPOSICION")
       mata(i, 2) = rmesa.Fields("COPERACION")
       mata(i, 3) = rmesa.Fields("TV")
       mata(i, 4) = rmesa.Fields("EMISION")
       mata(i, 5) = rmesa.Fields("SERIE")
       mata(i, 6) = rmesa.Fields("NO_TITULOS_")
       mata(i, 7) = rmesa.Fields("P_SUCIO")
       mata(i, 8) = rmesa.Fields("VAL_PIP_S")
       mata(i, 9) = Abs(mata(i, 7) - mata(i, 8))
       If mata(i, 8) <> 0 And mata(i, 7) = 0 Then alerta = "Hay valuaciones nulas"
       If mata(i, 9) > 0.001 Then
       nodif = nodif + 1
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
   nodif = 0
End If

End Sub

Function ValOpcTasa(ByVal fecha As Date, _
                    ByRef matfl() As Variant, _
                    ByVal tc As Double, _
                    ByVal strike As Double, _
                    ByVal pc As Integer, _
                    ByRef vvol() As Double, _
                    ByRef curva1() As propCurva, _
                    ByRef curva2() As propCurva, _
                    ByVal tinterpol As Integer, _
                    ByVal estiloOpc As Integer, _
                    ByVal topccp As Integer)

    'function para valuar las opciones de tasa caplets o floorlets
    'no implica la suma de varios caplets
    Dim i    As Integer, nc As Integer
    Dim dxv1 As Integer
    Dim pbc0 As Double
    Dim valor As Double
    Dim tdesc As Double
    Dim tcupon As Double

    nc = 12
    ReDim matval(1 To nc, 1 To 1) As Variant
    ReDim dvcp(1 To nc) As Long
    valor = 0

    For i = 1 To nc
        dvcp(i) = matfl(i, 3) - fecha
        pc = matfl(i, 3) - matfl(i, 2)
        If i = 1 Then
           tdesc = CalculaTasa(curva1, dvcp(i), tinterpol)
           If topccp = "C" Then
                matval(i, 1) = matfl(i, 3) * Maximo(tc - strike, 0) * pc / 360 / (1 + tdesc * dvcp(i) / 360)
            ElseIf topccp = "P" Then
                matval(i, 1) = matfl(i, 3) * Maximo(strike - tc, 0) * pc / 360 / (1 + tdesc * dvcp(i) / 360)
            End If
        Else
            tcupon = TFutura(curva1, dvcp(i) - pc, pc, tinterpol)    'tasa futura
            tdesc = CalculaTasa(curva2, dvcp(i), tinterpol)         'tasa subyacente
            pbc0 = (pc / 360) / (1 + tdesc * dvcp(i) / 360)
            matval(i, 1) = pbc0 * POpcion(matfl(i, 3), tcupon, strike, dvcp(i) - pc, 0, 0, vvol(1, 1), estiloOpc, topccp)
        End If
        valor = valor + matval(i, 1)
    Next i

    ValOpcTasa = valor
End Function

Function POpcion(ByVal vn As Double, _
                 ByVal precio As Double, _
                 ByVal strike As Double, _
                 ByVal dxv As Integer, _
                 ByVal Tasa1 As Double, _
                 ByVal Tasa2 As Double, _
                 ByVal volopc As Double, _
                 ByVal topccp As Integer, _
                 ByVal estiloOpc As String) As Double
    'esta funcion da el POpcion de los distintos tipos de
    'opciones que se pueden dar
    'clasificacion de los instrumentos

    Dim s    As Double
    Dim x    As Double
    Dim r1   As Double
    Dim r2   As Double
    Dim B    As Double
    Dim v    As Double
    Dim FIni As Date

    s = precio
    x = strike
    r1 = Tasa1
    v = volopc

    Select Case topccp

        Case "C"       'call

            If dxv = 0 Then
                POpcion = vn * (s - x)
            Else
                r1 = (Logarit(1 + Tasa1 * dxv / 360)) * 360 / dxv
                r2 = (Logarit(1 + Tasa2 * dxv / 360)) * 360 / dxv
                B = r1 - r2

                Select Case estiloOpc

                    Case "E"
                        POpcion = vn * FBlackSM(s, x, dxv / 360, r1, B, v, topccp)

                    Case "A"
                        POpcion = vn * ValAmericanAprox(s, x, dxv / 360, r1, B, v, topccp)
                End Select

            End If

        Case "P"      'put

            If dxv = 0 Then
                POpcion = vn * (x - s)
            Else
                r1 = (Logarit(1 + Tasa1 * dxv / 360)) * 360 / dxv
                r2 = (Logarit(1 + Tasa2 * dxv / 360)) * 360 / dxv
                B = r1 - r2

                Select Case estiloOpc

                    Case "E"
                        POpcion = vn * FBlackSM(s, x, dxv / 360, r1, B, v, topccp)

                    Case "A"
                        POpcion = vn * ValAmericanAprox(s, x, dxv / 360, r1, B, v, topccp)
                End Select

            End If

    End Select

End Function


Function FBlackSM(ByVal st As Double, ByVal kk As Double, ByVal t As Double, ByVal rend As Double, ByVal dif As Double, ByVal sigma As Double, ByVal topccp As String)
Dim d1 As Double
Dim d2 As Double
    'black and scholes generalizada
    'da el valor de una opcion de venta put o de compra call estilo europea
    'con la formula de black and scholes
    'dif    diferencia entre la tasa libre de riesgo y los tasa de dividendos
    '       si los dividendos son nulos entonces dif=rend
    'rend   tasa libre de riesgo
    'sigma  volatilidad observada del subyacente
    d1 = (Logarit(st / kk) + (dif + sigma ^ 2 / 2) * t) / (sigma * t ^ 0.5)
    d2 = (Logarit(st / kk) + (dif - sigma ^ 2 / 2) * t) / (sigma * t ^ 0.5)

    If topccp = "C" Then
        FBlackSM = st * Exponen((dif - rend) * t) * DNormal(d1, 0, 1, 1) - kk * Exponen(-rend * t) * DNormal(d2, 0, 1, 1)
    ElseIf topccp = "P" Then
        FBlackSM = kk * Exponen(-rend * t) * (1 - DNormal(d2, 0, 1, 1)) - st * Exponen((dif - rend) * t) * (1 - DNormal(d1, 0, 1, 1))
    Else
        FBlackSM = 0
    End If

End Function

Public Function ValAmericanAprox(ByVal s As Double, ByVal x As Double, ByVal t As Double, ByVal r As Double, ByVal B As Double, ByVal v As Double, ByVal CallPutFlag As String) As Double

    If CallPutFlag = "C" Then
        ValAmericanAprox = ValAmericanCallAprox(s, x, t, r, B, v)
    ElseIf CallPutFlag = "P" Then
        ValAmericanAprox = ValAmericanPutAprox(s, x, t, r, B, v)
    End If

End Function

Private Function ValAmericanCallAprox(ByVal s As Double, ByVal x As Double, ByVal t As Double, ByVal r As Double, ByVal B As Double, ByVal v As Double) As Double
    'funcion para la valuacion de un call tipo americano
    'se encuentra Sk de tal manera que optimiza la solucion

    Dim Sk As Double, n As Double, valK As Double
    Dim d1 As Double, Q2 As Double, a2 As Double

    If B >= r Then
        ValAmericanCallAprox = FBlackSM(s, x, t, r, B, v, "C")
    Else
        Sk = Kc(CDbl(x), CDbl(t), CDbl(r), CDbl(B), CDbl(v))
        n = 2 * B / v ^ 2                                           '
        valK = 2 * r / (v ^ 2 * (1 - Exponen(-r * t)))
        d1 = (Logarit(Sk / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
        Q2 = (-(n - 1) + Sqr((n - 1) ^ 2 + 4 * valK)) / 2
        a2 = (Sk / Q2) * (1 - Exponen((B - r) * t) * DNormal(d1, 0, 1, 1))

        If s < Sk Then
            ValAmericanCallAprox = FBlackSM(s, x, t, r, B, v, "C") + a2 * (s / Sk) ^ Q2
        Else
            ValAmericanCallAprox = (s - x)
        End If
    End If

End Function

Private Function ValAmericanPutAprox(ByVal s As Double, ByVal x As Double, ByVal t As Double, ByVal r As Double, ByVal B As Double, ByVal v As Double) As Double

    'funcion para la valuación de un put tipo americano
    Dim Sk As Double, n As Double, valK As Double
    Dim d1 As Double, Q1 As Double, a1 As Double

    Sk = Kp(CDbl(x), CDbl(t), CDbl(r), CDbl(B), CDbl(v))
    n = 2 * B / v ^ 2
    valK = 2 * r / (v ^ 2 * (1 - Exponen(-r * t)))
    d1 = (Logarit(Sk / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
    Q1 = (-(n - 1) - Sqr((n - 1) ^ 2 + 4 * valK)) / 2
    a1 = -(Sk / Q1) * (1 - Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1))

    If s > Sk Then
        ValAmericanPutAprox = FBlackSM(s, x, t, r, B, v, "P") + a1 * (s / Sk) ^ Q1
    Else
        ValAmericanPutAprox = (x - s)
    End If

End Function

Private Function Kc(x As Double, _
                    t As Double, _
                    r As Double, _
                    B As Double, _
                    v As Double) As Double

    '// Newton Raphson algorithm to solve for the critical commodity price for a Call
    Dim n   As Double, m As Double

    Dim Su  As Double, Si As Double

    Dim h2  As Double, valK As Double

    Dim d1  As Double, Q2 As Double, q2u As Double

    Dim LHS As Double, RHS As Double

    Dim bi  As Double, e As Double
    
    '// Calculation of seed value, Si
    n = 2 * B / v ^ 2
    m = 2 * r / v ^ 2
    q2u = (-(n - 1) + Sqr((n - 1) ^ 2 + 4 * m)) / 2
    Su = x / (1 - 1 / q2u)
    h2 = -(B * t + 2 * v * Sqr(t)) * x / (Su - x)
    Si = x + (Su - x) * (1 - Exponen(h2))

    valK = 2 * r / (v ^ 2 * (1 - Exponen(-r * t)))
    d1 = (Logarit(Si / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
    Q2 = (-(n - 1) + Sqr((n - 1) ^ 2 + 4 * valK)) / 2
    LHS = Si - x
    RHS = FBlackSM(Si, x, t, r, B, v, "C") + (1 - Exponen((B - r) * t) * DNormal(d1, 0, 1, 1)) * Si / Q2
    bi = Exponen((B - r) * t) * DNormal(d1, 0, 1, 1) * (1 - 1 / Q2) + (1 - Exponen((B - r) * t) * DNormal(d1, 0, 1, 1) / (v * Sqr(t))) / Q2
    e = 0.000001

    '// Newton Raphson algorithm for finding critical price Si
    While Abs(LHS - RHS) / x > e

        Si = (x + RHS - bi * Si) / (1 - bi)
        d1 = (Logarit(Si / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
        LHS = Si - x
        RHS = FBlackSM(Si, x, t, r, B, v, "C") + (1 - Exponen((B - r) * t) * DNormal(d1, 0, 1, 1)) * Si / Q2
        bi = Exponen((B - r) * t) * DNormal(d1, 0, 1, 1) * (1 - 1 / Q2) + (1 - Exponen((B - r) * t) * DNormal(d1, 0, 1, 0) / (v * Sqr(t))) / Q2

    Wend

    Kc = Si
End Function

Private Function Kp(x As Double, _
                    t As Double, _
                    r As Double, _
                    B As Double, _
                    v As Double) As Double

    Dim n   As Double, m As Double

    Dim Su  As Double, Si As Double

    Dim h1  As Double, valK As Double

    Dim d1  As Double, q1u As Double, Q1 As Double

    Dim LHS As Double, RHS As Double

    Dim bi  As Double, e As Double
    
    '// Calculation of seed value, Si
    n = 2 * B / v ^ 2
    m = 2 * r / v ^ 2
    q1u = (-(n - 1) - Sqr((n - 1) ^ 2 + 4 * m)) / 2
    Su = x / (1 - 1 / q1u)
    h1 = (B * t - 2 * v * Sqr(t)) * x / (x - Su)
    Si = Su + (x - Su) * Exponen(h1)
    
    valK = 2 * r / (v ^ 2 * (1 - Exponen(-r * t)))
    d1 = (Logarit(Si / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
    Q1 = (-(n - 1) - Sqr((n - 1) ^ 2 + 4 * valK)) / 2
    LHS = x - Si
    RHS = FBlackSM(Si, x, t, r, B, v, "P") - (1 - Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1)) * Si / Q1
    bi = -Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1) * (1 - 1 / Q1) - (1 + Exponen((B - r) * t) * DNormal(-d1, 0, 1, 0) / (v * Sqr(t))) / Q1
    e = 0.000001

    '// Newton Raphson algorithm for finding critical price Si
    While Abs(LHS - RHS) / x > e

        Si = (x - RHS + bi * Si) / (1 + bi)
        d1 = (Logarit(Si / x) + (B + v ^ 2 / 2) * t) / (v * Sqr(t))
        LHS = x - Si
        RHS = FBlackSM(Si, x, t, r, B, v, "P") - (1 - Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1)) * Si / Q1
        bi = -Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1) * (1 - 1 / Q1) - (1 + Exponen((B - r) * t) * DNormal(-d1, 0, 1, 1) / (v * Sqr(t))) / Q1

    Wend

    Kp = Si
End Function

Sub GenSubProcValDeriv(ByVal fecha As Date, ByVal txtport As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim i As Long
Dim noreg As Long
Dim contar As Long
Dim txtborra As String
Dim txtcadena As String
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim cposicion As Long
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim coperacion As String
Dim tipopos As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & TablaFlujosRichard & " WHERE FECHA_CORTE = " & txtfecha
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
txtborra = txtborra & " AND PARAMETRO9 = '" & id_val & "'"
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
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Detalle valuación derivado", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, id_val, "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
       DoEvents
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   exito = False
End If
End Sub

