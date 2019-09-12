Attribute VB_Name = "modProcesos"
Option Explicit

Sub GenSubProcEscEstresTaylor(ByVal fecha As Date, ByVal fecha1 As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtinserta As String
Dim mata() As Date
Dim contar As Long
Dim txtfiltro As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfecha4 As String
Dim txtborra As String
Dim txttabla As String
Dim i As Long

exito = False
txttabla = DetermTablaSubproc(id_tabla)
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO =  " & id_subproc
ConAdo.Execute txtborra
contar = DeterminaMaxRegSubproc(id_tabla)
mata = RealizarPartEsc(fecha, fecha1, 1)
For i = 1 To UBound(mata, 1)
    contar = contar + 1
    txtfecha1 = Format(fecha, "dd/mm/yyyy")
    txtfecha2 = Format(mata(i, 1), "dd/mm/yyyy")
    txtfecha3 = Format(mata(i, 2), "dd/mm/yyyy")
    txtfecha4 = "TO_DATE('" & Format(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtinserta = CrearCadInsSub(fecha, id_subproc, contar, "Esc estres de Taylor", txtfecha1, txtfecha2, txtfecha3, txtport, txtgrupoport, "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtinserta
Next i
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub

Sub CrearSubProcValExtremos(ByVal fecha As Date, ByVal finicio As Date, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtinserta As String
Dim contar As Long
Dim txtfiltro As String
Dim txtfecha As String
Dim txtborra As String
Dim i As Long

exito = False
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaValExtO & " WHERE FECHA = " & txtfecha
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO =  " & id_subproc
ConAdo.Execute txtborra
contar = DeterminaMaxRegSubproc(id_tabla)
Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To UBound(MatCaracFRiesgo, 1)
    contar = contar + 1
    txtinserta = CrearCadInsSub(fecha, id_subproc, contar, "Calc Rend Extremo", MatCaracFRiesgo(i).nomFactor, MatCaracFRiesgo(i).plazo, CStr(finicio), "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtinserta
Next i
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub

Sub GenSimEscCVaREstres(ByVal dtfecha As Date, ByVal id_proc As Integer, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal txtnomarch As String, ByVal id_tabla As Integer)
    Dim noport As Integer
    Dim bl_exito As Boolean, bl_exito1 As Boolean
    Dim lambda As Double
    Dim i As Integer, k As Integer, j As Integer
    Dim jj As Integer, indice As Integer
    Dim p As Integer
    Dim contar As Long
    Dim dtfechaf, dtfechav As Date
    Dim MatPrecios() As Double
    Dim txtsalida As String
    Dim txthorap As String, txtcadena As String
    Dim noesce As Integer
    Dim matesce() As Variant
    Dim matfechas() As Date
    Dim txtfiltro As String
    Dim nomarch As String
    
    contar = DeterminaMaxRegSubproc(id_tabla)
    matesce = LeerEscEstres(txtnomarch)
    matfechas = LeerFechasEsc(txtnomarch)
    NoFechas = UBound(matfechas, 1)
    noesce = UBound(matesce, 1)
    For k = 1 To noesce
        For p = 1 To NoFechas
            contar = contar + 1
            txtcadena = CrearCadInsSub(matfechas(p, 1), id_proc, contar, "Cálculo de CVaR con estrés", dtfecha, matfechas(p, 1), dtfecha, txtportCalc2, matesce(k, 1), noesc, htiempo, "", "", "", "", "", id_tabla)
            ConAdo.Execute txtcadena
        Next p
    Next k

End Sub

Sub EjecucionSubprocesos(ByVal id_tabla As Integer)
If ActivarControlErrores Then
   On Error GoTo hayerror
End If
    Dim bl_exito  As Boolean
    Dim bl_exito1 As Boolean
    Dim bl_exito2 As Boolean
    Dim sibloqueo As Boolean
    Dim matop()   As Variant
    Dim mattxt()  As String
    Dim txtmsg As String
    Dim fecha As Date
    Dim final As Boolean
    Do While True
        matop = ObSubProcPend(bl_exito, id_tabla)
        If bl_exito Then
           If Not EsVariableVacia(matop(2)) Then
           Call BloquearSubProc(matop(2), id_tabla, sibloqueo)
           If sibloqueo Then
              final = True
              fecha = CDate(matop(16))
              If matop(1) = 52 Then
                  Call CalculoCVaRID(CDate(matop(4)), CDate(matop(4)), CDate(matop(4)), CStr(matop(5)), CStr(matop(6)), Val(matop(7)), Val(matop(8)), Val(matop(9)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 53 Then
                  Call CalculoEfRetroSwapPas(CDate(matop(4)), CDate(matop(5)), CStr(matop(6)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 54 Then
                  Call CalculoEfRetroSwapAct(CDate(matop(4)), CDate(matop(5)), matop(6), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 55 Then
                  Call CalculoEfRetroSwapActPas(CDate(matop(4)), CDate(matop(5)), matop(6), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 56 Then
                  Call EficienciaRetroFwd(CDate(matop(4)), CDate(matop(5)), matop(6), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 57 Then
                  Call CalculoEfRetroProxySwap(CDate(matop(4)), CDate(matop(5)), matop(6), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 58 Then
                  Call CEficProsSwapPasiva(CDate(matop(4)), CStr(matop(5)), txtmsg, bl_exito2)
              ElseIf matop(1) = 59 Then
                  Call CEficProsSwapActiva(CDate(matop(4)), CStr(matop(5)), txtmsg, bl_exito2)
              ElseIf matop(1) = 60 Then
                  Call CEficProsSwapActivaPasiva(CDate(matop(4)), CStr(matop(5)), txtmsg, bl_exito2)
              ElseIf matop(1) = 61 Then
                  Call CalculaEficProsFWD(CDate(matop(4)), CStr(matop(5)), txtmsg, bl_exito2)
              ElseIf matop(1) = 62 Then
                  Call CEficProsProxySwap(CDate(matop(4)), CStr(matop(5)), txtmsg, bl_exito2)
              ElseIf matop(1) = 63 Then
                 Call CrearEscExtFR(fecha, matop(4), Val(matop(5)), CDate(matop(6)), txtmsg, bl_exito2)
              ElseIf matop(1) = 64 Then
                 Call DeterminaPort(fecha, fecha, Val(matop(4)), txtmsg, bl_exito2)
              ElseIf matop(1) = 65 Then
                 Call GenPortDerivContrap(fecha, Val(matop(4)), matop(5), Val(matop(6)), txtmsg, bl_exito2)
              ElseIf matop(1) = 66 Then
                 Call DeterPosSwapsxContrap(fecha, matop(4), "Swap Contrap " & matop(4), Val(matop(5)))
                 Call DeterPosSwapsContrapTOper(fecha, matop(4), "IRS", "IRS Contrap " & matop(4), Val(matop(5)))
                 Call DeterPosSwapsContrapTOper(fecha, matop(4), "CCS", "CCS Contrap " & matop(4), Val(matop(5)))
                 txtmsg = "El proceso finalizo correctamente"
                 bl_exito2 = True
              ElseIf matop(1) = 67 Then
                 Call ProcValOper(fecha, fecha, fecha, CStr(matop(4)), CStr(matop(5)), Val(matop(6)), CDate(matop(7)), CStr(matop(8)), CStr(matop(9)), Val(matop(10)), CStr(matop(11)), Val(matop(12)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 68 Then
                 Call CalcValSubportPos(fecha, fecha, fecha, matop(4), matop(5), CStr(matop(6)), Val(matop(7)), bl_exito2)
              ElseIf matop(1) = 69 Then
                 Call CalcPyG1Oper(fecha, fecha, fecha, CStr(matop(4)), CStr(matop(5)), Val(matop(6)), CDate(matop(7)), CStr(matop(8)), CStr(matop(9)), Val(matop(10)), CStr(matop(11)), Val(matop(12)), Val(matop(13)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 70 Then
                  Call ConsolidaPyGSubport(fecha, fecha, fecha, CStr(matop(4)), matop(5), matop(6), Val(matop(7)), Val(matop(8)), txtmsg, bl_exito2)
              ElseIf matop(1) = 71 Then
                 Call CalcEscEstresOper(fecha, CStr(matop(4)), CStr(matop(5)), Val(matop(6)), CDate(matop(7)), CStr(matop(8)), CStr(matop(9)), Val(matop(10)), CStr(matop(11)), txtmsg, bl_exito2)
              ElseIf matop(1) = 72 Then
                 If EsArrayVacio(matResEscEstres) Or FechaEscEstres <> fecha Then
                    matResEscEstres = LeerEscEstresS(fecha)
                    FechaEscEstres = fecha
                 End If
                 Call CalcEscEstresSubPort(fecha, matop(4), matop(5), matop(6), matResEscEstres, txtmsg, final, bl_exito2)
              ElseIf matop(1) = 73 Then
                 Call CalcSensibOper(fecha, CStr(matop(4)), CStr(matop(5)), Val(matop(6)), CDate(matop(7)), CStr(matop(8)), CStr(matop(9)), Val(matop(10)), CStr(matop(11)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 74 Then
                 Call CalcSensibPort(fecha, CStr(matop(4)), matop(5), matop(6), txtmsg, bl_exito2)
              ElseIf matop(1) = 75 Then
                Call CalcEstresTaylor(CDate(matop(4)), CDate(matop(5)), CDate(matop(6)), matop(7), CStr(matop(8)), 1, txtmsg, bl_exito2)
              ElseIf matop(1) = 76 Then
                  Call CalcPyGMontOper(fecha, CStr(matop(4)), CStr(matop(5)), Val(matop(6)), CDate(matop(7)), CStr(matop(8)), CStr(matop(9)), Val(matop(10)), CStr(matop(11)), Val(matop(12)), Val(matop(13)), Val(matop(14)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 77 Then
                  Call CalcPyGMontSubport(fecha, CStr(matop(4)), CStr(matop(5)), CStr(matop(6)), Val(matop(7)), Val(matop(8)), Val(matop(9)), txtmsg, bl_exito2)
              ElseIf matop(1) = 78 Then
                 Call CalcExpMaxFwd(fecha, Val(matop(4)), CDate(matop(5)), CStr(matop(6)), CStr(matop(7)), Val(matop(8)), CStr(matop(9)), 0.995, txtmsg, bl_exito2)
              ElseIf matop(1) = 79 Then
                 Call CalcLimContrapSwap1(fecha, Val(matop(4)), CDate(matop(5)), CStr(matop(6)), CStr(matop(7)), Val(matop(8)), CStr(matop(9)), Val(matop(10)), CDate(matop(11)), CDate(matop(12)), Val(matop(13)), Int(matop(14)), CDate(matop(15)), txtmsg, final, bl_exito2)
              ElseIf matop(1) = 80 Then
                  Call DetCurvaValMaxOper(fecha, CStr(matop(9)), Val(matop(10)), 79, id_tabla, txtmsg, final, bl_exito2)
              ElseIf matop(1) = 81 Then
                  Call CalcLimContrapSwap2(fecha, Val(matop(4)), CDate(matop(5)), CStr(matop(6)), CStr(matop(7)), Val(matop(8)), CStr(matop(9)), Val(matop(10)), CDate(matop(11)), CDate(matop(12)), Val(matop(13)), Int(matop(14)), CDate(matop(15)), 80, id_tabla, txtmsg, final, bl_exito2)
              ElseIf matop(1) = 82 Then
                  Call DeterminaEscValMax(fecha, CStr(matop(9)), Val(matop(10)), 81, id_tabla, txtmsg, final, bl_exito2)
              ElseIf matop(1) = 83 Then
                   Call CalcCVAOper(fecha, Val(matop(4)), CDate(matop(5)), CStr(matop(6)), CStr(matop(7)), Val(matop(8)), CStr(matop(9)), Val(matop(10)), Val(matop(11)), txtmsg, bl_exito2)
              ElseIf matop(1) = 84 Then
                  Call GenResCVA(fecha, Val(matop(4)), Val(matop(5)), Val(matop(6)), CStr(matop(7)), CStr(matop(8)), 83, id_tabla, txtmsg, bl_exito2)
              ElseIf matop(1) = 85 Then
                  Call GenResCVAPos(fecha, Val(matop(4)), Val(matop(5)), Val(matop(6)), matop(7), txtmsg, bl_exito2)
              ElseIf matop(1) = 86 Then
                   Call ProcValPos2(fecha, matop(4), matop(5), Val(matop(6)), txtmsg, bl_exito2)
              ElseIf matop(1) = 87 Then
                  Call CalculoCVaREstres(CDate(matop(4)), CDate(matop(5)), CDate(matop(6)), matop(7), matop(8), matop(9), Val(matop(10)), Val(matop(11)), 0.97, txtmsg, bl_exito2)
              ElseIf matop(1) = 88 Then
                 Call SubprocCalculoPyGPort(fecha, fecha, fecha, matop(8), matop(7), Val(matop(9)), Val(matop(10)), txtmsg, bl_exito2)
              ElseIf matop(1) = 89 Then
                  Call CalcularBacktesting(fecha, fecha, matop(7), txtmsg, bl_exito2)
              ElseIf matop(1) = 90 Then
                  Call CalcEscMakeW(fecha, Val(matop(4)), CDate(matop(5)), CStr(matop(6)), CStr(matop(7)), Val(matop(8)), CStr(matop(9)), Val(matop(10)), CDate(matop(11)), CDate(matop(12)), Val(matop(13)), txtmsg, bl_exito2)
              ElseIf matop(1) = 91 Then
                  Call CalcPyG1OperVR(fecha, fecha, fecha, matop(4), matop(5), Val(matop(6)), CDate(matop(7)), matop(8), matop(9), Val(matop(10)), Val(matop(11)), Val(matop(12)), Val(matop(13)), Val(matop(14)), txtmsg, bl_exito2)
              ElseIf matop(1) = 92 Then
                   Call CalcCVAMD(fecha, matop(4), Val(matop(5)), Val(matop(6)), Val(matop(7)), Val(matop(8)), matop(9), txtmsg, bl_exito2)
              ElseIf matop(1) = 93 Then
                   Call ProcCalculoWRW(fecha, matop(4), txtmsg, bl_exito2)
              ElseIf matop(1) = 94 Then
                   Call GenerarFlujosSwapsVFD2(fecha, Val(matop(4)), CDate(matop(5)), matop(6), matop(7), Val(matop(8)), matop(9), Val(matop(10)), txtmsg, bl_exito2)
              End If
              Call DesbloquearSubProc(matop(2), id_tabla, txtmsg, final, bl_exito2)
              If final And bl_exito2 Then
                 MensajeProc = "Se termino el proceso " & matop(2) & " del " & fecha
              ElseIf final And Not bl_exito2 Then
                 MensajeProc = "El proceso " & matop(2) & " del " & fecha & " tiene errores"
                 Call GuardaDatosBitacora(4, "Subproceso", matop(1), matop(2), NomUsuario, fecha, MensajeProc, id_tabla)
              ElseIf Not final And Not bl_exito2 Then
                 MensajeProc = "No se termino el proceso " & matop(2) & " del " & fecha & " tiene errores"
                 Call GuardaDatosBitacora(4, "Subproceso", matop(1), matop(2), NomUsuario, fecha, MensajeProc, id_tabla)
                 Call ActUHoraUsuario
                 SiActTProc = False
                 Exit Sub
              End If
           End If
           End If
        Else
           Exit Do
        End If
    Loop
    Call ActUHoraUsuario
    On Error GoTo 0
    Exit Sub
hayerror:
    MsgBox "EjecucionSubprocesos " & error(Err())
End Sub

Sub CalculoCVaRID(ByVal dtfecha1 As Date, _
                      ByVal dtfecha2 As Date, _
                      ByVal dtfecha3 As Date, _
                      ByVal txtport As String, _
                      ByVal txtescfr As String, _
                      ByVal noesc As Integer, _
                      ByVal htiempo As Integer, _
                      ByVal nconf As Double, _
                      ByVal txtmsg As String, _
                      ByRef final As Boolean, _
                      ByRef bl_exito As Boolean)

    Dim txtinserta As String, txtfecha As String
    Dim i As Integer, indice As Integer
    Dim j As Integer
    Dim noport As Integer
    Dim mattxt() As String
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim bl_exito1 As Boolean
    Dim fechax As Date
    Dim mata() As Variant
    Dim fecha1 As Date
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim noreg As Integer
    Dim finicio As Date
    Dim hinicio As Date
    Dim contar As Long
    Dim exito As Boolean
    Dim valor As Double
    Dim txtborra As String
    Dim txtmsg0 As String
    Dim txtmsg2 As String
    Dim exito2 As Boolean
    Dim exitofr As Boolean
    
    finicio = Date
    hinicio = Time
    noport = 4
    ReDim matport(1 To noport) As String
    matport(1) = "CONSOLIDADO ID"
    matport(2) = "DERIVADOS DE NEGOCIACION ID"
    matport(3) = "DERIVADOS ESTRUCTURALES ID"
    matport(4) = "DERIVADOS NEGOCIACION RECLASIFICACION ID"
    fechax = PBD1(dtfecha2, 1, "MX")
    Call VerifCargaFR2(fechax, noesc + htiempo, exitofr)
    mattxt = CrearFiltroPosPort(dtfecha1, txtport)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, bl_exito1)
    If UBound(matpos, 1) <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       If exito2 Then
          FechaPos = dtfecha1
          Call RutinaCargaFR(fechax, bl_exito)
          Call AnexarDatosVPrecios(dtfecha3, matposmd)
          indice = BuscarValorArray(fechax, MatFactRiesgo, 1)
          Call RutinaVaRHistórico1(fechax, dtfecha3, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
          Call GuardarEscHist(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito)
          txtfecha = "TO_DATE('" & Format(dtfecha1, "dd/mm/yyyy") & "','DD/MM/YYYY')"
          For i = 1 To UBound(matport, 1)
              Call ConsolidaPyGSubport(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, matport(i), noesc, htiempo, txtmsg, bl_exito)
              txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & matport(i) & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='CVARH' AND NCONF = " & 1 - nconf
              ConAdo.Execute txtborra
              valor = CalcularCVaRPyG(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, matport(i), noesc, htiempo, 1 - nconf, exito)
              If exito Then Call InsertaRegVaR(dtfecha1, dtfecha1, dtfecha1, txtport, matport(i), txtescfr, "CVARH", noesc, htiempo, 0, 1 - nconf, 0, valor)
          Next i
          ReDim mata(1 To 2, 1 To 1) As Variant
          contar = 0
          For i = 1 To UBound(matpos, 1)
             If matpos(i).tipopos = 3 Then
                contar = contar + 1
                ReDim Preserve mata(1 To 2, 1 To contar) As Variant
                mata(1, contar) = matpos(i).c_operacion
                mata(2, contar) = matpos(i).HoraRegOp
             End If
          Next i
          If contar <> 0 Then
             mata = MTranV(mata)
             Call GuardaResCVaRIKOS3(dtfecha1, mata, matport, conAdoBD, "", bl_exito)
             Call ValidarOperaciones2(mata, finicio, hinicio, Date, Time)
             bl_exito = True
             final = True
          Else
             MsgBox "No pasaron las operaciones a la posicion total para el calculo intradia"
             final = True
             exito = False
          End If
       Else
           final = True
           exito = False
           txtmsg = txtmsg2
       End If
    Else
      txtmsg = "No hay registros en la posicion"
      final = True
      exito = True
    End If
End Sub


Sub CalcEstresTaylor(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByVal htiempo As Long, ByRef txtmsg As String, ByRef bl_exito As Boolean)
Dim mattxt() As String
Dim fechax As Date
Dim exito1 As Boolean
Dim exito As Boolean
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
    Call VerifCargaFR(fecha1, fecha2)
    Call PruebaEstresT2(fecha, fecha1, fecha2, htiempo, txtport, "Normal", txtgrupoport, exito)
    bl_exito = exito
End Sub

Sub PruebaEstresT2(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal htiempo As Long, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByRef exito As Boolean)
Dim i As Long
Dim p As Long
Dim j As Long
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim indice As Long
Dim MatFactR1() As Double
Dim mats() As Variant
Dim tvar As Integer
Dim ndiassim As Integer
Dim nocols As Integer
Dim matrends() As Double
Dim MatDefC() As Variant
Dim sumval As Double
Dim matesc() As Variant
Dim matfechassh() As Date
Dim txtcadena As String
Dim matx() As Double
Dim delta() As Variant
Dim matb() As Integer
Dim l As Integer

'rutina calculo escenarios estres taylor
 SiValFR = False
 tvar = 0
 
'se filtran los factores de riesgo a los que es sensible el portafolio

ndiassim = UBound(MatFactRiesgo, 1)
nocols = UBound(MatFactRiesgo, 2) 'no de factores
'se extrae la matriz con las fechas de los escenarios
If UBound(MatFactRiesgo, 1) > htiempo Then
   MatFactR1 = CargaFR1Dia(fecha, exito)
'se calculan las sensibilidades de la posicion
   MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
   If UBound(MatGruposPortPos, 1) <> 0 Then
      For i = 1 To UBound(MatGruposPortPos, 1)
          delta = LeerSensibPort(fecha, txtport, txtportfr, MatGruposPortPos(i, 3))
          noreg = UBound(delta, 1)
          If noreg <> 0 Then
    'se extraen las sensibilidades de todo el portafolio en un vector
             ReDim matressim(1 To UBound(MatFactRiesgo, 1) - htiempo, 1 To 2) As Variant
    'se extrae la matriz con los escenarios a calcular
             matfechassh = ConvArVtDT(ExtraeSubMatrizV(MatFactRiesgo, 1, 1, 1 + htiempo, UBound(MatFactRiesgo, 1)))
             matx = ConvArVtDbl(ExtraeSubMatrizV(MatFactRiesgo, 2, UBound(MatFactRiesgo, 2), 1, UBound(MatFactRiesgo, 1)))
             Call GenRends3(matx, htiempo, matfechassh, matrends, matb)
             matesc = FiltrarFactSens(matrends, delta, MatDefC)
    'se multiplica la sensibilidad por el escenario
             
             For j = 1 To UBound(matrends, 1)
                 matressim(j, 1) = matfechassh(j, 1)             'fecha del escenario
                 sumval = 0
                 For p = 1 To UBound(matrends, 2)
                     For l = 1 To UBound(delta, 1)
                     If delta(l, 1) = MatCaracFRiesgo(p).indFactor Then
                        If matb(j, p) = 1 Then
                           sumval = sumval + delta(l, 2) * delta(l, 3) * matrends(j, p)
                        Else
                           sumval = sumval + delta(l, 3) * matrends(j, p)
                        End If
                     End If
                     Next l
                 Next p
                 matressim(j, 2) = sumval
                 MensajeProc = "Generando los escenarios de estres por aproximación lineal"
             Next j
             
             'For j = 1 To UBound(matrends, 1)
             '    matressim(j, 1) = matfechassh(j, 1)             'fecha del escenario
             '    sumval = 0
             '    For p = 1 To UBound(delta, 1)
             '        'If matb(j, p) = 1 Then
             '           sumval = sumval + delta(p, 2) * delta(p, 3) * matesc(j, p)
             '        'Else
             '        '   sumval = sumval + delta(p, 3) * matesc(j, p)
             '        'End If
             '    Next p
             '    matressim(j, 2) = sumval                           'el valor del escenario
             '    MensajeProc = "Generando los escenarios de estres por aproximación lineal"
             'Next j
             Call GuardarEscEstresAprox(fecha, fecha1, fecha2, txtport, txtportfr, MatGruposPortPos(i, 3), matressim)
          End If
      Next i
      exito = True
   Else
   
   
   End If
Else
   exito = False
End If
End Sub

Sub CalculoCVaREstres(ByVal dtfecha1 As Date, _
                      ByVal dtfecha2 As Date, _
                      ByVal dtfecha3 As Date, _
                      ByVal txtport As String, _
                      ByVal txtescfr As String, _
                      ByVal txtgrupoport As String, _
                      ByVal noesc As Integer, _
                      ByVal htiempo As Integer, _
                      ByVal nconf As Double, _
                      ByVal txtmsg As String, _
                      ByRef bl_exito As Boolean)

    Dim txtinserta As String, txtfecha As String
    Dim i As Integer, indice As Integer
    Dim j As Integer
    Dim indice1 As Integer, indice2 As Integer
    Dim noport As Integer
    Dim mattxt() As String
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim mata() As Variant
    Dim mfriesgo() As Variant
    Dim exito2 As Boolean
    Dim txtmsg0 As String
    Dim txtmsg2 As String
    Dim exitofr As Boolean
    
    'Dim parval As paramvalpos
    'Dim mrvalflujo() As resValFlujo
    Call VerifCargaFR2(dtfecha2, noesc + htiempo, exitofr)
    Dim bl_exito1 As Boolean
    If FechaPos <> dtfecha1 Then
       mattxt = CrearFiltroPosPort(dtfecha1, txtport)
       Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, bl_exito1)
       FechaPos = dtfecha1
    End If
    If UBound(matpos, 1) <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       If exito2 Then
          Call RutinaCargaFR(dtfecha2, bl_exito)
          mfriesgo = MatFactRiesgo
          If txtescfr <> "Normal" Then
             mata = LeerEscFR(txtescfr, dtfecha2)
             If UBound(mata, 1) <> 0 Then
                FechaMatFactR1 = dtfecha2
                indice1 = BuscarValorArray(dtfecha2, mfriesgo, 1)
                For i = 1 To NoFactores
                    indice2 = BuscarValorArray(MatCaracFRiesgo(i).indFactor, mata, 1)
                    If indice2 <> 0 Then
                       MatFactR1(i, 1) = mata(indice2, 2)
                       mfriesgo(indice1, i + 1) = mata(indice2, 2)
                    End If
                Next i
                Call AnexarDatosVPrecios(dtfecha3, matposmd)
                Call RutinaVaRHistórico1(dtfecha2, dtfecha3, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
                Call GuardarEscHist(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, bl_exito)
                MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
                For j = 1 To UBound(MatGruposPortPos, 1)
                    Call ConsolidaPyGSubport(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, MatGruposPortPos(j, 3), noesc, htiempo, txtmsg, bl_exito)
                Next j
                Call GeneraResCVaRPos(dtfecha1, dtfecha1, dtfecha1, txtport, txtescfr, txtgrupoport, noesc, htiempo, nconf, txtmsg, bl_exito)
                txtmsg = "El proceso finalizo correctamente"
                bl_exito = True
             Else
                MsgBox "No se encontro la simulacion de factores"
             End If
          End If
       End If
    End If
End Sub



Sub CalcCVaRIDOper(ByVal fecha As Date, _
                      ByVal txtport As String, _
                      ByVal txtescfr As String, _
                      ByVal tipopos As Integer, _
                      ByVal fechareg As Date, _
                      ByVal txtnompos As String, _
                      ByVal horareg As String, _
                      ByVal cposicion As Integer, _
                      ByVal coperacion As String, _
                      ByVal noesc As Integer, _
                      ByVal htiempo As Integer, _
                      ByVal nconf As Double, _
                      ByVal txtmsg As String, _
                      ByRef bl_exito As Boolean)

    Dim txtinserta As String, txtfecha As String
    Dim i As Integer, indice As Integer
    Dim j As Integer
    Dim noport As Integer
    Dim mattxt() As String
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim bl_exito1 As Boolean
    Dim f_factor As Date
    Dim mata() As Variant
    Dim fecha1 As Date
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim noreg As Integer
    Dim finicio As Date
    Dim hinicio As Date
    Dim contar As Long
    Dim exito As Boolean
    Dim valor As Double
    Dim txtborra As String
    Dim txtmsg0 As String
    Dim txtmsg2 As String
    Dim exito2 As Boolean
    Dim exitofr As Boolean
    
    finicio = Date
    hinicio = Time
    
    f_factor = PBD1(fecha, 1, "MX")
    Call VerifCargaFR2(f_factor, noesc + htiempo, exitofr)
    mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, bl_exito1)
    If UBound(matpos, 1) <> 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       If exito2 Then
          FechaPos = fecha
          Call RutinaCargaFR(f_factor, bl_exito)
          Call AnexarDatosVPrecios(fecha, matposmd)
          indice = BuscarValorArray(f_factor, MatFactRiesgo, 1)
          Call RutinaVaRHistórico1(f_factor, fecha, noesc, htiempo, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatVal0T, MatPyGT)
          Call GuardarEscHist(fecha, fecha, f_factor, txtport, txtescfr, noesc, htiempo, matpos, MatVal0T, MatPyGT, exito)
       Else
           exito = False
           txtmsg = "La operacion " & coperacion & " no esta en los catalogos de valuacion"
       End If
    Else
      MsgBox "No hay registros en la posicion"
      exito = True
    End If
End Sub


Function GenCuadroEscEstres2(ByVal fecha As Date) As Variant()
Dim noport As Integer
Dim noesc As Integer
Dim matportp() As String
Dim matesc() As String
Dim i As Integer
Dim j As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim fecha3 As Date

noport = 4
noesc = 11


ReDim matesc(1 To noesc) As String
matesc(1) = "3 desv est"
matesc(2) = "Ad Hoc 1"
matesc(3) = "Ad Hoc 2"
matesc(4) = "Global 1"
matesc(5) = "Global 2"
matesc(6) = "Global 3"
matesc(7) = "Global 4"
matesc(8) = "Deuda Estatal alarmante"
matesc(9) = "Elecciones EU 1"
matesc(10) = "Elecciones EU 2"
matesc(11) = "Jueves Negro"

ReDim mats(1 To noport, 1 To noesc + 7) As Variant
fecha1 = DeterminaFechaTaylor(fecha, txtportCalc2, "PI CONSERVADOS A VENCIMIENTO", "Normal", 1)
fecha2 = DeterminaFechaTaylor(fecha, txtportCalc2, "PI CONSERVADOS A VENCIMIENTO", "Normal", 2)
fecha3 = DeterminaFechaTaylor(fecha, txtportCalc2, "PI CONSERVADOS A VENCIMIENTO", "Normal", 3)
For i = 1 To noport
    mats(i, 1) = CLng(fecha) & MatPortSegRiesgo(i + 8, 1)
    mats(i, 2) = CLng(fecha)
    mats(i, 3) = i
    mats(i, 4) = MatPortSegRiesgo(8, 1)
    For j = 1 To noesc
        mats(i, j + 4) = LeerResEscEstres(fecha, txtportCalc2, MatPortSegRiesgo(i + 8, 1), matesc(j))
    Next j
    mats(i, noesc + 5) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i + 8, 1), fecha1)
    mats(i, noesc + 6) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i + 8, 1), fecha2)
    mats(i, noesc + 7) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i + 8, 1), fecha3)
Next i
GenCuadroEscEstres2 = mats
End Function


Sub CargarCalif(ByRef matpos() As propPosMD, ByVal dtfecha As Date)
   Dim i As Integer, j As Integer, indice As Integer
   Dim matvp() As New propVecPrecios
   Dim mindvp() As Variant

    matvp = LeerVPrecios(dtfecha, mindvp)
    For i = 1 To UBound(matpos, 1)
        indice = 0
        For j = 1 To UBound(matvp, 1)

            If matpos(i).tValorMD = matvp(j).tv And matpos(i).emisionMD = matvp(j).emision And matpos(i).serieMD = matvp(j).serie Then
                indice = j
                Exit For
            End If
        Next j

        If matpos(i).C_Posicion = "8" Or matpos(i).C_Posicion = "9" Then
            If indice <> 0 Then
                If matpos(i).tValorMD = "LD" Or matpos(i).tValorMD = "M" Or matpos(i).tValorMD = "S" Or matpos(i).tValorMD = "IM" Or matpos(i).tValorMD = "IQ" Or matpos(i).tValorMD = "IS" Or matpos(i).tValorMD = "BI" Or matpos(i).tValorMD = "PI" Or matpos(i).tValorMD = "2U" Then
                   matpos(i).CalifTMD = "GUB"
                Else
                   matpos(i).CalifTMD = DetermCalif2(matvp(indice).calif_fitch, matvp(indice).calif_fitch, matvp(indice).calif_fitch, matvp(indice).calif_fitch)
                End If
            Else
                MsgBox "no se califico " & matpos(i).cEmisionMD
            End If
        End If

    Next i
End Sub

Function DetermCalif(ByRef matcalif() As Variant, ByVal indice As Integer) As String
Dim clave As String

If Not EsVariableVacia(matcalif(indice, 5)) Then
clave = matcalif(indice, 5)
   If clave = "A" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "A-" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "AA-" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "mxA" Then
      DetermCalif = "mxA"
   ElseIf clave = "mxA+" Then
      DetermCalif = "mxA+"
   ElseIf clave = "mxA-1+" Then
      DetermCalif = "mxA-1+"
   ElseIf clave = "mxA-2" Then
      DetermCalif = "mxA-2"
   ElseIf clave = "mxAA" Then
      DetermCalif = "mxAA"
   ElseIf clave = "mxAA-" Then
      DetermCalif = "mxAA-"
   ElseIf clave = "mxAA+" Then
      DetermCalif = "mxAA+"
   ElseIf clave = "mxAAA" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "NA" Then
      DetermCalif = "NA"
   ElseIf clave = "RETIRADA" Then
      DetermCalif = "NA"
   End If
ElseIf Not EsVariableVacia(matcalif(indice, 6)) Then
   clave = matcalif(indice, 6)
   If clave = "A" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "A-(mex)" Then
      DetermCalif = "mxA-"
   ElseIf clave = "A+" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "A+(mex)" Then
      DetermCalif = "mxA+"
   ElseIf clave = "AA-" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "AA(mex)" Then
      DetermCalif = "mxAA"
   ElseIf clave = "AA-(mex)" Then
      DetermCalif = "mxAA-"
   ElseIf clave = "AA+(mex)" Then
      DetermCalif = "mxAA+"
   ElseIf clave = "AAA(mex)" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "AAA/1(mex)F" Then
      DetermCalif = "AAA/1(mex)F"
   ElseIf clave = "AAA/2(mex)F" Then
      DetermCalif = "AAA/2(mex)F"
   ElseIf clave = "F1(mex)" Then
      DetermCalif = "mxA-1"
   ElseIf clave = "F1+(mex)" Then
      DetermCalif = "mxA-1+"
   ElseIf clave = "NA" Then
      DetermCalif = "NA"
   ElseIf clave = "RETIRADA" Then
      DetermCalif = "NA"
   End If
ElseIf Not EsVariableVacia(matcalif(indice, 7)) Then
   clave = matcalif(indice, 7)
   If clave = "A1" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "A2" Then
      DetermCalif = "mxAA+"
   ElseIf clave = "Aa1.mx" Then
      DetermCalif = "mxAA+"
   ElseIf clave = "Aa2.mx" Then
      DetermCalif = "mxAA"
   ElseIf clave = "Aa3" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "Aa3.mx" Then
      DetermCalif = "mxAA-"
   ElseIf clave = "Aaa.mx" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "MX-1" Then
      DetermCalif = "mxA-1+"
   ElseIf clave = "MX-2" Then
      DetermCalif = "mxA-1"
   ElseIf clave = "NA" Then
      DetermCalif = "NA"
   End If
ElseIf Not EsVariableVacia(matcalif(indice, 8)) Then
   clave = matcalif(indice, 8)
   If clave = "HR A+" Then
      DetermCalif = "mxA+"
   ElseIf clave = "HR AA" Then
      DetermCalif = "mxAA"
   ElseIf clave = "HR AA-" Then
      DetermCalif = "mxAA-"
   ElseIf clave = "HR AA+" Then
      DetermCalif = "mxAA+"
   ElseIf clave = "HR AAA" Then
      DetermCalif = "mxAAA"
   ElseIf clave = "HR+1" Then
      DetermCalif = "mxA-1+"
   ElseIf clave = "HR1" Then
      DetermCalif = "mxA-1+"
   ElseIf clave = "HR4" Then
      DetermCalif = "mxB"
    ElseIf clave = "NA" Then
      DetermCalif = "NA"
   ElseIf clave = "RETIRADA" Then
      DetermCalif = "NA"
   End If
 
End If

End Function

Function DetermCalif2(ByVal clave1 As String, ByVal clave2 As String, ByVal clave3 As String, ByVal clave4 As String) As String
If Not EsVariableVacia(clave1) Then
   If clave1 = "A" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave1 = "A-" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave1 = "AA-" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave1 = "mxA" Then
      DetermCalif2 = "mxA"
   ElseIf clave1 = "mxA+" Then
      DetermCalif2 = "mxA+"
   ElseIf clave1 = "mxA-1+" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave1 = "mxA-2" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave1 = "mxAA" Then
      DetermCalif2 = "mxAA"
   ElseIf clave1 = "mxAA-" Then
      DetermCalif2 = "mxAA-"
   ElseIf clave1 = "mxAA+" Then
      DetermCalif2 = "mxAA+"
   ElseIf clave1 = "mxAAA" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave1 = "NA" Then
      DetermCalif2 = "NA"
   ElseIf clave1 = "RETIRADA" Then
      DetermCalif2 = "NA"
   End If
ElseIf Not EsVariableVacia(clave2) Then
   If clave2 = "A" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "A-(mex)" Then
      DetermCalif2 = "mxA-"
   ElseIf clave2 = "A+" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "A+(mex)" Then
      DetermCalif2 = "mxA+"
   ElseIf clave2 = "AA-" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "AA(mex)" Then
      DetermCalif2 = "mxAA"
   ElseIf clave2 = "AA-(mex)" Then
      DetermCalif2 = "mxAA-"
   ElseIf clave2 = "AA+(mex)" Then
      DetermCalif2 = "mxAA+"
   ElseIf clave2 = "AAA(mex)" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "AAA/1(mex)F" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "AAA/2(mex)F" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave2 = "F1(mex)" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave2 = "F1+(mex)" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave2 = "NA" Then
      DetermCalif2 = "NA"
   ElseIf clave2 = "RETIRADA" Then
      DetermCalif2 = "NA"
   End If
ElseIf Not EsVariableVacia(clave3) Then
   If clave3 = "A1" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave3 = "A2" Then
      DetermCalif2 = "mxAA+"
   ElseIf clave3 = "Aa1.mx" Then
      DetermCalif2 = "mxAA+"
   ElseIf clave3 = "Aa2.mx" Then
      DetermCalif2 = "mxAA"
   ElseIf clave3 = "Aa3" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave3 = "Aa3.mx" Then
      DetermCalif2 = "mxAA-"
   ElseIf clave3 = "Aaa.mx" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave3 = "MX-1" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave3 = "MX-2" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave3 = "NA" Then
      DetermCalif2 = "NA"
   End If
ElseIf Not EsVariableVacia(clave4) Then
   If clave4 = "HR A+" Then
      DetermCalif2 = "mxA+"
   ElseIf clave4 = "HR AA" Then
      DetermCalif2 = "mxAA"
   ElseIf clave4 = "HR AA-" Then
      DetermCalif2 = "mxAA-"
   ElseIf clave4 = "HR AA+" Then
      DetermCalif2 = "mxAA+"
   ElseIf clave4 = "HR AAA" Then
      DetermCalif2 = "mxAAA"
   ElseIf clave4 = "HR+1" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave4 = "HR1" Then
      DetermCalif2 = "Cplazo"
   ElseIf clave4 = "HR4" Then
      DetermCalif2 = "mxB"
    ElseIf clave4 = "NA" Then
      DetermCalif2 = "NA"
   ElseIf clave4 = "RETIRADA" Then
      DetermCalif2 = "NA"
   End If
End If

End Function

Sub GeneraLSubpEscEstres(ByVal dtfecha As Date, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Integer
Dim noreg As Integer
Dim hinicio As String
Dim hfinal As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim idsubp As Integer

        idsubp = 110
        txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & idsubp
        ConAdo.Execute txtborra
        contar = DeterminaMaxRegSubproc(id_tabla)
        matf = GenFechasEscEstres(dtfecha)
        noreg = UBound(matf, 1)
        For i = 1 To noreg
                contar = contar + 1
                txtcadena = CrearCadInsSub(dtfecha, idsubp, contar, "Esc estrés aprox", i, matf(i, 1), matf(i, 2), "", "", "", "", "", "", "", "", "", id_tabla)
                ConAdo.Execute txtcadena
        Next i

End Sub

Sub GenListaPyGPortPos(ByVal fecha As Date, ByVal id_subproc As Integer, ByVal txtport As String, ByVal txtgrupo As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_tabla As Integer)
Dim contar As Integer
Dim txtfiltro As String
Dim txtfecha1 As String
Dim txtcadena As String
txtfecha1 = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
contar = DeterminaMaxRegSubproc(id_tabla)
contar = contar + 1
txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Consolida p y g pot portafolio", fecha, fecha, fecha, txtport, "", "Normal", txtgrupo, noesc, htiempo, "", "", "", id_tabla)
ConAdo.Execute txtcadena

End Sub

Sub ExpPyGPortPostxt(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupo As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef exito As Boolean)
On Error GoTo hayerror
Dim valor As Variant
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim nomarch As String
Dim matc() As String
Dim txtcadena As String
Dim cposicion As Integer
Dim fregistro As Date
Dim coperacion As String
Dim j As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtgrupo & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matpl(1 To noesc) As Double
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma = 0
   nomarch = DirResVaR & "\escenarios pyg.txt"
   Open nomarch For Output As #1

   For i = 1 To noreg
       valt01 = rmesa.Fields(8)
       valor = rmesa.Fields(9).GetChunk(rmesa.Fields(9).ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       txtcadena = ""
       txtcadena = txtcadena & cposicion & Chr(9)
       txtcadena = txtcadena & fregistro & Chr(9)
       txtcadena = txtcadena & coperacion & Chr(9)
       txtcadena = txtcadena & valt01 & Chr(9)
       For j = 1 To noesc
           txtcadena = txtcadena & matc(j) & Chr(9)
       Next j
       Print #1, txtcadena
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Exportando las p y g del portafolio " & txtgrupo & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   Close #1
   rmesa.Close
  
   
End If
Exit Sub

hayerror:

End Sub

Function FiltrarFactSens(ByRef matren() As Double, ByRef matsen() As Variant, ByRef MatDeffact() As Variant) As Variant()
Dim nofr As Long
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim kk As Long
Dim contar As Long

'en funcion del indice se ordena la matriz matren y se procede a seleccionar los escenarios'
'mas importantes para valuar el portafolio
nofr = UBound(matsen, 1)      'el no de factores de riesgo determinados por el var markowitz
noreg = UBound(matren, 1)     'no de escenarios en la matriz de escenarios
ReDim matb(1 To noreg, 1 To 1) As Variant
ReDim MatDeffact(1 To 6, 1 To 1) As Variant
'se copian las fechas en una matriz auxiliar
For i = 1 To noreg
    matb(i, 1) = matren(i, 1)
Next i
contar = 0
For i = 1 To nofr
'se exploran todos los factores de riesgo
    For j = 1 To NoFactores
        If matsen(i, 1) = MatCaracFRiesgo(j).indFactor Then
           contar = contar + 1
           ReDim Preserve matb(1 To noreg, 1 To contar) As Variant
           For kk = 1 To noreg
               matb(kk, i) = matren(kk, j)
           Next kk
'se devuelve una matriz con los resultados del filtro
ReDim Preserve MatDeffact(1 To 6, 1 To contar) As Variant
MatDeffact(1, contar) = MatCaracFRiesgo(j).indFactor
MatDeffact(2, contar) = MatCaracFRiesgo(j).tfactor
MatDeffact(3, contar) = MatCaracFRiesgo(j).nomFactor
MatDeffact(4, contar) = MatCaracFRiesgo(j).plazo
MatDeffact(5, contar) = MatCaracFRiesgo(j).tinterpol
MatDeffact(6, contar) = MatCaracFRiesgo(j).descFactor
Exit For
End If
Next j
Next i

FiltrarFactSens = matb
End Function

Sub GuardarEscEstresAprox(ByVal dtfecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByRef mata() As Variant)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtborra As String
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim numbloques As Integer
Dim txtfiltro As String
Dim i As Long
Dim noreg As Long
Dim largo As Long
Dim leftover As Long
Dim txttexto As String
Dim RInterfIKOS As New ADODB.recordset

  txtfecha = "TO_DATE('" & Format(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
  txtfecha1 = "TO_DATE('" & Format(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
  txtfecha2 = "TO_DATE('" & Format(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"
  txtborra = "DELETE FROM  " & TablaResEstresAprox & " WHERE FECHA = " & txtfecha
  txtborra = txtborra & " AND FECHA1 = " & txtfecha1
  txtborra = txtborra & " AND FECHA2 = " & txtfecha2
  txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
  txtborra = txtborra & " AND ESC_FR = '" & txtportfr & "'"
  txtborra = txtborra & " AND SUBPORTAFOLIO = '" & txtsubport & "'"
  ConAdo.Execute txtborra
  txtfiltro = "SELECT * FROM " & TablaResEstresAprox
  RInterfIKOS.Open txtfiltro, ConAdo, 1, 3
  noreg = UBound(mata, 1)
         txtcadena1 = ""
         txtcadena2 = ""
         For i = 1 To noreg - 1
             txtcadena1 = txtcadena1 & mata(i, 1) & ","
             txtcadena2 = txtcadena2 & mata(i, 2) & ","
         Next i
         txtcadena1 = txtcadena1 & mata(noreg, 1)
         txtcadena2 = txtcadena2 & mata(noreg, 2)
         RInterfIKOS.AddNew
         RInterfIKOS.Fields("FECHA") = CLng(dtfecha)          'la fecha de proceso
         RInterfIKOS.Fields("FECHA1") = CLng(fecha1)          'fecha inicial de escenarios
         RInterfIKOS.Fields("FECHA2") = CLng(fecha2)          'fecha final
         RInterfIKOS.Fields("PORTAFOLIO") = txtport           'nombre del portafolio
         RInterfIKOS.Fields("ESC_FR") = txtportfr             'ESCENARIOS
         RInterfIKOS.Fields("SUBPORTAFOLIO") = txtsubport     'nombre del portafolio
         Call GuardarElementoClob(txtcadena1, RInterfIKOS, "AFECHAC")
         Call GuardarElementoClob(txtcadena2, RInterfIKOS, "ACALCULOS")
         RInterfIKOS.Update
         RInterfIKOS.Close
         MensajeProc = "Guardando los resultados de simulacion de estres por aproximacion lineal"
End Sub

Function LeerEscFR(ByVal txtnomesc As String, ByVal fecha As Date) As Variant()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaEscFR & " WHERE ID_ESCENARIO = '" & txtnomesc & "' AND FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("FACTOR")
       mata(i, 2) = rmesa.Fields("VALOR")
       rmesa.MoveNext
   Next i
   rmesa.Close
   mata = RutinaOrden(mata, 1, SRutOrden)
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If

LeerEscFR = mata
End Function


Sub GenProcEfecProsSwap(ByVal fecha As Date, ByVal txtport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim contar As Long
Dim txttabla As String
txttabla = DetermTablaSubproc(id_tabla)

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtfecha1 = Format(fecha, "dd/mm/yyyy")
    txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Efect Cob Prospectiva Swap", txtfecha1, txtport, "", "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena

End Sub

