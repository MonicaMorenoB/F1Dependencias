Attribute VB_Name = "modCVA"
Option Explicit

Sub GeneraLSubprocCVA(ByVal dtfecha As Date, _
                      ByVal id_subproc As Integer, _
                      ByVal txtport As String, _
                      ByVal noesc As Integer, _
                      ByVal htiempo As Integer, _
                      ByVal id_tabla As Integer, _
                      ByRef txtmsg As String, _
                      ByRef exito As Boolean)
'genera la lista de subprocesos para el calculo del CVA de la posicion de derivados
'datos de entrada:
'dtfecha      - fecha del proceso
'id_subproc   - id de subproceso al que se llamara en la rutina de ejecución de subprocesos
'txtport      - portafolio de posicion que se utilizara para la creacion de subprocesos
'noesc        - numero de escenarios a calcular en simulacion historica
'htiempo      - no de dias para el calculo de rendimientos
'id_tabla     - tabla a la que se apuntara para la creacion de subprocesos
'txtmsg       - la

                      
    Dim i As Long
    Dim j As Long
    Dim noreg As Long
    Dim contar As Long
    Dim txtfecha As String
    Dim txtborra As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim tipopos As Integer
    Dim fechareg As Date
    Dim txtnompos As String
    Dim txthorareg As String
    Dim C_Posicion As Integer
    Dim c_operacion As String
    Dim rmesa As New ADODB.recordset
    
    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    ConAdo.Execute txtborra
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
    txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For i = 1 To noreg
           contar = contar + 1
           tipopos = rmesa.Fields("TIPOPOS")
           fechareg = rmesa.Fields("FECHAREG")
           txtnompos = rmesa.Fields("NOMPOS")
           txthorareg = rmesa.Fields("HORAREG")
           C_Posicion = rmesa.Fields("CPOSICION")
           c_operacion = rmesa.Fields("COPERACION")
           Call GenRegSubpCVA(id_subproc, contar, dtfecha, tipopos, fechareg, txtnompos, txthorareg, C_Posicion, c_operacion, noesc, htiempo, id_tabla)
           rmesa.MoveNext
       Next i
       rmesa.Close
    End If
    txtmsg = "El proceso finalizo correctamente"
    exito = True
End Sub

Sub GenRegSubpCVA(ByVal id_proc As Integer, ByVal contar As Long, ByVal dtfecha As Date, ByVal tipopos As Integer, ByVal dtfechar As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_tabla As Integer)
    Dim txtfecha  As String
    Dim txtcadena As String
    txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "Cálculo de CVA", tipopos, dtfechar, txtnompos, horareg, cposicion, coperacion, noesc, htiempo, "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
End Sub

Function CrearCadInsSub(ByVal fecha As Date, ByVal id_proc As Long, ByVal contar As Long, ByVal descrip As String, ByVal par1 As Variant, ByVal par2 As Variant, ByVal par3 As Variant, ByVal par4 As Variant, ByVal par5 As Variant, ByVal par6 As Variant, ByVal par7 As Variant, ByVal par8 As Variant, ByVal par9 As Variant, ByVal par10 As Variant, ByVal par11 As Variant, ByVal par12 As Variant, ByVal id_tabla As Integer)
Dim txtcadena As String
Dim txtfecha As String
Dim txttabla As String
   
    txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & DetermTablaSubproc(id_tabla) & " VALUES("
    txtcadena = txtcadena & id_proc & ","                        'clave de la tarea
    txtcadena = txtcadena & contar & ","                         'folio de la tarea
    txtcadena = txtcadena & "'" & descrip & "',"                 'descripcion de la tarea
    If Not EsVariableVacia(par1) Then
       txtcadena = txtcadena & "'" & par1 & "',"                 'parametro 1
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par2) Then
       txtcadena = txtcadena & "'" & par2 & "',"                 'parametro 2
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par3) Then
       txtcadena = txtcadena & "'" & par3 & "',"                 'parametro 3
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par4) Then
       txtcadena = txtcadena & "'" & par4 & "',"                 'parametro 4
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par5) Then
       txtcadena = txtcadena & "'" & par5 & "',"                 'parametro 5
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par6) Then
       txtcadena = txtcadena & "'" & par6 & "',"                 'parametro 6
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par7) Then
       txtcadena = txtcadena & "'" & par7 & "',"                 'parametro 7
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par8) Then
       txtcadena = txtcadena & "'" & par8 & "',"                 'parametro 8
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par9) Then
       txtcadena = txtcadena & "'" & par9 & "',"                 'parametro 9
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par10) Then
       txtcadena = txtcadena & "'" & par10 & "',"                 'parametro 10
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par11) Then
       txtcadena = txtcadena & "'" & par11 & "',"                 'parametro 11
    Else
       txtcadena = txtcadena & "null,"
    End If
    If Not EsVariableVacia(par12) Then
       txtcadena = txtcadena & "'" & par12 & "',"                 'parametro 12
    Else
       txtcadena = txtcadena & "null,"
    End If

    txtcadena = txtcadena & txtfecha & ","                       'fecha proceso
    txtcadena = txtcadena & "null,"                              'fecha de inicio
    txtcadena = txtcadena & "null,"                              'hora de inicio
    txtcadena = txtcadena & "null,"                              'fecha final
    txtcadena = txtcadena & "null,"                              'hora final
    txtcadena = txtcadena & "'N',"                               'bloqueada
    txtcadena = txtcadena & "'N',"                               'finalizada
    txtcadena = txtcadena & "'N',"                               'exito
    txtcadena = txtcadena & "null,"                              'comentario
    txtcadena = txtcadena & "null,"                              'usuario
    txtcadena = txtcadena & "null)"                              'direccion ip
    CrearCadInsSub = txtcadena
End Function

Sub GenSubprocWRW(ByVal dtfecha As Date, ByVal id_tabla As Integer)
    Dim i As Integer
    Dim noreg As Integer
    Dim txtfecha As String
    Dim txtport As String
    Dim txtcadena As String
    Dim txtborra As String
    Dim txtfiltro As String
    Dim contar As Long
    Dim rmesa As New ADODB.recordset
    Dim id_proc As Integer
    id_proc = 93
   
       txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtborra = "DELETE FROM " & TablaResWRW & " WHERE FECHA = " & txtfecha
       ConAdo.Execute txtborra
       txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
       ConAdo.Execute txtborra
       contar = DeterminaMaxRegSubproc(id_tabla)
       For i = 1 To UBound(MatContrapartes, 1)
           txtport = "Deriv Contrap " & MatContrapartes(i, 1)
           txtfiltro = "SELECT COUNT(*) FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
           rmesa.Open txtfiltro, ConAdo
           noreg = rmesa.Fields(0).value
           rmesa.Close
           If noreg <> 0 Then
              contar = contar + 1
              txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "WRW", MatContrapartes(i, 1), "", "", "", "", "", "", "", "", "", "", "", id_tabla)
              ConAdo.Execute txtcadena
           End If
       Next i
End Sub

Sub CalcEPE(ByVal dtfecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtcadena As String
Dim i As Integer
Dim noreg As Integer, j As Integer
Dim suma1 As Double, suma2 As Double
Dim contar1 As Integer, contar2 As Integer
Dim nomarch As String
Dim valmax1 As Double
Dim valmax2 As Double
Dim matb() As Variant
Dim rmesa As New ADODB.recordset

   nomarch = "d:\Reporte EPE " & Format$(dtfecha, "yyyymmdd") & ".txt"
   txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #3
   txtcadena = "Sector" & Chr$(9) & "Contraparte" & Chr$(9) & "Menor a 1 año" & Chr$(9) & "Plazo remanente" & Chr$(9) & "Estres 1" & Chr$(9) & "Estres 2"
   Print #3, txtcadena
   For i = 1 To UBound(MatContrapartes, 1)
       txtfiltro = "SELECT * FROM " & TablaResCVA & " WHERE FECHA = " & txtfecha & " AND ID_CONTRAP = '" & MatContrapartes(i, 1) & "' AND ID_CALCULO = 'CVA' AND CPOSICION = 'DER' AND SUMAVP > 0 ORDER BY DXV"
       txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          ReDim mata(1 To noreg, 1 To 2) As Variant
          rmesa.Open txtfiltro, ConAdo
          For j = 1 To noreg
              mata(j, 1) = rmesa.Fields("DXV")      'dias por vencer
              mata(j, 2) = rmesa.Fields("SUMAVP")   'valor promedio positivo
              rmesa.MoveNext
          Next j
          rmesa.Close
          suma1 = 0: suma2 = 0
          contar1 = 0: contar2 = 0
          For j = 1 To noreg
              If mata(j, 1) <= 365 Then
                 suma1 = suma1 + mata(j, 2)
                 contar1 = contar1 + 1
              End If
              suma2 = suma2 + mata(j, 2)
              contar2 = contar2 + 1
          Next j
          If contar1 <> 0 Then suma1 = suma1 / contar1
          If contar2 <> 0 Then suma2 = suma2 / contar2
          matb = RutinaOrden(mata, 2, SRutOrden)
          valmax1 = Maximo(matb(noreg, 2), 0)
          If noreg > 1 Then valmax2 = Maximo(matb(noreg - 1, 2), 0)
          txtcadena = CLng(dtfecha) & Chr(9) & CLng(dtfecha) & MatContrapartes(i, 3) & Chr(9) & MatContrapartes(i, 6) & Chr$(9) & MatContrapartes(i, 3) & Chr$(9) & suma1 & Chr$(9) & suma2 & Chr(9) & valmax1 & Chr(9) & valmax2
          Print #3, txtcadena
       End If
   Next i
   Close #3
End Sub

Function GenRepEPE1(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal Sector As String)
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim contar As Integer
Dim contar1 As Integer
Dim contar2 As Integer
Dim matb() As Variant
Dim matr() As Variant
Dim matv() As Double
Dim valor As Double
Dim suma1 As Double, suma2 As Double, suma3 As Double
Dim valmax1 As Double
Dim valmax2 As Double
Dim rmesa As New ADODB.recordset

ReDim matr(1 To 8, 1 To 1) As Variant
   txtfecha1 = "to_date('" & Format$(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format$(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(MatContrapartes, 1)
       If MatContrapartes(i, 6) = Sector Then
          txtfiltro2 = "SELECT * FROM " & TablaResCVA & " WHERE FECHA = " & txtfecha2 & " AND ID_CONTRAP = '" & MatContrapartes(i, 1) & "' AND CPOSICION ='DER' AND ID_CALCULO = 'CVA' AND SUMAVP >0 ORDER BY DXV"
          txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
          rmesa.Open txtfiltro1, ConAdo
          noreg = rmesa.Fields(0)
          rmesa.Close
          suma1 = 0: suma2 = 0: suma3 = 0: valmax1 = 0: valmax2 = 0
          If noreg <> 0 Then
             ReDim mata(1 To noreg, 1 To 2) As Variant
             rmesa.Open txtfiltro2, ConAdo
             For j = 1 To noreg
                 mata(j, 1) = rmesa.Fields("DXV")      'dias por vencer
                 mata(j, 2) = rmesa.Fields("SUMAVP")   'valor promedio positivo
                 rmesa.MoveNext
             Next j
             rmesa.Close
             suma1 = 0: suma2 = 0
             contar1 = 0: contar2 = 0
             For j = 1 To noreg
                 If mata(j, 1) <= 365 Then
                    suma1 = suma1 + mata(j, 2)
                    contar1 = contar1 + 1
                 End If
                 suma2 = suma2 + mata(j, 2)
                 contar2 = contar2 + 1
             Next j
             If contar1 <> 0 Then suma1 = suma1 / contar1
             If contar2 <> 0 Then suma2 = suma2 / contar2
             matb = RutinaOrden(mata, 2, SRutOrden)
             valmax1 = Maximo(matb(noreg, 2), 0)
             If noreg > 1 Then valmax2 = Maximo(matb(noreg - 1, 2), 0)
             txtfiltro2 = "SELECT * FROM " & TablaResCVA & " WHERE FECHA = " & txtfecha1 & " AND ID_CONTRAP = '" & MatContrapartes(i, 1) & "' AND CPOSICION = 'DER' AND ID_CALCULO = 'CVA' AND SUMAVP > 0 ORDER BY DXV"
             txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
             rmesa.Open txtfiltro1, ConAdo
             noreg1 = rmesa.Fields(0)
             rmesa.Close
             suma3 = 0
             If noreg1 <> 0 Then
                rmesa.Open txtfiltro2, ConAdo
                For j = 1 To noreg1
                    valor = rmesa.Fields("SUMAVP")   'valor promedio positivo
                    suma3 = suma3 + valor
                    rmesa.MoveNext
                Next j
                suma3 = suma3 / noreg1
                rmesa.Close
             Else
                suma3 = 0
             End If
          End If
          matv = LeerResValPort(fecha2, "TOTAL", "Deriv Contrap " & MatContrapartes(i, 1), 2)
          If UBound(matv, 1) <> 0 Then
          If suma1 <> 0 Or suma2 <> 0 Or valmax1 <> 0 Or valmax2 <> 0 Or suma3 <> 0 Or matv(1) <> 0 Then
             contar = contar + 1
             ReDim Preserve matr(1 To 8, 1 To contar) As Variant
             matr(1, contar) = MatContrapartes(i, 3)
             matr(2, contar) = MatContrapartes(i, 1)
             If UBound(matv, 1) <> 0 Then
                matr(3, contar) = matv(1) / 1000000
             Else
                matr(3, contar) = 0
             End If
             matr(4, contar) = suma1 / 1000000
             matr(5, contar) = suma2 / 1000000
             matr(6, contar) = valmax1 / 1000000
             matr(7, contar) = valmax2 / 1000000
             matr(8, contar) = suma3 / 1000000
          End If
          End If
      End If
   Next i
   matr = MTranV(matr)
   matr = RutinaOrden(matr, 1, SRutOrden)
   GenRepEPE1 = matr
End Function


Sub ProcCalculoWRW(ByVal dtfecha As Date, ByVal id_contrap As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
    Dim i           As Integer, j As Integer
    Dim noreg As Integer
    Dim ll As Integer
    Dim escala      As String
    Dim Sector      As String
    Dim sector2      As String
    Dim Threshold   As Double
    Dim mmtransfer  As Double
    Dim califica    As Integer
    Dim mtrans()    As Double
    Dim mrecupera() As Variant
    Dim noesc       As Integer
    Dim tinterpol   As Integer
    Dim mata()      As Variant
    Dim matv1()      As Variant
    Dim matv2()      As Variant
    Dim vrecupera1  As Double
    Dim matprob1()  As Double
    Dim matprob2()  As Double
    Dim matprob3()  As Double
    Dim matprob4()  As Double
    Dim curva1()    As propCurva
    Dim contar()    As Integer
    Dim vtasa       As Double
    Dim txtfecha    As String
    Dim txtborra    As String
    Dim txtinserta  As String
    Dim bl_exito    As Boolean
    Dim suma1()      As Double
    Dim suma2()      As Double
    Dim vlambda     As Double
    Dim mtmgar1()    As Variant
    Dim mtmgar2()    As Variant
    Dim matff()     As Date
    Dim matpl() As Variant
    Dim noff As Integer
    Dim sicalc As Boolean
    Dim MatFactoresR() As Double
    Dim valcva As Double
    Dim exito1 As Boolean

    vlambda = 1
    noesc = 500
    tinterpol = 1
    matpl = LeerResPLCVAContrap(dtfecha, id_contrap, noesc, 1, matff)
    noff = UBound(matff, 1)
    noreg = UBound(matpl, 1)
    If noreg <> 0 And noff <> 0 Then
       ReDim matpl1(1 To noff, 1 To noesc + 1) As Double
       ReDim matpl2(1 To noff, 1 To noesc + 1) As Double
       For i = 1 To noff
           matpl1(i, 1) = matff(i, 1)
           matpl2(i, 1) = matff(i, 1)
       Next i
       MatDerivSinLMargen = CargaDerivSinLMargen(dtfecha)
       For i = 1 To noreg
           sicalc = DetermOperBlack(matpl(i, 1))
           For ll = 1 To noff
                If matpl(i, 2) = matff(ll, 1) Then
                   For j = 1 To noesc
                       If Not sicalc Then
                          matpl1(ll, j + 1) = matpl1(ll, j + 1) + matpl(i, j + 2)
                       Else
                          matpl2(ll, j + 1) = matpl2(ll, j + 1) + matpl(i, j + 2)
                       End If
                   Next j
                End If
            Next ll
            AvanceProc = i / noreg
            MensajeProc = "Procesando los p&l del CVA del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
            DoEvents
        Next i
        Call DeterminaParCVA(dtfecha, id_contrap, escala, Sector, sector2, Threshold, mmtransfer, califica, mtrans, mrecupera, exito1)
        vrecupera1 = Recuperacion(califica, mrecupera, sector2)
        matprob1 = CalcProbDefault(califica, mtrans, noff)  'probabilidad normal
        matprob2 = CalcProbDefault(1, mtrans, noff)         'probabilidad soberana
        matprob3 = CalcProbDefault(califica, mtrans, noff)  'probabilidad normal
        matprob4 = CalcProbDefault(1, mtrans, noff)         'probabilidad soberana
        For i = 1 To noff
             matprob1(i) = matprob1(i) * (1 - vrecupera1)
             matprob2(i) = matprob1(i) * (1 - vrecupera1)
             'Print #1, id_contrap & Chr(9) & matprob1(i) & Chr(9) & matprob2(i) & Chr(9) & matprob3(i); Chr(9) & matprob4(i)
        Next i
        
        curva1 = LeerCurvaC(dtfecha, "DESC IRS")
        ReDim matresul(1 To noff, 1 To 5) As Variant
        ReDim suma(1 To noff) As Double
        ReDim contar(1 To noff) As Integer
        ReDim matpera(1 To noff, 1 To 9) As Double
        ReDim matper(1 To noff, 1 To 9) As Double
        ReDim mtmgar1(1 To noff, 1 To noesc) As Variant
        ReDim mtmgar2(1 To noff, 1 To noesc) As Variant
        For i = 1 To noff
            For j = 1 To noesc
                If matpl1(i, j + 1) > Threshold And mmtransfer <> 0 Then
                    mtmgar1(i, j) = matpl1(i, j + 1) - Int((matpl1(i, j + 1) - Threshold) / mmtransfer) * mmtransfer
                Else
                    mtmgar1(i, j) = matpl1(i, j + 1)
                End If
                mtmgar2(i, j) = matpl2(i, j + 1)
            Next j
        Next i
        For i = 1 To noff
            suma(i) = 0
            contar(i) = 0
            'suma2(i) = 0
            'contar2(i) = 0
            For j = 1 To noesc
                If mtmgar1(i, j) + mtmgar2(i, j) > 0 Then
                    'solo se suman las marcas a mercado positivas
                    suma(i) = suma(i) + mtmgar1(i, j) + mtmgar2(i, j)
                    contar(i) = contar(i) + 1
                End If
            Next j
            If contar(i) <> 0 Then
                suma(i) = suma(i) / contar(i)
            Else
                suma(i) = 0
            End If
        Next i

        ReDim percen(1 To 9) As Double
        percen(1) = 0.8
        percen(2) = 0.84
        percen(3) = 0.85
        percen(4) = 0.9
        percen(5) = 0.96
        percen(6) = 0.97
        percen(7) = 0.975
        percen(8) = 0.98
        percen(9) = 0.99
        For i = 1 To noff
            matv1 = ExtraeSubMatV(mtmgar1, 1, noesc, i, i)
            matv2 = ExtraeSubMatV(mtmgar2, 1, noesc, i, i)
            ReDim matv(1 To UBound(matv1, 2), 1 To 1) As Variant
            For j = 1 To UBound(matv1, 2)
                matv(j, 1) = matv1(1, j) + matv2(1, j)
            Next j
            For j = 1 To 9
                If Abs(Round(contar(i) * percen(j)) - contar(i) * percen(j)) > 0.00001 Then
                    matpera(i, j) = Round(contar(i) * percen(j) + 0.5, 0)
                Else
                    matpera(i, j) = Round(contar(i) * percen(j), 0)
                End If
                matper(i, j) = ExtNPos(matv, matpera(i, j))
            Next j
        Next i
        ReDim matfact(1 To noff, 1 To 2) As Double
        ReDim matvar1(1 To 9) As Double
        ReDim matvar2(1 To 9) As Double
        valcva = 0
        For i = 1 To noff
            matresul(i, 1) = matff(i, 1) - dtfecha
            matresul(i, 2) = suma(i)
            vtasa = CalculaTasa(curva1, matresul(i, 1), tinterpol)
            matresul(i, 3) = 1 / (1 + vtasa * matresul(i, 1) / 360)                             'FD
            matresul(i, 4) = matresul(i, 2) * matprob1(i) / (1 + vtasa * matresul(i, 1) / 360)   'VTOTAL
            valcva = valcva + matresul(i, 4)
            If matprob1(i) <> 0 Then
                matfact(i, 1) = 1 - tanh(matprob2(i) / matprob1(i))
            Else
                matfact(i, 1) = 0
            End If

            If matprob3(i) <> 0 Then
                matfact(i, 2) = 1 - tanh(matprob4(i) / matprob3(i))
            Else
                matfact(i, 2) = 0
            End If
            For j = 1 To 9
                matvar1(j) = matvar1(j) + (suma(i) + vlambda * (matper(i, j) - suma(i)) * matfact(i, 1)) * matprob1(i) * matresul(i, 3)
                matvar2(j) = matvar2(j) + (suma(i) + vlambda * (matper(i, j) - suma(i)) * matfact(i, 2)) * matprob1(i) * matresul(i, 3)
            Next j
        Next i
  
        For i = 1 To 9
            matvar1(i) = matvar1(i) - valcva
            matvar2(i) = matvar2(i) - valcva
        Next i
        txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtinserta = "INSERT INTO " & TablaResWRW & " VALUES("
        txtinserta = txtinserta & txtfecha & ","
        txtinserta = txtinserta & id_contrap & ","
        txtinserta = txtinserta & valcva & ","
        For j = 1 To 9
            txtinserta = txtinserta & matvar1(j) & ","
        Next j
        For j = 1 To 9
            txtinserta = txtinserta & matvar2(j) & ","
        Next j
        txtinserta = txtinserta & "'" & NomUsuario & "')"
        ConAdo.Execute txtinserta
        txtmsg = "Proceso finalizado correctamente"
        exito = True
    Else
        exito = True
    End If
 
End Sub

Function ExtNPos(ByRef mata() As Variant, ByVal nr As Integer)
    Dim matb() As Variant
    Dim contar As Integer, i As Integer
    matb = RutinaOrden(mata, 1, 3)
    contar = 1
    For i = 1 To UBound(matb, 1)
        If matb(i, 1) > 0 And contar = nr Then
            ExtNPos = matb(i, 1)
            Exit Function
        ElseIf matb(i, 1) > 0 Then
            contar = contar + 1
        End If
    Next i
    ExtNPos = 0
End Function

Sub CalcCVAOper(ByVal fecha As Date, ByVal tipopos As Integer, ByVal fechar As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal noesc As Long, ByVal htiempo As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim indice0 As Long
Dim indice As Long
Dim fecha1 As Date
Dim mrvalflujo() As resValFlujo
Dim matx() As Variant
Dim matx1() As Double
Dim matrends() As Double
Dim parval As ParamValPos
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim curva1() As New propCurva
Dim mtrans() As Double
Dim matfshist() As Variant
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim mattxt() As String
Dim mtmoper() As Double
Dim MPrecio() As resValIns
Dim suma As Double
Dim matfechas1() As Date
Dim matfrsim() As Double
Dim txtmsg2 As String
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg3 As String
Dim exitofr As Boolean

   exito = False
   ValExacta = False
   Call VerifCargaFR2(fecha, noesc + htiempo, exitofr)
   mattxt = CrearFiltroPosOperPort(tipopos, fechar, txtnompos, horareg, cposicion, coperacion)
   Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
   noreg = UBound(matpos, 1)
   If noreg <> 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
      If exito2 Then
         SiIncTasaCVig = False
         matfechas1 = DetFechasCalcM(fecha, 3, matpos, matposswaps, matposfwd, matposdeuda)
         If noescSH <> noesc Or htiempoSH <> htiempo Or fechaSH <> fecha Then
            Call PrevCVaRHistorico(fecha, noesc, htiempo, MatFactRiesgo, matrendsSH, matBndSH)
            noescSH = noesc
            htiempoSH = htiempo
            fechaSH = fecha
         End If
         If fecha <> FechaMatFactR1 Then
            MatFactR1 = CargaFR1Dia(fecha, exito1)
            FechaMatFactR1 = fecha
         End If
     'Se carga la estructura de tasas para ese día de la matriz vector tasas
         Set parval = DeterminaPerfilVal("HISTORICO")
         For i = 1 To UBound(matfechas1, 1)
             ReDim mtmoper(1 To noesc) As Double
             For j = 1 To noesc
                 matfrsim = GenEscHist2(MatFactR1, matrendsSH, matBndSH, j)
                 parval.perfwd = matfechas1(i) - fecha                       'periodo forward
                 MPrecio = CalcValuacion(matfechas1(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matfrsim, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
                 suma = 0
                 For p = 1 To noreg
                     suma = suma + MPrecio(p).mtm_sucio
                 Next p
                 mtmoper(j) = suma
             Next j
             Call GuardarEscCVA(fecha, tipopos, fechar, txtnompos, horareg, cposicion, coperacion, matfechas1(i), noesc, htiempo, mtmoper)
             AvanceProc = i / UBound(matfechas1, 1)
             MensajeProc = "Clave de operacion " & coperacion & " Fecha de valuacion " & matfechas1(i) & " " & Format(AvanceProc, "##0.00 %")
             DoEvents
         Next i
         exito = True
         txtmsg = "Proceso finalizado correctamente"
         SiIncTasaCVig = True
      Else
        exito = False
        txtmsg = txtmsg2
      End If
   End If
End Sub

Sub CalcEscMakeW(ByVal fecha As Date, ByVal tipopos As Integer, ByVal fechar As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal htiempo As Long, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal orden As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim fecha0 As Date
Dim exito1 As Boolean
Dim indice0 As Long
Dim indice As Long
Dim mrvalflujo() As resValFlujo
Dim matx() As Variant
Dim matx1() As Double
Dim matrends() As Double
Dim parval As ParamValPos
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim curva1() As New propCurva
Dim mtrans() As Double
Dim matfechas1() As Date
Dim matfshist() As Variant
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim noesc As Integer
Dim matb() As Integer
Dim mattxt() As String
Dim mtmoper() As Double
Dim MPrecio() As resValIns
Dim suma As Double
Dim matfrsim() As Double
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String


   indice0 = BuscarValorArray(fecha1, MatFechasVaR, 1)
   fecha0 = MatFechasVaR(indice0 - htiempo, 1)
   exito = False
   ValExacta = False
   SiIncTasaCVig = False
   Call VerifCargaFR(fecha0, fecha2)
   noesc = UBound(MatFactRiesgo, 1) - htiempo
   mattxt = CrearFiltroPosOperPort(tipopos, fechar, txtnompos, horareg, cposicion, coperacion)
   Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
   noreg = UBound(matpos, 1)
   If noreg <> 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
      matfechas1 = DetFechasCalcM2(fecha, 28, matpos, matposswaps, matposfwd, matposdeuda)
      If exito2 Then
         'If noescSH <> noesc Or htiempoSH <> htiempo Or fechaSH <> fecha2 Then
            Call PrevCVaRHistorico(fecha2, noesc, htiempo, MatFactRiesgo, matrendsSH, matb)
            noescSH = noesc
            htiempoSH = htiempo
            fechaSH = fecha
         'End If
         If fecha <> FechaMatFactR1 Then
            MatFactR1 = CargaFR1Dia(fecha, exito1)
            FechaMatFactR1 = fecha
         End If
     'Se carga la estructura de tasas para ese día de la matriz vector tasas
         Set parval = DeterminaPerfilVal("HISTORICO")
         ReDim mtmoper(1 To noesc) As Double
         For i = 1 To UBound(matfechas1, 1)
             parval.perfwd = matfechas1(i) - fecha
             'parval.perfwd = 0
             For j = 1 To noesc
                 matfrsim = GenEscHist2(MatFactR1, matrendsSH, matb, j)
                 MPrecio = CalcValuacion(matfechas1(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matfrsim, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
                 suma = 0
                 For p = 1 To noreg
                     suma = suma + MPrecio(p).mtm_sucio
                 Next p
                 mtmoper(j) = suma
                 AvanceProc = j / noesc
                 MensajeProc = "Clave de operacion " & coperacion & " Fecha de valuacion " & fecha & " " & Format(AvanceProc, "##0.00 %")
                 DoEvents
             Next j
             Call GuardarEscMakeW(fecha, tipopos, fechar, txtnompos, horareg, cposicion, coperacion, matfechas1(i), fecha1, fecha2, orden, noesc, htiempo, mtmoper)
         Next i
         SiIncTasaCVig = True
         txtmsg = "El proceso finalizo correctamente"
         exito = True
      Else
         exito = False
         txtmsg = txtmsg2
      End If
   End If
End Sub

Sub GuardarEscCVA(ByVal dtfecha As Date, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal fechaf As Date, ByVal noesc As Long, ByVal htiempo As Long, ByRef mtm() As Double)
    Dim txtfecha As String
    Dim txtfechaf As String
    Dim txtfechar As String
    Dim txtfiltro As String
    Dim txtcadena As String
    Dim txttexto As String
    Dim i As Long
    Dim largo As Long
    Dim largoseg As Long
    Dim noseg As Long
    Dim residuo As Long
    
        txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtfechaf = "TO_DATE('" & Format$(fechaf, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtfechar = "TO_DATE('" & Format$(fechareg, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtcadena = ""
        For i = 1 To noesc - 1
            txtcadena = txtcadena & mtm(i) & ","
        Next i
            txtcadena = txtcadena & mtm(noesc)
            RegResCVA.AddNew
            RegResCVA.Fields("FECHA") = CLng(dtfecha)
            RegResCVA.Fields("TIPOPOS") = tipopos                  'tipo de posicion
            RegResCVA.Fields("FECHAREG") = CLng(fechareg)          'fecha de registro
            RegResCVA.Fields("NOMPOS") = txtnompos                 'NOMBRE DE LA POSICION
            RegResCVA.Fields("HORAREG") = horareg                  'HORAREG
            RegResCVA.Fields("CPOSICION") = cposicion              'clave de la posicion
            RegResCVA.Fields("COPERACION") = coperacion            'clave de operacion
            RegResCVA.Fields("FECHA_F") = CLng(fechaf)             'fecha futura de valuacion
            RegResCVA.Fields("NO_ESC") = noesc                     'numero de escenarios
            RegResCVA.Fields("H_TIEMPO") = htiempo                 'horizonte de tiempo
            RegResCVA.Fields("USUARIO") = NomUsuario               'nombre del usuario
            Call GuardarElementoClob(txtcadena, RegResCVA, "VECTOR_PYG")
            RegResCVA.Update

End Sub


Sub GuardarEscMakeW(ByVal dtfecha As Date, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal fechaf As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal grupo As Integer, ByVal noesc As Long, ByVal htiempo As Long, ByRef mtm() As Double)
    Dim txtfecha As String
    Dim txtfechaf As String
    Dim txtfechar As String
    Dim txtfiltro As String
    Dim txtcadena As String
    Dim txttexto As String
    Dim i As Long
    Dim largo As Long
    Dim largoseg As Long
    Dim noseg As Long
    Dim residuo As Long
    
        txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtfechaf = "TO_DATE('" & Format$(fechaf, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtfechar = "TO_DATE('" & Format$(fechareg, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtcadena = ""
        For i = 1 To noesc - 1
            txtcadena = txtcadena & mtm(i) & ","
        Next i
            txtcadena = txtcadena & mtm(noesc)
            RegResMakeW.AddNew
            RegResMakeW.Fields("FECHA") = CLng(dtfecha)
            RegResMakeW.Fields("TIPOPOS") = tipopos                  'tipo de posicion
            RegResMakeW.Fields("FECHAREG") = CLng(fechareg)          'fecha de registro
            RegResMakeW.Fields("NOMPOS") = txtnompos                 'NOMBRE DE LA POSICION
            RegResMakeW.Fields("HORAREG") = horareg                  'HORAREG
            RegResMakeW.Fields("CPOSICION") = cposicion              'clave de la posicion
            RegResMakeW.Fields("COPERACION") = coperacion            'clave de operacion
            RegResMakeW.Fields("FECHA_F") = CLng(fechaf)             'fecha futura de valuacion
            RegResMakeW.Fields("FECHA1") = CLng(fecha1)              'fecha futura de valuacion
            RegResMakeW.Fields("FECHA2") = CLng(fecha2)              'fecha futura de valuacion
            RegResMakeW.Fields("GRUPO") = grupo                      'fecha futura de valuacion
            RegResMakeW.Fields("NO_ESC") = noesc                     'numero de escenarios
            RegResMakeW.Fields("H_TIEMPO") = htiempo                 'horizonte de tiempo
            RegResMakeW.Fields("USUARIO") = NomUsuario               'nombre del usuario
            Call GuardarElementoClob(txtcadena, RegResMakeW, "VECTORPYG")
            RegResMakeW.Update

End Sub

Function DetParamContrapCVA(ByVal fecha As Date, ByVal clave As Long, ByVal indice As Integer) As Variant
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim j As Long
Dim contar As Integer

noreg = UBound(MatTresholdContrap, 1)
ReDim mata(1 To 7, 1 To 1) As Variant
contar = 0
For i = 1 To 7
mata(i, 1) = 0
Next i
For i = 1 To noreg
    If clave = MatTresholdContrap(i, 1) Then
       contar = contar + 1
       ReDim Preserve mata(1 To 7, 1 To contar) As Variant
       For j = 1 To 7
          mata(j, contar) = MatTresholdContrap(i, j)
       Next j
    End If
Next i
mata = MTranV(mata)
mata = RutinaOrden(mata, 3, SRutOrden)
noreg1 = UBound(mata, 1)
If fecha < mata(1, 3) Then
   DetParamContrapCVA = 0
ElseIf fecha >= mata(noreg1, 3) Then
  DetParamContrapCVA = mata(noreg1, indice)
Else
For i = 1 To noreg1 - 1
If fecha >= mata(i, 3) And fecha < mata(i + 1, 3) Then
  DetParamContrapCVA = mata(i, indice)
  Exit For
End If
Next i
End If
End Function

Sub DeterminaParCVA(ByVal dtfecha As Date, _
                    ByVal id_contrap As Integer, _
                    ByRef escala As String, _
                    ByRef Sector As String, _
                    ByRef sector2 As String, _
                    ByRef Threshold As Double, _
                    ByRef mmtransfer As Double, _
                    ByRef califica As Integer, _
                    ByRef mtrans() As Double, _
                    ByRef mrecupera() As Variant, ByRef exito As Boolean)

    Dim i          As Integer
    Dim indice     As Integer
    Dim indcontrap As Integer
    Dim cmon       As String
    Dim mattc()    As Variant
    Dim tCambio    As Double
    
    For i = 1 To UBound(MatContrapartes, 1)
        If id_contrap = MatContrapartes(i, 1) Then
            indcontrap = i
            Exit For
        End If
    Next i
    If indcontrap <> 0 Then
       Sector = MatContrapartes(indcontrap, 6)
       sector2 = MatContrapartes(indcontrap, 7)
       cmon = DetParamContrapCVA(dtfecha, id_contrap, 7)
       If cmon = "USD" Then
          mattc = Leer1FactorR(dtfecha, dtfecha, "DOLAR PIP FIX", 0)
          tCambio = mattc(1, 2)
       Else
          tCambio = 1
       End If
       Threshold = DetParamContrapCVA(dtfecha, id_contrap, 4) * tCambio
       mmtransfer = DetParamContrapCVA(dtfecha, id_contrap, 6) * tCambio
       If Sector = "F" Then
          Call DetermCalifContrapF(dtfecha, id_contrap, califica, escala)
       ElseIf Sector = "NF" Then
          Call DetermCalifContrapNF(dtfecha, id_contrap, califica, escala)
       End If
       If escala = "I" Then
          mtrans = CargarMatTrans(dtfecha, "I")
          mrecupera = CargaRecuperacion(dtfecha, PrefijoBD & TablaRecInt)
          exito = True
       ElseIf escala = "N" Then
          mtrans = CargarMatTrans(dtfecha, "N")
          mrecupera = CargaRecuperacion(dtfecha, PrefijoBD & TablaRecNacional)
          exito = True
       Else
          exito = False
       End If
      
    Else
       MsgBox "No se encontro la contraparte en la base de datos"
       exito = False
    End If
End Sub

Sub DeterminaCalifContrap(ByVal dtfecha As Date, ByVal id_contrap As Integer, ByRef califs As String, ByVal escala As String)

Dim i As Integer
Dim indcontrap As Integer
Dim califica As Integer
Dim Sector As String
Dim cmon As String
Dim mattc() As Variant
Dim tCambio As Double

For i = 1 To UBound(MatContrapartes, 1)
    If id_contrap = MatContrapartes(i, 1) Then
       indcontrap = i
       Exit For
    End If
Next i
If indcontrap <> 0 Then
   Sector = MatContrapartes(indcontrap, 6)
   If Sector = "F" Then
      Call DetermCalifContrapF(dtfecha, id_contrap, califica, escala)
   ElseIf Sector = "NF" Then
      Call DetermCalifContrapNF(dtfecha, id_contrap, califica, escala)
   End If
   califs = ConvCalNum2Str(califica)
End If
End Sub



Sub DetermCalifContrapF(ByVal fecha As Date, ByVal id_contrap As String, ByRef calif As Integer, ByRef escala As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim valor1 As Integer
Dim valor2 As Integer
Dim valor3 As Integer
Dim valor4 As Integer
Dim valors1 As String
Dim valors2 As String
Dim valors3 As String
Dim valors4 As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaCalifContrapF & " WHERE ID_CONTRAP = " & id_contrap
txtfiltro2 = txtfiltro2 & " AND FECHA IN ("
txtfiltro2 = txtfiltro2 & " SELECT MAX(FECHA) FROM " & PrefijoBD & TablaCalifContrapF & " WHERE "
txtfiltro2 = txtfiltro2 & " ID_CONTRAP = " & id_contrap & " AND FECHA <= " & txtfecha & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   escala = rmesa.Fields(3)
   valors1 = rmesa.Fields(4)
   valors2 = rmesa.Fields(5)
   valors3 = rmesa.Fields(6)
   valors4 = rmesa.Fields(7)
   rmesa.Close
   If escala = "I" Then
      valor1 = ConvCalSrt2Num(TradCalifEscSPG(valors1))
      valor2 = ConvCalSrt2Num(TradCalifEscFitchG(valors2))
      valor3 = ConvCalSrt2Num(TradCalifEscMdyG(valors3))
      valor4 = ConvCalSrt2Num(TradCalifEscHRG(valors4))
   Else
      valor1 = ConvCalSrt2Num(TradCalifEscSP(valors1))
      valor2 = ConvCalSrt2Num(TradCalifEscFitch(valors2))
      valor3 = ConvCalSrt2Num(TradCalifEscMdy(valors3))
      valor4 = ConvCalSrt2Num(TradCalifEscHR(valors4))
   End If
   calif = DefinEscMin(valor1, valor2, valor3, valor4, 0)
Else
  escala = ""
  calif = 0
End If

End Sub

Sub DetermCalifContrapNF(ByVal fecha As Date, ByVal id_contrap As String, ByRef calif As Integer, ByRef escala As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaCalifContrapNF & " WHERE ID_CONTRAP = " & id_contrap
txtfiltro2 = txtfiltro2 & " AND FECHA IN ("
txtfiltro2 = txtfiltro2 & " SELECT MAX(FECHA) FROM " & PrefijoBD & TablaCalifContrapNF & " WHERE "
txtfiltro2 = txtfiltro2 & " ID_CONTRAP = " & id_contrap & " AND FECHA <= " & txtfecha & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   escala = rmesa.Fields(3)
   calif = ConvCalSrt2Num(rmesa.Fields(4))
   rmesa.Close
Else
   escala = ""
   calif = 0
End If


End Sub

Sub DetermCalifContrapEm(ByVal fecha As Date, ByVal id_contrap As String, ByRef calif As Integer, ByRef escala As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaCalifContrapEmision & " WHERE ID_CONTRAP = " & id_contrap
txtfiltro2 = txtfiltro2 & " AND FECHA IN ("
txtfiltro2 = txtfiltro2 & " SELECT MAX(FECHA) FROM " & PrefijoBD & TablaCalifContrapEmision & " WHERE "
txtfiltro2 = txtfiltro2 & " ID_CONTRAP = " & id_contrap & " AND FECHA <= " & txtfecha & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   escala = rmesa.Fields(3)
   calif = ConvCalSrt2Num(rmesa.Fields(4))
   rmesa.Close
Else
   escala = ""
   calif = 0
End If


End Sub


Function ConvCalSrt2Num(ByVal cal As String)
Dim i As Integer
ConvCalSrt2Num = 0
For i = 1 To UBound(MatCalificaciones, 1)
If cal = MatCalificaciones(i, 2) Then
   ConvCalSrt2Num = MatCalificaciones(i, 1)
   Exit Function
End If
Next i
End Function

Function ConvCalNum2Str(ByVal num As Integer)
Dim i As Integer
For i = 1 To UBound(MatCalificaciones, 1)
If num = MatCalificaciones(i, 1) Then
   ConvCalNum2Str = MatCalificaciones(i, 2)
   Exit Function
End If
Next i
ConvCalNum2Str = "ND"
End Function

Sub CalcCVAEmMD(ByVal dtfecha As Date, _
                ByVal txtport As String, _
                ByVal txtsubport As String, _
                ByVal c_emision As String, _
                ByVal emision As String, _
                ByVal noesc As Integer, _
                ByVal htiempo As Integer, _
                ByRef valcva As Double, _
                ByRef califica As Integer, _
                ByRef est_calif As Integer, _
                ByRef escala As String, _
                ByRef Sector As String, _
                ByRef txtmsg As String, _
                ByRef exito As Boolean)
    
    Dim matv() As Double
    Dim txtportfr As String
    Dim dvx As Integer
    Dim nperiodo As Integer
    Dim vrecupera As Double
    Dim mtrans() As Double
    Dim valmaxesc As Double
    Dim fechaven As Date
    Dim dxv As Long
    Dim i As Long
    Dim matprob() As Double
    Dim id_contrap As Integer
    Dim mmtransfer As Double
    Dim mrecupera() As Variant
    Dim exito1 As Boolean
    Dim tv As String
    Dim calif_sp As String
    Dim calif_fitch As String
    Dim calif_mdy As String
    Dim calif_hr As String
   
    Dim indice As Long
    txtportfr = "Normal"
    If dtfecha <> FechaVPrecios Or EsArrayVacio(MatVPreciosT) Then
       MatVPreciosT = LeerPVPrecios(dtfecha)
       FechaVPrecios = dtfecha
    End If
       indice = BuscarValorArray(c_emision, MatVPreciosT, 6)
       fechaven = 0
       If indice <> 0 Then
          fechaven = MatVPreciosT(indice, 8)
          tv = MatVPreciosT(indice, 15)
          calif_sp = MatVPreciosT(indice, 10)
          calif_fitch = MatVPreciosT(indice, 11)
          calif_mdy = MatVPreciosT(indice, 12)
          calif_hr = MatVPreciosT(indice, 13)
       End If
       matv = LeerPyGHistSubport(dtfecha, dtfecha, dtfecha, txtport, txtsubport, txtportfr, noesc, htiempo, 1)
       If UBound(matv, 1) > 0 Then
          valmaxesc = matv(1, 1)
          For i = 1 To UBound(matv, 1)
              valmaxesc = Maximo(valmaxesc, matv(i, 1))
          Next i
          'valmaxesc = Maximo(valmaxesc, 0)
          If valmaxesc < 0 Then MsgBox "es posicion corta"
          escala = ""
          Sector = ""
          Call EscalaySector(dtfecha, tv, emision, escala, Sector)
          If escala = "I" Then
             mrecupera = CargaRecuperacion(dtfecha, PrefijoBD & TablaRecInt)
             mtrans = CargarMatTrans(dtfecha, "I")
          Else
             mrecupera = CargaRecuperacion(dtfecha, PrefijoBD & TablaRecNacional)
             mtrans = CargarMatTrans(dtfecha, "N")
          End If
          califica = DetCalificacionLPMD(calif_sp, calif_fitch, calif_mdy, calif_hr) + est_calif
          vrecupera = Recuperacion(califica, mrecupera, Sector)
          Dim matresul(1 To 6) As Variant
          dxv = fechaven - dtfecha
          nperiodo = Int(dxv / 90) + 1
          matprob = CalcProbDefault1(califica, mtrans, nperiodo)
          matresul(1) = dxv                                          'dxv
          matresul(2) = valmaxesc                                    'ESCENARIO
          matresul(3) = matprob(nperiodo)                            'prob
          matresul(4) = vrecupera                                    'recuperacion
          valcva = valmaxesc * matprob(nperiodo) * (1 - vrecupera)
       End If

End Sub

Function DetCalificacionLPMD(ByVal calif_sp As String, calif_ft As String, ByVal calif_mdy As String, ByVal calif_hr As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim calif1  As Integer
Dim calif2  As Integer
Dim calif3  As Integer
Dim calif4  As Integer
Dim mata() As Variant
Dim i As Integer
calif1 = 0
calif2 = 0
calif3 = 0
calif4 = 0
   mata = EquivalCalCortoLargo()
   For i = 1 To UBound(mata, 1)
    If mata(i, 3) = calif_sp Then
       calif1 = mata(i, 2)
    End If
    If mata(i, 3) = calif_ft Then
       calif2 = mata(i, 2)
    End If
    If mata(i, 3) = calif_mdy Then
       calif3 = mata(i, 2)
    End If
    If mata(i, 3) = calif_hr Then
       calif4 = mata(i, 2)
    End If

   Next i
DetCalificacionLPMD = Maximo(calif1, Maximo(calif2, Maximo(calif3, calif4)))

End Function

Function EquivalCalCortoLargo()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer
Dim i As Integer
Dim j As Integer

txtfiltro2 = "select * from " & PrefijoBD & TablaEscCortoLargo
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       For j = 1 To 6
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
EquivalCalCortoLargo = mata
End Function

Sub EscalaySector(ByVal fecha As Date, ByVal tv As String, ByVal emision As String, ByRef escala As String, ByRef Sector As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset
Dim matr() As Variant
Dim exito3 As Boolean

matr = LeerTablaSectorEscEm(fecha)
escala = ""
Sector = ""
Call AnexarSectorEsc(fecha, tv, emision, escala, Sector, matr, exito3)
If EsVariableVacia(escala) Or EsVariableVacia(Sector) Then
 MsgBox "no se encontro la emision " & emision & " en la tabla "
End If
End Sub


Sub CalcCVAMD(ByVal dtfecha As Date, _
              ByVal id_contrap As Integer, _
              ByVal op_pos As Integer, _
              ByVal noesc As Integer, _
              ByVal htiempo As Integer, _
              ByVal inc_cal As Integer, _
              ByVal txtestres As String, _
              ByRef txtmsg As String, _
              ByRef exito As Boolean)
    
    Dim dtfecha1 As Date
    Dim escala         As String
    Dim Sector         As String
    Dim sector2         As String
    Dim Threshold      As Double
    Dim mmtransfer     As Double
    Dim califica       As Integer
    Dim mtrans()       As Double
    Dim mrecupera()    As Variant
    Dim matx()         As Variant
    Dim m_f_riesgo()         As Double
    Dim matfechas() As Date
    Dim matrends() As Double
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim MatFactR1()    As Double
    Dim parval       As New ParamValPos
    Dim MatVPrecios()  As Variant
    Dim i As Integer
    Dim noreg As Integer
    Dim noregfr As Integer
    Dim ll As Integer
    Dim MPrecio() As resValIns
    Dim matprob() As Double
    Dim suma As Double
    Dim txtfecha As String
    Dim txtborra As String
    
    Dim txtinserta As String
    Dim indice1 As Integer
    Dim dxv As Integer
    Dim nperiodo As Integer
    Dim matfshist() As Double
    Dim vrecupera As Double
    Dim MatFactoresR() As Double
    Dim mrvalflujo() As resValFlujo
    Dim mattxt() As String
    Dim exito1 As Boolean
    Dim exito2 As Boolean
    Dim exito3 As Boolean
    Dim txtfiltro As String
    Dim noreg2 As Long
    Dim txtport As String
    Dim txtmsg2 As String
    Dim txtpos As String
    Dim matb() As Integer

    Dim txtmsg0 As String
    Dim txtmsg3 As String
    Dim siesfv As Boolean
    Dim rmesa As New ADODB.recordset
   
    siesfv = EsFechaVaR(dtfecha)
    dtfecha1 = DetFechaFNoEsc(dtfecha, noesc + htiempo)
    Call VerifCargaFR(dtfecha1, dtfecha)
    ValExacta = False
    If op_pos = 1 Then
       txtport = "Em x contrap " & id_contrap & " MD"
       txtpos = "MD"
    ElseIf op_pos = 2 Then
       txtport = "Em x contrap " & id_contrap & " PIDV"
       txtpos = "PIDV"
    ElseIf op_pos = 3 Then
       txtport = "Em x contrap " & id_contrap & " PICV"
       txtpos = "PICV"
    End If
    txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro = "SELECT COUNT(*) FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
    txtfiltro = txtfiltro & " AND PORTAFOLIO = '" & txtport & "'"
    rmesa.Open txtfiltro, ConAdo
    noreg2 = rmesa.Fields(0).value
    rmesa.Close
    If noreg2 = 0 Then
       txtmsg = "No hay registros en la posicion"
       exito = False
       Exit Sub
    End If
    mattxt = CrearFiltroPosPort(dtfecha, txtport)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
    If UBound(matpos, 1) > 0 Then
        Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
        If exito2 Then
           MatVPrecios = LeerPVPrecios(dtfecha)
           Call DeterminaParCVA(dtfecha, id_contrap, escala, Sector, sector2, Threshold, mmtransfer, califica, mtrans, mrecupera, exito1)
           If exito1 Then
              califica = califica + inc_cal
              Set parval = DeterminaPerfilVal("HISTORICO")
              Call CompletarPosMesaD(matposmd, MatVPrecios)
              indice1 = BuscarValorArray(dtfecha, MatFactRiesgo, 1)
              matx = ExtraerSMatFR(indice1, noesc + htiempo, MatFactRiesgo, True, SiFactorRiesgo)
              m_f_riesgo = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
              matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 2, UBound(matx, 1)))
              Call GenRends3(m_f_riesgo, htiempo, matfechas, matrends, matb)
           'se extrae la submatriz sin la columna de dtfecha
              MatFactoresR = ExtVecFactRiesgo(indice1, MatFactRiesgo)
           'Se carga la estructura de tasas para ese día de la matriz vector tasas
              noreg = UBound(matpos, 1)
              noregfr = UBound(MatFactRiesgo, 1)
              ReDim matescmax(1 To noreg, 1 To 2) As Variant
              For i = 1 To noreg
                  matescmax(i, 2) = matposmd(i).fVencMD
              Next i
              ReDim mattxt(1 To noreg) As String
              For i = 1 To noesc
                  MatFactR1 = GenEscHist2(MatFactoresR, matrends, matb, i)
                  MPrecio = CalcValuacion(dtfecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
                  For ll = 1 To noreg
                      matescmax(ll, 1) = Maximo(matescmax(ll, 1), MPrecio(ll).mtm_sucio)
                      mattxt(ll) = mattxt(ll) & "," & MPrecio(ll).mtm_sucio
                  Next ll
                  AvanceProc = i / noesc
                  MensajeProc = "Generando valuaciones máximas " & id_contrap & " fecha de valuación " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
                  DoEvents
              Next i
              vrecupera = Recuperacion(califica, mrecupera, Sector)
              ReDim matresul(1 To noreg, 1 To 6) As Variant
              For i = 1 To noreg
                  dxv = matescmax(i, 2) - dtfecha
                  nperiodo = Int(dxv / 90) + 1
                  matprob = CalcProbDefault1(califica, mtrans, nperiodo)
                  suma = suma + matescmax(i, 1) * matprob(nperiodo) * (1 - vrecupera)
                  matresul(i, 1) = dxv                                         'dxv
                  matresul(i, 2) = matescmax(i, 1)                             'ESCENARIO
                  matresul(i, 3) = matprob(nperiodo)                           'prob
                  matresul(i, 4) = vrecupera                                   'recuperacion
                  matresul(i, 5) = 1                                           'fd
                  matresul(i, 6) = matescmax(i, 1) * matprob(nperiodo) * (1 - vrecupera)
              Next i
      
              txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
              For i = 1 To noreg
                  txtborra = "DELETE FROM " & TablaPYGCVAMD & " WHERE FECHA  = " & txtfecha
                  txtborra = txtborra & " AND CPOSICION = '" & matpos(i).C_Posicion & "' AND COPERACION = '" & matpos(i).c_operacion & "'"
                  RegResCVAMD.AddNew
                  RegResCVAMD.Fields("FECHA").value = CLng(dtfecha)
                  RegResCVAMD.Fields("TIPOPOS").value = matpos(i).tipopos
                  RegResCVAMD.Fields("FECHAREG").value = matpos(i).fechareg
                  RegResCVAMD.Fields("NOMPOS").value = matpos(i).nompos
                  RegResCVAMD.Fields("HORAREG").value = matpos(i).HoraRegOp
                  RegResCVAMD.Fields("CPOSICION").value = matpos(i).C_Posicion
                  RegResCVAMD.Fields("COPERACION").value = matpos(i).c_operacion
                  RegResCVAMD.Fields("NO_ESC").value = noesc
                  RegResCVAMD.Fields("H_TIEMPO").value = htiempo
                  Call GuardarElementoClob(mattxt(i), RegResCVAMD, "VECTORPYG")
                  RegResCVAMD.Update
              Next i
              
              
              txtborra = "DELETE FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " AND ID_CONTRAP = '" & id_contrap & "'"
              txtborra = txtborra & " AND CPOSICION = '" & txtpos & "' AND ID_CALCULO = '" & txtestres & "'"
              ConAdo.Execute txtborra
              For i = 1 To noreg
                  txtinserta = "INSERT INTO " & TablaResCVA & " VALUES("
                  txtinserta = txtinserta & txtfecha & ","
                  txtinserta = txtinserta & "'" & txtestres & "',"                   'tipo de calculo de cva
                  txtinserta = txtinserta & "'" & id_contrap & "',"                  'clave de contraparte
                  txtinserta = txtinserta & "'" & txtpos & "',"                      'POSICION
                  txtinserta = txtinserta & matresul(i, 1) & ","                     'DXV
                  txtinserta = txtinserta & califica & ","                           'calificacion
                  txtinserta = txtinserta & "'" & escala & "',"                      'escala
                  txtinserta = txtinserta & ReemplazaVacioValor(matresul(i, 2), 0) & ","        'valor del escenario maximo
                  txtinserta = txtinserta & matresul(i, 3) & ","                                'probabilidad
                  txtinserta = txtinserta & matresul(i, 4) & ","                                'recuperacion
                  txtinserta = txtinserta & matresul(i, 5) & ","                                'fd
                  txtinserta = txtinserta & matresul(i, 6) & ")"                                'vp
                  ConAdo.Execute txtinserta
              Next i
              txtmsg = "Proceso finalizado correctamente"
              exito = True
          Else
             txtmsg = "una calificacion no esta actualizada"
             exito = False
        End If
        Else
           exito = False
       End If
    Else
        txtmsg = "No hay registros en la posicion"
        exito = False
    End If

End Sub

Function Recuperacion(ByVal califica As Integer, ByRef mrecupera() As Variant, ByVal Sector As String) As Double
Dim ncolumn As Integer
Dim i As Long
If Sector = "IF" Then               'instuciones financieras
   ncolumn = 1
ElseIf Sector = "EyM C/PART" Then   'estados, municipios y organismos con participaciones
   ncolumn = 2
ElseIf Sector = "EyM S/PART" Then   'estados, municipios y organismos con fuentes de pago dif a participaciones
   ncolumn = 3
ElseIf Sector = "FPP" Then          'proyectos con fuente de pago propia
   ncolumn = 4
ElseIf Sector = "EMP" Then          'empresas y contratistas
   ncolumn = 5
Else
 MsgBox "No es un sector valido " & Sector
End If
For i = 1 To UBound(mrecupera, 1)
    If mrecupera(i, 1) = califica Then
       Recuperacion = mrecupera(i, ncolumn + 1)
       Exit Function
     End If
Next i
End Function

Function LeerResCVA(ByVal dtfecha As Date, ByVal id_contrap As String, ByVal cpos As String, ByVal txtestres As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro2 = "SELECT * FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " and ID_CONTRAP = '" & id_contrap & "' AND CPOSICION = '" & cpos & "' AND ID_CALCULO = '" & txtestres & "' ORDER BY DXV"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        ReDim mata(1 To noreg, 1 To 6) As Variant
        rmesa.Open txtfiltro2, ConAdo
        For i = 1 To noreg
            mata(i, 1) = rmesa.Fields("DXV")           'dxv
            mata(i, 2) = rmesa.Fields("SUMAVP")        'suma vp
            mata(i, 3) = rmesa.Fields("PROBAB")        'probab
            mata(i, 4) = rmesa.Fields("RECUPERA")      'RECUPERACION
            mata(i, 5) = rmesa.Fields("FDESC")         'FACTOR descuento
            mata(i, 6) = rmesa.Fields("VPTOTAL")       'vp total
            rmesa.MoveNext
        Next i

        rmesa.Close
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If
    LeerResCVA = mata
End Function

Function LeerResCVA2(ByVal dtfecha As Date, ByVal id_contrap As String, ByVal txtestres As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro2 = "SELECT * FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " and ID_CONTRAP = '" & id_contrap & "' AND ID_CALCULO = '" & txtestres & "' ORDER BY DXV"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        ReDim mata(1 To noreg, 1 To 6) As Variant
        rmesa.Open txtfiltro2, ConAdo
        For i = 1 To noreg
            mata(i, 1) = rmesa.Fields("DXV")        'dxv
            mata(i, 2) = rmesa.Fields("SUMAVP")     'suma vp
            mata(i, 3) = rmesa.Fields("PROBAB")     'probab
            mata(i, 4) = rmesa.Fields("RECUPERA")   'RECUPERACION
            mata(i, 5) = rmesa.Fields("FDESC")      'fd
            mata(i, 6) = rmesa.Fields("VPTOTAL")     'vp total
            rmesa.MoveNext
        Next i
        rmesa.Close
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If
    LeerResCVA2 = mata
End Function

Function LeerResCVA3(ByVal dtfecha As Date, ByVal cpos As String, ByVal txtestres As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro2 = "SELECT * FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " AND CPOSICION = '" & cpos & "' AND ID_CALCULO = '" & txtestres & "' ORDER BY DXV"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        ReDim mata(1 To noreg, 1 To 6) As Variant
        rmesa.Open txtfiltro2, ConAdo
        For i = 1 To noreg
            mata(i, 1) = rmesa.Fields("DXV")          'dxv
            mata(i, 2) = rmesa.Fields("SUMAVP")       'suma vp
            mata(i, 3) = rmesa.Fields("PROBAB")       'probab
            mata(i, 4) = rmesa.Fields("RECUPERA")     'RECUPERACION
            mata(i, 5) = rmesa.Fields("FDESC")        'fd
            mata(i, 6) = rmesa.Fields("VPTOTAL")      'vp total
            rmesa.MoveNext
        Next i

        rmesa.Close
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If
    LeerResCVA3 = mata
End Function

Sub GenResCVA(ByVal dtfecha As Date, ByVal idcontrap As Integer, ByVal nconf As Double, ByVal incesc As Integer, ByVal txtestres As String, ByVal txtnestres As String, ByVal id_procp As Integer, ByVal opcion As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
    Dim i           As Integer, j As Integer, noreg As Integer
    Dim ll As Integer
    Dim escala      As String
    Dim Sector      As String
    Dim sector2      As String
    Dim Threshold   As Double
    Dim mmtransfer  As Double
    Dim califica    As Integer
    Dim mtrans()    As Double
    Dim mrecupera() As Variant
    Dim suma        As Double
    'Dim suma2        As Double
    Dim contar      As Long
    'Dim contar2      As Integer
    Dim noesc       As Integer
    Dim tinterpol   As Integer
    Dim matpl()     As Variant
    Dim matff()     As Date
    Dim noff        As Integer
    Dim vrecupera   As Double
    Dim matprob()   As Double
    Dim curva1()    As propCurva

    Dim vtasa       As Double
    Dim txtfecha    As String
    Dim txtborra    As String
    Dim txtinserta  As String
    Dim sicalc As Boolean
    Dim MatFactoresR() As Double
    Dim noind As Integer
    Dim exito1 As Boolean
 
    noesc = 500
    tinterpol = 1
    noind = ValidaResPLCVAContrap(dtfecha, idcontrap, id_procp, opcion)
    If noind = 0 Then
       matpl = LeerResPLCVAContrap(dtfecha, idcontrap, noesc, 1, matff)
       If UBound(matpl, 1) > 0 Then
          Call DeterminaParCVA(dtfecha, idcontrap, escala, Sector, sector2, Threshold, mmtransfer, califica, mtrans, mrecupera, exito1)
          If exito1 Then
             noreg = UBound(matpl, 1)
             noff = UBound(matff, 1)
             ReDim matpl1(1 To noff, 1 To noesc + 1) As Double
             ReDim matpl2(1 To noff, 1 To noesc + 1) As Double
             For i = 1 To noff
                 matpl1(i, 1) = matff(i, 1)
                 matpl2(i, 1) = matff(i, 1)
             Next i
             ReDim matres(1 To noff, 1 To noesc) As Double
             ReDim matres1(1 To noff, 1 To noesc) As Double
             ReDim matres2(1 To noff, 1 To noesc) As Double
             ReDim matresul(1 To noff, 1 To 6) As Variant
             MatDerivSinLMargen = CargaDerivSinLMargen(dtfecha)
             For i = 1 To noreg
                 sicalc = DetermOperBlack(matpl(i, 1))
                 For ll = 1 To noff
                     If matpl(i, 2) = matff(ll, 1) Then
                        For j = 1 To noesc
                            If Not sicalc Then
                               matpl1(ll, j + 1) = matpl1(ll, j + 1) + matpl(i, j + 2) 'sigue la regla de trheshold y monto minimo de garantia
                            Else
                               matpl2(ll, j + 1) = matpl2(ll, j + 1) + matpl(i, j + 2)
                            End If
                       Next j
                     End If
                 Next ll
                 AvanceProc = i / noreg
                 MensajeProc = "Procesando los p&l del CVA del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
                 DoEvents
             Next i
             For i = 1 To noff
                 For j = 1 To noesc
                     If matpl1(i, j + 1) > Threshold And mmtransfer <> 0 Then
                        matres1(i, j) = matpl1(i, j + 1) - Int((matpl1(i, j + 1) - Threshold) / mmtransfer) * mmtransfer
                     Else
                        matres1(i, j) = matpl1(i, j + 1)
                     End If
                     matres2(i, j) = matpl2(i, j + 1)
                 Next j
             Next i
             califica = Minimo(califica + incesc, 18)
             vrecupera = Recuperacion(califica, mrecupera, sector2)
             matprob = CalcProbDefault(califica, mtrans, noff)
             curva1 = LeerCurvaC(dtfecha, "DESC IRS")
             For i = 1 To noff
                 suma = 0
                 contar = 0
                 'suma2 = 0
                 'contar2 = 0
                 matresul(i, 1) = matff(i, 1) - dtfecha
                 If txtestres = "CVA" Then
                     For j = 1 To noesc
                         If matres1(i, j) + matres2(i, j) > 0 Then
                            'solo se suman las marcas a mercado positivas
                            suma = suma + matres1(i, j) + matres2(i, j)
                            contar = contar + 1
                         End If
                         'If matres2(i, j) > 0 Then
                          'solo se suman las marcas a mercado positivas
                         '   suma2 = suma2 + matres2(i, j)
                         '   contar2 = contar2 + 1
                         'End If
                     Next j
                     If contar <> 0 Then suma = suma / contar
                     'If contar2 <> 0 Then suma2 = suma2 / contar2
                 Else
                     ReDim matesc(1 To noesc, 1 To 1) As Double
                     For j = 1 To noesc
                         matesc(j, 1) = matres1(i, j) + matres2(i, j)
                     Next j
                     suma = Maximo(CPercentil2(nconf, matesc, 0, 0, True), 0)
                 End If
                 vtasa = CalculaTasa(curva1, matresul(i, 1), tinterpol)
                 matresul(i, 2) = suma                                                             'el promedio de las marcas a mercado positivas
                 matresul(i, 3) = matprob(i)                                                        'prob*(1-vrecupera)
                 matresul(i, 4) = vrecupera
                 matresul(i, 5) = 1 / (1 + vtasa * matresul(i, 1) / 360)                            'FD
                 matresul(i, 6) = matresul(i, 2) * matprob(i) * (1 - vrecupera) / (1 + vtasa * matresul(i, 1) / 360) 'VTOTAL
             Next i
             txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
             txtborra = "DELETE FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " AND ID_CONTRAP = '" & idcontrap & "' AND CPOSICION = 'DER' AND ID_CALCULO = '" & txtnestres & "'"
             ConAdo.Execute txtborra
             For i = 1 To noff
                 txtinserta = "INSERT INTO " & TablaResCVA & " VALUES("
                 txtinserta = txtinserta & txtfecha & ","                     'fecha
                 txtinserta = txtinserta & "'" & txtnestres & "',"            'tipo de calculo de cva
                 txtinserta = txtinserta & "'" & idcontrap & "',"             'clave de contraparte
                 txtinserta = txtinserta & "'DER',"                           'portafolio
                 txtinserta = txtinserta & matresul(i, 1) & ","               'periodo forward
                 txtinserta = txtinserta & califica & ","                     'calificacion
                 txtinserta = txtinserta & "'" & escala & "',"                'escala
                 txtinserta = txtinserta & ReemplazaVacioValor(matresul(i, 2), 0) & ","    'promedio de las simulaciones positivas
                 txtinserta = txtinserta & matresul(i, 3) & ","               'probabilidad de default
                 txtinserta = txtinserta & matresul(i, 4) & ","               'recuperacion
                 txtinserta = txtinserta & matresul(i, 5) & ","               'factor de descuento
                 txtinserta = txtinserta & matresul(i, 6) & ")"               'valor presente
                 ConAdo.Execute txtinserta
                 AvanceProc = i / noreg
                 MensajeProc = "Guardando los resultados de CVA en la base de datos " & Format$(AvanceProc, "##0.00 %")
                 DoEvents
             Next i
             MensajeProc = "El Proceso finalizo correctamente"
             exito = True
          Else
          MensajeProc = ""
             If califica = 0 Then MensajeProc = "No hay calificacion para esta contraparte" & idcontrap
             MensajeProc = MensajeProc & " no hay datos de threshold para la contraparte " & idcontrap
             exito = False
          End If
       Else
       
          MensajeProc = "No hay datos para la contraparte " & idcontrap
          txtmsg = MensajeProc
          exito = True
       End If
    Else
        exito = False
       MensajeProc = "No se han terminado todos los subprocesos"
    End If
End Sub

Sub GenResCVAPos(ByVal dtfecha As Date, ByVal txtnompos As String, ByVal nconf As Double, ByVal incesc As Integer, ByVal txtestres As String, ByRef txtmsg As String, ByRef exito As Boolean)
    Dim i           As Integer, j As Integer, noreg As Integer
    Dim ll As Integer
    Dim escala      As String
    Dim Sector      As String
    Dim sector2      As String
    Dim Threshold   As Double
    Dim mmtransfer  As Double
    Dim califica    As Integer
    Dim mtrans()    As Double
    Dim mrecupera() As Variant
    Dim suma1        As Double
    Dim suma2        As Double
    Dim contar1      As Integer
    Dim contar2      As Integer
    Dim noesc       As Integer
    Dim tinterpol   As Integer
    Dim matpl()     As Variant
    Dim matff()     As Date
    Dim noff        As Integer
    Dim vrecupera   As Double
    Dim matprob()   As Double
    Dim curva1()    As propCurva

    Dim vtasa       As Double
    Dim txtfecha    As String
    Dim txtborra    As String
    Dim txtinserta  As String
    Dim sicalc As Boolean
    Dim MatFactoresR() As Double
    Dim noind As Integer
    Dim exito1 As Boolean
    Dim idcontrap As Integer
 
    noesc = 500
    tinterpol = 1
    'noind = 0ValidaResPLCVAContrap(dtfecha, idcontrap, 110)
    noind = 0
    If noind = 0 Then
       Call DeterminaParCVA(dtfecha, idcontrap, escala, Sector, sector2, Threshold, mmtransfer, califica, mtrans, mrecupera, exito1)
       escala = "N"
       Sector = "F"
       Threshold = 0
       mmtransfer = 0
       califica = 1
       mtrans = CargarMatTrans(dtfecha, "N")
       mrecupera = CargaRecuperacion(dtfecha, PrefijoBD & TablaRecNacional)
       exito1 = True
       If exito1 Then
       matpl = LeerResPLCVAPos(dtfecha, txtnompos, noesc, 1, matff)
       If UBound(matpl, 1) > 0 Then
          noreg = UBound(matpl, 1)
          noff = UBound(matff, 1)
       ReDim matpl1(1 To noff, 1 To noesc + 1) As Double
       ReDim matpl2(1 To noff, 1 To noesc + 1) As Double
          For i = 1 To noff
              matpl1(i, 1) = matff(i, 1)
              matpl2(i, 1) = matff(i, 1)
          Next i
          califica = Minimo(califica + incesc, 18)
          vrecupera = Recuperacion(califica, mrecupera, Sector)
          matprob = CalcProbDefault(califica, mtrans, noff)
          If dtfecha <> FechaArchCurvas Or EsArrayVacio(MatCurvasT) Then
             MatCurvasT = LeerCurvaCompleta(dtfecha, exito)
             If exito Then FechaArchCurvas = dtfecha
          End If
          curva1 = CrearCurva(dtfecha, "DESC IRS", MatCurvasT, MatFactoresR, True)
       ReDim matres(1 To noff, 1 To noesc) As Double
       ReDim matres1(1 To noff, 1 To noesc) As Double
       ReDim matres2(1 To noff, 1 To noesc) As Double
       ReDim matresul(1 To noff, 1 To 5) As Variant
       MatDerivSinLMargen = CargaDerivSinLMargen(dtfecha)
          For i = 1 To noreg
              sicalc = DetermOperBlack(matpl(i, 1))
              For ll = 1 To noff
                  If matpl(i, 2) = matff(ll, 1) Then
                     For j = 1 To noesc
                         If Not sicalc Then
                            matpl1(ll, j + 1) = matpl1(ll, j + 1) + matpl(i, j + 2) 'sigue la regla de trheshold y monto minimo de garantia
                         Else
                            matpl2(ll, j + 1) = matpl2(ll, j + 1) + matpl(i, j + 2)
                         End If
                     Next j
                  End If
              Next ll
              AvanceProc = i / noreg
              MensajeProc = "Procesando los p&l del CVA del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
              DoEvents
          Next i
          For i = 1 To noff
              For j = 1 To noesc
                  If matpl1(i, j + 1) > Threshold And mmtransfer <> 0 Then
                     matres1(i, j) = matpl1(i, j + 1) - Int((matpl1(i, j + 1) - Threshold) / mmtransfer) * mmtransfer
                  Else
                     matres1(i, j) = matpl1(i, j + 1)
                  End If
                  matres2(i, j) = matpl2(i, j + 1)
              Next j
          Next i
          For i = 1 To noff
              suma1 = 0
              contar1 = 0
              suma2 = 0
              contar2 = 0
              matresul(i, 1) = matff(i, 1) - dtfecha
              If txtestres = "CVA" Then
                  For j = 1 To noesc
                      If matres1(i, j) > 0 Then
                         'solo se suman las marcas a mercado positivas
                         suma1 = suma1 + matres1(i, j)
                         contar1 = contar1 + 1
                      End If
                      If matres2(i, j) > 0 Then
                       'solo se suman las marcas a mercado positivas
                         suma2 = suma2 + matres2(i, j)
                         contar2 = contar2 + 1
                      End If
                  Next j
                  If contar1 <> 0 Then suma1 = suma1 / contar1
                  If contar2 <> 0 Then suma2 = suma2 / contar2
              Else
                ReDim matesc1(1 To noesc, 1 To 1) As Double
                ReDim matesc2(1 To noesc, 1 To 1) As Double
                  For j = 1 To noesc
                      matesc1(j, 1) = matres1(i, j)
                      matesc2(j, 1) = matres2(i, j)
                  Next j
                  suma1 = Maximo(CPercentilCVaR(nconf, matesc1, 0, 0, True), 0)
                  suma2 = Maximo(CPercentilCVaR(nconf, matesc2, 0, 0, True), 0)
              End If
              vtasa = CalculaTasa(curva1, matresul(i, 1), tinterpol)
              matresul(i, 2) = suma1 + suma2                                                     'el promedio de las marcas a mercado positivas
              matresul(i, 3) = matprob(i)                                                        'prob
              matresul(i, 4) = vrecupera                                                         'prob
              matresul(i, 5) = 1 / (1 + vtasa * matresul(i, 1) / 360)                            'FD
              matresul(i, 6) = matresul(i, 2) * matprob(i) * (1 - vrecupera) / (1 + vtasa * matresul(i, 1) / 360) 'VTOTAL
          Next i
          txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
          txtborra = "DELETE FROM " & TablaResCVA & " WHERE FECHA  = " & txtfecha & " AND id_contrap = '" & txtnompos & "' AND CPOSICION = 'DER' AND ID_CALCULO = '" & txtestres & "'"
          ConAdo.Execute txtborra
          For i = 1 To noff
              txtinserta = "INSERT INTO " & TablaResCVA & " VALUES("
              txtinserta = txtinserta & txtfecha & ","
              txtinserta = txtinserta & "'" & txtestres & "',"                          'tipo de calculo de cva
              txtinserta = txtinserta & "'" & txtnompos & "',"                          'clave de contraparte
              txtinserta = txtinserta & "'DER',"                                        'portafolio
              txtinserta = txtinserta & matresul(i, 1) & ","                            'periodo forward
              txtinserta = txtinserta & matresul(i, 1) & ","                            'calificacion
              txtinserta = txtinserta & matresul(i, 1) & ","                            'escala
              txtinserta = txtinserta & ReemplazaVacioValor(matresul(i, 2), 0) & ","    'promedio de las simulaciones positivas
              txtinserta = txtinserta & matresul(i, 3) & ","                            'probabilidad de default
              txtinserta = txtinserta & matresul(i, 4) & ","                            'recuperacion
              txtinserta = txtinserta & matresul(i, 5) & ","                            'factor de descuento
              txtinserta = txtinserta & matresul(i, 6) & ")"                            'valor presente
              ConAdo.Execute txtinserta
              AvanceProc = i / noreg
              MensajeProc = "Guardando los resultados de CVA en la base de datos " & Format$(AvanceProc, "##0.00 %")
              DoEvents
          Next i
          MensajeProc = "El Proceso finalizo correctamente"
          exito = True
          End If
      Else
          MensajeProc = "No hay datos en el catalogo " & PrefijoBD & TablaTreshCont & " para la contraparte " & txtnompos
          txtmsg = MensajeProc
          exito = False
      End If
   Else
       exito = False
      MensajeProc = "No se han terminado todos los subprocesos"
   End If
End Sub

Function ValidaResPLCVAContrap(ByVal fecha As Date, idcontrap As Integer, idproc As Integer, ByVal opcion As Integer)
Dim txtfiltro As String
Dim txtfecha As String
Dim txtport As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtport = "Deriv Contrap " & idcontrap
txtfiltro = "SELECT COUNT(*) FROM " & DetermTablaSubproc(opcion) & " WHERE ID_SUBPROCESO = " & idproc
txtfiltro = txtfiltro & " AND FECHAP = " & txtfecha
txtfiltro = txtfiltro & " AND FINALIZADO = 'N'"
txtfiltro = txtfiltro & " AND (PARAMETRO1,PARAMETRO2,PARAMETRO3) IN"
txtfiltro = txtfiltro & " (SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPortPosicion & " "
txtfiltro = txtfiltro & " WHERE FECHA_PORT = " & txtfecha
txtfiltro = txtfiltro & " AND PORTAFOLIO =  '" & txtport & "')"
rmesa.Open txtfiltro, ConAdo
ValidaResPLCVAContrap = rmesa.Fields(0)
rmesa.Close
End Function

Function LeerResPLCVAContrap(ByVal dtfecha As Date, ByVal id_contrap As Integer, ByVal noesc As Long, ByVal htiempo As Long, ByRef matf() As Date)
   
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i          As Integer, j As Integer, noreg As Integer
    Dim valor      As String
    Dim mata()     As Variant
    Dim matc()     As String
    Dim dxv        As Integer
    Dim noff As Integer
    Dim txtport As String
    Dim rmesa As New ADODB.recordset
    'primero obtenemos las fechas futuras para las que realizo el calculo
    txtport = "Deriv Contrap " & id_contrap
    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT FECHA_F from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND NO_ESC = " & noesc
    txtfiltro1 = txtfiltro1 & " AND H_TIEMPO = " & htiempo
    txtfiltro1 = txtfiltro1 & " AND (CPOSICION,FECHAREG,COPERACION) IN ("
    txtfiltro1 = txtfiltro1 & "SELECT CPOSICION,FECHAREG,COPERACION FROM "
    txtfiltro1 = txtfiltro1 & "" & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS = 1)"
    'txtfiltro1 = txtfiltro1 & " AND COPERACION IN (SELECT COPERACION FROM " & tablaposselecc & " WHERE FECHA = " & txtfecha & " )"
    txtfiltro1 = txtfiltro1 & " GROUP BY FECHA_F ORDER BY FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noff = rmesa.Fields(0)
    rmesa.Close
    If noff <> 0 Then
       rmesa.Open txtfiltro1, ConAdo
       ReDim matf(1 To noff, 1 To 1) As Date
       For i = 1 To noff
           matf(i, 1) = rmesa.Fields(0)
           rmesa.MoveNext
       Next i
       rmesa.Close
    Else
       ReDim matf(0 To 0) As Date
    End If
    txtfiltro1 = "SELECT * from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND NO_ESC = " & noesc
    txtfiltro1 = txtfiltro1 & " AND H_TIEMPO = " & htiempo
    txtfiltro1 = txtfiltro1 & " AND (CPOSICION,FECHAREG,COPERACION) IN ("
    txtfiltro1 = txtfiltro1 & "SELECT CPOSICION,FECHAREG,COPERACION FROM "
    txtfiltro1 = txtfiltro1 & "" & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS = 1)"
    'txtfiltro1 = txtfiltro1 & " AND COPERACION IN (SELECT COPERACION FROM " & TablaPosSelecc & " WHERE FECHA = " & txtfecha & " )"
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION,FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        ReDim mata(1 To noreg, 1 To noesc + 2) As Variant
        rmesa.Open txtfiltro1, ConAdo
        For i = 1 To noreg
            mata(i, 1) = rmesa.Fields("COPERACION")     'clave de operacion
            mata(i, 2) = rmesa.Fields("FECHA_F")       'fecha forward
            If rmesa.Fields("VECTOR_PYG").ActualSize <> 0 Then
               valor = rmesa.Fields("VECTOR_PYG").GetChunk(rmesa.Fields("VECTOR_PYG").ActualSize)
               matc = EncontrarSubCadenas(valor, ",")
               For j = 1 To noesc
                   mata(i, j + 2) = CDbl(matc(j))
               Next j
            Else
            valor = ""
            End If
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo los p&l del CVA del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
            DoEvents
        Next i
        rmesa.Close
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If
LeerResPLCVAContrap = mata
End Function

Function LeerResPLCVAPos(ByVal dtfecha As Date, ByVal txtnompos As String, ByVal noesc As Long, ByVal htiempo As Long, ByRef matf() As Date)
   
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i          As Integer, j As Integer, noreg As Integer
    Dim valor      As String
    Dim mata()     As Variant
    Dim matc()     As String
    Dim dxv        As Integer
    Dim noff As Integer
    Dim txtport As String
    Dim rmesa As New ADODB.recordset
    'primero obtenemos las fechas futuras para las que realizo el calculo
    txtport = "POSICION " & txtnompos
    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT FECHA_F from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND NOMPOS = '" & txtnompos & "'"
    txtfiltro1 = txtfiltro1 & " AND NO_ESC = " & noesc
    txtfiltro1 = txtfiltro1 & " AND H_TIEMPO = " & htiempo
    txtfiltro1 = txtfiltro1 & " GROUP BY FECHA_F ORDER BY FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noff = rmesa.Fields(0)
    rmesa.Close
    If noff <> 0 Then
       rmesa.Open txtfiltro1, ConAdo
       ReDim matf(1 To noff, 1 To 1) As Date
       For i = 1 To noff
           matf(i, 1) = rmesa.Fields(0)
           rmesa.MoveNext
       Next i
       rmesa.Close
    Else
       ReDim matf(0 To 0) As Date
    End If
    txtfiltro1 = "SELECT * from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND NOMPOS = '" & txtnompos & "'"
    txtfiltro1 = txtfiltro1 & " AND NO_ESC = " & noesc
    txtfiltro1 = txtfiltro1 & " AND H_TIEMPO = " & htiempo
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION,FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        ReDim mata(1 To noreg, 1 To noesc + 2) As Variant
        rmesa.Open txtfiltro1, ConAdo
        For i = 1 To noreg
            mata(i, 1) = rmesa.Fields(3)   'clave de operacion
            mata(i, 2) = rmesa.Fields(4)   'fecha forward
            valor = rmesa.Fields("VECTOR_PYG").GetChunk(rmesa.Fields("VECTOR_PYG").ActualSize)
            matc = EncontrarSubCadenas(valor, ",")
            For j = 1 To noesc
                mata(i, j + 2) = CDbl(matc(j))
            Next j
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo los p&l del CVA del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
            DoEvents
        Next i
        rmesa.Close
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If
LeerResPLCVAPos = mata
End Function

Function DetermOperBlack(ByVal coper As String)
Dim i As Integer
DetermOperBlack = False
For i = 1 To UBound(MatDerivSinLMargen, 1)
    If coper = MatDerivSinLMargen(i, 1) Then
       DetermOperBlack = True
       Exit Function
    End If
Next i
End Function

Function RepCVA(ByVal dt_fecha1 As Date, ByVal dt_fecha2 As Date, ByVal Sector As String)
    Dim escala      As String
    Dim sector2 As String
    Dim Threshold   As Double
    Dim mmtransfer  As Double
    Dim califica    As Integer
    Dim mtrans()    As Double
    Dim mrecupera() As Variant
    Dim txtfecha As String
    Dim txtnomarch As String
    Dim txtcadena As String
    Dim i As Integer, j As Integer
    Dim id_contrap As String
    Dim vrecupera  As Double
    Dim mata() As Variant
    Dim matb() As Variant
    Dim matc() As Variant
    Dim matd() As Variant
    Dim mate() As Variant
    Dim suma1 As Double, suma2 As Double
    Dim suma3 As Double, suma4 As Double
    Dim sumat1 As Double, sumat2 As Double
    Dim exito1 As Boolean
    Dim nivel As String
    Dim contar As Integer
    Dim matv() As Double
    Dim matr() As New resCVA
    Dim sector1 As String
    ReDim matr(1 To 1)
    contar = 0
    For i = 1 To UBound(MatContrapartes, 1)
        id_contrap = MatContrapartes(i, 1)
        If (Sector = "F" And MatContrapartes(i, 6) = "F") Or (Sector = "NF" And MatContrapartes(i, 6) = "EM") Or (Sector = "NF" And MatContrapartes(i, 6) = "NF") Then
           mata = LeerResCVA(dt_fecha2, id_contrap, "DER", "CVA")
           matb = LeerResCVA(dt_fecha2, id_contrap, "MD", "CVA")
           matc = LeerResCVA(dt_fecha2, id_contrap, "PIDV", "CVA")
           matd = LeerResCVA(dt_fecha2, id_contrap, "PICV", "CVA")
           mate = LeerResCVA2(dt_fecha1, id_contrap, "CVA")
           If UBound(mata, 1) <> 0 Or UBound(matb, 1) <> 0 Or UBound(matc, 1) <> 0 Or UBound(matd, 1) <> 0 Or UBound(mate, 1) <> 0 Then
               Call DeterminaParCVA(dt_fecha2, id_contrap, escala, sector1, sector2, Threshold, mmtransfer, califica, mtrans, mrecupera, exito1)
               If escala <> "0" Then
                  contar = contar + 1
                  ReDim Preserve matr(1 To contar)
                  'vrecupera = Recuperacion(califica, mrecupera, Sector)
                  'se leen los resultados de derivados y md
                  suma1 = 0: suma2 = 0: suma3 = 0: suma4 = 0: sumat1 = 0: sumat2 = 0
                  If UBound(mata, 1) > 0 Then
                      For j = 1 To UBound(mata, 1)
                          suma1 = suma1 + Val(mata(j, 6))
                      Next j
                  End If
                  If UBound(matb, 1) > 0 Then
                     For j = 1 To UBound(matb, 1)
                         suma2 = suma2 + Val(matb(j, 6))
                     Next j
                  End If
                  If UBound(matc, 1) > 0 Then
                     For j = 1 To UBound(matc, 1)
                         suma3 = suma3 + Val(matc(j, 6))
                     Next j
                  End If
                  If UBound(matd, 1) > 0 Then
                     For j = 1 To UBound(matd, 1)
                         suma4 = suma4 + Val(matd(j, 6))
                     Next j
                  End If
                  sumat2 = suma1 + suma2 + suma3 + suma4
                  If UBound(mate, 1) > 0 Then
                     For j = 1 To UBound(mate, 1)
                         sumat1 = sumat1 + Val(mate(j, 6))
                     Next j
                  End If
                  matv = LeerResValPort(dt_fecha2, "TOTAL", "Deriv Contrap " & id_contrap, 2)
                  If UBound(mata, 1) > 0 Or UBound(matb, 1) > 0 Or UBound(matc, 1) <> 0 Or UBound(matd, 1) <> 0 Or UBound(mate, 1) <> 0 Then
                     matr(contar).descrip = MatContrapartes(i, 3)
                     matr(contar).escala = escala
                     matr(contar).calif = califica
                     If UBound(matv, 1) <> 0 Then
                        matr(contar).mtm = matv(1) / 1000000
                     Else
                        matr(contar).mtm = 0
                     End If
                     matr(contar).cvaderiv = suma1 / 1000000
                     matr(contar).cvamd = suma2 / 1000000
                     matr(contar).cvapidv = suma3 / 1000000
                     matr(contar).cvapicv = suma4 / 1000000
                     matr(contar).cvatotal = sumat2 / 1000000
                     matr(contar).cvat_1 = sumat1 / 1000000
                  End If
               End If
           End If
        End If
    Next i
    RepCVA = matr
End Function

Function RepCVA2(ByVal dt_fecha As Date)
    Dim txtfecha As String
    Dim txtcadena As String
    Dim i As Integer, j As Integer, l As Integer
    Dim mata() As String
    Dim matb() As String
    Dim matr() As Variant
    Dim suma As Double
    
    ReDim mata(1 To 4, 1 To 3) As String
    ReDim matb(1 To 4, 1 To 3) As String
    ReDim mats(1 To 5, 1 To 4) As Variant
    mats(1, 1) = "Derivados"
    mats(2, 1) = "MD"
    mats(3, 1) = "PIDV"
    mats(4, 1) = "PICV"
    mats(5, 1) = "Total"
    
    mata(1, 1) = "DER": matb(1, 1) = "CVA": mata(1, 2) = "DER": matb(1, 2) = "Estr1": mata(1, 3) = "DER": matb(1, 3) = "Estr2"
    mata(2, 1) = "MD": matb(2, 1) = "CVA": mata(2, 2) = "MD": matb(2, 2) = "Estr1": mata(2, 3) = "MD": matb(2, 3) = "Estr2"
    mata(3, 1) = "PIDV": matb(3, 1) = "CVA": mata(3, 2) = "PIDV": matb(3, 2) = "Estr1": mata(3, 3) = "PIDV": matb(3, 3) = "Estr2"
    mata(4, 1) = "PICV": matb(4, 1) = "CVA": mata(4, 2) = "PICV": matb(4, 2) = "Estr1": mata(4, 3) = "PICV": matb(4, 3) = "Estr2"
    For i = 1 To 4
        For j = 1 To 3
            matr = LeerResCVA3(dt_fecha, mata(i, j), matb(i, j))
            If UBound(matr, 1) <> 0 Then
               suma = 0
               For l = 1 To UBound(matr, 1)
                   suma = suma + Val(matr(l, 6))
               Next l
               mats(i, j + 1) = suma / 1000000
            End If
            mats(5, j + 1) = mats(5, j + 1) + mats(i, j + 1)
        Next j
        
    Next i
    RepCVA2 = mats
End Function


Function ExtVecFactRiesgo(ByVal ind As Integer, ByRef mata() As Variant) As Double()

    'se cargan las tasas a una matriz pivote
    'sin la dtfecha
    Dim SiValoresCero As Boolean
    Dim j As Integer, n As Integer
    n = UBound(mata, 2)
    If n > 1 Then
        ReDim mt(1 To n - 1, 1 To 1) As Double
        For j = 1 To NoFactores
            If mata(ind, j + 1) = 0 Or IsNull(mata(ind, j + 1)) Then
               mt(j, 1) = 0
               SiValoresCero = True
            Else
               mt(j, 1) = CDbl(mata(ind, j + 1))
            End If
        Next j

        ExtVecFactRiesgo = mt
    End If
End Function

Sub GenSubpConsolCVA(ByVal id_proc As Integer, ByVal contar As Long, ByVal fecha As Date, ByVal id_contrap As Integer, ByVal nconf As Double, ByVal id_estres As Integer, ByVal t_cva As String, ByVal txtnestres As String, ByVal id_tabla As Integer)
    Dim txtfecha  As String
    Dim txtcadena As String
    txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consolida resultados CVA", id_contrap, nconf, id_estres, t_cva, txtnestres, "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
End Sub

Function LeerContrapFecha(ByVal fecha As Date)
Dim cmdSp As New ADODB.Command
Dim recordset As New ADODB.recordset
Dim contar As Integer
   With cmdSp
        .ActiveConnection = ConAdo
        .CommandType = adCmdStoredProc
        .CommandText = "OBTENERCONTRAPFECHA"
   End With
   cmdSp.Parameters.Append cmdSp.CreateParameter("FECHAX", adDBDate, adParamInput, , Format(fecha, "dd/mm/yyyy"))
   Set recordset = cmdSp.Execute
   contar = 0
   ReDim mata(1 To 1) As Integer
   While (Not recordset.EOF)
       contar = contar + 1
       ReDim Preserve mata(1 To contar) As Integer
       mata(contar) = recordset.Fields(0)
       recordset.MoveNext
   Wend
   recordset.Close
   Set cmdSp = Nothing
LeerContrapFecha = mata
End Function

