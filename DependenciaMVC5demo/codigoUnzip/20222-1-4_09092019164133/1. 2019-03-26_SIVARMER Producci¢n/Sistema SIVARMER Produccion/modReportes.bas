Attribute VB_Name = "modReportes"
Option Explicit

Sub GenRepValDer(ByVal fecha As Date, ByVal opc As Integer)
Dim matres() As Variant
Dim nomarch3 As String
Dim txttitulo As String
Dim i As Long
Dim j As Long
Dim contar As Long
Dim txtcadena As String
Dim exitoarch As Boolean

matres = LeerResValDeriv(fecha, txtportCalc1, opc, contar)
MsgBox "Hay " & contar & " diferencias de valuación"
If UBound(matres, 1) <> 0 Then
   nomarch3 = DirResVaR & "\Reporte Valuación derivados " & opc & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch3
   frmCalVar.CommonDialog1.ShowSave
   nomarch3 = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch3, 1, exitoarch)
   If exitoarch Then
   Print #1, "Operaciones al " & Format(fecha, "dd/mm/yyyy")
   txttitulo = "Clave de operación" & Chr(9) & "Pata activa SIVARMER" & Chr(9) & "Pata pasiva SIVARMER" & Chr(9) & "MTM SIVARMER"
   txttitulo = txttitulo & Chr(9) & "Pata activa IKOS" & Chr(9) & "Pata pasiva IKOS" & Chr(9) & "MTM IKOS" & Chr(9) & "Dif Pata activa" & Chr(9) & "Dif Pata Pasiva" & Chr(9) & "Dif MTM"
   txttitulo = txttitulo & Chr(9) & "Diferencia %" & Chr(9) & "Fecha de inicio" & Chr(9) & "Fecha de vencimiento" & Chr(9) & "Intencion" & Chr(9) & "Producto" & Chr(9) & "Tipo swap" & Chr(9) & "Contraparte"
   Print #1, txttitulo
   For i = 1 To UBound(matres, 1)
       txtcadena = ""
       For j = 1 To UBound(matres, 2)
           txtcadena = txtcadena & matres(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se genero el archivo " & nomarch3
   End If
Else
 MsgBox "No hay datos de valuación de derivados para esta fecha"
End If

End Sub

Sub GenRepDurPos(ByVal fecha As Date)
Dim matres() As Variant
Dim nomarch3 As String
Dim i As Integer
Dim jj As Integer
Dim txttitulo As String
Dim txtcadena As String
Dim exitoarch As Boolean

matres = LeerDuracionPos(fecha)
If UBound(matres, 1) > 1 Then
   nomarch3 = DirResVaR & "\Duracion posicion " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch3
   frmCalVar.CommonDialog1.ShowSave
   nomarch3 = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch3, 1, exitoarch)
   If exitoarch Then
   Print #1, "Operaciones al " & Format(fecha, "dd/mm/yyyy")
   txttitulo = "Clave posicion" & Chr(9) & "Clave operacion" & Chr(9) & "Tipo operacion" & Chr(9) & "Duracion activa" & Chr(9) & "Duracion pasiva"
   Print #1, txttitulo
   For i = 1 To UBound(matres, 1)
       txtcadena = ""
       For jj = 1 To 5
           txtcadena = txtcadena & matres(i, jj) & Chr(9)
       Next jj
       Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se genero el archivo " & nomarch3
   End If
Else
  MsgBox "No hay datos de duracion para esta fecha"
End If
End Sub

Function LeerAnalisisFR(ByVal fecha As Date, ByVal nesc As Integer) As Variant()
Dim txtfecha As String
Dim txtcadena As String
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = "" & TablaAnalisisFRO & " WHERE FECHA = " & txtfecha & "  AND NDATOS = " & nesc & " ORDER BY CONCEPTO, PLAZO"
txtfiltro1 = "select count(*) from " & txtcadena
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 7) As Variant
   txtfiltro2 = "select * from " & txtcadena
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("FECHA")
       mata(i, 2) = rmesa.Fields("CONCEPTO")
       mata(i, 3) = rmesa.Fields("PLAZO")
       mata(i, 4) = rmesa.Fields("VALORT1")
       mata(i, 5) = rmesa.Fields("VALORT")
       mata(i, 6) = rmesa.Fields("INCOBSERVADO")
       mata(i, 7) = rmesa.Fields("VOLATILIDAD")    'volatilidad del factor
       rmesa.MoveNext
       MensajeProc = "Leyendo los factores de riesgo " & Format(AvanceProc, "#,##0.00 %")
   Next i
   rmesa.Close
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerAnalisisFR = mata
End Function

Function LeerEfecRetro(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & TablaEficRetro & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(3)          'clave de operacion
       mata(i, 2) = rmesa.Fields(4)          'tipo de operacino
       mata(i, 3) = rmesa.Fields(5)          'fecha de vencimiento
       mata(i, 5) = Format(rmesa.Fields(7) / 100, "##0.00 %")       'valor efectividad
       mata(i, 6) = rmesa.Fields(6)          'tipo de efectividad
       mata(i, 4) = Val(mata(i, 3) - fecha)  'dias de vencimiento
      rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If

LeerEfecRetro = mata
End Function

Sub GeneraFlujosValSwaps(ByVal fecha As Date)
Dim bl_exito As Boolean
Dim txtport As String
Dim fecha1 As Date
Dim tipopos As Integer
Dim nomarch1 As String
Dim txtcadena As String
Dim i As Long
Dim exito As Boolean
Dim exitoarch As Boolean

'rutina que genera los flujos de la posicion de swaps sin descontar
SiAnexarFlujosSwaps = True
txtport = "SWAPS"
fecha1 = fecha - 10
'primero se procede a leer la interfase de
ValExacta = True
Call CrearMatFRiesgo2(fecha1, fecha, MatFactRiesgo, "", exito)
tipopos = 1
Call CalculaValPos(fecha, fecha, fecha, txtport, 2, bl_exito)  'procesando la informacion de la fecha
'se procede a actualizar una tabla con estos valores y despues se realiza un filtro
If bl_exito Then
   nomarch1 = DirResVaR & "\Flujos swaps valuados " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch1
   frmCalVar.CommonDialog1.ShowSave
   nomarch1 = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch1, 1, exitoarch)
   If exitoarch Then
   Print #1, "Flujos al " & fecha
   txtcadena = "Clave de operación" & Chr(9) & "Pata" & Chr(9) & "Inicio del flujo" & Chr(9) & "Final del flujo" & Chr(9) & "Fecha de descuento" & Chr(9) & "Saldo" & Chr(9) & "Amortizacion" & Chr(9) & "Intereses generados per" & Chr(9) & "Intereses acumulados per" & Chr(9) & "Intereses Pagados per" & Chr(9) & "Intereses acumulados sig per" & Chr(9) & "Pago total" & Chr(9) & "Tasa descuento" & Chr(9) & "Factor descuento" & Chr(9) & "VP"
   Print #1, txtcadena
   For i = 2 To UBound(MatValFlujosD, 1)
       txtcadena = MatValFlujosD(i).c_operacion & Chr(9)                       'Clave de operación
       txtcadena = txtcadena & MatValFlujosD(i).t_pata & Chr(9)                'pata
       txtcadena = txtcadena & MatValFlujosD(i).fecha_ini & Chr(9)             'inicio del flujo
       txtcadena = txtcadena & MatValFlujosD(i).fecha_fin & Chr(9)             'final del flujo
       txtcadena = txtcadena & MatValFlujosD(i).fecha_desc & Chr(9)            'fecha de descuento
       txtcadena = txtcadena & MatValFlujosD(i).saldo_periodo & Chr(9)         'saldo
       txtcadena = txtcadena & MatValFlujosD(i).amortizacion & Chr(9)          'amortizacion
       txtcadena = txtcadena & MatValFlujosD(i).int_gen_periodo & Chr(9)       'intereses generados
       txtcadena = txtcadena & MatValFlujosD(i).int_acum_periodo & Chr(9)      'intereses acumulados
       txtcadena = txtcadena & MatValFlujosD(i).int_pag_periodo & Chr(9)       'intereses pagados
       txtcadena = txtcadena & MatValFlujosD(i).int_acum_sig_periodo & Chr(9)  'intereses acumulados sig periodo
       txtcadena = txtcadena & MatValFlujosD(i).pago_total & Chr(9)            'pago total=amort+int
       txtcadena = txtcadena & MatValFlujosD(i).t_desc & Chr(9)                'tasa descuento
       txtcadena = txtcadena & MatValFlujosD(i).factor_desc & Chr(9)           'factor descuento
       txtcadena = txtcadena & MatValFlujosD(i).valor_presente & Chr(9)        'valor presente
       Print #1, txtcadena
   Next i
   Close #1
   End If
End If
MsgBox "Fin de proceso"
End Sub

Function DetSwapsInicianPer(fecha1 As Date, fecha2 As Date)
   'se validan los swaps que entraron en el mes
   Dim noreg1 As Long
   Dim txtcadena1 As String
   Dim txtcadena2 As String
   Dim txtfecha1 As String
   Dim txtfecha2 As String
   Dim i As Long
   Dim ccontra1 As Integer
   Dim indice As Long
   Dim indice2 As Long
   Dim txtmonact As String
   Dim txtmonpas As String
   Dim txtmon As String
   Dim mnocional As Double
   Dim estruc As String
   Dim rmesa As New ADODB.recordset
   
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtcadena2 = "SELECT COPERACION,FECHAREG,FINICIO,FVALUACION, INTENCION, FVENCIMIENTO,ID_CONTRAP,ESTRUCTURAL FROM " & TablaPosSwaps
   txtcadena2 = txtcadena2 & " WHERE TIPOPOS = 1 AND CPOSICION = " & ClavePosDeriv
   txtcadena2 = txtcadena2 & " AND FECHAREG = FINICIO AND FINICIO > " & txtfecha1 & " AND FINICIO <= " & txtfecha2
   txtcadena1 = "SELECT COUNT(*) FROM (" & txtcadena2 & ")"
   rmesa.Open txtcadena1, ConAdo
   noreg1 = rmesa.Fields(0)
   rmesa.Close
   If noreg1 <> 0 Then
      rmesa.Open txtcadena2, ConAdo
   ReDim mata(1 To noreg1, 1 To 13) As Variant
      For i = 1 To noreg1
          mata(i, 1) = rmesa.Fields("COPERACION")      'Clave de operación
          mata(i, 2) = rmesa.Fields("FECHAREG")        'fecha de registro
          mata(i, 3) = rmesa.Fields("FINICIO")         'fecha de inicio
          mata(i, 4) = rmesa.Fields("FVALUACION")      'tipo de operacion
          estruc = rmesa.Fields("ESTRUCTURAL")
          If rmesa.Fields("INTENCION") = "N" And estruc = "S" Then
             mata(i, 5) = "Negociación estructural"
          ElseIf rmesa.Fields("INTENCION") = "N" And estruc = "N" Then
             mata(i, 5) = "Negociación"
          ElseIf rmesa.Fields("INTENCION") = "C" Then
             mata(i, 5) = "Cobertura"
          End If
          mata(i, 6) = rmesa.Fields("FVENCIMIENTO")    'fecha vencimiento
          indice2 = BuscarValorArray(mata(i, 4), MatTValSwaps1, 1)
          If indice2 <> 0 Then
             txtmonact = ReemplazaVacioValor(MatTValSwaps1(indice2, 13), "")
             txtmonpas = ReemplazaVacioValor(MatTValSwaps1(indice2, 14), "")
          End If
          If indice2 = 0 Then MsgBox "no se determino la moneda"
          Call DetermMonySaldoSwap(mata(i, 1), mata(i, 2), ReemplazaVacioValor(txtmonact, ""), ReemplazaVacioValor(txtmonpas, ""), txtmon, mnocional)
          mata(i, 7) = txtmon
          mata(i, 8) = Round(mnocional / 1000000, 2)
          ccontra1 = rmesa.Fields("ID_CONTRAP")
          indice = BuscarValorArray(ccontra1, MatContrapartes, 1)
          If indice <> 0 Then
             mata(i, 9) = MatContrapartes(indice, 3)
          End If
          rmesa.MoveNext
      Next i
      rmesa.Close
   End If
DetSwapsInicianPer = mata
End Function

Sub DetermMonySaldoSwap(ByVal coperacion As String, ByVal fechareg As Date, ByVal txtmonact As String, ByVal txtmonpas As String, ByRef txtmons As String, ByRef mnocional As Double)
          If Not EsVariableVacia(txtmonact) Then
             txtmons = DetermMonedaSwap(txtmonact)
             mnocional = ObtPrimeFlujoSwap(coperacion, fechareg, "B")
          ElseIf Not EsVariableVacia(txtmonpas) Then
             txtmons = DetermMonedaSwap(txtmonpas)
             mnocional = ObtPrimeFlujoSwap(coperacion, fechareg, "C")
          Else
             txtmons = "MXP"
             mnocional = ObtPrimeFlujoSwap(coperacion, fechareg, "C")
          End If

End Sub

Function DetermMonedaSwap(ByVal txtmon As String)
         If txtmon = "DOLAR PIP FIX" Then
            DetermMonedaSwap = "USD"
         ElseIf txtmon = "EURO BM" Then
            DetermMonedaSwap = "EUR"
         ElseIf txtmon = "YEN BM" Then
            DetermMonedaSwap = "YEN"
         Else
            DetermMonedaSwap = txtmon
         End If
End Function

Function DetSwapCambIntencion(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim contar As Integer
Dim noreg2 As Integer
Dim noreg3 As Integer
Dim txtfechar As String
Dim rmesa As New ADODB.recordset

txtfiltro = "SELECT COPERACION FROM (SELECT COPERACION, COUNT(DISTINCT INTENCION) AS NOINT FROM " & TablaPosSwaps
txtfiltro = txtfiltro & " WHERE TIPOPOS = 1 GROUP BY COPERACION) WHERE NOINT > 1 ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg, 1 To 1) As Variant
rmesa.MoveFirst
For i = 1 To noreg
mata(i, 1) = rmesa.Fields(0)
rmesa.MoveNext
Next i
rmesa.Close
contar = 0
ReDim matc(1 To 9, 1 To 1) As Variant
For i = 1 To noreg
    txtfiltro = "SELECT FECHAREG,INTENCION FROM " & TablaPosSwaps & " WHERE COPERACION = '" & mata(i, 1) & "' AND TIPOPOS = 1 ORDER BY FECHAREG,INTENCION"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro, ConAdo
       rmesa.MoveFirst
       ReDim matb(1 To noreg2, 1 To 2) As Variant
       For j = 1 To noreg2
          matb(j, 1) = rmesa.Fields(0)
          matb(j, 2) = rmesa.Fields(1)
          rmesa.MoveNext
       Next j
       rmesa.Close
       For j = 1 To noreg2 - 1
           If matb(j, 2) <> matb(j + 1, 2) And matb(j + 1, 1) >= fecha1 And matb(j + 1, 1) <= fecha2 Then
              txtfechar = "to_date('" & Format(matb(j + 1, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
              contar = contar + 1
             ReDim Preserve matc(1 To 9, 1 To contar) As Variant
              txtfiltro = "SELECT COPERACION,INTENCION, FECHAREG, FINICIO, FVENCIMIENTO, CPRODUCTO, ID_CONTRAP FROM " & TablaPosSwaps & " WHERE COPERACION = '" & mata(i, 1) & "' AND FECHAREG = " & txtfechar & "  AND TIPOPOS = 1"
              txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
              rmesa.Open txtfiltro1, ConAdo
              noreg2 = rmesa.Fields(0)
              rmesa.Close
              If noreg2 <> 0 Then
                 rmesa.Open txtfiltro, ConAdo
                 rmesa.MoveFirst
                 For p = 1 To noreg2
                     matc(1, contar) = rmesa.Fields("FINICIO")
                     matc(2, contar) = rmesa.Fields("CPRODUCTO")
                     matc(3, contar) = rmesa.Fields("INTENCION")
                     matc(4, contar) = rmesa.Fields("FVENCIMIENTO")
                     matc(5, contar) = ""    'MONEDA
                     matc(6, contar) = 0
                     matc(7, contar) = rmesa.Fields("ID_CONTRAP")
                     matc(8, contar) = ""
                     matc(9, contar) = rmesa.Fields("COPERACION")
                     
                     rmesa.MoveNext
                 Next p
                 rmesa.Close
              End If
           End If
       Next j
    End If
Next i
If contar <> 0 Then
   matc = MTranV(matc)
Else
ReDim matc(0 To 0, 0 To 0) As Variant
End If

End If
DetSwapCambIntencion = matc
End Function

Function DetFwdCambIntencion(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim noreg2 As Integer
Dim noreg3 As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim contar As Integer
Dim txtfechar As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COPERACION FROM (SELECT COPERACION, COUNT(DISTINCT INTENCION) AS NOINT FROM " & TablaPosFwd & " WHERE TIPOPOS = 1 GROUP BY COPERACION) WHERE NOINT > 1 ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg, 1 To 1) As Variant
rmesa.MoveFirst
For i = 1 To noreg
mata(i, 1) = rmesa.Fields(0)
rmesa.MoveNext
Next i
rmesa.Close
contar = 0
ReDim matc(1 To 7, 1 To 1) As Variant
For i = 1 To noreg
    txtfiltro = "SELECT FECHAREG,INTENCION FROM " & TablaPosFwd & " WHERE COPERACION = '" & mata(i, 1) & "' AND TIPOPOS = 1 AND FECHAREG >= " & txtfecha1 & " AND FECHAREG <= " & txtfecha2 & " ORDER BY FECHAREG,INTENCION"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro, ConAdo
       rmesa.MoveFirst
       ReDim matb(1 To noreg2, 1 To 2) As Variant
       For j = 1 To noreg2
          matb(j, 1) = rmesa.Fields(0)
          matb(j, 2) = rmesa.Fields(1)
          rmesa.MoveNext
       Next j
       rmesa.Close
       For j = 1 To noreg2 - 1
           If matb(j, 2) = "C" And matb(j + 1, 2) = "N" Then
             txtfechar = "to_date('" & Format(matb(j + 1, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
             contar = contar + 1
             ReDim Preserve matc(1 To 7, 1 To contar) As Variant
             txtfiltro = "SELECT COPERACION,INTENCION,FECHAREG,FINICIO,FVENCIMIENTO,CPRODUCTO,ID_CONTRAP FROM " & TablaPosFwd & " WHERE COPERACION = '" & mata(i, 1) & "' AND FECHAREG = " & txtfechar & "  AND TIPOPOS = 1"
             txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
             rmesa.Open txtfiltro1, ConAdo
             noreg3 = rmesa.Fields(0)
             rmesa.Close
             If noreg3 <> 0 Then
                rmesa.Open txtfiltro, ConAdo
                rmesa.MoveFirst
                For p = 1 To noreg3
                  matc(1, contar) = rmesa.Fields(0)
                  matc(2, contar) = rmesa.Fields(1)
                  matc(3, contar) = rmesa.Fields(2)
                  matc(4, contar) = rmesa.Fields(3)
                  matc(5, contar) = rmesa.Fields(4)
                  matc(6, contar) = rmesa.Fields(5)
                  matc(7, contar) = rmesa.Fields(6)
                  rmesa.MoveNext
                Next p
                rmesa.Close
             End If
          End If
       Next j
    End If
Next i
If contar <> 0 Then
   DetFwdCambIntencion = MTranV(matc)
Else
ReDim matc(0 To 0, 0 To 0) As Variant
DetFwdCambIntencion = matc
End If
End If

End Function

Function CalcResEfCob(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtcadena As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim noreg1 As Integer
Dim indice As Integer
Dim noinst As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT COPERACION,EFIC_RETRO FROM " & TablaEficRetro & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro2, ConAdo
ReDim matg(1 To noreg, 1 To 10) As Variant
For i = 1 To noreg
    matg(i, 1) = rmesa.Fields(0)              'clave de la operacion
    matg(i, 2) = Val(rmesa.Fields(1)) / 100   'eficiencia
    rmesa.MoveNext
Next i
rmesa.Close
'se obtiene ahora la clave de producto
For i = 1 To noreg
    txtfiltro2 = "SELECT FVALUACION FROM " & TablaPosSwaps & " WHERE COPERACION = '" & matg(i, 1) & "' AND TIPOPOS = 1"
    txtfiltro1 = "SELECT count(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matg(i, 3) = rmesa.Fields(0)     'clave de producto
       rmesa.Close
       indice = BuscarValorArray(matg(i, 3), MatGruposDeriv, 1)
       If indice <> 0 Then
          matg(i, 4) = MatGruposDeriv(indice, 2)
       Else
        MsgBox "no se ha agrupado el producto " & matg(i, 3) & " de la operacion " & matg(i, 1)
       End If
    End If
    txtfiltro2 = "SELECT CPRODUCTO FROM " & TablaPosFwd & " WHERE COPERACION = '" & matg(i, 1) & "' AND FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
    txtfiltro1 = "SELECT count(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matg(i, 3) = rmesa.Fields(0)     'clave de producto
       rmesa.Close
       indice = BuscarValorArray(matg(i, 3), MatGruposDeriv, 1)
       If indice <> 0 Then
          matg(i, 4) = MatGruposDeriv(indice, 2)
       Else
         MsgBox "no se ha agrupado el producto " & matg(i, 3) & " de la operacion " & matg(i, 1)
       End If
    End If
    If EsVariableVacia(matg(i, 4)) Then MsgBox "No se encontro la operacion " & matg(i, 1)
Next i

ReDim matj(1 To 10, 1 To 10) As Variant
matj(1, 1) = "CCS JPY-MXN"
matj(2, 1) = "CCS MXN-UDI"
matj(3, 1) = "CCS MXN-USD"
matj(4, 1) = "CCS UDI-MXN"
matj(5, 1) = "CCS USD-MXN"
matj(6, 1) = "FWD MXN/EUR"
matj(7, 1) = "FWD MXN/USD"
matj(8, 1) = "IRS MXN-MXN"
matj(9, 1) = "IRS USD-USD"
matj(10, 1) = "Total General"
For i = 1 To 9
    matj(i, 4) = 1000
    For j = 1 To noreg
        If matg(j, 4) = matj(i, 1) Then
           matj(i, 2) = matj(i, 2) + 1                    'no de instrumentos
           noinst = noinst + 1
           matj(i, 3) = matj(i, 3) + matg(j, 2)           'suma
           matj(i, 4) = Minimo(matj(i, 4), matg(j, 2))    'minimo
           matj(i, 5) = Maximo(matj(i, 5), matg(j, 2))    'maximo
        End If
    Next j
    If matj(i, 2) <> 0 Then matj(i, 3) = matj(i, 3) / matj(i, 2)
Next i
matj(10, 4) = 1000
For i = 1 To 9
    matj(10, 2) = matj(10, 2) + matj(i, 2)
    matj(10, 3) = matj(10, 3) + matj(i, 3) * matj(i, 2)
    matj(10, 4) = Minimo(matj(10, 4), matj(i, 4))
    matj(10, 5) = Maximo(matj(10, 5), matj(i, 5))
Next i
matj(10, 3) = matj(10, 3) / matj(10, 2)
Else
ReDim matj(0 To 0, 0 To 0) As Variant
End If
CalcResEfCob = matj
End Function

Function CalcResPosDeriv(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim noreg1 As Integer
Dim indice As Integer
Dim noinst As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
'se obtiene ahora la clave de producto
  txtcadena2 = "SELECT FVALUACION FROM " & TablaPosSwaps & " WHERE (TIPOPOS,CPOSICION,COPERACION,FECHAREG) "
  txtcadena2 = txtcadena2 & "IN (SELECT TIPOPOS,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG "
  txtcadena2 = txtcadena2 & "FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
  txtcadena2 = txtcadena2 & " GROUP BY TIPOPOS,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
  txtcadena2 = txtcadena2 & "AND FVENCIMIENTO > " & txtfecha & " AND TIPOPOS = 1"
  txtcadena1 = "SELECT COUNT(*) FROM (" & txtcadena2 & ")"
  rmesa.Open txtcadena1, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
    If noreg <> 0 Then
    
       rmesa.Open txtcadena2, ConAdo
       For i = 1 To noreg
       'matg(i, 3) = RMesa.Fields(0)     'clave de producto
       rmesa.MoveNext
       Next i
       rmesa.Close
       'indice = BuscarValorArray(matg(i, 3), MatGruposDeriv, 1)
       If indice <> 0 Then
          'matg(i, 4) = MatGruposDeriv(indice, 2)
       Else
        'MsgBox "no se ha agrupado el producto " & matg(i, 3) & " de la operacion " & matg(i, 1)
       End If
    End If
    'txtcadena = "SELECT count(*) FROM " & TablaPosFwd & " WHERE COPERACION = '" & matg(i, 1) & "' AND FECHA = " & txtfecha
    rmesa.Open txtcadena2, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 Then
       'txtcadena2 = "SELECT CPRODUCTO FROM " & TablaPosFwd & " WHERE COPERACION = '" & matg(i, 1) & "' AND FECHA = " & txtfecha
       rmesa.Open txtcadena2, ConAdo
       
       'matg(i, 3) = RMesa.Fields(0)     'clave de producto
       rmesa.Close
       'indice = BuscarValorArray(matg(i, 3), MatGruposDeriv, 1)
       If indice <> 0 Then
          'matg(i, 4) = MatGruposDeriv(indice, 2)
       Else
         'MsgBox "no se ha agrupado el producto " & matg(i, 3) & " de la operacion " & matg(i, 1)
       End If
    End If
   'If EsVariableVacia(matg(i, 4)) Then MsgBox "No se encontro la operacion " & matg(i, 1)
   
'Next i

ReDim matj(1 To 10, 1 To 10) As Variant
matj(1, 1) = "CCS JPY-MXN"
matj(2, 1) = "CCS MXN-UDI"
matj(3, 1) = "CCS MXN-USD"
matj(4, 1) = "CCS UDI-MXN"
matj(5, 1) = "CCS USD-MXN"
matj(6, 1) = "FWD MXN/EUR"
matj(7, 1) = "FWD MXN/USD"
matj(8, 1) = "IRS MXN-MXN"
matj(9, 1) = "IRS USD-USD"
matj(10, 1) = "Total General"
'For i = 1 To 9
'matj(i, 4) = 1000
'For j = 1 To noreg
'If matg(j, 4) = matj(i, 1) Then
  'matj(i, 2) = matj(i, 2) + 1                    'no de instrumentos
  'noinst = noinst + 1
  'matj(i, 3) = matj(i, 3) + matg(j, 2)           'suma
  'matj(i, 4) = Minimo(matj(i, 4), matg(j, 2))    'minimo
  'matj(i, 5) = Maximo(matj(i, 5), matg(j, 2))    'maximo
'End If
'Next j
'If matj(i, 2) <> 0 Then matj(i, 3) = matj(i, 3) / matj(i, 2)
'Next i
'matj(10, 4) = 1000
'For i = 1 To 9
'    matj(10, 2) = matj(10, 2) + matj(i, 2)
'    matj(10, 3) = matj(10, 3) + matj(i, 3) * matj(i, 2)
'    matj(10, 4) = Minimo(matj(10, 4), matj(i, 4))
'    matj(10, 5) = Maximo(matj(10, 5), matj(i, 5))
'Next i
'matj(10, 3) = matj(10, 3) / matj(10, 2)
''Else
'ReDim matj(0 To 0, 0 To 0) As Variant
'End If
'CalcResEfCob = matj
End Function

Function ObtPrimeFlujoSwap(coperacion, fechareg, tpata)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa2 As New ADODB.recordset

txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT FINICIO, SALDO FROM " & TablaFlujosSwapsO & " WHERE COPERACION = '" & coperacion & "'"
txtfiltro = txtfiltro & " AND TIPOPOS = 1 AND FECHAREG = " & txtfecha & " AND TPATA = '" & tpata & "' ORDER BY FINICIO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa2.Open txtfiltro1, ConAdo
noreg = rmesa2.Fields(0)
rmesa2.Close
If noreg <> 0 Then
   rmesa2.Open txtfiltro, ConAdo
   rmesa2.MoveFirst
ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa2.Fields("FINICIO")
       mata(i, 2) = rmesa2.Fields("SALDO")
       rmesa2.MoveNext
   Next i
   rmesa2.Close
   ObtPrimeFlujoSwap = mata(1, 2)
Else
   ObtPrimeFlujoSwap = 0
End If
End Function

Function ObtFlujoSwapFecha(ByVal coperacion As String, ByVal fechareg As Date, ByVal fecha As Date, ByVal tpata As String)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro2 As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa2 As New ADODB.recordset

txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT FINICIO, SALDO FROM " & TablaFlujosSwapsO & " WHERE COPERACION = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG = " & txtfecha & " AND TPATA = '" & tpata & "'"
txtfiltro2 = txtfiltro2 & " AND FINICIO<= " & txtfecha1 & " AND FFINAL > " & txtfecha1
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa2.Open txtfiltro1, ConAdo
noreg = rmesa2.Fields(0)
rmesa2.Close
If noreg <> 0 Then
   rmesa2.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa2.Fields("FINICIO")
       mata(i, 2) = rmesa2.Fields("SALDO")
       rmesa2.MoveNext
   Next i
   rmesa2.Close
   ObtFlujoSwapFecha = mata(1, 2)
Else
   ObtFlujoSwapFecha = 0
End If
End Function

Function ObtTasaSwapFecha(ByVal coperacion As String, ByVal fechareg As Date, ByVal fecha As Date, ByVal tpata As String)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro2 As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa2 As New ADODB.recordset

txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT FINICIO, TASA FROM " & TablaFlujosSwapsO & " WHERE COPERACION = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG = " & txtfecha & " AND TPATA = '" & tpata & "'"
txtfiltro2 = txtfiltro2 & " AND FINICIO<= " & txtfecha1 & " AND FFINAL > " & txtfecha1
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa2.Open txtfiltro1, ConAdo
noreg = rmesa2.Fields(0)
rmesa2.Close
If noreg <> 0 Then
   rmesa2.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa2.Fields("FINICIO")
       mata(i, 2) = rmesa2.Fields("TASA")
       rmesa2.MoveNext
   Next i
   rmesa2.Close
   ObtTasaSwapFecha = mata(1, 2)
Else
   ObtTasaSwapFecha = 0
End If
End Function



Function DetPosVenceSwaps(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim contar As Integer
Dim indice As Integer
Dim monact As String
Dim monpas As String
Dim txtmons As String
Dim mnocional As Double
Dim estruc As String
Dim ccontra1 As Integer
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COPERACION FROM " & TablaPosSwaps & " WHERE TIPOPOS = 1  AND FVENCIMIENTO > " & txtfecha1 & " AND FVENCIMIENTO <= " & txtfecha2 & " GROUP BY COPERACION ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(0)
       mata(i, 2) = DetFechaReg(mata(i, 1), fecha2)
       rmesa.MoveNext
Next i
rmesa.Close
ReDim matb(1 To 11, 1 To 1) As Variant
contar = 0
For i = 1 To noreg
    txtfecha3 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro = "SELECT FECHAREG,COPERACION,FINICIO,FVALUACION,INTENCION,FVENCIMIENTO,ID_CONTRAP,ESTRUCTURAL FROM " & TablaPosSwaps & " WHERE COPERACION = '" & mata(i, 1) & "' AND FECHAREG = " & txtfecha3 & " AND TIPOPOS = 1 AND FVENCIMIENTO > " & txtfecha1 & " AND FVENCIMIENTO <= " & txtfecha2
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 Then
       contar = contar + 1
       ReDim Preserve matb(1 To 11, 1 To contar) As Variant
       rmesa.Open txtfiltro, ConAdo
       rmesa.MoveFirst
       For j = 1 To noreg1
            matb(1, contar) = rmesa.Fields("COPERACION")
            matb(2, contar) = rmesa.Fields("FECHAREG")
            matb(3, contar) = rmesa.Fields("FINICIO")
            matb(4, contar) = rmesa.Fields("FVALUACION")
            estruc = rmesa.Fields("ESTRUCTURAL")
            If rmesa.Fields("INTENCION") = "N" And estruc = "S" Then
               matb(5, contar) = "Negociación Estructural"
            ElseIf rmesa.Fields("INTENCION") = "N" And estruc = "N" Then
               matb(5, contar) = "Negociación"
            ElseIf rmesa.Fields("INTENCION") = "C" Then
               matb(5, contar) = "Cobertura"
            End If
            matb(6, contar) = rmesa.Fields("FVENCIMIENTO")
            indice = BuscarValorArray(matb(4, contar), MatTValSwaps1, 1)
            If indice <> 0 Then
                monact = ReemplazaVacioValor(MatTValSwaps1(indice, 13), "")
                monpas = ReemplazaVacioValor(MatTValSwaps1(indice, 14), "")
                Call DetermMonySaldoSwap(matb(1, contar), matb(2, contar), monact, monpas, txtmons, mnocional)
                matb(7, contar) = txtmons
                matb(8, contar) = mnocional / 1000000
            End If
            ccontra1 = rmesa.Fields("ID_CONTRAP")
            indice = BuscarValorArray(ccontra1, MatContrapartes, 1)
            If indice <> 0 Then
               matb(9, contar) = MatContrapartes(indice, 3)
            End If
            rmesa.MoveNext
       Next j
    rmesa.Close
End If
Next i
If contar <> 0 Then
DetPosVenceSwaps = MTranV(matb)
Else
ReDim matb(0 To 0, 0 To 0) As Variant
DetPosVenceSwaps = matb
End If
Else
ReDim matb(0 To 0, 0 To 0) As Variant
DetPosVenceSwaps = matb
End If
End Function

Function detFwdUnwind(fecha1 As Date, fecha2 As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim noreg0 As Long
Dim contar As Long
Dim mata() As Variant
Dim matb() As Variant
Dim fechaa As Date
Dim fechab As Date
Dim f_val As Date
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfechaa As String
Dim txtfechab As String
Dim norega As Long
Dim noregb As Long
Dim rmesa As New ADODB.recordset
contar = 0
'se obtiene la lista de operaciones en el mes
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT COPERACION,INTENCION,FINICIO,FVENCIMIENTO,CPRODUCTO FROM " & TablaPosFwd & " WHERE FECHAREG >=" & txtfecha1 & " AND FECHAREG <= " & txtfecha2
txtfiltro2 = txtfiltro2 & " GROUP BY COPERACION,INTENCION,FINICIO,FVENCIMIENTO,CPRODUCTO ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 7) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(0)     'clave de operacion
       mata(i, 2) = rmesa.Fields(1)     'intencion
       mata(i, 3) = rmesa.Fields(2)     'fecha de inicio
       mata(i, 4) = rmesa.Fields(3)     'fecha vencimiento
       mata(i, 5) = rmesa.Fields(4)     'tipo de operacion
       rmesa.MoveNext
   Next i
   rmesa.Close
   contar = 0
   ReDim matb(1 To 5, 1 To 1) As Variant
   For i = 1 To noreg
       txtfiltro2 = "SELECT MAX(FECHAREG) AS FECHA FROM " & TablaPosFwd
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG >= " & txtfecha1 & " AND COPERACION = '" & mata(i, 1) & "'"
       txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha2
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg0 = rmesa.Fields(0)
       rmesa.Close
       If noreg0 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          If Not EsVariableVacia(rmesa.Fields("FECHA")) Then
             f_val = rmesa.Fields("FECHA")
             If f_val < fecha2 Then
                contar = contar + 1
             ReDim Preserve matb(1 To 5, 1 To contar) As Variant
                matb(1, contar) = mata(i, 1)
                matb(2, contar) = mata(i, 2)
                matb(3, contar) = mata(i, 3)
                matb(4, contar) = mata(i, 4)
                matb(5, contar) = mata(i, 5)
             End If
          End If
          rmesa.Close
       End If
   Next i
   If contar <> 0 Then
      matb = MTranV(matb)
   Else
    ReDim matb(0 To 0, 0 To 0) As Variant
   End If
   detFwdUnwind = matb
End If
End Function

Function DetFwdsPactados(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro2 As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim indice As Integer
Dim estruc As String
Dim idcontrap As Integer
Dim fecham As Date
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT FECHAREG, COPERACION,INTENCION,FINICIO,"
txtfiltro2 = txtfiltro2 & "FVENCIMIENTO,M_NOCIONAL,CPRODUCTO,ID_CONTRAP,ESTRUCTURAL"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosFwd & " WHERE (FECHAREG,COPERACION) IN"
txtfiltro2 = txtfiltro2 & "(SELECT FECHAREG,COPERACION FROM "
txtfiltro2 = txtfiltro2 & "(SELECT MIN(FECHAREG) AS FECHAREG,COPERACION FROM "
txtfiltro2 = txtfiltro2 & TablaPosFwd & " WHERE TIPOPOS = 1 GROUP BY COPERACION)"
txtfiltro2 = txtfiltro2 & " WHERE  FECHAREG > " & txtfecha1 & " AND FECHAREG <= " & txtfecha2 & ") AND TIPOPOS = 1"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 10) As Variant
rmesa.MoveFirst
For i = 1 To noreg
    fecham = rmesa.Fields("FECHAREG")
    mata(i, 1) = rmesa.Fields("COPERACION")
    mata(i, 2) = Minimo(rmesa.Fields("FINICIO"), fecham)
    mata(i, 3) = rmesa.Fields("CPRODUCTO")
    estruc = rmesa.Fields("ESTRUCTURAL")
    If rmesa.Fields("INTENCION") = "N" And estruc = "S" Then
       mata(i, 4) = "Negociación estructural"
    ElseIf rmesa.Fields("INTENCION") = "N" And estruc = "N" Then
       mata(i, 4) = "Negociación"
    ElseIf rmesa.Fields("INTENCION") = "C" Then
       mata(i, 4) = "Cobertura"
    End If
    mata(i, 5) = rmesa.Fields("FVENCIMIENTO")
    mata(i, 6) = Right(mata(i, 3), 3)
    mata(i, 7) = rmesa.Fields("M_NOCIONAL") / 1000000
    idcontrap = rmesa.Fields("ID_CONTRAP")
    indice = BuscarValorArray(idcontrap, MatContrapartes, 1)
    If indice <> 0 Then
       mata(i, 8) = MatContrapartes(indice, 3)
    End If

   rmesa.MoveNext
Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
DetFwdsPactados = mata
End Function



Function LeerGarantias(ByVal fecha As Date) As Variant()
Dim nomarch As String
Dim siarch As Boolean
Dim noreg As Long
Dim i As Long
Dim nocampos As Integer
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset

  frmCalVar.CommonDialog1.FileName = nomarch
  frmCalVar.CommonDialog1.ShowSave
  nomarch = frmCalVar.CommonDialog1.FileName
  siarch = VerifAccesoArch(nomarch)
  If siarch Then

Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset("HOJA1$", dbOpenDynaset, dbReadOnly)
'se cargan las dos curvas necesarias para este proceso
If registros1.RecordCount <> 0 Then
   registros1.MoveLast
   noreg = registros1.RecordCount
   nocampos = registros1.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   registros1.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = LeerTAccess(registros1, 0, i)   'clave de la contraparte
       mata(i, 2) = LeerTAccess(registros1, 1, i)   'monto total
       mata(i, 3) = LeerTAccess(registros1, 2, i)   'monto en pesos
       mata(i, 4) = LeerTAccess(registros1, 3, i)   'monto en dolares
       mata(i, 5) = LeerTAccess(registros1, 4, i)   'colateral
       registros1.MoveNext
   Next i
   registros1.Close
   base1.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerGarantias = mata
End Function

Function LeerProbInc(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim valor As Double
Dim rmesa As New ADODB.recordset

   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM SRL_PROBA_INC WHERE FECHA = " & txtfecha
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      valor = rmesa.Fields(9)
      rmesa.Close
   Else
     valor = 0
   End If
   If valor = 0 Then MsgBox "El Valor de la clausula de incumplimiento es 0"
   LeerProbInc = valor
End Function

Sub CalcCCLBanobras(ByVal fecha As Date)
Dim indice As Integer
Dim fechax As Date
Dim matx1() As Variant
Dim matx2() As Variant
Dim matx3() As Variant
Dim matx4() As Variant
Dim matc() As Variant
Dim txtfecha As String
Dim matcolat() As Variant
Dim valmtmprob As Double
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim noreg3 As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim signo As Integer
Dim saldo As Double
Dim tcaplica As Double
Dim indice2 As Long
Dim mm As Integer
Dim k As Integer
Dim txtcadena As String
Dim ClaveBan As String
Dim nomarch As String
Dim txtborra As String
Dim txtinserta As String
Dim rmesa As New ADODB.recordset

    frmProgreso.Show
    fechax = fecha
    Do While True
       indice = BuscarValorArray(fechax, MatFechasFR, 1)
       If indice <> 0 Then
          Exit Do
       Else
          fechax = fechax - 1
       End If
    Loop
    matx1 = Leer1FactorR(fechax, fechax, "DOLAR PIP FIX", 0)
    matx2 = Leer1FactorR(fechax, fechax, "UDI", 0)
    matx3 = Leer1FactorR(fechax, fechax, "EURO BM", 0)
    matx4 = Leer1FactorR(fechax, fechax, "YEN BM", 0)
    matcolat = LeerGarantias(fecha)
    txtfecha = Format(fecha, "yyyymmdd")
    valmtmprob = LeerProbInc(fecha)
'generar la valuacion de la posicion de derivados en funcion de datos del sistema y de ikos derivados
    txtfiltro = "select * from " & TablaValDeriv & " WHERE FECHA = " & txtfecha & " and IS_NUMBER(CLAVE) = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    rmesa.Open txtfiltro, ConAdo
    rmesa.MoveFirst
    ReDim mata(1 To noreg, 1 To 23) As Variant
    For i = 1 To noreg
        mata(i, 1) = rmesa.Fields(0)                 'fecha
        mata(i, 2) = rmesa.Fields(1)                 'clave de operacion
        mata(i, 3) = rmesa.Fields(3)                 'clave tipo producto
        mata(i, 4) = rmesa.Fields(4)                 'fecha de vencimiento
        mata(i, 5) = Val(rmesa.Fields(5))            'val pata activa sivarmer
        mata(i, 6) = Val(rmesa.Fields(6))            'val pata pasiva sivarmer
        mata(i, 7) = Val(rmesa.Fields(7))            'mtm sivarmer
        mata(i, 8) = Val(rmesa.Fields(8))            'val pata activa ikos
        mata(i, 9) = Val(rmesa.Fields(9))            'val pata pasiva ikos
        mata(i, 10) = Val(rmesa.Fields(10))          'mtm ikos
        mata(i, 11) = rmesa.Fields(12)               'id_contraparte
        mata(i, 12) = Maximo(mata(i, 4) - fecha, 0)  'plazo del instrumento
        rmesa.MoveNext
    Next i
rmesa.Close
signo = 1
For i = 1 To noreg
    mata(i, 13) = 0                         'factor por prob incumplimiento
    If mata(i, 2) = "38" Then mata(i, 13) = valmtmprob
    mata(i, 14) = mata(i, 5)                'val pata activa ajustada por pi
    mata(i, 15) = mata(i, 6) + mata(i, 13)  'val pata pasiva ajustada por pi
    mata(i, 16) = mata(i, 14) - mata(i, 15) 'mtm ajustada por pi
    indice = BuscarValorArray(mata(i, 3), MatTValSwaps1, 1)
    If indice <> 0 Then
       mata(i, 17) = "Swap"
       If EsVariableVacia(MatTValSwaps1(indice, 16)) Then
          mata(i, 18) = "MXN"                                'tipo cambio activa
       Else
          mata(i, 18) = Trim(MatTValSwaps1(indice, 16))       'tipo cambio activa
       End If
       If EsVariableVacia(MatTValSwaps1(indice, 17)) Then
          mata(i, 19) = "MXN"                                'tipo cambio pasiva
       Else
          mata(i, 19) = Trim(MatTValSwaps1(indice, 17))       'tipo cambio pasiva
       End If
       If mata(i, 18) = mata(i, 19) Then
          mata(i, 20) = "Tasa"
       Else
          mata(i, 20) = "FX"
       End If
       mata(i, 21) = kpt(mata(i, 12), mata(i, 20))           'formula de ponderacion
       fechax = DetFechaReg(mata(i, 2), fecha)
       If fechax <> 0 Then
          saldo = ObtSaldoSwapFecha(fecha, fechax, mata(i, 2), "C")
          If mata(i, 19) <> "MXN" Then
             tcaplica = ObtTCCCL(fecha, mata(i, 19), matx1, matx2, matx3, matx4)
          Else
             tcaplica = 1
          End If
          'xdp(p)
          mata(i, 22) = saldo * tcaplica * mata(i, 21) / 12 ^ 0.5
       End If
       'si mtm >0 entonces
       If mata(i, 16) >= 0 Then
          mata(i, 23) = mata(i, 18)
       ElseIf mata(i, 16) < 0 Then
          mata(i, 23) = mata(i, 19)
       End If
    Else
    'la operacion es un forward de tipo de cambio
      indice2 = BuscarValorArray(mata(i, 3), MatTValFwdsTC1, 1)
      'If mata(i, 11) = "40110" Then MsgBox mata(i, 2)
      'aguas estoy suponiendo que solo tengo operaciones largas
      If indice2 <> 0 Then
         If signo = 1 Then
            mata(i, 18) = Trim(MatTValFwdsTC1(indice2, 14))         'tipo cambio activa
            mata(i, 19) = "MXN"                                    'tipo cambio pasiva
         Else
            mata(i, 18) = "MXN"                                    'tipo cambio pasiva
            mata(i, 19) = Trim(MatTValFwdsTC1(indice2, 14))         'tipo cambio activa
         End If
         mata(i, 17) = "Fwd"
         mata(i, 20) = "FX"                               'todos los forwards son de tc
         mata(i, 21) = kpt(mata(i, 12), mata(i, 20))      'formula de ponderacion
         saldo = ObtSaldoFwdFecha(fecha, mata(i, 2))
         tcaplica = ObtTCCCL(fecha, MatTValFwdsTC1(indice2, 14), matx1, matx2, matx3, matx4)
         'xdp(p)
         mata(i, 22) = saldo * tcaplica * mata(i, 21) / 12 ^ 0.5
         If mata(i, 16) < 0 Then
            mata(i, 23) = mata(i, 19)
         ElseIf mata(i, 16) >= 0 Then
            mata(i, 23) = mata(i, 18)
         End If
      Else
         MsgBox "No se clasifico la operación " & mata(i, 2)
      End If
    End If
Next i

matc = ObtFactUnicos(mata, 11)
noreg3 = UBound(matc, 1)
'se agrupan por activo y pasivo por swap y fwd y por tasa y fx
ReDim matp(1 To noreg3, 1 To 16, 1 To 4) As Variant
ReDim mates(1 To 4, 1 To 2) As Variant
ReDim matmon(1 To 4) As String
matmon(1) = "MXN,UDI": matmon(2) = "DOLAR PIP FIX": matmon(3) = "EURO BM": matmon(4) = "YEN BM"

mates(1, 1) = "Tasa": mates(1, 2) = "Swap"
mates(2, 1) = "Tasa": mates(2, 2) = "Fwd"
mates(3, 1) = "FX": mates(3, 2) = "Swap"
mates(4, 1) = "FX": mates(4, 2) = "Fwd"
'suma de patas activas y pasivas, se segmentan los swaps por moneda y se suman las partes correspondientes
For p = 1 To 4
For i = 1 To noreg3
    For j = 1 To noreg
       For mm = 1 To 4
       If matc(i, 1) = mata(j, 11) And InStr(matmon(mm), mata(j, 18)) <> 0 And mata(j, 20) = mates(p, 1) And mata(j, 17) = mates(p, 2) Then
          matp(i, 3 * mm - 2, p) = matp(i, 3 * mm - 2, p) + mata(j, 14)
       End If
       If matc(i, 1) = mata(j, 11) And InStr(matmon(mm), mata(j, 19)) <> 0 And mata(j, 20) = mates(p, 1) And mata(j, 17) = mates(p, 2) Then
          matp(i, 3 * mm - 1, p) = matp(i, 3 * mm - 1, p) + mata(j, 15)
       End If
       Next mm
    Next j
Next i
'el neto de activas y pasivas
For i = 1 To noreg3
    matp(i, 3, p) = matp(i, 1, p) - matp(i, 2, p)
    matp(i, 6, p) = matp(i, 4, p) - matp(i, 5, p)
    matp(i, 9, p) = matp(i, 7, p) - matp(i, 8, p)
    matp(i, 12, p) = matp(i, 10, p) - matp(i, 11, p)
Next i
Next p
ReDim matr(1 To 5, 1 To 3, 1 To 26) As Variant
'suma de fwds de tasa y fx
For p = 1 To 4
If p = 2 Or p = 4 Then
For i = 1 To noreg3
For k = 1 To 5
    matr(k, 1, 1) = matr(k, 1, 1) + matp(i, 3 * k - 2, p)
    matr(k, 2, 1) = matr(k, 2, 1) + matp(i, 3 * k - 1, p)
    matr(k, 3, 1) = matr(k, 3, 1) + matp(i, 3 * k, p)
Next k
Next i
End If
'suma de swaps de tasa y fx
If p = 1 Or p = 3 Then
For i = 1 To noreg3
For k = 1 To 5
matr(k, 1, 2) = matr(k, 1, 2) + matp(i, 3 * k - 2, p) 'activo mxn + udis
matr(k, 2, 2) = matr(k, 2, 2) + matp(i, 3 * k - 1, p) 'pasivo mxn +udis
matr(k, 3, 2) = matr(k, 3, 2) + matp(i, 3 * k, p)  'neto
Next k
Next i
End If
Next p


'se agrupan por valuaciones positivas y negativas, se agrupan por moneda dominante,en este caso,si el instrumento
'es moneda nacional y otra moneda, domina la moneda foranea
ReDim mats(1 To noreg3, 1 To 53, 1 To 4) As Variant
For p = 1 To 4
    For i = 1 To noreg3
        mats(i, 1, p) = matc(i, 1)
    Next i
For i = 1 To noreg3
    For j = 1 To noreg
        For mm = 1 To 4
            If mats(i, 1, p) = mata(j, 11) And InStr(matmon(mm), mata(j, 23)) <> 0 And mata(j, 20) = mates(p, 1) And mata(j, 17) = mates(p, 2) Then
               mats(i, 13 + mm, p) = mats(i, 13 + mm, p) + mata(j, 22)     'factor de ponderacion
               If mata(j, 16) >= 0 Then
                  mats(i, 3 * mm - 1, p) = mats(i, 3 * mm - 1, p) + mata(j, 16) 'mtm positivos
               ElseIf mata(j, 16) < 0 Then
                  mats(i, 3 * mm, p) = mats(i, 3 * mm, p) + mata(j, 16)         'mtm negativos
               End If
            End If
         Next mm
    Next j
Next i
For i = 1 To noreg3
    mats(i, 4, p) = mats(i, 2, p) + mats(i, 3, p)                   'mtm en pesos
    mats(i, 7, p) = mats(i, 5, p) + mats(i, 6, p)                   'mtm en dolares
    mats(i, 10, p) = mats(i, 8, p) + mats(i, 9, p)                  'mtm en euros
    mats(i, 13, p) = mats(i, 11, p) + mats(i, 12, p)                'mtm en yenes
Next i
Next p

'se calculan calculos necesarios para el report
For p = 1 To 4
    For i = 1 To noreg3
'salidas
'sd(p)
    mats(i, 18, p) = Abs(Minimo(0, mats(i, 4, p)))
    mats(i, 19, p) = Abs(Minimo(0, mats(i, 7, p)))
    mats(i, 20, p) = Abs(Minimo(0, mats(i, 10, p)))
    mats(i, 21, p) = Abs(Minimo(0, mats(i, 13, p)))
    mats(i, 22, p) = Abs(Minimo(0, mats(i, 4, p) + mats(i, 7, p) + mats(i, 10, p) + mats(i, 13, p)))
'xpd(p)
    mats(i, 23, p) = mats(i, 14, p)
    mats(i, 24, p) = mats(i, 15, p)
    mats(i, 25, p) = mats(i, 16, p)
    mats(i, 26, p) = mats(i, 17, p)
    mats(i, 27, p) = mats(i, 14, p) + mats(i, 15, p) + mats(i, 16, p) + mats(i, 17, p)
'ngr(p)
If mats(i, 3, p) <> 0 Then
mats(i, 28, p) = mats(i, 18, p) / Abs(mats(i, 3, p))
Else
mats(i, 28, p) = 1
End If
If mats(i, 6, p) <> 0 Then
mats(i, 29, p) = mats(i, 19, p) / Abs(mats(i, 6, p))
Else
mats(i, 29, p) = 1
End If
If mats(i, 9, p) <> 0 Then
mats(i, 30, p) = mats(i, 20, p) / mats(i, 9, p)
Else
mats(i, 30, p) = 1
End If
If mats(i, 12, p) <> 0 Then
mats(i, 31, p) = mats(i, 21, p) / mats(i, 12, p)
Else
mats(i, 31, p) = 1
End If
If Abs((mats(i, 3, p) + mats(i, 6, p) + mats(i, 9, p) + mats(i, 12, p))) <> 0 Then
   mats(i, 32, p) = mats(i, 22, p) / Abs((mats(i, 3, p) + mats(i, 6, p) + mats(i, 9, p) + mats(i, 12, p)))
Else
   mats(i, 32, p) = 1
End If
'exposicion
   mats(i, 33, p) = (0.4 + 0.6 * mats(i, 28, p)) * mats(i, 23, p)
   mats(i, 34, p) = (0.4 + 0.6 * mats(i, 29, p)) * mats(i, 24, p)
   mats(i, 35, p) = (0.4 + 0.6 * mats(i, 30, p)) * mats(i, 25, p)
   mats(i, 36, p) = (0.4 + 0.6 * mats(i, 31, p)) * mats(i, 26, p)
   mats(i, 37, p) = (0.4 + 0.6 * mats(i, 32, p)) * mats(i, 27, p)
'ED(p)
  mats(i, 38, p) = Maximo(0, mats(i, 4, p))
  mats(i, 39, p) = Maximo(0, mats(i, 7, p))
  mats(i, 40, p) = Maximo(0, mats(i, 10, p))
  mats(i, 41, p) = Maximo(0, mats(i, 13, p))
  mats(i, 42, p) = Maximo(0, mats(i, 4, p) + mats(i, 7, p) + mats(i, 10, p) + mats(i, 13, p))
Next i
Next p

ReDim matt(1 To noreg3, 1 To 55, 1 To 2) As Variant
'swaps
For p = 1 To 4
If p = 1 Or p = 3 Then
For i = 1 To noreg3
    For mm = 1 To 25
        matt(i, mm, 1) = matt(i, mm, 1) + mats(i, mm + 17, p)
    Next mm
Next i
End If
Next p
'se anexa el colaretal a la tabla y se calculan resultados derivados
For i = 1 To noreg3
    For j = 1 To UBound(matcolat, 1)
    If matc(i, 1) = matcolat(j, 1) Then
       If matcolat(j, 5) = "Entregado" Then
          matt(i, 26, 1) = matcolat(j, 3)   'colateral en pesos
          matt(i, 27, 1) = matcolat(j, 4)   'colateral en dolares (expresado en pesos)
          matt(i, 28, 1) = 0
          matt(i, 29, 1) = 0
          matt(i, 30, 1) = matcolat(j, 2)
       Else                 'recibido
          matt(i, 31, 1) = matcolat(j, 3)  'colateral en pesos
          matt(i, 32, 1) = matcolat(j, 4)  'colateral en dolares (expresado en pesos)
          matt(i, 33, 1) = 0
          matt(i, 34, 1) = 0
          matt(i, 35, 1) = matcolat(j, 2)
       End If
    End If
    Next j
Next i
'forwards

For p = 1 To 4
If p = 2 Or p = 4 Then
For i = 1 To noreg3
    For mm = 1 To 25
        matt(i, mm, 2) = matt(i, mm, 2) + mats(i, mm + 17, p)
    Next mm
Next i
End If
Next p

ReDim matsum(1 To 35, 1 To 2) As Variant
For mm = 1 To 2
For i = 1 To noreg3
    'SDcc = max(sd(p)-ed(p),0)
    matt(i, 36, mm) = Maximo(matt(i, 1, mm) - matt(i, 21, mm), 0)
    matt(i, 37, mm) = Maximo(matt(i, 2, mm) - matt(i, 22, mm), 0)
    matt(i, 38, mm) = Maximo(matt(i, 3, mm) - matt(i, 23, mm), 0)
    matt(i, 39, mm) = Maximo(matt(i, 4, mm) - matt(i, 24, mm), 0)
    matt(i, 40, mm) = Maximo(matt(i, 5, mm) - matt(i, 25, mm), 0)
    'max(ed(p)-sd(p),0)
    matt(i, 41, mm) = Maximo(matt(i, 21, mm) - matt(i, 1, mm), 0)
    matt(i, 42, mm) = Maximo(matt(i, 22, mm) - matt(i, 2, mm), 0)
    matt(i, 43, mm) = Maximo(matt(i, 23, mm) - matt(i, 3, mm), 0)
    matt(i, 44, mm) = Maximo(matt(i, 24, mm) - matt(i, 4, mm), 0)
    matt(i, 45, mm) = Maximo(matt(i, 25, mm) - matt(i, 5, mm), 0)
    'sdi
    'max(sd(p) mxn + exposicion mxn-colateral entregado mxn - max(colateral entregado usd - exposicion usd - sd(p) usd,0),0)
    matt(i, 46, mm) = Maximo(matt(i, 1, mm) + matt(i, 16, mm) - matt(i, 26, mm) - Maximo(matt(i, 27, mm) - matt(i, 17, mm) - matt(i, 2, mm), 0), 0)
    'max(sd(p) usd + exposicion usd - colateral entregado usd - max(colateral entregado mxn - exposicion mxn - sd(p) mxn,0),0)
    matt(i, 47, mm) = Maximo(matt(i, 2, mm) + matt(i, 17, mm) - matt(i, 27, mm) - Maximo(matt(i, 26, mm) - matt(i, 16, mm) - matt(i, 1, mm), 0), 0)
    matt(i, 48, mm) = Maximo(matt(i, 3, mm) + matt(i, 18, mm) - matt(i, 28, mm), 0)
    matt(i, 49, mm) = Maximo(matt(i, 4, mm) + matt(i, 19, mm) - matt(i, 29, mm), 0)
    matt(i, 50, mm) = Maximo(matt(i, 5, mm) + matt(i, 20, mm) - matt(i, 30, mm), 0)
    'END(p)
    'max(min(ED(p) mxn - colateral recibido mxn - max(colateral recibido usd - ED(p) usd,0),sd(p) mxn),0)
    matt(i, 51, mm) = Maximo(Minimo(matt(i, 21, mm) - matt(i, 31, mm) - Maximo(matt(i, 32, 1) - matt(i, 22, mm), 0), matt(i, 1, mm)), 0)
    'max(min(ED(p) usd - colateral recibido usd - max(colateral recibido mxn - ED(p) mxn,0),sd(p) usd),0)
    matt(i, 52, mm) = Maximo(Minimo(matt(i, 22, mm) - matt(i, 32, mm) - Maximo(matt(i, 31, 1) - matt(i, 21, mm), 0), matt(i, 2, mm)), 0)
    matt(i, 53, mm) = Maximo(Minimo(matt(i, 23, mm) - matt(i, 33, mm), matt(i, 3, mm)), 0)
    matt(i, 54, mm) = Maximo(Minimo(matt(i, 24, mm) - matt(i, 34, mm), matt(i, 4, mm)), 0)
    matt(i, 55, mm) = Maximo(Minimo(matt(i, 25, mm) - matt(i, 35, mm), matt(i, 5, mm)), 0)
    'suma exposicion
    matsum(1, mm) = matsum(1, mm) + matt(i, 16, mm)
    matsum(2, mm) = matsum(2, mm) + matt(i, 17, mm)
    matsum(3, mm) = matsum(3, mm) + matt(i, 18, mm)
    matsum(4, mm) = matsum(4, mm) + matt(i, 19, mm)
    matsum(5, mm) = matsum(5, mm) + matt(i, 20, mm)
    'suma max(sd(p)-ed(p),0)
    matsum(6, mm) = matsum(6, mm) + matt(i, 36, mm)
    matsum(7, mm) = matsum(7, mm) + matt(i, 37, mm)
    matsum(8, mm) = matsum(8, mm) + matt(i, 38, mm)
    matsum(9, mm) = matsum(9, mm) + matt(i, 39, mm)
    matsum(10, mm) = matsum(10, mm) + matt(i, 40, mm)
    'max(ed(p)-sd(p),0)
    matsum(11, mm) = matsum(11, mm) + matt(i, 41, mm)
    matsum(12, mm) = matsum(12, mm) + matt(i, 42, mm)
    matsum(13, mm) = matsum(13, mm) + matt(i, 43, mm)
    matsum(14, mm) = matsum(14, mm) + matt(i, 44, mm)
    matsum(15, mm) = matsum(15, mm) + matt(i, 45, mm)
    'suma SDi
    matsum(16, mm) = matsum(16, mm) + matt(i, 46, mm)
    matsum(17, mm) = matsum(17, mm) + matt(i, 47, mm)
    matsum(18, mm) = matsum(18, mm) + matt(i, 48, mm)
    matsum(19, mm) = matsum(19, mm) + matt(i, 49, mm)
    matsum(20, mm) = matsum(20, mm) + matt(i, 50, mm)
    'suma END
    matsum(21, mm) = matsum(21, mm) + matt(i, 51, mm)
    matsum(22, mm) = matsum(22, mm) + matt(i, 52, mm)
    matsum(23, mm) = matsum(23, mm) + matt(i, 53, mm)
    matsum(24, mm) = matsum(24, mm) + matt(i, 54, mm)
    matsum(25, mm) = matsum(25, mm) + matt(i, 55, mm)
    'colateral entregado
    matsum(26, mm) = matsum(26, mm) + matt(i, 26, mm)
    matsum(27, mm) = matsum(27, mm) + matt(i, 27, mm)
    matsum(28, mm) = matsum(28, mm) + matt(i, 28, mm)
    matsum(29, mm) = matsum(29, mm) + matt(i, 29, mm)
    matsum(30, mm) = matsum(30, mm) + matt(i, 30, mm)
    'colateral recibido
    matsum(31, mm) = matsum(31, mm) + matt(i, 31, mm)
    matsum(32, mm) = matsum(32, mm) + matt(i, 32, mm)
    matsum(33, mm) = matsum(33, mm) + matt(i, 33, mm)
    matsum(34, mm) = matsum(34, mm) + matt(i, 34, mm)
    matsum(35, mm) = matsum(35, mm) + matt(i, 35, mm)
    
  Next i
Next mm
Open DirResVaR & "\Debug Reporte CCL " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1

ReDim matenc(1 To 55) As String
matenc(1) = "sd(p)"
matenc(6) = "xpd(p)"
matenc(11) = "ngr(p)"
matenc(16) = "Exposición"
matenc(21) = "ed(p)"
matenc(26) = "Colateral entregado"
matenc(31) = "Colateral recibido"
matenc(36) = "SDcc"
matenc(41) = "EDcc"
matenc(46) = "SDi"
matenc(51) = "END"


txtcadena = "CONTRAPARTE" & Chr(9)
For i = 1 To 55
txtcadena = txtcadena & matenc(i) & Chr(9)
Next i
Print #1, txtcadena
txtcadena = "0" & Chr(9)
     For i = 1 To 55
         txtcadena = txtcadena & i & Chr(9)
     Next i
     Print #1, txtcadena
     For i = 1 To noreg3
         txtcadena = matc(i, 1) & Chr(9)
         For j = 1 To 55
             txtcadena = txtcadena & matt(i, j, 1) & Chr(9)
         Next j
         Print #1, txtcadena
     Next i
Print #1, ""
For i = 1 To noreg3
    txtcadena = matc(i, 1) & Chr(9)
    For j = 1 To 55
        txtcadena = txtcadena & matt(i, j, 2) & Chr(9)
    Next j
    Print #1, txtcadena
Next i
Close #1

ReDim matrep(1 To 100, 1 To 25) As Variant

mata(1, 1) = "10359"
mata(1, 2) = "Forwards"
mata(2, 1) = "15671"
mata(2, 2) = "Forwards"
mata(3, 1) = "10360"
mata(3, 2) = "Opciones"
mata(4, 1) = "15672"
mata(4, 2) = "Opciones"
mata(5, 1) = "10361"
mata(5, 2) = "Swaps"
mata(6, 1) = "15673"
mata(6, 2) = "Swaps"
mata(7, 1) = "10362"
mata(7, 2) = "Derivados crediticios"
mata(8, 1) = "15674"
mata(8, 2) = "Derivados crediticios"
mata(9, 1) = "10363"
mata(9, 2) = "Operaciones estructuradas"
mata(10, 1) = "15675"
mata(10, 2) = "Operaciones estructuradas"
mata(11, 1) = "10364"
mata(11, 2) = "Paquetes de instrumentos derivados"
mata(12, 1) = "15676"
mata(12, 2) = "Paquetes de instrumentos derivados"
mata(13, 1) = "10365"
mata(13, 2) = "Total de flujos de salida a valor de mercado que pueden compensarse por formar parte de un contrato marco de compensación"
mata(14, 1) = "15677"
mata(14, 2) = "Total de flujos de entrada a valor de mercado que pueden compensarse por formar parte de un contrato marco de compensación"
mata(15, 1) = "10366"
mata(15, 2) = "Total de flujos de salida a valor de mercado compensados con los flujos de entrada a valor de mercado por formar parte de un contrato marco de compensación, en operaciones de derivados"
mata(16, 1) = "15678"
mata(16, 2) = "Total de flujos de entrada a valor de mercado compensados con los flujos de salida a valor de mercado por formar parte de un contrato marco de compensación, en operaciones de derivados"
mata(17, 1) = "10367"
mata(17, 2) = "Derivados por exposición potencial (add-on) para los cuales NO se tiene celebrado un contrato marco de compensación"
mata(18, 1) = "15679"
mata(18, 2) = "Total de flujos de entrada a valor de mercado para los cuales NO se tiene un contrato marco de compensación netos de las garantías de nivel 1, 2A y 2B recibidos  sobre los cuales se tengan derechos de uso o reutilización"
mata(19, 1) = "10368"
mata(19, 2) = "Derivados por exposición potencial (add-on) para los cuales se tenga celebrado un contrato marco de compensación"
mata(20, 1) = "15681"
mata(20, 2) = "Total de flujos de entrada a valor de mercado para los cuales se tiene un contrato marco de compensación netos de las garantías de nivel 1, 2A y 2B recibidos sobre los cuales se tengan derechos de uso o reutilización"
mata(21, 1) = "10369"
mata(21, 2) = "Total de flujos de salida a valor de mercado más flujos de salida por derivados por exposición potencial (add-on) para los cuales NO se tiene un contrato marco de compensación. Estos flujos deberán presentarse netos de las garantías de nivel 1, 2A y 2B entregados"
mata(22, 1) = "15682"
mata(22, 2) = "VALOR DE MERCADO DE LAS GARANTÍAS DE NIVEL 1, 2A Y 2B RECIBIDOS SOBRE LOS CUALES SE TENGAN DERECHOS DE USO O REUTILIZACIÓN"
mata(23, 1) = "10370"
mata(23, 2) = "Total de flujos de salida a valor de mercado compensados con los flujos de entrada a valor de mercado por formar parte de un contrato marco de compensación, más flujos de salida por derivados por exposición potencial (add-on) para los cuales se tenga celebrado un contrato marco de compensación. Estos flujos deberán presentarse netos de las garantías de nivel 1, 2A y 2B entregados"
mata(24, 1) = "15683"
mata(24, 2) = "Suma de las garantías de nivel 1, 2A y 2B recibidos por operaciones de derivados para las cuales NO se tiene celebrado un contrato marco de compensación"
mata(25, 1) = "10371"
mata(25, 2) = "VALOR DE MERCADO DE las garantías DE NIVEL 1, 2A Y 2B ENTREGADOS"
mata(26, 1) = "10561"
mata(26, 2) = "VALOR DE MERCADO DE LAS GARANTÍAS OTORGADOS EN OPERACIONES DE DERIVADOS Y EN OTRAS OPERACIONES"
ClaveBan = "037009"
ReDim mattabla(1 To 4, 1 To 7, 1 To 30) As Variant

For i = 1 To 26
For j = 1 To 4
mattabla(j, 1, i) = fecha
mattabla(j, 2, i) = ClaveBan
mattabla(j, 3, i) = ClaveBan
mattabla(j, 4, i) = mata(i, 1)
mattabla(j, 6, i) = 1
mattabla(j, 7, i) = 0
Next j
mattabla(1, 5, i) = "MXN"
mattabla(2, 5, i) = "USD"
mattabla(3, 5, i) = "EUR"
mattabla(4, 5, i) = "JPY"
Next i

For i = 1 To 4
    mattabla(i, 7, 1) = Round(matr(i, 2, 1) / 1000, 0)                                   'pasiva fwds
    mattabla(i, 7, 2) = Round(matr(i, 1, 1) / 1000, 0)                                   'activa fwds
    mattabla(i, 7, 5) = Round(matr(i, 2, 2) / 1000, 0)                                   'pasiva swaps
    mattabla(i, 7, 6) = Round(matr(i, 1, 2) / 1000, 0)                                   'activa swaps
    mattabla(i, 7, 13) = Round(matr(i, 2, 1) / 1000, 0) + Round(matr(i, 2, 2) / 1000, 0) 'pasivas swaps + fwds
    mattabla(i, 7, 14) = Round(matr(i, 1, 1) / 1000, 0) + Round(matr(i, 1, 2) / 1000, 0) 'activas swaps + fwds
    mattabla(i, 7, 15) = Round((matsum(5 + i, 1) + matsum(5 + i, 2)) / 1000, 0)          'pasivas swaps + fwds
    mattabla(i, 7, 16) = Round((matsum(10 + i, 1) + matsum(10 + i, 2)) / 1000, 0)        'pasivas swaps + fwds
    mattabla(i, 7, 19) = Round((matsum(i, 1) + matsum(i, 2)) / 1000, 0)                  'exposicion
    mattabla(i, 7, 20) = Round((matsum(20 + i, 1) + matsum(20 + i, 2)) / 1000, 0)        'suma END
    mattabla(i, 7, 22) = Round(matsum(30 + i, 1) / 1000, 0)                              'suma colateral recibido
    mattabla(i, 7, 23) = Round((matsum(15 + i, 1) + matsum(15 + i, 2)) / 1000, 0)        'suma SDi
    mattabla(i, 7, 25) = Round(matsum(25 + i, 1) / 1000, 0)                              'suma colateral entregado
    mattabla(i, 7, 26) = Round(matsum(25 + i, 1) / 1000, 0)                              'suma colateral entregado
Next i

nomarch = DirResVaR & "\Reporte CCL " & Format(fecha, "yyyy-mm-dd") & ".txt"
Open nomarch For Output As #1
For i = 1 To 26
    Print #1, mata(i, 2)
    Print #1, mata(i, 1)
    For j = 1 To 4
        txtcadena = ""
        For p = 1 To 7
            txtcadena = txtcadena & mattabla(j, p, i) & Chr(9)
        Next p
        Print #1, txtcadena
    Next j
    Print #1, ""
Next i
Close #1
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & "LIQ_REP_CCL" & " WHERE FECHA = " & txtfecha
ConAdo.Execute txtborra

For i = 1 To 26
    For j = 1 To 4
    txtinserta = "INSERT INTO " & "LIQ_REP_CCL" & " VALUES("
    txtinserta = txtinserta & txtfecha & ","
    txtinserta = txtinserta & mattabla(j, 2, i) & ","
    txtinserta = txtinserta & mattabla(j, 3, i) & ","
    txtinserta = txtinserta & mattabla(j, 4, i) & ","
    txtinserta = txtinserta & "'" & mattabla(j, 5, i) & "',"
    txtinserta = txtinserta & mattabla(j, 6, i) & ","
    txtinserta = txtinserta & mattabla(j, 7, i) & ")"
    ConAdo.Execute txtinserta
    Next j
Next i
Unload frmProgreso

End Sub

Function ObtTCCCL(ByVal fecha As Date, ByVal texto As String, matx1, matx2, matx3, matx4)
If texto = "DOLAR PIP FIX" Then
   ObtTCCCL = matx1(1, 2)
ElseIf texto = "UDI" Then
   ObtTCCCL = matx2(1, 2)
ElseIf texto = "EURO BM" Then
   ObtTCCCL = matx3(1, 2)
ElseIf texto = "YEN BM" Then
   ObtTCCCL = matx4(1, 2)
Else
   ObtTCCCL = 0
End If
End Function

Function ObtSaldoFwdFecha(ByVal fecha As Date, ByVal coperacion As String)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COPERACION,MNOCIONAL FROM " & TablaPosFwd & " WHERE COPERACION = '" & coperacion & "' GROUP BY COPERACION,MNOCIONAL"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
rmesa.Open txtfiltro, ConAdo
rmesa.MoveFirst
ReDim mata(1 To noreg) As Variant
For i = 1 To noreg
mata(i) = rmesa.Fields(1)
rmesa.MoveNext
Next i
rmesa.Close
ObtSaldoFwdFecha = mata(1)
End Function

Function ObtSaldoSwapFecha(ByVal fecha As Date, ByVal fechar As Date, ByVal coperacion As String, ByVal tpos As String)
Dim txtfecha As Date
Dim txtfechar As Date
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechar = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT FINICIO, FFINAL, NOCIONAL FROM " & TablaFlujosSwapsO & " WHERE COPERACION = '" & coperacion & "' AND FECHAREG = " & txtfechar & " AND TPATA = '" & tpos & "' AND TIPOPOS ='1' ORDER BY FINICIO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro, ConAdo
rmesa.MoveFirst
ReDim mata(1 To noreg, 1 To 3) As Variant
For i = 1 To noreg
    mata(i, 1) = rmesa.Fields(0)
    mata(i, 2) = rmesa.Fields(1)
    mata(i, 3) = rmesa.Fields(2)
rmesa.MoveNext
Next i
rmesa.Close
If fecha < mata(1, 1) Then
   ObtSaldoSwapFecha = mata(1, 3)
ElseIf fecha >= mata(noreg, 2) Then
      ObtSaldoSwapFecha = 0
Else
For i = 1 To noreg
    If fecha >= mata(i, 1) And fecha < mata(i, 2) Then
       ObtSaldoSwapFecha = mata(i, 3)
       Exit Function
    End If
Next i
End If
Else
  ObtSaldoSwapFecha = 0
End If
End Function

Function DetFechaReg(ByVal coperacion As String, ByVal fecha As Date)
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim i As Integer
Dim noreg As Integer
Dim RInterfIKOS As New ADODB.recordset

txtfiltro = "SELECT FECHAREG FROM " & TablaPosSwaps & " WHERE COPERACION = '" & coperacion & "' ORDER BY FECHAREG"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
RInterfIKOS.Open txtfiltro1, ConAdo
noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
RInterfIKOS.Open txtfiltro, ConAdo
RInterfIKOS.MoveFirst
ReDim mata(1 To noreg) As Variant
For i = 1 To noreg
mata(i) = RInterfIKOS.Fields(0)
RInterfIKOS.MoveNext
Next i
RInterfIKOS.Close
If fecha < mata(1) Then
   DetFechaReg = 0
   Exit Function
End If
If fecha >= mata(noreg) Then
   DetFechaReg = mata(noreg)
   Exit Function
End If
For i = 1 To noreg - 1
    If fecha >= mata(i) And fecha < mata(i + 1) Then
       DetFechaReg = mata(i)
       Exit Function
    End If
Next i
End Function

Function kpt(plazo, suby)
If suby = "Tasa" Then
   If plazo > 1800 Then
      kpt = 0.015
   ElseIf plazo > 360 Then
      kpt = 0.005
   ElseIf plazo >= 0 Then
      kpt = 0
   End If
ElseIf suby = "FX" Then
   If plazo > 1800 Then
      kpt = 0.075
   ElseIf plazo > 360 Then
      kpt = 0.05
   ElseIf plazo >= 0 Then
      kpt = 0.01
   End If
End If
End Function


Sub LeerEscEstresTaylor(ByVal fecha As Date, ByVal txtport As String)

    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noport As Long
    Dim noreg As Long
    Dim dxv        As Integer
    Dim indice     As Integer

    Dim matc()     As String
    Dim matd()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim txtcadfechas As String
    Dim txtcadesc As String
    Dim rmesa As New ADODB.recordset
    
    txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro2 = "SELECT SUBPORTAFOLIO from " & TablaResEstresAprox & " WHERE FECHA = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "' GROUP BY SUBPORTAFOLIO ORDER BY SUBPORTAFOLIO"
    txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noport = rmesa.Fields(0)
    rmesa.Close
    If noport <> 0 Then
       txtnomarch = DirResVaR & "\Resultados Esc estres Taylor " & Format$(fecha, "yyyy-mm-dd") & ".TXT"
       frmReportes.CommonDialog1.FileName = txtnomarch
       frmReportes.CommonDialog1.ShowSave
       txtnomarch = frmReportes.CommonDialog1.FileName
       Open txtnomarch For Output As #1
       rmesa.Open txtfiltro2, ConAdo
       ReDim matport(1 To noport, 1 To 1) As String
       For i = 1 To noport
           matport(i, 1) = rmesa.Fields(0)
           rmesa.MoveNext
       Next i
       rmesa.Close
       For i = 1 To noport
           txtfiltro2 = "SELECT * from " & TablaResEstresAprox & " WHERE FECHA = " & txtfecha
           txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
           txtfiltro2 = txtfiltro2 & " AND SUBPORTAFOLIO = '" & matport(i, 1) & "'"
           txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
           rmesa.Open txtfiltro1, ConAdo
           noreg = rmesa.Fields(0)
           rmesa.Close
           If noreg <> 0 Then
              rmesa.Open txtfiltro2, ConAdo
              txtcadfechas = ""
              txtcadesc = ""
              For j = 1 To noreg
                  txtcadfechas = txtcadfechas & "," & rmesa.Fields(6).GetChunk(rmesa.Fields(6).ActualSize)
                  txtcadesc = txtcadesc & "," & rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
                  matc = EncontrarSubCadenas(txtcadfechas, ",")
                  matd = EncontrarSubCadenas(txtcadesc, ",")
                  rmesa.MoveNext
                  AvanceProc = i / noreg
                  MensajeProc = "Leyendo las p&l del " & fecha & " " & Format$(AvanceProc, "##0.00 %")
              Next j
              rmesa.Close
              Print #1, matport(i, 1)
              Print #1, "fecha" & Chr(9) & "valor"
              For j = 1 To UBound(matc, 1)
                  Print #1, matc(j) & Chr(9) & matd(j)
              Next j
              Print #1, ""
           End If
       Next i
       Close #1
    Else
     MsgBox "no hay datos"
    End If

End Sub

Function GenCuadroEscEstres(ByVal fecha As Date, ByVal txtport As String) As Variant()
Dim noport As Integer
Dim noesc As Integer
Dim matportp() As String
Dim matesc() As String
Dim i As Integer
Dim j As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim fecha3 As Date

noport = UBound(MatPortSegRiesgo, 1)
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
fecha1 = DeterminaFechaTaylor(fecha, txtportCalc2, txtportBanobras, "Normal", 2)
fecha2 = DeterminaFechaTaylor(fecha, txtportCalc2, txtportBanobras, "Normal", 3)
fecha3 = DeterminaFechaTaylor(fecha, txtportCalc2, txtportBanobras, "Normal", 4)
For i = 1 To noport
    mats(i, 1) = CLng(fecha) & MatPortSegRiesgo(i, 1)
    mats(i, 2) = CLng(fecha)
    mats(i, 3) = i
    mats(i, 4) = MatPortSegRiesgo(i, 1)
    For j = 1 To noesc
        mats(i, j + 4) = LeerResEscEstres(fecha, txtport, MatPortSegRiesgo(i, 1), matesc(j))
    Next j
    mats(i, noesc + 5) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i, 1), fecha1)
    mats(i, noesc + 6) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i, 1), fecha2)
    mats(i, noesc + 7) = LeerEscEstresTaylor2(fecha, txtportCalc2, "Normal", MatPortSegRiesgo(i, 1), fecha3)
Next i
GenCuadroEscEstres = mats
End Function

Function DeterminaFechaTaylor(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal txtportfr As String, ByVal indice As Integer) As Date
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noreg As Long

    Dim matc()     As String
    Dim matd()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim txtcadfechas As String
    Dim txtcadesc As String
    Dim rmesa As New ADODB.recordset
    
    txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro2 = "SELECT * from " & TablaResEstresAprox & " WHERE FECHA = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND ESC_FR = '" & txtportfr & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORTAFOLIO = '" & txtsubport & "'"
    txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        rmesa.Open txtfiltro2, ConAdo
        For i = 1 To noreg
            txtcadfechas = txtcadfechas & "," & rmesa.Fields(6).GetChunk(rmesa.Fields(6).ActualSize)
            txtcadesc = txtcadesc & "," & rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo las p&l del " & fecha & " " & Format$(AvanceProc, "##0.00 %")
        Next i
        rmesa.Close
        matc = EncontrarSubCadenas(txtcadfechas, ",")
        matd = EncontrarSubCadenas(txtcadesc, ",")
        ReDim mate(1 To UBound(matd, 1), 1 To 2) As Variant
        For i = 1 To UBound(matd, 1)
            mate(i, 1) = CDate(ReemplazaVacioValor(matc(i), 0))
            mate(i, 2) = CDbl(ReemplazaVacioValor(matd(i), 0))
        Next i
        mate = RutinaOrden(mate, 2, SRutOrden)
        DeterminaFechaTaylor = mate(indice, 1)
        Exit Function
 
    End If
DeterminaFechaTaylor = 0
End Function

Function LeerPyGHistSubport(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal txtportfr As String, ByVal noesc As Long, ByVal htiempo As Integer, ByVal tesc As Integer)

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim noreg As Integer
Dim l As Integer
Dim matc() As String
Dim valt01 As Double
Dim valor As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format$(f_pos, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format$(f_factor, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha3 = "TO_DATE('" & Format$(f_val, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha3
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
   If tesc = 0 Then
      valt01 = 0
   ElseIf tesc = 1 Then
      valt01 = rmesa.Fields("VALT0")
   Else
     MsgBox "Opcion no valida"
   End If
   valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
   matc = EncontrarSubCadenas(valor, ",")
   For l = 1 To UBound(matc, 1)
        matv(l, 1) = valt01 + CDbl(matc(l))
   Next l
   rmesa.Close
Else
   ReDim matv(0 To 0, 0 To 0) As Double
End If
LeerPyGHistSubport = matv

End Function

Function CalcularCVaRMarginal(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noesc As Long
Dim htiempo As Integer
Dim nconf As Double
Dim matv1() As Double, matv2() As Double, matv3() As Double, matv4() As Double, matv5() As Double, matv6() As Double
Dim matd2() As Double, matd3() As Double, matd4() As Double, matd5() As Double, matd6() As Double
Dim valor1 As Double, valor2 As Double, valor3 As Double, valor4 As Double, valor5 As Double, valor6 As Double

noesc = 500
htiempo = 1
nconf = 0.03
matv1 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "CONSOLIDADO", "Normal", noesc, htiempo, 0)
matv2 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "MERCADO DE DINERO", "Normal", noesc, htiempo, 0)
matv3 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "MESA DE CAMBIOS", "Normal", noesc, htiempo, 0)
matv4 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "DERIVADOS DE NEGOCIACION", "Normal", noesc, htiempo, 0)
matv5 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "DERIVADOS ESTRUCTURALES", "Normal", noesc, htiempo, 0)
matv6 = LeerPyGHistSubport(fecha, fecha, fecha, "TOTAL", "DERIVADOS NEGOCIACION RECLASIFICACION", "Normal", noesc, htiempo, 0)
If UBound(matv2, 1) <> 0 Then
   matd2 = MResta(matv1, matv2)
Else
   matd2 = matv1
End If
If UBound(matv3, 1) <> 0 Then
   matd3 = MResta(matv1, matv3)
Else
   matd3 = matv1
End If
If UBound(matv4, 1) <> 0 Then
   matd4 = MResta(matv1, matv4)
Else
   matd4 = matv1
End If
If UBound(matv5, 1) <> 0 Then
   matd5 = MResta(matv1, matv5)
Else
   matd5 = matv1
End If
If UBound(matv6, 1) <> 0 Then
   matd6 = MResta(matv1, matv6)
Else
   matd6 = matv1
End If
valor1 = CPercentilCVaR(nconf, matv1, 0, 0, True)
valor2 = CPercentilCVaR(nconf, matd2, 0, 0, True)
valor3 = CPercentilCVaR(nconf, matd3, 0, 0, True)
valor4 = CPercentilCVaR(nconf, matd4, 0, 0, True)
valor5 = CPercentilCVaR(nconf, matd5, 0, 0, True)
valor6 = CPercentilCVaR(nconf, matd6, 0, 0, True)
ReDim mata(1 To 1, 1 To 6) As Variant
mata(1, 1) = fecha
mata(1, 2) = valor1 - valor2
mata(1, 3) = valor1 - valor3
mata(1, 4) = valor1 - valor4
mata(1, 5) = valor1 - valor5
mata(1, 6) = valor1 - valor6

CalcularCVaRMarginal = mata
End Function

Function LeerResEscEstres(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal txtestres As String) As Double
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim valor As Double
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaResEscEstresPort & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_ESTRES = '" & txtestres & "'"
txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor = rmesa.Fields("VALOR")
   rmesa.Close
Else
   valor = 0
End If
LeerResEscEstres = valor
End Function

Function LeerEscEstresTaylor2(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal fechax As Date) As Double
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noreg As Long
    Dim dxv        As Integer
    Dim indice     As Integer
    Dim matc()     As String
    Dim matd()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim txtcadfechas As String
    Dim txtcadesc As String
    Dim rmesa As New ADODB.recordset
    
    txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro2 = "SELECT * from " & TablaResEstresAprox & " WHERE FECHA = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND ESC_FR = '" & txtportfr & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORTAFOLIO = '" & txtsubport & "'"
    txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        rmesa.Open txtfiltro2, ConAdo
        For i = 1 To noreg
            txtcadfechas = txtcadfechas & "," & rmesa.Fields(6).GetChunk(rmesa.Fields(6).ActualSize)
            txtcadesc = txtcadesc & "," & rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo las p&l del " & fecha & " " & Format$(AvanceProc, "##0.00 %")
        Next i
        rmesa.Close
        matc = EncontrarSubCadenas(txtcadfechas, ",")
        matd = EncontrarSubCadenas(txtcadesc, ",")
        For i = 1 To UBound(matc, 1)
            If CDate(ReemplazaVacioValor(matc(i), 0)) = fechax And Not EsVariableVacia(matd(i)) Then
               LeerEscEstresTaylor2 = CDbl(matd(i))
               Exit Function
            End If
        Next i
    End If
LeerEscEstresTaylor2 = 0
End Function

Function LeerResPIDVDer(ByVal fecha As Date)
Dim noreg0 As Integer
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtport As String
Dim i As Integer
Dim j As Integer
Dim matem() As Variant
Dim noesc As Integer
Dim htiempo As Integer
Dim rmesa As New ADODB.recordset

noesc = 500
htiempo = 1

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosPIDV & " AND TIPOPOS = 1 GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matem(1 To noreg + 1, 1 To 5) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       matem(i, 1) = CLng(fecha) & "_" & Format(i, "00")
       matem(i, 2) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT COPERACION FROM " & TablaPosSwaps & " WHERE C_EM_PIDV = '" & matem(i, 2) & "'"
       txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY COPERACION ORDER BY COPERACION"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg0 = rmesa.Fields(0)
       rmesa.Close
       txtcadena = "' "
       If noreg0 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          For j = 1 To noreg0
              txtcadena = txtcadena & rmesa.Fields("COPERACION") & ","
              rmesa.MoveNext
          Next j
          rmesa.Close
          If Len(txtcadena) <> 0 Then
              If Right(txtcadena, 1) = "," Then txtcadena = Left(txtcadena, Len(txtcadena) - 1)
          End If
       End If
       matem(i, 3) = txtcadena
       txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha & " AND SUBPORT =  'PIDV " & matem(i, 2) & "'"
       txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg0 = rmesa.Fields(0)
       rmesa.Close
       If noreg0 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          matem(i, 4) = rmesa.Fields("MTM_SUCIO")
          rmesa.Close
       End If
       txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND SUBPORT =  'PIDV " & matem(i, 2) & "+DERIV'"
       txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH' AND NCONF = .03"
       txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg0 = rmesa.Fields(0)
       rmesa.Close
       If noreg0 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          matem(i, 5) = rmesa.Fields("VALOR")
          rmesa.Close
       End If
   Next i
   txtport = "PI DISPONIBLES PARA LA VENTA"
   matem(noreg + 1, 1) = CLng(fecha) & "_00"
   matem(noreg + 1, 2) = " "
   txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha & " AND SUBPORT = '" & txtport & "'"
   txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg0 = rmesa.Fields(0)
   rmesa.Close
   If noreg0 <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      matem(noreg + 1, 4) = rmesa.Fields("MTM_SUCIO")
      rmesa.Close
   End If
   txtport = "PIDV+DERIVADOS"
   txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND SUBPORT =  '" & txtport & "'"
   txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH' AND NCONF = 0.03"
   txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg0 = rmesa.Fields(0)
   rmesa.Close
   If noreg0 <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      matem(noreg + 1, 5) = rmesa.Fields("VALOR")
      rmesa.Close
   End If
Else
ReDim matem(0 To 0, 0 To 0) As Variant
End If
LeerResPIDVDer = matem
End Function

Function RepDerivPI(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim noreg1 As Long
Dim contar As Long
Dim i As Long
Dim j As Long
Dim matem() As String
Dim txtcadena As String
Dim txtsubport As String
Dim rmesa As New ADODB.recordset
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtmsg As String
Dim exito As Boolean
Dim matpl() As Double
Dim indice2 As Long
Dim txtmonpas As String
Dim califs As String
Dim escala As String
Dim saldo As Double
Dim moneda As Double
Dim matm() As Variant

txtport = "TOTAL"
noesc = 500
htiempo = 1
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION,EMISION,CPOSICION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION,EMISION,CPOSICION,TOPERACION ORDER BY CPOSICION,C_EMISION,TOPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 4) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = rmesa.Fields("EMISION")
       matem(i, 3) = rmesa.Fields("CPOSICION")
       matem(i, 4) = rmesa.Fields("TOPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   ReDim mata(1 To 19, 1 To 1) As Variant
   contar = 0
   For i = 1 To UBound(matem, 1)
       txtsubport = "DERIVADOS PIDV " & matem(i, 1)
       txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT =" & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='" & txtsubport & "' AND CPOSICION = " & ClavePosDeriv
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          ReDim Preserve mata(1 To 19, 1 To contar + noreg1) As Variant
          For j = 1 To noreg1
              mata(1, contar + j) = fecha
              mata(2, contar + j) = contar + j
              mata(3, contar + j) = CLng(mata(1, contar + j)) & " " & mata(2, contar + j)
              mata(4, contar + j) = rmesa.Fields("FECHAREG")
              mata(5, contar + j) = rmesa.Fields("COPERACION")
              rmesa.MoveNext
          Next j
          rmesa.Close
          For j = 1 To noreg1
              txtfecha1 = "TO_DATE('" & Format$(mata(4, contar + j), "DD/MM/YYYY") & "','DD/MM/YYYY')"
              txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 1"
              txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha1
              txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & mata(5, contar + j) & "'"
              rmesa.Open txtfiltro2, ConAdo
              mata(6, contar + j) = rmesa.Fields("ID_CONTRAP")
              mata(7, contar + j) = DeterminaContraparte(mata(6, contar + j))
              mata(8, contar + j) = rmesa.Fields("TC_ACTIVA")
              mata(9, contar + j) = rmesa.Fields("ST_ACTIVA")
              mata(10, contar + j) = ObtTasaSwapFecha(mata(5, contar + j), mata(4, contar + j), fecha, "B")
              mata(11, contar + j) = ObtTasaSwapFecha(mata(5, contar + j), mata(4, contar + j), fecha, "C")
              mata(12, contar + j) = rmesa.Fields("FVALUACION")
              indice2 = BuscarValorArray(mata(12, contar + j), MatTValSwaps1, 1)
              If indice2 <> 0 Then
                 txtmonpas = ReemplazaVacioValor(MatTValSwaps1(indice2, 14), "")
              End If
              saldo = ObtFlujoSwapFecha(mata(5, contar + j), mata(4, contar + j), fecha, "C")
              If txtmonpas = "DOLAR PIP FIX" Then
                 matm = Leer1FactorR(fecha, fecha, "DOLAR PIP FIX", 0)
                 moneda = matm(1, 2)
              ElseIf txtmonpas = "EURO PIP" Then
                 matm = Leer1FactorR(fecha, fecha, "EURO PIP", 0)
                 moneda = matm(1, 2)
              ElseIf txtmonpas = "UDI" Then
                 matm = Leer1FactorR(fecha, fecha, "UDI", 0)
                 moneda = matm(1, 2)
              ElseIf EsVariableVacia(txtmonpas) Then
                 moneda = 1
              Else
                 MsgBox "No se definio la moneda" & txtmonpas
              End If
              mata(13, contar + j) = saldo * moneda
              mata(14, contar + j) = rmesa.Fields("FVENCIMIENTO")
              rmesa.Close
              Call DeterminaCalifContrap(fecha, mata(6, contar + j), califs, escala)
              mata(15, contar + j) = califs
          Next j
          For j = 1 To noreg1
              txtfiltro2 = "SELECT * FROM " & TablaValPos & " WHERE TIPOPOS = 1"
              txtfiltro2 = txtfiltro2 & " AND FECHAP = " & txtfecha
              txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosDeriv
              txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & mata(5, contar + j) & "'"
              rmesa.Open txtfiltro2, ConAdo
              mata(16, contar + j) = rmesa.Fields("MTM_S")
              rmesa.Close
              matpl = LeerPyG1Oper(fecha, txtport, "Normal", ClavePosDeriv, mata(5, contar + j), noesc, htiempo)
              mata(17, contar + j) = CPercentilCVaR(0.03, matpl, 0, 0, True)
          Next j
          For j = 1 To noreg1
              If j = 1 Then
                 txtsubport = "PIDV " & matem(i, 1) & "+DERIV"
                 mata(18, contar + j) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, 0.03, 0, "CVARH", exito)
                 mata(19, contar + j) = matem(i, 1)
              Else
                 mata(18, contar + j) = 0
                 mata(19, contar + j) = "  "
              End If
          Next j
          contar = contar + noreg1
        End If
   Next i
   RepDerivPI = MTranV(mata)
End If
End Function

Sub ValidarValPosMD(ByVal fecha As Date, ByRef mata() As Variant)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset
Dim nodif As Long
Dim valsiv As Double
Dim valpip As Double
Dim valdif As Double

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT A.CPOSICION,A.COPERACION,A.P_SUCIO,A.VAL_PIP_S,A.NO_TITULOS_,C.TV,C.EMISION,C.SERIE from " & TablaValPos & " A"
txtfiltro2 = txtfiltro2 & " LEFT JOIN (SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
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
nodif = 0
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst

   For i = 1 To noreg
       valsiv = rmesa.Fields("P_SUCIO")
       valpip = rmesa.Fields("VAL_PIP_S")
       valdif = Abs(valsiv - valpip)
       If valdif > 0.001 Then
          nodif = nodif + 1
          ReDim mata(1 To 9, 1 To nodif) As Variant
          mata(1, nodif) = rmesa.Fields("CPOSICION")
          mata(2, nodif) = rmesa.Fields("COPERACION")
          mata(3, nodif) = rmesa.Fields("TV")
          mata(4, nodif) = rmesa.Fields("EMISION")
          mata(5, nodif) = rmesa.Fields("SERIE")
          mata(6, nodif) = rmesa.Fields("NO_TITULOS_")
          mata(7, nodif) = rmesa.Fields("P_SUCIO")
          mata(8, nodif) = rmesa.Fields("VAL_PIP_S")
          mata(9, nodif) = Abs(mata(7, nodif) - mata(8, nodif))
       End If
       rmesa.MoveNext
   Next i
   mata = MTranV(mata)
   rmesa.Close
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
End If

End Sub

Function CrearCadSQLValPosDeriv(ByVal fecha As Date, ByVal txtport As String, ByVal id_pos As Integer, ByVal id_val As Integer)
Dim txtfecha As String
Dim txtfiltro As String

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT A.COPERACION,A.INTENCION,A.MTM_S,A.VAL_ACT_S,A.VAL_PAS_S,"
txtfiltro = txtfiltro & "A.CPOSICION,A.MTM_IKOS,A.VAL_ACT_IKOS,A.VAL_PAS_IKOS,B.FINICIO,B.FVENCIMIENTO,B.CPRODUCTO,B.FVALUACION,B.ID_CONTRAP"
txtfiltro = txtfiltro & " from " & TablaValPos & " A"
txtfiltro = txtfiltro & " JOIN " & TablaPosSwaps & " B"
txtfiltro = txtfiltro & " ON A.FECHAREG = B.FECHAREG "
txtfiltro = txtfiltro & " AND A.CPOSICION = B.CPOSICION "
txtfiltro = txtfiltro & " AND A.COPERACION = B.COPERACION "
txtfiltro = txtfiltro & " WHERE A.FECHAP = " & txtfecha
txtfiltro = txtfiltro & " AND A.FECHAFR = " & txtfecha
txtfiltro = txtfiltro & " AND A.FECHAV = " & txtfecha
txtfiltro = txtfiltro & " AND A.PORTAFOLIO = '" & txtport & "'"
txtfiltro = txtfiltro & " AND A.CPOSICION = " & id_pos
txtfiltro = txtfiltro & " AND A.ID_VALUACION = " & id_val
txtfiltro = txtfiltro & " AND B.TIPOPOS = 1"
txtfiltro = txtfiltro & " UNION "
txtfiltro = txtfiltro & " SELECT A.COPERACION,A.INTENCION,A.MTM_S,A.VAL_ACT_S,A.VAL_PAS_S,"
txtfiltro = txtfiltro & "A.CPOSICION,A.MTM_IKOS,A.VAL_ACT_IKOS,A.VAL_PAS_IKOS,B.FINICIO,B.FVENCIMIENTO,B.CPRODUCTO,B.FVALUACION,B.ID_CONTRAP"
txtfiltro = txtfiltro & " FROM " & TablaValPos & " A"
txtfiltro = txtfiltro & " JOIN ("
txtfiltro = txtfiltro & "SELECT FECHAREG,CPOSICION,COPERACION,FINICIO,FVENCIMIENTO,CPRODUCTO,"
txtfiltro = txtfiltro & "CPRODUCTO AS FVALUACION,ID_CONTRAP FROM " & TablaPosFwd & " WHERE TIPOPOS = 1) B"
txtfiltro = txtfiltro & " ON A.FECHAREG = B.FECHAREG "
txtfiltro = txtfiltro & " AND A.CPOSICION = B.CPOSICION "
txtfiltro = txtfiltro & " AND A.COPERACION = B.COPERACION "
txtfiltro = txtfiltro & " WHERE A.FECHAP = " & txtfecha
txtfiltro = txtfiltro & " AND A.FECHAFR = " & txtfecha
txtfiltro = txtfiltro & " AND A.FECHAV = " & txtfecha
txtfiltro = txtfiltro & " AND A.PORTAFOLIO = '" & txtport & "'"
txtfiltro = txtfiltro & " AND A.CPOSICION = " & id_pos
txtfiltro = txtfiltro & " AND A.ID_VALUACION = " & id_val
txtfiltro = txtfiltro & " ORDER BY COPERACION"

CrearCadSQLValPosDeriv = txtfiltro
End Function
