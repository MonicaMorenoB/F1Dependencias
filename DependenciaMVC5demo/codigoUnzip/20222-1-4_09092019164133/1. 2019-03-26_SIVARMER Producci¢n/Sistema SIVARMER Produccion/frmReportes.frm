VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportes 
   Caption         =   "Reportes de SIVARMER"
   ClientHeight    =   8175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15345
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   510
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mReporte50 
      Caption         =   "Reportes de factores de riesgo"
      Begin VB.Menu mReporte51 
         Caption         =   "Analisis de factores del dia"
      End
      Begin VB.Menu mReporte52 
         Caption         =   "Extracción de curvas"
      End
      Begin VB.Menu mReporte53 
         Caption         =   "Factores de riesgo por nodos"
      End
      Begin VB.Menu mHistoria 
         Caption         =   "Datos del vector de precios"
      End
   End
   Begin VB.Menu mReporte1 
      Caption         =   "Reportes de posición"
      Begin VB.Menu mAnaMovDeuda 
         Caption         =   "Analisis de movimientos de la posición de Dueda"
      End
      Begin VB.Menu mReporte101 
         Caption         =   "Caracteristicas y flujos de Swaps"
      End
      Begin VB.Menu mReporte02 
         Caption         =   "Caracteristicas de la posicion de fwds"
      End
      Begin VB.Menu mReporte003 
         Caption         =   "Caracteristicas y flujos de posiciones primarias"
      End
      Begin VB.Menu mReporteROD 
         Caption         =   "Resumen de operaciones de derivados"
      End
   End
   Begin VB.Menu mReporte2 
      Caption         =   "Reportes de Valuacion"
      Begin VB.Menu mreportecsp 
         Caption         =   "Comparacion valuación MD SIVARMER vs. PIP"
      End
      Begin VB.Menu mReporteCVSP2065 
         Caption         =   "Comparación valuación Pensiones SIVARMER vs. PIP"
      End
      Begin VB.Menu mcompvalsi 
         Caption         =   "Valuacion Derivados SIVARMER vs. IKOS"
      End
      Begin VB.Menu mvalDerivTC 
         Caption         =   "Valuacion de derivados por moneda"
      End
      Begin VB.Menu mReporte103 
         Caption         =   "Valuación por contraparte"
      End
      Begin VB.Menu mReporteFVS 
         Caption         =   "Flujos valuados de swaps"
      End
      Begin VB.Menu mReporte22 
         Caption         =   "Duracion de la posicion"
      End
      Begin VB.Menu mReporteTCCF 
         Caption         =   "Tasa de cupones que cortan a una fecha"
      End
      Begin VB.Menu mSaldoDeriv 
         Caption         =   "Saldo de Derivados"
      End
   End
   Begin VB.Menu mReporte3 
      Caption         =   "Reportes de VaR"
      Begin VB.Menu mreporte11 
         Caption         =   "VaR de la posicion de Mercado de Dinero"
      End
      Begin VB.Menu mReporteVMD 
         Caption         =   "VaR de la Mesa de Dinero"
      End
      Begin VB.Menu mReporteVT 
         Caption         =   "VaR de la Tesoreria"
      End
      Begin VB.Menu mReporte33 
         Caption         =   "VaR de la posición de Derivados"
      End
      Begin VB.Menu mReporte55 
         Caption         =   "Escenarios de simulacion historica por instrumento"
      End
      Begin VB.Menu mReporte100 
         Caption         =   "Escenarios de simulacion historica por subportafolios"
      End
      Begin VB.Menu mReporte200 
         Caption         =   "Escenarios de simulación Montecarlo por instrumento"
      End
      Begin VB.Menu mReporte300 
         Caption         =   "Escenarios de simulación Montecarlo por subportafolios"
      End
      Begin VB.Menu mReporteHVS 
         Caption         =   "Histórico del CVaR por subportafolios"
      End
      Begin VB.Menu mReporte301 
         Caption         =   "Sensibilidades calculadas"
      End
      Begin VB.Menu mRiesgoF 
         Caption         =   "Cuadro riesgo por factor"
      End
      Begin VB.Menu mReporteEET 
         Caption         =   "Escenarios de estres de Taylor"
      End
      Begin VB.Menu mReporteBt 
         Caption         =   "Backtesting"
      End
      Begin VB.Menu mEscHisBack 
         Caption         =   "Escenarios historicos port Banobras"
      End
      Begin VB.Menu mExpChol 
         Caption         =   "Exportar Matriz Choleski"
      End
   End
   Begin VB.Menu mReporte4 
      Caption         =   "Reportes de CVA"
      Begin VB.Menu mReporte123 
         Caption         =   "Reporte de CVA"
      End
      Begin VB.Menu mgenrepEPE 
         Caption         =   "Generar reportes EPE"
      End
      Begin VB.Menu mReporteWRW 
         Caption         =   "Reporte de Wrong Risk Way"
      End
      Begin VB.Menu mMatTran 
         Caption         =   "Matrices de transicion"
      End
      Begin VB.Menu mReporte201 
         Caption         =   "Generar p&g de proceso de CVA"
      End
      Begin VB.Menu mExppgcva1 
         Caption         =   "Exportar pyg CVA 1 contraparte"
      End
      Begin VB.Menu mexppygmd 
         Caption         =   "Exportar pyg CVA Deuda"
      End
   End
   Begin VB.Menu mReporte5 
      Caption         =   "Reportes de efectividad"
      Begin VB.Menu mRepEfPros 
         Caption         =   "Reportes de eficiencia prospectiva"
      End
      Begin VB.Menu mReporteER 
         Caption         =   "Reporte de efectividad retrospectiva"
      End
   End
   Begin VB.Menu mReporte6 
      Caption         =   "Reportes Varios"
      Begin VB.Menu msql 
         Caption         =   "Exportacion de Cadenas sql de poscion"
      End
      Begin VB.Menu mReporteVR 
         Caption         =   "Valor de reemplazo"
      End
      Begin VB.Menu mLimCont 
         Caption         =   "Límites de contraparte de swaps"
      End
      Begin VB.Menu mLimContFwd 
         Caption         =   "Limites de contraparte de fwds de tipo de cambio"
      End
      Begin VB.Menu mDetalleLContrap 
         Caption         =   "Detalle de calculo de limites de contraparte"
      End
      Begin VB.Menu mReporteERRE 
         Caption         =   "Exportar resultados a resumen Ejecutivo.xlsx"
      End
      Begin VB.Menu mReporteERtxt 
         Caption         =   "Exportar resultados a datos resumen ejecutivo.txt"
      End
      Begin VB.Menu mExpResPI 
         Caption         =   "Exportar resultados a Resumen PI"
      End
      Begin VB.Menu mRepValProc 
         Caption         =   "Reporte de validacion de procesos"
      End
      Begin VB.Menu mExpCurvas 
         Caption         =   "Exportar curvas a resumen ejecutivo"
      End
      Begin VB.Menu mObCurvasRE 
         Caption         =   "Obtener curvas Resumen ejecutivo"
      End
      Begin VB.Menu mGenComenRE 
         Caption         =   "Generar comentario Resumen Ejecutivo"
      End
      Begin VB.Menu mExportCAIR 
         Caption         =   "Exportar datos de Reporte CAIR"
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Unload(Cancel As Integer)
SiActTProc = False
End Sub

Private Sub mAnaMovDeuda_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmAnalisisPosMD.Show
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mcompvalsi_Click()
Dim fecha As Date
Dim txtfecha As String
'ahora se procede a hacer un reporte para su envio a tesoreria
Screen.MousePointer = 11
 txtfecha = InputBox("Dame la fecha de la valuación ", , Date)
 If IsDate(txtfecha) Then
    fecha = CDate(txtfecha)
    Call GenRepValDer(fecha, 1)
    Call GenRepValDer(fecha, 2)
 End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mDetalleLContrap_Click()
Dim fecha As Date
Dim txtfecha As String
Dim coperacion As String
txtfecha = InputBox("dame la fecha a cargar", , Date)
coperacion = InputBox("Dame la clave de operacion", , "000")
If IsDate(txtfecha) Then
Screen.MousePointer = 11
frmProgreso.Show
  fecha = CDate(txtfecha)
  'Call RepDetalleCLimiteC2(fecha, coperacion, TablaLimContrap1)
  Call RepDetalleCLimiteC(fecha, coperacion, TablaLimContrap1)
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End If
End Sub

Private Sub mEscHisBack_Click()
Dim tfecha1 As String
Dim tfecha2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim txtcadena As String
Dim txtcadena1 As String
Dim mats() As String
Dim noesc As Integer
Dim rmesa As New ADODB.recordset
Dim txtnomarch As String

noesc = 500
Screen.MousePointer = 11
tfecha1 = InputBox("Dame la fecha inicial ", , Date)
tfecha2 = InputBox("Dame la fecha final ", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
   fecha1 = CDate(tfecha1)
   fecha2 = CDate(tfecha2)
   txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtfiltro2 = "select * FROM " & TablaPLEscHistPort & " WHERE F_POSICION >= " & txtfecha1
   txtfiltro2 = txtfiltro2 & " AND F_POSICION <= " & txtfecha2
   txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = 'TOTAL'"
   txtfiltro2 = txtfiltro2 & " AND SUBPORT = 'CONSOLIDADO'"
   txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES ='Normal' and noesc = " & noesc & " AND HTIEMPO = 1"
   txtfiltro2 = txtfiltro2 & " ORDER BY F_POSICION"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      txtnomarch = "D:\esc hist banobras " & Format(fecha2, "yyyy-mm-dd") & ".TXT"
      frmReportes.CommonDialog1.FileName = txtnomarch
      frmReportes.CommonDialog1.ShowSave
      txtnomarch = frmReportes.CommonDialog1.FileName
      Open txtnomarch For Output As #1
      rmesa.Open txtfiltro2, ConAdo
      For i = 1 To noreg
          txtcadena = rmesa.Fields("F_POSICION") & Chr(9)
          txtcadena1 = rmesa.Fields("DATOS")
          mats = EncontrarSubCadenas(txtcadena1, ",")
          For j = 1 To UBound(mats, 1)
              txtcadena = txtcadena & mats(j) & Chr(9)
          Next j
          Print #1, txtcadena
          rmesa.MoveNext
      Next i
      rmesa.Close
      Close #1
   Else
      MsgBox "No se encontraron registros"
   End If
End If
Screen.MousePointer = 0
MsgBox "Fin del proceso"
End Sub

Private Sub mExpChol_Click()
Dim fecha As Date
Dim tfecha As String
Dim txtport As String
Dim i As Integer, j As Integer
Dim txtcadena As String
Dim nomarch As String

tfecha = InputBox("Dame la fecha de calculo", , Date)
txtport = InputBox("Dame la fecha de calculo", , "NEGOCIACION + INVERSION")
If IsDate(tfecha) And Not EsVariableVacia(txtport) Then
   fecha = CDate(tfecha)
   Screen.MousePointer = 11
   Call LeerMatCholeski(fecha, txtport, matordenMont, mmediasMont, MatCholeski)
   If UBound(matordenMont, 1) <> 0 Then
   nomarch = DirResVaR & "\choleski " & txtport & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmReportes.CommonDialog1.FileName = nomarch
   frmReportes.CommonDialog1.ShowSave
   nomarch = CommonDialog1.FileName
   Open nomarch For Output As #1
      txtcadena = ""
      For i = 1 To UBound(matordenMont, 1)
        txtcadena = txtcadena & matordenMont(i, 1) & Chr(9)
      Next i
      Print #1, txtcadena
      Print #1, ""
      txtcadena = ""
      For i = 1 To UBound(mmediasMont, 1)
        txtcadena = txtcadena & mmediasMont(i, 1) & Chr(9)
      Next i
      Print #1, txtcadena
      Print #1, ""
      For i = 1 To UBound(MatCholeski, 1)
          txtcadena = ""
          For j = 1 To UBound(MatCholeski, 2)
              txtcadena = txtcadena & MatCholeski(i, j) & Chr(9)
          Next j
          Print #1, txtcadena
      Next i
      Print #1, ""
     
   Close #1
   End If
   Screen.MousePointer = 0
   MsgBox "Fin de proceso"
End If
End Sub

Private Sub mExpCurvas_Click()
Dim txtfecha As String
Dim strconexion As String
Dim fecha As Date
Dim mcurvas() As Variant
Dim conadoex As New ADODB.Connection
SiActTProc = True
Screen.MousePointer = 11
txtfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(txtfecha) Then
   frmProgreso.Show
   fecha = CDate(txtfecha)
   mcurvas = LeerCurvasRE(fecha)
   If UBound(mcurvas, 1) <> 0 Then
      strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DirReportes & "\Reporte Ejecutivo 2018.xlsx"
      strconexion = strconexion & ";Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
      conadoex.ConnectionString = strconexion
      conadoex.Open
      Call ActCurvasRE(mcurvas, "Curvas$", conadoex, RegExcel)
      conadoex.Close
   End If
   Unload frmProgreso
End If
SiActTProc = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mExportCAIR_Click()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim tfecha1 As String
Dim tfecha2 As String
Dim mata() As Variant
Dim matv1() As Variant
Dim matv2() As Variant
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim matres() As Variant

Screen.MousePointer = 11
tfecha1 = InputBox("Dame la fecha del mes anterior", , Date)
tfecha2 = InputBox("Dame la fecha del mes de cierre", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
fecha1 = CDate(tfecha1)
fecha2 = CDate(tfecha2)
mata = ObtSensibCAIR(fecha1, fecha2, 500)
Open "d:\Resultados CAIR  " & Format(fecha2, "yyyy-mm-dd") & ".txt" For Output As #1
Print #1, "Sensibilidades"
Print #1, "Factor de riesgo" & Chr(9) & fecha2 & Chr(9) & fecha1 & Chr(9) & "Variación"
For i = 1 To UBound(mata, 1)
    txtcadena = mata(i, 5) & Chr(9)
    txtcadena = txtcadena & mata(i, 6) / 4 & Chr(9)
    txtcadena = txtcadena & mata(i, 9) / 4 & Chr(9)
    txtcadena = txtcadena & (mata(i, 6) - mata(i, 9)) / 4 & Chr(9)
    Print #1, txtcadena
Next i
Print #1, ""
Print #1, "Factor de riesgo" & Chr(9) & "Mayor cambio negativo en 1 año PB" & Chr(9) & "Sensibilidad mdp" & Chr(9) & "Mayor cambio positivo en 1 año PB" & Chr(9) & "Sensibilidad mdp"
For i = 1 To UBound(mata, 1)
    txtcadena = mata(i, 5) & Chr(9)
    txtcadena = txtcadena & mata(i, 7) * 10000 & Chr(9)
    txtcadena = txtcadena & mata(i, 6) * mata(i, 7) * 100 & Chr(9)
    txtcadena = txtcadena & mata(i, 8) * 10000 & Chr(9)
    txtcadena = txtcadena & mata(i, 6) * mata(i, 8) * 100 & Chr(9)
    Print #1, txtcadena
Next i
Print #1, ""
matv1 = ObtVolatilCAIR(fecha1, fecha2)
matv2 = ObtVolatilCAIR2(fecha1, fecha2)
Print #1, "Volatilidades"
Print #1, "Factor de riesgo" & Chr(9) & fecha2 & Chr(9) & fecha1
For i = 1 To UBound(matv1, 1)
    txtcadena = matv1(i, 3) & Chr(9)
    txtcadena = txtcadena & matv1(i, 4) & Chr(9)
    txtcadena = txtcadena & matv1(i, 5) & Chr(9)
    Print #1, txtcadena
Next i
For i = 1 To Minimo(UBound(matv2, 1), 10)
    txtcadena = matv2(i, 3) & Chr(9)
    txtcadena = txtcadena & matv2(i, 4) & Chr(9)
    txtcadena = txtcadena & matv2(i, 5) & Chr(9)
    Print #1, txtcadena
Next i
Print #1, ""
matres = ObtenerLimMarcoOp(fecha1, fecha2)
If UBound(matres, 1) <> 0 Then
   Print #1, "Limites de Marco de operación"
   For i = 1 To UBound(matres, 1)
       txtcadena = ""
       For j = 1 To UBound(matres, 2) - 1
           txtcadena = txtcadena & matres(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
End If
matres = DetermPlazoPromMO(fecha1, fecha2)
If UBound(matres, 1) <> 0 Then
   Print #1, "Plazos promedio MO"
   For i = 1 To UBound(matres, 1)
       txtcadena = ""
       For j = 1 To UBound(matres, 2)
           txtcadena = txtcadena & matres(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
End If
matres = DeterminaPorcCalif(fecha2)
If UBound(matres, 1) <> 0 Then
   Print #1, "Distribucion de instrumentos por calificacion en MO"
   For i = 1 To UBound(matres, 1)
       Print #1, matres(i, 1) & Chr(9) & matres(i, 2)
   Next i
End If
MsgBox "Fin de proceso"


Close #1
End If
MsgBox "Fin del proceso"
Screen.MousePointer = 0
End Sub

Function ObtSensibCAIR(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal noesc As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fecha0 As Date
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim j As Integer
Dim indice As Integer
Dim matf() As Variant
Dim valmax As Double
Dim valmin As Double
Dim rmesa As New ADODB.recordset

txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT FACTOR,CURVA, PLAZO, DESCRIPCION, DERIVADA/100000000 AS DERIV FROM"
txtfiltro2 = txtfiltro2 & " (SELECT * FROM " & TablaSensibPort & " WHERE FECHA= " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal' AND PORTAFOLIO='NEGOCIACION + INVERSION'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT='CONSOLIDADO' AND TVALOR<>'T CAMBIO'"
txtfiltro2 = txtfiltro2 & " ORDER BY ABS(DERIVADA) DESC) WHERE ROWNUM<11 ORDER BY DERIV DESC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 9) As Variant
   For i = 1 To noreg
       mata(i, 1) = i
       mata(i, 2) = rmesa.Fields("FACTOR")
       mata(i, 3) = rmesa.Fields("CURVA")
       mata(i, 4) = rmesa.Fields("PLAZO")
       mata(i, 5) = rmesa.Fields("DESCRIPCION")
       mata(i, 6) = rmesa.Fields("DERIV")
       rmesa.MoveNext
   Next i
   rmesa.Close
   indice = BuscarValorArray(fecha2, MatFechasVaR, 1)
   fecha0 = MatFechasVaR(indice - noesc, 1)
   For i = 1 To noreg
       matf = Leer1FactorR(fecha0, fecha2, mata(i, 3), mata(i, 4))
       valmax = 0
       valmin = 0
       For j = 2 To UBound(matf, 1)
           valmax = Maximo(valmax, matf(j, 2) - matf(j - 1, 2))
           valmin = Minimo(valmin, matf(j, 2) - matf(j - 1, 2))
       Next j
       mata(i, 7) = valmin
       mata(i, 8) = valmax
   Next i
   For i = 1 To noreg
       txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       txtfiltro2 = "SELECT FACTOR,CURVA, PLAZO, DESCRIPCION, DERIVADA/100000000 AS DERIV FROM"
       txtfiltro2 = txtfiltro2 & " (SELECT * FROM " & TablaSensibPort & " WHERE FECHA= " & txtfecha1
       txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal' AND PORTAFOLIO='NEGOCIACION + INVERSION'"
       txtfiltro2 = txtfiltro2 & " AND SUBPORT='CONSOLIDADO'"
       txtfiltro2 = txtfiltro2 & " AND FACTOR = '" & mata(i, 2) & "')"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          mata(i, 9) = rmesa.Fields("DERIV")
          rmesa.Close
       End If
   Next i
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtSensibCAIR = mata
End Function

Function ObtVolatilCAIR(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT FACTOR, DESCRIPCION, ABS(VOLATIL)*VALOR*100 AS VOL FROM " & TablaSensibPort & " WHERE"
txtfiltro2 = txtfiltro2 & " FECHA = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal' AND PORTAFOLIO='NEGOCIACION + INVERSION' AND SUBPORT='CONSOLIDADO'"
txtfiltro2 = txtfiltro2 & " AND TVALOR='T CAMBIO' AND CURVA IN ('DOLAR PIP FIX','EURO PIP','YEN PIP') ORDER BY VOL DESC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = i
       mata(i, 2) = rmesa.Fields("FACTOR")
       mata(i, 3) = rmesa.Fields("DESCRIPCION")
       mata(i, 4) = rmesa.Fields("VOL")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT ABS(VOLATIL)*VALOR*100 AS VOL FROM " & TablaSensibPort & " WHERE"
       txtfiltro2 = txtfiltro2 & " FECHA = " & txtfecha1
       txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal'"
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO='NEGOCIACION + INVERSION' AND SUBPORT='CONSOLIDADO'"
       txtfiltro2 = txtfiltro2 & " AND TVALOR='T CAMBIO' AND FACTOR = '" & mata(i, 2) & "'"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          mata(i, 5) = rmesa.Fields("VOL")
          rmesa.Close
       End If
   Next i
   ObtVolatilCAIR = mata
End If
End Function

Function ObtVolatilCAIR2(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"

txtfiltro2 = "SELECT FACTOR, DESCRIPCION, ABS(VOLATIL)*VALOR*10000 AS VOL FROM " & TablaSensibPort & " WHERE"
txtfiltro2 = txtfiltro2 & " FECHA = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal' AND PORTAFOLIO='NEGOCIACION + INVERSION'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT='CONSOLIDADO' AND TVALOR<>'T CAMBIO' ORDER BY VOL DESC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = i
       mata(i, 2) = rmesa.Fields("FACTOR")
       mata(i, 3) = rmesa.Fields("DESCRIPCION")
       mata(i, 4) = rmesa.Fields("VOL")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT  ABS(VOLATIL)*VALOR*10000 AS VOL FROM " & TablaSensibPort & " WHERE"
       txtfiltro2 = txtfiltro2 & " FECHA = " & txtfecha1
       txtfiltro2 = txtfiltro2 & " AND PORT_FR='Normal'"
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO='NEGOCIACION + INVERSION' AND SUBPORT='CONSOLIDADO'"
       txtfiltro2 = txtfiltro2 & " AND FACTOR = '" & mata(i, 2) & "'"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          mata(i, 5) = rmesa.Fields("VOL")
          rmesa.Close
       End If
   Next i
   ObtVolatilCAIR2 = mata
End If
End Function


Private Sub mExppgcva1_Click()
Dim fecha As Date
Dim coperacion As Integer
fecha = #12/31/2018#

Call ExpPyGCVA2(fecha, coperacion)
End Sub

Private Sub mexppygmd_Click()
Dim dtfecha As Date
Screen.MousePointer = 11
dtfecha = #12/31/2018#
Call ExpPyGCVAMD(dtfecha)
Screen.MousePointer = 0
End Sub

Private Sub mExpResPI_Click()
Dim nomarch As String
Dim conadoex As New ADODB.Connection
Dim sihayarch As Boolean
Dim noesc As Integer
Dim htiempo As Integer
Dim txtport As String
Dim tfecha As String
Dim fecha As Date
Dim mata() As Variant
Dim matb() As Variant
Dim matv() As Variant
Dim strconexion As String

noesc = 500
htiempo = 1
tfecha = InputBox("Dame la fecha de calculo", , Date)
If IsDate(tfecha) Then
   Screen.MousePointer = 11
   fecha = CDate(tfecha)
   txtport = "PORTAFOLIO DE INVERSION"
   nomarch = DirReportes & "\Resultados Port Inv 2019.xlsx"
   sihayarch = VerifAccesoArch(nomarch)
   If sihayarch Then
        strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & nomarch
        strconexion = strconexion & ";Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
        conadoex.ConnectionString = strconexion
        conadoex.Open
        matb = LeerEscHistRE(fecha, txtportCalc1, "Normal", txtport, noesc, htiempo)
        If UBound(matb, 1) <> 0 Then
           matv = RutinaOrden(matb, 2, SRutOrden)
           Call GuardaEscHistExcel(fecha, matb(1, 1), matb(UBound(matb, 1), 1), matv, "ESCENARIOS$", 16, conadoex, RegExcel)
        End If
        matb = LeerResValVARPI(fecha)
        Call GuardaResPI(matb, "MTMCVAR$", conadoex, RegExcel)
        mata = RepPortPosEm(fecha, ClavePosPIDV)
        Call GuardaDetValPI(mata, "PIDV$", conadoex, RegExcel)
        mata = RepPortPosEm(fecha, ClavePosPICV)
        Call GuardaDetValPI(mata, "PICV$", conadoex, RegExcel)
        mata = LeerEscEstresPI(fecha)
        Call GuardaEscEstPI(mata, "ESTRES$", conadoex, RegExcel)
        mata = LeerSensibNuevo(fecha, txtportCalc2, "Normal", txtport)
        Call PegarSensResEx(fecha, mata, txtport, "SENSIB$", conadoex, RegExcel)
        mata = RepDerivPI(fecha)
        If UBound(mata, 1) <> 0 Then
           Call GuardaResDerivPI(mata, "DERIV$", conadoex, RegExcel)
        End If
        conadoex.Close
   End If
   MsgBox "Fin de proceso"
   Screen.MousePointer = 0
End If
End Sub

Function LeerEscEstresPI(ByVal fecha As Date)
Dim txtport As String
Dim txtsubport As String
txtport = "TOTAL"
txtsubport = "PORTAFOLIO DE INVERSION"
ReDim mata(1 To 1, 1 To 7) As Variant
mata(1, 1) = fecha
mata(1, 2) = LeerResEscEstres(fecha, txtport, txtsubport, "3 desv est")
mata(1, 3) = LeerResEscEstres(fecha, txtport, txtsubport, "Ad Hoc 1")
mata(1, 4) = LeerResEscEstres(fecha, txtport, txtsubport, "Crisis global 1")
mata(1, 5) = LeerResEscEstres(fecha, txtport, txtsubport, "Elecciones EU 1")
mata(1, 6) = LeerResEscEstres(fecha, txtport, txtsubport, "Elecciones EU 2")
mata(1, 7) = LeerResEscEstres(fecha, txtport, txtsubport, "FED disminuye estimulos de crecimiento")
LeerEscEstresPI = mata
End Function

Sub GuardaResDerivPI(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
noreg = UBound(mata, 1)
If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 2), 0) & ","
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 3), 0) & "',"
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 7), 0) & "',"
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 15), 0) & "',"
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 16), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 13), 0) & ","
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 12), " ") & "',"
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 8), 0) & "',"
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 9), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 10), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 11), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(CLng(mata(i, 14)), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 17), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 18), 0) & ","
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 19), 0) & "')"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub



Sub GuardaEscEstPI(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
noreg = UBound(mata, 1)
If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 2), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 3), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 4), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 5), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 6), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 7), 0) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub


Function LeerResValVARPI(ByVal fecha As Date)
Dim txtport As String
Dim txtsubport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim nconf As Double
Dim txttvar As String
Dim id_val As Integer
Dim matv() As Double
Dim exito As Boolean
noesc = 500
htiempo = 1
nconf = 0.03
txttvar = "CVARH"
txtport = "TOTAL"
id_val = 1
Dim mata(1 To 1, 1 To 12) As Variant
mata(1, 1) = fecha
txtsubport = "PI CONSERVADOS A VENCIMIENTO"
matv = LeerResValPort(fecha, txtport, txtsubport, id_val)
mata(1, 2) = matv(1)
txtsubport = "PI DISPONIBLES PARA LA VENTA"
matv = LeerResValPort(fecha, txtport, txtsubport, id_val)
mata(1, 3) = matv(1)
txtsubport = "DERIVADOS PI"
matv = LeerResValPort(fecha, txtport, txtsubport, id_val)
If UBound(matv, 1) <> 0 Then
mata(1, 4) = matv(1)
Else
mata(1, 4) = 0
End If

txtsubport = "PORTAFOLIO DE INVERSION"
mata(1, 5) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, nconf, 0, txttvar, exito)
txtsubport = "PI CONSERVADOS A VENCIMIENTO"
mata(1, 6) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, nconf, 0, txttvar, exito)
txtsubport = "PI DISPONIBLES PARA LA VENTA"
mata(1, 7) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, nconf, 0, txttvar, exito)
txtsubport = "DERIVADOS PI"
mata(1, 8) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, nconf, 0, txttvar, exito)
txtsubport = "PIDV+DERIVADOS"
mata(1, 9) = LeerResVaR(fecha, txtport, "Normal", txtsubport, noesc, htiempo, 0, nconf, 0, txttvar, exito)
LeerResValVARPI = mata
End Function

Sub GuardaDetValPI(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)

If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          txtcadena = txtcadena & i & ","
          txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & " " & i & "',"
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 2), 0) & "',"
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 3), 0) & "',"
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 4), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 5), 0) & ","
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 6), 0) & "',"
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 7), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 8), 0) & ","
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 9), 0) & "',"
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 10), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(CLng(mata(i, 11)), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 12), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 13), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 14), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 15), 0) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub


Sub GuardaResPI(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)

If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 2), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 3), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 4), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 5), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 6), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 7), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 8), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 9), 0) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub



Private Sub mGenComenRE_Click()
Dim txtfecha As String
Dim fecha As Date
Dim valposdiv As Double
Dim tCambio As Double
Dim mata() As Variant
Dim matb() As Variant
Dim txtmsg As String
Dim txtnomarch As String
txtfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(txtfecha) Then
   frmProgreso.Show
   fecha = CDate(txtfecha)
   mata = LeerParamRE(fecha)
   tCambio = mata(1, 3)
   matb = LeerDesgloseDiv(fecha, valposdiv)
   txtmsg = GenerarComentarioRE(fecha, valposdiv / tCambio)
   txtnomarch = "Comentario RE " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmReportes.CommonDialog1.FileName = txtnomarch
   frmReportes.CommonDialog1.ShowSave
   txtnomarch = frmReportes.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   Print #1, txtmsg
   Close #1
   Unload frmProgreso
End If
MsgBox "Fin de proceso"
frmReportes.SetFocus
End Sub

Private Sub mgenrepEPE_Click()
Dim i As Integer
Dim tfecha1 As String
Dim tfecha2 As String

Dim fecha1 As Date
Dim fecha2 As Date
Dim mata() As Variant
Dim valor As Integer
Dim porcen As String
Dim suma11 As Double, suma12 As Double, suma13 As Double, suma14 As Double, suma15 As Double
Dim suma21 As Double, suma22 As Double, suma23 As Double, suma24 As Double, suma25 As Double
Dim suma31 As Double, suma32 As Double, suma33 As Double, suma34 As Double, suma35 As Double
Screen.MousePointer = 11
suma11 = 0: suma12 = 0: suma13 = 0: suma14 = 0: suma15 = 0
suma21 = 0: suma22 = 0: suma23 = 0: suma24 = 0: suma25 = 0
suma31 = 0: suma32 = 0: suma33 = 0: suma34 = 0: suma35 = 0
tfecha1 = InputBox("Dame la primera fecha", , Date)
tfecha2 = InputBox("Dame la primera fecha", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
fecha1 = CDate(tfecha1)
fecha2 = CDate(tfecha2)
Open "d:\Reportes EPE " & Format(fecha2, "yyyy-mm-dd") & ".txt" For Output As #1
mata = GenRepEPE1(fecha1, fecha2, "F")
Print #1, "Contraparte" & Chr(9) & "Plazo menor a un año" & Chr(9) & "Plazo Total" & Chr(9) & "Función creciente" & Chr(9) & "Plazo total (t-1)" & Chr(9) & "Cambio"
If UBound(mata, 1) <> 0 Then
   For i = 1 To UBound(mata, 1)
       If mata(i, 5) > mata(i, 4) Then
          valor = 1
       Else
          valor = 0
       End If
       Print #1, mata(i, 1) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & valor & Chr(9) & mata(i, 8) & Chr(9) & mata(i, 5) - mata(i, 8)
       suma11 = suma11 + mata(i, 4)
       suma12 = suma12 + mata(i, 5)
       suma13 = suma13 + mata(i, 8)
       suma14 = suma14 + mata(i, 5) - mata(i, 8)
   Next i
   Print #1, "Total contrapartes financieras" & Chr(9) & suma11 & Chr(9) & suma12 & Chr(9) & Chr(9) & suma13 & Chr(9) & suma14
End If
mata = GenRepEPE1(fecha1, fecha2, "NF")
If UBound(mata, 1) <> 0 Then
For i = 1 To UBound(mata, 1)
    If mata(i, 5) > mata(i, 4) Then
       valor = 1
    Else
       valor = 0
    End If
    Print #1, mata(i, 1) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & valor & Chr(9) & mata(i, 8) & Chr(9) & mata(i, 5) - mata(i, 8)
    suma21 = suma21 + mata(i, 4)
    suma22 = suma22 + mata(i, 5)
    suma23 = suma23 + mata(i, 8)
    suma24 = suma24 + mata(i, 5) - mata(i, 8)

Next i
Print #1, "Total contrapartes no financieras" & Chr(9) & suma21 & Chr(9) & suma22 & Chr(9) & Chr(9) & suma23 & Chr(9) & suma24
End If
suma31 = suma11 + suma21
suma32 = suma12 + suma22
suma33 = suma13 + suma23
suma34 = suma14 + suma24
Print #1, "Total" & Chr(9) & suma31 & Chr(9) & suma32 & Chr(9) & Chr(9) & suma33 & Chr(9) & suma34
Print #1, ""
suma11 = 0: suma12 = 0: suma13 = 0: suma14 = 0: suma15 = 0
suma21 = 0: suma22 = 0: suma23 = 0: suma24 = 0: suma25 = 0
suma31 = 0: suma32 = 0: suma33 = 0: suma34 = 0: suma35 = 0
mata = GenRepEPE1(fecha1, fecha2, "F")
Print #1, "Contraparte" & Chr(9) & "EPE Plazo Total" & Chr(9) & "MTM" & Chr(9) & "% Exposición" & Chr(9) & "Esc de estrés 1" & Chr(9) & "Esc de estrés 2"
For i = 1 To UBound(mata, 1)
    If mata(i, 3) > 0 Then
       porcen = Format(mata(i, 5) / mata(i, 3) - 1, "##0.00 %")
    Else
       porcen = "NA"
    End If
    Print #1, mata(i, 1) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 3) & Chr(9) & porcen & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) - mata(i, 8)
    suma11 = suma11 + mata(i, 5)
    suma12 = suma12 + mata(i, 3)
    suma13 = suma13 + mata(i, 6)
    suma14 = suma14 + mata(i, 7) - mata(i, 8)
Next i
Print #1, "Total contrapartes financieras" & Chr(9) & suma11 & Chr(9) & suma12 & Chr(9) & Chr(9) & suma13 & Chr(9) & suma14
mata = GenRepEPE1(fecha1, fecha2, "NF")
For i = 1 To UBound(mata, 1)
    If mata(i, 3) > 0 Then
       porcen = Format(mata(i, 5) / mata(i, 3) - 1, "##0.00 %")
    Else
       porcen = "NA"
    End If
    Print #1, mata(i, 1) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 3) & Chr(9) & porcen & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) - mata(i, 8)
    suma21 = suma21 + mata(i, 5)
    suma22 = suma22 + mata(i, 3)
    suma23 = suma23 + mata(i, 6)
    suma24 = suma24 + mata(i, 7) - mata(i, 8)
Next i
Print #1, "Total contrapartes no financieras" & Chr(9) & suma21 & Chr(9) & suma22 & Chr(9) & Chr(9) & suma23 & Chr(9) & suma24
Close #1
Else

End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mHistoria_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmHistPrecios.Show
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mLimCont_Click()
Dim fecha As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim j As Integer
Dim txtnomarch As String
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
txtfecha = InputBox("Dame la fecha a obtener", , Date)
frmProgreso.Show
If IsDate(txtfecha) Then
   fecha = CDate(txtfecha)
   SiActTProc = True
   txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtfiltro2 = "SELECT * FROM " & TablaResLimContrap & " WHERE FECHA = " & txtfecha
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      txtnomarch = DirResVaR & "\Res Lim Contrap NF " & Format(fecha, "yyyy-mm-dd") & ".txt"
      frmCalVar.CommonDialog1.FileName = txtnomarch
      frmCalVar.CommonDialog1.ShowSave
      txtnomarch = frmCalVar.CommonDialog1.FileName
      Open txtnomarch For Output As #1
      rmesa.Open txtfiltro2, ConAdo
      Print #1, "Fecha" & Chr(9) & "Clave de operacion" & Chr(9) & "Valuacion maxima" & Chr(9) & "Fecha de curva" & Chr(9) & "Fecha val max" & Chr(9) & "Escenario estres maximo" & Chr(9) & "Fecha esc max" & Chr(9) & "Fecha val estres max"
   ReDim mata(1 To noreg, 1 To 8) As Variant
      For i = 1 To noreg
          mata(i, 1) = rmesa.Fields("FECHA")
          mata(i, 2) = rmesa.Fields("COPERACION")
          mata(i, 3) = rmesa.Fields("VALMAX1")
          mata(i, 4) = rmesa.Fields("FESCMAX1")
          mata(i, 5) = rmesa.Fields("FVALMAX1")
          mata(i, 6) = rmesa.Fields("VALMAX2")
          mata(i, 7) = rmesa.Fields("FESCMAX2")
          mata(i, 8) = rmesa.Fields("FVALMAX2")
          rmesa.MoveNext
          txtcadena = ""
          For j = 1 To 8
              txtcadena = txtcadena & mata(i, j) & Chr(9)
          Next j
          Print #1, txtcadena
      Next i
      rmesa.Close
      Close #1
      Call ActUHoraUsuario
      SiActTProc = False

   Else
      MsgBox "No hay registros para esta fecha"
   End If
End If
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mLimContFwd_Click()
Dim fecha As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtnomarch As String
Dim txtfecha As String
Dim tfecha As Date
Dim exitoarch As Boolean
Dim txtcadena As String
Dim i As Integer, noreg As Integer
Dim rmesa As New ADODB.recordset

tfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(tfecha) Then
   Screen.MousePointer = 11
   fecha = CDate(tfecha)
   txtnomarch = "Res calc lim cont fwds " & Format(fecha, "yyyy-mm-dd")
   frmReportes.CommonDialog1.FileName = txtnomarch
   frmReportes.CommonDialog1.ShowSave
   txtnomarch = frmReportes.CommonDialog1.FileName
   Call VerificarSalidaArchivo(txtnomarch, 1, exitoarch)
   If exitoarch Then
      txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtfiltro2 = "SELECT * FROM " & TablaExpFwds & " WHERE FECHA = " & txtfecha
      txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg = rmesa.Fields(0)
      rmesa.Close
      If noreg <> 0 Then
         rmesa.Open txtfiltro2, ConAdo
         For i = 1 To noreg
             txtcadena = rmesa.Fields("FECHA") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("FECHAREG") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("CPOSICION") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("COPERACION") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("FVALUACION") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("VOLATILIDAD") & Chr(9)
             txtcadena = txtcadena & rmesa.Fields("VALOR")
             rmesa.MoveNext
             Print #1, txtcadena
         Next i
         rmesa.Close
      End If
      Close #1
      Screen.MousePointer = 0
      MsgBox "Proceso finalizado"
   End If
   
End If
End Sub

Private Sub mMatTran_Click()
Dim tfecha As String
Dim fecha As Date
Dim i As Integer
Dim j As Integer
Dim mtrans() As Double
Dim txtcadena As String

tfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Open "d:\matrices.txt" For Output As #1
   mtrans = CargarMatTrans(fecha, "I")
   For i = 1 To UBound(mtrans, 1)
       txtcadena = ""
       For j = 1 To UBound(mtrans, 2)
           txtcadena = txtcadena & mtrans(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Print #1, ""
   mtrans = CargarMatTrans(fecha, "N")
   For i = 1 To UBound(mtrans, 1)
       txtcadena = ""
       For j = 1 To UBound(mtrans, 1)
           txtcadena = txtcadena & mtrans(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Print #1, ""
   
   Close #1
   
End If
MsgBox "Fin de proceso"
End Sub

Private Sub mObCurvasRE_Click()
Dim txtfecha As String
Dim strconexion As String
Dim fecha As Date
Dim mcurvas() As Variant
Dim txtnomarch As String
Dim i As Long
Dim j As Long
Dim txtcadena As String
SiActTProc = True
Screen.MousePointer = 11
txtfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(txtfecha) Then
   frmProgreso.Show
   fecha = CDate(txtfecha)
   mcurvas = LeerCurvasRE(fecha)
   If UBound(mcurvas, 1) <> 0 Then
      txtnomarch = DirResVaR & "\Curvas RE " & Format(fecha, "yyyy-mm-dd") & ".txt"
      frmReportes.CommonDialog1.FileName = txtnomarch
      frmReportes.CommonDialog1.ShowSave
      txtnomarch = frmReportes.CommonDialog1.FileName
      Open txtnomarch For Output As #1
      For i = 1 To UBound(mcurvas, 1)
          txtcadena = ""
          For j = 1 To UBound(mcurvas, 2)
              txtcadena = txtcadena & mcurvas(i, j) & Chr(9)
          Next j
          Print #1, txtcadena
      Next i
      Close #1
   End If
   Unload frmProgreso
End If
SiActTProc = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub mRepEfPros_Click()
frmRepEfPros.Show
End Sub

Private Sub mReporte003_Click()
Dim fecha As Date
Dim txtfecha As String
Screen.MousePointer = 11
txtfecha = InputBox("dame la fecha a cargar", , Date)
If IsDate(txtfecha) Then
   fecha = CDate(txtfecha)
   frmProgreso.Show
   SiActTProc = True
   Call ExpCaracFlujosDeuda(fecha)
   Call ActUHoraUsuario
   SiActTProc = False
   Unload frmProgreso
Else
   MsgBox "No es una fecha valida"
End If
Screen.MousePointer = 0
End Sub

Private Sub mReporte02_Click()
Dim txtfecha As String
Dim fecha As Date
Dim dfecha As String
Dim txtfiltro As String
Dim txtcadena As String
Dim mata() As New propPosFwd
Dim i As Integer
Dim j As Integer
Dim txtnomarch As String
dfecha = InputBox("Dame la fecha", , Date)
If IsDate(dfecha) Then
   fecha = CDate(dfecha)
   Screen.MousePointer = 11
   txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha
   mata = LeerTablaPosFwd(txtfiltro)
   txtnomarch = "d:\pos fwds " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmReportes.CommonDialog1.FileName = txtnomarch
   frmReportes.CommonDialog1.ShowSave
   txtnomarch = frmReportes.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   txtcadena = "Clave de operacion" & Chr(9)
   txtcadena = txtcadena & "Intencion" & Chr(9)
   txtcadena = txtcadena & "Estructural" & Chr(9)
   txtcadena = txtcadena & "Reclasificacion" & Chr(9)
    txtcadena = txtcadena & "Fecha de inicio" & Chr(9)
   txtcadena = txtcadena & "Fecha de vencimiento" & Chr(9)
   txtcadena = txtcadena & "Monto nocional" & Chr(9)
   txtcadena = txtcadena & "Activa/Pasiva" & Chr(9)
   txtcadena = txtcadena & "Tipo forward" & Chr(9)
   txtcadena = txtcadena & "Strike" & Chr(9)
   txtcadena = txtcadena & "Contraparte"
   Print #1, txtcadena
   For i = 1 To UBound(mata, 1)
      txtcadena = ""
      
      txtcadena = txtcadena & mata(i).c_operacion & Chr(9)
      txtcadena = txtcadena & mata(i).intencion & Chr(9)
      txtcadena = txtcadena & mata(i).EstructuralFwd & Chr(9)
      txtcadena = txtcadena & mata(i).ReclasificaFwd & Chr(9)
      txtcadena = txtcadena & mata(i).FCompraFwd & Chr(9)
      txtcadena = txtcadena & mata(i).FVencFwd & Chr(9)
      txtcadena = txtcadena & mata(i).MontoNocFwd & Chr(9)
      txtcadena = txtcadena & mata(i).Tipo_Mov & Chr(9)
      txtcadena = txtcadena & mata(i).ClaveProdFwd & Chr(9)
      txtcadena = txtcadena & mata(i).PAsignadoFwd & Chr(9)
      txtcadena = txtcadena & mata(i).ID_ContrapFwd & Chr(9)
     Print #1, txtcadena
   Next i
Close #1
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mReporte100_Click()
frmPyGSubport.Show 1
End Sub

Private Sub mReporte101_Click()
Dim txtfecha As String
Dim fecha As Date
Screen.MousePointer = 11
txtfecha = InputBox("dame la fecha a cargar", , Date)
If IsDate(txtfecha) Then
   fecha = CDate(txtfecha)
   frmProgreso.Show
   SiActTProc = True
   Call ExpCaracFlujosSwaps(fecha)
   Call ActUHoraUsuario
   SiActTProc = False
   Unload frmProgreso
Else
   MsgBox "No es una fecha valida"
End If
 MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mReporte103_Click()
frmValContrap.Show 1
End Sub

Private Sub mreporte11_Click()
Dim tfecha As String
Dim fecha As Date
Screen.MousePointer = 11
'primero se obtiene la fecha del reporte
tfecha = InputBox("Dame la fecha del reporte", , Date)
If IsDate(tfecha) Then
   Call ImpRepVaRMDT(CDate(tfecha))
End If
MsgBox "Fin de proceso"
frmReportes.SetFocus
Screen.MousePointer = 0
End Sub



Private Sub mReporte123_Click()

   Dim dt_fecha1 As Date
   Dim dt_fecha2 As Date
   Dim txtfecha1 As String
   Dim txtfecha2 As String
   Dim txtnomarch As String
   Dim txtcadena As String
   Dim mata() As New resCVA
   Dim matb() As Variant
   Dim i As Integer
   Dim sumacf As New resCVA
   Dim sumacnf As New resCVA
   Dim sumat As New resCVA
   Dim exitoarch As Boolean
   Screen.MousePointer = 11
 
    'las recuperaciones corresponden a las columnas 4 y 6
    txtfecha1 = InputBox("Dame la primera fecha ", , Date)
    txtfecha2 = InputBox("Dame la segunda fecha ", , Date)
    If IsDate(txtfecha1) And IsDate(txtfecha2) Then
       dt_fecha1 = CDate(txtfecha1)
       dt_fecha2 = CDate(txtfecha2)
       SiActTProc = True
       frmProgreso.Show
       txtnomarch = DirResVaR & "\Reporte CVA " & Format$(dt_fecha2, "yyyy-mm-dd") & ".txt"
       frmCalVar.CommonDialog1.FileName = txtnomarch
       frmCalVar.CommonDialog1.ShowSave
       txtnomarch = frmCalVar.CommonDialog1.FileName
       Call VerificarSalidaArchivo(txtnomarch, 1, exitoarch)
       If exitoarch Then
       txtcadena = "Contraparte" & Chr$(9) & "Escala" & Chr$(9) & "Calificación" & Chr$(9) & "MtM (mdp)" & Chr$(9) & "Derivados" & Chr$(9) & "MD" & Chr$(9) & "PIDV" & Chr$(9) & "PICV" & Chr$(9) & "Total" & Chr$(9) & "Cambio"
       Print #1, txtcadena
       mata = RepCVA(dt_fecha1, dt_fecha2, "F")
       For i = 1 To UBound(mata, 1)
          Print #1, mata(i).descrip & Chr(9) & mata(i).escala & Chr(9) & mata(i).calif & Chr(9) & mata(i).mtm & Chr(9) & mata(i).cvaderiv & Chr(9) & mata(i).cvamd & Chr(9) & mata(i).cvapidv & Chr(9) & mata(i).cvapicv & Chr(9) & mata(i).cvatotal & Chr(9) & mata(i).cvatotal - mata(i).cvat_1
          sumacf.mtm = sumacf.mtm + mata(i).mtm
          sumacf.cvaderiv = sumacf.cvaderiv + mata(i).cvaderiv
          sumacf.cvamd = sumacf.cvamd + mata(i).cvamd
          sumacf.cvapidv = sumacf.cvapidv + mata(i).cvapidv
          sumacf.cvapicv = sumacf.cvapicv + mata(i).cvapicv
          sumacf.cvatotal = sumacf.cvatotal + mata(i).cvatotal
          sumacf.cvat_1 = sumacf.cvat_1 + mata(i).cvat_1
          sumacf.cva_dif = sumacf.cva_dif + mata(i).cvatotal - mata(i).cvat_1
       Next i
       Print #1, "Total de Contrapartes Financieras" & Chr(9) & Chr(9) & Chr(9) & sumacf.mtm & Chr(9) & sumacf.cvaderiv & Chr(9) & sumacf.cvamd & Chr(9) & sumacf.cvapidv & Chr(9) & sumacf.cvapicv & Chr(9) & sumacf.cvatotal & Chr(9) & sumacf.cva_dif
       mata = RepCVA(dt_fecha1, dt_fecha2, "NF")
       For i = 1 To UBound(mata, 1)
          Print #1, mata(i).descrip & Chr(9) & mata(i).escala & Chr(9) & mata(i).calif & Chr(9) & mata(i).mtm & Chr(9) & mata(i).cvaderiv & Chr(9) & mata(i).cvamd & Chr(9) & mata(i).cvapidv & Chr(9) & mata(i).cvapicv & Chr(9) & mata(i).cvatotal & Chr(9) & mata(i).cvatotal - mata(i).cvat_1
          sumacnf.mtm = sumacnf.mtm + mata(i).mtm
          sumacnf.cvaderiv = sumacnf.cvaderiv + mata(i).cvaderiv
          sumacnf.cvamd = sumacnf.cvamd + mata(i).cvamd
          sumacnf.cvapidv = sumacnf.cvapidv + mata(i).cvapidv
          sumacnf.cvapicv = sumacnf.cvapicv + mata(i).cvapicv
          sumacnf.cvatotal = sumacnf.cvatotal + mata(i).cvatotal
          sumacnf.cvat_1 = sumacnf.cvat_1 + mata(i).cvat_1
          sumacnf.cva_dif = sumacnf.cva_dif + mata(i).cvatotal - mata(i).cvat_1
       Next i
       Print #1, "Total de Contrapartes no Financieras" & Chr(9) & Chr(9) & Chr(9) & sumacnf.mtm & Chr(9) & sumacnf.cvaderiv & Chr(9) & sumacnf.cvamd & Chr(9) & sumacnf.cvapidv & Chr(9) & sumacnf.cvapicv & Chr(9) & sumacnf.cvatotal & Chr(9) & sumacnf.cva_dif
          sumat.mtm = sumacf.mtm + sumacnf.mtm
          sumat.cvaderiv = sumacf.cvaderiv + sumacnf.cvaderiv
          sumat.cvamd = sumacf.cvamd + sumacnf.cvamd
          sumat.cvapidv = sumacf.cvapidv + sumacnf.cvapidv
          sumat.cvapicv = sumacf.cvapicv + sumacnf.cvapicv
          sumat.cvatotal = sumacf.cvatotal + sumacnf.cvatotal
          sumat.cvat_1 = sumacf.cvat_1 + sumacnf.cvat_1
          sumat.cva_dif = sumacf.cva_dif + sumacnf.cva_dif
       Print #1, "Total Global" & Chr(9) & Chr(9) & Chr(9) & sumat.mtm & Chr(9) & sumat.cvaderiv & Chr(9) & sumat.cvamd & Chr(9) & sumat.cvapidv & Chr(9) & sumat.cvapicv & Chr(9) & sumat.cvatotal & Chr(9) & sumat.cva_dif
       Print #1, ""
       matb = RepCVA2(dt_fecha2)
       Print #1, "CVA" & Chr(9) & Format(dt_fecha2) & Chr(9) & "Esc. Estrés 1" & Chr(9) & "Esc. Estrés 2"
       For i = 1 To UBound(matb, 1)
           Print #1, matb(i, 1) & Chr(9) & matb(i, 2) & Chr(9) & matb(i, 3) & Chr(9) & matb(i, 4)
       Next i
       Close #1
       End If
       Unload frmProgreso
       MsgBox "Fin de proceso"
       Call ActUHoraUsuario
       SiActTProc = False
    End If
    Screen.MousePointer = 0
End Sub

Private Sub mReporte200_Click()
frmPyGMontOper.Show
End Sub


Sub ExpPyGCVA(ByVal dtfecha As Date)
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noreg As Long
    Dim dxv        As Integer
    Dim indice     As Integer
    Dim txtcontrap As String
    Dim matc()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim id_contrap As Integer
    Dim rmesa As New ADODB.recordset
    
    SiActTProc = True
    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT * from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION, FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        txtnomarch = DirResVaR & "\Reporte de simulaciones CVA " & Format$(dtfecha, "yyyy-mm-dd") & ".TXT"
        frmReportes.CommonDialog1.FileName = txtnomarch
        frmReportes.CommonDialog1.ShowSave
        txtnomarch = frmReportes.CommonDialog1.FileName
        Open txtnomarch For Output As #1
        ReDim matb(1 To 3) As Variant
        rmesa.Open txtfiltro1, ConAdo

        For i = 1 To noreg
            matb(1) = rmesa.Fields("COPERACION")    'CLAVE operacion
            matb(2) = rmesa.Fields("FECHA_F")    'fecha forward
            dxv = matb(2) - dtfecha
            matb(3) = rmesa.Fields("VECTOR_PYG").GetChunk(rmesa.Fields("VECTOR_PYG").ActualSize)
            matc = EncontrarSubCadenas(matb(3), ",")
            txtcadena = matb(1) & Chr$(9) & dxv & Chr$(9)
            For j = 1 To UBound(matc, 1)
                txtcadena = txtcadena & matc(j) & Chr$(9)
            Next j
            Print #1, txtcadena
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo las p&l del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
        Next i
        Close #1
        rmesa.Close
    End If

End Sub


Sub ExpPyGCVA2(ByVal dtfecha As Date, ByVal coperacion As String)
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noreg As Long
    Dim dxv        As Integer
    Dim indice     As Integer
    Dim txtcontrap As String
    Dim matc()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim id_contrap As Integer
    Dim rmesa As New ADODB.recordset
    
    If IsDate(txtfecha) Then
    SiActTProc = True

    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT * from " & TablaPLEscCVA & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND COPERACION = '" & coperacion & "'  ORDER BY FECHA_F"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        txtnomarch = DirResVaR & "\Reporte de simulaciones CVA " & coperacion & "  " & Format$(dtfecha, "yyyy-mm-dd") & ".TXT"
        frmReportes.CommonDialog1.FileName = txtnomarch
        frmReportes.CommonDialog1.ShowSave
        txtnomarch = frmReportes.CommonDialog1.FileName
        Open txtnomarch For Output As #1
        ReDim matb(1 To 3) As Variant
        rmesa.Open txtfiltro1, ConAdo

        For i = 1 To noreg
            matb(1) = rmesa.Fields("COPERACION")    'CLAVE operacion
            matb(2) = rmesa.Fields("FECHA_F")    'fecha forward
            dxv = matb(2) - dtfecha
            matb(3) = rmesa.Fields("VECTOR_PYG").GetChunk(rmesa.Fields("VECTOR_PYG").ActualSize)
            matc = EncontrarSubCadenas(matb(3), ",")
            txtcadena = matb(1) & Chr$(9) & dxv & Chr$(9)
            For j = 1 To UBound(matc, 1)
                txtcadena = txtcadena & matc(j) & Chr$(9)
            Next j
            Print #1, txtcadena
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo las p&l del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
        Next i
        Close #1
        rmesa.Close
    End If
    Unload frmProgreso
    Call ActUHoraUsuario
    SiActTProc = False
    End If
    MsgBox "Fin de proceso"
    Screen.MousePointer = 0

End Sub

Sub ExpPyGCVAMD(ByVal dtfecha As Date)
    Dim txtfecha   As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim i       As Long
    Dim j As Long
    Dim noreg As Long
    Dim dxv        As Integer
    Dim indice     As Integer
    Dim txtcontrap As String
    Dim matc()     As String
    Dim txtcadena  As String
    Dim txtnomarch As String
    Dim id_contrap As Integer
    Dim rmesa As New ADODB.recordset
    
    SiActTProc = True

    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT * from " & TablaPYGCVAMD & " WHERE FECHA = " & txtfecha
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION"
    txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
    rmesa.Open txtfiltro2, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        txtnomarch = DirResVaR & "\Reporte de simulaciones CVA MD  " & Format$(dtfecha, "yyyy-mm-dd") & ".TXT"
        frmReportes.CommonDialog1.FileName = txtnomarch
        frmReportes.CommonDialog1.ShowSave
        txtnomarch = frmReportes.CommonDialog1.FileName
        Open txtnomarch For Output As #1
        ReDim matb(1 To 3) As Variant
        rmesa.Open txtfiltro1, ConAdo

        For i = 1 To noreg
            matb(1) = rmesa.Fields("CPOSICION")    'CLAVE operacion
            matb(2) = rmesa.Fields("COPERACION")    'CLAVE operacion
            matb(3) = rmesa.Fields("VECTORPYG").GetChunk(rmesa.Fields("VECTORPYG").ActualSize)
            matc = EncontrarSubCadenas(matb(3), ",")
            txtcadena = matb(1) & Chr$(9) & matb(2) & Chr$(9)
            For j = 1 To UBound(matc, 1)
                txtcadena = txtcadena & matc(j) & Chr$(9)
            Next j
            Print #1, txtcadena
            rmesa.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Leyendo las p&l del " & dtfecha & " " & Format$(AvanceProc, "##0.00 %")
        Next i
        Close #1
        rmesa.Close
End If
End Sub



Private Sub mReporte201_Click()
  Dim dtfecha As Date
  Dim tfecha As String
  tfecha = InputBox("Dame la fecha de calculo ", , Date)
  If IsDate(tfecha) Then
     dtfecha = CDate(tfecha)
     Screen.MousePointer = 11
     frmProgreso.Show
     Call ExpPyGCVA(dtfecha)
     Unload frmProgreso
     Screen.MousePointer = 0
  End If
End Sub

Private Sub mReporte22_Click()
Dim fecha As Date
Dim txtfecha As String
Screen.MousePointer = 11
 txtfecha = InputBox("Dame la fecha de la posicion ", , Date)
 If IsDate(txtfecha) Then
  fecha = CDate(txtfecha)
  Call GenRepDurPos(fecha)
 End If
Screen.MousePointer = 0
End Sub

Private Sub mReporte300_Click()
frmPyGMontSubport.Show
End Sub

Private Sub mReporte301_Click()
Dim mata() As Variant
Dim txtport As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim fecha As Date
Dim noreg As Integer
Dim nocampos As Integer
Dim txtsubport As String
Dim i As Integer
Dim j As Integer
Dim nomarch As String
Dim txtsalida As String
Dim exitoarch As Boolean

txtport = "CONSOLIDADO"
txtfecha = InputBox("Indica la fecha de los datos a obtener", , Date)
txtport = InputBox("Indicame la posicion", , txtport)
If IsDate(txtfecha) Then
frmProgreso.Show
fecha = CDate(txtfecha)
mata = LeerSensibNuevo(fecha, txtportCalc2, "Normal", txtport)
'mata = LeerSensibNuevo(fecha, "FID 2065", "Normal", "FID 2065")
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg <> 0 Then
 frmSensibilidades.Show
 frmSensibilidades.MSFlexGrid1.Cols = 12
 frmSensibilidades.MSFlexGrid1.Rows = noreg + 1
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 1) = "Descripcion"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 2) = "Factor"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 3) = "Curva"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 4) = "Plazo"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 5) = "Tipo factor"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 6) = "Valor"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 7) = "Derivada"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 8) = "Vol %"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 9) = "VaR"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 10) = "Sensibilidad"
 frmSensibilidades.MSFlexGrid1.TextMatrix(0, 11) = "Volatilidad"
 For i = 1 To noreg
  For j = 1 To 11
  If Not EsVariableVacia(mata(i, j)) Then
   frmSensibilidades.MSFlexGrid1.TextMatrix(i, j) = mata(i, j)
  Else
   frmSensibilidades.MSFlexGrid1.TextMatrix(i, j) = ""
  End If
  Next j
  
 Next i
  txtfecha1 = Format(fecha, "yyyymmdd")
  nomarch = DirResVaR & "\Sensibilidades " & txtport & " " & txtfecha1 & ".txt"
  frmCalVar.CommonDialog1.FileName = nomarch
  frmCalVar.CommonDialog1.ShowSave
  nomarch = frmCalVar.CommonDialog1.FileName
  Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
  If exitoarch Then
  txtsalida = "Descripcion" & Chr(9) & "Factor" & Chr(9) & "Curva" & Chr(9) & "Plazo" & Chr(9) & "Tipo valor " & Chr(9) & "Valor" & Chr(9) & "Derivada" & Chr(9) & "Vol %" & Chr(9) & "VaR" & Chr(9) & "Sensbilidad" & Chr(9) & "Volatilidad"
  Print #1, txtsalida
  For i = 1 To UBound(mata, 1)
  txtsalida = ""
   For j = 1 To 11
    txtsalida = txtsalida & mata(i, j) & Chr(9)
   Next j
   Print #1, txtsalida
  Next i
  Close #1
  MsgBox "Se generaron las sensibilidades en el archivo " & nomarch
  End If
 Else
    MsgBox "No hay datos de sensibilidades para esta fecha " & fecha
 End If
End If
Unload frmProgreso
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub

Private Sub mReporte33_Click()
Dim fecha As Date
Dim tfecha As String
Screen.MousePointer = 11
'primero se obtiene la fecha del reporte
tfecha = InputBox("Dame la fecha del reporte", , Date)
If IsDate(tfecha) Then
 fecha = CDate(tfecha)
 Call ImpRepDeriv(fecha)
End If
Screen.MousePointer = 0
End Sub

Private Sub mReporte51_Click()
Dim nesc As Integer
Dim tfecha As String
Dim nomarch As String
Dim fecha As Date
Dim mata() As Variant
Dim i As Long
Dim valfa As Double
Dim vobservada As Double
Dim exitoarch As Boolean

Screen.MousePointer = 11
nesc = 500
valfa = NormalInv(0.99)
tfecha = InputBox("Dame la fecha a recuperar", , Date)
If IsDate(tfecha) Then
fecha = CDate(tfecha)
mata = LeerAnalisisFR(fecha, nesc)
If UBound(mata, 1) > 0 Then
   nomarch = DirResVaR & "\Analisis factores " & Format(fecha, "yyyymmdd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
   If exitoarch Then
   Print #1, "Factor" & Chr(9) & "Plazo" & Chr(9) & "Valor t-1" & Chr(9) & "Valor t" & Chr(9) & "Variación observada" & Chr(9) & "Variación esperada" & Chr(9) & "Nivel de variación"
   For i = 1 To UBound(mata, 1)
       If mata(i, 7) <> 0 Then
          vobservada = Abs(mata(i, 6) / mata(i, 7))
          If vobservada >= valfa And vobservada < 3 Then
             Print #1, mata(i, 2) & Chr(9) & mata(i, 3) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) & Chr(9) & "Mayor a 2.33 desv est"
          ElseIf vobservada >= 3 And vobservada < 6 Then
             Print #1, mata(i, 2) & Chr(9) & mata(i, 3) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) & Chr(9) & "Mayor a 3 desv est"
          ElseIf vobservada >= 6 Then
             Print #1, mata(i, 2) & Chr(9) & mata(i, 3) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) & Chr(9) & "Mayor a 6 desv est"
          End If
       Else
         'MsgBox mata(i, 2) & " " & mata(i, 3)
       End If
   Next i
   Close #1
   MsgBox "Se creo el archivo " & nomarch
   End If
Else
   MsgBox "No hay datos de analisis de factores para esta fecha"
End If
End If
Screen.MousePointer = 0
End Sub

Private Sub mReporte52_Click()
frmCurvas.Show
End Sub

Private Sub mReporte53_Click()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim exito As Boolean
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

   SiActTProc = True
   txtfecha1 = InputBox("dame la fecha inicial ", , Date)
   txtfecha2 = InputBox("dame la fecha final ", , Date)
   If IsDate(txtfecha1) And IsDate(txtfecha2) Then
      fecha1 = CDate(txtfecha1)
      fecha2 = CDate(txtfecha2)
      Screen.MousePointer = 11
          frmProgreso.Show
          Call CrearMatFRiesgo3(fecha1, fecha2, MatFactRiesgo, "", exito)
          Unload frmProgreso
          frmHistFRiesgo.Show
      Screen.MousePointer = 0
   End If
   Call ActUHoraUsuario
   SiActTProc = False
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mReporte55_Click()
frmPyGOper.Show 1
End Sub

Private Sub mReporteBt_Click()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim nomarch As String
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim siesfv As Boolean
Dim txtcadena As String
Dim mata() As Double
Dim exitoarch As Boolean

Screen.MousePointer = 11

txtfecha1 = InputBox("dame la fecha inicial", , Date)
txtfecha2 = InputBox("dame la fecha final", , Date)
If IsDate(txtfecha1) And IsDate(txtfecha2) Then
   fecha1 = CDate(txtfecha1)
   fecha2 = CDate(txtfecha2)
  MatGruposPortPos = CargaGruposPortPos("REPORTE PRINCIPAL")
  nomarch = DirResVaR & "\Backtesting " & Format(fecha1, "YYYY-MM-DD") & " - " & Format(fecha2, "YYYY-MM-DD") & ".txt"
  frmCalVar.CommonDialog1.FileName = nomarch
  frmCalVar.CommonDialog1.ShowSave
  nomarch = frmCalVar.CommonDialog1.FileName
  Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
  If exitoarch Then
  txtcadena = "fecha" & Chr(9)
  For i = 1 To UBound(MatGruposPortPos, 1)
      txtcadena = txtcadena & MatGruposPortPos(i, 3) & Chr(9)
  Next i
  Print #1, txtcadena
  fecha = fecha1
  Do While fecha <= fecha2
     siesfv = EsFechaVaR(fecha)
     If siesfv Then
     txtcadena = fecha & Chr(9)
     For i = 1 To UBound(MatGruposPortPos, 1)
         mata = LeerBackPort(fecha, txtportCalc2, MatGruposPortPos(i, 3))
         txtcadena = txtcadena & mata(3) & Chr(9)
     Next i
     Print #1, txtcadena
     End If
     fecha = fecha + 1
  Loop
  Close #1
  End If
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mreportecsp_Click()
Dim fecha As Date
Dim tfecha As String
Dim txtfecha As String
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim nomarch As String
Dim txtcadena As String
Dim nodif As Long
Dim noreg As Long
Dim alerta As String
Dim exitoarch As Boolean

Screen.MousePointer = 11
tfecha = InputBox("Indica la fecha a obtener", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Call LeerValPosMD(fecha, mata, nodif, alerta)
   If Len(alerta) <> 0 Then MsgBox alerta
   MsgBox "Número de diferencias: " & nodif
   noreg = UBound(mata, 1)
   nomarch = DirResVaR & "\Comparacion Val MD SIVARMER vs. PIP " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
   If exitoarch Then
   txtcadena = "Clave de posicion" & Chr(9) & "Clave de operación" & Chr(9) & "TV" & Chr(9) & "Emision" & Chr(9) & "Serie" & Chr(9) & "No. de títulos" & Chr(9) & "Val. SIVARMER" & Chr(9) & "Val. PIP" & Chr(9) & "Diferencia"
   Print #1, txtcadena
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To 9
   txtcadena = txtcadena & mata(i, j) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   
   MsgBox "Se creo el archivo " & nomarch
   MsgBox "Fin de proceso"
   Else
     MsgBox "No se pudo crear el archivo " & nomarch
   End If
End If
Screen.MousePointer = 0
End Sub

Private Sub mReporteCVSP2065_Click()
Dim fecha As Date
Dim tfecha As String
Dim txtfecha As String
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim nomarch As String
Dim txtcadena As String
Dim nodif As Long
Dim noreg As Long

Screen.MousePointer = 11
tfecha = InputBox("Indica la fecha a obtener", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Call LeerValPosPension(fecha, ClavePosPension1, mata, nodif)
   MsgBox "Número de diferencias: " & nodif
   noreg = UBound(mata, 1)
   nomarch = DirResVaR & "\Comparacion Val fid. 2065 SIVARMER vs. PIP " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #1
   txtcadena = "Clave de posicion" & Chr(9) & "Clave de operación" & Chr(9) & "TV" & Chr(9) & "Emision" & Chr(9) & "Serie" & Chr(9) & "No. de títulos" & Chr(9) & "Val. SIVARMER" & Chr(9) & "Val. PIP" & Chr(9) & "Diferencia"
   Print #1, txtcadena
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To 9
   txtcadena = txtcadena & mata(i, j) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se creo el archivo " & nomarch
   Call LeerValPosPension(fecha, ClavePosPension2, mata, nodif)
   MsgBox "Número de diferencias: " & nodif
   noreg = UBound(mata, 1)
   nomarch = DirResVaR & "\Comparacion Val fid. 2160 SIVARMER vs. PIP " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #1
   txtcadena = "Clave de posicion" & Chr(9) & "Clave de operación" & Chr(9) & "TV" & Chr(9) & "Emision" & Chr(9) & "Serie" & Chr(9) & "No. de títulos" & Chr(9) & "Val. SIVARMER" & Chr(9) & "Val. PIP" & Chr(9) & "Diferencia"
   Print #1, txtcadena
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To 9
   txtcadena = txtcadena & mata(i, j) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se creo el archivo " & nomarch

   MsgBox "Fin de proceso"
End If
Screen.MousePointer = 0
End Sub

Private Sub mReporteEET_Click()
Dim txtfecha As String
Dim fecha As Date
    txtfecha = InputBox("Dame la fecha de calculo ", , Date)
    If IsDate(txtfecha) Then
    fecha = CDate(txtfecha)
    frmProgreso.Show
    SiActTProc = True
    Call LeerEscEstresTaylor(fecha, txtportCalc2)
    Call ActUHoraUsuario
    SiActTProc = False
    Unload frmProgreso
    MsgBox "Fin de proceso"
    End If

End Sub

Private Sub mReporteER_Click()
Dim fecha As Date
Dim txtfecha As String
Dim nomarch As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim mata() As Variant

Screen.MousePointer = 11
txtfecha = InputBox("Dame la fecha de la efectividad", , Date)
If IsDate(txtfecha) Then
   fecha = CDate(txtfecha)
   frmProgreso.Show
   mata = LeerEfecRetro(fecha)
   If UBound(mata, 1) <> 0 Then
      nomarch = DirResVaR & "\efec retro " & Format(fecha, "yyyy-mm-dd") & ".txt"
      frmCalVar.CommonDialog1.FileName = nomarch
      frmCalVar.CommonDialog1.ShowSave
      nomarch = frmCalVar.CommonDialog1.FileName
      Open nomarch For Output As #1
      txtcadena = "Clave de la operación" & Chr(9) & "Tipo de operación derivada" & Chr(9) & "Fecha de vencimiento" & Chr(9) & "Días por vencer" & Chr(9) & "Eficiencia retrospectiva" & Chr(9) & "Tipo de cálculo de eficiencia"
      Print #1, txtcadena
      For i = 1 To UBound(mata, 1)
      txtcadena = ""
      For j = 1 To UBound(mata, 2)
      txtcadena = txtcadena & mata(i, j) & Chr(9)
      Next j
      Print #1, txtcadena
      Next i
      Close #1
      MsgBox "Se genero el archivo " & nomarch
   Else
      MsgBox "No hay registros para este dia"
   End If
   Unload frmProgreso
End If

MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub mReporteERRE_Click()
Dim mata() As Double
Dim matb() As Variant
Dim matv() As Variant
Dim valposdiv As Double
Dim fecha1 As Date, fecha2 As Date
Dim sisensib As Boolean
Dim siactvar As Boolean
Dim pint1 As Object
Dim pint2 As Object
Dim siactback As Boolean
Dim siactrespos As Boolean
Dim siactpar As Boolean
Dim siactposdiv As Boolean
Dim siacthcurva As Boolean
Dim siacteh As Boolean
Dim siactvalder As Boolean
Dim siactvar100 As Boolean
Dim siactvcont As Boolean
Dim fecha As Date
Dim txtfecha As String
Dim siesfv As Boolean
Dim nomarch As String
Dim sihayarch As Boolean
Dim strconexion As String
Dim ind As Long
Dim i As Integer
Dim matrp() As Variant
Dim exito As Boolean
Dim suma1 As Double
Dim suma2 As Double
Dim noesc1 As Integer
Dim nosim As Long
Dim htiempo1 As Integer
Dim noesc2 As Integer
Dim htiempo2 As Integer
Dim conadoex As New ADODB.Connection

'rutina de exportacion de datos del sistema de VaR a un archivo de excel
'se ordenan los resultados de var de la forma deseada
If PerfilUsuario = "ADMINISTRADOR" Then
SiActTProc = True
Screen.MousePointer = 11
Set pint1 = frmProgreso.Label2
Set pint1 = frmProgreso.Picture1
Set pint2 = frmProgreso.Picture2

 sisensib = True
 siactvar = True
 siactback = True
 siactrespos = True
 siactpar = True
 siactposdiv = True
 siacthcurva = True
 siacteh = True
 siactvalder = True
 siactvar100 = True
 noesc1 = 500
 htiempo1 = 1
 noesc2 = 500
 htiempo2 = 20
 nosim = 10000

txtfecha = InputBox("dame la fecha a actualizar ", , Date)
If IsDate(txtfecha) Then
   fecha = CDate(txtfecha)
   siesfv = EsFechaVaR(fecha)
   If siesfv Then
      frmProgreso.Show
      Call CrearMatFRiesgo2(fecha, fecha, MatFactRiesgo, "", exito)
'se abre la aplicacion de excel y el archivo de resumen ejecutivo
     nomarch = DirReportes & "\" & NomArchRVaR
     sihayarch = VerifAccesoArch(nomarch)
     If sihayarch Then

     strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & nomarch
     strconexion = strconexion & ";Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
     conadoex.ConnectionString = strconexion
     conadoex.Open

'se abre el archivo de excel para exportar la información de sensibilidades
     If sisensib Then
        matv = LeerSensibNuevo(fecha, txtportCalc2, "Normal", txtportBanobras)
        Call PegarSensResEx(fecha, matv, txtportBanobras, "SENSIB1$", conadoex, RegExcel)
        matv = LeerSensibNuevo(fecha, txtportCalc2, "Normal", "PI CONSERVADOS A VENCIMIENTO")
        Call PegarSensResEx(fecha, matv, "PI CONSERVADOS A VENCIMIENTO", "SENSIB2$", conadoex, RegExcel)
     End If

     If siactvar Then
  'resultados de VaR y escenarios de estres
        For i = 1 To UBound(MatPortSegRiesgo, 1)
            mata = LeerResVarPort(fecha, "Normal", MatPortSegRiesgo(i, 1), noesc1, htiempo1, nosim)
            Call ActVaRRE(fecha, mata, MatPortSegRiesgo(i, 2) & "$", MatPortSegRiesgo(i, 1), conadoex, RegExcel)
        Next i
        
        matb = CalcularCVaRMarginal(fecha)
        Call GuardaResCVaRMarginal(matb, "CVaRMarginal$", conadoex, RegExcel)
        matb = LeerResPIDVDer(fecha)
        Call GuardaResPIDVDer(matb, "PIDV+DER2$", conadoex, RegExcel)
        matb = ObtResCVaREstruc1(fecha)
        suma1 = 0
        For i = 1 To UBound(matb, 1)
            suma1 = suma1 + matb(i, 4)
        Next i
        Call GuardaResEstruct1(matb, "ESTRUC+REL1$", conadoex, RegExcel)
        matb = ObtResCVaREstruc2(fecha)
        suma2 = matb(1, 2)
        If Abs(suma1 - suma2) > 0.01 Then MsgBox "No esta bien construido el portafolio de derivados estructurales"
        Call GuardaResEstruct2(matb, "ESTRUC+REL2$", conadoex, RegExcel)
  'VaR preventivo
       mata = LeerResCVaRPrev(fecha, 20, 1)
       Call GuardarResVaRPrev(fecha, mata, "VARPREV$", conadoex, RegExcel)
  'cvar acumulado
       mata = LeerCVaRPortPos(fecha, txtportCalc2, noesc2, htiempo2, 0, 0.03, 0, "CVARH")
       Call GuardarResVaRAcum(fecha, mata, "CVaRACUM$", conadoex, RegExcel)
  'var exponencial
       mata = LeerResCVaRExp(fecha, "REPORTE PRINCIPAL", 500)
       Call GuardarResVaRExp(fecha, mata, "VAREXP$", conadoex, RegExcel)
     
 End If
 
'resultados del backtesting
If siactback Then
   fecha1 = PBD1(fecha, 1, "MX")
   mata = LeerHistBack(fecha1)
   Call ExportaBackExcel(fecha1, mata, "Backtesting$", conadoex, RegExcel)
End If
'se busca la fecha en el vector de tasas
 ind = BuscarValorArray(fecha, MatFechasVaR, 1)
 If ind <> 0 Then
    If siactrespos Then
'resumen de las posiciones de md e inversion
       matv = LeerResPosMD(fecha, "REPORTE MD")  'leyendo valor posicion activa, pasiva y mtm
       Call GuardaResPosExcel(fecha, matv, "Posicion$", conadoex, RegExcel)
       matv = LeerResPosInversion(fecha)   'leyendo valor posicion activa, pasiva y mtm
       Call GuardaResPosExcel(fecha, matv, "Posicion2$", conadoex, RegExcel)
    End If
    If siactvalder Then
'la valuacion de derivados
       mata = LeerValDerivadosExcel(fecha, 1)
       Call GuardaValDerivadosExcel(fecha, mata, "VALDER$", conadoex, RegExcel)
       mata = LeerValDerivadosExcel(fecha, 2)
       Call GuardaValDerivadosExcel(fecha, mata, "VALDER2$", conadoex, RegExcel)
    End If
  If siactposdiv Then
    'posicion x divisas
     matb = LeerDesgloseDiv(fecha, valposdiv)
     Call ExpDesgloseDiv(fecha, matb, "posxdivACMLE$", conadoex, RegExcel)
  End If
  If siactvar100 Then    'var al 100%
     matv = LeerVaR100MDResumen(fecha, "REPORTE MD", noesc1, valposdiv)
     Call GuardaVaR100MDResumen(fecha, matv, "VaR100_Liq$", conadoex, RegExcel)
  End If
  If siactpar Then
    'parametros para calcular limites
     matb = LeerParamRE(fecha)
     Call GuardaParamRE(matb, "Parametros$", conadoex, RegExcel)
  End If
     If siacteh Then
  'se actualizan los escenarios generados en la simulacion de var
       matb = LeerEscHistRE(fecha, txtportCalc1, "Normal", txtportBanobras, noesc1, htiempo1)
       If UBound(matb, 1) <> 0 Then
          matv = RutinaOrden(matb, 2, SRutOrden)
          Call GuardaEscHistExcel(fecha, matb(1, 1), matb(UBound(matb, 1), 1), matv, "Eschist1$", 10, conadoex, RegExcel)
       End If
       matb = LeerEscHistRE(fecha, txtportCalc1, "Normal", "PI CONSERVADOS A VENCIMIENTO", noesc1, htiempo1)
       If UBound(matb, 1) <> 0 Then
          matv = RutinaOrden(matb, 2, SRutOrden)
          Call GuardaEscHistExcel(fecha, matb(1, 1), matb(UBound(matb, 1), 1), matv, "Eschist2$", 10, conadoex, RegExcel)
       End If
     End If
     'escenarios de estres
     matv = GenCuadroEscEstres(fecha, txtportCalc1)
     Call GuardaReEscEstresExcel(matv, "Estres2$", conadoex, RegExcel)
     
'valuaciones por contraparte
     matv = LeerValContraparte(fecha, 1)
     Call GuardaValContraparte(matv, "VALCONT1$", conadoex, RegExcel)
     matv = LeerValContraparte(fecha, 2)
     Call GuardaValContraparte(matv, "VALCONT2$", conadoex, RegExcel)
     
 End If
 MensajeProc = "Se exportaron los resultados del reporte ejecutivo del " & fecha
 Call GuardaDatosBitacora(3, "Consulta", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
 conadoex.Close
Else
  MsgBox "No hay acceso al archivo " & nomarch
End If
Unload frmProgreso
MsgBox "Proceso terminado"
Call ActUHoraUsuario
SiActTProc = False
End If
End If
Screen.MousePointer = 0
Else
MsgBox "No tiene usted accceso a este modulo"
End If
End Sub

Private Sub mReporteERtxt_Click()
Dim txtfecha As String
Dim p As Integer
Dim valposdiv As Double
Dim nomarch As String
Dim mata() As Variant
Dim matb() As Double
Dim matc() As Variant
Dim txtmensaje As String
Dim i As Integer
Dim noesc1 As Integer
Dim htiempo1 As Integer
Dim txtclave As String
Dim txtcadena As String
Dim j As Integer
Dim noreg As Integer
Dim matrp() As Variant
Dim fecha1 As Date
Dim txtmsg As String
Dim tCambio As Double
Dim exitoarch As Boolean
Dim nosim As Long

Screen.MousePointer = 11
SiActTProc = True
noesc1 = 500
htiempo1 = 1
nosim = 10000

Dim fecha As Date
' lectura de curvas de irs y cetes
txtfecha = InputBox("dame la fecha  ", , Date)
If IsDate(txtfecha) Then
fecha = CDate(txtfecha)
frmProgreso.Show
nomarch = DirResVaR & "\Datos Resumen Ejecutivo " & Format(fecha, "YYYY-MM-DD") & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
'sensibilidades de la posicion
mata = LeerSensibNuevo(fecha, txtportCalc2, "Normal", txtportBanobras)
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Sensibilidades de la posición"
   Print #1, txtmensaje
   noreg = UBound(mata, 1)
   For i = 1 To Minimo(noreg, 20)
       txtclave = Format(i, "000") & CLng(fecha) & txtportBanobras
       txtcadena = txtclave & Chr(9)
       txtcadena = txtcadena & txtportBanobras & Chr(9)            'nombre del portafolio
       txtcadena = txtcadena & fecha & Chr(9)                       'fecha
       txtcadena = txtcadena & mata(noreg - i + 1, 1) & Chr(9)      'nombre del factor
       txtcadena = txtcadena & mata(noreg - i + 1, 3) & Chr(9)      'curva
       txtcadena = txtcadena & mata(noreg - i + 1, 4) & Chr(9)      'plazo
       txtcadena = txtcadena & mata(noreg - i + 1, 6) & Chr(9)      'valor
       txtcadena = txtcadena & mata(noreg - i + 1, 11) & Chr(9)     'volatilidad
       txtcadena = txtcadena & mata(noreg - i + 1, 10)              'sensibilidad
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
mata = LeerSensibNuevo(fecha, txtportCalc2, "Normal", "PI CONSERVADOS A VENCIMIENTO")

If UBound(mata, 1) <> 0 Then
   txtmensaje = "Sensibilidades de la posición"
   Print #1, txtmensaje
   noreg = UBound(mata, 1)
   For i = 1 To Minimo(noreg, 20)
       txtclave = Format(i, "000") & CLng(fecha) & "PI CONSERVADOS A VENCIMIENTO"
       txtcadena = txtclave & Chr(9)
       txtcadena = txtcadena & "PI CONSERVADOS A VENCIMIENTO" & Chr(9)               'nombre del portafolio
       txtcadena = txtcadena & fecha & Chr(9)                       'fecha
       txtcadena = txtcadena & mata(noreg - i + 1, 1) & Chr(9)      'nombre del factor
       txtcadena = txtcadena & mata(noreg - i + 1, 3) & Chr(9)      'curva
       txtcadena = txtcadena & mata(noreg - i + 1, 4) & Chr(9)      'plazo
       txtcadena = txtcadena & mata(noreg - i + 1, 6) & Chr(9)      'valor
       txtcadena = txtcadena & mata(noreg - i + 1, 11) & Chr(9)     'volatilidad
       txtcadena = txtcadena & mata(noreg - i + 1, 10)              'sensibilidad
       Print #1, txtcadena
   Next i
   Print #1, ""
End If

'parametros del resumen ejecutivo
  mata = LeerParamRE(fecha)
  If UBound(mata, 1) <> 0 Then
   txtmensaje = "Parametros del resumen ejecutivo"
   Print #1, txtmensaje
   For i = 1 To UBound(mata, 1)
       txtcadena = ""
       For j = 1 To UBound(mata, 2)
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
tCambio = mata(1, 3)
'detalle de la posicion de divisas
mata = LeerDesgloseDiv(fecha, valposdiv)
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Detalle de la posición de divisas"
   Print #1, txtmensaje
   For i = 1 To UBound(mata, 1)
       txtcadena = ""
       For j = 1 To UBound(mata, 2)
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
 'var por portafolio
 txtmensaje = "VaR por portafolio"
 Print #1, txtmensaje
  For p = 1 To UBound(MatPortSegRiesgo, 1)
      txtmensaje = "VaR " & MatPortSegRiesgo(p, 1)
      Print #1, txtmensaje
      txtcadena = "Fecha" & Chr(9)
      txtcadena = txtcadena & "Lim inf VaR Hist" & Chr(9)
      txtcadena = txtcadena & "Lim sup VaR Hist" & Chr(9)
      txtcadena = txtcadena & "Lim inf VaR Mark" & Chr(9)
      txtcadena = txtcadena & "Lim sup VaR Mark" & Chr(9)
      txtcadena = txtcadena & "Lim inf VaR Mont" & Chr(9)
      txtcadena = txtcadena & "Lim sup VaR Mont" & Chr(9)
      Print #1, txtcadena
      matb = LeerResVarPort(fecha, "Normal", MatPortSegRiesgo(p, 1), noesc1, htiempo1, nosim)
      If UBound(matb, 1) <> 0 Then
         txtcadena = fecha & Chr(9)
         For i = 1 To UBound(matb, 1)
             txtcadena = txtcadena & matb(i) & Chr(9)
         Next i
         Print #1, txtcadena
         Print #1, ""
       End If
  Next p
'cvar acumulado
      matb = LeerCVaRPortPos(fecha, txtportCalc2, 500, 20, 0, 0.03, 0, "CVARH")
      txtmensaje = "CVaR acumulado"
      Print #1, txtmensaje
      txtcadena = "Fecha" & Chr(9)
      For i = 1 To UBound(MatPortSegRiesgo, 1)
          txtcadena = txtcadena & MatPortSegRiesgo(i, 1) & Chr(9)
      Next i
      Print #1, txtcadena
      txtcadena = fecha & Chr(9)
      For i = 1 To UBound(matb, 1)
          txtcadena = txtcadena & matb(i) & Chr(9)
      Next i
      Print #1, txtcadena
      Print #1, ""
'var preventivo
matb = LeerResCVaRPrev(fecha, 20, 1)
If UBound(matb, 1) <> 0 Then
   txtmensaje = "VaR Preventivo"
   Print #1, txtmensaje
   txtmensaje = "Fecha" & Chr(9)
   For j = 1 To UBound(MatPortSegRiesgo, 1)
       txtmensaje = txtmensaje & "VaR " & MatPortSegRiesgo(j, 1) & Chr(9)
   Next j
   Print #1, txtmensaje
   txtmensaje = fecha & Chr(9)
   For j = 1 To UBound(matb, 1)
       txtmensaje = txtmensaje & matb(j) & Chr(9)
   Next j
   Print #1, txtmensaje
   Print #1, ""
End If

'var exponecial
matb = LeerResCVaRExp(fecha, "REPORTE PRINCIPAL", 500)
If UBound(matb, 1) <> 0 Then
     txtmensaje = "VaR exponencial"
     Print #1, txtmensaje
     txtmensaje = "Fecha" & Chr(9)
     For j = 1 To UBound(MatPortSegRiesgo, 1)
         txtmensaje = txtmensaje & "VaR " & MatPortSegRiesgo(j, 1) & Chr(9)
     Next j
     Print #1, txtmensaje
     txtmensaje = fecha & Chr(9)
     For j = 1 To UBound(matb, 1)
         txtmensaje = txtmensaje & matb(j) & Chr(9)
     Next j
     Print #1, txtmensaje
     Print #1, ""
End If
     
'escenarios de estres 1
      mata = GenCuadroEscEstres(fecha, txtportCalc1)
      Print #1, "Escenarios de estres"
      txtcadena = "Clave" & Chr(9)
      txtcadena = txtcadena & "Fecha" & Chr(9)
      txtcadena = txtcadena & "Orden" & Chr(9)
      txtcadena = txtcadena & "Portafolio" & Chr(9)
      txtcadena = txtcadena & "3 desv estandar" & Chr(9)
      txtcadena = txtcadena & "Ad Hoc 1" & Chr(9)
      txtcadena = txtcadena & "Ad Hoc 2" & Chr(9)
      txtcadena = txtcadena & "Global 1" & Chr(9)
      txtcadena = txtcadena & "Global 2" & Chr(9)
      txtcadena = txtcadena & "Global 3" & Chr(9)
      txtcadena = txtcadena & "Global 4" & Chr(9)
      txtcadena = txtcadena & "Deuda Estatal alarmante" & Chr(9)
      txtcadena = txtcadena & "Elecc pres EU 1" & Chr(9)
      txtcadena = txtcadena & "Elecc pres EU 2" & Chr(9)
      txtcadena = txtcadena & "Jueves Negro" & Chr(9)
      txtcadena = txtcadena & "Taylor 1" & Chr(9)
      txtcadena = txtcadena & "Taylor 2" & Chr(9)
      txtcadena = txtcadena & "Taylor 3" & Chr(9)
   
      Print #1, txtcadena
      For i = 1 To UBound(mata, 1)
          txtcadena = ""
          For j = 1 To UBound(mata, 2)
              If Not EsVariableVacia(mata(i, j)) Then
                 txtcadena = txtcadena & CStr(mata(i, j)) & Chr(9)
              Else
                 txtcadena = txtcadena & 0 & Chr(9)
              End If
          Next j
          Print #1, txtcadena
      Next i
      Print #1, ""
  

'valuacion de derivados
  matb = LeerValDerivadosExcel(fecha, 1)
  If UBound(matb, 1) <> 0 Then
     txtmensaje = "Valuacion de Derivados 1"
     Print #1, txtmensaje
     txtcadena = fecha & Chr(9)
     For i = 1 To UBound(matb, 1)
         txtcadena = txtcadena & matb(i) & Chr(9)
     Next i
     Print #1, txtcadena
     Print #1, ""
  End If
  matb = LeerValDerivadosExcel(fecha, 2)
  If UBound(matb, 1) <> 0 Then
     txtmensaje = "Valuacion de Derivados 2"
     Print #1, txtmensaje
     txtcadena = fecha & Chr(9)
     For i = 1 To UBound(matb, 1)
         txtcadena = txtcadena & matb(i) & Chr(9)
     Next i
     Print #1, txtcadena
     Print #1, ""
  End If

'var al 100%
  mata = LeerVaR100MDResumen(fecha, "REPORTE MD", 500, valposdiv)
  If UBound(mata, 1) <> 0 Then
     txtmensaje = "VaR al 100%"
     Print #1, txtmensaje
     txtmensaje = "Fecha" & Chr(9)
     For i = 1 To UBound(mata, 1)
        txtmensaje = txtmensaje & mata(i, 1) & Chr(9)
     Next i
     Print #1, txtmensaje
     txtcadena = fecha & Chr(9)
     For i = 1 To UBound(mata, 1)
         txtcadena = txtcadena & mata(i, 2) & Chr(9)
     Next i
     Print #1, txtcadena
     Print #1, ""
  End If
  
'escenarios historicos del portafolio consolidado
mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", txtportBanobras, 500, 1)
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Escenarios historicos del portafolio de BANOBRAS"
   Print #1, txtmensaje
   Print #1, ""
   txtcadena = ""
   txtcadena = txtcadena & fecha & Chr(9)                    'la fecha
   txtcadena = txtcadena & mata(1, 1) & Chr(9)               'la fecha inicial
   txtcadena = txtcadena & mata(UBound(mata, 1), 1) & Chr(9) 'la fecha final
   mata = RutinaOrden(mata, 2, SRutOrden)
   For i = 1 To 10
       txtcadena = txtcadena & mata(i, 1) & Chr(9)
   Next i
   For i = 1 To 10
       txtcadena = txtcadena & mata(i, 2) & Chr(9)
   Next i
   
   Print #1, txtcadena
   Print #1, ""
End If

'escenarios historicos del PI CONSERVADOS A VENCIMIENTO
mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", "PI CONSERVADOS A VENCIMIENTO", 500, 1)
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Escenarios historicos del PI CONSERVADOS A VENCIMIENTO"
   Print #1, txtmensaje
   Print #1, ""
   txtcadena = ""
   txtcadena = txtcadena & fecha & Chr(9)                    'la fecha
   txtcadena = txtcadena & mata(1, 1) & Chr(9)               'la fecha inicial
   txtcadena = txtcadena & mata(UBound(mata, 1), 1) & Chr(9) 'la fecha final
   mata = RutinaOrden(mata, 2, SRutOrden)
   For i = 1 To 10
       txtcadena = txtcadena & mata(i, 1) & Chr(9)
   Next i
   For i = 1 To 10
       txtcadena = txtcadena & mata(i, 2) & Chr(9)
   Next i
   
   Print #1, txtcadena
   Print #1, ""
End If
'resumen de posicion de mercado de dinero
mata = LeerResPosMD(fecha, "REPORTE MD") 'leyendo dxv dxvc mtm
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Resumen de la posicion de mercado de dinero"
   Print #1, txtmensaje
   txtcadena = fecha & Chr(9)
   For i = 1 To UBound(mata, 1)
   For j = 1 To UBound(mata, 2)
    txtcadena = txtcadena & mata(i, j) & Chr(9)
   Next j
   Next i
   Print #1, txtcadena
   Print #1, ""
End If
'resumen posicion de inversion
mata = LeerResPosInversion(fecha)
If UBound(mata, 1) <> 0 Then
   txtmensaje = "Resumen de la posicion de inversión"
   Print #1, txtmensaje
   txtcadena = fecha & Chr(9)
   For i = 1 To UBound(mata, 1)
   For j = 1 To UBound(mata, 2)
    txtcadena = txtcadena & mata(i, j) & Chr(9)
   Next j
   Next i
   Print #1, txtcadena
   Print #1, ""
End If


'backtesting
fecha1 = PBD1(fecha, 1, "MX")
matb = LeerHistBack(fecha1)
If UBound(matb, 1) <> 0 Then
       Print #1, "Backtesting"
       txtcadena = "Fecha" & Chr(9)
       noreg = UBound(MatGruposPortPos, 1)
       For i = 1 To noreg - 5
           txtcadena = txtcadena & MatGruposPortPos(i, 3) & Chr(9)
       Next i
       Print #1, txtcadena
       txtcadena = fecha1 & Chr(9)
       For i = 1 To UBound(matb, 1)
           txtcadena = txtcadena & matb(i) & Chr(9)
       Next i
       Print #1, txtcadena
       Print #1, ""
End If
matc = LeerValContraparte(fecha, 1)
If UBound(matc, 1) <> 0 Then
   Print #1, "Valuacion por contraparte mercado"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       txtcadena = txtcadena & matc(i, 4) & Chr(9)
       txtcadena = txtcadena & matc(i, 5) & Chr(9)
       txtcadena = txtcadena & matc(i, 6)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
matc = LeerValContraparte(fecha, 2)
If UBound(matc, 1) <> 0 Then
   Print #1, "Valuacion por contraparte Banxico"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       txtcadena = txtcadena & matc(i, 4) & Chr(9)
       txtcadena = txtcadena & matc(i, 5) & Chr(9)
       txtcadena = txtcadena & matc(i, 6)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
matc = LeerResPIDVDer(fecha)
If UBound(matc, 1) <> 0 Then
   Print #1, "CVaR PIDV+DERIVADOS"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       txtcadena = txtcadena & matc(i, 4) & Chr(9)
       txtcadena = txtcadena & matc(i, 5)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If

'var marginal
matc = CalcularCVaRMarginal(fecha)
If UBound(matc, 1) <> 0 Then
   Print #1, "CVaR Marginal"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       txtcadena = txtcadena & matc(i, 4) & Chr(9)
       txtcadena = txtcadena & matc(i, 5) & Chr(9)
       txtcadena = txtcadena & matc(i, 6)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If

matc = ObtResCVaREstruc1(fecha)
If UBound(matc, 1) <> 0 Then
   Print #1, "CVaR estructural 1"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       txtcadena = txtcadena & matc(i, 4) & Chr(9)
       txtcadena = txtcadena & matc(i, 5) & Chr(9)
       txtcadena = txtcadena & matc(i, 6) & Chr(9)
       txtcadena = txtcadena & matc(i, 7)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If
matc = ObtResCVaREstruc2(fecha)
If UBound(matc, 1) <> 0 Then
   Print #1, "CVaR estructural 2"
   For i = 1 To UBound(matc, 1)
       txtcadena = matc(i, 1) & Chr(9)
       txtcadena = txtcadena & matc(i, 2) & Chr(9)
       txtcadena = txtcadena & matc(i, 3) & Chr(9)
       Print #1, txtcadena
   Next i
   Print #1, ""
End If

If tCambio <> 0 Then
  txtmsg = GenerarComentarioRE(fecha, valposdiv / tCambio)
Else
  txtmsg = ""
End If
Print #1, txtmsg
Close #1
End If
Unload frmProgreso
End If
MsgBox "Fin del proceso"
Call ActUHoraUsuario
SiActTProc = False
Screen.MousePointer = 0
End Sub

Private Sub mReporteFVS_Click()
Dim fecha As Date
Dim txtfecha As String

Screen.MousePointer = 11
txtfecha = InputBox("dame la fecha a generar", , Date)
If IsDate(txtfecha) Then
   SiActTProc = True
   fecha = CDate(txtfecha)
   frmProgreso.Show
   Call GeneraFlujosValSwaps(fecha)
   Unload frmProgreso
   Call ActUHoraUsuario
   SiActTProc = False
Else
   MsgBox "No es una fecha valida"
End If
Screen.MousePointer = 0

End Sub

Private Sub mReporteHVS_Click()
Screen.MousePointer = 11
frmResumenVaR.Show
Screen.MousePointer = 0
End Sub

Private Sub mReporteROD_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim fechaa As Date
Dim fechab As Date

Dim tfecha1 As String
Dim tfecha2 As String

Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfechaa As String
Dim txtfechab As String
Dim txtcadena As String
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim contar As Integer
Dim i As Integer
Dim j As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim noreg3 As Integer
Dim noreg4 As Integer
Dim noreg5 As Integer
Dim noreg6 As Integer
Dim noreg7 As Integer
Dim noreg8 As Integer
Dim norega As Integer
Dim noregb As Integer
Dim ccontra1 As String
Dim indice As Long
Dim indice2 As Long
Dim mnocional As Double
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim matd() As Variant
Dim mate() As Variant
Dim matf() As Variant
Dim matg() As Variant
Dim matj() As Variant
Dim matjj() As Variant
Dim txtsalida As String
Dim nomarch As String
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
tfecha1 = InputBox("Dame la ultima fecha del mes anterior ", , Date)
tfecha2 = InputBox("Dame la ultima fecha del mes ", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
   frmProgreso.Show
   fecha1 = CDate(tfecha1)
   fecha2 = CDate(tfecha2)
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
'swaps que iniciaron en el periodo
mata = DetSwapsInicianPer(fecha1, fecha2)
'se validan los swaps que cambiaron de intencion en el mes
matb = DetSwapCambIntencion(fecha1, fecha2)
noreg2 = UBound(matb, 1)

'SE ENUMERAN LOS SWAPS QUE VENCIERON EN EL MES
matc = DetPosVenceSwaps(fecha1, fecha2)
noreg3 = UBound(matc, 1)
For i = 1 To noreg3
    txtfecha3 = "to_date('" & Format(matc(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "SELECT MAX(FFINAL) FROM " & TablaFlujosSwapsO & " WHERE FECHAREG = " & txtfecha3 & " AND COPERACION  = '" & matc(i, 1) & "' AND TIPOPOS ='1'"
    rmesa.Open txtcadena, ConAdo
    matc(i, 10) = rmesa.Fields(0)
    rmesa.Close
Next i
contar = 0
'se cuentan las operaciones unwind de swaps
For i = 1 To noreg3
    If matc(i, 6) < matc(i, 10) Then
       matc(i, 11) = True
    Else
       matc(i, 11) = False
    End If
Next i
'fwds que entraron
matd = DetFwdsPactados(fecha1, fecha2)
noreg4 = UBound(matd, 1)
'fwds que cambiaron de intencion
mate = DetFwdCambIntencion(fecha1, fecha2)
'fwds que vencen
matf = DetFwdsVencen(fecha1, fecha2)
noreg5 = UBound(mate, 1)
'forwards unwind
matg = detFwdUnwind(fecha1, fecha2)


'RESUMEN EFICIENCIA COBERTURA
matjj = CalcResEfCob(fecha1)
matj = CalcResEfCob(fecha2)

noreg8 = UBound(matj, 1)

nomarch = DirResVaR & "\Resumen operaciones derivados periodo " & Format(fecha1, "yyyy-mm-dd") & "  " & Format(fecha2, "yyyy-mm-dd") & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Open nomarch For Output As #1
If UBound(mata, 1) <> 0 Then
   Print #1, "Swaps pactados"
   Print #1, "Fecha de concertación" & Chr(9) & "Producto" & Chr(9) & "Intención" & Chr(9) & "Vencimiento" & Chr(9) & "Moneda" & Chr(9) & "Nocional (cifras en millones)" & Chr(9) & "Contraparte" & Chr(9) & "Observaciones" & Chr(9) & "Clave operacion (no se incluye)"
   For i = 1 To UBound(mata, 1)
       txtsalida = mata(i, 3) & Chr(9) & mata(i, 4) & Chr(9) & mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) & Chr(9) & mata(i, 8) & Chr(9) & mata(i, 9) & Chr(9) & Chr(9) & mata(i, 1)
       Print #1, txtsalida
   Next i
   Print #1, ""
End If
If noreg4 <> 0 Then
   Print #1, "Forwards que entraron"
   Print #1, "Fecha de concertación" & Chr(9) & "Producto" & Chr(9) & "Intención" & Chr(9) & "Vencimiento" & Chr(9) & "Moneda" & Chr(9) & "Nocional (cifras en millones)" & Chr(9) & "Contraparte" & Chr(9) & "Observaciones" & Chr(9) & "Clave operacion (no se incluye)"
   For i = 1 To UBound(matd, 1)
       txtsalida = matd(i, 2) & Chr(9) & matd(i, 3) & Chr(9) & matd(i, 4) & Chr(9) & matd(i, 5) & Chr(9) & matd(i, 6) & Chr(9) & matd(i, 7) & Chr(9) & matd(i, 8) & Chr(9) & Chr(9) & matd(i, 1)
       Print #1, txtsalida
   Next i
Print #1, ""
End If
If UBound(matb, 1) <> 0 Then
   Print #1, "Swaps que cambiaron de intencion"
   Print #1, "fecha de concertación" & Chr(9) & "Producto" & Chr(9) & "Intención" & Chr(9) & "Vencimiento" & Chr(9) & "Moneda" & Chr(9) & "Nocional" & Chr(9) & " Contraparte " & Chr(9) & "Observaciones" & Chr(9) & "Clave de operación"
   For i = 1 To UBound(matb, 1)
       txtsalida = ""
       txtsalida = matb(i, 1) & Chr(9) & matb(i, 2) & Chr(9) & matb(i, 3) & Chr(9) & matb(i, 4) & Chr(9) & matb(i, 5) & Chr(9) & matb(i, 6) & Chr(9) & matb(i, 7) & Chr(9) & matb(i, 8) & Chr(9) & matb(i, 9)
       Print #1, txtsalida
   Next i
   Print #1, ""
End If
If UBound(matc, 1) <> 0 Then
Print #1, "Swaps que vencieron"
Print #1, "Fecha de concertación" & Chr(9) & "Producto" & Chr(9) & "Intencion" & Chr(9) & "Vencimiento" & Chr(9) & "Moneda" & Chr(9) & "Nocional" & Chr(9) & "Contraparte" & Chr(9) & "Observaciones " & Chr(9) & "Fecha del ultimo flujos" & Chr(9) & "Unwind" & Chr(9) & "Clave de operacion"
For i = 1 To UBound(matc, 1)
    txtsalida = matc(i, 3) & Chr(9) & matc(i, 4) & Chr(9) & matc(i, 5) & Chr(9) & matc(i, 6) & Chr(9) & matc(i, 7) & Chr(9) & matc(i, 8) & Chr(9) & matc(i, 9) & Chr(9) & Chr(9) & matc(i, 10) & Chr(9) & matc(i, 11) & Chr(9) & matc(i, 1)
    Print #1, txtsalida
Next i
Print #1, ""
End If
If UBound(mate, 1) <> 0 Then
Print #1, "Forwards que cambiaron de intencion"
Print #1, "Clave de operación" & Chr(9) & "Intención" & Chr(9) & "Fecha cambio" & Chr(9) & "Fecha de inicio" & Chr(9) & "Fecha de vencimiento" & Chr(9) & "Tipo de operacion" & Chr(9) & "Contraparte"
For i = 1 To UBound(mate, 1)
    txtsalida = ""
    For j = 1 To UBound(mate, 2)
        txtsalida = txtsalida & mate(i, j) & Chr(9)
    Next j
    Print #1, txtsalida
Next i
Print #1, ""
End If
If UBound(matf, 1) <> 0 Then
   Print #1, "Forwards que vencieron"
   Print #1, "Fecha de concertación" & Chr(9) & "Producto" & Chr(9) & "Intención" & Chr(9) & "Vencimiento" & Chr(9) & "Moneda" & Chr(9) & "Nocional (cifras en millones)" & Chr(9) & "Contraparte" & Chr(9) & "Observaciones" & Chr(9) & "Clave de operacion"
   For i = 1 To UBound(matf, 1)
       txtsalida = matf(i, 2) & Chr(9) & matf(i, 3) & Chr(9) & matf(i, 4) & Chr(9) & matf(i, 5) & Chr(9) & matf(i, 6) & Chr(9) & matf(i, 7) & Chr(9) & matf(i, 8) & Chr(9) & matf(i, 9) & Chr(9) & matf(i, 1)
       Print #1, txtsalida
   Next i
   Print #1, ""
End If

If UBound(matg, 1) <> 0 Then
   Print #1, "Forwards que vencieron anticipadamente"
End If
If noreg8 <> 0 And Not EsArrayVacio(matjj) Then
Print #1, "Eficiencia de la cobertura"
Print #1, "Producto" & Chr(9) & "# de instrumentos" & fecha2 & Chr(9) & "# de instrumentos" & fecha1 & Chr(9) & "Promedio" & Chr(9) & "Mínimo" & Chr(9) & "Máximo"
For i = 1 To UBound(matj, 1)
    If matj(i, 2) <> 0 Then
       txtsalida = matj(i, 1) & Chr(9) & matj(i, 2) & Chr(9) & matjj(i, 2) & Chr(9) & Format(matj(i, 3), "#000.00 %") & Chr(9) & Format(matj(i, 4), "#000.00 %") & Chr(9) & Format(matj(i, 5), "#000.00 %")
       Print #1, txtsalida
    End If
Next i



End If
Close #1
Unload frmProgreso
MsgBox "Fin de proceso"
End If
Screen.MousePointer = 0
End Sub

Function DetFwdsVencen(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim noreg6 As Integer
Dim noreg7 As Integer
Dim i As Integer
Dim estruc As String
Dim idcontrap As Integer
Dim indice As Integer
Dim f_val As Date
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena2 = "SELECT COPERACION,INTENCION,FINICIO,FVENCIMIENTO,CPRODUCTO,M_NOCIONAL,ID_CONTRAP,ESTRUCTURAL FROM " & TablaPosFwd
txtcadena2 = txtcadena2 & " WHERE (FECHAREG,COPERACION) IN (SELECT MAX(FECHAREG),COPERACION FROM " & TablaPosFwd
txtcadena2 = txtcadena2 & " WHERE FVENCIMIENTO > " & txtfecha1 & " AND FVENCIMIENTO <= " & txtfecha2
txtcadena2 = txtcadena2 & " AND TIPOPOS = 1 GROUP BY COPERACION) ORDER BY COPERACION"
txtcadena1 = "SELECT COUNT(*) FROM (" & txtcadena2 & ")"
rmesa.Open txtcadena1, ConAdo
noreg6 = rmesa.Fields(0)
rmesa.Close
If noreg6 <> 0 Then
   rmesa.Open txtcadena2, ConAdo
   ReDim matf(1 To noreg6, 1 To 9) As Variant
   For i = 1 To noreg6
       matf(i, 1) = rmesa.Fields("COPERACION")                'clave de operacion
       matf(i, 2) = rmesa.Fields("FINICIO")                   'fecha de inicio
       matf(i, 3) = rmesa.Fields("CPRODUCTO")                 'tipo de operacion
       estruc = rmesa.Fields("ESTRUCTURAL")                   'tipo de operacion
       If rmesa.Fields("INTENCION") = "N" And estruc = "S" Then               'INTENCION
          matf(i, 4) = "Negociación estructural"
       ElseIf rmesa.Fields("INTENCION") = "N" And estruc = "N" Then               'INTENCION
         matf(i, 4) = "Negociación"
       ElseIf rmesa.Fields("INTENCION") = "C" Then     'INTENCION
          matf(i, 4) = "Cobertura"
       ElseIf rmesa.Fields("INTENCION") = "R" Then     'INTENCION
          matf(i, 4) = "Reclasifiacion"
       End If
       matf(i, 5) = rmesa.Fields("FVENCIMIENTO")           'fecha vencimiento
       matf(i, 6) = Right(matf(i, 3), 3)
       matf(i, 7) = rmesa.Fields("M_NOCIONAL") / 1000000  'monto nocional
       idcontrap = rmesa.Fields("ID_CONTRAP")           'contraparte
       indice = BuscarValorArray(idcontrap, MatContrapartes, 1)
       If indice <> 0 Then
          matf(i, 8) = MatContrapartes(indice, 3)
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg6
       txtcadena2 = "SELECT max(FECHAREG) as MFECHAREG,FVENCIMIENTO FROM " & TablaPosFwd
       txtcadena2 = txtcadena2 & " WHERE FVENCIMIENTO > " & txtfecha1 & " AND FVENCIMIENTO <= " & txtfecha2
       txtcadena2 = txtcadena2 & " AND COPERACION = '" & matf(i, 1) & "'"
       txtcadena2 = txtcadena2 & " AND TIPOPOS = 1 GROUP BY FVENCIMIENTO"
       txtcadena1 = "SELECT COUNT(*) FROM (" & txtcadena2 & ")"
       rmesa.Open txtcadena1, ConAdo
       noreg7 = rmesa.Fields(0)
       rmesa.Close
       rmesa.Open txtcadena2, ConAdo
       f_val = rmesa.Fields("MFECHAREG")
       If rmesa.Fields("FVENCIMIENTO") - f_val > 3 Then
          matf(i, 9) = "Vencimiento anticipado"
       Else
          matf(i, 9) = "Vencimiento natural"
       End If
       rmesa.Close
   Next i
Else
  ReDim matf(0 To 0, 0 To 0) As Variant
End If
DetFwdsVencen = matf

End Function

Private Sub mReporteTCCF_Click()
Dim fecha As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
txtfecha = InputBox("Dame la fecha", , Date)
fecha = CDate(txtfecha)
txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "select * from " & TablaFlujosSwapsO & " where FINICIO = " & txtfecha1 & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg, 1 To 9) As Variant
rmesa.MoveFirst
For i = 1 To noreg
mata(i, 1) = rmesa.Fields(1) '
mata(i, 2) = rmesa.Fields(5) '
mata(i, 3) = rmesa.Fields(6) '
mata(i, 4) = rmesa.Fields(7) '
mata(i, 5) = rmesa.Fields(8) '
mata(i, 6) = rmesa.Fields(9) '
mata(i, 7) = rmesa.Fields(12) '
mata(i, 8) = rmesa.Fields(13) '
mata(i, 9) = rmesa.Fields(14) '
rmesa.MoveNext
Next i
rmesa.Close


frmCuponesCortan.Show
frmCuponesCortan.MSFlexGrid1.Rows = noreg + 1
frmCuponesCortan.MSFlexGrid1.Cols = 10
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 1) = "Fecha de registro"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 2) = "Clave de operacion"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 3) = "Tipo de posicion"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 4) = "Fecha de inicio"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 5) = "Fecha final"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 6) = "Fecha de descuento"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 7) = "Saldo"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 8) = "Amortizacion"
frmCuponesCortan.MSFlexGrid1.TextMatrix(0, 9) = "Tasa cupon"
For i = 1 To noreg
    For j = 1 To 9
        frmCuponesCortan.MSFlexGrid1.TextMatrix(i, j) = mata(i, j)
    Next j
Next i
MsgBox "Fin del proceso"
Screen.MousePointer = 0

End Sub

Private Sub mReporteVMD_Click()
Dim tfecha As String
Dim fecha As Date
Screen.MousePointer = 11
'primero se obtiene la fecha del reporte
tfecha = InputBox("Dame la fecha del reporte", , Date)
If IsDate(tfecha) Then
   Call ImpRepVaRMD(CDate(tfecha))
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mReporteVR_Click()
Screen.MousePointer = 11
  frmRepValReemplazo.Show
Screen.MousePointer = 0
End Sub

Private Sub mReporteVT_Click()
Dim tfecha As String
Dim fecha As Date
Screen.MousePointer = 11
'primero se obtiene la fecha del reporte
tfecha = InputBox("Dame la fecha del reporte", , Date)
If IsDate(tfecha) Then
   Call ImpRepVaRTeso(CDate(tfecha))
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub mReporteWRW_Click()
Screen.MousePointer = 11
       Dim nomarch As String
       Dim tfecha As String
       Dim txtcadena As String
       Dim valcvatot As Double
       Dim txtfecha As String
       Dim txtfiltro1 As String
       Dim txtfiltro2 As String
       Dim noreg As Long
       Dim i As Long
       Dim j As Long
       Dim fecha As Date
       Dim rmesa As New ADODB.recordset
       
       tfecha = InputBox("Dame la fecha del proceso", , Date)
       If IsDate(tfecha) Then
       fecha = CDate(tfecha)
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT * FROM " & TablaResWRW & " WHERE FECHA = " & txtfecha
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          ReDim mata(1 To noreg, 1 To 22) As Variant
          rmesa.Open txtfiltro2, ConAdo
          ReDim matresul1(1 To 9, 1 To 2) As Double
          ReDim matresul2(1 To 9, 1 To 2) As Double
          valcvatot = 0
          For i = 1 To noreg
              mata(i, 1) = rmesa.Fields("ID_CONTRAP")
              mata(i, 2) = rmesa.Fields("VAL_CVA")
              For j = 1 To 9
                  mata(i, j + 2) = rmesa.Fields(j + 2)
                  matresul1(j, 1) = matresul1(j, 1) + mata(i, j + 2)
              Next j
              For j = 1 To 9
                  mata(i, j + 11) = rmesa.Fields(j + 11)
                  matresul2(j, 1) = matresul2(j, 1) + mata(i, j + 11)
              Next j
              rmesa.MoveNext
              valcvatot = valcvatot + mata(i, 2)
          Next i
          rmesa.Close
          Dim matpor(1 To 9) As Double
          matpor(1) = 0.8: matpor(2) = 0.84: matpor(3) = 0.85: matpor(4) = 0.9
          matpor(5) = 0.96: matpor(6) = 0.97: matpor(7) = 0.975: matpor(8) = 0.98
          matpor(9) = 0.99
          nomarch = DirResVaR & "\Reporte Wrong Risk Way " & Format$(fecha, "yyyy-mm-dd") & ".txt"
          frmCalVar.CommonDialog1.FileName = nomarch
          frmCalVar.CommonDialog1.ShowSave
          nomarch = frmCalVar.CommonDialog1.FileName
          Open nomarch For Output As #3
          txtcadena = "Sector" & Chr(9) & "Contraparte" & Chr$(9) & "CVA" & Chr$(9)
          For i = 1 To 9
              txtcadena = txtcadena & Format$(matpor(i), "##0.00 %") & Chr$(9)
          Next i
          For i = 1 To 9
              txtcadena = txtcadena & Format$(matpor(i), "##0.00 %") & Chr$(9)
          Next i
          Print #3, txtcadena
          For i = 1 To noreg
              txtcadena = mata(i, 1) & Chr(9)
              txtcadena = txtcadena & mata(i, 2) & Chr(9)
              For j = 1 To 9
                  txtcadena = txtcadena & mata(i, j + 2) & Chr(9)
              Next j
              For j = 1 To 9
                  txtcadena = txtcadena & mata(i, j + 11) & Chr(9)
              Next j
              Print #3, txtcadena
          Next i
          Print #3, ""
          txtcadena = "CVA Derivados " & Chr$(9) & valcvatot
          Print #3, txtcadena
          txtcadena = Chr$(9) & "Con recuperacion" & Chr$(9) & Chr$(9) & "Sin recuperación"
          Print #3, txtcadena
          txtcadena = "alfa" & Chr$(9) & "WWR" & Chr$(9) & "WWR/CVA" & Chr$(9) & "WWR" & Chr$(9) & "WWR/CVA"
          Print #3, txtcadena
          For i = 1 To 9
              matresul1(i, 2) = matresul1(i, 1) / valcvatot
              matresul2(i, 2) = matresul2(i, 1) / valcvatot
              txtcadena = matpor(i) & Chr$(9) & matresul1(i, 1) & Chr$(9) & matresul1(i, 2) & Chr$(9) & matresul2(i, 1) & Chr$(9) & matresul2(i, 2)
              Print #3, txtcadena
          Next i
          Close #3
          MsgBox "Fin de proceso"
       Else
         MsgBox "No hay registros"
       End If
       End If
       
Screen.MousePointer = 0
End Sub

Private Sub mRepValProc_Click()
'rutina para obtener los datos de validacion de procesos de resumen ejecutivo
Dim tfecha As String
Dim fecha As Date
Dim suma As Double
Dim txtnomarch As String
Dim noesc As Integer
Dim mata() As Variant
Dim txtcadena As String
Dim i As Long
Dim j As Long

tfecha = InputBox("Dame la fecha de proceso", , Date)
 If IsDate(tfecha) Then
    Screen.MousePointer = 11
    noesc = 500
    fecha = CDate(tfecha)
    txtnomarch = DirResVaR & "\Resultados de validacion de procesos " & Format(fecha, "yyyy-mm-dd") & ".txt"
    frmReportes.CommonDialog1.FileName = txtnomarch
    frmReportes.CommonDialog1.ShowSave
    txtnomarch = frmReportes.CommonDialog1.FileName
    Open txtnomarch For Output As #1
    Call ValidarValPosMD(fecha, mata)
    If UBound(mata, 1) > 0 Then
       Print #1, "Diferencias de valuación entre SIVARMER y PIP para los instrumentos de Deuda"
       txtcadena = "Clave de posicion" & Chr(9) & "Clave de operación" & Chr(9) & "TV" & Chr(9) & "Emision" & Chr(9) & "Serie" & Chr(9) & "No. de títulos" & Chr(9) & "Val. SIVARMER" & Chr(9) & "Val. PIP" & Chr(9) & "Diferencia"
       Print #1, txtcadena
       For i = 1 To UBound(mata, 1)
           txtcadena = ""
           For j = 1 To UBound(mata, 2)
               txtcadena = txtcadena & mata(i, j) & Chr(9)
           Next j
           Print #1, txtcadena
       Next i
    End If
    Print #1, ""
    suma = ValidarMtMMercadoDinero(fecha, "TOTAL")
    Print #1, "Diferencia entre MtM port Mercado Dinero y Subportafolios: " & Chr(9) & suma
    suma = ValidarResVaR(fecha, "TOTAL", 500, 1)
    Print #1, "Diferencia entre CVaR port Banobras y Subportafolios, escenarios " & noesc & " horizonte " & 1 & ": " & Chr(9) & suma
    suma = ValidarResVaR(fecha, "NEGOCIACION + INVERSION", 500, 20)
    Print #1, "Diferencia entre CVaR port Banobras y Subportafolios, escenarios " & noesc & " horizonte " & 20 & ": " & Chr(9) & suma
    Close #1
    Screen.MousePointer = 0
    MsgBox "Fin de proceso"
 End If

End Sub

Private Sub mRiesgoF_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 MatFechasPos = ObFechasPos(txtportBanobras)
 frmAnalisisEvVaR.Show
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mSaldoDeriv_Click()
Dim tfecha As String
Dim fecha As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtport As String
Dim fechareg As Date
Dim cposicion As Integer
Dim coperacion As String
Dim fvaluacion As String
Dim monedaact As String
Dim monedapas As String
Dim saldoact As Double
Dim saldopas As Double
Dim mata() As Integer
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim no_ctp As Integer
Dim rmesa As New ADODB.recordset
Dim txtnomarch As String
Dim txtcadena As String
Dim siswap As Boolean
Dim indice As Integer
Dim txtcontrap As String
Dim sector As String
Dim toperacion As Integer
Dim montonoc As Double

tfecha = InputBox("Dame la fecha del proceso", , Date)
Screen.MousePointer = 11
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   mata = LeerContrapFecha(fecha)
   no_ctp = UBound(mata, 1)
   ReDim matsum(1 To no_ctp, 1 To 12) As Variant
   For i = 1 To no_ctp
       indice = BuscarValorArray(mata(i), MatContrapartes, 1)
       If indice <> 0 Then
          txtcontrap = MatContrapartes(indice, 3)
          sector = MatContrapartes(indice, 6)
       End If
       matsum(i, 1) = txtcontrap
       matsum(i, 2) = sector
       txtport = "Deriv Contrap " & mata(i)
       txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          For j = 1 To noreg
              fechareg = rmesa.Fields("FECHAREG")
              cposicion = rmesa.Fields("CPOSICION")
              coperacion = rmesa.Fields("COPERACION")
              monedaact = ""
              monedapas = ""
              saldoact = 0
              saldopas = 0
              siswap = DeterminaSiEsSwap(fechareg, coperacion)
              If siswap Then
                 Call ObtenerParamDeriv(cposicion, coperacion, fechareg, fvaluacion)
                 For l = 1 To UBound(MatTValSwaps1, 1)
                     If MatTValSwaps1(l, 1) = fvaluacion Then
                        monedaact = ReemplazaVacioValor(MatTValSwaps1(l, 13), "")
                        monedapas = ReemplazaVacioValor(MatTValSwaps1(l, 14), "")
                     End If
                 Next l
                 saldoact = ObtFlujoSwapFecha(coperacion, fechareg, fecha, "B")
                 saldopas = ObtFlujoSwapFecha(coperacion, fechareg, fecha, "C")
              Else
                 'Call ObtenerParamFwd(cposicion, coperacion, fechareg, fvaluacion, toperacion, montonoc)
                 'If toperacion = 1 Then
                 '   saldoact = montonoc
                 '   saldopas = 0
                 '   If fvaluacion = "FWD MXN/USD" Then
                 '      monedaact = "DOLAR PIP FIX"
                 '   Else
                 '      monedaact = "EURO BM D"
                 '   End If
                 '   monedapas = ""
                 'Else
                 '   saldoact = 0
                 '   saldopas = montonoc
                 '   If fvaluacion = "FWD MXN/USD" Then
                 '      monedapas = "DOLAR PIP FIX"
                 '   Else
                 '      monedapas = "EURO BM D"
                 '   End If
                 '   monedaact = ""
                 'End If
              End If
              If monedaact = "" Then
                matsum(i, 3) = matsum(i, 3) + saldoact
              ElseIf monedaact = "DOLAR PIP FIX" Then
                matsum(i, 5) = matsum(i, 5) + saldoact
              ElseIf monedaact = "YEN BM D" Then
                matsum(i, 7) = matsum(i, 7) + saldoact
              ElseIf monedaact = "UDI" Then
                matsum(i, 9) = matsum(i, 9) + saldoact
              ElseIf monedaact = "EURO BM D" Then
                matsum(i, 11) = matsum(i, 11) + saldoact
              End If
              If monedapas = "" Then
                matsum(i, 4) = matsum(i, 4) + saldopas
              ElseIf monedapas = "DOLAR PIP FIX" Then
                matsum(i, 6) = matsum(i, 6) + saldopas
              ElseIf monedapas = "YEN BM D" Then
                matsum(i, 8) = matsum(i, 8) + saldopas
              ElseIf monedapas = "UDI" Then
                matsum(i, 10) = matsum(i, 10) + saldopas
              ElseIf monedapas = "EURO BM D" Then
                matsum(i, 12) = matsum(i, 12) + saldopas
              End If
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
   Next i
   txtnomarch = "D:\salida" & Format(fecha, "yyyy-mm-dd") & ".txt"
   Open txtnomarch For Output As #1
   Print #1, "Contraparte" & Chr(9) & "Sector" & Chr(9) & "Pesos activa" & Chr(9) & "Pesos pasiva" & Chr(9) & "Dolares activa" & Chr(9) & "Dolares pasiva" & Chr(9) & "Yenes activa" & Chr(9) & "Yenes pasiva" & Chr(9) & "Udis activa" & Chr(9) & "Udis pasiva" & Chr(9) & "Euros activa" & Chr(9) & "Euros pasiva"
   For i = 1 To no_ctp
       txtcadena = ""
       For j = 1 To 12
           txtcadena = txtcadena & matsum(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Sub ObtenerParamDeriv(ByVal cposicion As Integer, ByVal coperacion As String, ByVal fechareg As Date, ByRef fvaluacion As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer

  txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 1 AND FECHAREG = " & txtfecha
  txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion & " AND COPERACION ='" & coperacion & "'"
  txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
  rmesa.Open txtfiltro1, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
  If noreg > 0 Then
     rmesa.Open txtfiltro2, ConAdo
     fvaluacion = rmesa.Fields("FVALUACION")
     rmesa.Close
  Else
    fvaluacion = ""
  End If

End Sub

Sub ObtenerParamFwd(ByVal cposicion As Integer, ByVal coperacion As String, ByVal fechareg As Date, ByRef fvaluacion As String, ByRef toperacion As Integer, ByRef montonoc As Double)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer

  txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE TIPOPOS = 1 AND FECHAREG = " & txtfecha
  txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion & " AND COPERACION ='" & coperacion & "'"
  txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
  rmesa.Open txtfiltro1, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
  If noreg > 0 Then
     rmesa.Open txtfiltro2, ConAdo
     fvaluacion = rmesa.Fields("CPRODUCTO")
     toperacion = rmesa.Fields("TOPERACION")
     montonoc = rmesa.Fields("M_NOCIONAL")
     rmesa.Close
  Else
    fvaluacion = ""
  End If

End Sub


Private Sub msql_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim rmesa As New ADODB.recordset
Dim txtnomarch As String

txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaSQLPort
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   txtnomarch = "d:\var_tc_sql_port.txt"
   frmReportes.CommonDialog1.FileName = txtnomarch
   frmReportes.CommonDialog1.ShowSave
   txtnomarch = frmReportes.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   For i = 1 To noreg
       txtcadena = rmesa.Fields(0) & Chr(9) & rmesa.Fields(1) & Chr(9) & rmesa.Fields(2)
       'For j = 1 To Len(txtcadena)
        '   MsgBox j & " " & Mid(txtcadena, j, 1) & " " & Asc(Mid(txtcadena, j, 1))
        '   MsgBox Mid(txtcadena, j, 1)
       'N'ext j
       txtcadena = ReemplazaCadenaTexto(txtcadena, Chr(10), " ")
       txtcadena = ReemplazaCadenaTexto(txtcadena, Chr(11), " ")
       txtcadena = ReemplazaCadenaTexto(txtcadena, Chr(12), " ")
       txtcadena = ReemplazaCadenaTexto(txtcadena, Chr(13), " ")
       Print #1, txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
   Close #1
   MsgBox "Fin de proceso"
End If

End Sub

Private Sub mvalDerivTC_Click()
'la funcion de este proceso es obtener la valuacion de la posicion
'de derivados para una fecha determinada
'para realizar este proceso primero lee
'clave de operacion
'tipo de operacion
'valuacion activa
'valuacion pasiva
'de la tabla VAR_TD_VAL_POS de la posicion de derivados
'luego consulta las tablas de posicion VAR_POS_SWAPS_3 y VAR_POS_FWDS_6 para obtener
'la clave de valuacion a aplicar a cada operacion
'despues consulta la informacion de las tablas catalogo VAR_TC_VAL_SWAPS2 y VAR_TC_VAL_FWDS2
'para determinar:
'si es swap, el tipo de tasa en cada pata y la moneda de cada pata
'si es forward de tipo de cambio, el sentido de la operacion y la moneda a aplicar

'datos de entrada: la fecha de proceso
'condiciones para la fecha: fecha laborable en mexico
'salida: un archivo de texto con el resultado
'cuando se ejecuta: a peticion de usuario
'utilidad de la salida: Para su uso final en los reportes de riesgo.

Dim tfecha As String
Dim fecha As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim noreg1 As Long
Dim mata() As Variant
Dim indice As Long
Dim indice1 As Long
Dim txtcadena As String
Dim txtnomarch As String

tfecha = InputBox("Dame la fecha a procesar", , Date)
If IsDate(tfecha) Then
fecha = CDate(tfecha)
Screen.MousePointer = 11
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT COPERACION,T_OPERACION,VAL_ACT_S,VAL_PAS_S FROM " & TablaValPos
txtfiltro2 = txtfiltro2 & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1  AND CPOSICION = 4 AND ID_VALUACION = 2 ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
    rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 15) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("COPERACION")
       mata(i, 2) = rmesa.Fields("T_OPERACION")
       mata(i, 3) = rmesa.Fields("VAL_ACT_S")
       mata(i, 4) = rmesa.Fields("VAL_PAS_S")
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtfiltro2 = "SELECT COPERACION,FVALUACION FROM " & TablaPosSwaps & " "
   txtfiltro2 = txtfiltro2 & " WHERE (FECHAREG,CPOSICION,COPERACION) IN(SELECT FECHAREG,CPOSICION,COPERACION "
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='SWAPS') ORDER BY COPERACION"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg1 = rmesa.Fields(0)
   rmesa.Close
   If noreg1 <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      ReDim matb(1 To noreg1, 1 To 2) As Variant
      For i = 1 To noreg1
          matb(i, 1) = rmesa.Fields("COPERACION")
          matb(i, 2) = rmesa.Fields("FVALUACION")
          rmesa.MoveNext
      Next i
      rmesa.Close
      For i = 1 To noreg
          indice1 = BuscarValorArray(mata(i, 1), matb, 1)
          If indice1 <> 0 Then
             mata(i, 5) = matb(indice1, 2)
          End If
      Next i
   End If
   txtfiltro2 = "SELECT COPERACION,CPRODUCTO FROM " & TablaPosFwd & " "
   txtfiltro2 = txtfiltro2 & " WHERE (FECHAREG,CPOSICION,COPERACION) IN(SELECT FECHAREG,CPOSICION,COPERACION "
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='FORWARDS') ORDER BY COPERACION"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg1 = rmesa.Fields(0)
   rmesa.Close
   If noreg1 <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      ReDim matb(1 To noreg1, 1 To 2) As Variant
      For i = 1 To noreg1
          matb(i, 1) = rmesa.Fields("COPERACION")
          matb(i, 2) = rmesa.Fields("CPRODUCTO")
          rmesa.MoveNext
      Next i
      rmesa.Close
      For i = 1 To noreg
          indice1 = BuscarValorArray(mata(i, 1), matb, 1)
          If indice1 <> 0 Then
             mata(i, 5) = matb(indice1, 2)
          End If
      Next i
   End If
      For i = 1 To noreg
          For j = 1 To UBound(MatTValSwaps2, 1)
              If mata(i, 5) = MatTValSwaps2(j, 1) Then
                 mata(i, 6) = ReemplazaVacioValor(MatTValSwaps2(j, 5), "")      'CURVA DESC ACTIVA
                 mata(i, 7) = ReemplazaVacioValor(MatTValSwaps2(j, 9), "")      'CURVA PAGO ACTIVA
                 mata(i, 8) = ReemplazaVacioValor(MatTValSwaps2(j, 7), "")      'CURVA DESC PASIVA
                 mata(i, 9) = ReemplazaVacioValor(MatTValSwaps2(j, 11), "")     'CURVA PAGO PASIVA
                 mata(i, 10) = ReemplazaVacioValor(MatTValSwaps2(j, 13), "")    'MONEDA ACTIVA
                 mata(i, 11) = ReemplazaVacioValor(MatTValSwaps2(j, 14), "")    'MONEDA PASIVA
                 mata(i, 12) = DetTTasa(mata(i, 6), mata(i, 7))
                 mata(i, 13) = DetTTasa(mata(i, 8), mata(i, 9))
                 Exit For
              End If
          Next j
      Next i
      For i = 1 To noreg
          For j = 1 To UBound(MatTValFwdsTC2, 1)
              If mata(i, 5) = MatTValFwdsTC2(j, 1) Then
                 mata(i, 14) = ReemplazaVacioValor(MatTValFwdsTC2(j, 11), "")       'MONEDA
                 Exit For
              End If
          Next j
      Next i
       
      ReDim matres(1 To 20, 1 To 6) As Variant
      matres(1, 1) = "Tasa Fija MXN Activa"
      matres(2, 1) = "Tasa Fija MXM Pasiva"
      matres(3, 1) = "Tasa Variable MXN Activa"
      matres(4, 1) = "Tasa Variable MXN Pasiva"
      matres(5, 1) = "TASA FIJA UDIS ACTIVA"
      matres(6, 1) = "TASA FIJA UDIS PASIVA"
      matres(7, 1) = "TASA VARIABLE UDIS ACTIVA"
      matres(8, 1) = "TASA VARIABLE UDIS PASIVA"
      matres(9, 1) = "TASA FIJA DOLARES ACTIVA"
      matres(10, 1) = "TASA FIJA DOLARES PASIVA"
      matres(11, 1) = "TASA VARIABLE DOLARES ACTIVA"
      matres(12, 1) = "TASA VARIABLE DOLARES PASIVA"
      matres(13, 1) = "TASA FIJA YENES ACTIVA"
      matres(14, 1) = "TASA FIJA YENES PASIVA"
      matres(15, 1) = "TASA VARIABLE YENES ACTIVA"
      matres(16, 1) = "TASA VARIABLE YENES PASIVA"
      matres(17, 1) = "Forwards USD Activa"
      matres(18, 1) = "Forwards USD Pasivo"
      matres(19, 1) = "Forwards EUR Activa"
      matres(20, 1) = "Forwards EUR Pasivo"
      For i = 1 To noreg
      'tasa fija mn activa
      If mata(i, 12) = "TF" And mata(i, 10) = "" Then
         matres(1, 2) = matres(1, 2) + mata(i, 3)
      End If
      'tasa fija mn pasiva
      If mata(i, 13) = "TF" And mata(i, 11) = "" Then
         matres(2, 2) = matres(2, 2) + mata(i, 4)
      End If
      'tasa variable mn activa
      If mata(i, 12) = "TV" And mata(i, 10) = "" Then
         matres(3, 2) = matres(3, 2) + mata(i, 3)
      End If
      'tasa variable mn pasiva
      If mata(i, 13) = "TV" And mata(i, 11) = "" Then
         matres(4, 2) = matres(4, 2) + mata(i, 4)
      End If
      'tasa fija udis activa
      If mata(i, 12) = "TF" And mata(i, 10) = "UDI" Then
         matres(5, 2) = matres(5, 2) + mata(i, 3)
      End If
      'tasa fija udis pasiva
      If mata(i, 13) = "TF" And mata(i, 11) = "UDI" Then
         matres(6, 2) = matres(6, 2) + mata(i, 4)
      End If
      'tasa variable udis activa
      If mata(i, 12) = "TV" And mata(i, 10) = "UDI" Then
         matres(7, 2) = matres(7, 2) + mata(i, 3)
      End If
      'tasa variable udis pasiva
      If mata(i, 13) = "TV" And mata(i, 11) = "UDI" Then
         matres(8, 2) = matres(8, 2) + mata(i, 4)
      End If
      'tasa fija dolares activa
      If mata(i, 12) = "TF" And mata(i, 10) = "DOLAR PIP FIX" Then
         matres(9, 2) = matres(9, 2) + mata(i, 3)
      End If
      'tasa fija dolares pasiva
      If mata(i, 13) = "TF" And mata(i, 11) = "DOLAR PIP FIX" Then
         matres(10, 2) = matres(10, 2) + mata(i, 4)
      End If
      'tasa variable dolares activa
      If mata(i, 12) = "TV" And mata(i, 10) = "DOLAR PIP FIX" Then
         matres(11, 2) = matres(11, 2) + mata(i, 3)
      End If
      'tasa variable dolares pasiva
      If mata(i, 13) = "TV" And mata(i, 11) = "DOLAR PIP FIX" Then
         matres(12, 2) = matres(12, 2) + mata(i, 4)
      End If
      'tasa fija yenes activa
      If mata(i, 12) = "TF" And mata(i, 10) = "YEN BM" Then
         matres(13, 2) = matres(13, 2) + mata(i, 3)
      End If
      'tasa fija yenes pasiva
      If mata(i, 13) = "TF" And mata(i, 11) = "YEN BM" Then
         matres(14, 2) = matres(14, 2) + mata(i, 4)
      End If
      'tasa variable yenes activa
      If mata(i, 12) = "TV" And mata(i, 10) = "YEN BM" Then
         matres(15, 2) = matres(15, 2) + mata(i, 3)
      End If
      'tasa variable yenes pasiva
      If mata(i, 13) = "TV" And mata(i, 11) = "YEN BM" Then
         matres(16, 2) = matres(16, 2) + mata(i, 4)
      End If
      'forwards usd activa
      If mata(i, 2) = 1 And mata(i, 14) = "DOLAR PIP FIX" Then
         matres(17, 2) = matres(17, 2) + mata(i, 3) - mata(i, 4)
      End If
      'forwards usd pasiva
      If mata(i, 2) = -1 And mata(i, 14) = "DOLAR PIP FIX" Then
         matres(18, 2) = matres(18, 2) + mata(i, 4) - mata(i, 3)
      End If
      'forwards euro activa
      If mata(i, 2) = 1 And mata(i, 14) = "EURO BM" Then
         matres(19, 2) = matres(19, 2) + mata(i, 3) - mata(i, 4)
      End If
      'forwards euro pasiva
      If mata(i, 2) = -1 And mata(i, 14) = "EURO BM" Then
         matres(20, 2) = matres(20, 2) + mata(i, 4) - mata(i, 3)
      End If
    
      Next i
      txtnomarch = DirResVaR & "\Val deriv x moneda " & Format(fecha, "yyyy-mm-dd") & ".txt"
      frmReportes.CommonDialog1.FileName = txtnomarch
      frmReportes.CommonDialog1.ShowSave
      txtnomarch = frmReportes.CommonDialog1.FileName
      Open txtnomarch For Output As #1
      For i = 1 To 20
      txtcadena = ""
      For j = 1 To 4
      txtcadena = txtcadena & matres(i, j) & Chr(9)
      Next j
      Print #1, txtcadena
      Next i
      Close #1
   
End If
Screen.MousePointer = 0
MsgBox "Fin del proceso"
End If
End Sub

Function DetTTasa(ByVal txtcdesc As String, ByVal txtcpago As String)

If Not EsVariableVacia(txtcdesc) And Not EsVariableVacia(txtcpago) Then
   DetTTasa = "TV"
ElseIf Not EsVariableVacia(txtcdesc) And EsVariableVacia(txtcpago) Then
   DetTTasa = "TF"
End If
End Function
