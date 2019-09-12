VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalVar 
   Caption         =   "Sistema de Cálculo VAR Banco Nacional Obras y Servicios Publicos"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   -2040
   ClientWidth     =   17565
   Icon            =   "frmcalvar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   17565
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Catalogos"
      Height          =   1215
      Left            =   300
      TabIndex        =   6
      Top             =   7500
      Width           =   4305
      Begin VB.CommandButton Command2 
         Caption         =   "Carga manual de catalogos en memoria"
         Height          =   600
         Left            =   300
         TabIndex        =   7
         Top             =   390
         Width           =   1500
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   9240
      Width           =   17565
      _ExtentX        =   30983
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   25321
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame31 
      Caption         =   "Derivados"
      Height          =   1215
      Left            =   300
      TabIndex        =   0
      Top             =   6000
      Width           =   4305
      Begin VB.CommandButton Command38 
         Caption         =   "Validacion de operaciones IKOS"
         Height          =   600
         Left            =   300
         TabIndex        =   1
         Top             =   390
         Width           =   1500
      End
   End
   Begin VB.PictureBox Inet1 
      Height          =   480
      Left            =   -408
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   -408
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   65535
      Left            =   3030
      Top             =   30
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      Picture         =   "frmcalvar.frx":0CCA
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   34799617
      CurrentDate     =   41736
   End
   Begin VB.Timer Timer2 
      Interval        =   50000
      Left            =   3480
      Top             =   30
   End
   Begin VB.Menu mprincipal 
      Caption         =   "Menu principal"
      Begin VB.Menu mListaP 
         Caption         =   "Generar lista de procesos"
      End
      Begin VB.Menu mImportardatos 
         Caption         =   "Ejecución de procesos"
      End
      Begin VB.Menu mSubprocesos 
         Caption         =   "Lista de subprocesos"
      End
      Begin VB.Menu mReportes 
         Caption         =   "Reportes"
      End
      Begin VB.Menu mSimuladas 
         Caption         =   "Importación de posiciones simuladas"
      End
      Begin VB.Menu mParametros 
         Caption         =   "Parámetros del sistema"
      End
      Begin VB.Menu mCatalogos 
         Caption         =   "Catalogos del sistema"
         Begin VB.Menu mEdCatalogos 
            Caption         =   "Actualización de Catalogos"
         End
         Begin VB.Menu mActSQLPort 
            Caption         =   "Actualizacion de Cadenas SQL Portafolios"
         End
      End
      Begin VB.Menu mBitacora 
         Caption         =   "Bitacora de Operación"
      End
      Begin VB.Menu mSalir 
         Caption         =   "Salir del sistema"
      End
   End
   Begin VB.Menu mmodulos 
      Caption         =   "Procesos reporte diario"
      Begin VB.Menu CalcCVaRt_2 
         Caption         =   "Calculo de CVaR Balanza t-2"
      End
      Begin VB.Menu mBack 
         Caption         =   "Historico del Backtesting"
         Visible         =   0   'False
      End
      Begin VB.Menu mValSwap 
         Caption         =   "Valuacion de swap"
      End
      Begin VB.Menu mValPosPrim 
         Caption         =   "Valuacion de posicion primaria"
      End
      Begin VB.Menu mReproceso 
         Caption         =   "Reproceso de operaciones"
      End
      Begin VB.Menu mActValIKOS 
         Caption         =   "Actualizacion de valuaciòn IKOS"
      End
      Begin VB.Menu mMarcoOp 
         Caption         =   "Marco de operacion"
      End
      Begin VB.Menu mMonitorVO 
         Caption         =   "Monitor de validación de operaciones"
      End
   End
   Begin VB.Menu mProcEfect 
      Caption         =   "Procesos efectividad"
      Begin VB.Menu mCargapp 
         Caption         =   "Carga de posiciones primarias"
      End
      Begin VB.Menu mefretro2 
         Caption         =   "Eficiencia retrospectiva portafolio"
      End
      Begin VB.Menu mPPrimaria 
         Caption         =   "Posiciones Primarias"
      End
      Begin VB.Menu mEficiencia2 
         Caption         =   "Eficiencia prospectiva de 1 swap"
      End
      Begin VB.Menu mEficProsFwd 
         Caption         =   "Eficiencia Prospectiva fwd"
      End
   End
   Begin VB.Menu mProcesos2 
      Caption         =   "Procesos Manuales"
      Begin VB.Menu mDALM 
         Caption         =   "Obtener datos ALM"
         Visible         =   0   'False
      End
      Begin VB.Menu mSimCVaR 
         Caption         =   "Simulaciones de CVaR"
      End
      Begin VB.Menu mGenAlMont 
         Caption         =   "Generar numeros aleatorios Montecarlo"
      End
      Begin VB.Menu mAnalisisBC 
         Caption         =   "Analisis de la Balanza Cambiaria"
      End
      Begin VB.Menu mprocmanual 
         Caption         =   "Procesos manuales"
      End
      Begin VB.Menu mCVaRFP 
         Caption         =   "CVaR Fondo de Pensiones"
      End
   End
   Begin VB.Menu mProcFR 
      Caption         =   "Procesos Factores de riesgo"
      Begin VB.Menu mCargaFR 
         Caption         =   "Importacion de factores de riesgo"
      End
      Begin VB.Menu mExtraTIIE 
         Caption         =   "Extrapolar yields referenciadas a TIIE"
      End
      Begin VB.Menu mGenDatosEm 
         Caption         =   "Generar datos de emision"
      End
      Begin VB.Menu mCrearFR 
         Caption         =   "Crear nuevo factor de riesgo"
      End
      Begin VB.Menu mGenSimFactR 
         Caption         =   "Generar simulaciones de factores de riesgo"
      End
   End
   Begin VB.Menu mProcEm 
      Caption         =   "Procesos de emisiones"
      Begin VB.Menu mCargarFEE 
         Caption         =   "Cargar flujos de emisiones especiales"
      End
      Begin VB.Menu mFlujosPIP 
         Caption         =   "Obtener flujos de archivos PIP"
      End
   End
   Begin VB.Menu mProcInfMen 
      Caption         =   "Procesos Informes mensuales"
      Begin VB.Menu CalcCVaRt_9 
         Caption         =   "Calculo de CVaR Balanza t-9"
      End
      Begin VB.Menu mVReemplazo 
         Caption         =   "Valor de reemplazo"
      End
      Begin VB.Menu mCCVA 
         Caption         =   "Calculo de CVA"
      End
      Begin VB.Menu mAVMCVAR 
         Caption         =   "Analisis de validez del modelo de CVaR"
      End
      Begin VB.Menu LecMinusPlus 
         Caption         =   "Lectura de Plus-Minusvalias"
      End
      Begin VB.Menu mCalcMW 
         Caption         =   "Calculo de Make Whole"
      End
      Begin VB.Menu mCLimC 
         Caption         =   "Calculo de limite contraparte"
      End
   End
   Begin VB.Menu mmodulos2 
      Caption         =   "Procesos para otras areas"
      Begin VB.Menu mexport1 
         Caption         =   "Exportacion de val deriv a VAR_VALDERIV1"
      End
      Begin VB.Menu mactswr3r 
         Caption         =   "Actualizar SW_RIESGO3_RESPALDO"
      End
      Begin VB.Menu mexpVFD 
         Caption         =   "Exportar flujos a VAR_FLUJOSD"
      End
   End
   Begin VB.Menu mUtilerias 
      Caption         =   "Utilerias"
      Begin VB.Menu msistema1 
         Caption         =   "Mensajes de sistema"
      End
      Begin VB.Menu mbdt 
         Caption         =   "Modelo BDT"
         Visible         =   0   'False
      End
      Begin VB.Menu mABalanza 
         Caption         =   "Analisis posicion divisas"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mAcerca 
         Caption         =   "Acerca del sistema"
      End
   End
End
Attribute VB_Name = "frmCalVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatTotales() As Variant
Dim Fechasback() As Date
Dim NoDiasBack As Long
Dim VarFecha As Integer

Sub VerTablasVaRMontecarlo(ByRef matp() As Variant, ByRef matrends() As Double, ByRef matcov() As Double, ByRef Matroot() As Double, ByRef obj As MSFlexGrid)
Dim noreg As Long
Dim i As Long
Dim j As Long

noreg = UBound(matcov, 1)
obj.Rows = noreg + 1
obj.Cols = noreg + 1
For i = 1 To noreg
    For j = 1 To noreg
        obj.TextMatrix(i, j) = 9
    Next j
Next i
End Sub

Sub GenMArcoOp(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim suma1 As Double
Dim suma2 As Double
Dim suma3 As Double
Dim suma4 As Double
Dim suma5 As Double
Dim suma6 As Double
Dim suma7 As Double
Dim suma8 As Double
Dim suma9 As Double
Dim suma10 As Double
Dim matposmd() As New propPosMD
Dim variable1 As Double
Dim variable2 As Double
Dim variable3 As Double
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

   txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
   txtfiltro = "select count(*) from " & TablaVecPrecios & " WHERE FECHA = " & txtfecha
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      txtfiltro1 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND (TIPOPOS = 1 OR TIPOPOS = 4) AND (CPOSICION =1 OR CPOSICION =2)"
      matposmd = LeerBaseMD(txtfiltro1)
      MatVPreciosT = LeerPVPrecios(fecha)
      Call CompletarPosMesaD(matposmd, MatVPreciosT)
      suma1 = 0
      suma2 = 0
      suma3 = 0
      suma4 = 0
      suma5 = 0
      suma6 = 0
      suma7 = 0
      suma8 = 0
      suma9 = 0
      suma10 = 0
      For i = 1 To UBound(matposmd, 1)
         If matposmd(i).Tipo_Mov = "1" Then matposmd(i).Signo_Op = 1
         If matposmd(i).Tipo_Mov = "4" Then matposmd(i).Signo_Op = -1
         'tasa fija
         If (matposmd(i).reglaCuponMD = "Tasa Fija" Or matposmd(i).tValorMD = "I" Or matposmd(i).tValorMD = "BI") And (matposmd(i).Tipo_Mov = "1" Or matposmd(i).Tipo_Mov = "4") Then
            suma1 = suma1 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'en directo
         If (matposmd(i).Tipo_Mov = "1" Or matposmd(i).Tipo_Mov = "4") Then
            suma2 = suma2 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'cetes
         If (matposmd(i).tValorMD = "BI") And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma3 = suma3 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'bonos y udibonos
         If (matposmd(i).tValorMD = "S" Or matposmd(i).tValorMD = "M") And (matposmd(i).Tipo_Mov = "1" Or matposmd(i).Tipo_Mov = "4") Then
            suma4 = suma4 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
        'banca de desarrollo
         variable1 = matposmd(i).emisionMD = "NAFIN" Or matposmd(i).emisionMD = "BACMEXT" Or matposmd(i).emisionMD = "SHF" Or matposmd(i).emisionMD = "BANSEFI" Or matposmd(i).emisionMD = "NAFF"
         'certificado bursatil
         variable2 = matposmd(i).tValorMD = "90" Or matposmd(i).tValorMD = "91" Or matposmd(i).tValorMD = "92" Or matposmd(i).tValorMD = "93" Or matposmd(i).tValorMD = "94" Or matposmd(i).tValorMD = "95" Or matposmd(i).tValorMD = "96" Or matposmd(i).tValorMD = "97" Or matposmd(i).tValorMD = "98" Or matposmd(i).tValorMD = "F" Or matposmd(i).tValorMD = "CD" Or matposmd(i).tValorMD = "CF"
         'paraestatales
         variable3 = matposmd(i).emisionMD = "CFE" Or matposmd(i).emisionMD = "PEMEX"
         'prlv de banca de desarrollo
         If (variable1 And matposmd(i).tValorMD = "I") And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma5 = suma5 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'certificados  de banca de desarrollo
         If (variable1 And variable2) And matposmd(i).reglaCuponMD = "Tasa Fija" And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma6 = suma6 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'prlvs de bancos privados
         If (matposmd(i).tValorMD = "I" And Not variable1 And Not variable3) And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma7 = suma7 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'certificados  de bancos privados
         If (variable2 And Not variable1 And Not variable3) And matposmd(i).reglaCuponMD = "Tasa Fija" And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma8 = suma8 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         'paraestatales menor a un año
         If variable3 And matposmd(i).fVencMD - fecha < 365 And matposmd(i).reglaCuponMD = "Tasa Fija" And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma9 = suma9 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
         If variable3 And matposmd(i).fVencMD - fecha >= 365 And matposmd(i).reglaCuponMD = "Tasa Fija" And (matposmd(i).Tipo_Mov = 1 Or matposmd(i).Tipo_Mov = 4) Then
            suma10 = suma10 + matposmd(i).noTitulosMD * matposmd(i).valSucioPIP * matposmd(i).Signo_Op
         End If
      Next i

         txtcadena = fecha & Chr(9) & suma1 & Chr(9) & suma2 & Chr(9)
         txtcadena = txtcadena & suma3 & Chr(9) & suma4 & Chr(9)
         txtcadena = txtcadena & suma5 & Chr(9) & suma6 & Chr(9)
         txtcadena = txtcadena & suma7 & Chr(9) & suma8 & Chr(9)
         txtcadena = txtcadena & suma9 & Chr(9) & suma10
        
         Print #1, txtcadena
         
   End If
End Sub

Function TraduceClaveSwapDeuda(ByVal txtent As String, ByVal id_toper As Integer)
'objetivo de esta funcion: asignar una clave de producto a una posicion de deuda
'que tiene asociada una posición derivada, actualmente las claves estan definidas
'dentro del codigo, lo mas recomentable seria asociar las claves con un catalogo

'variables
'txtent   -   es la clave de tipo de swap que usa el sistema ikos derivados
'id_toper -   es el tipo de operacion que se introdujo 1-activa, 4-pasiva

Dim txtclave As String
Dim txtclave1 As String
Dim i As Long

    txtclave = TraduceDerivadoEstandar2(txtent)
    If Not EsVariableVacia(txtclave) Then
       For i = 1 To UBound(MatRelSwapIS, 1)
           If txtclave = MatRelSwapIS(i, 1) And id_toper = 1 Then
              txtclave1 = MatRelSwapIS(i, 7)
           ElseIf txtclave = MatRelSwapIS(i, 1) And id_toper = 4 Then
              txtclave1 = MatRelSwapIS(i, 6)
           End If
       Next i
     Else
       MsgBox "No se ha definido " & txtclave
       txtclave1 = ""
     End If
 TraduceClaveSwapDeuda = txtclave1
End Function

Sub GenProcBack(ByVal dtfecha As Date, ByVal id_proc As Integer, ByVal txtport As String, ByVal id_tabla As Integer)
    Dim noport As Integer
    Dim bl_exito As Boolean, bl_exito1 As Boolean
    Dim i As Integer, k As Integer, j As Integer
    Dim jj As Integer, indice As Integer
    Dim p As Integer
    Dim contar As Long
    Dim dtfechaf, dtfechav As Date
    Dim txtnomarch As String, txtsalida As String, txtfecha1 As String, txthorap As String, txtcadena As String
    Dim txtfecha2 As String
    Dim txtfiltro As String
    Dim txttabla As String
    
    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_proc, contar, "Cálculo de Backtesting", dtfecha, dtfecha, dtfecha, txtport, "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub


Private Sub CalcCVaRt_2_Click()
Dim tfecha As String
Dim fecha As Date
Dim fecha0 As Date
Dim txtport As String
Dim txtportfr As String
Dim txtmsg As String
Dim noesc As Integer
Dim htiempo As Integer
Dim nconf As Double
Dim exito As Boolean
Dim valor As Double
Dim valpos As Double
Dim vallim As Double
Dim txtcadena As String
Dim txtnomarch As String

txtport = "Balanza t-2"
txtportfr = "Normal"
noesc = 500
htiempo = 1
nconf = 0.97

tfecha = InputBox("Dame la fecha de calculo ", , Date)
If IsDate(tfecha) Then
   frmProgreso.Show
   SiActTProc = True
   fecha = CDate(tfecha)
   fecha0 = PBD1(fecha, 1, "MX")
   Call CrearPosBalanza(fecha0, 10, "posxdiv50115041", txtport)
   FechaPosRiesgo = 0
   txtNomPosRiesgo = ""
   Call SubprocCalculoPyGPort(fecha0, fecha, fecha, txtport, txtportfr, noesc, htiempo, txtmsg, exito)
   Call ConsolidaPyGSubport(fecha0, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, txtmsg, exito)
   valor = CalcularCVaRPyG(fecha0, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, 1 - nconf, exito)
   valpos = ValuarPosBalUSD(fecha0, fecha, 10)
   CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
   'fechacn = DevFechaLimite(CDate(fecha), MatCapitalSist, "CAPITAL NETO B")
    vallim = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MC")
    txtcadena = "La revisión de la posición cambiaria tomada de la balanza contable para el " & Day(fecha0)
    txtcadena = txtcadena & " de " & Mestxt(Month(fecha0)) & " de " & Year(fecha0) & " da como resultado una posición dolarizada de $"
    txtcadena = txtcadena & Format(valpos / 1000000, "###,###,##0.00") & " mdd con un CVaR de $" & Format(Abs(valor) / 1000000, "##0.00") & " mdp, lo que implica un consumo del "
    txtcadena = txtcadena & Format(Abs(valor) / (CapitalNeto * vallim), "#,##0.00%")
    MsgBox txtcadena
    txtnomarch = DirResVaR & "\poscam " & Format(fecha, "yyyy-mm-dd") & ".txt"
    frmCalVar.CommonDialog1.FileName = txtnomarch
    frmCalVar.CommonDialog1.ShowSave
    txtnomarch = frmCalVar.CommonDialog1.FileName
    Open txtnomarch For Output As #1
    Print #1, txtcadena
    Close #1
    Unload frmProgreso
    SiActTProc = False
End If

MsgBox "Fin de proceso"
End Sub

Function ValuarPosBalUSD(ByVal f_pos As Date, ByVal f_val As Date, ByVal cposicion As Integer)
Dim mata() As Variant
Dim valusd As Double
Dim valyen As Double
Dim valeur As Double
Dim valor As Double
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

mata = Leer1FactorR(f_val, f_val, "DOLAR PIP FIX", 0)
valusd = mata(1, 2)
mata = Leer1FactorR(f_val, f_val, "YEN PIP", 0)
valyen = mata(1, 2)
mata = Leer1FactorR(f_val, f_val, "EURO PIP", 0)
valeur = mata(1, 2)

txtfecha = "TO_DATE('" & Format(f_pos, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion & " ORDER BY COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
rmesa.Open txtfiltro2, ConAdo
ReDim matb(1 To noreg, 1 To 2) As Double
For i = 1 To noreg
  If rmesa.Fields(7) = 1 Then
    matb(i, 1) = rmesa.Fields(7)
   Else
   matb(i, 1) = -1
   End If
    matb(i, 2) = rmesa.Fields(12)
rmesa.MoveNext
Next i
rmesa.Close
valor = matb(1, 1) * matb(1, 2)
valor = valor + matb(2, 1) * matb(2, 2) * valyen / valusd
valor = valor + matb(3, 1) * matb(3, 2) * valeur / valusd

ValuarPosBalUSD = valor

End Function

Sub CrearPosBalanza(ByVal fecha As Date, ByVal cposicion As Integer, ByVal txttabla As String, ByVal txtport As String)
Dim txtfecha As String
Dim nomarch As String
Dim noreg As Integer
Dim i As Integer
Dim txtcadena As String
Dim txtmsg As String
Dim exito1 As Boolean
Dim txtfiltro() As String
Dim matpos() As propPosRiesgo
Dim valdolares As Double
Dim valeuros As Double
Dim valyenes As Double
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
nomarch = DirReportes & "\" & NomArchRVaR
Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset("select * from [" & txttabla & "$] where FECHA = " & CLng(fecha), dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 ReDim mata(1 To 3, 1 To 5) As Variant
   If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg = registros1.RecordCount
      registros1.MoveFirst
      For i = 1 To noreg
          valdolares = registros1.Fields(1)
          valyenes = registros1.Fields(2)
          valeuros = registros1.Fields(3)
      Next i
   End If
   registros1.Close
   base1.Close

mata(1, 1) = FSigno(valdolares)
mata(1, 2) = "*C"
mata(1, 3) = "MXPUSD"
mata(1, 4) = "FIX"
mata(1, 5) = Abs(valdolares)

mata(2, 1) = FSigno(valyenes)
mata(2, 2) = "*C"
mata(2, 3) = "MXPJPY"
mata(2, 4) = "JPY"
mata(2, 5) = Abs(valyenes)

mata(3, 1) = FSigno(valeuros)
mata(3, 2) = "*C"
mata(3, 3) = "MXPEUR"
mata(3, 4) = "EUR"
mata(3, 5) = Abs(valeuros)



txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
ConAdo.Execute "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & cposicion
For i = 1 To 3
    txtcadena = "INSERT INTO " & TablaPosDiv & " VALUES("
    txtcadena = txtcadena & "1,"
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & "'Real',"
    txtcadena = txtcadena & "'000000',"
    txtcadena = txtcadena & "'N',"
    txtcadena = txtcadena & cposicion & ","
    txtcadena = txtcadena & i & ","
    txtcadena = txtcadena & mata(i, 1) & ","
    txtcadena = txtcadena & "'" & mata(i, 2) & "',"
    txtcadena = txtcadena & "'" & mata(i, 3) & "',"
    txtcadena = txtcadena & "'" & mata(i, 4) & "',"
    txtcadena = txtcadena & "'" & mata(i, 2) & mata(i, 3) & mata(i, 4) & "',"
    txtcadena = txtcadena & mata(i, 5) & ","
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & "0,"
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "null)"
    ConAdo.Execute txtcadena
Next i
ConAdo.Execute "DELETE FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
For i = 1 To 3
    txtcadena = "INSERT INTO " & TablaPortPosicion & " VALUES("
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & "'" & txtport & "',"
    txtcadena = txtcadena & "1,"
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & "'Real',"
    txtcadena = txtcadena & "'000000',"
    txtcadena = txtcadena & cposicion & ","
    txtcadena = txtcadena & i & ")"
    ConAdo.Execute txtcadena
Next i

End Sub


Private Sub CalcCVaRt_9_Click()
Dim tfecha As String
Dim fecha As Date
Dim fecha0 As Date
Dim txtport As String
Dim txtportfr As String
Dim cposicion As Integer
Dim txtnomarch As String
Dim txtcadena As String
Dim nconf As Double
Dim valor As Double
Dim valpos As Double
Dim vallim As Double
Dim htiempo As Integer
Dim noesc As Integer
Dim txtmsg As String
Dim exito As Boolean
cposicion = 11
txtport = "Balanza t-9"
txtportfr = "Normal"
noesc = 500
htiempo = 1
nconf = 0.97
tfecha = InputBox("Dame la fecha de calculo ", , Date)
If IsDate(tfecha) Then
   frmProgreso.Show
   fecha = CDate(tfecha)
   fecha0 = fecha
   Call CrearPosBalanza(fecha, cposicion, "posxdiv5011t9", txtport)
   FechaPosRiesgo = 0
   txtNomPosRiesgo = ""
   Call SubprocCalculoPyGPort(fecha0, fecha, fecha, txtport, txtportfr, noesc, htiempo, txtmsg, exito)
   Call ConsolidaPyGSubport(fecha0, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, txtmsg, exito)
   valor = CalcularCVaRPyG(fecha0, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, 1 - nconf, exito)
   valpos = ValuarPosBalUSD(fecha0, fecha, cposicion)
   CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
   vallim = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MC")
   txtcadena = "La revisión de la posición cambiaria tomada de la balanza contable para el " & Day(fecha0)
   txtcadena = txtcadena & " de " & Month(fecha0) & " de " & Year(fecha0) & " da como resultado una posición dolarizada de $"
   txtcadena = txtcadena & Format(valpos / 1000000, "###,###,##0.00") & " mdd con un CVaR de $" & Format(Abs(valor) / 1000000, "##0.00") & " mdp, lo que implica un consumo del "
   txtcadena = txtcadena & Format(Abs(valor) / (CapitalNeto * vallim), "##0.00%")
   MsgBox txtcadena
   txtnomarch = "d:\riesgo poscam" & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = txtnomarch
   frmCalVar.CommonDialog1.ShowSave
   txtnomarch = frmCalVar.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   Print #1, txtcadena
   Close #1
   Unload frmProgreso
End If
End Sub


Sub ActTasasCuponNuevo(ByVal fecha As Date)
Dim exito As Boolean
Dim mata() As Variant
Dim matf1() As Date
Dim matf2() As Date
Dim txtmsg As String
Dim i As Long
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txfecha As String
mata = Leer1FactorR(#1/1/2003#, fecha, "LIBOR 6M PIP", 0)
matf1 = LeerFSinTasa(fecha, "B", "LIBOR6M[2]")
matf2 = LeerFSinTasa(fecha, "C", "LIBOR6M[2]")
If UBound(matf1, 1) <> 0 Then
   For i = 1 To UBound(matf1, 1)
       Call ActTCFlujosSwaps3(matf1(i), mata, "B", "LIBOR6M[2]", txtmsg, exito)
   Next i
End If
If UBound(matf2, 1) <> 0 Then
   For i = 1 To UBound(matf2, 1)
       Call ActTCFlujosSwaps3(matf2(i), mata, "C", "LIBOR6M[2]", txtmsg, exito)
   Next i
End If
End Sub

Function LeerFSinTasa(ByVal fecha As Date, ByVal txtpos As String, ByVal txttref As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim matf() As Date
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT FINICIO FROM " & TablaFlujosSwapsO & " WHERE FINICIO IN ("
txtfiltro2 = txtfiltro2 & "(select FINICIO from " & TablaFlujosSwapsO & " where (coperacion,tpata,finicio) in "
txtfiltro2 = txtfiltro2 & "(select a.coperacion,a.tpata,a.finicio "
txtfiltro2 = txtfiltro2 & "from " & TablaFlujosSwapsO & " a join " & TablaPosSwaps & " b "
txtfiltro2 = txtfiltro2 & "on a.tipopos = b.tipopos and "
txtfiltro2 = txtfiltro2 & "a.cposicion = b.cposicion and "
txtfiltro2 = txtfiltro2 & "a.coperacion = b.coperacion "
txtfiltro2 = txtfiltro2 & "where a.tpata = '" & txtpos & "' and "
If txtpos = "B" Then
   txtfiltro2 = txtfiltro2 & "b.TC_ACTIVA = '" & txttref & "' "
Else
   txtfiltro2 = txtfiltro2 & "b.TC_PASIVA = '" & txttref & "' "
End If
txtfiltro2 = txtfiltro2 & "and a.tasa =0 "
txtfiltro2 = txtfiltro2 & " AND a.FINICIO <= " & txtfecha & "))) GROUP BY FINICIO ORDER BY FINICIO"
txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matf(1 To noreg) As Date
   For i = 1 To noreg
       matf(i) = rmesa.Fields("FINICIO")
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim matf(0 To 0) As Date
End If
LeerFSinTasa = matf
End Function


Function algo(ByRef mat1(), ByRef mat2())

End Function






Function FiltrarSwapPos(ByRef matpos() As propPosRiesgo)
Dim contar As Integer
Dim i As Integer
Dim j As Integer
Dim noreg As Integer

'esta rutina filtra la posicion por la clave de la posicion o mesa
ReDim mata(1 To 1) As New propPosRiesgo
noreg = UBound(matpos, 1)
contar = 0
For i = 1 To noreg
If Val(matpos(i).C_Posicion) = ClavePosDeriv And matpos(i, CSiPosPrim) = "N" Then
contar = contar + 1
ReDim Preserve mata(1 To contar) As New propPosRiesgo
mata(contar).c_operacion = matpos(i).c_operacion
End If
Next i
If contar = 0 Then
 ReDim matb(0 To 0, 0 To 0) As Variant
 FiltrarSwapPos = matb
 MensajeProc = "no existe posicion para la fecha de analisis"
Else
 FiltrarSwapPos = mata
End If
End Function


Function FiltrarPosClave(ByRef matpos() As propPosRiesgo, ByVal clave As Integer, ByVal indice As Integer)
Dim contar As Integer
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
'esta rutina filtra la posicion por la clave de la posicion o mesa
ReDim mata(1 To 1) As New propPosRiesgo

noreg = UBound(matpos, 1)
contar = 0
For i = 1 To noreg
If Val(matpos(i, indice)) = clave Then
contar = contar + 1
ReDim Preserve mata(1 To contar) As New propPosRiesgo
mata(j) = matpos(i)
End If
Next i
If contar = 0 Then
 ReDim matb(0 To 0, 0 To 0) As Variant
 FiltrarPosClave = matb
 MensajeProc = "no existe posicion para la fecha de analisis"
Else
 FiltrarPosClave = mata
End If
End Function

Function FiltrarPosClave2(ByRef matpos() As propPosRiesgo, ByVal clave As Integer, ByVal indice As Integer)
Dim contar As Integer
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim mata() As New propPosRiesgo

'esta rutina filtra la posicion por la clave de la posicion o mesa
ReDim mata(1 To 1) As New propPosRiesgo

noreg = UBound(matpos, 1)
contar = 0
For i = 1 To noreg
If matpos(i, indice) = clave Or matpos(i, indice) = "PRIMARIA " & clave Or matpos(i, indice) = "PRIMARIA " & clave & " A" Or matpos(i, indice) = "PRIMARIA " & clave & " P" Then
contar = contar + 1
ReDim Preserve mata(1 To contar) As New propPosRiesgo
mata(contar) = matpos(i)
End If
Next i
If contar = 0 Then
 ReDim matb(0 To 0, 0 To 0) As Variant
 FiltrarPosClave2 = matb
 MsgBox "no existe posicion para la fecha de analisis"
Else
 FiltrarPosClave2 = mata
End If
End Function


Private Sub Command36_Click()
Dim fecha As Date
Dim exito As Boolean
Dim txtmsg As String
If PerfilUsuario = "ADMINISTRADOR" Then
Screen.MousePointer = 11
fecha = FechaPos
SiActTProc = True

 Call GuardaResVaRIKOS2(fecha, conAdoBD, txtmsg, exito)
 MsgBox "Proceso terminado"
 Call ActUHoraUsuario
 SiActTProc = False
Screen.MousePointer = 0
Else
MsgBox "No tiene acceso a este modulo"
End If
End Sub

Sub GenerarHistVaR(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim fecha As Date
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim siesfv As Boolean
Dim nomarch As String
Dim mata() As Variant
Dim suma As Double
Dim noesc As Integer
noesc = 500

   ReDim matb(1 To NoPortafolios) As String
   
   matb(1) = "TOTALES"
   matb(2) = "TOTAL MERCADO DE DINERO"
   matb(3) = "TOTAL MESA DE CAMBIOS"
   matb(4) = "DERIVADOS DE NEGOCIACION"
   matb(5) = "DERIVADOS ESTRUCTURALES"
   matb(6) = "DERIVADOS 10"
   fecha1 = #1/1/2012#
   fecha2 = #10/1/2013#
   fecha = fecha1
   Open DirResVaR & "\resumen cvar.txt" For Output As #1
   txtcadena = ""
   For i = 1 To 7
   txtcadena = txtcadena & matb(i) & Chr(9)
   Next i
   Print #1, txtcadena
   Do While fecha <= fecha2
      siesfv = EsFechaVaR(fecha)
      If siesfv Then
         txtcadena = fecha & Chr(9)
         nomarch = DirResVaR & "\Sim hist CONSOLIDADO " & noesc & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
         mata = LeerArchTexto(nomarch, Chr(9), "")
         For j = 1 To 7
             suma = 0
             For i = 1 To UBound(mata, 2)
                 If mata(1, i) = matb(j) Then
                    ReDim matc(1 To UBound(mata, 1) - 1, 1 To 1) As Variant
                    For p = 1 To UBound(mata, 1) - 1
                        matc(p, 1) = Val(mata(p + 1, i))
                    Next p
                    matc = RutinaOrden(matc, 1, 1)
                    suma = 0
                    For p = 1 To 8
                        suma = suma + matc(p, 1)
                    Next p
                    Exit For
                 End If
             Next i
             txtcadena = txtcadena & suma / 8 & Chr(9)
         Next j
         Print #1, txtcadena
      End If
      fecha = fecha + 1
   Loop
   Close #1


End Sub




Private Sub Command38_Click()

If PerfilUsuario = "ADMINISTRADOR" Then
SiActTProc = True
Screen.MousePointer = 11
  frmValidaOper.Show
  Call ActUHoraUsuario
Else
 MsgBox "No tiene acceso autorizado a este modulo"
End If
Screen.MousePointer = 0
End Sub

Sub AgregarProductoBase(ByVal indice As Integer, ByVal txtclave As String)
Dim pcuponact As Integer
Dim pcuponpas As Integer
Dim tcambioact As Double
Dim tcambiopas As Double
Dim curvadescact As String
Dim curvadescpas As String
Dim txtbase As String
Dim txtcadena As String
Dim curvapagact As String
Dim curvapagpas As String


If MatDerEstandar(indice, 8) = "TIIE28[0]" Then
   pcuponact = 28
ElseIf MatDerEstandar(indice, 8) = "TIIE91[0]" Then
   pcuponact = 91
ElseIf MatDerEstandar(indice, 8) = "LIBOR1M[0]" Then
   pcuponact = 30
ElseIf MatDerEstandar(indice, 8) = "LIBOR3M[0]" Then
   pcuponact = 90
ElseIf MatDerEstandar(indice, 8) = "LIBOR6M[0]" Then
   pcuponact = 180
Else
   pcuponact = 0
End If
If MatDerEstandar(indice, 9) = "TIIE28[0]" Then
   pcuponpas = 28
ElseIf MatDerEstandar(indice, 9) = "TIIE91[0]" Then
   pcuponpas = 91
ElseIf MatDerEstandar(indice, 8) = "LIBOR1M[0]" Then
   pcuponpas = 30
ElseIf MatDerEstandar(indice, 9) = "LIBOR3M[0]" Then
   pcuponpas = 90
ElseIf MatDerEstandar(indice, 9) = "LIBOR6M[0]" Then
   pcuponpas = 180
Else
   pcuponpas = 0
End If

If MatDerEstandar(indice, 4) = "USD" Then
   tcambioact = "DOLAR PIP FIX"
ElseIf MatDerEstandar(indice, 4) = "UDI" Then
   tcambioact = "UDI"
Else

End If

If MatDerEstandar(indice, 5) = "USD" Then
   tcambiopas = "DOLAR PIP FIX"
ElseIf MatDerEstandar(indice, 5) = "UDI" Then
   tcambiopas = "UDI"
Else
   tcambiopas = ""
End If
curvadescact = ""
curvadescpas = ""


txtcadena = "INSERT INTO " & txtbase & " VALUES("
txtcadena = txtcadena & "'" & MatDerEstandar(indice, 2) & "',"
txtcadena = txtcadena & "'SWAP',"
txtcadena = txtcadena & "null,"
txtcadena = txtcadena & "null,"
txtcadena = txtcadena & "null,"
txtcadena = txtcadena & pcuponact & ","
txtcadena = txtcadena & pcuponpas & ","
txtcadena = txtcadena & "'" & curvadescact & "',"
txtcadena = txtcadena & "1,"
txtcadena = txtcadena & "'" & curvadescpas & "',"
txtcadena = txtcadena & "1,"
txtcadena = txtcadena & "" & curvapagact & "',"
txtcadena = txtcadena & "1,"
txtcadena = txtcadena & "'" & curvapagpas & "',"
txtcadena = txtcadena & "1,"
txtcadena = txtcadena & "'" & tcambioact & "',"
txtcadena = txtcadena & "'" & tcambiopas & "')"
ConAdo.Execute txtcadena


End Sub


Sub ImpReporteEfProsFwd(ByVal fecha As Date, ByVal cemision As String, ByVal eficpros As Double)
Dim nomarch1 As String
Dim txtcadena As String
Dim j As Integer
Dim noreg As Integer
Dim exitoarch As Boolean

   nomarch1 = DirResVaR & "\Resumen eficiencia prospectiva fwd " & cemision & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch1
   frmCalVar.CommonDialog1.ShowSave
   nomarch1 = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch1, 5, exitoarch)
   If exitoarch Then
      txtcadena = "Fecha" & Chr(9)
      txtcadena = txtcadena & "No. de simulaciones" & Chr(9)
      txtcadena = txtcadena & "No. de aciertos"
      Print #5, txtcadena
      noreg = UBound(MatResEficFwd, 1)
      For j = 2 To noreg
      txtcadena = MatResEficFwd(j, 1) & Chr(9)
      txtcadena = txtcadena & 10 & Chr(9)
      txtcadena = txtcadena & MatResEficFwd(j, 2)
      Print #5, txtcadena
      Next j
      Print #5, "Porcentaje de aciertos" & Chr(9) & Format(eficpros, "##0.00 %")
      Close #5
      MsgBox "Se debe de tomar el archivo " & nomarch1 & ", pegarse en un documento de word y enviarse al area de Derivados."
   End If
End Sub






Function AnalisisPosEfec(ByRef matpos() As Variant)
Dim noreg As Integer

'como se define la matriz mata
'1 en la primera celda Clave de operación
'2 en la segunda celda clave pos primaria activa
'3 en la tercera celda clave pos primaria pasiva
'4 el tipo de efectividad
noreg = UBound(matpos, 1)

If noreg = 1 Then 'es un forward
 ReDim mata(1 To 1, 1 To 4) As Variant
   mata(1, 1) = matpos(1).c_operacion   'la Clave de operación
   mata(1, 4) = 4                        'propia de un forward
   AnalisisPosEfec = mata
ElseIf noreg = 2 Then
ReDim mata(1 To 1, 1 To 4) As Variant
   If Left(matpos(1).c_operacion, 8) = "PRIMARIA" Then   'solo la posicion primaria puede ser negativa
      mata(1, 1) = matpos(2).c_operacion   'la Clave de operación
      If matpos(1).IFlujoActSwap <> 0 And matpos(1).FFlujoActSwap <> 0 Then 'activo
         mata(1, 2) = matpos(1).c_operacion 'activo
         mata(1, 4) = 2                    'eficiencia de una pos activa
      ElseIf matpos(1).IFlujoPasSwap <> 0 And matpos(1).FFlujoPasSwap <> 0 Then
         mata(1, 3) = matpos(1).c_operacion 'pasivo
         mata(1, 4) = 1                    'eficiencia de una pos pasiva
      End If
   ElseIf Left(matpos(2).c_operacion, 8) = "PRIMARIA" Then    'solo la posicion primaria puede ser negativa
      mata(1, 1) = matpos(1).c_operacion 'la Clave de operación
      If matpos(2).Tipo_Mov = 1 Then
         mata(1, 3) = matpos(2).c_operacion 'activa
         mata(1, 4) = 2
      Else
         mata(1, 3) = matpos(2).c_operacion 'pasiva
         mata(1, 4) = 1                    'pasiva
      End If
   End If
 AnalisisPosEfec = mata
ElseIf noreg = 3 Then
 ReDim mata(1 To 1, 1 To 4) As Variant
 mata(1, 4) = 3
 If matpos(1).Signo_Op = 1 Then 'el primer registro
  If matpos(2).IFlujoActSwap <> 0 And matpos(2).FFlujoActSwap And matpos(3).IFlujoPasSwap <> 0 And matpos(3).FFlujoPasSwap Then
   mata(1, 1) = matpos(1).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  ElseIf matpos(2).IFlujoPasSwap <> 0 And matpos(2).FFlujoPasSwap And matpos(3).IFlujoActSwap <> 0 And matpos(3).FFlujoActSwap Then
   mata(1, 1) = matpos(1).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  End If
 ElseIf matpos(2).Signo_Op = 1 Then 'el segundo registro
  If matpos(1).IFlujoActSwap <> 0 And matpos(1).FFlujoActSwap And matpos(3).IFlujoPasSwap <> 0 And matpos(3).FFlujoPasSwap Then
   mata(1, 1) = matpos(1).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  ElseIf matpos(1).IFlujoPasSwap <> 0 And matpos(1).FFlujoPasSwap And matpos(2).IFlujoActSwap <> 0 And matpos(2).FFlujoActSwap Then
   mata(1, 1) = matpos(1).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  End If
 ElseIf matpos(3).Signo_Op = 1 Then 'el primer registro
  If matpos(1).IFlujoActSwap <> 0 And matpos(1).FFlujoActSwap And matpos(2).IFlujoPasSwap <> 0 And matpos(2).FFlujoPasSwap Then
   mata(1, 1) = matpos(3).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  ElseIf matpos(1).IFlujoPasSwap <> 0 And matpos(1).FFlujoPasSwap And matpos(2).IFlujoActSwap <> 0 And matpos(2).FFlujoActSwap Then
   mata(1, 1) = matpos(1).c_operacion
   mata(1, 2) = matpos(2).c_operacion
   mata(1, 3) = matpos(3).c_operacion
  End If
 End If
 AnalisisPosEfec = mata
Else
 MsgBox "No se puede determinar el tipo de calculo de efectividad"
End If
End Function

Function AnalisisPosEfec2(ByRef matpos() As Variant) As Variant()
Dim contar As Integer
Dim i As Integer
Dim noreg As Integer

'en la matriz mata se definiran las posiciones individuales
contar = 0
ReDim mata(1 To 1) As Variant
noreg = UBound(matpos, 1)
For i = 1 To noreg
If "" & Val(matpos(i).c_operacion) & "" = matpos(i).c_operacion Then 'esta es una posicion neta
 contar = contar + 1
 ReDim Preserve mata(1 To contar) As Variant
 mata(contar) = matpos(i).c_operacion
End If
Next i
AnalisisPosEfec2 = mata
End Function


Sub GenBackPortPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtgrupo As String, ByRef exito As Boolean)
On Error GoTo hayerror
Dim valor As Variant
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim suma1 As Double
Dim suma2 As Double
Dim suma3 As Double
Dim i As Integer
Dim noreg As Integer
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResBack & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion
txtfiltro2 = txtfiltro2 & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtgrupo & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma1 = 0: suma2 = 0: suma3 = 0
   For i = 1 To noreg
       valor1 = rmesa.Fields(5)
       valor2 = rmesa.Fields(6)
       valor3 = rmesa.Fields(7)
       suma1 = suma1 + valor1
       suma2 = suma2 + valor2
       suma3 = suma3 + valor3
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Sumando las las p y g del portafolio " & txtgrupo & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   rmesa.Close
   txtborra = "DELETE FROM " & TablaBackPort & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtgrupo & "'"
   ConAdo.Execute txtborra
   txtcadena = "INSERT INTO " & TablaBackPort & " VALUES ("
   txtcadena = txtcadena & txtfecha & ","                    'la fecha de proceso
   txtcadena = txtcadena & "'" & txtport & "',"              'el portafolio
   txtcadena = txtcadena & "'" & txtgrupo & "',"             'el subportafolio
   txtcadena = txtcadena & suma1 & ","                       'valuacion de escenario base
   txtcadena = txtcadena & suma2 & ","                       'valuacion de escenario base
   txtcadena = txtcadena & suma3 & ")"                       'valuacion de escenario base
   ConAdo.Execute txtcadena
End If
exito = True
On Error GoTo 0
Exit Sub
hayerror:
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Private Sub Command53_Click()
Screen.MousePointer = 11
If FechaEval <> 0 Then
  MsgBox "Atencion. Si ajusta la valuacion de las emisisiones se perderar los parametros de estres cargados en la sesion"
  MatFactRiesgo = TransfFRSplit(FechaEval, MatFactRiesgo)
End If
Screen.MousePointer = 0
End Sub

Sub GeneraResCVaRSim(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef exito As Boolean)
Dim txtfecha As String
Dim valor As Double
Dim txtborra As String
Dim txttvar As String
Dim i As Integer
Dim j As Integer
Dim noport As Integer
Dim noesc1 As Long
Dim nconf As Double

txttvar = "CVARH"
exito = False

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
For j = 1 To UBound(MatGruposPortPos, 1)
    noesc1 = 250: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
    noesc1 = 500: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
    noesc1 = 750: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
    noesc1 = 1000: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
    noesc1 = 1250: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
    noesc1 = 1500: nconf = 0.03
    valor = CalcularCVaRPyG2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), 1500, 1, noesc1, nconf, exito)
    If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & valor
Next j
End Sub

Sub GeneraResVaRExpSim(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal lambda As Double, ByRef exito As Boolean)
Dim txtfecha As String
Dim valor As Double
Dim txtborra As String
Dim i As Integer
Dim j As Integer
Dim noesc1 As Long
exito = False

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For j = 1 To UBound(MatGruposPortPos, 1)
       noesc1 = 250
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
       noesc1 = 500
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
       noesc1 = 750
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
       noesc1 = 1000
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
       noesc1 = 1250
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
       noesc1 = 1500
       valor = CalcularVaRExp2(fecha, txtport, txtportfr, MatGruposPortPos(j, 3), noesc, htiempo, noesc1, nconf, lambda, exito)
       If valor <> 0 Then Print #1, fecha & Chr(9) & MatGruposPortPos(j, 3) & Chr(9) & noesc1 & Chr(9) & nconf & Chr(9) & lambda & Chr(9) & valor
    Next j
End If
End Sub

Private Sub Form_Load()
  If OpcionBDatos = 1 Then
       frmCalVar.Caption = "Sistema VaR de Mercado Banobras (Producción)"
  ElseIf OpcionBDatos = 2 Then
       frmCalVar.Caption = "Sistema VaR de Mercado Banobras (Desarrollo)"
  ElseIf OpcionBDatos = 3 Then
       frmCalVar.Caption = "Sistema VaR de Mercado Banobras (DRP)"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   MensajeProc = NomUsuario & " ha salido del sistema"
   Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
   Call DesbloquearUsuario(NomUsuario)
   RGuardarPL.Close
   RegResCVA.Close
   ConAdo.Close
   conAdoBD.Close
   End
End Sub

Private Sub LecMinusPlus_Click()
frmAnalisisPlusMinus.Show
End Sub

Private Sub mABalanza_Click()
frmABalanza.Show 1
End Sub

Private Sub mAcerca_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
 frmAcerca.Show 1
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mADFact_Click()
End Sub


Private Sub mAFRiesgo_Click()
End Sub

Private Sub mActSQLPort_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
   frmEdSQLPort.Show
End If
End Sub

Private Sub mactswr3r_Click()
Dim tfecha As String
Dim fecha As Date
tfecha = InputBox("Dame la fecha del proceso", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Screen.MousePointer = 11
   Call Lee_Ikos_Derivados(fecha)
   Screen.MousePointer = 0
   MsgBox "Fin de proceso"
End If

End Sub

Private Sub mActValIKOS_Click()
Dim tfecha As String
Dim fecha As Date
tfecha = InputBox("Dame la fecha del proceso", , Date)
If IsDate(tfecha) Then
   SiActTProc = True
   fecha = CDate(tfecha)
   Screen.MousePointer = 11
   frmProgreso.Show
   Call ActValIKOSValSVM(fecha)
   Unload frmProgreso
   SiActTProc = False
   Screen.MousePointer = 0
   MsgBox "Fin de proceso"
End If
End Sub

Private Sub mAnalisisBC_Click()
 frmABalanza.Show
End Sub

Private Sub mAutoVAR_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11

Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mAPrecios_Click()
End Sub

Private Sub mAVMCVAR_Click()
frmAVCVAR.Show 1
End Sub

Private Sub mBack_Click()
 frmHistBack.Show 1
End Sub



Private Sub mBitacora_Click()
frmBitacora.Show
End Sub


Private Sub mCalcMW_Click()
frmCMW.Show
End Sub

Private Sub mCargaFR_Click()
Dim nomarch As String
Dim sihayarch As Boolean

Screen.MousePointer = 11
 frmCalVar.CommonDialog1.ShowOpen
 nomarch = frmCalVar.CommonDialog1.FileName
 sihayarch = VerifAccesoArch(nomarch)
 If sihayarch Then
   SiActTProc = True
   frmProgreso.Show
   Call ImpFRExcel(nomarch, "Hoja1$")
   Unload frmProgreso
   MsgBox "Fin de proceso"
   Call ActUHoraUsuario
   SiActTProc = False
 Else
   MsgBox "El archivo " & nomarch & " no esta disponible"
 End If
Screen.MousePointer = 0

End Sub

Private Sub mCargarPosF_Click()

End Sub

Private Sub mCargapp_Click()
frmCargaPosP.Show
End Sub

Private Sub mCargarFEE_Click()
Dim nomarch As String
Dim sihayarch As Boolean
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Integer
Dim encontro As Boolean
Dim Base As DAO.Database
Dim registros As DAO.recordset

Screen.MousePointer = 11

MsgBox "Esta rutina sube los flujos de emisiones especiales a una tabla en oracle"
frmEjecucionProc2.CommonDialog1.ShowOpen
nomarch = frmEjecucionProc2.CommonDialog1.FileName
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
frmProgreso.Show
Set Base = OpenDatabase(nomarch, dbDriverNoPrompt, True, VersExcel)
Set registros = Base.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
If registros.RecordCount <> 0 Then
registros.MoveLast
noreg = registros.RecordCount
nocampos = registros.Fields.Count
ReDim mata(1 To noreg, 1 To nocampos) As Variant
registros.MoveFirst
For i = 1 To noreg
For j = 1 To nocampos
mata(i, j) = LeerTAccess(registros, j - 1, i)
Next j
registros.MoveNext
Next i
End If
registros.Close
Base.Close
ReDim matem(1 To 2, 0 To 0) As Variant
contar = 0
For i = 1 To noreg
    If contar <> 0 Then
       encontro = False
       For j = 1 To contar
         If matem(1, j) = mata(i, 1) And matem(2, j) = mata(i, 2) Then
            encontro = True
            Exit For
         End If
       Next j
       If Not encontro Then
          contar = contar + 1
          ReDim Preserve matem(1 To 2, 0 To contar) As Variant
          matem(1, contar) = mata(i, 1)
          matem(2, contar) = mata(i, 2)
       End If
    Else
       contar = contar + 1
       ReDim Preserve matem(1 To 2, 0 To contar) As Variant
       matem(1, contar) = mata(i, 1)
       matem(2, contar) = mata(i, 2)
    End If
Next i
matem = MTranV(matem)
If noreg <> 0 Then
 Call GuardaFlujosEmEsp(matem, mata)
 SiCargaFEmMD = True
End If
Unload frmProgreso
Else
 MsgBox "No hay acceso al archivo " & nomarch
End If
MsgBox "Proceso terminado"
Screen.MousePointer = 0

End Sub


Private Sub mCNombre_Click()
Dim i As Long
Dim noreg As Long

Screen.MousePointer = 11
noreg = 1000000
ReDim mata(1 To noreg, 1 To 1) As Variant
For i = 1 To noreg
mata(i, 1) = Rnd * noreg
Next
mata = RutinaOrden(mata, 1, SRutOrden)
For i = 1 To noreg
If i > 1 Then
If mata(i - 1, 1) > mata(i, 1) Then
MsgBox "hay un error en el algoritmo " & i
End If
End If
Next i
Screen.MousePointer = 0
End Sub



Private Sub mCCVA_Click()
frmCVA.Show
End Sub

Private Sub mCLimC_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
  frmCLimiteC.Show
Else
  MsgBox "No tiene acceso a este modulo"
End If

End Sub

Private Sub mCrearFR_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
frmCrearFR.Show
Else
MsgBox "No tiene acceso a este modulo"
End If
End Sub


Private Sub mDetalles_Click()
Dim i As Integer
Dim j As Integer

    frmDesglosePos.Show
    Call VerDetallesPosicion(frmDesglosePos.MSFlexGrid1, MatPosRiesgo)
    If Not EsArrayVacio(MatNodosFREx) Then
    frmDesglosePos.MSFlexGrid2.Rows = UBound(MatNodosFREx, 1) + 1
    For i = 1 To UBound(MatNodosFREx, 1)
        frmDesglosePos.MSFlexGrid2.TextMatrix(i, 1) = MatNodosFREx(i, 1)
    Next i
    End If
End Sub

Private Sub mDetallesVMark_Click()
If Not EsArrayVacio(MatCovar1) Then
   frmVaRMark.Show
Else
   MsgBox "No se ha calculado ningun VaR Mark"
End If
End Sub

Private Sub mCVaRFP_Click()
 frmFondoPensiones.Show
End Sub

Private Sub mDALM_Click()
Dim fecha As Date
Dim txtfecha As String
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim mata() As Variant
Dim matcampos() As String
fecha = #8/31/2018#
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
Open "d:\Datos alm " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
txtfiltro = "SELECT * FROM LIQ_TESORERIA WHERE FECHA_CORTE = " & txtfecha
Call LeerBDatosO(txtfiltro, mata, matcampos)
txtcadena = ""
For i = 1 To UBound(matcampos, 1)
    txtcadena = txtcadena & matcampos(i) & Chr(9)
Next i
Print #1, txtcadena
For i = 1 To UBound(mata, 1)
     txtcadena = ""
     For j = 1 To UBound(matcampos, 1)
         txtcadena = txtcadena & mata(i, j) & Chr(9)
     Next j
     Print #1, txtcadena
Next i
Print #1, ""


txtfiltro = "SELECT * FROM LIQ_REPORTOS WHERE FECHA_CORTE = " & txtfecha
Call LeerBDatosO(txtfiltro, mata, matcampos)
txtcadena = ""
For i = 1 To UBound(matcampos, 1)
    txtcadena = txtcadena & matcampos(i) & Chr(9)
Next i
Print #1, txtcadena
For i = 1 To UBound(mata, 1)
     txtcadena = ""
     For j = 1 To UBound(matcampos, 1)
         txtcadena = txtcadena & mata(i, j) & Chr(9)
     Next j
     Print #1, txtcadena
Next i
Print #1, ""

txtfiltro = "SELECT * FROM " & TablaPosMesaIKOS & " WHERE F_POSICION = '" & Format(fecha, "YYYYMMDD") & "'"
Call LeerBDatosO(txtfiltro, mata, matcampos)
txtcadena = ""
For i = 1 To UBound(matcampos, 1)
    txtcadena = txtcadena & matcampos(i) & Chr(9)
Next i
Print #1, txtcadena
For i = 1 To UBound(mata, 1)
     txtcadena = ""
     For j = 1 To UBound(matcampos, 1)
         txtcadena = txtcadena & mata(i, j) & Chr(9)
     Next j
     Print #1, txtcadena
Next i


Close #1
MsgBox "Fin de proceso"
End Sub

Sub LeerBDatosO(ByVal txtfiltro As String, ByRef mata() As Variant, ByRef matcampos() As String)
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   ReDim matcampos(1 To nocampos) As String
   
   For i = 1 To nocampos
       matcampos(i) = rmesa.Fields(i - 1).Name
   Next i
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1).value
       Next j
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
End If
End Sub


Private Sub mEdCatalogos_Click()
   If PerfilUsuario = "ADMINISTRADOR" Then
      frmCatalogos.Show
   Else
      MsgBox "No tiene permisos para actualizar los catalogos del sistema"
  End If
End Sub

Private Sub mEficiencia2_Click()
Dim id_tabla As Integer
Dim tfecha As String
Dim txtfecha As String
Dim mateficpros() As Variant
Dim matpos1() As Variant
Dim fecha As Date
Dim eficpros As Double
Dim indice As Integer
Dim i As Integer
Dim coperacion As String   'clave de operacion en el sistema ikos de un swap
Dim txtport As String
Dim exito As Boolean
Dim exito2 As Boolean
Dim txtmsg2 As String
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg3 As String
id_tabla = 3

Screen.MousePointer = 11
  tfecha = InputBox("Dame la fecha de la simulacion", , Date)
  coperacion = InputBox("Dame la clave de operacion", , "0000")
  frmProgreso.Show
  SiActTProc = True
  indice = 0
  For i = 1 To UBound(MatRelSwapsPrim, 1)
      If coperacion = MatRelSwapsPrim(i).coperacion Then
         indice = i
         Exit For
      End If
  Next i

  If IsDate(tfecha) And indice <> 0 Then
     fecha = tfecha
     txtport = "Port Efect pros " & coperacion
     Call GenPortEfect(fecha, coperacion, MatRelSwapsPrim(indice).c_ppactiva, MatRelSwapsPrim(indice).c_pppasiva, MatRelSwapsPrim(indice).c_pswap, txtport)
     'crear el portafolio de efectividad en var_td_port_pos_3
'generar el proceso de efectividad que le corresponde a la opeacion
     If MatRelSwapsPrim(indice).t_efect = 1 Then
        Call GenProcEfecProsSwap(fecha, txtport, 58, id_tabla)
     ElseIf MatRelSwapsPrim(indice).t_efect = 2 Then
        Call GenProcEfecProsSwap(fecha, txtport, 59, id_tabla)
     ElseIf MatRelSwapsPrim(indice).t_efect = 3 Then
        Call GenProcEfecProsSwap(fecha, txtport, 60, id_tabla)
     ElseIf MatRelSwapsPrim(indice).t_efect = 5 Then
        Call GenProcEfecProsSwap(fecha, txtport, 62, id_tabla)
     End If
  Else
     MsgBox "No esta definida la relacion de efectividad en " & PrefijoBD & TablaPosPrimarias
  End If
  SiActTProc = False
  Unload frmProgreso
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub



Private Sub mEficProsFwd_Click()
Dim txtfecha As String
Dim txtcoperacion As String
Dim fecha As Date
Dim txtmsg As String
Dim txtport As String
Dim exito As Boolean

Screen.MousePointer = 11
  txtfecha = InputBox("Dame la fecha de la simulacion", , FechaEval)
  txtcoperacion = InputBox("Dame la clave de operacion", , "0000")
  If IsDate(txtfecha) And Not EsVariableVacia(txtcoperacion) Then
     frmProgreso.Show
     fecha = CDate(txtfecha)
     Call CalculaEficProsFWD(fecha, txtport, txtmsg, exito)
     Unload frmProgreso
  End If
Screen.MousePointer = 0
End Sub


Private Sub mefretro2_Click()
Dim tfecha As String
Dim txtmsg As String
Dim exito As Boolean
Dim fecha As Date
Screen.MousePointer = 11
tfecha = InputBox("Dame la fecha", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Call ProcEficRetro(fecha, txtmsg, exito)
End If
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub mexport1_Click()
Dim tfecha1 As String
Dim tfecha2 As String
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
tfecha1 = InputBox("Dame la primera fecha", , Date)
tfecha2 = InputBox("Dame la segunda fecha", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
   fecha1 = CDate(tfecha1)
   fecha2 = CDate(tfecha2)
   SiActTProc = True
   Screen.MousePointer = 11
   frmProgreso.Show
   Call ExportarValDerivCF(fecha1, fecha2)
   Unload frmProgreso
   Call ActUHoraUsuario
End If
SiActTProc = False
MsgBox "Fin del proceso"
frmCalVar.SetFocus
Screen.MousePointer = 0
End Sub

Private Sub mexpVFD_Click()
Dim fechatxt As String
Dim bl_exito As Boolean
Dim fecha As Date
Dim txtmsg As String
Dim exito As Boolean
Dim id_tabla As Integer
If PerfilUsuario = "ADMINISTRADOR" Then
   Screen.MousePointer = 11
'rutina que genera los flujos de la posicion de swaps sin descontar
   fechatxt = ReemplazaVacioValor(InputBox("Dame la fecha de calculo", , Date), "")
   id_tabla = ReemplazaVacioValor(InputBox("Dame la tabla de subprocesos(1 o 2)", , 1), 0)
   If IsDate(fechatxt) And (id_tabla = 1 Or id_tabla = 2) Then
      SiActTProc = True
      fecha = CDate(fechatxt)
      frmProgreso.Show
      Call GenSubProcValDeriv(fecha, "SWAPS", 2, 94, id_tabla, txtmsg, exito)
      Unload frmProgreso
      Call ActUHoraUsuario
      SiActTProc = False
      MsgBox "Fin de proceso"
   End If
   Screen.MousePointer = 0
Else
  MsgBox "No tiene acceso a este modulo"
End If
End Sub

Private Sub mExtraTIIE_Click()
frmExtrapolYieldTIIE.Show
End Sub

Private Sub mFlujosPIP_Click()
Dim fecha As Date
Dim tfecha As String
Dim dirarch As String
tfecha = InputBox("Dame la fecha", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   dirarch = DirTemp
   Screen.MousePointer = 11
   frmProgreso.Show
       Call ImportFlujosPIP(fecha, dirarch)
   Unload frmProgreso
   MsgBox "Fin de proceso"
   Screen.MousePointer = 0
End If

End Sub

Private Sub mGenAlMont_Click()
Screen.MousePointer = 11
frmProgreso.Show
 Call GenNormMont(#1/1/2018#, 700, 10000)
Unload frmProgreso
Screen.MousePointer = 0
End Sub

Private Sub mGenDatosEm_Click()
frmGenDatosEm.Show
End Sub

Private Sub mGenSimFactR_Click()
  frmSimFactR.Show
End Sub

Private Sub mListaP_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
   frmProcesos.Show
Else
  MsgBox "No puede acceder a la generacion de procesos"
End If
End Sub


Private Sub mMontecarlo_Click()

End Sub



Private Sub mOpDeriv_Click()
End Sub


Private Sub mMarcoOp_Click()
frmMarcoOp.Show
End Sub

Private Sub mMonitorVO_Click()
  SiActTProc = True
   frmMonitorValidación.Show
End Sub

Private Sub mParametros_Click()
frmParametros.Show 1
End Sub

Private Sub mImportardatos_Click()
Screen.MousePointer = 11
If PerfilUsuario = "ADMINISTRADOR" Then
 frmEjecucionProc.Show
Else
 MsgBox "Usted no tiene acceso a este modulo del sistema"
End If

End Sub

Private Sub mRegresar_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Unload Me
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub mPPrimaria_Click()
frmPPrimaria.Show 1
End Sub


Private Sub mprocesos3_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
frmEjecucionProc2.Show
Else
MsgBox "no puede ejecutar procesos secundarios"
End If
End Sub


Private Sub mRepDerv_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtport As String
Dim tipopos As String
Dim txtfecha As String
Dim curva1() As New propCurva
Dim curva2() As New propCurva
Dim curva3() As New propCurva
Dim mcurvat() As Variant
Dim bl_exito As Boolean
Dim noreg As Integer
Dim noregpos As Integer
Dim noflujos As Long
Dim txtpossim As String
Dim fechaid As Date
Dim i As Integer
Dim j As Integer
Dim contar As Integer
Dim dxv As Integer
Dim factor As Double
Dim valusd As Double
Dim valudi As Double
Dim valyen As Double
Dim nomarch As String
Dim txtsalida As String
Dim txtcadena As String
Dim exito As Boolean
Dim exitoarch As Boolean


SiAnexarFlujosSwaps = True
txtfecha = InputBox("Dame la fecha del reporte", , Date)
If IsDate(txtfecha) Then
fecha = CDate(txtfecha)
fecha1 = Date - 10
fecha2 = Date
ValExacta = True  'valuacion exacta
frmProgreso.Show
Call CrearMatFRiesgo2(fecha1, fecha2, MatFactRiesgo, "", exito)
Unload frmProgreso
txtport = "SWAPS"
tipopos = 1

'primero se procede a leer la interfase de
Call CalculaValPos(fecha, fecha, fecha, txtport, 1, bl_exito) 'cargando la posicion y valuando a la fecha
curva1 = CrearCurva(fecha, "DOLAR PIP FIX", mcurvat, MatFactR1, "N")
valusd = CalculaTasa(curva1, 0, 1)
curva2 = CrearCurva(fecha, "UDI", mcurvat, MatFactR1, "N")
valudi = CalculaTasa(curva2, 0, 1)
curva3 = CrearCurva(fecha, "YEN BM", mcurvat, MatFactR1, "N")
valyen = CalculaTasa(curva3, 0, 1)

noreg = UBound(MatContrapartes1, 1)
noregpos = UBound(MatPosRiesgo, 1)
noflujos = UBound(MatValFlujosD, 2)
ReDim mattabla(1 To noreg, 1 To 15) As Variant   'consolidado por contraparte
ReDim mattabla2(1 To noreg, 1 To 15) As Variant   'consolidado por contraparte
ReDim matpos1(1 To noregpos, 1 To 8) As Variant
contar = 0

For i = 1 To noregpos
    matpos1(i, 1) = MatPosRiesgo(i).c_operacion        'Clave de operación
    'matpos1(i, 3) = MatPosriesgo(i)           'fecha de vencimiento
    'matpos1(i, 5) = MatPosriesgo(i).TCambio2Swap        'tipo de cambio pasiva
    For j = 1 To noflujos
 'se filtran los flujos de la emision
    If matpos1(i, 1) = MatValFlujosD(2, j) And (MatValFlujosD(5, j) <= fecha And fecha < MatValFlujosD(6, j)) And MatValFlujosD(3, j) = "C" Then
       matpos1(i, 6) = MatValFlujosD(7, j)          'saldo pasivo vigente
    End If
 Next j
Next i
For i = 1 To noreg
    mattabla(i, 1) = MatContrapartes1(i, 3)                 'descripcion de la contraparte
    mattabla2(i, 1) = MatContrapartes1(i, 3)                'descripcion de la contraparte
    If MatContrapartes1(i, 5) = "MXN" Then
       mattabla(i, 2) = MatContrapartes1(i, 4) * 1000000      'limite de la contraparte en pesos
       mattabla2(i, 2) = MatContrapartes1(i, 4) * 100000      'limite de la contraparte en pesos
    ElseIf MatContrapartes1(i, 5) = "USD" Then
       mattabla(i, 2) = MatContrapartes1(i, 4) * 1000000 * valusd  'limite de la contraparte en pesos
       mattabla2(i, 2) = MatContrapartes1(i, 4) * 1000000 * valusd  'limite de la contraparte en pesos
    End If
mattabla(i, 3) = 0
mattabla(i, 3) = 0
mattabla2(i, 3) = 0
For j = 1 To noregpos
    If MatContrapartes1(i, 2) = matpos1(j, 2) Then
       If matpos1(j, 5) = "UDI" Then      'suma de patas en udis
          mattabla(i, 3) = mattabla(i, 3) + Minimo(matpos1(j, 4), 0) 'pata en udis
          mattabla2(i, 3) = mattabla2(i, 3) + matpos1(j, 4)           'pata en udis
       ElseIf matpos1(j, 5) = "UDI" Then  'suma de patas en dolares
          mattabla(i, 4) = mattabla(i, 4) + Minimo(matpos1(j, 4), 0) 'pata en USD
          mattabla2(i, 4) = mattabla2(i, 4) + matpos1(j, 4)          'pata en USD
       Else
          mattabla(i, 5) = mattabla(i, 5) + Minimo(matpos1(j, 4), 0) 'pata en PESOS
          mattabla2(i, 5) = mattabla2(i, 5) + matpos1(j, 4)          'pata en PESOS
       End If
 'suma total
       mattabla(i, 6) = mattabla(i, 3) + mattabla(i, 4) + mattabla(i, 5)
       mattabla2(i, 6) = mattabla2(i, 3) + mattabla2(i, 4) + mattabla2(i, 5)
    End If
Next j
mattabla(i, 7) = Abs(mattabla(i, 6)) 'valor absoluto de mtm
mattabla2(i, 7) = Abs(mattabla2(i, 6)) 'valor absoluto de mtm
'calculo de EPD y SD
For j = 1 To noregpos  'se encuentran los saldos de la parte pasiva
  If matpos1(j, 4) < 0 Then
  If MatContrapartes1(i, 2) = matpos1(j, 2) Then
   dxv = matpos1(j, 3) - fecha
   If dxv / 365 <= 1 Then
   factor = 0
   ElseIf dxv / 365 > 1 And dxv / 365 <= 5 Then
   factor = 0.005
   ElseIf dxv / 365 > 5 Then
   factor = 0.015
   End If
   If matpos1(j, 5) = "UDI" Then
    mattabla(i, 8) = mattabla(i, 8) + matpos1(j, 6) * valudi * factor / 12
   ElseIf matpos1(j, 5) = "DOLAR PIP FIX" Then
    mattabla(i, 9) = mattabla(i, 9) + matpos1(j, 6) * valusd * factor / 12
   Else
    mattabla(i, 10) = mattabla(i, 10) + matpos1(j, 6) * factor / 12
   End If
   mattabla(i, 11) = mattabla(i, 8) + mattabla(i, 9) + mattabla(i, 10)
  End If
  End If
Next j
If mattabla(i, 7) > 0 Then
 mattabla(i, 12) = mattabla(i, 7) + mattabla(i, 11)
 mattabla(i, 13) = Maximo(0, mattabla(i, 12) - mattabla(i, 2))
End If
Next i
'se crea el archivo con los resultados
nomarch = DirResVaR & "\Regulatorio " & Format(fecha, "yyyy-mm-dd") & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName

Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
Print #1, "Dolar fix " & Chr(9) & valusd
Print #1, "Valor UDI " & Chr(9) & valudi
Print #1, "Valor Yen " & Chr(9) & valyen

txtsalida = "Swap" & Chr(9) & "Contraparte" & Chr(9) & "Dias venc" & Chr(9) & "MTM" & Chr(9)
txtsalida = txtsalida & "Saldo p pasiva" & Chr(9) & "Moneda p pasiva"
Print #1, txtsalida
For i = 1 To noregpos
txtsalida = ""
txtsalida = txtsalida & matpos1(i, 1) & Chr(9)         'clave de operacion
txtsalida = txtsalida & matpos1(i, 2) & Chr(9)         'clave contraparte
txtsalida = txtsalida & matpos1(i, 3) - fecha & Chr(9) 'fecha
txtsalida = txtsalida & matpos1(i, 4) & Chr(9)         'valuacion ikos
txtsalida = txtsalida & matpos1(i, 6) & Chr(9)         'saldo fecha
txtsalida = txtsalida & matpos1(i, 5)                  'moneda pos pasiva
Print #1, txtsalida
Next i
Print #1, ""

txtsalida = "Clave de operacion" & Chr(9) & "Pata" & Chr(9)
txtsalida = txtsalida & "Inicio flujo" & Chr(9) & "Fin flujo" & Chr(9)
txtsalida = txtsalida & "Saldo (MO)" & Chr(9) & "Amortizacion (MO)" & Chr(9)
txtsalida = txtsalida & "Intereses (MO)" & Chr(9) & "Amortizacion+int (MO)" & Chr(9)
txtsalida = txtsalida & "Saldo (pesos)" & Chr(9) & "Amortizacion (pesos)" & Chr(9)
txtsalida = txtsalida & "Intereses (pesos)" & Chr(9) & "Amortizacion+int (pesos)" & Chr(9)

Print #1, txtsalida
 For i = 1 To UBound(MatValFlujosD, 2)
  txtcadena = ""
  txtcadena = txtcadena & MatValFlujosD(2, i) & Chr(9)                        'Clave de operación
  txtcadena = txtcadena & MatValFlujosD(3, i) & Chr(9)                        'pata
  txtcadena = txtcadena & MatValFlujosD(5, i) & Chr(9)                        'inicio del flujo
  txtcadena = txtcadena & MatValFlujosD(6, i) & Chr(9)                        'final del flujo
  txtcadena = txtcadena & MatValFlujosD(7, i) & Chr(9)                        'Saldo(MO)
  txtcadena = txtcadena & MatValFlujosD(8, i) & Chr(9)                        'amortizacion (MO)
  txtcadena = txtcadena & MatValFlujosD(14, i) & Chr(9)                       'intereses (MO)
  txtcadena = txtcadena & MatValFlujosD(15, i) & Chr(9)                       'amortizacion + intereses (MO)
  txtcadena = txtcadena & MatValFlujosD(17, i) & Chr(9)                       'moneda
  txtcadena = txtcadena & MatValFlujosD(7, i) * MatValFlujosD(17, i) & Chr(9) 'Saldo(pesos)
  txtcadena = txtcadena & MatValFlujosD(8, i) * MatValFlujosD(17, i) & Chr(9) 'amortizacion (pesos)
  txtcadena = txtcadena & MatValFlujosD(14, i) * MatValFlujosD(17, i) & Chr(9) 'intereses (pesos)
  txtcadena = txtcadena & MatValFlujosD(15, i) * MatValFlujosD(17, i) & Chr(9) 'amortizacion + intereses (pesos)
  Print #1, txtcadena
 Next i
Print #1, ""
txtsalida = "Contraparte" & Chr(9) & "Threshold (Pesos)" & Chr(9) & "MTM en UDIS (pesos)" & Chr(9)
txtsalida = txtsalida & "MTM USD (pesos)" & Chr(9) & "MTM PESOS" & Chr(9)
txtsalida = txtsalida & "MTM TOTAL" & Chr(9) & "abs(MTM negativo)" & Chr(9)
txtsalida = txtsalida & "Nocionales UDIS" & Chr(9) & "Nocionales USD" & Chr(9)
txtsalida = txtsalida & "Nocionales pesos" & Chr(9) & "Total nocionales" & Chr(9)
txtsalida = txtsalida & "EPD" & Chr(9) & "SD"
Print #1, txtsalida
For i = 1 To noreg
txtsalida = ""
For j = 1 To 13
txtsalida = txtsalida & mattabla(i, j) & Chr(9)
Next j
Print #1, txtsalida
Next i
txtsalida = "Contraparte" & Chr(9) & "Threshold (Pesos)" & Chr(9) & "MTM en UDIS (pesos)" & Chr(9)
txtsalida = txtsalida & "MTM USD (pesos)" & Chr(9) & "MTM PESOS" & Chr(9)
txtsalida = txtsalida & "MTM TOTAL" & Chr(9) & "abs(MTM)"
Print #1, txtsalida
For i = 1 To noreg
txtsalida = ""
For j = 1 To 7
txtsalida = txtsalida & mattabla2(i, j) & Chr(9)
Next j
Print #1, txtsalida
Next i
Close #1
End If
End If
SiAnexarFlujosSwaps = False
MsgBox "Proceso terminado"
Screen.MousePointer = 0
End Sub

Private Sub mRepMD_Click()
End Sub


Private Sub mprocmanual_Click()
frmEjecucionProc2.Show
End Sub

Private Sub mReportes_Click()
SiActTProc = True
frmReportes.Show
End Sub

Private Sub mReproceso_Click()
Dim fecha As Date
Dim txtcadena As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim mata() As Variant
Dim coperacion As String
Dim noreg1 As Long
Dim noreg2 As Long
Dim noreg3 As Long
Dim noreg4 As Long
Dim noreg5 As Long
Dim noreg As Integer
Dim nocampos As Integer
Dim nomarch As String
Dim i As Long
Dim j As Long
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim txttabla1 As String
Dim txttabla2 As String
Dim resp As Integer
Dim tfecha As String
Dim fecha1 As Date
txttabla1 = DetermTablaSubproc(1)
txttabla2 = TablaProcesos1

Screen.MousePointer = 11
'1 fecha de registro
'2 clave de posicion
'3 clave de operacion
resp = MsgBox("Advertencia, esta accion restablecera todos los procesos de consolidación de datos. ¿Desea continuar?!", vbYesNo)
If resp = 6 Then
   tfecha = InputBox("Dame la fecha de reproceso", , Date)
   If IsDate(tfecha) Then
      fecha1 = CDate(tfecha)
      nomarch = ""
      frmCalVar.CommonDialog1.FileName = nomarch
      frmCalVar.CommonDialog1.ShowOpen
      nomarch = frmCalVar.CommonDialog1.FileName
      frmProgreso.Show
      If Not EsVariableVacia(nomarch) Then
      Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
      Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
      If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg = registros1.RecordCount
      registros1.MoveFirst
      nocampos = registros1.Fields.Count
      ReDim mata(1 To noreg, 1 To 3) As Variant
 For i = 1 To noreg
     For j = 1 To nocampos
         mata(i, j) = LeerTAccess(registros1, j - 1, i)
     Next j
     registros1.MoveNext
 Next i
 registros1.Close
 base1.Close
 End If

For i = 1 To noreg
    txtfecha = "TO_DATE('" & Format$(mata(i, 1), "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtcadena = "DELETE FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha & " AND PORTAFOLIO = 'TOTAL' AND ESC_FR = 'Normal' AND CPOSICION = " & mata(i, 2) & " AND COPERACION = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg1
    txtcadena = "DELETE FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = 'TOTAL' AND ESC_FACTORES = 'Normal' AND CPOSICION = " & mata(i, 2) & " AND COPERACION = '" & mata(i, 3) & "' and TIPOPOS = 1"
    ConAdo.Execute txtcadena, noreg2
    txtcadena = "DELETE FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = 'NEGOCIACION + INVERSION' AND ESC_FACTORES = 'Normal' AND CPOSICION = " & mata(i, 2) & " AND COPERACION = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg3
    txtcadena = "DELETE FROM " & TablaSensibN & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = 'NEGOCIACION + INVERSION' AND PORT_FR = 'Normal' AND ID_POSICION = " & mata(i, 2) & " AND ID_OPERACION = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg4
    txtcadena = "DELETE FROM " & TablaResEscEstres & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = 'NEGOCIACION + INVERSION' AND CPOSICION = " & mata(i, 2) & " AND COPERACION = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg5
    'valuacion x operacion
    txtcadena = "UPDATE " & txttabla1 & " SET BLOQUEADO = 'N', FINALIZADO = 'N', EXITO = 'N' WHERE ID_SUBPROCESO = 67 AND FECHAP = " & txtfecha & " AND PARAMETRO1 = 'TOTAL' AND PARAMETRO2 = 'Normal' AND PARAMETRO7 = '" & mata(i, 2) & "' AND PARAMETRO8 ='" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg1
    'calculo de pyg pot operacion 1
    txtcadena = "UPDATE " & txttabla1 & " SET BLOQUEADO = 'N', FINALIZADO = 'N', EXITO = 'N' WHERE ID_SUBPROCESO = 69 AND FECHAP = " & txtfecha & " AND PARAMETRO1 = 'TOTAL' AND PARAMETRO2 = 'Normal' AND PARAMETRO7 ='" & mata(i, 2) & "' AND PARAMETRO8 = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg2
    'calculo de pyg pot operacion 2
    txtcadena = "UPDATE " & txttabla1 & " SET BLOQUEADO = 'N', FINALIZADO = 'N', EXITO = 'N' WHERE ID_SUBPROCESO = 69 AND FECHAP = " & txtfecha & " AND PARAMETRO1 = 'NEGOCIACION + INVERSION' AND PARAMETRO2 = 'Normal' AND PARAMETRO7 ='" & mata(i, 2) & "' AND PARAMETRO8 = '" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg3
    'estres por operacion
    txtcadena = "UPDATE " & txttabla1 & " SET BLOQUEADO = 'N', FINALIZADO = 'N', EXITO ='N' WHERE ID_SUBPROCESO = 71 AND FECHAP = " & txtfecha & " AND PARAMETRO1 = 'TOTAL' AND PARAMETRO2 = 'Normal' AND PARAMETRO7 = " & mata(i, 2) & " AND PARAMETRO8 ='" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg4
    'sensibilidades por operacion
    txtcadena = "UPDATE " & txttabla1 & " SET BLOQUEADO = 'N', FINALIZADO = 'N', EXITO ='N' WHERE ID_SUBPROCESO = 73 AND FECHAP = " & txtfecha & " AND PARAMETRO1 = 'NEGOCIACION + INVERSION' AND PARAMETRO2 = 'Normal' AND PARAMETRO7 = " & mata(i, 2) & " AND PARAMETRO8 ='" & mata(i, 3) & "'"
    ConAdo.Execute txtcadena, noreg5
   Next i
   txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtcadena = "UPDATE " & txttabla2 & " SET FINALIZADO = 'N' WHERE (ID_TAREA = 27"
   txtcadena = txtcadena & " OR ID_TAREA = 28"
   txtcadena = txtcadena & " OR ID_TAREA = 29"
   txtcadena = txtcadena & " OR ID_TAREA = 30"
   txtcadena = txtcadena & " OR ID_TAREA = 31"
   txtcadena = txtcadena & " OR ID_TAREA = 33"
   txtcadena = txtcadena & " OR ID_TAREA = 34"
   txtcadena = txtcadena & " OR ID_TAREA = 35"
   txtcadena = txtcadena & " OR ID_TAREA = 36"
   txtcadena = txtcadena & " OR ID_TAREA = 38"
   txtcadena = txtcadena & " OR ID_TAREA = 40"
   txtcadena = txtcadena & " OR ID_TAREA = 41"
   txtcadena = txtcadena & " OR ID_TAREA = 42"
   txtcadena = txtcadena & " OR ID_TAREA = 43"
   txtcadena = txtcadena & " OR ID_TAREA = 44"
   txtcadena = txtcadena & " OR ID_TAREA = 45"
   txtcadena = txtcadena & " OR ID_TAREA = 46"
   txtcadena = txtcadena & ") AND FECHAP = " & txtfecha1
   ConAdo.Execute txtcadena
   Unload frmProgreso
   End If
   MsgBox "Fin de proceso"
   End If
End If
Screen.MousePointer = 0
End Sub

Private Sub mSalir_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
End 'finalizar la aplicacion
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mSensibilidades_Click()
End Sub


Private Sub mSimCVaR_Click()
frmSimCVaR.Show
End Sub

Private Sub mSimuladas_Click()
frmPosSimuladas.Show
End Sub

Private Sub msistema1_Click()
frmMensajes.Show
End Sub

Private Sub mSubprocesos_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
frmSubprocesos.Show
Else
  MsgBox "No tiene permisos para modificar los subprocesos"
End If
End Sub

Private Sub mValorR_Click()
frmValorReemplazo.Show 1
End Sub

Private Sub mValPosPrim_Click()
frmDetValPosPrim.Show
End Sub

Private Sub mValSwap_Click()
frmDetValSWap.Show
End Sub


Private Sub mVolatilidades_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

Screen.MousePointer = 11
If Not EsArrayVacio(MatFactRiesgo) Then
If IsArray(MatFactRiesgo) And UBound(MatFactRiesgo, 1) > 0 Then
 frmVolatilidades.Show
 frmCalVar.Hide
End If
Else
MsgBox "Este modulo solo funciona si se han cargado previamente datos"
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub



Private Sub mVReemplazo_Click()
If PerfilUsuario = "ADMINISTRADOR" Then
frmValorReemplazo.Show
Else
MsgBox "No tiene permisos para acceder a este modulo"
End If
End Sub

Private Sub Option15_Click()
 PrecioLimpio = 1
End Sub

Private Sub Option16_Click()
 PrecioLimpio = 0
End Sub


Private Sub Option34_Click()
ValExacta = False
End Sub

Private Sub Option36_Click()
ValExacta = True
End Sub

Private Sub Option39_Click()
ValEficiencia = True
End Sub

Private Sub Option4_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ImpResVaRMark(ByVal fecha As Date, ByVal txtport As String, ByRef mcov1() As Double, ByRef mcov2() As Double, ByRef mnomf() As Variant, ByRef MatSens() As Variant)
Dim n As Integer
Dim m As Integer
Dim anio As Integer
Dim mes As Integer
Dim dia As Integer
Dim nomarch As String
Dim i As Integer
Dim j As Integer
Dim exitoarch As Boolean
'se imprimen en un archivo de texto algunos resultados
'para la verificacion del var
If IsArray(mcov1) And IsArray(mnomf) Then
 n = UBound(mcov1, 1)
 m = UBound(mnomf, 2)
If n > 0 Then
'se construye el nombre del archivo
 anio = Format(fecha, "yy")
 mes = Format(fecha, "mm")
 dia = Format(fecha, "dd")
 nomarch = DirResVaR & "\sensibilidades\Sensibilidades " & txtport & Format(fecha, "yymmdd") & ".txt"
 Call VerificarSalidaArchivo(nomarch, 10, exitoarch)
 If exitoarch Then
 Print #10, "Covarianzas Mod normal"
 For i = 1 To n
 Print #10, mnomf(1, i) & Chr(9);
 Next i
 Print #10, ""
 For i = 1 To n
 For j = 1 To m
 Print #10, mcov1(i, j) & Chr(9);
 Next j
 Print #10, ""
 Next i
 Print #10, "Covarianzas Mod Exponencial"
 For i = 1 To n
 Print #10, mnomf(1, i) & Chr(9);
 Next i
 Print #10, ""
 For i = 1 To n
 For j = 1 To m
 Print #10, mcov2(i, j) & Chr(9);
 Next j
 Print #10, ""
 Next i

 Print #10, "Sensibilidades del portafolio"
 n = UBound(MatSens, 1)
 m = UBound(MatSens, 2)
 For i = 1 To m
 Print #10, mnomf(1, i) & Chr(9) & mnomf(2, i) & Chr(9) & mnomf(5, i) & Chr(9) & MatSens(1, i) / 1000000
 Next i
 Close #10
 End If
End If
End If
End Sub

Function AnexarTCupon(ByRef matpos() As propPosMD, ByRef matv() As Variant)
Dim mata() As New propPosMD
Dim noreg As Integer
Dim i As Integer
Dim txtcadena As String
Dim indice As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
mata = matpos
noreg = UBound(mata, 1)
For i = 1 To noreg
txtcadena = mata(i).cEmisionMD
If Not EsVariableVacia(txtcadena) Then
indice = 0
indice = BuscarValorArray(txtcadena, matv, 17)
If indice <> 0 Then
If mata(i).tCuponMD <> matv(indice, 15) Then
 mata(i).tCuponMD = matv(indice, 15)
End If
If mata(i).fVencMD <> matv(indice, 12) Then
 mata(i).fVencMD = matv(indice, 12)
End If
If mata(i).vNominalMD <> matv(indice, 11) Then
 mata(i).vNominalMD = matv(indice, 11)
End If
Else
 'MsgBox "Falta informacion en el vector de precios " & txtcadena
End If
End If
Next i
AnexarTCupon = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Private Sub Text10_KeyPress(KeyAscii As Integer)
 SiCargoFactR = False
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
SiCargoFactR = False
End Sub

Sub BajarArchivosPIP(ByVal fecha As Date, ByVal nomarch1 As String, ByVal nomarch2 As String, ByVal nomarch3 As String)
Dim sihayarch1 As String
Dim sihayarch2 As String
Dim sihayarch3 As String
Dim txtmsg As String
Dim sihay As Boolean
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 If Not SiExisteVecPrecios Then
    Call TransferirVect1(fecha, "PIP", "M", "XLS", DirVPreciosZ, DirTemp, DirVPrecios, "S", "S", "batchvp1.bat", txtmsg, sihay)
 End If
 sihayarch1 = VerifAccesoArch(DirVPrecios & "\" & nomarch1)
 If sihayarch1 And Not SiExisteVecPrecios Then
    SiExisteVecPrecios = True
 End If
 If Not SiExisteVecPreciosMD Then
    Call TransferirVect2(fecha, "VMD", "CSV", DirVPreciosZ, DirTemp, DirVPrecios, "S", "S", "batchvp2.bat", txtmsg, sihay)
 End If
 sihayarch2 = VerifAccesoArch(DirVPrecios & "\" & nomarch2)
 If sihayarch2 And Not SiExisteVecPreciosMD Then
    SiExisteVecPreciosMD = True
 End If
 If Not SiExisteCurvas Then
  Call ObtenerArchSFTP(fecha, "CURVAS", "", "XLS", DirCurvasZ, DirTemp, DirCurvas, "batchcurvas.bat", txtmsg, sihay)
 End If
 sihayarch3 = VerifAccesoArch(DirVPrecios & "\" & nomarch3)
If sihayarch3 And Not SiExisteCurvas Then
   SiExisteCurvas = True
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub MuestraBacktesting(ByRef matresback() As Variant, ByRef objeto As MSFlexGrid, ByRef matdback() As Variant)
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Integer
Dim simostrar As Boolean

 nocampos = UBound(matdback, 1)
 objeto.Rows = 2 * UBound(matdback, 3) + 1
 objeto.Cols = 2
 For i = 1 To UBound(matdback, 3)
  objeto.TextMatrix(2 * i, 0) = matresback(1, i)   'la fecha del backtesting
 Next i
 contar = 0
 For j = 1 To nocampos
 simostrar = False
 For i = 1 To UBound(matdback, 3)
 If matdback(j, 3, i) <> 0 Then
  simostrar = True
  contar = contar + 1
  objeto.Cols = contar + 1
  Exit For
 End If
 Next i
 If simostrar Then
  For i = 1 To UBound(matdback, 3)
   objeto.TextMatrix(2 * i - 1, contar) = Format(matdback(j, 2, i), "###,###,###,###,###,##0.00")
   objeto.TextMatrix(2 * i, contar) = Format(matdback(j, 7, i), "###,###,###,###,###,##0.00")
  Next i
 End If
 Next j

End Sub


Private Sub Option40_Click()
ValEficiencia = False
End Sub

Private Sub Timer1_Timer()
Dim uhora As Double
Dim tiempo As Double
If Not SiActTProc Then
   uhora = LeyendoEstadoUsuario(NomUsuario)
   tiempo = CDbl(Now) - uhora
   If tiempo * 24 * 60 > 10 Then
      MensajeProc = "La sesion ha estado inactiva mas de 10 minutos. Cerrando."
      Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
      Call DesbloquearUsuario(NomUsuario)
      MsgBox MensajeProc
      End
   End If
End If
End Sub


Private Sub Timer2_Timer()
If SiActTProc Then
  Call ActUHoraUsuario
End If
End Sub

