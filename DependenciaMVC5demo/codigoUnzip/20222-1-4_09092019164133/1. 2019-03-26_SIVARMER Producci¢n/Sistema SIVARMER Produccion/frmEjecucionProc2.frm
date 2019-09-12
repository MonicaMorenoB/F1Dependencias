VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEjecucionProc2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecución de procesos adicionales"
   ClientHeight    =   9360
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8385
      Left            =   200
      TabIndex        =   0
      Top             =   810
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   14790
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Procesos de CVaR"
      TabPicture(0)   =   "frmEjecucionProc2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Factores de riesgo"
      TabPicture(1)   =   "frmEjecucionProc2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Posiciones"
      TabPicture(2)   =   "frmEjecucionProc2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Cálculo de CVA"
      TabPicture(3)   =   "frmEjecucionProc2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Simulaciones"
      TabPicture(4)   =   "frmEjecucionProc2.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame16"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Make Whole"
      TabPicture(5)   =   "frmEjecucionProc2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1605
         Left            =   -74460
         TabIndex        =   24
         Top             =   780
         Width           =   13755
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Height          =   375
            Left            =   6930
            TabIndex        =   28
            Top             =   810
            Width           =   1305
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   555
            Left            =   4290
            TabIndex        =   27
            Top             =   750
            Width           =   2085
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   2490
            TabIndex        =   26
            Text            =   "Text3"
            Top             =   690
            Width           =   1575
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   480
            TabIndex        =   25
            Text            =   "Combo2"
            Top             =   690
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "VaR Montecarlo"
         Height          =   2265
         Left            =   -74610
         TabIndex        =   17
         Top             =   2550
         Width           =   13995
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   555
            Left            =   7050
            TabIndex        =   23
            Top             =   750
            Width           =   1605
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   525
            Left            =   4410
            TabIndex        =   22
            Top             =   780
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2190
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   900
            Width           =   1875
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   390
            TabIndex        =   19
            Text            =   "Combo1"
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Porfafolio de posicion"
            Height          =   195
            Left            =   2190
            TabIndex        =   21
            Top             =   510
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   135
            Left            =   420
            TabIndex        =   18
            Top             =   510
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Otros"
         Height          =   1200
         Left            =   -74760
         TabIndex        =   12
         Top             =   4300
         Width           =   14000
         Begin VB.CommandButton Command17 
            Caption         =   "Generar script efectividad"
            Height          =   700
            Left            =   4380
            TabIndex        =   14
            Top             =   300
            Width           =   1500
         End
         Begin VB.CommandButton Command60 
            Caption         =   "Obtener flujos creditos del SIC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   1500
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Efectividad prospectiva para operacion simulada"
         Height          =   1200
         Left            =   200
         TabIndex        =   6
         Top             =   5600
         Width           =   14000
         Begin VB.CommandButton Command21 
            Caption         =   "Calcular ef prospectiva"
            Height          =   555
            Left            =   300
            TabIndex        =   9
            Top             =   360
            Width           =   1905
         End
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   480
            Width           =   1665
         End
         Begin VB.TextBox Text27 
            Height          =   315
            Left            =   4440
            TabIndex        =   7
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   2580
            TabIndex        =   11
            Top             =   270
            Width           =   450
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de la posicion simulada"
            Height          =   195
            Left            =   4470
            TabIndex        =   10
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Flujos"
         Height          =   1200
         Left            =   -74800
         TabIndex        =   3
         Top             =   1700
         Width           =   14000
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3810
            TabIndex        =   5
            Text            =   "S:\FlujosPIP"
            Top             =   630
            Width           =   5205
         End
         Begin VB.CommandButton Command24 
            Caption         =   "Obtener flujos de archivos PIP"
            Height          =   645
            Left            =   2160
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Flujos de Swaps"
         Height          =   1200
         Left            =   -74800
         TabIndex        =   1
         Top             =   400
         Width           =   14000
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   330
            TabIndex        =   15
            Top             =   540
            Width           =   1905
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Generacion de calculos de forwards en archivo txt"
            Height          =   615
            Left            =   5430
            TabIndex        =   2
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   390
            TabIndex        =   16
            Top             =   300
            Width           =   450
         End
      End
   End
   Begin VB.Menu mfuncion1 
      Caption         =   "Funciones 1"
      Begin VB.Menu LeerPlusMinus 
         Caption         =   "Lectura de Plus-Minus"
      End
      Begin VB.Menu mSQLTC 
         Caption         =   "Extracción del catalogo de cadenas SQL"
      End
   End
End
Attribute VB_Name = "frmEjecucionProc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Sub algo()
Dim tfecha As String
Dim fecha As Date
Dim txtfecha As String
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
tfecha = InputBox("Fecha de los calculos", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT COUNT(*) from " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha, "YYYYMMDD")
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
If noreg <> 0 Then
  txtfiltro = "SELECT * from " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha, "YYYYMMDD")
 rmesa.Open txtfiltro, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 
 rmesa.MoveNext
 Next i
 rmesa.Close
 
 Open DirResVaR & "\Script VaR.sql" For Output As #1
 txtcadena = "DELETE FROM " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha, "yyyymmdd")
 Print #1, txtcadena
 For i = 1 To noreg
     txtcadena = "INSERT INTO " & TablaVaRIKOS & " VALUES("
     txtcadena = txtcadena & mata(i, 1) & ","               'FECHA DE OPERACION
     txtcadena = txtcadena & mata(i, 2) & ","               'Clave de operación
     txtcadena = txtcadena & mata(i, 3) & ","               'VAR GLOBAL
     txtcadena = txtcadena & mata(i, 4) & ","               'VAR DEL PORTAFOLIO
     txtcadena = txtcadena & mata(i, 5) & ","               'VAR ESTRUCTURAL
     txtcadena = txtcadena & mata(i, 6) & ","               'LIM VAR GLOBAL
     txtcadena = txtcadena & mata(i, 7) & ","               'LIM VAR PORTAFOLIO
     txtcadena = txtcadena & mata(i, 8) & ");"              'LIM VAR ESTRUCTURAL
  Print #1, txtcadena
  Call MostrarMensajeSistema("Copiando VAR del dia " & Format(fecha, "dd/mm/yyyy") & " " & Format(AvanceProc, "##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
  DoEvents
 Next i
  Print #1, "COMMIT;"
 Close #1
MsgBox "se insertaron " & noreg & " registros. Proceso terminado"
End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Command12_Click()
Dim mattiie28() As Variant
Dim mattiie91() As Variant
Dim noreg As Long
Dim i As Long
Dim txtfecha1 As String
Dim fecha1 As Date
Dim finicioem As Date
Dim pc As Integer
Dim concepto As String
Dim cemision As String
cemision = ""
'se carga toda la historia disponible de tiie a 28 dias
mattiie28 = Leer1FactorRC("TIIE28 PIP", 0)
mattiie91 = Leer1FactorRC("TIIE91 PIP", 0)
fecha1 = #1/1/2008#
Screen.MousePointer = 11
frmProgreso.Show
    concepto = ""
    st = 0
    finicioem = #1/1/2008#
       Call ExtrapolaYieldEmTIIE(cemision, concepto, st, fecha1, finicioem, finicioem, mattiie28)
Unload frmProgreso
 MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub


Sub codigohuerfano1()
Dim fecha As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim txtfechac1 As String
Dim txtfechac2 As String
Dim txtfechac3 As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset


If PerfilUsuario = "ADMINISTRADOR" Then
Screen.MousePointer = 11
txtfecha1 = InputBox("La fecha a exportar ", , Date)
fecha = CDate(txtfecha1)
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro = "SELECT * FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
   rmesa.Open txtfiltro, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   rmesa.MoveFirst
 For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 rmesa.MoveNext
 Next i
 rmesa.Close
 Call InConexOracle(NomServQA, conAdo1)
 conAdo1.Execute "DELETE FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
 For i = 1 To noreg
  txtfechac1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfechac2 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfechac3 = "to_date('" & Format(mata(i, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
  txtcadena = txtcadena & txtfechac1 & ","                'FECHA
  txtcadena = txtcadena & txtfechac2 & ","                'FECHA
  txtcadena = txtcadena & txtfechac3 & ","                'FECHA
  txtcadena = txtcadena & "'" & mata(i, 4) & "',"         'Clave de operación
  For j = 5 To 16
  If Len(Trim(mata(i, j))) <> 0 Then
  txtcadena = txtcadena & mata(i, j) & ","
  Else
  txtcadena = txtcadena & "null,"
  End If
  Next j
  txtcadena = txtcadena & mata(i, 17) & ","
  txtcadena = txtcadena & mata(i, 18) & ")"               'FECHA
' MsgBox txtcadena
 conAdo1.Execute txtcadena
 Next i
conAdo1.Close
End If
MsgBox "Proceso terminado " & noreg & " registros"
Screen.MousePointer = 0
Else
MsgBox "No tiene acceso a este modulo"
End If


End Sub

Private Sub Command1_Click()
Dim fecha As Date
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtmsg As String
Dim exito As Boolean

fecha = Combo1.Text
txtport = Text1.Text
noesc = 500
htiempo = 1
Screen.MousePointer = 11
frmProgreso.Show
Call CalculoMatCholeski(fecha, txtport, "Normal", noesc, htiempo, txtmsg, exito)
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command17_Click()
Dim tfecha As String
Dim fecha As Date
Dim txtfiltro As String
Dim txtfecha As String
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim matb() As Variant
Dim txtarch As String
Dim txtcadena As String
Dim txtfechac1 As String
Dim txtfechac2 As String
Dim txtfechac3 As String
Dim fecha1 As Date
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
tfecha = InputBox("Dame la fecha de los calculos", , Date)

If IsDate(tfecha) Then
 fecha = CDate(tfecha)
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "SELECT COUNT(*) from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
 rmesa.Open txtfiltro, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
If noreg <> 0 Then
  txtfiltro = "SELECT * from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
 rmesa.Open txtfiltro, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos + 1) As Variant
 For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 mata(i, nocampos + 1) = CLng(mata(i, 1)) & " " & mata(i, 4)
 If mata(i, 17) < 80 Or mata(i, 17) > 125 Then
    MsgBox "Alerta. Una efectividad no es 80-125 emision " & mata(i, 4)
 End If
 rmesa.MoveNext
 Next i
 rmesa.Close
 
 matb = RutinaOrden(mata, nocampos + 1, 1)
 
 For i = 2 To UBound(matb, 1)
     If matb(i - 1, nocampos + 1) = matb(i, nocampos + 1) Then
        MsgBox "se repite la efectividad de la operación " & matb(i, 1) & "   " & matb(i, 4)
        Exit Sub
     End If
 Next i
 txtarch = DirResVaR & "\script efec.sql"
 Open txtarch For Output As #1
 txtcadena = "DELETE FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
 Print #1, txtcadena
 For i = 1 To noreg
     txtfechac1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfechac2 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfechac3 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
     txtcadena = txtcadena & txtfechac1 & ","                'FECHA
     txtcadena = txtcadena & txtfechac2 & ","                'FECHA
     txtcadena = txtcadena & txtfechac3 & ","                'FECHA
     txtcadena = txtcadena & "'" & mata(i, 4) & "',"         'Clave de operación
     For j = 5 To 16
     If Len(Trim(mata(i, j))) <> 0 Then
        txtcadena = txtcadena & mata(i, j) & ","
     Else
        txtcadena = txtcadena & "null,"
     End If
  Next j
  txtcadena = txtcadena & mata(i, 17) & ","
  txtcadena = txtcadena & mata(i, 18) & ");"               'FECHA
  Print #1, txtcadena
  AvanceProc = i / noreg
  MensajeProc = "Copiando efectividad del dia " & Format(fecha1, "dd/mm/yyyy") & " " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
  Print #1, "COMMIT;"
 Close #1
MsgBox "Se genero el archivo " & txtarch
End If
End If
Screen.MousePointer = 0

End Sub


Private Sub Command2_Click()
Dim fecha As Date
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim nosim As Long
Dim txtmsg As String
Dim exito As Boolean
Screen.MousePointer = 11
fecha = Combo1.Text
txtport = Text1.Text
noesc = 500
htiempo = 1
nosim = 10000
Call GenSubprocPyGMontOper(fecha, txtport, "Normal", noesc, htiempo, nosim, 76, 1, txtmsg, exito)
Screen.MousePointer = 0
End Sub

Private Sub Command21_Click()
Dim fecha As Date
Dim txtport As String
Dim txtmsg As String
Dim bl_exito2 As Boolean
fecha = CDate(Text26.Text)
txtport = Text27.Text
Screen.MousePointer = 11
frmProgreso.Show
SiActTProc = True
  Call CEficProsSwapsPort(fecha, txtport, txtmsg, bl_exito2)
SiActTProc = False
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command28_Click()
Dim nomarch As String
 frmEjecucionProc2.CommonDialog1.ShowOpen
 nomarch = frmEjecucionProc2.CommonDialog1.FileName

End Sub




Private Sub Command3_Click()
Dim fecha As Date
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtmsg As String
Dim exito As Boolean
Dim nconf As Double
noesc = 500
htiempo = 1
nconf = 0.97
fecha = Combo2.Text
txtport = Text3.Text

Screen.MousePointer = 11
frmProgreso.Show
Call CalcSensibPort2(fecha, "NEGOCIACION + INVERSION", "Normal", txtport, txtmsg, exito)
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command33_Click()
Dim matposfwd() As New propPosFwd
Dim fecha As Date
Dim txtmsg As String
Dim exito As Boolean
Dim txtfecha As String
txtfecha = InputBox("Dame la fecha", , Date)
SiActTProc = True
fecha = CDate(txtfecha)
Call GenerarDatosFwds(fecha, MatPosRiesgo, matposfwd, txtmsg, exito)
Call ActUHoraUsuario
SiActTProc = False
MsgBox "Fin de proceso"
End Sub


Private Sub Command38_Click()
Dim txtfecha As String
Dim dtfecha As Date
txtfecha = InputBox("Dame la fecha de calculo", "", Date)
If IsDate(txtfecha) Then
   SiActTProc = True
   Screen.MousePointer = 11
      dtfecha = CDate(txtfecha)
      Call CalcEPE(dtfecha)
   Screen.MousePointer = 0
   Call ActUHoraUsuario
   SiActTProc = False
End If
MsgBox "Fin de proceso"
End Sub

Private Sub Command58_Click()
Dim nomarch As String
 frmEjecucionProc2.CommonDialog1.ShowOpen
 nomarch = frmEjecucionProc2.CommonDialog1.FileName
' Text20.Text = nomarch
End Sub



Private Sub Command8_Click()
Dim fecha As Date
Dim txtportfr As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matpr() As New resValIns
Dim matport1() As String
Dim matport2() As String
Dim txtnompos As String
Dim txtport2 As String
Dim txttvar As String
Dim txtmsg As String
Dim noesc As Long
Dim htiempo As Long
Dim i As Long
Dim nconf As Double
Dim valor As Double
Dim exito As Boolean


'fecha = CDate(Text18.Text)
'txtnompos = Text19.Text
'txtportfr = Text21.Text
txttvar = "CVARH"
'noesc = Val(Text22.Text)
'htiempo = Val(Text23.Text)
'nconf = 0.97


Screen.MousePointer = 11
SiActTProc = True
ValExacta = True
frmProgreso.Show
Call ActTCFlujosSwaps(fecha, 2, "N", txtmsg, exito)
Call RutinaValPos(fecha, fecha, fecha, matpos, txtnompos, 1, matpr, txtmsg, exito)
If exito Then
   Call GuardarResValPos(fecha, fecha, fecha, txtnompos, txtportfr, matpos, matposmd, matpr, 1, exito)
   Call SubprocCalculoPyGPos(fecha, fecha, fecha, txtnompos, txtportfr, noesc, htiempo, txtmsg, exito)
   Call GenVaRMark2(fecha, txtnompos, txtportfr, matport1, noesc, htiempo, nconf, txtmsg, exito)
   Call LeerResultadosSimPos(fecha, txtnompos, noesc, htiempo)
End If

Unload frmProgreso
Call ActUHoraUsuario
SiActTProc = False
ValExacta = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub


Private Sub Command4_Click()
Dim fecha As Date
Dim txtport As String
Dim noesc As Long
Dim htiempo As Integer
Dim nconf As Double
Dim txtmsg As String
Dim exito As Boolean

Screen.MousePointer = 11
Call GenVaRMark(fecha, txtport, "Normal", "GRUPO PRINCIPAL", noesc, htiempo, nconf, txtmsg, exito)
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

