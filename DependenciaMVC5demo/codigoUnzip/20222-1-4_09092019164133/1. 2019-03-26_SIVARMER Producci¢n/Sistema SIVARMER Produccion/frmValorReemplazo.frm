VERSION 5.00
Begin VB.Form frmValorReemplazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valor de reemplazo"
   ClientHeight    =   3135
   ClientLeft      =   -30
   ClientTop       =   360
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Subprocesos"
      Height          =   1065
      Left            =   2790
      TabIndex        =   7
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Subprocesos 2"
         Height          =   225
         Left            =   200
         TabIndex        =   9
         Top             =   690
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Subprocesos 1"
         Height          =   195
         Left            =   200
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar resultados VR"
      Height          =   700
      Left            =   2130
      TabIndex        =   6
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calcular costo de reemplazo"
      Height          =   700
      Left            =   3930
      TabIndex        =   5
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar subprocesos Valor Reemplazo"
      Height          =   700
      Left            =   180
      TabIndex        =   4
      Top             =   2400
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   200
      TabIndex        =   2
      Text            =   "30"
      Top             =   1110
      Width           =   1932
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   420
      Width           =   1932
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dias de prospección"
      Height          =   195
      Left            =   200
      TabIndex        =   3
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de la posicion"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   180
      Width           =   1500
   End
End
Attribute VB_Name = "frmValorReemplazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim nocontrap As Integer
Dim i As Integer
Dim mtm0 As Double
Dim mtm1 As Double
Dim cvar1 As Double
Dim cvar2 As Double
Dim noesc As Integer
Dim htiempo As Integer
Dim nconf As Double
Dim dtfecha As Date
Dim txtport As String
Dim txtsubport As String
Dim txtinserta As String
Dim txtfecha As String
Dim txtborra As String
Dim txtescfr As String
Dim exito As Boolean
SiActTProc = True
dtfecha = CDate(Combo2.Text)
noesc = 500
htiempo = 1
nconf = 0.97
txtport = "DERIVADOS"
txtescfr = "Normal"
Screen.MousePointer = 11
frmProgreso.Show
nocontrap = UBound(MatContrapartes, 1)
       ReDim matres(1 To nocontrap + 3, 1 To 8) As Variant
       For i = 1 To UBound(MatContrapartes, 1)
           txtsubport = "Deriv Contrap " & MatContrapartes(i, 1)
           matres(i, 1) = MatContrapartes(i, 1)                   'id contraparte
           matres(i, 2) = MatContrapartes(i, 3)                   'descripcion
           matres(i, 3) = MatContrapartes(i, 6)                   'sector
           Call GenPyGPortVR(dtfecha, txtport, txtescfr, txtsubport, noesc, htiempo, exito)
           Call CalcularCVaR_VR(dtfecha, txtport, txtescfr, txtsubport, noesc, htiempo, 1 - nconf, exito, mtm0, mtm1, cvar1)
           Call CalcularCVaR_VR(dtfecha, txtport, txtescfr, txtsubport, noesc, htiempo, nconf, exito, mtm0, mtm1, cvar2)
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
           Call GenPyGPortVR(dtfecha, txtport, txtescfr, matb(i), noesc, htiempo, exito)
           Call CalcularCVaR_VR(dtfecha, txtport, txtescfr, matb(i), noesc, htiempo, 1 - nconf, exito, mtm0, mtm1, cvar1)
           Call CalcularCVaR_VR(dtfecha, txtport, txtescfr, matb(i), noesc, htiempo, nconf, exito, mtm0, mtm1, cvar2)
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
           If matres(i, 4) <> 0 Then
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
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim dt_fecha As Date
Dim p_fwd As Integer
Dim exito As Boolean
Dim txtnompos As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtport As String
Dim txtportfr As String
Dim txtmsg As String
Dim id_tabla As Integer

dt_fecha = CDate(Combo2.Text)
p_fwd = Val(Text1.Text)
txtport = "DERIVADOS"
txtnompos = "Real"
txtportfr = "Normal"
noesc = 500
htiempo = 1
If Option1.value Then
   id_tabla = 1
Else
   id_tabla = 2
End If
Screen.MousePointer = 11
frmProgreso.Show
SiActTProc = True
Call GenSubCalcPyGOperVR(dt_fecha, txtport, txtnompos, txtportfr, noesc, htiempo, p_fwd, 91, id_tabla, txtmsg, exito)
SiActTProc = True
Unload frmProgreso
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub Command3_Click()
Dim fecha As Date
Dim tfondeo As Double

Dim noesc As Integer
Dim nconf As Double

Screen.MousePointer = 11
   fecha = CDate(Combo2.Text)
   
   tfondeo = InputBox("Dame la tasa de fondeo diario", , 0.05)
   tfondeo = Val(tfondeo) / 100
   noesc = 500
   nconf = 0.97
   Call CalcCostoReemplazo(fecha, tfondeo, noesc, nconf)
   MsgBox "Fin del proceso"
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim noreg As Long
Dim i As Long
noreg = UBound(MatFechasVaR, 1)
Combo2.Clear
For i = 1 To noreg
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
