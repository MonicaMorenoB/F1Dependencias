VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDetValSWap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Características del swap"
   ClientHeight    =   11325
   ClientLeft      =   -30
   ClientTop       =   330
   ClientWidth     =   13680
   Icon            =   "frmDetValSWap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11325
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1830
      TabIndex        =   16
      Top             =   1020
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir flujos"
      Height          =   600
      Left            =   6150
      TabIndex        =   15
      Top             =   300
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Valuar operacion"
      Height          =   600
      Left            =   4290
      TabIndex        =   12
      Top             =   300
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1830
      TabIndex        =   11
      Top             =   600
      Width           =   1965
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   180
      Width           =   1965
   End
   Begin VB.Frame Frame7 
      Caption         =   "Valuacion de Swap"
      Height          =   9540
      Left            =   240
      TabIndex        =   0
      Top             =   1530
      Width           =   12750
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5460
         TabIndex        =   18
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox Text16 
         Height          =   288
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2172
      End
      Begin VB.TextBox Text18 
         Height          =   300
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   2268
      End
      Begin VB.TextBox Text20 
         Height          =   324
         Left            =   5550
         TabIndex        =   1
         Top             =   1290
         Width           =   2220
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7245
         Left            =   210
         TabIndex        =   4
         Top             =   2010
         Width           =   12330
         _ExtentX        =   21749
         _ExtentY        =   12779
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de la valuación"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   1650
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor pata pasiva"
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de inicio"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final"
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Valor de la Pata activa"
         Height          =   195
         Left            =   5460
         TabIndex        =   5
         Top             =   210
         Width           =   1605
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de factores"
      Height          =   195
      Left            =   300
      TabIndex        =   17
      Top             =   1080
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de valuacion"
      Height          =   195
      Left            =   330
      TabIndex        =   14
      Top             =   690
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de operacion"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   270
      Width           =   1380
   End
End
Attribute VB_Name = "frmDetValSWap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command2_Click()
Dim noreg As Integer
Dim nocols As Integer
Dim i As Integer
Dim j As Integer
Dim nomarch As String
Dim txtcadena As String
Dim coperacion As Integer
Dim fecha As Date
Dim exitoarch As Boolean

Screen.MousePointer = 11
noreg = MSFlexGrid1.Rows
nocols = MSFlexGrid1.Cols
coperacion = Val(Text1.Text)
fecha = CDate(Combo1.Text)
nomarch = DirResVaR & "\flujos swap " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
For i = 1 To noreg
txtcadena = ""
For j = 1 To nocols
txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i - 1, j - 1) & Chr(9)
Next j
Print #1, txtcadena
Next i
Close #1
End If
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim coperacion As Integer
Dim exito As Boolean
Dim f_pos As Date
Dim fechaval As Date
Dim fechaf As Date
Dim fechareg As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposswaps() As New propPosSwaps
Dim matpr() As New resValIns
Dim i As Long
Dim noreg As Long
Dim j As Integer
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

coperacion = Val(Text1.Text)
fechaval = CDate(Combo1.Text)
f_pos = CDate(Combo1.Text)
fechaf = CDate(Combo2.Text)
txtfecha = "TO_DATE('" & Format(f_pos, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT MAX(FECHAREG) AS FECHAREG from " & TablaPosSwaps & " WHERE COPERACION = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND FECHAREG <= " & txtfecha & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   fechareg = ReemplazaVacioValor(rmesa.Fields("FECHAREG"), 0)
   rmesa.Close
   If fechareg <> 0 Then
   Screen.MousePointer = 11
   Call RutinaValOper(f_pos, fechaf, fechaval, matpos, matposmd, matposswaps, 1, fechareg, "Real", "000000", 4, coperacion, matpr, 1, txtmsg, exito)
   MSFlexGrid1.Rows = UBound(MatResValFlujo, 1) + 1
   MSFlexGrid1.Cols = 21
   MSFlexGrid1.TextMatrix(0, 0) = "Clave de operacion"
   MSFlexGrid1.TextMatrix(0, 1) = "Pata"
   MSFlexGrid1.TextMatrix(0, 2) = "Fecha de inicio de cupon"
   MSFlexGrid1.TextMatrix(0, 3) = "Fecha de vencimiento de cupon"
   MSFlexGrid1.TextMatrix(0, 4) = "Fecha de descuento de cupon"
   MSFlexGrid1.TextMatrix(0, 5) = "Dias inicio de cupon"
   MSFlexGrid1.TextMatrix(0, 6) = "Dias de vencimiento de cupon"
   MSFlexGrid1.TextMatrix(0, 7) = "Dias de descuento de cupon"
   MSFlexGrid1.TextMatrix(0, 8) = "Saldo del periodo"
   MSFlexGrid1.TextMatrix(0, 9) = "Amortizacion"
   MSFlexGrid1.TextMatrix(0, 10) = "Tasa del periodo"
   MSFlexGrid1.TextMatrix(0, 11) = "Sobretasa"
   MSFlexGrid1.TextMatrix(0, 12) = "Tasa cupon a aplicar"
   MSFlexGrid1.TextMatrix(0, 13) = "Intereses periodo anterior"
   MSFlexGrid1.TextMatrix(0, 14) = "Intereses Generados"
   MSFlexGrid1.TextMatrix(0, 15) = "Intereses pagados"
   MSFlexGrid1.TextMatrix(0, 16) = "Intereses sig periodo"
   MSFlexGrid1.TextMatrix(0, 17) = "pago total"
   MSFlexGrid1.TextMatrix(0, 18) = "tasa de descuento"
   MSFlexGrid1.TextMatrix(0, 19) = "factor de descuento"
   MSFlexGrid1.TextMatrix(0, 20) = "valor presente"
   
   For i = 1 To UBound(MatResValFlujo, 1)
          MSFlexGrid1.TextMatrix(i, 0) = MatResValFlujo(i).c_operacion
          MSFlexGrid1.TextMatrix(i, 1) = MatResValFlujo(i).t_pata
          MSFlexGrid1.TextMatrix(i, 2) = MatResValFlujo(i).fecha_ini
          MSFlexGrid1.TextMatrix(i, 3) = MatResValFlujo(i).fecha_fin
          MSFlexGrid1.TextMatrix(i, 4) = MatResValFlujo(i).fecha_desc
          MSFlexGrid1.TextMatrix(i, 5) = MatResValFlujo(i).dxv1
          MSFlexGrid1.TextMatrix(i, 6) = MatResValFlujo(i).dxv2
          MSFlexGrid1.TextMatrix(i, 7) = MatResValFlujo(i).dxv3
          MSFlexGrid1.TextMatrix(i, 8) = MatResValFlujo(i).saldo_periodo
          MSFlexGrid1.TextMatrix(i, 9) = MatResValFlujo(i).amortizacion
          MSFlexGrid1.TextMatrix(i, 10) = MatResValFlujo(i).tc_aplicar
          MSFlexGrid1.TextMatrix(i, 11) = MatResValFlujo(i).sobretasa
          MSFlexGrid1.TextMatrix(i, 12) = MatResValFlujo(i).tc_aplicar + MatResValFlujo(i).sobretasa
          MSFlexGrid1.TextMatrix(i, 13) = MatResValFlujo(i).int_acum_periodo - MatResValFlujo(i).int_gen_periodo
          MSFlexGrid1.TextMatrix(i, 14) = MatResValFlujo(i).int_gen_periodo
          MSFlexGrid1.TextMatrix(i, 15) = MatResValFlujo(i).int_pag_periodo
          MSFlexGrid1.TextMatrix(i, 16) = MatResValFlujo(i).int_acum_sig_periodo
          MSFlexGrid1.TextMatrix(i, 17) = MatResValFlujo(i).pago_total
          MSFlexGrid1.TextMatrix(i, 18) = MatResValFlujo(i).t_desc
          MSFlexGrid1.TextMatrix(i, 19) = MatResValFlujo(i).factor_desc
          MSFlexGrid1.TextMatrix(i, 20) = MatResValFlujo(i).valor_presente
   Next i
   End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
MSFlexGrid1.Cols = 8
MSFlexGrid1.Rows = 2
MSFlexGrid1.TextMatrix(0, 0) = "No Periodo"
MSFlexGrid1.TextMatrix(0, 1) = "Fecha inicio"
MSFlexGrid1.TextMatrix(0, 2) = "fecha final"
MSFlexGrid1.TextMatrix(0, 3) = "tasa aplica periodo"
MSFlexGrid1.TextMatrix(0, 4) = "Monto deuda"
MSFlexGrid1.TextMatrix(0, 5) = "Pago de capital"
MSFlexGrid1.TextMatrix(0, 6) = "Intereses Generados"
MSFlexGrid1.TextMatrix(0, 7) = "Pago Total"
noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
Combo1.ListIndex = 1
Combo2.ListIndex = 1

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmCalVar.Visible = True
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

