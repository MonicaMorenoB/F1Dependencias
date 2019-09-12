VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDetValPosPrim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valuación de posicion primaria"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar valuación a texto"
      Height          =   615
      Left            =   5370
      TabIndex        =   5
      Top             =   390
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular valuación"
      Height          =   615
      Left            =   2970
      TabIndex        =   4
      Top             =   360
      Width           =   1665
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5565
      Left            =   195
      TabIndex        =   3
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9816
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   2
      Top             =   1470
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Top             =   870
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   200
      TabIndex        =   0
      Top             =   270
      Width           =   2000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Detalle de la valuación"
      Height          =   195
      Left            =   200
      TabIndex        =   9
      Top             =   2130
      Width           =   1620
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de valuación"
      Height          =   195
      Left            =   200
      TabIndex        =   8
      Top             =   1230
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de los factores"
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   690
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Clave de operacion"
      Height          =   195
      Left            =   200
      TabIndex        =   6
      Top             =   90
      Width           =   1380
   End
End
Attribute VB_Name = "frmDetValPosPrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposswaps() As New propPosSwaps
Dim matpr() As New resValIns
Dim coperacion As String
Dim exito As Boolean
Dim f_pos As Date
Dim fechaval As Date
Dim fechaf As Date
Dim fechareg As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim j As Integer
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

coperacion = Text1.Text
fechaval = CDate(Combo1.Text)
f_pos = CDate(Combo1.Text)
fechaf = CDate(Combo2.Text)
txtfecha = "TO_DATE('" & Format(f_pos, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT MAX(FECHAREG) AS FECHAREG from " & TablaPosDeuda & " WHERE COPERACION = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND FECHAREG <= " & txtfecha & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   fechareg = rmesa.Fields("FECHAREG")
   rmesa.Close
   Screen.MousePointer = 11
   Call RutinaValOper(f_pos, fechaf, fechaval, matpos, matposmd, matposswaps, 1, fechareg, "Real", "000000", 7, coperacion, matpr, 1, txtmsg, exito)
   MSFlexGrid1.Rows = UBound(MatResValFlujo, 1) + 1
   MSFlexGrid1.Cols = 20
   MSFlexGrid1.TextMatrix(0, 0) = "Clave de operacion"
   MSFlexGrid1.TextMatrix(0, 1) = "Fecha de inicio de cupon"
   MSFlexGrid1.TextMatrix(0, 2) = "Fecha de vencimiento de cupon"
   MSFlexGrid1.TextMatrix(0, 3) = "Fecha de descuento de cupon"
   MSFlexGrid1.TextMatrix(0, 4) = "Dias inicio de cupon"
   MSFlexGrid1.TextMatrix(0, 5) = "Dias de vencimiento de cupon"
   MSFlexGrid1.TextMatrix(0, 6) = "Dias de descuento de cupon"
   MSFlexGrid1.TextMatrix(0, 7) = "Saldo del periodo"
   MSFlexGrid1.TextMatrix(0, 8) = "Amortizacion"
   MSFlexGrid1.TextMatrix(0, 9) = "Tasa del periodo"
   MSFlexGrid1.TextMatrix(0, 10) = "Sobretasa"
   MSFlexGrid1.TextMatrix(0, 11) = "Tasa cupon a aplicar"
   MSFlexGrid1.TextMatrix(0, 12) = "Intereses periodo anterior"
   MSFlexGrid1.TextMatrix(0, 13) = "Intereses Generados"
   MSFlexGrid1.TextMatrix(0, 14) = "Intereses pagados"
   MSFlexGrid1.TextMatrix(0, 15) = "Intereses sig periodo"
   MSFlexGrid1.TextMatrix(0, 16) = "pago total"
   MSFlexGrid1.TextMatrix(0, 17) = "tasa de descuento"
   MSFlexGrid1.TextMatrix(0, 18) = "factor de descuento"
   MSFlexGrid1.TextMatrix(0, 19) = "valor presente"
   
   For i = 1 To UBound(MatResValFlujo, 1)
          MSFlexGrid1.TextMatrix(i, 0) = MatResValFlujo(i).c_operacion
          MSFlexGrid1.TextMatrix(i, 1) = MatResValFlujo(i).fecha_ini
          MSFlexGrid1.TextMatrix(i, 2) = MatResValFlujo(i).fecha_fin
          MSFlexGrid1.TextMatrix(i, 3) = MatResValFlujo(i).fecha_desc
          MSFlexGrid1.TextMatrix(i, 4) = MatResValFlujo(i).dxv1
          MSFlexGrid1.TextMatrix(i, 5) = MatResValFlujo(i).dxv2
          MSFlexGrid1.TextMatrix(i, 6) = MatResValFlujo(i).dxv3
          MSFlexGrid1.TextMatrix(i, 7) = MatResValFlujo(i).saldo_periodo
          MSFlexGrid1.TextMatrix(i, 8) = MatResValFlujo(i).amortizacion
          MSFlexGrid1.TextMatrix(i, 9) = MatResValFlujo(i).tc_aplicar
          MSFlexGrid1.TextMatrix(i, 10) = MatResValFlujo(i).sobretasa
          MSFlexGrid1.TextMatrix(i, 11) = MatResValFlujo(i).tc_aplicar + MatResValFlujo(i).sobretasa
          MSFlexGrid1.TextMatrix(i, 12) = MatResValFlujo(i).int_acum_periodo - MatResValFlujo(i).int_gen_periodo
          MSFlexGrid1.TextMatrix(i, 13) = MatResValFlujo(i).int_gen_periodo
          MSFlexGrid1.TextMatrix(i, 14) = MatResValFlujo(i).int_pag_periodo
          MSFlexGrid1.TextMatrix(i, 15) = MatResValFlujo(i).int_acum_sig_periodo
          MSFlexGrid1.TextMatrix(i, 16) = MatResValFlujo(i).pago_total
          MSFlexGrid1.TextMatrix(i, 17) = MatResValFlujo(i).t_desc
          MSFlexGrid1.TextMatrix(i, 18) = MatResValFlujo(i).factor_desc
          MSFlexGrid1.TextMatrix(i, 19) = MatResValFlujo(i).valor_presente
   Next i
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim noreg As Integer
Dim nocols As Integer
Dim i As Integer
Dim j As Integer
Dim nomarch As String
Dim txtcadena As String
Dim coperacion As String
Dim fecha As Date
Dim exitoarch As Boolean

Screen.MousePointer = 11
noreg = MSFlexGrid1.Rows
nocols = MSFlexGrid1.Cols
coperacion = Text1.Text
fecha = CDate(Combo1.Text)
nomarch = DirResVaR & "\flujos operacion " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
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

Private Sub Form_Load()
Dim noreg As Integer
Dim i As Integer

noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
