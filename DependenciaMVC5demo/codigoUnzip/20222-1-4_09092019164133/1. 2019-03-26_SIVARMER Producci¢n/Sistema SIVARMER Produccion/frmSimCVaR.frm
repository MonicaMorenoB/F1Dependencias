VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSimCVaR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulaciones de CVaR"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Estres de factores"
      Height          =   3000
      Left            =   200
      TabIndex        =   14
      Top             =   4860
      Width           =   11500
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7530
         TabIndex        =   19
         Top             =   690
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5070
         TabIndex        =   18
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por fecha especifica"
         Height          =   375
         Left            =   7500
         TabIndex        =   17
         Top             =   210
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por escenario simulado"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   2565
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2295
         Left            =   360
         TabIndex        =   15
         Top             =   660
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Posicion simulada"
      Height          =   3000
      Left            =   200
      TabIndex        =   11
      Top             =   1560
      Width           =   11500
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   9000
         TabIndex        =   13
         Top             =   300
         Width           =   2025
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2505
         Left            =   200
         TabIndex        =   12
         Top             =   300
         Width           =   8000
         _ExtentX        =   14129
         _ExtentY        =   4419
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ejecucion de procesos"
      Height          =   1305
      Left            =   300
      TabIndex        =   7
      Top             =   8370
      Width           =   8445
      Begin VB.CommandButton Command3 
         Caption         =   "Generar CVaR por subportafolio"
         Height          =   645
         Left            =   6390
         TabIndex        =   23
         Top             =   300
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Generar subprocesos pyg por subport"
         Height          =   555
         Left            =   4470
         TabIndex        =   22
         Top             =   420
         Width           =   1545
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Generar subproc de gen de p y g"
         Height          =   645
         Left            =   2520
         TabIndex        =   9
         Top             =   390
         Width           =   1665
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar subprocesos valuacion"
         Height          =   585
         Left            =   300
         TabIndex        =   8
         Top             =   390
         Width           =   1965
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "CVaR para posicion simulada"
      Height          =   1200
      Left            =   200
      TabIndex        =   0
      Top             =   240
      Width           =   11500
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2970
         TabIndex        =   24
         Top             =   570
         Width           =   2025
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   8460
         TabIndex        =   20
         Text            =   ".03"
         Top             =   600
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   330
         TabIndex        =   10
         Top             =   630
         Width           =   2115
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   7770
         TabIndex        =   2
         Text            =   "1"
         Top             =   600
         Width           =   585
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   5970
         TabIndex        =   1
         Text            =   "500"
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Percentil"
         Height          =   195
         Left            =   8670
         TabIndex        =   21
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de calculo"
         Height          =   165
         Left            =   2970
         TabIndex        =   6
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "H de tiempo"
         Height          =   195
         Left            =   7650
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "No escenarios"
         Height          =   195
         Left            =   5910
         TabIndex        =   4
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Posición"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmSimCVaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfiltro As String
Dim txtfecha As String
Dim noreg As Integer
Dim txtport As String
Dim txtportfr As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim rmesa As New ADODB.recordset

SiActTProc = True
'fecha1 = CDate(Combo1.Text)
'fecha2 = CDate(Combo2.Text)

Screen.MousePointer = 11
frmProgreso.Show
fecha = fecha1
txtport = Text1.Text
txtportfr = Text3.Text
Do While fecha <= fecha2
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT COUNT(*) FROM " & TablaPortPosicion & "  WHERE"
   txtfiltro = txtfiltro & " FECHA = " & txtfecha
   txtfiltro = txtfiltro & " AND PORTAFOLIO = '" & txtport & "'"
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      Call GenSubpValPosicion(121, fecha, txtport, txtportfr, 1, 1)
   End If
   fecha = fecha + 1
Loop
Unload frmProgreso
Screen.MousePointer = 0
Call ActUHoraUsuario
SiActTProc = False
MsgBox "Fin de proceso"

End Sub

Private Sub Command8_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim fecha As Date
Dim indice As Integer
Dim txtportfr As String
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer

Screen.MousePointer = 11
'fecha1 = CDate(Combo1.Text)
'fecha2 = CDate(Combo2.Text)
txtport = Text10.Text
txtportfr = Text14.Text
noesc = Val(Text11.Text)
htiempo = Val(Text12.Text)
fecha = fecha1
Do While fecha <= fecha2
   indice = BuscarValorArray(fecha, MatFechasVaR, 1)
   If indice <> 0 Then
      Call GenProcCVaR(fecha, 122, txtport, txtportfr, noesc, htiempo, 1)
   End If
   fecha = fecha + 1
Loop
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

txtfiltro2 = "SELECT NOMPOS FROM (SELECT NOMPOS FROM " & TablaPosMD & " WHERE TIPOPOS = 2 GROUP BY NOMPOS"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT NOMPOS FROM " & TablaPosDiv & " WHERE TIPOPOS = 2 GROUP BY NOMPOS"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT NOMPOS FROM " & TablaPosSwaps & " WHERE TIPOPOS = 2 GROUP BY NOMPOS"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT NOMPOS FROM " & TablaPosFwd & " WHERE TIPOPOS =2 GROUP BY NOMPOS) GROUP BY NOMPOS ORDER BY NOMPOS"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Rows = noreg + 1
   For i = 1 To noreg
   MSFlexGrid1.TextMatrix(i, 1) = rmesa.Fields("NOMPOS")
   rmesa.MoveNext
   Next i
   rmesa.Close
End If
Combo1.Clear
Combo3.Clear
noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo3.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
