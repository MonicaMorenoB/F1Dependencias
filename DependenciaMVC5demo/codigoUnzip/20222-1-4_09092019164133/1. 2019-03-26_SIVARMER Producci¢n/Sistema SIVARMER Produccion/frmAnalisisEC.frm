VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResultadosER 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eficiencia retrospectiva"
   ClientHeight    =   11805
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   15270
   Icon            =   "frmAnalisisEC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11805
   ScaleWidth      =   15270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   11340
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   820
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   21749
            Text            =   ""
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultados"
      Height          =   10905
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   14775
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   2100
         TabIndex        =   2
         Top             =   540
         Width           =   1740
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   1692
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   8760
         Left            =   210
         TabIndex        =   3
         Top             =   1650
         Width           =   14040
         _ExtentX        =   24765
         _ExtentY        =   15452
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Resultados retrospectiva"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al periodo"
         Height          =   195
         Left            =   2100
         TabIndex        =   5
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del periodo"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   48
      Top             =   -360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmResultadosER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim indice As Integer
Dim noreg As Integer
Dim i As Integer


Screen.MousePointer = 11
MSFlexGrid2.Rows = 2
MSFlexGrid2.Cols = 2
indice = 1
noreg = UBound(MatEficFwds, 1)
MSFlexGrid2.Rows = noreg + 1
MSFlexGrid2.Cols = 17
MSFlexGrid2.TextMatrix(0, 0) = "Swap"
MSFlexGrid2.TextMatrix(0, 1) = "Orden"
MSFlexGrid2.TextMatrix(0, 2) = "Fecha de valuacion"
MSFlexGrid2.TextMatrix(0, 3) = "Tipo de cambio t0"
MSFlexGrid2.TextMatrix(0, 4) = "Tipo cambio pactado"
MSFlexGrid2.TextMatrix(0, 5) = "Val fwd inicio"
MSFlexGrid2.TextMatrix(0, 6) = "Val fwd "
MSFlexGrid2.TextMatrix(0, 7) = "puntos fwd"
MSFlexGrid2.TextMatrix(0, 8) = "Val Fwd teorico"
MSFlexGrid2.TextMatrix(0, 9) = "Eficiencia fwd teorico"
MSFlexGrid2.TextMatrix(0, 10) = "Devengo"
MSFlexGrid2.TextMatrix(0, 11) = "Val Fwd teorico-devengo"
MSFlexGrid2.TextMatrix(0, 12) = "Eficiencia 2"
MSFlexGrid2.TextMatrix(0, 13) = "Fwd teorico * curva"
MSFlexGrid2.TextMatrix(0, 14) = "Fwd negociado * curva"
MSFlexGrid2.TextMatrix(0, 15) = "Eficiencia fwd hipotetico"

For i = 1 To noreg

MSFlexGrid2.TextMatrix(i, 0) = MatEficFwds(i, 1, indice)
MSFlexGrid2.TextMatrix(i, 1) = Format(MatEficFwds(i, 2, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 2) = MatEficFwds(i, 3, indice)
MSFlexGrid2.TextMatrix(i, 3) = Format(MatEficFwds(i, 4, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 4) = Format(MatEficFwds(i, 5, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 5) = Format(MatEficFwds(i, 6, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 6) = Format(MatEficFwds(i, 7, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 7) = Format(MatEficFwds(i, 8, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 8) = Format(MatEficFwds(i, 9, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 9) = Format(MatEficFwds(i, 10, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 10) = Format(MatEficFwds(i, 11, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 11) = Format(MatEficFwds(i, 12, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 12) = Format(MatEficFwds(i, 13, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 13) = Format(MatEficFwds(i, 14, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 14) = Format(MatEficFwds(i, 15, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 15) = Format(MatEficFwds(i, 16, indice), "###,###,###,###,###,##0.00000")
MSFlexGrid2.TextMatrix(i, 16) = Format(MatEficFwds(i, 17, indice), "###,###,###,###,###,##0.00000")
'MSFlexGrid2.TextMatrix(i, 17) = Format(MatEficFwds(i, 18, indice), "###,###,###,###,###,##0.00000")
'MSFlexGrid2.TextMatrix(i, 18) = Format(MatEficFwds(i, 19, indice), "###,###,###,###,###,##0.00000")
Next i
Screen.MousePointer = 0
End Sub



Private Sub Command3_Click()
Dim eficpros As Double
Dim i As Integer
Dim noreg As Integer
Dim matesc() As Variant
Dim mata() As Variant

Screen.MousePointer = 11
Dim fecha As Date
  Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
For i = 1 To 10
MsgBox "requiere antencion urgente"
 'mata = CrearMatEscenariosO(matfechasext(i), matfechasext(i), MatCaracFRiesgo, MatResFRiesgo, "t", frmAnalisisEC.Picture1, frmAnalisisEC.Picture2, frmAnalisisEC.StatusBar1.Panels(3))
 matesc = UnirMatrices(matesc, mata, 1)
Next i
 'Call CalculaEficProsFWD(fecha, MatPosRiesgo, eficpros)
 Screen.MousePointer = 0
End Sub


Private Sub Command8_Click()
Dim mateficpros() As Variant
Dim efecpros As Double
Dim txtfecha As String
Dim fechaval As Date
Dim nofval As Integer
Dim noemisiones As Integer
Dim i As Integer
Dim j As Integer
Dim etiq1 As Panels
Dim cikos As String
Dim matemisiones() As Variant
Dim txtport As String
Dim txtmsg As String
Dim exito As Boolean

Screen.MousePointer = 11
txtfecha = InputBox("Dame la fecha a tomar como escenario tabla ", , Date)
If IsDate(txtfecha) Then
fechaval = CDate(txtfecha)
Text2.Text = fechaval
 Set etiq1 = frmResultadosER.StatusBar1.Panels(3)
 Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
 Call CEficProsSwapsPort(fechaval, txtport, txtmsg, exito)
nofval = UBound(mateficpros, 1)
'matemisiones = ObtFactUnicos(MatPos).cEmisionMD
noemisiones = UBound(matemisiones, 1)

frmResultadosER.MSFlexGrid2.Rows = 2
frmResultadosER.MSFlexGrid2.Cols = 2
frmResultadosER.MSFlexGrid2.Rows = nofval + 1
frmResultadosER.MSFlexGrid2.Cols = noemisiones + 1

For j = 1 To noemisiones + 1
If j <> 1 Then frmResultadosER.MSFlexGrid2.TextMatrix(0, j - 1) = matemisiones(j - 1, 1)
For i = 2 To nofval
 frmResultadosER.MSFlexGrid2.TextMatrix(i - 1, j - 1) = mateficpros(i, j)
Next i
If j <> 1 Then frmResultadosER.MSFlexGrid2.TextMatrix(nofval, j - 1) = Format(MREFProsSwap(j - 1, 2) * 100, "##00.00") & "%"
Next j
frmResultadosER.MSFlexGrid2.TextMatrix(nofval, 0) = "Porcentaje de aciertos"
MsgBox "Proceso terminado. Se generaron los archivos " & DirResVaR & "\resultados eficiencia prospectiva " & Format(FechaEval, "yyyymmdd") & ".txt"
End If
Screen.MousePointer = 0
End Sub


Private Sub Trans_Click()
Dim fecha1 As Date
Dim fecha2 As Date

Screen.MousePointer = 11
 fecha1 = CDate(Text1.Text)
 fecha2 = CDate(Text2.Text)
 Call InConexOracle("alm2", conAdo2)
 Call GuardaResEfRetroSwaps(fecha1, fecha1, fecha2, MatResEficSwaps, conAdo2)
 Call GuardaResEfRetroFwds(fecha2, MatResEficFwd, conAdo2)
 conAdo2.Close
Screen.MousePointer = 0
End Sub

