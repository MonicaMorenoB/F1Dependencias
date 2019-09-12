VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDesglosePos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de la Posición"
   ClientHeight    =   10275
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9180
      Left            =   150
      TabIndex        =   1
      Top             =   750
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   16193
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Detalle de la posición"
      TabPicture(0)   =   "frmDesglosePos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Factores de riesgo exactos"
      TabPicture(1)   =   "frmDesglosePos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid2"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   7260
         Left            =   -74844
         TabIndex        =   3
         Top             =   504
         Width           =   8124
         _ExtentX        =   14340
         _ExtentY        =   12806
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8430
         Left            =   150
         TabIndex        =   2
         Top             =   420
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   14870
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.CommandButton BotonArchivoPos 
      Caption         =   "Exportar resultados a archivo de texto"
      Height          =   492
      Left            =   192
      TabIndex        =   0
      Top             =   144
      Width           =   1668
   End
End
Attribute VB_Name = "frmDesglosePos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
 MSFlexGrid1.Width = Maximo(frmDesglosePos.Width - 600, 0)
 MSFlexGrid1.Height = Maximo(frmDesglosePos.Height - 2500, 0)
 MSFlexGrid2.Width = Maximo(frmDesglosePos.Width - 600, 0)
 MSFlexGrid2.Height = Maximo(frmDesglosePos.Height - 2500, 0)

End Sub


Private Sub MSFlexGrid1_DblClick()
'esta rutina muestra el detalle de la operacion asi como su valuacion
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim mrvalflujo() As resValFlujo
Dim parval As ParamValPos
Dim indice As Long
Dim noreg As Long
Dim nocols As Long
Dim matp() As Double
Dim i As Integer
Dim j As Integer
Dim exito As Boolean
Dim txtmsg As String

Screen.MousePointer = 11
   indice = MSFlexGrid1.row
   Set parval = DeterminaPerfilVal("VALUACION")
   parval.indpos = indice
   matp = CalcValuacion(FechaEval, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
   If matpos(indice).C_Posicion = 4 Then
      frmDetValSWap.Show
      noreg = UBound(mrvalflujo, 1)
      nocols = UBound(mrvalflujo, 2)
      frmDetValSWap.MSFlexGrid1.Rows = noreg + 1
      frmDetValSWap.MSFlexGrid1.Cols = nocols + 1
      For i = 1 To noreg
          For j = 1 To nocols
              frmDetValSWap.MSFlexGrid1.TextMatrix(i, j) = mrvalflujo(i, j)
           Next j
      Next i
   End If
Screen.MousePointer = 0
End Sub
