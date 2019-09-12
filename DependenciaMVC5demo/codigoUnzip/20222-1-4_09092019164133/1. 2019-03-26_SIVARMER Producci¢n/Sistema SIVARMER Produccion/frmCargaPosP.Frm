VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCargaPosP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de posiciones primarias"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   14685
   Icon            =   "frmCargaPosP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3900
   ScaleWidth      =   14685
   Begin VB.Frame Frame1 
      Caption         =   "Proxy swaps"
      Height          =   1185
      Left            =   200
      TabIndex        =   8
      Top             =   1800
      Width           =   14000
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5730
         TabIndex        =   12
         Top             =   510
         Width           =   7845
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   5250
         TabIndex        =   11
         Top             =   480
         Width           =   405
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2190
         TabIndex        =   10
         Top             =   510
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Importar"
         Height          =   615
         Left            =   300
         TabIndex        =   9
         Top             =   390
         Width           =   1485
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   330
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Deuda emisiones y posiciones primarias"
      Height          =   1200
      Left            =   200
      TabIndex        =   1
      Top             =   420
      Width           =   14000
      Begin VB.CommandButton Command23 
         Caption         =   "Importar"
         Height          =   585
         Left            =   150
         TabIndex        =   5
         Top             =   390
         Width           =   1635
      End
      Begin VB.CommandButton Command62 
         Caption         =   "..."
         Height          =   288
         Left            =   5100
         TabIndex        =   4
         Top             =   600
         Width           =   300
      End
      Begin VB.TextBox Text24 
         Height          =   288
         Left            =   5430
         TabIndex        =   3
         Top             =   570
         Width           =   8145
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación:"
         Height          =   195
         Left            =   4260
         TabIndex        =   7
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de registro"
         Height          =   195
         Left            =   2220
         TabIndex        =   6
         Top             =   390
         Width           =   1230
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   3375
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   926
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "30/10/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "30/10/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   20717
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCargaPosP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fecha As Date
Dim nomarchswaps As String
Dim noreg As Long

Screen.MousePointer = 11
  fecha = CDate(Combo1.Text)
  nomarchswaps = Text1.Text
  Call CrearPosSwapsSimArch(fecha, nomarchswaps, 1, "Real", "000000", noreg)
  Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub Command2_Click()
Dim nomarch As String
 CommonDialog1.ShowOpen
 nomarch = CommonDialog1.FileName
 Text1.Text = nomarch
End Sub

Private Sub Command23_Click()
Dim nomarch As String
Dim exito As Boolean
Dim noreg As Long
Dim fecha As Date
 nomarch = Text24.Text
 fecha = CDate(Combo3.Text)
 Screen.MousePointer = 11
 frmProgreso.Show
    Call ImpPosPrimSwapsArch(fecha, nomarch, noreg, exito)
 Unload frmProgreso
 Screen.MousePointer = 0
 MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command62_Click()
Dim nomarch As String
 CommonDialog1.ShowOpen
 nomarch = CommonDialog1.FileName
 Text24.Text = nomarch
End Sub

Private Sub Form_Load()
Dim noreg As Integer
noreg = UBound(MatFechasVaR, 1)
Combo1.Clear
Combo3.Clear
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo3.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
