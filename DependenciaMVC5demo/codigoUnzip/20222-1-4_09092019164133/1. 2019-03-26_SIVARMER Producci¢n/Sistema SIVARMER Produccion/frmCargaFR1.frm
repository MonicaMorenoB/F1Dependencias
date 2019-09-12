VERSION 5.00
Begin VB.Form frmCargaFR1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factores de riesgo exactos"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   675
      Left            =   2880
      TabIndex        =   6
      Top             =   2700
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar historia factores"
      Height          =   735
      Left            =   390
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   2025
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   4665
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2000
         TabIndex        =   3
         Top             =   1110
         Width           =   2000
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2000
         TabIndex        =   2
         Top             =   390
         Width           =   2000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final"
         Height          =   195
         Left            =   200
         TabIndex        =   5
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de inicio"
         Height          =   195
         Left            =   200
         TabIndex        =   4
         Top             =   510
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmCargaFR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Screen.MousePointer = 11
If IsDate(Combo1.Text) And IsDate(Combo2.Text) Then
  fecha1 = CDate(Combo1.Text)
  fecha2 = CDate(Combo2.Text)
  CargaTasas = True
  'se cargan los plazos de las curvas
  txtpossim = Combo1.Text
  Unload Me
  Screen.MousePointer = 11
  frmProgreso.Show
   Call NuevaCFriesgo(fecha1, fecha2)
   Unload frmProgreso
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
  Combo1.Clear
  Combo2.Clear
  noreg = UBound(MatFechasVaR, 1)
  For i = 1 To noreg
      Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
      Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
  Next i
  Combo1.Text = MatFechasVaR(noreg, 1)
  Combo2.Text = MatFechasVaR(noreg, 1)
End Sub
