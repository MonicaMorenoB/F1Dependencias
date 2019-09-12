VERSION 5.00
Begin VB.Form frmPyGSubport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de P y G por subportafolios"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5190
      TabIndex        =   12
      Text            =   "Normal"
      Top             =   630
      Width           =   2145
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   200
      TabIndex        =   9
      Top             =   600
      Width           =   4320
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   200
      TabIndex        =   4
      Text            =   "1"
      Top             =   3780
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   200
      TabIndex        =   3
      Text            =   "500"
      Top             =   3030
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   2
      Top             =   2430
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1560
      Width           =   4290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar informacion"
      Height          =   555
      Left            =   5640
      TabIndex        =   0
      Top             =   3210
      Width           =   1425
   End
   Begin VB.Label Label6 
      Caption         =   "Escenario de factores de riesgo"
      Height          =   165
      Left            =   5220
      TabIndex        =   11
      Top             =   390
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   10
      Top             =   300
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Agrupados por"
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   1230
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   2250
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Horizonte de tiempo"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   3510
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No de escenarios"
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   2820
      Width           =   1245
   End
End
Attribute VB_Name = "frmPyGSubport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtsubport As String
Dim txtport As String
Dim fecha As Date
Dim noesc As Integer
Dim htiempo As Integer
Dim txtportfr As String

Screen.MousePointer = 11
txtsubport = Combo1.Text
txtport = Combo3.Text
fecha = CDate(Combo2.Text)
noesc = Val(Text1.Text)
htiempo = Val(Text2.Text)
txtportfr = Text3.Text
Unload frmPyGSubport
Call LeerPyGHistPortPos(fecha, txtport, txtportfr, txtsubport, noesc, htiempo)
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Integer
Combo3.AddItem "TOTAL"
Combo3.AddItem "NEGOCIACION + INVERSION"
For i = 1 To UBound(MatListaPortPos, 1)
    Combo1.AddItem MatListaPortPos(i, 1)
Next i
For i = 1 To UBound(MatFechasVaR, 1)
    Combo2.AddItem MatFechasVaR(UBound(MatFechasVaR, 1) - i + 1, 1)
Next i
End Sub
