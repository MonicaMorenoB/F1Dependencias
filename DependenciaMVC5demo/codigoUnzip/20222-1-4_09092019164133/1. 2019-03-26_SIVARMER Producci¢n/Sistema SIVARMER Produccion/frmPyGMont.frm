VERSION 5.00
Begin VB.Form frmPyGMontOper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escenarios Montecarlo por operación"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   200
      TabIndex        =   11
      Top             =   1600
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   200
      TabIndex        =   9
      Text            =   "10000"
      Top             =   3400
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   200
      TabIndex        =   8
      Text            =   "1"
      Top             =   2800
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   200
      TabIndex        =   6
      Text            =   "500"
      Top             =   2190
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   4
      Top             =   1000
      Width           =   3000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   3
      Top             =   400
      Width           =   4000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar archivo"
      Height          =   615
      Left            =   3900
      TabIndex        =   0
      Top             =   3120
      Width           =   1725
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   12
      Top             =   1400
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No de simulaciones"
      Height          =   195
      Left            =   195
      TabIndex        =   10
      Top             =   3200
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Horizonte de tiempo"
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   2600
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No de escenarios"
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   2000
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subportafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   2
      Top             =   800
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   200
      Width           =   660
   End
End
Attribute VB_Name = "frmPyGMontOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtport As String
Dim txtsubport As String
Dim fecha As Date
Dim noesc As Integer
Dim htiempo As Integer
Dim nosim As Integer

Screen.MousePointer = 11
If IsDate(Combo3.Text) And Not EsVariableVacia(Combo1.Text) Then
txtport = Combo1.Text
txtsubport = Combo2.Text
fecha = CDate(Combo3.Text)
noesc = Val(Text1.Text)
htiempo = Val(Text2.Text)
nosim = Val(Text3.Text)
Unload frmPyGOper
SiActTProc = True
frmProgreso.Show
Call LeerPyGMontOper(fecha, txtport, "Normal", txtsubport, noesc, htiempo, nosim)
Unload frmProgreso
MsgBox "Fin de proceso"
Call ActUHoraUsuario
SiActTProc = False
Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Dim i As Integer

Combo1.AddItem "TOTAL"
Combo1.AddItem "NEGOCIACION + INVERSION"
For i = 1 To UBound(MatPortPosicion, 1)
 Combo2.AddItem MatPortPosicion(i, 2)
Next i
For i = 1 To UBound(MatFechasVaR, 1)
 Combo3.AddItem MatFechasVaR(UBound(MatFechasVaR, 1) - i + 1, 1)
Next i

End Sub
