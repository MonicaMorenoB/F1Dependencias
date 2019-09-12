VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPyGOper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de las P y G de un portafolio"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5310
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   200
      TabIndex        =   9
      Top             =   240
      Width           =   5205
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   8
      Top             =   1620
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   3360
      TabIndex        =   6
      Top             =   3330
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   200
      TabIndex        =   4
      Text            =   "1"
      Top             =   3390
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   200
      TabIndex        =   3
      Text            =   "250"
      Top             =   2520
      Width           =   1755
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Top             =   870
      Width           =   5145
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio de posición"
      Height          =   195
      Left            =   200
      TabIndex        =   10
      Top             =   30
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   1290
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Horizonte de tiempo"
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   2940
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "No de escenarios"
      Height          =   195
      Left            =   200
      TabIndex        =   2
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subportafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   0
      Top             =   600
      Width           =   930
   End
End
Attribute VB_Name = "frmPyGOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtport As String
Dim txtgrupo As String
Dim noesc As Integer
Dim htiempo As Integer
Dim fecha As Date
Dim tfecha As String



txtport = Combo3.Text
txtgrupo = Combo1.Text
If IsDate(Combo2.Text) And Not EsVariableVacia(txtport) And Not EsVariableVacia(txtgrupo) Then
Screen.MousePointer = 11
fecha = CDate(Combo2.Text)
noesc = Val(Text1.Text)
htiempo = Val(Text2.Text)
Unload frmPyGOper
SiActTProc = True
frmProgreso.Show
Call LeerPyGPostxt(fecha, txtport, "Normal", txtgrupo, noesc, htiempo)
Unload frmProgreso
Call ActUHoraUsuario
SiActTProc = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
 Combo3.AddItem "TOTAL"
 Combo3.AddItem "NEGOCIACION + INVERSION"
For i = 1 To UBound(MatPortPosicion, 1)
  Combo1.AddItem MatPortPosicion(i, 2)
Next i
For i = 1 To UBound(MatFechasVaR, 1)
 Combo2.AddItem MatFechasVaR(UBound(MatFechasVaR, 1) - i + 1, 1)
Next i

End Sub
