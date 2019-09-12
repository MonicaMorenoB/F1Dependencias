VERSION 5.00
Begin VB.Form frmCurvas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de curvas"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Extraer informacion"
      Height          =   705
      Left            =   630
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   1800
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   1170
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   1620
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Curva a extraer"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmCurvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim indice As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim idcurva As Integer
Dim nomcurva As String

Screen.MousePointer = 11
txtfecha1 = Combo2.Text
txtfecha2 = Combo3.Text
indice = Combo1.ListIndex
If IsDate(txtfecha1) And IsDate(txtfecha2) And indice >= 0 Then
   SiActTProc = True
   fecha1 = CDate(txtfecha1)
   fecha2 = CDate(txtfecha2)
   idcurva = MatCatCurvas(indice + 1, 1)
   nomcurva = MatCatCurvas(indice + 1, 2)
   frmProgreso.Show
   Call ObtenerHistCurvas(fecha1, fecha2, idcurva, nomcurva)
   Unload frmProgreso
   Call ActUHoraUsuario
   SiActTProc = False
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub


Private Sub Form_Load()
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer

noreg1 = UBound(MatCatCurvas, 1)
For i = 1 To noreg1
Combo1.AddItem MatCatCurvas(i, 2)
Next i
Combo1.Text = ""
noreg2 = UBound(MatFechasVaR, 1)
For i = 1 To noreg2
Combo2.AddItem MatFechasVaR(noreg2 - i + 1, 1)
Combo3.AddItem MatFechasVaR(noreg2 - i + 1, 1)
Next i
Combo2.Text = MatFechasVaR(noreg2, 1)
Combo3.Text = MatFechasVaR(noreg2, 1)
End Sub
