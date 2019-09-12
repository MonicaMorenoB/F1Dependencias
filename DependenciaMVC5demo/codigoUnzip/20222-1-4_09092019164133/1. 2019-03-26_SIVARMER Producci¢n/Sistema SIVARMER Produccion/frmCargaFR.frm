VERSION 5.00
Begin VB.Form frmCargaFR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de factores de riesgo"
   ClientHeight    =   4110
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar FR por nodos"
      Height          =   645
      Left            =   150
      TabIndex        =   8
      Top             =   3150
      Width           =   1680
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   2040
      TabIndex        =   5
      Top             =   216
      Width           =   1700
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2370
      TabIndex        =   1
      Top             =   3150
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de factores de riesgo"
      Height          =   1900
      Left            =   200
      TabIndex        =   0
      Top             =   864
      Width           =   3540
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   200
         TabIndex        =   7
         Top             =   1300
         Width           =   1700
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   200
         TabIndex        =   6
         Top             =   600
         Width           =   1700
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha final"
         Height          =   188
         Left            =   200
         TabIndex        =   3
         Top             =   1000
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inicio"
         Height          =   188
         Left            =   200
         TabIndex        =   2
         Top             =   300
         Width           =   1164
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Portafolio de factores"
      Height          =   192
      Left            =   200
      TabIndex        =   4
      Top             =   264
      Width           =   1668
   End
End
Attribute VB_Name = "frmCargaFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtportfr As String
Dim exito As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If Len(Trim(Combo1.Text)) <> 0 Then
 If IsDate(Combo2.Text) And IsDate(Combo3.Text) Then
 SiActTProc = True
  fecha1 = CDate(Combo2.Text)
  fecha2 = CDate(Combo3.Text)
  'se cargan los plazos de las curvas
  txtportfr = Combo1.Text
  Unload Me
  Screen.MousePointer = 11
  frmProgreso.Show
   Call CrearMatFRiesgo2(fecha1, fecha2, MatFactRiesgo, "", exito)
  Unload frmProgreso
  Call ActUHoraUsuario
  SiActTProc = False
  Screen.MousePointer = 0
 End If
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim n As Integer
Dim i As Integer
Dim noreg As Integer

  Combo1.Clear
  n = UBound(MatPortCurvas, 1)
  For i = 1 To n
  Combo1.AddItem MatPortCurvas(i, 1)
  Next i
  Combo1.Text = NombrePortFR
  noreg = UBound(MatFechasVaR, 1)
  For i = 1 To noreg
  Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
  Combo3.AddItem MatFechasVaR(noreg - i + 1, 1)
  Next i
  Combo2.Text = MatFechasVaR(noreg, 1)
  Combo3.Text = MatFechasVaR(noreg, 1)
End Sub
