VERSION 5.00
Begin VB.Form frmConsolidado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Consolidado"
   ClientHeight    =   3135
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2490
      TabIndex        =   5
      Top             =   2430
      Width           =   1089
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   1290
      TabIndex        =   4
      Top             =   2430
      Width           =   1045
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   300
      TabIndex        =   2
      Top             =   700
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A la fecha:"
      Height          =   176
      Left            =   300
      TabIndex        =   1
      Top             =   1200
      Width           =   682
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "De la fecha:"
      Height          =   176
      Left            =   300
      TabIndex        =   0
      Top             =   400
      Width           =   781
   End
End
Attribute VB_Name = "frmConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Unload Me
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Load()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmConsolidado.Left = (Screen.Width - frmConsolidado.Width) / 2
frmConsolidado.top = (Screen.Height - frmConsolidado.Height) / 2
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
