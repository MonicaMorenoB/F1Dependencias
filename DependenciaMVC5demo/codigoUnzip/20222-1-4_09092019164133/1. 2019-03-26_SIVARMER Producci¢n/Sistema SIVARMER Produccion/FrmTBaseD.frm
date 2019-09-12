VERSION 5.00
Begin VB.Form frmTBaseD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de acceso"
   ClientHeight    =   3030
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Oracle desarrollo"
      Height          =   192
      Left            =   624
      TabIndex        =   7
      Top             =   960
      Width           =   2148
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar"
      Height          =   396
      Left            =   1512
      TabIndex        =   6
      Top             =   2352
      Width           =   1740
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   300
      Left            =   3816
      TabIndex        =   5
      Top             =   1800
      Width           =   612
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   648
      TabIndex        =   3
      Top             =   1800
      Width           =   3108
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Access"
      Height          =   192
      Left            =   624
      TabIndex        =   2
      Top             =   1272
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Oracle producción"
      Height          =   192
      Left            =   648
      TabIndex        =   1
      Top             =   672
      Value           =   -1  'True
      Width           =   1956
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ubicación"
      Height          =   192
      Left            =   672
      TabIndex        =   4
      Top             =   1536
      Width           =   732
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bases de datos"
      Height          =   192
      Left            =   648
      TabIndex        =   0
      Top             =   288
      Width           =   1140
   End
End
Attribute VB_Name = "frmTBaseD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
If Option1.value Then
 OpcionBDatos = 2
ElseIf Option3.value Then
 OpcionBDatos = 1
ElseIf Option2.value Then
 DirBases = frmTBaseD.Text1.Text
End If
Unload Me
End Sub

Private Sub Form_Load()
frmTBaseD.Text1.Text = DirBases
End Sub
