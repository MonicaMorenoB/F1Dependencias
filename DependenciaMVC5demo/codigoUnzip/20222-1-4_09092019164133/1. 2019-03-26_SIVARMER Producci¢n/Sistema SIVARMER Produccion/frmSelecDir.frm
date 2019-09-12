VERSION 5.00
Begin VB.Form frmSelecDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar directorio"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   200
      TabIndex        =   3
      Top             =   240
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3660
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   390
      TabIndex        =   1
      Top             =   3690
      Width           =   1500
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   200
      TabIndex        =   0
      Top             =   690
      Width           =   4635
   End
End
Attribute VB_Name = "frmSelecDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
DirSalida = Dir1.List(Dir1.ListIndex)
Unload Me
End Sub

Private Sub Command2_Click()
DirSalida = ""
Unload Me
End Sub

Private Sub Drive1_Change()
On Error Resume Next
  Dir1.Path = Drive1.Drive
On Error GoTo 0
End Sub
