VERSION 5.00
Begin VB.Form frmListaOpR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de operaciones relacionadas"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar"
      Height          =   555
      Left            =   4260
      TabIndex        =   1
      Top             =   390
      Width           =   1785
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Elegir la operacion correcta:"
      Height          =   195
      Left            =   810
      TabIndex        =   2
      Top             =   180
      Width           =   1980
   End
End
Attribute VB_Name = "frmListaOpR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   IndOperR = Combo1.ListIndex + 1
   Unload Me
End Sub
