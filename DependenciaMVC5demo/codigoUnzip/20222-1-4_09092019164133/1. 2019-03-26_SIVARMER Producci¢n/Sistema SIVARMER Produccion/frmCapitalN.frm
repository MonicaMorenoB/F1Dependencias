VERSION 5.00
Begin VB.Form frmCapitalN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capital Neto y Basico"
   ClientHeight    =   3420
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2500
      TabIndex        =   7
      Top             =   912
      Width           =   1716
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   200
      TabIndex        =   6
      Top             =   936
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   468
      Left            =   168
      TabIndex        =   4
      Top             =   2808
      Width           =   1620
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   2500
      TabIndex        =   3
      Top             =   1752
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   200
      TabIndex        =   1
      Top             =   1728
      Width           =   2076
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "al"
      Height          =   192
      Left            =   2500
      TabIndex        =   9
      Top             =   600
      Width           =   132
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vigencia del "
      Height          =   192
      Left            =   200
      TabIndex        =   8
      Top             =   624
      Width           =   936
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Correspodiente al mes de "
      Height          =   192
      Left            =   240
      TabIndex        =   5
      Top             =   168
      Width           =   1896
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Capital Basico"
      Height          =   192
      Left            =   2500
      TabIndex        =   2
      Top             =   1488
      Width           =   1044
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Capital Neto"
      Height          =   192
      Left            =   200
      TabIndex        =   0
      Top             =   1440
      Width           =   888
   End
End
Attribute VB_Name = "frmCapitalN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim resp As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
resp = InputBox("Estan correctos los datos?", , "S")
If UCase(resp) Then

End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 Cancel = 0
 frmMensajes.Hide
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 Cancel = True
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

