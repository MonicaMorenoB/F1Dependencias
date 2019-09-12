VERSION 5.00
Begin VB.Form frmProgreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Avance del Proceso"
   ClientHeight    =   1275
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   7680
   Icon            =   "frmProgreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   260
      Left            =   55
      ScaleHeight     =   195
      ScaleWidth      =   7440
      TabIndex        =   0
      Top             =   750
      Width           =   7500
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   250
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   0
         TabIndex        =   1
         Top             =   0
         Width           =   66
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo la informacion"
      Height          =   192
      Left            =   3048
      TabIndex        =   3
      Top             =   144
      Width           =   1668
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2112
      TabIndex        =   2
      Top             =   120
      Width           =   852
   End
End
Attribute VB_Name = "frmProgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmProgreso.top = Screen.Height * 2 / 3
frmProgreso.Left = (Screen.Width - frmProgreso.Width) / 2
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
  Picture2.Width = Picture1.Width * AvanceProc
  Label2.Caption = MensajeProc
End Sub

