VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSensibilidades 
   Caption         =   "Sensibilidades calculadas"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   8115
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   14314
      _Version        =   393216
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "frmSensibilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 MSFlexGrid1.Width = frmSensibilidades.Width - 500
 MSFlexGrid1.Height = frmSensibilidades.Height - 1500
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

