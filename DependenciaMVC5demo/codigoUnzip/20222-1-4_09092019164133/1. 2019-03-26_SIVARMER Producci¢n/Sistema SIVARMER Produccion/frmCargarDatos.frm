VERSION 5.00
Begin VB.Form frmCargarDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema para el calculo del VAR"
   ClientHeight    =   1524
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1524
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   276
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   -24
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.PictureBox Picture1 
      Height          =   372
      Left            =   72
      ScaleHeight     =   324
      ScaleWidth      =   7668
      TabIndex        =   1
      Top             =   912
      Width           =   7716
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   384
         Left            =   110
         Picture         =   "frmCargarDatos.frx":0000
         ScaleHeight     =   384
         ScaleWidth      =   384
         TabIndex        =   3
         Top             =   0
         Width           =   384
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF00FF&
         Height          =   324
         Left            =   0
         ScaleHeight     =   324
         ScaleWidth      =   24
         TabIndex        =   2
         Top             =   0
         Width           =   25
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando datos"
      Height          =   176
      Left            =   3311
      TabIndex        =   0
      Top             =   242
      Width           =   1067
   End
End
Attribute VB_Name = "frmCargarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
frmCargarDatos.Top = (Screen.Height - frmCargarDatos.Height) / 2
frmCargarDatos.Left = (Screen.Width - frmCargarDatos.Width) / 2
Label1.Left = (frmCargarDatos.Width - Label1.Width) / 2
frmCargarDatos.Refresh
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

