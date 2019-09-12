VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPasswordN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Contraseña"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7785
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1245
      Left            =   5580
      TabIndex        =   7
      Top             =   1140
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2196
      _Version        =   393217
      TextRTF         =   $"frmPasswordN.frx":0000
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1000
      TabIndex        =   5
      Top             =   300
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   585
      Left            =   180
      TabIndex        =   4
      Top             =   2670
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2070
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   5000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña"
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   1830
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva contraseña"
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   900
      Width           =   1320
   End
End
Attribute VB_Name = "frmPasswordN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtusuario As String
Dim txtpass1 As String
Dim txtpass2 As String
Dim indice As Integer
Dim idusuario As Integer
Dim sicambiarpass As Boolean
Dim MatUsuarioso() As Variant

txtusuario = Text3.Text
txtpass1 = Text1.Text
txtpass2 = Text2.Text
MatUsuarioso = RutinaOrden(MatUsuarios, 2, SRutOrden)
indice = BuscarValorVector(txtusuario, MatUsuarioso, 2)
idusuario = MatUsuarioso(indice, 1)
sicambiarpass = ValidarCreacionPass(idusuario, txtpass1, txtpass2)
If sicambiarpass Then
   Call CambiarPassVigente(idusuario, txtpass2, "N", Date)
   Unload Me
 End If
End Sub


Private Sub Form_Load()
Text3.Text = txtUsuarioCC
End Sub
