VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de conbtraseña"
   ClientHeight    =   4260
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1365
      Left            =   5970
      TabIndex        =   9
      Top             =   1140
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2408
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPassword.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1095
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   5000
   End
   Begin VB.TextBox Text3 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2445
      Width           =   5000
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   288
      Left            =   975
      TabIndex        =   1
      Top             =   330
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   468
      Left            =   240
      TabIndex        =   0
      Top             =   2970
      Width           =   1548
   End
   Begin VB.Label lblContraseñaActual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña actual"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   840
      Width           =   1290
   End
   Begin VB.Label lblNuevaContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva contraseña"
      Height          =   195
      Left            =   285
      TabIndex        =   7
      Top             =   1530
      Width           =   1320
   End
   Begin VB.Label lblConfirmarContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar contraseña"
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   2130
      Width           =   1530
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Left            =   285
      TabIndex        =   5
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtusuario As String
Dim txtpass1 As String
Dim txtpass2 As String
Dim txtpass3 As String
Dim indice As Integer
Dim idusuario As Integer
Dim sicambiarpass As Boolean
Dim MatUsuarioso() As Variant

txtusuario = Text4.Text
txtpass1 = Text1.Text
txtpass2 = Text2.Text
txtpass3 = Text3.Text
MatUsuarioso = RutinaOrden(MatUsuarios, 2, SRutOrden)
indice = BuscarValorVector(txtusuario, MatUsuarioso, 2)
idusuario = MatUsuarioso(indice, 1)
sicambiarpass = ValidarCambioPass(idusuario, txtpass1, txtpass2, txtpass3)
If sicambiarpass Then
   Call CambiarPassVigente(idusuario, txtpass2, "S", Date)
   Unload Me
End If
End Sub

Private Sub Form_Load()
Text4.Text = txtUsuarioCC
End Sub

