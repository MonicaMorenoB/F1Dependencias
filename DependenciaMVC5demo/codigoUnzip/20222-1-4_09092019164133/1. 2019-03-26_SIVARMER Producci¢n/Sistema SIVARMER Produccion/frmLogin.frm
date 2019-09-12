VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema VaR Mercado (SIVARMER)"
   ClientHeight    =   2775
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5595
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1639.98
   ScaleMode       =   0  'User
   ScaleWidth      =   5250.56
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Ejecución de subprocesos 3"
      Height          =   195
      Left            =   1598
      TabIndex        =   8
      Top             =   1861
      Width           =   2344
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Ejecución de subprocesos 2"
      Height          =   195
      Left            =   1598
      TabIndex        =   7
      Top             =   1530
      Width           =   2344
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ejecución de subprocesos 1"
      Height          =   195
      Left            =   1598
      TabIndex        =   6
      Top             =   1170
      Width           =   2344
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   156
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1344
      TabIndex        =   4
      Top             =   2200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   3000
      TabIndex        =   5
      Top             =   2200
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   552
      Width           =   2325
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4170
      Picture         =   "frmLogin.frx":0CCA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   525
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre de usuario:"
      Height          =   192
      Index           =   0
      Left            =   108
      TabIndex        =   0
      Top             =   192
      Width           =   1416
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      Height          =   192
      Index           =   1
      Left            =   107
      TabIndex        =   2
      Top             =   540
      Width           =   864
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
Check2.value = 0
Check3.value = 0
End Sub

Private Sub Check2_Click()
Check1.value = 0
Check3.value = 0
End Sub

Private Sub Check3_Click()
Check1.value = 0
Check2.value = 0
End Sub

Private Sub cmdCancel_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
    'establecer la variable global a "N"
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Unload Me
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub cmdOK_Click()
Dim exito As Boolean
Dim txtusuario As String
Dim txtcontraseña As String

txtusuario = txtUserName.Text
txtcontraseña = txtPassword.Text
If Check1.value Then
   ejecutaSubproc1 = True
Else
   ejecutaSubproc1 = False
End If
If Check2.value Then
   ejecutaSubproc2 = True
Else
   ejecutaSubproc2 = False
End If
If Check3.value Then
   ejecutaSubproc3 = True
Else
   ejecutaSubproc3 = False
End If


Call ValidarAcceso2(txtUserName, txtcontraseña, exito)
If exito Then
   Unload Me
End If
End Sub

Sub ValidarAcceso2(ByVal txtusuario As String, ByVal txtcontraseña As String, ByRef bl_exito As Boolean)
Dim siex As Boolean
Dim edocta As String
Dim siacceso As Boolean
Dim idusuario As Integer
Dim txtnombre As String
Dim indice0 As Integer
Dim i As Integer

'comprobar si la contraseña es correcta
'se busca el usuario en la tabla de datos
  MatUsuarios = LeerUsuariosSistema2()
  indice0 = 0
  For i = 1 To UBound(MatUsuarios, 1)
      If MatUsuarios(i, 2) = txtusuario Then
         indice0 = i
         Exit For
      End If
  Next i
  If indice0 <> 0 Then     'se encontro el usuario en la tabla de datos
     If MatUsuarios(indice0, 5) = "S" Then    'usuario vigente
        If MatUsuarios(indice0, 6) = "N" Or ejecutaSubproc1 Or ejecutaSubproc2 Or ejecutaSubproc3 Then      'usuario que puede entrar
        'se verifica que tenga un password vigente
           Call ValUsuarioDirActivo(txtusuario, siex, edocta, txtnombre)
           If siex Then
              If edocta = "Habilitado" Then
                 siacceso = IsGoodPWD("BANOBRAS", txtusuario, txtcontraseña)
                 If siacceso Then                                 'password valido
                    idusuario = MatUsuarios(indice0, 1)
                    NomUsuario = MatUsuarios(indice0, 2)
                    PerfilUsuario = MatUsuarios(indice0, 4) 'nivel del usuario
                    LoginSucceeded = True
                    bl_exito = True
                    Exit Sub
                 Else
                    bl_exito = False
                    Call AgregarIntentoFallido(indice0)  'password incorrecto
                 End If
              Else
                 bl_exito = False
                 MsgBox "El usuario " & txtusuario & " esta inhabilitado en directorio activo de windows"
              End If
           Else
              bl_exito = False
              MsgBox "El usuario " & txtusuario & " no existe en el directorio activo de windows"
           End If
        Else
           bl_exito = False
           MsgBox "El usuario " & MatUsuarios(indice0, 2) & " esta bloqueado. Llame a mesa de servicio"
        End If
     Else         'usuario sin acceso
        bl_exito = False
        MsgBox "El usuario " & MatUsuarios(indice0, 2) & " no tiene acceso": End
     End If
  Else          'el usuario no existe
     bl_exito = False
     MsgBox "La clave de usuario " & txtusuario & " no existe. Vuelva a intentarlo", , "Inicio de sesión"
     frmLogin.txtUserName.SetFocus
     'SendKeys "{Home}+{End}"
    Exit Sub
  End If
End Sub

Sub AgregarIntentoFallido(ByVal indice As Integer)
      If NoIntFallidos(indice) < 2 Then
         MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
         NoIntFallidos(indice) = NoIntFallidos(indice) + 1
      ElseIf NoIntFallidos(indice) >= 2 Then
        MsgBox "La contraseña no es válida.", , "Inicio de sesión"
        MsgBox "Hizo más de tres intentos de acceso. " & MatUsuarios(indice, 2) & " no tiene acceso al sistema"
        LoginSucceeded = False
        Unload Me
        Exit Sub
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
      End If

End Sub

Private Sub Form_Load()
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 MatUsuarios = LeerUsuariosSistema2()
noreg = UBound(MatUsuarios, 1)
ReDim NoIntFallidos(1 To noreg) As Integer
If OpcionBDatos = 1 Then
   frmLogin.Caption = "Sistema VaR de Mercado Banobras (Producción)"
ElseIf OpcionBDatos = 2 Then
   frmLogin.Caption = "Sistema VaR de Mercado Banobras (Desarrollo)"
ElseIf OpcionBDatos = 3 Then
   frmLogin.Caption = "Sistema VaR de Mercado Banobras (DRP)"
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Image1_Click()
txtPassword.PasswordChar = "*"
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtPassword.PasswordChar = ""
End Sub
