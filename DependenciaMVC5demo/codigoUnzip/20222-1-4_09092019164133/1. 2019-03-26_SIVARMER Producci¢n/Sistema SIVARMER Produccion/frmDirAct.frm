VERSION 5.00
Begin VB.Form frmDirAct 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   1890
      TabIndex        =   0
      Top             =   1260
      Width           =   990
   End
End
Attribute VB_Name = "frmDirAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim objetoUsuario, gruposSeguridad
    Dim ultimoInicioSesion As String
    Dim dominio As String
    Dim nombreUsuario As String
    Dim estadoCuenta As String
    Dim gruposSeguridadUsuario As String
    
    dominio = InputBox("Nombre del dominio Windows Server", "")

    nombreUsuario = InputBox("Nombre de usuario del dominio", "")
    
 '   On Error GoTo cError
 
    On Error Resume Next
    
    Set objetoUsuario = GetObject("WinNT://" + dominio + "/" + nombreUsuario + ",user")
    If Err.Number = 0 Then
        If objetoUsuario.AccountDisabled = True Then
            estadoCuenta = "Deshabilitado"
            ultimoInicioSesion = "No existe"
        Else
            estadoCuenta = "Habilitado"
            ultimoInicioSesion = objetoUsuario.Get("Lastlogin")
        End If
        
        gruposSeguridad = ""
        For Each gruposSeguridad In objetoUsuario.Groups
            If gruposSeguridadUsuario = "" Then
              gruposSeguridadUsuario = gruposSeguridad.Name
            Else
              gruposSeguridadUsuario = gruposSeguridadUsuario + ", " + gruposSeguridad.Name
            End If
        Next
        'Mostramos los datos del usuario
        MsgBox "Nombre completo: " & objetoUsuario.Get("Fullname") & vbCrLf & _
            "Descripción: " & objetoUsuario.Get("Description") & vbCrLf & _
            "Nombre: " & objetoUsuario.Get("Name") & vbCrLf & _
            "Carpeta de inicio: " & objetoUsuario.Get("HomeDirectory") & vbCrLf & _
            "Script de inicio: " & objetoUsuario.Get("LoginScript") & vbCrLf & _
            "Último inicio de sesión: " & ultimoInicioSesion & vbCrLf & _
            "Perfil: " & objetoUsuario.Get("Profile") & vbCrLf & _
            "Estado de la cuenta: " & estadoCuenta & vbCrLf & _
            "Grupos seguridad: " & gruposSeguridadUsuario, vbInformation + vbOKOnly
        Set objetoUsuario = Nothing
    Else
        MsgBox "No existe el usuario " + nombreUsuario + " o el dominio " + dominio, vbExclamation + vbOKOnly
    End If
    
'cSalir:
'    Exit Sub
'
'cError:
'    MsgBox "Error " + CStr(Err.Number) + " " + Err.Description
'    GoTo cSalir
End Sub

