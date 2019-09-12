VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de usuarios"
   ClientHeight    =   7215
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   40000
      Left            =   7350
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bitácora del sistema"
      Height          =   700
      Left            =   6300
      TabIndex        =   21
      Top             =   900
      Width           =   1400
   End
   Begin VB.CommandButton cmdConsultarBitacora 
      Caption         =   "Consultar bitacora de IF"
      Height          =   700
      Left            =   7800
      TabIndex        =   20
      Top             =   900
      Width           =   1400
   End
   Begin VB.Timer Timer1 
      Interval        =   40000
      Left            =   8460
      Top             =   120
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   700
      Left            =   4020
      TabIndex        =   13
      Top             =   900
      Width           =   1400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar usuario"
      Height          =   700
      Left            =   1710
      TabIndex        =   11
      Top             =   900
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear usuario"
      Height          =   700
      Left            =   200
      TabIndex        =   2
      Top             =   900
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   5100
      Left            =   180
      TabIndex        =   1
      Top             =   1800
      Width           =   7740
      Begin VB.Frame Frame3 
         Caption         =   "Bloqueado"
         Enabled         =   0   'False
         Height          =   800
         Left            =   4350
         TabIndex        =   17
         Top             =   4100
         Width           =   3015
         Begin VB.OptionButton Option4 
            Caption         =   "No"
            Height          =   195
            Left            =   1860
            TabIndex        =   19
            Top             =   400
            Width           =   1065
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Si"
            Height          =   195
            Left            =   360
            TabIndex        =   18
            Top             =   400
            Width           =   945
         End
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   288
         Left            =   100
         TabIndex        =   15
         Top             =   1992
         Width           =   6000
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         Height          =   1035
         ItemData        =   "frmUsuarios.frx":0000
         Left            =   120
         List            =   "frmUsuarios.frx":0002
         TabIndex        =   14
         Top             =   2688
         Width           =   4500
      End
      Begin VB.Frame Frame2 
         Caption         =   "Habilitado"
         Enabled         =   0   'False
         Height          =   800
         Left            =   216
         TabIndex        =   8
         Top             =   4100
         Width           =   2844
         Begin VB.OptionButton Option2 
            Caption         =   "No"
            Height          =   195
            Left            =   1536
            TabIndex        =   10
            Top             =   400
            Width           =   945
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Si"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   400
            Width           =   900
         End
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   288
         Left            =   100
         TabIndex        =   4
         Top             =   1300
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   288
         Left            =   100
         TabIndex        =   3
         Top             =   600
         Width           =   6000
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de tipos de usuarios"
         Height          =   204
         Left            =   120
         TabIndex        =   16
         Top             =   2448
         Width           =   2004
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de usuario"
         Height          =   192
         Left            =   200
         TabIndex        =   7
         Top             =   1700
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
         Height          =   192
         Left            =   200
         TabIndex        =   6
         Top             =   1000
         Width           =   1524
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre completo del usuario"
         Height          =   192
         Left            =   200
         TabIndex        =   5
         Top             =   300
         Width           =   2316
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   192
      TabIndex        =   0
      Top             =   360
      Width           =   5000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de usuarios"
      Height          =   195
      Left            =   195
      TabIndex        =   12
      Top             =   90
      Width           =   1185
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdUsuarioP As Integer

Private Sub cmdConsultarBitacora_Click()
frmBitacoraIF.Show 1
End Sub

Private Sub Combo1_Click()
 IdUsuarioP = Combo1.ListIndex + 1
 Call MuestraDatosUsuario(IdUsuarioP)
End Sub

Private Sub Command1_Click()
Dim txtnomusuario As String
Dim txtusuario As String
Dim txtgrupo As String
Dim txthabil As String
Dim txtbloquea As String
Dim i As Integer
Dim indice As Integer
Dim siex As Boolean
Dim siguarda As Boolean
Dim edocta As String

If Command1.Caption = "Crear usuario" Then
   Combo1.Enabled = False
   Combo1.Text = ""
   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
   Command1.Caption = "Guarda usuario"
   Command2.Enabled = False
   Text2.Enabled = True
   List1.Enabled = True
   Frame1.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   Option1.value = True
   Option4.value = True
ElseIf Command1.Caption = "Guarda usuario" Then
   txtnomusuario = Text1.Text
   txtusuario = Text2.Text
   txtgrupo = Text3.Text
   If Option1.value Then
      txthabil = "S"
   Else
      txthabil = "N"
   End If
   If Option3.value Then
      txtbloquea = "S"
   Else
      txtbloquea = "N"
   End If
   Call ValUsuarioDirActivo(txtusuario, siex, edocta, txtnomusuario)
   If siex And edocta = "Habilitado" Then
      siguarda = ValidarUsuarioN(txtusuario, txtnomusuario, txtgrupo, txthabil, txtbloquea)
      If siguarda Then
         Combo1.Enabled = True
         Command1.Caption = "Crear usuario"
         Command2.Enabled = True
         Text2.Enabled = False
         List1.Enabled = False
         Frame1.Enabled = True
         Frame2.Enabled = False
         Frame3.Enabled = False
         indice = 0
         For i = 1 To UBound(MatUsuarios, 1)
             indice = Maximo(indice, MatUsuarios(i, 1))
         Next i
         indice = indice + 1
         Call AgregarUsuario(indice, txtusuario, txtnomusuario, txtgrupo, txthabil, txtbloquea)
         Call GenerarParamUsuario(txtusuario)
         MatUsuarios = LeerUsuariosSistema2()
         Combo1.Clear
         For i = 1 To UBound(MatUsuarios, 1)
             Combo1.AddItem MatUsuarios(i, 3)
         Next i
         MensajeProc = "Se creo la cuenta del usuario " & txtusuario
         Call GuardaDatosBitacora(5, "Administracion", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
      End If
      IdUsuarioP = 0
   End If
End If
End Sub

Private Sub Command2_Click()
Dim indice As Integer
Dim txtnomusuario As String
Dim txtusuario As String
Dim txtgrupo As String
Dim txthabil As String
Dim txtbloquea As String
Dim i As Integer
Dim edocta As String
Dim siex As Boolean
Dim siguarda As Boolean


indice = Combo1.ListIndex + 1
If indice <> 0 Then
   If Command2.Caption = "Editar usuario" And indice >= 1 Then
      Combo1.Enabled = False
      Frame1.Enabled = True
      Frame2.Enabled = True
      Frame3.Enabled = True
      Command2.Caption = "Guarda usuario"
      Command1.Enabled = False
Print List1.Enabled = True
   ElseIf Command2.Caption = "Guarda usuario" Then
      txtnomusuario = Text1.Text
      txtusuario = Text2.Text
      txtgrupo = Text3.Text
      If Option1.value Then
         txthabil = "S"
      Else
         txthabil = "N"
      End If
      If Option3.value Then
         txtbloquea = "S"
      Else
         txtbloquea = "N"
      End If
      Call ValUsuarioDirActivo(txtusuario, siex, edocta, txtnomusuario)
      If siex And edocta = "Habilitado" Then
         siguarda = ValidarUsuarioE(txtusuario, txtnomusuario, txtgrupo, txthabil, txtbloquea)
         If siguarda Then
            Combo1.Enabled = True
            Frame1.Enabled = False
            Frame2.Enabled = False
            Frame3.Enabled = False
            Command2.Caption = "Editar usuario"
            Command1.Enabled = True
            List1.Enabled = False
            Call GuardaDatosUsuario(txtusuario, txtnomusuario, txtgrupo, txthabil, txtbloquea)
            MatUsuarios = LeerUsuariosSistema2()
            Combo1.Clear
            For i = 1 To UBound(MatUsuarios, 1)
                If Not EsVariableVacia(MatUsuarios(i, 3)) Then
                   Combo1.AddItem MatUsuarios(i, 3)
                End If
            Next i
            MensajeProc = "Se modifico la cuenta del usuario " & txtusuario
            Call GuardaDatosBitacora(5, "Administracion", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
            IdUsuarioP = 0
         End If
      End If
  
   Else
    MsgBox "No se puede editar el usuario 'sistemas'"
   End If
End If
End Sub

Private Sub Command3_Click()
frmBitacora.Show 1
End Sub

Private Sub Command4_Click()
If Command1.Caption = "Guarda usuario" Then
   Combo1.Enabled = True
   Command1.Caption = "Crear usuario"
   Command2.Enabled = True
   Text1.Enabled = False
   Text2.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   List1.Enabled = False
ElseIf Command2.Caption = "Guarda usuario" Then
   Combo1.Enabled = True
   Command1.Enabled = True
   Command2.Caption = "Editar usuario"
   Text1.Enabled = False
   Text2.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
 List1.Enabled = False
End If
End Sub


Private Sub Form_Load()
Dim i As Integer


MatParamSistema = LeerParametrosSist()
Combo1.Clear
For i = 1 To UBound(MatUsuarios, 1)
 Combo1.AddItem MatUsuarios(i, 3)
Next i
List1.AddItem "ADMUSUARIOS"
List1.AddItem "ADMINISTRADOR"
List1.AddItem "USUARIO"
List1.AddItem "REPORTES"
List1.AddItem "BITACORA"
If OpcionBDatos = 1 Then
   frmUsuarios.Caption = "Administración de usuarios (Producción)"
ElseIf OpcionBDatos = 2 Then
   frmUsuarios.Caption = "Administración de usuarios (Desarrollo)"
ElseIf OpcionBDatos = 3 Then
   frmUsuarios.Caption = "Administración de usuarios (DRP)"
End If
End Sub

Sub MuestraDatosUsuario(indice)
 frmUsuarios.Text1.Text = MatUsuarios(indice, 3)
 frmUsuarios.Text2.Text = MatUsuarios(indice, 2)
 frmUsuarios.Text3.Text = MatUsuarios(indice, 4)
 If MatUsuarios(indice, 5) = "S" Then
  frmUsuarios.Option1.value = True
 Else
  frmUsuarios.Option2.value = True
 End If
 If MatUsuarios(indice, 6) = "S" Then
  frmUsuarios.Option3.value = True
 Else
  frmUsuarios.Option4.value = True
 End If
End Sub

Sub GuardaDatosUsuario(ByVal txtusuario As String, ByVal txtnomusuario As String, ByVal txtgrupo As String, ByVal txthabil As String, ByVal txtbloquea As String)
Dim txtinsert As String

 txtinsert = "UPDATE " & TablaUsuarios & " SET "
 txtinsert = txtinsert & "NOMBRE = '" & txtnomusuario & "',"             'nombre del usuario
 txtinsert = txtinsert & "GRUPO ='" & txtgrupo & "',"                     'grupo al que pertenece
 txtinsert = txtinsert & "ACCESO ='" & txthabil & "',"                    'habilitado
 txtinsert = txtinsert & "ENLINEA = '" & txtbloquea & "'"                 'EN LINEA
 txtinsert = txtinsert & " WHERE USUARIO = '" & txtusuario & "'"
 ConAdo.Execute txtinsert
End Sub

Sub AgregarUsuario(ByVal indice As Integer, ByVal txtusuario As String, ByVal txtnomusuario As String, ByVal txtgrupo As String, ByVal txthabil As String, ByVal txtbloquea As String)
Dim txtinsert As String

txtinsert = "INSERT INTO " & TablaUsuarios & " VALUES("
txtinsert = txtinsert & indice & ","                    'indice del usuario
txtinsert = txtinsert & "'" & txtusuario & "',"         'clave del usuario
txtinsert = txtinsert & "'" & txtnomusuario & "',"      'nombre largo del usuario
txtinsert = txtinsert & "'" & txtgrupo & "',"           'grupo al que pertenece
txtinsert = txtinsert & "'" & txthabil & "',"           'usuario vigente en el sistema
txtinsert = txtinsert & "'" & txtbloquea & "',"         'usuario bloqueado en el sistema
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null,"
txtinsert = txtinsert & "null)"
ConAdo.Execute txtinsert
End Sub


Function ValidarUsuarioN(ByVal txtusuario As String, ByVal txtnomusuario As String, ByVal txtgrupo As String, ByVal txthabil As String, ByVal txtbloquea As String) As Boolean
Dim i As Integer

If Len(Trim(txtnomusuario)) = 0 Then
 MsgBox "El nombre del usuario no puede ser nula"
 ValidarUsuarioN = False
 Exit Function
End If
If Len(Trim(txtusuario)) = 0 Then
MsgBox "La clave de usuario no puede ser nula"
 ValidarUsuarioN = False
 Exit Function
End If
If Len(Trim(txtgrupo)) = 0 Then
   MsgBox "No se escogio un perfil"
 ValidarUsuarioN = False
 Exit Function
End If
If Len(Trim(txthabil)) = 0 Then
 ValidarUsuarioN = False
 Exit Function
End If
If Len(Trim(txtbloquea)) = 0 Then
 ValidarUsuarioN = False
 Exit Function
End If
For i = 1 To UBound(MatUsuarios, 1)
 If txtusuario = MatUsuarios(i, 2) Then
  MsgBox "La clave de usuario ya existe en la base de datos"
  ValidarUsuarioN = False
  Exit Function
 End If
Next i
ValidarUsuarioN = True
End Function

Function ValidarUsuarioE(ByVal txtusuario As String, ByVal txtnomusuario As String, ByVal txtgrupo As String, ByVal txthabil As String, ByVal txtbloquea As String) As Boolean
If Len(Trim(txtnomusuario)) = 0 Then
   MsgBox "El nombre de usuario no puede estar vacio"
   ValidarUsuarioE = False
   Exit Function
End If
If Len(Trim(txtusuario)) = 0 Then
   MsgBox "La clave de usuario no puede estar vacia"
   ValidarUsuarioE = False
   Exit Function
End If
If Len(Trim(txtgrupo)) = 0 Then
   MsgBox "No se ha definido un perfil"
   ValidarUsuarioE = False
   Exit Function
End If
If Len(Trim(txthabil)) = 0 Then
 ValidarUsuarioE = False
 Exit Function
End If
If Len(Trim(txtbloquea)) = 0 Then
 ValidarUsuarioE = False
 Exit Function
End If
ValidarUsuarioE = True
End Function


Sub BorrarUsuario(ByVal txtusuario As String)
ConAdo.Execute "DELETE FROM " & TablaUsuarios & " WHERE USUARIO = '" & txtusuario & "'"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call MostrarMensajeSistema(NomUsuario & " ha salido del sistema", frmUsuarios.Label2, 1, Date, Time, NomUsuario)
Call DesbloquearUsuario(NomUsuario)
End
End Sub

Private Sub List1_Click()
Text3.Text = List1.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim txtusuario As String
Dim siex As Boolean
Dim edocta As String
Dim txtnomusuario As String

If KeyAscii = 13 Then
   txtusuario = Text2.Text
   Call ValUsuarioDirActivo(txtusuario, siex, edocta, txtnomusuario)
   If siex Then
      If edocta = "Habilitado" Then
         Text1.Text = txtnomusuario
      Else
         MsgBox "El usuario " & txtusuario & " no esta habilitado en el Directorio Activo"
      End If
   Else
      MsgBox "El usuario " & txtusuario & " no existe en el dominio BANOBRAS"
   End If
End If
End Sub

Private Sub Timer1_Timer()
Dim uhora As Double
Dim tiempo As Double
If Not SiActTProc Then
   uhora = LeyendoEstadoUsuario(NomUsuario)
   tiempo = CDbl(Now) - uhora
   If tiempo * 24 * 60 > 5 Then
      MensajeProc = "La sesion ha estado inactiva mas de 5 minutos. Cerrando."
      Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
      Call DesbloquearUsuario(NomUsuario)
      MsgBox MensajeProc
      End
   End If
End If
End Sub

Private Sub Timer2_Timer()
Dim txtfecha As String
Dim txthora As String
Dim txtcadena As String

If SiActTProc Then
   txtfecha = "TO_DATE('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txthora = "TO_DATE('" & Format(Time, "HH:MM:SS") & "','HH24:MI:SS')"
   txtcadena = "UPDATE " & TablaUsuarios & " SET FUREPORTE = " & txtfecha & ", HUREPORTE = " & txthora & " WHERE USUARIO = '" & NomUsuario & "'"
   ConAdo.Execute txtcadena
   txtcadena = "UPDATE " & TablaSesiones & " SET F_ACTIVIDAD = " & txtfecha & ", H_ACTIVIDAD = " & txthora & " WHERE ID_SESION = '" & Id_Sesion & "'"
   ConAdo.Execute txtcadena
End If
End Sub
