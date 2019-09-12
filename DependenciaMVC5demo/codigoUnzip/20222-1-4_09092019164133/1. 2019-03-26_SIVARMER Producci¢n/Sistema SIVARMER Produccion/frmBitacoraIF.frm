VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBitacoraIF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora de intentos fallidos"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10545
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   510
      Width           =   2235
   End
   Begin VB.CommandButton cmdGuardarInformacion 
      Caption         =   "Exportar información a archivo de texto"
      Height          =   600
      Left            =   4230
      TabIndex        =   3
      Top             =   270
      Width           =   1400
   End
   Begin VB.CommandButton cmdLeerInformacion 
      Caption         =   "Leer informacion"
      Height          =   600
      Left            =   2700
      TabIndex        =   2
      Top             =   240
      Width           =   1400
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6105
      Left            =   195
      TabIndex        =   0
      Top             =   1005
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   10769
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   210
      Width           =   435
   End
End
Attribute VB_Name = "frmBitacoraIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGuardarInformacion_Click()
Dim fecha As Date
Dim nomarch As String
Dim noreg As Integer
Dim txtcadena As String
Dim i As Integer

If IsDate(Combo1.Text) Then
Screen.MousePointer = 11
fecha = CDate(Combo1.Text)
 nomarch = "Bitacora de accesos fallidos " & Format(fecha, "yyyy-mm-dd") & ".txt"
 frmCalVar.CommonDialog1.FileName = nomarch
 frmCalVar.CommonDialog1.ShowSave
 nomarch = frmCalVar.CommonDialog1.FileName
Open nomarch For Output As #1
noreg = MSFlexGrid1.Rows - 1
txtcadena = "Fecha" & Chr(9)
txtcadena = txtcadena & "Hora" & Chr(9)
txtcadena = txtcadena & "Usuario" & Chr(9)
txtcadena = txtcadena & "No intentos" & Chr(9)
Print #1, txtcadena
For i = 1 To noreg
txtcadena = MSFlexGrid1.TextMatrix(i, 1) & Chr(9)
txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i, 2) & Chr(9)
txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i, 3) & Chr(9)
txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i, 4) & Chr(9)
Print #1, txtcadena
Next i
Close #1
MensajeProc = "Se exporto la informacion de la bitacora a un archivo " & fecha
Call GuardaDatosBitacora(3, "Consulta", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
Screen.MousePointer = 0
End If
End Sub

Private Sub cmdLeerInformacion_Click()
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim fecha As Date
Dim mata() As Variant

Screen.MousePointer = 11
txtfecha = Combo1.Text
If Len(Trim(txtfecha)) <> 0 Then
   fecha = CDate(txtfecha)
   mata = LeerBitacoraIF(fecha)
   If UBound(mata, 1) <> 0 Then
      noreg = UBound(mata, 1)
      MSFlexGrid1.Cols = 5
      MSFlexGrid1.Rows = noreg + 1
      MSFlexGrid1.TextMatrix(0, 1) = "FECHA DE OPERACION"
      MSFlexGrid1.TextMatrix(0, 2) = "HORA DE OPERACION"
      MSFlexGrid1.TextMatrix(0, 3) = "USUARIO"
      MSFlexGrid1.TextMatrix(0, 4) = "NO DE INTENTOS FALLIDOS"
      MSFlexGrid1.ColWidth(0) = 100
      MSFlexGrid1.ColWidth(1) = 2000
      MSFlexGrid1.ColWidth(2) = 2000
      MSFlexGrid1.ColWidth(3) = 2000
      MSFlexGrid1.ColWidth(4) = 2000
      For i = 1 To noreg
          MSFlexGrid1.TextMatrix(i, 1) = mata(i, 1)
          MSFlexGrid1.TextMatrix(i, 2) = mata(i, 2)
          MSFlexGrid1.TextMatrix(i, 3) = mata(i, 3)
          MSFlexGrid1.TextMatrix(i, 4) = mata(i, 4)
      Next i
   End If
    MensajeProc = "Se leyo la informacion de la bitacora de accesos fallidos"
    Call GuardaDatosBitacora(3, "Consulta", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Call ListarFBitIF
If OpcionBDatos = 1 Then
   frmBitacoraIF.Caption = "Bitácora de Intentos fallidos (Producción)"
ElseIf OpcionBDatos = 2 Then
   frmBitacoraIF.Caption = "Bitácora de Intentos fallidos (Desarrollo)"
ElseIf OpcionBDatos = 3 Then
  frmBitacoraIF.Caption = "Bitácora de Intentos Fallidos (DRP)"
End If
End Sub

Sub ListarFBitIF()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfiltro = "SELECT FECHA FROM " & TablaBitacoraIF & " GROUP BY FECHA ORDER BY FECHA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg) As Date
For i = 1 To noreg
  mata(i) = rmesa.Fields(0)
rmesa.MoveNext
Next i
rmesa.Close
For i = 1 To noreg
   Combo1.AddItem mata(noreg - i + 1)
Next i
End If
End Sub

