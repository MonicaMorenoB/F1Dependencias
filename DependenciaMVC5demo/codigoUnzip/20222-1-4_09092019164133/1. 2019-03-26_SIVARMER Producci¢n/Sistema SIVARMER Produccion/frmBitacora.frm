VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBitacora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora de Operación"
   ClientHeight    =   10515
   ClientLeft      =   -30
   ClientTop       =   360
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3030
      TabIndex        =   5
      Top             =   300
      Width           =   1605
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   12270
      Top             =   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cargar registros"
      Height          =   600
      Left            =   5190
      TabIndex        =   3
      Top             =   210
      Width           =   1700
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   2
      Top             =   310
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar bitacora a archivo de texto"
      Height          =   600
      Left            =   7800
      TabIndex        =   1
      Top             =   150
      Width           =   1700
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   9330
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   16457
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de evento"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   90
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   4
      Top             =   100
      Width           =   450
   End
End
Attribute VB_Name = "frmBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fecha As Date
Dim nomarch As String
Dim txtcadena As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer

If IsDate(Combo1.Text) Then
Screen.MousePointer = 11
fecha = CDate(Combo1.Text)
 nomarch = "Bitacora de operacion " & Format(fecha, "yyyy-mm-dd") & ".txt"
 frmCalVar.CommonDialog1.FileName = nomarch
 frmCalVar.CommonDialog1.ShowSave
 nomarch = frmCalVar.CommonDialog1.FileName
Open nomarch For Output As #1
noreg = MSFlexGrid1.Rows - 1
txtcadena = ""
For i = 1 To 11
   txtcadena = txtcadena & MSFlexGrid1.TextMatrix(0, i) & Chr(9)
Next i
Print #1, txtcadena
For i = 1 To noreg
    txtcadena = ""
    For j = 1 To 11
        txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i, j) & Chr(9)
    Next j
    Print #1, txtcadena
Next i
Close #1
MensajeProc = "Se exporto la informacion de la bitacora a un archivo " & fecha
Screen.MousePointer = 0
End If
End Sub

Private Sub Command2_Click()
Dim fecha As Date
Dim idevento As Integer
Dim txtfecha As String
Dim mata() As Variant
Dim noreg As Integer
Dim i As Long
Dim j As Long

Screen.MousePointer = 11
txtfecha = Combo1.Text
If Combo2.Text = "Acceso" Then
   idevento = 1
ElseIf Combo2.Text = "Proceso" Then
   idevento = 2
ElseIf Combo2.Text = "Consulta" Then
   idevento = 3
Else
   idevento = 0
End If
If Len(Trim(txtfecha)) <> 0 And idevento <> 0 Then
 fecha = CDate(txtfecha)
 mata = LeerBitacoraOp(fecha, idevento)
If UBound(mata, 1) <> 0 Then
 noreg = UBound(mata, 1)
 frmBitacora.MSFlexGrid1.Cols = 12
 frmBitacora.MSFlexGrid1.Rows = noreg + 1
 frmBitacora.MSFlexGrid1.TextMatrix(0, 1) = "Tipo de evento"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 2) = "ID Proceso"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 3) = "Descripcion"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 4) = "Usuario"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 5) = "Direccion IP"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 6) = "Fecha"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 7) = "Fecha de inicio"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 8) = "Hora de inicio"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 9) = "Fecha final"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 10) = "hora final"
 frmBitacora.MSFlexGrid1.TextMatrix(0, 11) = "Observacion"
 frmBitacora.MSFlexGrid1.ColWidth(0) = 100
 frmBitacora.MSFlexGrid1.ColWidth(1) = 2000
 frmBitacora.MSFlexGrid1.ColWidth(2) = 3000
 frmBitacora.MSFlexGrid1.ColWidth(3) = 3000
 frmBitacora.MSFlexGrid1.ColWidth(4) = 4000
 
 For i = 1 To noreg
     For j = 1 To 11
        frmBitacora.MSFlexGrid1.TextMatrix(i, j) = mata(i, j)
     Next j
 Next i
 End If
 MensajeProc = "Se leyo la informacion de la bitacora"
Call GuardaDatosBitacora(3, "Consulta", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
Call ListarFBitacora
Combo2.Clear
Combo2.AddItem "Acceso"
Combo2.AddItem "Proceso"
Combo2.AddItem "Consulta"

If OpcionBDatos = 1 Then
   frmBitacora.Caption = "Bitácora del sistema (Producción)"
ElseIf OpcionBDatos = 2 Then
   frmBitacora.Caption = "Bitácora del sistema (Desarrollo)"
ElseIf OpcionBDatos = 3 Then
  frmBitacora.Caption = "Bitácora del sistema (DRP)"
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
MSFlexGrid1.Width = Maximo(frmBitacora.Width - 500, 0)
MSFlexGrid1.Height = Maximo(frmBitacora.Height - 1200, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)

If PerfilUsuario = "BITACORA" Then
  MensajeProc = NomUsuario & " ha salido del sistema"
  Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
  Call DesbloquearUsuario(NomUsuario)
  RGuardarPL.Close
  RegResCVA.Close
  RegResMakeW.Close
  conAdo.Close
  conAdoBD.Close
  End
End If
End Sub

Private Sub Timer1_Timer()
Dim uhora As Double
Dim tiempo As Double

If PerfilUsuario = "BITACORA" Then
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

Sub ListarFBitacora()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer

txtfiltro = "SELECT FECHAP FROM " & TablaBitacora & " GROUP BY FECHAP ORDER BY FECHAP"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
RFlujos.Open txtfiltro1, conAdo
noreg = RFlujos.Fields(0)
RFlujos.Close
If noreg <> 0 Then
   RFlujos.Open txtfiltro, conAdo
ReDim mata(1 To noreg) As Date
   RFlujos.MoveFirst
   For i = 1 To noreg
       mata(i) = RFlujos.Fields(0)
       RFlujos.MoveNext
   Next i
RFlujos.Close
For i = 1 To noreg
   Combo1.AddItem mata(noreg - i + 1)
Next i
End If
End Sub

