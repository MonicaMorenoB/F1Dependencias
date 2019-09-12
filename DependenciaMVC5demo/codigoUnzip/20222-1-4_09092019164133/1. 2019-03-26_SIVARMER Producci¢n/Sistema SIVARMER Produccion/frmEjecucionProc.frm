VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEjecucionProc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ejecución de procesos"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15090
   Icon            =   "frmEjecucionProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   15090
   Begin VB.Frame Frame1 
      Caption         =   "Grupo de procesos"
      Height          =   765
      Left            =   6690
      TabIndex        =   13
      Top             =   810
      Width           =   4485
      Begin VB.OptionButton Option2 
         Caption         =   "Procesos 2"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   330
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Procesos 1"
         Height          =   285
         Left            =   300
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar procesos pendientes"
      Height          =   650
      Left            =   2100
      TabIndex        =   8
      Top             =   900
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   400
      Width           =   2115
   End
   Begin VB.Frame lista 
      Caption         =   "Lista de procesos"
      Height          =   8865
      Left            =   180
      TabIndex        =   5
      Top             =   1600
      Width           =   14745
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   300
         TabIndex        =   10
         Top             =   660
         Width           =   1755
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cambiar estado procesos"
         Height          =   650
         Left            =   3000
         TabIndex        =   9
         Top             =   400
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7365
         Left            =   195
         TabIndex        =   6
         Top             =   1200
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   12991
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de procesos"
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   330
         Width           =   1365
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   12300
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CheckBox Check46 
      Caption         =   "Realizar procesos ya ejecutados"
      Height          =   195
      Left            =   7860
      TabIndex        =   4
      Top             =   480
      Width           =   2835
   End
   Begin VB.CheckBox Check37 
      Caption         =   "Respetar dependencia de procesos"
      Height          =   195
      Left            =   7830
      TabIndex        =   3
      Top             =   90
      Value           =   1  'Checked
      Width           =   3195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ejecutar 1 proceso pendiente"
      Height          =   650
      Left            =   270
      TabIndex        =   2
      Top             =   900
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11940
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Ejecutar procesos en automatico"
      Height          =   650
      Left            =   3900
      TabIndex        =   0
      Top             =   900
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   13500
      Top             =   90
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   11430
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   34799617
      CurrentDate     =   41736
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de ejecución"
      Height          =   195
      Left            =   270
      TabIndex        =   11
      Top             =   195
      Width           =   1410
   End
End
Attribute VB_Name = "frmEjecucionProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public noreg As Integer
Dim FechaImportacion As Date
Dim VarFecha As Integer



Private Sub B_Click()
End Sub

Private Sub Combo2_Click()
Dim fecha As Date
Dim id_proc As Integer
If Option1.value Then
   id_proc = 1
ElseIf Option2.value Then
   id_proc = 2
End If
If IsDate(Combo2.Text) Then
   fecha = CDate(Combo2.Text)
   Call MostrarEstadoProcesos(fecha, id_proc)
End If
End Sub

Sub MostrarEstadoProcesos(fecha, ByVal opcion As Integer)
Dim matpr() As Variant
Dim i As Integer
Dim noreg As Integer

matpr = ObtProcesosFecha(fecha, opcion)
noreg = UBound(matpr, 1)
frmEjecucionProc.MSFlexGrid1.Clear
frmEjecucionProc.MSFlexGrid1.Visible = True
frmEjecucionProc.MSFlexGrid1.Rows = noreg + 1
frmEjecucionProc.MSFlexGrid1.Cols = 12
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 0) = "ID Proceso"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 1) = "Descripcion"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 2) = "Bloqueada"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 3) = "Finalizada"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 4) = "Comentario"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 5) = "Usuario"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 6) = "Direccion IP"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 7) = "Fecha de inicio"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 8) = "Hora de inicio"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 9) = "Fecha final"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 10) = "Hora final"
frmEjecucionProc.MSFlexGrid1.TextMatrix(0, 11) = "Tiempo de proceso"


frmEjecucionProc.MSFlexGrid1.ColWidth(0) = 1000
frmEjecucionProc.MSFlexGrid1.ColWidth(1) = 3000

For i = 2 To 11
frmEjecucionProc.MSFlexGrid1.ColWidth(i) = 1000
Next i
frmEjecucionProc.MSFlexGrid1.ColWidth(4) = 3000

For i = 1 To noreg
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 0) = ReemplazaVacioValor(matpr(i, 1), "")  'id tarea
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 1) = ReemplazaVacioValor(matpr(i, 3), "")  'Descripcion
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 2) = ReemplazaVacioValor(matpr(i, 29), "") 'bloqueada
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 3) = ReemplazaVacioValor(matpr(i, 30), "") 'finalizada
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 4) = ReemplazaVacioValor(matpr(i, 31), "") 'Comentario
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 5) = ReemplazaVacioValor(matpr(i, 32), "") 'usuario
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 6) = ReemplazaVacioValor(matpr(i, 33), "") 'direccion ip
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 7) = ReemplazaVacioValor(matpr(i, 25), "") 'fecha de inicio
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 8) = ReemplazaVacioValor(matpr(i, 26), "") 'hora de inicio
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 9) = ReemplazaVacioValor(matpr(i, 27), "") 'fecha fin
    frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 10) = ReemplazaVacioValor(matpr(i, 28), "") 'hora fin
    If Not EsVariableVacia(matpr(i, 25)) And Not EsVariableVacia(matpr(i, 26)) And Not EsVariableVacia(matpr(i, 27)) And Not EsVariableVacia(matpr(i, 28)) Then
       frmEjecucionProc.MSFlexGrid1.TextMatrix(i, 11) = TiempoProc(matpr(i, 25), matpr(i, 26), matpr(i, 27), matpr(i, 28))  'tiempo del proceso
    End If

Next i

End Sub

Private Sub Combo2_DblClick()
 Call Combo2_Click
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Combo2_Click
End Sub

Private Sub Command1_Click()
Dim tiempo1 As Date
Dim tiempo2 As Date
Dim tiempo3 As Date
Dim accion1 As Boolean
Dim accion2 As Boolean
Dim fecha As Date
Dim exito As Boolean
Dim id_proc As Integer
If Option1.value Then
   id_proc = 1
ElseIf Option2.value Then
   id_proc = 2
End If

SiActTProc = True
Screen.MousePointer = 11
Command1.Enabled = False
If IsDate(Combo1.Text) Then
   tiempo1 = Now()
   fecha = CDate(Combo1.Text)
   If frmEjecucionProc.Check37.value Then
      accion1 = False
   Else
      accion1 = True
   End If
   If frmEjecucionProc.Check46.value Then
      accion2 = False
   Else
      accion2 = True
   End If
   lista.Enabled = False
   frmProgreso.Show
     Call SecProcesosManual2(fecha, accion1, accion2, id_proc)
     Combo2.Text = fecha
     Call MostrarEstadoProcesos(fecha, id_proc)
   Unload frmProgreso
   lista.Enabled = True
   frmEjecucionProc.SetFocus
tiempo2 = Now()
tiempo3 = tiempo2 - tiempo1
End If
MsgBox "Proceso terminado " & Format(tiempo3, "hh:mm:ss")
Call ActUHoraUsuario
SiActTProc = False
frmEjecucionProc.Command2.Enabled = True
Command1.Enabled = True
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim tiempo1 As Date
Dim tiempo2 As Date
Dim tiempo3 As Date
Dim accion1 As Boolean
Dim accion2 As Boolean
Dim fecha As Date
Dim exito As Boolean
Dim id_proc As Integer

If frmEjecucionProc.Option1.value Then
    id_proc = 1
ElseIf frmEjecucionProc.Option2.value Then
    id_proc = 2
End If
SiActTProc = True
Screen.MousePointer = 11
Command2.Enabled = False
If IsDate(Combo1.Text) Then
   tiempo1 = Now()
   fecha = CDate(Combo1.Text)
   If frmEjecucionProc.Check37.value Then
      accion1 = False
   Else
      accion1 = True
   End If
   If frmEjecucionProc.Check46.value Then
      accion2 = False
   Else
      accion2 = True
   End If
   lista.Enabled = False
   frmProgreso.Show
     Call SecProcesosManual1(fecha, accion1, accion2, id_proc)
     Combo2.Text = fecha
     Call MostrarEstadoProcesos(fecha, id_proc)
   Unload frmProgreso
   lista.Enabled = True
tiempo2 = Now()
tiempo3 = tiempo2 - tiempo1
End If
frmEjecucionProc.SetFocus
MsgBox "Proceso terminado " & Format(tiempo3, "hh:mm:ss")
Call ActUHoraUsuario
SiActTProc = False
frmEjecucionProc.Command2.Enabled = True
Command2.Enabled = True

Screen.MousePointer = 0
End Sub

Private Sub Command33_Click()
Dim nomarch As String
Dim matdatos() As Variant
Dim i As Long
Screen.MousePointer = 11
NoFechas = 54
ReDim matfecha(1 To NoFechas) As Date
matfecha(1) = #8/29/2003#
matfecha(2) = #9/30/2003#
matfecha(3) = #10/31/2003#
matfecha(4) = #11/28/2003#
matfecha(5) = #12/31/2003#
matfecha(6) = #1/30/2004#
matfecha(7) = #2/27/2004#
matfecha(8) = #3/31/2004#
matfecha(9) = #4/30/2004#
matfecha(10) = #5/31/2004#
matfecha(11) = #6/30/2004#
matfecha(12) = #7/30/2004#
matfecha(13) = #8/31/2004#
matfecha(14) = #9/30/2004#
matfecha(15) = #10/29/2004#
matfecha(16) = #11/30/2004#
matfecha(17) = #12/31/2004#
matfecha(18) = #1/31/2005#
matfecha(19) = #2/28/2005#
matfecha(20) = #3/31/2005#
matfecha(21) = #4/29/2005#
matfecha(22) = #5/31/2005#
matfecha(23) = #6/30/2005#
matfecha(24) = #7/29/2005#
matfecha(25) = #8/31/2005#
matfecha(26) = #9/30/2005#
matfecha(27) = #10/31/2005#
matfecha(28) = #11/30/2005#
matfecha(29) = #12/30/2005#
matfecha(30) = #1/31/2006#
matfecha(31) = #2/28/2006#
matfecha(32) = #3/31/2006#
matfecha(33) = #4/28/2006#
matfecha(34) = #5/31/2006#
matfecha(35) = #6/30/2006#
matfecha(36) = #7/31/2006#
matfecha(37) = #8/31/2006#
matfecha(38) = #9/29/2006#
matfecha(39) = #10/31/2006#
matfecha(40) = #11/30/2006#
matfecha(41) = #12/29/2006#
matfecha(42) = #1/31/2007#
matfecha(43) = #2/28/2007#
matfecha(44) = #3/30/2007#
matfecha(45) = #4/30/2007#
matfecha(46) = #5/31/2007#
matfecha(47) = #6/29/2007#
matfecha(48) = #7/31/2007#
matfecha(49) = #8/31/2007#
matfecha(50) = #9/28/2007#
matfecha(51) = #10/31/2007#
matfecha(52) = #11/30/2007#
matfecha(53) = #12/31/2007#
matfecha(54) = #1/31/2008#

nomarch = DirResVaR & "\desc irs 2003.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 1 To 5
 Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
nomarch = DirResVaR & "\desc irs 2004.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 6 To 17
  Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
nomarch = DirResVaR & "\desc irs 2005.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 18 To 29
  Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
nomarch = DirResVaR & "\desc irs 2006.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 30 To 41
  Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
nomarch = DirResVaR & "\desc irs 2007.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 42 To 53
  Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
nomarch = DirResVaR & "\desc irs 2008.txt"
matdatos = LeerArchTexto(nomarch, Chr(9), "Leyendo el archivo " & nomarch)
For i = 54 To 54
  Call CompletarCurvaIRS(matdatos, matfecha(i))
Next i
Screen.MousePointer = 0
End Sub

Sub CompletarCurvaIRS(ByRef based() As Variant, ByVal fecha As Date)
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim i As Long
Dim columna As Long
Dim j As Long
Dim nomarch1 As String


noreg = UBound(based, 1)
columna = UBound(based, 2)
For i = 1 To columna
If CDate(based(1, i)) = fecha Then
 ReDim curva(1 To noreg - 1, 1 To 2) As Variant
 For j = 1 To noreg - 1
 curva(j, 1) = Val(based(j + 1, i))
 curva(j, 2) = Val(based(j + 1, 1))
Next j

nomarch1 = DirCurvas & "\CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
Set base1 = OpenDatabase(nomarch1, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset)
registros1.MoveFirst
For j = 1 To 12000
registros1.Edit
 Call GrabarTAccess(registros1, 39, Val(curva(j, 1)), i)
registros1.Update
registros1.MoveNext
Next j
registros1.Close
base1.Close
Exit For
End If
Next i

End Sub

Private Sub Command35_Click()
Dim nomarch As String
Dim sihayarch As Boolean
Dim fecha1 As Date
Dim fecha2 As Date


Screen.MousePointer = 11
MsgBox "Carga los escenarios definidos como mas extremos en la tabla historica"
Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
nomarch = DirResVaR & "\tabla esc estres.xlsx"
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
fecha1 = #1/4/1996#
fecha2 = #3/11/1997#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #1/4/1996#
fecha2 = #9/2/1998#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #9/15/1999#
fecha2 = #9/15/1999#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #9/12/2001#
fecha2 = #9/12/2001#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #4/16/2002#
fecha2 = #4/16/2002#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #10/30/1997#
fecha2 = #10/30/1997#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #8/27/1998#
fecha2 = #8/27/1998#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #9/10/1998#
fecha2 = #9/10/1998#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #9/11/2001#
fecha2 = #9/11/2001#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
fecha1 = #4/27/2007#
fecha2 = #4/27/2007#
Call CargaEscEstresOra(fecha1, fecha2, nomarch, "t")
Else
 MsgBox "no existe el archivo " & nomarch
End If
Screen.MousePointer = 0
End Sub


Private Sub Command3_Click()
Dim resp As Integer
Dim fecha As Date
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim txtejecuta As String
Dim txttabla As String

If frmEjecucionProc.Option1.value Then
   txttabla = TablaProcesos1
ElseIf frmEjecucionProc.Option2.value Then
   txttabla = TablaProcesos2
End If


Screen.MousePointer = 11
resp = MsgBox("Desea cambiar el estado de algunos procesos", vbYesNo)
If resp = 6 Then
fecha = CDate(Combo2.Text)
txtfecha = "TO_DATE('" & Format(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
noreg = MSFlexGrid1.Rows - 1
ReDim mata(1 To noreg, 1 To 3) As Variant
For i = 1 To noreg
    mata(i, 1) = CLng(MSFlexGrid1.TextMatrix(i, 0))
    mata(i, 2) = MSFlexGrid1.TextMatrix(i, 3)
    mata(i, 3) = MSFlexGrid1.TextMatrix(i, 2)
Next i
For i = 1 To noreg
    txtejecuta = "UPDATE " & txttabla & " SET FINALIZADO ='" & mata(i, 2) & "', BLOQUEADO ='" & mata(i, 3) & "' WHERE FECHAP = " & txtfecha & " AND ID_TAREA = " & mata(i, 1)
    ConAdo.Execute txtejecuta
    'txtejecuta = "UPDATE " & txttabla & " SET BLOQUEADO ='" & mata(i, 3) & "' WHERE FECHAP = " & txtfecha & " AND ID_TAREA = " & mata(i, 1)
    'ConAdo.Execute txtejecuta
Next i
MsgBox "Se cambio el estados de algunos procesos"
End If
Screen.MousePointer = 0
End Sub

Private Sub Command45_Click()
If Command45.Caption = "Ejecutar procesos en automatico" Then
  Timer1.Enabled = True
  Command45.Caption = "Desabilitar ejecucion"
  SiActTProc = True
ElseIf Command45.Caption = "Desabilitar ejecucion" Then
  Timer1.Enabled = False
  Command45.Caption = "Ejecutar procesos en automatico"
  Call ActUHoraUsuario
  SiActTProc = False
End If
End Sub


Private Sub Command46_Click()
Screen.MousePointer = 11
   Call CopiaBaseOracle("var_flujos_deuda3", TablaFlujosDeudaO)
Screen.MousePointer = 0
End Sub

Private Sub Command48_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim i As Long
Dim noreg1 As Long
Dim nomarch As String
Dim sihayarch As Boolean
Dim mata() As Variant
Dim exitoarch As Boolean

Screen.MousePointer = 11
fecha1 = #1/1/2008#
fecha2 = #4/2/2009#
fecha = fecha1
ReDim matc(1 To 12001) As Variant
matc(1) = "Plazo" & Chr(9)
For i = 1 To 12000
matc(i + 1) = i & Chr(9)
Next i

Do While fecha <= fecha2
nomarch = DirResVaR & "\jtsmed" & Format(fecha, "yyyymmdd") & "_final.txt"
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
mata = LeerArchTexto(nomarch, ",", "Leyendo ")
noreg1 = UBound(mata, 1)
matc(1) = matc(1) & Format(fecha, "dd/mm/yyyy") & Chr(9)
For i = 1 To 12000
If i <= noreg1 Then
matc(i + 1) = matc(i + 1) & mata(i, 2) * 100 & Chr(9)
Else
matc(i + 1) = matc(i + 1) & 0 & Chr(9)
End If
Next i
End If
fecha = fecha + 1
Loop
Call VerificarSalidaArchivo(DirResVaR & "\salida.txt", 1, exitoarch)
If exitoarch Then
For i = 1 To 12001
Print #1, matc(i)
Next i
Close #1
End If
Screen.MousePointer = 0
End Sub


Private Sub Command49_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena As String
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
fecha = #6/19/2013#
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro = "SELECT * FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha
 rmesa.Open txtfiltro, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 rmesa.MoveNext
 Next i
 rmesa.Close
 fecha1 = fecha + 1
 txtfecha1 = "TO_DATE('" & Format(fecha1, "dd/mm/yyyy") & "','DD/MM/YYYY')"
 txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha1
 ConAdo.Execute txtborra
 For i = 1 To noreg
     txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
     txtcadena = txtcadena & txtfecha1 & ","            'fecha
     txtcadena = txtcadena & "'" & mata(i, 2) & "',"    'concepto
     txtcadena = txtcadena & mata(i, 3) & ","           'plazo
     txtcadena = txtcadena & mata(i, 4) & ","           'valor
     txtcadena = txtcadena & "'" & mata(i, 5) & "')"    'indice
     ConAdo.Execute txtcadena
  Next i
End If
MsgBox "fin de proceso"
Screen.MousePointer = 0

End Sub


Sub GenerarFRLiquidez(ByVal fecha As Date)
Dim icete28 As Integer
Dim icete91 As Integer
Dim icete182 As Integer
Dim icete364 As Integer
Dim itiie28 As Integer
Dim itiie91 As Integer
Dim IUDI As Integer
Dim iccp As Integer
Dim icpp As Integer
Dim exito As Boolean
Dim i As Long
Dim indice As Long
Dim valorudi As Double
Dim noreg As Long
Dim nomarch2 As String
Dim txtcadena As String
Dim nodat1 As Long
Dim matcurvas() As Variant
Dim exitoarch As Boolean

icete28 = 0: icete91 = 0: icete182 = 0: itiie28 = 0: itiie91 = 0: IUDI = 0
For i = 1 To NoFactores
 If MatCaracFRiesgo(i).indFactor = "CETES IMP 28" Then
  icete28 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "CETES IMP 91" Then
  icete91 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "CETES IMP 182" Then
  icete182 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "CETES IMP 360" Then
  icete364 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "TIIE 28 0" Then
  itiie28 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "TIIE 91 0" Then
  itiie91 = i
 ElseIf MatCaracFRiesgo(i).indFactor = "UDI 0" Then
  IUDI = i
 ElseIf MatCaracFRiesgo(i).indFactor = "CCP 0" Then
  iccp = i
 ElseIf MatCaracFRiesgo(i).indFactor = "CPP 0" Then
  icpp = i
 End If
Next i
If icete28 = 0 Or icete91 = 0 Or icete182 = 0 Or itiie28 = 0 Or itiie91 = 0 Or IUDI = 0 Then
 MsgBox "no se tienen todos los datos para calcular la proyeccion de tasas"
Else
indice = BuscarValorArray(fecha, MatFactRiesgo, 1)
If indice <> 0 Then
valorudi = MatFactRiesgo(indice, IUDI + 1)
'se lee el archivo de curvas de ese día
matcurvas = LeerCurvaCompleta(fecha, exito)
noreg = 12000
ReDim mata(1 To noreg, 1 To 4)
For i = 1 To noreg
 mata(i, 1) = i                                  'plazo
 mata(i, 2) = matcurvas(i + 1, 20)                    'curva de cetes imp
 mata(i, 3) = matcurvas(i + 1, 21)                    'curva de tiie
 mata(i, 4) = matcurvas(i + 1, 44)                    'curva tasa real
Next i

'se procede a calcular las tasas futuras que se necesitan

nomarch2 = DirResVaR & "\Tasas alm " & Format(fecha, "yyyy-mm-dd") & ".txt"
CommonDialog1.FileName = nomarch2
CommonDialog1.ShowSave
nomarch2 = CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch2, 1, exitoarch)
If exitoarch Then
txtcadena = "FECHA" & Chr(9)
txtcadena = txtcadena & "CETES 28" & Chr(9)
txtcadena = txtcadena & "CETES 91" & Chr(9)
txtcadena = txtcadena & "CETES 182" & Chr(9)
txtcadena = txtcadena & "TIIE 28" & Chr(9)
txtcadena = txtcadena & "TIIE 91" & Chr(9)
txtcadena = txtcadena & "UDI" & Chr(9)
txtcadena = txtcadena & "CCP" & Chr(9)
txtcadena = txtcadena & "CPP"
Print #1, txtcadena
txtcadena = ""
'la fecha
txtcadena = txtcadena & MatFactRiesgo(indice, 1) & Chr(9)
'cetes 28 dias
txtcadena = txtcadena & MatFactRiesgo(indice, icete28 + 1) & Chr(9)
'cetes 91 dias
txtcadena = txtcadena & MatFactRiesgo(indice, icete91 + 1) & Chr(9)
'cetes 182 dias
txtcadena = txtcadena & MatFactRiesgo(indice, icete182 + 1) & Chr(9)
'tiie 28
txtcadena = txtcadena & MatFactRiesgo(indice, itiie28 + 1) & Chr(9)
'tiie 91
txtcadena = txtcadena & MatFactRiesgo(indice, itiie91 + 1) & Chr(9)
'udi
txtcadena = txtcadena & MatFactRiesgo(indice, IUDI + 1) & Chr(9)
'ccp
txtcadena = txtcadena & MatFactRiesgo(indice, iccp + 1) & Chr(9)
'cpp
txtcadena = txtcadena & MatFactRiesgo(indice, icpp + 1)
Print #1, txtcadena

nodat1 = noreg - 182
ReDim tasaf(1 To nodat1, 1 To 7) As Variant
For i = 1 To nodat1
 tasaf(i, 1) = fecha + i
 'se generan las tasas futuras cuando existan datos
 If mata(i + 28, 2) <> 0 And mata(i, 2) <> 0 Then
  tasaf(i, 2) = ((1 + mata(i + 28, 2) * (i + 28) / 360) / (1 + mata(i, 2) * i / 360) - 1) * 360 / 28
 Else
  tasaf(i, 2) = tasaf(i - 1, 2)
 End If
 If mata(i + 91, 2) <> 0 And mata(i, 2) <> 0 Then
  tasaf(i, 3) = ((1 + mata(i + 91, 2) * (i + 91) / 360) / (1 + mata(i, 2) * i / 360) - 1) * 360 / 91
 Else
  tasaf(i, 3) = tasaf(i - 1, 3)
 End If
 If mata(i + 182, 2) <> 0 And mata(i, 2) <> 0 Then
  tasaf(i, 4) = ((1 + mata(i + 182, 2) * (i + 182) / 360) / (1 + mata(i, 2) * i / 360) - 1) * 360 / 182
 Else
  tasaf(i, 4) = tasaf(i - 1, 4)
 End If
 If mata(i + 28, 3) <> 0 And mata(i, 3) <> 0 Then
  tasaf(i, 5) = ((1 + mata(i + 28, 3) * (i + 28) / 360) / (1 + mata(i, 3) * i / 360) - 1) * 360 / 28
 Else
  tasaf(i, 5) = tasaf(i - 1, 5)
 End If
 If mata(i + 182, 3) <> 0 And mata(i, 3) <> 0 Then
  tasaf(i, 6) = ((1 + mata(i + 182, 3) * (i + 182) / 360) / (1 + mata(i, 3) * i / 360) - 1) * 360 / 182
 Else
  tasaf(i, 6) = tasaf(i - 1, 6)
 End If
 If mata(i, 2) <> 0 And mata(i, 4) <> 0 Then
  tasaf(i, 7) = valorudi * (1 + mata(i, 2) * i / 360) / (1 + mata(i, 4) * i / 360)
 Else
  tasaf(i, 7) = tasaf(i - 1, 7)
 End If
 txtcadena = ""
 txtcadena = txtcadena & tasaf(i, 1) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 2) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 3) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 4) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 5) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 6) & Chr(9)
 txtcadena = txtcadena & tasaf(i, 7)
 Print #1, txtcadena
Next i
Close #1
MensajeProc = "Se genero el archivo " & nomarch2
MsgBox MensajeProc
End If
Else
MsgBox "Falta la fecha solicitada en la tabla de factores de riesgo"
End If
End If
End Sub

Private Sub Command60_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fecha As Date
Dim fechan As Date
Dim fechan1 As Date
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim nomarch As String
Dim txtcadena As String
Dim nocampos As Long
Dim rmesa As New ADODB.recordset
Dim exitoarch As Boolean

Screen.MousePointer = 11


 If ActivarControlErrores Then
 On Error GoTo ControlErrores
 End If
 fecha = CDate(Combo1.Text)
 fechan = Val(Format(fecha, "yyyymmdd"))
 
 '====================================================
 txtfiltro1 = "select count(*) from LIQ_T_TABLA_AMORT WHERE FECHA_PROCESO = " & fechan
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
 txtfiltro2 = "select * from LIQ_T_TABLA_AMORT WHERE FECHA_PROCESO = " & fechan
 rmesa.Open txtfiltro2, ConAdo
 rmesa.MoveFirst
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos + 1) As Variant
 For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 mata(i, nocampos + 1) = "C" & Format(mata(i, 1), "0000000") & "F" & Format(mata(i, 3), "0000000")
 rmesa.MoveNext
 Next i
 rmesa.Close
 mata = RutinaOrden(mata, nocampos + 1, 1)
 ReDim matb(1 To noreg, 1 To 5) As Variant
 For i = 1 To noreg
     matb(i, 1) = mata(i, 1)
     If i <> 1 Then
        If mata(i - 1, 1) = mata(i, 1) Then
           matb(i, 2) = mata(i - 1, 4)
        Else
           matb(i, 2) = mata(i, 4) - 28
        End If
     Else
        matb(i, 2) = mata(i, 4) - 28
     End If
     matb(i, 3) = mata(i, 4)    'fecha corte
     matb(i, 4) = mata(i, 5)    'saldo
     If i <> noreg Then
        If mata(i, 1) <> mata(i + 1, 1) Then
           matb(i, 5) = mata(i, 5)
        Else
           matb(i, 5) = mata(i, 5) - mata(i + 1, 5)
        End If
     Else
        matb(i, 5) = mata(i, 5)
     End If
 
 Next i
   fechan1 = Val(Format(fecha, "yyyy-mm-dd"))
   nomarch = DirResVaR & "\Tabla flujos SIC " & fechan1 & ".txt"
   nomarch = DirResVaR & "\Tabla flujos SIC " & fechan1 & ".txt"
   Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
   If exitoarch Then
   txtcadena = "Credito" & Chr(9) & "Fecha Inicio flujo" & Chr(9) & "Fecha Fin flujo" & Chr(9) & "Saldo" & Chr(9) & "Amort"
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To 5
   txtcadena = txtcadena & matb(i, j) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se genero el archivo " & nomarch
   MsgBox "Fin de proceso"
   End If
 Else
   ReDim mata(0 To 0, 0 To 0) As Variant
 End If
   
 On Error GoTo 0
Screen.MousePointer = 0
Exit Sub
ControlErrores:
 MsgBox error(Err())
 On Error GoTo 0
Screen.MousePointer = 0
End Sub

Function LeerArchExcel(ByVal nomarch As String, ByVal nomtabla As String) As Variant()
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset(nomtabla, dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
   mata(i, j) = LeerTAccess(registros1, j - 1, i)
  Next j
 registros1.MoveNext
 Next i
 Else
 ReDim mata(0 To 0, 0 To 0) As Variant
 End If
 registros1.Close
 base1.Close
 LeerArchExcel = mata
End Function

Function CSerie(ByVal texto As String)
If texto = "9" Then
   CSerie = "09"
Else
   CSerie = texto
End If
End Function

Function transPosPension(ByRef mata() As Variant) As Variant()
Dim noreg As Long
Dim i As Long
Dim fecha As Date
Dim f_val As Date
Dim serie As String
Dim emision As String
Dim cemision As String
Dim indice As Long
Dim toperacion As Integer
Dim fcompra As Date
Dim fvence As Date
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim cposicion As Integer

noreg = UBound(mata, 1)

ReDim matb(1 To noreg, 1 To 14) As Variant
For i = 1 To noreg
    fecha = ConvertirTextoFecha(mata(i, 1), 0) 'fecha de compra
    If f_val <> fecha Or EsArrayVacio(matvp) Then
       matvp = LeerVPrecios(fecha, mindvp)
       f_val = fecha
    End If
    serie = CSerie(mata(i, 4))
    If mata(i, 2) = "IM" Or mata(i, 2) = "IQ" Or mata(i, 2) = "M" Or mata(i, 2) = "PI" Or mata(i, 2) = "IS" Or mata(i, 2) = "LD" Or mata(i, 2) = "S" Or mata(i, 2) = "BI" Then
       emision = ""
    Else
      emision = mata(i, 3)
    End If
    cemision = mata(i, 2) & emision & serie
    indice = BuscarValorArray(cemision, matvp, 18)
    If indice <> 0 Then
    matb(i, 1) = fecha            'fecha
    matb(i, 2) = "N"              'intencion
    If mata(i, 5) = 0 Then
       toperacion = 1
    Else
       toperacion = 2
    End If
    matb(i, 3) = toperacion      'tipo de operacion
    matb(i, 4) = mata(i, 2)      'tipo valor
    matb(i, 5) = mata(i, 3)      'emision
    matb(i, 6) = mata(i, 4)      'serie
    matb(i, 7) = cemision        'clave de emision
    matb(i, 8) = mata(i, 6)      'no de titulos
    If Not EsVariableVacia(mata(i, 9)) Then
      fcompra = ConvertirTextoFecha(mata(i, 9), 0)
    Else
      fcompra = 0
    End If
    matb(i, 9) = fcompra              'fecha de compra
    matb(i, 10) = mata(i, 13)         'p compra
    If Not EsVariableVacia(mata(i, 7)) Then
    matb(i, 11) = Val(mata(i, 7))  'tasa premio
    Else
    matb(i, 11) = 0
    End If
    If toperacion = 1 Then
       fvence = matvp(indice, 12)
    Else
       fvence = ConvertirTextoFecha(mata(i, 9), 0)
    End If
    matb(i, 12) = fvence         'fecha vencimiento
    If Val(mata(i, 12)) = 2065 Then
       cposicion = 12
    Else
       cposicion = 13
    End If
    matb(i, 13) = cposicion         'clave de posicion
    matb(i, 14) = mata(i, 11)       'clave del subportafolio
    Else
    MsgBox "no se encontro la emision en el vector de precios " & cemision
    End If
Next i
transPosPension = matb
End Function



Private Sub Form_Load()
Dim i As Long
Dim matfechasp() As Date

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
RespetarSecProc = True
 frmEjecucionProc.Left = (Screen.Width - frmEjecucionProc.Width) / 2
 frmEjecucionProc.top = (Screen.Height - frmEjecucionProc.Height) / 2
 SiExisteVecPrecios = False
 SiExisteCurvas = False
 MonthView2.value = Date
       If OpcionBDatos = 1 Then
          frmEjecucionProc.Caption = "Ejecución de Procesos (Producción)"
       ElseIf OpcionBDatos = 2 Then
          frmEjecucionProc.Caption = "Ejecución de Procesos (Desarrollo)"
       ElseIf OpcionBDatos = 3 Then
          frmEjecucionProc.Caption = "Ejecución de Procesos (DRP)"
       End If
       matfechasp = FechasProcG(1)
       noreg = UBound(matfechasp, 1)
    For i = 1 To noreg
      Combo1.AddItem matfechasp(i, 1)
      Combo2.AddItem matfechasp(i, 1)
    Next i
 
Screen.MousePointer = 0
On Error GoTo 0
frmEjecucionProc.Show
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
 MatFechasVaR = LeerFechasVaRT()
  frmCalVar.Show
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim txtmsg As String
 Select Case State
        Case 0
           txtmsg = "Sin estado que reportar"
        Case 1
           txtmsg = "Estableciendo una conexion"
        Case icHostResolved
           txtmsg = icHostResolved
        Case icConnecting
          txtmsg = "Conectandose al servidor"
        Case icConnected
          txtmsg = "Conectado al servidor"
        Case icRequesting
          txtmsg = "Haciendo solicitud al servidor"
        Case icRequestSent
          txtmsg = "solicitud enviada"
        Case icReceivingResponse
          txtmsg = "Recibiendo respuesta"
        Case icResponseReceived
           txtmsg = "respuesta recibida"
        Case icDisconnecting
           txtmsg = "Desconectandose del servidor"
        Case icDisconnected
           txtmsg = "Desconectado del servidor"
        Case icError
           If Inet1.ResponseCode = 12003 Then
              siacarchftp = False
           ElseIf Inet1.ResponseCode = 12300 Then
              siacarchftp = False
           ElseIf Inet1.ResponseCode = 80 Then
              siacarchftp = False
           ElseIf Inet1.ResponseCode = 123 Then
              siacarchftp = False
           End If
           txtmsg = Inet1.ResponseCode & " " & Inet1.ResponseInfo
        Case icResponseCompleted
          txtmsg = "Proceso finalizado"
        End Select
        txtmsgFTP = txtmsg
        
End Sub

Sub CargarArchivoWeb()
Dim bDone As Boolean
Dim FileSize As Long
Dim vtData As Variant
Dim tempArray As Variant
 bDone = False
            'Para saber el tamaño del fichero en bytes
            'Creamos y abrimos un nuevo archivo en modo binario
            Open "d:\salida.zip" For Binary As #1
          
            ' Leemos de a 1 Kbytes. El segundo parámetro indica _
            el tipo de fichero. Tipo texto o tipo Binario, en este caso binario
            vtData = Inet1.GetChunk(1024, icByteArray)
            DoEvents
        
            'Si el tamaño del fichero es 0 ponemos bDone en _
            True para que no entre en el bucle
            If Len(vtData) = 0 Then
                bDone = True
            End If
              
              
            Do While Not bDone
                'Almacenamos en un array el contenido del archivo que se va leyendo
                tempArray = vtData
                'Escribimos los datos en el archivo
                Put #1, , tempArray
                'Leemos  datos de a 1 kb (1024 bytes)
                vtData = Inet1.GetChunk(1024, icByteArray)
           
                DoEvents
                'Aumentamos la barra de progreso
              
                If Len(vtData) = 0 Then
                    bDone = True
                End If
            Loop
  
        Close #1
        MsgBox "Archivo descargado correctamente", vbInformation
  
End Sub

Sub Mostrar_Lista(lista As String)
    Dim i As Long
    Dim Archivos() As String
    Dim Sfile As String
      
    'Desgloza la cadena y la almacena en un vector
    Archivos = Split(lista, vbCrLf)
          
    'REcorre el vector
    For i = 0 To UBound(Archivos)
          
        If Archivos(i) <> vbNullString Then
              
            ' ..es un directorio
            If Right(Archivos(i), 1) = "/" Then
                frmCalVar.Print "[ Directorio ] " & Archivos(i)
            Else
            '..  es un archivo
                frmCalVar.Print "[ Archivo ] " & Archivos(i)
            End If
        End If
    Next
End Sub


Private Sub TabStrip1_Click()

End Sub

Private Sub Combo1_DblClick()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 MonthView2.Visible = True
 MonthView2.Left = 200
 MonthView2.top = 800
 MonthView2.ZOrder 0
 VarFecha = 1
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub



Private Sub MonthView2_DateDblClick(ByVal DateDblClicked As Date)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If VarFecha = 1 Then
Combo1.Text = CDate(DateDblClicked)
ElseIf VarFecha = 2 Then
Combo2.Text = CDate(DateDblClicked)
End If
MonthView2.Visible = False
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub MSFlexGrid1_DblClick()
Dim indice As Integer
indice = MSFlexGrid1.row
If (MSFlexGrid1.col = 2 Or MSFlexGrid1.col = 3) And indice >= 1 Then
If MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "S" Then
   MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "N"
ElseIf MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "N" Then
   MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "S"
End If
End If

End Sub

Private Sub Option1_Click()
Dim noreg As Integer
Dim matfechasp() As Date
Dim i As Integer
       matfechasp = FechasProcG(1)
       noreg = UBound(matfechasp, 1)
       Combo1.Clear
       Combo2.Clear
    For i = 1 To noreg
      Combo1.AddItem matfechasp(i, 1)
      Combo2.AddItem matfechasp(i, 1)
    Next i

End Sub

Private Sub Option2_Click()
Dim noreg As Integer
Dim matfechasp() As Date
Dim i As Integer
       matfechasp = FechasProcG(2)
    noreg = UBound(matfechasp, 1)
    Combo1.Clear
    Combo2.Clear
    For i = 1 To noreg
      Combo1.AddItem matfechasp(i, 1)
      Combo2.AddItem matfechasp(i, 1)
    Next i

End Sub

Private Sub Timer1_Timer()
Dim accion1 As Boolean
Dim accion2 As Boolean
If Check37.value = 0 Then
   accion1 = True
Else
   accion1 = False
End If
If Check46.value = 0 Then
   accion2 = True
Else
   accion2 = False
End If

If frmEjecucionProc.Enabled Then
   Screen.MousePointer = 11
   frmProgreso.Show
   Call SecProcesosAuto(accion1, accion2, 1)
   Unload frmProgreso
   Screen.MousePointer = 0
End If
End Sub

