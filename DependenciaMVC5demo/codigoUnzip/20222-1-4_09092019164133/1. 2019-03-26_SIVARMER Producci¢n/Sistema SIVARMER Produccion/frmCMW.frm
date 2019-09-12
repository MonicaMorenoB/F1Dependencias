VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCMW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make whole"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Obtener p&g por operacion"
      Height          =   700
      Left            =   7710
      TabIndex        =   20
      Top             =   7500
      Width           =   2000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12150
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tasa a buscar"
      Height          =   825
      Left            =   200
      TabIndex        =   17
      Top             =   6510
      Width           =   3615
      Begin VB.OptionButton Option4 
         Caption         =   "Pasiva"
         Height          =   195
         Left            =   2010
         TabIndex        =   19
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Activa"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   390
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subprocesos"
      Height          =   675
      Left            =   8700
      TabIndex        =   14
      Top             =   6330
      Width           =   3765
      Begin VB.OptionButton Option2 
         Caption         =   "Subprocesos 2"
         Height          =   195
         Left            =   1950
         TabIndex        =   16
         Top             =   300
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Subprocesos 1"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Value           =   -1  'True
         Width           =   1605
      End
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   7700
      TabIndex        =   11
      Top             =   5900
      Width           =   2500
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6500
      TabIndex        =   10
      Top             =   5900
      Width           =   1000
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4700
      TabIndex        =   8
      Top             =   5900
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2000
      TabIndex        =   5
      Top             =   5900
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   200
      TabIndex        =   4
      Top             =   5900
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Obtener resultados"
      Height          =   700
      Left            =   5160
      TabIndex        =   3
      Top             =   7500
      Width           =   2000
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   195
      TabIndex        =   2
      Top             =   450
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8705
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Determinar sobretasa de equilibrio"
      Height          =   700
      Left            =   200
      TabIndex        =   1
      Top             =   7500
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear subprocesos make whole"
      Height          =   700
      Left            =   2640
      TabIndex        =   0
      Top             =   7500
      Width           =   2000
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Clave de operación"
      Height          =   195
      Left            =   7700
      TabIndex        =   13
      Top             =   5600
      Width           =   1380
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Clave posicion"
      Height          =   195
      Left            =   6500
      TabIndex        =   12
      Top             =   5600
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hora reg"
      Height          =   195
      Left            =   4700
      TabIndex        =   9
      Top             =   5600
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nombre pos"
      Height          =   195
      Left            =   2000
      TabIndex        =   7
      Top             =   5600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de registro"
      Height          =   225
      Left            =   200
      TabIndex        =   6
      Top             =   5600
      Width           =   1365
   End
End
Attribute VB_Name = "frmCMW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Screen.MousePointer = 11
Dim fecha As Date
Dim fecha0 As Date
Dim dtfechar As Date
Dim txtfecha As String
Dim txtcadena As String
Dim id_proc As Integer
Dim coperacion As String
Dim cposicion As Integer
Dim tipopos As Integer
Dim txtnompos As String
Dim horareg As String
Dim txtfiltro As String
Dim contar As Long
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda

Dim exito As Boolean
Dim matfechas1() As Date
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim htiempo As Integer
Dim noreg As Integer
Dim id_tabla As Integer
Dim fechareg As Date
Dim txtmsg As String

If Option1.value Then
   id_tabla = 1
Else
   id_tabla = 2
End If
id_proc = 90
htiempo = 251
fecha0 = #12/31/2008#
If IsDate(Text2.Text) Then
   tipopos = 2
   fechareg = CDate(Text2.Text)
   fecha = fechareg
   txtnompos = Text3.Text
   horareg = Text4.Text
   cposicion = Val(Text5.Text)
   coperacion = Text6.Text
   contar = DeterminaMaxRegSubproc(id_tabla)
   mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
   Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg, exito)
   If UBound(matpos, 1) <> 0 Then
       matfechas1 = GenPartFechasEsc(fecha0, fecha, 50)
       txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       ConAdo.Execute "DELETE FROM " & TablaPLEscMW & " WHERE FECHA = " & txtfecha & " AND COPERACION = '" & coperacion & "'"
       For j = 1 To UBound(matfechas1, 1)
           contar = contar + 1
           txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Calculo de pyg futuras", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, htiempo, matfechas1(j, 1), matfechas1(j, 2), j, "", "", id_tabla)
           ConAdo.Execute txtcadena
       Next j
   End If
   MsgBox "Fin de proceso"
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim tipopos As Integer
Dim fecha As Date
Dim fechar As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
 Dim tasa As Double
Dim txtcadena As String
Dim exito As Boolean
Dim exito1 As Boolean

Dim txtmsg As String
Screen.MousePointer = 11
tipopos = 2
fechar = CDate(Text2.Text)
txtnompos = Text3.Text
horareg = Text4.Text
cposicion = Val(Text5.Text)
coperacion = Text6.Text
fecha = fechar

If Option3.value Then
      txtcadena = "UPDATE " & TablaPosSwaps & " SET ST_ACTIVA = 0"
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      'txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & 0
      'txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      'txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      'txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      'txtcadena = txtcadena & " AND TPATA = 'B'"
      'ConAdo.Execute txtcadena
   tasa = DetermTasaEquilibrio(fecha, tipopos, fechar, txtnompos, horareg, cposicion, coperacion, 1)
   MsgBox "la tasa es del " & Format(tasa, "##0.000000 %")
   If tasa <> 0 Then
      txtcadena = "UPDATE " & TablaPosSwaps & " SET ST_ACTIVA = " & tasa
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      'txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
      'txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      'txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      'txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      'txtcadena = txtcadena & " AND TPATA = 'B'"
      'ConAdo.Execute txtcadena
   End If
Else
      txtcadena = "UPDATE " & TablaPosSwaps & " SET ST_PASIVA = 0"
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      'txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & 0
      'txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      'txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      'txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      'txtcadena = txtcadena & " AND TPATA = 'C'"
      'ConAdo.Execute txtcadena
   tasa = DetermTasaEquilibrio(fecha, tipopos, fechar, txtnompos, horareg, cposicion, coperacion, 2)
   MsgBox "la tasa es del " & Format(tasa, "##0.000000 %")
   If tasa <> 0 Then
      txtcadena = "UPDATE " & TablaPosSwaps & " SET ST_PASIVA = " & tasa
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      'txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
      'txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      'txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      'txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      'txtcadena = txtcadena & " AND TPATA = 'C'"
      'ConAdo.Execute txtcadena
   End If

End If
MsgBox "Fin de proceso"



Screen.MousePointer = 0

End Sub

Private Sub Command3_Click()
Dim fecha As Date
Dim coperacion As String
Dim fecha_f As Date
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim txtnomarch As String

Screen.MousePointer = 11
frmProgreso.Show
fecha = CDate(Text2.Text)
coperacion = Text6.Text
mata = ObtPyGTasaMRef(fecha, coperacion, 0.05)
If UBound(mata, 1) <> 0 Then
   txtnomarch = "d:\escenarios " & coperacion & ".txt"
   CommonDialog1.FileName = txtnomarch
   CommonDialog1.ShowSave
   txtnomarch = CommonDialog1.FileName
   Open txtnomarch For Output As #1
   For i = 1 To UBound(mata, 1)
       txtcadena = ""
       For j = 1 To UBound(mata, 2)
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
Unload frmProgreso
Screen.MousePointer = 0
MsgBox "Fin de proceso"

End Sub

Private Sub Command4_Click()
Dim fecha As Date
Dim coperacion As String
Dim fecha_f As Date
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim txtnomarch As String

Screen.MousePointer = 11
frmProgreso.Show
fecha = CDate(Text2.Text)
coperacion = Text6.Text
mata = ObtPyGTasaMRef2(fecha, coperacion)
If UBound(mata, 1) <> 0 Then
   txtnomarch = "d:\escenarios " & coperacion & ".txt"
   CommonDialog1.FileName = txtnomarch
   CommonDialog1.ShowSave
   txtnomarch = CommonDialog1.FileName
   Open txtnomarch For Output As #1
   For i = 1 To UBound(mata, 1)
       txtcadena = ""
       For j = 1 To UBound(mata, 2)
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
Unload frmProgreso
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub Form_Load()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset
Dim noreg As Integer

txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 2"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 6) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("TIPOPOS")
       mata(i, 2) = rmesa.Fields("FECHAREG")
       mata(i, 3) = rmesa.Fields("NOMPOS")
       mata(i, 4) = rmesa.Fields("HORAREG")
       mata(i, 5) = rmesa.Fields("CPOSICION")
       mata(i, 6) = rmesa.Fields("COPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Rows = noreg + 1
   MSFlexGrid1.Cols = 1
   MSFlexGrid1.Cols = 6
   MSFlexGrid1.ColWidth(0) = 1000
   MSFlexGrid1.ColWidth(1) = 1000
   MSFlexGrid1.ColWidth(2) = 3000
   MSFlexGrid1.ColWidth(3) = 1000
   MSFlexGrid1.ColWidth(4) = 1000
   MSFlexGrid1.ColWidth(5) = 3000
   For i = 1 To noreg
       For j = 1 To 6
           MSFlexGrid1.TextMatrix(i, j - 1) = mata(i, j)
       Next j
   Next i
End If

End Sub


Private Sub MSFlexGrid1_DblClick()
Dim indice1 As Integer
Dim indice2 As Integer
indice1 = MSFlexGrid1.MouseRow
indice2 = MSFlexGrid1.MouseCol
If indice1 <> 0 Then
   Text2.Text = MSFlexGrid1.TextMatrix(indice1, 1)
   Text3.Text = MSFlexGrid1.TextMatrix(indice1, 2)
   Text4.Text = MSFlexGrid1.TextMatrix(indice1, 3)
   Text5.Text = MSFlexGrid1.TextMatrix(indice1, 4)
   Text6.Text = MSFlexGrid1.TextMatrix(indice1, 5)
End If

End Sub
