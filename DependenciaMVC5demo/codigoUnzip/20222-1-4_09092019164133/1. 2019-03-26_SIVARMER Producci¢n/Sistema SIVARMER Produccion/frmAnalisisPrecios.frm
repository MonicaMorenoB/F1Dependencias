VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia de precios de proveedor"
   ClientHeight    =   8160
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   11880
   Icon            =   "frmAnalisisPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   7788
      Left            =   72
      TabIndex        =   0
      Top             =   144
      Width           =   11652
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   1536
         TabIndex        =   4
         Top             =   480
         Width           =   2685
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar resultados archivo texto"
         Height          =   516
         Left            =   6210
         TabIndex        =   3
         Top             =   420
         Width           =   1572
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1188
      End
      Begin VB.ComboBox Combo3 
         Height          =   288
         Left            =   4416
         TabIndex        =   1
         Top             =   480
         Width           =   1308
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6360
         Left            =   108
         TabIndex        =   5
         Top             =   1212
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   11218
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Historia de la Emisión"
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   936
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "EMISIÓN"
         Height          =   192
         Left            =   1560
         TabIndex        =   8
         Top             =   252
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TIPO VALOR"
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   252
         Width           =   948
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "SERIE"
         Height          =   192
         Left            =   4464
         TabIndex        =   6
         Top             =   216
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmHistPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MatEmision() As String
Dim MatSerie() As String
Dim MatPrecio() As Variant

Private Sub Combo1_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se filtra la tabla y se agrupan los precios por
'emision. Dependiendo de que emision se seleccione se
Screen.MousePointer = 11
Combo3.Clear
txtfiltro1 = "SELECT COUNT(DISTINCT SERIE) FROM " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "'"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
txtfiltro2 = "Select SERIE from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "' GROUP BY SERIE ORDER BY SERIE"
rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 3) As Variant
rmesa.MoveFirst
ReDim MatSerie(1 To noreg) As String
For i = 1 To noreg
MatSerie(i) = rmesa.Fields("SERIE")
Combo3.AddItem MatSerie(i)
rmesa.MoveNext
Next i
rmesa.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub LlenarSerie(ByVal serie As String)
Dim txtfiltro As String
Dim noreg As Integer
Dim txtcadena As String
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'este es el segundo filtro,de aqui se procede a
'filtrar la informacion por emision y serie
Screen.MousePointer = 11
txtfiltro = "SELECT COUNT(DISTINCT FECHA) FROM " & TablaVecPrecios & " WHERE EMISION = '" & Combo1.Text & "' AND SERIE = '" & serie & "' ORDER BY Fecha"
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
txtfiltro = "SELECT FECHA, EMISION, SERIE, MIN(INTERESESMD) AS PINT, MIN(PCUPON) AS PPCUPON, MIN(PSUCIO) AS PPSUCIO, MIN(PLIMPIO) AS PPLIMPIO, MIN(DVENCIMIENTO) AS PDV, MIN(TASASOBRET) AS PTASA, MIN(VNOMINAL) AS PVNOMINAL, MIN(TCUPON) AS PTCUPON FROM " & TablaVecPrecios & " WHERE EMISION = '" & Combo1.Text & "' AND SERIE = '" & serie & "' GROUP BY FECHA, EMISION, SERIE ORDER BY Fecha"
rmesa.Open txtfiltro, ConAdo
rmesa.MoveFirst
ReDim MatPrecio(1 To noreg, 1 To 12) As Variant
'MSFlexGrid1.Rows = Maximo(MSFlexGrid1.Rows, noreg + 2)

txtcadena = serie & " - FECHA" & Chr(9)
txtcadena = txtcadena & serie & " - PRECIO SUCIO MD" & Chr(9)
txtcadena = txtcadena & serie & " - PRECIO LIMPIO MD" & Chr(9)
txtcadena = txtcadena & serie & " - TASA CUPON" & Chr(9)
txtcadena = txtcadena & serie & " - DIAS VENCIMIENTO" & Chr(9)
txtcadena = txtcadena & serie & " - PERIODO CUPON" & Chr(9)
txtcadena = txtcadena & serie & " - TASA DE REFERENCIA" & Chr(9)
txtcadena = txtcadena & serie & " - TASA / SOBRETASA"
Print #1, txtcadena

'MSFlexGrid1.TextMatrix(0, 8 * valK - 7) = serie & " - FECHA"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 6) = serie & " - PRECIO SUCIO MD"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 5) = serie & " - PRECIO LIMPIO MD"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 4) = serie & " - TASA CUPON"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 3) = serie & " - DIAS VENCIMIENTO"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 2) = serie & " - PERIODO CUPON"
'MSFlexGrid1.TextMatrix(0, 8 * valK - 1) = serie & " - TASA DE REFERENCIA"
'MSFlexGrid1.TextMatrix(0, 8 * valK) = serie & " - TASA / SOBRETASA"
For i = 1 To noreg
 txtcadena = rmesa.Fields("FECHA") & Chr(9)
 txtcadena = txtcadena & rmesa.Fields("PPSUCIO") & Chr(9)
 txtcadena = txtcadena & rmesa.Fields("PPLIMPIO") & Chr(9)
 txtcadena = txtcadena & rmesa.Fields("PTCUPON") & Chr(9)
 If Not IsNull(rmesa.Fields("PDV")) Then
  txtcadena = txtcadena & rmesa.Fields("PDV") & Chr(9)
 Else
  txtcadena = txtcadena & 0 & Chr(9)
 End If
 If Not IsNull(rmesa.Fields("PPCupon")) Then
  txtcadena = txtcadena & rmesa.Fields("PPCupon") & Chr(9)
 Else
  txtcadena = txtcadena & 0 & Chr(9)
 End If
 If Not IsNull(rmesa.Fields("PTASA")) Then
  txtcadena = txtcadena & 100 * rmesa.Fields("PTASA") & Chr(9)
 Else
  txtcadena = txtcadena & 0 & Chr(9)
 End If
 If Not IsNull(rmesa.Fields("PTASA")) Then
  txtcadena = txtcadena & rmesa.Fields("PTASA") * 100
 Else
  txtcadena = txtcadena & 0
 End If
Print #1, txtcadena
'MSFlexGrid1.TextMatrix(i, 8 * valK - 7) = RMesa.Fields("FECHA")
'MSFlexGrid1.TextMatrix(i, 8 * valK - 6) = RMesa.Fields("PPSUCIO")
'MSFlexGrid1.TextMatrix(i, 8 * valK - 5) = RMesa.Fields("PPLIMPIO")
'MSFlexGrid1.TextMatrix(i, 8 * valK - 4) = RMesa.Fields("PTCUPON")
'If Not IsNull(RMesa.Fields("PDV")) Then MSFlexGrid1.TextMatrix(i, 8 * valK - 3) = RMesa.Fields("PDV")
'If Not IsNull(RMesa.Fields("PPCupon")) Then MSFlexGrid1.TextMatrix(i, 8 * valK - 2) = RMesa.Fields("PPCupon")
'If Not IsNull(RMesa.Fields("PTASA")) Then MSFlexGrid1.TextMatrix(i, 8 * valK - 1) = 100 * RMesa.Fields("PTASA")
'If Not IsNull(RMesa.Fields("PTASA")) Then MSFlexGrid1.TextMatrix(i, 8 * valK) = RMesa.Fields("PTASA") * 100
rmesa.MoveNext
Next i
End If
rmesa.Close
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo2_Click()
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
Combo1.Clear
Combo3.Clear
txtfiltro = "SELECT count(DISTINCT EMISION) from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "'"
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro = "SELECT EMISION from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' GROUP BY EMISION ORDER BY EMISION"
 rmesa.Open txtfiltro, ConAdo
 rmesa.MoveFirst
 ReDim MatEmision(1 To noreg) As String
 For i = 1 To noreg
 MatEmision(i) = rmesa.Fields("EMISION")
 Combo1.AddItem MatEmision(i)
 rmesa.MoveNext
 Next i
 rmesa.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub


Private Sub Combo3_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
Call TitulosEstadisticas1
Call TituloPantalla
 'Call LimpiarPantallas
txtfiltro1 = "SELECT COUNT(*) FROM " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "' AND SERIE = '" & Combo3.Text & "'"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
MSFlexGrid1.Rows = noreg + 1
If noreg <> 0 Then
txtfiltro2 = "Select * from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "' AND SERIE = '" & Combo3.Text & "' ORDER BY FECHA"
rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 3) As Variant
rmesa.MoveFirst
ReDim MatSerie(1 To noreg) As String
For i = 1 To noreg
MSFlexGrid1.TextMatrix(i, 0) = rmesa.Fields("FECHA")
MSFlexGrid1.TextMatrix(i, 1) = rmesa.Fields("FEMISION")
MSFlexGrid1.TextMatrix(i, 2) = rmesa.Fields("FVENCIMIENTO")
MSFlexGrid1.TextMatrix(i, 3) = rmesa.Fields("PSUCIO")
MSFlexGrid1.TextMatrix(i, 4) = rmesa.Fields("PLIMPIO")
MSFlexGrid1.TextMatrix(i, 5) = rmesa.Fields("TASASOBRET")
MSFlexGrid1.TextMatrix(i, 6) = rmesa.Fields("DVENCIMIENTO")
MSFlexGrid1.TextMatrix(i, 7) = rmesa.Fields("VNOMINAL")
MSFlexGrid1.TextMatrix(i, 8) = rmesa.Fields("TCUPON")
MSFlexGrid1.TextMatrix(i, 9) = rmesa.Fields("PCUPON")
MSFlexGrid1.TextMatrix(i, 10) = rmesa.Fields("YIELD")
MSFlexGrid1.TextMatrix(i, 11) = rmesa.Fields("REGLA_CUPON")
rmesa.MoveNext
Next i
rmesa.Close
End If
Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
Dim noreg As Integer
Dim nocol As Integer
Dim i As Integer
Dim j As Integer
Dim nomarch As String
Dim txtcadena As String
Dim exitoarch As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'este es el segundo filtro,de aqui se procede a
'filtrar la informacion por emision y serie
Screen.MousePointer = 11
noreg = MSFlexGrid1.Rows
nocol = MSFlexGrid1.Cols
nomarch = DirResVaR & "\hist precios " & Combo2.Text & Combo1.Text & Combo3.Text & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To nocol
    txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i - 1, j - 1) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se creo el archivo " & nomarch
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TituloPantalla()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
MSFlexGrid1.Cols = 12
MSFlexGrid1.RowHeight(0) = 900
MSFlexGrid1.ColWidth(0) = 1300
MSFlexGrid1.ColWidth(1) = 1300
MSFlexGrid1.ColWidth(2) = 1300
MSFlexGrid1.ColWidth(3) = 1300
MSFlexGrid1.ColWidth(4) = 1300
MSFlexGrid1.ColWidth(5) = 1300
MSFlexGrid1.ColWidth(6) = 1300
MSFlexGrid1.ColWidth(7) = 1300
MSFlexGrid1.ColWidth(8) = 1300
MSFlexGrid1.ColWidth(9) = 1300
MSFlexGrid1.ColWidth(10) = 1300
MSFlexGrid1.ColWidth(11) = 1300

MSFlexGrid1.TextMatrix(0, 0) = "FECHA"
MSFlexGrid1.TextMatrix(0, 1) = "FECHA DE EMISION"
MSFlexGrid1.TextMatrix(0, 2) = "FECHA DE VENCIMIENTO"
MSFlexGrid1.TextMatrix(0, 3) = "PRECIO SUCIO"
MSFlexGrid1.TextMatrix(0, 4) = "PRECIO LIMPIO"
MSFlexGrid1.TextMatrix(0, 5) = "TASA / SOBRETASA"
MSFlexGrid1.TextMatrix(0, 6) = "DIAS POR VENCER"
MSFlexGrid1.TextMatrix(0, 7) = "VALOR NOMINAL"
MSFlexGrid1.TextMatrix(0, 8) = "TASA CUPON"
MSFlexGrid1.TextMatrix(0, 9) = "PERIODO CUPON"
MSFlexGrid1.TextMatrix(0, 10) = "TASA DE REFERENCIA/YIELD"
MSFlexGrid1.TextMatrix(0, 11) = "REGLA CUPON"

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TitulosEstadisticas1()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
MSFlexGrid1.Rows = 10
MSFlexGrid1.Cols = 5
MSFlexGrid1.TextMatrix(1, 0) = "Precio"
MSFlexGrid1.TextMatrix(2, 0) = "Tasa de referencia"
MSFlexGrid1.TextMatrix(3, 0) = "Sobretasa"
MSFlexGrid1.TextMatrix(0, 1) = "Media"
MSFlexGrid1.TextMatrix(0, 2) = "Desviacion Estandar"
MSFlexGrid1.TextMatrix(0, 3) = "Mínimo"
MSFlexGrid1.TextMatrix(0, 4) = "Máximo"

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
SiActTProc = True
Screen.MousePointer = 11
Combo1.Clear
Combo2.Clear
Combo3.Clear
txtfiltro2 = "SELECT TV from " & TablaVecPrecios & " GROUP BY TV ORDER BY TV"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   ReDim MatTV(1 To noreg, 1 To 1) As String
   For i = 1 To noreg
       MatTV(i, 1) = rmesa.Fields("TV")
       Combo2.AddItem MatTV(i, 1)
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
 Screen.MousePointer = 0
Call ActUHoraUsuario
SiActTProc = False
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmCalVar.Visible = True
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
