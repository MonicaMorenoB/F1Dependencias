VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSubprocesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de subprocesos"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   15750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Desbloqueado procesos"
      Height          =   675
      Left            =   4200
      TabIndex        =   15
      Top             =   1380
      Width           =   1545
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14250
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tabla de subprocesos"
      Height          =   915
      Left            =   9270
      TabIndex        =   11
      Top             =   120
      Width           =   5955
      Begin VB.OptionButton Option6 
         Caption         =   "Subprocesos 3"
         Height          =   405
         Left            =   3810
         TabIndex        =   14
         Top             =   330
         Width           =   1605
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Subprocesos 2"
         Height          =   255
         Left            =   1890
         TabIndex        =   13
         Top             =   360
         Width           =   1545
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Subprocesos 1"
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar subprocesos"
      Height          =   645
      Left            =   2370
      TabIndex        =   10
      Top             =   1350
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar subprocesos pendientes"
      Height          =   675
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   1665
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generar analisis rendimiento"
      Height          =   675
      Left            =   7200
      TabIndex        =   8
      Top             =   1410
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar a archivo de texto"
      Height          =   700
      Left            =   9270
      TabIndex        =   7
      Top             =   1380
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6015
      Left            =   195
      TabIndex        =   3
      Top             =   2160
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   10610
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   825
      Left            =   200
      TabIndex        =   0
      Top             =   150
      Width           =   8415
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6150
         TabIndex        =   5
         Top             =   350
         Width           =   1905
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Todos"
         Height          =   195
         Left            =   4050
         TabIndex        =   4
         Top             =   350
         Width           =   1000
      End
      Begin VB.OptionButton Option2 
         Caption         =   "En proceso"
         Height          =   195
         Left            =   2580
         TabIndex        =   2
         Top             =   350
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pendientes sin procesar"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   350
         Value           =   -1  'True
         Width           =   2200
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   5550
         TabIndex        =   6
         Top             =   350
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmSubprocesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim tfecha As String
Dim fecha As Date
Dim opcions As Integer
Dim id_tabla As Integer
tfecha = Combo1.Text
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   If Option1.value Then
      opcions = 1
   ElseIf Option2.value Then
      opcions = 2
   ElseIf Option3.value Then
      opcions = 3
   End If
   If Option4.value Then
      id_tabla = 1
   ElseIf Option5.value Then
      id_tabla = 2
   ElseIf Option6.value Then
      id_tabla = 3
   End If
 
   Screen.MousePointer = 11
   Call MostrarListaSubprocesos(fecha, opcions, id_tabla)
   Screen.MousePointer = 0
End If
End Sub

Sub MostrarListaSubprocesos(ByVal fecha As Date, ByVal opcions As Integer, ByVal opcion As Integer)
On Error GoTo hayerror
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(opcion)

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If opcions = 1 Then
   txtfiltro2 = "SELECT * FROM " & txttabla & " WHERE FINALIZADO = 'N' AND BLOQUEADO ='N' AND FECHAP = " & txtfecha & " ORDER BY FOLIO"
ElseIf opcions = 2 Then
   txtfiltro2 = "SELECT * FROM " & txttabla & " WHERE FINALIZADO = 'N' AND BLOQUEADO ='S' AND FECHAP = " & txtfecha & " ORDER BY FOLIO"
ElseIf opcions = 3 Then
   txtfiltro2 = "SELECT * FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " ORDER BY FOLIO"
Else
 MsgBox "no se eligio una opcion"
 Exit Sub
End If
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
  noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       If Not rmesa.EOF Then
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
       End If
   Next i
   rmesa.Close
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Cols = 1
   MSFlexGrid1.Rows = noreg + 1
   MSFlexGrid1.Cols = 14
   MSFlexGrid1.TextMatrix(0, 0) = "Consecutivo"
   MSFlexGrid1.TextMatrix(0, 1) = "ID de proceso"
   MSFlexGrid1.TextMatrix(0, 2) = "Folio"
   MSFlexGrid1.TextMatrix(0, 3) = "Descripcion"
   MSFlexGrid1.TextMatrix(0, 4) = "Comentario"
   MSFlexGrid1.TextMatrix(0, 5) = "Usuario"
   MSFlexGrid1.TextMatrix(0, 6) = "Direccion IP"
   MSFlexGrid1.TextMatrix(0, 7) = "Bloqueado"
   MSFlexGrid1.TextMatrix(0, 8) = "Finalizado"
   MSFlexGrid1.TextMatrix(0, 9) = "Fechas de inicio"
   MSFlexGrid1.TextMatrix(0, 10) = "Hora de inicio"
   MSFlexGrid1.TextMatrix(0, 11) = "Fecha final"
   MSFlexGrid1.TextMatrix(0, 12) = "Hora final"
   MSFlexGrid1.TextMatrix(0, 13) = "Tiempo de proceso"
   MSFlexGrid1.ColWidth(3) = 3000
   For i = 1 To noreg
       MSFlexGrid1.TextMatrix(i, 0) = i                                       'consecutivo
       MSFlexGrid1.TextMatrix(i, 1) = ReemplazaVacioValor(mata(i, 1), "")    'id_subproceso
       MSFlexGrid1.TextMatrix(i, 2) = ReemplazaVacioValor(mata(i, 2), "")    'folio
       MSFlexGrid1.TextMatrix(i, 3) = ReemplazaVacioValor(mata(i, 3), "")    'descripcion
       MSFlexGrid1.TextMatrix(i, 4) = ReemplazaVacioValor(mata(i, 22), "")   'comentario
       MSFlexGrid1.TextMatrix(i, 5) = ReemplazaVacioValor(mata(i, 23), "")   'usuario
       MSFlexGrid1.TextMatrix(i, 6) = ReemplazaVacioValor(mata(i, 24), "")   'direccion ip
       MSFlexGrid1.TextMatrix(i, 7) = ReemplazaVacioValor(mata(i, 20), "")   'bloqueado
       MSFlexGrid1.TextMatrix(i, 8) = ReemplazaVacioValor(mata(i, 21), "")   'finalizado
       MSFlexGrid1.TextMatrix(i, 9) = ReemplazaVacioValor(mata(i, 16), "")   'fecha de inicio
       MSFlexGrid1.TextMatrix(i, 10) = ReemplazaVacioValor(mata(i, 17), "")   'hora de inicio
       MSFlexGrid1.TextMatrix(i, 11) = ReemplazaVacioValor(mata(i, 18), "")   'fecha final
       MSFlexGrid1.TextMatrix(i, 12) = ReemplazaVacioValor(mata(i, 19), "")  'hora final
       If Not EsVariableVacia(mata(i, 16)) And Not EsVariableVacia(mata(i, 17)) And Not EsVariableVacia(mata(i, 18)) And Not EsVariableVacia(mata(i, 19)) Then
          MSFlexGrid1.TextMatrix(i, 13) = TiempoProc(mata(i, 16), mata(i, 17), mata(i, 18), mata(i, 19))  'tiempo de proceso
       End If
   Next i
Else
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Rows = 2
   MSFlexGrid1.Cols = 13
   MSFlexGrid1.TextMatrix(0, 0) = "ID de proceso"
   MSFlexGrid1.TextMatrix(0, 1) = "Folio"
   MSFlexGrid1.TextMatrix(0, 2) = "Descripcion"
   MSFlexGrid1.TextMatrix(0, 3) = "Comentario"
   MSFlexGrid1.TextMatrix(0, 4) = "Usuario"
   MSFlexGrid1.TextMatrix(0, 5) = "Direccion IP"
   MSFlexGrid1.TextMatrix(0, 6) = "Bloqueado"
   MSFlexGrid1.TextMatrix(0, 7) = "Finalizado"
   MSFlexGrid1.TextMatrix(0, 8) = "Fechas de inicio"
   MSFlexGrid1.TextMatrix(0, 9) = "Hora de inicio"
   MSFlexGrid1.TextMatrix(0, 10) = "Fecha final"
   MSFlexGrid1.TextMatrix(0, 11) = "Hora final"
   MSFlexGrid1.TextMatrix(0, 12) = "Tiempo de proceso"
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Private Sub Command1_Click()
Dim txtborra As String
Dim opcion As Integer
Dim id_tabla As Integer
Screen.MousePointer = 11
   If Option4.value Then
      id_tabla = 1
   ElseIf Option5.value Then
      id_tabla = 2
   ElseIf Option6.value Then
      id_tabla = 3
   End If
txtborra = "UPDATE " & DetermTablaSubproc(id_tabla) & " SET BLOQUEADO = 'N' WHERE BLOQUEADO = 'S'"
ConAdo.Execute txtborra
Screen.MousePointer = 0

End Sub

Private Sub Command2_Click()
Dim txtcadena As String
Dim i As Long
Dim j As Integer
Screen.MousePointer = 11
If MSFlexGrid1.Rows > 1 Then
   Open "d:\salida.txt" For Output As #1
   For i = 1 To MSFlexGrid1.Rows
       txtcadena = ""
       For j = 1 To MSFlexGrid1.Cols
           txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i - 1, j - 1) & ","
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim fecha As Date
Dim txtfecha As String
Dim matsubproc() As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim matpr() As Variant
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim noreg1 As Long
Dim txtcadena As String
Dim txttabla As String
Dim txtnomarch As String
Dim rmesa As New ADODB.recordset
Dim id_tabla As Integer
Dim valor1 As Double
Dim valor2 As Double
Dim suma1 As Double
Dim suma2 As Double

If Option4.value Then
   id_tabla = 1
ElseIf Option5.value Then
   id_tabla = 2
ElseIf Option6.value Then
   id_tabla = 3
End If
txttabla = DetermTablaSubproc(id_tabla)
fecha = CDate(Combo1.Text)
Screen.MousePointer = 11
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT ID_SUBPROCESO,DESCRIPCION FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " GROUP BY ID_SUBPROCESO,DESCRIPCION"
txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matpr(1 To noreg, 1 To 10) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       matpr(i, 1) = rmesa.Fields("ID_SUBPROCESO")
       matpr(i, 2) = rmesa.Fields("DESCRIPCION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
      txtfiltro2 = "SELECT MIN(FINICIO),MIN(HINICIO),MAX(FFINAL),MAX(HFINAL),SUM(FFINAL-FINICIO) AS SUMA1,SUM(HFINAL-HINICIO) AS SUMA2 FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & matpr(i, 1)
      txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg1 = rmesa.Fields(0)
      rmesa.Close
      If noreg1 <> 0 Then
         rmesa.Open txtfiltro2, ConAdo
         matpr(i, 3) = rmesa.Fields("MIN(FINICIO)")
         matpr(i, 4) = rmesa.Fields("MIN(HINICIO)")
         matpr(i, 5) = rmesa.Fields("MAX(FFINAL)")
         matpr(i, 6) = rmesa.Fields("MAX(HFINAL)")
         If Not EsVariableVacia(matpr(i, 3)) And Not EsVariableVacia(matpr(i, 4)) And Not EsVariableVacia(matpr(i, 5)) And Not EsVariableVacia(matpr(i, 6)) Then
            valor1 = TiempoProc(matpr(i, 3), matpr(i, 4), matpr(i, 5), matpr(i, 6))
         Else
            valor1 = 0
         End If
         matpr(i, 7) = Format(valor1, "HH:MM:SS")
         suma1 = ReemplazaVacioValor(rmesa.Fields("SUMA1"), 0)
         suma2 = ReemplazaVacioValor(rmesa.Fields("SUMA2"), 0)
         valor2 = suma1 / 24 + suma2
         matpr(i, 8) = Format(valor2, "HH:MM:SS")
         If valor1 <> 0 Then
            matpr(i, 9) = 1 - valor2 / valor1
         Else
            matpr(i, 9) = 0
         End If
         rmesa.Close
      End If
   Next i
   txtnomarch = "Analisis tiempos subprocesos " & Format(fecha, "yyyy-mm-dd") & ".txt"
   CommonDialog1.FileName = txtnomarch
   CommonDialog1.ShowSave
   txtnomarch = CommonDialog1.FileName
   Open txtnomarch For Output As #1
   Print #1, "ID subproceso" & Chr(9) & "Descripcion" & Chr(9) & "Fecha de inicio" & Chr(9) & "Hora de inciio" & Chr(9) & "Fecha final" & Chr(9) & "Hora final" & Chr(9) & "Tiempo acumulado" & Chr(9) & "Tiempo por operacion"
   For i = 1 To noreg
       txtcadena = ""
       For j = 1 To 9
           txtcadena = txtcadena & matpr(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim txtborra As String
Dim opcion As Integer
Dim id_tabla As Integer
Dim valor As Integer
valor = MsgBox("Desea realizar el borrado?", vbYesNo)
If valor = 6 Then
Screen.MousePointer = 11
   If Option4.value Then
      id_tabla = 1
   ElseIf Option5.value Then
      id_tabla = 2
   ElseIf Option6.value Then
      id_tabla = 3
   End If
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FINALIZADO = 'N'"
ConAdo.Execute txtborra
Screen.MousePointer = 0
End If
End Sub

Private Sub Command5_Click()
Dim txtborra As String
Dim id_tabla As Integer
Dim valor As Integer
   If Option4.value Then
      id_tabla = 1
   ElseIf Option5.value Then
      id_tabla = 2
   ElseIf Option6.value Then
      id_tabla = 3
   End If
valor = MsgBox("Desea realizar el borrado?", vbYesNo)
If valor = 6 Then
   Screen.MousePointer = 11
   txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla)
   ConAdo.Execute txtborra
   MsgBox "Borrado finalizado"
   Screen.MousePointer = 0
End If

End Sub

Private Sub Form_Load()
Dim i As Long
Dim noreg As Long
noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
   Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
MSFlexGrid1.Cols = 12
MSFlexGrid1.TextMatrix(0, 0) = "ID de proceso"
MSFlexGrid1.TextMatrix(0, 1) = "Folio"
MSFlexGrid1.TextMatrix(0, 2) = "Descripcion"
MSFlexGrid1.TextMatrix(0, 3) = "Usuario"
MSFlexGrid1.TextMatrix(0, 4) = "Direccion IP"
MSFlexGrid1.TextMatrix(0, 5) = "Bloqueado"
MSFlexGrid1.TextMatrix(0, 6) = "Finalizado"
MSFlexGrid1.TextMatrix(0, 7) = "Fechas de inicio"
MSFlexGrid1.TextMatrix(0, 8) = "Hora de inicio"
MSFlexGrid1.TextMatrix(0, 9) = "Fecha final"
MSFlexGrid1.TextMatrix(0, 10) = "Hora final"
MSFlexGrid1.TextMatrix(0, 11) = "Tiempo de proceso"
End Sub
