VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdSQLPort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros de posicion de SIVARMER"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   16140
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5910
      ItemData        =   "frmEdSQLPort.frx":0000
      Left            =   14200
      List            =   "frmEdSQLPort.frx":0002
      TabIndex        =   9
      Top             =   1410
      Width           =   1875
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6200
      TabIndex        =   8
      Top             =   600
      Width           =   7500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4500
      TabIndex        =   7
      Top             =   600
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   600
      Left            =   7900
      TabIndex        =   4
      Top             =   7700
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear nueva cadena"
      Height          =   600
      Left            =   6200
      TabIndex        =   3
      Top             =   7700
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar cadena"
      Height          =   600
      Left            =   4500
      TabIndex        =   2
      Top             =   7700
      Width           =   1500
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6075
      Left            =   4500
      TabIndex        =   1
      Top             =   1410
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   10716
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmEdSQLPort.frx":0004
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7125
      Left            =   200
      TabIndex        =   0
      Top             =   450
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   12568
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Lista de palabras clave"
      Height          =   195
      Left            =   14200
      TabIndex        =   12
      Top             =   1100
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Left            =   6200
      TabIndex        =   11
      Top             =   300
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ID Cadena"
      Height          =   195
      Left            =   4500
      TabIndex        =   10
      Top             =   300
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contenido"
      Height          =   195
      Left            =   4500
      TabIndex        =   6
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de filtros"
      Height          =   195
      Left            =   200
      TabIndex        =   5
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmEdSQLPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim txtcadena As String
Dim txtcadena1 As String
Dim tfecha As String
Dim fecha As Date
Dim txtinserta As String
Dim id_cadena As Integer
Dim nombre As String
id_cadena = Text1.Text
nombre = Text2.Text
txtcadena = RichTextBox1.Text
tfecha = InputBox("Dame la fecha de validación de la cadena", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   siact = ValidaCadSQl(txtcadena, fecha)
   If siact Then
      txtcadena = ReemplazaCadenaTexto(txtcadena, "'", "''")
      txtinserta = "UPDATE " & PrefijoBD & TablaSQLPort & " SET CADENASQL = '" & txtcadena & "' WHERE ID_CADENA = " & id_cadena & " AND PORTAFOLIO = '" & nombre & "'"
      ConAdo.Execute txtinserta
   Else
      MsgBox "La cadena sql no esta bien construida"
   End If
End If
End Sub

Private Sub Command2_Click()
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim tfecha As String
Dim fecha As Date
Dim txtnombre As String
Dim txtinserta As String
Dim contar As Integer
Dim rmesa As New ADODB.recordset

txtcadena1 = RichTextBox1.Text
tfecha = InputBox("Dame la fecha de validación de la cadena", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)

   siact = ValidaCadSQl(txtcadena1, fecha)
   If siact Then
      txtcadena1 = ReemplazaCadenaTexto(txtcadena1, "'", "''")
      txtcadena2 = "SELECT MAX(ID_CADENA) FROM " & PrefijoBD & TablaSQLPort
      rmesa.Open txtcadena2, ConAdo
      contar = rmesa.Fields(0)
      rmesa.Close
      txtnombre = Text2.Text
      txtinserta = "INSERT INTO " & PrefijoBD & TablaSQLPort & " VALUES("
      txtinserta = txtinserta & contar + 1 & ","
      txtinserta = txtinserta & "'" & txtnombre & "',"
      txtinserta = txtinserta & "'" & txtcadena1 & "')"
      MsgBox txtinserta
      ConAdo.Execute txtinserta
   Else
      MsgBox "La cadena sql no esta bien construida"
   End If
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim noreg As Integer
MatSQLPort = CargaTablaD(PrefijoBD & TablaSQLPort, "SQL de definicion de portafolios", 1)
noreg = UBound(MatSQLPort, 1)
MSFlexGrid1.Rows = noreg + 1
MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(1) = 2500
For i = 1 To UBound(MatSQLPort, 1)
    MSFlexGrid1.TextMatrix(i, 0) = MatSQLPort(i, 1)
    MSFlexGrid1.TextMatrix(i, 1) = MatSQLPort(i, 2)
Next i
List1.AddItem "sql_txtfechareg"
List1.AddItem "sql_tipopos"
List1.AddItem "sql_ClavePosMD"
List1.AddItem "sql_ClavePosTeso"
List1.AddItem "sql_ClavePosMC"
List1.AddItem "sql_ClavePosDeriv"
List1.AddItem "sql_ClavePosPIDV"
List1.AddItem "sql_ClavePosPICV"
List1.AddItem "sql_ClavePosPID"
List1.AddItem "sql_ClavePosPenMD"
List1.AddItem "sql_TablaPosMD"
List1.AddItem "sql_TablaPosDiv"
List1.AddItem "sql_TablaPosSwaps"
List1.AddItem "sql_TablaPosFwd"
List1.AddItem "sql_TablaPosDeuda"
List1.AddItem "sql_TablaCatContrap"







End Sub

Private Sub MSFlexGrid1_DblClick()
 Dim indice As Integer
 indice = MSFlexGrid1.row
 Text1.Text = MatSQLPort(indice, 1)
 Text2.Text = MatSQLPort(indice, 2)
 RichTextBox1.Text = MatSQLPort(indice, 3)
End Sub

Function ValidaCadSQl(ByVal txtcadena As String, ByVal fecha As Date)
Dim noreg As Long
Dim txtfiltro2 As String
Dim txtfiltro1 As String
Dim txtfecha As String
Dim rmesa As New ADODB.recordset
On Error GoTo hayerror
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = TraducirCadenaSQL(txtcadena, txtfecha, 1)
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
ValidaCadSQl = True
Exit Function
hayerror:
ValidaCadSQl = False
End Function
