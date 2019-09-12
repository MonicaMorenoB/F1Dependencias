VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCLimiteC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo de limite de contraparte"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Tasa de equilibrio"
      Height          =   1305
      Left            =   10230
      TabIndex        =   32
      Top             =   6840
      Width           =   2205
      Begin VB.OptionButton Option8 
         Caption         =   "Pasiva"
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   840
         Width           =   945
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Activa"
         Height          =   255
         Left            =   300
         TabIndex        =   33
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Valuacion prospectiva"
      Height          =   1485
      Left            =   10110
      TabIndex        =   29
      Top             =   5100
      Width           =   2805
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   210
         TabIndex        =   31
         Top             =   810
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Limitar valuacion prospectiva"
         Height          =   195
         Left            =   270
         TabIndex        =   30
         Top             =   300
         Width           =   2475
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tabla de subprocesos"
      Height          =   1215
      Left            =   10020
      TabIndex        =   26
      Top             =   3600
      Width           =   2835
      Begin VB.OptionButton Option6 
         Caption         =   "Subprocesos 2"
         Height          =   285
         Left            =   360
         TabIndex        =   28
         Top             =   750
         Width           =   1755
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Subprocesos 1"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   330
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8205
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   14473
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Posicion total"
      TabPicture(0)   =   "frmCLimiteC.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Simulaciones"
      TabPicture(1)   =   "frmCLimiteC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "MSFlexGrid1"
      Tab(1).Control(6)=   "Text5"
      Tab(1).Control(7)=   "Text4"
      Tab(1).Control(8)=   "Text3"
      Tab(1).Control(9)=   "Text2"
      Tab(1).Control(10)=   "Text1"
      Tab(1).Control(11)=   "Command4"
      Tab(1).Control(12)=   "Command3"
      Tab(1).ControlCount=   13
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   480
         TabIndex        =   35
         Top             =   1140
         Width           =   3500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Determinar tasa de equilibrio"
         Height          =   705
         Left            =   -74640
         TabIndex        =   25
         Top             =   6450
         Width           =   2000
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Generar subprocesos calculo límite"
         Height          =   765
         Left            =   -72450
         TabIndex        =   24
         Top             =   6390
         Width           =   2000
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -74460
         TabIndex        =   18
         Top             =   5610
         Width           =   1485
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   -72960
         TabIndex        =   17
         Top             =   5610
         Width           =   1500
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   -71355
         TabIndex        =   16
         Top             =   5610
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         Height          =   345
         Left            =   -69810
         TabIndex        =   15
         Top             =   5640
         Width           =   1500
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -68115
         TabIndex        =   14
         Top             =   5610
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar subprocesos"
         Height          =   700
         Left            =   240
         TabIndex        =   12
         Top             =   5910
         Width           =   2000
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones"
         Height          =   3405
         Left            =   210
         TabIndex        =   7
         Top             =   2430
         Width           =   4095
         Begin VB.OptionButton Option1 
            Caption         =   "Una contraparte"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   1260
            Width           =   1545
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   200
            TabIndex        =   9
            Top             =   2220
            Width           =   3500
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Todas las contrapartes"
            Height          =   195
            Left            =   200
            TabIndex        =   8
            Top             =   400
            Width           =   2355
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contraparte"
            Height          =   195
            Left            =   200
            TabIndex        =   11
            Top             =   1860
            Width           =   825
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Subprocesos limites contraparte forwards"
         Height          =   705
         Left            =   2460
         TabIndex        =   6
         Top             =   5910
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4725
         Left            =   -74850
         TabIndex        =   13
         Top             =   510
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8334
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de calculo"
         Height          =   195
         Left            =   450
         TabIndex        =   36
         Top             =   810
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de registro"
         Height          =   195
         Left            =   -74460
         TabIndex        =   23
         Top             =   5310
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre pos"
         Height          =   195
         Left            =   -72960
         TabIndex        =   22
         Top             =   5310
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hora de registro"
         Height          =   195
         Left            =   -71355
         TabIndex        =   21
         Top             =   5310
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "clave pos"
         Height          =   195
         Left            =   -69645
         TabIndex        =   20
         Top             =   5310
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "clave de operacion"
         Height          =   195
         Left            =   -68085
         TabIndex        =   19
         Top             =   5310
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opción"
      Height          =   1125
      Left            =   10080
      TabIndex        =   3
      Top             =   1770
      Width           =   2865
      Begin VB.CheckBox Check1 
         Caption         =   "Considerar tipos de cambios"
         Height          =   255
         Left            =   270
         TabIndex        =   4
         Top             =   510
         Value           =   1  'Checked
         Width           =   2475
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de calculo"
      Height          =   1395
      Left            =   10050
      TabIndex        =   0
      Top             =   210
      Width           =   2925
      Begin VB.OptionButton Option4 
         Caption         =   "Mínima exposicion"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   900
         Width           =   1845
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Máxima exposición"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmCLimiteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo3_Click()
Dim fecha As Date
Dim mata() As String
Dim i As Integer
fecha = CDate(Combo3.Text)
mata = ObtenerContrapNoFinSwaps(fecha)
Combo1.Clear
For i = 1 To UBound(mata, 1)
    Combo1.AddItem mata(i, 2)
Next i
End Sub

Private Sub Command1_Click()
Dim exito As Boolean
Dim txtfecha As String
Dim tfecha As String
Dim fecha As Date
Dim fecha0 As Date
Dim fechax As Date
Dim opc_fecha As Integer
Dim i As Integer
Dim id_contrap As String
Dim exito1 As Boolean
Dim txtpossim As String
Dim txtmsg As String
Dim id_tabla As Integer
Dim opcion_c As Integer
Dim txtborra As String
Dim txtport As String
Dim rmesa As New ADODB.recordset
Dim mata() As String

Screen.MousePointer = 11
tfecha = Combo3.Text
If Option3.value Then
   id_tabla = 1
ElseIf Option6.value Then
   id_tabla = 2
End If
If Not Check2.value Then
   opc_fecha = 1
Else
   opc_fecha = 2
End If
If Option2.value Then
   opcion_c = 0
Else
   opcion_c = 1
End If
If IsDate(tfecha) Then
   fecha0 = #1/2/2008#
   fechax = Date
   fecha = CDate(tfecha)
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   SiActTProc = True
   frmProgreso.Show
   If Option5.value Then
      mata = ObtenerContrapNoFinSwaps(fecha)
      If UBound(mata, 1) <> 0 Then
         txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & 79
         ConAdo.Execute txtborra
         txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & 80
         ConAdo.Execute txtborra
         txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & 81
         ConAdo.Execute txtborra
         txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & 82
         ConAdo.Execute txtborra
         txtborra = "DELETE FROM " & TablaLimContrap1 & " WHERE FECHA  = " & txtfecha
         ConAdo.Execute txtborra
         txtborra = "DELETE FROM " & TablaLimContrap2 & " WHERE FECHA  = " & txtfecha
         ConAdo.Execute txtborra
         For i = 1 To UBound(mata, 1)
             id_contrap = mata(i, 1)
             txtport = "Swap Contrap " & id_contrap
             Call GenerarLSubpLimC1(fecha0, fecha, txtport, opcion_c, opc_fecha, fechax, 79, id_tabla)
             Call GenLSubConsolLimC1(fecha, txtport, 80, id_tabla)
             Call GenerarLSubpLimC2(fecha0, fecha, txtport, opcion_c, opc_fecha, fechax, 81, id_tabla)
             Call GenLSubConsolLimC2(fecha, txtport, 82, id_tabla)
         Next i
      End If
   ElseIf Option1.value Then        'una contraparte
      For i = 1 To UBound(mata, 1)
          If Combo1.Text = mata(i, 2) Then
             id_contrap = mata(i, 1)
             Call GenerarLSubpLimC1(fecha0, fecha, txtport, opcion_c, opc_fecha, fechax, 79, id_tabla)
             Call GenLSubConsolLimC1(fecha, txtport, 80, id_tabla)
             Call GenerarLSubpLimC2(fecha0, fecha, txtport, opcion_c, opc_fecha, fechax, 81, id_tabla)
             Call GenLSubConsolLimC2(fecha, txtport, 82, id_tabla)
             Exit For
         End If
      Next i
   End If
   Call ActUHoraUsuario
   SiActTProc = False
End If
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub


Private Sub Command2_Click()
Dim fecha As Date
Dim txtfecha As String
Dim id_tabla As Integer
Screen.MousePointer = 11
  txtfecha = Combo3.Text
  If Option3.value Then
     id_tabla = 1
  ElseIf Option6.value Then
     id_tabla = 2
  End If
   SiActTProc = True
   fecha = CDate(txtfecha)
   frmProgreso.Show
   Call GenerarSubpMaxExpFwds(fecha, 78, id_tabla)
   Unload frmProgreso
   SiActTProc = False
   MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim tipopos As Integer
Dim fechar As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtcadena As String
Dim tasa As Double
Dim opcion As Integer
If Option7.value Then
   opcion = 1
Else
   opcion = 2
End If
Screen.MousePointer = 11
If IsDate(Text1.Text) And Not EsVariableVacia(Text2.Text) Then
tipopos = 2
fechar = CDate(Text1.Text)
txtnompos = Text2.Text
horareg = Text3.Text
cposicion = Val(Text4.Text)
coperacion = Text5.Text
tasa = 0

If opcion = 1 Then
   txtcadena = "UPDATE " & TablaPosSwaps & " SET TC_ACTIVA = '" & Format(tasa, "##0.00 %") & "'"
   txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
   txtcadena = txtcadena & " AND CPOSICION = " & cposicion
   txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
   ConAdo.Execute txtcadena
   txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
   txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
   txtcadena = txtcadena & " AND CPOSICION = " & cposicion
   txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
   txtcadena = txtcadena & " AND TPATA = 'B'"
   ConAdo.Execute txtcadena
ElseIf opcion = 2 Then
   txtcadena = "UPDATE " & TablaPosSwaps & " SET TC_PASIVA = '" & Format(tasa, "##0.00 %") & "'"
   txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
   txtcadena = txtcadena & " AND CPOSICION = " & cposicion
   txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
   ConAdo.Execute txtcadena
   txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
   txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
   txtcadena = txtcadena & " AND CPOSICION = " & cposicion
   txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
   txtcadena = txtcadena & " AND TPATA = 'C'"
   ConAdo.Execute txtcadena
End If
tasa = DetermTasaEquilibrio(fechar, tipopos, fechar, txtnompos, horareg, cposicion, coperacion, opcion)
MsgBox "la tasa es del " & Format(tasa, "##0.000000 %")
If tasa <> 0 Then
   If opcion = 1 Then
      txtcadena = "UPDATE " & TablaPosSwaps & " SET TC_ACTIVA = '" & Format(tasa, "##0.00 %") & "'"
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      txtcadena = txtcadena & " AND TPATA = 'B'"
      ConAdo.Execute txtcadena
    ElseIf opcion = 2 Then
      txtcadena = "UPDATE " & TablaPosSwaps & " SET TC_PASIVA = '" & Format(tasa, "##0.00 %") & "'"
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtcadena
      txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tasa
      txtcadena = txtcadena & " WHERE TIPOPOS = " & tipopos
      txtcadena = txtcadena & " AND CPOSICION = " & cposicion
      txtcadena = txtcadena & " AND COPERACION = '" & coperacion & "'"
      txtcadena = txtcadena & " AND TPATA = 'C'"
      ConAdo.Execute txtcadena
   End If
End If
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim tipopos As Integer
Dim txtpossim  As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim fecha As Date
Dim fecha0 As Date
Dim fecha1 As Date
Dim opc_fecha As Integer
Dim opc_calc As Integer
Dim id_tabla As Integer
Dim indice As Long
tipopos = 2
If Option3.value Then
   id_tabla = 1
ElseIf Option6.value Then
   id_tabla = 2
End If
If Option2 Then
   opc_calc = 0
Else
   opc_calc = 1
End If
If Not Check2.value Then
   opc_fecha = 1
Else
   opc_fecha = 2
End If
fecha0 = #1/2/2008#
fecha1 = CDate(Text1.Text)
indice = BuscarValorArray(fecha0, MatFechasVaR, 1)
If indice <> 0 Then
   Screen.MousePointer = 11
   fecha = CDate(Text1.Text)
   txtpossim = Text2.Text
   horareg = Text3.Text
   cposicion = Text4.Text
   coperacion = Text5.Text
   Call GenerarLSubpLimCPosSim1(fecha0, tipopos, fecha, txtpossim, horareg, cposicion, coperacion, opc_calc, opc_fecha, fecha1, 79, id_tabla)
   Call GenLSubConsolLimCPosSim1(tipopos, fecha, txtpossim, horareg, cposicion, coperacion, opc_calc, 80, id_tabla)
   Call GenerarLSubpLimCPosSim2(fecha0, tipopos, fecha, txtpossim, horareg, cposicion, coperacion, opc_calc, opc_fecha, fecha1, 81, id_tabla)
   Call GenLSubConsolLimCPosSim2(tipopos, fecha, txtpossim, horareg, cposicion, coperacion, opc_calc, 82, id_tabla)
   Screen.MousePointer = 0
End If
MsgBox "Fin de proceso"
End Sub

Private Sub Form_Load()
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

noreg1 = UBound(MatFechasVaR, 1)
For i = 1 To noreg1
    Combo3.AddItem MatFechasVaR(noreg1 - i + 1, 1)
Next i
Combo3.Text = MatFechasVaR(noreg1, 1)
  
txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 2 ORDER BY FECHAREG,CPOSICION,COPERACION"
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
   MSFlexGrid1.TextMatrix(0, 0) = "Tipo de posicion"
   MSFlexGrid1.TextMatrix(0, 1) = "Fecha de registro"
   MSFlexGrid1.TextMatrix(0, 2) = "Nombre de la posicion"
   MSFlexGrid1.TextMatrix(0, 3) = "Hora de registro"
   MSFlexGrid1.TextMatrix(0, 4) = "Clave de posicion"
   MSFlexGrid1.TextMatrix(0, 5) = "Clave de operacion"
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
   Text1.Text = MSFlexGrid1.TextMatrix(indice1, 1)
   Text2.Text = MSFlexGrid1.TextMatrix(indice1, 2)
   Text3.Text = MSFlexGrid1.TextMatrix(indice1, 3)
   Text4.Text = MSFlexGrid1.TextMatrix(indice1, 4)
   Text5.Text = MSFlexGrid1.TextMatrix(indice1, 5)
End If
End Sub

Private Sub Option1_Click()
Combo1.Enabled = True
'Combo2.Enabled = False
Combo3.Enabled = True
End Sub

Private Sub Option2_Click()
Combo1.Enabled = False
'Combo2.Enabled = False
Combo3.Enabled = True
End Sub

Private Sub Option3_Click()
Combo1.Enabled = False
'Combo2.Enabled = True
Combo3.Enabled = True
End Sub
