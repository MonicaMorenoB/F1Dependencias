VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSimFactR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estres de factores de riesgo"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Aplicar Incrementos"
      Height          =   585
      Left            =   13230
      TabIndex        =   20
      Top             =   1380
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar escenario de estres"
      Height          =   585
      Left            =   5520
      TabIndex        =   19
      Top             =   180
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar escenarios estres archivo excel"
      Height          =   615
      Left            =   7200
      TabIndex        =   16
      Top             =   180
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modo de estres de escenario"
      Height          =   645
      Left            =   200
      TabIndex        =   10
      Top             =   1350
      Width           =   6495
      Begin VB.OptionButton Option5 
         Caption         =   "Por tipo de factor"
         Height          =   195
         Left            =   4020
         TabIndex        =   13
         Top             =   300
         Width           =   1635
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Por curva"
         Height          =   195
         Left            =   2310
         TabIndex        =   12
         Top             =   300
         Width           =   1275
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por factor"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11370
      TabIndex        =   8
      Text            =   ".001"
      Top             =   1470
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de incremento"
      Height          =   645
      Left            =   7140
      TabIndex        =   5
      Top             =   1290
      Width           =   3795
      Begin VB.OptionButton Option2 
         Caption         =   "Porcentual"
         Height          =   195
         Left            =   2100
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aritmetico"
         Height          =   225
         Left            =   330
         TabIndex        =   6
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   600
      Left            =   9060
      TabIndex        =   4
      Top             =   210
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2600
      TabIndex        =   2
      Top             =   210
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2600
      TabIndex        =   1
      Top             =   810
      Width           =   2500
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6945
      Left            =   195
      TabIndex        =   14
      Top             =   2490
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   12250
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   6795
      Left            =   7410
      TabIndex        =   15
      Top             =   2490
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   11986
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factor individual"
      Height          =   195
      Left            =   7500
      TabIndex        =   18
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de factores"
      Height          =   195
      Left            =   200
      TabIndex        =   17
      Top             =   2220
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor del incremento"
      Height          =   195
      Left            =   11400
      TabIndex        =   9
      Top             =   1230
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de los factores"
      Height          =   195
      Left            =   200
      TabIndex        =   3
      Top             =   300
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del escenario generado"
      Height          =   195
      Left            =   200
      TabIndex        =   0
      Top             =   870
      Width           =   2265
   End
End
Attribute VB_Name = "frmSimFactR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim fecha As Date
fecha = CDate(Combo1.Text)
  Call MostrarDatos(fecha)
End Sub

Private Sub Combo1_DblClick()
Dim fecha As Date
  Call MostrarDatos(fecha)
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim txtnomarch As String
Dim noesce As Integer
Dim matesce() As Variant
Dim exito1 As Boolean
Dim matfechas() As Date

Screen.MousePointer = 11
frmProgreso.Show
SiActTProc = True
txtnomarch = "d:\escenarios sim 2017-08-31.xlsx"
frmCalVar.CommonDialog1.FileName = txtnomarch
frmCalVar.CommonDialog1.ShowOpen
txtnomarch = frmCalVar.CommonDialog1.FileName

    matesce = LeerEscEstres(txtnomarch)
    matfechas = LeerFechasEsc(txtnomarch)
    NoFechas = UBound(matfechas, 1)
    noesce = UBound(matesce, 1)
    For j = 1 To NoFechas
        MatFactR1 = CargaFR1Dia(matfechas(j, 1), exito1)
        For i = 1 To noesce
            ReDim mata(1 To NoFactores, 1 To 2) As Variant
            For l = 1 To NoFactores
               mata(l, 1) = MatCaracFRiesgo(l).indFactor
               If MatCaracFRiesgo(l).tfactor = "TASA" Or MatCaracFRiesgo(l).tfactor = "TASA REAL" Or MatCaracFRiesgo(l).tfactor = "TASA EXT" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 2)
               ElseIf MatCaracFRiesgo(l).tfactor = "TASA REF" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 3)
               ElseIf MatCaracFRiesgo(l).tfactor = "TASA REF EXT" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 3)
               ElseIf MatCaracFRiesgo(l).tfactor = "YIELD IS" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 4)
               ElseIf MatCaracFRiesgo(l).tfactor = "YIELD" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 5)
               ElseIf MatCaracFRiesgo(l).tfactor = "SOBRETASA" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 6)
               ElseIf MatCaracFRiesgo(l).tfactor = "T CAMBIO" And MatCaracFRiesgo(l).indFactor = "DOLAR PIP FIX 0" Then
                  mata(l, 2) = MatFactR1(l, 1) + matesce(i, 7)
               ElseIf MatCaracFRiesgo(l).tfactor = "T CAMBIO YEN" Then
                  If matesce(i, 1) <> "Normal" Then mata(l, 2) = MatFactR1(l, 1) + 0
               ElseIf MatCaracFRiesgo(l).tfactor = "T CAMBIO" Then
                  mata(l, 2) = MatFactR1(l, 1) + 0
               ElseIf MatCaracFRiesgo(l).tfactor = "INDICE" Then
                  mata(l, 2) = MatFactR1(l, 1) + 0
               ElseIf MatCaracFRiesgo(l).tfactor = "UDI" Then
                  mata(l, 2) = MatFactR1(l, 1) + 0
               Else
                  MsgBox "no se clasifico " & MatCaracFRiesgo(l).indFactor
               End If
          Next l
          Call GuardarEscFR(matesce(i, 1), matfechas(j, 1), mata)
        Next i
    Next j
Unload frmProgreso
Call ActUHoraUsuario
SiActTProc = False
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Sub GuardarEscFR(ByVal txtnomfr As String, ByVal fecha As Date, mata() As Variant)
Dim txtinserta As String
Dim txtfecha As String
Dim txtborra As String
Dim i As Long

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaEscFR & " WHERE ID_ESCENARIO = '" & txtnomfr & "' AND FECHA = " & txtfecha
ConAdo.Execute txtborra
For i = 1 To NoFactores
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtinserta = "INSERT INTO " & TablaEscFR & " VALUES("
    txtinserta = txtinserta & "'" & txtnomfr & "',"
    txtinserta = txtinserta & txtfecha & ","
    txtinserta = txtinserta & "'" & mata(i, 1) & "',"
    txtinserta = txtinserta & mata(i, 2) & ")"
    ConAdo.Execute txtinserta
Next i
End Sub


Private Sub Command2_Click()
Dim txtnomfr As String
Dim fecha As Date
Dim i As Long

txtnomfr = Text3.Text
fecha = CDate(Combo1.Text)
Screen.MousePointer = 11
ReDim mata(1 To NoFactores, 1 To 2) As Variant
For i = 1 To NoFactores
    mata(i, 1) = MSFlexGrid2.TextMatrix(i, 0)
    mata(i, 2) = CDbl(MSFlexGrid2.TextMatrix(i, 4))
Next i
Call GuardarEscFR(txtnomfr, fecha, mata)
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub



Sub MostrarDatos(ByVal fecha As Date)
Dim i As Long
Dim indice As Long
Dim noreg As Long
Dim noreg1 As Long
Dim exito1 As Boolean

indice = BuscarValorArray(fecha, MatFechasVaR, 1)
If indice <> 0 Then
   MatFactR1 = CargaFR1Dia(fecha, exito1)
   MSFlexGrid2.Cols = 5
   MSFlexGrid2.Rows = NoFactores + 1
   MSFlexGrid2.TextMatrix(0, 0) = "Factor de riesgo y nodo"
   MSFlexGrid2.TextMatrix(0, 1) = "Factor de riesgo"
   MSFlexGrid2.TextMatrix(0, 2) = "Tipo de factor"
   MSFlexGrid2.TextMatrix(0, 3) = "Valor original"
   MSFlexGrid2.TextMatrix(0, 4) = "Valor alterado"
   MSFlexGrid2.ColWidth(0) = 2000
   MSFlexGrid2.ColWidth(1) = 2000
   MSFlexGrid2.ColWidth(2) = 2000
   MSFlexGrid2.ColWidth(3) = 1000
   MSFlexGrid2.ColWidth(4) = 1000
   For i = 1 To NoFactores
       MSFlexGrid2.TextMatrix(i, 0) = MatCaracFRiesgo(i).indFactor
       MSFlexGrid2.TextMatrix(i, 1) = MatCaracFRiesgo(i).nomFactor
       MSFlexGrid2.TextMatrix(i, 2) = MatCaracFRiesgo(i).tfactor
       MSFlexGrid2.TextMatrix(i, 3) = MatFactR1(i, 1)
       MSFlexGrid2.TextMatrix(i, 4) = MatFactR1(i, 1)
   Next i
End If

End Sub


Private Sub Command4_Click()
Screen.MousePointer = 0
Dim indice2, indice1 As Integer
Dim incr As Double
Dim i As Long
Dim j As Long

If Option3.value Then
   For i = 1 To NoFactores
       incr = MSFlexGrid1.TextMatrix(i, 1)
       MSFlexGrid2.TextMatrix(i, 4) = MSFlexGrid2.TextMatrix(i, 3) + incr
   Next i
ElseIf Option4.value Then
     For i = 1 To NoFactores
         For j = 1 To MSFlexGrid1.Rows - 1
             If MSFlexGrid2.TextMatrix(i, 1) = MSFlexGrid1.TextMatrix(j, 0) Then
                incr = CDbl(MSFlexGrid1.TextMatrix(j, 1))
                MSFlexGrid2.TextMatrix(i, 4) = MSFlexGrid2.TextMatrix(i, 3) + incr
             End If
         Next j
     Next i
   
ElseIf Option5.value Then
     For i = 1 To NoFactores
         For j = 1 To MSFlexGrid1.Rows - 1
            If MSFlexGrid2.TextMatrix(i, 2) = MSFlexGrid1.TextMatrix(j, 0) Then
               incr = CDbl(MSFlexGrid1.TextMatrix(j, 1))
               MSFlexGrid2.TextMatrix(i, 4) = MSFlexGrid2.TextMatrix(i, 3) + incr
            End If
         Next j
     Next i
End If
  
  
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Long
Combo1.Clear
For i = 1 To UBound(MatFechasVaR, 1)
    Combo1.AddItem MatFechasVaR(UBound(MatFechasVaR, 1) - i + 1, 1)
Next i
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim indice2, indice1 As Integer
Dim incr As Double
Dim i As Long
  indice1 = MSFlexGrid1.MouseRow
  indice2 = MSFlexGrid1.MouseCol
  incr = Val(Text2.Text)
  
     For i = 1 To MSFlexGrid1.Rows - 1
         If i = indice1 And indice2 = 2 Then
            MSFlexGrid1.TextMatrix(i, 1) = MSFlexGrid1.TextMatrix(i, 1) - incr
         End If
         If i = indice1 And indice2 = 3 Then
            MSFlexGrid1.TextMatrix(i, 1) = MSFlexGrid1.TextMatrix(i, 1) + incr
         End If
     Next i
  
  
End Sub

Private Sub MSFlexGrid2_DblClick()
Dim i As Integer, indice1 As Integer, indice2 As Integer
Dim vinc As Double
  indice1 = MSFlexGrid2.MouseRow
  indice2 = MSFlexGrid2.MouseCol
  vinc = Val(Text2.Text)
  If indice1 <> 0 And indice2 = 1 Then
     MSFlexGrid2.TextMatrix(indice1, 3) = Val(MSFlexGrid2.TextMatrix(indice1, 3)) - vinc
     For i = 1 To NoFactores
         If MatResFRiesgo(indice1, 1) = MatCaracFRiesgo(i).nomFactor Then
            MSFlexGrid1.TextMatrix(i, 4) = Val(MSFlexGrid1.TextMatrix(i, 4)) - vinc
         End If
      Next i
  End If
  If indice1 <> 0 And indice2 = 2 Then
     MSFlexGrid2.TextMatrix(indice1, 3) = Val(MSFlexGrid2.TextMatrix(indice1, 3)) + vinc
     For i = 1 To NoFactores
         If MatResFRiesgo(indice1, 1) = MatCaracFRiesgo(i).nomFactor Then
            MSFlexGrid1.TextMatrix(i, 4) = Val(MSFlexGrid1.TextMatrix(i, 4)) + vinc
         End If
      Next i
  End If

End Sub

Private Sub Option3_Click()
 Call MostrarTabla1(1)
End Sub

Sub MostrarTabla1(ByVal opcion As Integer)
Dim i As Long
Dim noreg1 As Long
If opcion = 1 Then
   MSFlexGrid1.Cols = 4
   MSFlexGrid1.Rows = NoFactores + 1
   MSFlexGrid1.TextMatrix(0, 0) = "Factor de riesgo"
   MSFlexGrid1.TextMatrix(0, 1) = "Incremento a aplicar"
   MSFlexGrid1.TextMatrix(0, 2) = "Decrementar 10"
   MSFlexGrid1.TextMatrix(0, 3) = "Incrementar 10"
   MSFlexGrid1.ColWidth(0) = 2500
   MSFlexGrid1.ColWidth(1) = 2000
   MSFlexGrid1.ColWidth(2) = 500
   MSFlexGrid1.ColWidth(3) = 500
   For i = 1 To NoFactores
       MSFlexGrid1.TextMatrix(i, 0) = MatCaracFRiesgo(i).indFactor
       MSFlexGrid1.TextMatrix(i, 1) = 0
   Next i

ElseIf opcion = 2 Then
   noreg1 = UBound(MatResFRiesgo, 1)
   MSFlexGrid1.Cols = 4
   MSFlexGrid1.Rows = noreg1 + 1
   MSFlexGrid1.TextMatrix(0, 0) = "Nombre factor"
   MSFlexGrid1.TextMatrix(0, 1) = "Incremento a aplicar"
   MSFlexGrid1.TextMatrix(0, 2) = "decrementar 10"
   MSFlexGrid1.TextMatrix(0, 3) = "Incrementar 10"
   MSFlexGrid1.ColWidth(0) = 2500
   MSFlexGrid1.ColWidth(1) = 2000
   MSFlexGrid1.ColWidth(2) = 500
   MSFlexGrid1.ColWidth(3) = 500
   For i = 1 To noreg1
       MSFlexGrid1.TextMatrix(i, 0) = MatResFRiesgo(i).nomFactor
       MSFlexGrid1.TextMatrix(i, 1) = 0
   Next i
ElseIf opcion = 3 Then
   MSFlexGrid1.Cols = 4
   MSFlexGrid1.Rows = 1
   MSFlexGrid1.Rows = 12
   MSFlexGrid1.TextMatrix(0, 0) = "Tipo de factor"
   MSFlexGrid1.TextMatrix(0, 1) = "Incremento a aplicar"
   MSFlexGrid1.TextMatrix(0, 2) = "decrementar 10"
   MSFlexGrid1.TextMatrix(0, 3) = "Incrementar 10"
   MSFlexGrid1.ColWidth(0) = 2500
   MSFlexGrid1.ColWidth(1) = 2000
   MSFlexGrid1.ColWidth(2) = 500
   MSFlexGrid1.ColWidth(3) = 500
   MSFlexGrid1.TextMatrix(1, 0) = "INDICE"
   MSFlexGrid1.TextMatrix(2, 0) = "SOBRETASA"
   MSFlexGrid1.TextMatrix(3, 0) = "T CAMBIO"
   MSFlexGrid1.TextMatrix(4, 0) = "TASA"
   MSFlexGrid1.TextMatrix(5, 0) = "TASA EXT"
   MSFlexGrid1.TextMatrix(6, 0) = "TASA REAL"
   MSFlexGrid1.TextMatrix(7, 0) = "TASA REF"
   MSFlexGrid1.TextMatrix(8, 0) = "TASA REF EXT"
   MSFlexGrid1.TextMatrix(9, 0) = "UDI"
   MSFlexGrid1.TextMatrix(10, 0) = "YIELD"
   MSFlexGrid1.TextMatrix(11, 0) = "YIELD IS"
   For i = 1 To 11
   MSFlexGrid1.TextMatrix(i, 1) = 0
   Next i
   
End If

End Sub

Private Sub Option4_Click()
 Call MostrarTabla1(2)
End Sub

Private Sub Option5_Click()
 Call MostrarTabla1(3)
End Sub
