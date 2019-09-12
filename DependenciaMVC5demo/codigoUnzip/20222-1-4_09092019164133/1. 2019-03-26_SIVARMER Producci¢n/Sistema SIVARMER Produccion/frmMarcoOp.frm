VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMarcoOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marco de Operación"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6990
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Obtener posición para reportes MO"
      Height          =   765
      Left            =   2580
      TabIndex        =   6
      Top             =   1020
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   765
      Left            =   270
      TabIndex        =   5
      Top             =   1080
      Width           =   1755
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2370
      TabIndex        =   4
      Top             =   420
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generar reporte marco Op PIDV"
      Height          =   765
      Left            =   300
      TabIndex        =   3
      Top             =   2100
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Determinar plazo promedio"
      Height          =   885
      Left            =   2340
      TabIndex        =   2
      Top             =   2100
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   100
      TabIndex        =   0
      Top             =   480
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   100
      TabIndex        =   1
      Top             =   210
      Width           =   450
   End
End
Attribute VB_Name = "frmMarcoOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Dim fecha1 As Date
Dim fecha2 As Date
fecha1 = #9/3/2018#
fecha2 = #9/4/2018#
Screen.MousePointer = 11
Call Carga_VAR_MO(fecha1, fecha2)
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim tfecha1 As String
Dim tfecha2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim i As Integer
Dim j As Integer
Dim matres() As Variant
Dim txtcadena As String

tfecha1 = InputBox("Dame la primera fecha", , Date)
tfecha2 = InputBox("Dame la primera fecha", , Date)
If IsDate(tfecha1) And IsDate(tfecha2) Then
   fecha1 = CDate(tfecha1)
   fecha2 = CDate(tfecha2)
   matres = DetermPlazoPromMO(fecha1, fecha2)
   Open "D:\RESULTADOS plazo promedio.TXT" For Output As #1
   For i = 1 To UBound(matres, 1)
       txtcadena = ""
       For j = 1 To UBound(matres, 2)
           txtcadena = txtcadena & matres(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
    Next i
    Close #1
End If
MsgBox "Fin de proceso"

End Sub

Private Sub Command3_Click()
Dim fecha As Date
Dim txtfecha As String
Dim txtfiltroa As String
Dim txtfiltrob As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltroc As String
Dim txtborra As String
Dim txtcadena As String
Dim noreg As Integer
Dim i As Integer
Dim indice As Integer
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim calif As String
Dim rmesa As New ADODB.recordset


Screen.MousePointer = 11
fecha = InputBox("Dame la fecha de calculo", , Date)
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltroa = "SELECT * FROM " & TablaPosMD
txtfiltroa = txtfiltroa & " WHERE FECHAREG = " & txtfecha
txtfiltroa = txtfiltroa & " AND TIPOPOS=1 "
txtfiltroa = txtfiltroa & " AND (TOPERACION = 1 OR TOPERACION = 4)"
    
txtfiltrob = "SELECT * FROM " & TablaValPos
txtfiltrob = txtfiltrob & " WHERE FECHAP = " & txtfecha
txtfiltrob = txtfiltrob & " AND ID_VALUACION = 1 "
txtfiltrob = txtfiltrob & " AND ESC_FR = 'Normal' "
txtfiltrob = txtfiltrob & " AND (CPOSICION = " & ClavePosMD & " OR CPOSICION = " & ClavePosTeso & " OR CPOSICION=8 OR CPOSICION= 9)"

txtfiltroc = "SELECT TV,EMISION,SERIE,TASA FROM " & TablaVecPrecios & " WHERE FECHA = " & txtfecha
txtfiltroc = txtfiltroc & " "

txtfiltro2 = "SELECT t1.TV, t1.EMISION, t1.SERIE, t1.CPOSICION, t1.TOPERACION, "
txtfiltro2 = txtfiltro2 & "t1.COPERACION, t1.NO_TITULOS,t1.C_EMISION, t2.P_SUCIO, t2.VAL_PIP_S, t2.DUR_ACT,"
txtfiltro2 = txtfiltro2 & "t3.REGLA_CUPON"
txtfiltro2 = txtfiltro2 & "FROM (" & txtfiltroa & ") t1 "
txtfiltro2 = txtfiltro2 & "INNER JOIN (" & txtfiltrob & ") t2 "
txtfiltro2 = txtfiltro2 & "ON t1.COPERACION = t2.COPERACION "
txtfiltro2 = txtfiltro2 & "INNER JOIN (" & txtfiltroc & ") t3 "
txtfiltro2 = txtfiltro2 & "ON t1.TV = t3.TV "
txtfiltro2 = txtfiltro2 & "AND t1.EMISION = t3.EMISION "
txtfiltro2 = txtfiltro2 & "AND t1.SERIE = t3.SERIE "
txtfiltro2 = txtfiltro2 & "ORDER BY t1.TV, t1.EMISION, t1.SERIE"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 20) As Variant
   For i = 1 To noreg
       mata(i, 1) = fecha
       mata(i, 2) = rmesa.Fields("TV")
       mata(i, 3) = rmesa.Fields("EMISION")
       mata(i, 4) = rmesa.Fields("SERIE")
       mata(i, 5) = rmesa.Fields("NO_TITULOS")
       mata(i, 6) = rmesa.Fields("CPOSICION")
       mata(i, 7) = rmesa.Fields("C_EMISION")
       mata(i, 8) = rmesa.Fields("P_SUCIO")
       mata(i, 9) = rmesa.Fields("VAL_PIP_S")
       mata(i, 10) = rmesa.Fields("DUR_ACT")
       mata(i, 11) = rmesa.Fields("REGLA_CUPON")
       mata(i, 12) = rmesa.Fields("CALIF_SP")
       mata(i, 13) = rmesa.Fields("CALIF_FITCH")
       mata(i, 14) = rmesa.Fields("CALIF_MOODYS")
       mata(i, 15) = rmesa.Fields("CALIF_HR")
       mata(i, 15) = rmesa.Fields("MONTO_EMITIDO")
       mata(i, 15) = rmesa.Fields("MONTO_CIRCULACION")
     
       rmesa.MoveNext
   Next i
   rmesa.Close
   matvp = LeerVPrecios(fecha, mindvp)
   
   For i = 1 To noreg
       mata(i, 11) = DetermMOSector(mata(i, 2) & "_" & mata(i, 3))
       mata(i, 12) = DetermMOTipo(mata(i, 2) & "_" & mata(i, 3))
       indice = BuscarValorArray(mata(i, 7), matvp, 22)
       If indice <> 0 Then
          mata(i, 11) = matvp(indice, 18)             'CALIF sp
          mata(i, 12) = matvp(indice, 19)             'calif moodys
          mata(i, 13) = matvp(indice, 20)             'calif fitch
          mata(i, 14) = matvp(indice, 21)             'calif hr
          mata(i, 15) = matvp(indice, 12)             'FECHA DE VENCIMIENTO
          mata(i, 16) = (matvp(indice, 12) - fecha) / 365  'PLAZO
          mata(i, 17) = matvp(indice, 11)            'Valor nominal
          calif = AsignaCalif(mata(i, 11), mata(i, 13), mata(i, 12), mata(i, 14))
          If matvp(indice).regla_cupon = "Tasa fija" Or matvp(indice).regla_cupon = "NA" Then
             mata(i, 18) = "TF"
          Else
             mata(i, 18) = "TV"
          End If
       End If
   Next i
   txtborra = "DELETE FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
   ConAdo.Execute txtborra
   For i = 1 To noreg
       txtcadena = "INSERT INTO " & TablaDetalleMo & " VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & mata(i, 2) & ","       'TV
       txtcadena = txtcadena & mata(i, 3) & ","       'EMISION
       txtcadena = txtcadena & mata(i, 4) & ","       'SERIE
       txtcadena = txtcadena & mata(i, 5) & ","       'NO TITULOS
       txtcadena = txtcadena & mata(i, 6) & ","       'CLAVE DE POSICION
       txtcadena = txtcadena & mata(i, 8) & ","       'calificacion moodys
       txtcadena = txtcadena & mata(i, 9) & ","       'calificacion sp
       txtcadena = txtcadena & mata(i, 10) & ","      'calificacion fitch
       txtcadena = txtcadena & mata(i, 11) & ","      'calificacion hr
       txtcadena = txtcadena & mata(i, 12) & ","       'sector
       txtcadena = txtcadena & mata(i, 13) & ","       'TIPO
       txtcadena = txtcadena & mata(i, 14) & ","       'plazo
       txtcadena = txtcadena & mata(i, 15) & ","       'tasa
       txtcadena = txtcadena & mata(i, 16) & ","       'valor nominal
       txtcadena = txtcadena & mata(i, 17) & ","       'valor nominal
   
   Next i
End If
Print "Marco de operación Mercado de Dinero"
Print #1, "Fecha: " & fecha
Print #1, "Portafolio de inversión"
Print #1, "Cifras en millones de pesos"
Print #1, "Posicion en tasa Fija"
Print #1, "Detalle del portafolio de inversión de la Tesorería de instrumentos conservados a vencimiento"
Print #1, "Instrumento" & Chr(9) & "Tipo de contraparte" & Chr(9) & "Tipo de emisión" & Chr(9) & "Monto" & Chr(9) & "Duración"
For i = 1 To noreg
Next i


Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub


Private Sub Command4_Click()
Dim tfecha As Date
Dim fecha As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim mata() As Variant
Dim rmesa As New ADODB.recordset
Dim txtnomarch As String


tfecha = InputBox("Dame la fecha del marco", , Date)
If IsDate(tfecha) Then
   fecha = CDate(tfecha)
   Screen.MousePointer = 11
   txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtfiltro2 = "SELECT * FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 16)
      For i = 1 To noreg
          mata(i, 1) = rmesa.Fields("FECHA")
          mata(i, 2) = rmesa.Fields("TV")
          mata(i, 3) = rmesa.Fields("EMISION")
          mata(i, 4) = rmesa.Fields("SERIE")
          mata(i, 5) = rmesa.Fields("N_TITULOS")
          mata(i, 6) = rmesa.Fields("MESA")
          mata(i, 7) = rmesa.Fields("CALIF_MOODYS")
          mata(i, 8) = rmesa.Fields("CALIF_SP")
          mata(i, 9) = rmesa.Fields("CALIF_FITCH")
          mata(i, 10) = rmesa.Fields("CALIF_HR")
          mata(i, 11) = rmesa.Fields("SECTOR")
          mata(i, 12) = rmesa.Fields("TIPO")
          mata(i, 13) = rmesa.Fields("PLAZO")
          mata(i, 14) = rmesa.Fields("TASA")
          mata(i, 15) = rmesa.Fields("VN")
          mata(i, 16) = rmesa.Fields("PSUCIO_SIVARMER")
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtnomarch = "D:\MO " & Format(fecha, "YYYY-MM-DD") & ".TXT"
      CommonDialog1.FileName = txtnomarch
      CommonDialog1.ShowSave
      txtnomarch = CommonDialog1.FileName
      Open txtnomarch For Output As #1
      For i = 1 To noreg
          txtcadena = mata(i, 1) & Chr(9) & mata(i, 2) & Chr(9) & mata(i, 3) & Chr(9) & "'" & CStr(mata(i, 4)) & Chr(9)
          txtcadena = txtcadena & mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7) & Chr(9) & mata(i, 8) & Chr(9)
          txtcadena = txtcadena & mata(i, 9) & Chr(9) & mata(i, 10) & Chr(9) & mata(i, 11) & Chr(9) & mata(i, 12) & Chr(9)
          txtcadena = txtcadena & mata(i, 13) & Chr(9) & mata(i, 14) & Chr(9) & mata(i, 15) & Chr(9) & mata(i, 16)
          Print #1, txtcadena
      Next i
      Close #1
End If
MsgBox "Fin de proceso"
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Long
Dim noreg As Long
noreg = UBound(MatFechasVaR, 1)
For i = 1 To UBound(MatFechasVaR, 1)
     Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
     Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
