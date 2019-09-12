VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAnalisisEvVaR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analisis de la evolución del VaR"
   ClientHeight    =   11070
   ClientLeft      =   2205
   ClientTop       =   585
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   12045
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen"
      Height          =   9780
      Left            =   96
      TabIndex        =   6
      Top             =   1000
      Width           =   11625
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9120
         Left            =   195
         TabIndex        =   7
         Top             =   330
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   16087
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear archivo texto"
      Height          =   636
      Left            =   9750
      TabIndex        =   5
      Top             =   270
      Width           =   1596
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar análisis"
      Height          =   636
      Left            =   7800
      TabIndex        =   4
      Top             =   270
      Width           =   1692
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2500
      TabIndex        =   1
      Top             =   550
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   550
      Width           =   2000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   195
      Left            =   4890
      TabIndex        =   9
      Top             =   90
      Width           =   660
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha final:"
      Height          =   156
      Left            =   2500
      TabIndex        =   3
      Top             =   300
      Width           =   1188
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha inicial:"
      Height          =   156
      Left            =   200
      TabIndex        =   2
      Top             =   300
      Width           =   1188
   End
End
Attribute VB_Name = "frmAnalisisEvVaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtport As String
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim mata() As Variant
Dim matb() As Variant
Dim matfact() As Variant
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim indice1 As Long
Dim indice2 As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
SiActTProc = True
Screen.MousePointer = 11
'se procede a leer la informacion correspondiente para
'estos 2 periodos se procede despues a empalmarla por
'factores de riesgo
'Nota: si la rutina no funciona quiere decir que han entrado como factores de riesgo
'variables que no tienen descripcion, hay que poner una descripcion de la
'variable en el catalogos portafolio factores r
fecha1 = CDate(Combo1.Text)
fecha2 = CDate(Combo2.Text)
txtport = Combo3.Text
'se leen las sensibilidades calculadas y volatilidades
mata = LeerSensibNuevo(fecha1, txtportCalc2, "Normal", txtport)
'mata = TransSensib(mata)
matb = LeerSensibNuevo(fecha2, txtportCalc2, "Normal", txtport)
'matb = TransSensib(matb)

If UBound(mata, 1) > 0 And UBound(matb, 1) > 0 Then
   noreg1 = UBound(mata, 1)
   noreg2 = UBound(matb, 1)
   mata = RutinaOrden(mata, 1, SRutOrden)
   matb = RutinaOrden(matb, 1, SRutOrden)
   ReDim matfriesgo(1 To noreg1 + noreg2, 1 To 1) As Variant
   For i = 1 To noreg1
       matfriesgo(i, 1) = mata(i, 1)
   Next i
   For i = 1 To noreg2
       matfriesgo(i + noreg1, 1) = matb(i, 1)
   Next i
   matfact = ObtFactUnicos(matfriesgo, 1)
   noreg = UBound(matfact, 1)
   ReDim matresul(1 To noreg, 1 To 15) As Variant
   For i = 1 To noreg
       matresul(i, 1) = matfact(i, 1)
       indice1 = BuscarValorArray(matresul(i, 1), mata, 1)
       If indice1 <> 0 Then
          matresul(i, 2) = mata(indice1, 3)       'curva
          matresul(i, 3) = mata(indice1, 4)       'plazo
          matresul(i, 4) = mata(indice1, 6)       'valor
          matresul(i, 7) = mata(indice1, 8)       'volatilidad
          matresul(i, 10) = mata(indice1, 7)      'derivada (en millones)
          matresul(i, 13) = mata(indice1, 9)     'var
       End If
       indice2 = BuscarValorArray(matresul(i, 1), matb, 1)
       If indice2 <> 0 Then
          matresul(i, 2) = matb(indice2, 3)       'curva
          matresul(i, 3) = matb(indice2, 4)       'plazo
          matresul(i, 5) = matb(indice2, 6)       'valor
          matresul(i, 8) = matb(indice2, 8)       'volatilidad
          matresul(i, 11) = matb(indice2, 7)      'derivada (en millones)
          matresul(i, 14) = matb(indice2, 9)      'var
       End If
 'incrementos en los conceptos anteriores
       If matresul(i, 4) <> 0 Then matresul(i, 6) = matresul(i, 5) / matresul(i, 4) - 1
       If matresul(i, 7) <> 0 Then matresul(i, 9) = matresul(i, 8) / matresul(i, 7) - 1
       If matresul(i, 10) <> 0 Then matresul(i, 12) = matresul(i, 11) / matresul(i, 10) - 1
       If matresul(i, 13) <> 0 Then matresul(i, 15) = matresul(i, 14) / matresul(i, 13) - 1
   Next i

MSFlexGrid1.Rows = noreg + 1
MSFlexGrid1.Cols = 15
MSFlexGrid1.TextMatrix(0, 0) = "Factor"
MSFlexGrid1.TextMatrix(0, 1) = "Curva"
MSFlexGrid1.TextMatrix(0, 2) = "Plazo"

MSFlexGrid1.TextMatrix(0, 3) = "Valor factor " & fecha1
MSFlexGrid1.TextMatrix(0, 4) = "Valor factor " & fecha2
MSFlexGrid1.TextMatrix(0, 5) = "Inc factor "
MSFlexGrid1.TextMatrix(0, 6) = "Valor vol " & fecha1
MSFlexGrid1.TextMatrix(0, 7) = "Valor vol " & fecha2
MSFlexGrid1.TextMatrix(0, 8) = "Inc vol "
MSFlexGrid1.TextMatrix(0, 9) = "Valor Derivada " & fecha1
MSFlexGrid1.TextMatrix(0, 10) = "Valor Derivada " & fecha2
MSFlexGrid1.TextMatrix(0, 11) = "Inc Derivada"
MSFlexGrid1.TextMatrix(0, 12) = "Valor var " & fecha1
MSFlexGrid1.TextMatrix(0, 13) = "Valor var " & fecha2
MSFlexGrid1.TextMatrix(0, 14) = "Inc var"

For i = 1 To noreg
MSFlexGrid1.TextMatrix(i, 0) = matresul(i, 1)
MSFlexGrid1.TextMatrix(i, 1) = matresul(i, 2)
MSFlexGrid1.TextMatrix(i, 2) = matresul(i, 3)
For j = 4 To 5
 MSFlexGrid1.TextMatrix(i, j - 1) = Format(matresul(i, j), "###,###,##0.000000")
Next j
 MSFlexGrid1.TextMatrix(i, 5) = Format(100 * matresul(i, 6), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 6) = Format(matresul(i, 7), "###,###,##0.00000000")
 MSFlexGrid1.TextMatrix(i, 7) = Format(matresul(i, 8), "###,###,##0.00000000")
 MSFlexGrid1.TextMatrix(i, 8) = Format(100 * matresul(i, 9), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 9) = Format(matresul(i, 10), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 10) = Format(matresul(i, 11), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 11) = Format(100 * matresul(i, 12), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 12) = Format(matresul(i, 13), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 13) = Format(matresul(i, 14), "###,###,##0.000000")
 MSFlexGrid1.TextMatrix(i, 14) = Format(100 * matresul(i, 15), "###,###,##0.000000")
Next i
Else
MsgBox "No hay datos para realizar el calculo"
End If
Unload frmProgreso
Screen.MousePointer = 0
Call ActUHoraUsuario
SiActTProc = False
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Command2_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim nomarch As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim txtport As String
Dim exitoarch As Boolean

Screen.MousePointer = 11
fecha1 = CDate(Combo1.Text)
fecha2 = CDate(Combo2.Text)
txtport = Combo3.Text
 nomarch = DirResVaR & "\Análisis de VaR " & txtport & "  " & Format(fecha1, "yyyy-mm-dd") & "  " & Format(fecha2, "yyyy-mm-dd") & ".txt"
 frmCalVar.CommonDialog1.FileName = nomarch
 frmCalVar.CommonDialog1.ShowSave
 nomarch = frmCalVar.CommonDialog1.FileName
 Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
   noreg = frmAnalisisEvVaR.MSFlexGrid1.Rows
   For i = 1 To noreg
   txtcadena = ""
   For j = 1 To 15
       txtcadena = txtcadena & frmAnalisisEvVaR.MSFlexGrid1.TextMatrix(i - 1, j - 1) & Chr(9)
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "Se creo el archivo " & nomarch
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
For i = UBound(MatFechasPos, 1) To 1 Step -1
Combo1.AddItem MatFechasPos(i, 1)
Combo2.AddItem MatFechasPos(i, 1)
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

