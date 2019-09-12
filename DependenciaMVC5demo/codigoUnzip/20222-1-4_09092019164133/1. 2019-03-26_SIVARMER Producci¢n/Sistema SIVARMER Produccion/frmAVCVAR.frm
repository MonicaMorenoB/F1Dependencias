VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAVCVAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analisis validez modelo CVaR"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6615
      Left            =   105
      TabIndex        =   5
      Top             =   840
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   11668
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAVCVAR.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   100
      TabIndex        =   2
      Top             =   400
      Width           =   2025
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2835
      TabIndex        =   1
      Top             =   400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   675
      Left            =   10710
      TabIndex        =   0
      Top             =   330
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial"
      Height          =   195
      Left            =   100
      TabIndex        =   4
      Top             =   200
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha fin"
      Height          =   195
      Left            =   2895
      TabIndex        =   3
      Top             =   200
      Width           =   660
   End
End
Attribute VB_Name = "frmAVCVAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
Dim noreg As Integer

noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub

Private Sub Command1_Click()
Dim matv() As Variant
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim noreg3 As Integer
Dim noesc2 As Integer
Dim i As Integer
Dim j As Integer
Dim noesc As Integer
Dim nosim As Integer
Dim htiempo As Integer
Dim contar As Integer
Dim suma0 As Double
Dim suma1 As Double
Dim suma2 As Double
Dim noexcep  As Integer
Dim nconf As Double
Dim estadis1 As Double
Dim estkupiec As Double
Dim valor As Double
Dim valx As Double
Dim txtcadena As String
Dim txtcadena1 As String
Dim matc() As String
Dim txtport1 As String
Dim txtport2 As String
Dim txtport3 As String

txtport1 = "NEGOCIACION + INVERSION"
txtport2 = "CONSOLIDADO"
txtport3 = "TOTAL"
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
SiActTProc = True
fecha1 = CDate(Combo1.Text)
fecha2 = CDate(Combo2.Text)
RichTextBox1.Text = ""
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
'PRIMERO SE obtienen los p&L reales generados por la posicion en una ventana de tiempo
txtfiltro = "SELECT * from " & TablaBackPort & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "AND PORTAFOLIO = '" & txtport1 & "'  AND SUBPORT = '" & txtport2 & "' ORDER BY FECHA"
txtfiltro1 = "SELECT count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim matpyl(1 To noreg, 1 To 8) As Variant
   For i = 1 To noreg
       matpyl(i, 1) = rmesa.Fields(0)
       matpyl(i, 2) = rmesa.Fields(3)
       matpyl(i, 3) = rmesa.Fields(4)
       matpyl(i, 4) = (matpyl(i, 3) - matpyl(i, 2))
       AvanceProc = i / noreg
       MensajeProc = "Leyendo los resultados de backtesting " & Format(AvanceProc, "##0.00 %")
       rmesa.MoveNext
   Next i
   rmesa.Close

End If
fecha = fecha1
contar = 0
noesc = 500
htiempo = 1
'se obtiene al mismo tiempo los escenarios generados para la fecha especifica
ReDim matesc(1 To noesc, 1 To noreg) As Variant
For i = 1 To noreg
   txtfecha = "to_date('" & Format(matpyl(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT * from " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
   txtfiltro = txtfiltro & " AND PORTAFOLIO = '" & txtport3 & "'  AND SUBPORT = '" & txtport2 & "'"
   txtfiltro = txtfiltro & " AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
   txtfiltro1 = "SELECT count(*) from (" & txtfiltro & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg3 = rmesa.Fields(0)
   rmesa.Close
   If noreg3 <> 0 Then
      rmesa.Open txtfiltro, ConAdo
      noesc2 = rmesa.Fields("NOESC")
      txtcadena1 = rmesa.Fields("DATOS")
      matc = EncontrarSubCadenas(txtcadena1, ",")
      For j = 1 To noesc2
          matesc(j, i) = CDbl(matc(j))
      Next j
      rmesa.Close
   End If
   DoEvents
Next i
nosim = 2000
Randomize Timer
suma0 = 0
suma1 = 0
noexcep = 0
ReDim matsim(1 To nosim, 1 To 3) As Variant
For i = 1 To noreg
    matv = ExtVecMatV(matesc, i, 0)
    matv = RutinaOrden(matv, 1, SRutOrden)
    nconf = 0.03                                                                'nivel de confianza
    matpyl(i, 5) = CPercentil2(nconf, ConvArVtDbl(matv), 0, 0, False)           'var
    matpyl(i, 6) = CPercentilCVaR(nconf, ConvArVtDbl(matv), 0, 0, False)        'cvar
    If matpyl(i, 4) < matpyl(i, 5) Then
       matpyl(i, 7) = 1
       noexcep = noexcep + 1
    Else
       matpyl(i, 7) = 0
    End If
    matpyl(i, 8) = matpyl(i, 7) / matpyl(i, 6) * matpyl(i, 4)
    suma0 = suma0 + matpyl(i, 7)
    suma1 = suma1 + matpyl(i, 8)
    For j = 1 To nosim
        matsim(j, 1) = CPercentil2(Rnd(), ConvArVtDbl(matv), 0, 0, False)
        If matsim(j, 1) < matpyl(i, 5) Then
           matsim(j, 2) = matsim(j, 2) + 1
           matsim(j, 3) = matsim(j, 3) + matsim(j, 1) / matpyl(i, 6)
        End If
    Next j
    
Next i
estadis1 = -suma1 / suma0 + 1
estkupiec = -2 * Log(((1 - nconf) ^ (noreg - noexcep) * nconf ^ noexcep) / ((1 - (noexcep / noreg)) ^ (noreg - noexcep) * (noexcep / noreg) ^ noexcep))

suma2 = 0
For i = 1 To nosim
    If matsim(i, 2) <> 0 Then
       valor = -matsim(i, 3) / matsim(i, 2) + 1
    Else
       valor = 0
   End If
   If valor < estadis1 Then
      suma2 = suma2 + 1
   End If
Next i
valx = suma2 / nosim
txtcadena = "Nivel de confianza: " & nconf & Chr(13)
txtcadena = txtcadena & "No de dias para el analisis: " & noreg & Chr(13)
txtcadena = txtcadena & "No de excepciones al CVaR: " & noexcep & Chr(13)
txtcadena = txtcadena & "No de simulaciones: " & nosim & Chr(13)
txtcadena = txtcadena & "Valor del estadistico: " & valx & Chr(13)
If valx > nconf Then
   txtcadena = txtcadena & "No hay información suficiente para rechazar el modelo"
Else
   txtcadena = txtcadena & "El modelo no es adecuado"
End If
RichTextBox1.Text = txtcadena
Call ActUHoraUsuario
SiActTProc = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Function Gamma(ByVal x As Double) As Double
Dim valor As Double
Dim pr As Double
'esta funcion solo existe para numeros no negativos
If x >= 0 And x <= 1 Then
   Gamma = GammaC(x + 1) / x
ElseIf x > 1 And x <= 2 Then
   Gamma = GammaC(x)
Else
   valor = x
   pr = 1
   Do While valor > 2
      pr = pr * (valor - 1)
      valor = valor - 1
   Loop
   Gamma = pr * GammaC(valor)
End If

End Function

Function igf(ByVal s As Integer, ByVal Z As Double) As Double
Dim sc As Double
Dim sum As Double
Dim nom As Double
Dim denom As Double
Dim i As Integer

    If Z < 0# Then
       igf = 0#
       Exit Function
    End If
    sc = 1 / s
    sc = sc * Z ^ s
    sc = sc * Exp(-Z)
 
    sum = 1#
    nom = 1#
    denom = 1#
    For i = 0 To 200
        nom = nom * Z
        s = s + 1
        denom = denom * s
        sum = sum + (nom / denom)
    Next i
    igf = sum * sc
End Function

Function GammaC(ByVal x As Double) As Double
Dim suma As Double
Dim i As Integer

'la variable independiente debe estar en el rango 0,1
ReDim matc(1 To 8) As Double
matc(1) = -0.577191652
matc(2) = 0.988205891
matc(3) = -0.897056937
matc(4) = 0.918206857
matc(5) = -0.756704078
matc(6) = 0.482199394
matc(7) = -0.193527818
matc(8) = 0.035868343
suma = 1
For i = 1 To 8
    suma = suma + matc(i) * (x - 1) ^ i
Next i
GammaC = suma

End Function

Function ChiSQ(ByVal x As Integer, ByVal ngl As Integer) As Double
Dim i As Integer
Dim m As Integer
Dim s As Double

If ngl Mod 2 = 0 Then
   m = ngl / 2 - 1
   ReDim mu(0 To m + 1) As Double
   mu(0) = 1: s = 0:
   For i = 0 To m
      s = s + mu(i)
      mu(i + 1) = mu(i) * x / (2 * (i + 1))
   Next i
   ChiSQ = 1 - s * Exp(-x / 2)
Else
   m = (ngl - 1) / 2
   ReDim mu(0 To m) As Double
   mu(0) = 1: s = 0:
   For i = 0 To m - 1
      s = s + mu(i)
      mu(i + 1) = mu(i) * x / (2 * (i + 1) + 1)
   Next i
   ChiSQ = 2 * DNormal(x ^ 0.5, 0, 1, 1) + (2 * x / Pi) ^ 0.5 * Exp(-x / 2) * s - 1
End If

End Function
