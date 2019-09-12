VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAnalisisPosMD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de la Posición de Deuda"
   ClientHeight    =   10260
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4860
      TabIndex        =   11
      Top             =   450
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1635
      Left            =   6180
      TabIndex        =   10
      Top             =   7600
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   2884
      _Version        =   393217
      TextRTF         =   $"frmAnalisisPosMD.frx":0000
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1665
      Left            =   200
      TabIndex        =   9
      Top             =   7600
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   2937
      _Version        =   393217
      TextRTF         =   $"frmAnalisisPosMD.frx":008B
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar a archivo texto"
      Height          =   600
      Left            =   2370
      TabIndex        =   7
      Top             =   930
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar proceso"
      Height          =   600
      Left            =   200
      TabIndex        =   6
      Top             =   900
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opcion"
      Height          =   1275
      Left            =   9390
      TabIndex        =   3
      Top             =   180
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Valor v precios"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   750
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Títulos"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   450
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   450
      Width           =   2000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -72
      Top             =   -96
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5775
      Left            =   200
      TabIndex        =   8
      Top             =   1620
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   10186
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Portafolio"
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comparar posiciones de las fechas:"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   90
      Width           =   2505
   End
End
Attribute VB_Name = "frmAnalisisPosMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 If KeyAscii = 13 Then

 End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Command1_Click()
Dim tfecha1 As String
Dim tfecha2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtport As String
Screen.MousePointer = 11
 tfecha1 = Combo1.Text
 tfecha2 = Combo2.Text
 txtport = Combo3.Text
 If IsDate(tfecha1) And IsDate(tfecha2) And Not EsVariableVacia(txtport) Then
    fecha1 = CDate(tfecha1)
    fecha2 = CDate(tfecha2)
    Call CompararPosicionesMD(fecha1, fecha2, txtport)
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim noreg As Integer
Dim nocols As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim nomarch As String
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim exitoarch As Boolean

Screen.MousePointer = 11
noreg = frmAnalisisPosMD.MSFlexGrid1.Rows
nocols = frmAnalisisPosMD.MSFlexGrid1.Cols
fecha1 = CDate(Combo1.Text)
fecha2 = CDate(Combo2.Text)
nomarch = DirResVaR & "\Analisis pos MD " & Format(fecha1, "yyyy-mm-dd") & "  " & Format(fecha2, "yyyy-mm-dd") & ".txt"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
For i = 1 To noreg
txtcadena = ""
For j = 1 To nocols
txtcadena = txtcadena & frmAnalisisPosMD.MSFlexGrid1.TextMatrix(i - 1, j - 1) & Chr(9)
Next j
Print #1, txtcadena
Next i
Close #1
MsgBox "Se creo el archivo " & nomarch
End If
Screen.MousePointer = 0
End Sub






Private Sub Form_Load()
Dim noreg As Long
Dim i As Long
noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
Combo3.AddItem "MERCADO DE DINERO"
Combo3.AddItem "MESA DE DINERO"
Combo3.AddItem "TESORERIA"
Combo3.AddItem "PIDV"
Combo3.AddItem "PICV"
Combo3.AddItem "PORTAFOLIO DE INVERSION"
End Sub

Sub CompararPosicionesMD(ByRef fecha1 As Date, ByRef fecha2 As Date, ByVal txtport As String)
Dim mata1() As propPosMD
Dim mata2() As propPosMD
Dim matunion() As New propPosMD
Dim noreg As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim indice As Integer
Dim textpos1 As String
Dim textpos2 As String
Dim textneg1 As String
Dim textneg2 As String
Dim texto1 As String
Dim texto2 As String
Dim nogrp As Integer
Dim notit As String
Dim montonoc As String

'esta es la clave para el filtrado
Call LeerYUnirPosMD(fecha1, fecha2, mata1, mata2, matunion, txtport)
'se obtienen las emisiones presentes en la posicion
noreg = UBound(matunion, 1)
noreg1 = UBound(mata1, 1)  'mesa1
noreg2 = UBound(mata2, 1)  'mesa2

ReDim matresumen(1 To noreg, 1 To 15) As Variant
For i = 1 To noreg
    matresumen(i, 1) = matunion(i).cEmisionMD
    matresumen(i, 2) = matunion(i).tValorMD
    matresumen(i, 3) = matunion(i).emisionMD
    matresumen(i, 4) = matunion(i).serieMD
Next i

'leer el vector de precios de la fecha2
For i = 1 To noreg
    matresumen(i, 5) = 0
    matresumen(i, 8) = 0
    matresumen(i, 11) = 0
    For j = 1 To noreg1 'fecha 1
        If mata1(j).cEmisionMD = matresumen(i, 1) And (mata1(j).Tipo_Mov = 1 Or mata1(j).Tipo_Mov = 6) Then
           matresumen(i, 5) = matresumen(i, 5) + mata1(j).noTitulosMD    'compra en directo
        End If
        If mata1(j).cEmisionMD = matresumen(i, 1) And (mata1(j).Tipo_Mov = 4 Or mata1(j).Tipo_Mov = 7) Then
           matresumen(i, 5) = matresumen(i, 5) - mata1(j).noTitulosMD    'venta en directo
        End If
        If mata1(j).cEmisionMD = matresumen(i, 1) And mata1(j).Tipo_Mov = 2 Then
           matresumen(i, 8) = matresumen(i, 8) + mata1(j).noTitulosMD    'compra en reporto
        End If
       If mata1(j).cEmisionMD = matresumen(i, 1) And mata1(j).Tipo_Mov = 3 Then
           matresumen(i, 11) = matresumen(i, 11) + mata1(j).noTitulosMD    'venta en reporto
       End If
    Next j
    matresumen(i, 6) = 0
    matresumen(i, 9) = 0
    matresumen(i, 12) = 0
    For j = 1 To noreg2   'fecha 2
        If mata2(j).cEmisionMD = matresumen(i, 1) And (mata2(j).Tipo_Mov = 1 Or mata2(j).Tipo_Mov = 6) Then
           matresumen(i, 6) = matresumen(i, 6) + mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And (mata2(j).Tipo_Mov = 4 Or mata2(j).Tipo_Mov = 7) Then
           matresumen(i, 6) = matresumen(i, 6) - mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And mata2(j).Tipo_Mov = 2 Then
           matresumen(i, 9) = matresumen(i, 9) + mata2(j).noTitulosMD
        End If
        If mata2(j).cEmisionMD = matresumen(i, 1) And mata2(j).Tipo_Mov = 3 Then
           matresumen(i, 12) = matresumen(i, 12) + mata2(j).noTitulosMD
        End If
    Next j
matresumen(i, 7) = matresumen(i, 6) - matresumen(i, 5)
matresumen(i, 10) = matresumen(i, 9) - matresumen(i, 8)
matresumen(i, 13) = matresumen(i, 12) - matresumen(i, 11)

Next i
frmAnalisisPosMD.MSFlexGrid1.Rows = noreg + 1
frmAnalisisPosMD.MSFlexGrid1.Cols = 11
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 0) = "Emisión"
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 1) = "Compra Directo " & fecha1
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 2) = "Compra Directo " & fecha2
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 3) = "Variación CD"
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 4) = "Compra Reporto " & fecha1
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 5) = "Compra Reporto " & fecha2
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 6) = "Variacion CR"
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 7) = "Venta Reporto " & fecha1
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 8) = "Venta Reporto " & fecha2
frmAnalisisPosMD.MSFlexGrid1.TextMatrix(0, 9) = "Variacion VR"
For i = 1 To noreg
    frmAnalisisPosMD.MSFlexGrid1.TextMatrix(i, 0) = matresumen(i, 1)
    For j = 2 To 10
        frmAnalisisPosMD.MSFlexGrid1.TextMatrix(i, j - 1) = Format(matresumen(i, j + 3), "###,###,###,###,###,###,##0")
    Next j
Next i
'SE realiza el analisis de la posicion
matvp = LeerVPrecios(fecha2, mindvp)

For i = 1 To noreg
    If matresumen(i, 4) <> 0 Then
        indice = BuscarValorArray(matresumen(i, 1), mindvp, 1)
        If indice <> 0 Then
           matresumen(i, 14) = matvp(mindvp(indice, 2)).psucio      'precio sucio pip
           matresumen(i, 15) = matresumen(i, 7) * matresumen(i, 14) 'marca a mercado
        End If
    End If
Next i

nogrp = 12
ReDim Matinst(1 To nogrp, 1 To 5) As Variant
Matinst(1, 1) = "Certificados bursátiles"
Matinst(2, 1) = "Papel CFE"
Matinst(3, 1) = "Papel PEMEX"
Matinst(4, 1) = "bonos BPAG 28"
Matinst(5, 1) = "bonos BPAG 91"
Matinst(6, 1) = "bonos IPAB con cupón semestral"
Matinst(7, 1) = "Bondes D"
Matinst(8, 1) = "bonos M"
Matinst(9, 1) = "Udibonos"
Matinst(10, 1) = "Cetes"
Matinst(11, 1) = "PRLVs"
Matinst(12, 1) = "Bonos USD"


For i = 1 To noreg
    indice = DetermInstGrupo(matresumen(i, 2), matresumen(i, 3), matresumen(i, 4))
    If indice <> 0 Then
       Matinst(indice, 4) = Matinst(indice, 4) + matresumen(i, 7)    'no de titulos
       Matinst(indice, 5) = Matinst(indice, 5) + matresumen(i, 15)   'valor de mercado
    Else
      MsgBox "No pudo determinar el tipo de instrumento"
    End If
Next i

textpos1 = ""
textpos2 = ""
textneg1 = ""
textneg2 = ""
For i = 1 To nogrp
    If Truncar(Abs(Matinst(i, 4)) / 1000000, 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)) / 1000000, 2), "#,###,###,##0.0") & " millones de"
    ElseIf Truncar(Abs(Matinst(i, 4)) / 1000000, 2) < 1 And Truncar(Abs(Matinst(i, 4)) / 1000, 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)) / 1000, 2), "###,##0.00") & " mil"
    ElseIf Truncar(Abs(Matinst(i, 4)) / 1000, 2) < 1 And Truncar(Abs(Matinst(i, 4)), 2) >= 1 Then
       notit = Format(Truncar(Abs(Matinst(i, 4)), 2), "###,##0.00")
    End If
    If Truncar(Abs(Matinst(i, 5)) / 1000000, 2) > 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)) / 1000000, 2), "###,##0.00") & " mdp"
    ElseIf Truncar(Abs(Matinst(i, 5)) / 1000000, 2) < 1 And Truncar(Abs(Matinst(i, 5)) / 1000, 2) >= 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)) / 1000, 2), "###,##0.00") & " mil pesos"
    ElseIf Truncar(Abs(Matinst(i, 5)) / 1000, 2) < 1 And Truncar(Abs(Matinst(i, 5)), 2) >= 1 Then
       montonoc = Format(Truncar(Abs(Matinst(i, 5)), 2), "###,##0.0") & " pesos"
    End If
   
    If Matinst(i, 5) > 0 Then
       textpos1 = textpos1 & notit & " títulos ($" & montonoc & ") en " & Matinst(i, 1) & ", "
       textpos2 = textpos2 & "$" & montonoc & " en " & Matinst(i, 1) & ", "
    End If
    If Matinst(i, 5) < 0 Then
       textneg1 = textneg1 & notit & " títulos ($" & montonoc & ") en " & Matinst(i, 1) & ", "
       textneg2 = textneg2 & "$" & montonoc & " en " & Matinst(i, 1) & ", "
    End If
Next i

If Len(textpos1) <> 0 And Len(textneg1) <> 0 Then
   texto1 = "La posición de Mercado de Dinero aumentó " & textpos1 & " y disminuyó " & textneg1
ElseIf Len(textpos1) <> 0 And Len(textneg1) = 0 Then
   texto1 = "La posición de Mercado de Dinero aumentó " & textpos1
ElseIf Len(textpos1) = 0 And Len(textneg1) <> 0 Then
   texto1 = "La posición de Mercado de Dinero disminuyó " & textneg1
End If

If Len(textpos2) <> 0 And Len(textneg2) <> 0 Then
   texto2 = "La posición de Mercado de Dinero aumentó " & textpos2 & " y disminuyó " & textneg2
ElseIf Len(textpos2) <> 0 And Len(textneg2) = 0 Then
   texto2 = "La posición de Mercado de Dinero aumentó " & textpos2
ElseIf Len(textpos2) = 0 And Len(textneg2) <> 0 Then
   texto2 = "La posición de Mercado de Dinero disminuyó " & textneg2
End If


RichTextBox1.Text = texto1
RichTextBox2.Text = texto2
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub



