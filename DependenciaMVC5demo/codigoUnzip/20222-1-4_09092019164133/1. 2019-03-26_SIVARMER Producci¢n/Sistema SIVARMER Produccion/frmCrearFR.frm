VERSION 5.00
Begin VB.Form frmCrearFR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear factor de riesgo"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   200
      TabIndex        =   11
      Text            =   "1"
      Top             =   2610
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   200
      TabIndex        =   10
      Top             =   1800
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear factor de riesgo"
      Height          =   525
      Left            =   3500
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   1100
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   200
      TabIndex        =   2
      Top             =   1100
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   400
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   400
      Width           =   2000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Incremento al factor"
      Height          =   195
      Left            =   200
      TabIndex        =   12
      Top             =   2250
      Width           =   1410
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Plazo nuevo factor"
      Height          =   195
      Left            =   200
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Factor a copiar"
      Height          =   195
      Left            =   200
      TabIndex        =   8
      Top             =   100
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   800
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial"
      Height          =   195
      Left            =   3000
      TabIndex        =   6
      Top             =   100
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto nuevo factor"
      Height          =   225
      Left            =   200
      TabIndex        =   4
      Top             =   800
      Width           =   1515
   End
End
Attribute VB_Name = "frmCrearFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtconcepto1 As String
Dim txtconcepto2 As String
Dim plazox As Integer
Dim incfact  As Double
Dim fecha1 As Date
Dim fecha2 As Date
Dim i As Integer
Dim indice As Integer
Dim noreg As Integer
Dim txtfecha As String
Dim txtborra As String
Dim txtcadena As String
Dim nocampos As Integer
Dim txtorden As String
Dim exito As Boolean

Screen.MousePointer = 11
   txtconcepto1 = Combo1.Text
   txtconcepto2 = Text1.Text
   plazox = Val(Text2.Text)
   incfact = Val(Text3.Text)
   If IsDate(Combo2.Text) And IsDate(Combo4.Text) Then
   fecha1 = CDate(Combo2.Text)
   fecha2 = CDate(Combo4.Text)
   Unload Me
   frmProgreso.Show
   Call CrearMatFRiesgo2(fecha1, fecha2, MatFactRiesgo, "", exito)
   For i = 1 To UBound(MatCaracFRiesgo, 1)
       If MatCaracFRiesgo(i).indFactor = txtconcepto1 Then
          indice = i
          Exit For
       End If
   Next i

noreg = UBound(MatFactRiesgo, 1)
nocampos = 2
ReDim matdatos(1 To noreg, 1 To nocampos) As Variant
For i = 1 To noreg
    matdatos(i, 1) = MatFactRiesgo(i, 1)
    matdatos(i, 2) = Val(MatFactRiesgo(i, indice + 1)) * incfact
Next i
For i = 1 To noreg
    txtfecha = "to_date('" & Format(matdatos(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha
    txtborra = txtborra & " AND CONCEPTO = '" & txtconcepto2 & "'"
    txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
    txtcadena = txtcadena & txtfecha & ","
    txtcadena = txtcadena & "'" & txtconcepto2 & "',"
    txtcadena = txtcadena & plazox & ","
    txtcadena = txtcadena & matdatos(i, 2) & ","
    txtorden = CLng(matdatos(i, 1)) & txtconcepto2 & "0000000"
    txtcadena = txtcadena & "'" & txtorden & "')"
    ConAdo.Execute txtborra
    ConAdo.Execute txtcadena
Next i
End If
Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim nofechas1 As Integer

For i = 1 To UBound(MatCaracFRiesgo, 1)
   Combo1.AddItem MatCaracFRiesgo(i).indFactor
Next i
nofechas1 = UBound(MatFechasVaR, 1)
For i = 1 To nofechas1
   Combo2.AddItem MatFechasVaR(nofechas1 - i + 1, 1)
   Combo4.AddItem MatFechasVaR(nofechas1 - i + 1, 1)
Next i

End Sub

