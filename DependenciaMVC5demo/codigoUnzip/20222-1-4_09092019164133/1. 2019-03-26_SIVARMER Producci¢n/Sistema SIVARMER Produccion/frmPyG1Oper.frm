VERSION 5.00
Begin VB.Form frmPyGMontSubport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esc de perdidas y ganancias Mont por subport"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6405
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   200
      TabIndex        =   11
      Top             =   1600
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   200
      TabIndex        =   9
      Text            =   "10000"
      Top             =   3400
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   8
      Top             =   1000
      Width           =   4000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   7
      Top             =   400
      Width           =   4000
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   200
      TabIndex        =   5
      Text            =   "1"
      Top             =   2800
      Width           =   2000
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   200
      TabIndex        =   3
      Text            =   "500"
      Top             =   2200
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar archivo"
      Height          =   675
      Left            =   4440
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   200
      TabIndex        =   12
      Top             =   1400
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No de simulaciones"
      Height          =   195
      Left            =   200
      TabIndex        =   10
      Top             =   3200
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Horizonte de tiempo"
      Height          =   195
      Left            =   200
      TabIndex        =   6
      Top             =   2600
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "No de escenarios"
      Height          =   195
      Left            =   200
      TabIndex        =   4
      Top             =   2000
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subportafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   1
      Top             =   800
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   195
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   660
   End
End
Attribute VB_Name = "frmPyGMontSubport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fecha As Date
Dim txtport As String
Dim txtsubport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim nosim As Long

txtport = Combo1.Text
txtsubport = Combo2.Text
fecha = CDate(Combo3.Text)
noesc = Val(Text3.Text)
htiempo = Val(Text4.Text)
nosim = Val(Text1.Text)
Screen.MousePointer = 11
Call LeerPyGMontSubport(fecha, txtport, "Normal", txtsubport, noesc, htiempo, nosim)
Screen.MousePointer = 0
MsgBox "Fin de proceso"
Unload Me

End Sub

Sub LeerPyGMontSubport(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Long)
Dim siesfv As Boolean
Dim fecha0 As Date
Dim matf() As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim indice As Long
Dim i As Long
Dim j As Long
Dim l As Long
Dim noreg As Long
Dim nomarch As String
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

MatGruposPortPos = CargaGruposPortPos(txtsubport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   ReDim mattxt(1 To noesc + 2) As String
   siesfv = EsFechaVaR(f_pos)
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   mattxt(1) = "Subportafolio" & Chr(9)
   mattxt(2) = "MtM t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 2) = matf(i, 1) & Chr(9)
   Next i
   txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(MatGruposPortPos, 1)
       txtfiltro2 = "SELECT * FROM " & TablaPyGMontPort & " WHERE FECHA = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
       txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & MatGruposPortPos(i, 3) & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro2 = txtfiltro2 & " AND NOSIM = " & nosim
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       mattxt(1) = mattxt(1) & MatGruposPortPos(i, 3) & Chr(9)
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          mattxt(2) = mattxt(2) & rmesa.Fields("VALT0") & Chr(9)
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              mattxt(l + 2) = mattxt(l + 2) & CDbl(matc(l)) & Chr(9)
          Next l
          rmesa.Close
       Else
          For j = 1 To noesc + 1
              mattxt(j + 1) = mattxt(j + 1) & 0 & Chr(9)
          Next j
       End If
   Next i
   nomarch = DirResVaR & "\Escenarios p y g Montecarlo " & txtport & " subport " & txtsubport & " esc fr " & txtportfr & " no esc " & noesc & " " & Format(f_pos, "YYYY-MM-DD") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #1
   For i = 1 To noesc + 2
       Print #1, mattxt(i)
   Next i
   Close #1
End If

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo1.AddItem "TOTAL"
Combo1.AddItem "NEGOCIACION + INVERSION"
Dim i As Integer

For i = 1 To UBound(MatListaPortPos, 1)
    Combo2.AddItem MatListaPortPos(i, 1)
Next i
For i = 1 To UBound(MatFechasVaR, 1)
    Combo3.AddItem MatFechasVaR(UBound(MatFechasVaR, 1) - i + 1, 1)
Next i

End Sub
