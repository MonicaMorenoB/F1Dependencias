VERSION 5.00
Begin VB.Form frmGeneradorParam 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   675
      Left            =   5520
      TabIndex        =   3
      Top             =   2190
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar flujos emision"
      Height          =   735
      Left            =   3330
      TabIndex        =   2
      Top             =   2190
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   780
      TabIndex        =   1
      Top             =   2190
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1485
      Left            =   390
      TabIndex        =   0
      Top             =   300
      Width           =   11925
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   720
         Width           =   1755
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   750
         Width           =   1965
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   750
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo valor"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmGeneradorParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se filtra la tabla y se agrupan los precios por
'emision. Dependiendo de que emision se seleccione se
Screen.MousePointer = 11
Combo2.Clear
txtfiltro1 = "SELECT COUNT(DISTINCT SERIE) FROM " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "'"
RMesa.Open txtfiltro1, ConAdo
noreg = RMesa.Fields(0)
RMesa.Close
If noreg <> 0 Then
txtfiltro2 = "Select SERIE from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "' GROUP BY SERIE ORDER BY SERIE"
RMesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 3) As Variant
RMesa.MoveFirst
ReDim MatSerie(1 To noreg) As String
For i = 1 To noreg
MatSerie(i) = RMesa.Fields("SERIE")
Combo2.AddItem MatSerie(i)
RMesa.MoveNext
Next i
RMesa.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub

Private Sub Combo2_Change()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se filtra la tabla y se agrupan los precios por
'emision. Dependiendo de que emision se seleccione se
Screen.MousePointer = 11
Combo3.Clear
txtfiltro1 = "SELECT COUNT(DISTINCT SERIE) FROM " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "'"
RMesa.Open txtfiltro1, ConAdo
noreg = RMesa.Fields(0)
RMesa.Close
If noreg <> 0 Then
txtfiltro2 = "Select SERIE from " & TablaVecPrecios & " WHERE TV = '" & Combo2.Text & "' AND EMISION = '" & Trim(Combo1.Text) & "' GROUP BY SERIE ORDER BY SERIE"
RMesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 3) As Variant
RMesa.MoveFirst
ReDim MatSerie(1 To noreg) As String
For i = 1 To noreg
MatSerie(i) = RMesa.Fields("SERIE")
Combo3.AddItem MatSerie(i)
RMesa.MoveNext
Next i
RMesa.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub


Private Sub Command1_Click()
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
SiActTProc = True
Screen.MousePointer = 11
Combo1.Clear
Combo2.Clear
Combo3.Clear
txtfiltro = "SELECT count(DISTINCT TV) from " & TablaVecPrecios
RMesa.Open txtfiltro, ConAdo
noreg = RMesa.Fields(0)
RMesa.Close
If noreg <> 0 Then
 txtfiltro = "SELECT TV from " & TablaVecPrecios & " GROUP BY TV ORDER BY TV"
 RMesa.Open txtfiltro, ConAdo
 RMesa.MoveFirst
 ReDim MatTV(1 To noreg) As String
 For i = 1 To noreg
 MatTV(i) = RMesa.Fields("tv")
 Combo1.AddItem MatTV(i)
 RMesa.MoveNext
 Next i
 RMesa.Close
End If
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
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim tv As String
Dim emision As String
Dim serie As String
Dim fecha1 As String
Dim fecha2 As String
Dim pcupon As String
Dim txtemision As String
Dim matfl() As Variant
Screen.MousePointer = 11
  tv = Combo1.Text
  emision = Combo1.Text
  serie = Combo2.Text
  txtemision = GeneraClaveEmision(tv, emision, serie)
  txtfiltro2 = "SELECT F_EMISION,FVENCIMIENTO FROM " & TablaVecPrecios & " WHERE CLAVEEMISION = '" & txtemision & "'"
  txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
  RMesa.Open txtfiltro2, ConAdo
  noreg = RMesa.Fields(0)
  RMesa.Close
  If noreg <> 0 Then
     RMesa.Open txtfiltro2, ConAdo
     fecha1 = RMesa.Fields("finicio")
     fecha2 = RMesa.Fields("fvencimiento")
     pcupon = RMesa.Fields("pcupon")
     RMesa.Close
     matfl = CalcFlujosEmision(fecha1, fecha2, saldo, pcupon)
     Call GuardaFlujosMD(txtemision, finicio, matfl)
  End If
  
Screen.MousePointer = 0
End Sub

