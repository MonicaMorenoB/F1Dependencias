VERSION 5.00
Begin VB.Form frmGenDatosEm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar flujos de emision"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   675
      Left            =   5280
      TabIndex        =   3
      Top             =   2100
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar flujos emision"
      Height          =   675
      Left            =   3150
      TabIndex        =   2
      Top             =   2040
      Width           =   1785
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   750
      TabIndex        =   1
      Top             =   2070
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1485
      Left            =   240
      TabIndex        =   0
      Top             =   210
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   4800
         TabIndex        =   9
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión"
         Height          =   195
         Left            =   2550
         TabIndex        =   5
         Top             =   510
         Width           =   540
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
Attribute VB_Name = "frmGenDatosEm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
Combo2.Clear

txtfiltro2 = "SELECT EMISION from " & TablaVecPrecios & " WHERE TV = '" & Combo1.Text & "' GROUP BY EMISION ORDER BY EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   ReDim MatEmision(1 To noreg) As String
   For i = 1 To noreg
       MatEmision(i) = rmesa.Fields("EMISION")
       Combo2.AddItem MatEmision(i)
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo2_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se filtra la tabla y se agrupan los precios por
'emision. Dependiendo de que emision se seleccione se
Screen.MousePointer = 11
Combo3.Clear
txtfiltro2 = "Select SERIE from " & TablaVecPrecios & " WHERE TV = '" & Combo1.Text & "' AND EMISION = '" & Trim(Combo2.Text) & "' GROUP BY SERIE ORDER BY SERIE"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 3) As Variant
rmesa.MoveFirst
ReDim MatSerie(1 To noreg) As String
For i = 1 To noreg
MatSerie(i) = rmesa.Fields("SERIE")
Combo3.AddItem MatSerie(i)
rmesa.MoveNext
Next i
rmesa.Close
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
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
SiActTProc = True
Screen.MousePointer = 11
Combo1.Clear
Combo2.Clear
Combo3.Clear
txtfiltro = "SELECT count(DISTINCT TV) from " & TablaVecPrecios
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro = "SELECT TV from " & TablaVecPrecios & " GROUP BY TV ORDER BY TV"
 rmesa.Open txtfiltro, ConAdo
 rmesa.MoveFirst
 ReDim MatTV(1 To noreg) As String
 For i = 1 To noreg
 MatTV(i) = rmesa.Fields("tv")
 Combo1.AddItem MatTV(i)
 rmesa.MoveNext
 Next i
 rmesa.Close
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
Dim fecha1 As Date
Dim fecha2 As Date
Dim pcupon As Integer
Dim noreg As Integer
Dim txtemision As String
Dim matfl() As Variant
Dim saldo As Double
Dim rmesa As New ADODB.recordset

Screen.MousePointer = 11
  tv = Combo1.Text
  emision = Combo2.Text
  serie = Combo3.Text
  txtemision = GeneraClaveEmision(tv, emision, serie)
  txtfiltro2 = "SELECT FEMISION,FVENCIMIENTO,PCUPON,VNOMINAL FROM " & TablaVecPrecios & " WHERE CLAVE_EMISION = '" & txtemision & "'"
  txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
  rmesa.Open txtfiltro1, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
  If noreg <> 0 Then
     rmesa.Open txtfiltro2, ConAdo
     fecha1 = rmesa.Fields("FEMISION")
     fecha2 = rmesa.Fields("FVENCIMIENTO")
     saldo = rmesa.Fields("VNOMINAL")
     pcupon = 182
     rmesa.Close
     If (fecha2 - fecha1) Mod pcupon = 0 Then
        matfl = CalcFlujosEmision(fecha1, fecha2, saldo, pcupon)
        finicio = fecha1
        Call GuardaFlujosMD(txtemision, finicio, matfl)
     Else
       MsgBox "No puedo calcular cupones irregulares"
     End If
  End If
  
Screen.MousePointer = 0
End Sub

