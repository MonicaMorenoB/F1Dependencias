VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepEfPros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de eficiencia prospectiva"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3420
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar reporte"
      Height          =   615
      Left            =   2730
      TabIndex        =   2
      Top             =   1290
      Width           =   1605
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha - Clave de operación"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   1965
   End
End
Attribute VB_Name = "frmRepEfPros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim rmesa As New ADODB.recordset
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim coperacion As String
Dim fecha As Date
Dim noreg As Integer
Dim i As Integer
Dim txtnomarch As String
Dim txtcadena As String
Dim suma1 As Double
Dim suma2 As Double
If Not EsVariableVacia(Combo1.Text) Then
Screen.MousePointer = 11
coperacion = extCoperCadena(Combo1.Text)
fecha = CDate(extFechaCadena(Combo1.Text))

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaResEfectPros & " WHERE F_CALCULO = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "' ORDER BY FECHA1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   suma1 = 0
   suma2 = 0
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(3)
       mata(i, 2) = rmesa.Fields(4)
       mata(i, 3) = rmesa.Fields(5)
       suma1 = suma1 + mata(i, 2)
       suma2 = suma2 + mata(i, 3)
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtnomarch = DirResVaR & "\Resultados efect prospectiva " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmRepEfPros.CommonDialog1.FileName = txtnomarch
   frmRepEfPros.CommonDialog1.ShowSave
   txtnomarch = frmRepEfPros.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   For i = 1 To noreg
       If (i - 1) Mod 45 = 0 Then Print #1, "Fecha" & Chr(9) & "No. de simulaciones" & Chr(9) & "No. de aciertos"
       txtcadena = mata(i, 1) & Chr(9) & mata(i, 2) & Chr(9) & mata(i, 3)
       Print #1, txtcadena
       If i Mod 45 = 0 Then Print #1, ""
   Next i
   If suma1 <> 0 Then
      Print #1, "Porcentaje de aciertos" & Chr(9) & Chr(9) & Format(suma2 / suma1, "##0.00 %")
   Else
      Print #1, "Porcentaje de aciertos" & Chr(9) & Chr(9) & Format(suma2 / suma1, "##0.00 %")
   End If
   Close #1
   MsgBox "Fin de proceso"
   Screen.MousePointer = 0
Else
   MsgBox "No se encontraron datos"
End If
End If
End Sub

Function extCoperCadena(txtcadena)
Dim i As Integer
Dim largo As Integer
For i = 1 To Len(txtcadena)
    If Mid(txtcadena, i, 1) = " " Then
       largo = Len(txtcadena) - i
       extCoperCadena = Right(txtcadena, largo)
       Exit Function
    End If
Next i
End Function

Function extFechaCadena(txtcadena)
Dim i As Integer
Dim largo As Integer
For i = 1 To Len(txtcadena)
    If Mid(txtcadena, i, 1) = " " Then
       largo = Len(txtcadena) - i
       extFechaCadena = Mid(txtcadena, 1, i - 1)
       Exit Function
    End If
Next i
End Function


Private Sub Form_Load()
Dim rmesa As New ADODB.recordset
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
txtfiltro2 = "SELECT F_CALCULO,COPERACION FROM " & TablaResEfectPros & " GROUP BY F_CALCULO,COPERACION ORDER BY F_CALCULO,COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("F_CALCULO")
       mata(i, 2) = rmesa.Fields("COPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   Combo1.Clear
   For i = 1 To noreg
       Combo1.AddItem mata(i, 1) & " " & mata(i, 2)
   Next i
  
End If
End Sub
