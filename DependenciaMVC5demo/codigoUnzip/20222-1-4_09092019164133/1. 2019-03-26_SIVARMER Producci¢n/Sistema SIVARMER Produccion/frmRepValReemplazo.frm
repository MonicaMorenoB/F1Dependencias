VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepValReemplazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte del valor de reemplazo"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2910
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   585
      Left            =   510
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Top             =   450
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha del reporte"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   210
      Width           =   1245
   End
End
Attribute VB_Name = "frmRepValReemplazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dtfecha As Date
If IsDate(Combo1.Text) Then
   dtfecha = CDate(Combo1.Text)
   Call GenRepValReemplazo(dtfecha)
   Unload Me
End If
End Sub

Sub GenRepValReemplazo(ByVal dtfecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtnomarch As String
Dim txtcadena As String
Dim txttabla As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim matb() As Variant
Dim rmesa As New ADODB.recordset
Dim exitoarch As Boolean

txtfecha = "to_date('" & Format(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResVReemplazo & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 7) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(3)
       mata(i, 2) = rmesa.Fields(2)
       mata(i, 3) = rmesa.Fields(4)
       mata(i, 4) = rmesa.Fields(5)
       mata(i, 5) = rmesa.Fields(6)
       mata(i, 6) = rmesa.Fields(7)
       mata(i, 7) = rmesa.Fields(8)
       rmesa.MoveNext
   Next i
   rmesa.Close
txtfecha = "to_date('" & Format(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResCalcVReemplazo & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg1 <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matb(1 To noreg1, 1 To 3) As Variant
   For i = 1 To noreg1
       matb(i, 1) = rmesa.Fields(1)
       matb(i, 2) = rmesa.Fields(9)
       matb(i, 3) = rmesa.Fields(10)
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
   
   
   txtnomarch = DirResVaR & "\Rep Val Reemplazo " & Format(dtfecha, "YYYY-MM-DD") & ".txt"
   CommonDialog1.FileName = txtnomarch
   CommonDialog1.ShowSave
   txtnomarch = CommonDialog1.FileName
   Call VerificarSalidaArchivo(txtnomarch, 1, exitoarch)
   If exitoarch Then
   Print #1, "Sector" & Chr(9) & "Contraparte" & Chr(9) & "mtm t" & Chr(9) & "mtm t+1" & Chr(9) & "CVaR" & Chr(9) & "Anti CVaR" & Chr(9) & "Valor de reemplazo"
   For i = 1 To noreg
       txtcadena = ""
       For j = 1 To 7
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   If noreg1 <> 0 Then
      Print #1, ""
      Print #1, "Tasa de fondeo: " & Chr(9) & Format(matb(1, 1), "##0.00 %")
      Print #1, "Costo reemplazo val pos: " & Chr(9) & Format(matb(1, 2) / 1000000, "###,###,##0.00")
      Print #1, "Costo reemplazo val neg: " & Chr(9) & Format(matb(1, 3) / 1000000, "###,###,##0.00")
   End If
   Close #1
   MsgBox "Se genero el archivo " & txtnomarch
   End If
End If
End Sub


Private Sub Form_Load()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim txttabla As String
Dim i As Integer
Dim rmesa As New ADODB.recordset

 txtfiltro2 = "SELECT FECHA FROM " & TablaResVReemplazo & " GROUP BY FECHA ORDER BY FECHA DESC"
 txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       Combo1.AddItem rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub
