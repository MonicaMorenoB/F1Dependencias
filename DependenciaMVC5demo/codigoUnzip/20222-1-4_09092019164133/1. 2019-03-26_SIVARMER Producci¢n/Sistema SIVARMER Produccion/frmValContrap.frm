VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmValContrap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valuacion por contraparte"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6585
      Left            =   180
      TabIndex        =   4
      Top             =   1770
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11615
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exportar a archivo de texto"
      Height          =   700
      Left            =   7350
      TabIndex        =   3
      Top             =   360
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar reporte"
      Height          =   700
      Left            =   4740
      TabIndex        =   2
      Top             =   360
      Width           =   1500
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1830
      TabIndex        =   1
      Top             =   390
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   390
      Width           =   1605
   End
End
Attribute VB_Name = "frmValContrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nofechas1 As Integer
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim noreg1 As Integer
Dim matv() As Variant
Dim matcont() As Variant

Screen.MousePointer = 11
nofechas1 = 7
ReDim matf(1 To nofechas1, 1 To 1) As Date
matf(1, 1) = #3/31/2015#
matf(2, 1) = #6/30/2015#
matf(3, 1) = #9/30/2015#
matf(4, 1) = #12/31/2015#
matf(5, 1) = #1/29/2016#
matf(6, 1) = #2/29/2016#
matf(7, 1) = #3/31/2016#
noreg1 = UBound(MatContrapartes, 1)
matcont = ObtFactUnicos(MatContrapartes, 5)
ReDim matres(1 To noreg1, 1 To nofechas1 + 1) As Variant
For i = 1 To UBound(matcont, 1)
    matres(i, 1) = matcont(i, 1)
Next i
For i = 1 To nofechas1
    matv = LeerValContraparte(matf(i, 1), 1)
    For j = 1 To noreg1
        For p = 1 To UBound(matv, 1)
            If matv(p, 2) = matres(j, 1) Then
               matres(j, i + 1) = matv(p, 7) / 1000000
            End If
        Next p
    Next j
Next i
MSFlexGrid1.Rows = UBound(matcont, 1) + 1
MSFlexGrid1.Cols = nofechas1 + 1
For i = 1 To nofechas1
MSFlexGrid1.TextMatrix(0, i) = matf(i, 1)
Next i
For i = 1 To UBound(matcont, 1)
For j = 1 To nofechas1 + 1
MSFlexGrid1.TextMatrix(i, j - 1) = matres(i, j)
Next j
Next i
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

Screen.MousePointer = 11
Open "d:\resultados.txt" For Output As #1
For i = 1 To MSFlexGrid1.Rows
txtcadena = ""
For j = 1 To MSFlexGrid1.Cols
txtcadena = txtcadena & MSFlexGrid1.TextMatrix(i - 1, j - 1) & Chr(9)
Next j
Print #1, txtcadena
Next i
Close #1
Screen.MousePointer = 0
End Sub

