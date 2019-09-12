VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistEfcob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia eficiencia cobertura"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9975
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   12795
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar historia archivo texto"
         Height          =   705
         Left            =   2850
         TabIndex        =   2
         Top             =   375
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   555
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   7665
         Left            =   120
         TabIndex        =   3
         Top             =   1605
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   13520
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
         Height          =   7665
         Left            =   6600
         TabIndex        =   4
         Top             =   1635
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   13520
         _Version        =   393216
      End
      Begin VB.Label Label8 
         Caption         =   "Eficiencia Prospectiva"
         Height          =   285
         Left            =   6735
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Eficiencia Retrospectiva"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1230
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No de Derivado"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   240
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmHistEfcob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo2_Click()
Dim noswap  As Integer

Screen.MousePointer = 11
 noswap = Combo2.Text
 Call ObtenerHistEficiencia(noswap)
Screen.MousePointer = 0

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Dim noswap As Integer
Dim mata() As Variant
Dim i As Integer
Dim j As Integer
Dim noreg As Integer

If KeyAscii = 13 Then
Screen.MousePointer = 11
 noswap = Combo2.Text
 mata = ObtenerHistEficiencia(noswap)
 If UBound(mata, 1) > 0 Then
 noreg = UBound(mata, 1)
 MSFlexGrid4.Rows = 2
 MSFlexGrid4.Cols = 2
 MSFlexGrid4.Rows = noreg + 1
 MSFlexGrid4.Cols = 19
 For i = 1 To noreg
 For j = 1 To 18
  If Len(Trim(mata(i, j))) <> 0 Then
  MSFlexGrid4.TextMatrix(i, j) = mata(i, j)
  End If
 Next j
 Next i
 Else
 MsgBox "No hay datos de eficiencia para la operación " & noswap
 End If
Screen.MousePointer = 0
End If
End Sub

Private Sub Command1_Click()
Dim noswap As Integer
Dim nomarch As String
Dim mata() As Variant
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim noreg As Integer
Dim nofilas As Integer
Dim exitoarch As Boolean

Screen.MousePointer = 11
 noswap = Trim(Combo2.Text)
 nomarch = "Eficiencia retro Swap " & noswap & ".txt"
 frmCalVar.CommonDialog1.FileName = nomarch
 frmCalVar.CommonDialog1.ShowSave
 nomarch = frmCalVar.CommonDialog1.FileName
 mata = ObtenerHistEficiencia(noswap)
 noreg = UBound(mata, 1)
 nofilas = UBound(mata, 2)
 Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
 If exitoarch Then
 For i = 1 To noreg
 txtcadena = ""
  For j = 1 To nofilas
   txtcadena = txtcadena & mata(i, j) & Chr(9)
  Next j
 Print #1, txtcadena
 Next i
 Close #1
 End If
Screen.MousePointer = 0
End Sub

