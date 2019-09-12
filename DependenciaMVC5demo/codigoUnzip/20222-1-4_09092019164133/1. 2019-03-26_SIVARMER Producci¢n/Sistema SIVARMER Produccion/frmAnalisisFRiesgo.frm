VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistFRiesgo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia de Factores de Riesgo"
   ClientHeight    =   10530
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10005
   Icon            =   "frmAnalisisFRiesgo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   10005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Factores de riesgo"
      Height          =   10080
      Left            =   96
      TabIndex        =   0
      Top             =   168
      Width           =   9750
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   8500
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   15002
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exportar historia a archivo CSV"
         Height          =   700
         Left            =   216
         TabIndex        =   1
         Top             =   216
         Width           =   1500
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8500
         Left            =   4992
         TabIndex        =   2
         Top             =   1200
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   15002
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lista de factores"
         Height          =   192
         Left            =   288
         TabIndex        =   5
         Top             =   1000
         Width           =   1176
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Historia"
         Height          =   192
         Left            =   5040
         TabIndex        =   3
         Top             =   1000
         Width           =   552
      End
   End
End
Attribute VB_Name = "frmHistFRiesgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
Dim nofilas As Integer
Dim nocols As Integer
Dim inicio As Integer
Dim final As Integer
Dim fechaa As String
Dim fechab As String
Dim nomarch As String
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim exitoarch As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
nofilas = UBound(MatFactRiesgo, 1)
nocols = UBound(MatFactRiesgo, 2)
inicio = 1
final = nocols - 1

fechaa = Format(MatFactRiesgo(1, 1), "yyyy-mm-dd")
fechab = Format(MatFactRiesgo(nofilas, 1), "yyyy-mm-dd")
nomarch = DirResVaR & "\Hist FR " & fechaa & " - " & fechab & ".csv"
frmCalVar.CommonDialog1.FileName = nomarch
frmCalVar.CommonDialog1.ShowSave
nomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
txtcadena = "fecha,"
For i = inicio To final
txtcadena = txtcadena & MatCaracFRiesgo(i).nomFactor & ","
Next i
Print #1, txtcadena
txtcadena = ","
For i = inicio To final
txtcadena = txtcadena & MatCaracFRiesgo(i).plazo & ","
Next i
Print #1, txtcadena
'se guarda la historia de los valores
For i = 1 To nofilas
txtcadena = MatFactRiesgo(i, 1) & ","
For j = inicio To final
 txtcadena = txtcadena & MatFactRiesgo(i, j + 1) & ","
Next j
Print #1, txtcadena
Next i
Close #1
MsgBox "se creo el archivo " & nomarch
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim noreg As Integer
Dim i As Integer


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
noreg = UBound(MatCaracFRiesgo, 1)
MSFlexGrid2.Rows = noreg + 1
MSFlexGrid2.ColWidth(1) = 3000
For i = 1 To noreg
 MSFlexGrid2.TextMatrix(i, 1) = MatCaracFRiesgo(i).indFactor
Next i
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmCalVar.Show
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub MSFlexGrid2_DblClick()
Dim noreg As Integer
Dim i As Integer
Dim indice As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
indice = MSFlexGrid2.row

noreg = UBound(MatFactRiesgo, 1)
MSFlexGrid1.Rows = noreg + 1
MSFlexGrid1.Cols = 2
MSFlexGrid1.TextMatrix(0, 1) = MatCaracFRiesgo(indice).indFactor
For i = 1 To noreg
 MSFlexGrid1.TextMatrix(i, 0) = MatFactRiesgo(i, 1)
 MSFlexGrid1.TextMatrix(i, 1) = MatFactRiesgo(i, indice + 1)
Next i

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
