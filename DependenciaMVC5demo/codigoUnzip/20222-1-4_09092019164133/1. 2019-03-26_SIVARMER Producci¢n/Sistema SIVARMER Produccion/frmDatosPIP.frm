VERSION 5.00
Begin VB.Form frmDatosPIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descarga de Datos de PIP"
   ClientHeight    =   5220
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command3"
      Height          =   396
      Left            =   5208
      TabIndex        =   11
      Top             =   4440
      Width           =   444
   End
   Begin VB.TextBox Text4 
      Height          =   420
      Left            =   3504
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   4440
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   396
      Left            =   2832
      TabIndex        =   9
      Top             =   4464
      Width           =   444
   End
   Begin VB.TextBox Text3 
      Height          =   420
      Left            =   1152
      TabIndex        =   8
      Text            =   "c:\salida.csv"
      Top             =   4464
      Width           =   1620
   End
   Begin VB.TextBox Text2 
      Height          =   348
      Left            =   3480
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1320
      Width           =   1600
   End
   Begin VB.TextBox Text1 
      Height          =   324
      Left            =   1176
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1368
      Width           =   1600
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   1600
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   1176
      TabIndex        =   2
      Top             =   1824
      Width           =   1600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Archivo de curvas"
      Height          =   500
      Left            =   3480
      TabIndex        =   1
      Top             =   300
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vector de precios"
      Height          =   500
      Left            =   1200
      TabIndex        =   0
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha a obtener"
      Height          =   188
      Left            =   3480
      TabIndex        =   7
      Top             =   1104
      Width           =   1596
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha a obtener"
      Height          =   188
      Left            =   1176
      TabIndex        =   6
      Top             =   1128
      Width           =   1596
   End
End
Attribute VB_Name = "frmDatosPIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim fecha As Date
Dim mata() As propVecPrecios
Dim mindvp() As Variant
Dim nomarch As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim exitoarch As Boolean

Screen.MousePointer = 11
fecha = CDate(Text1.Text)
mata = LeerVPrecios(fecha, mindvp)
If UBound(mata, 1) > 0 Then
nomarch = Text3.Text
Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
If exitoarch Then
   For i = 1 To UBound(mata, 1)
   txtcadena = ""
   For j = 1 To UBound(mata, 2)
   txtcadena = txtcadena & mata(i, j) & ","
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "se guardo el archivo" & nomarch
End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim fecha As Date
Dim mata() As Variant
Dim nomarch As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim exito As Boolean
Dim exitoarch As Boolean

Screen.MousePointer = 11
fecha = CDate(Text2.Text)
mata = LeerCurvaCompleta(fecha, exito)
If UBound(mata, 1) > 0 Then
   nomarch = Text4.Text
   Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
   If exitoarch Then
   For i = 1 To UBound(mata, 1)
   txtcadena = ""
   For j = 1 To UBound(mata, 2)
   txtcadena = txtcadena & mata(i, j) & ","
   Next j
   Print #1, txtcadena
   Next i
   Close #1
   MsgBox "se guardo el archivo" & nomarch
   End If
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Text1.Text = Date
Text2.Text = Date
End Sub
