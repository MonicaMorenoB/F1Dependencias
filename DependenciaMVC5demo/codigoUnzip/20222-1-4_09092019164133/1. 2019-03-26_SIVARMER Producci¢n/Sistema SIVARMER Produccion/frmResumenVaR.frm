VERSION 5.00
Begin VB.Form frmResumenVaR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia de VaR"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Obtener historia"
      Height          =   735
      Left            =   4020
      TabIndex        =   8
      Top             =   2700
      Width           =   2145
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   200
      TabIndex        =   6
      Top             =   3090
      Width           =   1905
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   200
      TabIndex        =   5
      Top             =   2340
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   200
      TabIndex        =   3
      Top             =   1260
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   390
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      Height          =   195
      Left            =   200
      TabIndex        =   7
      Top             =   2790
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial"
      Height          =   195
      Left            =   200
      TabIndex        =   4
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Agrupacion"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   930
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Portafolio"
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmResumenVaR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtport As String
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim nconf As Double
Dim noesc As Integer
Dim htiempo As Integer
Dim txtcadena As String
Dim i As Integer
Dim siesfv As Boolean
Dim valor As Double
Dim txtgrupoport As String

Screen.MousePointer = 11
txtport = Combo1.Text
txtgrupoport = Combo2.Text
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
fecha1 = CDate(Combo3.Text)
fecha2 = CDate(Combo4.Text)
nconf = 0.97
noesc = 500
htiempo = 1
Open DirResVaR & "\RESUMEN CVaR.txt" For Output As #1
txtcadena = "fecha" & Chr(9)
For i = 1 To UBound(MatGruposPortPos, 1)
   txtcadena = txtcadena & MatGruposPortPos(i, 3) & Chr(9)
Next i
Print #1, txtcadena
fecha = fecha1
Do While fecha <= fecha2
   siesfv = EsFechaVaR(fecha)
   If siesfv Then
      txtcadena = fecha & Chr(9)
      For i = 1 To UBound(MatGruposPortPos, 1)
          valor = LeerCVaRHist(fecha, txtport, MatGruposPortPos(i, 3), 1 - nconf, noesc, htiempo)
          txtcadena = txtcadena & valor & Chr(9)
      Next i
      Print #1, txtcadena
   End If
   fecha = fecha + 1
Loop
Close #1
MsgBox "Fin de proceso"
Unload Me
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim noreg3 As Integer

Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
noreg1 = UBound(MatPortPosicion, 1)
If noreg1 <> 0 Then
   For i = 1 To noreg1
       Combo1.AddItem MatPortPosicion(i, 2)
   Next i
End If
noreg2 = UBound(MatListaPortPos, 1)
For i = 1 To noreg2
 Combo2.AddItem MatListaPortPos(i, 1)
Next i
noreg3 = UBound(MatFechasVaR, 1)
For i = 1 To noreg3
    Combo3.AddItem MatFechasVaR(noreg3 - i + 1, 1)
    Combo4.AddItem MatFechasVaR(noreg3 - i + 1, 1)
Next i
End Sub
