VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPPrimaria 
   Caption         =   "Posiciones Primarias"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9120
      Left            =   192
      TabIndex        =   0
      Top             =   288
      Width           =   10545
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1395
         Left            =   3630
         TabIndex        =   7
         Top             =   840
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2461
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalles de la posicion primaria"
         Height          =   5625
         Left            =   420
         TabIndex        =   6
         Top             =   3030
         Width           =   9465
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   3795
            Left            =   300
            TabIndex        =   14
            Top             =   1380
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   6694
            _Version        =   393216
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   6060
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   600
            Width           =   2205
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   3000
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   690
            Width           =   2295
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   300
            TabIndex        =   9
            Text            =   "Combo2"
            Top             =   720
            Width           =   2235
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "% de cobertura de tasa"
            Height          =   195
            Left            =   6030
            TabIndex        =   13
            Top             =   330
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Porcentaje de cobertura de saldos"
            Height          =   165
            Left            =   2940
            TabIndex        =   10
            Top             =   390
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1335
         Left            =   390
         TabIndex        =   3
         Top             =   1380
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "Fwd"
            Height          =   255
            Left            =   270
            TabIndex        =   5
            Top             =   810
            Width           =   1845
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Swap"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   330
            Width           =   2325
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   450
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Posiciones Primarias asociadas"
         Height          =   195
         Left            =   3630
         TabIndex        =   8
         Top             =   570
         Width           =   2205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave de IKOs"
         Height          =   195
         Left            =   450
         TabIndex        =   2
         Top             =   480
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmPPrimaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nomarch As String
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim matd() As Variant
Dim sihayarch As Boolean
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim nocolum As Integer
Dim contar As Integer
Dim txttexto As String
Dim exitoarch As Boolean

Screen.MousePointer = 11
nomarch = DirResVaR & "\Tabla_Amort_feb2009.txt"
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
mata = LeerArchTexto(nomarch, ",", "")
'se debe de crear una clave unica de ordenacion de la tabla que permita ordenar los flujos por fecha
noreg = UBound(mata, 1)
nocolum = UBound(mata, 2)
ReDim matb(1 To noreg, 1 To 6) As Variant
'se agrega una columna de ordenacion de los flujos
For i = 1 To noreg
 matb(i, 1) = mata(i, 1)                         'clave del credito
 matb(i, 3) = CDate(mata(i, 5))                  'fecha de fin del flujo
 matb(i, 5) = Val(mata(i, 6))                    'amortizacion
 matb(i, 6) = matb(i, 1) & CLng(matb(i, 3))      'clave de ordenacion
Next i
'se procede a ordenar los flujos con la clave de ordenacion
Screen.MousePointer = 11
mata = RutinaOrden(mata, 1, SRutOrden)
matb = RutinaOrden(matb, 6, SRutOrden)
'se obtiene la clave de los creditos en la tabla
matc = ObtFactUnicos(mata, 1)
noreg1 = UBound(matc, 1)
ReDim matd(1 To 3, 1 To 1) As Variant
'se hace una exploracion de toda la tabla para encontrar el inicio y el final de
'cada uno de los creditos
contar = 0
For i = 1 To noreg
'solo hay de tres sabores i=1 i entre 1 y noreg e i=noreg
If i = 1 Then
contar = contar + 1
ReDim Preserve matd(1 To 3, 1 To contar) As Variant
'se determina el inicio del primer credito
matd(1, contar) = mata(i, 1)
matd(2, contar) = i
ElseIf i > 1 And i < noreg Then
'se determina el inicio del credito en i y el final del credito anterior
If mata(i - 1, 1) <> mata(i, 1) Then
 contar = contar + 1
 ReDim Preserve matd(1 To 3, 1 To contar) As Variant
 matd(1, contar) = mata(i, 1)
 matd(2, contar) = i
 matd(3, contar - 1) = i - 1
End If
ElseIf i = noreg Then

If mata(i - 1, 1) <> mata(i, 1) Then
 contar = contar + 1
 ReDim Preserve matd(1 To 3, 1 To contar) As Variant
 matd(1, contar) = mata(i, 1)
 matd(2, contar) = i
 matd(3, contar) = i
 matd(3, contar - 1) = i - 1
Else
 matd(3, contar) = i
End If

End If
Next i
matd = MTranV(matd)
noreg2 = UBound(matd, 1)
'ahora se procede a reconstruir los flujos de los creditos, ya que solo llega
For i = 1 To noreg2
For j = matd(i, 3) To matd(i, 2) Step -1
If j = matd(i, 3) Then
 matb(j, 4) = matb(j, 5)                      'saldo al final
Else
 matb(j, 4) = matb(j, 5) + matb(j + 1, 4)     'saldo intermedio
End If
If j <> matd(i, 2) Then
 matb(j, 2) = matb(j - 1, 3)
End If
Next j
Next i
Call VerificarSalidaArchivo(DirResVaR & "\salida.txt", 1, exitoarch)
If exitoarch Then
For i = 1 To noreg
txttexto = ""
For j = 1 To 6
 txttexto = txttexto & matb(i, j) & Chr(9)
Next j
Print #1, txttexto
Next i
Close #1
MsgBox "Proceso terminado"
End If
End If
Screen.MousePointer = 0

End Sub

