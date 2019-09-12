VERSION 5.00
Begin VB.Form frmParametros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros del sistema"
   ClientHeight    =   8655
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Datos del servidor PIP"
      Height          =   2475
      Left            =   4830
      TabIndex        =   79
      Top             =   5400
      Width           =   4605
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   200
         TabIndex        =   85
         Top             =   1950
         Width           =   4000
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   200
         TabIndex        =   83
         Top             =   1230
         Width           =   4000
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   200
         TabIndex        =   81
         Top             =   540
         Width           =   4000
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Left            =   200
         TabIndex        =   84
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   200
         TabIndex        =   82
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Servidor"
         Height          =   195
         Left            =   210
         TabIndex        =   80
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Restaurar valores predeterminados"
      Height          =   600
      Left            =   7680
      TabIndex        =   72
      Top             =   8000
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Caption         =   "Archivos de resultados"
      Height          =   7000
      Left            =   9600
      TabIndex        =   32
      Top             =   90
      Width           =   4600
      Begin VB.CommandButton Command28 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   78
         Top             =   3500
         Width           =   400
      End
      Begin VB.CommandButton Command27 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   77
         Top             =   2900
         Width           =   400
      End
      Begin VB.CommandButton Command26 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   76
         Top             =   2300
         Width           =   400
      End
      Begin VB.CommandButton Command25 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   75
         Top             =   1700
         Width           =   400
      End
      Begin VB.CommandButton Command24 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   74
         Top             =   500
         Width           =   400
      End
      Begin VB.CommandButton Command23 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   73
         Top             =   1100
         Width           =   400
      End
      Begin VB.TextBox Text21 
         Height          =   315
         Left            =   100
         TabIndex        =   65
         Top             =   3500
         Width           =   4000
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   90
         TabIndex        =   37
         Top             =   500
         Width           =   4000
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   100
         TabIndex        =   36
         Top             =   1100
         Width           =   4000
      End
      Begin VB.TextBox Text25 
         Height          =   285
         Left            =   100
         TabIndex        =   35
         Top             =   1700
         Width           =   4000
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   100
         TabIndex        =   34
         Top             =   2300
         Width           =   4000
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   100
         TabIndex        =   33
         Top             =   2900
         Width           =   4000
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Instalacion de Winzip"
         Height          =   195
         Left            =   105
         TabIndex        =   71
         Top             =   2730
         Width           =   1515
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Carpeta de resultados"
         Height          =   195
         Left            =   105
         TabIndex        =   70
         Top             =   2070
         Width           =   1545
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Reportes en Excel"
         Height          =   195
         Left            =   105
         TabIndex        =   69
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Access de Catalogos"
         Height          =   195
         Left            =   105
         TabIndex        =   68
         Top             =   870
         Width           =   1500
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Resumen Ejecutivo"
         Height          =   195
         Left            =   105
         TabIndex        =   67
         Top             =   3330
         Width           =   2205
      End
      Begin VB.Label Label22 
         Caption         =   "Directorio batch"
         Height          =   165
         Left            =   100
         TabIndex        =   66
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Archivos de posición"
      Height          =   5235
      Left            =   4800
      TabIndex        =   23
      Top             =   100
      Width           =   4600
      Begin VB.CommandButton Command22 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   64
         Top             =   4700
         Width           =   400
      End
      Begin VB.CommandButton Command21 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   63
         Top             =   4100
         Width           =   400
      End
      Begin VB.CommandButton Command20 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   62
         Top             =   3480
         Width           =   400
      End
      Begin VB.CommandButton Command19 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   61
         Top             =   2900
         Width           =   400
      End
      Begin VB.CommandButton Command18 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   60
         Top             =   2300
         Width           =   400
      End
      Begin VB.CommandButton Command17 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   59
         Top             =   1700
         Width           =   400
      End
      Begin VB.CommandButton Command16 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   58
         Top             =   1100
         Width           =   400
      End
      Begin VB.CommandButton Command15 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   57
         Top             =   500
         Width           =   400
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   100
         TabIndex        =   31
         Top             =   500
         Width           =   4000
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   100
         TabIndex        =   30
         Top             =   1100
         Width           =   4000
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   100
         TabIndex        =   29
         Top             =   1700
         Width           =   4000
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   100
         TabIndex        =   28
         Top             =   2300
         Width           =   4000
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   90
         TabIndex        =   27
         Top             =   2900
         Width           =   4000
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   100
         TabIndex        =   26
         Top             =   3500
         Width           =   4000
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   100
         TabIndex        =   25
         Top             =   4100
         Width           =   4000
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   100
         TabIndex        =   24
         Top             =   4700
         Width           =   4000
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Pensiones 2"
         Height          =   195
         Left            =   100
         TabIndex        =   56
         Top             =   4470
         Width           =   870
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Pensiones 1"
         Height          =   195
         Left            =   100
         TabIndex        =   55
         Top             =   3900
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Primaria forwards TC"
         Height          =   195
         Left            =   100
         TabIndex        =   54
         Top             =   3200
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Primaria swaps"
         Height          =   195
         Left            =   100
         TabIndex        =   53
         Top             =   2650
         Width           =   1050
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Forwards TC"
         Height          =   195
         Left            =   100
         TabIndex        =   52
         Top             =   2000
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Swaps"
         Height          =   195
         Left            =   100
         TabIndex        =   51
         Top             =   1450
         Width           =   480
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Divisas"
         Height          =   195
         Left            =   100
         TabIndex        =   50
         Top             =   850
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Mercado de Dinero"
         Height          =   195
         Left            =   100
         TabIndex        =   49
         Top             =   250
         Width           =   1365
      End
   End
   Begin VB.Frame Frame29 
      Caption         =   "Archivos de factores de riesgo"
      Height          =   7000
      Left            =   100
      TabIndex        =   2
      Top             =   100
      Width           =   4600
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   44
         Top             =   5700
         Width           =   400
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   43
         Top             =   5200
         Width           =   400
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   42
         Top             =   4600
         Width           =   400
      End
      Begin VB.CommandButton Command10 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   41
         Top             =   4100
         Width           =   400
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   315
         Left            =   4100
         TabIndex        =   40
         Top             =   3500
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   39
         Top             =   2900
         Width           =   400
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   38
         Top             =   2300
         Width           =   400
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   100
         TabIndex        =   15
         Top             =   1700
         Width           =   4000
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   100
         TabIndex        =   14
         Top             =   1100
         Width           =   4000
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   500
         Width           =   4000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   12
         Top             =   500
         Width           =   400
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   11
         Top             =   1100
         Width           =   400
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   300
         Left            =   4100
         TabIndex        =   10
         Top             =   1700
         Width           =   400
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   100
         TabIndex        =   9
         Top             =   2300
         Width           =   4000
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   100
         TabIndex        =   8
         Top             =   2900
         Width           =   4000
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   100
         TabIndex        =   7
         Top             =   3500
         Width           =   4000
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   4100
         Width           =   4000
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   100
         TabIndex        =   5
         Top             =   4600
         Width           =   4000
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   100
         TabIndex        =   4
         Top             =   5200
         Width           =   4000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   100
         TabIndex        =   3
         Top             =   5700
         Width           =   4000
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Vector analitico"
         Height          =   195
         Left            =   105
         TabIndex        =   48
         Top             =   4400
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Directorio de trabajo"
         Height          =   195
         Left            =   105
         TabIndex        =   47
         Top             =   5505
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vector analitico zip"
         Height          =   195
         Left            =   105
         TabIndex        =   46
         Top             =   5000
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   30
         Left            =   300
         TabIndex        =   45
         Top             =   5070
         Width           =   435
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Curvas zip"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Curvas"
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CSV Curvas IKOS"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vectores"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vectores zip"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   2670
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Flujos Emisiones"
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   3300
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Flujos Emisiones zip"
         Height          =   195
         Left            =   105
         TabIndex        =   16
         Top             =   3900
         Width           =   1395
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   600
      Left            =   12800
      TabIndex        =   1
      Top             =   8000
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   600
      Left            =   11100
      TabIndex        =   0
      Top             =   8000
      Width           =   1500
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'se actualizan los directorios
 DirCurvasCSV1 = frmParametros.Text18.Text
 DirCurvas = frmParametros.Text11.Text
 DirCurvasZ = frmParametros.Text14.Text
 DirVPrecios = frmParametros.Text1.Text
 DirVPreciosZ = frmParametros.Text4.Text
 DirFlujosEm = frmParametros.Text5.Text
 DirFlujosEmZ = frmParametros.Text7.Text
 DirVAnalitico = frmParametros.Text8.Text
 DirVAnaliticoZ = frmParametros.Text9.Text
 DirTemp = frmParametros.Text10.Text
 DirIndFecha = frmParametros.Text12.Text
 DirPosMesaD = frmParametros.Text2.Text
 DirPosDiv = frmParametros.Text3.Text
 DirPosSwaps = frmParametros.Text13.Text
 DirPosFwdTC = frmParametros.Text15.Text
 DirPosPrimSwaps = frmParametros.Text16.Text
 DirPosPrimFwd = frmParametros.Text17.Text
 DirPosPensiones = frmParametros.Text19.Text
 DirPosPensiones2 = frmParametros.Text20.Text

 DirArchBat = frmParametros.Text23.Text
 DirBases = frmParametros.Text24.Text
 DirReportes = frmParametros.Text25.Text
 DirResVaR = frmParametros.Text26.Text
 DirWinRAR = frmParametros.Text27.Text
 NomArchRVaR = frmParametros.Text21.Text
 NomSRVPIP = frmParametros.Text6.Text
 usersftpPIP = frmParametros.Text12.Text
 passsftpPIP = frmParametros.Text22.Text

  
 ReDim matpu(1 To 28, 1 To 2) As String
 matpu(1, 1) = "DirCurvasCSV1"
 matpu(1, 2) = DirCurvasCSV1
 matpu(2, 1) = "DirCurvas"
 matpu(2, 2) = DirCurvas
 matpu(3, 1) = "DirCurvasZ"
 matpu(3, 2) = DirCurvasZ
 matpu(4, 1) = "DirVPrecios"
 matpu(4, 2) = DirVPrecios
 matpu(5, 1) = "DirVPreciosZ"
 matpu(5, 2) = DirVPreciosZ
 matpu(6, 1) = "DirFlujosEm"
 matpu(6, 2) = DirFlujosEm
 matpu(7, 1) = "DirFlujosEmZ"
 matpu(7, 2) = DirFlujosEmZ
 matpu(8, 1) = "DirVAnalitico"
 matpu(8, 2) = DirVAnalitico
 matpu(9, 1) = "DirVAnaliticoZ"
 matpu(9, 2) = DirVAnaliticoZ
 matpu(10, 1) = "DirTemp"
 matpu(10, 2) = DirTemp
 matpu(11, 1) = "DirIndFecha"
 matpu(11, 2) = DirIndFecha
 matpu(12, 1) = "DirPosMesaD"
 matpu(12, 2) = DirPosMesaD
 matpu(13, 1) = "DirPosDiv"
 matpu(13, 2) = DirPosDiv
 matpu(14, 1) = "DirPosSwaps"
 matpu(14, 2) = DirPosSwaps
 matpu(15, 1) = "DirPosFwdTC"
 matpu(15, 2) = DirPosFwdTC
 matpu(16, 1) = "DirPosPrimSwaps"
 matpu(16, 2) = DirPosPrimSwaps
 matpu(17, 1) = "DirPosPrimFwd"
 matpu(17, 2) = DirPosPrimFwd
 matpu(18, 1) = "DirPosPensiones"
 matpu(18, 2) = DirPosPensiones
 matpu(19, 1) = "DirPosPensiones2"
 matpu(19, 2) = DirPosPensiones2
 matpu(20, 1) = "DirArchBat"
 matpu(20, 2) = DirArchBat
 matpu(21, 1) = "DirBases"
 matpu(21, 2) = DirBases
 matpu(22, 1) = "DirReportes"
 matpu(22, 2) = DirReportes
 matpu(23, 1) = "DirResVaR"
 matpu(23, 2) = DirResVaR
 matpu(24, 1) = "DirWinRAR"
 matpu(24, 2) = DirWinRAR
 matpu(25, 1) = "NomArchRVaR"
 matpu(25, 2) = NomArchRVaR
 matpu(26, 1) = "NomSRVPIP"
 matpu(26, 2) = NomSRVPIP
 matpu(27, 1) = "usersftpPIP"
 matpu(27, 2) = usersftpPIP
 matpu(28, 1) = "passsftpPIP"
 matpu(28, 2) = passsftpPIP
 Call GuardarParamUsuario(matpu)
  Unload Me
 
End Sub

Sub GuardarParamUsuario(ByRef matp() As String)
Dim noreg As Integer
Dim i As Integer
Dim txtcadena As String

ConAdo.Execute "DELETE FROM " & TablaParamUsuario & " WHERE USUARIO = '" & NomUsuario & "'"
noreg = UBound(matp, 1)
For i = 1 To noreg
    txtcadena = "INSERT INTO " & TablaParamUsuario & " VALUES("
    txtcadena = txtcadena & "'" & NomUsuario & "',"
    txtcadena = txtcadena & "'" & matp(i, 1) & "',"
    txtcadena = txtcadena & "'" & matp(i, 2) & "')"
    ConAdo.Execute txtcadena
Next i
End Sub

Private Sub Command10_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text7.Text = DirSalida
End If
End Sub

Private Sub Command11_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text8.Text = DirSalida
End If
End Sub

Private Sub Command12_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text9.Text = DirSalida
End If
End Sub

Private Sub Command13_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text10.Text = DirSalida
End If
End Sub

Private Sub Command14_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text12.Text = DirSalida
End If
End Sub

Private Sub Command15_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text2.Text = DirSalida
End If
End Sub

Private Sub Command16_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text3.Text = DirSalida
End If
End Sub

Private Sub Command17_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text13.Text = DirSalida
End If
End Sub

Private Sub Command18_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text15.Text = DirSalida
End If
End Sub

Private Sub Command19_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text16.Text = DirSalida
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command20_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text17.Text = DirSalida
End If
End Sub

Private Sub Command21_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text19.Text = DirSalida
End If
End Sub

Private Sub Command22_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text20.Text = DirSalida
End If
End Sub

Private Sub Command23_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text24.Text = DirSalida
End If
End Sub

Private Sub Command24_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text23.Text = DirSalida
End If
End Sub

Private Sub Command25_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text25.Text = DirSalida
End If
End Sub

Private Sub Command26_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text26.Text = DirSalida
End If
End Sub

Private Sub Command27_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text27.Text = DirSalida
End If
End Sub

Private Sub Command28_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text21.Text = DirSalida
End If
End Sub

Private Sub Command3_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text18.Text = DirSalida
End If
End Sub

Private Sub Command4_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text11.Text = DirSalida
End If
End Sub

Private Sub Command5_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text14.Text = DirSalida
End If
End Sub

Private Sub Command6_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text1.Text = DirSalida
End If
End Sub

Private Sub Command7_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text4.Text = DirSalida
End If
End Sub

Private Sub Command8_Click()
frmSelecDir.Show 1
If DirSalida <> "" Then
  Text5.Text = DirSalida
End If
End Sub

Private Sub Command9_Click()
Dim sicam As Integer
sicam = MsgBox("Esta accion reestablece los parametros para el ususario, desea continuar", vbYesNo)
If sicam = 6 Then
  Call GenerarParamUsuario(NomUsuario)
  Call ValidarParamUsuario(NomUsuario)
  Call CargaParamPantalla
End If
End Sub

Private Sub Form_Load()
  Call CargaParamPantalla
End Sub

Sub CargaParamPantalla()
frmParametros.Text18.Text = DirCurvasCSV1
frmParametros.Text11.Text = DirCurvas
frmParametros.Text14.Text = DirCurvasZ
frmParametros.Text1.Text = DirVPrecios
frmParametros.Text4.Text = DirVPreciosZ
frmParametros.Text5.Text = DirFlujosEm
frmParametros.Text7.Text = DirFlujosEmZ
frmParametros.Text8.Text = DirVAnalitico
frmParametros.Text9.Text = DirVAnaliticoZ
frmParametros.Text10.Text = DirTemp

frmParametros.Text12.Text = DirIndFecha

frmParametros.Text2.Text = DirPosMesaD
frmParametros.Text3.Text = DirPosDiv
frmParametros.Text13.Text = DirPosSwaps
frmParametros.Text15.Text = DirPosFwdTC
frmParametros.Text16.Text = DirPosPrimSwaps
frmParametros.Text17.Text = DirPosPrimFwd
frmParametros.Text19.Text = DirPosPensiones
frmParametros.Text20.Text = DirPosPensiones2

frmParametros.Text23.Text = DirArchBat
frmParametros.Text24.Text = DirBases
frmParametros.Text25.Text = DirReportes
frmParametros.Text26.Text = DirResVaR
frmParametros.Text27.Text = DirWinRAR
frmParametros.Text21.Text = NomArchRVaR

frmParametros.Text6.Text = NomSRVPIP
frmParametros.Text12.Text = usersftpPIP
frmParametros.Text22.Text = passsftpPIP


End Sub

