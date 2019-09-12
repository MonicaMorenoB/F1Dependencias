VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVaRMark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del VaR Markowitz"
   ClientHeight    =   7140
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6660
      Left            =   144
      TabIndex        =   0
      Top             =   312
      Width           =   11004
      _ExtentX        =   19394
      _ExtentY        =   11748
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "Rendimientos"
      TabPicture(0)   =   "frmVaRMark.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Covarianzas"
      TabPicture(1)   =   "frmVaRMark.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSFlexGrid3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Sensibilidades"
      TabPicture(2)   =   "frmVaRMark.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Estadisticas muestra"
      TabPicture(3)   =   "frmVaRMark.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSFlexGrid15"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   6012
         Left            =   -74880
         TabIndex        =   1
         Top             =   504
         Width           =   10764
         _ExtentX        =   18997
         _ExtentY        =   10583
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   5130
         Left            =   120
         TabIndex        =   2
         Top             =   1350
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   9049
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   6156
         Left            =   -74856
         TabIndex        =   3
         Top             =   384
         Width           =   10644
         _ExtentX        =   18759
         _ExtentY        =   10848
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid15 
         Height          =   5892
         Left            =   -74832
         TabIndex        =   4
         Top             =   576
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   10398
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
End
Attribute VB_Name = "frmVaRMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Call frmVaRMark.LlenarMatricesVaRMark
   Call frmVaRMark.MuestraMatrizCov(MatCovar1)
End Sub



Sub MuestraMatrizCov(matriz)
Dim i As Integer
Dim j As Integer
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(matriz, 2)
MSFlexGrid3.Rows = noreg + 1
MSFlexGrid3.Cols = noreg + 1
For i = 1 To noreg
MSFlexGrid3.TextMatrix(i, 0) = MatSensib1(1, i)
MSFlexGrid3.TextMatrix(0, i) = MatSensib1(1, i)
For j = 1 To noreg
MSFlexGrid3.TextMatrix(i, j) = matriz(i, j)
Next j
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub LlenarMatricesVaRMark()
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim noreg2 As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(MatCovar1, 1)
MSFlexGrid2.Rows = noreg + 1
MSFlexGrid2.Cols = 4
MSFlexGrid2.TextMatrix(0, 0) = "Factor"
MSFlexGrid2.TextMatrix(0, 1) = "Valor"
MSFlexGrid2.TextMatrix(0, 2) = "Volaltilidad"
MSFlexGrid2.TextMatrix(0, 3) = "Sensibilidad Factor"
MSFlexGrid2.ColWidth(0) = 1500
MSFlexGrid2.ColWidth(1) = 1500
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid2.ColWidth(3) = 1500
For i = 1 To noreg
    MSFlexGrid2.TextMatrix(i, 0) = MatSensib1(1, i)   'nombre del factor
    MSFlexGrid2.TextMatrix(i, 1) = MatSensib6(1, i)   'valor del factor
    MSFlexGrid2.TextMatrix(i, 2) = MatSensib6(1, i) * (MatCovar1(i, i)) ^ 0.5 * 100
    MSFlexGrid2.TextMatrix(i, 3) = Format(MatSensNum(1, i) / 1000, "###,###,###,###,##0.0000000")
Next i

noreg1 = UBound(MatRendimientos, 1)
noreg2 = UBound(MatRendimientos, 2)
MSFlexGrid4.Rows = noreg1 + 1
MSFlexGrid4.Cols = noreg2 + 1
For i = 1 To noreg2
MSFlexGrid4.TextMatrix(0, i) = MatSensib1(1, i)
Next i
For i = 1 To noreg1
MSFlexGrid4.TextMatrix(i, 0) = i
 For j = 1 To noreg2
  MSFlexGrid4.TextMatrix(i, j) = MatRendimientos(i, j)
 Next j
Next i

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
