VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMonitorValidación 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor validaciones intradia"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6885
      Left            =   195
      TabIndex        =   3
      Top             =   1860
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   12144
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Efectividad"
      TabPicture(0)   =   "frmAnalisisECP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CVaR Intradia"
      TabPicture(1)   =   "frmAnalisisECP.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSFlexGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlexGrid3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   2925
         Left            =   4000
         TabIndex        =   6
         Top             =   500
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5159
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   6225
         Left            =   195
         TabIndex        =   5
         Top             =   495
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   10980
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6015
         Left            =   -74730
         TabIndex        =   4
         Top             =   720
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   10610
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   585
      Left            =   300
      TabIndex        =   0
      Top             =   990
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   180
      Width           =   450
   End
End
Attribute VB_Name = "frmMonitorValidación"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fecha As Date
Dim txtfecha As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer
Dim i As Integer
Screen.MousePointer = 11
fecha = CDate(Combo1.Text)
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CLAVE_SWAP IN (SELECT COPERACION FROM " & TablaResEfectPros & " WHERE F_CALCULO = " & txtfecha & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("CLAVE_SWAP")
       mata(i, 2) = rmesa.Fields("EFIC_PRO")
       rmesa.MoveNext
   Next i
   rmesa.Close
   MSFlexGrid1.Rows = noreg + 1
   MSFlexGrid1.Cols = 2
   MSFlexGrid1.TextMatrix(0, 0) = "Clave de operacion"
   MSFlexGrid1.TextMatrix(0, 1) = "Efectividad calculada"
   MSFlexGrid1.ColWidth(0) = 1000
   MSFlexGrid1.ColWidth(1) = 3000
   For i = 1 To noreg
       MSFlexGrid1.TextMatrix(i, 0) = mata(i, 1)
       MSFlexGrid1.TextMatrix(i, 1) = mata(i, 2)
   Next i
End If
txtfiltro2 = "SELECT * FROM " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha, "YYYYMMDD")
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
MSFlexGrid2.Rows = 1
MSFlexGrid2.Cols = 1
MSFlexGrid2.Rows = 2
MSFlexGrid2.Cols = 2
MSFlexGrid3.Rows = 1
MSFlexGrid3.Cols = 1
MSFlexGrid3.Rows = 2
MSFlexGrid3.Cols = 2
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matb(1 To noreg, 1 To 9) As Variant
   For i = 1 To noreg
       matb(i, 1) = rmesa.Fields("CVESWAP")
       matb(i, 2) = rmesa.Fields("VAR_GLOBAL")
       matb(i, 3) = rmesa.Fields("VAR_PORTAFOLIO")
       matb(i, 4) = rmesa.Fields("VAR_ESTRUCTURAL")
       matb(i, 5) = rmesa.Fields("VAR_RECLASIFICA")
       matb(i, 6) = rmesa.Fields("LIM_VAR_GLOBAL")
       matb(i, 7) = rmesa.Fields("LIM_VAR_PORTAFOLIO")
       matb(i, 8) = rmesa.Fields("LIM_VAR_ESTRUCTURAL")
       matb(i, 9) = rmesa.Fields("LIM_VAR_RECLASIFICA")
       rmesa.MoveNext
   Next i
   rmesa.Close
   MSFlexGrid2.Rows = noreg + 1
   MSFlexGrid2.Cols = 2
   MSFlexGrid2.TextMatrix(0, 0) = "Orden"
   MSFlexGrid2.TextMatrix(0, 1) = "Clave de operación"
   MSFlexGrid2.ColWidth(0) = 1000
   MSFlexGrid2.ColWidth(1) = 2000
   For i = 1 To noreg
   MSFlexGrid2.TextMatrix(i, 0) = i
   MSFlexGrid2.TextMatrix(i, 1) = matb(i, 1)
   Next i

   MSFlexGrid3.Rows = 5
   MSFlexGrid3.Cols = 4
   MSFlexGrid3.ColWidth(0) = 3000
   MSFlexGrid3.ColWidth(1) = 2000
   MSFlexGrid3.ColWidth(2) = 2000
   MSFlexGrid3.ColWidth(3) = 2000
   MSFlexGrid3.TextMatrix(0, 0) = "Portafolio"
   MSFlexGrid3.TextMatrix(1, 0) = "Banobras"
   MSFlexGrid3.TextMatrix(2, 0) = "Derivados de Negociación"
   MSFlexGrid3.TextMatrix(3, 0) = "Derivados Estructurales"
   MSFlexGrid3.TextMatrix(4, 0) = "Deriv Negociación por Reclasificación"
   MSFlexGrid3.TextMatrix(0, 1) = "CVaR"
   MSFlexGrid3.TextMatrix(0, 2) = "Límite de CVaR"
   MSFlexGrid3.TextMatrix(0, 3) = "Consumo de límite"
   MSFlexGrid3.TextMatrix(1, 1) = Format(matb(1, 2) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(2, 1) = Format(matb(1, 3) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(3, 1) = Format(matb(1, 4) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(4, 1) = Format(matb(1, 5) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(1, 2) = Format(matb(1, 6) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(2, 2) = Format(matb(1, 7) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(3, 2) = Format(matb(1, 8) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(4, 2) = Format(matb(1, 9) / 1000000, "###,###,###,###,###,##0.00")
   MSFlexGrid3.TextMatrix(1, 3) = Format(matb(1, 2) / matb(1, 6), "##0.00 %")
   MSFlexGrid3.TextMatrix(2, 3) = Format(matb(1, 3) / matb(1, 7), "##0.00 %")
   MSFlexGrid3.TextMatrix(3, 3) = Format(matb(1, 4) / matb(1, 8), "##0.00 %")
   MSFlexGrid3.TextMatrix(4, 3) = Format(matb(1, 5) / matb(1, 9), "##0.00 %")
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
  Combo1.Text = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
SiActTProc = False
End Sub
