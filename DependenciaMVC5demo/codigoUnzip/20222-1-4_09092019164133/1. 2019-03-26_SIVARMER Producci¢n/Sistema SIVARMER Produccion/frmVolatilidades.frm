VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVolatilidades 
   Caption         =   "Análisis de Volatilidades"
   ClientHeight    =   8310
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11880
   Icon            =   "frmVolatilidades.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   25
      Top             =   7920
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7668
      Left            =   72
      TabIndex        =   0
      Top             =   96
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   13547
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "Volatilidades"
      TabPicture(0)   =   "frmVolatilidades.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Gráficos"
      TabPicture(1)   =   "frmVolatilidades.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Cálculo de Lambda Óptima"
      TabPicture(2)   =   "frmVolatilidades.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSFlexGrid1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmVolatilidades.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   2685
         Left            =   6750
         TabIndex        =   38
         Top             =   690
         Width           =   4605
         Begin VB.CommandButton Command1 
            Caption         =   "Minimizar Error Varianza"
            Height          =   675
            Left            =   288
            TabIndex        =   41
            Top             =   1620
            Width           =   1725
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   150
            TabIndex        =   40
            Top             =   570
            Width           =   2115
         End
         Begin VB.Label Label7 
            Caption         =   "Error de la Varianza"
            Height          =   165
            Left            =   210
            TabIndex        =   39
            Top             =   300
            Width           =   2205
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2805
         Left            =   -74580
         TabIndex        =   33
         Top             =   3450
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4948
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parametros"
         Height          =   2475
         Left            =   -74880
         TabIndex        =   29
         Top             =   390
         Width           =   5115
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   1980
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   180
            TabIndex        =   35
            Text            =   "99"
            Top             =   1290
            Width           =   2000
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   180
            TabIndex        =   30
            Top             =   600
            Width           =   2000
         End
         Begin VB.Label Label5 
            Caption         =   "Factor Riesgo 1"
            Height          =   225
            Left            =   270
            TabIndex        =   37
            Top             =   1710
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nivel de Confianza"
            Height          =   195
            Left            =   210
            TabIndex        =   32
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Días para el Error de la Varianza"
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   330
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Volatilidades"
         Height          =   3564
         Left            =   -74832
         TabIndex        =   23
         Top             =   408
         Width           =   11652
         Begin VB.PictureBox MSChart1 
            Height          =   3132
            Left            =   96
            ScaleHeight     =   3075
            ScaleWidth      =   11325
            TabIndex        =   24
            Top             =   300
            Width           =   11388
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Limites diarios"
         Height          =   3372
         Left            =   -74856
         TabIndex        =   21
         Top             =   4104
         Width           =   11652
         Begin VB.PictureBox MSChart4 
            Height          =   2904
            Left            =   84
            ScaleHeight     =   2850
            ScaleWidth      =   11370
            TabIndex        =   22
            Top             =   264
            Width           =   11436
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Características"
         Height          =   3108
         Left            =   100
         TabIndex        =   5
         Top             =   300
         Width           =   6504
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   2700
            TabIndex        =   42
            Top             =   2760
            Width           =   2500
         End
         Begin VB.Frame Frame2 
            Caption         =   "Rendimientos"
            Height          =   996
            Left            =   2370
            TabIndex        =   26
            Top             =   360
            Width           =   2000
            Begin VB.OptionButton Option2 
               Caption         =   "Logaritmicos"
               Height          =   192
               Left            =   100
               TabIndex        =   28
               Top             =   600
               Value           =   -1  'True
               Width           =   1236
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Aritmeticos"
               Height          =   192
               Left            =   100
               TabIndex        =   27
               Top             =   250
               Width           =   1236
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   100
            TabIndex        =   19
            Top             =   2760
            Width           =   2500
         End
         Begin VB.TextBox Text10 
            Height          =   288
            Left            =   100
            TabIndex        =   14
            Text            =   "365"
            Top             =   1635
            Width           =   2000
         End
         Begin VB.ComboBox txtNoVolatil 
            Height          =   315
            Left            =   100
            TabIndex        =   13
            Text            =   "30"
            Top             =   1035
            Width           =   2000
         End
         Begin VB.Frame Frame12 
            Caption         =   "Formula para el calculo"
            Height          =   1700
            Left            =   4410
            TabIndex        =   8
            Top             =   390
            Width           =   2000
            Begin VB.TextBox txtFactCaida 
               Height          =   288
               Left            =   90
               TabIndex        =   11
               Text            =   ".94"
               Top             =   1290
               Width           =   1500
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Normal"
               Height          =   192
               Left            =   100
               TabIndex        =   10
               Top             =   300
               Value           =   -1  'True
               Width           =   1265
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Ponderado exponencial"
               Height          =   385
               Left            =   100
               TabIndex        =   9
               Top             =   600
               Width           =   1177
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Factor de decaimiento"
               Height          =   176
               Left            =   100
               TabIndex        =   12
               Top             =   1100
               Width           =   1474
            End
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   100
            TabIndex        =   7
            Top             =   435
            Width           =   2000
         End
         Begin VB.TextBox Text7 
            Height          =   288
            Left            =   100
            TabIndex        =   6
            Text            =   "99"
            Top             =   2235
            Width           =   1991
         End
         Begin VB.Label Label6 
            Caption         =   "Factor Riesgo 2"
            Height          =   225
            Left            =   2670
            TabIndex        =   43
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Factor Riesgo 1:"
            Height          =   195
            Left            =   100
            TabIndex        =   20
            Top             =   2580
            Width           =   1170
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Días de calculo de eficiencia"
            Height          =   195
            Left            =   100
            TabIndex        =   18
            Top             =   1440
            Width           =   2070
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Última fecha:"
            Height          =   180
            Left            =   100
            TabIndex        =   17
            Top             =   225
            Width           =   855
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Días Calc. Vol."
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nivel de Confianza"
            Height          =   195
            Left            =   100
            TabIndex        =   15
            Top             =   2025
            Width           =   1350
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Eficiencia"
         Height          =   3100
         Left            =   5900
         TabIndex        =   3
         Top             =   3600
         Width           =   5700
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
            Height          =   2700
            Left            =   88
            TabIndex        =   4
            Top             =   250
            Width           =   5500
            _ExtentX        =   9710
            _ExtentY        =   4763
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Volatilidades"
         Height          =   3100
         Left            =   100
         TabIndex        =   1
         Top             =   3600
         Width           =   5700
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
            Height          =   2700
            Left            =   110
            TabIndex        =   2
            Top             =   250
            Width           =   5500
            _ExtentX        =   9710
            _ExtentY        =   4763
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Lambdas Optimas"
         Height          =   345
         Left            =   -74520
         TabIndex        =   34
         Top             =   3060
         Width           =   1815
      End
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mSalir 
         Caption         =   "Regresar al menu principal"
      End
   End
End
Attribute VB_Name = "frmVolatilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
Combo3.Text = Combo1.Text
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Command1_Click()
Dim ind1 As Integer
Dim ind2 As Integer
Dim fvol As Date
Dim ndias1 As Integer
Dim novolatil As Integer
Dim lambda As Double
Dim valor As Double
Dim dvalor As Double
Dim errorm As Double
Dim s As Double
Dim t As Double
Dim x As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
 
 ind1 = frmVolatilidades.Combo1.ListIndex + 2
 ind2 = frmVolatilidades.Combo3.ListIndex + 2
 fvol = CDate(frmVolatilidades.Combo4.Text)
 ndias1 = Int(Val(frmVolatilidades.Text10.Text))
 novolatil = Val(frmVolatilidades.txtNoVolatil.Text)
 lambda = Val(frmVolatilidades.txtFactCaida.Text)

valor = ErrVar(fvol, ndias1, novolatil, ind1, ind2, 1, lambda)
dvalor = DErrVar(fvol, ndias1, novolatil, ind1, ind2, 1, lambda)
errorm = 0.0000001
s = 0
Do
valor = ErrVar(fvol, ndias1, novolatil, ind1, ind2, 1, lambda)
dvalor = DErrVar(fvol, ndias1, novolatil, ind1, ind2, 1, lambda)
t = s
s = s - (valor - x) / dvalor
Loop Until Abs(t - s) < errorm
Call MostrarMensajeSistema(valor & " " & dvalor, frmProgreso.Label2, 1, Date, Time, NomUsuario)
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
For i = 2 To NoVectores
Combo1.AddItem MatVectores(i)
Combo3.AddItem MatVectores(i)
Next i
txtNoVolatil.AddItem 30
txtNoVolatil.AddItem 60
txtNoVolatil.AddItem 90
txtNoVolatil.AddItem 120
ListarFechas
Call ListarFRiesgo1
noreg = UBound(MatFactRiesgo, 1)
Combo4.Text = MatFactRiesgo(noreg, 1)
txtFactCaida.Text = 0.94
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
frmCalVar.Visible = True
frmCalVar.Enabled = True
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub mSalir_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Unload Me
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ListarFechas()
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
For i = 1 To NoDatVecMer
If Not IsNull(MatFactRiesgo(i, 1)) Then Combo4.AddItem MatFactRiesgo(i, 1)
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo4_Click()
Dim tvol As Integer
Dim indice1 As Integer
Dim indice2 As Integer
Dim ndias1 As Integer
Dim fvol As Date
Dim lambda As Double
Dim nconf As Double
Dim novolatil As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
 If frmVolatilidades.Option3 Then
 tvol = 0
 ElseIf frmVolatilidades.Option4 Then
 tvol = 1
 End If
 indice1 = frmVolatilidades.Combo1.ListIndex + 2
 indice2 = frmVolatilidades.Combo3.ListIndex + 2
 fvol = CDate(frmVolatilidades.Combo4.Text)
 lambda = Val(frmVolatilidades.txtFactCaida.Text)
 Call CalculaVolatilidad(fvol, ndias1, nconf, novolatil, indice1, indice2, 0, lambda)
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
Dim tvol As Integer
Dim indice1 As Integer
Dim indice2 As Integer
Dim fvol As Date
Dim ndias1 As Integer
Dim nconf As Double
Dim novolatil As Integer
Dim lambda As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If KeyAscii = 13 Then
Screen.MousePointer = 11
 If frmVolatilidades.Option3 Then
 tvol = 0
 ElseIf frmVolatilidades.Option4 Then
 tvol = 1
 End If
 indice1 = frmVolatilidades.Combo1.ListIndex + 2
 indice2 = frmVolatilidades.Combo3.ListIndex + 2
 fvol = CDate(frmVolatilidades.Combo4.Text)
 Call CalculaVolatilidad(fvol, ndias1, nconf, novolatil, indice1, indice2, 0, lambda)
Screen.MousePointer = 0
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
Dim tvol As Integer
Dim indice1 As Integer
Dim indice2 As Integer
Dim fvol As Date
Dim ndias1 As Integer
Dim nconf As Double
Dim novolatil As Integer
Dim lambda As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
If KeyAscii = 13 Then
 If frmVolatilidades.Option3 Then
 tvol = 0
 ElseIf frmVolatilidades.Option4 Then
 tvol = 1
 End If
 indice1 = frmVolatilidades.Combo1.ListIndex + 2
 indice2 = frmVolatilidades.Combo3.ListIndex + 2
 fvol = CDate(frmVolatilidades.Combo4.Text)
 ndias1 = Int(Val(frmVolatilidades.Text10.Text))
 nconf = Val(frmVolatilidades.Text7.Text) / 100
 novolatil = Val(frmVolatilidades.txtNoVolatil.Text)
 lambda = Val(frmVolatilidades.txtFactCaida.Text)
 frmProgreso.Show
 Call CalculaVolatilidad(fvol, ndias1, nconf, novolatil, indice1, indice2, 0, lambda)
 Unload frmProgreso
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Text7_KeyPress(KeyAscii As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
Call Text10_KeyPress(13)
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub txtNoVolatil_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
Call Text10_KeyPress(13)
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub txtNoVolatil_KeyPress(KeyAscii As Integer)
Dim tvol As Integer
Dim indice1 As Integer
Dim indice2 As Integer
Dim fvol As Date
Dim ndias1 As Integer
Dim nconf As Double
Dim novolatil As Integer
Dim lambda As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
If KeyAscii = 13 Then
 indice1 = frmVolatilidades.Combo1.ListIndex + 2
 indice2 = frmVolatilidades.Combo3.ListIndex + 2
 fvol = CDate(frmVolatilidades.Combo4.Text)
 ndias1 = Int(Val(frmVolatilidades.Text10.Text))
 nconf = Val(frmVolatilidades.Text7.Text) / 100
 novolatil = Val(frmVolatilidades.txtNoVolatil.Text)
 lambda = 0.94
 frmProgreso.Show
 Call CalculaVolatilidad(fvol, ndias1, nconf, novolatil, indice1, indice2, 0, lambda)
 Unload frmProgreso
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ListarFRiesgo1()
Dim i As Integer
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(MatCaracFRiesgo, 1)
For i = 1 To noreg
Combo1.AddItem MatCaracFRiesgo(i).indFactor
Combo3.AddItem MatCaracFRiesgo(i).indFactor
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

