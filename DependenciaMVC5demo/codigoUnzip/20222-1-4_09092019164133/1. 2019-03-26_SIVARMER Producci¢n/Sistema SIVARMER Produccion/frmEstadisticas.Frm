VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de ánalisis de instrumentos financieros"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   12195
   Icon            =   "frmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8475
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7572
      Left            =   72
      TabIndex        =   1
      Top             =   108
      Width           =   11952
      _ExtentX        =   21087
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   423
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Análisis estadistico"
      TabPicture(0)   =   "frmEstadisticas.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Gráficos"
      TabPicture(1)   =   "frmEstadisticas.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Instrumento a analizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3322
         Left            =   -74928
         TabIndex        =   32
         Top             =   350
         Width           =   7194
         Begin VB.ComboBox Combo3 
            Height          =   264
            Left            =   132
            TabIndex        =   43
            Top             =   1518
            Width           =   1969
         End
         Begin VB.ComboBox Combo1 
            Height          =   264
            Left            =   100
            TabIndex        =   37
            Top             =   462
            Width           =   4180
         End
         Begin VB.ComboBox Combo2 
            Height          =   264
            Left            =   100
            TabIndex        =   36
            Top             =   950
            Width           =   2000
         End
         Begin VB.Frame Frame1 
            Caption         =   "Rendimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1152
            Left            =   100
            TabIndex        =   33
            Top             =   1900
            Width           =   2052
            Begin VB.OptionButton Option2 
               Caption         =   "Logaritmicos"
               Height          =   288
               Left            =   100
               TabIndex        =   35
               Top             =   648
               Value           =   -1  'True
               Width           =   1380
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Aritmeticos"
               Height          =   288
               Left            =   100
               TabIndex        =   34
               Top             =   250
               Width           =   1188
            End
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2354
            Left            =   2222
            TabIndex        =   38
            Top             =   836
            Width           =   4917
            _ExtentX        =   8678
            _ExtentY        =   4154
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador"
            Height          =   192
            Left            =   100
            TabIndex        =   41
            Top             =   250
            Width           =   672
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de días"
            Height          =   192
            Left            =   100
            TabIndex        =   40
            Top             =   750
            Width           =   792
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta el día:"
            Height          =   176
            Left            =   99
            TabIndex        =   39
            Top             =   1254
            Width           =   825
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prueba Kolmogorov-Smirnov"
         Height          =   3100
         Left            =   -69566
         TabIndex        =   28
         Top             =   3784
         Width           =   5200
         Begin VB.TextBox Text8 
            Height          =   288
            Left            =   1600
            TabIndex        =   29
            Top             =   240
            Width           =   1800
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   2300
            Left            =   88
            TabIndex        =   30
            Top             =   638
            Width           =   5005
            _ExtentX        =   8837
            _ExtentY        =   4048
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor del estadístico:"
            Height          =   192
            Left            =   100
            TabIndex        =   31
            Top             =   300
            Width           =   1500
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estadistica descriptiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   -67652
         TabIndex        =   13
         Top             =   418
         Width           =   3300
         Begin VB.TextBox Text1 
            Height          =   288
            Left            =   1900
            TabIndex        =   20
            Top             =   300
            Width           =   1300
         End
         Begin VB.TextBox Text2 
            Height          =   288
            Left            =   1900
            TabIndex        =   19
            Top             =   700
            Width           =   1300
         End
         Begin VB.TextBox Text3 
            Height          =   288
            Left            =   1900
            TabIndex        =   18
            Top             =   1100
            Width           =   1300
         End
         Begin VB.TextBox Text4 
            Height          =   288
            Left            =   1900
            TabIndex        =   17
            Top             =   1500
            Width           =   1300
         End
         Begin VB.TextBox Text5 
            Height          =   288
            Left            =   1900
            TabIndex        =   16
            Top             =   1900
            Width           =   1300
         End
         Begin VB.TextBox Text6 
            Height          =   288
            Left            =   1900
            TabIndex        =   15
            Top             =   2300
            Width           =   1300
         End
         Begin VB.TextBox Text12 
            Height          =   288
            Left            =   1900
            TabIndex        =   14
            Top             =   2664
            Width           =   1300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Media"
            Height          =   192
            Left            =   100
            TabIndex        =   27
            Top             =   300
            Width           =   456
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desviación estandar"
            Height          =   192
            Left            =   100
            TabIndex        =   26
            Top             =   700
            Width           =   1488
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Coeficiente de asimetría"
            Height          =   192
            Left            =   100
            TabIndex        =   25
            Top             =   1100
            Width           =   1728
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Curtosis"
            Height          =   192
            Left            =   100
            TabIndex        =   24
            Top             =   1500
            Width           =   576
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimo"
            Height          =   192
            Left            =   100
            TabIndex        =   23
            Top             =   1900
            Width           =   516
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Máximo"
            Height          =   192
            Left            =   100
            TabIndex        =   22
            Top             =   2300
            Width           =   564
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño de la muestra"
            Height          =   192
            Left            =   120
            TabIndex        =   21
            Top             =   2712
            Width           =   1620
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Histograma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3100
         Left            =   -74890
         TabIndex        =   8
         Top             =   3800
         Width           =   5200
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   2300
            Left            =   110
            TabIndex        =   42
            Top             =   660
            Width           =   5000
            _ExtentX        =   8837
            _ExtentY        =   4048
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
         End
         Begin VB.TextBox txtFactEsc 
            Height          =   288
            Left            =   1500
            TabIndex        =   9
            Top             =   250
            Width           =   418
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de invervalos:"
            Height          =   187
            Left            =   2255
            TabIndex        =   12
            Top             =   297
            Width           =   1265
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Factor de Escala"
            Height          =   187
            Left            =   110
            TabIndex        =   11
            Top             =   363
            Width           =   1221
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   176
            Left            =   3641
            TabIndex        =   10
            Top             =   319
            Width           =   33
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Rendimientos"
         Height          =   3300
         Left            =   100
         TabIndex        =   6
         Top             =   350
         Width           =   5300
         Begin VB.PictureBox MSChart1 
            Height          =   3000
            Left            =   100
            ScaleHeight     =   2940
            ScaleWidth      =   5040
            TabIndex        =   7
            Top             =   240
            Width           =   5100
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Histograma"
         Height          =   3300
         Left            =   5522
         TabIndex        =   4
         Top             =   352
         Width           =   5300
         Begin VB.PictureBox MSChart3 
            Height          =   3003
            Left            =   110
            ScaleHeight     =   2940
            ScaleWidth      =   5040
            TabIndex        =   5
            Top             =   242
            Width           =   5104
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Distribucion Normal Teorica vs. Empírica"
         Height          =   3300
         Left            =   110
         TabIndex        =   2
         Top             =   3718
         Width           =   5300
         Begin VB.PictureBox MSChart2 
            Height          =   3000
            Left            =   100
            ScaleHeight     =   2940
            ScaleWidth      =   5040
            TabIndex        =   3
            Top             =   240
            Width           =   5100
         End
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   528
      Left            =   0
      TabIndex        =   0
      Top             =   7956
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   926
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "11/01/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "11/01/2018"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   16325
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstadisticas.frx":047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstadisticas.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstadisticas.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmEstadisticas.frx":0DC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mActualizar 
         Caption         =   "Validación de datos"
      End
      Begin VB.Menu mImprime 
         Caption         =   "Imprimir Análisis"
      End
      Begin VB.Menu mSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mAcerca 
         Caption         =   "Acerca del programa"
      End
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub AnalisisEstadistico(ByRef mata() As Double)
Dim tabla() As Variant

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'mata  es la matriz de rendimientos
Call CalculaEstadisticas(mata)
MatHistograma = GenHistograma(mata, FEscala)
NoIntervalos = UBound(MatHistograma, 1)
Call AnalisisKolmogorov(mata)
Call LlenarTablaRends(MSFlexGrid1, tabla, mata)
VerDatosEst
Call ActGraficos(mata)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub GraficarRendimientos(ByRef mata() As Double)
Dim n As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
n = UBound(mata, 1)
grafico.chartType = VtChChartType2dLine
grafico.ColumnCount = 1
grafico.RowCount = n
For i = 1 To n
grafico.Column = 1
grafico.Row = i
grafico.Data = mata(i, 1)
grafico.RowLabel = mata(i, 1)
Next i
grafico.Title = Left("Rendimientos diarios - " & Combo1.Text, 80)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ListarFRiesgo()
Dim noreg As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(MatNodosFRiesgo, 1)
For i = 1 To noreg
Combo1.AddItem MatNodosFRiesgo(i).indFactor
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub LlenarTablaRends(ByRef rejilla1 As MSFlexGrid, ByRef tabla() As Variant, ByRef mata() As Double)
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se llena la tabla con los valores y los
'rendimientos de la serie
rejilla1.Rows = NoDias + 2
rejilla1.Cols = 3
rejilla1.ColWidth(0) = 1000
rejilla1.ColWidth(1) = 1900
rejilla1.ColWidth(2) = 1900
'rejilla1.TextMatrix(0, 1) = MatVectores(indice + 1)
rejilla1.TextMatrix(0, 2) = "Rendimientos"
For i = 1 To NoDias
For j = 1 To 2
If Not EsVariableVacia(tabla(i, j)) Then
   rejilla1.TextMatrix(i, j - 1) = tabla(i, j)
End If
If i <> 1 Then rejilla1.TextMatrix(i, 2) = mata(i - 1, 1)
Next j
Next i
rejilla1.RowHeight(0) = 800
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ActGraficos(ByRef mata() As Double)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Set grafico = MSChart1
Call GraficarRendimientos(mata)
Set grafico = MSChart2
Call GraficarHistograma(MatHistograma, grafico, nombre)
Set grafico = MSChart3
Call GraficoKolmogorov(mata, grafico, 1)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub VerDatosEst()
Dim i As Integer
Dim contador As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Text1.Text = Format(media, "##0.0000%")
Text2.Text = Format(DesvEst, "##0.0000%")
Text3.Text = Format(CoefAsimetria, "##0.0000")
Text4.Text = Format(Curtosis, "##0.0000")
Text5.Text = Format(ValorMinimo, "##0.0000%")
Text6.Text = Format(ValorMaximo, "##0.0000%")
Text12.Text = NoDias
Label18.Caption = NoIntervalos
Text8.Text = Kolmogorov
MSFlexGrid2.Rows = 1
MSFlexGrid2.Rows = 2
MSFlexGrid2.FixedRows = 1
MSFlexGrid2.Rows = NoIntervalos + 2
MSFlexGrid2.Cols = 5
'se colocan los datos del histograma
Call TablaHistograma(MatHistograma, MSFlexGrid2)
With MSFlexGrid3
.Rows = 1
.Rows = 2
.FixedRows = 1
.Rows = contador + 1
.Cols = 3
.Col = 1
For i = 1 To contador
.Row = i
.Col = 0
.Text = Format(matorden(i, 1), "####0.#####")
.Col = 1
.Text = Format(matorden(i, 2), "####0.#####")
.Col = 2
.Text = Format(matorden(i, 3), "####0.#####")
Next i
End With
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub AnalisisKolmogorov(ByRef mata() As Double)
Dim n As Long
Dim i As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
' es en esta rutina donde se calcula el estadistico de kolmogorov
'asi como las distribuciones teoricas y empiricas de la muestra
matorden = GenerarEmpirica(mata, 0, 0.95)
n = UBound(matorden, 1)
For i = 1 To n
matorden(i, 3) = DNormal(matorden(i, 1), media, DesvEst, 1)
Call MostrarMensajeSistema(media & " " & DesvEst & " " & matorden(i, 1) & " " & matorden(i, 3), frmProgreso.Label2, 0#, Date, Time, NomUsuario)
Next i
'se procede a calcular el valor del Estadistico de Kolmogorov-Smirnov
Kolmogorov = 0
For i = 1 To n
Kolmogorov = Maximo(Kolmogorov, Abs(matorden(i, 2) - matorden(i, 3)))
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Sub TitulosKolmogorov()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
MSFlexGrid3.Cols = 3
MSFlexGrid3.Row = 0
MSFlexGrid3.Col = 0
MSFlexGrid3.Text = "Valor X"
MSFlexGrid3.Col = 1
MSFlexGrid3.Text = "Distribucion empírica"
MSFlexGrid3.Col = 2
MSFlexGrid3.Text = "Distribucion teórica"
MSFlexGrid3.RowHeight(0) = 600
MSFlexGrid3.ColWidth(0) = 1000
MSFlexGrid3.ColWidth(1) = 1300
MSFlexGrid3.ColWidth(2) = 1300
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TrazarNormal(ByVal lix As Long, ByVal liy As Long, ByVal lsx As Long, ByVal lsy As Long, objeto1)
Dim limitel As Long
Dim limiter As Long
Dim limiteu  As Long
Dim limited As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'con este programa trazamos la curva normal
objeto1.DrawWidth = 2
limitel = 50
limiter = objeto1.ScaleWidth - 50
limiteu = 150
limited = objeto1.ScaleHeight - 50
objeto1.Line (lix, liy)-(lsx, lsy), , B

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo1_Click()

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
CalculaEstadisticasMuestra
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub CalculaEstadisticas(ByRef mata() As Double)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
' en esta rutina se calculan las estadisticas basicas del instrumento a analizar
'se calcula el maximo, ValorMinimo y media de la serie


ValorMaximo = VMaximo(mata, 1, 1)
ValorMinimo = VMinimo(mata, 1, 1)
media = CMedia2(mata, 1, "c")
'la desviacion estandar
DesvEst = Sqr(CVarianza2(mata, 1, "c"))
'se calcula la curtosis de la serie
Curtosis = CCurtosis2(mata, 1, "c")
'el coeficiente de asimetria
CoefAsimetria = CCAsimetria2(mata, 1, "c")
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
If KeyAscii = 13 Then
CalculaEstadisticasMuestra
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo2_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
CalculaEstadisticasMuestra
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If KeyAscii = 13 Then
Combo2_Click
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Combo3_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Combo2_Click
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If KeyAscii = 13 Then Combo2_Click
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub



Private Sub Form_Load()
Dim noreg As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se procede a cargar y formatear
'la hoja de análisis
frmEstadisticas.Caption = NomSistem & " Analisis estadistico de series de datos"
ListarFRiesgo
Combo2.AddItem 30
Combo2.AddItem 60
Combo2.AddItem 90
Combo2.AddItem 120
SiCargoSerie = False
ListarFechasVal
noreg = UBound(MatFactRiesgo, 1)
Combo3.Text = MatFactRiesgo(noreg, 1)
txtFactEsc.Text = 1
Call TitulosHistograma(frmEstadisticas.MSFlexGrid2)
TitulosKolmogorov
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
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub



Private Sub mAcerca_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
frmAcerca.Show 1
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

Private Sub Option1_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
CalculaEstadisticasMuestra
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Private Sub Option2_Click()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
CalculaEstadisticasMuestra
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub CalculaEstadisticasMuestra()
Dim nombre As String
Dim matrends() As Double
Dim indice As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim tabla() As Double
'SE PROCEDE A VERIFICAR LOS DATOS DE ENTRADA
' SI ESTAN BIEN, SE PROCEDE A CARGAR LA SERIE
'EN CUESTION
VerificaParametros
If Procede Then
nombre = Combo1.Text
NoDias = Val(Combo2.Text)
Dias = CDbl(NoDias)
FechaFinal = CDate(Combo3.Text)
FEscala = Val(txtFactEsc.Text)
If indice <= 0 Then indice = Combo1.ListIndex + 2
frmProgreso.Show
IFecha = BuscarValorVector(FechaFinal, MatFactRiesgo, 1)
If IFecha = 0 Then
 Call MostrarMensajeSistema("Falta la fecha en la tabla de datos, se hara con la ultima fecha de la tabla", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 IFecha = NoDatVecMer
End If
Combo3.Text = MatFactRiesgo(IFecha, 1)
If IFecha < NoDias + 1 Then
 Call MostrarMensajeSistema("No ha suficientes datos para realizar los calculos", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Procede2 = False
 GoTo inicio
End If
'se lee la serie con los datos a analizar
tabla = ExtSerieFR(MatFactRiesgo, IFecha, indice, NoDias)
Procede2 = True
Unload frmProgreso
inicio:
If Procede2 Then
'se calcula la matriz de rendimientos
If Option2.value Then
matrends = CalculaRendimientoColumna(tabla, 2)
ElseIf Option1.value Then
matrends = CalculaRendimientoColumna(tabla, 2)
End If
Call AnalisisEstadistico(matrends)
frmEstadisticas.SetFocus
MSChart1.Enabled = True
MSChart2.Enabled = True
MSChart3.Enabled = True
End If
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
Sub VerificaParametros()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
If Combo1.ListIndex < 0 Then
 Call MostrarMensajeSistema("no se ha seleccionado un indicador", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Procede = False
 Exit Sub
End If

If Val(Combo2.Text) = 0 Then
 Call MostrarMensajeSistema("No de días invalido", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Procede = False
 Exit Sub
End If

If Not IsDate(Combo3.Text) And Val(Combo3.Text) = 0 Then
 Call MostrarMensajeSistema("La fecha de corte no es válida", frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Procede = False
 Exit Sub
End If
Procede = True
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ListarFechasVal()
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
For i = 1 To NoDatVecMer
If Not IsNull(MatFactRiesgo(i, 1)) Then
Combo3.AddItem MatFactRiesgo(i, 1)
End If
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
