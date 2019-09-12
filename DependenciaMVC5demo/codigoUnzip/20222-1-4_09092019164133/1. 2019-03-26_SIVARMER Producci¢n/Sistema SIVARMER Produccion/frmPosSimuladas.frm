VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPosSimuladas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación de posiciones simuladas"
   ClientHeight    =   5010
   ClientLeft      =   -30
   ClientTop       =   360
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Caption         =   "Posiciones simuladas"
      Height          =   4680
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10830
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9030
         Top             =   540
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   6660
         TabIndex        =   19
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   1600
         TabIndex        =   18
         Text            =   "H:\Riesgos de Mercado\Posiciones\Primarias\"
         Top             =   2130
         Width           =   4995
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1170
         TabIndex        =   16
         Top             =   3540
         Width           =   2715
      End
      Begin VB.TextBox Text12 
         Height          =   288
         Left            =   1600
         TabIndex        =   10
         Text            =   "H:\Riesgos de Mercado\Posiciones\Mesa dinero\"
         Top             =   330
         Width           =   4980
      End
      Begin VB.TextBox Text13 
         Height          =   300
         Left            =   1600
         TabIndex        =   9
         Text            =   "H:\Riesgos de Mercado\Posiciones\Divisas\"
         Top             =   800
         Width           =   5010
      End
      Begin VB.TextBox Text14 
         Height          =   288
         Left            =   1600
         TabIndex        =   8
         Text            =   "H:\Riesgos de Mercado\Posiciones\Derivados\"
         Top             =   1250
         Width           =   4995
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Importar datos"
         Height          =   444
         Left            =   9150
         TabIndex        =   7
         Top             =   4110
         Width           =   1524
      End
      Begin VB.TextBox Text15 
         Height          =   288
         Left            =   1620
         TabIndex        =   6
         Top             =   4050
         Width           =   5040
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   6660
         TabIndex        =   5
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   6630
         TabIndex        =   4
         Top             =   870
         Width           =   1065
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   6660
         TabIndex        =   3
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox Text18 
         Height          =   288
         Left            =   1600
         TabIndex        =   2
         Text            =   "H:\Riesgos de Mercado\Posiciones\Derivados\"
         Top             =   1700
         Width           =   4995
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Examinar"
         Height          =   300
         Left            =   6660
         TabIndex        =   1
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   3540
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Posicion primaria"
         Height          =   195
         Left            =   200
         TabIndex        =   17
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Mesa de Dinero"
         Height          =   192
         Left            =   200
         TabIndex        =   15
         Top             =   350
         Width           =   1152
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Mesa de Cambios"
         Height          =   192
         Left            =   200
         TabIndex        =   14
         Top             =   850
         Width           =   1320
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Swaps"
         Height          =   192
         Left            =   200
         TabIndex        =   13
         Top             =   1300
         Width           =   492
      End
      Begin VB.Label Label14 
         Caption         =   "Nombre de la Simulación"
         Height          =   510
         Left            =   270
         TabIndex        =   12
         Top             =   4020
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Forwards"
         Height          =   192
         Left            =   200
         TabIndex        =   11
         Top             =   1750
         Width           =   672
      End
   End
End
Attribute VB_Name = "frmPosSimuladas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nomarch As String
  CommonDialog1.FileName = nomarch
  CommonDialog1.ShowOpen
  nomarch = CommonDialog1.FileName
  Text2.Text = nomarch

End Sub

Private Sub Command37_Click()
Dim fecha As Date
Dim noreg As Integer
Dim nomposicion As String
Dim exito As Boolean
Dim nr1 As Long
Dim nr2 As Long
Dim nr3 As Long
Dim nr4 As Long
Dim nr5 As Long
Dim txtmsg As String
Dim nomarchmd As String
Dim nomarchmc As String
Dim nomarchswaps As String
Dim nomarchfwd As String
Dim nomarchposprim As String
Dim txtclave As String
Dim intencion As String
Dim sihayarch1 As Boolean
Dim sihayarch2 As Boolean
Dim sihayarch3 As Boolean
Dim sihayarch4 As Boolean
Dim sihayarch5 As Boolean


Screen.MousePointer = 11
fecha = CDate(Text1.Text)
nomarchmd = Trim(Text12.Text)
nomarchmc = Trim(Text13.Text)
nomarchswaps = Trim(Text14.Text)
nomarchfwd = Trim(Text18.Text)
nomarchposprim = Trim(Text2.Text)
nomposicion = Trim(Text15.Text)
If Len(Trim(Text15.Text)) <> 0 Then
If Len(Trim(Text1.Text)) <> 0 Then
fecha = CDate(Trim(Text1.Text))
  'acceso a las operaciones via servidor oracle
 sihayarch1 = VerifAccesoArch(nomarchmd)
 If sihayarch1 = True And Right(nomarchmd, 1) <> "\" Then
    Call ImpPosMDineroSim(fecha, nomarchmd, nomposicion, nr1, exito)
 End If
 sihayarch2 = VerifAccesoArch(nomarchmc)
 If sihayarch2 = True And Right(nomarchmc, 1) <> "\" Then
    Call ImpPosMCambiosSim(fecha, nomarchmc, nomposicion, nr2, exito)
 End If
  sihayarch3 = VerifAccesoArch(nomarchfwd)
  If sihayarch3 = True And Right(nomarchfwd, 1) <> "\" Then
     Call ImpFwdSimArch(fecha, nomarchfwd, nomposicion, nr3, exito)
  End If
  sihayarch4 = VerifAccesoArch(nomarchswaps)
  If sihayarch4 = True And Right(nomarchswaps, 1) <> "\" Then
     Call CrearPosSwapsSimArch(fecha, nomarchswaps, 2, nomposicion, "000000", nr4)
  End If
  sihayarch5 = VerifAccesoArch(nomarchposprim)
  If sihayarch5 = True And Right(nomarchposprim, 1) <> "\" Then
     Call ImpPosPrimArc(fecha, nomposicion, nomarchposprim, 2, nr5, txtmsg, exito)
  End If
Else
  MsgBox "xx"
End If
Else
MsgBox "No se definio un nombre para la posición de simulacion"
End If
MsgBox "Proceso terminado"
Screen.MousePointer = 0

End Sub

Private Sub Command38_Click()
Dim nomarch As String

  CommonDialog1.FileName = nomarch
  CommonDialog1.ShowOpen
  nomarch = CommonDialog1.FileName
  Text18.Text = nomarch
End Sub

Private Sub Command40_Click()
Dim nomarch As String

 CommonDialog1.FileName = nomarch
 CommonDialog1.ShowOpen
 nomarch = CommonDialog1.FileName
 Text12.Text = nomarch

End Sub

Private Sub Command41_Click()
Dim nomarch As String
 CommonDialog1.FileName = nomarch
 CommonDialog1.ShowOpen
 nomarch = CommonDialog1.FileName
 Text13.Text = nomarch

End Sub

Private Sub Command42_Click()
Dim nomarch As String
 CommonDialog1.FileName = nomarch
 CommonDialog1.ShowOpen
 nomarch = CommonDialog1.FileName
 Text14.Text = nomarch

End Sub

Private Sub Form_Load()
Text1.Text = Date
End Sub

