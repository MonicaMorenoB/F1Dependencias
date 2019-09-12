VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCatalogos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogos del Sistema"
   ClientHeight    =   7290
   ClientLeft      =   -30
   ClientTop       =   330
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9540
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Catalogos"
      Height          =   6795
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   9945
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   255
         Left            =   9000
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Marco de operacion"
         Height          =   700
         Left            =   2610
         TabIndex        =   11
         Top             =   4020
         Width           =   1905
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Portafolios de posición"
         Height          =   700
         Left            =   2500
         TabIndex        =   10
         Top             =   3100
         Width           =   2000
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Parametros sistema"
         Height          =   700
         Left            =   2500
         TabIndex        =   9
         Top             =   2200
         Width           =   2000
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Factores de riesgo"
         Height          =   700
         Left            =   2500
         TabIndex        =   8
         Top             =   1300
         Width           =   2000
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Derivados"
         Height          =   700
         Left            =   300
         TabIndex        =   7
         Top             =   4000
         Width           =   2000
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Contrapartes"
         Height          =   700
         Left            =   300
         TabIndex        =   6
         Top             =   3100
         Width           =   2000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Procesos"
         Height          =   700
         Left            =   300
         TabIndex        =   5
         Top             =   2200
         Width           =   2000
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   660
         Width           =   7965
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Valuación"
         Height          =   700
         Left            =   300
         TabIndex        =   3
         Top             =   1290
         Width           =   2000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Todos los catalogos"
         Height          =   700
         Left            =   300
         TabIndex        =   1
         Top             =   5130
         Width           =   2000
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   705
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dirbase As String
Dim nomarch As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Screen.MousePointer = 11
 MsgBox "Se sincronizan los catalogos de Access y Oracle"
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call SincCatAccessOracle(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No hay acceso al archivo " & nomarch
 End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Private Sub Command10_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatMO(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatValuacion(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 dirbase = Text1.Text
 nomarch = dirbase & "\" & TablaCatalogosA
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatProcesos(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatContrapartes(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command5_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatDerivados(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command6_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatFactRiesgo(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatParametros(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command8_Click()
Dim dirbase As String
Dim nomarch As String

Screen.MousePointer = 11
 nomarch = Text1.Text
 If VerifAccesoArch(nomarch) Then
    frmProgreso.Show
    SiActTProc = True
    Call CatEstReportes(nomarch)
    Call LeerCatalogos
    Unload frmProgreso
    MsgBox "Proceso terminado"
    Call ActUHoraUsuario
    SiActTProc = False
 Else
    MsgBox "No se encuentra el archivo " & nomarch
 End If
Screen.MousePointer = 0
End Sub

Private Sub Command9_Click()
frmCatalogos.CommonDialog1.FileName = Text1.Text
frmCatalogos.CommonDialog1.ShowOpen
Text1.Text = frmCatalogos.CommonDialog1.FileName
End Sub

Private Sub Form_Load()
   Text1.Text = DirBases & "\VAR_CATALOGOSN.MDB"
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

