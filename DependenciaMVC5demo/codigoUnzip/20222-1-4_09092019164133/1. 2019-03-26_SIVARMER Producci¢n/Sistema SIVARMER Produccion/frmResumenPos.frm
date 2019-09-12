VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtrapolYieldTIIE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extrapolacion de yields referenciadas a TIIE"
   ClientHeight    =   3210
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   7035
   Icon            =   "frmResumenPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3990
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar historia"
      Height          =   825
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   345
      Left            =   6180
      TabIndex        =   4
      Top             =   1590
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   300
      TabIndex        =   3
      Top             =   1590
      Width           =   5745
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   300
      TabIndex        =   2
      Top             =   1020
      Width           =   2200
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   2200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   780
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ubicacion del archivo"
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   1380
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de inicio"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmExtrapolYieldTIIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtnomarch As String
Dim base1      As DAO.Database
Dim i          As Integer, noreg As Integer
Dim j As Integer
Dim registros1 As DAO.recordset

Screen.MousePointer = 11
fecha1 = CDate(Combo1.Text)
fecha2 = CDate(Combo2.Text)
SiActTProc = True
txtnomarch = Text1.Text
If VerifAccesoArch(txtnomarch) Then
 frmProgreso.Show
    Set base1 = OpenDatabase(txtnomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
    
    If registros1.RecordCount <> 0 Then
        registros1.MoveLast
        noreg = registros1.RecordCount
        ReDim mata(1 To noreg, 1 To 6) As Variant
        registros1.MoveFirst
        For i = 1 To noreg
          For j = 1 To 6
            mata(i, j) = LeerTAccess(registros1, j - 1, i)
          Next j
            registros1.MoveNext
        Next i

        registros1.Close
        base1.Close
        Call ExtrapolTTIIE(fecha1, fecha2, mata)
    End If
 End If
 Unload frmProgreso
 Call ActUHoraUsuario
 SiActTProc = False
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim noreg As Long
noreg = UBound(MatFechasVaR, 1)
Combo1.Clear
Combo2.Clear
For i = 1 To noreg
 Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
 Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub


