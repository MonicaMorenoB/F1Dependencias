VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnalisisPlusMinus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lectura de plusvalias y minusvalias"
   ClientHeight    =   4725
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   630
      TabIndex        =   4
      Top             =   810
      Width           =   3555
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3435
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "frmFunciones.frx":0000
      Left            =   360
      List            =   "frmFunciones.frx":0002
      TabIndex        =   2
      Top             =   2340
      Width           =   6585
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Consolidar archivos"
      Height          =   645
      Left            =   8160
      TabIndex        =   1
      Top             =   1290
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar archivos"
      Height          =   495
      Left            =   8070
      TabIndex        =   0
      Top             =   510
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10230
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAnalisisPlusMinus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
CommonDialog1.ShowOpen
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim noreg As Integer
Dim txtdir As String
Dim lista As String
txtdir = Dir1.Path
lista = Dir(txtdir & "\*.*", vbArchive)
List1.Clear
Do While lista <> ""
   List1.AddItem lista
   lista = Dir
Loop
End Sub

Private Sub Command3_Click()
Dim txtnomarch() As String
Dim txtarch As String
Dim i As Long
Dim noreg As Integer
Dim txtdir As String

Screen.MousePointer = 11
noreg = List1.ListCount
txtdir = Dir1.Path
ReDim txtnomarch(1 To noreg)
For i = 1 To List1.ListCount
    txtnomarch(i) = txtdir & "\" & List1.List(i - 1)
Next i
txtarch = "Consolidacion plus-minus.txt"
CommonDialog1.FileName = txtarch
CommonDialog1.ShowOpen
txtarch = CommonDialog1.FileName
Open txtarch For Output As #2
Call LeerPlusMinuss(txtnomarch)
Close #2
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
