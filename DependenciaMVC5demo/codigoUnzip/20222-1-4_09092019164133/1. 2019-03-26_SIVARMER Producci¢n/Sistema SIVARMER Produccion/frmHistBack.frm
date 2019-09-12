VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHistBack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historia del backtesting"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   10485
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   12615
      Begin VB.Frame Frame14 
         Caption         =   "Parametros"
         Height          =   5448
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   2316
         Begin VB.CommandButton Command13 
            Caption         =   "Mostrar Resumen"
            Height          =   600
            Left            =   216
            TabIndex        =   7
            Top             =   1584
            Width           =   1500
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Graficar Resultados"
            Height          =   600
            Left            =   216
            TabIndex        =   6
            Top             =   2280
            Width           =   1500
         End
         Begin VB.ComboBox Combo12 
            Height          =   315
            Left            =   144
            TabIndex        =   5
            Top             =   984
            Width           =   1956
         End
         Begin VB.ComboBox Combo11 
            Height          =   315
            Left            =   144
            TabIndex        =   4
            Top             =   432
            Width           =   1932
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Crear HTML"
            Height          =   600
            Left            =   264
            TabIndex        =   3
            Top             =   3072
            Width           =   1500
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Crear PDF"
            Height          =   600
            Left            =   240
            TabIndex        =   2
            Top             =   3936
            Width           =   1500
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "al día"
            Height          =   192
            Left            =   96
            TabIndex        =   9
            Top             =   720
            Width           =   396
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Del día"
            Height          =   192
            Left            =   96
            TabIndex        =   8
            Top             =   204
            Width           =   516
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid19 
         Height          =   9825
         Left            =   2565
         TabIndex        =   10
         Top             =   300
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   17330
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmHistBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim noreg As Integer
Dim i As Integer

noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo11.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo12.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
