VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVaRMont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados VaR Montecarlo"
   ClientHeight    =   8070
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7380
      Left            =   72
      TabIndex        =   0
      Top             =   120
      Width           =   8748
      _ExtentX        =   15425
      _ExtentY        =   13018
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "Matriz de Choleski"
      TabPicture(0)   =   "frmVaRMont.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSFlexGrid12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Estadisticas de la simulaciones"
      TabPicture(1)   =   "frmVaRMont.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFlexGrid17"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Histograma simulaciones"
      TabPicture(2)   =   "frmVaRMont.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChart1"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox MSChart1 
         Height          =   6612
         Left            =   -74808
         ScaleHeight     =   6555
         ScaleWidth      =   8205
         TabIndex        =   7
         Top             =   456
         Width           =   8268
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar a archivo de texto"
         Height          =   660
         Left            =   4992
         TabIndex        =   6
         Top             =   1248
         Width           =   2436
      End
      Begin VB.TextBox Text1 
         Height          =   288
         Left            =   4968
         TabIndex        =   4
         Top             =   816
         Width           =   2460
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid12 
         Height          =   6120
         Left            =   216
         TabIndex        =   1
         Top             =   696
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   10795
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid17 
         Height          =   6720
         Left            =   -74784
         TabIndex        =   3
         Top             =   408
         Width           =   8388
         _ExtentX        =   14817
         _ExtentY        =   11853
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Distancia matricial entre Cov y MatM"
         Height          =   312
         Left            =   4968
         TabIndex        =   5
         Top             =   432
         Width           =   3132
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Matriz para descomposición"
         Height          =   192
         Left            =   264
         TabIndex        =   2
         Top             =   408
         Width           =   2028
      End
   End
End
Attribute VB_Name = "frmVaRMont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

