VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVaRHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles del VaR Historico"
   ClientHeight    =   8790
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6516
      Left            =   192
      TabIndex        =   0
      Top             =   144
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   11483
      _Version        =   393216
      Tab             =   1
      TabHeight       =   420
      TabCaption(0)   =   "Rendimientos"
      TabPicture(0)   =   "frmVaRHistorico.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label50"
      Tab(0).Control(1)=   "MSFlexGrid26"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Histograma"
      TabPicture(1)   =   "frmVaRHistorico.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label29"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSFlexGrid14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid26 
         Height          =   5136
         Left            =   -74760
         TabIndex        =   1
         Top             =   720
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   9049
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid14 
         Height          =   5508
         Left            =   168
         TabIndex        =   3
         Top             =   672
         Width           =   4188
         _ExtentX        =   7382
         _ExtentY        =   9710
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Datos de Histograma"
         Height          =   192
         Left            =   96
         TabIndex        =   4
         Top             =   360
         Width           =   1536
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Rendimientos Observados"
         Height          =   192
         Left            =   -74712
         TabIndex        =   2
         Top             =   432
         Width           =   1932
      End
   End
End
Attribute VB_Name = "frmVaRHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

