VERSION 5.00
Begin VB.Form frmEjecSubproc3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecución de subprocesos de CVaR Intradia"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4620
      Top             =   90
   End
End
Attribute VB_Name = "frmEjecSubproc3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Screen.MousePointer = 11
    frmProgreso.Show
    Call EjecucionSubprocesos(3)
    Unload frmProgreso
    Screen.MousePointer = 0
End Sub
