VERSION 5.00
Begin VB.Form frmEjecSubproc1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecucion de subprocesos 1"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   420
      Top             =   120
   End
End
Attribute VB_Name = "frmEjecSubproc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
   MensajeProc = NomUsuario & " ha salido del sistema"
   Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
   RGuardarPL.Close
   RegResCVA.Close
   conAdo.Close
   conAdoBD.Close
   End
End Sub

Private Sub Timer1_Timer()
    Screen.MousePointer = 11
    frmProgreso.Show
    Call EjecucionSubprocesos(1)
    Unload frmProgreso
    Screen.MousePointer = 0
End Sub
