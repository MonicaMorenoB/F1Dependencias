VERSION 5.00
Begin VB.Form frmEjecSubproc2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecución de subprocesos 2"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   390
      Top             =   300
   End
End
Attribute VB_Name = "frmEjecSubproc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
   MensajeProc = NomUsuario & " ha salido del sistema"
   Call GuardaDatosBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc, 2)
   RGuardarPL.Close
   RegResCVA.Close
   ConAdo.Close
   conAdoBD.Close
   End

End Sub

Private Sub Timer1_Timer()
    Screen.MousePointer = 11
    frmProgreso.Show
    Call EjecucionSubprocesos(2)
    Unload frmProgreso
    Screen.MousePointer = 0
End Sub
