VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResumenPos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalles de la Posicion"
   ClientHeight    =   8856
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   10752
   Icon            =   "frmDetallesPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8856
   ScaleWidth      =   10752
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Resumen de la posición"
      Height          =   8412
      Left            =   120
      TabIndex        =   0
      Top             =   168
      Width           =   10428
      Begin VB.CommandButton Command16 
         Caption         =   "Crear PDF VaR Mesa Dinero"
         Height          =   500
         Left            =   3288
         TabIndex        =   10
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Crear HTLM"
         Height          =   500
         Left            =   6576
         TabIndex        =   9
         Top             =   264
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exportar Archivo Texto"
         Height          =   500
         Left            =   168
         TabIndex        =   8
         Top             =   240
         Width           =   1236
      End
      Begin VB.Frame Frame26 
         Caption         =   "Tipo de Volatilidad"
         Height          =   672
         Left            =   192
         TabIndex        =   4
         Top             =   7488
         Width           =   4596
         Begin VB.OptionButton Option36 
            Caption         =   "Constante"
            Height          =   192
            Left            =   100
            TabIndex        =   7
            Top             =   300
            Width           =   1236
         End
         Begin VB.OptionButton Option37 
            Caption         =   "Exponencial"
            Height          =   192
            Left            =   1368
            TabIndex        =   6
            Top             =   300
            Width           =   1188
         End
         Begin VB.OptionButton Option38 
            Caption         =   "Max ambos"
            Height          =   192
            Left            =   2784
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1452
         End
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Guardar datos VaR"
         Height          =   516
         Left            =   1512
         TabIndex        =   3
         Top             =   240
         Width           =   1668
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Crear PDF VaR Fondo Pensiones"
         Height          =   468
         Left            =   4872
         TabIndex        =   2
         Top             =   264
         Width           =   1452
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid18 
         Height          =   6432
         Left            =   192
         TabIndex        =   1
         Top             =   936
         Width           =   10104
         _ExtentX        =   17822
         _ExtentY        =   11345
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
      End
   End
End
Attribute VB_Name = "frmResumenPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command16_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 Call ImpresionVaRMesa
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Command29_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
If NivelUsuario = "ADMINISTRADOR" Then
 If TBaseDatos = 0 Then
 Call GuardarResVaRO(frmCalVar.StatusBar1.Panels(3))
 Else
 Call GuardarResVaRA(frmCalVar.StatusBar1.Panels(3))
 End If
Else
MsgBox "no puede"
End If
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Command3_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 Call CreartxtVaR(FechaPos)
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Command8_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 If frmResumenPos.Option36.Value Then opcion = 1
 If frmResumenPos.Option37.Value Then opcion = 2
 If frmResumenPos.Option38.Value Then opcion = 3
 Call ImpresionVaRFPensiones(opcion)
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Form_Load()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Call CambiarCuadro(frmResumenPos, "frmresumenpos.SSTab1", "frmresumenpos.MSFlexGrid1")
Call ResetearPantallasSistema
Call TitulosResumenPosicion(frmResumenPos.MSFlexGrid18, 1, 1)
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Form_Resize()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Call CambiarCuadro(frmResumenPos, "frmresumenpos.SSTab1", "frmresumenpos.MSFlexGrid1")
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub MSFlexGrid1_DblClick()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
IndicePos = MSFlexGrid1.Row
 frmTitulos.Show 1
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Sub CambiarCuadro(objeto, objeto1, objeto2)
if siactivarcontrolerrores then
On error goto ControlErrores
end if
On Error Resume Next
objeto1.Left = 200
objeto1.Top = 200
objeto2.Left = 200
objeto2.Top = 600
objeto1.Width = objeto.Width - 600
objeto1.Height = objeto.Height - 1200
objeto2.Width = objeto1.Width - 300
objeto2.Height = objeto1.Height - 800
On Error GoTo 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Option36_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 opcion = 1
 Call VerResumenVAR(frmResumenPos.MSFlexGrid18, opcion)
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Option37_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 opcion = 2
 Call VerResumenVAR(frmResumenPos.MSFlexGrid18, opcion)
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub

Private Sub Option38_Click()
if siactivarcontrolerrores then
On error goto ControlErrores
end if
Screen.MousePointer = 11
 opcion = 3
 Call VerResumenVAR(frmResumenPos.MSFlexGrid18, opcion)
Screen.MousePointer = 0
on error goto 0
Exit Sub
ControlErrores:
msgbox error(err())
on error goto 0
End Sub
