VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcesos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de procesos"
   ClientHeight    =   9045
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Generación de procesos 2"
      Height          =   600
      Left            =   6240
      TabIndex        =   9
      Top             =   300
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deseleccionar todos"
      Height          =   600
      Left            =   10170
      TabIndex        =   8
      Top             =   270
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar todos"
      Height          =   600
      Left            =   8400
      TabIndex        =   7
      Top             =   270
      Width           =   1500
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2250
      TabIndex        =   4
      Top             =   420
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generación de procesos 1"
      Height          =   600
      Left            =   4500
      TabIndex        =   3
      Top             =   300
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   200
      TabIndex        =   1
      Top             =   420
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7485
      Left            =   195
      TabIndex        =   0
      Top             =   1080
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   13203
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   9660
      TabIndex        =   6
      Top             =   390
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   113246209
      CurrentDate     =   41736
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final"
      Height          =   195
      Left            =   2250
      TabIndex        =   5
      Top             =   180
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de inicio"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Integer
For i = 1 To UBound(MatCatProcesos, 1)
   MSFlexGrid1.TextMatrix(i, 2) = "S"
Next i
End Sub

Private Sub Command2_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim indice As Long
Dim sigen As Boolean
Dim txtmsg As String
Dim exito1 As Boolean

Screen.MousePointer = 11
Command2.Enabled = False
If IsDate(Combo1.Text) And IsDate(Combo2.Text) Then
   SiActTProc = True
   fecha1 = CDate(Combo1.Text)
   fecha2 = CDate(Combo2.Text)
   fecha = fecha1
  'se generan las tareas para proceso de factores
   Do While fecha <= fecha2 And fecha <= Date
      If fecha < Date Or (fecha = Date And Time > #2:00:00 PM#) Then
         indice = BuscarValorArray(fecha, MatFechasTareas1, 1)
         If indice <> 0 And UBound(MatCatProcesos, 1) <> 0 Then
            Call GenProcesosDia(fecha, MatCatProcesos, 1)
         Else
            MsgBox "No se puede generar tareas diarias para esta fecha"
         End If
      End If
      Call GenerarFechasVaR(fecha, fecha, txtmsg, exito1)
      fecha = fecha + 1
      DoEvents
   Loop
   MatFechasVaR = LeerFechasVaRT()

   Call ActUHoraUsuario
   SiActTProc = False
 End If
 Command2.Enabled = True
 frmProcesos.SetFocus
 MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
Dim i As Integer
For i = 1 To UBound(MatCatProcesos, 1)
   MSFlexGrid1.TextMatrix(i, 2) = "N"
Next i
End Sub

Private Sub Command4_Click()
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim indice As Long
Dim sigen As Boolean
Dim txtmsg As String
Dim exito1 As Boolean

Screen.MousePointer = 11
Command4.Enabled = False
If IsDate(Combo1.Text) And IsDate(Combo2.Text) Then
   SiActTProc = True
   fecha1 = CDate(Combo1.Text)
   fecha2 = CDate(Combo2.Text)
   fecha = fecha1
  'se generan las tareas para proceso de factores
   Do While fecha <= fecha2 And fecha <= Date
      If fecha < Date Or (fecha = Date And Time > #2:00:00 PM#) Then
         indice = BuscarValorArray(fecha, MatFechasTareas1, 1)
         If indice <> 0 And UBound(MatCatProcesos, 1) <> 0 Then
            Call GenProcesosDia(fecha, MatCatProcesos, 2)
         Else
            MensajeProc = "No se puede generar tareas diarias para esta fecha"
         End If
      End If
      Call GenerarFechasVaR(fecha, fecha, txtmsg, exito1)
      fecha = fecha + 1
      DoEvents
   Loop
   MatFechasVaR = LeerFechasVaRT()

   Call ActUHoraUsuario
   SiActTProc = False
 End If
 Command4.Enabled = True
 frmProcesos.SetFocus
 MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim mata() As Date
Dim noreg As Integer
Dim i As Integer
Combo1.Clear
noreg = UBound(MatFechasVaR, 1)
For i = 1 To noreg
    Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
    Combo2.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
Combo1.Text = Date
Combo2.Text = Date
noreg = UBound(MatCatProcesos, 1)
MSFlexGrid1.Rows = noreg + 1
MSFlexGrid1.Cols = 3
MSFlexGrid1.ColWidth(0) = 1500
MSFlexGrid1.ColWidth(1) = 2500
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.TextMatrix(0, 0) = "No de tarea"
MSFlexGrid1.TextMatrix(0, 1) = "Descripcion"
MSFlexGrid1.TextMatrix(0, 2) = "Generar"
For i = 1 To noreg
    MSFlexGrid1.TextMatrix(i, 0) = MatCatProcesos(i, 1)
    MSFlexGrid1.TextMatrix(i, 1) = MatCatProcesos(i, 3)
    MSFlexGrid1.TextMatrix(i, 2) = "S"
Next i

End Sub


Private Sub MSFlexGrid1_DblClick()
Dim indice As Integer
indice = MSFlexGrid1.row
If (MSFlexGrid1.col = 2 Or MSFlexGrid1.col = 3) And indice >= 1 Then
If MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "S" Then
   MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "N"
ElseIf MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "N" Then
   MSFlexGrid1.TextMatrix(indice, MSFlexGrid1.col) = "S"
End If
End If
End Sub
