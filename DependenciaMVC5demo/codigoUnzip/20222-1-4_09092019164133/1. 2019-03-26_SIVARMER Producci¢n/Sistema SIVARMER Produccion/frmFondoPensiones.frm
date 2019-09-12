VERSION 5.00
Begin VB.Form frmFondoPensiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondo de pensiones"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Procesos"
      Height          =   6645
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   9735
      Begin VB.CommandButton Command11 
         Caption         =   "Valuación del Fondo de Pensiones"
         Height          =   800
         Left            =   5400
         TabIndex        =   13
         Top             =   2200
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   390
         TabIndex        =   11
         Top             =   570
         Width           =   2325
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Factores de riesgo incompletos"
         Height          =   800
         Left            =   8000
         TabIndex        =   10
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Lectura fondo pensiones sep 2018"
         Height          =   800
         Left            =   2130
         TabIndex        =   9
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Lectura pos Fondo Pensiones Nueva Estructura"
         Height          =   800
         Left            =   4050
         TabIndex        =   8
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Calculo de CVaR Fondo de Pensiones"
         Height          =   645
         Left            =   300
         TabIndex        =   7
         Top             =   3360
         Width           =   1500
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Generacion cuadros comparacion de valuación"
         Height          =   800
         Left            =   8000
         TabIndex        =   6
         Top             =   2200
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Actualizar Calificaciones vector precios"
         Height          =   800
         Left            =   2070
         TabIndex        =   5
         Top             =   2200
         Width           =   1500
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Actualizar calificaciones manualmente"
         Height          =   800
         Left            =   3720
         TabIndex        =   4
         Top             =   2200
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Correccion de errores en posición"
         Height          =   800
         Left            =   300
         TabIndex        =   3
         Top             =   2200
         Width           =   1500
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Validar datos posicion"
         Height          =   800
         Left            =   5910
         TabIndex        =   2
         Top             =   1200
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Lectura posicion Fondo pensiones"
         Height          =   800
         Left            =   300
         TabIndex        =   1
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   300
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmFondoPensiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'autor angel esteban castañeda castelan
'rutina para corregir errores de la posición del fondo de pensiones

Dim txtcadena As String
Dim txtfecha As String
Screen.MousePointer = 11
txtcadena = "DELETE FROM  " & TablaPosMD & " WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION IS NULL"
ConAdo.Execute txtcadena
txtcadena = "DELETE FROM  " & TablaPosDiv & " WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION IS NULL"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='00518',C_EMISION ='92FEFA00518' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '92FEFA518'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='00718',C_EMISION ='92FEFA00718' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '92FEFA718'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01218',C_EMISION ='92FEFA01218' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '92FEFA1218'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01518',C_EMISION ='92FEFA01518' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '92FEFA1518'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01118',C_EMISION ='93DAIMLER01118' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93DAIMLER1118'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='02018',C_EMISION ='93DAIMLER02018' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93DAIMLER2018'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='02318',C_EMISION ='93DAIMLER02318' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93DAIMLER2318'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='02618',C_EMISION ='93DAIMLER02618' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93DAIMLER2618'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01018', C_EMISION ='93NRF01018' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93NRF1018'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01318', C_EMISION ='93NRF01318' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93NRF1318'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01318',C_EMISION ='93PCARFM01318' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93PCARFM1318'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  EMISION = 'DOIHICB', SERIE ='13',C_EMISION ='91DOIHICB13' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91DOIHICB1313'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='09-4', C_EMISION ='91KIMBER09-4' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91KIMBER9-4'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='09', C_EMISION ='95CFECB09' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95CFECB9'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01118', C_EMISION ='93VWLEASE01118' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93VWLEASE1118'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18',EMISION ='SCOTIAB',C_EMISION ='94SCOTIAB18' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94BSCOTIAB18'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='03118', C_EMISION ='93VWLEASE03118' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93VWLEASE3118'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='02118', C_EMISION ='93PCARFM02118' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93PCARFM2118'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='01518', C_EMISION ='93PCARFM01518' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '93PCARFM1518'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='09',C_EMISION ='95CFECB09' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95CFECB9'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='09',C_EMISION ='95CFEHCB09' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95CFEHCB9'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-6',C_EMISION ='95FEFA17-6' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95FEFA43268'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='0624', C_EMISION ='JEAMX0624' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = 'JEAMX624'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='220113', C_EMISION ='LD220113' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = 'LD22013'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='10-2', C_EMISION ='90GDFCB10-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '90GDFCB43141'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='13-2', C_EMISION ='91AC13-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91AC43144'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='10-2', C_EMISION ='91AMX10-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91AMX43141'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='13-2', C_EMISION ='91BBVALMX13-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91BBVALMX43144'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-2', C_EMISION ='91BBVALMX18-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91BBVALMX43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-2', C_EMISION ='91DAIMLER18-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91DAIMLER43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-2', C_EMISION ='91FUNO17-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91FUNO43148'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='15-2', C_EMISION ='91GICSA15-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91GICSA43146'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='09-4', C_EMISION ='91KIMBER09-4' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91KIMBER43199'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-2', C_EMISION ='91LALA18-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91LALA43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-3', C_EMISION ='91LALA18-3' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91LALA43177'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-2', C_EMISION ='91LIVEPOL17-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91LIVEPOL43148'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-2', C_EMISION ='91VWLEASE17-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91VWLEASE43148'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-2', C_EMISION ='91VWLEASE18-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '91VWLEASE43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='12-3', C_EMISION ='94BACMEXT12-3' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94BACMEXT43171'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='10-2', C_EMISION ='94BANAMEX10-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94BANAMEX43141'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-2', C_EMISION ='94HSBC17-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94HSBC43148'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='16-2', C_EMISION ='94MULTIVA16-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94MULTIVA43147'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='17-3', C_EMISION ='94SCOTIAB17-3' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '94SCOTIAB43176'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='10-2', C_EMISION ='95CFE10-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95CFE43141'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='10-2', C_EMISION ='95CFECB10-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95CFECB43141'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-2', C_EMISION ='95FEFA18-2' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95FEFA43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='18-4', C_EMISION ='95FEFA18-4' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95FEFA43208'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  SERIE ='11-3', C_EMISION ='95PEMEX11-3' WHERE (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & ") AND C_EMISION = '95PEMEX43170'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'CD', EMISION ='BACMEXT', SERIE = '17-2',C_EMISION='CDBACMEXT17-2' WHERE C_EMISION = 'CDBACMEXT43148'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='BAC', SERIE = '6-10',C_EMISION='D8BAC6-10' WHERE C_EMISION = 'D8BAC43379'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='CACIB', SERIE = '1-10',C_EMISION='D8CACIB1-10' WHERE C_EMISION = 'D8CACIB43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='GS', SERIE = '1-10',C_EMISION='D8GS1-10' WHERE C_EMISION = 'D8GS43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'CD', EMISION ='BACMEXT', SERIE = '18-2',C_EMISION='CDBACMEXT18-2' WHERE C_EMISION = 'CDBACMEXT43149'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='BAC', SERIE = '6-10',C_EMISION='D8BAC6-10' WHERE C_EMISION = 'D8BAC43379'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='CACIB', SERIE = '1-10',C_EMISION='D8CACIB1-10' WHERE C_EMISION = 'D8CACIB43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='GS', SERIE = '1-10',C_EMISION='D8GS1-10' WHERE C_EMISION = 'D8GS43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='CACIB', SERIE = '1-10',C_EMISION='D8CACIB1-10' WHERE C_EMISION = 'D8CACIB43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='GS', SERIE = '1-10',C_EMISION='D8GS1-10' WHERE C_EMISION = 'D8GS43374'"
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET TV = 'D8', EMISION ='BAC', SERIE = '6-10',C_EMISION='D8BAC6-10' WHERE C_EMISION = 'D8BAC43379'"
ConAdo.Execute txtcadena

txtfecha = "TO_DATE('" & Format$(#12/5/2018#, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtcadena = "UPDATE " & TablaPosMD & " SET  EMISION ='CETELEM',C_EMISION= '91CETELEM17' WHERE C_EMISION = '91BNPPPF17' AND FECHAREG >=" & txtfecha
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  EMISION ='CETELEM',C_EMISION= '91CETELEM17-2' WHERE C_EMISION = '91BNPPPF17-2' AND FECHAREG >=" & txtfecha
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  EMISION ='CETELEM',C_EMISION= '91CETELEM18' WHERE C_EMISION = '91BNPPPF18' AND FECHAREG >=" & txtfecha
ConAdo.Execute txtcadena
txtcadena = "UPDATE " & TablaPosMD & " SET  EMISION ='CETELEM',C_EMISION= '91CETELEM18-2' WHERE C_EMISION = '91BNPPPF18-2' AND FECHAREG >=" & txtfecha
ConAdo.Execute txtcadena
txtfecha = "TO_DATE('" & Format$(#12/4/2018#, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtcadena = "UPDATE " & TablaPosMD & " SET SERIE = 'BPE',C_EMISION= '52GBMCREBPE' WHERE C_EMISION = '52GBMCREBP' AND FECHAREG >=" & txtfecha
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command10_Click()
Dim txtfecha As String
Dim txtborra As String
Dim matarch1() As String
Dim nomtabla1() As String

Dim matarch2() As String
Dim nomtabla2() As String

Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim notabla1 As Integer
Dim notabla2 As Integer
Dim nreg1 As Long
Dim nreg2 As Long
Dim exito As Boolean
Dim contar As Long

SiActTProc = True
noreg = 1
    Dim mata() As Variant
    Screen.MousePointer = 11
    frmProgreso.Show
    ReDim matfecha(1 To noreg) As Date
    matfecha(1) = #9/28/2018#
    For i = 1 To noreg
       Call ImpPosPensiones2(matfecha(i))
    Next i
    Unload frmProgreso
    Screen.MousePointer = 0
    SiActTProc = False
    MsgBox "Fin de proceso"

End Sub

Private Sub Command11_Click()
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim fecha As Date
Dim exito As Boolean
Dim txtport As String
Dim txtportfr As String
Dim txtmsg As String

If IsDate(Combo1.Text) Then
   fecha = CDate(Combo1.Text)
   frmProgreso.Show
   txtportfr = "Normal"
   SiActTProc = True
   ValExacta = True
   MatCurvasT = LeerCurvaCompleta(fecha, exito)
   Call GenPortPensiones(fecha)
   txtport = "FID 2065"
   matport1 = DefinePortFP1()
   Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, 1, txtmsg, exito)
   If exito Then Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, 1, exito)
   Call GenValPortPosPension(fecha, fecha, fecha, txtport, txtport, txtportfr, 1, exito)
   For i = 1 To UBound(matport1, 1)
       Call GenValPortPosPension(fecha, fecha, fecha, txtport, matport1(i, 1), txtportfr, 1, exito)
   Next i
   matport1 = DefinePortFP2()
   txtport = "FID 2160"
   Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, 1, txtmsg, exito)
   If exito Then Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, 1, exito)
   Call GenValPortPosPension(fecha, fecha, fecha, txtport, txtport, txtportfr, 1, exito)
   For i = 1 To UBound(matport1, 1)
       Call GenValPortPosPension(fecha, fecha, fecha, txtport, matport1(i, 1), txtportfr, 1, exito)
   Next i
   SiActTProc = False
   ValExacta = False
   Unload frmProgreso
End If
MsgBox "Fin del proceso"
End Sub

Private Sub Command2_Click()
Dim txtfecha As String
Dim txtborra As String
Dim matarch1() As String
Dim nomtabla1() As String

Dim matarch2() As String
Dim nomtabla2() As String

Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim notabla1 As Integer
Dim notabla2 As Integer
Dim nreg1 As Long
Dim nreg2 As Long
Dim exito As Boolean
Dim contar As Long

SiActTProc = True
noreg = 5


    Dim mata() As Variant
    Screen.MousePointer = 11
    frmProgreso.Show
    ReDim matfecha(1 To noreg) As Date
    matfecha(1) = #4/30/2018#
    matfecha(2) = #5/31/2018#
    matfecha(3) = #6/29/2018#
    matfecha(4) = #7/31/2018#
    matfecha(5) = #8/31/2018#
    For i = 1 To noreg
       Call ImpPosPensiones(matfecha(i))
    Next i
    Unload frmProgreso
    Screen.MousePointer = 0
    SiActTProc = False
    MsgBox "Fin de proceso"

End Sub

Private Sub Command3_Click()
Dim fecha As Date
If IsDate(Combo1.Text) Then
fecha = CDate(Combo1.Text)
    SiActTProc = True
    Screen.MousePointer = 11
    frmProgreso.Show
    Open "d:\VAR_TC_VAL_ST_CUPON_2 " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
    Open "d:\VAR_TC_VAL_BONOS_2 " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #2
    Open "d:\VAR_TC_VAL_IND " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #3
    Open "d:\VAR_TC_IND_VPRECIOS " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #4
    Open "d:\VAR_flujos_md " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #5
    Open "d:\VAR_port_fr_2 " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #6
    Open "d:\inst sin modelo val " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #7
    Open "d:\inst que no estan el el vector " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #8
    Call ValidarPosPension1(fecha, ClavePosPension1)
    Call ValidarPosPension1(fecha, ClavePosPension2)
    DoEvents
   Close #1
   Close #2
   Close #3
   Close #4
   Close #5
   Close #6
   Close #7
   Close #8
   Unload frmProgreso
   SiActTProc = False
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim fecha As Date
Dim matfechas() As Date
Dim i As Integer
Dim noreg As Integer
noreg = 7
ReDim matfechas(1 To noreg) As Date
fecha = CDate(Combo1.Text)
Screen.MousePointer = 11
frmProgreso.Show
  Call ActCalifPosFP1(fecha, ClavePosPension1)
  Call ActCalifPosFP2(fecha, ClavePosPension1)
  Call ActCalifPosFP1(fecha, ClavePosPension2)
  Call ActCalifPosFP2(fecha, ClavePosPension2)
Unload frmProgreso
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Sub ActCalifPosFP1(ByVal fecha As Date, ByVal cposicion As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim i As Long
Dim calif As String
Dim siflujos As String
Dim rmesa As New ADODB.recordset
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 5) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("COPERACION")
       mata(i, 2) = rmesa.Fields("TV")
       mata(i, 3) = rmesa.Fields("EMISION")
       mata(i, 4) = rmesa.Fields("SERIE")
       mata(i, 5) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   matvp = LeerVPrecios(fecha, mindvp)
   For i = 1 To UBound(mata, 1)
        indice = BuscarValorArray(mata(i, 5), mindvp, 1)
        calif = DefinirCalifEmFP(fecha, mata(i, 5), mindvp(indice, 2), matvp)
        siflujos = ConvBolStr(DetermSiEmFlujos(mata(i, 2), mata(i, 3), mata(i, 4)))
        txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtcadena = "UPDATE " & TablaPosMD & " SET CALIFICACION = '" & calif & "', SI_FLUJOS ='" & siflujos & "'"
        txtcadena = txtcadena & " WHERE FECHAREG = " & txtfecha
        txtcadena = txtcadena & " AND C_EMISION = '" & mata(i, 5) & "'"
        txtcadena = txtcadena & " AND CPOSICION = " & cposicion
        txtcadena = txtcadena & " AND COPERACION= '" & mata(i, 1) & "'"
        ConAdo.Execute txtcadena
    Next i
End If
End Sub

Sub ActCalifPosFP2(ByVal fecha As Date, ByVal cposicion As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim calif As String
Dim i As Long
Dim rmesa As New ADODB.recordset
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 5) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("COPERACION")
       mata(i, 2) = rmesa.Fields("TV")
       mata(i, 3) = rmesa.Fields("EMISION")
       mata(i, 4) = rmesa.Fields("SERIE")
       mata(i, 5) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   matvp = LeerVPrecios(fecha, mindvp)
   For i = 1 To UBound(mata, 1)
       indice = BuscarValorArray(mata(i, 5), mindvp, 1)
       calif = DefinirCalifEmFP(fecha, mata(i, 5), mindvp(indice, 2), matvp)
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "UPDATE " & TablaPosDiv & " SET CALIFICACION = '" & calif & "'"
       txtcadena = txtcadena & " WHERE FECHAREG = " & txtfecha
       txtcadena = txtcadena & " AND C_EMISION = '" & mata(i, 5) & "'"
       txtcadena = txtcadena & " AND CPOSICION = " & cposicion
       txtcadena = txtcadena & " AND COPERACION= '" & mata(i, 1) & "'"
       ConAdo.Execute txtcadena
   Next i
End If
End Sub



Private Sub Command5_Click()
Dim txtfecha As String
Dim txtcadena As String
Dim nomarch As String
Dim sihayarch As Boolean
Dim mata() As Variant
Dim i As Long

Screen.MousePointer = 11
frmEjecucionProc2.CommonDialog1.ShowOpen
nomarch = frmEjecucionProc2.CommonDialog1.FileName
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   SiActTProc = True
   frmProgreso.Show
   mata = LeerArchCalificaciones(nomarch)
   For i = 1 To UBound(mata, 1)
       txtfecha = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "UPDATE " & TablaPosMD & " SET CALIFICACION = '" & mata(i, 3) & "'"
       txtcadena = txtcadena & " WHERE FECHAREG = " & txtfecha
       txtcadena = txtcadena & " AND EMISION = '" & mata(i, 2) & "'"
       txtcadena = txtcadena & " AND (CPOSICION = 5 OR CPOSICION = 6)"
       ConAdo.Execute txtcadena
       txtcadena = "UPDATE " & TablaPosDiv & " SET CALIFICACION = '" & mata(i, 3) & "'"
       txtcadena = txtcadena & " WHERE FECHAREG = " & txtfecha
       txtcadena = txtcadena & " AND EMISION = '" & mata(i, 2) & "'"
       txtcadena = txtcadena & " AND (CPOSICION = 5 OR CPOSICION = 6)"
       ConAdo.Execute txtcadena
   Next i
   SiActTProc = False
   Unload frmProgreso
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0

End Sub

Private Sub Command6_Click()
Dim fecha As Date
If IsDate(Combo1.Text) Then
   SiActTProc = True
   fecha = CDate(Combo1.Text)
   Screen.MousePointer = 11
   frmProgreso.Show
   Call GenValuacionFP1(fecha)
   Call GenValuacionFP2(fecha)
   Unload frmProgreso
   SiActTProc = False
End If
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
Dim noreg As Integer
Dim txtmsg As String
Dim matport1() As String
Dim matport2() As String
Dim txtport1 As String
Dim txtport2 As String
Dim i As Integer
Dim fecha As Date
Dim noesc As Integer
Dim htiempo As Integer

htiempo = 1
If IsDate(Combo1.Text) Then
   fecha = CDate(Combo1.Text)
   If fecha < #6/29/2018# Then
      noesc = 250
   Else
      noesc = 500
   End If
   Screen.MousePointer = 11
   SiActTProc = True
   ValExacta = True
   frmProgreso.Show
   matport1 = DefinePortFP1()
   txtport1 = "FID 2065"
   Call GenPortPensiones(fecha)
   Call ProcCalculoCVaRFondoP(fecha, txtport1, matport1, False, DirTemp, noesc, htiempo, txtmsg)
   matport2 = DefinePortFP2()
   txtport2 = "FID 2160"
   Call GenPortPensionesb(fecha)
   Call ProcCalculoCVaRFondoP(fecha, txtport2, matport2, False, DirTemp, noesc, htiempo, txtmsg)
   Unload frmProgreso
   Call ActUHoraUsuario
   SiActTProc = False
   ValExacta = False
End If
MsgBox "Fin de proceso"
frmFondoPensiones.SetFocus
Screen.MousePointer = 0
End Sub

Private Sub Command22_Click()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fecha As Date
Dim txtfecha As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim txtcadena As String

fecha = #4/30/2018#
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosPension1
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 14) As Variant
   For i = 1 To noreg
       mata(i, 1) = Format(rmesa.Fields("FECHAREG"), "YYYYMMDD")
       mata(i, 2) = rmesa.Fields("COPERACION")
       mata(i, 3) = "'" & rmesa.Fields("TV")
       mata(i, 4) = rmesa.Fields("EMISION")
       mata(i, 5) = "'" & rmesa.Fields("SERIE")
       If rmesa.Fields("TOPERACION") = 1 Then
          mata(i, 6) = 0
       Else
          mata(i, 6) = 1
       End If
       mata(i, 7) = rmesa.Fields("NO_TITULOS")
       mata(i, 8) = rmesa.Fields("T_REPORTO")
       mata(i, 9) = rmesa.Fields("F_COMPRA")
       mata(i, 10) = rmesa.Fields("F_VENC_OPER")
       mata(i, 11) = rmesa.Fields("SUBPORT_1")
       mata(i, 12) = rmesa.Fields("SUBPORT2")
       If rmesa.Fields("CPOSICION") = ClavePosPension1 Then
          mata(i, 13) = 2065
       Else
          mata(i, 13) = 2160
       End If
       mata(i, 14) = rmesa.Fields("P_COMPRA")
       rmesa.MoveNext
   Next i
   rmesa.Close
   Open "D:\SALIDA.TXT" For Output As #1
   For i = 1 To noreg
       txtcadena = ""
       For j = 1 To 14
           txtcadena = txtcadena & mata(i, j) & Chr(9)
       Next j
       Print #1, txtcadena
   Next i
   Close #1
End If
   MsgBox "Fin de proceso"

End Sub

Private Sub Command8_Click()
Dim noreg1 As Integer
Dim i As Integer
noreg1 = 7
    ReDim matfecha(1 To noreg1) As Date
   matfecha(1) = #4/30/2018#
   matfecha(2) = #5/31/2018#
   matfecha(3) = #6/29/2018#
   matfecha(4) = #7/31/2018#
   matfecha(5) = #8/31/2018#
   matfecha(6) = #9/28/2018#
   matfecha(7) = #10/31/2018#
   SiActTProc = True
   Screen.MousePointer = 11
   frmProgreso.Show
   Open "d:\factores r incompletos.txt" For Output As #7
   For i = 1 To noreg1
       Call ValidarPosPension3(matfecha(i), ClavePosPension1, 7)
       Call ValidarPosPension3(matfecha(i), ClavePosPension2, 7)
       DoEvents
   Next i
   Close #7
Unload frmProgreso
MsgBox "Fin de proceso"
SiActTProc = False
Screen.MousePointer = 0

End Sub

Private Sub Command9_Click()
Dim mata1() As Variant
Dim mata2() As Variant
Dim txtfecha As String
Dim txtnomarch1 As String
Dim txtnomarch2 As String
Dim txthojacalc1 As String
Dim txthojacalc2 As String
Dim txtborra As String
Dim fecha As Date
Dim contar1 As Long
Dim contar2 As Long
Dim nreg1 As Long
Dim nreg2 As Long
Dim exito As Boolean
If IsDate(Combo1.Text) Then
   fecha = CDate(Combo1.Text)
   txtnomarch1 = "d:\fp\Fid. 2065 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
   txtnomarch2 = "d:\fp\Fid. 2160 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
txthojacalc1 = "Hoja1"
txthojacalc2 = "Hoja1"
Screen.MousePointer = 11
frmProgreso.Show
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
ConAdo.Execute txtborra
mata1 = LeerHojaCalc5(txtnomarch1, txthojacalc1)
mata1 = depurartablafp3(mata1, fecha, ClavePosPension1, contar1)
Call ImpPosFid(mata1, 2065, nreg1, exito)
mata2 = LeerHojaCalc6(txtnomarch2, txthojacalc2)
mata2 = depurartablafp3(mata2, fecha, ClavePosPension2, contar2)
Call ImpPosFid(mata2, 2160, nreg1, exito)
Unload frmProgreso
End If
Screen.MousePointer = 0
MsgBox "Fin de proceso"
End Sub

Private Sub Form_Load()
Dim noreg As Integer
Dim i As Integer
noreg = UBound(MatFechasVaR, 1)
Combo1.Clear
For i = 1 To noreg
  Combo1.AddItem MatFechasVaR(noreg - i + 1, 1)
Next i
End Sub
