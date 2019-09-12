VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmABalanza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis de Balanza"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13470
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar analisis Balanza"
      Height          =   675
      Left            =   10260
      TabIndex        =   6
      Top             =   660
      Width           =   1785
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   7920
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   660
      Width           =   1755
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6030
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   630
      Width           =   1485
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   450
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   570
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar balanza en BD"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7005
      Left            =   510
      TabIndex        =   0
      Top             =   1800
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   12356
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   210
      Width           =   450
   End
End
Attribute VB_Name = "frmABalanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim nomarch As String
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim fechareg As Date
Dim fecharegtxt As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim mes As Integer
Dim txtfecha As String
Dim horareg As String
Dim txtcadena As String
Dim txtnomtabla As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtborra As String


Screen.MousePointer = 11
nomarch = "d:\mov volantes 2019-01.xls"
frmABalanza.CommonDialog1.FileName = nomarch
frmABalanza.CommonDialog1.ShowOpen
nomarch = frmABalanza.CommonDialog1.FileName
sihayarch = VerifAccesoArch(nomarch)
 

 If sihayarch Then
    frmProgreso.Show
    txtnomtabla = "mov_volantes"
    mes = 1
    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset(txtnomtabla & "$", dbOpenDynaset)
    registros1.MoveLast
    noreg = registros1.RecordCount
    ReDim mata(1 To 10, 1 To 1) As Variant
    registros1.MoveFirst
    contar = 0
    For i = 1 To noreg
        If Not EsVariableVacia(registros1.Fields(4)) Then
           contar = contar + 1
           ReDim Preserve mata(1 To 10, 1 To contar) As Variant
           mata(1, contar) = registros1.Fields(5)                                 'fecha de afectacion
           mata(2, contar) = registros1.Fields(6)                                 'fecha de registro
           mata(3, contar) = Right(registros1.Fields(1), 2)                       'moneda
           mata(4, contar) = Right(registros1.Fields(2), 4)                       'cuenta
           mata(5, contar) = Trim(registros1.Fields(4))                           'volante
           mata(6, contar) = CDbl(ReemplazaVacioValor(registros1.Fields(7), 0))   'cargo
           mata(7, contar) = CDbl(ReemplazaVacioValor(registros1.Fields(8), 0))   'abono
           mata(8, contar) = Trim(registros1.Fields(9))                           'concepto
           mata(9, contar) = Trim(registros1.Fields(10))                          'autorizo
           mata(10, contar) = Trim(registros1.Fields(11))                         'elaboro
        End If
        registros1.MoveNext
    Next i
    registros1.Close
    base1.Close
    mata = MTranV(mata)
    txtfecha = "TO_DATE('" & Format(#1/1/2019#, "DD/MM/YYYY") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaMovBalanza & "WHERE FECHA >= " & txtfecha
    For i = 1 To contar
        txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtfecha2 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtcadena = "INSERT into " & TablaMovBalanza & " VALUES("
        txtcadena = txtcadena & i & ","
        txtcadena = txtcadena & mes & ","
        txtcadena = txtcadena & txtfecha1 & ","
        txtcadena = txtcadena & txtfecha2 & ","
        txtcadena = txtcadena & "'" & mata(i, 3) & "',"
        txtcadena = txtcadena & "'" & mata(i, 4) & "',"
        txtcadena = txtcadena & "'" & mata(i, 5) & "',"
        txtcadena = txtcadena & mata(i, 6) & ","
        txtcadena = txtcadena & mata(i, 7) & ","
        txtcadena = txtcadena & "'" & mata(i, 8) & "',"
        txtcadena = txtcadena & "'" & mata(i, 9) & "',"
        txtcadena = txtcadena & "'" & mata(i, 10) & "')"
        ConAdo.Execute txtcadena
    Next i
 End If
 Unload frmProgreso
MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg0 As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
Dim j As Integer
Dim indice As Long
Dim rmesa As New ADODB.recordset
Dim contar As Long
Dim suma As Double

Screen.MousePointer = 11
fecha1 = #1/1/2019#
fecha2 = #1/31/2019#
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaFechasVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg0 = rmesa.Fields(0)
rmesa.Close
If noreg0 <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matf(1 To noreg0, 1 To 1) As Variant
   For i = 1 To noreg0
      matf(i, 1) = rmesa.Fields(0)
      rmesa.MoveNext
   Next i
   rmesa.Close
End If
txtfiltro2 = "SELECT * FROM " & TablaMovBalanza
txtfiltro2 = txtfiltro2 & " WHERE F_AFECTACION >= " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND F_AFECTACION <= " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND MONEDA = '" & "02" & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg1 <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg1, 1 To 11) As Variant
   suma = 0
   For i = 1 To noreg1
       mata(i, 1) = rmesa.Fields(0)     'id de volante
       mata(i, 2) = rmesa.Fields(2)     'fecha de afectacion
       mata(i, 3) = rmesa.Fields(3)     'fecha de registro
       mata(i, 4) = rmesa.Fields(6)     'CUENTA
       mata(i, 5) = rmesa.Fields(7)     'CARGO
       mata(i, 6) = rmesa.Fields(8)     'abono
       mata(i, 7) = rmesa.Fields(9)     'concepto
       mata(i, 8) = rmesa.Fields(10)    'autorizo
       mata(i, 9) = rmesa.Fields(11)    'elaboro
       suma = suma + mata(i, 5)
       rmesa.MoveNext
   Next i
   rmesa.Close
   contar = 0
   For i = 1 To noreg1
       For j = 1 To noreg1
           If mata(i, 5) = mata(j, 6) And mata(i, 2) = mata(j, 2) And mata(i, 5) <> 0 And mata(i, 11) <> "C" And mata(j, 11) <> "C" Then
              contar = contar + 1
              mata(j, 10) = mata(i, 1)
              mata(j, 11) = "C"
              mata(i, 11) = "C"
              Exit For
           End If
       Next j
   Next i
   MsgBox "No. de volantes conciliados " & contar
   Open "d:\resultados.txt" For Output As #1
   suma = 0
   For i = 1 To noreg1
   If mata(i, 11) <> "C" Then
     Print #1, mata(i, 1) & Chr(9) & mata(i, 2) & Chr(9) & mata(i, 3) & Chr(9) & mata(i, 4) & Chr(9) & -mata(i, 5) & Chr(9) & mata(i, 6) & Chr(9) & mata(i, 7)
     suma = suma + mata(i, 6) - mata(i, 5)
   End If
   Next i
   Close #1
End If

MsgBox "Fin de proceso"
Screen.MousePointer = 0
End Sub

