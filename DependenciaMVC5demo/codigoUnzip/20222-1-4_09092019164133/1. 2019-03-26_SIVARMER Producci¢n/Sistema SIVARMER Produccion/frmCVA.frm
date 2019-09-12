VERSION 5.00
Begin VB.Form frmCVA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo de CVA"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   825
      Left            =   4350
      TabIndex        =   10
      Top             =   3630
      Width           =   1125
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Validacion de contrapartes CVA"
      Height          =   615
      Left            =   4260
      TabIndex        =   9
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tabla de subprocesos"
      Height          =   1455
      Left            =   3870
      TabIndex        =   6
      Top             =   600
      Width           =   2685
      Begin VB.OptionButton Option2 
         Caption         =   "Subprocesos 2"
         Height          =   195
         Left            =   510
         TabIndex        =   8
         Top             =   810
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Subprocesos 1"
         Height          =   195
         Left            =   510
         TabIndex        =   7
         Top             =   390
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CVA para posicion simulada"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesos"
      Height          =   4575
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   2925
      Begin VB.CommandButton Command4 
         Caption         =   "Generar subprocesos consolidar CVA Derivados"
         Height          =   800
         Left            =   200
         TabIndex        =   4
         Top             =   2340
         Width           =   2000
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar subprocesos de Wrong Risk Way"
         Height          =   800
         Left            =   200
         TabIndex        =   3
         Top             =   3450
         Width           =   2000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Generar subrocesos CVA Deuda"
         Height          =   800
         Left            =   200
         TabIndex        =   2
         Top             =   1380
         Width           =   2000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar subprocesos CVA Derivados"
         Height          =   800
         Left            =   200
         TabIndex        =   1
         Top             =   360
         Width           =   2000
      End
   End
End
Attribute VB_Name = "frmCVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim dtfecha As Date
Dim txtfecha As String
Dim txtmsg As String
Dim exito As Boolean
If Option1.value Then
   id_tabla = 1
Else
   id_tabla = 2
End If
txtfecha = InputBox("Dame la fecha", , Date)
If IsDate(txtfecha) Then
   Screen.MousePointer = 11
   dtfecha = CDate(txtfecha)
   frmProgreso.Show
   Call GeneraLSubprocCVA(dtfecha, 83, "DERIVADOS", 500, 1, id_tabla, txtmsg, exito)
   Unload frmProgreso
   Screen.MousePointer = 0
End If
MsgBox "Fin de proceso"
End Sub

Private Sub Command2_Click()
    Dim txtfecha As String
    Dim vrecupera As Double
    Dim valcva    As Double
    Dim dtfecha As Date
    Dim dtfecha1 As Date
    Dim noesc As Integer, i As Integer, j As Integer
    Dim htiempo As Integer
    Dim siesfv As Boolean
    Dim suma As Double
    Dim noreg As Long
    Dim txtfiltro As String
    Dim mata() As String
    Dim exito As Boolean
    Dim contar As Long
    Dim txtcadena As String
    Dim id_tabla As Integer
    If frmCVA.Option1.value Then
       id_tabla = 1
    ElseIf frmCVA.Option2.value Then
       id_tabla = 2
    End If
    Screen.MousePointer = 11
    txtfecha = InputBox("Dame la fecha de calculo ", , Date)
    If IsDate(txtfecha) Then
        SiActTProc = True
        frmProgreso.Show
        dtfecha = CDate(txtfecha)
        noesc = 500
        htiempo = 1
        txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
        txtborra = txtborra & " AND ID_SUBPROCESO = 92"
        ConAdo.Execute txtborra
        txtborra = "DELETE FROM " & TablaPYGCVAMD & " WHERE FECHA = " & txtfecha
        ConAdo.Execute txtborra
        contar = DeterminaMaxRegSubproc(id_tabla)
        For j = 1 To 3
            Call CrearPortEmxContrap(dtfecha, mata, j)
            For i = 1 To UBound(mata, 1)
                contar = contar + 1
                txtcadena = CrearCadInsSub(dtfecha, 92, contar, "CVA MD", mata(i), j, noesc, htiempo, 0, "CVA", "", "", "", "", "", "", id_tabla)
                ConAdo.Execute txtcadena
                contar = contar + 1
                txtcadena = CrearCadInsSub(dtfecha, 92, contar, "CVA MD", mata(i), j, noesc, htiempo, 1, "Estr1", "", "", "", "", "", "", id_tabla)
                ConAdo.Execute txtcadena
                contar = contar + 1
                txtcadena = CrearCadInsSub(dtfecha, 92, contar, "CVA MD", mata(i), j, noesc, htiempo, 2, "Estr2", "", "", "", "", "", "", id_tabla)
                ConAdo.Execute txtcadena
                DoEvents
            Next i
        Next j
        Call ActUHoraUsuario
        SiActTProc = False
        Unload frmProgreso
        MsgBox "Fin de proceso"
    End If
    Screen.MousePointer = 0

End Sub

Private Sub Command3_Click()
   Dim dtfecha As Date
   Dim txtfecha As String
   Dim id_tabla As Integer
   If frmCVA.Option1.value Then
      id_tabla = 1
   ElseIf frmCVA.Option2.value Then
      id_tabla = 2
   End If
   txtfecha = InputBox("Dame la fecha a calcular", , Date)
   If IsDate(txtfecha) Then
      dtfecha = CDate(txtfecha)
      SiActTProc = True
      Screen.MousePointer = 11
      frmProgreso.Show
      Call GenSubprocWRW(dtfecha, id_tabla)
      Unload frmProgreso
      Screen.MousePointer = 0
      Call ActUHoraUsuario
      SiActTProc = False
      MsgBox "Proceso finalizado"
   End If

End Sub

Private Sub Command4_Click()
Dim dtfecha As Date
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim i As Long
Dim id_proc As Integer
Dim contar As Long
Dim mata() As Integer
Dim id_tabla As Integer
Dim txttabla As String
If frmCVA.Option1.value Then
   id_tabla = 1
ElseIf frmCVA.Option2.value Then
   id_tabla = 2
End If
Screen.MousePointer = 11
   txtfecha = InputBox("Dame la fecha a calcular", , Date)
   id_proc = 84
   If IsDate(txtfecha) Then
   dtfecha = CDate(txtfecha)
   SiActTProc = True
   frmProgreso.Show
   mata = LeerContrapFecha(dtfecha)
   txtfecha = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
   ConAdo.Execute txtborra
   txtborra = "DELETE FROM " & TablaResCVA & " WHERE FECHA = " & txtfecha & " AND CPOSICION = 'DER'"
   ConAdo.Execute txtborra
   contar = DeterminaMaxRegSubproc(id_tabla)
      For i = 1 To UBound(mata, 1)
          contar = contar + 1
          Call GenSubpConsolCVA(id_proc, contar, dtfecha, mata(i), 0, 0, "CVA", "CVA", id_tabla)
          contar = contar + 1
          Call GenSubpConsolCVA(id_proc, contar, dtfecha, mata(i), 0.95, 1, "Estr1", "Estr1", id_tabla)
          contar = contar + 1
          Call GenSubpConsolCVA(id_proc, contar, dtfecha, mata(i), 0.99, 2, "Estr2", "Estr2", id_tabla)
      Next i
   Call ActUHoraUsuario
   SiActTProc = False
   Unload frmProgreso
   MsgBox "Fin de proceso"
   End If
Screen.MousePointer = 0

End Sub

Private Sub Command5_Click()
Dim txtnompos As String
Dim fecha As Date
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda

Dim i As Long
Dim contar As Long
Dim exito As Boolean
Dim txtborra As String
Dim txtfecha As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtmsg0 As String

Screen.MousePointer = 11
txtnompos = Text16.Text
fecha = CDate(Text17.Text)
noesc = 500
htiempo = 1
frmProgreso.Show
    mattxt = CrearFiltroPosSim(txtnompos)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
    If Not EsArrayVacio(matpos) Then
       txtborra = "DELETE FROM " & TablaPLEscCVA & "WHERE FECHA = " & txtfecha
       txtborra = txtborra & " AND NOMPOS = '" & txtnompos & "'"
       For i = 1 To UBound(matpos, 1)
           contar = contar + 1
           Call GenRegSubpCVA(110, contar, fecha, matpos(i).tipopos, matpos(i).fechareg, matpos(i).nompos, matpos(i).HoraRegOp, matpos(i).C_Posicion, matpos(i).c_operacion, noesc, htiempo, 1)
       Next i
    End If
    MensajeProc = "El proceso finalizo correctamente"
Unload frmProgreso
Screen.MousePointer = 0
End Sub

Private Sub Command6_Click()
Dim fecha As Date
Dim mata() As String
mata = ValidarContrapartesPosMD(#2/28/2019#)
End Sub

Private Sub Command7_Click()
Screen.MousePointer = 11
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txtsubport As String
Dim rmesa As New ADODB.recordset
Dim fecha As Date
Dim txtport As String
Dim noesc As Integer
Dim htiempo As Integer
Dim txtmsg As String
Dim exito As Boolean
Dim valcva As Double
Dim califica As Integer
Dim escala As String
Dim Sector As String
Dim txtnomarch As String

txtport = "TOTAL"
fecha = #2/28/2019#
noesc = 500
htiempo = 1
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION,EMISION,CPOSICION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND  (CPOSICION =  " & ClavePosMD & " OR CPOSICION = " & ClavePosTeso & " OR CPOSICION =" & ClavePosPIDV & " OR CPOSICION = " & ClavePosPICV & ")"
txtfiltro2 = txtfiltro2 & " AND (TOPERACION =1 OR TOPERACION = 4)"
txtfiltro2 = txtfiltro2 & " GROUP BY C_EMISION,EMISION,CPOSICION,TOPERACION ORDER BY CPOSICION,EMISION,C_EMISION,TOPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 4) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = rmesa.Fields("EMISION")
       matem(i, 3) = rmesa.Fields("CPOSICION")
       matem(i, 4) = rmesa.Fields("TOPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtnomarch = "D:\RES CVA MD " & Format(fecha, "YYYY-MM-DD") & ".TXT"
   frmCalVar.CommonDialog1.FileName = txtnomarch
   frmCalVar.CommonDialog1.ShowSave
   txtnomarch = frmCalVar.CommonDialog1.FileName
   Open txtnomarch For Output As #1
   Print #1, "Clave de posicion" & Chr(9) & "Emisión" & Chr(9) & "Clave de emision" & Chr(9) & "CVA" & Chr(9) & "Sector" & Chr(9) & "Escala" & Chr(9) & "Calificacion"
   For i = 1 To UBound(matem, 1)
       txtsubport = "EM " & matem(i, 1) & " POS " & matem(i, 3) & " T_OP " & matem(i, 4)
       Call CalcCVAEmMD(fecha, txtport, txtsubport, matem(i, 1), matem(i, 2), noesc, htiempo, valcva, califica, 0, escala, Sector, txtmsg, exito)
       Print #1, "CVA" & Chr(9) & matem(i, 3) & Chr(9) & matem(i, 2) & Chr(9) & matem(i, 1) & Chr(9) & valcva & Chr(9) & Sector & Chr(9) & escala & Chr(9) & califica
       Call CalcCVAEmMD(fecha, txtport, txtsubport, matem(i, 1), matem(i, 2), noesc, htiempo, valcva, califica, 1, escala, Sector, txtmsg, exito)
       Print #1, "1 notch" & Chr(9) & matem(i, 3) & Chr(9) & matem(i, 2) & Chr(9) & matem(i, 1) & Chr(9) & valcva & Chr(9) & Sector & Chr(9) & escala & Chr(9) & califica
       Call CalcCVAEmMD(fecha, txtport, txtsubport, matem(i, 1), matem(i, 2), noesc, htiempo, valcva, califica, 2, escala, Sector, txtmsg, exito)
       Print #1, "2 notch" & Chr(9) & matem(i, 3) & Chr(9) & matem(i, 2) & Chr(9) & matem(i, 1) & Chr(9) & valcva & Chr(9) & Sector & Chr(9) & escala & Chr(9) & califica
   Next i
   Close #1
End If
Screen.MousePointer = 0
MsgBox "Fin de proceso"

End Sub
