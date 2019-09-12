VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmValidaOper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación de operaciones intradia"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2430
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2010
      TabIndex        =   8
      Top             =   90
      Width           =   2205
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Importar posicion IKOS a SIVARMER"
      Height          =   600
      Left            =   390
      TabIndex        =   5
      Top             =   8190
      Width           =   1700
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   200
      TabIndex        =   0
      Top             =   500
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   12991
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Negociacion"
      TabPicture(0)   =   "frmValidaOper.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSFlexGrid1"
      Tab(0).Control(1)=   "Command2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cobertura"
      TabPicture(1)   =   "frmValidaOper.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MSFlexGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de calculo"
         Height          =   1000
         Left            =   6960
         TabIndex        =   9
         Top             =   5940
         Width           =   8955
         Begin VB.OptionButton Option5 
            Caption         =   "Proxy swap"
            Height          =   195
            Left            =   6990
            TabIndex        =   14
            Top             =   400
            Width           =   1245
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Forward"
            Height          =   255
            Left            =   5700
            TabIndex        =   13
            Top             =   400
            Width           =   1005
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Activo y pasivo"
            Height          =   195
            Left            =   3500
            TabIndex        =   12
            Top             =   400
            Width           =   1500
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Activa"
            Height          =   195
            Left            =   2000
            TabIndex        =   11
            Top             =   400
            Width           =   1305
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pasiva"
            Height          =   195
            Left            =   200
            TabIndex        =   10
            Top             =   400
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cargar posiciones primarias swaps"
         Height          =   600
         Left            =   2190
         TabIndex        =   6
         Top             =   6540
         Width           =   1700
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Generar proceso para validar operaciones"
         Height          =   600
         Left            =   -74700
         TabIndex        =   4
         Top             =   6540
         Width           =   1700
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar proceso para validar operaciones"
         Height          =   600
         Left            =   240
         TabIndex        =   3
         Top             =   6540
         Width           =   1700
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   5145
         Left            =   195
         TabIndex        =   2
         Top             =   510
         Width           =   15825
         _ExtentX        =   27914
         _ExtentY        =   9075
         _Version        =   393216
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5685
         Left            =   -74805
         TabIndex        =   1
         Top             =   510
         Width           =   15765
         _ExtentX        =   27808
         _ExtentY        =   10028
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de validación:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "frmValidaOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Screen.MousePointer = 11
Dim noreg As Integer
Dim i As Integer
Dim coperacion As String
Dim txtport As String
Dim fecha As Date
Dim horareg As String
Dim txtnompos As String
Dim id_tabla As Integer
If IsDate(Text1.Text) Then
Command1.Enabled = False
fecha = CDate(Text1.Text)
noreg = MSFlexGrid2.Rows - 1
id_tabla = 3
If noreg <> 0 And MSFlexGrid2.Cols > 3 Then
   For i = 1 To noreg
       horareg = MSFlexGrid2.TextMatrix(i, 5)
       txtnompos = "Intradia " & Format(fecha, "dd/mm/yyyy")
       If MSFlexGrid2.TextMatrix(i, 8) = "S" Then
          coperacion = MSFlexGrid2.TextMatrix(i, 1)
          If MSFlexGrid2.TextMatrix(i, 7) = "Swap" Then
             txtport = "Efec Pros swap " & coperacion
             If Option1.value Or Option2.value Or Option3.value Then
                Call DetSubportEfecCobSwap(fecha, txtport, 3, txtnompos, horareg, coperacion)

             ElseIf Option5.value Then
                Call DetSubportEfecCobPSwap(fecha, txtport, 3, txtnompos, horareg, coperacion)
             End If
             If Option1.value Then
               Call GenProcEfecProsSwap(fecha, txtport, 58, id_tabla)
             ElseIf Option2.value Then
               Call GenProcEfecProsSwap(fecha, txtport, 59, id_tabla)
             ElseIf Option3.value Then
               Call GenProcEfecProsSwap(fecha, txtport, 60, id_tabla)
             ElseIf Option5.value Then
               Call GenProcEfecProsSwap(fecha, txtport, 62, id_tabla)
             End If
          ElseIf MSFlexGrid2.TextMatrix(i, 7) = "Fwd" Then
             txtport = "Efec Pros fwd " & coperacion
             Call DetSubportEfecCobFwd(fecha, txtport, 3, txtnompos, horareg, coperacion)
             Call GenProcEfecFwd(fecha, txtport, 61, id_tabla)
          ElseIf Option5.value Then
          
          End If
       End If
   Next i
End If
Command1.Enabled = True
End If
MsgBox "Se crearon los subprocesos de efectividad"
Screen.MousePointer = 0
End Sub

Sub GenProcEfecFwd(ByVal fecha As Date, ByVal txtport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim contar As Long
Dim txttabla As String
txttabla = DetermTablaSubproc(id_tabla)

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtfecha1 = Format(fecha, "dd/mm/yyyy")
    'txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Efect Cob Prospectiva Fwd", tipopos, txtfecha1, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
    txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Efect Cob Prospectiva Fwd", txtfecha1, txtport, "", "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
    
End Sub

Sub DetSubportEfecCobFwd(ByVal fecha As Date, ByVal txtport As String, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByVal coperacion As String)
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtoper(1 To 1) As String
Dim cpos(1 To 1) As Integer
Dim hreg(1 To 1) As String
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtoper(1) = coperacion
cpos(1) = ClavePosDeriv
hreg(1) = horareg
txtcadena = "DELETE FROM " & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtcadena
For i = 1 To 1
    txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de la posicion
    txtcadena = txtcadena & "'" & txtport & "',"                  'el nombre del portafolio
    txtcadena = txtcadena & "'" & tipopos & "',"                  'tipo de posicion
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de registro
    txtcadena = txtcadena & "'" & txtnompos & "',"                'nombre de la posicion
    txtcadena = txtcadena & "'" & hreg(i) & "',"                  'la hora de registro
    txtcadena = txtcadena & cpos(i) & ","                         'la clave de posicion
    txtcadena = txtcadena & "'" & txtoper(i) & "')"               'la clave de OPERACION
    ConAdo.Execute txtcadena
Next i

End Sub


Sub DetSubportEfecCobSwap(ByVal fecha As Date, ByVal txtport As String, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByVal coperacion As String)
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtoper(1 To 4) As String
Dim cpos(1 To 4) As Integer
Dim hreg(1 To 4) As String
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtoper(1) = coperacion
txtoper(2) = "PRIMARIA " & coperacion
txtoper(3) = "PRIMARIA " & coperacion & " A"
txtoper(4) = "PRIMARIA " & coperacion & " P"
cpos(1) = ClavePosDeriv
cpos(2) = ClavePosDeuda
cpos(3) = ClavePosDeuda
cpos(4) = ClavePosDeuda

hreg(1) = horareg
hreg(2) = "000000"
hreg(3) = "000000"
hreg(4) = "000000"

txtcadena = "DELETE FROM " & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtcadena
For i = 1 To 4
    txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de la posicion
    txtcadena = txtcadena & "'" & txtport & "',"                  'el nombre del portafolio
    txtcadena = txtcadena & "'" & tipopos & "',"                  'tipo de posicion
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de registro
    txtcadena = txtcadena & "'" & txtnompos & "',"                'nombre de la posicion
    txtcadena = txtcadena & "'" & hreg(i) & "',"                  'la hora de registro
    txtcadena = txtcadena & cpos(i) & ","                         'la clave de posicion
    txtcadena = txtcadena & "'" & txtoper(i) & "')"               'la clave de OPERACION
    ConAdo.Execute txtcadena
Next i

End Sub

Sub DetSubportEfecCobPSwap(ByVal fecha As Date, ByVal txtport As String, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByVal coperacion As String)
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtoper(1 To 4) As String
Dim cpos(1 To 2) As Integer
Dim hreg(1 To 2) As String
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtoper(1) = coperacion
txtoper(2) = "PS " & coperacion
cpos(1) = ClavePosDeriv
cpos(2) = ClavePosDeuda
hreg(1) = horareg
hreg(2) = "000000"

txtcadena = "DELETE FROM " & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtcadena
For i = 1 To 2
    txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de la posicion
    txtcadena = txtcadena & "'" & txtport & "',"                  'el nombre del portafolio
    txtcadena = txtcadena & "'" & tipopos & "',"                  'tipo de posicion
    txtcadena = txtcadena & txtfecha & ","                        'la fecha de registro
    txtcadena = txtcadena & "'" & txtnompos & "',"                'nombre de la posicion
    txtcadena = txtcadena & "'" & hreg(i) & "',"                  'la hora de registro
    txtcadena = txtcadena & cpos(i) & ","                         'la clave de posicion
    txtcadena = txtcadena & "'" & txtoper(i) & "')"               'la clave de OPERACION
    ConAdo.Execute txtcadena
Next i

End Sub



Private Sub Command2_Click()
Dim noreg As Integer
Dim i As Integer
Dim coperacion As String
Dim txtport As String
Dim mata() As Variant
Dim fecha As Date
Dim fecha0 As Date
Dim txtfecha As String
Dim txtfechar As String
Dim txtcadena As String
Dim txtborra As String
Dim txtnompos As String
Dim id_tabla As Integer
id_tabla = 3
If IsDate(Text1.Text) Then
   Command2.Enabled = False
   fecha = CDate(Text1.Text)
   mata = ObtOperGRid(fecha, MSFlexGrid1)
   noreg = UBound(mata, 1)
If noreg <> 0 Then
Screen.MousePointer = 11
   fecha0 = PBD1(fecha, 1, "MX")
   txtport = "CONSOLIDADO ID"
   txtnompos = "Intradia " & Format(fecha, "dd/mm/yyyy")
   Call GenProcCVaRID(fecha, txtport, 52, id_tabla)
   Call DeterminaPortCons(fecha, fecha0, txtport)
   Call DeterminaPortMD(fecha, fecha0, "MERCADO DE DINERO ID", "N")
   Call DeterminaPortMC(fecha, fecha0, "MESA DE CAMBIOS ID", "N")
   Call DeterminaPortDer2(fecha, fecha0, "DERIVADOS DE NEGOCIACION ID", "N", "N", "N")
   Call DeterminaPortDer2(fecha, fecha0, "DERIVADOS ESTRUCTURALES ID", "N", "S", "N")
   Call DeterminaPortFwds2(fecha, fecha0, "DERIVADOS NEGOCIACION RECLASIFICACION ID", "N", "N", "S")
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(mata, 1)
       txtfechar = "to_date('" & Format(mata(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'fecha
       txtcadena = txtcadena & "'" & txtport & "',"                     'PORTAFOLIO
       txtcadena = txtcadena & mata(i, 5) & ","                         'tipo de posición
       txtcadena = txtcadena & txtfechar & ","                          'fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'Nombre de la posicion
       txtcadena = txtcadena & "'" & mata(i, 7) & "',"                  'hora de registro
       txtcadena = txtcadena & mata(i, 1) & ","                         'clave de posicion
       txtcadena = txtcadena & mata(i, 2) & ")"                         'clave de operacion
       ConAdo.Execute txtcadena
       If mata(i, 6) = "N" And mata(i, 3) <> "R" Then
          txtfechar = "to_date('" & Format(mata(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
          txtcadena = txtcadena & txtfecha & ","                           'fecha
          txtcadena = txtcadena & "'DERIVADOS DE NEGOCIACION ID',"         'PORTAFOLIO
          txtcadena = txtcadena & mata(i, 5) & ","                         'tipo de posición
          txtcadena = txtcadena & txtfechar & ","                          'fecha de registro
          txtcadena = txtcadena & "'" & txtnompos & "',"                   'Nombre de la posicion
          txtcadena = txtcadena & "'" & mata(i, 7) & "',"                  'hora de registro
          txtcadena = txtcadena & mata(i, 1) & ","                         'clave de posicion
          txtcadena = txtcadena & mata(i, 2) & ")"                         'clave de operacion
          ConAdo.Execute txtcadena
       ElseIf mata(i, 6) = "S" And mata(i, 3) <> "R" Then
          txtfechar = "to_date('" & Format(mata(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
          txtcadena = txtcadena & txtfecha & ","                           'fecha
          txtcadena = txtcadena & "'DERIVADOS ESTRUCTURALES ID',"          'PORTAFOLIO
          txtcadena = txtcadena & mata(i, 5) & ","                         'tipo de posición
          txtcadena = txtcadena & txtfechar & ","                          'fecha de registro
          txtcadena = txtcadena & "'" & txtnompos & "',"                   'Nombre de la posicion
          txtcadena = txtcadena & "'" & mata(i, 7) & "',"                  'hora de registro
          txtcadena = txtcadena & mata(i, 1) & ","                         'clave de posicion
          txtcadena = txtcadena & mata(i, 2) & ")"                         'clave de operacion
          ConAdo.Execute txtcadena
       End If
       If mata(i, 3) = "R" Then
          txtfechar = "to_date('" & Format(mata(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
          txtcadena = txtcadena & txtfecha & ","                           'fecha
          txtcadena = txtcadena & "'DERIVADOS NEGOCIACION RECLASIFICACION ID',"  'PORTAFOLIO
          txtcadena = txtcadena & mata(i, 5) & ","                         'tipo de posición
          txtcadena = txtcadena & txtfechar & ","                          'fecha de registro
          txtcadena = txtcadena & "'" & txtnompos & "',"                   'Nombre de la posicion
          txtcadena = txtcadena & "'" & mata(i, 7) & "',"                  'hora de registro
          txtcadena = txtcadena & mata(i, 1) & ","                         'clave de posicion
          txtcadena = txtcadena & mata(i, 2) & ")"                         'clave de operacion
          ConAdo.Execute txtcadena
       End If
   Next i
   MsgBox "Se creo el subproceso de CVaR"
   Screen.MousePointer = 0
Else
   MsgBox "No se seleccionaron operaciones"
End If
Command2.Enabled = True
End If

End Sub

Sub GenProcCVaRID2(ByVal fecha As Date, ByVal txtport As String, ByVal noesc As Long, ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean, ByVal id_tabla As Integer)
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim contar As Long
txtfiltro2 = "SELECT * FROM "
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtportfr As String
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)

If noreg <> 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, 61, contar, "Calc CVaR ID Oper", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, noesc, htiempo, "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   exito = False
End If





End Sub

Function ObtOperGRid(ByVal fecha As Date, ByRef msfl As MSFlexGrid) As Variant()
Dim contar As Integer
Dim noreg As Integer
Dim i As Integer

contar = 0
noreg = msfl.Rows - 1
If noreg <> 0 Then
  ReDim mata(1 To 7, 1 To 1) As Variant
  For i = 1 To noreg
    If msfl.TextMatrix(i, 8) = "S" Then
       contar = contar + 1
       ReDim Preserve mata(1 To 7, 1 To contar) As Variant
       mata(1, contar) = ClavePosDeriv          'clave de posicion
       mata(2, contar) = msfl.TextMatrix(i, 1)  'clave de operacion
       mata(3, contar) = msfl.TextMatrix(i, 2)  'intencion
       mata(4, contar) = fecha                  'fecha reg
       mata(5, contar) = 3                      'tipo pos
       mata(6, contar) = msfl.TextMatrix(i, 6)  'reclasificacion
       mata(7, contar) = msfl.TextMatrix(i, 5)  'hora de registro
    End If
  Next i
  mata = MTranV(mata)
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
If contar = 0 Then
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtOperGRid = mata
End Function

Sub GenProcCVaRID(ByVal fecha As Date, ByVal txtport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim contar As Long
Dim txttabla As String

txttabla = DetermTablaSubproc(id_tabla)

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtfecha1 = Format(fecha, "dd/mm/yyyy")
    txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Cálculo CVaR ID", txtfecha1, txtport, "Simulación", 500, 1, 0.97, "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena

End Sub


Private Sub Command3_Click()
Dim fecha As Date
Dim txtpossim As String
Dim txtfecha As String
Dim nrn As Integer
Dim nrc As Integer
If IsDate(Text1.Text) Then
   Command3.Enabled = False
   fecha = CDate(Text1.Text)
   txtpossim = "Intradia " & Format(fecha, "DD/MM/YYYY")
   Screen.MousePointer = 11
       Call ImportarPosID(fecha, txtpossim, nrn, nrc)
      MsgBox "Se importo la información de " & nrn + nrc & " operaciones"
   Command3.Enabled = True
   Screen.MousePointer = 0
End If
End Sub

Private Sub Command4_Click()
Dim exito As Boolean
Dim noreg As Long
Dim fecha As Date
Dim txtnompos As String
Dim txtnomarch As String
Dim sihayarch As Boolean
Dim txtmsg As String
If Option1.value Or Option2.value Or Option3.value Or Option5.value Then
If IsDate(Text1.Text) Then
   fecha = CDate(Text1.Text)
   txtnompos = "Intradia " & Format(fecha, "dd/mm/yyyy")
   Screen.MousePointer = 11
   If Option1.value Or Option2.value Or Option3.value Then
      frmValidaOper.CommonDialog1.DialogTitle = "Abrir Posicion primaria de swap"
      frmValidaOper.CommonDialog1.FileName = txtnomarch
      frmValidaOper.CommonDialog1.ShowOpen
      txtnomarch = frmValidaOper.CommonDialog1.FileName
      frmProgreso.Show
      Call ImpPosPrimArc(fecha, txtnompos, txtnomarch, 3, noreg, txtmsg, exito)
      Unload frmProgreso
   ElseIf Option5.value Then
      CommonDialog1.ShowOpen
      txtnomarch = CommonDialog1.FileName
      sihayarch = VerifAccesoArch(txtnomarch)
      If sihayarch Then
         frmProgreso.Show
         Call CrearPosSwapsSimArch(fecha, txtnomarch, 3, txtnompos, "000000", noreg)
         MsgBox "Se cargo " & noreg & " operaciones primarias"
         Unload frmProgreso
      End If
   End If
   If exito Then
    'Call ValidarTOpPrim(fecha)
   End If
   Screen.MousePointer = 0
End If
End If
End Sub

Sub ValidarTOpPrim(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim indice As Integer
Dim txttoper As String
Dim coperacion As String
Dim txtactualiza As String
Dim fvalua As String
Dim tipopos As Integer
Dim rmesa As New ADODB.recordset

tipopos = 3

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPosDeuda & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS ='" & tipopos & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("COPERACION") 'clave de operacion
       mata(i, 2) = rmesa.Fields("TOPERACION") 'tipo de operacion
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       coperacion = DetCoper(mata(i, 1))
       If Not EsVariableVacia(coperacion) Then
          fvalua = DetFValxSwapAsociado(tipopos, fecha, coperacion, "000000", mata(i, 2))
          If Not EsVariableVacia(fvalua) Then
             txtactualiza = "UPDATE " & TablaPosDeuda & " SET CPRODUCTO = '" & fvalua & "' "
             txtactualiza = txtactualiza & " WHERE COPERACION = '" & mata(i, 1) & "'"
             ConAdo.Execute txtactualiza
          Else
             MsgBox "No se determino el tipo producto de la operacion " & mata(i, 1)
          End If
       Else
          MsgBox "No hay un swap asociado a la operacion " & mata(i, 1)
       End If
   Next i
End If
End Sub


Private Sub Form_Load()
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim matd() As Variant
Dim mats1() As Variant
Dim mats2() As Variant
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim fecha As Date
fecha = Date
  If OpcionBDatos = 1 Then
       frmValidaOper.Caption = "Validación de operaciones intradia (Producción)"
  ElseIf OpcionBDatos = 2 Then
       frmValidaOper.Caption = "Validación de operaciones intradia (Desarrollo)"
  ElseIf OpcionBDatos = 3 Then
       frmValidaOper.Caption = "Validación de operaciones intradia (DRP)"
  End If

Call IniciarConexOracle(conAdo2, BDIKOS)
mata = LeerInterfSwapsIKOS2(fecha, "N", conAdo2)
matb = ImpFwdSimIkos2(fecha, "N", "R", conAdo2)
matc = LeerInterfSwapsIKOS2(fecha, "C", conAdo2)
matd = ImpFwdSimIkos2(fecha, "C", "C", conAdo2)
mats1 = UnirTablas(mata, matb, 1)
mats2 = UnirTablas(matc, matd, 1)
conAdo2.Close
If UBound(mats1, 1) <> 0 Then
   noreg1 = UBound(mats1, 1)
   MSFlexGrid1.Rows = noreg1 + 1
   MSFlexGrid1.Cols = 9
   MSFlexGrid1.TextMatrix(0, 1) = "Clave de operacion"
   MSFlexGrid1.TextMatrix(0, 2) = "Intencion"
   MSFlexGrid1.TextMatrix(0, 3) = "Clave de operacion"
   MSFlexGrid1.TextMatrix(0, 4) = "Estado"
   MSFlexGrid1.TextMatrix(0, 5) = "Hora de registro en IKOS"
   MSFlexGrid1.TextMatrix(0, 6) = "Estructural"
   MSFlexGrid1.TextMatrix(0, 7) = "Tipo de operacion"
   MSFlexGrid1.TextMatrix(0, 8) = "Validar"
   For j = 1 To 8
       MSFlexGrid1.ColWidth(j) = 1800
   Next j
   For i = 1 To noreg1
   For j = 1 To 7
   MSFlexGrid1.TextMatrix(i, j) = mats1(i, j)
   Next j
   Next i

End If
If UBound(mats2, 1) <> 0 Then
   noreg2 = UBound(mats2, 1)
   MSFlexGrid2.Rows = noreg2 + 1
   MSFlexGrid2.Cols = 9
   MSFlexGrid2.TextMatrix(0, 1) = "Clave de operacion"
   MSFlexGrid2.TextMatrix(0, 2) = "Intencion"
   MSFlexGrid2.TextMatrix(0, 3) = "Clave de de operacion"
   MSFlexGrid2.TextMatrix(0, 4) = "Estado"
   MSFlexGrid2.TextMatrix(0, 5) = "Hora de registro en IKOS"
   MSFlexGrid2.TextMatrix(0, 6) = "Estructural"
   MSFlexGrid2.TextMatrix(0, 7) = "Tipo de operacion"
   MSFlexGrid2.TextMatrix(0, 8) = "Validar"
   For j = 1 To 7
       MSFlexGrid2.ColWidth(j) = 1800
   Next j
   For i = 1 To noreg2
   For j = 1 To 7
   MSFlexGrid2.TextMatrix(i, j) = mats2(i, j)
   Next j
   Next i
End If
Text1.Text = fecha
End Sub


Private Sub MSFlexGrid1_DblClick()
Dim indice1 As Integer
Dim indice2 As Integer

indice1 = MSFlexGrid1.col
indice2 = MSFlexGrid1.row
If indice1 = 8 Then
   If MSFlexGrid1.TextMatrix(indice2, indice1) = "S" Then
      MSFlexGrid1.TextMatrix(indice2, indice1) = "N"
   Else
      MSFlexGrid1.TextMatrix(indice2, indice1) = "S"
   End If
End If
End Sub


Private Sub MSFlexGrid2_DblClick()
Dim indice1 As Integer
Dim indice2 As Integer
indice1 = MSFlexGrid2.col
indice2 = MSFlexGrid2.row
If indice1 = 8 Then
   If MSFlexGrid2.TextMatrix(indice2, indice1) = "S" Then
      MSFlexGrid2.TextMatrix(indice2, indice1) = "N"
   Else
      MSFlexGrid2.TextMatrix(indice2, indice1) = "S"
   End If
End If

End Sub
