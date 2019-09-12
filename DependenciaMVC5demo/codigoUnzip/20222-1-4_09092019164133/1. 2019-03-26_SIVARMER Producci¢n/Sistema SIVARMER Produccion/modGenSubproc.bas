Attribute VB_Name = "modGenSubproc"
Option Explicit

Sub GenSubProcValOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim i As Long
Dim noreg As Long
Dim contar As Long
Dim txtborra As String
Dim txtcadena As String
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim cposicion As Long
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim coperacion As String
Dim tipopos As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND ID_VALUACION = " & id_val
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
txtborra = txtborra & " AND PARAMETRO9 = '" & id_val & "'"
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
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
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de operación", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, id_val, "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
       DoEvents
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   exito = False
End If
End Sub

Sub GenSubpValPosicion(ByVal id_proc As Integer, ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_val As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txttabla As String
txttabla = DetermTablaSubproc(id_tabla)

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Valuación de posicion", txtport, txtportfr, id_val, "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
   
End Sub

