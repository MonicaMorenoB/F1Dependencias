Attribute VB_Name = "modPortPensiones"
Option Explicit

Sub GenPortPensiones(ByVal fecha As Date)
Dim txtfecha As String
Dim txtborra As String
Dim tipopos As Integer
tipopos = 1

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
txtborra = txtborra & " AND TIPOPOS = " & tipopos
txtborra = txtborra & " AND CPOSICION = " & ClavePosPension1
ConAdo.Execute txtborra

Call DetermPosPension(fecha, "FID 2065", tipopos, ClavePosPension1)
Call DetermPosPensionAdmin(fecha, "BANAMEX 2065", ClavePosPension1, tipopos, "BANAMEX")
Call DetermPosPensionAdmin(fecha, "BANOBRAS 2065", ClavePosPension1, tipopos, "BANOBRAS")
Call DetermPosPensionAdmin(fecha, "BANORTE 2065", ClavePosPension1, tipopos, "BANORTE")
Call DetermPosPensionAdmin(fecha, "EVERCORE 2065", ClavePosPension1, 1, "EVERCORE")
Call DetermPosPensionAdmin(fecha, "GBM 2065", ClavePosPension1, 1, "GBM")
Call DetermPosPensionAdmin(fecha, "SANTANDER 2065", ClavePosPension1, 1, "SANTANDER")
Call DetermPosPensionAdmin(fecha, "VECTOR 2065", ClavePosPension1, 1, "VECTOR")

Call DetermPosPensionContrato(fecha, ClavePosPension1, 1, "983")
Call DetermPosPensionContrato(fecha, ClavePosPension1, 1, "984")
Call DetermPosPensionContrato(fecha, ClavePosPension1, 1, "985")
Call DetermPosPensionContrato(fecha, ClavePosPension1, 1, "986")
Call DetermPosPensionContrato(fecha, ClavePosPension1, 1, "987")

Call DetermPosPensionTPapel(fecha, "Capitales y SI 2065", "CAPITALES", 1, ClavePosPension1, 1)
Call DetermPosPensionTPapel(fecha, "Papel Guber 2065", "GUBERNAMENTAL", 1, ClavePosPension1, 1)
Call DetermPosPensionTPapel(fecha, "Papel Privado 2065", "PAPEL PRIVADO", 1, ClavePosPension1, 1)
Call DetermPosPensionTPapel(fecha, "Reporto PG 2065", "REPORTO GUB", 1, ClavePosPension1, 2)
Call DetermPosPensionTPapel(fecha, "CBICS Y PIC 2065", "CBICS Y PIC", 1, ClavePosPension1, 1)
Call DetermPosPensionTPapel(fecha, "Reporto PP 2065", "REPORTO PRIVADO", 1, ClavePosPension1, 2)


Call DetermPosPensionTCalif(fecha, "AAA 2065", 1, ClavePosPension1, "AAA")
Call DetermPosPensionTCalif(fecha, "AA 2065", 1, ClavePosPension1, "AA")
Call DetermPosPensionTCalif(fecha, "A 2065", 1, ClavePosPension1, "A")
Call DetermPosPensionTCalif(fecha, "BBB+ 2065", 1, ClavePosPension1, "BBB+")
Call DetermPosPensionTCalif(fecha, "BBB 2065", 1, ClavePosPension1, "BBB")
Call DetermPosPensionTCalif(fecha, "Corto plazo 2065", 1, ClavePosPension1, "CP")
Call DetermPosPensionTCalif(fecha, "NA 2065", 1, ClavePosPension1, "ND")
Call DetermPosPensionTCalif(fecha, "Gobierno Fed 2065", 1, ClavePosPension1, "GF")


End Sub


Sub GenPortPensionesb(ByVal fecha As Date)
Dim txtfecha As String
Dim txtborra As String
Dim tipopos As Integer
tipopos = 1

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
txtborra = txtborra & " AND TIPOPOS = " & tipopos
txtborra = txtborra & " AND CPOSICION = " & ClavePosPension2
ConAdo.Execute txtborra


Call DetermPosPension(fecha, "FID 2160", 1, ClavePosPension2)

Call DetermPosPensionAdmin(fecha, "ACTINVER 2160", ClavePosPension2, 1, "ACTINVER")
Call DetermPosPensionAdmin(fecha, "BANAMEX 2160", ClavePosPension2, 1, "BANAMEX")
Call DetermPosPensionAdmin(fecha, "BANOBRAS 2160", ClavePosPension2, 1, "BANOBRAS")
Call DetermPosPensionAdmin(fecha, "GBM 2160", ClavePosPension2, 1, "GBM")
Call DetermPosPensionAdmin(fecha, "VECTOR 2160", ClavePosPension2, 1, "VECTOR")

Call DetermPosPensionContrato(fecha, ClavePosPension2, 1, "988")
Call DetermPosPensionContrato(fecha, ClavePosPension2, 1, "989")
Call DetermPosPensionContrato(fecha, ClavePosPension2, 1, "990")
Call DetermPosPensionContrato(fecha, ClavePosPension2, 1, "1111")

Call DetermPosPensionTPapel(fecha, "Capitales y SI 2160", "CAPITALES", 1, ClavePosPension2, 1)
Call DetermPosPensionTPapel(fecha, "Papel Guber 2160", "GUBERNAMENTAL", 1, ClavePosPension2, 1)
Call DetermPosPensionTPapel(fecha, "Papel Privado 2160", "PAPEL PRIVADO", 1, ClavePosPension2, 1)
Call DetermPosPensionTPapel(fecha, "Reporto PG 2160", "REPORTO GUB", 1, ClavePosPension2, 2)
Call DetermPosPensionTPapel(fecha, "CBICS Y PIC 2160", "CBICS Y PIC", 1, ClavePosPension2, 1)
Call DetermPosPensionTPapel(fecha, "Reporto PP 2160", "REPORTO PRIVADO", 1, ClavePosPension2, 2)

Call DetermPosPensionTCalif(fecha, "AAA 2160", 1, ClavePosPension2, "AAA")
Call DetermPosPensionTCalif(fecha, "AA 2160", 1, ClavePosPension2, "AA")
Call DetermPosPensionTCalif(fecha, "A 2160", 1, ClavePosPension2, "A")
Call DetermPosPensionTCalif(fecha, "BBB+ 2160", 1, ClavePosPension2, "BBB+")
Call DetermPosPensionTCalif(fecha, "BBB 2160", 1, ClavePosPension2, "BBB")
Call DetermPosPensionTCalif(fecha, "Corto plazo 2160", 1, ClavePosPension2, "CP")
Call DetermPosPensionTCalif(fecha, "NA 2160", 1, ClavePosPension2, "ND")
Call DetermPosPensionTCalif(fecha, "Gobierno Fed 2160", 1, ClavePosPension2, "GF")

End Sub





Sub GenPortPensiones2(ByVal fecha As Date, matport1, matport2)
Dim txtfecha As String
Dim txtborra As String
Dim tipopos As Integer
Dim i As Long
tipopos = 1

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
txtborra = txtborra & " AND TIPOPOS = " & tipopos
txtborra = txtborra & " AND (CPOSICION = " & ClavePosPension1 & " OR CPOSICION = " & ClavePosPension2 & " )"
ConAdo.Execute txtborra
Call DetermPosPension(fecha, "FID 2065", tipopos, ClavePosPension1)
For i = 1 To UBound(matport1, 1)
   Call DetermPosPensionTV(fecha, matport1(i, 1), ClavePosPension1, tipopos, matport1(i, 2), 1)
Next i
Call DetermPosPension(fecha, "FID 2160", tipopos, ClavePosPension2)
For i = 1 To UBound(matport2, 1)
   Call DetermPosPensionTV(fecha, matport2(i, 1), ClavePosPension2, tipopos, matport2(i, 2), 1)
Next i
End Sub


Sub DetermPosPensionTV(ByVal fecha As Date, ByVal txtport As String, ByVal id_pos As Integer, ByVal tipo_pos As Integer, ByVal tv As String, ByVal top As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
    
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TV = '" & tv & "'"
txtfiltro2 = txtfiltro2 & " AND TOPERACION = " & top
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TV = '" & tv & "'"
txtfiltro2 = txtfiltro2 & " AND TOPERACION = " & top

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub



Sub DetermPosPensionAdmin(ByVal fecha As Date, ByVal txtport As String, ByVal id_pos As Integer, ByVal tipo_pos As Integer, ByVal txtadmin As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
    
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND SUBPORT2 = '" & txtadmin & "'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND SUBPORT2 = '" & txtadmin & "'"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub

Sub DetermPosPensionContrato(ByVal fecha As Date, ByVal id_pos As Integer, ByVal tipo_pos As Integer, ByVal txtcontrato As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND SUBPORT_1 = '" & txtcontrato & "'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND SUBPORT1 = '" & txtcontrato & "'"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close

If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtcontrato & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub

Sub DetermPosPensionTCalif(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer, ByVal calif As String)
'capitales y sociedades de inversion
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"

txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD
txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND CALIFICACION = '" & calif & "'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv
txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND CALIFICACION = '" & calif & "'"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DetermPosPensionTPapel2(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer)
'PAPEL guber
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 1"
txtfiltro2 = txtfiltro2 & " AND (TV = 'BI'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IM'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IQ'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IS'"
txtfiltro2 = txtfiltro2 & " OR TV = 'LD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'M'"
txtfiltro2 = txtfiltro2 & " OR TV = 'S'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CDVITOT'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CEDEVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFE'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFEHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CIENCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GCDMXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='MICHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXACA'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PAMMCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PEMEX'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVICB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='VERACB')"

txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 1"
txtfiltro2 = txtfiltro2 & " AND (TV = 'BI'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IM'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IQ'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IS'"
txtfiltro2 = txtfiltro2 & " OR TV = 'LD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'M'"
txtfiltro2 = txtfiltro2 & " OR TV = 'S'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CDVITOT'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CEDEVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFE'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFEHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CIENCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GCDMXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='MICHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXACA'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PAMMCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PEMEX'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVICB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVIS'"

txtfiltro2 = txtfiltro2 & " OR EMISION ='VERACB')"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DetermPosPensionTPapel3(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer)
'papel privado
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 1"
txtfiltro2 = txtfiltro2 & " AND (TV = '90'"
txtfiltro2 = txtfiltro2 & " OR TV = 'I'"
txtfiltro2 = txtfiltro2 & " OR TV = 'D2'"
txtfiltro2 = txtfiltro2 & " OR TV = 'D8'"
txtfiltro2 = txtfiltro2 & " OR TV = 'CD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'JE'"
txtfiltro2 = txtfiltro2 & " OR TV = 'JI'"
txtfiltro2 = txtfiltro2 & " OR TV = '91'"
txtfiltro2 = txtfiltro2 & " OR TV = '93'"
txtfiltro2 = txtfiltro2 & " OR TV = '94'"
txtfiltro2 = txtfiltro2 & " OR TV = '95')"
txtfiltro2 = txtfiltro2 & " AND (EMISION <> 'CDVITOT'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CEDEVIS'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFE'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFECB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFEHCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CIENCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GCDMXCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GDFCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GDFECB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'MICHCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'OAXACA'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'OAXCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'PAMMCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'PEMEX'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'TFOVICB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <>'TFOVIS'"
txtfiltro2 = txtfiltro2 & " AND EMISION <>'VERACB')"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND (TV = '90'"
txtfiltro2 = txtfiltro2 & " OR TV = 'I'"
txtfiltro2 = txtfiltro2 & " OR TV = 'D2'"
txtfiltro2 = txtfiltro2 & " OR TV = 'D8'"
txtfiltro2 = txtfiltro2 & " OR TV = 'CD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'JE'"
txtfiltro2 = txtfiltro2 & " OR TV = 'JI'"
txtfiltro2 = txtfiltro2 & " OR TV = '91'"
txtfiltro2 = txtfiltro2 & " OR TV = '93'"
txtfiltro2 = txtfiltro2 & " OR TV = '94'"
txtfiltro2 = txtfiltro2 & " OR TV = '95')"
txtfiltro2 = txtfiltro2 & " AND (EMISION <> 'CDVITOT'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CEDEVIS'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFE'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFECB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CFEHCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'CIENCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GCDMXCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GDFCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'GDFECB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'MICHCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'OAXACA'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'OAXCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'PAMMCB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'PEMEX'"
txtfiltro2 = txtfiltro2 & " AND EMISION <> 'TFOVICB'"
txtfiltro2 = txtfiltro2 & " AND EMISION <>'TFOVIS'"
txtfiltro2 = txtfiltro2 & " AND EMISION <>'VERACB')"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DetermPosPensionTPapel4(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer)
'reporto papel guber
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 2"
txtfiltro2 = txtfiltro2 & " AND (TV = 'LD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'M'"
txtfiltro2 = txtfiltro2 & " OR TV = 'S'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CDVITOT'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CEDEVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFE'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFEHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CIENCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GCDMXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='MICHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXACA'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PAMMCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PEMEX'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVICB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='VERACB')"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 2"
txtfiltro2 = txtfiltro2 & " AND (TV = 'LD'"
txtfiltro2 = txtfiltro2 & " OR TV = 'M'"
txtfiltro2 = txtfiltro2 & " OR TV = 'S'"
txtfiltro2 = txtfiltro2 & " OR TV = 'IS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CDVITOT'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CEDEVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFE'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CFEHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='CIENCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GCDMXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='GDFECB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='MICHCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXACA'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='OAXCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PAMMCB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='PEMEX'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVICB'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='TFOVIS'"
txtfiltro2 = txtfiltro2 & " OR EMISION ='VERACB')"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext

   Next i
   rmesa.Close
End If

End Sub

Sub DetermPosPensionTPapel(ByVal fecha As Date, ByVal txtnomport As String, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer, ByVal toper As Integer)
'reporto papel guber
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
txtfiltro2 = ConstrClaveSQL(fecha, txtport, tipo_pos, id_pos, toper)
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtnomport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext

   Next i
   rmesa.Close
End If

End Sub

Function ConstrClaveSQL(ByVal fecha As Date, txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer, ByVal toper As Integer)
Dim txtfecha As String
Dim txtcadena As String
Dim i As Integer

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtcadena = txtcadena & " AND TIPOPOS = " & tipo_pos
txtcadena = txtcadena & " AND CPOSICION = " & id_pos
txtcadena = txtcadena & " AND TOPERACION = " & toper
txtcadena = txtcadena & " AND ("
For i = 1 To UBound(MatGruposPapelFP, 1)
    If txtport = MatGruposPapelFP(i, 1) Then
       If MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND EMISION = '" & MatGruposPapelFP(i, 3) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND EMISION = '" & MatGruposPapelFP(i, 3) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " EMISION = '" & MatGruposPapelFP(i, 3) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " EMISION = '" & MatGruposPapelFP(i, 3) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       End If
       txtcadena = txtcadena & " OR "
    End If
Next i
txtcadena = Left(txtcadena, Len(txtcadena) - 4) & ") UNION "
txtcadena = txtcadena & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtcadena = txtcadena & " AND TIPOPOS = " & tipo_pos
txtcadena = txtcadena & " AND CPOSICION = " & id_pos
txtcadena = txtcadena & " AND TOPERACION = " & toper
txtcadena = txtcadena & " AND ("
For i = 1 To UBound(MatGruposPapelFP, 1)
    If txtport = MatGruposPapelFP(i, 1) Then
       If MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND EMISION = '" & MatGruposPapelFP(i, 3) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND EMISION = '" & MatGruposPapelFP(i, 3) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " EMISION = '" & MatGruposPapelFP(i, 3) & "' AND SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       ElseIf MatGruposPapelFP(i, 2) <> "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " TV = '" & MatGruposPapelFP(i, 2) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) <> "*" And MatGruposPapelFP(i, 4) = "*" Then
          txtcadena = txtcadena & " EMISION = '" & MatGruposPapelFP(i, 3) & "'"
       ElseIf MatGruposPapelFP(i, 2) = "*" And MatGruposPapelFP(i, 3) = "*" And MatGruposPapelFP(i, 4) <> "*" Then
          txtcadena = txtcadena & " SERIE = '" & MatGruposPapelFP(i, 4) & "'"
       End If
       txtcadena = txtcadena & " OR "
    End If
Next i
txtcadena = Left(txtcadena, Len(txtcadena) - 4) & ")"
ConstrClaveSQL = txtcadena
End Function



Sub DetermPosPensionTPapel6(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer)
'REPORTO PAPEL privado
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 2"
txtfiltro2 = txtfiltro2 & " AND (TV = 'I'"
txtfiltro2 = txtfiltro2 & " OR TV = '50'"
txtfiltro2 = txtfiltro2 & " OR TV = '90')"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " AND TOPERACION = 2"
txtfiltro2 = txtfiltro2 & " AND (TV = 'I'"
txtfiltro2 = txtfiltro2 & " OR TV = '50'"
txtfiltro2 = txtfiltro2 & " OR TV = '90')"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DetermPosPension(ByVal fecha As Date, ByVal txtport As String, ByVal tipo_pos As Integer, ByVal id_pos As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipo_pos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_pos
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & txtport & "',"
       txtcadena = txtcadena & tipopos & ","
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & "'" & txtnompos & "',"
       txtcadena = txtcadena & "'" & horareg & "',"
       txtcadena = txtcadena & cposicion & ","
       txtcadena = txtcadena & "'" & coperacion & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub

