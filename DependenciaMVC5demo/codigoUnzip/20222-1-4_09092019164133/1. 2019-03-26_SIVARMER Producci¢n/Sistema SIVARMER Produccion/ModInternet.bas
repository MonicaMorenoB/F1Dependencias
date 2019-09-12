Attribute VB_Name = "Internet"
'En este modulo se cargan todas las funciones y
'rutinas para conectarse a internet
Global siacarchftp As Boolean
Global txtmsgFTP As String

Sub DescargarFTP2(ByVal txtnomarch1 As String, ByVal txtnomarch2 As String, ByRef txtmsg As String, ByRef exito As Boolean)
On Error GoTo hayerror
If Len(Trim(txtnomarch1)) <> 0 And Len(Trim(txtnomarch2)) <> 0 Then
frmEjecucionProc.Inet1.Protocol = icFTP
frmEjecucionProc.Inet1.URL = NomSRVPIP
frmEjecucionProc.Inet1.UserName = usersftpPIP
frmEjecucionProc.Inet1.Password = passsftpPIP
Call frmEjecucionProc.Inet1.Execute(, "cd definitivo")
Do While frmEjecucionProc.Inet1.StillExecuting
  DoEvents
Loop
siacarchftp = True
Call frmEjecucionProc.Inet1.Execute(, "get " & txtnomarch1 & " " & txtnomarch2)
Do While frmEjecucionProc.Inet1.StillExecuting
  DoEvents
Loop
If siacarchftp Then
   exito = True
Else
  exito = False
End If
txtmsg = txtmsgFTP
Else
 exito = False
End If
Exit Sub
hayerror:
  exito = False
  txtmsg = error(Err())
End Sub

