Attribute VB_Name = "FuncionesSistema"
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
 hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
 lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
 lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
 ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
 ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
 lpStartupInfo As STARTUPINFO, lpProcessInformation As _
 PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
 (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
 (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Function ExecCmd(cmdline$)
Dim ret&
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
    Dim Proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    start.cb = Len(start)
    ret& = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, Proc)
    ret& = WaitForSingleObject(Proc.hProcess, INFINITE)
    Call GetExitCodeProcess(Proc.hProcess, ret&)
    Call CloseHandle(Proc.hThread)
    Call CloseHandle(Proc.hProcess)
    ExecCmd = ret&
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CalcValuacion(ByVal fecha As Date, _
ByRef matpos() As propPosRiesgo, _
ByRef matposmd() As propPosMD, _
ByRef matposdiv() As propPosDiv, _
ByRef matposswaps() As propPosSwaps, _
ByRef matposfwd() As propPosFwd, _
ByRef mflujosswp() As estFlujosDeuda, _
ByRef matposdeuda() As propPosDeuda, _
ByRef mflujosd() As estFlujosDeuda, _
ByRef mfriesgo1() As Double, _
ByRef curvacompl() As Variant, _
ByRef parval As ParamValPos, _
ByRef mrvalflujo() As resValFlujo, _
ByRef txtmsg As String, ByRef _
exito As Boolean)

Dim indice As Long
Dim txtmensaje As String
Dim txtmsg1 As String
Dim noreg As Long
Dim inicio As Long
Dim fin As Long
Dim contar0 As Long
Dim valor1 As Long
Dim valor2 As Long
Dim exito1 As Boolean
'esta es la definicion de parval
'funcion para valuar la posición
'aparte hay que cargar de manera sucesiva varios parametros de la
'posicion como los flujos de los swap o las mfriesgo1 para el calculo
'matpos          posicion que hay que valuar
'mflujosswp    flujos de las emisiones
'fecha           fecha de valuacion
'mfriesgo1       matriz con las mfriesgo1 a aplicar
'tprecio         1 precio limpio    0 precio sucio
'valunico        indica si se va a valuar el portafolio o solo la posicion indicada por contar0
If SiAgregarDatosFwd Then
   ReDim MatParamFwds(1 To 5, 1 To 1) As Variant
End If

If SiAnexarFlujosSwaps Then
   ReDim MatValFlujosD(1 To 1) As resValFlujoExt
End If

noreg = UBound(matpos, 1)
'resval1
'1 valuacion sucio
'2 mtm sucio
'3 parte pasiva
'4 i dev activa
'5 i dev pasiva

If noreg <> 0 Then
If parval.indpos <> 0 Then
   valor1 = parval.indpos
   valor2 = parval.indpos
Else
   valor1 = 1
   valor2 = noreg
End If
ReDim resval1(1 To noreg) As New resValIns
exito1 = True
For contar0 = valor1 To valor2          'comienza el bucle de calculo COMIENZA EL BUCLE
'estas instrucciones solo aplican para los instrumentos que tienen cupon
    resval1(contar0).mtm_sucio = 0
    Select Case matpos(contar0).fValuacion
    Case "VREPORTO"
         Call ProcReporto(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "ACCION"
         Call ProcAccion(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "TIPO CAMBIO"
         Call ProcTCambio(fecha, matposdiv, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "IPAB"   'BONDES y BONOS PARECIDOS
         Call ProcIPAB(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "IPAB Y"   'BONDES y BONOS PARECIDOS
         Call ProcIPABY(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "BONO TASA FIJA" 'BONO TASA FIJA
         Call ProcBonoTFija(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "BONO TASA FIJA Y" 'BONO TASA FIJA YIELD
         Call ProcBonoYTF(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "BONO TASA FIJA Y PCF" 'BONO TASA FIJA YIELD pc fijo
         Call ProcBonoYTPCF(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "BONO STC Y"       'bono con sobretasa en el cupon
         Call ProcBonoSTCupon(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "B C CERO"
         Call ProcBonoCCero(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "B C CERO Y"
         Call ProcBonoCCeroY(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "BONDESD"   'BondesD
         Call ProcBondesD(fecha, matposmd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "DEUDA"
         Call ProcValDeuda(fecha, matposdeuda, mflujosd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval, mrvalflujo)
    Case "SWAP"
         Call ProcValSwap(fecha, matposswaps, mflujosswp, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval, mrvalflujo)
    Case "FWD TC"
         Call ProcFwdTC(fecha, matposfwd, mfriesgo1, curvacompl, resval1, contar0, matpos(contar0).IndPosicion, parval)
    Case "", Null
         MensajeProc = "No se ha catalogado el titulo: Posicion-> " & matpos(contar0).C_Posicion & " Clave de operacion -> " & matpos(contar0).c_operacion & " . Modificar la tabla VALUACION"
    Case Else
         MensajeProc = "No se ha definido una función de valuación: " & matpos(contar0).C_Posicion & " " & matpos(contar0).c_operacion & " "
    End Select
If Not exito1 Then
   ReDim resval1(0 To 0) As New resValIns
   txtmsg = txtmsg & "," & MensajeProc
   CalcValuacion = resval1
   exito = False
   Exit Function
End If
   If parval.simostrarap Then
      AvanceProc = contar0 / (valor2 - valor1 + 1)
      MensajeProc = txtmensaje & " " & Format(AvanceProc, "##0.00 %")
   End If
DoEvents
Next contar0
exito = True
Else
 ReDim resval1(0 To 0) As New resValIns
 exito = False
End If
CalcValuacion = resval1
End Function

Sub ProcAccion(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Integer, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim durac As Double
Dim vdv01 As Double
Dim valacc() As Variant

    valacc = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)

    resval(contar0).pu_sucio = valacc(1, 1)     'PRECIO SUCIO
    resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    End If
    resval(contar0).pu_limpio = resval(contar0).pu_sucio
    resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    End If
    durac = 0: vdv01 = 0
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).duractiva = durac
       resval(contar0).dv01activa = vdv01
    Else
       resval(contar0).durpasiva = durac
       resval(contar0).dv01pasiva = vdv01
    End If



End Sub

Sub ProcBonoYTPCF(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Integer, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curvatasa() As propCurva
Dim matfl() As New estFlujosMD
Dim durac As Double
Dim vdv01 As Double
Dim fvenc As Date
Dim dxv As Integer
Dim yield As Double
Dim tcupon As Double
Dim pc As Double
Dim diasc As Integer
Dim tCambio As Double
Dim califica As Integer
Dim escala As String
Dim recupera As Double

'se valua con la curva
    califica = matpos(indice0).CalifLP
    escala = matpos(indice0).escala
    recupera = matpos(indice0).recupera
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
    matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
    yield = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
    tcupon = matpos(indice0).tCuponVigenteMD
    pc = matpos(indice0).pCuponMD
    If Not EsVariableVacia(matpos(indice0).tCambioMD) Then
       tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
    Else
       tCambio = 1
    End If
    resval(contar0).pu_sucio = tCambio * PBonoYieldPCFijo(fechaval, matfl, tcupon, yield, pc, califica, escala, recupera, 0)
    If parval.sicalcPE Then resval(contar0).p_esperada = tCambio * PBonoYieldPCFijo(fechaval, matfl, tcupon, yield, pc, califica, escala, recupera, 1) * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    End If
    resval(contar0).pu_limpio = resval(contar0).pu_sucio
    resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    End If
    If parval.sicalcdur Then
       durac = DurBonoY(fechaval, matfl, tcupon, pc, yield)
    End If
    If parval.sicalcdv01 Then
       vdv01 = DV01BonoY(fechaval, matfl, tcupon, pc, yield)
    End If
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).duractiva = durac
       resval(contar0).dv01activa = vdv01
    Else
       resval(contar0).durpasiva = durac
       resval(contar0).dv01pasiva = vdv01
    End If

End Sub

Sub ProcFwdTC(ByVal fechaval As Date, ByRef matpos() As propPosFwd, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Integer, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curvat1() As propCurva
Dim curvat2() As propCurva
Dim curvat3() As propCurva
Dim tCambio As Double
Dim resval1() As Variant
Dim pstrike As Double
Dim fvenc As Date
Dim dxv As Integer
Dim durac As Double
Dim vdv01 As Double

    pstrike = matpos(indice0).PAsignadoFwd
    fvenc = matpos(indice0).FVencFwd
    dxv = Maximo(fvenc - fechaval, 0)
    
    curvat1 = CrearCurva(fechaval, matpos(indice0).FRiesgo1Fwd, curvacompl, mfriesgo1, parval.siValExc)
    curvat2 = CrearCurva(fechaval, matpos(indice0).FRiesgo3Fwd, curvacompl, mfriesgo1, parval.siValExc)
    curvat3 = CrearCurva(fechaval, matpos(indice0).FRiesgo2Fwd, curvacompl, mfriesgo1, parval.siValExc)
    tCambio = ObtieneFRiesgo(matpos(indice0).TCambioFwd, mfriesgo1)
    resval(contar0).pu_sucio = ValFwdDiv(pstrike, dxv, parval.perfwd, curvat1, curvat2, curvat3, tCambio, matpos(indice0).TInterpol1Fwd, resval1)
       resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).MontoNocFwd * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).ps_activa = resval1(1) * matpos(indice0).MontoNocFwd
          resval(contar0).ps_pasiva = resval1(2) * matpos(indice0).MontoNocFwd
       Else
          resval(contar0).ps_activa = resval1(2) * matpos(indice0).MontoNocFwd
          resval(contar0).ps_pasiva = resval1(1) * matpos(indice0).MontoNocFwd
       End If
       resval(contar0).pu_limpio = resval(contar0).pu_sucio
       resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).MontoNocFwd * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).pl_activa = resval1(1) * matpos(indice0).MontoNocFwd
          resval(contar0).pl_pasiva = resval1(2) * matpos(indice0).MontoNocFwd
       Else
          resval(contar0).pl_activa = resval1(2) * matpos(indice0).MontoNocFwd
          resval(contar0).pl_pasiva = resval1(1) * matpos(indice0).MontoNocFwd
       End If
       durac = dxv: vdv01 = 0
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).duractiva = durac
          resval(contar0).dv01activa = vdv01
       Else
          resval(contar0).durpasiva = durac
           resval(contar0).dv01pasiva = vdv01
      End If
End Sub

Sub ProcTCambio(ByVal fechaval As Date, ByRef matpos() As propPosDiv, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Integer, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curva1() As New propCurva
Dim valacc As Double
Dim durac As Double
Dim vdv01 As Double
    
    valacc = ObtieneFRiesgo(matpos(indice0).TCambioDiv, mfriesgo1)

       resval(contar0).pu_sucio = valacc
       resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).MontoNocDiv * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).MontoNocDiv
       Else
          resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).MontoNocDiv
       End If
       resval(contar0).pu_limpio = resval(contar0).pu_sucio
       resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).MontoNocDiv * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).MontoNocDiv
       Else
          resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).MontoNocDiv
       End If
       durac = 0: vdv01 = 0
       If matpos(indice0).Signo_Op = 1 Then
           resval(contar0).duractiva = durac
           resval(contar0).dv01activa = vdv01
       Else
           resval(contar0).durpasiva = durac
           resval(contar0).dv01pasiva = vdv01
       End If

    
End Sub

Sub ProcBonoYTF(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curvatasa() As propCurva
Dim matfl() As New estFlujosMD
Dim tcupon As Double
Dim fvenc As Date
Dim dxv As Integer
Dim pc As Integer
Dim yield As Double
Dim durac As Double
Dim vdv01 As Double
Dim tCambio As Double
Dim califica As Integer
Dim escala As String
Dim recupera As Double

'se valua con la curva
    califica = matpos(indice0).CalifLP
    escala = matpos(indice0).escala
    recupera = matpos(indice0).recupera
    tcupon = matpos(indice0).tCuponVigenteMD
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    pc = matpos(indice0).pCuponMD
    matpos(indice0).dVencMD = dxv
    matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
    matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
    yield = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
    If Not EsVariableVacia(matpos(indice0).tCambioMD) Then
       tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
    Else
       tCambio = 1
    End If
    resval(contar0).pu_sucio = tCambio * PBonoYield(fechaval, matfl, tcupon, yield, pc, califica, escala, recupera, 0)
    If parval.sicalcPE Then resval(contar0).p_esperada = tCambio * PBonoYield(fechaval, matfl, tcupon, yield, pc, califica, escala, recupera, 1) * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    End If
    resval(contar0).pu_limpio = resval(contar0).pu_sucio
    resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).pl_activa = resval(contar0).ps_pasiva * matpos(indice0).noTitulosMD
    Else
       resval(contar0).pl_pasiva = resval(contar0).ps_pasiva * matpos(indice0).noTitulosMD
    End If
    If parval.sicalcdur Then
       durac = DurBonoY(fechaval, matfl, tcupon, pc, yield)
    End If
    If parval.sicalcdv01 Then vdv01 = DV01BonoY(fechaval, matfl, tcupon, pc, yield)
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).duractiva = durac
       resval(contar0).dv01activa = vdv01
    Else
       resval(contar0).durpasiva = durac
       resval(contar0).dv01pasiva = vdv01
    End If
 End Sub

Sub ProcBonoCCero(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curva1() As New propCurva
Dim valcam() As Variant
Dim vn As Double
Dim fvenc As Date
Dim dxv As Integer
Dim tCambio As Double
Dim stasa As Double
Dim durac As Double
Dim vdv01 As Double
Dim calif As Integer
Dim escala As String
Dim recupera As Double
    calif = matpos(indice0).CalifLP
    escala = matpos(indice0).escala
    recupera = matpos(indice0).recupera
    vn = matpos(indice0).vNominalMD
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVCuponMD = 0
    matpos(indice0).dVencMD = dxv
' se crea la curva con que se descuenta
     curva1 = CrearCurva(fechaval, matpos(indice0).fRiesgo1MD, curvacompl, mfriesgo1, parval.siValExc)
     If Len(Trim(matpos(indice0).tCambioMD)) <> 0 Then
        tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
     Else
        tCambio = 1
     End If
     resval(contar0).pu_sucio = tCambio * PBonoC0(fechaval, vn, curva1, stasa, dxv, parval.perfwd, matpos(indice0).tInterpol1MD, calif, escala, recupera, 0)
     If parval.sicalcPE Then resval(contar0).p_esperada = tCambio * PBonoC0(fechaval, vn, curva1, stasa, dxv, parval.perfwd, matpos(indice0).tInterpol1MD, calif, escala, recupera, 1) * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     End If
     resval(contar0).pu_limpio = resval(contar0).pu_sucio
     resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     End If
     If parval.sicalcdur Then durac = (dxv) / 360   'dias de vencimiento del titulo
     If parval.sicalcdv01 Then vdv01 = DV01BonoC0(vn, curva1, dxv, parval.perfwd, matpos(indice0).tInterpol1MD)
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).duractiva = durac
        resval(contar0).dv01activa = vdv01
     Else
        resval(contar0).durpasiva = durac
        resval(contar0).dv01pasiva = vdv01
     End If

End Sub

Sub ProcBonoCCeroY(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curva1() As New propCurva
Dim valcam() As Variant
Dim tprecio As Integer
Dim pfwd As Integer
Dim sicalcdur As Boolean
Dim sicalcdv01 As Boolean
Dim vn As Double
Dim fvenc As Date
Dim dxv As Integer
Dim tCambio As Double
Dim stasa As Double
Dim durac As Double
Dim vdv01 As Double
Dim yield As Double

    vn = matpos(indice0).vNominalMD
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVCuponMD = 0
    matpos(indice0).dVencMD = dxv
' se crea la curva con que se descuenta
     yield = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
     If Len(Trim(matpos(indice0).tCambioMD)) <> 0 Then
        tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
     Else
        tCambio = 1
     End If
     resval(contar0).pu_sucio = tCambio * PBonoC0Y(vn, yield, stasa, dxv)
     resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     End If
     resval(contar0).pu_limpio = resval(contar0).pu_sucio
     resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     End If
     If parval.sicalcdur Then durac = (dxv) / 360   'dias de vencimiento del titulo
     If parval.sicalcdv01 Then vdv01 = 0 'DV01BonoC0(vn, curva1, dxv, pfwd, matpos(indice0).tInterpol1MD)
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).duractiva = durac
        resval(contar0).dv01activa = vdv01
     Else
        resval(contar0).durpasiva = durac
        resval(contar0).dv01pasiva = vdv01
     End If
End Sub

Sub ProcReporto(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim curva1() As New propCurva
Dim fcompra As Date
Dim fvenc As Date
Dim pasignado As Double
Dim dvr1 As Integer
Dim dxv As Integer
Dim tpremio As Double
Dim tlarga As Double
Dim tcorta As Double
Dim tdesc As Double
Dim durac As Double
Dim vdv01 As Double
Dim tCambio As Double
    fcompra = matpos(indice0).fCompraMD
    fvenc = matpos(indice0).fVencMD
    pasignado = matpos(indice0).pAsignadoMD
    dvr1 = Maximo(fvenc - fcompra, 0)
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    matpos(indice0).dVCuponMD = 0
    If Len(Trim(matpos(indice0).tReporto)) <> 0 Then
       tpremio = matpos(indice0).tReporto
    Else
       tpremio = 0
    End If
    curva1 = CrearCurva(fechaval, matpos(indice0).fRiesgo1MD, curvacompl, mfriesgo1, parval.siValExc)
    tlarga = CalculaTasa(curva1, dxv + parval.perfwd, matpos(indice0).tInterpol1MD)
     If parval.perfwd <> 0 Then
        tcorta = CalculaTasa(curva1, parval.perfwd, matpos(indice0).tInterpol1MD)
     Else
        tcorta = 0
     End If
     If dxv <> 0 Then
        tdesc = ((1 + tlarga * (dxv + parval.perfwd) / 360) / (1 + tcorta * parval.perfwd / 360) - 1) * 360 / dxv
     Else
        tdesc = 0
     End If
     If Not EsVariableVacia(matpos(indice0).tCambioMD) Then
        tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
     Else
        tCambio = 1
     End If
     resval(contar0).pu_sucio = pasignado * (1 + tpremio * dvr1 / 360) / (1 + tdesc * dxv / 360) * tCambio
     resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
     End If
     resval(contar0).pu_limpio = resval(contar0).pu_sucio
     resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     Else
        resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
     End If
     If parval.sicalcdur Then durac = dxv / 360  'dias de vencimiento del titulo
     If parval.sicalcdv01 Then vdv01 = pasignado * (1 + tpremio * dvr1 / 360) * (1 / (1 + (tdesc + 0.0001) * dxv / 360) - 1 / (1 + tdesc * dxv / 360))
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).duractiva = durac
        resval(contar0).dv01activa = vdv01
     Else
        resval(contar0).durpasiva = durac
        resval(contar0).dv01pasiva = vdv01
     End If

End Sub

Sub ProcValDeuda(ByVal fechaval As Date, ByRef matpos() As propPosDeuda, ByRef mfldeuda() As estFlujosDeuda, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos, ByRef mrvalflujo() As resValFlujo)
Dim matval() As Double
Dim parval1 As New paramValFlujo
Dim matn(0 To 0) As propCurva
Dim curvatasa1() As propCurva
Dim curvatasa2() As propCurva
Dim tCambio1 As Double

Dim idevd As Double
Dim durac As Double
Dim vdv01 As Double

'se leen 4 curvas para la valuacion del swap
     If Len(Trim(matpos(indice0).FRiesgo1Deuda)) <> 0 Then
        curvatasa1 = CrearCurva(fechaval, matpos(indice0).FRiesgo1Deuda, curvacompl, mfriesgo1, parval.siValExc)
     Else
        curvatasa1 = matn
     End If
     If Len(Trim(matpos(indice0).FRiesgo2Deuda)) <> 0 Then
        curvatasa2 = CrearCurva(fechaval, matpos(indice0).FRiesgo2Deuda, curvacompl, mfriesgo1, parval.siValExc)
     Else
        curvatasa2 = matn
     End If
     If Len(Trim(matpos(indice0).TCambioDeuda)) <> 0 Then
        tCambio1 = ObtieneFRiesgo(matpos(indice0).TCambioDeuda, mfriesgo1)
     Else
        tCambio1 = 1
     End If
     Call DefinirParamValDeuda(matpos, indice0, parval.perfwd, parval.si_int_flujos, tCambio1, parval1)
     resval(contar0).pu_sucio = ValDeuda(fechaval, mfldeuda, curvatasa1, curvatasa2, parval1, mrvalflujo)
     resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).Signo_Op
     idevd = IDevDeuda(fechaval, mfldeuda, curvatasa1, curvatasa2, parval1)
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).ps_activa = resval(contar0).pu_sucio
        resval(contar0).ps_pasiva = 0
     Else
        resval(contar0).ps_activa = 0
        resval(contar0).ps_pasiva = resval(contar0).pu_sucio
     End If
     resval(contar0).pu_limpio = resval(contar0).pu_sucio - idevd
     resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).Signo_Op
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).pl_activa = resval(contar0).pu_limpio
     Else
        resval(contar0).pl_pasiva = resval(contar0).pu_limpio
     End If
     If matpos(indice0).Signo_Op = 1 Then
        resval(contar0).duractiva = durac
        resval(contar0).dv01activa = vdv01
     Else
        resval(contar0).durpasiva = durac
        resval(contar0).dv01pasiva = vdv01
     End If
     
End Sub

Sub ProcBonoSTCupon(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Integer, ByRef parval As ParamValPos)
Dim curva1() As New propCurva
Dim curva2() As New propCurva
Dim matfl() As New estFlujosMD
Dim fvenc As Date
Dim tcupon As Double
Dim pc As Integer
Dim stcupon As Double
Dim yield As Double
Dim tc0 As Double
Dim tCambio As Double
Dim durac As Double
Dim vdv01 As Double
Dim dxv As Long
Dim califica As Integer
Dim escala As String
Dim recupera As Double

    califica = matpos(indice0).CalifLP
    escala = matpos(indice0).escala
    recupera = matpos(indice0).recupera
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    pc = matpos(indice0).pCuponMD
    stcupon = matpos(indice0).sTCuponMD
    matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
    matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
    yield = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
    tcupon = ObtieneFRiesgo(matpos(indice0).fRiesgo2MD, mfriesgo1)
    tc0 = matpos(indice0).tCuponVigenteMD
    If Not EsVariableVacia(matpos(indice0).tCambioMD) Then
       tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
    Else
       tCambio = 1
    End If
    resval(contar0).pu_sucio = tCambio * PBonoStCY(fechaval, matfl, tc0, tcupon, stcupon, yield, pc, False, califica, escala, recupera, 0)
    If parval.sicalcPE Then resval(contar0).p_esperada = tCambio * PBonoStCY(fechaval, matfl, tc0, tcupon, stcupon, yield, pc, False, califica, escala, recupera, 1) * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    End If
    resval(contar0).pu_limpio = resval(contar0).pu_sucio - IDevBonoCurva(fechaval, matfl, tc0)
    resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op   'MTM LIMPIO
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    End If
    If parval.sicalcdur Then
       durac = DurBonoStCY(fechaval, matfl, tc0, tcupon, stcupon, yield, pc)
    End If
    If parval.sicalcdv01 Then
       vdv01 = DV01BonoStCY(fechaval, matfl, tc0, tcupon, stcupon, yield, pc)
    End If
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).duractiva = durac
       resval(contar0).dv01activa = vdv01
    Else
       resval(contar0).durpasiva = durac
       resval(contar0).dv01pasiva = vdv01
    End If
End Sub

Sub ProcBonoTFija(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef mprecios() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim matfl() As New estFlujosMD
Dim curvatasa() As propCurva
Dim fvenc As Date
Dim dxv As Integer
Dim tcupon As Double
Dim pc As Integer
Dim durac As Double
Dim vdv01 As Double
Dim tCambio As Double

    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    tcupon = matpos(indice0).tCuponVigenteMD
    pc = matpos(indice0).pCuponMD
       matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
       matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
       curvatasa = CrearCurva(fechaval, matpos(indice0).fRiesgo1MD, curvacompl, mfriesgo1, parval.siValExc)
       If Not EsVariableVacia(matpos(indice0).tCambioMD) Then
          tCambio = ObtieneFRiesgo(matpos(indice0).tCambioMD, mfriesgo1)
       Else
          tCambio = 1
       End If
       mprecios(contar0).pu_sucio = tCambio * PBonoCurva(fechaval, tcupon, pc, parval.perfwd, matfl, curvatasa, matpos(indice0).tInterpol1MD)
       mprecios(contar0).mtm_sucio = mprecios(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          mprecios(contar0).ps_activa = mprecios(contar0).pu_sucio * matpos(indice0).noTitulosMD
       Else
          mprecios(contar0).ps_pasiva = mprecios(contar0).pu_sucio * matpos(indice0).noTitulosMD
       End If
       mprecios(contar0).pu_limpio = mprecios(contar0).pu_sucio - IDevBonoCurva(fechaval, matfl, tcupon)
       mprecios(contar0).mtm_limpio = mprecios(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          mprecios(contar0).ps_activa = mprecios(contar0).pu_sucio * matpos(indice0).noTitulosMD
       Else
          mprecios(contar0).pl_pasiva = mprecios(contar0).pu_limpio * matpos(indice0).noTitulosMD
       End If
       If parval.sicalcdur Then durac = DurBonoCurva(fechaval, matfl, tcupon, pc, curvatasa, matpos(indice0).tInterpol1MD)
       If parval.sicalcdv01 Then vdv01 = DV01BonoC(fechaval, matfl, tcupon, pc, curvatasa, matpos(indice0).tInterpol1MD)
       If matpos(indice0).Signo_Op = 1 Then
          mprecios(contar0).duractiva = durac
          mprecios(contar0).dv01activa = vdv01
       Else
          mprecios(contar0).durpasiva = durac
          mprecios(contar0).dv01pasiva = vdv01
       End If

End Sub

Sub ProcBondesD(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim matfl() As New estFlujosMD
Dim curvast() As New propCurva
Dim fcompra As Date
Dim fvenc As Date
Dim dxv As Integer
Dim pc1 As Integer
Dim tr As Double
Dim intdev As Double
Dim durac As Double
Dim vdv01 As Double

    fcompra = matpos(indice0).fCompraMD
    fvenc = matpos(indice0).fVencMD
    dxv = Maximo(fvenc - fechaval, 0)
    pc1 = matpos(indice0).pCuponMD
    matpos(indice0).dVencMD = dxv
   'esta formula da los dias que le faltan al cupon vigente
   'If pc1 <> 0 Then matpos(indice0, CDuracAct) = dxv + pc1 + Int(-dxv / pc1) * pc1
    tr = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
    curvast = CrearCurva(fechaval, matpos(indice0).fRiesgo2MD, curvacompl, mfriesgo1, parval.siValExc)
    st = CalculaTasa(curvast, dxv, matpos(indice0).tInterpol2MD)
    intdev = Round(matpos(indice0).intDevengMD, 8)
    matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
    matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
    If parval.mVBondesD = 1 Then      'valuacion con redondeo
       resval(contar0).pu_sucio = PBondesDV1(fechaval, matfl, intdev, tr, st, pc1)
    ElseIf parval.mVBondesD = 2 Then  'valuacion sin redondeo
       resval(contar0).pu_sucio = PBondesDV2(fechaval, matfl, intdev, tr, st, pc1)
    End If
    resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    Else
       resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
    End If
    resval(contar0).pu_limpio = resval(contar0).pu_sucio - intdev
    resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
    If matpos(indice0).Signo_Op = 1 Then
       resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    Else
      resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
    End If
    If parval.sicalcdur And pc1 <> 0 Then durac = DurBondesD(fechaval, matfl, intdev, tr, st, pc1)
    If parval.sicalcdv01 Then vdv01 = DV01BondesD(fechaval, matfl, intdev, tr, st, pc1)
    If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).duractiva = durac
          resval(contar0).dv01activa = vdv01
       Else
          resval(contar0).durpasiva = durac
          resval(contar0).dv01pasiva = vdv01
       End If

End Sub

Sub ProcIPAB(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim matfl() As New estFlujosMD
Dim curva2() As New propCurva
Dim fvenc As Date
Dim dxv As Integer
Dim tcuponvig As Double
Dim pc1 As Integer
Dim pc As Integer
Dim yield As Double
Dim durac As Double
Dim vdv01 As Double
    
    fvenc = matpos(indice0).fVencMD
    matpos(indice0).dVencMD = dxv
    tcuponvig = matpos(indice0).tCuponVigenteMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    pc1 = matpos(indice0).pCuponMD

       matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
       matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
       yield = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
       curva2 = CrearCurva(fechaval, matpos(indice0).fRiesgo2MD, curvacompl, mfriesgo1, parval.siValExc)
       st = CalculaTasa(curva2, dxv, matpos(indice0).tInterpol2MD)
       'st = ObtieneFRiesgo(matpos(indice0).fRiesgo2MD, mfriesgo1,exitofr)
       resval(contar0).pu_sucio = PIPABV1(fechaval, matfl, tcuponvig, yield, st, pc1)
        resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
        If matpos(indice0).Signo_Op = 1 Then
           resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
        Else
           resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
        End If
        resval(contar0).pu_limpio = resval(contar0).pu_sucio - IDevIPAB(fechaval, matfl, tcuponvig, pc1)
        resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
        If matpos(indice0).Signo_Op = 1 Then
           resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
        Else
           resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
        End If
        If parval.sicalcdur And pc1 <> 0 Then durac = DurIPAB(fechaval, matfl, tcuponvig, yield, st, pc1)
        'If parval.sicalcdurM And pc1 <> 0 Then duracm = DurIPAB(fechaval, matfl, tcuponvig, yield, st, pc1)
        If parval.sicalcdv01 Then vdv01 = DV01IPAB(fechaval, matfl, tcuponvig, yield, st, pc1)
        If matpos(indice0).Signo_Op = 1 Then
           resval(contar0).duractiva = durac
           resval(contar0).dv01activa = vdv01
        Else
           resval(contar0).durpasiva = durac
           resval(contar0).dv01pasiva = vdv01
        End If
  
End Sub

Sub ProcIPABY(ByVal fechaval As Date, ByRef matpos() As propPosMD, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos)
Dim matfl() As New estFlujosMD
Dim curva1() As New propCurva

Dim fvenc As Date
Dim tcupon As Double
Dim dxv As Integer
Dim pc1 As Integer
Dim pc As Integer
Dim tasaref As Double
Dim yield As Double
Dim durac As Double
Dim vdv01 As Double

    fvenc = matpos(indice0).fVencMD
    tcupon = matpos(indice0).tCuponVigenteMD
    dxv = Maximo(fvenc - fechaval, 0)
    matpos(indice0).dVencMD = dxv
    pc1 = matpos(indice0).pCuponMD
       matfl = ObtenerFlujosMD(matpos(indice0).iFlujoMD, matpos(indice0).fFlujoMD, MatFlujosMD)
       matpos(indice0).dVCuponMD = CalcDVCBono(fechaval, matfl)
       tasaref = ObtieneFRiesgo(matpos(indice0).fRiesgo1MD, mfriesgo1)
       yield = ObtieneFRiesgo(matpos(indice0).fRiesgo2MD, mfriesgo1)
       resval(contar0).pu_sucio = PIPABYield(fechaval, matfl, tcupon, tasaref, yield, pc1)
       resval(contar0).mtm_sucio = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).ps_activa = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
       Else
          resval(contar0).ps_pasiva = resval(contar0).pu_sucio * matpos(indice0).noTitulosMD
       End If
       resval(contar0).pu_limpio = resval(contar0).pu_sucio - IDevIPABY(fechaval, matfl, tcupon, pc1)
       resval(contar0).mtm_limpio = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD * matpos(indice0).Signo_Op
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).pl_activa = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
       Else
          resval(contar0).pl_pasiva = resval(contar0).pu_limpio * matpos(indice0).noTitulosMD
       End If
       If parval.sicalcdur Then
          If pc1 <> 0 Then durac = DurIPABY(fechaval, matfl, tcupon, tasaref, yield, pc1)
       End If
       If parval.sicalcdv01 Then vdv01 = DV01IPABY(fechaval, matfl, tcupon, tasaref, yield, pc1)
       If matpos(indice0).Signo_Op = 1 Then
          resval(contar0).duractiva = durac
          resval(contar0).dv01activa = vdv01
       Else
          resval(contar0).durpasiva = durac
          resval(contar0).dv01pasiva = vdv01
       End If

End Sub

Sub DeterminaCalifRecupContrap(ByVal id_contrap As String)
Dim califica As Integer
Dim Sector As String
Dim vrecupera As Double
Dim mrecupera() As Variant
vrecupera = Recuperacion(califica, mrecupera, Sector)


End Sub

Sub ProcValSwap(ByVal fecha As Date, ByRef matpos() As propPosSwaps, _
                ByRef mflujosswaps() As estFlujosDeuda, ByRef mfriesgo1() As Double, ByRef curvacompl() As Variant, ByRef resval() As resValIns, ByVal contar0 As Long, ByVal indice0 As Long, ByRef parval As ParamValPos, ByRef mrvalflujo() As resValFlujo)
Dim matn(0 To 0) As propCurva
Dim parval1 As New paramValFlujo
Dim parval2 As New paramValFlujo
Dim curvatasa1() As propCurva
Dim curvatasa2() As propCurva
Dim curvatasa3() As propCurva
Dim curvatasa4() As propCurva
Dim matval() As Double
Dim matdv() As Variant
Dim dxv As Integer
Dim tCambio1 As Double
Dim tCambio2 As Double
Dim idevs() As Double
Dim duraswap() As Double
Dim duract As Double
Dim durpas As Double
Dim vdv01act As Double
Dim vdv01pas As Double
Dim califica As Integer
Dim recupera As Double

If parval.sicalcVE Then
   Call DeterminaCalifRecupContrap(matpos(indice0).ID_ContrapSwap)
End If

    If Len(Trim(matpos(indice0).FRiesgo1Swap)) <> 0 Then
       curvatasa1 = CrearCurva(fecha, matpos(indice0).FRiesgo1Swap, curvacompl, mfriesgo1, parval.siValExc)
    Else
       curvatasa1 = matn
    End If
     If Len(Trim(matpos(indice0).FRiesgo2Swap)) <> 0 Then
        curvatasa2 = CrearCurva(fecha, matpos(indice0).FRiesgo2Swap, curvacompl, mfriesgo1, parval.siValExc)
     Else
        curvatasa2 = matn
     End If
     If Len(Trim(matpos(indice0).FRiesgo3Swap)) <> 0 Then
        curvatasa3 = CrearCurva(fecha, matpos(indice0).FRiesgo3Swap, curvacompl, mfriesgo1, parval.siValExc)
     Else
        curvatasa3 = matn
     End If
     If Len(Trim(matpos(indice0).FRiesgo4Swap)) <> 0 Then
        curvatasa4 = CrearCurva(fecha, matpos(indice0).FRiesgo4Swap, curvacompl, mfriesgo1, parval.siValExc)
     Else
        curvatasa4 = matn
     End If
     If parval.siTCambio Then
        If Len(Trim(matpos(indice0).TCambio1Swap)) <> 0 Then
           tCambio1 = ObtieneFRiesgo(matpos(indice0).TCambio1Swap, mfriesgo1)
        Else
           tCambio1 = 1
        End If
        If Len(Trim(matpos(indice0).TCambio2Swap)) <> 0 Then
           tCambio2 = ObtieneFRiesgo(matpos(indice0).TCambio2Swap, mfriesgo1)
        Else
           tCambio2 = 1
        End If
     Else
        tCambio1 = 1
        tCambio2 = 1
     End If
     Call DefinirParamValSwap(matpos, indice0, parval.perfwd, parval.si_int_flujos, tCambio1, tCambio2, parval1, parval2)
     With resval(contar0)
          .pu_sucio = ValSwap(fecha, mflujosswaps, curvatasa1, curvatasa2, curvatasa3, curvatasa4, parval1, parval2, matval, mrvalflujo)
          .mtm_sucio = .pu_sucio * matpos(indice0).Signo_Op
          If matpos(indice0).Signo_Op = 1 Then
             .ps_activa = matval(1)      'valuacion pata activa
             .ps_pasiva = matval(2)      'valuacion pata pasiva
          Else
             .ps_activa = matval(2)      'valuacion pata activa
             .ps_pasiva = matval(1)      'valuacion pata pasiva
          End If
          '.valve = ValSwap(fecha, mflujosswaps, curvatasa1, curvatasa2, curvatasa3, curvatasa4, parval1, parval2, matval, mrvalflujo)
          '.valVEact = ValSwap(fecha, mflujosswaps, curvatasa1, curvatasa2, curvatasa3, curvatasa4, parval1, parval2, matval, mrvalflujo)
          If parval.siPLimpio Then
             idevs = IDevSwap(fecha, mflujosswaps, curvatasa1, curvatasa2, curvatasa3, curvatasa4, parval1, parval2)
             .pu_limpio = .pu_sucio - idevs(1) + idevs(2)
             .mtm_limpio = .pu_limpio * matpos(indice0).Signo_Op
             If matpos(indice0).Signo_Op = 1 Then
                .pl_activa = matval(1) - idevs(1)    'valuacion pata activa limpia
                .pl_pasiva = matval(2) - idevs(2)    'valuacion pata pasiva  limpia
             Else
                .pl_activa = matval(2) - idevs(2)    'valuacion pata activa
                .pl_pasiva = matval(1) - idevs(1)    'valuacion pata pasiva
             End If
          End If
          If parval.sicalcdur Then
             duraswap = CDurSwap(fecha, mflujosswaps, curvatasa1, curvatasa2, curvatasa3, curvatasa4, parval1, parval2)
             If matpos(indice0).Signo_Op = 1 Then
                If matval(1) <> 0 Then duract = duraswap(1)
                If matval(2) <> 0 Then durpas = duraswap(2)
             Else
                If matval(2) <> 0 Then duract = duraswap(2)
                If matval(1) <> 0 Then durpas = duraswap(1)
             End If
             If matpos(indice0).Signo_Op = 1 Then
                .duractiva = duract
                .durpasiva = durpas
                .dv01activa = vdv01act
                .dv01pasiva = vdv01pas
             Else
                .duractiva = durpas
                .durpasiva = duract
                .dv01activa = vdv01pas
                .dv01pasiva = vdv01act
             End If
          End If
     End With
      If SiAnexarFlujosSwaps Then Call AnexarFlujosSwaps2(MatValFlujosD, mrvalflujo, fecha, matpos(indice0).intencion, matpos(indice0).IntercIFSwap, matpos(indice0).IntercFFSwap, tCambio1, tCambio2, matpos(indice0).TCambio1Swap, matpos(indice0).TCambio2Swap)
End Sub

Function DetFactRPos(ByVal fechaval As Date, ByRef matpos() As propPosRiesgo, ByRef mflujosswaps() As Variant, ByRef mflujosdeuda() As Variant, ByRef parval As ParamValPos, ByVal indice As Long)
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd

Dim parval1 As New paramValFlujo
Dim parval2 As New paramValFlujo
Dim tprecio As Integer
Dim pfwd As Integer
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim inicio As Long
Dim fin As Long
Dim contar As Long
Dim contar0 As Long
Dim fechareg As Date
Dim vn As Double
Dim fcompra As Date
Dim fvenc As Date
Dim pcompra As Double
Dim tpremio As Double
Dim TOPERACION As Integer
Dim tcupon As Double
Dim pc As Integer
Dim pasignado As Double
Dim tint As Double
Dim st As Double
Dim dvr1 As Integer
Dim dxv As Integer
Dim nc As Integer
Dim matfl() As New estFlujosMD
Dim tCambio1 As Double
Dim tCambio2 As Double
Dim matn() As Variant
Dim noreg1 As Long
Dim matc() As Variant

'esta es la definicion de matpar



'funcion para valuar la posición
'aparte hay que cargar de manera sucesiva varios parametros de la
'posicion como los flujos de los swap o las mfriesgo1 para el calculo
'matpos         posicion que hay que valuar
'mflujosswaps     flujos de las emisiones
'fechaval       fechaval de valuacion
'tprecio        1 precio limpio    0 precio sucio
'valunico       indica si se va a valuar el portafolio o solo la posicion indicada por contar0

noreg = UBound(matpos, 1)
ReDim MatFRPos(1 To noreg, 1 To 2) As Variant
'MPRECIOS
'1 VALUACION
'2 parte activa
'3 parte pasiva
'4 i dev activa
'5 i dev pasiva
If noreg <> 0 Then

contar = 0
ReDim MatNodosFREx(1 To 4, 0 To contar)
For contar0 = 1 To noreg           'comienza el bucle de calculo COMIENZA EL BUCLE
    If UBound(MatNodosFREx, 2) <> 1 Then contar = UBound(MatNodosFREx, 2)
    fechareg = matpos(contar0).fechareg
    vn = matposmd(contar0).vNominalMD
    fcompra = matposmd(contar0).fCompraMD
    fvenc = matposmd(contar0).fVencMD
    pcompra = matposmd(contar0).pAsignadoMD
    tpremio = matposmd(contar0).tReporto
    TOPERACION = matposmd(contar0).Tipo_Mov
    tcupon = matposmd(contar0).tCuponVigenteMD
    pc = matposmd(contar0).pCuponMD
    pasignado = matposmd(contar0).pAsignadoMD
    tint = 0
    st = 0
    matposmd(contar0).dVencMD = 0
    dxv = Maximo(fvenc - fechaval, 0)
    dvr1 = Maximo(fvenc - fcompra, 0)
    matposmd(contar0).dVencMD = Maximo(fvenc - fechaval, 0)
'estas instrucciones solo aplican para los instrumentos que tienen cupon
    If pc <> 0 Then
       nc = -Int(-dxv / pc)
       matposmd(contar0).dVCuponMD = Maximo(dxv - pc * (nc - 1), 0)
    Else
       matposmd(contar0).dVCuponMD = 0
    End If
Select Case matpos(contar0).fValuacion
Case "VREPORTO"
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " " & dxv + pfwd
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = dxv + pfwd
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " " & Format(dxv + pfwd, "000000")
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = dxv + pfwd
     If pfwd > 0 Then
        contar = contar + 1
        ReDim Preserve MatNodosFREx(0 To contar)
        MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " " & pfwd
        MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
        MatNodosFREx(3, contar) = parval.perfwd
        MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " " & Format(parval.perfwd, "000000")
        MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0).fRiesgo1MD
        MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & parval.perfwd
     End If
Case "ACCION"
     contar = contar + 1
     ReDim Preserve MatNodosFREx(0 To contar)
     MatNodosFREx(1, contar) = matposdiv(contar0).TCambioDiv & " 0"
     MatNodosFREx(2, contar) = matposdiv(contar0).TCambioDiv
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposdiv(contar0).TCambioDiv & " 000000"
     MatFRPos(contar0, 1) = matposdiv(contar0).TCambioDiv
     MatFRPos(contar0, 2) = 0
Case "TIPO CAMBIO"
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposdiv(contar0).TCambioDiv & " 0"
     MatNodosFREx(2, contar) = matposdiv(contar0).TCambioDiv
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposdiv(contar0).TCambioDiv & " 000000"
     MatFRPos(contar0, 1) = matposdiv(contar0).TCambioDiv
     MatFRPos(contar0, 2) = 0
Case "BONDE"   'BONDES y BONOS PARECIDOS
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0, CSobreT) & " " & dxv
     MatNodosFREx(2, contar) = matposmd(contar0, CSobreT)
     MatNodosFREx(3, contar) = dxv
     MatNodosFREx(4, contar) = matposmd(contar0, CSobreT) & " " & Format(dxv, "000000")
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0, CSobreT)
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & dxv
Case "BONO D81"   'BONO D8 referenciado a tiie
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0, CSobreT) & " " & dxv
     MatNodosFREx(2, contar) = matposmd(contar0, CSobreT)
     MatNodosFREx(3, contar) = dxv
     MatNodosFREx(4, contar) = matposmd(contar0, CSobreT) & " " & Format(dxv, "000000")
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0, CSobreT)
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & dxv
Case "BONO TASA FIJA" 'BONO TASA FIJA
     matfl = ObtenerFlujosMD(matposmd(contar0).iFlujoMD, matposmd(contar0).fFlujoMD, MatFlujosMD)
     Call DNBonoCurva(fechaval, contar0, parval.perfwd, matfl, matposmd(contar0).fRiesgo1MD)
Case "BONO TASA FIJA EXT" '
     matfl = ObtenerFlujosMD(matposmd(contar0).iFlujoMD, matposmd(contar0).fFlujoMD, MatFlujosMD)
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).tCambioMD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).tCambioMD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).tCambioMD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).tCambioMD
     MatFRPos(contar0, 2) = 0
     Call DNBonoCurva(fechaval, contar0, parval.perfwd, matfl, matposmd(contar0).fRiesgo1MD)
Case "BONO TASA FIJA Y" 'BONO TASA FIJA YIELD
  'se valua con la curva
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
Case "BONO TASA FIJA Y EXT" 'BONO YIELD EXT
  'se valua con la curva
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).tCambioMD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).tCambioMD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).tCambioMD & " 000000"
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0).tCambioMD & ","
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & 0
Case "BONO STC Y"       'certificados bursátiles
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo3MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo3MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo3MD & " 000000"
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0).fRiesgo3MD
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & 0
Case "B C CERO"
' se crea la curva con que se descuenta
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " " & dxv + parval.perfwd
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = dxv + parval.perfwd
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " " & Format(dxv + parval.perfwd, "000000")
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = dxv + parval.perfwd
     If parval.perfwd <> 0 Then
        contar = contar + 1
        ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
        MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " " & parval.perfwd
        MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
        MatNodosFREx(3, contar) = parval.perfwd
        MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " " & Format(parval.perfwd, "000000")
        MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0).fRiesgo1MD
        MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & parval.perfwd
     End If
Case "BREMS"   'BondesD
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0).fRiesgo1MD & " 0"
     MatNodosFREx(2, contar) = matposmd(contar0).fRiesgo1MD
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposmd(contar0).fRiesgo1MD & " 000000"
     MatFRPos(contar0, 1) = matposmd(contar0).fRiesgo1MD
     MatFRPos(contar0, 2) = 0
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposmd(contar0, CSobreT) & " " & dxv
     MatNodosFREx(2, contar) = matposmd(contar0, CSobreT)
     MatNodosFREx(3, contar) = dxv
     MatNodosFREx(4, contar) = matposmd(contar0, CSobreT) & " " & Format(dxv, "000000")
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposmd(contar0, CSobreT)
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & dxv
Case "SWAP"
   If Not EsVariableVacia(matposswaps(contar0).TCambio1Swap) Then
      contar = contar + 1
      ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
      MatNodosFREx(1, contar) = matposswaps(contar0).TCambio1Swap & " 0"
      MatNodosFREx(2, contar) = matposswaps(contar0).TCambio1Swap
      MatNodosFREx(3, contar) = 0
      MatNodosFREx(4, contar) = matposswaps(contar0).TCambio1Swap & " 000000"
      MatFRPos(contar0, 1) = matposswaps(contar0).TCambio1Swap
      MatFRPos(contar0, 2) = 0
   End If
   If Not EsVariableVacia(matposswaps(contar0).TCambio2Swap) Then
      contar = contar + 1
      ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
      MatNodosFREx(1, contar) = matposswaps(contar0).TCambio2Swap & " 0"
      MatNodosFREx(2, contar) = matposswaps(contar0).TCambio2Swap
      MatNodosFREx(3, contar) = 0
      MatNodosFREx(4, contar) = matposswaps(contar0).TCambio2Swap & " 000000"
      MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposswaps(contar0).TCambio2Swap
      MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & 0

   End If
   Call DefinirParamValSwap(matposswaps, contar0, parval.perfwd, parval.si_int_flujos, tCambio1, tCambio2, parval1, parval2)
   'Call DetNodosSwap(fechaval, contar0, parval1, parval2, mflujosswaps, MatPosSwaps(contar0).FRiesgo1Swap, MatPosSwaps(contar0).FRiesgo2Swap, matpos(contar0).fRiesgo2MD, matpos(contar0).fRiesgo2MD)
 
Case "DEUDA"
  'se leen 4 curvas para la valuacion del swap
     'Call DefinirParamValSwap(MatPosDeuda, contar0, parval.perfwd, tCambio1, tCambio2, parval1, parval2)
  
Case "FWD IND"

Case "FWD TC"
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposfwd(contar0).FRiesgo1Fwd & " " & dxv + parval.perfwd
     MatNodosFREx(2, contar) = matposfwd(contar0).FRiesgo1Fwd
     MatNodosFREx(3, contar) = dxv + parval.perfwd
     MatNodosFREx(4, contar) = matposfwd(contar0).FRiesgo1Fwd & " " & Format(dxv + parval.perfwd, "000000")
     MatFRPos(contar0, 1) = matposfwd(contar0).FRiesgo1Fwd
     MatFRPos(contar0, 2) = dxv + parval.perfwd
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposfwd(contar0).FRiesgo3Fwd & " " & dxv + parval.perfwd
     MatNodosFREx(2, contar) = matposfwd(contar0).FRiesgo3Fwd
     MatNodosFREx(3, contar) = dxv + parval.perfwd
     MatNodosFREx(4, contar) = matposfwd(contar0).FRiesgo3Fwd & " " & Format(dxv + parval.perfwd, "000000")
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposfwd(contar0).FRiesgo3Fwd
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & dxv + parval.perfwd
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(1, contar) = matposfwd(contar0).FRiesgo2Fwd & " " & dxv + parval.perfwd
     MatNodosFREx(2, contar) = matposfwd(contar0).FRiesgo2Fwd
     MatNodosFREx(3, contar) = dxv + parval.perfwd
     MatNodosFREx(4, contar) = matposfwd(contar0).FRiesgo2Fwd & " " & Format(dxv + parval.perfwd, "000000")
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposfwd(contar0).FRiesgo2Fwd
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & dxv + parval.perfwd
     contar = contar + 1
     ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
     MatNodosFREx(contar).nomFactor = matposfwd(contar0).TCambioFwd & " 0"
     MatNodosFREx(2, contar) = matposfwd(contar0).TCambioFwd
     MatNodosFREx(3, contar) = 0
     MatNodosFREx(4, contar) = matposfwd(contar0).TCambioFwd & " 000000"
     MatFRPos(contar0, 1) = MatFRPos(contar0, 1) & "," & matposfwd(contar0).TCambioFwd
     MatFRPos(contar0, 2) = MatFRPos(contar0, 2) & "," & 0
Case "FWD TASA"

Case "", Null
   MensajeProc = "No se ha catalogado el titulo: Posicion-> " & matpos(contar0).C_Posicion & " . Modificar la tabla VALUACION"
   MsgBox MensajeProc
Case Else
  MensajeProc = "No se ha definido como valuar: " & contar0 & " " & matpos(contar0).fValuacion & " . Se debe modificar el sistema."
End Select
   AvanceProc = contar0 / noreg
   DoEvents
Next contar0
End If
For i = 1 To contar
  If Len(Trim(MatNodosFREx(2, i))) = 0 Then MsgBox "hay un factor de riesgo nulo"
Next i


'se obtienen los factores distintos por curva o indice

ReDim MatCaracFRiesgo(1 To UBound(matn, 1)) As propNodosFRiesgo
For i = 1 To UBound(matn, 1)
    For j = 1 To UBound(MatNodosFREx, 1)
        If matn(i, 1) = MatNodosFREx(j, 4) Then
           MatCaracFRiesgo(i).indFactor = MatNodosFREx(j, 1)
           MatCaracFRiesgo(i).tfactor = "Indice"
           MatCaracFRiesgo(i).nomFactor = MatNodosFREx(j, 2)
           MatCaracFRiesgo(i).plazo = MatNodosFREx(j, 3)
           MatCaracFRiesgo(i).tinterpol = 1
           MatCaracFRiesgo(i).descFactor = "Desc larga"
           'MatCaracFRiesgo(i, 7) = MatNodosFREx(j, 4)
           Exit For
        End If
    Next j
Next i
NoFactores = UBound(MatCaracFRiesgo, 1)
ReDim SiFactorRiesgo(1 To NoFactores) As Boolean
For i = 1 To NoFactores
    SiFactorRiesgo(i) = True
Next i

NoGruposFR = UBound(matc, 1)
ReDim MatResFRiesgo(1 To NoGruposFR) As New resPropFRiesgo
ReDim MatResFRiesgo1(1 To NoGruposFR) As Variant
For i = 1 To NoGruposFR
    For j = 1 To NoFactores
        If matc(i, 1) = MatCaracFRiesgo(j).nomFactor Then
           MatResFRiesgo(i, 2) = MatResFRiesgo(i, 2) + 1
        End If
    Next j
Next i
For i = 1 To NoGruposFR
    MatResFRiesgo(i, 1) = matc(i, 1)
    If i <> 1 Then
       MatResFRiesgo(i, 3) = MatResFRiesgo(i - 1, 3) + MatResFRiesgo(i, 2)
    Else
       MatResFRiesgo(i, 3) = MatResFRiesgo(i, 2)
    End If
    MatResFRiesgo(i, 4) = "TASA"
    MatResFRiesgo(i, 5) = 1
Next i
noreg1 = 0
For i = 1 To NoGruposFR
    noreg1 = Maximo(noreg1, MatResFRiesgo(i, 2))
Next i
ReDim MatPlazos(1 To noreg1, 1 To NoGruposFR) As Long
ReDim MatDescripFR(1 To noreg1, 1 To NoGruposFR) As Variant

For i = 1 To NoGruposFR
    For j = 1 To MatResFRiesgo(i, 2)
        If i <> 1 Then
           MatPlazos(j, i) = MatCaracFRiesgo(MatResFRiesgo(i - 1, 3) + j).plazo  'plazo
           MatDescripFR(j, i) = MatCaracFRiesgo(MatResFRiesgo(i - 1, 3) + j).descFactor  'descripcion
        Else
           MatPlazos(j, i) = MatCaracFRiesgo(j).plazo     'plazo
           MatDescripFR(j, i) = MatCaracFRiesgo(j).descFactor     'descripcion
        End If
    Next j
Next i
MensajeProc = "Se encontraron " & UBound(MatCaracFRiesgo, 1) & " factores de riesgo"
End Function

Sub DetNodosSwap(ByVal fecha As Date, ByVal indice As Integer, ByRef parval1 As paramValFlujo, ByRef parval2 As paramValFlujo, ByRef mflujos() As estFlujosDeuda, ByVal txtcdesc1 As String, ByVal txtcpago1 As String, ByVal txtcdesc2 As String, ByVal txtcpago2 As String)

    If Not EsVariableVacia(txtcdesc1) And Not EsVariableVacia(txtcpago1) Then
        Call detNodosDeudaTVar(fecha, indice, parval1, mflujos, txtcdesc1, txtcpago1)
    Else
        Call detNodosDeudaTFija(fecha, indice, parval1, mflujos, txtcdesc1)
    End If
    If Not EsVariableVacia(txtcdesc2) And Not EsVariableVacia(txtcpago2) Then
        Call detNodosDeudaTVar(fecha, indice, parval2, mflujos, txtcdesc2, txtcpago2)
    Else
        Call detNodosDeudaTFija(fecha, indice, parval2, mflujos, txtcdesc2)
    End If
End Sub

Sub detNodosDeudaTFija(ByVal fecha As Date, ByVal indice As Integer, ByRef parval As paramValFlujo, ByRef matflujos() As estFlujosDeuda, ByVal txtcurva1 As String)
Dim tinterpol As Integer
Dim indicex As Long
Dim i As Long
Dim contar As Long
Dim mrvalflujo() As resValFlujo

'para el calculo del valor presente de una deuda de tasa variable con amortizaciones
'curva1 para el descuento

'primero determinamos el flujos

'1   clave de ikos
'2   pata
'3   fecha de inicio
'4   fecha final
'5   fecha pago intereses
'6   pago intereses
'7   aplicar int todo el saldo
'8   saldo
'9   amortizacion
'10  tasa texto
'11  spread y hasta aqui son datos de entrada
'datos derivados
'12  dias inicio cupon
'13  dias fin cupon
'14  dias desc flujo
'15  periodo cupon
'16  periodo cupon efectivo
'17  saldo efectivo a aplicar intereses
'18  intereses generados en el periodo
'19  intereses acumulados en el periodo
'20  int pagados en el periodo
'21  int acumulados sig periodo
'22  pago total
'23  tasa descuento pago
'24  factor descuento
'25  valor presente
'26  tipo de pata
indicex = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indicex <= 0 Then
   Exit Sub
End If
contar = UBound(MatNodosFREx, 2)
mrvalflujo = CrearTablaAmortiza(fecha, indicex, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
For i = 1 To parval.finFLujo - indicex + 1
    If mrvalflujo(i, 14) + parval.perfwd > 0 Then
       contar = contar + 1
 ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
       MatNodosFREx(1, contar) = txtcurva1 & " " & mrvalflujo(i, 14) + parval.perfwd
       MatNodosFREx(2, contar) = txtcurva1
       MatNodosFREx(3, contar) = mrvalflujo(i, 14) + parval.perfwd
       MatNodosFREx(4, contar) = txtcurva1 & " " & Format(mrvalflujo(i, 14) + parval.perfwd, "000000")
       MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva1 & ","
       MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & mrvalflujo(i, 14) + parval.perfwd & ","
    End If
Next i
End Sub

Sub detNodosDeudaTVar(ByVal fecha As Date, ByVal indice As Integer, ByRef parval As paramValFlujo, ByRef matflujos() As estFlujosDeuda, ByVal txtcurva1 As String, ByVal txtcurva2 As String)
Dim curvacp() As propCurva
Dim mrvalflujo() As resValFlujo
Dim indicex As Long
Dim contar As Long
Dim i As Long

'se determina la estructura de la tabla de amortizacion

indicex = EstIndTAmortiza(fecha, parval.inFLujo, parval.finFLujo, matflujos)
If indicex <= 0 Then 'no se debe de valuar si ya finalizo
   Exit Sub
End If
'se dimensiona una matriz donde se colocan el desglose de los resultados
mrvalflujo = CrearTablaAmortiza(fecha, indicex, parval.finFLujo, parval.intFinFlujos, parval.sobret, matflujos)
contar = UBound(MatNodosFREx, 2)
If parval.perfwd > 0 Then
   contar = contar + 1
   ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
   MatNodosFREx(1, contar) = txtcurva2 & " " & parval.perfwd
   MatNodosFREx(2, contar) = txtcurva2
   MatNodosFREx(3, contar) = parval.perfwd
   MatNodosFREx(4, contar) = txtcurva2 & " " & Format(parval.perfwd, "000000")
   MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva2 & ","
   MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & parval.perfwd & ","
End If
For i = 1 To parval.finFLujo - indicex + 1
    mrvalflujo(i, 16) = mrvalflujo(i, 15)                              'periodo cupon a aplicar
    If fecha < mrvalflujo(i, 3) Then                                  'tiene que ser una tasa forward
       mrvalflujo(i, 10) = TasaFwdCurva(mrvalflujo(i, 12) + parval.perfwd, mrvalflujo(i, 12) + parval.pcref + parval.perfwd, curvacp, parval.modint2)
       If mrvalflujo(i, 12) + parval.perfwd > 0 Then
       contar = contar + 1
       ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
       MatNodosFREx(1, contar) = txtcurva2 & " " & mrvalflujo(i, 12) + parval.perfwd
       MatNodosFREx(2, contar) = txtcurva2
       MatNodosFREx(3, contar) = mrvalflujo(i, 12) + parval.perfwd
       MatNodosFREx(4, contar) = txtcurva2 & " " & Format(mrvalflujo(i, 12) + parval.perfwd, "000000")
       MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva2 & ","
       MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & mrvalflujo(i, 12) + parval.perfwd & ","
       End If
       If mrvalflujo(i, 12) + parval.pcref + parval.perfwd > 0 Then
       contar = contar + 1
       ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
       MatNodosFREx(1, contar) = txtcurva2 & " " & mrvalflujo(i, 12) + parval.pcref + parval.perfwd
       MatNodosFREx(2, contar) = txtcurva2
       MatNodosFREx(3, contar) = mrvalflujo(i, 12) + parval.pcref + parval.perfwd
       MatNodosFREx(4, contar) = txtcurva2 & " " & Format(mrvalflujo(i, 12) + parval.pcref + parval.perfwd, "000000")
       MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva2 & ","
       MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & mrvalflujo(i, 12) + parval.pcref + parval.perfwd & ","
       

       End If
    Else
       If mrvalflujo(i, 10) = 0 Then
          mrvalflujo(i, 10) = TasaFwdCurva(parval.perfwd, parval.pcref + parval.perfwd, curvacp, parval.modint2)
          contar = contar + 1
          ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
          MatNodosFREx(1, contar) = txtcurva2 & " " & parval.pcref + parval.perfwd
          MatNodosFREx(2, contar) = txtcurva2
          MatNodosFREx(3, contar) = parval.pcref + parval.perfwd
          MatNodosFREx(4, contar) = txtcurva2 & " " & Format(parval.pcref + parval.perfwd, "000000")
          MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva2 & ","
          MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & parval.pcref + parval.perfwd & ","

       End If
    End If
'saldo al que se le aplicaran los intereses
'intereses generados periodo = (saldo+intereses per ant) * (tasa+spread) * parval.pcref/360
 'intereses acumulados periodo
 'intereses pagados en el periodo= hasta el total de intereses acumulados
 'intereses acumulados sig periodo=intereses acumulados-intereses pagados
 'pago total sin descontar=amortizacion+intereses pagados
 'tasa de descuento
       If mrvalflujo(i, 14) + parval.perfwd > 0 Then
          contar = contar + 1
       ReDim Preserve MatNodosFREx(1 To 4, 0 To contar)
          MatNodosFREx(1, contar) = txtcurva1 & " " & mrvalflujo(i, 14) + parval.perfwd
          MatNodosFREx(2, contar) = txtcurva1
          MatNodosFREx(3, contar) = mrvalflujo(i, 14) + parval.perfwd
          MatNodosFREx(4, contar) = txtcurva1 & " " & Format(mrvalflujo(i, 14) + parval.perfwd, "000000")
          MatFRPos(indice, 1) = MatFRPos(indice, 1) & "," & txtcurva1 & ","
          MatFRPos(indice, 2) = MatFRPos(indice, 2) & "," & mrvalflujo(i, 14) + parval.perfwd & ","
       End If
Next i

End Sub

Function FechaFinMesCercana(ByVal fecha As Date) As Date
Dim año As Integer
Dim mes As Integer
Dim fecha1 As Date
Dim fecha2 As Date

   año = Year(fecha)
   mes = Month(fecha)
   fecha1 = DateSerial(año, mes, 1) - 1   'ultimo dia mes anterior
   fecha2 = DateSerial(año, mes + 1, 1) - 1  'ultimo dia de este mes
   If Abs(fecha - fecha1) < Abs(fecha - fecha2) Then
      FechaFinMesCercana = fecha1
   Else
      FechaFinMesCercana = fecha2
   End If
End Function

Sub DefinirParamValSwap(ByRef matpos() As propPosSwaps, ByVal contar0 As Integer, ByVal pfwd As Integer, ByVal si_int_flujos As Boolean, ByVal tCambio1 As Double, ByVal tCambio2 As Double, ByRef parval1 As paramValFlujo, ByRef parval2 As paramValFlujo)

        parval1.pcref = Val(matpos(contar0).PCuponActSwap)            'cupon 1
        parval1.perfwd = pfwd                                         'periodo fwd
        If si_int_flujos Then
           parval1.intInFlujos = "S"                                  'intercambio inicial flujos
           parval1.intFinFlujos = "S"                                 'intercambio intermedio y final flujos
        Else
           parval1.intInFlujos = matpos(contar0).IntercIFSwap         'intercambio inicial flujos
           parval1.intFinFlujos = matpos(contar0).IntercFFSwap        'intercambio intermedio y final flujos
        End If
        parval1.acumInt = matpos(contar0).RIntAct                     'reinversion de intereses
        parval1.modint1 = matpos(contar0).TInterpol1Swap              't interpol curva desc 1
        If Not EsVariableVacia(matpos(contar0).TInterpol3Swap) Then
           parval1.modint2 = matpos(contar0).TInterpol3Swap      't interpol curva pago 1
        Else
           parval1.modint2 = 0
        End If
        
        parval1.inFLujo = matpos(contar0).IFlujoActSwap           'inicio flujos activa
        parval1.finFLujo = matpos(contar0).FFlujoActSwap          'fin flujos activa
        parval1.tCambio = tCambio1                                'tipo cambio 1
        parval1.sobret = matpos(contar0).STActiva                 'sobretasa activa
        parval1.convInt = matpos(contar0).ConvIntAct              'CONV INT ACTIVA
        
        parval2.pcref = Val(matpos(contar0).PCuponPasSwap)        'cupon2
        parval2.perfwd = pfwd                                     'periodo fwd
        If si_int_flujos Then
           parval2.intInFlujos = "S"                                  'intercambio inicial flujos
           parval2.intFinFlujos = "S"                                 'intercambio intermedio y final flujos
        Else
           parval2.intInFlujos = matpos(contar0).IntercIFSwap        'intercambio inicial flujos
           parval2.intFinFlujos = matpos(contar0).IntercFFSwap       'intercambio intermedio y final flujos
        End If
        parval2.acumInt = matpos(contar0).RIntPas                 'reinversion de intereses
        parval2.modint1 = matpos(contar0).TInterpol2Swap          't interpol curva desc 2
        If Not EsVariableVacia(matpos(contar0).TInterpol4Swap) Then
           parval2.modint2 = Val(matpos(contar0).TInterpol4Swap)     't interpol curva pago 2
        Else
           parval2.modint2 = 0
        End If
        parval2.inFLujo = matpos(contar0).IFlujoPasSwap           'inicio flujos pasiva
        parval2.finFLujo = matpos(contar0).FFlujoPasSwap          'fin flujos pasiva
        parval2.tCambio = tCambio2                                'tipo cambio 2
        parval2.sobret = matpos(contar0).STPasiva                 'sobretasa activa
        parval2.convInt = matpos(contar0).ConvIntPas              'CONVENCION DE INTERESES pasiva
        parval2.sicalcVE = True
End Sub

Sub DefinirParamValDeuda(ByRef matpos() As propPosDeuda, ByVal contar0 As Integer, ByVal pfwd As Long, ByVal si_int_flujos As Boolean, ByVal tCambio1 As Double, ByRef parval1 As paramValFlujo)
        parval1.pcref = matpos(contar0).PCuponDeuda                    'cupon 1
        parval1.perfwd = pfwd                                          'periodo fwd
        If si_int_flujos Then
           parval1.intInFlujos = "S"                                   'intercambio inicial flujos
           parval1.intFinFlujos = "S"                                  'intercambio intermedio y final flujos
         Else
           parval1.intInFlujos = matpos(contar0).InteriDeuda           'intercambio inicial flujos
           parval1.intFinFlujos = matpos(contar0).InterfDeuda          'intercambio intermedio y final flujos
        End If
        parval1.acumInt = matpos(contar0).RintDeuda                    'reinversion de intereses
        parval1.modint1 = matpos(contar0).TInterpol1Deuda              't interpol curva desc 1
        parval1.modint2 = matpos(contar0).TInterpol2Deuda              't interpol curva pago 1
        parval1.inFLujo = matpos(contar0).IFlujoDeuda                  'inicio flujos activa
        parval1.finFLujo = matpos(contar0).FFlujoDeuda                 'fin flujos activa
        parval1.tCambio = tCambio1                                     'tipo cambio 1
        'parval1(12) = matpos(contar0).TRefDeuda                       'tasa cupon
        parval1.sobret = matpos(contar0).SpreadDeuda                   'sobretasa activa
        parval1.convInt = matpos(contar0).ConvIntDeuda                 'CONVENCION DE CALCULO DE INTERESES
End Sub

Public Function EsArrayVacio(aarray) As Boolean
Dim valor As Long
    On Error GoTo opciones
    valor = UBound(aarray)
    If valor = 0 Then
       EsArrayVacio = True
    Else
       EsArrayVacio = False
    End If
    Exit Function
opciones:
    EsArrayVacio = True ' Error 9 (Subscript out of range)
    On Error GoTo 0
       
End Function

Function EsArrayValAnexar(mata)
If Not EsArrayVacio(mata) Then
 If UBound(mata, 1) > 0 And UBound(mata, 2) > 0 Then
  EsArrayValAnexar = True
 Else
  EsArrayValAnexar = False
 End If
Else
 EsArrayValAnexar = False
End If
End Function

Function CalcDVCBono(ByVal fecha As Date, ByRef matfl() As estFlujosMD) As Integer
Dim valor As Integer
Dim i As Integer
valor = 0
For i = 1 To UBound(matfl, 1)
    If fecha >= matfl(i).finicio And fecha < matfl(i).ffin Then
       valor = matfl(i).ffin - fecha
       Exit For
    End If
Next i
CalcDVCBono = valor
End Function

Function EsFechaVaR(ByVal fecha As Date) As Boolean
Dim txtfiltro As String
Dim txtfecha As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaFechasVaR & " WHERE FECHA = " & txtfecha
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   EsFechaVaR = True
Else
   EsFechaVaR = False
End If
End Function

Function DetFechaFNoEsc(ByVal fecha As Date, ByVal noesc As Long) As Date
On Error GoTo hayerror
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT FECHA,ORDEN FROM (SELECT FECHA,ROWNUM AS ORDEN FROM"
txtfiltro2 = txtfiltro2 & " (SELECT FECHA FROM " & TablaFechasVaR & " WHERE FECHA <= " & txtfecha
txtfiltro2 = txtfiltro2 & " ORDER BY FECHA DESC)) WHERE ORDEN = " & noesc
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   DetFechaFNoEsc = rmesa.Fields("FECHA")
   rmesa.Close
Else
   DetFechaFNoEsc = 0
End If
Exit Function
hayerror:
  MsgBox "DetFechaFNoEsc " & error(Err())
End Function
