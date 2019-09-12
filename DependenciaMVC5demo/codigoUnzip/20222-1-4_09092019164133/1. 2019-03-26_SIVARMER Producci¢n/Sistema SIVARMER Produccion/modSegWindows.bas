Attribute VB_Name = "modSegWindows"
Option Explicit

Private Type IPINFO

    dwAddr As Long ' Dirección Ip
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1 As Integer
    unused2 As Integer

End Type
  
Private Type MIB_IPADDRTABLE

    dEntrys As Long 'Numero de entradas de la tabla
    mIPInfo(5) As IPINFO 'Array de entradas de direcciones Ip

End Type
  
Private Type IP_Array

    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long

End Type
  
'Función Api CopyMemory
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)
  
'Función Api GetIpAddrTable para obtener la tabla de direcciones IP
Private Declare Function GetIpAddrTable _
                Lib "IPHlpApi" (pIPAdrTable As Byte, _
                                pdwSize As Long, _
                                ByVal Sort As Long) As Long
  
'Función para Convertir el valor de tipo Long a un string
Private Function ConvertirDirecciónAstring(longAddr As Long) As String

    Dim myByte(3) As Byte 'array de tipo Byte
    Dim Cnt       As Long
      
    CopyMemory myByte(0), longAddr, 4

    For Cnt = 0 To 3
        ConvertirDirecciónAstring = ConvertirDirecciónAstring + CStr(myByte(Cnt)) + "."
    Next Cnt

    ConvertirDirecciónAstring = Left$(ConvertirDirecciónAstring, Len(ConvertirDirecciónAstring) - 1)
End Function


Function EliminarPassViejo(ByRef mata() As Variant, ByVal orden As Integer) As Variant()
Dim noreg As Integer
Dim clave As Integer
Dim i As Integer
Dim j As Integer

'contar cuantos password hay en el historial
noreg = UBound(mata, 1)
'si hay "orden" se elimina el password mas viejo
If noreg >= orden Then
 clave = orden
Else
 clave = noreg + 1
End If
ReDim matb(1 To clave, 1 To 6) As Variant
For i = 1 To clave - 1
  For j = 1 To 5
  matb(i + 1, j) = mata(i, j)
  Next j
Next i
EliminarPassViejo = matb
End Function


Function ValidartxtPalin(ByVal txtcadena As String) As Boolean
Dim largo As Integer
Dim largo1 As Integer
Dim texto1 As String
Dim texto2 As String


largo = Len(txtcadena)
If largo Mod 2 = 0 Then
   largo1 = largo / 2
   texto1 = Mid$(txtcadena, 1, largo1)
   texto2 = InvCadena(Mid$(txtcadena, largo1 + 1, largo1))
Else
largo1 = (largo - 1) / 2
   texto1 = Mid$(txtcadena, 1, largo1)
   texto2 = InvCadena(Mid$(txtcadena, largo1 + 1, largo1))
End If
If texto1 = texto2 Then
   ValidartxtPalin = True
Else
   ValidartxtPalin = False
End If
End Function

Function Valtxt2Carac(ByVal txtcadena As String) As Boolean
Dim largo As Integer
Dim i As Integer

largo = Len(txtcadena)
For i = 1 To largo - 1
If Mid$(txtcadena, i, 1) = Mid$(txtcadena, i + 1, 1) Then
Valtxt2Carac = True
Exit Function
End If
Next i
Valtxt2Carac = False
End Function

Function InvCadena(ByVal txtcadena As String) As String
Dim largo As Integer
Dim i As Integer
Dim txtsalida As String

largo = Len(txtcadena)
txtsalida = ""
For i = 1 To largo
txtsalida = txtsalida & Mid$(txtcadena, largo - i + 1, 1)
Next i
InvCadena = txtsalida
End Function

Function ContrCMay(ByVal txtcadena As String)
Dim largo As Integer
Dim i As Integer

largo = Len(txtcadena)
For i = 1 To largo
    If InStr(txtCadMay, Mid$(txtcadena, i, 1)) <> 0 Then
       ContrCMay = True
       Exit Function
    End If
Next i
ContrCMay = False
End Function

Function ContrCMin(txtcadena)
Dim largo As Integer
Dim i As Integer

largo = Len(txtcadena)
For i = 1 To largo
    If InStr(txtCadMin, Mid$(txtcadena, i, 1)) <> 0 Then
       ContrCMin = True
       Exit Function
    End If
Next i
ContrCMin = False
End Function

Function ContrCNum(txtcadena)
Dim largo As Integer
Dim i As Integer

largo = Len(txtcadena)
For i = 1 To largo
    If InStr(txtCadNum, Mid$(txtcadena, i, 1)) <> 0 Then
       ContrCNum = True
       Exit Function
    End If
Next i
ContrCNum = False
End Function

Function ContrCCEsp(ByVal txtcadena As String)
Dim largo As Integer
Dim i As Integer

largo = Len(txtcadena)
For i = 1 To largo
    If InStr(txtCadCarEsp, Mid$(txtcadena, i, 1)) <> 0 Then
       ContrCCEsp = True
       Exit Function
    End If
Next i
ContrCCEsp = False
End Function

Function Busc2CarEsp(ByVal txtcadena As String)
Dim i As Integer

For i = 1 To Len(txtcadena) - 1
    If InStr(txtCadCarEsp, Mid$(txtcadena, i, 1)) <> 0 And InStr(txtCadCarEsp, Mid$(txtcadena, i + 1, 1)) <> 0 Then
       Busc2CarEsp = True
       Exit Function
    End If
Next i
Busc2CarEsp = False
End Function

Function Busc2Num(ByVal txtcadena As String)
Dim i As Integer

For i = 1 To Len(txtcadena) - 1
    If InStr(txtCadNum, Mid$(txtcadena, i, 1)) <> 0 And InStr(txtCadNum, Mid$(txtcadena, i + 1, 1)) <> 0 Then
       Busc2Num = True
       Exit Function
    End If
Next i
Busc2Num = False
End Function

Function Busc4Min(ByVal txtcadena As String)
Dim i As Integer

For i = 1 To Len(txtcadena) - 3
    If InStr(txtCadMin, Mid$(txtcadena, i, 1)) <> 0 And InStr(txtCadMin, Mid$(txtcadena, i + 1, 1)) <> 0 And InStr(txtCadMin, Mid$(txtcadena, i + 2, 1)) <> 0 And InStr(txtCadMin, Mid$(txtcadena, i + 3, 1)) <> 0 Then
       Busc4Min = True
       Exit Function
    End If
Next i
Busc4Min = False
End Function

Function Busc4May(ByVal txtcadena As String)
Dim i As Integer
For i = 1 To Len(txtcadena) - 3
    If InStr(txtCadMay, Mid$(txtcadena, i, 1)) <> 0 And InStr(txtCadMay, Mid$(txtcadena, i + 1, 1)) <> 0 And InStr(txtCadMay, Mid$(txtcadena, i + 2, 1)) <> 0 And InStr(txtCadMay, Mid$(txtcadena, i + 3, 1)) <> 0 Then
       Busc4May = True
       Exit Function
    End If
Next i
Busc4May = False
End Function


Public Function Encrypt(ByVal Word As String, ByVal Key As String, _
Optional ByVal Mode As Boolean = False) As String
    Dim w As Long, k As Long, p As Long, j As Long, NuChr As Long
    Dim Cd As String, Kd As String, Rd As String
    w = Len(Word)
    k = Len(Key)
    ' Modalidad de cifrado...
    If Mode = False Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) + Asc(Kd)
            If NuChr > 255 Then
                NuChr = NuChr - 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function
    End If
    ' Modalidad de descifrado...
    If Mode = True Then
        For j = 1 To w
            Cd = Mid(Word, j, 1)
            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) - Asc(Kd)
            If NuChr < 0 Then
                NuChr = NuChr + 255
            End If
            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function
    End If
End Function

Function SHA(ByVal sMessage) As String
    Dim i, result(32), Temp(8) As Double, fraccubeprimes, hashValues
    Dim done512, index512, words(64) As Double, index32, mask(4)
    Dim s0, s1, T1, T2, maj, ch, strLen
    Dim txtcadena As String
 
    mask(0) = 4294967296#
    mask(1) = 16777216
    mask(2) = 65536
    mask(3) = 256
 
    hashValues = Array( _
        1779033703, 3144134277#, 1013904242, 2773480762#, _
        1359893119, 2600822924#, 528734635, 1541459225)
 
    fraccubeprimes = Array( _
        1116352408, 1899447441, 3049323471#, 3921009573#, 961987163, 1508970993, 2453635748#, 2870763221#, _
        3624381080#, 310598401, 607225278, 1426881987, 1925078388, 2162078206#, 2614888103#, 3248222580#, _
        3835390401#, 4022224774#, 264347078, 604807628, 770255983, 1249150122, 1555081692, 1996064986, _
        2554220882#, 2821834349#, 2952996808#, 3210313671#, 3336571891#, 3584528711#, 113926993, 338241895, _
        666307205, 773529912, 1294757372, 1396182291, 1695183700, 1986661051, 2177026350#, 2456956037#, _
        2730485921#, 2820302411#, 3259730800#, 3345764771#, 3516065817#, 3600352804#, 4094571909#, 275423344, _
        430227734, 506948616, 659060556, 883997877, 958139571, 1322822218, 1537002063, 1747873779, _
        1955562222, 2024104815, 2227730452#, 2361852424#, 2428436474#, 2756734187#, 3204031479#, 3329325298#)
 
    If IsNull(sMessage) Then
       sMessage = ""
    End If
    strLen = Len(sMessage) * 8
    sMessage = sMessage & Chr(128)
    done512 = False
    index512 = 0
 
    If (Len(sMessage) Mod 64) < 56 Then
        sMessage = sMessage & String(56 - (Len(sMessage) Mod 64), Chr(0))
    ElseIf (Len(sMessage) Mod 64) > 56 Then
        sMessage = sMessage & String(120 - (Len(sMessage) Mod 64), Chr(0))
    End If
    sMessage = sMessage & Chr(0) & Chr(0) & Chr(0) & Chr(0)
 
    sMessage = sMessage & Chr(Int((strLen / mask(0) - Int(strLen / mask(0))) * 256))
    sMessage = sMessage & Chr(Int((strLen / mask(1) - Int(strLen / mask(1))) * 256))
    sMessage = sMessage & Chr(Int((strLen / mask(2) - Int(strLen / mask(2))) * 256))
    sMessage = sMessage & Chr(Int((strLen / mask(3) - Int(strLen / mask(3))) * 256))
 
    Do Until done512
        For i = 0 To 15
            words(i) = Asc(Mid(sMessage, index512 * 64 + i * 4 + 1, 1)) * mask(1) + Asc(Mid(sMessage, index512 * 64 + i * 4 + 2, 1)) * mask(2) + Asc(Mid(sMessage, index512 * 64 + i * 4 + 3, 1)) * mask(3) + Asc(Mid(sMessage, index512 * 64 + i * 4 + 4, 1))
        Next
 
        For i = 16 To 63
            s0 = largeXor(largeXor(rightRotate(words(i - 15), 7, 32), rightRotate(words(i - 15), 18, 32), 32), Int(words(i - 15) / 8), 32)
            s1 = largeXor(largeXor(rightRotate(words(i - 2), 17, 32), rightRotate(words(i - 2), 19, 32), 32), Int(words(i - 2) / 1024), 32)
            words(i) = Mod32Bit(words(i - 16) + s0 + words(i - 7) + s1)
        Next
 
        For i = 0 To 7
            Temp(i) = hashValues(i)
        Next
 
        For i = 0 To 63
            s0 = largeXor(largeXor(rightRotate(Temp(0), 2, 32), rightRotate(Temp(0), 13, 32), 32), rightRotate(Temp(0), 22, 32), 32)
            maj = largeXor(largeXor(largeAnd(Temp(0), Temp(1), 32), largeAnd(Temp(0), Temp(2), 32), 32), largeAnd(Temp(1), Temp(2), 32), 32)
            T2 = Mod32Bit(s0 + maj)
            s1 = largeXor(largeXor(rightRotate(Temp(4), 6, 32), rightRotate(Temp(4), 11, 32), 32), rightRotate(Temp(4), 25, 32), 32)
            ch = largeXor(largeAnd(Temp(4), Temp(5), 32), largeAnd(largeNot(Temp(4), 32), Temp(6), 32), 32)
            T1 = Mod32Bit(Temp(7) + s1 + ch + fraccubeprimes(i) + words(i))
 
            Temp(7) = Temp(6)
            Temp(6) = Temp(5)
            Temp(5) = Temp(4)
            Temp(4) = Mod32Bit(Temp(3) + T1)
            Temp(3) = Temp(2)
            Temp(2) = Temp(1)
            Temp(1) = Temp(0)
            Temp(0) = Mod32Bit(T1 + T2)
        Next
 
        For i = 0 To 7
            hashValues(i) = Mod32Bit(hashValues(i) + Temp(i))
        Next
 
        If (index512 + 1) * 64 >= Len(sMessage) Then done512 = True
        index512 = index512 + 1
    Loop
 
    For i = 0 To 31
        result(i) = Int((hashValues(i \ 4) / mask(i Mod 4) - Int(hashValues(i \ 4) / mask(i Mod 4))) * 256)
    Next
 txtcadena = ""
 For i = 1 To 32
     txtcadena = txtcadena & Format(result(i - 1), "000")
 Next i
    SHA = txtcadena
End Function
 
Function Mod32Bit(value)
    Mod32Bit = Int((value / 4294967296# - Int(value / 4294967296#)) * 4294967296#)
End Function
 
Function rightRotate(value, amount, totalBits)
    'To leftRotate, make amount = totalBits - amount
    Dim i
    rightRotate = 0
 
    For i = 0 To (totalBits - 1)
        If i >= amount Then
            rightRotate = rightRotate + (Int((value / (2 ^ (i + 1)) - Int(value / (2 ^ (i + 1)))) * 2)) * 2 ^ (i - amount)
        Else
            rightRotate = rightRotate + (Int((value / (2 ^ (i + 1)) - Int(value / (2 ^ (i + 1)))) * 2)) * 2 ^ (totalBits - amount + i)
        End If
    Next
End Function
 
Function largeXor(value, xorValue, totalBits)
    Dim i, a, B
    largeXor = 0
 
    For i = 0 To (totalBits - 1)
        a = (Int((value / (2 ^ (i + 1)) - Int(value / (2 ^ (i + 1)))) * 2))
        B = (Int((xorValue / (2 ^ (i + 1)) - Int(xorValue / (2 ^ (i + 1)))) * 2))
        If a <> B Then
            largeXor = largeXor + 2 ^ i
        End If
    Next
End Function
 
Function largeNot(value, totalBits)
    Dim i, a
    largeNot = 0
 
    For i = 0 To (totalBits - 1)
        a = Int((value / (2 ^ (i + 1)) - Int(value / (2 ^ (i + 1)))) * 2)
        If a = 0 Then
            largeNot = largeNot + 2 ^ i
        End If
    Next
End Function
 
Function largeAnd(value, andValue, totalBits)
    Dim i, a, B
    largeAnd = 0
 
    For i = 0 To (totalBits - 1)
        a = Int((value / (2 ^ (i + 1)) - Int(value / (2 ^ (i + 1)))) * 2)
        B = (Int((andValue / (2 ^ (i + 1)) - Int(andValue / (2 ^ (i + 1)))) * 2))
        If a = 1 And B = 1 Then
            largeAnd = largeAnd + 2 ^ i
        End If
    Next
End Function

Sub GuardarIFBit(fecha, hora, txtusuario, noint)
Dim txtcadena As String
Dim txtfecha As String
Dim txthora As String


txtcadena = "INSERT INTO " & TablaBitacoraIF & " VALUES("
txtfecha = "TO_DATE('" & Format(fecha, "DD/MM/YYYY") & "','dd/mm/yyyy')"
txthora = "TO_DATE('" & Format(hora, "HH:MM:SS") & "','HH24:MI:SS')"
txtcadena = txtcadena & txtfecha & ","
txtcadena = txtcadena & txthora & ","
txtcadena = txtcadena & "'" & txtusuario & "',"
txtcadena = txtcadena & noint & ")"
ConAdo.Execute txtcadena
End Sub

Function RecuperarIP() As String
  
    Dim ret        As Long, Tel As Long
    Dim bBytes()   As Byte
    Dim TempList() As String
    Dim TempIP     As String
    Dim Tempi      As Long
    Dim Listing    As MIB_IPADDRTABLE
    Dim L3         As String
  
    On Error GoTo errSub
      
    GetIpAddrTable ByVal 0&, ret, True
  
    If ret <= 0 Then Exit Function
    ReDim bBytes(0 To ret - 1) As Byte
    ReDim TempList(0 To ret - 1) As String
  
    'recuperamos la tabla con las ip
    GetIpAddrTable bBytes(0), ret, False
    CopyMemory Listing.dEntrys, bBytes(0), 4
  
    For Tel = 0 To Listing.dEntrys - 1
        'Copiamos la estructura entera a la lista
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        TempList(Tel) = ConvertirDirecciónAstring(Listing.mIPInfo(Tel).dwAddr)
    Next Tel
  
    TempIP = TempList(0)

    For Tempi = 0 To Listing.dEntrys - 1
        L3 = Left$(TempList(Tempi), 3)

        If L3 <> "169" And L3 <> "127" And L3 <> "192" Then
            TempIP = TempList(Tempi)
        End If

    Next Tempi
      
    RecuperarIP = TempIP
    Exit Function

errSub:
    RecuperarIP = ""
End Function

Sub ValUsuarioDirActivo(ByVal txtnomusuario As String, ByRef siex As Boolean, ByRef edocta As String, ByRef txtnombre As String)
    Dim objetoUsuario, gruposSeguridad
    Dim ultimoInicioSesion As String
    Dim dominio As String
    
    dominio = "BANOBRAS"

    On Error Resume Next
    
    Set objetoUsuario = GetObject("WinNT://" + dominio + "/" + txtnomusuario + ",user")
    If Err.Number = 0 Then
       siex = True
        If objetoUsuario.AccountDisabled = True Then
           edocta = "Deshabilitado"
        Else
           edocta = "Habilitado"
        End If
        'Mostramos los datos del usuario
        txtnombre = objetoUsuario.Get("Fullname")

        Set objetoUsuario = Nothing
    Else
        siex = False
    End If

End Sub

Function IsGoodPWD(DomainName As String, sUserName As String, _
  chkPassword As String) As Boolean
On Error GoTo MyError:

If Not EsVariableVacia(chkPassword) Then
'====================================================
'=====================================================
'   Purpose:    To determin if a password given is the correct network password for the specified user
'
'   Syntax:     IsGoodPWD(User Name, Domain Name, Password)
'
'   Arguments:
'               username        -- login ID of user to verify
'               DomainName      -- Domain name the user and group reside in
'                           (Can also use IP Address of Primary Domain Controller)
'               chkPassword     -- Password to verify against the domain
'
'   Example:    IsGoodPWD("myusername", "mydomain", "mypass123")
'==============================================================
'=========================================================

    Dim dso As IADsOpenDSObject
    Dim UserObj As IADs
    Set dso = GetObject("LDAP:")
    Set UserObj = dso.OpenDSObject("LDAP://" & DomainName, sUserName, chkPassword, ADS_SECURE_AUTHENTICATION)
    IsGoodPWD = True
Else
    IsGoodPWD = False
End If
Exit Function
MyError:
    IsGoodPWD = False
End Function


