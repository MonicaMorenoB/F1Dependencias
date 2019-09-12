Attribute VB_Name = "MClob"

Option Explicit

      Const BLOCK_SIZE = 16384
    

      Sub WriteFromBinary(ByVal F As Long, fld As ADODB.Field, ByVal FieldSize As Long)
      Dim Data() As Byte, BytesRead As Long
        Do While FieldSize <> BytesRead
          If FieldSize - BytesRead < BLOCK_SIZE Then
            Data = fld.GetChunk(FieldSize - BLOCK_SIZE)
            BytesRead = FieldSize
          Else
            Data = fld.GetChunk(BLOCK_SIZE)
            BytesRead = BytesRead + BLOCK_SIZE
          End If
          Put #F, , Data
        Loop
      End Sub

      Sub WriteFromUnsizedBinary(ByVal F As Long, fld As ADODB.Field)
      Dim Data() As Byte, Temp As Variant
        Do
          Temp = fld.GetChunk(BLOCK_SIZE)
          If IsNull(Temp) Then Exit Do
          Data = Temp
          Put #F, , Data
        Loop While LenB(Temp) = BLOCK_SIZE
      End Sub

      Sub WriteFromText(ByVal F As Long, fld As ADODB.Field, ByVal FieldSize As Long)
      Dim datos As String, CharsRead As Long
        Do While FieldSize <> CharsRead
          If FieldSize - CharsRead < BLOCK_SIZE Then
            datos = fld.GetChunk(FieldSize - BLOCK_SIZE)
            CharsRead = FieldSize
          Else
            datos = fld.GetChunk(BLOCK_SIZE)
            CharsRead = CharsRead + BLOCK_SIZE
          End If
          Put #F, , datos
        Loop
      End Sub

      Sub WriteFromUnsizedText(ByVal F As Long, fld As ADODB.Field)
      Dim Data As String, Temp As Variant
        Do
          Temp = fld.GetChunk(BLOCK_SIZE)
          If IsNull(Temp) Then Exit Do
          Data = Temp
          Put #F, , Data
        Loop While Len(Temp) = BLOCK_SIZE
      End Sub

      Sub FileToBlob(ByVal FName As String, fld As ADODB.Field, Optional Threshold As Long = 1048576)
      '
      ' Assumes file exists
      ' Assumes calling routine does the UPDATE
      ' File cannot exceed approx. 2Gb in size
      '
      Dim F As Long, Data() As Byte, FileSize As Long
        F = FreeFile
        Open FName For Binary As #F
        FileSize = LOF(F)
        Select Case fld.Type
          Case adLongVarBinary
            If FileSize > Threshold Then
              ReadToBinary F, fld, FileSize
            Else
              Data = InputB(FileSize, F)
              fld.value = Data
            End If
          Case adLongVarChar, adLongVarWChar
            If FileSize > Threshold Then
              ReadToText F, fld, FileSize
            Else
              fld.value = Input(FileSize, F)
            End If
        End Select
        Close #F
      End Sub

      Sub ReadToBinary(ByVal F As Long, fld As ADODB.Field, _
                       ByVal FileSize As Long)
      Dim Data() As Byte, BytesRead As Long
        Do While FileSize <> BytesRead
          If FileSize - BytesRead < BLOCK_SIZE Then
            Data = InputB(FileSize - BytesRead, F)
            BytesRead = FileSize
          Else
            Data = InputB(BLOCK_SIZE, F)
            BytesRead = BytesRead + BLOCK_SIZE
          End If
          fld.AppendChunk Data
        Loop
      End Sub

      Sub ReadToText(ByVal F As Long, fld As ADODB.Field, _
                     ByVal FileSize As Long)
      Dim Data As String, CharsRead As Long
        Do While FileSize <> CharsRead
          If FileSize - CharsRead < BLOCK_SIZE Then
            Data = Input(FileSize - CharsRead, F)
            CharsRead = FileSize
          Else
            Data = Input(BLOCK_SIZE, F)
            CharsRead = CharsRead + BLOCK_SIZE
          End If
          fld.AppendChunk Data
        Loop
      End Sub

