Attribute VB_Name = "MExif"
Option Explicit
Public Type IFHeader
    ByteOrder(0 To 1) As Byte    ' &H4949 = "II" (Intel little endian) or &H4D4D = "MM" (Motorola big endian)
    'IFId(0 To 1)      As Byte   ' &H002A =  42
    IFId01            As Integer ' &H002A =  42
    OffsetIFD0        As Long    ' the file pos of the IFD-0-structure
End Type                         '8 Bytes
Public Type IFVersion
    A1 As Byte ' 0 = &H30
    A2 As Byte ' 2 = &H32
    B1 As Byte ' 1 = &H31
    B2 As Byte ' 0 = &H30
End Type
'stellt einen Bruch dar bspw für Verschlusszeit 1/50
Public Type IFRational
    Numerator   As Long 'muß für signed und unsigned herhalten
    Denominator As Long
End Type
'Public Type TCurrency
'    Value As Currency
'End Type
'die verschiedenen Datentypen die es in Exif gibt
' 1 = BYTE An 8-bit unsigned integer.,
' 2 = ASCII An 8-bit byte containing one 7-bit ASCII code. The final byte is terminated with NULL.,
' 3 = SHORT A 16-bit (2-byte) unsigned integer,
' 4 = LONG A 32-bit (4-byte) unsigned integer,
' 5 = RATIONAL Two LONGs. The first LONG is the numerator and the second LONG expresses the denominator.,
' 7 = UNDEFINED An 8-bit byte that can take any value depending on the field definition,
' 9 = SLONG A 32-bit (4-byte) signed integer (2's complement notation),
'10 = SRATIONAL Two SLONGs. The first SLONG is the numerator and the second SLONG is the denominator.

Public Enum IFDataType
    dtByte = 1       '8 Bit unsigned byte value
    dtASCII = 2      'ein Stringtyp
    dtShort = 3      '16 bit unsigned Integer
    dtLong = 4       '32 Bit unsigned Integer
    dtRational = 5   'Bruchtyp 2*32Bit unsinged Integer siehe IFTypeRational
    dtSByte = 6      '8 Bit signed byte value
    dtUndefined2 = 7 '8 Bit byte value
    dtSShort = 8     '16 bit signed Integer
    dtSLong = 9      '32 Bit signed Integer
    dtSRational = 10 'Bruchtyp 2*32Bit signed Integer siehe IFTypeRational
    dtFloat = 11     'Single Precision 32bit
    dtDouble = 12    'Double Precision 64bit
End Enum
Public Type IFDEntry
    Tag         As Integer ' 2 ' Tag is an enum and contains the meaning of the data see MTagIF, MTagExif, MTagGPS
    DataType    As Integer ' 2 ' As ExifType aber dann wäre es standardmäßig 32bit
    Count       As Long    ' 4 ' Number of Values
    ValueOffset As Long    ' 4 ' Value für IFDataType 1,3,4,9, Offset for 2,5,10
                     ' Sum: 12 ' Offset vom Start des Tiff-Headers
End Type
'ein Wert steckt entweder in IFDEntry.ValueOffset oder wenn dort ein Offset steht
'dann schreiben wir den Wert in den Variant IFDEntryValue.Value hinein
Public Type IFDEntryValue
    Entry As IFDEntry     ' 12
    Value As Variant      ' 16
End Type             ' Sum: 28
'IFD = Image File Directory
Public Type IFD
    Count     As Integer
    Entries() As IFDEntryValue
    OffsetNextIFD As Long 'der Dateioffset zur nächsten IFD-Strktur IFD_0 -> IFD_1
End Type
Public Const C_ExifHeader As String = "Exif"
Private Declare Sub GetMem1 Lib "msvbvm60" (ByRef pSrc As Any, ByRef pDst As Any)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByRef pSrc As Any, ByRef pDst As Any)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByRef pSrc As Any, ByRef pDst As Any)
Private Declare Sub GetMem8 Lib "msvbvm60" (ByRef pSrc As Any, ByRef pDst As Any)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef pDst As Any, ByRef pScr As Any, ByVal BytLength As Long)
Public Const MaxTagStrLen As Long = 28

Public Function GetPosition(ByVal ebr As FilEBinReader, ByVal strVal As String) As Long
    ' Searches the position of strVal in the file and returns it.
    ' Here it will be used to search for the string "Exif" in the file.
    Dim bytValue() As Byte: bytValue = StrConv(strVal, vbFromUnicode)
    ReDim bytbuffer(0 To 1023) As Byte
    
    ebr.ReadBytBuffer bytbuffer
    
    Dim i As Long, j As Long, u As Long: u = UBound(bytValue)
    GetPosition = -1
    For i = 0 To UBound(bytbuffer) - u
        If bytbuffer(i) = bytValue(0) Then
            For j = 0 To u
                If bytbuffer(i + j) <> bytValue(j) Then
                    Exit For
                Else
                    If (j = u) And (bytbuffer(i + u) = bytValue(u)) Then
                        GetPosition = i
                        'found it leave immediately
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Public Function Align4(ByVal val As Long) As Long
    Dim d As Long: d = val Mod 4
    If d Then Align4 = val + (4 - d) Else Align4 = val
End Function

' v ############################## v '    IFHeader    ' v ############################## v '
Public Function ReadIFHeader(ByRef this As IFHeader, _
                             ByVal ebr As FilEBinReader, _
                             ByVal OffsetIFHeader As Long) As Boolean
Try: On Error GoTo Catch
    'Halt, so ist das Käse!
    'man muss zuerst die ersten2bytes die die endianness angeben lesen, und dann erst die folgenden bytes lesen
    'ebr.BaseStream.Position = OffsetIFHeader + 1
    ebr.ReadBytBuffer this.ByteOrder, OffsetIFHeader + 1
    ebr.Endianness = MExif.IFHeader_Endianness(this)
    this.IFId01 = ebr.ReadInt16
    this.OffsetIFD0 = ebr.ReadInt32
    ReadIFHeader = True
    Exit Function
Catch: ErrHandler "ReadIFHeader"
End Function

Public Property Get IFHeader_Endianness(ByRef this As IFHeader) As EEndianness
    With this
        If .ByteOrder(0) = &H49 Then
            IFHeader_Endianness = EEndianness.IntelLittleEndian
            'we are on Intel hence no rotation needed
        Else
            IFHeader_Endianness = EEndianness.MotorolaBigEndian
            'we are on Intel hence we have to rotate the offset-bytes
        End If
    End With
End Property
Public Property Let IFHeader_Endianness(ByRef this As IFHeader, ByVal RHS As EEndianness)
    With this
        If RHS = EEndianness.IntelLittleEndian Then
            .ByteOrder(0) = &H49 ' = Asc("I")
            .ByteOrder(1) = &H49
        Else
            .ByteOrder(0) = &H4D ' = Asc("M")
            .ByteOrder(1) = &H4D
        End If
    End With
End Property
Public Property Get IFHeader_IsEqual(this As IFHeader, other As IFHeader) As Boolean
    Dim B As Boolean
    With this
        B = .ByteOrder(0) = other.ByteOrder(0): If Not B Then Exit Property
        B = .ByteOrder(1) = other.ByteOrder(1): If Not B Then Exit Property
        B = .IFId01 = other.IFId01:             If Not B Then Exit Property
        'B = .IFId01(0) = other.IFId(0): If Not B Then Exit Property
        'B = .IFId(1) = other.IFId(1): If Not B Then Exit Property
        B = .OffsetIFD0 = other.OffsetIFD0:     If Not B Then Exit Property
    End With
End Property
' ^ ############################## ^ '    IFHeader    ' ^ ############################## ^ '

' v ############################## v '      IFD       ' v ############################## v '
Public Function ReadIFD(ByRef this As IFD, _
                        ByVal ebr As FilEBinReader, _
                        ByVal OffsetIFD As Long, _
                        ByVal OffsetIFHeader As Long) As Boolean
Try: On Error GoTo Catch
    Dim i As Long
    With this
        .Count = ebr.ReadInt16(OffsetIFD + OffsetIFHeader + 1)
        ReDim .Entries(0 To .Count - 1)
        Dim p As Long
        Dim bl As Long
        For i = 0 To .Count - 1
            .Entries(i).Entry.Tag = ebr.ReadInt16
            .Entries(i).Entry.DataType = ebr.ReadInt16
            .Entries(i).Entry.Count = ebr.ReadInt32
            .Entries(i).Entry.ValueOffset = ebr.ReadInt32
        Next
        .OffsetNextIFD = ebr.ReadInt32
        For i = 0 To .Count - 1
            ReadIFD = ReadIFDEntryValue(.Entries(i), ebr, OffsetIFHeader)
        Next
        ReadIFD = True
    End With
    Exit Function
Catch: ErrHandler "ReadIFD", this.OffsetNextIFD
End Function

Public Property Get IFD_ValueByTag(this As IFD, ByVal aTag As TagIF) As Variant
Try: On Error GoTo Catch
    Dim i As Long
    With this
        For i = 0 To .Count - 1
            If .Entries(i).Entry.Tag = aTag Then
                IFD_ValueByTag = IFDEntryValue_GetValue(.Entries(i))
                Exit Property
            End If
        Next
    End With
    Exit Property
Catch: ErrHandler "IFD_ValueByTag", """" & CStr(IFD_ValueByTag) & """"
End Property
' ^ ############################## ^ '      IFD       ' ^ ############################## ^ '

' v ############################## v ' IFDEntryValue  ' v ############################## v '
Public Function ReadIFDEntryValue(ByRef this As IFDEntryValue, _
                                  ByVal ebr As FilEBinReader, _
                                  ByVal OffsetIFDHeader As Long) As Boolean
Try: On Error GoTo Catch
    'der Array Entries mit IFEntry wurde schon gelesen jetzt werden die Werte gelesen die nicht in ValueOffset
    'enthalten sind sondern an einem bestimmten Offset (dem ValueOffset) in der Datei liegen.
    'für die verschiedenen Datentypen wird je eine Variable angelegt.
    'der Rational wird als Currency (der im Variant steckt) zurückgegeben, da beide 64Bit haben.
    Dim i As Long
    'Dim v As Variant
    Dim bytval As Byte, intVal As Integer, lngVal As Long, ratVal As Currency, strVal As String
    Dim sngVal As Single, dblVal As Double
    Dim ofs As Long
    With this
        ofs = .Entry.ValueOffset + OffsetIFDHeader + 1
        Select Case .Entry.DataType
        Case IFDataType.dtByte, IFDataType.dtSByte, IFDataType.dtUndefined2
                                ' muß nur gelesen werden wenn Length > 4, weil
                                ' zwischen 1 und 4 Byte stehen die Werte schon in ValueOffset
                                If .Entry.Count > 4 Then
                                    'den ersten immer mit ofs lesen
                                    'die anderen gehen dann ohne
                                    'Get FNr, ofs, bytval
                                    'ReDim .Value(0 To .Entry.Count - 1) As Byte
                                    
                                    'hmm komisch, das müßte doch auch viel eifacher so gehen, oder?
                                    ReDim bytes(0 To .Entry.Count - 1) As Byte
                                    ebr.ReadBytBuffer bytes, ofs
                                    .Value = bytes
                                    
                                    '.Value(0) = bytval
                                    'For i = 1 To .Entry.Count - 1
                                    '    Get aFNr, , bytval
                                    '    .Value(i) = bytval
                                    'Next
                                End If
                                
        Case IFDataType.dtASCII:
                                ' muß nur gelesen werden wenn Length > 4, weil
                                ' zwischen 1 und 4 Character stehen die Werte schon in ValueOffset
                                If .Entry.Count > 4 Then
                                    'strVal = Space$(.Entry.Count)
                                    'Get FNr, ofs, strVal
                                    'strVal = StrConv(strVal, vbUnicode)
                                    strVal = ebr.ReadString(.Entry.Count, ofs)
                                    Dim p As Long
                                    p = InStr(1, strVal, vbNullChar, vbBinaryCompare)
                                    If p > 0 Then
                                        .Value = Left$(strVal, p - 1)
                                    Else
                                        .Value = Trim$(strVal)
                                    End If
                                End If
        Case IFDataType.dtShort, IFDataType.dtSShort:
                                ' muß nur gelesen werden wenn Length > 2, weil
                                ' zwischen 1 und 2 Shorts stehen die Werte schon in ValueOffset
                                If .Entry.Count > 2 Then
                                    'Get FNr, ofs, intVal
                                    ReDim Integers(0 To .Entry.Count - 1) As Integer
                                    'ReDim .Value(0 To .Entry.Count - 1) As Integer
                                    Integers(0) = ebr.ReadInt16(ofs)
                                    For i = 1 To .Entry.Count - 1
                                        Integers(i) = ebr.ReadInt16
                                    Next
                                    '.Value(0) = intVal
                                    'For i = 1 To .Entry.Count - 1
                                    '    Get aFNr, , intVal
                                    '    .Value(i) = intVal
                                    'Next
                                    .Value = Integers
                                End If
        Case dtLong, dtSLong:
                                'den Long nur lesen falls Length > 1
                                If .Entry.Count > 1 Then
                                    'Get FNr, ofs, lngVal
                                    ReDim Longs(0 To .Entry.Count - 1) As Long
                                    Longs(0) = ebr.ReadInt32(ofs)
                                    For i = 1 To .Entry.Count - 1
                                        'Get FNr, , lngVal
                                        '.Value(i) = lngVal
                                        Longs(i) = ebr.ReadInt32
                                    Next
                                    .Value = Longs
                                End If
        Case dtRational, dtSRational
                                'der Rational muß immer vom Offset gelesen werden
                                Dim rat As IFRational
                                'Get FNr, ofs, ratVal
                                rat.Numerator = ebr.ReadInt32(ofs)
                                rat.Denominator = ebr.ReadInt32
                                GetMem8 rat, ratVal
                                If .Entry.Count = 1 Then
                                    .Value = ratVal
                                Else
                                    ReDim .Value(0 To .Entry.Count - 1) As Currency
                                    .Value(0) = ratVal
                                    For i = 1 To .Entry.Count - 1
                                        rat.Numerator = ebr.ReadInt32
                                        rat.Denominator = ebr.ReadInt32
                                        GetMem8 rat, ratVal
                                        'Get FNr, , ratVal
                                        .Value(i) = ratVal
                                    Next
                                End If
        Case dtDouble:
                                'der Double muß immer vom Offset gelesen werden
                                'Get FNr, ofs, dblVal
                                dblVal = ebr.ReadDouble(ofs)
                                If .Entry.Count = 1 Then
                                    .Value = dblVal
                                Else
                                    ReDim .Value(0 To .Entry.Count - 1) As Double
                                    .Value(0) = dblVal
                                    For i = 1 To .Entry.Count - 1
                                        'Get FNr, , dblVal
                                        dblVal = ebr.ReadDouble
                                        .Value(i) = dblVal
                                    Next
                                End If
        Case dtFloat:
                                'den Single nur lesen falls Length > 1
                                If .Entry.Count > 1 Then
                                    'Get FNr, ofs, sngVal
                                    sngVal = ebr.ReadSingle(ofs)
                                    ReDim .Value(0 To .Entry.Count - 1) As Single
                                    .Value(0) = sngVal
                                    For i = 1 To .Entry.Count - 1
                                        'Get FNr, , sngVal
                                        sngVal = ebr.ReadSingle
                                        .Value(i) = sngVal
                                    Next
                                End If
        End Select
    End With
    ReadIFDEntryValue = True
    Exit Function
Catch: ErrHandler "ReadIFDEntryValue", "this.Entry.Count: " & CStr(this.Entry.Count)
End Function

Public Function IFDEntryValue_GetValue(this As IFDEntryValue) As Variant
Try: On Error GoTo Catch
    Dim dt As IFDataType
    Dim RetVarVal As Variant
    Dim i As Long
    With this
        dt = .Entry.DataType
        Select Case dt
        Case IFDataType.dtASCII
            If .Entry.Count <= 4 Then
                'den string aus ValueOffset rauslesen
                Dim s As String: s = Space$(.Entry.Count)
                Call CopyMem(ByVal StrPtr(s), .Entry.ValueOffset, .Entry.Count)
                s = StrConv(s, vbUnicode)
                Dim p As Long
                p = InStr(1, s, vbNullChar, vbBinaryCompare)
                If p > 0 Then
                    s = Left$(s, p - 1)
                Else
                    s = Trim$(s)
                End If
                RetVarVal = s
            Else
                RetVarVal = .Value
            End If
        'Länge 1
        Case IFDataType.dtByte, IFDataType.dtSByte, IFDataType.dtUndefined2
            If .Entry.Count = 1 Then
                'nur einen Wert übergeben
                RetVarVal = CByte(.Value)
            ElseIf .Entry.Count <= 4 Then
                'das Array erst erzeugen und die einzelnen Elemente aus dem ValueOffset rauslesen
                ReDim B(0 To .Entry.Count - 1) As Byte
                Call CopyMem(ByVal VarPtr(B(0)), .Entry.ValueOffset, .Entry.Count)
                RetVarVal = B 'v
            Else 'if .Entry.Count
                'es wird ein Array von Daten übergeben
                RetVarVal = .Value
            End If
        'Länge 2
        Case IFDataType.dtShort, IFDataType.dtSShort ', IFDataType.dtUndefined2
            If .Entry.Count = 1 Then
                If dt = IFDataType.dtSShort Then
                    RetVarVal = CInt(.Entry.ValueOffset)
                Else
                    RetVarVal = CLng(.Entry.ValueOffset)
                End If
            ElseIf .Entry.Count = 2 Then
                'jetzt auch ein Array von Daten übergeben
                'die allerdings erst aus dem Long rausgelesen werden müssen
                ReDim v(0 To 1) As Integer
                Call GetMem2(ByVal VarPtr(.Entry.ValueOffset), v(0))
                Call GetMem2(ByVal VarPtr(.Entry.ValueOffset) + 2, v(1))
                RetVarVal = v ' .Entry.Value
            Else
                'es wird ein Array von Daten übergeben
                RetVarVal = .Value
            End If
        'Länge 4
        Case IFDataType.dtFloat, IFDataType.dtLong, IFDataType.dtSLong
            If .Entry.Count = 1 Then
                RetVarVal = .Entry.ValueOffset
            Else
                RetVarVal = .Value
            End If
        'Länge 8
        Case IFDataType.dtRational, IFDataType.dtSRational, IFDataType.dtDouble
            If .Entry.Count = 1 Then
                'es wird ein Array von Daten übergeben
                RetVarVal = .Value
            Else
                'es wird ein einzelner Wert übergeben
                RetVarVal = .Value
            End If
        End Select
    End With
    IFDEntryValue_GetValue = RetVarVal
    Exit Function
Catch: ErrHandler "IFDEntryValue_GetValue", RetVarVal
End Function
Public Function IFDEntryValue_IsEqual(this As IFDEntryValue, other As IFDEntryValue) As Boolean
    With this
        '.Entry
    End With
End Function
' ^ ############################## ^ ' IFDEntryValue  ' ^ ############################## ^ '


' v ############################## v '   all ToStr    ' v ############################## v '
Public Function IFRational_ToStr(ByVal v As Variant) As String
    Dim r As IFRational: Call GetMem8(ByVal VarPtr(v) + 8, r)
    IFRational_ToStr = CStr(r.Numerator) & "/" & CStr(r.Denominator)
End Function
Public Function IFD_ToStr(this As IFD, Optional ByVal Index As Long) As String
Try: On Error GoTo Catch
    Dim i As Long
    Dim s As String, c As String
    Dim dt As IFDataType
    With this
        s = s & "Count: " & CStr(.Count) & vbCrLf
        s = s & " Nr:  Tag-Name                     Tag-ID  Type                Count  Offset  Value" & vbCrLf
        For i = 0 To .Count - 1
            c = CStr(i)
            s = s & Space(3 - Len(c)) & c & ": "
            s = s & IFDEntryValue_ToStr(.Entries(i)) & vbCrLf
        Next
        s = s & "OffsetNextIFD: " & CStr(.OffsetNextIFD) & vbCrLf
    End With
    IFD_ToStr = s
    Exit Function
Catch: ErrHandler "IFD_ToStr", s
End Function
Public Function IFDEntryValue_ToStr(this As IFDEntryValue) As String
Try: On Error GoTo Catch
    Dim s As String, c As String
    Dim dt As IFDataType
    With this
        With .Entry
            c = "&H" & Hex$(.Tag)
            s = s & " " & TagIF_ToStr(.Tag) & Space(7 - Len(c)) & c & "  "
            s = s & IFDataType_ToStr(.DataType) & "  "
            c = CStr(.Count)
            s = s & Space(6 - Len(c)) & c & "    "
            dt = .DataType
        End With
        'in zwei Schritten zuerst ob Offset geschrieben werden soll
        c = vbNullString
        Select Case dt
        Case IFDataType.dtASCII, IFDataType.dtByte, IFDataType.dtSByte, IFDataType.dtUndefined2
            If .Entry.Count > 4 Then
                c = CStr(.Entry.ValueOffset)
            End If
        Case IFDataType.dtShort, IFDataType.dtSShort
            If .Entry.Count > 2 Then
                c = CStr(.Entry.ValueOffset)
            End If
        Case IFDataType.dtFloat, IFDataType.dtLong, IFDataType.dtSLong
            If .Entry.Count > 1 Then
                c = CStr(.Entry.ValueOffset)
            End If
        Case IFDataType.dtRational, IFDataType.dtSRational, IFDataType.dtDouble
            c = CStr(.Entry.ValueOffset)
        End Select
        'dann den Wert dazuschreiben
        s = s & Space(4 - Len(c)) & c & "  "
        Dim v As Variant
        v = IFDEntryValue_GetValue(this)
        If IsArray(v) Then
            s = s & IFValueArray_ToStr(dt, v)
        Else
            Select Case dt
            Case IFDataType.dtRational, IFDataType.dtSRational
                s = s & IFRational_ToStr(v)
            Case IFDataType.dtASCII
                s = s & """" & v & """"
            Case Else
                s = s & CStr(v)
            End Select
        End If
    End With
    IFDEntryValue_ToStr = s
    Exit Function
Catch: ErrHandler "IFDEntryValue_ToStr", s
End Function
Public Function IFValueArray_ToStr(ByVal dt As IFDataType, ByRef vArr As Variant) As String
Try: On Error GoTo Catch
    Dim s As String
    Dim i As Long
    If IsArray(vArr) Then
        s = "{"
        If dt = dtRational Or dt = dtSRational Then
            For i = LBound(vArr) To UBound(vArr) - 1
                s = s & IFRational_ToStr(vArr(i)) & "; "
            Next
            s = s & IFRational_ToStr(vArr(i)) & "}"
        Else
            '{0; 0; 0; ...; 0}
            Dim u0 As Long: u0 = UBound(vArr)
            Const ShowMaxBytes As Long = 32
            Dim u As Long: u = Min(ShowMaxBytes, u0)
            For i = LBound(vArr) To u - 1
                s = s & CStr(vArr(i)) & "; "
            Next
            If u0 > 32 Then
                s = s & "...;"
            End If
            s = s & CStr(vArr(u0)) & "}"
        End If
    End If
    IFValueArray_ToStr = s
    Exit Function
Catch: ErrHandler "IFValueArray_ToStr", s
End Function
Public Function IFDataType_ToStr(ByVal this As IFDataType) As String
    Dim s As String
    Select Case this
    Case dtByte:       s = "Byte"
    Case dtASCII:      s = "ASCII"
    Case dtShort:      s = "Short"
    Case dtLong:       s = "Unsigned Long"
    Case dtRational:   s = "Unsigned Rational"
    Case dtSLong:      s = "Signed Long"
    Case dtSRational:  s = "Signed Rational"
    Case dtFloat:      s = "Float (Single)"
    Case dtDouble:     s = "Double"
    Case dtSByte:      s = "Signed Byte"
    Case dtSShort:     s = "Signed Short"
    Case dtUndefined2: s = "Undefined (Byte)"
    Case Else:         s = "Undefined"
    End Select
    IFDataType_ToStr = s & Space(17 - Len(s))
End Function

'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal FncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional ByVal bLoud As Boolean = True, _
                            Optional ByVal bErrLog As Boolean = False, _
                            Optional ByVal vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    ErrHandler = MError.ErrHandler("MExif", FncName, AddInfo, bLoud, bErrLog, vbDecor)
End Function
'
'Private Function PErrHandler(ByVal FncName As String, ByVal AddErrMsg As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
'    PErrHandler = MError.ErrHandler("MExif", FncName, AddErrMsg, Buttons)
'End Function
