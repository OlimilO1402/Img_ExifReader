Attribute VB_Name = "MByteSwapper"
Option Explicit
' Ein SafeArray-Descriptor dient in VB als ein universaler Zeiger
' ein Lightweight-Object ganz ohne Interface
Private Type TUDTPtr
    pSA        As Long
    Reserved   As Long ' z.B. für vbVarType oder IRecordInfo
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements  As Long
    lLBound    As Long
End Type

Private Enum SAFeature
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4

    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20
    FADF_HAVEIID = &H40
    FADF_HAVEVARTYPE = &H80

    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800
    FADF_RESERVED = &HF008
End Enum

' die TByteSwapper Lightweight-Object Struktur
Public Type TByteSwapper
    pB      As TUDTPtr
    tmpByte As Byte
    B()     As Byte
End Type
Private Type TSafeArrayPtr
    pSAPtr As TUDTPtr
    pSA()  As TUDTPtr
End Type
Private Declare Sub PutMem4 Lib "msvbvm60" ( _
    ByRef pDst As Any, _
    ByVal Src As Long)

Private Declare Sub GetMem4 Lib "msvbvm60" ( _
    ByRef pSrc As Any, _
    ByRef pDst As Any)

Public Declare Function ArrPtr Lib "msvbvm60" _
                        Alias "VarPtr" ( _
                        ByRef pArr() As Any) As Long

Public Declare Sub SBO_Rotate2 Lib "SwapByteOrder.dll" _
             Alias "SwapByteOrder16" (ByRef Ptr As Any)
             
Public Declare Sub SBO_Rotate4 Lib "SwapByteOrder.dll" _
             Alias "SwapByteOrder32" (ByRef Ptr As Any)
             
Public Declare Sub SBO_Rotate8 Lib "SwapByteOrder.dll" _
             Alias "SwapByteOrder64" (ByRef Ptr As Any)
             
Public Declare Function SBO_RotateArray Lib "SwapByteOrder.dll" _
             Alias "SwapByteOrderArray" (ByRef Value() As Any) As Long

Public Declare Sub SBO_RotateUDTArray Lib "SwapByteOrder.dll" _
             Alias "SwapByteOrderUDTArray" ( _
             ByRef Arr() As Any, ByRef udtDescription() As Integer)

Private Sub New_UDTPtr(ByRef this As TUDTPtr, _
                       ByVal Feature As SAFeature, _
                       ByVal bytesPerElement As Long, _
                       Optional ByVal CountElements As Long = 1, _
                       Optional ByVal lLBound As Long = 0)
    ' erzeugt ein neues UDTPtr-Lightweight-Object
    ' nur als Sub wegen VarPtr(cDims)
    With this
        .pSA = VarPtr(.cDims)
        .cDims = 1
        .cbElements = bytesPerElement
        .fFeatures = CInt(Feature)
        .cElements = CountElements
        .lLBound = lLBound
    End With
    
End Sub
Private Sub New_SafeArrayPtr(this As TSafeArrayPtr)
    ' erzeugt ein neues SafeArrayPtr-Lightweight-Object
    With this
        Call New_UDTPtr(.pSAPtr, SAFeature.FADF_EMBEDDED Or SAFeature.FADF_STATIC Or SAFeature.FADF_RECORD, LenB(.pSAPtr))
        Call PutMem4(ByVal ArrPtr(.pSA), .pSAPtr.pSA)
    End With
End Sub
Private Sub DeleteSafeArrayPtr(this As TSafeArrayPtr)
    ' löscht ein SafeArrayPtr-Lightweight-Object
    With this
        Call PutMem4(ByVal ArrPtr(.pSA), 0)
    End With
End Sub
Private Property Let SAPtr(this As TSafeArrayPtr, ByVal RHS As Long)
    ' schreibt den Zeiger auf eine SafeArrayDescriptor-Struktur in
    ' SafeArrayPtr-Lightweight-Object
    Dim p As Long
    Call GetMem4(ByVal RHS, p)
    this.pSAPtr.pvData = p - 8
    ' -8 weil zuerst pSA und Reserved und dann kommt erst
    ' der Anfang der SafeArrayDesc-Struktur mit cDims
End Property
Private Property Get VarSAPtr(ByRef vArr As Variant) As Long
    Call PutMem4(VarSAPtr, VarPtr(vArr) + 8)
End Property

Public Sub New_ByteSwapper(this As TByteSwapper, Optional ByVal CountBytes As Long = 2)
    ' erzeugt ein neues ByteSwapper Lightweight-Object
    With this
        Call New_UDTPtr(.pB, FADF_EMBEDDED Or FADF_STATIC, 1, CountBytes)
        Call PutMem4(ByVal ArrPtr(.B), .pB.pSA)
    End With
End Sub

Public Sub DeleteByteSwapper(this As TByteSwapper)
    'löscht den Zeiger im Array der ByteSwapper-Struktur
    With this
        Call PutMem4(ByVal ArrPtr(.B), 0)
    End With
End Sub

Public Sub Rotate2(this As TByteSwapper)
    ' vertauscht (rotiert) zwei Bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(1)
        .B(1) = .tmpByte
    End With
End Sub
Public Sub Rotate4(this As TByteSwapper)
    ' vertauscht vier Bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(3)
        .B(3) = .tmpByte
        
        .tmpByte = .B(1)
        .B(1) = .B(2)
        .B(2) = .tmpByte
    End With
End Sub
Public Sub Rotate8(this As TByteSwapper)
    ' vertauscht acht Bytes
    With this
        .tmpByte = .B(0)
        .B(0) = .B(7)
        .B(7) = .tmpByte
        
        .tmpByte = .B(1)
        .B(1) = .B(6)
        .B(6) = .tmpByte
        
        .tmpByte = .B(2)
        .B(2) = .B(5)
        .B(5) = .tmpByte
        
        .tmpByte = .B(3)
        .B(3) = .B(4)
        .B(4) = .tmpByte
    End With
End Sub
Public Sub Rotate(this As TByteSwapper)
    ' Rotiert die Bytes in Abhängigkeit der Anzahl der Bytes
    ' bzw der Größe der Variable in Bytes
    With this
        Select Case .pB.cElements
        Case 2
            .tmpByte = .B(0)
            .B(0) = .B(1)
            .B(1) = .tmpByte
        Case 4
            .tmpByte = .B(0)
            .B(0) = .B(3)
            .B(3) = .tmpByte
            
            .tmpByte = .B(1)
            .B(1) = .B(2)
            .B(2) = .tmpByte
        Case 8
            .tmpByte = .B(0)
            .B(0) = .B(7)
            .B(7) = .tmpByte
            
            .tmpByte = .B(1)
            .B(1) = .B(6)
            .B(6) = .tmpByte
            
            .tmpByte = .B(2)
            .B(2) = .B(5)
            .B(5) = .tmpByte
            
            .tmpByte = .B(3)
            .B(3) = .B(4)
            .B(4) = .tmpByte
        End Select
    End With
End Sub

Public Sub RotateArray(this As TByteSwapper, vArr)
    ' Rotiert die Bytes der Elemente eines beliebigen Arrays
    ' das in dem Variant vArr übergeben wird.
    ' Das Array kann vom Typ Integer, Long, Currency, Single oder Double sein.
    ' soll stattdessen ein Array aus UD-Type Elementen behandelt werden so kann
    ' die Funktion RotateUDTArray (siehe unten) verwendet werden .
    If Not IsArray(vArr) Then Exit Sub
    Dim pSA As TSafeArrayPtr: Call New_SafeArrayPtr(pSA)
    SAPtr(pSA) = VarSAPtr(vArr)
    Dim i  As Long
    Dim p  As Long
    Dim pc As Long
    Dim ub As Long: ub = UBound(vArr)
    Dim lb As Long: lb = LBound(vArr)
    Dim cnt As Long: cnt = ub - lb + 1
    If cnt > 0 Then
        With this
            If .pB.pvData = 0 Then
                .pB.pvData = pSA.pSA(0).pvData 'VarPtr(vArr(0))
            End If
            .pB.cElements = pSA.pSA(0).cbElements 'LenB(vArr(0))
            Select Case .pB.cbElements
            Case 2
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 2
                    .tmpByte = .B(0)
                    .B(0) = .B(1)
                    .B(1) = .tmpByte
                Next
            Case 4
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 4
                    .tmpByte = .B(0)
                    .B(0) = .B(3)
                    .B(3) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(2)
                    .B(2) = .tmpByte
                Next
            Case 8
                For i = lb To ub
                    .pB.pvData = .pB.pvData + 8
                    .tmpByte = .B(0)
                    .B(0) = .B(7)
                    .B(7) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(6)
                    .B(6) = .tmpByte
                    
                    .tmpByte = .B(2)
                    .B(2) = .B(5)
                    .B(5) = .tmpByte
                    
                    .tmpByte = .B(3)
                    .B(3) = .B(4)
                    .B(4) = .tmpByte
                Next
            End Select
        End With
    End If
    Call DeleteSafeArrayPtr(pSA)
End Sub

'Eine entsprechende Funktion in der SwapByteOrderDll könnte in etwa so aussehen
'wie diese Funktion
'allerdings statt pData und Count könnte das Array direkt angegeben werden,
'As Any machts möglich.
'eine entsprechende Deklaration könnte so aussehen:
'Public Declare Sub RotateUDTArr Lib "SwapByteOrder.dll" _
'             Alias "SwapByteOrderUDTArray" ( _
'             ByRef ArrayOfUDType() As Any, ByRef udtDescription() As Integer)

Public Sub RotateUDTArray(this As TByteSwapper, _
                          ByVal pData As Long, _
                          ByVal Count As Long, _
                          ByRef udtDescription() As Integer)
    ' Rotiert die Elemente eines Array vom Typ eines beliebigen UD-Types
    ' this:  der ByteSwapper
    ' pData: der Zeiger auf das erste Element im Array (verwende VarPtr())
    ' Count: die Anzahl der Elemente im Array
    ' udtDescription(): liefert eine Beschreibung des UD-Types.
    '                   Der Wert der Integer-Elemente im Array repräsentiert
    '                   die Größe der einzelnen Variablen-Elemente des UD-Types
    '                   verwende dazu die Funktion LenB.
    '                   Variablen des UD-Types die nicht gedreht werden sollen,
    '                   müssen negativ angegeben werden.
    '                   Achtung: es müssen auch Padbytes berücksichtigt werden
    '
    Dim i As Long, j As Long
    Dim CountUDTElements As Long: CountUDTElements = UBound(udtDescription) + 1
    Dim udtLength As Long, ValLength As Long
    For i = 0 To CountUDTElements - 1
        udtLength = udtLength + Abs(udtDescription(i))
    Next
    With this
        .pB.pvData = pData
        For i = 0 To Count - 1
            For j = 0 To CountUDTElements - 1
                ValLength = udtDescription(j)
                .pB.cElements = Abs(ValLength)
                Select Case ValLength
                Case 2
                    .tmpByte = .B(0)
                    .B(0) = .B(1)
                    .B(1) = .tmpByte
                Case 4
                    .tmpByte = .B(0)
                    .B(0) = .B(3)
                    .B(3) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(2)
                    .B(2) = .tmpByte
                Case 8
                    .tmpByte = .B(0)
                    .B(0) = .B(7)
                    .B(7) = .tmpByte
                    
                    .tmpByte = .B(1)
                    .B(1) = .B(6)
                    .B(6) = .tmpByte
                    
                    .tmpByte = .B(2)
                    .B(2) = .B(5)
                    .B(5) = .tmpByte
                    
                    .tmpByte = .B(3)
                    .B(3) = .B(4)
                    .B(4) = .tmpByte
                End Select
                .pB.pvData = .pB.pvData + .pB.cElements 'Abs(ValLength)
            Next
        Next
    End With
End Sub

Public Function SwapBytesInt16(i As Integer) As Integer
    SBO_Rotate2 i: SwapBytesInt16 = i
End Function
Public Function SwapBytesInt32(i As Long) As Long
    SBO_Rotate4 i: SwapBytesInt32 = i
End Function

