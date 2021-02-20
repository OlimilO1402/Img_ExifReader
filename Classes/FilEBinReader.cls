VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilEBinReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_FNr As Integer
Private m_se As EEndianness
Private m_en As EEndianness

Private Declare Sub SwapByteOrder16 Lib "SwapByteOrder.dll" (ByRef Ptr As Any)
Private Declare Sub SwapByteOrder32 Lib "SwapByteOrder.dll" (ByRef Ptr As Any)
Private Declare Sub SwapByteOrder64 Lib "SwapByteOrder.dll" (ByRef Ptr As Any)
Private Declare Function SwapByteOrderArray Lib "SwapByteOrder.dll" (ByRef Value() As Any) As Long
Private Declare Sub SwapByteOrderUDTArray Lib "SwapByteOrder.dll" (ByRef Arr() As Any, ByRef udtDescription() As Integer)

Friend Sub New_(PathFileName As String)
    m_FNr = OOpen(PathFileName)
    m_se = MMain.SystemEndianness
End Sub

Private Sub Class_Terminate()
    CClose
End Sub

Public Function OOpen(PathFileName As String) As Integer
Try: On Error GoTo Catch
    OOpen = FreeFile
    Open PathFileName For Binary Access Read As OOpen
Catch:
End Function

Public Sub CClose()
    Close m_FNr
End Sub

Public Property Get Endianness() As EEndianness
    Endianness = m_en
End Property
Public Property Let Endianness(ByVal Value As EEndianness)
    m_en = Value
    Debug.Print Endianness_ToStr(m_en)
End Property

Public Property Let Position(ByVal Value As Long)
Try: On Error GoTo Catch
    Seek m_FNr, Value
    Exit Property
Catch: ErrHandler "Let_Position"
End Property
Public Property Get Position() As Long
Try: On Error GoTo Catch
    Position = Seek(m_FNr)
    Exit Property
Catch: ErrHandler "Get_Position"
End Property

Public Function ReadInt16(Optional ByVal pos As Long = -1) As Integer
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , ReadInt16
    If m_se = m_en Then Exit Function
    SwapByteOrder16 ReadInt16
    Exit Function
Catch: ErrHandler "ReadInt16"
End Function

Public Function ReadInt32(Optional ByVal pos As Long = -1) As Long
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , ReadInt32
    If m_se = m_en Then Exit Function
    SwapByteOrder32 ReadInt32
    Exit Function
Catch: ErrHandler "ReadInt32"
End Function

Public Function ReadCur64(Optional ByVal pos As Long = -1) As Currency
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , ReadCur64
    If m_se = m_en Then Exit Function
    SwapByteOrder64 ReadCur64
    Exit Function
Catch: ErrHandler "ReadCur64"
End Function

Public Function ReadSingle(Optional ByVal pos As Long = -1) As Single
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , ReadSingle
    If m_se = m_en Then Exit Function
    SwapByteOrder32 ReadSingle
    Exit Function
Catch: ErrHandler "ReadSingle"
End Function

Public Function ReadDouble(Optional ByVal pos As Long = -1) As Double
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , ReadDouble
    If m_se = m_en Then Exit Function
    SwapByteOrder64 ReadDouble
    Exit Function
Catch: ErrHandler "ReadDouble"
End Function

Public Function ReadString(ByVal Length As Long, Optional ByVal pos As Long = -1) As String
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    'ReDim bytbuffer(0 To Length - 1) As Byte
    ReadString = String$(Length, vbNullChar)
    Get m_FNr, , ReadString
    'ReadString = StrConv(bytbuffer, vbUnicode)
    Exit Function
Catch: ErrHandler "ReadString"
End Function

Public Sub ReadBytBuffer(inout_bytBuffer() As Byte, Optional ByVal pos As Long = -1)
Try: On Error GoTo Catch
    If pos >= 0 Then Me.Position = pos
    Get m_FNr, , inout_bytBuffer
    Exit Sub
Catch: ErrHandler "ReadBytBuffer"
End Sub


'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal FncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional ByVal bLoud As Boolean = True, _
                            Optional ByVal bErrLog As Boolean = False, _
                            Optional ByVal vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    ErrHandler = MError.ErrHandler(Me, FncName, AddInfo, bLoud, bErrLog, vbDecor)
End Function

