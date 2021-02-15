VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EBinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_st As Stream
Private m_se As Endianness
Private m_en As Endianness

Friend Sub New_(st As Stream)
    Set m_st = st: m_st.OOpen
    m_se = MMain.SystemEndianness
End Sub

Public Property Get ReadEndianness() As Endianness
    ReadEndianness = m_en
End Property
Public Property Let ReadEndianness(ByVal Value As Endianness)
    m_en = Value
    MsgBox Endianness_ToStr(m_en)
End Property

Public Function ReadInt16(Optional ByVal pos As Long = -1) As Integer
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(ReadInt16), 2
    If m_se = m_en Then Exit Function
    MByteSwapper.SBO_Rotate2 ReadInt16
    Exit Function
Catch: ErrHandler "ReadInt16"
End Function

Public Function ReadInt32(Optional ByVal pos As Long = -1) As Long
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(ReadInt32), 4
    If m_se = m_en Then Exit Function
    MByteSwapper.SBO_Rotate4 ReadInt32
    Exit Function
Catch: ErrHandler "ReadInt32"
End Function

Public Function ReadCur64(Optional ByVal pos As Long = -1) As Currency
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(ReadCur64), 8
    If m_se = m_en Then Exit Function
    MByteSwapper.SBO_Rotate8 ReadCur64
    Exit Function
Catch: ErrHandler "ReadCur64"
End Function

Public Function ReadSingle(Optional ByVal pos As Long = -1) As Single
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(ReadSingle), 4
    If m_se = m_en Then Exit Function
    MByteSwapper.SBO_Rotate4 ReadSingle
    Exit Function
Catch: ErrHandler "ReadSingle"
End Function

Public Function ReadDouble(Optional ByVal pos As Long = -1) As Double
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(ReadDouble), 8
    If m_se = m_en Then Exit Function
    MByteSwapper.SBO_Rotate8 ReadDouble
    Exit Function
Catch: ErrHandler "ReadDouble"
End Function

Public Function ReadString(ByVal Length As Long, Optional ByVal pos As Long = -1) As String
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    ReDim bytbuffer(0 To Length - 1) As Byte
    'ReadString = String$(Length, vbNullChar)
    m_st.Read VarPtr(bytbuffer(0)), Length
    ReadString = StrConv(bytbuffer, vbUnicode)
    Exit Function
Catch: ErrHandler "ReadString"
End Function

Public Sub ReadByteBuf(inout_bytBuffer() As Byte, Optional ByVal pos As Long = -1)
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read VarPtr(inout_bytBuffer(0)), UBound(inout_bytBuffer) + 1
    Exit Sub
Catch: ErrHandler "ReadByteBuf"
End Sub

Public Sub ReadBytes(ByVal pDst As Long, ByVal BytLength As Long, Optional ByVal pos As Long = -1)
Try: On Error GoTo Catch
    If pos >= 0 Then m_st.Position = pos
    m_st.Read pDst, BytLength
    Exit Sub
Catch: ErrHandler "ReadBytes"
End Sub

Public Property Get BaseStream() As Stream
    Set BaseStream = m_st
End Property

'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal FncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional ByVal bLoud As Boolean = True, _
                            Optional ByVal bErrLog As Boolean = False, _
                            Optional ByVal vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    ErrHandler = MError.ErrHandler(Me, FncName, AddInfo, bLoud, bErrLog, vbDecor)
End Function
