VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Stream
Public Enum FileMode
    fmAppend
    fmBinary
    fmInput
    fmOutput
    fmRandom
End Enum
Public Enum FileAccess
    faNone
    faRead = 1
    faWrite = 2
    faReadWrite = 3 'Read Or Write 'ist die Standardeinstellung
End Enum
Public Enum FileShare
    fsNone = 0
    fsLockRead = 1
    fsLockWrite = 2
    fsLockReadWrite = 3
    fsShared = 16
End Enum
Private Enum ReadKind
    ReadInput      'Read one thing
    ReadLineInput  'Read one line
    ReadInputFkt   'Read the whole stuff to the end
End Enum
Private m_FNm     As String
Private m_FNr     As Integer
Private m_FMode   As FileMode
Private m_FAccess As FileAccess
Private m_FShare  As FileShare
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal BytLen As Long)

Friend Sub New_(aPathFileName As String, Optional FMode As FileMode = fmRandom, Optional FAccess As FileAccess = faReadWrite, Optional FShare As FileShare = fsNone)
    m_FNm = aPathFileName: m_FMode = FMode: m_FAccess = FAccess: m_FShare = FShare
End Sub

Private Sub Class_Terminate()
    Stream_Flush
    Stream_CClose
End Sub

Private Property Get Stream_IsEOF() As Boolean
    Stream_IsEOF = EOF(m_FNr)
End Property

Private Sub Stream_Read(ByVal pDst As Long, ByVal BytLength As Long)
Try: On Error GoTo Catch
    If pDst = 0 Or BytLength = 0 Then Exit Sub
    ReDim bytbuffer(0 To BytLength - 1) As Byte
    Get m_FNr, , bytbuffer
    RtlMoveMemory ByVal pDst, bytbuffer(0), BytLength
    Exit Sub
Catch: ErrHandler "Read"
End Sub
Private Sub Stream_WWrite(ByVal pSrc As Long, ByVal BytLength As Long)
Try: On Error GoTo Catch
    If pSrc = 0 Or BytLength = 0 Then Exit Sub
    ReDim bytbuffer(0 To BytLength - 1) As Byte
    RtlMoveMemory bytbuffer(0), ByVal pSrc, BytLength
    Put m_FNr, , bytbuffer(0)
    Exit Sub
Catch: ErrHandler "Write"
End Sub

Private Property Let Stream_Position(ByVal Value As Long)
Try: On Error GoTo Catch
    Seek m_FNr, Value
    Exit Property
Catch: ErrHandler "Let_Position"
End Property
Private Property Get Stream_Position() As Long
Try: On Error GoTo Catch
    Stream_Position = Seek(m_FNr)
    Exit Property
Catch: ErrHandler "Get_Position"
End Property

Private Property Get Stream_Length() As Long
    If m_FNr Then Stream_Length = LOF(m_FNr) Else Stream_Length = FileLen(m_FNm)
End Property
Private Sub Stream_Flush()
End Sub
Private Sub Stream_CClose()
    Close m_FNr
    m_FNr = 0
End Sub

Private Sub Stream_OOpen()
    If m_FNr Then Exit Sub 'nur einmal öffnen
Try: On Error GoTo Catch
    m_FNr = FreeFile
    Select Case m_FMode
    Case FileMode.fmAppend
        Select Case m_FAccess
        Case FileAccess.faNone
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Append As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Append Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Append Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Append Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Append Shared As m_FNr
            End Select
        'Case FileAccess.faRead 'Not Possible
        '    Select Case m_FShare
        '    Case FileShare.fsNone:          Open m_FNm For Append Access Read As m_FNr
        '    Case FileShare.fsLockRead:      Open m_FNm For Append Access Read Lock Read As m_FNr
        '    Case FileShare.fsLockWrite:     Open m_FNm For Append Access Read Lock Write As m_FNr
        '    Case FileShare.fsLockReadWrite: Open m_FNm For Append Access Read Lock Read Write As m_FNr
        '    Case FileShare.fsShared:        Open m_FNm For Append Access Read Shared As m_FNr
        '    End Select
        Case FileAccess.faWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Append Access Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Append Access Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Append Access Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Append Access Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Append Access Write Shared As m_FNr
            End Select
        Case FileAccess.faReadWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Append Access Read Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Append Access Read Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Append Access Read Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Append Access Read Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Append Access Read Write Shared As m_FNr
            End Select
        End Select
    Case FileMode.fmBinary
        Select Case m_FAccess
        Case FileAccess.faNone
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Binary As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Binary Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Binary Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Binary Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Binary Shared As m_FNr
            End Select
        Case FileAccess.faRead
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Binary Access Read As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Binary Access Read Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Binary Access Read Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Binary Access Read Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Binary Access Read Shared As m_FNr
            End Select
        Case FileAccess.faWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Binary Access Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Binary Access Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Binary Access Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Binary Access Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Binary Access Write Shared As m_FNr
            End Select
        Case FileAccess.faReadWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Binary Access Read Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Binary Access Read Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Binary Access Read Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Binary Access Read Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Binary Access Read Write Shared As m_FNr
            End Select
        End Select
    Case FileMode.fmInput
        Select Case m_FAccess
        Case FileAccess.faNone
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Input As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Input Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Input Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Input Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Input Shared As m_FNr
            End Select
        Case FileAccess.faRead
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Input Access Read As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Input Access Read Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Input Access Read Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Input Access Read Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Input Access Read Shared As m_FNr
            End Select
        'Case FileAccess.faWrite 'Not Possible
        '    Select Case m_FShare
        '    Case FileShare.fsNone:          Open m_FNm For Input Access Write As m_FNr
        '    Case FileShare.fsLockRead:      Open m_FNm For Input Access Write Lock Read As m_FNr
        '    Case FileShare.fsLockWrite:     Open m_FNm For Input Access Write Lock Write As m_FNr
        '    Case FileShare.fsLockReadWrite: Open m_FNm For Input Access Write Lock Read Write As m_FNr
        '    Case FileShare.fsShared:        Open m_FNm For Input Access Write Shared As m_FNr
        '    End Select
        'Case FileAccess.faReadWrite 'Not Possible
        '    Select Case m_FShare
        '    Case FileShare.fsNone:          Open m_FNm For Input Access Read Write As m_FNr
        '    Case FileShare.fsLockRead:      Open m_FNm For Input Access Read Write Lock Read As m_FNr
        '    Case FileShare.fsLockWrite:     Open m_FNm For Input Access Read Write Lock Write As m_FNr
        '    Case FileShare.fsLockReadWrite: Open m_FNm For Input Access Read Write Lock Read Write As m_FNr
        '    Case FileShare.fsShared:        Open m_FNm For Input Access Read Write Shared As m_FNr
        '    End Select
        End Select
    Case FileMode.fmOutput
        Select Case m_FAccess
        Case FileAccess.faNone
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Output As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Output Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Output Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Output Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Output Shared As m_FNr
            End Select
        'Case FileAccess.faRead
        '    Select Case m_FShare 'Not Possible
        '    Case FileShare.fsNone:          Open m_FNm For Output Access Read As m_FNr
        '    Case FileShare.fsLockRead:      Open m_FNm For Output Access Read Lock Read As m_FNr
        '    Case FileShare.fsLockWrite:     Open m_FNm For Output Access Read Lock Write As m_FNr
        '    Case FileShare.fsLockReadWrite: Open m_FNm For Output Access Read Lock Read Write As m_FNr
        '    Case FileShare.fsShared:        Open m_FNm For Output Access Read Shared As m_FNr
        '    End Select
        Case FileAccess.faWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Output Access Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Output Access Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Output Access Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Output Access Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Output Access Write Shared As m_FNr
            End Select
        'Case FileAccess.faReadWrite
        '    Select Case m_FShare 'Not Possible
        '    Case FileShare.fsNone:          Open m_FNm For Output Access Read Write As m_FNr
        '    Case FileShare.fsLockRead:      Open m_FNm For Output Access Read Write Lock Read As m_FNr
        '    Case FileShare.fsLockWrite:     Open m_FNm For Output Access Read Write Lock Write As m_FNr
        '    Case FileShare.fsLockReadWrite: Open m_FNm For Output Access Read Write Lock Read Write As m_FNr
        '    Case FileShare.fsShared:        Open m_FNm For Output Access Read Write Shared As m_FNr
        '    End Select
        End Select
    Case FileMode.fmRandom
        Select Case m_FAccess
        Case FileAccess.faNone
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Random As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Random Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Random Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Random Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Random Shared As m_FNr
            End Select
        Case FileAccess.faRead
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Random Access Read As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Random Access Read Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Random Access Read Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Random Access Read Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Random Access Read Shared As m_FNr
            End Select
        Case FileAccess.faWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Random Access Write As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Random Access Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Random Access Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Random Access Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Random Access Write Shared As m_FNr
            End Select
        Case FileAccess.faReadWrite
            Select Case m_FShare
            Case FileShare.fsNone:          Open m_FNm For Random Access Read As m_FNr
            Case FileShare.fsLockRead:      Open m_FNm For Random Access Read Write Lock Read As m_FNr
            Case FileShare.fsLockWrite:     Open m_FNm For Random Access Read Write Lock Write As m_FNr
            Case FileShare.fsLockReadWrite: Open m_FNm For Random Access Read Write Lock Read Write As m_FNr
            Case FileShare.fsShared:        Open m_FNm For Random Access Read Write Shared As m_FNr
            End Select
        End Select
    End Select
    'Stream_OOpen = True
    Exit Sub
Catch: ErrHandler "Open", """" & m_FNm & """" & vbCrLf & "FileMode: " & FileMode_ToStr(m_FMode) & "; FileAccess: " & FileAccess_ToStr(m_FAccess) & "; FileShare: " & FileShare_ToStr(m_FShare)
End Sub

Private Function FileMode_ToStr(e As FileMode) As String
    Dim s As String
    Select Case e
    Case FileMode.fmAppend: s = "Append"
    Case FileMode.fmBinary: s = "Binary"
    Case FileMode.fmInput:  s = "Input"
    Case FileMode.fmOutput: s = "Output"
    Case FileMode.fmRandom: s = "Random"
    End Select
    FileMode_ToStr = s
End Function
Private Function FileAccess_ToStr(e As FileAccess) As String
    Dim s As String
    Select Case e
    Case FileAccess.faNone:      s = "None"
    Case FileAccess.faRead:      s = "Read"
    Case FileAccess.faWrite:     s = "Write"
    Case FileAccess.faReadWrite: s = "ReadWrite"
    End Select
    FileAccess_ToStr = s
End Function
Private Function FileShare_ToStr(e As FileAccess) As String
    Dim s As String
    Select Case e
    Case FileShare.fsNone:          s = "None"
    Case FileShare.fsLockRead:      s = "LockRead"
    Case FileShare.fsLockWrite:     s = "LockWrite"
    Case FileShare.fsLockReadWrite: s = "LockReadWrite"
    Case FileShare.fsShared:        s = "Shared"
    End Select
    FileShare_ToStr = s
End Function
Private Function ReadKind_ToStr(e As FileAccess) As String
    Dim s As String
    Select Case e
    Case ReadKind.ReadInput:     s = "ReadInput"
    Case ReadKind.ReadLineInput: s = "ReadLineInput"
    Case ReadKind.ReadInputFkt:  s = "ReadInputFkt"
    End Select
    ReadKind_ToStr = s
End Function

'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal FncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional ByVal bLoud As Boolean = True, _
                            Optional ByVal bErrLog As Boolean = False, _
                            Optional ByVal vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    ErrHandler = MError.ErrHandler(Me, FncName, AddInfo, bLoud, bErrLog, vbDecor)
End Function

