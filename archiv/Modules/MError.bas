Attribute VB_Name = "MError"
Option Explicit
Public Errors As Collection

Public Function ErrHandler(ObjOrModName, FncName As String, _
                            Optional AddErrMsg As String, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean, _
                            Optional Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
    If IsObject(ObjOrModName) Then ObjOrModName = TypeName(ObjOrModName)
    Dim msg As String
    msg = "An error occured in " & ObjOrModName & "::" & FncName & vbCrLf & _
           Err.Number & " " & Err.Description & _
           IIf(Len(AddErrMsg), vbCrLf & AddErrMsg, "")
    If bLoud Then
        ErrHandler = MsgBox(msg, Buttons Or vbCritical)
    End If
    If bErrLog Then
        If Errors Is Nothing Then Set Errors = New Collection
        Errors.Add msg
    End If
End Function

Public Property Get LastError() As String
    If Errors Is Nothing Then Exit Property
    If Errors.count Then
        LastError = Errors.Item(Errors.count)
    End If
End Property
