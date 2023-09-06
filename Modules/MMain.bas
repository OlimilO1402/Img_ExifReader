Attribute VB_Name = "MMain"
Option Explicit

Public Enum EEndianness ' determines ByteOrder; in which order bytes have to be read
    IntelLittleEndian   ' = &H4949
    MotorolaBigEndian   ' = &H4D4D
End Enum


Sub Main()
    FMain.Show
End Sub

Public Property Get SystemEndianness() As EEndianness
    SystemEndianness = EEndianness.IntelLittleEndian
End Property

Public Function Endianness_ToStr(e As EEndianness) As String
    Dim s As String
    Select Case e
    Case EEndianness.IntelLittleEndian: s = "IntelLittleEndian"
    Case EEndianness.MotorolaBigEndian: s = "MotorolaBigEndian"
    End Select
    Endianness_ToStr = s
End Function
'
'Public Function Max(V1, V2)
'    If V1 > V2 Then Max = V1 Else Max = V2
'End Function
'
'Public Function Min(V1, V2)
'    If V1 < V2 Then Min = V1 Else Min = V2
'End Function
'
