Attribute VB_Name = "MNew"
Option Explicit

Public Function TaggedImageFile(aFileName As String) As TaggedImageFile
    Set TaggedImageFile = New TaggedImageFile: TaggedImageFile.New_ aFileName
End Function

Public Function FileStream(aPathFileName As String, _
                            Optional FMode As FileMode = fmRandom, _
                            Optional FAccess As FileAccess = faReadWrite, _
                            Optional FShare As FileShare = fsNone) As FileStream
    Set FileStream = New FileStream: FileStream.New_ aPathFileName, FMode, FAccess, FShare
End Function

'Public Function EndianBinaryReader(st As Stream) As EndianBinaryReader
'    Set EndianBinaryReader = New EndianBinaryReader: EndianBinaryReader.New_ st
'End Function
Public Function EBinaryReader(st As Stream) As EBinaryReader
    Set EBinaryReader = New EBinaryReader: EBinaryReader.New_ st
End Function

