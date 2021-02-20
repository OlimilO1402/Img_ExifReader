Attribute VB_Name = "MNew"
Option Explicit

Public Function TaggedImageFile(aFileName As String) As TaggedImageFile
    Set TaggedImageFile = New TaggedImageFile: TaggedImageFile.New_ aFileName
End Function

Public Function FilEBinReader(PathFileName As String) As FilEBinReader
    Set FilEBinReader = New FilEBinReader: FilEBinReader.New_ PathFileName
End Function

