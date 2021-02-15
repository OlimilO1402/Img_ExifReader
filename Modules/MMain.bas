Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    FrmMain.Show
End Sub

Public Property Get SystemEndianness() As Endianness
    SystemEndianness = Endianness.IntelLittleEndian
End Property
