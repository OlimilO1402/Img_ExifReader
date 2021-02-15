VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub OOpen()
End Sub

Public Sub Read(ByVal pDst As Long, ByVal BytLength As Long)
End Sub
Public Sub WWrite(ByVal pSrc As Long, ByVal BytLength As Long)
End Sub

Public Property Get Position() As Long
End Property
Public Property Let Position(ByVal Value As Long)
End Property

Public Property Get IsEOF() As Boolean
End Property

Public Property Get Length() As Long
End Property
Public Sub Flush()
End Sub

Public Sub CClose()
End Sub
