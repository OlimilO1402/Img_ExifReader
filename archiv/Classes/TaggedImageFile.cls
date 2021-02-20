VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaggedImageFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' stellt eine Klasse dar für Bilddateien mit integrierten Tags und
' zum Bild gehörende beschreibende im Bild integrierte Daten die bei
' Weitergabe des Bildes mit dem Bild verknüpft bleiben sollen.
' die Klasse soll Lese -und Schreibzugriffe auf diese Daten bereitstellen
' Bilder im Format mit der Endung *.tif und *.jpg
'
Private m_FileName    As String
Private m_IFHeader    As IFHeader
Private m_IFD_0       As IFD
Private m_IFD_Exif    As IFD
Private m_IFD_GPS     As IFD
Private m_IFD_Interop As IFD
Private m_IFD_1       As IFD
'Private m_Values      As Collection 'speichert alle Values anhand des Tags

Friend Sub New_(aFileName As String)
    m_FileName = aFileName
End Sub

#If AssumeNoBigEndBytSwap Then
Public Function Read() As Boolean
Try: On Error GoTo Finally

    Dim OffsetExif       As Long
    Dim OffsetIFHeader   As Long
    Dim OffsetIFDExif    As Long
    Dim OffsetIFDGPS     As Long
    Dim OffsetIFDInterop As Long
    
    
    Dim FNr As Integer: FNr = FreeFile
    Open m_FileName For Binary Access Read As FNr
    
    OffsetExif = MExif.GetPosition(FNr, MExif.C_ExifHeader)
    If OffsetExif < 0 Then
        PErrHandler "Read", "Exif-tag not found.", False
        Exit Function
    End If
    
    OffsetIFHeader = Align4(OffsetExif + Len(MExif.C_ExifHeader))             'must be a multiple of 4 +1!
    If OffsetIFHeader >= 0 Then _
        Read = ReadIFHeader(m_IFHeader, FNr, OffsetIFHeader)
        Read = ReadIFD(m_IFD_0, FNr, m_IFHeader.OffsetIFD0, OffsetIFHeader)
    
    OffsetIFDExif = MExif.IFD_ValueByTag(m_IFD_0, itExifIFDOffset)
    If OffsetIFDExif <> 0 Then _
        Read = ReadIFD(m_IFD_Exif, FNr, OffsetIFDExif, OffsetIFHeader)
    
    OffsetIFDGPS = MExif.IFD_ValueByTag(m_IFD_0, itGPSIFDOffset)
    If OffsetIFDGPS <> 0 Then _
        Read = ReadIFD(m_IFD_GPS, FNr, OffsetIFDGPS, OffsetIFHeader)
    
    OffsetIFDInterop = MExif.IFD_ValueByTag(m_IFD_Exif, itInteropIFDOffset)
    If OffsetIFDInterop <> 0 Then _
        Read = ReadIFD(m_IFD_Interop, FNr, OffsetIFDInterop, OffsetIFHeader)
    
    If m_IFD_0.OffsetNextIFD <> 0 Then _
        Read = ReadIFD(m_IFD_1, FNr, m_IFD_0.OffsetNextIFD, OffsetIFHeader)
    
    Read = True
Finally:
    Close #FNr
    If Err <> 0 Then
        Read = False
        ErrHandler "Read", """" & m_FileName & """" '& vbclrf
    End If
End Function

#Else

Public Function Read() As Boolean
Try: On Error GoTo Finally
    
    Dim OffsetExif       As Long
    Dim OffsetIFHeader   As Long
    Dim OffsetIFDExif    As Long
    Dim OffsetIFDGPS     As Long
    Dim OffsetIFDInterop As Long
    
    Dim fst As Stream:        Set fst = MNew.FileStream(m_FileName, fmBinary, faRead)
    Dim ebr As EBinaryReader: Set ebr = MNew.EBinaryReader(fst)
    
    OffsetExif = MExif.GetPosition(ebr, MExif.C_ExifHeader)
    If OffsetExif < 0 Then
        ErrHandler "Read", "Exif-tag not found.", False
        Exit Function
    End If
    
    OffsetIFHeader = Align4(OffsetExif + Len(MExif.C_ExifHeader))           'must be a multiple of 4 +1!
    If OffsetIFHeader >= 0 Then _
        Read = ReadIFHeader(m_IFHeader, ebr, OffsetIFHeader)
        
    Read = ReadIFD(m_IFD_0, ebr, m_IFHeader.OffsetIFD0, OffsetIFHeader)
    
    OffsetIFDExif = MExif.IFD_ValueByTag(m_IFD_0, itExifIFDOffset)
    If OffsetIFDExif <> 0 Then _
        Read = ReadIFD(m_IFD_Exif, ebr, OffsetIFDExif, OffsetIFHeader)
    
    OffsetIFDGPS = MExif.IFD_ValueByTag(m_IFD_0, itGPSIFDOffset)
    If OffsetIFDGPS <> 0 Then _
        Read = ReadIFD(m_IFD_GPS, ebr, OffsetIFDGPS, OffsetIFHeader)
    
    OffsetIFDInterop = MExif.IFD_ValueByTag(m_IFD_Exif, itInteropIFDOffset)
    If OffsetIFDInterop <> 0 Then _
        Read = ReadIFD(m_IFD_Interop, ebr, OffsetIFDInterop, OffsetIFHeader)
    
    If m_IFD_0.OffsetNextIFD <> 0 Then _
        Read = ReadIFD(m_IFD_1, ebr, m_IFD_0.OffsetNextIFD, OffsetIFHeader)
    
    Read = True
Finally:
    fst.CClose
    If Err <> 0 Then
        Read = False
        ErrHandler "Read", """" & m_FileName & """" '& vbclrf
    End If
End Function
#End If

Friend Property Get IFHeader() As IFHeader
    IFHeader = m_IFHeader
End Property
Friend Property Get IFD0() As IFD
    IFD0 = m_IFD_0
End Property
Friend Property Get IFDExif() As IFD
    IFDExif = m_IFD_Exif
End Property
Friend Property Get IFDGPS() As IFD
    IFDGPS = m_IFD_GPS
End Property
Friend Property Get IFDInterop() As IFD
    IFDInterop = m_IFD_Interop
End Property
Friend Property Get IFD1() As IFD
    IFD1 = m_IFD_1
End Property

Public Function ToStr() As String
    Dim s As String
Try: On Error GoTo Catch
    If m_IFD_0.Count > 0 Then
        s = s & "IFD_0" & vbCrLf & "=====" & vbCrLf
        s = s & MExif.IFD_ToStr(m_IFD_0)
    End If
    If m_IFD_Exif.Count > 0 Then
        s = s & "IFD_Exif" & vbCrLf & "========" & vbCrLf
        s = s & MTagExif.IFDExif_ToStr(m_IFD_Exif)
    End If
    If m_IFD_GPS.Count > 0 Then
        s = s & "IFD_GPS" & vbCrLf & "=======" & vbCrLf
        s = s & MTagGPS.IFDGPS_ToStr(m_IFD_GPS)
    End If
    If m_IFD_Interop.Count > 0 Then
        s = s & "IFD_Interop" & vbCrLf & "===========" & vbCrLf
        s = s & MExif.IFD_ToStr(m_IFD_Interop)
    End If
    If m_IFD_1.Count > 0 Then
        s = s & "IFD_1" & vbCrLf & "=====" & vbCrLf
        s = s & MExif.IFD_ToStr(m_IFD_1)
    End If
    ToStr = s '& r
    Exit Function
Catch:
    If Err Then ErrHandler "ToStr", s
End Function

Public Function IsEQual(other As TaggedImageFile) As Boolean
    '
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
