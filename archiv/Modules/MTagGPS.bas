Attribute VB_Name = "MTagGPS"
Option Explicit

Public Enum TagGPS
    itGPSVersionID = 0         ' &H0  BYTE        4 ' GPS tag version
    itGPSLatitudeRef = 1       ' &H1  ASCII       2 ' North or South Latitude
    itGPSLatitude = 2          ' &H2  RATIONAL    3 ' Latitude
    itGPSLongitudeRef = 3      ' &H3  ASCII       2 ' East or West Longitude
    itGPSLongitude = 4         ' &H4  RATIONAL    3 ' Longitude
    itGPSAltitudeRef = 5       ' &H5  BYTE        1 ' Altitude reference
    itGPSAltitude = 6          ' &H6  RATIONAL    1 ' Altitude
    itGPSTimeStamp = 7         ' &H7  RATIONAL    3 ' GPS time (atomic clock)
    itGPSSatellites = 8        ' &H8  ASCII     Any ' GPS satellites used for measurement
    itGPSStatus = 9            ' &H9  ASCII       2 ' GPS receiver status
    itGPSMeasureMode = 10      ' &HA  ASCII       2 ' GPS measurement mode
    itGPSDOP = 11              ' &HB  RATIONAL    1 ' Measurement precision
    itGPSSpeedRef = 12         ' &HC  ASCII       2 ' Speed unit
    itGPSSpeed = 13            ' &HD  RATIONAL    1 ' Speed of GPS receiver
    itGPSTrackRef = 14         ' &HE  ASCII       2 ' Reference for direction of movement
    itGPSTrack = 15            ' &HF  RATIONAL    1 ' Direction of movement
    itGPSImgDirectionRef = 16  ' &H10 ASCII       2 ' Reference for direction of image
    itGPSImgDirection = 17     ' &H11 RATIONAL    1 ' Direction of image
    itGPSMapDatum = 18         ' &H12 ASCII     Any ' Geodetic survey data used
    itGPSDestLatitudeRef = 19  ' &H13 ASCII       2 ' Reference for latitude of destination
    itGPSDestLatitude = 20     ' &H14 RATIONAL    3 ' Latitude of destination
    itGPSDestLongitudeRef = 21 ' &H15 ASCII       2 ' Reference for longitude of destination
    itGPSDestLongitude = 22    ' &H16 RATIONAL    3 ' Longitude of destination
    itGPSDestBearingRef = 23   ' &H17 ASCII       2 ' Reference for bearing of destination
    itGPSDestBearing = 24      ' &H18 RATIONAL    1 ' Bearing of destination
    itGPSDestDistanceRef = 25  ' &H19 ASCII       2 ' Reference for distance to destination
    itGPSDestDistance = 26     ' &H1A RATIONAL    1 ' Distance to destination
    itGPSProcessingMethod = 27 ' &H1B UNDEFINED Any ' Name of GPS processing method
    itGPSAreaInformation = 28  ' &H1C UNDEFINED Any ' Name of GPS area
    itGPSDateStamp = 29        ' &H1D ASCII      11 ' GPS date
    itGPSDifferential = 30     ' &H1E SHORT       1 ' GPS differential correction
End Enum

Public Function TagGPS_ToStr(ByVal t As TagGPS) As String
    Dim s As String
    Select Case t
    Case itGPSVersionID:                s = "GPSVersionID"         ' &H0  BYTE        4 ' GPS tag version
    Case itGPSLatitudeRef:              s = "GPSLatitudeRef"       ' &H1  ASCII       2 ' North or South Latitude
    Case itGPSLatitude:                 s = "GPSLatitude"          ' &H2  RATIONAL    3 ' Latitude
    Case itGPSLongitudeRef:             s = "GPSLongitudeRef"      ' &H3  ASCII       2 ' East or West Longitude
    Case itGPSLongitude:                s = "GPSLongitude"         ' &H4  RATIONAL    3 ' Longitude
    Case itGPSAltitudeRef:              s = "GPSAltitudeRef"       ' &H5  BYTE        1 ' Altitude reference
    Case itGPSAltitude:                 s = "GPSAltitude"          ' &H6  RATIONAL    1 ' Altitude
    Case itGPSTimeStamp:                s = "GPSTimeStamp"         ' &H7  RATIONAL    3 ' GPS time (atomic clock)
    Case itGPSSatellites:               s = "GPSSatellites"        ' &H8  ASCII     Any ' GPS satellites used for measurement
    Case itGPSStatus:                   s = "GPSStatus"            ' &H9  ASCII       2 ' GPS receiver status
    Case itGPSMeasureMode:              s = "GPSMeasureMode"       ' &HA  ASCII       2 ' GPS measurement mode
    Case itGPSDOP:                      s = "GPSDOP"               ' &HB  RATIONAL    1 ' Measurement precision
    Case itGPSSpeedRef:                 s = "GPSSpeedRef"          ' &HC  ASCII       2 ' Speed unit
    Case itGPSSpeed:                    s = "GPSSpeed"             ' &HD  RATIONAL    1 ' Speed of GPS receiver
    Case itGPSTrackRef:                 s = "GPSTrackRef"          ' &HE  ASCII       2 ' Reference for direction of movement
    Case itGPSTrack:                    s = "GPSTrack"             ' &HF  RATIONAL    1 ' Direction of movement
    Case itGPSImgDirectionRef:          s = "GPSImgDirectionRef"   ' &H10 ASCII       2 ' Reference for direction of image
    Case itGPSImgDirection:             s = "GPSImgDirection"      ' &H11 RATIONAL    1 ' Direction of image
    Case itGPSMapDatum:                 s = "GPSMapDatum"          ' &H12 ASCII     Any ' Geodetic survey data used
    Case itGPSDestLatitudeRef:          s = "GPSDestLatitudeRef"   ' &H13 ASCII       2 ' Reference for latitude of destination
    Case itGPSDestLatitude:             s = "GPSDestLatitude"      ' &H14 RATIONAL    3 ' Latitude of destination
    Case itGPSDestLongitudeRef:         s = "GPSDestLongitudeRef"  ' &H15 ASCII       2 ' Reference for longitude of destination
    Case itGPSDestLongitude:            s = "GPSDestLongitude"     ' &H16 RATIONAL    3 ' Longitude of destination
    Case itGPSDestBearingRef:           s = "GPSDestBearingRef"    ' &H17 ASCII       2 ' Reference for bearing of destination
    Case itGPSDestBearing:              s = "GPSDestBearing"       ' &H18 RATIONAL    1 ' Bearing of destination
    Case itGPSDestDistanceRef:          s = "GPSDestDistanceRef"   ' &H19 ASCII       2 ' Reference for distance to destination
    Case itGPSDestDistance:             s = "GPSDestDistance"      ' &H1A RATIONAL    1 ' Distance to destination
    Case itGPSProcessingMethod:         s = "GPSProcessingMethod"  ' &H1B UNDEFINED Any ' Name of GPS processing method
    Case itGPSAreaInformation:          s = "GPSAreaInformation"   ' &H1C UNDEFINED Any ' Name of GPS area
    Case itGPSDateStamp:                s = "GPSDateStamp"         ' &H1D ASCII      11 ' GPS date
    Case itGPSDifferential:             s = "GPSDifferential"      ' &H1E SHORT       1 ' GPS differential correction
    Case Else:                          s = "unkown"
    End Select
    TagGPS_ToStr = s & Space(MaxTagStrLen - Len(s))
End Function

Public Function IFDGPS_ToStr(this As IFD, Optional ByVal Index As Long) As String
Try: On Error GoTo Catch
    Dim i As Long
    Dim s As String
    Dim dt As IFDataType
    With this
        s = s & "Count: " & CStr(.count) & vbCrLf
        For i = 0 To .count - 1 ' UBound(.Entries)
            s = s & " " & CStr(i) & ": " & vbCrLf
            s = s & IFDGPSEntryValue_ToStr(.Entries(i)) & vbCrLf
        Next
        s = s & " OffsetNextIFD: " & CStr(.OffsetNextIFD) & vbCrLf
    End With
    IFDGPS_ToStr = s
    Exit Function
Catch: ErrHandler "IFDGPS_ToStr", s
End Function
Public Function IFDGPSEntryValue_ToStr(this As IFDEntryValue) As String
Try: On Error GoTo Catch
    Dim s As String
    Dim dt As IFDataType
    With this
        With .Entry
            s = s & "  Tag:    " & MTagGPS.TagGPS_ToStr(.Tag) & " &H" & Hex$(.Tag) & vbCrLf
            s = s & "  Type:   " & IFDataType_ToStr(.DataType) & vbCrLf
            s = s & "  Count:  " & CStr(.count) & vbCrLf
            dt = .DataType
        End With
        'in zwei Schritten zuerst ob Offset geschrieben werden soll
        Select Case dt
        Case IFDataType.dtASCII, IFDataType.dtByte, IFDataType.dtSByte, IFDataType.dtUndefined2
            If .Entry.count > 4 Then
                s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
            End If
        Case IFDataType.dtShort, IFDataType.dtSShort
            If .Entry.count > 2 Then
                s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
            End If
        Case IFDataType.dtFloat, IFDataType.dtLong, IFDataType.dtSLong
            If .Entry.count > 1 Then
                s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
            End If
        Case IFDataType.dtRational, IFDataType.dtSRational, IFDataType.dtDouble
            s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
        End Select
        'dann den Wert dazuschreiben
        s = s & "  Value:  "
        Dim v As Variant
        v = MExif.IFDEntryValue_GetValue(this)
        If IsArray(v) Then
            s = s & IFValueArray_ToStr(dt, v)
        Else
            Select Case dt
            Case IFDataType.dtRational, IFDataType.dtSRational
                s = s & IFRational_ToStr(v)
            Case IFDataType.dtASCII
                s = s & """" & v & """"
            Case Else
                s = s & CStr(v)
            End Select
        End If
    End With
    IFDGPSEntryValue_ToStr = s
    Exit Function
Catch: ErrHandler "IFDGPSEntryValue_ToStr", s
End Function

'##############################'   Locale ErrHandler   '##############################'
Private Function ErrHandler(ByVal FncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional ByVal bLoud As Boolean = True, _
                            Optional ByVal bErrLog As Boolean = False, _
                            Optional ByVal vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical _
                            ) As VbMsgBoxResult
    ErrHandler = MError.ErrHandler("MExif", FncName, AddInfo, bLoud, bErrLog, vbDecor)
End Function
'
'Private Function PErrHandler(ByVal FncName As String, ByVal AddErrMsg As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
'    PErrHandler = MError.ErrHandler("MGPSTag", FncName, AddErrMsg, Buttons)
'End Function
'
