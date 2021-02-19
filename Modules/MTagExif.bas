Attribute VB_Name = "MTagExif"
Option Explicit
'                                                                                                         Tag-ID
                                          'Tag-Name                        Field-Name                    Dec      Hex      Type              Count
Public Enum TagExif
                                          'A. Tags Relating to Version
    itExifVersion = &H9000                'Exif version                    ExifVersion                   36864    9000     UNDEFINED         4
    itFlashpixVersion = &HA000            'Supported Flashpix version      FlashpixVersion               40960    A000     UNDEFINED         4

                                          'B. Tag Relating to Image Data Characteristics
    itColorSpace = &HA001                 'Color space information         ColorSpace                    40961    A001     SHORT             1

                                          'C. Tags Relating to Image Configuration
    itComponentsConfiguration = &H9101    'Meaning of each component       ComponentsConfiguration       37121    9101     UNDEFINED         4
    itCompressedBitsPerPixel = &H9102     'Image compression mode          CompressedBitsPerPixel        37122    9102     RATIONAL          1
    itPixelXDimension = &HA002            'Valid image width               PixelXDimension               40962    A002     SHORT or LONG     1
    itPixelYDimension = &HA003            'Valid image height              PixelYDimension               40963    A003     SHORT or LONG     1

                                          'D. Tags Relating to User Information
    itMakerNote = &H927C                  'Manufacturer notes              MakerNote                     37500    927C     UNDEFINED         Any
    itUserComment = &H9286                'User comments                   UserComment                   37510    9286     UNDEFINED         Any

                                          'E. Tag Relating to Related File Information
    itRelatedSoundFile = &HA004           'Related audio file              RelatedSoundFile              40964    A004     ASCII             13

                                          'F. Tags Relating to Date and Time
                                          'Date and time of original data
    itDateTimeOriginal = &H9003           'generation                      DateTimeOriginal              36867    9003     ASCII             20
                                          'Date and time of digital data
    itDateTimeDigitized = &H9004          'generation                      DateTimeDigitized             36868    9004     ASCII             20
    itSubSecTime = &H9290                 'DateTime subseconds             SubSecTime                    37520    9290     ASCII             Any
    itSubSecTimeOriginal = &H9291         'DateTimeOriginal subseconds     SubSecTimeOriginal            37521    9291     ASCII             Any
    itSubSecTimeDigitized = &H9292        'DateTimeDigitized subseconds    SubSecTimeDigitized           37522    9292     ASCII             Any

                                          'G. Tags Relating to Picture-Taking Conditions
    itExposureTime = &H829A               'Exposure time                   ExposureTime                  33434    829A     RATIONAL          1
    itFNumber = &H829D                    'F number                        FNumber                       33437    829D     RATIONAL          1
    itExposureProgram = &H8822            'Exposure program                ExposureProgram               34850    8822     SHORT             1
    itSpectralSensitivity = &H8824        'Spectral sensitivity            SpectralSensitivity           34852    8824     ASCII             Any
    itISOSpeedRatings = &H8827            'ISO speed rating                ISOSpeedRatings               34855    8827     SHORT             Any
    itOECF = &H8828                       'Optoelectric conversion factor  OECF                          34856    8828     UNDEFINED         Any
    itShutterSpeedValue = &H9201          'Shutter speed                   ShutterSpeedValue             37377    9201     SRATIONAL         1
    itApertureValue = &H9202              'Aperture                        ApertureValue                 37378    9202     RATIONAL          1
    itBrightnessValue = &H9203            'Brightness                      BrightnessValue               37379    9203     SRATIONAL         1
    itExposureBiasValue = &H9204          'Exposure bias                   ExposureBiasValue             37380    9204     SRATIONAL         1
    itMaxApertureValue = &H9205           'Maximum lens aperture           MaxApertureValue              37381    9205     RATIONAL          1
    itSubjectDistance = &H9206            'Subject distance                SubjectDistance               37382    9206     RATIONAL          1
    itMeteringMode = &H9207               'Metering mode                   MeteringMode                  37383    9207     SHORT             1
    itLightSource = &H9208                'Light source                    LightSource                   37384    9208     SHORT             1
    itFlash = &H9209                      'Flash                           Flash                         37385    9209     SHORT             1
    itFocalLength = &H920A                'Lens focal length               FocalLength                   37386    920A     RATIONAL          1
    itSubjectArea = &H9214                'Subject area                    SubjectArea                   37396    9214     SHORT             2 or 3 or 4
    itFlashEnergy = &HA20B                'Flash energy                    FlashEnergy                   41483    A20B     RATIONAL          1
    itSpatialFrequencyResponse = &HA20C   'Spatial frequency response      SpatialFrequencyResponse      41484    A20C     UNDEFINED         Any
    itFocalPlaneXResolution = &HA20E      'Focal plane X resolution        FocalPlaneXResolution         41486    A20E     RATIONAL          1
    itFocalPlaneYResolution = &HA20F      'Focal plane Y resolution        FocalPlaneYResolution         41487    A20F     RATIONAL          1
    itFocalPlaneResolutionUnit = &HA210   'Focal plane resolution unit     FocalPlaneResolutionUnit      41488    A210     SHORT             1
    itSubjectLocation = &HA214            'Subject location                SubjectLocation               41492    A214     SHORT             2
    itExposureIndex = &HA215              'Exposure index                  ExposureIndex                 41493    A215     RATIONAL          1
    itSensingMethod = &HA217              'Sensing method                  SensingMethod                 41495    A217     SHORT             1
    itFileSource = &HA300                 'File source                     FileSource                    41728    A300     UNDEFINED         1
    itSceneType = &HA301                  'Scene type                      SceneType                     41729    A301     UNDEFINED         1
    itCFAPattern = &HA302                 'CFA pattern                     CFAPattern                    41730    A302     UNDEFINED         Any
    itCustomRendered = &HA401             'Custom image processing         CustomRendered                41985    A401     SHORT             1
    itExposureMode = &HA402               'Exposure mode                   ExposureMode                  41986    A402     SHORT             1
    itWhiteBalance = &HA403               'White balance                   WhiteBalance                  41987    A403     SHORT             1
    itDigitalZoomRatio = &HA404           'Digital zoom ratio              DigitalZoomRatio              41988    A404     RATIONAL          1
    itFocalLengthIn35mmFilm = &HA405      'Focal length in 35 mm film      FocalLengthIn35mmFilm         41989    A405     SHORT             1
    itSceneCaptureType = &HA406           'Scene capture type              SceneCaptureType              41990    A406     SHORT             1
    itGainControl = &HA407                'Gain control                    GainControl                   41991    A407     RATIONAL          1
    itContrast = &HA408                   'Contrast                        Contrast                      41992    A408     SHORT             1
    itSaturation = &HA409                 'Saturation                      Saturation                    41993    A409     SHORT             1
    itSharpness = &HA40A                  'Sharpness                       Sharpness                     41994    A40A     SHORT             1
    itDeviceSettingDescription = &HA40B   'Device settings description     DeviceSettingDescription      41995    A40B     UNDEFINED         Any
    itSubjectDistanceRange = &HA40C       'Subject distance range          SubjectDistanceRange          41996    A40C     SHORT             1
                                        
                                          'H.other Tags
    itImageUniqueID = &HA420              'Unique image ID                 ImageUniqueID                 42016    A420     ASCII             33

End Enum

Public Function TagExif_ToStr(ByVal this As TagExif) As String
    Dim s As String
    Select Case this
                                                                            'A. Tags Relating to Version
    Case itExifVersion:                 s = "ExifVersion"                   'Exif version                    ExifVersion                   36864    9000     UNDEFINED         4
    Case itFlashpixVersion:             s = "FlashpixVersion"               'Supported Flashpix version      FlashpixVersion               40960    A000     UNDEFINED         4

                                                                            'B. Tag Relating to Image Data Characteristics
    Case itColorSpace:                  s = "ColorSpace"                    'Color space information         ColorSpace                    40961    A001     SHORT             1

                                                                            'C. Tags Relating to Image Configuration
    Case itComponentsConfiguration:     s = "ComponentsConfiguration"       'Meaning of each component       ComponentsConfiguration       37121    9101     UNDEFINED         4
    Case itCompressedBitsPerPixel:      s = "CompressedBitsPerPixel"        'Image compression mode          CompressedBitsPerPixel        37122    9102     RATIONAL          1
    Case itPixelXDimension:             s = "PixelXDimension"               'Valid image width               PixelXDimension               40962    A002     SHORT or LONG     1
    Case itPixelYDimension:             s = "PixelYDimension"               'Valid image height              PixelYDimension               40963    A003     SHORT or LONG     1

                                                                            'D. Tags Relating to User Information
    Case itMakerNote:                   s = "MakerNote"                     'Manufacturer notes              MakerNote                     37500    927C     UNDEFINED         Any
    Case itUserComment:                 s = "UserComment"                   'User comments                   UserComment                   37510    9286     UNDEFINED         Any

                                                                            'E. Tag Relating to Related File Information
    Case itRelatedSoundFile:            s = "RelatedSoundFile"              'Related audio file              RelatedSoundFile              40964    A004     ASCII             13
                                                                         
                                                                            'F. Tags Relating to Date and Time
                                                                            'Date and time of original data
    Case itDateTimeOriginal:            s = "DateTimeOriginal"              'generation                      DateTimeOriginal              36867    9003     ASCII             20
                                                                            'Date and time of digital data
    Case itDateTimeDigitized:           s = "DateTimeDigitized"             'generation                      DateTimeDigitized             36868    9004     ASCII             20
    Case itSubSecTime:                  s = "SubSecTime"                    'DateTime subseconds             SubSecTime                    37520    9290     ASCII             Any
    Case itSubSecTimeOriginal:          s = "SubSecTimeOriginal"            'DateTimeOriginal subseconds     SubSecTimeOriginal            37521    9291     ASCII             Any
    Case itSubSecTimeDigitized:         s = "SubSecTimeDigitized"           'DateTimeDigitized subseconds    SubSecTimeDigitized           37522    9292     ASCII             Any

                                                                            'G. Tags Relating to Picture-Taking Conditions
    Case itExposureTime:                s = "ExposureTime"                  'Exposure time                   ExposureTime                  33434    829A     RATIONAL          1
    Case itFNumber:                     s = "FNumber"                       'F number                        FNumber                       33437    829D     RATIONAL          1
    Case itExposureProgram:             s = "ExposureProgram"               'Exposure program                ExposureProgram               34850    8822     SHORT             1
    Case itSpectralSensitivity:         s = "SpectralSensitivity"           'Spectral sensitivity            SpectralSensitivity           34852    8824     ASCII             Any
    Case itISOSpeedRatings:             s = "ISOSpeedRatings"               'ISO speed rating                ISOSpeedRatings               34855    8827     SHORT             Any
    Case itOECF:                        s = "OECF"                          'Optoelectric conversion factor  OECF                          34856    8828     UNDEFINED         Any
    Case itShutterSpeedValue:           s = "ShutterSpeedValue"             'Shutter speed                   ShutterSpeedValue             37377    9201     SRATIONAL         1
    Case itApertureValue:               s = "ApertureValue"                 'Aperture                        ApertureValue                 37378    9202     RATIONAL          1
    Case itBrightnessValue:             s = "BrightnessValue"               'Brightness                      BrightnessValue               37379    9203     SRATIONAL         1
    Case itExposureBiasValue:           s = "ExposureBiasValue"             'Exposure bias                   ExposureBiasValue             37380    9204     SRATIONAL         1
    Case itMaxApertureValue:            s = "MaxApertureValue"              'Maximum lens aperture           MaxApertureValue              37381    9205     RATIONAL          1
    Case itSubjectDistance:             s = "SubjectDistance"               'Subject distance                SubjectDistance               37382    9206     RATIONAL          1
    Case itMeteringMode:                s = "MeteringMode"                  'Metering mode                   MeteringMode                  37383    9207     SHORT             1
    Case itLightSource:                 s = "LightSource"                   'Light source                    LightSource                   37384    9208     SHORT             1
    Case itFlash:                       s = "Flash"                         'Flash                           Flash                         37385    9209     SHORT             1
    Case itFocalLength:                 s = "FocalLength"                   'Lens focal length               FocalLength                   37386    920A     RATIONAL          1
    Case itSubjectArea:                 s = "SubjectArea"                   'Subject area                    SubjectArea                   37396    9214     SHORT             2 or 3 or 4
    Case itFlashEnergy:                 s = "FlashEnergy"                   'Flash energy                    FlashEnergy                   41483    A20B     RATIONAL          1
    Case itSpatialFrequencyResponse:    s = "SpatialFrequencyResponse"      'Spatial frequency response      SpatialFrequencyResponse      41484    A20C     UNDEFINED         Any
    Case itFocalPlaneXResolution:       s = "FocalPlaneXResolution"         'Focal plane X resolution        FocalPlaneXResolution         41486    A20E     RATIONAL          1
    Case itFocalPlaneYResolution:       s = "FocalPlaneYResolution"         'Focal plane Y resolution        FocalPlaneYResolution         41487    A20F     RATIONAL          1
    Case itFocalPlaneResolutionUnit:    s = "FocalPlaneResolutionUnit"      'Focal plane resolution unit     FocalPlaneResolutionUnit      41488    A210     SHORT             1
    Case itSubjectLocation:             s = "SubjectLocation"               'Subject location                SubjectLocation               41492    A214     SHORT             2
    Case itExposureIndex:               s = "ExposureIndex"                 'Exposure index                  ExposureIndex                 41493    A215     RATIONAL          1
    Case itSensingMethod:               s = "SensingMethod"                 'Sensing method                  SensingMethod                 41495    A217     SHORT             1
    Case itFileSource:                  s = "FileSource"                    'File source                     FileSource                    41728    A300     UNDEFINED         1
    Case itSceneType:                   s = "SceneType"                     'Scene type                      SceneType                     41729    A301     UNDEFINED         1
    Case itCFAPattern:                  s = "CFAPattern"                    'CFA pattern                     CFAPattern                    41730    A302     UNDEFINED         Any
    Case itCustomRendered:              s = "CustomRendered"                'Custom image processing         CustomRendered                41985    A401     SHORT             1
    Case itExposureMode:                s = "ExposureMode"                  'Exposure mode                   ExposureMode                  41986    A402     SHORT             1
    Case itWhiteBalance:                s = "WhiteBalance"                  'White balance                   WhiteBalance                  41987    A403     SHORT             1
    Case itDigitalZoomRatio:            s = "DigitalZoomRatio"              'Digital zoom ratio              DigitalZoomRatio              41988    A404     RATIONAL          1
    Case itFocalLengthIn35mmFilm:       s = "FocalLengthIn35mmFilm"         'Focal length in 35 mm film      FocalLengthIn35mmFilm         41989    A405     SHORT             1
    Case itSceneCaptureType:            s = "SceneCaptureType"              'Scene capture type              SceneCaptureType              41990    A406     SHORT             1
    Case itGainControl:                 s = "GainControl"                   'Gain control                    GainControl                   41991    A407     RATIONAL          1
    Case itContrast:                    s = "Contrast"                      'Contrast                        Contrast                      41992    A408     SHORT             1
    Case itSaturation:                  s = "Saturation"                    'Saturation                      Saturation                    41993    A409     SHORT             1
    Case itSharpness:                   s = "Sharpness"                     'Sharpness                       Sharpness                     41994    A40A     SHORT             1
    Case itDeviceSettingDescription:    s = "DeviceSettingDescription"      'Device settings description     DeviceSettingDescription      41995    A40B     UNDEFINED         Any
    Case itSubjectDistanceRange:        s = "SubjectDistanceRange"          'Subject distance range          SubjectDistanceRange          41996    A40C     SHORT             1
                                        
                                                                            'H.other Tags
    Case itImageUniqueID:               s = "ImageUniqueID"                 'Unique image ID                 ImageUniqueID                 42016    A420     ASCII             33
    Case Else:                          s = "unkown"
    End Select
    TagExif_ToStr = s & Space(MaxTagStrLen - Len(s))
End Function


Public Function IFDExif_ToStr(this As IFD, Optional ByVal Index As Long) As String
Try: On Error GoTo Catch
    Dim i As Long
    Dim s As String
    Dim dt As IFDataType
    With this
        s = s & "Count: " & CStr(.Count) & vbCrLf
        For i = 0 To .Count - 1 ' UBound(.Entries)
            s = s & " " & CStr(i) & ": " & vbCrLf
            s = s & IFDExifEntryValue_ToStr(.Entries(i)) & vbCrLf
        Next
        s = s & " OffsetNextIFD: " & CStr(.OffsetNextIFD) & vbCrLf
    End With
    IFDExif_ToStr = s
    Exit Function
Catch: ErrHandler "IFDExif_ToStr", s
End Function
Public Function IFDExifEntryValue_ToStr(this As IFDEntryValue) As String
Try: On Error GoTo Catch
    Dim s As String
    Dim dt As IFDataType
    With this
        With .Entry
            s = s & "  Tag:    " & TagExif_ToStr(.Tag) & " &H" & Hex$(.Tag) & vbCrLf
            s = s & "  Type:   " & IFDataType_ToStr(.DataType) & vbCrLf
            s = s & "  Count:  " & CStr(.Count) & vbCrLf
            dt = .DataType
        End With
        'in zwei Schritten zuerst ob Offset geschrieben werden soll
        Select Case dt
        Case IFDataType.dtASCII, IFDataType.dtByte, IFDataType.dtSByte, IFDataType.dtUndefined2
            If .Entry.Count > 4 Then
                s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
            End If
        Case IFDataType.dtShort, IFDataType.dtSShort
            If .Entry.Count > 2 Then
                s = s & "  Offset: " & CStr(.Entry.ValueOffset) & vbCrLf
            End If
        Case IFDataType.dtFloat, IFDataType.dtLong, IFDataType.dtSLong
            If .Entry.Count > 1 Then
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
    IFDExifEntryValue_ToStr = s
    Exit Function
Catch: ErrHandler "IFDExifEntryValue_ToStr", s
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
'Private Function PErrHandler(ByVal FncName As String, ByVal AddErrMsg As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
'    PErrHandler = MError.ErrHandler("MGPSTag", FncName, AddErrMsg, Buttons)
'End Function
'
'
