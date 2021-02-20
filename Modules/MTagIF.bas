Attribute VB_Name = "MTagIF"
'Option Explicit
'Public Enum TagIF
'    itGPSVersionID = &H0
'    itGPSLatitudeRef = &H1
'    itGPSLatitude = &H2
'    itGPSLongitudeRef = &H3
'    itGPSLongitude = &H4
'    itGPSAltitudeRef = &H5
'    itGPSAltitude = &H6
'    itGPSTimeStamp = &H7
'    itGPSSatellites = &H8
'    itGPSStatus = &H9
'    itGPSMeasureMode = &HA
'    itGPSDOP = &HB
'    itGPSSpeedRef = &HC
'    itGPSSpeed = &HD
'    itGPSTrackRef = &HE
'    itGPSTrack = &HF
'    itGPSImgDirectionRef = &H10
'    itGPSImgDirection = &H11
'    itGPSMapDatum = &H12
'    itGPSDestLatitudeRef = &H13
'    itGPSDestLatitude = &H14
'    itGPSDestLongitudeRef = &H15
'    itGPSDestLongitude = &H16
'    itGPSDestBearingRef = &H17
'    itGPSDestBearing = &H18
'    itGPSDestDistanceRef = &H19
'    itGPSDestDistance = &H1A
'    itGPSProcessingMethod = &H1B
'    itGPSAreaInformation = &H1C
'    itGPSDateStamp = &H1D
'    itGPSDifferential = &H1E
'    itImageWidth = &H100
'    itImageLength = &H101
'    itBitsPerSample = &H102
'    itCompression = &H103
'    itPhotometricInterpretaion = &H106
'    itImageDescription = &H10E
'    itMake = &H10F
'    itModel = &H110
'    itStripOffsets = &H111
'    itOrientation = &H112
'    itSamplesPerPixel = &H115
'    itRowsPerStrip = &H116
'    itStripByteCounts = &H117
'    itX -Resolution = &H11A
'    itY -Resolution = &H11B
'    itPlanarConfiguration = &H11C
'    itResolutionUnit = &H128
'    itTransferFunction = &H12D
'    itSoftwareUsed = &H131
'    itDateTime = &H132
'    itArtist = &H13B
'    itWhitePoint = &H13E
'    itPrimaryChromaticities = &H13F
'    itJPEGInterchangeFormat = &H201
'    itJPEGInterchangeFormatLength = &H202
'    itYCbCrCoefficients = &H211
'    itYCbCrSubSampling = &H212
'    itYCbCrPositioning = &H213
'    itReferenceBlackWhite = &H214
'    itCopyright = &H8298
'    itExposureTime = &H829A
'    itFNumber = &H829D
'    itIFDOffsetExif = &H8769
'    itExposureProgram = &H8822
'    itSpectralSensitivity = &H8824
'    itIFDOffsetGPS = &H8825
'    itISOSpeedRatings = &H8827
'    itOECF = &H8828
'    itSensitivityType = &H8830
'    itStandardOutputSensitivity = &H8831
'    itRecommendedExposureIndex = &H8832
'    itISOSpeed = &H8833
'    itISOSpeedLatitudeYYY = &H8834
'    itISOSpeedLatitudeZZZ = &H8835
'    itExifVersion = &H9000
'    itDateTimeOriginal = &H9003
'    itDateTimeDigitized = &H9004
'    itComponentsConfiguration = &H9101
'    itCompressedBitsPerPixel = &H9102
'    itShutterSpeedValue = &H9201
'    itApertureValue = &H9202
'    itBrightnessValue = &H9203
'    itExposureBiasValue = &H9204
'    itMaxApertureValue = &H9205
'    itSubjectDistance = &H9206
'    itMeteringMode = &H9207
'    itLightSource = &H9208
'    itFlash = &H9209
'    itFocalLength = &H920A
'    itSubjectArea = &H9214
'    itMakerNote = &H927C
'    itUserComment = &H9286
'    itSubSecTime = &H9290
'    itSubSecTimeOriginal = &H9291
'    itSubSecTimeDigitized = &H9292
'    itFlashpixVersion = &HA000
'    itColorSpace = &HA001
'    itPixelXDimension = &HA002
'    itPixelYDimension = &HA003
'    itRelatedSoundFile = &HA004
'    itIFDOffsetInterop = &HA005
'    itFlashEnergy = &HA20B
'    itSpatialFrequencyResponse = &HA20C
'    itFocalPlaneXResolution = &HA20E
'    itFocalPlaneYResolution = &HA20F
'    itFocalPlaneResolutionUnit = &HA210
'    itSubjectLocation = &HA214
'    itExposureIndex = &HA215
'    itSensingMethod = &HA217
'    itFileSource = &HA300
'    itSceneType = &HA301
'    itCFAPattern = &HA302
'    itCustomRendered = &HA401
'    itExposureMode = &HA402
'    itWhiteBalance = &HA403
'    itDigitalZoomRatio = &HA404
'    itFocalLengthIn35mmFilm = &HA405
'    itSceneCaptureType = &HA406
'    itGainControl = &HA407
'    itContrast = &HA408
'    itSaturation = &HA409
'    itSharpness = &HA40A
'    itDeviceSettingDescription = &HA40B
'    itSubjectDistanceRange = &HA40C
'    itImageUniqueID = &HA420
'    itCameraOwnerName = &HFFFF
'End Enum

'Public Enum TagIF
'    'itInteropIndex = &H1
'
''A) Tags relating to image data structure
'    itImageWidth = &H100                    ' Image width
'    itImageLength = &H101                   ' Image height
'    itBitsPerSample = &H102                 ' Number of bits per component
'    itCompression = &H103                   ' Compression scheme
'    itPhotometricInterpretaion = &H106      ' Pixel composition
'    itOrientation = &H112                   ' Orientation of image
'    itSamplesPerPixel = &H115               ' Number of components
'    itPlanarConfiguration = &H11C           ' Image data arrangement
'    itYCbCrSubSampling = &H212              ' Subsampling ratio of Y to C
'    itYCbCrPositioning = &H213              ' Y and Yc positioning
'    itXResolution = &H11A                   ' Image resolution in width direction
'    itYResolution = &H11B                   ' Image resolution in height direction
'    itResolutionUnit = &H128                ' ResolutionUnit
'
''B) Tags relating to recording offset
'    itStripOffsets = &H111                  ' Image Data location
'    itRowsPerStrip = &H116                  ' Number of rows per strip
'    itStripByteCounts = &H117               ' bytes per compressed strip
'    itJPEGInterchangeFormat = &H201         ' Offset of JPEG SOI
'    itJPEGInterchangeFormatLength = &H202   ' Bytes if JPEG data
'
'  'Tags relating to image data characteristics
'    itTransferFunction = &H12D              ' Transfer function
'    itWhitePoint = &H13E                    ' White point
'    itPrimaryChromaticities = &H13F         ' Primary Chromaticities
'    itYCbCrCoefficients = &H211             ' YCbCrCoefficients
'    itReferenceBlackWhite = &H214           ' ReferenceBlackWhite
'
'  'Other tags
'    itDateTime = &H132                      ' File change date and time
'    itImageDescription = &H10E              ' image title
'    itMake = &H10F                          ' manufacturer
'    itModel = &H110                         ' Image input equipment model
'    itSoftwareUsed = &H131                  ' Software used
'    itArtist = &H13B                        ' Person who created the image
'    'itSoftware = &H10E
'    itCopyright = &H8298
'
'    itExifIFDOffset = &H8769
'    itGPSIFDOffset = &H8825 '34853
'    itInteropIFDOffset = &HA005
'
''GPS Info IFD Pointer
''Tag = 34853 (8825.H)
''Type = LONG
''Count = 1
''Default = none
'
'    ' = &H0
'    ' = &H0
'End Enum '34
'
''könnte man in einem Array speichern oder in einer Collection Keys sind die Enum-Konstanten als Hex-String
'Public Function TagIF_ToStr(ByVal this As TagIF) As String
'    'das könnte man auch über eine Collection machen, da es sehr viele sein werden
'    Dim s As String
'    Select Case this
'    'Case TagIF.itInteropIndex: s = "InteropIndex"
'  'Tags relating to image data structure
'    Case TagIF.itImageWidth:                    s = "ImageWidth"
'    Case TagIF.itImageLength:                   s = "ImageLength"
'    Case TagIF.itBitsPerSample:                 s = "BitsPerSample"
'    Case TagIF.itCompression:                   s = "Compression"
'    Case TagIF.itPhotometricInterpretaion:      s = "PhotometricInterpretaion"
'    Case TagIF.itOrientation:                   s = "Orientation"
'    Case TagIF.itSamplesPerPixel:               s = "SamplesPerPixel"
'    Case TagIF.itPlanarConfiguration:           s = "PlanarConfiguration"
'    Case TagIF.itYCbCrSubSampling:              s = "YCbCrSubSampling"
'    Case TagIF.itYCbCrPositioning:              s = "YCbCrPositioning"
'    Case TagIF.itXResolution:                   s = "X-Resolution"
'    Case TagIF.itYResolution:                   s = "Y-Resolution"
'    Case TagIF.itResolutionUnit:                s = "ResolutionUnit"
'
'  'Tags relating to recording offset
'    Case TagIF.itStripOffsets:                  s = "StripOffsets"
'    Case TagIF.itRowsPerStrip:                  s = "RowsPerStrip"
'    Case TagIF.itStripByteCounts:               s = "StripByteCounts"
'    Case TagIF.itJPEGInterchangeFormat:         s = "JPEGInterchangeFormat"
'    Case TagIF.itJPEGInterchangeFormatLength:   s = "JPEGInterchangeFormatLength"
'
'  'Tags relating to image data characteristics
'    Case TagIF.itTransferFunction:              s = "TransferFunction"
'    Case TagIF.itWhitePoint:                    s = "WhitePoint"
'    Case TagIF.itPrimaryChromaticities:         s = "PrimaryChromaticities"
'    Case TagIF.itYCbCrCoefficients:             s = "YCbCrCoefficients"
'    Case TagIF.itReferenceBlackWhite:           s = "ReferenceBlackWhite"
'
'  'Other tags
'    Case TagIF.itDateTime:                      s = "DateTime"
'    Case TagIF.itImageDescription:              s = "ImageDescription"
'    Case TagIF.itMake:                          s = "Make"
'    Case TagIF.itModel:                         s = "Model"
'    Case TagIF.itSoftwareUsed:                  s = "SoftwareUsed"
'    Case TagIF.itArtist:                        s = "Artist"
'    Case TagIF.itCopyright:                     s = "Copyright"
'
'    Case TagIF.itExifIFDOffset:                 s = "IFDOffsetExif"
'    Case TagIF.itGPSIFDOffset:                  s = "IFDOffsetGPS"
'    Case TagIF.itInteropIFDOffset:              s = "IFDOffsetInterop"
'    'Case IFTag
'    Case Else:                                  s = "unknown"
'    End Select
'    TagIF_ToStr = s & Space(MaxTagStrLen - Len(s))
'End Function
'
