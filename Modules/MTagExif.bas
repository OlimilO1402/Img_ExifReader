Attribute VB_Name = "MTagExif"
Option Explicit
'                                                                                                         Tag-ID

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
End Enum '31

Public Enum TagIF
  'itInteropIndex = &H1
    
  'A) Tags relating to image data structure
    itImageWidth = &H100                    ' Image width
    itImageLength = &H101                   ' Image height
    itBitsPerSample = &H102                 ' Number of bits per component
    itCompression = &H103                   ' Compression scheme
    itPhotometricInterpretaion = &H106      ' Pixel composition
    itOrientation = &H112                   ' Orientation of image
    itSamplesPerPixel = &H115               ' Number of components
    itPlanarConfiguration = &H11C           ' Image data arrangement
    itYCbCrSubSampling = &H212              ' Subsampling ratio of Y to C
    itYCbCrPositioning = &H213              ' Y and Yc positioning
    itXResolution = &H11A                   ' Image resolution in width direction
    itYResolution = &H11B                   ' Image resolution in height direction
    itResolutionUnit = &H128                ' ResolutionUnit
  
  'B) Tags relating to recording offset
    itStripOffsets = &H111                  ' Image Data location
    itRowsPerStrip = &H116                  ' Number of rows per strip
    itStripByteCounts = &H117               ' bytes per compressed strip
    itJPEGInterchangeFormat = &H201         ' Offset of JPEG SOI
    itJPEGInterchangeFormatLength = &H202   ' Bytes if JPEG data
        
  'Tags relating to image data characteristics
    itTransferFunction = &H12D              ' Transfer function
    itWhitePoint = &H13E                    ' White point
    itPrimaryChromaticities = &H13F         ' Primary Chromaticities
    itYCbCrCoefficients = &H211             ' YCbCrCoefficients
    itReferenceBlackWhite = &H214           ' ReferenceBlackWhite
  
  'Other tags
    itDateTime = &H132                      ' File change date and time
    itImageDescription = &H10E              ' image title
    itMaker = &H10F                          ' manufacturer
    itModel = &H110                         ' Image input equipment model
    itSoftwareUsed = &H131                  ' Software used
    itArtist = &H13B                        ' Person who created the image
    'itSoftware = &H10E
    itCopyright = &H8298
    
    itIFDOffsetExif = &H8769
    itIFDOffsetGPS = &H8825 '34853
    itIFDOffsetInterop = &HA005
    
'GPS Info IFD Pointer
'Tag = 34853 (8825.H)
'Type = LONG
'Count = 1
'Default = none
    
    ' = &H0
    ' = &H0
End Enum '34

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
    'ifOptoElectricConvFact = &H8828
    itSensitivityType = &H8830
    itStandardOutputSensitivity = &H8831
    itRecommendedExposureIndex = &H8832
    itISOSpeed = &H8833
    itISOSpeedLatitudeYYY = &H8834
    itISOSpeedLatitudeZZZ = &H8835
    
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


    itCameraOwnerName = &HA430
    itBodySerialNumber = &HA431
    itLensSpecification = &HA432
    itLensMaker = &HA433
    itLensModel = &HA434
    itLensSerialnumber = &HA435
    'CompositeImage   = &HA????
    'SourceImageNumberOfCompositeImage     = &HA????
    'SourceExposureTimeOfCompositeImage     = &HA????
    itGamma = &HA500
          

'    MarkerSegments
'    StartOfImage = &HFFD8
'    ApplicationSegment1 = &HFFE1
'    ApplicationSegment2 = &HFFE2
'    DefineQuantizationTable = &HFFDB
'    DefineHuffmanTable = &HFFC4
'    DefineRstartInteroperability = &HFFDD
'    StartOfFrame = &HFFC0
'    StartOfScan = &HFFDA
'    EndOfImage = &HFFD9

End Enum '69

Public Function TagExif_ToStr(ByVal e As Integer) As String
    Dim s As String
    Select Case e
    Case TagGPS.itGPSVersionID:               s = "GPSVersionID"
    Case TagGPS.itGPSLatitudeRef:             s = "GPSLatitudeRef"
    Case TagGPS.itGPSLatitude:                s = "GPSLatitude"
    Case TagGPS.itGPSLongitudeRef:            s = "GPSLongitudeRef"
    Case TagGPS.itGPSLongitude:               s = "GPSLongitude"
    Case TagGPS.itGPSAltitudeRef:             s = "GPSAltitudeRef"
    Case TagGPS.itGPSAltitude:                s = "GPSAltitude"
    Case TagGPS.itGPSTimeStamp:               s = "GPSTimeStamp"
    Case TagGPS.itGPSSatellites:              s = "GPSSatellites"
    Case TagGPS.itGPSStatus:                  s = "GPSStatus"
    Case TagGPS.itGPSMeasureMode:             s = "GPSMeasureMode"
    Case TagGPS.itGPSDOP:                     s = "GPSDOP"
    Case TagGPS.itGPSSpeedRef:                s = "GPSSpeedRef"
    Case TagGPS.itGPSSpeed:                   s = "GPSSpeed"
    Case TagGPS.itGPSTrackRef:                s = "GPSTrackRef"
    Case TagGPS.itGPSTrack:                   s = "GPSTrack"
    Case TagGPS.itGPSImgDirectionRef:         s = "GPSImgDirectionRef"
    Case TagGPS.itGPSImgDirection:            s = "GPSImgDirection"
    Case TagGPS.itGPSMapDatum:                s = "GPSMapDatum"
    Case TagGPS.itGPSDestLatitudeRef:         s = "GPSDestLatitudeRef"
    Case TagGPS.itGPSDestLatitude:            s = "GPSDestLatitude"
    Case TagGPS.itGPSDestLongitudeRef:        s = "GPSDestLongitudeRef"
    Case TagGPS.itGPSDestLongitude:           s = "GPSDestLongitude"
    Case TagGPS.itGPSDestBearingRef:          s = "GPSDestBearingRef"
    Case TagGPS.itGPSDestBearing:             s = "GPSDestBearing"
    Case TagGPS.itGPSDestDistanceRef:         s = "GPSDestDistanceRef"
    Case TagGPS.itGPSDestDistance:            s = "GPSDestDistance"
    Case TagGPS.itGPSProcessingMethod:        s = "GPSProcessingMethod"
    Case TagGPS.itGPSAreaInformation:         s = "GPSAreaInformation"
    Case TagGPS.itGPSDateStamp:               s = "GPSDateStamp"
    Case TagGPS.itGPSDifferential:            s = "GPSDifferential"
    
    Case TagIF.itImageWidth:                  s = "ImageWidth"
    Case TagIF.itImageLength:                 s = "ImageLength"
    Case TagIF.itBitsPerSample:               s = "BitsPerSample"
    Case TagIF.itCompression:                 s = "Compression"
    Case TagIF.itPhotometricInterpretaion:    s = "PhotometricInterpretaion"
    Case TagIF.itImageDescription:            s = "ImageDescription"
    Case TagIF.itMaker:                       s = "Maker"
    Case TagIF.itModel:                       s = "Model"
    Case TagIF.itStripOffsets:                s = "StripOffsets"
    Case TagIF.itOrientation:                 s = "Orientation"
    Case TagIF.itSamplesPerPixel:             s = "SamplesPerPixel"
    Case TagIF.itRowsPerStrip:                s = "RowsPerStrip"
    Case TagIF.itStripByteCounts:             s = "StripByteCounts"
    Case TagIF.itXResolution:                 s = "X-Resolution"
    Case TagIF.itYResolution:                 s = "Y-Resolution"
    Case TagIF.itPlanarConfiguration:         s = "PlanarConfiguration"
    Case TagIF.itResolutionUnit:              s = "ResolutionUnit"
    Case TagIF.itTransferFunction:            s = "TransferFunction"
    Case TagIF.itSoftwareUsed:                s = "SoftwareUsed"
    Case TagIF.itDateTime:                    s = "DateTime"
    Case TagIF.itArtist:                      s = "Artist"
    Case TagIF.itWhitePoint:                  s = "WhitePoint"
    Case TagIF.itPrimaryChromaticities:       s = "PrimaryChromaticities"
    Case TagIF.itJPEGInterchangeFormat:       s = "JPEGInterchangeFormat"
    Case TagIF.itJPEGInterchangeFormatLength: s = "JPEGInterchangeFormatLength"
    Case TagIF.itYCbCrCoefficients:           s = "YCbCrCoefficients"
    Case TagIF.itYCbCrSubSampling:            s = "YCbCrSubSampling"
    Case TagIF.itYCbCrPositioning:            s = "YCbCrPositioning"
    Case TagIF.itReferenceBlackWhite:         s = "ReferenceBlackWhite"
    Case TagIF.itCopyright:                   s = "Copyright"
    
    Case TagExif.itExposureTime:              s = "ExposureTime"
    Case TagExif.itFNumber:                   s = "FNumber"
    
    Case TagIF.itIFDOffsetExif:               s = "IFDOffsetExif"
    
    Case TagExif.itExposureProgram:           s = "ExposureProgram"
    Case TagExif.itSpectralSensitivity:       s = "SpectralSensitivity"
    
    Case TagIF.itIFDOffsetGPS:                s = "IFDOffsetGPS"
        
    Case TagExif.itISOSpeedRatings:           s = "ISOSpeedRatings"
    Case TagExif.itOECF:                      s = "OECF"
    Case TagExif.itSensitivityType:           s = "SensitivityType"
    Case TagExif.itStandardOutputSensitivity: s = "StandardOutputSensitivity"
    Case TagExif.itRecommendedExposureIndex:  s = "RecommendedExposureIndex"
    Case TagExif.itISOSpeed:                  s = "ISOSpeed"
    Case TagExif.itISOSpeedLatitudeYYY:       s = "ISOSpeedLatitudeYYY"
    Case TagExif.itISOSpeedLatitudeZZZ:       s = "ISOSpeedLatitudeZZZ"
    Case TagExif.itExifVersion:               s = "ExifVersion"
    Case TagExif.itDateTimeOriginal:          s = "DateTimeOriginal"
    Case TagExif.itDateTimeDigitized:         s = "DateTimeDigitized"
    Case TagExif.itComponentsConfiguration:   s = "ComponentsConfiguration"
    Case TagExif.itCompressedBitsPerPixel:    s = "CompressedBitsPerPixel"
    Case TagExif.itShutterSpeedValue:         s = "ShutterSpeedValue"
    Case TagExif.itApertureValue:             s = "ApertureValue"
    Case TagExif.itBrightnessValue:           s = "BrightnessValue"
    Case TagExif.itExposureBiasValue:         s = "ExposureBiasValue"
    Case TagExif.itMaxApertureValue:          s = "MaxApertureValue"
    Case TagExif.itSubjectDistance:           s = "SubjectDistance"
    Case TagExif.itMeteringMode:              s = "MeteringMode"
    Case TagExif.itLightSource:               s = "LightSource"
    Case TagExif.itFlash:                     s = "Flash"
    Case TagExif.itFocalLength:               s = "FocalLength"
    Case TagExif.itSubjectArea:               s = "SubjectArea"
    Case TagExif.itMakerNote:                 s = "MakerNote"
    Case TagExif.itUserComment:               s = "UserComment"
    Case TagExif.itSubSecTime:                s = "SubSecTime"
    Case TagExif.itSubSecTimeOriginal:        s = "SubSecTimeOriginal"
    Case TagExif.itSubSecTimeDigitized:       s = "SubSecTimeDigitized"
    Case TagExif.itFlashpixVersion:           s = "FlashpixVersion"
    Case TagExif.itColorSpace:                s = "ColorSpace"
    Case TagExif.itPixelXDimension:           s = "PixelXDimension"
    Case TagExif.itPixelYDimension:           s = "PixelYDimension"
    Case TagExif.itRelatedSoundFile:          s = "RelatedSoundFile"
    
    Case TagIF.itIFDOffsetInterop:            s = "IFDOffsetInterop"
    
    Case TagExif.itFlashEnergy:               s = "FlashEnergy"
    Case TagExif.itSpatialFrequencyResponse:  s = "SpatialFrequencyResponse"
    Case TagExif.itFocalPlaneXResolution:     s = "FocalPlaneXResolution"
    Case TagExif.itFocalPlaneYResolution:     s = "FocalPlaneYResolution"
    Case TagExif.itFocalPlaneResolutionUnit:  s = "FocalPlaneResolutionUnit"
    Case TagExif.itSubjectLocation:           s = "SubjectLocation"
    Case TagExif.itExposureIndex:             s = "ExposureIndex"
    Case TagExif.itSensingMethod:             s = "SensingMethod"
    Case TagExif.itFileSource:                s = "FileSource"
    Case TagExif.itSceneType:                 s = "SceneType"
    Case TagExif.itCFAPattern:                s = "CFAPattern"
    Case TagExif.itCustomRendered:            s = "CustomRendered"
    Case TagExif.itExposureMode:              s = "ExposureMode"
    Case TagExif.itWhiteBalance:              s = "WhiteBalance"
    Case TagExif.itDigitalZoomRatio:          s = "DigitalZoomRatio"
    Case TagExif.itFocalLengthIn35mmFilm:     s = "FocalLengthIn35mmFilm"
    Case TagExif.itSceneCaptureType:          s = "SceneCaptureType"
    Case TagExif.itGainControl:               s = "GainControl"
    Case TagExif.itContrast:                  s = "Contrast"
    Case TagExif.itSaturation:                s = "Saturation"
    Case TagExif.itSharpness:                 s = "Sharpness"
    Case TagExif.itDeviceSettingDescription:  s = "DeviceSettingDescription"
    Case TagExif.itSubjectDistanceRange:      s = "SubjectDistanceRange"
    Case TagExif.itImageUniqueID:             s = "ImageUniqueID"
    Case TagExif.itCameraOwnerName:           s = "CameraOwnerName"
    Case Else: s = "unknown"
    End Select
    TagExif_ToStr = s & Space(MaxTagStrLen - Len(s))
End Function
