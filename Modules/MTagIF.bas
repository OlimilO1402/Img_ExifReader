Attribute VB_Name = "MTagIF"
Option Explicit

Public Enum TagIF
    itInteropIndex = &H1
    
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
    itMake = &H10F                          ' manufacturer
    itModel = &H110                         ' Image input equipment model
    itSoftwareUsed = &H131                  ' Software used
    itArtist = &H13B                        ' Person who created the image
    'itSoftware = &H10E
    itCopyright = &H8298
    
    
    
    itExifIFDOffset = &H8769
    itGPSIFDOffset = &H8825 '34853
    itInteropIFDOffset = &HA005
    
'GPS Info IFD Pointer
'Tag = 34853 (8825.H)
'Type = LONG
'Count = 1
'Default = none
    
    ' = &H0
    ' = &H0
End Enum

'könnte man in einem Array speichern oder in einer Collection Keys sind die Enum-Konstanten als Hex-String
Public Function TagIF_ToStr(ByVal this As TagIF) As String
    'das könnte man auch über eine Collection machen, da es sehr viele sein werden
    Dim s As String
    Select Case this
  'Tags relating to image data structure
    Case TagIF.itImageWidth:                    s = "ImageWidth"
    Case TagIF.itImageLength:                   s = "ImageLength"
    Case TagIF.itBitsPerSample:                 s = "BitsPerSample"
    Case TagIF.itCompression:                   s = "Compression"
    Case TagIF.itPhotometricInterpretaion:      s = "PhotometricInterpretaion"
    Case TagIF.itOrientation:                   s = "Orientation"
    Case TagIF.itSamplesPerPixel:               s = "SamplesPerPixel"
    Case TagIF.itPlanarConfiguration:           s = "PlanarConfiguration"
    Case TagIF.itYCbCrSubSampling:              s = "YCbCrSubSampling"
    Case TagIF.itYCbCrPositioning:              s = "YCbCrPositioning"
    Case TagIF.itXResolution:                   s = "X-Resolution"
    Case TagIF.itYResolution:                   s = "Y-Resolution"
    Case TagIF.itResolutionUnit:                s = "ResolutionUnit"
    
  'Tags relating to recording offset
    Case TagIF.itStripOffsets:                  s = "StripOffsets"
    Case TagIF.itRowsPerStrip:                  s = "RowsPerStrip"
    Case TagIF.itStripByteCounts:               s = "StripByteCounts"
    Case TagIF.itJPEGInterchangeFormat:         s = "JPEGInterchangeFormat"
    Case TagIF.itJPEGInterchangeFormatLength:   s = "JPEGInterchangeFormatLength"
    
  'Tags relating to image data characteristics
    Case TagIF.itTransferFunction:              s = "TransferFunction"
    Case TagIF.itWhitePoint:                    s = "WhitePoint"
    Case TagIF.itPrimaryChromaticities:         s = "PrimaryChromaticities"
    Case TagIF.itYCbCrCoefficients:             s = "YCbCrCoefficients"
    Case TagIF.itReferenceBlackWhite:           s = "ReferenceBlackWhite"
    
  'Other tags
    Case TagIF.itDateTime:                      s = "DateTime"
    Case TagIF.itImageDescription:              s = "ImageDescription"
    Case TagIF.itMake:                          s = "Make"
    Case TagIF.itModel:                         s = "Model"
    Case TagIF.itSoftwareUsed:                  s = "SoftwareUsed"
    Case TagIF.itArtist:                        s = "Artist"
    Case TagIF.itCopyright:                     s = "Copyright"
    
    Case TagIF.itExifIFDOffset:                 s = "IFDOffsetExif"
    Case TagIF.itGPSIFDOffset:                  s = "IFDOffsetGPS"
    Case TagIF.itInteropIFDOffset:              s = "IFDOffsetInterop"
    'Case IFTag
    Case Else:                                  s = "unkown"
    End Select
    TagIF_ToStr = s & Space(MaxTagStrLen - Len(s))
End Function

