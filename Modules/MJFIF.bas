Attribute VB_Name = "MJFIF"
Option Explicit
'https://de.wikipedia.org/wiki/JPEG_File_Interchange_Format
Public Enum TagJFIF
    jtSOI = &HFFD8     ' Start Of Image
    jtAPP0 = &HFFE0    ' JFIF tag
    
    'jtSOFn = &HFFCn   ' Start of Frame Marker, legt Art der Kompression fest:
    jtSOF0 = &HFFC0    ' Baseline DCT
    jtSOF1 = &HFFC1    ' Extended sequential DCT
    jtSOF2 = &HFFC2    ' Progressive DCT
    jtSOF3 = &HFFC3    ' Lossless (sequential)
    
    jtDHT = &HFFC4     ' Definition der Huffman-Tabellen
    
    jtSOF5 = &HFFC5    ' Differential sequential DCT
    jtSOF6 = &HFFC6    ' Differential progressive DCT
    jtSOF7 = &HFFC7    ' Differential lossless (sequential)
    
    jtJPG = &HFFC8     ' reserviert für JPEG extensions
    
    jtSOF9 = &HFFC9    ' Extended sequential DCT
    jtSOF10 = &HFFCA   ' Progressive DCT
    jtSOF11 = &HFFCB   ' Lossless (sequential)
    
    jtSOF13 = &HFFCD   ' Differential sequential DCT
    jtSOF14 = &HFFCE   ' Differential progressive DCT
    jtSOF15 = &HFFCF   ' Differential lossless (sequential)
    
    jtDAC = &HFFCC     ' Definition der arithmetischen Codierung
    
    jtDQT = &HFFDB     ' Definition der Quantisierungstabellen
    jtDRI = &HFFDD     ' Define Restart Interval
    
    jtAPP1 = &HFFE1    ' Exif-Daten
    
    jtAPP14 = &HFFEE   ' Oft für Copyright-Einträge
    
    'jtAPPn = &HFFEn   ' n=2..F allg. Zeiger
    
    jtCOM = &HFFFE     ' Kommentare
    jtSOS = &HFFDA     ' Start of Scan
    jtEOI = &HFFD9     ' End of Image
    
End Enum
    
Public Const C_JFIFHeader As String = "JFIF" & vbNullChar
Public Const C_JFXXHeader As String = "JFXX" & vbNullChar

'APP0-marker ist mandatory right after the SOI-marker
Public Type APP0MarkerJFIF
    length             As Integer 'Total APP0 field byte count, including the byte count value (2 bytes), but excluding the APP0 marker itself
    identifier(0 To 4) As Byte    '= X'4A', X'46', X'49', X'46', X'00' This zero terminated string ("JFIF") uniquely identifies this APP0 marker. This string shall have zero parity (bit 7=0).
    version(0 To 1)    As Byte    '= X'0102' The most significant byte is used for major revisions, the least significant byte for minor revisions. Version 1.02 is the current released revision.
    units              As Byte    '0: no units, X and Y specify the pixel aspect ratio
                                  '1: X and Y are dots per inch
                                  '2: X and Y are dots per cm
    Xdensity           As Integer 'Horizontal pixel density
    Ydensity           As Integer 'Vertical   pixel density
    Xthumbnail         As Byte    '
    Ythumbnail         As Byte    '
    'n = Xthumbnail * Ythumbnail
    RGB()              As Byte    '3n bytes 'Packed (24-bit) RGB values for the thumbnail pixels
End Type

'Color Space YCbCr (CCIR 601 256 levels)
'Do not gamma correct
'if using only 1 component, use Y

Public Type App0MarkerJFXX
    length             As Integer 'Total APP0 field byte count, including the byte count value (2 bytes), but excluding the APP0 marker itself
    identifier(0 To 4) As Byte    '= X'4A', X'46', X'58', X'58', X'00' This zero terminated string ("JFIF") uniquely identifies this APP0 marker. This string shall have zero parity (bit 7=0).
    extension_code     As Byte    'Code which identifies the extension.  In this version, the following extensions are defined:
                                  ' &H10:   Thumbnail coded using JPEG
                                  ' &H11:   Thumbnail stored using 1 byte/pixel
                                  ' &H13:   Thumbnail stored using 3 bytes/pixel
    extension_data()   As Byte
    '
End Type

'convert YCbCr from and to RGB
' Y = 256 *   E'y
'Cb = 256 * [ E'Cb ] + 128
'Cr = 256 * [ E'Cr ] + 128

'E'y : 0 ... 1
'E'Cb, E'Cr :-0.5 ... +0.5
'Y, Cb, and Cr must be clamped to 255 when they are maximum value

'RGB to YCbCr Conversion
'YCbCr (256 levels) can be computed directly from 8-bit RGB as follows:
'Y   =     0.299  R + 0.587  G + 0.114  B
'Cb  =   - 0.1687 R - 0.3313 G + 0.5    B + 128
'Cr  =     0.5    R - 0.4187 G - 0.0813 B + 128
'
'YCbCr to RGB Conversion
'RGB can be computed directly from YCbCr (256 levels) as follows:
'R = Y                    + 1.402   (Cr-128)
'G = Y - 0.34414 (Cb-128) - 0.71414 (Cr-128)
'B = Y + 1.772   (Cb-128)
