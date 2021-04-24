VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PicturePlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Edanmo
'http://www.mvps.org/emorcillo/en/code/vb6/loadimagegdip.shtml

' ----==== API Declarations ====----

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Declare Function GdiplusStartup _
    Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    
Private Declare Function GdiplusShutdown _
    Lib "GDIPlus" (ByVal token As Long) As Long
    
Private Declare Function GdipCreateBitmapFromFile _
    Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
    
Private Declare Function GdipDisposeImage _
    Lib "GDIPlus" (ByVal image As Long) As Long
    
Private Declare Function GdipCreateHBITMAPFromBitmap _
    Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long

Private Declare Sub OleCreatePictureIndirect _
    Lib "oleaut32" (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, lplpvObj As Object)

'------------------------------------------------------
' Procedure : LoadPicturePlus
' Purpose   : Loads an image using GDI+
' Returns   : The image loaded in a StdPicture object
' Author    : Eduardo A. Morcillo
'------------------------------------------------------
'
Public Function LoadPicturePlus(ByVal filename As String) As StdPicture

   ' Initialize GDI+
    Dim tSI As GdiplusStartupInput
    tSI.GdiplusVersion = 1
    Dim lRes As Long
    Dim lGDIP As Long
    lRes = GdiplusStartup(lGDIP, tSI)
    
    If lRes = 0 Then
    
        ' Open the image file
        Dim lBitmap As Long
        lRes = GdipCreateBitmapFromFile(StrPtr(filename), lBitmap)
        
        If lRes = 0 Then
        
            Dim hBitmap As Long
            
            ' Create a GDI bitmap
            lRes = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
            
            ' Create the StdPicture object
            Set LoadPicturePlus = HandleToPicture(hBitmap, vbPicTypeBitmap)
            
            ' Dispose the image
            GdipDisposeImage lBitmap
            
        End If
        
        ' Shutdown GDI+
        GdiplusShutdown lGDIP
    End If
    
    If lRes Then Err.Raise 5, , "Cannot load file"
   
End Function

'------------------------------------------------------
' Procedure : HandleToPicture
' Purpose   : Creates a StdPicture object to wrap a GDI
'             image handle
'------------------------------------------------------
'
Public Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As PictureTypeConstants, Optional ByVal hPal As Long = 0) As StdPicture
    
    ' Initialize the PICTDESC structure
    Dim tPictDesc As PICTDESC
    With tPictDesc
       .cbSizeOfStruct = Len(tPictDesc)
       .picType = ObjectType
       .hgdiObj = hGDIHandle
       .hPalOrXYExt = hPal
    End With
    
    ' Initialize the IPicture interface ID
    Dim IID_IPicture As IID
    With IID_IPicture
       .Data1 = &H7BF80981
       .Data2 = &HBF32
       .Data3 = &H101A
       .Data4(0) = &H8B
       .Data4(1) = &HBB
       .Data4(3) = &HAA
       .Data4(5) = &H30
       .Data4(6) = &HC
       .Data4(7) = &HAB
    End With
    
    'Create the object
    Dim oPicture As IPicture
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
   
    ' Return the picture object
    Set HandleToPicture = oPicture
    
End Function

