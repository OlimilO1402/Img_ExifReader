VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Exxifer"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   7695
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
   Begin VB.ComboBox CBFileName 
      Height          =   315
      ItemData        =   "Form1.frx":1782
      Left            =   1110
      List            =   "Form1.frx":1784
      TabIndex        =   2
      ToolTipText     =   "Select or dragdrop file here"
      Top             =   90
      Width           =   6735
   End
   Begin VB.TextBox TBExifData 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_IFFile As TaggedImageFile
'

'Private Sub BtnEquals_Click()
'    Dim dt As Single: dt = Timer
'    Dim p   As String:  p = App.Path & "\Resources\Exif.org\examples\"
'    Dim p1  As String: p1 = p & "IntelLittleEndian\"
'    Dim p2  As String: p2 = p & "MotorolaBigEndian\"
'
'    Dim PFN1 As String: PFN1 = p1 & "sony_DSC-HX400V.jpg" '"canon-ixus.jpg"
'    Dim PFN2 As String: PFN2 = p2 & "sony_DSC-HX400V_testingEquals.jpg" '"kodak-dc210.jpg" '"canon-ixus.jpg"
'
'    Dim tif1 As TaggedImageFile: Set tif1 = MNew.TaggedImageFile(PFN1)
'    Dim tif2 As TaggedImageFile: Set tif2 = MNew.TaggedImageFile(PFN2)
'    tif1.Read: tif2.Read
'    Dim B As Boolean: B = tif1.Equals(tif2)
'    dt = Timer - dt
'    MsgBox "Sind die beiden Dateien gleich? " & B & vbCrLf & dt
'
'    'ca 0,04 sek
'    '=> ca 25 Bilder / sek
'
'End Sub

'Private Sub Command1_Click()
'    'Dim l As Integer: l = -32768
'    'Dim u As Integer: u = 32767
'    ReDim tags(0 To 65535) As String
'    Dim i As Integer
'    Dim s As String
'    'Dim s1 As String
'    'Dim s2 As String
'    'Dim s3 As String
'    'On Error Resume Next
'
'    Dim maxL As Long
'    Dim ui As Long
'    For ui = 0 To 65535
'
'        i = CInt("&H" & Hex(ui))
'
'        s = ""
'
'        s = Trim(MTagExif.TagExif_ToStr(i))
'        If s = "unknown" Then
'            s = Trim(MTagGPS.TagGPS_ToStr(i))
'            If s = "unknown" Then
'                s = Trim(MTagIF.TagIF_ToStr(i))
'                If s = "unknown" Then
'                    s = ""
'                Else
'                    s = "TagIF.it" & s
'                End If
'            Else
'                s = "TagGPS.it" & s
'            End If
'        Else
'            s = "TagExif.it" & s
'        End If
'        If Len(s) Then s = "    Case " & s & " = &H" & Hex(i) & ":" & Space(37 - Len(s)) & "s = " & """" & s & """" & vbCrLf
'        tags(ui) = s
'
'    Next
'
'    s = "    Select Case e" & vbCrLf
'    s = s & Join(tags, "")
'    s = s & "    End Select"
'    TBExifData.Text = s & vbCrLf
'
'End Sub
'

Private Sub Form_Load()
    
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    
    Dim p   As String:  p = App.Path & "\Resources\Exif.org\examples\"
    Dim p1  As String: p1 = p & "IntelLittleEndian\"
    Dim p2  As String: p2 = p & "MotorolaBigEndian\"
    
    With CBFileName
        .AddItem p1 & "canon-ixus.jpg"
        .AddItem p1 & "canon-powershota5.jpg"
        .AddItem p1 & "Canon-PowerShotA40.jpg"
        .AddItem p1 & "Canon-PowerShot-S5-IS.JPG"
        .AddItem p1 & "fujifilm-dx10.jpg"
        .AddItem p1 & "fujifilm-mx1700.jpg"
        .AddItem p1 & "olympus-c960.jpg"
        .AddItem p1 & "olympus-d320l.jpg" 'only JFIF no Exif-data
        .AddItem p1 & "sanyo-vpcg250.jpg"
        .AddItem p1 & "sanyo-vpcsx550.jpg"
        .AddItem p1 & "sony_DSC-HX400V.jpg"
        .AddItem p1 & "sony-cybershot.jpg"
        
        .AddItem p2 & "fujifilm-finepix40i.jpg"
        .AddItem p2 & "kodak-dc210.jpg"
        .AddItem p2 & "kodak-dc240.jpg"
        .AddItem p2 & "nikon-D7000.jpg"
        .AddItem p2 & "nikon-e950.jpg"    'JFIF with Exif-data
        .AddItem p2 & "ricoh-rdc5300.jpg"
        .AddItem p2 & "sony-d700.jpg"
        .ListIndex = 0 '16
    End With
    'CBFileName.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Dim l As Single, t As Single, W As Single, H As Single
    Dim brdr As Single: brdr = 1 * Screen.TwipsPerPixelX
    l = brdr: t = BtnRead.Top: BtnRead.Move l, t
    l = BtnRead.Left + BtnRead.Width + brdr: t = CBFileName.Top
    W = Me.ScaleWidth - l - brdr: H = CBFileName.Height
    If W > 0 And H > 0 Then CBFileName.Move l, t, W
    l = TBExifData.Left:          t = TBExifData.Top: brdr = 0
    W = Me.ScaleWidth - l - brdr: H = Me.ScaleHeight - t - brdr
    If W > 0 And H > 0 Then TBExifData.Move l, t, W, H
End Sub

Private Sub CBFileName_Click()
    BtnRead_Click
End Sub

Private Sub BtnRead_Click()
    TBExifData.Text = ""
    Dim PFN As String
    PFN = CBFileName.Text
    Set m_IFFile = MNew.TaggedImageFile(PFN)
    If m_IFFile.Read Then
        TBExifData.Text = m_IFFile.ToStr
    Else
        Dim e As String: e = MError.LastError
        If Len(e) Then MsgBox e
        TBExifData.Text = "No exif data"
        'If Len(m_IFFile.ErrorInfo) Then MsgBox m_IFFile.ErrorInfo
    End If
End Sub

Private Sub CBFileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmPicViewer
End Sub

Private Sub TBExifData_DblClick()
    CBFileName_DblClick
End Sub

Private Sub CBFileName_DblClick()
    FrmPicViewer.Show
    FrmPicViewer.ShowPicture CBFileName.Text
End Sub

Private Sub TBExifData_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub OnOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    If Data.Files.Count = 0 Then Exit Sub
    CBFileName.Text = Data.Files(1)
    BtnRead.Value = True
End Sub

