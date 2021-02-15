VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Exxifer"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   7695
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TBFileName 
      Height          =   285
      Left            =   1560
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox TBExifData 
      BeginProperty Font 
         Name            =   "Courier New"
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
   Begin VB.CommandButton BtnRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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

Private Sub Form_Load()
    'TBFileName = "\\SOLS_DS\Daten\Stuff_saves\Bilder\ASUS ZenUI_2020-07\2020\P_20200117_092514.jpg"
    Dim p   As String:     p = App.Path & "\Exif.org\examples\"
    Dim fnm As String
    
    fnm = "canon-ixus.jpg"
    'fnm = "Canon-PowerShotA40.jpg"
    'fnm = "canon-powershota5.jpg" 'nur JFIF keine Exif-daten
    'fnm = "Canon-PowerShot-S5-IS.JPG"
    'fnm = "fujifilm-dx10.jpg"
    'fnm = "fujifilm-finepix40i.jpg" 'Motorola big endian Integer
    'fnm = "fujifilm-mx1700.jpg"
    'fnm = "kodak-dc210.jpg" 'Motorola big endian Integer
    'fnm = "kodak-dc240.jpg" 'Motorola big endian Integer
    'fnm = "nikon-e950.jpg" 'nur JFIF keine Exif-daten
    'fnm = "olympus-c960.jpg"
    'fnm = "olympus-d320l.jpg" 'nur JFIF keine Exif-daten
    'fnm = "ricoh-rdc5300.jpg" 'nur JFIF keine Exif-daten
    'fnm = "sanyo-vpcg250.jpg"
    'fnm = "sanyo-vpcsx550.jpg"
    'fnm = "sony-cybershot.jpg"
    'fnm = "sony-d700.jpg"
    TBFileName.Text = p & fnm
End Sub

Private Sub BtnRead_Click()
    TBExifData.Text = ""
    Dim pfn As String
    pfn = TBFileName.Text
    Set m_IFFile = MNew.TaggedImageFile(pfn)
    If m_IFFile.Read Then
        TBExifData.Text = m_IFFile.ToStr
    Else
        Dim e As String: e = MError.LastError
        If Len(e) Then MsgBox e
        'If Len(m_IFFile.ErrorInfo) Then MsgBox m_IFFile.ErrorInfo
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    L = TBFileName.Left:          t = TBFileName.Top
    W = Me.ScaleWidth - L - brdr: H = TBFileName.Height
    If W > 0 And H > 0 Then TBFileName.Move L, t, W, H
    L = Me.TBExifData.Left:       t = Me.TBExifData.Top: brdr = 0
    W = Me.ScaleWidth - L - brdr: H = Me.ScaleHeight - t - brdr
    If W > 0 And H > 0 Then Me.TBExifData.Move L, t, W, H
End Sub

Private Sub TBFileName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub TBExifData_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub OnOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    If Data.Files.count = 0 Then Exit Sub
    TBFileName.Text = Data.Files(1)
    BtnRead.Value = True
End Sub

