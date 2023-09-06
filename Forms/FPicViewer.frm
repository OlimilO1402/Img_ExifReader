VERSION 5.00
Begin VB.Form FPicViewer 
   Caption         =   "PicViewer"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   Icon            =   "FPicViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "FPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Pic As PicturePlus
'


Private Sub Form_Load()
    Set m_Pic = New PicturePlus
End Sub

Public Sub ShowPicture(aFileName As String)
    Set Picture1.Picture = m_Pic.LoadPicturePlus(aFileName)
End Sub

Private Sub Picture1_Resize()
    Me.ScaleHeight = Picture1.ScaleHeight
    Me.ScaleWidth = Picture1.ScaleWidth
    'Me.Move Me.Left, Me.Top, Picture1.ScaleHeight
End Sub
