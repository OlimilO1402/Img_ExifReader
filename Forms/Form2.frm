VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   9915
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Form2"
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
