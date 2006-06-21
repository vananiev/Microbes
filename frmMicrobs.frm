VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMicrobs 
   Caption         =   "Состояние микробов"
   ClientHeight    =   5370
   ClientLeft      =   3615
   ClientTop       =   4500
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   10770
   Begin RichTextLib.RichTextBox rtbInfor 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMicrobs.frx":0000
   End
End
Attribute VB_Name = "frmMicrobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    rtbInfor.Height = Abs(frmMicrobs.Height - 500)
    rtbInfor.Width = Abs(frmMicrobs.Width - 100)
End Sub
