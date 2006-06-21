VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl spgob 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   Picture         =   "spgob.ctx":0000
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   ToolboxBitmap   =   "spgob.ctx":0C42
   Begin VB.Timer tmrMoveStep 
      Interval        =   200
      Left            =   3360
      Top             =   600
   End
   Begin MSComctlLib.ImageList imlNAnim 
      Index           =   0
      Left            =   840
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlAnim 
      Index           =   0
      Left            =   120
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picPict 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picNPict 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1680
      ScaleHeight     =   89
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "spgob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mintPic As Integer
Private mobjPrm As Object
Public Move As Integer

Private Sub tmrMoveStep_Timer()
    If imlAnim(Move).ListImages.Count = 0 Then
        Exit Sub
    Else
        mintPic = mintPic + 1
        If mintPic = imlAnim(Move).ListImages.Count + 1 Then mintPic = 1
    End If
    picPict.Picture = imlAnim(Move).ListImages.Item(mintPic).Picture
    picNPict.Picture = imlNAnim(Move).ListImages.Item(mintPic).Picture
    UserControl.Height = imlAnim(Move).ListImages.Item(mintPic).Picture.Height \ 1.76
    UserControl.Width = imlAnim(Move).ListImages.Item(mintPic).Picture.Width \ 1.76
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 3
    picPict.Top = 0
    picPict.Left = 0
    picPict.Height = UserControl.ScaleHeight
    picPict.Width = UserControl.ScaleWidth
    picNPict.Top = 0
    picNPict.Left = 0
    picNPict.Height = picPict.Height
    picNPict.Width = picPict.Width
End Sub

Private Sub DoNPicture(Group As Integer, Index As Integer, Optional Key)
    Dim intX As Integer
    Dim intY As Integer
    Set picPict.Picture = imlAnim(Move).ListImages.Item(Index).Picture
    UserControl.Height = imlAnim(Move).ListImages.Item(Index).Picture.Height \ 1.76
    UserControl.Width = imlAnim(Move).ListImages.Item(Index).Picture.Width \ 1.76
    UserControl_Resize
    For intX = 0 To picPict.Width - 1
        For intY = 0 To picPict.Height - 1
            If picPict.Point(intX, intY) = vbBlack Then
                picNPict.PSet (intX, intY), vbWhite
            Else
                picNPict.PSet (intX, intY), vbBlack
            End If
        Next intY
    Next intX
    imlNAnim(Group).ListImages.Add Index, Key, picNPict.Image
End Sub

Public Sub AddPicture(Group As Integer, Index As Integer, NewPicture As StdPicture, Optional Key)
    If Index > imlAnim(Group).ListImages.Count + 1 Then Index = imlAnim(Group).ListImages.Count + 1
    If Index <= imlAnim(Group).ListImages.Count Then RemPicture Group, Index
    imlAnim(Group).ListImages.Add Index, Key, NewPicture
    DoNPicture Group, Index, Key
End Sub

Public Sub RemPicture(Group As Integer, Optional Index As Integer, Optional Key As String)
    If Key = "" And Index = 0 Then
        MsgBox "Picture isn`t delete, because Index or Key isn`t known", vbCritical, " Error"
        Exit Sub
    End If
    If Key = "" Then
        imlAnim(Group).ListImages.Remove Index
        imlNAnim(Group).ListImages.Remove Index
    Else
        imlAnim(Group).ListImages.Remove Key
        imlNAnim(Group).ListImages.Remove Key
    End If
End Sub

Public Sub ClsPicture(Group As Integer)
    imlAnim(Group).ListImages.Clear
    imlNAnim(Group).ListImages.Clear
End Sub

Public Function CntPicture(Group As Integer) As Integer
    CntPicture = imlAnim(Group).ListImages.Count
End Function

Public Sub AddGroup(Optional Group = "Next")
    If Group = "Next" Then Group = imlAnim.Count
    Load imlAnim(Group)
    Load imlNAnim(Group)
End Sub

Public Sub RemGroup(Optional Group = "Next")
    If Group = "Next" Then Group = imlAnim.Count - 1
    Unload imlAnim(Group)
    Unload imlNAnim(Group)
End Sub

Public Sub MoveStep(Optional Interval = 200, Optional Enable = True)
    tmrMoveStep.Interval = Interval
    tmrMoveStep.Enabled = Enable
End Sub

Public Function Positive() As Long
    Positive = picPict.hDC
End Function

Public Function Negative() As Long
    Negative = picNPict.hDC
End Function


Public Static Property Get Prm() As Object
    Set Prm = mobjPrm
End Property

Public Static Property Set Prm(ByVal objNewValue As Object)
    Set mobjPrm = objNewValue
End Property
