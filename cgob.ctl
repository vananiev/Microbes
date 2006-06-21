VERSION 5.00
Begin VB.UserControl cgob 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "cgob.ctx":0000
   Begin VB.Timer tmrFPS 
      Interval        =   200
      Left            =   1080
      Top             =   960
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "cgob.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "cgob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strControls() As Object
Dim intCount As Integer
Private mcolObjects As New Collection
Dim picFerst As StdPicture
Dim bLoad As Boolean
Public objNet As Object

Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal X As Long, ByVal у As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long _
) As Long

Private Sub tmrFPS_Timer()
    Dim lngRtn As Long
    Dim crlControl As Object
    Dim intCnt As Long
    Dim intX As Long
    Dim intY As Long
    Dim intHeight As Integer
    Dim intWidth As Integer
    If Not (UserControl.Ambient.UserMode) Then Exit Sub
    ReDim strControls(((Abs(objNet.fZ) + Abs(objNet.nZ)) + 1) * UserControl.Parent.spgob1.Count)
    For Each crlControl In UserControl.ParentControls
        If TypeName(crlControl) = "spgob" Then
            crlControl.Top = crlControl.prm.y - (crlControl.Height \ 2)
            crlControl.Left = crlControl.prm.X - (crlControl.Width \ 2)
            intX = (crlControl.prm.z + Abs(objNet.nZ)) * UserControl.Parent.spgob1.Count + crlControl.Index
            Set strControls(intX) = crlControl
        End If
    Next crlControl
    
    If Not (bLoad) Then Set picFerst = UserControl.Parent.Picture: bLoad = True
    Set UserControl.Parent.Picture = picFerst
    For intCnt = 0 To UBound(strControls)
        If Not (strControls(intCnt) Is Nothing) Then
            Set crlControl = strControls(intCnt)
            intX = crlControl.Left
            intY = crlControl.Top
            intHeight = crlControl.Height
            intWidth = crlControl.Width
            lngRtn = BitBlt(UserControl.Parent.hDC, intX, intY, intWidth, intHeight, crlControl.Negative, 0, 0, vbSrcAnd)
            lngRtn = BitBlt(UserControl.Parent.hDC, intX, intY, intWidth, intHeight, crlControl.Positive, 0, 0, vbSrcPaint)
        End If
    Next intCnt
End Sub

Public Function Add(Name As String, Parametors As Object) As Object
    Dim ObjectNew As Object
    Dim strMis As String
    Dim intCnt As Integer
    Dim bSearch As Boolean
    On Error GoTo Errors
    If UserControl.Parent.spgob1.Count = 1 And mcolObjects.Count = 0 Then
        Set ObjectNew = UserControl.Parent.spgob1(0)
        ObjectNew.prm.SName = "Ferst"
        mcolObjects.Add ObjectNew, UserControl.Parent.spgob1(0).prm.SName
    End If
    ' проверка на существование  объекта в коллекции
    intCnt = -1
    Do While Not (bSearch)
        intCnt = intCnt + 1
        On Error GoTo Errors
        strMis = UserControl.Parent.spgob1(intCnt).Top
    Loop
    Load UserControl.Parent.spgob1(intCnt)
    'запись объекта
    Set ObjectNew = UserControl.Parent.spgob1(intCnt)
    Set ObjectNew.prm = Parametors
    ObjectNew.Top = (intCnt Mod 20) * 5 + (intCnt \ 20) * 50
    ObjectNew.Left = (intCnt Mod 20) * 15 + 7
    ObjectNew.prm.SName = Name
    On Error GoTo Errors
    mcolObjects.Add ObjectNew, Name
    Set Add = ObjectNew
    Exit Function
Errors:
    Select Case Err.Number
        Case 438:   MsgBox "1:  Object spgob1(0) is not exist   ", vbCritical, "Error of object"
        Case 457:   MsgBox "2:  Name `" & Name & "` alredy exist", vbCritical, "Error of object"
                    Unload UserControl.Parent.spgob1(intCnt)
                    Set ObjectNew = Nothing
        Case 340:   bSearch = True
    End Select
    Resume Next
End Function

Public Sub Delete(Name As String)
    On Error Resume Next
    Unload UserControl.Parent.spgob1(mcolObjects.Item(Name).Index)
    mcolObjects.Remove Name
    Exit Sub
Errors:
    Select Case Err.Number
        Case 5:   MsgBox "3:  Object doesn`t delete. Name `" & Name & "` is not exist", vbCritical, "Error of object"
    End Select
End Sub

Public Function Count() As Long
    Count = mcolObjects.Count
End Function

Public Function Item(SName As String) As Object
Set Item = mcolObjects.Item(SName)
End Function

Public Function NewEnum() As IUnknown
Set NewEnum = mcolObjects.[_NewEnum]
End Function

Private Sub UserControl_Resize()
    UserControl.Height = 25 * 15
    UserControl.Width = 25 * 15
End Sub

Public Sub SetFps(Optional Intervel = 200)
    tmrFPS.Interval = Interval
End Sub
