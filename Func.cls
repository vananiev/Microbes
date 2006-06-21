VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim objNet As New Net

Public Sub LoadObj(Num As Integer)
    Dim intCnt As Integer
    For intCnt = 1 To Num
        frmMain.cgob1.Add Str(intCnt), New Param
        frmMain.spgob1(intCnt).addPicture 0, 1, LoadPicture("E:\Program Files\Microsoft Visual Studio\VB98\Program Files\Микробы\Image\Microb_1.GIF")
        frmMain.spgob1(intCnt).movestep
    Next intCnt
    objNet.Himic = 100
End Sub

Public Sub Dead(objA As Object, Reason As String)
    objA.prm.DeadReason = Reason
    objA.movestep , False
End Sub

Public Sub Eating(objA As Object)
    frmMain.cgob1.Delete objA.prm.SName
End Sub

Public Sub DoZord(objA As Object)
    objA.prm.ZOrd = (objA.prm.Z \ ((Abs(objNet.fZ) + Abs(objNet.nZ)) / frmMain.cgob1.Count)) + 1
    a = objA.prm.ZOrd
    b = objA.prm.Z
End Sub
