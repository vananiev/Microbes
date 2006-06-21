VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Microbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim intCount As Integer
Private mcolMicrobs As New Collection

Public Function Add(strName As String) As Microb
Dim MicrobNew As New Microb
MicrobNew.Name = strName
mcolMicrobs.Add MicrobNew, strName
Set Add = MicrobNew
End Function

Public Sub Delete(strName As String)
mcolMicrobs.Remove strName
End Sub

Public Function Count() As Long
Count = mcolMicrobs.Count
End Function

Public Function Item(strName As String) As Microb
Attribute Item.VB_UserMemId = 0
Set Item = mcolMicrobs.Item(strName)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mcolMicrobs.[_NewEnum]
End Function

Public Sub SetMicrobs(Count As Integer)
Randomize Timer
Dim btWork As Byte
Dim intCntF As Integer
Dim intCntS As Integer
Dim intCoord As Integer
Dim MicrobN As Microb
For intCount = 1 To Count
    Set MicrobN = Add(Str(intCount))
    With MicrobN
    .Age = 0
    .Mass = Rnd + 1
    .Kalories = 600
    .X = Rnd * 9000
    .Y = Rnd * 6500
    For intCntF = 1 To 15
        For intCntS = 1 To 2
           .setDNK intCntF, intCntS, Rnd * 255
        Next intCntS
    Next intCntF
    .setDNK 16, 1, 1
    .setDNK 16, 2, 1
    .setDNK 3, 1, 150
    .setDNK 3, 2, 150
    .setDNK 4, 1, 50
    .setDNK 4, 2, 50
    .setDNK 6, 1, 50
    .setDNK 6, 2, 50
    .setDNK 9, 1, 200
    .setDNK 9, 2, 200
    .setDNK 11, 1, Rnd * 100 + 150
    .setDNK 11, 2, Rnd * 100 + 150
    End With
Next intCount
End Sub
