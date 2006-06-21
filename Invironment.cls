VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Invironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Форма 9600*7200
Dim intCount As Integer
Dim btWork As Byte
Private mcolNet As New Collection

Public Function Add(strName As String) As Net
Dim NetNew As New Net
NetNew.Name = strName
mcolNet.Add NetNew, strName
Set Add = NetNew
End Function

Public Sub Delete(strName As String)
mcolNet.Remove strName
End Sub

Public Function Count() As Long
Count = mcolNet.Count
End Function

Public Function Item(strName As String) As Net
Attribute Item.VB_UserMemId = 0
Set Item = mcolNet.Item(strName)
End Function

Public Function NewEnum() As IUnknown
Set NewEnum = mcolNet.[_NewEnum]
End Function

Public Sub SetWeth()
Randomize Timer
Dim objNet As Net
For intCount = 1 To 4800
Set objNet = Add(Str(intCount))
With objNet
    .Light = 20
NewTemp:
    .Temprt = 120
    .Kislorod = 200
    .Uglcisl = 5
    .Himic = Rnd * 255
    .Hidros = 150
End With
Next intCount
End Sub
