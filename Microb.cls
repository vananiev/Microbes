VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Microb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name As String
Public Age As Single
Public Mass As Long
Public Kalories As Single
Public X As Integer
Public Y As Integer
Private mDNA() As Byte

Public Property Get TrDNA(Triplet As Long) As Byte
    DNA = mDNA(Triplet)
End Property
Public Sub setTrDNA(Triplet As Long, Value As Byte, Optional LenOfDNA = 100)
    If UBound(mDNA) < 0 Then ReDim mDNA(Int(Abs(LenOfDNA)))
    If UBound(mDNA) > Triplet Then MsgBox "Triplet " & Triplet & " is not exist", vbOKOnly, "DNA error"
    mDNA(Triplet) = DNA
End Sub
