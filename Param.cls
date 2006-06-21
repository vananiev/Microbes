VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public SName As String
Public Age As Single
Public Mass As Long
Public Kalories As Single
Public SpeedMove As Integer
Public Prey As Byte
Public DeadReason As String
Public X As Integer
Public Y As Integer
Public Z As Integer
Public ZOrd As Integer
Private mDNA(100) As Byte

Public Property Get TrDNA(Triplet As Long) As Byte
    TrDNA = mDNA(Triplet)
End Property

Public Property Let setTrDNA(Triplet As Long, Value As Byte)
    If UBound(mDNA) < Triplet Then MsgBox "Triplet " & Triplet & " is not exist", vbOKOnly, "DNA error"
    mDNA(Triplet) = Value
End Property

Private Sub Class_Initialize()
    Age = 100
    Mass = 1
    Kalories = 1000
    SpeedMove = 2
    X = Rnd * 600
    Y = Rnd * 400
    Z = 1
End Sub
