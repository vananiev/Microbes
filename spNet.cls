VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Net"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name As String
Public Light As Byte '������������
Public Temprt As Byte '�����������
Public Kislorod As Byte '��������
Public Uglcisl As Byte '�����������
Public Hidros As Byte '����
Public Himic As Byte '���.�-��
Public fZ As Long
Public nZ As Long

Private Sub Class_Initialize()
    fZ = 100
    nZ = -100
End Sub