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
Public Light As Byte 'освещенность
Public Temprt As Byte 'температура
Public Kislorod As Byte 'кислорол
Public Uglcisl As Byte 'углекислота
Public Hidros As Byte 'вода
Public Himic As Byte 'хим.в-ва
Public Micrb As Byte '0-травоядный 1-хищник
