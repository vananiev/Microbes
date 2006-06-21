VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSkor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Установите скорость времени"
   ClientHeight    =   2040
   ClientLeft      =   5115
   ClientTop       =   5205
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin MSComctlLib.Slider sldSkor 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   873
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   100
      TickFrequency   =   20
      Value           =   100
   End
   Begin VB.Label Label2 
      Caption         =   "<< Fast       Low>>"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "frmSkor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Hide
End Sub

Private Sub Form_Load()
  sldSkor.SelStart = Heard.tmrOfLife.Interval / 10
  Label1.Caption = Str(sldSkor.Value) & " %"
End Sub

Private Sub sldSkor_Change()
Label1.Caption = Str(sldSkor.Value) & " %"
If sldSkor = 100 Then Label1.Caption = "Now"
End Sub

