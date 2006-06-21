VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMicrob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Состояние микробa"
   ClientHeight    =   4725
   ClientLeft      =   2790
   ClientTop       =   3360
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin RichTextLib.RichTextBox rtbInfor 
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6800
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMicrob.frx":0000
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Введите имя микроба"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "frmMicrob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Hide
    cmdOk.Enabled = False
End Sub
