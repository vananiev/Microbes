VERSION 5.00
Begin VB.Form NumMicrobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Введите количество микробов"
   ClientHeight    =   1470
   ClientLeft      =   6510
   ClientTop       =   5955
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3975
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "50"
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "NumMicrobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Hide
End Sub
