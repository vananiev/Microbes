VERSION 5.00
Object = "{A16332AD-79A8-450A-BA49-19873FB846B0}#11.0#0"; "simpGraphObg.ocx"
Object = "*\A..\FAFF~1\ControlSimplgob.vbp"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDo 
      Interval        =   20
      Left            =   9000
      Top             =   6120
   End
   Begin ControlSimplgob.cgob cgob1 
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin simpGraphObg.spgob spgob1 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Dim objF As New Func
Dim objNet As New Net

Private Sub Form_Load()
    Set cgob1.objNet = objNet
    Set spgob1(0).prm = New Param
    spgob1(0).addPicture 0, 1, LoadPicture("E:\Program Files\Microsoft Visual Studio\VB98\Program Files\Микробы\Image\Microb_1.GIF")
    spgob1(0).movestep
    Set spgob1(0).prm = New Param
    objF.LoadObj 50
    Randomize Timer
End Sub

Private Sub tmrDo_Timer()
    Dim objA As Object
    Dim strWork As String
    For Each objA In Controls
        If TypeName(objA) = "spgob" Then
            If objA.prm.DeadReason = "" Then
                With objA.prm
                    .Age = .Age - tmrDo.Interval / 1000
                    .Kalories = .Kalories - (.Prey + 1) * .Mass
                    .X = .X + Rnd * .SpeedMove - (.SpeedMove \ 2)
                    .Y = .Y + Rnd * .SpeedMove - (.SpeedMove \ 2)
                    strWork = .Z
                    .Z = .Z + Rnd * .SpeedMove - (.SpeedMove \ 2)
                    If .Z > objNet.fZ Then
                        strWork = .Z
                        .Z = objNet.fZ
                        strWork = .Z
                    End If
                    If .Z < objNet.nZ Then .Z = objNet.nZ
                    If .X > 580 Then .X = 580
                    If .X < 20 Then .X = 20
                    If .Y > 380 Then .Y = 380
                    If .Y < 20 Then .Y = 20
                    
                    If .Age < 0 Then objF.Dead objA, "Dead of Age"
                    If .Kalories < 0 Then objF.Dead objA, "Dead of Kalories"
                End With
            End If
        End If
    Next objA
End Sub

