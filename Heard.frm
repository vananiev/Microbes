VERSION 5.00
Begin VB.Form Heard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Life"
   ClientHeight    =   6720
   ClientLeft      =   3480
   ClientTop       =   2220
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9480
   Begin VB.Timer tmrOfLife 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   6240
   End
   Begin VB.Label lblCount 
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label lblTimer 
      Height          =   255
      Left            =   8160
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Menu File 
      Caption         =   "����"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu a1 
         Caption         =   "-"
      End
      Begin VB.Menu Open 
         Caption         =   "�������"
      End
      Begin VB.Menu Save 
         Caption         =   "Coxpa����"
      End
   End
   Begin VB.Menu Prav 
      Caption         =   "������"
      Begin VB.Menu CMicrobs 
         Caption         =   "��������� ��������"
      End
      Begin VB.Menu CMicrob 
         Caption         =   "��������� �������"
      End
      Begin VB.Menu CInvir 
         Caption         =   "��������� �����"
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu Skoroct 
         Caption         =   "�������� �������"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "������"
   End
End
Attribute VB_Name = "Heard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMicrobs As New Microbs
Dim objMicrob As New Microb
Dim objInvir As New Invironment
Dim objNet As New Net
Dim intCount As Integer
Dim btNumMic  As Integer
Dim Circles(10000, 1) As Integer

Private Sub CInvir_Click()
frmMicrob.Show
frmMicrob.cmdOk = False
frmMicrob.rtbInfor.Text = objInvir.Count
 frmMicrob.rtbInfor.Text = objInvir.Item("1").Himic
For Each objNet In objInvir
With objNet
    frmMicrob.txtName = "�����"
    frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "������������:           " & .Light \ 255 * 100 & " %"
    frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "����:                " & .Hidros \ 255 * 100 & " %"
    frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "��������:                " & .Kislorod \ 255 * 100 & " %"
    frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�����������:             " & .Uglcisl \ 255 * 100 & " %"
    frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�����������:    " & .Temprt - 100 & " �" & vbCrLf
End With
Next objNet
End Sub

Private Sub CMicrob_Click()
Dim btbyte As Integer
Dim intWrk As Integer
Dim intCount2 As Integer
    frmMicrob.Show vbModal
    frmMicrob.Show
For Each objMicrob In objMicrobs
    With objMicrob
        If (Val(.Name) = Val(frmMicrob.txtName)) Then
            frmMicrob.rtbInfor.Text = ""
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���:                " & .Name
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�������:           " & Round(.Age, 2)
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�����:             " & .Mass
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�������:          " & .Kalories
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���������� x:  " & .X
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���������� y:  " & .Y & vbCrLf
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���:   "
            For intCount2 = 1 To 2
                For intCount = 1 To 16
                    For btbyte = -3 To 0
                        intWrk = (.DNK(intCount, intCount2) And 3 * 4 ^ Abs(btbyte)) / 4 ^ Abs(btbyte)
                        If intWrk = 0 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 1 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 2 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 3 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                    Next btbyte
                Next intCount
                frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "            "
            Next intCount2
       End If
    End With
Next objMicrob
If frmMicrob.rtbInfor.Text = "" Then frmMicrob.rtbInfor.Text = "������ � ����� ������ �� ����������"
End Sub

Private Sub CMicrobs_Click()
Dim strWork As String
Dim btbyte As Integer
Dim intWrk As Integer
Dim intCount2 As Integer
frmMicrobs.Show
Information.Show
frmMicrob.rtbInfor.Text = ""
frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & "����� ����������:       " & objMicrobs.Count & vbCrLf
For Each objMicrob In objMicrobs
With objMicrob
    intWrk = Val(.Name) / objMicrobs.Count * 100
     If intWrk > 100 Then intWrk = 100
    Information.Caption = "���������� " & Str(intWrk) & "%"
    Information.prbPr.Value = intWrk
    DoEvents
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "���:                " & .Name
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "�������:           " & Round(.Age, 2)
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "�����:             " & .Mass
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "�������:          " & .Kalories
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "���������� x:  " & .X
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "���������� y:  " & .Y & vbCrLf
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "���:   "
    For intCount2 = 1 To 2
        For intCount = 1 To 16
            For btbyte = -3 To 0
                intWrk = (.DNK(intCount, intCount2) And 3 * 4 ^ Abs(btbyte)) / 4 ^ Abs(btbyte)
                If intWrk = 0 Then frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & "�-"
                If intWrk = 1 Then frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & "�-"
                If intWrk = 2 Then frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & "�-"
                If intWrk = 3 Then frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & "�-"
            Next btbyte
        Next intCount
    frmMicrobs.rtbInfor.Text = frmMicrobs.rtbInfor.Text & vbCrLf & "            "
    Next intCount2
End With
Next objMicrob
Unload Information
End Sub

Private Sub Form_Load()
Dim n As Long
Randomize Timer
' ���������� ��� �����
Show
'���������� ����� - ��������
frmSplash.Show
DoEvents
For n = 1 To 200000#: Print "": Next n
' ������� �����-��������
Unload frmSplash
Cls

NewNum:
NumMicrobs.Show vbModal
btNumMic = Val(NumMicrobs.Text1)
If btNumMic > 10000 Or btNumMic < 1 Then MsgBox "�������� �����", vbCritical, "Error of number": GoTo NewNum
objInvir.SetWeth
objMicrobs.SetMicrobs (btNumMic)
tmrOfLife.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intWork As Integer
Dim intX As Integer
Dim intY As Integer
Dim intCount2 As Integer
Dim btbyte As Integer
Dim intWrk As Byte
If Button = 1 And Shift = 0 Then
    For intWork = 1 To btNumMic
        If Circles(Str(intWork), 0) = -1 Then GoTo NextFotot
        intX = Circles(Str(intWork), 0) - X
        intY = Circles(Str(intWork), 1) - Y
        If 10000 > (intX ^ 2 + intY ^ 2) Then '���� ���� ���������������� > �����. �������
        Set objMicrob = objMicrobs.Item(Str(intWork))
        With objMicrob
            frmMicrob.Show
            frmMicrob.txtName = .Name
            frmMicrob.rtbInfor.Text = ""
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���:                " & .Name
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�������:           " & Round(.Age, 2)
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�����:             " & .Mass
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "�������:          " & .Kalories
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���������� x:  " & .X
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���������� y:  " & .Y & vbCrLf
            frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "���:   "
            For intCount2 = 1 To 2
                For intCount = 1 To 16
                    For btbyte = -3 To 0
                        intWrk = (.DNK(intCount, intCount2) And 3 * 4 ^ Abs(btbyte)) / 4 ^ Abs(btbyte)
                        If intWrk = 0 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 1 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 2 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                        If intWrk = 3 Then frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & "�-"
                    Next btbyte
                Next intCount
                frmMicrob.rtbInfor.Text = frmMicrob.rtbInfor.Text & vbCrLf & "            "
            Next intCount2
        End With
        Exit Sub
        End If
NextFotot:
    Next intWork
End If
End Sub

Private Sub New_Click()
Randomize Timer
NewNum:
NumMicrobs.Show vbModal
btNumMic = Val(NumMicrobs.Text1)
If btNumMic > 10000 Or btNumMic < 1 Then MsgBox "�������� �����", vbCritical, "Error of number": GoTo NewNum

For Each objMicrob In objMicrobs
    objMicrobs.Delete (objMicrob.Name)
Next objMicrob
Cls
objMicrobs.SetMicrobs (btNumMic)
lblTimer.Caption = ""
End Sub

Private Sub Skoroct_Click()
frmSkor.Show vbModal
tmrOfLife.Interval = frmSkor.sldSkor.Value * 10
Unload frmSkor
End Sub

Private Sub tmrClear_Timer()
Cls
End Sub

Private Sub tmrOfLife_Timer()
lblTimer = Str(Val(lblTimer) + 1) & "  ���"
lblCount = objMicrobs.Count
Dim intCoord As Integer '��������� �� ���. ����� 80*60
For Each objMicrob In objMicrobs
    With objMicrob
    '������
        intCoord = Int(.Y / 120) * 60 + Int(.X / 120) + 1 '��������� �� ���. ����� 80*60
        If (.DNK(2, 1) \ 2 + .DNK(2, 2) \ 2) = objInvir.Item(Str(intCoord)).Himic Then KillMicrob .Name, .X, .Y '�������� ���. �-����
        If (.DNK(3, 1) \ 2 + .DNK(3, 2) \ 2) < objInvir.Item(Str(intCoord)).Temprt Then KillMicrob .Name, .X, .Y '�������� �����. ������������
        If (.DNK(4, 1) \ 2 + .DNK(4, 2) \ 2) > objInvir.Item(Str(intCoord)).Temprt Then KillMicrob .Name, .X, .Y '�������� �������
        If (.DNK(5, 1) \ 2 + .DNK(5, 2) \ 2) < objInvir.Item(Str(intCoord)).Uglcisl Then KillMicrob .Name, .X, .Y '�������� CO2
        If (.DNK(6, 1) \ 2 + .DNK(6, 2) \ 2) > objInvir.Item(Str(intCoord)).Kislorod Then KillMicrob .Name, .X, .Y '��. ������. O2
        If (.DNK(7, 1) \ 2 + .DNK(7, 2) \ 2) > objInvir.Item(Str(intCoord)).Hidros Then KillMicrob .Name, .X, .Y '��. ���������.H2O
        If (.DNK(8, 1) \ 2 + .DNK(8, 2) \ 2) < objInvir.Item(Str(intCoord)).Light Then KillMicrob .Name, .X, .Y '��. ������
        If (.DNK(10, 1) \ 2 + .DNK(10, 2) \ 2) * 2 < .Age Then KillMicrob .Name, .X, .Y '������ �� ��������
        If .Kalories = 0 Then KillMicrob .Name, .X, .Y '�������� �������
        '����������
        If (.DNK(13, 1) \ 128) * (.DNK(13, 2) \ 128) = 1 Then
        Else
            .Kalories = .Kalories + ((.DNK(9, 1) \ 2 + .DNK(9, 2) \ 2) * objInvir.Item(Str(intCoord)).Light \ 255) \ 4 '��� ��� ����� � ������ ��������. ������. �� 64
            If .Mass + ((.DNK(9, 1) \ 2 + .DNK(9, 2) \ 2) * objInvir.Item(Str(intCoord)).Light \ 255) \ 256 < (.DNK(1, 1) \ 64 + .DNK(1, 2) \ 64) Then .Mass = .Mass + ((.DNK(9, 1) \ 2 + .DNK(9, 2) \ 2) * objInvir.Item(Str(intCoord)).Light \ 255) \ 256
        End If
    End With
Next objMicrob
'��������
MoveMicr
'�����������
Dublicate
If objMicrobs.Count = 0 Then tmrOfLife.Enabled = False
 End Sub

Private Sub KillMicrob(Name As String, X As Integer, Y As Integer)
objMicrobs.Item(Name).setDNK 16, 1, 0
objMicrobs.Item(Name).setDNK 16, 2, 0
End Sub

Public Sub MoveMicr()
Dim intX As Long '���������� ����.X
Dim intY As Long '���������� ����. Y
Dim intR As Long '���������� ��������� ��������
Dim intCoord As Integer
Dim intWork As Integer
Dim intWork2 As Integer
Dim intSee As Integer

intWork2 = 12
For Each objMicrob In objMicrobs
    With objMicrob
        intCoord = Int(.Y / 120) * 60 + Int(.X / 120) + 1 '��������� �� ���. ����� 80*60
        If (.DNK(16, 1) + .DNK(16, 2)) = 0 Then GoTo Drw
         intR = Rnd * (.DNK(14, 1) \ 2 + .DNK(14, 2) \ 2)
         
          '��������� ��������
        intSee = MicrobSee(objMicrob)
        If intSee <= 0 Then intSee = intCoord
        If (.DNK(13, 1) \ 128) * (.DNK(13, 2) \ 128) = 0 Then '����������
            If objInvir.Item(Str(intSee)).Micrb = 1 Then '�������� �� ��������
                If (intSee - intCoord) < 0 Then intWork2 = 12.5
                If (intSee - intCoord) > 0 Then intWork2 = 11.5
            End If
        Else
            If objInvir.Item(Str(intSee)).Micrb = 0 Then '������������� ���������� ���������
                 If (intSee - intCoord) < 0 Then intWork2 = 11.5
                 If (intSee - intCoord) > 0 Then intWork2 = 12.5
                If (.DNK(16, 1) + .DNK(16, 2)) > 0 Then
                    For intWork = 1 To btNumMic
                        If Circles(Str(intWork), 0) = -1 Then GoTo NextFotot
                        intX = Circles(Str(intWork), 0) - .X
                        intY = Circles(Str(intWork), 1) - .Y
                        If (.DNK(11, 1) \ 2 + .DNK(11, 1) \ 2) ^ 2 > (intX ^ 2 + intY ^ 2) Then '���� ���� ���������������� > �����. �������
                            If (objMicrobs.Item(Str(intWork)).DNK(13, 1) \ 128) * (objMicrobs.Item(Str(intWork)).DNK(13, 2) \ 128) = 0 Then '���� ��������� ������ ��������
Eat:
                                Circle (.X, .Y), .Mass * 25, -2147483633  '����. ������� ��������� �������
                                Circle (objMicrobs.Item(Str(intWork)).X, objMicrobs.Item(Str(intWork)).Y), objMicrobs.Item(Str(intWork)).Mass * 25, -2147483633 '����. ������� ��������� ���������
                                .Kalories = .Kalories + objMicrobs.Item(Str(intWork)).Kalories * 0.9
                                If .Mass + objMicrobs.Item(Str(intWork)).Mass * 0.5 < (.DNK(1, 1) \ 64 + .DNK(1, 2) \ 64) Then .Mass = .Mass + objMicrobs.Item(Str(intWork)).Mass * 0.5
                                objMicrobs.Delete (Str(intWork))
                                Circles(Str(intWork), 0) = -1
                                Circles(Str(intWork), 1) = -1
                                Circle (.X, .Y), .Mass * 25, 16711680  '������ 1
                                GoTo ExitFind
                            End If
                            If (objMicrobs.Item(Str(intWork)).DNK(16, 1) + objMicrobs.Item(Str(intWork)).DNK(16, 2)) = 0 Then GoTo Eat '���� ��������� ������ ����
                            If objMicrobs.Item(Str(intWork)).Mass < .Mass Then GoTo Eat
                        End If
NextFotot:
                    Next intWork
                End If
ExitFind:
            End If
        End If
        
NewCrd:
        intX = Rnd * intR * Sgn(Int(Rnd * Rnd * Rnd * intWork2 - Rnd))
        intY = Sqr(intR * intR - intX * intX) * Sgn(Int(Rnd * Rnd * Rnd * intWork2 - Rnd))
        If .X + intX >= 9500 Then GoTo NewCrd
        If .X + intX <= 0 Then GoTo NewCrd
        If .Y + intY >= 6700 Then GoTo NewCrd
        If .Y + intY <= 0 Then GoTo NewCrd
        
        '����. ������� ���������
        Circle (.X, .Y), .Mass * 25, -2147483633
        
        .X = .X + intX
        .Y = .Y + intY
        intCoord = Int(.Y / 120) * 60 + Int(.X / 120) + 1 '��������� �� ���. ����� 80*60

        .Kalories = .Kalories - 1
        
Drw:
        .Age = .Age + 1
        '����������
        Circles(Val(.Name), 0) = .X
        Circles(Val(.Name), 1) = .Y
        If (.DNK(16, 1) + .DNK(16, 2)) = 0 Then
            If .Age > 100 Then '������� ������
                '����. ������� ���������
                Circle (.X, .Y), .Mass * 25, -2147483633
                Circles(Val(.Name), 0) = -1
                Circles(Val(.Name), 1) = -1
                objMicrobs.Delete (.Name)
            Else
                Circle (.X, .Y), .Mass * 25, 255 '����
                objInvir.Item(Str(intCoord)).Micrb = 0
            End If
        Else
            If (.DNK(13, 1) \ 128) * (.DNK(13, 2) \ 128) = 1 Then
                Circle (objMicrob.X, objMicrob.Y), .Mass * 25, 16711680  '������ 1
                objInvir.Item(Str(intCoord)).Micrb = 1
            Else
                Circle (objMicrob.X, objMicrob.Y), .Mass * 25, 32535     '���������� 0
                objInvir.Item(Str(intCoord)).Micrb = 0
            End If
        End If
    End With
Next objMicrob
End Sub

Private Function MicrobSee(objMicrob As Microb) As Integer
Dim intWork As Integer
Dim intX As Integer
Dim intY As Integer
    With objMicrob
    For intWork = 1 To btNumMic
        If Circles(Str(intWork), 0) = -1 Then GoTo NextFotot
        intX = Circles(Str(intWork), 0) - .X
        intY = Circles(Str(intWork), 1) - .Y
        If (.DNK(11, 1) \ 2 + .DNK(11, 1) \ 2) ^ 2 > (intX ^ 2 + intY ^ 2) Then '���� ���� ���������������� > �����. �������
               MicrobSee = Int(.Y / 120) * 60 + Int(.X / 120) + 1
        End If
NextFotot:
    Next intWork
End With
End Function

Private Sub Dublicate()
Dim intCntF As Integer
Dim intCntS As Integer
Dim intX As Integer '���������� ����.X
Dim intY As Integer '���������� ����. Y
For Each objMicrob In objMicrobs
    With objMicrob
        If .Age > 100 And .Age < 104 And (.DNK(16, 1) + .DNK(16, 2)) > 0 Then
            btNumMic = btNumMic + 1
            '����. ������� ���������
            Circle (.X, .Y), .Mass * 25, -2147483633
            objMicrobs.Add (Str(btNumMic))
            objMicrobs.Item(Str(btNumMic)).Age = 0
            objMicrobs.Item(Str(btNumMic)).Mass = Rnd + 1
            If .Mass - 2 > 0 Then .Mass = .Mass - 2
            objMicrobs.Item(Str(btNumMic)).Kalories = 600
            If .Kalories - 100 > 0 Then .Kalories = .Kalories - 100
            objMicrobs.Item(Str(btNumMic)).X = .X
            objMicrobs.Item(Str(btNumMic)).Y = .Y
             For intCntF = 1 To 15
                 For intCntS = 1 To 2
                    objMicrobs.Item(Str(btNumMic)).setDNK intCntF, intCntS, Rnd * 255
                Next intCntS
            Next intCntF
            objMicrobs.Item(Str(btNumMic)).setDNK 16, 1, 1
            objMicrobs.Item(Str(btNumMic)).setDNK 16, 2, 1
            objMicrobs.Item(Str(btNumMic)).setDNK 3, 1, 150
            objMicrobs.Item(Str(btNumMic)).setDNK 3, 2, 150
            objMicrobs.Item(Str(btNumMic)).setDNK 4, 1, 50
            objMicrobs.Item(Str(btNumMic)).setDNK 4, 2, 50
            objMicrobs.Item(Str(btNumMic)).setDNK 6, 1, 50
            objMicrobs.Item(Str(btNumMic)).setDNK 6, 2, 50
            objMicrobs.Item(Str(btNumMic)).setDNK 9, 1, 200
            objMicrobs.Item(Str(btNumMic)).setDNK 9, 2, 200
            objMicrobs.Item(Str(btNumMic)).setDNK 11, 1, Rnd * 100 + 150
            objMicrobs.Item(Str(btNumMic)).setDNK 11, 2, Rnd * 100 + 150
        End If
    End With
Next objMicrob
End Sub
