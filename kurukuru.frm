VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�N���N���R�b�V�["
   ClientHeight    =   5835
   ClientLeft      =   3555
   ClientTop       =   2310
   ClientWidth     =   6255
   FillStyle       =   0  '�h��Ԃ�
   BeginProperty Font 
      Name            =   "System"
      Size            =   13.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "kurukuru.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'Z ���ް
   ScaleHeight     =   389
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   417
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3600
      Picture         =   "kurukuru.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3600
      Picture         =   "kurukuru.frx":060C
      ScaleHeight     =   435
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   240
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3600
      Picture         =   "kurukuru.frx":0CCA
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   4440
      Picture         =   "kurukuru.frx":12B0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   5160
      Picture         =   "kurukuru.frx":1BF2
      ScaleHeight     =   915
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   240
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   3960
      Picture         =   "kurukuru.frx":2034
      Top             =   5040
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3960
      Picture         =   "kurukuru.frx":2596
      Top             =   4200
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   31
      X2              =   208
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Menu m01 
      Caption         =   "�Q�[��(&G)"
      Begin VB.Menu m11 
         Caption         =   "�X�^�[�g(&N)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu m17 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu m16 
         Caption         =   "�ꎞ��~(&P)"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu m14 
         Caption         =   "-"
      End
      Begin VB.Menu m02 
         Caption         =   "�N���N���R�b�V�[�̏I��(&X)"
      End
   End
   Begin VB.Menu m03 
      Caption         =   "���[��(&R)"
      Begin VB.Menu m35 
         Caption         =   "�t�B�[���h�̍L��(&F)"
         Begin VB.Menu m36 
            Caption         =   "����(&N)"
            Index           =   0
         End
         Begin VB.Menu m36 
            Caption         =   "�������炢(&M)"
            Index           =   1
         End
         Begin VB.Menu m36 
            Caption         =   "�L��(&W)"
            Index           =   2
         End
      End
      Begin VB.Menu m04 
         Caption         =   "�F�̐�(&C)"
         Begin VB.Menu m25 
            Caption         =   "&2"
            Index           =   0
         End
         Begin VB.Menu m25 
            Caption         =   "&3"
            Index           =   1
         End
         Begin VB.Menu m25 
            Caption         =   "&4"
            Index           =   2
         End
         Begin VB.Menu m25 
            Caption         =   "&5"
            Index           =   3
         End
         Begin VB.Menu m25 
            Caption         =   "&6"
            Index           =   4
         End
         Begin VB.Menu m25 
            Caption         =   "&10"
            Index           =   5
         End
         Begin VB.Menu m25 
            Caption         =   "&256"
            Index           =   6
         End
      End
      Begin VB.Menu m05 
         Caption         =   "�Ȃ��鐔(&J)"
         Begin VB.Menu m26 
            Caption         =   "&2"
            Index           =   0
         End
         Begin VB.Menu m26 
            Caption         =   "&3"
            Index           =   1
         End
         Begin VB.Menu m26 
            Caption         =   "&4"
            Index           =   2
         End
         Begin VB.Menu m26 
            Caption         =   "&5"
            Index           =   3
         End
         Begin VB.Menu m26 
            Caption         =   "&6"
            Index           =   4
         End
         Begin VB.Menu m26 
            Caption         =   "&10"
            Index           =   5
         End
      End
      Begin VB.Menu m29 
         Caption         =   "�ȂȂ߂�(&D)"
         Begin VB.Menu m30 
            Caption         =   "�Ȃ���(&T)"
            Index           =   0
         End
         Begin VB.Menu m30 
            Caption         =   "�Ȃ���Ȃ�(&F)"
            Index           =   1
         End
      End
      Begin VB.Menu m31 
         Caption         =   "�܂�����(&O)"
         Begin VB.Menu m32 
            Caption         =   "�Ȃ���(&T)"
            Index           =   0
         End
         Begin VB.Menu m32 
            Caption         =   "�Ȃ���Ȃ�(&F)"
            Index           =   1
         End
      End
      Begin VB.Menu m18 
         Caption         =   "���Ԑ���(&T)"
         Visible         =   0   'False
         Begin VB.Menu m19 
            Caption         =   "�Ȃ�(&F)"
            Index           =   0
         End
         Begin VB.Menu m19 
            Caption         =   "����(&T)"
            Index           =   1
         End
      End
      Begin VB.Menu m23 
         Caption         =   "���уR�b�V�[�̏o���p�x(&M)"
         Begin VB.Menu m24 
            Caption         =   "�o�Ȃ�(&D)"
            Index           =   0
         End
         Begin VB.Menu m24 
            Caption         =   "���܂ɏo��(&T)"
            Index           =   1
         End
         Begin VB.Menu m24 
            Caption         =   "�΂�΂�o��(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu m33 
         Caption         =   "-"
      End
      Begin VB.Menu m34 
         Caption         =   "������Ԃɖ߂�(&I)"
      End
   End
   Begin VB.Menu m22 
      Caption         =   "�I�v�V����(&O)"
      Begin VB.Menu m12 
         Caption         =   "�R�b�V�[�̍~�鑬��(&S)"
         Begin VB.Menu m13 
            Caption         =   "�������(&S)"
            Index           =   0
         End
         Begin VB.Menu m13 
            Caption         =   "���炽��(&N)"
            Index           =   1
         End
         Begin VB.Menu m13 
            Caption         =   "���[��(&F)"
            Index           =   2
         End
         Begin VB.Menu m13 
            Caption         =   "���[��(&V)"
            Index           =   3
         End
         Begin VB.Menu m13 
            Caption         =   "�΁[��(&M)"
            Index           =   4
         End
      End
      Begin VB.Menu m20 
         Caption         =   "���L�[���������Ƃ�(&D)"
         Begin VB.Menu m21 
            Caption         =   "���킶�헎����(&S)"
            Index           =   0
         End
         Begin VB.Menu m21 
            Caption         =   "��C�ɗ�����(&F)"
            Index           =   1
         End
      End
      Begin VB.Menu m27 
         Caption         =   "���ʉ�(&E)"
         Begin VB.Menu m28 
            Caption         =   "&ON"
            Index           =   0
         End
         Begin VB.Menu m28 
            Caption         =   "O&FF"
            Index           =   1
         End
      End
      Begin VB.Menu m37 
         Caption         =   "&NEXT�R�b�V�[�̃v���t�F�b�`"
         Begin VB.Menu m38 
            Caption         =   "���Ȃ�(&N)"
            Index           =   0
         End
         Begin VB.Menu m38 
            Caption         =   "&1�g����"
            Index           =   1
         End
         Begin VB.Menu m38 
            Caption         =   "&2�g�܂�"
            Index           =   2
         End
         Begin VB.Menu m38 
            Caption         =   "&3�g�܂�"
            Index           =   3
         End
         Begin VB.Menu m38 
            Caption         =   "&4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu m38 
            Caption         =   "&5"
            Index           =   5
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu m06 
      Caption         =   "�w���v(&H)"
      Begin VB.Menu m07 
         Caption         =   "�g�s�b�N�̌���(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu m10 
         Caption         =   "������@(&K)..."
      End
      Begin VB.Menu m08 
         Caption         =   "-"
      End
      Begin VB.Menu m09 
         Caption         =   "�o�[�W�������(&A)..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' �N���N���R�b�V�[
'
'
'  �쐬�F�z�� �N��
'
'
Option Explicit


Const Kabe = -2
Const Kara = -1

Const TenmetsuDelayTime = 0.3

Dim StageXSize
Dim StageYSize

Const StageXSizeMaxLimit = 15
Const StageYSizeMaxLimit = 20

Const ScoreDrawX = 500  '�A���\���̎n�܂�X
Const ScoreDrawY = 110  '�A���\���̎n�܂�Y


Const DrawNextColorMax = 5 '�z��̂��߂̍ō��l ���ۂ�3�܂�

Dim ScoreHideX

Dim FieldSizeScoreXHosei   '�t�B�[���h�T�C�Y�ɂ��X�R�A�\��X�̕␳

'          x  y
' small    4  7
' middle   6 11
' large    8 15

'Z(x,y) -2 �� , -1 ��A�������̂�1000�ȏ�
Dim Z(0 To StageXSizeMaxLimit + 1, -1 To StageYSizeMaxLimit + 1)
Dim KesuZ(1 To StageXSizeMaxLimit, 0 To StageYSizeMaxLimit)

Dim Ochiteru
Dim px(0 To 1), py(0 To 1)        ' �R�b�V�[�̂��A��
Dim NowColor(0 To 1)
Dim NextColor(0 To 1, 0 To 50)    ' ���̃R�b�V�[�̐F 50�͂Ȃ�ƂȂ�
Dim PuyoMuki                      ' �R�b�V�[�̌��� 0 �� 1 �� 2 �� 3 �E

Dim ColorMax
Dim Narabe
Dim ChibiKuruRitu                 ' ���уR�b�V�[���o�闦
Dim Mageru                        ' �Ȃ���̂�F�߂�
Dim PuyoColor(255)
Dim Score
Dim NokoriTime                    ' �^�C���A�^�b�N���̎c��^�C��
Dim Monokuro                      ' ���m�N���f�B�X�v���C���[�h
Dim Koma                          ' 1��̃^�C�}�[�ŉ��R�}�����邩
Dim QuickTurnFlag
Dim NextColorPrefetch             ' Next�R�b�V�[�̃v���t�F�b�`���������邩
'
' �R�b�V�[�������܂��B
'
Public Sub PutPuyo(x, y, c)

    Select Case c
    Case Kara
        DrawKara x, y
    Case Kabe
        DrawKabe x, y
    Case Is < 1000         '1000�����Ȃ畁�ʂ̃R�b�V�[������
        DrawKara x, y

        Form1.FillStyle = 0
        Form1.FillColor = PuyoColor(c)
        Circle (x * 30 + 15, y * 30 + 15), 14, PuyoColor(c), , , 0.95
        Form1.FillColor = 0
        Circle (x * 30 + 15 - 5, y * 30 + 15 - 5), 3, 0, , , 2
        Circle (x * 30 + 15 + 5, y * 30 + 15 - 5), 3, 0, , , 2
        If Monokuro Then         ' ���m�N�����[�h�Ȃ��
            Select Case c Mod 4  ' 4��ނ��ƂɌ��̌`��ς���
            Case 0
                Form1.FillStyle = 1
                Circle (x * 30 + 15, y * 30 + 15 + 2), 8, 0, -3.14, -0.01
            Case 1
                Form1.FillStyle = 0
                Circle (x * 30 + 15, y * 30 + 15 + 2), 8, 0, -3.14, -0.01
            Case 2
                Form1.FillStyle = 1
                Circle (x * 30 + 15, y * 30 + 15 + 6), 8, 0, , , 0.6
            Case 3
                Form1.FillStyle = 0
                Circle (x * 30 + 15, y * 30 + 15 + 6), 8, 0, , , 0.6
            End Select
        Else
            Form1.FillStyle = 1
            Circle (x * 30 + 15, y * 30 + 15 + 2), 8, 0, -3.14, -0.01
        End If
    Case Is >= 1000               ' ���уR�b�V�[������
        DrawKara x, y

        Form1.FillStyle = 0
        Form1.FillColor = PuyoColor(c - 1000)
        Circle (x * 30 + 15, y * 30 + 15), 8, PuyoColor(c - 1000), , , 0.95
        Form1.FillColor = 0
        Circle (x * 30 + 15 - 3, y * 30 + 15 - 3), 2, 0, , , 2
        Circle (x * 30 + 15 + 3, y * 30 + 15 - 3), 2, 0, , , 2
        If Monokuro Then
            Select Case c Mod 4  ' 4��ނ��ƂɌ��̌`��ς���
            Case 0
                Form1.FillStyle = 1
                Circle (x * 30 + 15, y * 30 + 15 + 2), 4, 0, -3.14, -0.01
            Case 1
                Form1.FillStyle = 0
                Circle (x * 30 + 15, y * 30 + 15 + 2), 4, 0, -3.14, -0.01
            Case 2
                Form1.FillStyle = 1
                Circle (x * 30 + 15, y * 30 + 15 + 4), 3, 0, , , 0.6
            Case 3
                Form1.FillStyle = 0
                Circle (x * 30 + 15, y * 30 + 15 + 4), 3, 0, , , 0.6
            End Select
        Else
            Form1.FillStyle = 1
            Circle (x * 30 + 15, y * 30 + 15 + 2), 4, 0, -3.14, -0.01
        End If
    End Select

End Sub

'
' �L�[���������Ƃ��̏���
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim PlusX, PlusY
    Dim ToPuyoMuki
    Dim i
    Dim OtosuKoma
    Static QuickTurnKeyCode

    'Debug.Print KeyCode, Shift, PuyoMuki

    If Ochiteru And m16.Checked = False Then
        Select Case KeyCode
        Case vbKeyS, vbKeyD, vbKeyNumpad8
            ImanoPuyoKesu
            QuickTurn
            ImanoPuyoKaku
        Case vbKeyC, vbKeyZ, vbKeyX, vbKeyNumpad5
            ' ��]
            Select Case KeyCode
            Case vbKeyNumpad5, vbKeyX
                ToPuyoMuki = (PuyoMuki - 1) And 3
            Case vbKeyZ, vbKeyC
                ToPuyoMuki = (PuyoMuki + 1) And 3
            End Select
            
            Select Case ToPuyoMuki
            Case 0
                PlusX = 0
                PlusY = -1
            Case 1
                PlusX = -1
                PlusY = 0
            Case 2
                PlusX = 0
                PlusY = 1
            Case 3
                PlusX = 1
                PlusY = 0
            End Select
            If Z(px(0) + PlusX, int2(py(0) + PlusY)) = Kara Then ' �܂킷
                PuyoMuki = ToPuyoMuki
                ImanoPuyoKesu
                px(1) = px(0) + PlusX
                py(1) = py(0) + PlusY
                ImanoPuyoKaku
            Else
                If PlusX <> 0 And Z(px(0) + -PlusX, int2(py(0))) = Kara Then '����Ă܂킷
                    PuyoMuki = ToPuyoMuki
                    ImanoPuyoKesu
                    px(1) = px(0)
                    py(1) = py(0)
                    px(0) = px(0) + -PlusX
                    ImanoPuyoKaku
                Else ' �܂킹�Ȃ��Ƃ�
                        ' �N�C�b�N�^�[�����
                    If Z(px(0) - 1, int2(py(0))) <> Kara And Z(px(1) - 1, int2(py(1))) <> Kara And Z(px(0) + 1, int2(py(0))) <> Kara And Z(px(1) + 1, int2(py(1))) <> Kara Then
                        'Beep  'for Debug
                        If QuickTurnFlag = True And KeyCode = QuickTurnKeyCode Then
                            ' �N�C�b�N�^�[��
                            ImanoPuyoKesu
                            QuickTurn
                            ImanoPuyoKaku
                            QuickTurnFlag = False
                        Else
                            QuickTurnKeyCode = KeyCode
                            QuickTurnFlag = True
                        End If
                    Else
                        ' ���ɉ������邽�߂ɉ񂹂Ȃ�
                    End If
                End If
            End If
        Case 39, 102                      ' �E�ֈړ�   �� = 39 , �e���L�[6 = 102
            If Z(px(0) + 1, int2(py(0))) = Kara And Z(px(1) + 1, int2(py(1))) = Kara Then
                ImanoPuyoKesu
                For i = 0 To 1
                    px(i) = px(i) + 1
                Next
                ImanoPuyoKaku
            End If
        Case 37, 100                      ' ���ֈړ�   �� = 37 , �e���L�[4 = 100
            If Z(px(0) - 1, int2(py(0))) = Kara And Z(px(1) - 1, int2(py(1))) = Kara Then
                ImanoPuyoKesu
                For i = 0 To 1
                    px(i) = px(i) - 1
                Next
                ImanoPuyoKaku
            End If
        Case 40, 98                       ' ���Ƃ�   �� = 40 , �e���L�[2 = 98
            OtosuKoma = Koma              ' ���Ƃ��R�}�̐������߂�
            OtosuKoma = 1 / 2            ' ��������1�ɂ���BVersion 54.05���� ' ��������1 / 2�ɂ���BVersion 54.06����
            ImanoPuyoKesu
            If m21(0).Checked = True Then
                If Z(px(0), int2(py(0) + OtosuKoma)) = Kara And Z(px(1), int2(py(1) + OtosuKoma)) = Kara Then
                    For i = 0 To 1
                        py(i) = py(i) + OtosuKoma
                    Next
                Else
                    ' ����DO�͍ő��񂵂����Ȃ��B
                    Do While (Z(px(0), int2(py(0) + OtosuKoma)) = Kara And Z(px(1), int2(py(1) + OtosuKoma)) = Kara)
                        For i = 0 To 1
                            py(i) = py(i) + OtosuKoma
                        Next
                    Loop
                    For i = 0 To 1
                        py(i) = int2(py(i))
                    Next
                End If
            Else
                Do While (Z(px(0), int2(py(0) + OtosuKoma)) = Kara And Z(px(1), int2(py(1) + OtosuKoma)) = Kara)
                    For i = 0 To 1
                        py(i) = py(i) + OtosuKoma
                    Next
                Loop
                For i = 0 To 1
                    py(i) = int2(py(i))
                Next
            End If
            ImanoPuyoKaku
        End Select
    End If

End Sub

'
' ��ԍŏ���
'
Private Sub Form_Load()
    
    Randomize
    
    ' ���W�X�g�����珉���l�ǂݍ���
    
    m25_Click Val(GetSetting("Kurukuru", "Config", "Color", 2))
    m26_Click Val(GetSetting("Kurukuru", "Config", "Narabe", 1))
    m24_Click Val(GetSetting("Kurukuru", "Config", "ChibiKuru", 1))
    m13_Click Val(GetSetting("Kurukuru", "Config", "KuruDownSpeed", 0))
    m21_Click Val(GetSetting("Kurukuru", "Config", "DownKey", 0))
    m19_Click Val(GetSetting("Kurukuru", "Config", "GameMode", 0))
    m28_Click Val(GetSetting("Kurukuru", "Config", "SoundEffect", 0))
    m30_Click Val(GetSetting("Kurukuru", "Config", "Naname", 1))
    m32_Click Val(GetSetting("Kurukuru", "Config", "Magaru", 0))
    m36_Click Val(GetSetting("Kurukuru", "Config", "FieldSize", 0))
    m38_Click Val(GetSetting("Kurukuru", "Config", "NextColorPrefetch", 1))
    
    Mageru = True
    Monokuro = (InStr(LCase(Command$), "/m") > 0)
    
    GameJunbi

    MakeField
    
    ' Form����ʒ����Ɉړ�����BGameJunbi��Form�̑傫����ς������ƁB
    Form1.Left = (Screen.Width - Form1.Width) / 2
    Form1.Top = (Screen.Height - Form1.Height) / 2

    DrawScore 0
    
    MakePuyoColor
    
End Sub
'
' �R�b�V�[���������߂ɐ����āA���邵������B
'
' �������@�Ȃ�
' �Ԃ��l�@�������߂ɂ��邵�������R�b�V�[�̐�
'         KesuZ �͕ύX����
'
Public Function KesuTameni()

    ReDim KariZ(0 To StageXSize + 1, -1 To StageYSize + 1)
    Dim x, y
    Dim xx, yy
    Dim xxx, yyy
    Dim CheckBangou
    Dim TunagetaKazu
    Dim STunagetaKazu
    Dim ReturnValue
    
        '*****************  ��������
    ReturnValue = 0
    CheckBangou = 0
    
    For x = 1 To StageXSize
        For y = 0 To StageYSize
            KesuZ(x, y) = 0
        Next
    Next
    
    
    If m32(0).Checked = True Then           '   �Ȃ��郂�[�h�Ȃ�
    
        For x = 1 To StageXSize
            For y = 0 To StageYSize
                If Z(x, y) <> Kara And KariZ(x, y) = 0 And Z(x, y) < 1000 Then
                    CheckBangou = CheckBangou + 1
                    KariZ(x, y) = CheckBangou
                    TunagetaKazu = 1
                    Do
                        STunagetaKazu = TunagetaKazu
                        For xxx = 1 To StageXSize
                            For yyy = 0 To StageYSize
                                If KariZ(xxx, yyy) = CheckBangou Then
                                    For xx = -1 To 1
                                        For yy = -1 To 1
                                        
                                            If NanameTrueOrFalse(m30(0).Checked, xx, yy) Then '
                                                If Z(xxx, yyy) = Z(xxx + xx, yyy + yy) And KariZ(xxx + xx, yyy + yy) <> CheckBangou Then
                                                    TunagetaKazu = TunagetaKazu + 1
                                                    KariZ(xxx + xx, yyy + yy) = CheckBangou
                                                End If
                                            End If
                                        
                                        Next
                                    Next
                                End If
                            Next
                        Next
                    Loop Until TunagetaKazu = STunagetaKazu
                End If
                
                If TunagetaKazu >= Narabe Then     ' Val �������ق����ǂ��̂��ȁB
                    For xx = 1 To StageXSize
                        For yy = 0 To StageYSize
                            If KariZ(xx, yy) = CheckBangou Then
                                ReturnValue = ReturnValue + 1
                                KariZ(xx, yy) = -1
                                KesuZ(xx, yy) = True
                            End If
                        Next
                    Next
                End If
            Next
        Next

    Else ' �Ȃ��Ȃ����[�h�Ȃ�

        
        For x = 1 To StageXSize
            For y = 0 To StageYSize
                KariZ(x, y) = 0
            Next
        Next
        
        For x = 1 To StageXSize
            For y = 0 To StageYSize
                If Z(x, y) <> Kara And Z(x, y) < 1000 Then
                    
                    For xx = -1 To 1
                        For yy = -1 To 1
                            
                            If NanameTrueOrFalse(m30(0).Checked, xx, yy) Then
                                TunagetaKazu = 1
                                CheckBangou = CheckBangou + 1
                                    Do While Z(x + xx * TunagetaKazu, y + yy * TunagetaKazu) = Z(x, y)
                                    KariZ(x + xx * TunagetaKazu, y + yy * TunagetaKazu) = CheckBangou
                                    TunagetaKazu = TunagetaKazu + 1
                                Loop
                            
                If TunagetaKazu >= Narabe Then     ' Val �������ق����ǂ��̂��ȁB
                    For xxx = 1 To StageXSize
                        For yyy = 0 To StageYSize
                            If KariZ(xxx, yyy) = CheckBangou Then
                                KariZ(xxx, yyy) = -1
                                KesuZ(xxx, yyy) = True
                            End If
                        Next
                    Next
                End If
                        
                        
                            End If
                        
                        Next
                    Next
                End If
                
            Next
        Next

                    For xxx = 1 To StageXSize
                        For yyy = 0 To StageYSize
                            If KesuZ(xxx, yyy) = True Then
                                ReturnValue = ReturnValue + 1
                            End If
                        Next
                    Next

    End If
    
    KesuTameni = ReturnValue

End Function

'
' �N���N���R�b�V�[�̏I��
'
Private Sub m02_Click()

    End
    
End Sub

'
' �w���v
'
Private Sub m07_Click()

    MsgBox "�܂��o���Ă��Ȃ��̂ł��B"
    
End Sub


'
' �ް�ޮݏ��
'
Private Sub m09_Click()

    About.Timer1.Enabled = False
    About.Show 1
    
End Sub

'
'���Ƃ����
'
Public Sub Otosu()

    Dim i, y, x
        
    For i = 0 To StageYSize
        For y = StageYSize To 0 Step -1
            For x = 1 To StageXSize
                If Z(x, y) = Kara Then
                    Z(x, y) = Z(x, y - 1)
'                    PutPuyo x, y, Z(x, y)
                    Z(x, y - 1) = Kara
'                    PutPuyo x, y - 1, Kara
                End If
            Next
        Next
    Next

End Sub

'
' �t�B�[���h�̏�Ԃ����̂܂ܕx���I�ɕ\�����܂��B
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub FugouHyouji()

    Dim y, x

    For y = 0 To StageYSize
        For x = 1 To StageXSize
            PutPuyo x, y, Z(x, y)
        Next
    Next

End Sub

'
' �R�b�V�[�̐F���Q�T�U�F�ݒ肵�܂��B
'
Public Sub MakePuyoColor()

    Dim i, j, Flag
    
    PuyoColor(0) = RGB(0, 255, 0)
    PuyoColor(1) = RGB(255, 255, 0)
    PuyoColor(2) = RGB(255, 0, 0)
    PuyoColor(3) = RGB(0, 160, 255)
    PuyoColor(4) = RGB(255, 128, 255)
    PuyoColor(5) = RGB(0, 255, 255)
    PuyoColor(6) = RGB(256, 128, 64)
    PuyoColor(7) = RGB(0, 160, 0)
    PuyoColor(8) = RGB(192, 192, 0)
    PuyoColor(9) = RGB(192, 64, 192)
    
    For i = 10 To 255
        Do
            DoEvents ' �i�v���[�v�h�~�p
            Flag = True
                
            Do
                PuyoColor(i) = RGB(Int(Rnd * 9) * 32, Int(Rnd * 9) * 32, Int(Rnd * 9) * 32)
            Loop Until ColorSa(RGB(0, 0, 0), PuyoColor(i)) > 200
            
            For j = 0 To i - 1
                'If ColorSa(PuyoColor(j), PuyoColor(i)) <= 10 Then
                If PuyoColor(j) = PuyoColor(i) Then
                    Flag = False
                    Exit For
                End If
            Next
        
        Loop Until Flag = True
    Next

End Sub

'
' �����ė��Ƃ��ĕ\������
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub KeshiteOtoshiteHyouji()

    Dim Rensame
    Dim KeshitaPuyoKazu
    Dim KeshitaPuyoKazuTotal
    Dim ScoreUp
    Dim ScoreUpTotal

    Rensame = 0
    KeshitaPuyoKazuTotal = 0
    DrawRensa 0, 0, 0

    Ochiteru = False

    Do
        Otosu   ' �P��ڂ̃��[�v�͒u�����̂𗎂Ƃ��A�̂��C�������B
        FugouHyouji
        Form1.Refresh
        Rensame = Rensame + 1
        KeshitaPuyoKazu = KesuTameni ' �������߂ɂ�������B
        If KeshitaPuyoKazu > 0 Then
            If Rensame >= 2 Then
                Delay TenmetsuDelayTime
                Form1.Refresh
            End If
            If m28(0).Checked Then
                SoundKurukuru
            End If
            Kesutenmetu
        End If
        Kesu    ' �����B
        If KeshitaPuyoKazu > 0 Then
            KeshitaPuyoKazuTotal = KeshitaPuyoKazuTotal + KeshitaPuyoKazu
            ScoreUp = ((2 ^ (Rensame - 1)) * KeshitaPuyoKazu) * 100  ' * 100 ��Version 54.07����
            DrawRensa Rensame, KeshitaPuyoKazu, ScoreUp
            ScoreUpTotal = ScoreUpTotal + ScoreUp
        End If
    Loop Until KeshitaPuyoKazu = 0  '�������R�b�V�[���O�ɂȂ�܂Ń��[�v�𑱂���B

    FugouHyouji

    If Rensame > 1 Then
        DrawRensaTotal Rensame, KeshitaPuyoKazuTotal, ScoreUpTotal
    End If

    Score = Score + ScoreUpTotal
    DrawScore Score

End Sub
'
' �Q�̕ϐ����r�v�`�o���܂��B
'
' �������@�Q�̕ϐ�
' �Ԃ��l�@�Ȃ�
'
Public Sub Swap(a, b)

    Dim c
    
    c = a
    a = b
    b = c

End Sub

'
' ���̃R�b�V�[��\�����܂��B
'
Public Sub PutNextPuyo()

    Dim i

    For i = 0 To DrawNextColorMax - 1
        Line ((StageXSize + 2.5 + i * 1.5) * 30, 2 * 30)-Step(30, 60), Form1.BackColor, BF
    Next

        For i = 0 To NextColorPrefetch - 1
            PutPuyo (StageXSize + 2.5 + i * 1.5), 3, Kara
            PutPuyo (StageXSize + 2.5 + i * 1.5), 2, Kara
        Next

    If m11.Enabled = False Then   ' �Q�[�����Ȃ��
        For i = 0 To NextColorPrefetch - 1
            PutPuyo StageXSize + 2.5 + i * 1.5, 3, NextColor(0, i)
            PutPuyo StageXSize + 2.5 + i * 1.5, 2, NextColor(1, i)
        Next
    End If
    
End Sub

'
' ������@�\��
'
Private Sub m10_Click()

    MsgBox "�E�ֈړ��F���@�܂��́@�U(�ݷ�)" & Chr(13) & "���ֈړ��F���@�܂��́@�S(�ݷ�)" & Chr(13) & "���Ƃ��F���@�܂��́@�Q(�ݷ�)" & Chr(13) & "����]�F�y�@�܂��́@�b" & Chr(13) & "�E��]�F�w�@�܂��́@�T(�ݷ�)", , "������@"

End Sub

'
' �t�B�[���h�����������܂��B
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub MakeField()

    Dim x, y

    For x = 0 To StageXSize + 1
        For y = -1 To StageYSize + 1
            If x = 0 Or x = StageXSize + 1 Or y = StageYSize + 1 Then
                Z(x, y) = Kabe
                PutPuyo x, y, Kabe
            Else
                Z(x, y) = Kara
                PutPuyo x, y, Kara
            End If
        Next
    Next

End Sub

'
' �X�^�[�g
'
Private Sub m11_Click()

    RuleHideShowChange False
    
    m11.Enabled = False   ' �X�^�[�g
    m16.Enabled = True    ' �ꎞ��~

    MakeField

    Score = 0
    DrawScore 0
    NokoriTime = 180
    
    m19_Click 0            ' ���Ԑ����������I�ɂȂ��ɂ���B Version 54.02����
    
    DrawScore 0
    If m19(1).Checked Then    '���Ԑ�������Ȃ�
        DrawTime NokoriTime
    End If
    
    ' �\���̃��Z�b�g
    Form1.Refresh
    
    Image1.Visible = False

    Timer1.Enabled = True  ' �Q�[���X�^�[�g

End Sub

'
' �R�b�V�[�̍~�鑬��
'
Private Sub m13_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "KuruDownSpeed", Index
    
    For i = m13.LBound To m13.UBound
        m13(i).Checked = False
    Next
    
    m13(Index).Checked = True

    Select Case Index
    Case 0
        Timer1.Interval = 25
        Koma = 1 / 30
    Case 1
        Timer1.Interval = 25
        Koma = 2 / 30
    Case 2
        Timer1.Interval = 25
        Koma = 4 / 30
    Case 3
        Timer1.Interval = 25
        Koma = 8 / 30
    Case 4
        Timer1.Interval = 25
        Koma = 15 / 30
    Case Else
        MsgBox "�o�O Timer��Interval�܂��"
        Stop
    End Select

End Sub

'
' �ꎞ��~
'
Private Sub m16_Click()

    If Ochiteru Then
        If m16.Checked = False Then
            m16.Checked = True
            Timer1.Enabled = False
        Else
            m16.Checked = False
            Timer1.Enabled = True
        End If
    End If
        
End Sub

'
' �Q�[�����[�h�ύX
'
' ���Ԑ����̂���E�Ȃ�
'
Private Sub m19_Click(Index As Integer)

    
    Index = 0    '���Ԑ������Ȃ��ɋ����I�ɐݒ�B  Version 54.02����
    
    Dim i
    
    SaveSetting "Kurukuru", "Config", "GameMode", Index
    
    For i = m19.LBound To m19.UBound
        m19(i).Checked = False
    Next
    
    m19(Index).Checked = True
    
    Select Case Index
    Case 0
        DrawTime -1 '����
    Case 1
        DrawTime 180
        Form1.Refresh
    End Select

End Sub


'
' ���L�[�̏����̕ύX
'
Private Sub m21_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "DownKey", Index
    
    For i = m21.LBound To m21.UBound
        m21(i).Checked = False
    Next
    m21(Index).Checked = True
    
        
End Sub

'
' ���уR�b�V�[�̏o���p�x��ς���
'
Private Sub m24_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "ChibiKuru", Index
    
    For i = m24.LBound To m24.UBound
        m24(i).Checked = False
    Next
    
    m24(Index).Checked = True
    
    Select Case Index
        Case 0
            ChibiKuruRitu = 1
        Case 1
            ChibiKuruRitu = 1.3
        Case 2
            ChibiKuruRitu = 2
    End Select

End Sub

'
' �F�̐���ς���
'
Private Sub m25_Click(Index As Integer)
    
    Dim i

    SaveSetting "Kurukuru", "Config", "Color", Index
        
    For i = m25.LBound To m25.UBound
        m25(i).Checked = False
    Next
    
    m25(Index).Checked = True
        
    Select Case Index
    Case 0 To 4
        ColorMax = Index + 2
    Case 5
        ColorMax = 10
    Case 6
        ColorMax = 256
    End Select
    
End Sub

'
' �Ȃ��鐔��ς���
'
Private Sub m26_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "Narabe", Index

    For i = m26.LBound To m26.UBound
        m26(i).Checked = False
    Next
    
    m26(Index).Checked = True
        
    Select Case Index
    Case 0 To 4
        Narabe = Index + 2
    Case 5
        Narabe = 10
    End Select
    
End Sub

'
' ���ʉ���ON��OFF�̐؂�ւ�
'
Private Sub m28_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "SoundEffect", Index
    
    For i = m28.LBound To m28.UBound
        m28(i).Checked = False
    Next
    
    m28(Index).Checked = True
    
End Sub

'
' �ȂȂ߂łȂ��邩
'
Private Sub m30_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "Naname", Index
    
    For i = m30.LBound To m30.UBound
        m30(i).Checked = False
    Next
    
    m30(Index).Checked = True
    
End Sub

'
' �܂����ĂȂ��邩
'
Private Sub m32_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "Magaru", Index
    
    For i = m32.LBound To m32.UBound
        m32(i).Checked = False
    Next
    
    m32(Index).Checked = True

End Sub


Private Sub m34_Click()

    SetRuleDefault
    
End Sub

'
' �t�B�[���h�T�C�Y��I�ԃ��j���[
'
Private Sub m36_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "FieldSize", Index
    
    For i = m36.LBound To m36.UBound
        m36(i).Checked = False
    Next
    
    m36(Index).Checked = True
    
    
    Select Case Index
        Case 0
            StageXSize = 4
            StageYSize = 7
        Case 1
            StageXSize = 6
            StageYSize = 11
        Case 2
            StageXSize = 8
            StageYSize = 15
    End Select
    
    ChangeFieldSize
    Cls               ' ����
    DrawScore 0
    PutNextPuyo
    MakeField
    Form1.Refresh     ' �\���̃��Z�b�g

End Sub

'
' NextColorPrefetch
'
Private Sub m38_Click(Index As Integer)

    Dim i
    
    SaveSetting "Kurukuru", "Config", "NextColorPrefetch", Index
    
    For i = m38.LBound To m38.UBound
        m38(i).Checked = False
    Next
    
    m38(Index).Checked = True
    
    NextColorPrefetch = Index
    
    PutNextPuyo   '�ύX�𑦍��ɔ��f����
    
End Sub

'
' ���C���̃^�C�}�[
'
' ���ׂĂ͂����ŊǗ����܂��B
'
Private Sub Timer1_Timer()

    Dim i
    Static SavedTimer
    
    'Debug.Print Timer

    If m19(1).Checked = True Then   ' �^�C���A�^�b�N���[�h�Ȃ玞�Ԃ����炵�܂��B
        If Int(Timer) <> SavedTimer Then
            SavedTimer = Int(Timer)
            NokoriTime = NokoriTime - 1
            DrawTime Sei(NokoriTime)
            If NokoriTime < 0 Then
                Z(FuruIchiX(StageXSize), 0) = Kabe
                Z(FuruIchiX(StageXSize), 1) = Kabe
            End If
        End If
    End If


    ' �΂��񂫂�[�`�F�b�N
    For i = 1 To StageXSize
        If Z(i, 1) <> Kara Then
            Z(FuruIchiX(StageXSize), 0) = Kabe
        End If
    Next
    
    If Z(FuruIchiX(StageXSize), 0) <> Kara Then  '�΂��񂫂�[
        
        '�^�C���I�[�o�[���Q�[���I�[�o�[��
        If NokoriTime < 0 Then
            Image1.Picture = Image2.Picture
        Else
            Image1.Picture = Image3.Picture
        End If
        
        ' Image1�̃Z���^�����O
        ToCenterImage1
        
        Timer1.Enabled = False
        Image1.Visible = True
        Timer2.Enabled = True
        
        GameJunbi
    Else
        If Not Ochiteru Then ' �����ĂȂ�
            Ochiteru = True
            px(0) = FuruIchiX(StageXSize)
            py(0) = 0
            px(1) = FuruIchiX(StageXSize)
            py(1) = -1
            NowColor(0) = NextColor(0, 0)
            NowColor(1) = NextColor(1, 0)
            MakeNextColor
            PuyoMuki = 0
            PutNextPuyo
        End If

        If Z(px(0), int2(py(0) + Koma)) = Kara And Z(px(1), int2(py(1) + Koma)) = Kara Then ' ���ɉ����������
            ImanoPuyoKesu
            For i = 0 To 1
                py(i) = py(i) + Koma
            Next
            ImanoPuyoKaku
        Else               ' ���ɉ��������
            Tsuita
        End If
    End If
    
End Sub
'
' �m�������R�b�V�[�̐F�����߂܂��B
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
' NextColor( ) �̓O���[�o���ϐ�
'
Public Sub MakeNextColor()

    Dim i, j
    
    For j = 0 To DrawNextColorMax - 1
        For i = 0 To 1
            NextColor(i, j) = NextColor(i, j + 1)
        Next
    Next
    
    For i = 0 To 1
        NextColor(i, DrawNextColorMax) = Int(Rnd * ColorMax) + (Int(Rnd * ChibiKuruRitu) * 1000)
    Next

End Sub

'
' �����Ă���Ƃ��̍��̃R�b�V�[������
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub ImanoPuyoKesu()

    Dim i
    
    For i = 0 To 1
        PutPuyo px(i), py(i), Kara
    Next

End Sub

'
' �����Ă���Ƃ��̍��̃R�b�V�[��`��
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub ImanoPuyoKaku()

    Dim i
    
    For i = 0 To 1
        PutPuyo px(i), py(i), NowColor(i)
    Next

End Sub

'
' �������B
'
' ����������������B
'
' �������@�Ȃ�
' �Ԃ��l�@�Ȃ�
'
Public Sub Tsuita()

    Dim i

    For i = 0 To 1
        Z(px(i), py(i)) = NowColor(i)
    Next
    
    KeshiteOtoshiteHyouji
    
    Ochiteru = False
    QuickTurnFlag = False

End Sub

'
' �Q�[�����n�߂鏀�������܂�
'
Public Sub GameJunbi()

'    Form1.Left = (Screen.Width - Form1.Width) / 2
'    Form1.Top = (Screen.Height - Form1.Height) / 2
    
    ChangeFieldSize
    
    Ochiteru = False
    
    RuleHideShowChange True
    
    m11.Enabled = True
    m16.Enabled = False
    m18.Enabled = True

    MakeNextColorPrefetch

End Sub

'
' ���̐��ɂ��܂�
'
Public Function Sei(x)

    If x < 0 Then
        Sei = 0
    Else
        Sei = x
    End If

End Function

'
' �Q�F�̂���̍�
'
Public Function ColorSa(c1, c2)

    ColorSa = Abs((Int(c1 / &H1) And &HFF) - (Int(c2 / &H1) And &HFF)) _
            + Abs((Int(c1 / &H100) And &HFF) - (Int(c2 / &H100) And &HFF)) _
            + Abs((Int(c1 / &H10000) And &HFF) - (Int(c2 / &H10000) And &HFF))
    
End Function



'
' ���������ɑ҂��܂��B
'
Public Sub Delay(Byou)

    Dim StartTimer
    Dim NowTimer

    StartTimer = Timer
    Do
        ' Debug.Print NowTimer - StartTimer
        NowTimer = Timer
        If NowTimer < StartTimer Then ' ���ɂ����ς������
            NowTimer = NowTimer + 86400
        End If
    Loop Until NowTimer - StartTimer > Byou

End Sub
'
' KesuZ(x,y) �̓��e�ɂ��������ď����܂��B
'
' ����̔�������܂��B
'
Public Sub Kesu()

    Dim x, y, xx, yy

    For x = 1 To StageXSize
        For y = 0 To StageYSize
            If KesuZ(x, y) = True Then
                PutPuyo x, y, Kara
                Z(x, y) = Kara
                
                ' ����̔���
                For xx = -1 To 1
                    For yy = -1 To 1
                        If Abs(xx + yy) = 1 Then ' �㉺���E�S����
                            If Z(x + xx, y + yy) >= 1000 Then
                                Z(x + xx, y + yy) = Z(x + xx, y + yy) - 1000
                                PutPuyo x + xx, y + yy, Z(x + xx, y + yy)
                            End If
                        End If
                    Next
                Next
            
            End If
        Next
    Next

End Sub

'
' �J��グ Int
'
Public Function int2(Num)

    int2 = Int(Num + 0.999)

End Function

'
' ����̂Ƃ��̔w�i�������܂�
'
Public Sub DrawKara(x, y)

    Dim Dummy
    
    Dummy = BitBlt(Form1.hDC, x * 30, y * 30, 30, 30, Picture1.hDC, 0, Int((y - Int(y)) * 30), SRCCOPY)

End Sub

'
' �ǂ������܂�
'
Public Sub DrawKabe(x, y)

    Dim Dummy
    
    Dummy = BitBlt(Form1.hDC, x * 30, y * 30, 30, 30, Picture2.hDC, 0, 0, SRCCOPY)

End Sub

'
' �����R�b�V�[��_�ł�����
'
Public Sub Kesutenmetu()

    Dim x, y

    For x = 1 To StageXSize
        For y = 0 To StageYSize
            If KesuZ(x, y) = True Then
                PutPuyo x, y, Kara
            End If
        Next
    Next
    Form1.Refresh
    
    Delay TenmetsuDelayTime
    
    For x = 1 To StageXSize
        For y = 0 To StageYSize
            If KesuZ(x, y) = True Then
                PutPuyo x, y, Z(x, y)
            End If
        Next
    Next
    Form1.Refresh
    
    Delay TenmetsuDelayTime


End Sub

'
' QuickTurn���܂��B
'
Public Sub QuickTurn()

    Select Case PuyoMuki
    Case 2
        PuyoMuki = 0
        py(1) = py(0) - 1
        py(0) = py(0) + 1
        py(1) = py(1) + 1
    Case 0
        PuyoMuki = 2
        py(1) = py(0) + 1
        py(0) = py(0) - 1
        py(1) = py(1) - 1
    End Select

End Sub

'
' DrawNumber����Ăяo����܂�
'
Public Sub DrawNumber2(x, y, n)

    Dim Dummy
        
    Const Xsize = 10
    
    Dummy = BitBlt(Form1.hDC, x, y, Xsize, 20, Picture3.hDC, n * Xsize, 0, SRCCOPY)

End Sub

'
' �X�R�A��\��
'
Public Sub DrawScore(n)

    Dim Dummy
    
    Const DrawTotalScoreY = 30    ' �ύX��
    Const ScoreNumberHeight = 20  ' �����炭�ύX��
    
    '�g�[�^���X�R�A�̂Ƃ��낾������ �Ȃ�ċ����Ȃ񂾂낤
    Line (ScoreHideX, DrawTotalScoreY)-(Form1.Width, DrawTotalScoreY + ScoreNumberHeight), Form1.BackColor, BF
    
    If n < 0 Then
        MsgBox "�X�R�A���}�C�i�X�ł��B"
        Stop
    Else
        ' ���l
        DrawNumber ScoreDrawX + FieldSizeScoreXHosei + 0 + 30, DrawTotalScoreY, n, 0
        '�_
        Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei + 0 + 30 + 15, DrawTotalScoreY, 60, 20, Picture5.hDC, 0, 0, SRCCOPY)
    End If

End Sub
'
' �c�莞�Ԃ�\��
'
Public Sub DrawTime(n)

    Dim Dummy

    If n < 0 Then
    Else
        DrawNumber ScoreDrawX + FieldSizeScoreXHosei, 80, n, 0
         '�b
        Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei + 15, 80, 20, 20, Picture4.hDC, 90, 0, SRCCOPY)
    End If

End Sub

'
' �A���󋵂̕\��
'
' ���� ���A���ڂ��A���������A�X�R�A
'
Public Sub DrawRensa(r, k, s)

    Dim Dummy
    
    Dim y
    
    If r = 0 Then
        '�g�[�^���X�R�A�̂Ƃ��낾������ �Ȃ�ċ����Ȃ񂾂낤
        Line (ScoreHideX, ScoreDrawY + 20)-(Form1.Width, Form1.Height), Form1.BackColor, BF
    Else
        y = ScoreDrawY + r * 20
        '�A��������
        DrawNumber ScoreDrawX + FieldSizeScoreXHosei - 120, y, r, 0
        Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 120 + 15, y, 20, 20, Picture4.hDC, 0, 0, SRCCOPY)
        DrawNumber ScoreDrawX + FieldSizeScoreXHosei - 70, y, k, 0
        Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 70 + 15, y, 20, 20, Picture4.hDC, 30, 0, SRCCOPY)
        DrawNumber ScoreDrawX + FieldSizeScoreXHosei - 0 + 30, y, s, 0
        Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 0 + 30 + 15, y, 60, 20, Picture5.hDC, 0, 0, SRCCOPY)
    End If
    
End Sub

'
' �A�����ʂ̕\��
'
' ���� ���A���ڂ��A���������A�X�R�A
'
Public Sub DrawRensaTotal(r, k, s)

    Dim Dummy

    Dim y
        
    y = ScoreDrawY + r * 20 + 10

    Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 120 + 15, y, 20, 20, Picture4.hDC, 117, 0, SRCCOPY)
    DrawNumber ScoreDrawX + FieldSizeScoreXHosei - 70, y, k, 0
    Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 70 + 15, y, 20, 20, Picture4.hDC, 30, 0, SRCCOPY)
    DrawNumber ScoreDrawX + FieldSizeScoreXHosei - 0 + 30, y, s, 0
    Dummy = BitBlt(Form1.hDC, ScoreDrawX + FieldSizeScoreXHosei - 0 + 30 + 15, y, 60, 20, Picture5.hDC, 0, 0, SRCCOPY)

End Sub

'
' �����������܂�
'
' Ketasuu��0�̂Ƃ��́A�O�[�������܂���B
'
Public Sub DrawNumber(x, y, n, Ketasuu)

    Dim i
    Dim KetaN
    Dim MaxKeta
    
    If Ketasuu = 0 Then
        MaxKeta = Int(Log10(n) + 0.001) '�덷��?
    Else
        MaxKeta = Ketasuu - 1
    End If
    
    KetaN = n
    For i = 0 To MaxKeta
        DrawNumber2 x - i * 10, y, (KetaN Mod 10)
        KetaN = Int(KetaN / 10)
    Next

End Sub

'
' 10���ɂ���log
'
Static Function Log10(n)

    If n = 0 Then
        Log10 = 0
    Else
        Log10 = Log(n) / Log(10)
    End If

End Function

'
' ���Ԃ�������Q�[���I�[�o�[�̕\��������
'
Private Sub Timer2_Timer()

    Image1.Visible = False
    Timer2.Enabled = False

End Sub



'
' ���[���̕ύX�ł��邩�ł��Ȃ�����ς��܂�
'
Public Sub RuleHideShowChange(TrueOrFalse As Boolean)

    Dim i
    
    For i = m19.LBound To m19.UBound
        m19(i).Enabled = TrueOrFalse
    Next
    
    For i = m25.LBound To m25.UBound
        m25(i).Enabled = TrueOrFalse
    Next
    
    For i = m26.LBound To m26.UBound
        m26(i).Enabled = TrueOrFalse
    Next
    
    For i = m30.LBound To m30.UBound
        m30(i).Enabled = TrueOrFalse
    Next
    
    For i = m32.LBound To m32.UBound
        m32(i).Enabled = TrueOrFalse
    Next
    
    For i = m36.LBound To m36.UBound
        m36(i).Enabled = TrueOrFalse
    Next
    
    For i = m24.LBound To m24.UBound   '���уR�b�V�[�̏o���p�x
        m24(i).Enabled = TrueOrFalse
    Next
    
    m34.Enabled = TrueOrFalse
    
End Sub

'
' �ȂȂ߃t���O��x,y����A�l��Ԃ��܂�
'
Public Function NanameTrueOrFalse(Flag, x, y)

    Dim ReturnValue
    
    ReturnValue = True
    
    If Flag Then
        If x = 0 And y = 0 Then
            ReturnValue = False
        End If
    Else
        If Not (Abs(x + y) = 1) Then
            ReturnValue = False
        End If
    End If
    
    NanameTrueOrFalse = ReturnValue
    
End Function

'
' ���[����������Ԃɂ��܂�
'
Public Sub SetRuleDefault()

    m25_Click 2 ' �F��4�F
    m26_Click 1 '�R�Ȃ��ŏ���
    'm13_Click Val(GetSetting("Kurukuru", "Config", "KuruDownSpeed", 0)) '�~�鑬��
    'm21_Click Val(GetSetting("Kurukuru", "Config", "DownKey", 0))
    m19_Click 0
    'm28_Click Val(GetSetting("Kurukuru", "Config", "SoundEffect", 0))
    m30_Click 1
    m32_Click 0
    m36_Click 0
    m24_Click 1 '���уR�b�V�[�̏o��p�x�͂��܂�

End Sub

'
' �t�B�[���h�T�C�Y����~��ʒu���v�Z���܂��B
'
Public Function FuruIchiX(StageXSize)

    FuruIchiX = Int(StageXSize / 2)

End Function

'
' �t�B�[���h�T�C�Y��ς��܂��B
'
Public Sub ChangeFieldSize()

    Form1.Width = Screen.TwipsPerPixelX * 30 * (StageXSize + 11)
    Form1.Height = Screen.TwipsPerPixelY * 30 * (StageYSize + 3.5)
    
    Line1.X2 = StageXSize * 30 + Line1.X1
    
    FieldSizeScoreXHosei = 30 * (StageXSize - 10)
    ScoreHideX = 30 * (StageXSize - 10) + ScoreDrawX - 140

End Sub

'
' Image1(GAME OVER��\������Image)���Z���^�����O���܂��B
'
Public Sub ToCenterImage1()

    Const Mibae = 4 '���h�����悭���邽�߂̒萔�B���l�̍����͂Ȃ�ƂȂ��B
    
    Image1.Left = Int(((Form1.Width / Screen.TwipsPerPixelX) - Image1.Width) / 2)
    Image1.Top = Int(((Form1.Height / Screen.TwipsPerPixelY) - Image1.Height) / 2)
    Image1.Top = Image1.Top - Int(Image1.Height / Mibae) '�Ȃ�ƂȂ��������ق������h��������B

End Sub

'
' Next�R�b�V�[�̐F���v���t�F�b�`���������߂܂��B
'
Public Sub MakeNextColorPrefetch()

    Dim i

    For i = 0 To DrawNextColorMax
        MakeNextColor
    Next

End Sub
