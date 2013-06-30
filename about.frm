VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   3045
   ClientLeft      =   3405
   ClientTop       =   2655
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   3045
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   515
      Left            =   3540
      TabIndex        =   0
      Top             =   2380
      Width           =   1015
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "Version X.XX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "About.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Unload About
    
End Sub

'
' 表示されてから
'
Private Sub Form_Activate()

    SoundKurukuru ' 鳴らす
    
End Sub

Private Sub Form_Load()

    Label1.Caption = "Version " & App.Major & "." & Right(("00" & App.Minor), 2)
        
    About.Top = Int((Screen.Height - About.Height) / 2)
    About.Left = Int((Screen.Width - About.Width) / 2)
    
End Sub

'
' Aboutフォームがフォーカスを失ったときに閉じる
'
Private Sub Form_LostFocus()

    Unload About
    
End Sub


Private Sub Timer1_Timer()

    Unload About

End Sub


