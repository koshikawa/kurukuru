Attribute VB_Name = "kurukuru"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal xsaki As Long, ByVal ysaki As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SRCCOPY = &HCC0020

'
' 始まり
'
Public Sub Main()

    Form1.Show
    About.Command1.Visible = False
    About.Show 0

End Sub

'
' クルクルってしゃべります。
'
Public Sub SoundKurukuru()

    Dim Dummy
    
    Dummy = sndPlaySound("kurukuru.wav", 3)

End Sub




