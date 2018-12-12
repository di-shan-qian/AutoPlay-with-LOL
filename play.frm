VERSION 5.00
Begin VB.Form play 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   6030
   ClientTop       =   4290
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4710
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1695
      Left            =   2160
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'模拟按键F2，鼠标右键点击，shift按下一次 ，时钟
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起
Const MOUSEEVENTF_RIGHTDOWN = &H8 '模拟鼠标右键按下
Const MOUSEEVENTF_RIGHTUP = &H10 '模拟鼠标右键抬起
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Dim Ti
Dim Tj



Private Sub Form_Load()

Ti = 1
Tj = 1

Sleep (10000)

'判断游戏开始，开启始终
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
ExitProcess 0
End Sub

Private Sub Timer1_Timer()
Call ct
    
If Ti < 30 Then
'f2.2.5min
    keybd_event vbKeyF2, MapVirtualKey(vbKeyF2, 0), 0, 0
    Sleep (77)
    
    SetCursorPos 847, 604
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Sleep (10)
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    If Tj = 1 Then
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), 0, 0
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        Sleep (10)
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), &H2, 0
        Tj = 0
    Else
        Tj = 1
        Call qwer
    End If
    
    keybd_event vbKeyF2, MapVirtualKey(vbKeyF2, 0), &H2, 0
    

ElseIf Ti >= 30 And Ti < 60 Then
'f3
    keybd_event vbKeyF3, MapVirtualKey(vbKeyF3, 0), 0, 0
    Sleep (77)
    
    SetCursorPos 847, 604
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Sleep (10)
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    If Tj = 1 Then
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), 0, 0
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        Sleep (10)
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), &H2, 0
        Tj = 0
    Else
        Tj = 1
        Call qwer
    End If
    keybd_event vbKeyF3, MapVirtualKey(vbKeyF3, 0), &H2, 0



ElseIf Ti >= 60 And Ti < 90 Then

'f4
 keybd_event vbKeyF4, MapVirtualKey(vbKeyF4, 0), 0, 0
    Sleep (77)
    
    SetCursorPos 847, 604
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Sleep (10)
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    If Tj = 1 Then
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), 0, 0
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        Sleep (10)
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), &H2, 0
        Tj = 0
    Else
        Tj = 1
        Call qwer
    End If
    keybd_event vbKeyF4, MapVirtualKey(vbKeyF4, 0), &H2, 0
    
    Call df
ElseIf Ti >= 90 And Ti < 120 Then
'f5
 keybd_event vbKeyF5, MapVirtualKey(vbKeyF5, 0), 0, 0
    Sleep (77)
    
    SetCursorPos 847, 604
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    Sleep (10)
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    
    If Tj = 1 Then
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), 0, 0
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        Sleep (10)
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
        keybd_event vbKeyShift, MapVirtualKey(vbKeyShift, 0), &H2, 0
        Tj = 0
    Else
        Tj = 1
        Call qwer
    End If
    keybd_event vbKeyF5, MapVirtualKey(vbKeyF5, 0), &H2, 0
Else
Ti = 1
End If
Ti = Ti + 1
'End If


If FindWindow(vbNullString, "League of Legends (TM) Client") = 0 Then
        
        Timer1.Enabled = False
        Unload Me

End If
End Sub
Sub ct()

'Ctrl qwer
            keybd_event vbKeyControl, MapVirtualKey(vbKeyControl, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyR, MapVirtualKey(vbKeyR, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyR, MapVirtualKey(vbKeyR, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyQ, MapVirtualKey(vbKeyQ, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyQ, MapVirtualKey(vbKeyQ, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyW, MapVirtualKey(vbKeyW, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyW, MapVirtualKey(vbKeyW, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyE, MapVirtualKey(vbKeyE, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyE, MapVirtualKey(vbKeyE, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyControl, MapVirtualKey(vbKeyControl, 0), &H2, 0


End Sub
Sub qwer()
            keybd_event vbKeyQ, MapVirtualKey(vbKeyQ, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyQ, MapVirtualKey(vbKeyQ, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyW, MapVirtualKey(vbKeyW, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyW, MapVirtualKey(vbKeyW, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyE, MapVirtualKey(vbKeyE, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyE, MapVirtualKey(vbKeyE, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyR, MapVirtualKey(vbKeyR, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyR, MapVirtualKey(vbKeyR, 0), &H2, 0
End Sub
Sub df()
            keybd_event vbKeyD, MapVirtualKey(vbKeyD, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyD, MapVirtualKey(vbKeyD, 0), &H2, 0
            Sleep (77)
            keybd_event vbKeyF, MapVirtualKey(vbKeyF, 0), 0, 0
            Sleep (77)
            keybd_event vbKeyF, MapVirtualKey(vbKeyF, 0), &H2, 0
End Sub

