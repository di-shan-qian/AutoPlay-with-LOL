VERSION 5.00
Begin VB.Form Ai 
   Caption         =   "自动打人机"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   2130
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   -52
      Top             =   637
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   -52
      Top             =   637
   End
   Begin VB.CommandButton PlayCmd 
      Caption         =   "开始"
      Height          =   1095
      Left            =   196
      TabIndex        =   2
      Top             =   300
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   1388
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   -52
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Ai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'开始匹配
'判断
'    匹配状态Step_01
'        成功
'            点击进入房间
'                成功
'                    选择英雄 Step_02
'                        成功
'                            判断游戏状态
'                                开始
'                                    开始游戏，判断游戏结束
'                                        结束
'                                            开始匹配
'                                        未结束
'                                            null
'                        失败
'                            开始匹配Step_01
'                失败
'                    开始匹配Step_01
'        失败
'
Option Explicit

'鼠标点击API===
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'=============
'找图API===
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long '释放DC
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte   '透明通道
End Type

Private Type BITMAPINFOHEADER
    biSize As Long          '位图大小
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer   '信息头长度
    biCompression As Long   '压缩方式
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
'=============
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'=====================
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim gameWnd1 As Long
Dim SW As Long
Dim SH As Long
Dim x As Long
Dim y As Long
Dim R As RECT
Dim Py As Integer
Dim F1 As Boolean
Dim F2 As Boolean
Dim F3 As Boolean
Dim F4 As Boolean
Dim cal As Integer
Dim OK As Boolean
Dim F7 As Boolean
'模拟鼠标左键单击
Private Sub M_L()
    mouse_event &H2, 0, 0, 0, 0
    Sleep (10)
    mouse_event &H4, 0, 0, 0, 0
End Sub
'模拟鼠标右键单击
Private Sub M_R()
    mouse_event &H8, 0, 0, 0, 0
    Sleep (10)
    mouse_event &H10, 0, 0, 0, 0
End Sub
'找图=======
Private Sub GetImageMemory(ByVal Pic As PictureBox, W As Long, H As Long, Memory() As Byte, bi As BITMAPINFO)
    
    With bi.bmiHeader
        .biCompression = 0&
        .biSize = Len(bi.bmiHeader)
        .biWidth = W
        .biHeight = -H
        .biBitCount = 32
        .biPlanes = 1
    End With
    ReDim Memory(3, 0 To W - 1, 0 To H - 1)
    GetDIBits Pic.hdc, Pic.Picture.Handle, 0&, H, Memory(0, 0, 0), bi, 0&
    ReleaseDC 0, Pic.hdc
End Sub
Private Sub saveMyScreen()
        Dim lngDesktopHwnd As Long
        Dim lngDesktopDC As Long

        Ai.Picture2.AutoRedraw = True
        Ai.Picture2.ScaleMode = vbPixels
        lngDesktopHwnd = GetDesktopWindow
        lngDesktopDC = GetDC(lngDesktopHwnd)

        Ai.Picture2.Width = Screen.Width
        Ai.Picture2.Height = Screen.Height
        Call BitBlt(Ai.Picture2.hdc, 0, 0, Screen.Width, Screen.Height, lngDesktopDC, 0, 0, SRCCOPY)
        Ai.Picture2.Picture = Ai.Picture2.Image
        Call ReleaseDC(lngDesktopHwnd, lngDesktopDC)
End Sub
Private Function FindPic(Left As Long, Top As Long, Right As Long, Bottom As Long, fileurl As String, SimRate As Long, intX As Long, intY As Long) As Boolean
    Dim zPic() As Byte, fPic() As Byte
    Dim zImg As BITMAPINFO, fImg As BITMAPINFO
    Dim Now As Long, Noh As Long
    Dim I As Long, J As Long, I2 As Long, J2 As Long
    Dim W As Long, H As Long
    Set Ai.Picture1.Picture = LoadPicture(fileurl)
    W = Ai.Picture1.ScaleWidth / Screen.TwipsPerPixelX
    H = Ai.Picture1.ScaleHeight / Screen.TwipsPerPixelY
    GetImageMemory Ai.Picture1, W, H, zPic(), zImg
    W = SW 'Right
    H = SH 'Bottom
    saveMyScreen
    GetImageMemory Ai.Picture2, W, H, fPic(), fImg
    Now = Round(UBound(zPic, 2) / 10) + 1
    Noh = Round(UBound(zPic, 3) / 10) + 1
    For J = Top To H - UBound(zPic, 3)
        For I = Left To W - UBound(zPic, 2)
            For J2 = 0 To UBound(zPic, 3) - 1 Step Noh '循环判断小图片
                For I2 = 0 To UBound(zPic, 2) - 1 Step Now

                    If SimRate < Abs(CInt(fPic(2, I + I2, J + J2)) - CInt(zPic(2, I2, J2))) Then GoTo ExitLine: 'R
                    If SimRate < Abs(CInt(fPic(1, I + I2, J + J2)) - CInt(zPic(1, I2, J2))) Then GoTo ExitLine: 'G
                    If SimRate < Abs(CInt(fPic(0, I + I2, J + J2)) - CInt(zPic(0, I2, J2))) Then GoTo ExitLine: 'b
                Next I2
            Next J2
            '
            intX = I
            intY = J
            FindPic = True
            I = W - UBound(zPic, 2)
            J = H - UBound(zPic, 3)
ExitLine:
        Next I
    Next J

End Function



Private Sub Command1_Click()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    SW = Screen.Width \ Screen.TwipsPerPixelX
    SH = Screen.Height \ Screen.TwipsPerPixelY
End Sub

'=============

Private Sub PlayCmd_Click()
    gameWnd1 = FindWindow(vbNullString, "League of Legends")
    GetWindowRect gameWnd1, R
    AppActivate "League of Legends"
    '开始找图“寻找对局”
    F1 = FindPic(R.Left, R.Top, R.Right, R.Bottom, "01.bmp", "9", x, y)
    If F1 Then
        SetCursorPos x + 40, y + 10
        Call M_L

    End If
    
    Call Step_01
    
End Sub


Private Sub Step_01()

        '判断是否匹配到队列
        
        '点击接受
        Do
        '每隔三秒检测是否匹配到对局，匹配到，退出循环
            Sleep 2000
            F2 = FindPic(R.Left, R.Top, R.Right, R.Bottom, "02.bmp", "9", x, y)
            
            If F2 Then
            
                SetCursorPos x + 40, y + 10
                Call M_L
                Timer1.Enabled = True
                Exit Do

            End If
            
        Loop
        
        Call Step_02 '选择英雄
   
End Sub

Private Sub Timer1_Timer()

    F3 = FindPic(R.Left, R.Top, R.Right, R.Bottom, "02.bmp", "9", x, y)
    
    If F3 Then
                SetCursorPos x + 40, y + 10
                Call M_L
                Call Step_02 '选择英雄
                            
    End If
    
    F4 = FindPic(R.Left, R.Top, R.Right, R.Bottom, "05.bmp", "9", x, y)
    
    If F4 Then
                SetCursorPos x + 40, y + 10
                Call M_L

                            
    End If
    
    '关闭条件
    If FindWindow(vbNullString, "League of Legends (TM) Client") <> 0 Then
        AppActivate "League of Legends (TM) Client"
        
        Py = Shell(App.Path & "\play.exe", 6)

        Timer1.Enabled = False
        
        Timer2.Enabled = True
    End If
End Sub

Private Sub Step_02()
    Sleep (7000)
    SetCursorPos R.Left + 520, R.Top + 108
    Sleep (2000)
    Call M_L
    cal = 0
    Sleep (2000)
    
    Do
        Sleep 2000
        SetCursorPos R.Left + 380 + cal * 100, R.Top + 168
        Call M_L
        Sleep (3000)
        If cal >= 6 Then
            Exit Do
        End If
        cal = cal + 1
    Loop


End Sub



Private Sub Timer2_Timer()
        OK = FindPic(R.Left, R.Top, R.Right, R.Bottom, "07.bmp", "9", x, y)
        'FindPic(R.Left, R.Top, R.Right, R.Bottom, "01.bmp", "9", x, y)
        If OK Then
            SetCursorPos x, y
            Call M_L
        End If
        
        F7 = FindPic(R.Left, R.Top, R.Right, R.Bottom, "06.bmp", "9", x, y)
        If F7 Then
            Timer2.Enabled = False
            SetCursorPos x + 100, y + 10
            Call M_L
            Sleep 3000
            Call M_L
            Call Step_01
        End If

End Sub

'1: 寻找对局1
'2: 接受
'3: 锁定
'4: 进入游戏标志
'5: 继续游戏
'6: 再玩一次1
'7: OK
