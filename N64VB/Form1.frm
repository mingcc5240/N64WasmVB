VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   609
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   748
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   8280
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   120
      ScaleHeight     =   7995
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === N64Wasm.dll 함수 선언 ===
Private Declare Function VB_N64_Init Lib "N64_Wasm.dll" (ByVal hParentWnd As Long, ByVal romPath As String) As Long
Private Declare Function VB_N64_RunFrame Lib "N64_Wasm.dll" () As Long
Private Declare Sub VB_N64_SetInput Lib "N64_Wasm.dll" (ByVal btnMask As Long, ByVal stickX As Long, ByVal stickY As Long)
Private Declare Sub VB_N64_SaveState Lib "N64_Wasm.dll" ()
Private Declare Sub VB_N64_LoadState Lib "N64_Wasm.dll" ()
Private Declare Sub VB_N64_SetDoubleSpeed Lib "N64_Wasm.dll" (ByVal enable As Long)
Private Declare Sub VB_N64_ShowFPS Lib "N64_Wasm.dll" (ByVal show As Long)
Private Declare Sub VB_N64_Shutdown Lib "N64_Wasm.dll" ()
' === 선언부에 추가 ===
Private Declare Sub VB_N64_Resize Lib "N64Wasm.dll" (ByVal width As Long, ByVal height As Long)

' 키보드 실시간 감지용 Win32 API
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

' === Windows API 선언 (파일 열기 대용량 다이얼로그) ===
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private m_bRunning As Boolean
Private m_SelectedRomPath As String

' 1. 외부 ROM 파일을 선택하는 파일 대화상자 함수
Private Function BrowseRomFile() As String
    Dim ofn As OPENFILENAME
    Dim sFile As String
    
    sFile = String(512, 0)
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.lpstrFilter = "N64 ROM Files (*.z64;*.v64;*.n64)|*.z64;*.v64;*.n64|All Files (*.*)|*.*"
    ofn.lpstrFile = sFile & String(512, 0)
    ofn.nMaxFile = Len(ofn.lpstrFile)
    ofn.lpstrTitle = "구동할 N64 롬 파일 선택"
    ofn.flags = &H4& ' OFN_HIDEREADONLY
    
    If GetOpenFileName(ofn) Then
        ' 널 문자 제거 후 경로 추출
        BrowseRomFile = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        BrowseRomFile = ""
    End If
End Function

Private Sub cmdStart_Click()
    Dim res As Long
    
    ' 이미 실행 중인 상태라면 중복 실행 방지
    If m_bRunning Then Exit Sub
    
    ' 파일 선택 다이얼로그 호출
    m_SelectedRomPath = BrowseRomFile()
    
    'MsgBox m_SelectedRomPath
    If m_SelectedRomPath = "" Then
        Exit Sub ' 사용자가 취소를 누른 경우
    End If
    
    DoEvents
    
    ' 선택한 롬 파일 경로를 DLL로 전달하여 초기화
    res = VB_N64_Init(Picture1.hWnd, m_SelectedRomPath)
    If res <= 0 Then
        MsgBox "초기화 실패! 에러 코드: " & res, vbCritical
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    m_bRunning = True
    'VB_N64_Resize Picture1.width, Picture1.height
    
    ' 메인 에뮬레이션 루프 실행
    Do While m_bRunning
        UpdateInputFromVB
        
        If VB_N64_RunFrame() = 0 Then
            m_bRunning = False
            Exit Do
        End If
        
        DoEvents
    Loop
End Sub
' 폼 크기를 마우스로 조절하거나 최대화할 때 자동으로 화면이 커지도록 처리
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Dim picW As Long, picH As Long
    
    ' 아래쪽 버튼 영역(50픽셀)을 제외한 나머지 화면 전체를 Picture1로 채움
    picW = Me.ScaleWidth - 20
    picH = Me.ScaleHeight - 70
    
    If picW > 100 And picH > 100 Then
        Picture1.Move 10, 10, picW, picH
        cmdStart.Move 10, picH + 20
        
        ' 실행 중이면 C++ DLL에도 변경된 크기를 즉시 알림
        If m_bRunning Then
            VB_N64_Resize picW, picH
        End If
    End If
End Sub
   
Private Sub Form_Load()
   Me.ScaleMode = vbPixels
   Picture1.ScaleMode = vbPixels
   Picture1.width = 640 + 10
   Picture1.height = 480 + 10
   Picture1.BackColor = vbBlack
      
   ' Me.ScaleMode = vbPixels
   ' Picture1.width = 640
   ' Picture1.height = 480
   ' Me.width = 1000 * Screen.TwipsPerPixelX
   ' Me.height = 800 * Screen.TwipsPerPixelY
    
   ' Picture1.BackColor = vbBlack
    cmdStart.Caption = "ROM 선택 및 실행"
End Sub

Private Sub cmdStart0_Click()
    Dim res As Long
    Dim sRom As String
    
    ' 실행할 ROM 파일 경로 (App.Path 기준)
    sRom = App.Path & "\waverace64.n64"
    
    ' Picture1.hWnd 안에 OpenGL 화면을 바인딩하여 시작!
    res = VB_N64_Init(Picture1.hWnd, sRom)
    If res <= 0 Then
        MsgBox "초기화 실패! 에러 코드: " & res, vbCritical
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    m_bRunning = True
    
    ' 메인 루프 실행
    Do While m_bRunning
        UpdateInputFromVB
        
        If VB_N64_RunFrame() = 0 Then
            m_bRunning = False
            Exit Do
        End If
        
        DoEvents
    Loop
End Sub

Private Sub UpdateInputFromVB()
    Dim mask As Long
    Dim sx As Long, sy As Long
    
    ' 방향키 -> N64 아날로그 스틱 (-32000 ~ +32000)
    If (GetAsyncKeyState(vbKeyLeft) And &H8000) <> 0 Then sx = -32000
    If (GetAsyncKeyState(vbKeyRight) And &H8000) <> 0 Then sx = 32000
    If (GetAsyncKeyState(vbKeyUp) And &H8000) <> 0 Then sy = -32000    ' main.cpp 기준 Up은 -32000
    If (GetAsyncKeyState(vbKeyDown) And &H8000) <> 0 Then sy = 32000   ' main.cpp 기준 Down은 +32000
    
    ' 버튼 매핑: Enter=Start(16), X=A(32), C=B(64), Z=Z트리거(128), A=L(256), S=R(512)
    If (GetAsyncKeyState(vbKeyReturn) And &H8000) <> 0 Then mask = mask Or 16
    If (GetAsyncKeyState(vbKeyX) And &H8000) <> 0 Then mask = mask Or 32
    If (GetAsyncKeyState(vbKeyC) And &H8000) <> 0 Then mask = mask Or 64
    If (GetAsyncKeyState(vbKeyZ) And &H8000) <> 0 Then mask = mask Or 128
    If (GetAsyncKeyState(vbKeyA) And &H8000) <> 0 Then mask = mask Or 256
    If (GetAsyncKeyState(vbKeyS) And &H8000) <> 0 Then mask = mask Or 512
    
    ' C-버튼 매핑: I=C_Up(1024), K=C_Down(2048), J=C_Left(4096), L=C_Right(8192)
    If (GetAsyncKeyState(vbKeyI) And &H8000) <> 0 Then mask = mask Or 1024
    If (GetAsyncKeyState(vbKeyK) And &H8000) <> 0 Then mask = mask Or 2048
    If (GetAsyncKeyState(vbKeyJ) And &H8000) <> 0 Then mask = mask Or 4096
    If (GetAsyncKeyState(vbKeyL) And &H8000) <> 0 Then mask = mask Or 8192
    
    ' F5 = 강제 세이브, F8 = 강제 로드, Space = 누르는 동안 2배속!
    If (GetAsyncKeyState(vbKeyF5) And &H8000) <> 0 Then VB_N64_SaveState
    If (GetAsyncKeyState(vbKeyF8) And &H8000) <> 0 Then VB_N64_LoadState
    If (GetAsyncKeyState(vbKeySpace) And &H8000) <> 0 Then
        VB_N64_SetDoubleSpeed 1
    Else
        VB_N64_SetDoubleSpeed 0
    End If
    
    VB_N64_SetInput mask, sx, sy
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_bRunning = False
    VB_N64_Shutdown
End Sub
