Attribute VB_Name = "modAPI"
Global Const NIM_ADD = &H0&
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const NIF_INFO = &H10
Global Const NIM_SETVERSION = &H4

Global Const WM_MOUSEMOVE = &H200
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_RBUTTONDBLCLK = &H206
Global Const WM_RBUTTONDOWN = &H204
Global Const WM_RBUTTONUP = &H205

Global Const WM_USER = &H400

Global Const NIN_SELECT = WM_USER
Global Const NINF_KEY = &H1
Global Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Global Const NIN_BALLOONSHOW = (WM_USER + 2)
Global Const NIN_BALLOONHIDE = (WM_USER + 3)
Global Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Global Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Global NI As NOTIFYICONDATA

Public Enum EBalloonIconTypes
 NIIF_NONE = 0
 NIIF_INFO = 1
 NIIF_WARNING = 2
 NIIF_ERROR = 3
 NIIF_NOSOUND = &H10
End Enum


Type NOTIFYICONDATA
 cbSize As Long             ' 4
 hwnd As Long               ' 8
 uID As Long                ' 12
 uFlags As Long             ' 16
 uCallbackMessage As Long   ' 20
 hIcon As Long              ' 24
 szTip As String * 128      ' 152
 dwState As Long            ' 156
 dwStateMask As Long        ' 160
 szInfo As String * 256     ' 416
 uTimeOutOrVersion As Long  ' 420
 szInfoTitle As String * 64 ' 484
 dwInfoFlags As Long        ' 488
 guidItem As Long           ' 492
End Type

Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ShowBalloonTip( _
 ByVal sMessage As String, _
 Optional ByVal sTitle As String, _
 Optional ByVal eIcon As EBalloonIconTypes, _
 Optional ByVal lTimeOutMs = 90000, _
 Optional ByVal bModal As Boolean _
)
  
  Dim lR As Long

  NI.szInfo = sMessage
  NI.szInfoTitle = sTitle
  NI.uTimeOutOrVersion = lTimeOutMs
  NI.dwInfoFlags = eIcon
  NI.uFlags = NIF_INFO
  lR = Shell_NotifyIconA(NIM_MODIFY, NI)
         
  frmMain.BalloonOpen = True
  'bModal is not catched since the Windows API is not able to make
  'Ballon Tips modal!  -fschneider
         
End Sub

