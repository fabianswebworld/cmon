VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CMon 1.2.1"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDisp 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3120
      Top             =   2235
   End
   Begin VB.Timer tmrPollIP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   2235
   End
   Begin VB.Timer tmrPollCOM 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4365
      Top             =   2280
   End
   Begin VB.OptionButton optCOM 
      Caption         =   "Modem (oder ISDN-Karte mit COM-Port)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cboCOM 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":0E42
      Left            =   2640
      List            =   "frmMain.frx":0E76
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2385
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Einstellungen speichern"
      Height          =   360
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   $"frmMain.frx":0EE1
      Top             =   2040
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin VB.PictureBox picTray 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   2115
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picTrayNormal 
      DragIcon        =   "frmMain.frx":0F6C
      Height          =   135
      Left            =   960
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2235
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTrayAnnoy 
      DragIcon        =   "frmMain.frx":10B6
      Height          =   375
      Left            =   3960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Beenden"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Beendet das Programm."
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtFBAddr 
      Height          =   300
      Left            =   2640
      TabIndex        =   0
      Text            =   "fritz.box"
      ToolTipText     =   "Hier die IP-Adresse oder den Hostnamen der FRITZ!Box Fon eintragen (Standard: fritz.box)"
      Top             =   1400
      Width           =   1935
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "&Aktivieren -> Tray-Icon"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Aktiviert die Verbindung zur FRITZ!Box Fon und minimiert das Programm ins Tray."
      Top             =   2520
      Width           =   2415
   End
   Begin VB.OptionButton optIP 
      Caption         =   "FRITZ!Box-Anrufmonitor"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1070
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame fraIP 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   4665
      Begin VB.Label lblFBA 
         BackStyle       =   0  'Transparent
         Caption         =   "IP-Adresse der FRITZ!Box:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraCOM 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   4680
      Begin VB.Label lblCOM 
         BackStyle       =   0  'Transparent
         Caption         =   "COM-Anschluss für Modem:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label lblWelcome 
      Caption         =   $"frmMain.frx":1200
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Info über CMon..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008F0000&
      Height          =   255
      Left            =   3210
      MouseIcon       =   "frmMain.frx":1293
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3060
      Width           =   1470
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2008-2013 Fabian Schneider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   285
      TabIndex        =   8
      Top             =   3060
      Width           =   3120
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "Menü"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Wiederherstellen"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZeilenFeld(1000) As String
Public ZeilenP, RingCount As Integer
Public BalloonOpen, FBConnected, CIDRing, NamedRing, NmbrdRing As Boolean
Public CIDNmbr, CIDName As String

Public bEcho As Boolean
Public bOK As Boolean
Public bRing As Boolean
Public bError As Boolean
Public LastRing As Double
 
Dim hSock As Long

Private Sub cmdExit_Click()
  If chkSave.Value = vbChecked Then
  Open "cmon.ini" For Output As #2
  Print #2, "[General]"
  Print #2, "FBAddress="; txtFBAddr
  Print #2, "COMPort="; cboCOM
  Print #2, "Mode=IP"
  Close
  End If

Unload Me
End
End Sub

Private Sub Form_Load()

  LastRing = 0
  RingCount = 0
  
  FBConnected = False
  On Error Resume Next
  
  Open "cmon.ini" For Input As 1
  
  Input #1, Dummy
  Input #1, FBTempstr
  If Not EOF(1) Then Input #1, COMTempstr
  If Not EOF(1) Then Input #1, ModeTempstr
  
  txtFBAddr = Split(FBTempstr, "=")(1)
  cboCOM = Split(COMTempstr, "=")(1)
  ModeTempstr = Split(ModeTempstr, "=")(1)
  
  If ModeTempstr = "IP" Then optIP.Value = True Else optCOM.Value = True
  
  Close
  
  picTray.DragIcon = picTrayNormal.DragIcon
  
  Dim Retval As Long, WSD As WSAData
   
  Retval = WSAStartup(&H202, WSD)
  If Retval <> 0 Then
    MsgBox "Der TCP-Stack konnte nicht initialisiert werden.", vbExclamation, "WinSock-Fehler"
    Unload Me
    Exit Sub
  End If
    
  If InStr(1, UCase(Command), "/TRAY") > 0 Then Call cmdInit_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call Disconnect(hSock)
  Call WSACleanup
End Sub

Private Sub cmdInit_Click()

' Erstmal im Tray "gemütlich machen"

  frmMain.WindowState = vbMinimized
  App.TaskVisible = False
 
  InTray = True
 
  NI.cbSize = 504
  NI.hwnd = picTray.hwnd
  NI.uID = 123
  NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
  NI.uCallbackMessage = WM_MOUSEMOVE
 
  ModifyTrayIcon

  Shell_NotifyIconA NIM_ADD, NI
  Shell_NotifyIconA NIM_SETVERSION, NI
 
  frmMain.Hide

  If FBConnected = False Then
    connectIP
 
    ' Empfang abwarten
    FBConnected = True
    tmrDisp.Enabled = True
  End If
  
End Sub

Private Sub connectIP()
  
  ' Mit FB verbinden auf Port 1012

  Dim ServerIP As String
 
  ' Eventuell vorherigen Socket schließen
  If hSock <> 0 Then
    Call Disconnect(hSock)
  End If
 
  ' ServerIP ermitteln
  ServerIP = GetIP(txtFBAddr)
  If ServerIP = "" Then
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    ShowBalloonTip "Die FRITZ!Box ist unter der angegebenen Adresse nicht erreichbar.", "CMon :: Fehler", NIIF_INFO, 900000, True
    Do
      DoEvents
      Sleep 500
    Loop Until BalloonOpen = False
    Unload Me: End
    Exit Sub
  End If
 
  ' Verbinden mit dem Server
  hSock = ConnectToServer(ServerIP, 1012)
  If hSock = -1 Then
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    ShowBalloonTip "Fehler bei der Verbindung mit der FRITZ!Box.", "CMon :: Fehler", NIIF_INFO, 900000, True
    Do
      DoEvents
      Sleep 500
    Loop Until BalloonOpen = False
    hSock = 0
    Exit Sub
  End If

  tmrPollIP.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

Private Sub lblInfo_Click()
  MsgBox "CMon 1.2.1" & vbCrLf & vbCrLf & "Kleiner Anrufmonitor für die FRITZ!Box-Fon-Familie von AVM, welcher ohne Java, .NET oder andere Laufzeitumgebungen auskommt. Dieses Programm ist Freeware. Die neueste Version finden Sie jeweils unter:" & vbCrLf & vbCrLf & "www.fabianswebworld.de" & vbCrLf & vbCrLf & "HINWEIS: FRITZ!Box und FRITZ!Box Fon sind eingetragene Markenzeichen der AVM GmbH, Berlin.", vbOKOnly + vbInformation, "Info über CMon"
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuRestore_Click()
  InTray = False
  App.TaskVisible = True
  tmrPollIP.Enabled = False
  tmrPollCOM.Enabled = False
  tmrDisp.Enabled = False
  frmMain.WindowState = OldWinState
  frmMain.Show
  Call Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

Private Sub tmrPollIP_Timer()
  On Error GoTo IpError
  Dim Zeile As String
  If DataComeIn(hSock) <= 1 Then
  ' Ankommenden Ruf von FB verarbeiten
    
    Zeile = GetData(hSock)

    If Zeile = "" Then Exit Sub
    SplitZeile = Split(Zeile, vbCrLf)
    
    i = 0
    
    Do
      ZeilenP = ZeilenP + 1
      ZeilenFeld(ZeilenP) = SplitZeile(i)
      i = i + 1
    Loop Until SplitZeile(i) = ""
    
  End If
    
  Exit Sub

IpError:
  'DEBUG
  'MsgBox Zeile
  Resume Next
  
End Sub

Private Sub tmrDisp_Timer()

  If ZeilenFeld(1) = "" Then Exit Sub
    
  SplitFeld = Split(ZeilenFeld(1), ";")
 
  If SplitFeld(4) = "COM" Then
    picTray.DragIcon = picTrayAnnoy.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    If SplitFeld(3) = "" Then SplitFeld(3) = "unbekannter Nummer"
    ShowBalloonTip "Eingehender Anruf von " & SplitFeld(3) & " an Modem (" & cboCOM & ")", _
                   "CMon :: Eingehender Ruf", NIIF_INFO, 900000, True
    GoTo CleanFeld
  End If
 
  If InStr(1, ZeilenFeld(1), "RING") > 0 Then
    picTray.DragIcon = picTrayAnnoy.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    If SplitFeld(3) = "" Then SplitFeld(3) = "unbekannter Nummer"
    ShowBalloonTip "Eingehender Anruf von " & SplitFeld(3) & " an eigener Nummer " & SplitFeld(4), _
                   "CMon :: Eingehender Ruf", NIIF_INFO, 900000, True
    Do
      DoEvents
      Sleep 500
    Loop Until BalloonOpen = False
  End If
  
  If InStr(1, ZeilenFeld(1), "CALL") > 0 Then
    picTray.DragIcon = picTrayAnnoy.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    ShowBalloonTip "Eigene Nummer " & SplitFeld(4) & " ruft gerade die Nummer " & SplitFeld(5) & " an.", _
                   "CMon :: Abgehender Ruf", NIIF_INFO, 900000, True
    Do
      DoEvents
      Sleep 500
    Loop Until BalloonOpen = False
  End If
  
  If InStr(1, ZeilenFeld(1), "DISCONNECT") > 0 Then
    picTray.DragIcon = picTrayAnnoy.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
    ShowBalloonTip "Das zuvor geführte Gespräch wurde soeben beendet.", _
                   "CMon :: Telefon aufgelegt", NIIF_INFO, 900000, True
    Do
      DoEvents
      Sleep 500
    Loop Until BalloonOpen = False
  End If
  
  
  picTray.DragIcon = picTrayNormal.DragIcon
  NI.hIcon = picTray.DragIcon
  NI.uFlags = NIF_MESSAGE Or NIF_ICON
  Shell_NotifyIconA NIM_MODIFY, NI


CleanFeld:
  ii = 0
  Do
    ii = ii + 1
    ZeilenFeld(ii) = ZeilenFeld(ii + 1)
  Loop Until ZeilenFeld(ii) = ""
  ZeilenP = 0

End Sub


Public Sub ModifyTrayIcon()
  tip = "CMon 1.2.1  © 2008-2013 Fabian Schneider" & Chr(0)
 
  NI.szTip = tip
  picTray.DragIcon = picTrayNormal.DragIcon
  NI.hIcon = picTray.DragIcon
  NI.dwInfoFlags = 0
  NI.szInfo = ""
  NI.szInfoTitle = ""
  NI.uTimeOutOrVersion = 0
  
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Msg As Long
  Msg = ScaleX(X, Me.ScaleMode, vbPixels)
  Select Case Msg
  
  Case WM_LBUTTONUP
    Call mnuRestore_Click
       
  Case WM_RBUTTONUP
    PopupMenu mnuIcon
  
  Case NIN_BALLOONHIDE
    BalloonOpen = False
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    picTray.DragIcon = picTrayNormal.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
  
  Case NIN_BALLOONTIMEOUT
    BalloonOpen = False
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    picTray.DragIcon = picTrayNormal.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
  
  
  Case NIN_BALLOONUSERCLICK
    BalloonOpen = False
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    picTray.DragIcon = picTrayNormal.DragIcon
    NI.hIcon = picTray.DragIcon
    NI.uFlags = NIF_MESSAGE Or NIF_ICON
    Shell_NotifyIconA NIM_MODIFY, NI
        
  End Select
  
End Sub


Private Sub Wait()
  Dim Start

  Start = Timer
  Do While Timer < Start + 2
    DoEvents
    If bOK Then
      Exit Sub
    End If
    If bError Then
      Exit Sub
    End If
  Loop
End Sub

