VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launchdeck"
   ClientHeight    =   390
   ClientLeft      =   825
   ClientTop       =   1380
   ClientWidth     =   3315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   3315
   Begin VB.PictureBox statusIcon 
      Height          =   375
      Left            =   3120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox statusHook 
      Height          =   255
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtEntry 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim T As NOTIFYICONDATA

Private Sub Form_Load()
    If Not GetAttr(App.Path & "\shortcuts\") And vbDirectory Then
        MkDir App.Path & "\shortcuts\"
    End If
    
    Dim intCount As Integer
    Dim fs, RAWFolder, File
    Set fs = CreateObject("Scripting.filesystemObject")
    Set RAWFolder = fs.GetFolder(App.Path & "\shortcuts\")
    intCount = 0
    For Each File In RAWFolder.Files
       intCount = intCount + 1
    Next
    
    If intCount = 0 Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You don't have any shortcuts. Would you like to import your Start Menu items?", vbQuestion + vbYesNo, "Launchdeck")
        If response = vbYes Then
            importStartMenuItems
        End If
    End If
    'Setup initial Tray Icon
    T.cbSize = Len(T)
    T.hWnd = statusHook.hWnd
    T.uId = 1&
    T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    T.ucallbackMessage = WM_MOUSEMOVE
    T.hIcon = statusIcon.Picture
    T.szTip = "Launchdeck" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, T
    
    'Hide this form
    Me.Hide
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Unload this form. Important: always end with "unload me".
    T.cbSize = Len(T)
    T.hWnd = statusHook.hWnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
    End
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim no As Boolean
    If txtEntry.Text = "/cfg" Or txtEntry.Text = "/config" Or txtEntry.Text = "/c" Or txtEntry.Text = "/" Then
        frmCfg.Show
    ElseIf txtEntry.Text = "/quit" Or txtEntry.Text = "/exit" Or txtEntry.Text = "/q" Or txtEntry.Text = "/e" Then
        Unload Me
        Exit Sub
    ElseIf txtEntry.Text = "/shutdown" Then
        no = ExitWindowsEx(EWX_SHUTDOWN, 0)
        Unload Me
    ElseIf txtEntry.Text = "/restart" Then
        no = ExitWindowsEx(EWX_REBOOT, 0)
        Unload Me
    Else
        Dim sh As Boolean
        Dim shortcut As String
        shortcut = App.Path & "\shortcuts\" & txtEntry.Text
        If Dir(shortcut) <> "" Then
           Call ShellExecute(0, "Open", shortcut, "", "", 1)
        Else
          If Dir(shortcut & ".lnk") <> "" Then
           shortcut = shortcut & ".lnk"
           Call ShellExecute(0, "Open", shortcut, "", "", 1)
          End If
        End If
    End If
    txtEntry.Text = ""
    Me.Hide
End If
End Sub

Private Sub statusHook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static State As Boolean
 Static Popped As Boolean
 Static Msg As Long
 
    Msg = X / Screen.TwipsPerPixelX
    If Popped = False Then
        'As we don't want any popups during the processing,
        'turn off the MouseMove event by setting popped=true
        Popped = True
        
        'Se the general section on how to interpret Msg
        Select Case Msg
            Case WM_LBUTTONDOWN
                State = Not State
                Me.Show
            
            Case WM_LBUTTONUP
            
            Case WM_RBUTTONDBLCLK
            
            Case WM_RBUTTONDOWN
            
            Case WM_RBUTTONUP
                
        End Select
        'OK to popup again
        Popped = False
    End If
End Sub
