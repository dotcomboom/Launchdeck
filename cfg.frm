VERSION 5.00
Begin VB.Form frmCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Launchdeck Configuration"
   ClientHeight    =   4020
   ClientLeft      =   780
   ClientTop       =   2130
   ClientWidth     =   4845
   Icon            =   "cfg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnRemove 
      Caption         =   "-"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import Start Menu Items"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox lstCommands 
      Height          =   840
      ItemData        =   "cfg.frx":0442
      Left            =   120
      List            =   "cfg.frx":0452
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.FileListBox filShortcuts 
      Height          =   2430
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblCommandTitle 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblCommandDesc 
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label lblCommands 
      Caption         =   "Commands:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblShortcuts 
      Caption         =   "Shortcuts:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnImport_Click()
btnImport.Enabled = False
btnRemove.Enabled = False
importStartMenuItems
filShortcuts.Refresh
btnImport.Enabled = True
btnRemove.Enabled = True
End Sub

Private Sub btnRemove_Click()
On Error GoTo no
Dim idx As Integer
Dim max As Integer
With filShortcuts
    max = filShortcuts.ListCount - 1
    For idx = 0 To max
    If filShortcuts.Selected(idx) Then
        Kill App.Path & "\shortcuts\" & filShortcuts.List(idx)
    End If
Next
End With
filShortcuts.Refresh

no:

End Sub

Private Sub filShortcuts_Click()
Dim idx As Integer
Dim max As Integer
Dim yay As Boolean
yay = False
With filShortcuts
    max = filShortcuts.ListCount - 1
    For idx = 0 To max
    If filShortcuts.Selected(idx) Then
        yay = True
    End If
Next
End With
    btnRemove.Enabled = yay
End Sub

Private Sub filShortcuts_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim intFile As Integer
  On Error GoTo no
  
  With Data
    For intFile = 1 To .Files.Count
      FileCopy Data.Files.Item(intFile), App.Path & "\shortcuts\" & GetFileNameFromPath(Data.Files.Item(intFile))
      filShortcuts.Refresh
    Next intFile
  End With
  
no:
  
End Sub

Private Sub Form_Load()
filShortcuts.Path = App.Path & "\shortcuts\"
End Sub

Private Sub lstCommands_Click()
If lstCommands.ListIndex > -1 Then
    lblCommandTitle.Caption = lstCommands.List(lstCommands.ListIndex)
    If lstCommands.ListIndex = 0 Then
    lblCommandDesc.Caption = "Opens the configuration window."
    ElseIf lstCommands.ListIndex = 1 Then
    lblCommandDesc.Caption = "Exits Launchdeck."
    ElseIf lstCommands.ListIndex = 2 Then
    lblCommandDesc.Caption = "Turns off your computer. (9x)"
    ElseIf lstCommands.ListIndex = 3 Then
    lblCommandDesc.Caption = "Restarts your computer. (9x)"
    End If
End If
End Sub
