VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Mavica Photo Viewer"
   ClientHeight    =   5115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6870
   Begin VB.PictureBox pnlDrive 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1545
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdCancelChange 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdChangeDrive 
         Caption         =   "Change"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.DriveListBox drvActive 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Select the Mavica disk drive"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.PictureBox pnlFolder 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   2400
      ScaleHeight     =   4065
      ScaleWidth      =   3945
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox cmbFormat 
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   1440
         List            =   "frmMain.frx":08FB
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelImport 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   3480
         Width           =   1215
      End
      Begin VB.DriveListBox drvImport 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   3480
         Width           =   1215
      End
      Begin VB.DirListBox dirImport 
         Height          =   2115
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Folder format"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Select output directory"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Image imgRect 
      BorderStyle     =   1  'Fixed Single
      Height          =   4920
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6600
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuDrive 
         Caption         =   "&Change drive"
      End
      Begin VB.Menu menuReload 
         Caption         =   "&Reload"
         Shortcut        =   ^R
      End
      Begin VB.Menu divide0 
         Caption         =   "-"
      End
      Begin VB.Menu menuImport 
         Caption         =   "&Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu divide1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErDisk 
         Caption         =   "Erase &disk"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuErImage 
         Caption         =   "Erase i&mage"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu divide2 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&?"
      Begin VB.Menu menuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: José Antonio Barranquero Fernández
'Date:   23rd February 2019
'Desc:   Application which searchs for a SONY Mavica disk and shows its contents

Const MIN_FORM_HEIGHT = 1290
Const IMAGE_MARGIN = 200
Const HORIZONTAL_RATIO = 4
Const VERTICAL_RATIO = 3

Private Sub Form_Load()
On Error Resume Next
    Me.WindowState = vbMaximized

    Call IO.Initialize(Me)
    
    Call IO.FileLoad(Me)
End Sub

'Resize the imgRect so the aspect ratio is not distorted
Private Sub Form_Resize()
On Error GoTo errorHandler
    If Me.WindowState = vbMinimized Then Exit Sub
        
    imgRect.Height = Me.ScaleHeight - IMAGE_MARGIN
    imgRect.Width = imgRect.Height * (HORIZONTAL_RATIO / VERTICAL_RATIO)            'Set imgRect to properly resize
    imgRect.Left = (Me.ScaleWidth / 2) - (imgRect.Width / 2)
    
    pnlDrive.Left = (Me.Width / 2) - (pnlDrive.Width / 2)
    pnlDrive.Top = (Me.Height / 2) - (pnlDrive.Height / 2)
    
    pnlFolder.Left = (Me.Width / 2) - (pnlFolder.Width / 2)
    pnlFolder.Top = (Me.Height / 2) - (pnlFolder.Height / 2)
    
    If Me.Width < imgRect.Width Then Me.Width = imgRect.Width + imgRect.Left
    
    Exit Sub
errorHandler:
    'Error 380: Invalid property value, height is negative
    If (Err.Number = 380) Then Me.Height = MIN_FORM_HEIGHT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then Call menuReload_Click: Exit Sub
        
    Select Case KeyCode
        Case vbKeyLeft
            MousePointer = vbHourglass: Call IO.ChangeImage(-1, imgRect): MousePointer = vbDefault
        Case vbKeyRight
            MousePointer = vbHourglass: Call IO.ChangeImage(1, imgRect): MousePointer = vbDefault
    End Select
End Sub

Private Sub menuDrive_Click()
    pnlDrive.Visible = Not pnlDrive.Visible
End Sub

Private Sub cmdChangeDrive_Click()
    pnlDrive.Visible = False
    Call IO.ChangeDrive(drvActive.drive, Me)
        
    Call IO.FileLoad(Me)
End Sub

Private Sub cmdCancelChange_Click()
    pnlDrive.Visible = False
End Sub

Private Sub mnuErDisk_Click()
    MousePointer = vbHourglass
    If IO.DeleteDisk Then
        Call IO.FileLoad(Me)
    End If
    MousePointer = vbDefault
End Sub

Private Sub mnuErImage_Click()
    MousePointer = vbHourglass
    If IO.DeleteImg Then
        Call IO.FileLoad(Me)
    End If
    MousePointer = vbDefault
End Sub

Private Sub pnlDrive_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call cmdCancelChange_Click
End Sub

Private Sub drvImport_Change()
    dirImport.path = Left$(drvImport.drive, 1) & ":"
End Sub

Private Sub menuImport_Click()
On Error GoTo errorHandler
    pnlFolder.Refresh
    pnlFolder.Visible = Not pnlFolder.Visible

errorHandler:
    If Err.Number = 380 Then INIwrite.RestoreFile
    If Err.Number = 68 Then INIwrite.RestoreFolder
    
    cmbFormat.ListIndex = INIread.ReadFormat
    dirImport.path = INIread.ReadFolder
    drvImport.drive = Left$(dirImport.path, 1) & ":"
End Sub

Private Sub cmdImport_Click()
On Error GoTo errorHandler
    INIwrite.WriteFormat cmbFormat.ListIndex
    INIwrite.WriteFolder dirImport.path
    
    pnlFolder.Visible = False
    
    MousePointer = vbHourglass
    Call IO.FileImport(IIf(Len(dirImport.path) <> 3, dirImport.path & "\", dirImport.path), cmbFormat.Text)
errorHandler:
    Select Case Err.Number
        Case 9      'Error 9:  No files
            MsgBox "No files found on drive", vbExclamation + vbOKOnly, "Import error"
        Case 53     'Error 53: Source files not found
            MsgBox "Original files not found", vbExclamation + vbOKOnly, "Import error"
        Case 75     'Error 75: Folder already exists
            MsgBox "Folder error", vbExclamation + vbOKOnly
        Case 76     'Error 76: File not found, serious shit, should not happen
            Err.Raise Err.Description
        Case 0
        Case Else
            MsgBox Err.Description, vbExclamation + vbOKOnly, "Import error"
    End Select
    MousePointer = vbDefault
End Sub

Private Sub cmdCancelImport_Click()
    pnlFolder.Visible = False
End Sub

Private Sub pnlFolder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call cmdCancelImport_Click
End Sub

Private Sub menuReload_Click()
    Call IO.FileLoad(Me)
End Sub

Private Sub menuExit_Click()
    Unload Me
End Sub

Private Sub menuAbout_Click()
    frmAbout.Show vbModal
End Sub
