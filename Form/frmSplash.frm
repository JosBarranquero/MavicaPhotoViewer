VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5400
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   480
      Top             =   1800
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    INIwrite.CheckFileExists
    Randomize
    tmrStart.Interval = Rnd() * 5000
End Sub

Private Sub tmrStart_Timer()
    tmrStart.Enabled = False
    frmMain.Show
    Unload Me
End Sub

