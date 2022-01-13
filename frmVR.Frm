VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "mswebdvd.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "XPDVD Player"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10755
   Icon            =   "frmVR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imglstNormal 
      Left            =   960
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   3342130
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":1674
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":1C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":21A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":2742
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":2CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":3276
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":3810
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":3DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":4344
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":4E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":5412
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":59AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":5F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":64E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":6A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":7014
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstHot 
      Left            =   240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   3342130
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":75AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":7B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":80E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":867C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":8C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":91B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":974A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":9CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":A27E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":A818
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":ADB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":B34C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":B8E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":BE80
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":C41A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":C9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":CF4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":D4E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":DA82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":E01C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVR.frx":E5B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imglstNormal"
      HotImageList    =   "imglstHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnUp"
            Object.ToolTipText     =   "Button Up"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnDn"
            Object.ToolTipText     =   "Button Down"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnLt"
            Object.ToolTipText     =   "Button Left"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BtnRt"
            Object.ToolTipText     =   "Button Right"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous Chapter"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rewind"
            Object.ToolTipText     =   "Rewind"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            Object.ToolTipText     =   "Play"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.ToolTipText     =   "Pause"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FF"
            Object.ToolTipText     =   "Fast Forward"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Chapter"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eject"
            Object.ToolTipText     =   "Eject"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mute"
            Object.ToolTipText     =   "Mute"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pref"
            Object.ToolTipText     =   "Preferences"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayList"
            Object.ToolTipText     =   "Playlist"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MovieInfo"
            Object.ToolTipText     =   "Movie Info"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "About XPDVD"
            ImageIndex      =   21
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   7000
      Left            =   8400
      Top             =   5900
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7800
      Top             =   5880
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   5670
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3757
            Object.ToolTipText     =   "DVD Title Playing"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1552
            MinWidth        =   1552
            Object.ToolTipText     =   "Play State"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1676
            MinWidth        =   1676
            Object.ToolTipText     =   "Sound State"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3757
            Text            =   "Title:      Chapter:"
            TextSave        =   "Title:      Chapter:"
            Object.ToolTipText     =   "Title and Chapter Number"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3757
            Text            =   "Time:"
            TextSave        =   "Time:"
            Object.ToolTipText     =   "DVD Running Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3757
            Text            =   "Length:"
            TextSave        =   "Length:"
            Object.ToolTipText     =   "DVD Total Running Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrMainFrm 
      Interval        =   10000
      Left            =   360
      Top             =   6720
   End
   Begin MSWEBDVDLibCtl.MSWebDVD ctlDVD 
      Height          =   6100
      Left            =   -60
      TabIndex        =   0
      Top             =   -405
      Visible         =   0   'False
      Width           =   10800
      _cx             =   19050
      _cy             =   10760
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   983055
      WindowlessActivation=   0   'False
   End
   Begin VB.Image Image1 
      Height          =   6480
      Left            =   0
      Picture         =   "frmVR.frx":EB50
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuPref 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuSpacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlaylist 
         Caption         =   "Show Play&list"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "Player &Controls"
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuFastForward 
         Caption         =   "Fast Forward"
      End
      Begin VB.Menu mnuRewind 
         Caption         =   "Rewind"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next Chapter"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous Chapter"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEject 
         Caption         =   "Eject"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Help Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Spacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovieInfo 
         Caption         =   "Movie &Info"
      End
      Begin VB.Menu mnuWebAddress 
         Caption         =   "Your_Site_Name.com"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iTotalTime As String
Dim strChap As String, strTitle As String, strTemp As String, strMute As String
Dim strRunningTime As String, strPlayState As String
Dim tempTitle As String, Response As Integer, Style As String

Private Sub btnExit_Click()
Unload Me
End
End Sub

Private Sub Form_Load()

'-----------Start Initial Settings-------------------------------------
'Get the DVD title information
GetDVDInfo

'Get the user preferences
GetPref

'Set the logo image & screen size
With Me
    If Toolbar1.Visible = True Then
        .Image1.Top = Top '+ Toolbar1.Height
    ElseIf Toolbar1.Visible = False Then
        .Image1.Top = 0
    End If
    .Image1.Left = 0
    .Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    .Top = (Screen.Height - Height) / 2   ' Center form vertically.
    .Width = Screen.Width * 0.725  ' Set width of form.
    .Height = Screen.Height * 0.75   ' Set height of form.
End With

'Hide the DVD control
ctlDVD.Visible = False

'Set initial panel display values
strTitle = "Title: 0"
strChap = "Chapter: 0"
strRunningTime = "Time: 00:00:00"
iTotalTime = "Length: 00:00:00"

'Set the statusbar's initial panel display
With StatusBar1.Panels
    .Item(1) = Mid$(TitleCaption, 18, Len(TitleCaption))
    .Item(4) = strTitle & Space$(5) & strChap
    .Item(5) = strRunningTime
    .Item(6) = iTotalTime
End With

'Set toolbar initial display
With Toolbar1
    For i = 1 To 26
        .Buttons(i).Enabled = False
    Next i
End With
'---------------------End Initial Settings--------------------------------
End Sub

Private Sub Form_Resize()

Image1.Height = Height - StatusBar1.Height
Image1.Width = Width
ctlDVD.Height = Me.Height - StatusBar1.Height
ctlDVD.Width = Me.Width

Refresh

End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
Pause
End Sub

Private Sub mnuContents_Click()
Call Shell("hh.exe help.chm", vbNormalFocus)
Pause
End Sub

Private Sub mnuEject_Click()
Me.ctlDVD.Eject
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuFastForward_Click()
mnuFastForward.Checked = Not (mnuFastForward.Checked)
    If Me.mnuFastForward.Checked = True Then
        Me.ctlDVD.PlayForwards (UserPref.intSpeed)
    ElseIf Me.mnuFastForward.Checked = False Then
        Me.ctlDVD.Play
    End If
End Sub

Private Sub mnuControls_Click() 'This is the control menu item

'Find state
If ctlDVD.PlayState = dvdState_Running Then
    RunningState
ElseIf ctlDVD.PlayState = dvdState_Paused Then
    PausedState
End If

    
End Sub

Private Sub mnuWebAddress_Click()
ShellExecute hwnd, "open", "http://www.your_site_name.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub mnuMovieInfo_Click()
frmAboutTitle.Show
Pause
End Sub

Private Sub mnuMute_Click()
If mnuMute.Checked = False Then mnuMute.Checked = True Else mnuMute.Checked = False
If Me.mnuMute.Checked = True Then
    Me.ctlDVD.Mute = True
ElseIf Me.mnuMute.Checked = False Then
    Me.ctlDVD.Mute = False
End If

End Sub

Private Sub mnuNext_Click()
Me.ctlDVD.PlayNextChapter
End Sub

Private Sub mnuPause_Click()
Pause
End Sub

Private Sub mnuPlay_Click()
If ctlDVD.PlayState = dvdState_Stopped Then
    Play
ElseIf ctlDVD.PlayState = dvdState_Paused Then
   ContPlay
End If
Image1.Visible = False
End Sub

Private Sub mnuPlaylist_Click()
frmPlaylist.Show
Pause
End Sub

Private Sub mnuPref_Click()
    frmPref.Show
    Me.ctlDVD.Pause
End Sub

Private Sub mnuPrevious_Click()
Me.ctlDVD.PlayPrevChapter
End Sub

Private Sub mnuRewind_Click()
mnuRewind.Checked = Not (mnuRewind.Checked)
    If Me.mnuRewind.Checked = True Then
        Me.ctlDVD.PlayBackwards (UserPref.intSpeed)
    ElseIf Me.mnuRewind.Checked = False Then
        Me.ctlDVD.Play
    End If
End Sub

Private Sub mnuStop_Click()
Stopped
End Sub

Private Sub Timer1_Timer()

On Error Resume Next

strRunningTime = "Time: " & Mid$(Me.ctlDVD.CurrentTime, 1, 8)
strTitle = "Title: " & Me.ctlDVD.CurrentTitle
strChap = "Chapter: " & Me.ctlDVD.CurrentChapter
iTotalTime = "Length: " & Left$(Me.ctlDVD.TotalTitleTime, 8)

If Me.ctlDVD.PlayState = dvdState_Paused Then
    strPlayState = "Paused"
    StatusBar1.Panels.Item(2).Bevel = sbrRaised
ElseIf Me.ctlDVD.PlayState = dvdState_Running Then
    strPlayState = "Playing"
    StatusBar1.Panels.Item(2).Bevel = sbrInset
    StatusBar1.Panels.Item(2).Enabled = True
ElseIf Me.ctlDVD.PlayState = dvdState_Stopped Then
    strPlayState = "Stopped"
    StatusBar1.Panels.Item(2).Bevel = sbrRaised
    StatusBar1.Panels.Item(2).Enabled = False
Else
    strPlayState = ""
End If

If Me.mnuMute.Checked = True Then
    strMute = "Sound Off"
    StatusBar1.Panels.Item(3).Bevel = sbrRaised
ElseIf Me.mnuMute.Checked = False Then
    strMute = "Sound On"
    StatusBar1.Panels.Item(3).Bevel = sbrInset
End If

If strTitle = 0 Then
    With StatusBar1.Panels
        .Item(2) = strPlayState
        .Item(3) = strMute
        .Item(4) = strTitle & Space$(5) & strChap
        .Item(5) = strRunningTime
        .Item(6) = iTotalTime
    End With
End If

End Sub

Private Sub Timer2_Timer()
If ctlDVD.DVDUniqueID = TitleID(0) Then
    frmNowShowing.Show
Else
    Timer2.Enabled = False
    Exit Sub
End If
'Play the video
Me.ctlDVD.Play
Timer2.Enabled = False
End Sub

Private Sub tmrMainFrm_Timer()
With Me
    .ctlDVD.Visible = True
    .StatusBar1.Enabled = True
    If ctlDVD.PlayState = dvdState_Running Then
        .mnuPlay.Enabled = False
        With .Toolbar1
            For i = 1 To 10
                .Buttons(i).Enabled = True
            Next i
            .Buttons(11).Enabled = False
            For i = 12 To 14
                .Buttons(i).Enabled = True
            Next i
            .Buttons(16).Enabled = False
            For i = 18 To 26
                .Buttons(i).Enabled = True
            Next i
        End With
    Else
        .mnuPlay.Enabled = True
    End If
    .ctlDVD.Visible = True
    .Caption = TitleCaption
    .Image1.Visible = False
End With

If UserPref.bFullScreen = True Then
    ctlDVD.FullScreenMode = True
    frmPref.cbFullScreen.Value = 1
ElseIf UserPref.bFullScreen = False Then
    ctlDVD.FullScreenMode = False
    frmPref.cbFullScreen.Value = 0
End If

tmrMainFrm.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

'Add the mute off button
'Add a menu drop down
Select Case Button.Key
    
    Case "Exit"
        Unload Me
        End
    Case "BtnUp"
        ctlDVD.SelectUpperButton
    Case "BtnDn"
        ctlDVD.SelectLowerButton
    Case "BtnLt"
        ctlDVD.SelectLeftButton
    Case "BtnRt"
        ctlDVD.SelectRightButton
    Case "Previous"
        ctlDVD.PlayPrevChapter
    Case "Rewind"
        ctlDVD.PlayBackwards (UserPref.intSpeed)
        mnuRewind.Checked = Not (mnuRewind.Checked)
        PlayOnly
    Case "Stop"
        Stopped
    Case "Play"
        Play
    Case "Pause"
        Pause
    Case "FF"
        ctlDVD.PlayForwards (UserPref.intSpeed)
        mnuFastForward.Checked = Not (mnuFastForward.Checked)
        PlayOnly
    Case "Next"
        ctlDVD.PlayNextChapter
    Case "Eject"
        ctlDVD.Eject
    Case "Mute"
        ctlDVD.Mute = Not (ctlDVD.Mute)
        If ctlDVD.Mute = True Then
            mnuMute.Checked = True
            StatusBar1.Panels(3).Bevel = sbrRaised
            strMute = "Sound Off"
        ElseIf ctlDVD.Mute = False Then
            mnuMute.Checked = False
            StatusBar1.Panels(3).Bevel = sbrInset
            strMute = "Sound On"
        End If
    Case "Pref"
        frmPref.Show
        Pause
    Case "PlayList"
        frmPlaylist.Show
        Pause
    Case "MovieInfo"
        frmAboutTitle.Show
        Pause
    Case "Help"
        Call Shell("hh.exe help.chm", vbNormalFocus)
        Pause
    Case "About"
        frmAbout.Show
        Pause
End Select

End Sub
