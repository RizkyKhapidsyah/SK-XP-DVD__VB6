VERSION 5.00
Begin VB.Form frmPref 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteBookmark 
      Caption         =   "Delete Bookmark"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display on Start up"
      Height          =   2055
      Left            =   2400
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      Begin VB.CheckBox cbHideTool 
         Caption         =   "Hide Tool bar"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox cbHideStatus 
         Caption         =   "Hide Status Bar"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox cbHideMenu 
         Caption         =   "Hide Menu Bar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox cbFullScreen 
         Caption         =   "Play Full screen"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Set Playback Speed"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "2x"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3x"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4x"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "5x"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6x"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "7x"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public strValue As String

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub cbFullScreen_Click()

IIf cbFullScreen.Value = 1, sPref = "Full Screen", sPref = "Normal"

End Sub

Private Sub cbHideStatus_Click()

IIf cbHideStatus.Value = 1, strValue = "True", strValue = "False"

End Sub

Private Sub cmdDeleteBookmark_Click()
frmXP.ctlDVD.DeleteBookmark
End Sub

Private Sub Form_Load()
Cfg$ = App.Path & "\config.ini"
up$ = "User Preferences"

'Get the user preferences from the ini file and display them on the form
strTempSpeed = GetFromIni(up$, "Playback Speed", Cfg$)
strTempFS = GetFromIni(up$, "Full Screen", Cfg$)
strTempHM = GetFromIni(up$, "Hide Menu", Cfg$)
strTempHS = GetFromIni(up$, "Hide Status", Cfg$)
strTempHT = GetFromIni(up$, "Hide Toolbar", Cfg$)
'Convert the ini settings read in
With UserPref
    .intSpeed = CInt(strTempSpeed)
    .bFullScreen = CBool(strTempFS)
    .bHideMenu = CBool(strTempHM)
    .bHideStatus = CBool(strTempHS)
    .bHideTool = CBool(strTempHT)
End With

If UserPref.bFullScreen = True Then
    cbFullScreen.Value = 1
Else
    cbFullScreen.Value = 0
End If

If UserPref.bHideMenu = True Then
    cbHideMenu.Value = 1
Else
    cbHideMenu.Value = 0
End If

If UserPref.bHideStatus = True Then
    cbHideMenu.Value = 1
Else
    cbHideMenu.Value = 0
End If

If UserPref.bHideTool = True Then
    cbHideTool.Value = 1
Else
    cbHideTool.Value = 0
End If

If UserPref.intSpeed = 2 Then
    Me.Option1(0).Value = True
ElseIf UserPref.intSpeed = 3 Then
    Me.Option1(1).Value = True
ElseIf UserPref.intSpeed = 4 Then
    Me.Option1(2).Value = True
ElseIf UserPref.intSpeed = 5 Then
    Me.Option1(3).Value = True
ElseIf UserPref.intSpeed = 6 Then
    Me.Option1(4).Value = True
ElseIf UserPref.intSpeed = 7 Then
    Me.Option1(5).Value = True
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmXP.ctlDVD.Play
ContPlay
End Sub

Private Sub OKButton_Click()
Cfg$ = App.Path & "\config.ini"

'Set the preferences and write the settings out to the ini file.
If cbFullScreen.Value = 1 Then
    UserPref.bFullScreen = True
    frmXP.ctlDVD.FullScreenMode = True
ElseIf cbFullScreen.Value = 0 Then
    UserPref.bFullScreen = False
    frmXP.ctlDVD.FullScreenMode = False
End If

If cbHideMenu.Value = 1 Then
    UserPref.bHideMenu = True
    With frmXP
        .mnuEdit.Visible = False
        .mnuHelp.Visible = False
        .mnuControls.Visible = False
    End With
ElseIf cbHideMenu.Value = 0 Then
    UserPref.bHideMenu = False
    With frmXP
        .mnuEdit.Visible = True
        .mnuHelp.Visible = True
        .mnuControls.Visible = True
    End With
End If

If cbHideStatus.Value = 1 Then
    UserPref.bHideStatus = True
    frmXP.StatusBar1.Visible = False
ElseIf cbHideStatus.Value = 0 Then
    UserPref.bHideStatus = False
    frmXP.StatusBar1.Visible = True
End If

If cbHideTool.Value = 1 Then
    UserPref.bHideTool = True
    frmXP.Toolbar1.Visible = False
ElseIf cbHideTool.Value = 0 Then
    UserPref.bHideTool = False
    frmXP.Toolbar1.Visible = True
End If

'Write out Preferences
f% = WriteToIni(up$, "Playback Speed", CStr(UserPref.intSpeed), Cfg$)
f% = WriteToIni(up$, "Full Screen", CStr(UserPref.bFullScreen), Cfg$)
f% = WriteToIni(up$, "Hide Menu", CStr(UserPref.bHideMenu), Cfg$)
f% = WriteToIni(up$, "Hide Status", CStr(UserPref.bHideStatus), Cfg$)
f% = WriteToIni(up$, "Hide Toolbar", CStr(UserPref.bHideTool), Cfg$)

frmXP.ctlDVD.Play

Unload Me

End Sub

Private Sub Option1_Click(index As Integer)
    
    Select Case index
        Case 0
            UserPref.intSpeed = 2
        Case 1
            UserPref.intSpeed = 3
        Case 2
            UserPref.intSpeed = 4
        Case 3
            UserPref.intSpeed = 5
        Case 4
            UserPref.intSpeed = 6
        Case 5
            UserPref.intSpeed = 7
    End Select
    
End Sub


