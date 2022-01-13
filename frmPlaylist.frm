VERSION 5.00
Begin VB.Form frmPlaylist 
   Caption         =   "Playlist"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playlist"
      Height          =   3855
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.ListBox lstPlaylist 
         Height          =   3375
         ItemData        =   "frmPlaylist.frx":0442
         Left            =   120
         List            =   "frmPlaylist.frx":0444
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make your selection from the DVD titles in your list >>"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image imgPicList 
      Height          =   4215
      Left            =   120
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pics(2) As String
Dim strTitleName As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'Identify the DVD title and issue commands
    
    Select Case strTitleName

        Case "Contact"
            If frmXP.ctlDVD.DVDUniqueID = TitleID(0) Then
                MsgBox "You are already playing this DVD title." & vbCrLf & "" & vbCrLf & "Enjoy!", _
                vbOKOnly + vbInformation + vbDefaultButton1, _
                "DVD Now Playing"
                With frmXP
                    .tmrMainFrm.Enabled = False
                    .ctlDVD.Visible = True
                    .ctlDVD.Play
                End With
            Else
                With frmXP
                    .ctlDVD.Stop
                    .ctlDVD.Eject
                    MsgBox "Please insert the DVD titled " & TitleName(0) & " into the DVD-ROM drive." _
                    & vbCrLf & "" & vbCrLf & "Enjoy your movie!", _
                    vbOKOnly + vbInformation + vbDefaultButton1, _
                    "DVD Selected"
                End With
                TitleCaption = "You are watching: " & TitleName(0)
                frmXP.StatusBar1.Panels.Item(1) = Mid$(TitleCaption, 18, Len(TitleCaption))
            End If
        Case "Snoop"
            If frmXP.ctlDVD.DVDUniqueID = TitleID(1) Then
                MsgBox "You are already playing this DVD title." & vbCrLf & "" & vbCrLf & "Enjoy!", _
                vbOKOnly + vbInformation + vbDefaultButton1, _
                "DVD Now Playing"
            Else
                With frmXP
                    .ctlDVD.Stop
                    .ctlDVD.Eject
                    MsgBox "Please insert the DVD titled " & TitleName(1) & " into the DVD-ROM drive." _
                    & vbCrLf & "" & vbCrLf & "Enjoy your movie!", _
                    vbOKOnly + vbInformation + vbDefaultButton1, _
                    "DVD Selected"
                End With
                TitleCaption = "You are watching: " & TitleName(1)
                frmXP.StatusBar1.Panels.Item(1) = Mid$(TitleCaption, 18, Len(TitleCaption))
            End If
            

    End Select

Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer

'Change this to read a directory and populate the Pics() with filenames
Pics(0) = App.Path & "\graphics\" & TitleID(0) & ".jpg"
Pics(1) = App.Path & "\graphics\snoop.jpg"

frmXP.ctlDVD.Pause

For i = 0 To 1
    lstPlaylist.AddItem (TitleName(i))
Next

If frmXP.ctlDVD.DVDUniqueID = TitleID(0) Then
    imgPicList.Picture = LoadPicture(Pics(0))
    lstPlaylist.Selected(0) = True
ElseIf frmXP.ctlDVD.DVDUniqueID = TitleID(1) Then
    imgPicList.Picture = LoadPicture(Pics(1))
    lstPlaylist.Selected(1) = True
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmXP.ctlDVD.Play
    With frmXP.Toolbar1
        For i = 8 To 10
            .Buttons(i).Enabled = True
        Next i
        .Buttons(11).Enabled = False
        For i = 12 To 15
            .Buttons(i).Enabled = True
        Next i
        .Buttons(16).Enabled = False
    End With
End Sub

Private Sub lstPlaylist_Click()
frmXP.ctlDVD.Pause

If lstPlaylist.Selected(0) = True Then
    imgPicList.Picture = LoadPicture(Pics(0))
    strTitleName = "Contact"
ElseIf lstPlaylist.Selected(1) = True Then
    imgPicList.Picture = LoadPicture(Pics(1))
    strTitleName = "Snoop"
End If

End Sub
