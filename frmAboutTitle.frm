VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAboutTitle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOW SHOWING"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   Icon            =   "frmAboutTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Return to Movie"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin TabDlg.SSTab SSTab1 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   6800
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Summary"
         TabPicture(0)   =   "frmAboutTitle.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtSummary"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Actors"
         TabPicture(1)   =   "frmAboutTitle.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtActors"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Additional Info"
         TabPicture(2)   =   "frmAboutTitle.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtAdditional"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txtAdditional 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   3015
            Left            =   -74760
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtActors 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   -74760
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtSummary 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   480
            Width           =   4215
         End
      End
   End
   Begin VB.Image imgPic 
      Appearance      =   0  'Flat
      Height          =   4815
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmAboutTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSummary As String
Dim strProducer As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim strTemp As String
Dim i As Integer
Dim pix As String, Snoop As String

'Set varaibles values
f% = FreeFile
pix = App.Path & "\graphics\" & TitleID(0) & ".jpg"

'Display info - This should be redone to simplify and make dynamic
Frame1.Caption = Mid$(TitleCaption, 18, Len(TitleCaption)) & " Movie Info"
If frmXP.ctlDVD.DVDUniqueID = TitleID(0) Then
    imgPic.Picture = LoadPicture(pix)
    Open App.Path & "\contact.txt" For Input As #f%
    Do Until EOF(f%)
        Input #f%, strTemp
        If strTemp = "[Summary Title]" Then
            Input #f%, strTemp
            MovieInfo.SumTitle = strTemp
        ElseIf strTemp = "[Summary]" Then
            Input #f%, strTemp
            MovieInfo.Summary = strTemp
        ElseIf strTemp = "[Producer]" Then
            Input #f%, strTemp
            MovieInfo.Producer = strTemp
        ElseIf strTemp = "[Actors]" Then
            ReDim MovieInfo.Actors(9)
            For i = 0 To 5
                Input #f%, MovieInfo.Actors(i)
            Next i
        End If
    Loop
    Close #f%
With txtActors
    .Font.Size = 10
    .FontName = "arial"
    .FontBold = True
    For i = 0 To 5
        .Text = .Text & MovieInfo.Actors(i) & vbCrLf
    Next
End With

With txtSummary
    .Font.Size = 10
    .FontName = "arial"
    .Text = MovieInfo.SumTitle & vbCrLf & vbCrLf
    .Text = .Text & MovieInfo.Summary
End With

With txtAdditional
    .Font.Size = 10
    .FontName = "arial"
    .Text = MovieInfo.Producer
End With


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
