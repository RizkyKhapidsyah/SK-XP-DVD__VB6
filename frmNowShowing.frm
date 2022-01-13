VERSION 5.00
Begin VB.Form frmNowShowing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOW SHOWING"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   ControlBox      =   0   'False
   Icon            =   "frmNowShowing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   6240
   End
   Begin VB.Image imgNowShow 
      Height          =   6735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "frmNowShowing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Place the imgNowShow code here to display the image of what is the
'current DVD Title
Me.imgNowShow.Picture = LoadPicture(App.Path & "\graphics\" & TitleID(0) & ".jpg")
End Sub

Private Sub Timer1_Timer()
'    frmXP.Show
    Unload Me
End Sub
