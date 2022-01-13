Attribute VB_Name = "Status"
Option Explicit
'This module houses some of the different states that the player
'would be in, and sets the correct settings.

Public Sub RunningState()
With frmXP
        .mnuPlay.Enabled = False
        .mnuPause.Enabled = True
        .mnuFastForward.Enabled = True
        .mnuRewind.Enabled = True
        .mnuNext.Enabled = True
        .mnuPrevious.Enabled = True
        .mnuStop.Enabled = True
        .mnuEject.Enabled = False
    End With
End Sub
Public Sub PausedState()
With frmXP
        .mnuPlay.Enabled = True
        .mnuPause.Enabled = False
        .mnuFastForward.Enabled = False
        .mnuRewind.Enabled = False
        .mnuNext.Enabled = False
        .mnuPrevious.Enabled = False
        .mnuStop.Enabled = False
        .mnuEject.Enabled = False
    End With
End Sub
Public Sub PlayOnly()
With frmXP
    .mnuPlay.Enabled = True
        With .Toolbar1
            For i = 7 To 10
                .Buttons(i).Enabled = False
            Next i
            .Buttons(11).Enabled = True
            For i = 12 To 14
                .Buttons(i).Enabled = False
            Next i
            .Buttons(16).Enabled = False
            For i = 18 To 26
                .Buttons(i).Enabled = True
            Next i
        End With
End With
End Sub
Public Sub ContPlay()
Dim i As Integer

With frmXP.ctlDVD
    .Play
    With frmXP.Toolbar1
        For i = 8 To 10
            .Buttons(i).Enabled = True
        Next i
        .Buttons(11).Enabled = False
        For i = 12 To 15
            .Buttons(i).Enabled = True
        Next i
        .Buttons(16).Enabled = False
        For i = 17 To 26
            .Buttons(i).Enabled = True
        Next i
    End With
End With
End Sub
Public Sub Pause()
frmXP.ctlDVD.Pause
    With frmXP.Toolbar1
        For i = 3 To 10
            .Buttons(i).Enabled = False
        Next i
       .Buttons(11).Enabled = True
        For i = 12 To 14
            .Buttons(i).Enabled = False
        Next i
    End With
End Sub
Public Sub Stopped()
With frmXP
    .ctlDVD.Stop
    .ctlDVD.Visible = False
    .mnuPlay.Enabled = True
    .Toolbar1.Buttons(11).Enabled = True
    .mnuPause.Enabled = False
    .Toolbar1.Buttons(12).Enabled = False
    .mnuFastForward.Enabled = False
    .Toolbar1.Buttons(13).Enabled = False
    .mnuRewind.Enabled = False
    .Toolbar1.Buttons(9).Enabled = False
    .mnuNext.Enabled = False
    .Toolbar1.Buttons(14).Enabled = False
    .mnuPrevious.Enabled = False
    .Toolbar1.Buttons(8).Enabled = False
    .mnuStop.Enabled = False
    .Toolbar1.Buttons(10).Enabled = False
    .mnuEject.Enabled = True
    .Toolbar1.Buttons(16).Enabled = True
    .Image1.Visible = True
End With
End Sub
Public Sub Play()

With frmXP
    .ctlDVD.Visible = True
    .ctlDVD.Play
    With .Toolbar1
        For i = 1 To 10
            .Buttons(i).Enabled = True
        Next i
        .Buttons(11).Enabled = False
        For i = 12 To 26
            .Buttons(i).Enabled = True
        Next i
    End With
End With

End Sub
