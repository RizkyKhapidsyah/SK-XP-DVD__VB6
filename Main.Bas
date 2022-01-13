Attribute VB_Name = "Main"
Option Explicit
Public sPref As String, f%, Cfg$, my$, up$
Public i As Integer, bMuteValue As Boolean
Public TitleID(2) As String, TitleName(2) As String, TitleCaption As String
Public strTempSpeed As String, strTempFS As String, strTempHM As String
Public strTempHS As String, strTempHT As String, strFile As String
Public iTitleMenu As Integer, iRootMenu As Integer, intSpeed As Integer
'Declaration of User Preferences type
Type Preferences
    intSpeed As Integer
    bFullScreen As Boolean
    bHideMenu As Boolean
    bHideStatus As Boolean
    bHideTool As Boolean
End Type
'Declaration of Movie Info type
Type MI
    SumTitle As String
    Summary As String
    Producer As String
    Actors() As String
End Type
'Setting the globals for the types
Global UserPref As Preferences
Global MovieInfo As MI
'used for shelling out to the default web browser
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
'Used for reading from and writing to the ini file
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

Public Sub GetTxt()
'This sub gets the DVD ID and DVD Title Name from a text file.
'The DVD ID is a property - DVDUniqueID - that identifies every
'DVD title.  With this you can individualize any DVD title complete
'with individualized graphics and movie info.
Dim strPath As String, temp As String

f% = FreeFile
strPath = App.Path & "\"
strFile = "id.txt"

Open strPath & strFile For Input As #f%
Line Input #f%, TitleID(0)
Line Input #f%, TitleName(0)
Close #f%

End Sub

Public Sub GetPref()
Cfg$ = App.Path & "\config.ini"
up$ = "User Preferences"

'Read in Pref data
strTempSpeed = GetFromIni(up$, "Playback Speed", Cfg$)
strTempFS = GetFromIni(up$, "Full Screen", Cfg$)
strTempHM = GetFromIni(up$, "Hide Menu", Cfg$)
strTempHS = GetFromIni(up$, "Hide Status", Cfg$)
strTempHT = GetFromIni(up$, "Hide Toolbar", Cfg$)

'Convert the Pref data
With UserPref
    .intSpeed = CInt(strTempSpeed)
    .bFullScreen = CBool(strTempFS)
    .bHideMenu = CBool(strTempHM)
    .bHideStatus = CBool(strTempHS)
    .bHideTool = CBool(strTempHT)
End With

'Set the preferences
If UserPref.bHideMenu = True Then
    frmPref.cbHideMenu.Value = 1
    With frmXP
        .mnuEdit.Visible = False
        .mnuHelp.Visible = False
        .mnuControls.Visible = False
    End With
ElseIf UserPref.bHideMenu = False Then
    frmPref.cbHideMenu.Value = 0
    With frmXP
        .mnuEdit.Visible = True
        .mnuHelp.Visible = True
        .mnuControls.Visible = True
    End With
End If

If UserPref.bHideStatus = True Then
    frmXP.StatusBar1.Visible = False
    frmPref.cbHideStatus.Value = 1
ElseIf UserPref.bHideStatus = False Then
    frmXP.StatusBar1.Visible = True
    frmPref.cbHideStatus.Value = 0
End If

If UserPref.bHideTool = True Then
    frmXP.Toolbar1.Visible = False
    frmPref.cbHideTool.Value = 1
ElseIf UserPref.bHideTool = False Then
    frmXP.Toolbar1.Visible = True
    frmPref.cbHideTool.Value = 0
End If

End Sub

Public Sub GetDVDInfo()
Dim msg As String, Style As String, Title As String, Response As String, MyString As String

'Find what DVD is in the drive -- This needs to be in a loop
GetTxt
If frmXP.ctlDVD.DVDUniqueID = TitleID(0) Then
    TitleCaption = "You are watching: " & TitleName(0)
Else
    TitleCaption = "I don't recognize this DVD!"
    msg = "This DVD Title is not in the Database!"
    msg = msg & vbCrLf & vbCrLf
    msg = msg & "Would you like to add it to the database?"
    Style = vbYesNo + vbCritical + vbDefaultButton1
    Title = "DVD Title not in Database"

    ' Display message.
    Response = MsgBox(msg, Style, Title)
    If Response = vbYes Then
       MyString = "Yes"
       MsgBox MyString & Space(1) & ".., " & vbCrLf & "Place code here for the database!"
    Else
       MyString = "No"
       End
    End If
End If


End Sub
