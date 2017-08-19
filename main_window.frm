
'==============================================================================================
' Main Form
'
' Author: Ben Pogrund
'
'==============================================================================================

' --- Dev notes --- '
' Integer = %
' Long = &
' Single = !
' Double = #
' Currency = @
' String = $
' ----------------- '

Option Explicit

' --- UI Variables --- '
Private bShowConfig As Boolean
Private bNewTabOffset As Boolean
Public bShftKeyDown As Boolean
Private Const MENU_HEIGHT = 300


Private Sub ctrl_Menu_Conn_Click()
  Connect CurrentProfIndx
End Sub

Private Sub ctrl_Menu_Disc_Click()
  Disconnect CurrentProfIndx
End Sub

'--- Loading ----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Private Sub Form_Load()
  'Init Config Controls
  '  Client Images
  icmbConfClient.ComboItems.Add , , "Warcraft 2", imgListClients.ListImages("W2BN").Index
  icmbConfClient.ComboItems.Add , , "Starcraft", imgListClients.ListImages("STAR").Index
  icmbConfClient.ComboItems.Add , , "Starcraft BW", imgListClients.ListImages("SEXP").Index
  icmbConfClient.ComboItems.Add , , "Diablo", imgListClients.ListImages("DRTL").Index
  icmbConfClient.ComboItems.Add , , "Diablo II", imgListClients.ListImages("D2DV").Index
  icmbConfClient.ComboItems.Add , , "Diablo II LOD", imgListClients.ListImages("D2XP").Index
  icmbConfClient.ComboItems.Add , , "Warcraft 3", imgListClients.ListImages("WAR3").Index
  icmbConfClient.ComboItems.Add , , "Warcraft 3XP", imgListClients.ListImages("W3XP").Index
  icmbConfClient.ComboItems.Add , , "Starcraft Japan", imgListClients.ListImages("JSTR").Index
  
  icmbConfClient.ComboItems.Item(1).Selected = True
  
  '  Role Images
  icmbConfRole.ComboItems.Add , , , 1
  icmbConfRole.ComboItems.Add , , , 2
  icmbConfRole.ComboItems.Add , , , 3
  
  icmbConfRole.ComboItems.Item(1).Selected = True
  
  ' Channel
  lvChannel(MAIN_TAB).ColumnHeaders.Add , , "main", 2800, lvwColumnLeft
  lblChannel(MAIN_TAB).Caption = "Channel"
  
  ' Input text
  InputText.SelColor = fcGray
  
  ' Editor
  EditorFrame.Left = 0
  EditorFrame.Visible = False
  
  Load EditorTab(1)
  EditorTab(1).Left = EditorTab(1).Left + 2505
  EditorTab(1).Visible = True
  EditorTab(1).Caption = "CDKeys.txt  "
  Load EditorBox(1)
  EditorBox(1).Left = EditorBox(1).Left + 2505
  EditorBox(1).Visible = True
  EditorBox(1).ZOrder BRING_TO_FRONT
  Load rtbEditor(1)
  rtbEditor(1).Visible = True
  
  Load EditorTab(2)
  EditorTab(2).Left = EditorTab(2).Left + 2505 * 2
  EditorTab(2).Visible = True
  EditorTab(2).Caption = "Servers.txt "
  Load EditorBox(2)
  EditorBox(2).Left = EditorBox(2).Left + 2505 * 2
  EditorBox(2).Visible = True
  EditorBox(2).ZOrder BRING_TO_FRONT
  Load rtbEditor(2)
  rtbEditor(2).Visible = True
  
  VersionMacro = "- Version " & App.Major & "." & App.Minor & "." & App.Revision & " -"
  ReDim Preserve pbuffer(0)
  Form_PostLoad
  
End Sub

Private Sub Form_PostLoad()
  'On Error GoTo Skip

  mainForm.Show
  
  ' Profile Loading
  bNewTabOffset = False
  CurrentProfIndx = MAIN_TAB
  CurrentEdtrIndx = 0
  LoadProfiles App.Path & "\profiles\"
  
skip:
  
  ' Window Init
  Me.Height = 13440         'Should load from config
  Me.Width = 16110          'Should load from config
  
  ' UI Init
  bShftKeyDown = False
  bMassConnect = False
  
  If ProfileCount = 0 Then
    ' New Profile
    ProfileTabs(MAIN_TAB).Caption = "New *       "
    ProfileCount = 1
    ReDim ProfileStructs(0)
    
    ' Init some necessary defaults
    With ProfileStructs(ProfileCount - 1)
      If .KeyInUseMsg = vbNullString Then .KeyInUseMsg = "Flavor Bot"
      If .Home = vbNullString Then .Home = "The Flavor Hideout"
      If .Trigger = vbNullString Then .Trigger = "."
    End With
  
    ShowConfig (True)
    AddChat MAIN_TAB, " * No profiles found." & vbCrLf & vbCrLf, fcYellow, , NO_LINE
  Else
    ' For loaded profiles
    ShowConfig (False)
    'AddChat MAIN_TAB, " > " & ProfileCount & " Profiles loaded" & vbCrLf, vbWhite
  End If
  
  ' Load Access list
  ReDim Preserve AccessList(0)
  AddChat MAIN_TAB, " > Loading Access List", vbWhite
  LoadAccessList
  AddChat MAIN_TAB, vbCrLf, vbWhite
  
  ' Welcome message
  GreetMsg MAIN_TAB
  
  ' Focus the main tab
  FocusTab MAIN_TAB
  
  ' Finish
  bNewTabOffset = True  ' Never took the time to figure out why this index is off
                        ' Its only off during this function call
  
  BringChannelToFront
    
  mainForm.InputText.SetFocus
  ReDim Preserve PreviousInputs(1)
End Sub

Private Sub LoadProfiles(strPath As String)
  Dim File As String
  Dim k As Integer
  
  ProfileCount = 0      ' Iniitalize global count
  k = 0                 ' Initialize iterator
  
  AddChat MAIN_TAB, " > ", vbWhite
  AddChat MAIN_TAB, File & "Loading Profiles", vbWhite, , NO_LINE
  File = Dir$(strPath & "*.ini")
  Do While Len(File)
    LoadProfile File, k
    AddChat MAIN_TAB, "    " & File & " loaded", fcBlue
    
    If k <> 0 Then NewTab
    ProfileTabs(k).Caption = TabFormat(Left(File, Len(File) - 4))
    
    File = Dir$
    k = k + 1
  Loop
  AddChat MAIN_TAB, "", vbWhite
End Sub

Private Sub LoadProfile(File As String, Index As Integer)
  Dim Section As String
  
  If Index >= ProfileCount Then
    ProfileCount = ProfileCount + 1
    ReDim Preserve ProfileStructs(ProfileCount - 1)
  End If
  
  Section = "Main"
  
  ProfileStructs(Index).Username = ReadINI(App.Path & "\profiles\" & File, Section, "Username")
  ProfileStructs(Index).Password = ReadINI(App.Path & "\profiles\" & File, Section, "Password")
  ProfileStructs(Index).Server = ReadINI(App.Path & "\profiles\" & File, Section, "Server")
  ProfileStructs(Index).Client = ReadINI(App.Path & "\profiles\" & File, Section, "Client")
  
  ProfileStructs(Index).CDKey = ReadINI(App.Path & "\profiles\" & File, Section, "CDKey")
  ProfileStructs(Index).XPKey = ReadINI(App.Path & "\profiles\" & File, Section, "XPKey")
  ProfileStructs(Index).BNLSServ = ReadINI(App.Path & "\profiles\" & File, Section, "BNLS")
  ProfileStructs(Index).bBNLS = BoolINI(App.Path & "\profiles\" & File, Section, "UseBNLS")
  ProfileStructs(Index).Role = ReadINI(App.Path & "\profiles\" & File, Section, "Role")
  
  Section = "Bot"
  
  ProfileStructs(Index).Trigger = ReadINI(App.Path & "\profiles\" & File, Section, "Trigger")
  ProfileStructs(Index).Home = ReadINI(App.Path & "\profiles\" & File, Section, "Home")
  ProfileStructs(Index).GreetOn = BoolINI(App.Path & "\profiles\" & File, Section, "GreetOn")
  ProfileStructs(Index).GreetMsg = ReadINI(App.Path & "\profiles\" & File, Section, "GreetMsg")
  ProfileStructs(Index).IdleOn = BoolINI(App.Path & "\profiles\" & File, Section, "IdleOn")
  ProfileStructs(Index).IdleMsg = ReadINI(App.Path & "\profiles\" & File, Section, "IdleMsg")
  ProfileStructs(Index).IdleTimer = ReadIntINI(App.Path & "\profiles\" & File, Section, "IdleTimer")
  ProfileStructs(Index).IdleElapsed = 0
  ProfileStructs(Index).KeyInUseMsg = ReadINI(App.Path & "\profiles\" & File, Section, "KeyInUseMsg")
  
  ProfileStructs(Index).bNotifySet = False ' Initialize blinky graphics as off
  
End Sub

Private Sub LoadConf(Index As Integer)
  If ProfileCount = 0 Then Exit Sub
  
  txtConfUsername.Text = ProfileStructs(Index).Username
  txtConfPassword.Text = ProfileStructs(Index).Password
  txtConfBNLS.Text = ProfileStructs(Index).BNLSServ

  Dim i As Long
  i = 1
  While i <= icmbConfClient.ComboItems.Count
    If icmbConfClient.ComboItems.Item(i).Text = ProfileStructs(Index).Client Then
      icmbConfClient.ComboItems.Item(i).Selected = True
    End If
    i = i + 1
  Wend
  
  icmbConfRole.ComboItems.Item(GetRoleIndex(ProfileStructs(Index).Role)).Selected = True
  
  'Load ComboBox Lists
  LoadConfServers
  LoadConfCDKeys
End Sub

Private Sub LoadBotConf(Index As Integer)
  If ProfileCount = 0 Then Exit Sub
  
  txtBotTrigger.Text = ProfileStructs(Index).Trigger
  txtBotHome.Text = ProfileStructs(Index).Home
  txtBotGreetMsg.Text = ProfileStructs(Index).GreetMsg
  ChkConfGreet.Value = Abs(CInt(ProfileStructs(Index).GreetOn))
  txtBotIdleMsg.Text = ProfileStructs(Index).IdleMsg
  chkConfIdle.Value = Abs(CInt(ProfileStructs(Index).IdleOn))
  txtBotKeySignature.Text = ProfileStructs(Index).KeyInUseMsg
  
  txtBotIdle_H.Text = Int(ProfileStructs(Index).IdleTimer / 3600)
  txtBotIdle_M.Text = Int(ProfileStructs(Index).IdleTimer / 60) Mod 60
  txtBotIdle_S.Text = ProfileStructs(Index).IdleTimer Mod 60
  
  If txtBotIdle_H.Text = "0" Then
    txtBotIdle_H.Text = ""
  Else
    If Len(txtBotIdle_M.Text) = 1 Then txtBotIdle_M.Text = "0" & txtBotIdle_M.Text
  End If
  
  If txtBotIdle_M.Text = "0" Then
    txtBotIdle_M.Text = ""
  Else
    If Len(txtBotIdle_S.Text) = 1 Then txtBotIdle_S.Text = "0" & txtBotIdle_S.Text
  End If
End Sub

Public Sub LoadConfServers()
  Dim FileNumb As Integer
  Dim FileLine As String
  FileNumb = FreeFile
  
  ' Load Server List
  cmbConfServer.Clear
  Open App.Path & "\txt\Servers.txt" For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Line Input #FileNumb, FileLine$
    cmbConfServer.AddItem FileLine, cmbConfServer.ListCount
  Loop
  Close #FileNumb
  cmbConfServer.Text = ProfileStructs(CurrentProfIndx).Server
  
End Sub

Public Sub LoadConfCDKeys()
  Dim KeyCount As Integer
  KeyCount = GetKeyCount(icmbConfClient.Text)
  
  ' Disable Unused Fields
  If KeyCount < 2 Then cmbConfXPKey.Enabled = False Else cmbConfXPKey.Enabled = True
  If KeyCount < 1 Then cmbConfCDKey.Enabled = False Else cmbConfCDKey.Enabled = True
  
  ' Load CDKeys
  Dim FileNumb As Integer
  Dim FileLine As String
  FileNumb = FreeFile
  
  cmbConfCDKey.Clear
  Open App.Path & "\txt\CDKeys.txt" For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Line Input #FileNumb, FileLine$
    If (KeyClientMatch(FileLine, icmbConfClient.Text)) Then
      cmbConfCDKey.AddItem FileLine, cmbConfCDKey.ListCount
    End If
  Loop
  Close #FileNumb
  cmbConfCDKey.Text = ProfileStructs(CurrentProfIndx).CDKey
  
  cmbConfXPKey.Clear
  Open App.Path & "\txt\CDKeys.txt" For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Line Input #FileNumb, FileLine$
    If (KeyClientMatch(FileLine, icmbConfClient.Text)) Then
      cmbConfXPKey.AddItem FileLine, cmbConfXPKey.ListCount
    End If
  Loop
  Close #FileNumb
  cmbConfXPKey.Text = ProfileStructs(CurrentProfIndx).XPKey
  
  If KeyCount < 2 Then cmbConfXPKey.Text = "N/A"
  If KeyCount < 1 Then cmbConfCDKey.Text = "N/A"
  
End Sub

Public Function KeyClientMatch(CDKey As String, Client As String) As Boolean
  KeyClientMatch = False
  Select Case Client
    Case "Starcraft", "Starcraft BW":
      If Len(CDKey) = 13 Then KeyClientMatch = True
    Case "Warcraft 2", "Diablo II", "Diablo II LOD":
      If Len(CDKey) = 16 Then KeyClientMatch = True
    Case "Warcraft 3", "Warcraft 3XP":
      If Len(CDKey) = 26 Then KeyClientMatch = True
  End Select
  
End Function

Private Function GetRoleIndex(str As String) As Integer
  Select Case str
    Case "Chat":      GetRoleIndex = 1
    Case "Trivia":    GetRoleIndex = 2
    Case "Tool":      GetRoleIndex = 3
    Case Else
      GetRoleIndex = 1
  End Select
  
End Function

Private Function GetRole(Index As Integer) As String
  Select Case Index
    Case 1:           GetRole = "Chat"
    Case 2:           GetRole = "Trivia"
    Case 3:           GetRole = "Tool"
  End Select
  
End Function

'--- Resizing ---------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Private Sub Form_Resize()
  If Height >= 8805 Then
    Resize_Height
  End If
  
  If Width >= 16110 Then
    Resize_Width
  Else
    configFrame.Left = 840
  End If
  
End Sub

Private Sub Resize_Width()
  Dim i As Integer
  
'Horizontal Properties
' Main Chat Window
  i = 0
  While i < rtbChat.Count
    rtbChat(i).Width = mainForm.Width - 4230
    i = i + 1
  Wend

' Channel
  i = 0
  While i < lvChannel.Count
    lvChannel(i).Left = mainForm.Width - 3885
    i = i + 1
  Wend
  
  channelBack.Left = mainForm.Width - 4005 - 120
  
' Input
  InputTextFrame.Width = mainForm.Width - 4305
  InputText.Width = mainForm.Width - 4305 - (12 * 15)

' Config Frame
  If bShowConfig Then configFrame.Left = (mainForm.Width - 14445) / 2
  If bShowConfig Then configBotFrame.Left = (mainForm.Width - 14445) / 2
  
' Editor
  EditorFrame.Width = mainForm.Width - 4080
  EditorCloseAll.Left = mainForm.Width - 7920
  EditorSaveAll.Left = mainForm.Width - 7920 + 1815

  i = 0
  While i < rtbEditor.Count
    rtbEditor(i).Width = mainForm.Width - 4230
    i = i + 1
  Wend
  
End Sub

Private Sub Resize_Height()
  Dim offset As Long
  Dim i As Integer
  offset = 0
  
'Vertical Properties
' Main Chat Window
  If bShowConfig Then offset = 3315
  i = 0
  While i < rtbChat.Count
    rtbChat(i).Height = mainForm.Height - 1875 - MENU_HEIGHT - offset
    i = i + 1
  Wend
  
' Channel
  i = 0
  While i < lvChannel.Count
    lvChannel(i).Height = mainForm.Height - 1875 - MENU_HEIGHT
    i = i + 1
  Wend
  KeyEditorBtn.Top = mainForm.Height - 1485
  channelBack.Height = mainForm.Height
  channelFrame.Height = mainForm.Height
  
' Input
  InputTextFrame.Top = mainForm.Height - 1170 - MENU_HEIGHT
  InputText.Top = mainForm.Height - 1170 - MENU_HEIGHT + (6 * 15)
  
' Config Frame
  If bShowConfig Then configFrame.Top = mainForm.Height - 4725
  If bShowConfig Then configBotFrame.Top = mainForm.Height - 4725
  
' Editor
  EditorFrame.Height = mainForm.Height - 1560 - offset
  EditorCloseAll.Top = EditorFrame.Height - 735 - (60 * 15)
  EditorSaveAll.Top = EditorFrame.Height - 735 - (60 * 15)
  
  i = 0
  While i < rtbEditor.Count
    rtbEditor(i).Height = EditorFrame.Height - 1440 - (60 * 15)
    i = i + 1
  Wend
  
End Sub

'--- Menu ---------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Private Sub ctrl_Menu_ConnectAll_Click()
  MassConnectTimer = True
  iMassConnectIndex = -1
  bMassConnect = True
  LoadConnectionIcons
End Sub

Private Sub ctrl_Config_Click()
  ShowConfig (True)
End Sub

Private Sub ctrl_Bot_Click()
  ShowBotConfig (True)
End Sub

Private Sub ctrl_Menu_Exit_Click()
  End
End Sub

Private Sub ctrl_Menu_Access_Click()
  LoadEditor 0, "Access.txt"
End Sub

Private Sub ctrl_Menu_CDKeys_Click()
  LoadEditor 1, "CDkeys.txt"
End Sub

Private Sub ctrl_Menu_Servers_Click()
  LoadEditor 2, "Servers.txt"
End Sub

'--- Controls ---------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Shift on tabs
  If bShftKeyDown = False And Shift = 1 And InputText.Text = "" Then
    bShftKeyDown = True
    LoadConnectionIcons
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
' UnShift on tabs
  If KeyCode = 16 Then
    bShftKeyDown = False
    ClearConnectionIcons
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub icmbConfClient_Click()
  LoadConfCDKeys
End Sub

Private Sub IdleTimer_Timer()
  Dim i As Integer
  While i < UBound(ProfileStructs)
    With ProfileStructs(i)
      If .iConnected = 2 And .IdleOn Then
        .IdleElapsed = .IdleElapsed + 1
        
        If .IdleElapsed >= .IdleTimer Then
          Send0x0E i, .IdleMsg
        End If
      Else
        .IdleElapsed = 0
      End If
    End With
    
    i = i + 1
  Wend
End Sub

Private Sub InputText_KeyDown(KeyCode As Integer, Shift As Integer)
' Messanger
  ' Recall Previous Input feature
  If KeyCode = 38 And PrevInputIndex <> 0 Then
    PrevInputIndex = PrevInputIndex - 1
    InputText.Text = PreviousInputs(PrevInputIndex)
    InputText.SelStart = 0
    InputText.SelLength = 99999
    InputText.SelColor = fcGray
    InputText.SelStart = 99999
    Exit Sub
  End If
  
  ' Ignore normal typing, detects enter for sending
  If KeyCode <> 13 Or InputText.Text = "" Then Exit Sub
  If ProfileStructs(CurrentProfIndx).iConnected <> 2 Then
    ' Offline commands? ...
    InputText.Text = ""
    Exit Sub
  End If
  
  PreviousInputs(UBound(PreviousInputs)) = InputText.Text
  ReDim Preserve PreviousInputs(UBound(PreviousInputs) + 1)
  PrevInputIndex = UBound(PreviousInputs)
  
  Dim Message As String
  Message = InputText.Text
  InputText.Text = vbNullString
  InputText.SelColor = fcGray
  
  ' Online commands
  If Left(Message, 1) = "/" Then
    Dim Length As Long
    Dim Command As String
    Dim Args As String
    Length = InStr(1, Message, " ") - 2
    If Length <= 0 Then Length = Len(Message) - 1
    Command = LCase(Mid(Message, 2, Length))
    Args = Mid(Message, Len(Command) + 3)
    
    ' Client Only Commands
    Dim Commands() As String
    Commands = Split(Message, " ")
    Select Case Command
      Case "profile":
        If UBound(Commands) = 1 Then
          Send0x26 CurrentProfIndx, Commands(1)
        Else
          Send0x26 CurrentProfIndx, ProfileStructs(CurrentProfIndx).Username
        End If
        Exit Sub
        
      Case "permissions":
        PrintPermissions CurrentProfIndx
        Exit Sub
        
      Case "spooflist", "spooftable":
        PrintSpoofs CurrentProfIndx
        Exit Sub
      
    End Select
    
    ' Client/Chat commands
    If ParseChatCommand(CurrentProfIndx, Command, Args) = True Then Exit Sub
    
  End If
  ' Normal chat message
  Send0x0E CurrentProfIndx, Message
End Sub

Private Sub chkConfSave_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  chkConfSave.Value = 0
  Dim File As String
  Dim FilePath As String
  Dim Section As String
  Dim CDKey As String, XPKey As String
  
  If cmbConfCDKey.Text <> "N/A" Then CDKey = cmbConfCDKey.Text Else CDKey = ""
  If cmbConfXPKey.Text <> "N/A" Then XPKey = cmbConfXPKey.Text Else XPKey = ""
  
  Section = "Main"
  
  File = "Profile" & CurrentProfIndx + 1 & ".ini"
  FilePath = App.Path & "\profiles\" & File
  
' Main config
  WriteINI Section, "Username", txtConfUsername.Text, FilePath
  WriteINI Section, "Password", txtConfPassword.Text, FilePath
  WriteINI Section, "Server", cmbConfServer.Text, FilePath
  WriteINI Section, "Client", icmbConfClient.Text, FilePath
  
  WriteINI Section, "CDKey", CDKey, FilePath
  WriteINI Section, "XPKey", XPKey, FilePath
  WriteINI Section, "BNLS", txtConfBNLS.Text, FilePath
  WriteINI Section, "UseBNLS", ChkConfBNLS.Value, FilePath
  WriteINI Section, "Role", GetRole(icmbConfRole.SelectedItem.Index), FilePath
  
  LoadProfile File, CurrentProfIndx
  ProfileTabs(CurrentProfIndx).Caption = TabFormat(Left(File, Len(File) - 4))
  ShowConfig (False)
  
  AddChat CurrentProfIndx, " > ", vbWhite
  AddChat CurrentProfIndx, "Config saved.", fcBlue, , NO_LINE
  
  
' Check if settings valid, then print system info

' ...
  
End Sub


Private Sub chkBotSave_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  chkBotSave.Value = 0
  Dim File As String
  Dim FilePath As String
  Dim Section As String
  Dim IdleTimeInSeconds As Integer
  
  Section = "Bot"
  
  File = "Profile" & CurrentProfIndx + 1 & ".ini"
  FilePath = App.Path & "\profiles\" & File
  
' Bot config
  WriteINI Section, "Trigger", txtBotTrigger.Text, FilePath
  WriteINI Section, "Home", txtBotHome.Text, FilePath
  WriteINI Section, "GreetOn", ChkConfGreet.Value, FilePath
  WriteINI Section, "GreetMsg", txtBotGreetMsg.Text, FilePath
  WriteINI Section, "IdleOn", chkConfIdle.Value, FilePath
  WriteINI Section, "IdleMsg", txtBotIdleMsg, FilePath
  WriteINI Section, "KeyInUseMsg", txtBotKeySignature.Text, FilePath
  
  Dim s As Integer, m As Integer, h As Integer
  If txtBotIdle_S.Text <> "" Then s = CInt(txtBotIdle_S.Text) Else s = 0
  If txtBotIdle_M.Text <> "" Then m = CInt(txtBotIdle_M.Text) Else m = 0
  If txtBotIdle_H.Text <> "" Then h = CInt(txtBotIdle_H.Text) Else h = 0
  
  IdleTimeInSeconds = s + m * 60 + h * 60 * 60
  
  WriteINI Section, "IdleTimer", CStr(IdleTimeInSeconds), FilePath
  
  LoadProfile File, CurrentProfIndx
  ProfileTabs(CurrentProfIndx).Caption = TabFormat(Left(File, Len(File) - 4))
  ShowConfig (False)
  
  AddChat CurrentProfIndx, " > ", vbWhite
  AddChat CurrentProfIndx, "Bot settings saved.", fcBlue, , NO_LINE
  
  
' Check if settings valid, then print system info

' ...
  
End Sub



Private Sub EditorTab_Click(Index As Integer)
  Dim Filename As String
  EditorTab(Index).Value = 0
  
  Select Case Index
    Case 0:   Filename = "Access.txt"
    Case 1:   Filename = "CDKeys.txt"
    Case 2:   Filename = "Servers.txt"
  End Select
  
  CurrentEdtrIndx = Index
  CurrentEdtrFile = Filename
  
  If FileStatus(Index) = EDITOR_FILE_CLOSED Then
    LoadEditor Index, Filename
  Else
    rtbEditor(Index).ZOrder BRING_TO_FRONT
  End If
    
  mainForm.EditorCoverObj.ZOrder BRING_TO_FRONT
  mainForm.EditorCoverObj.Left = 150 + (CurrentEdtrIndx * 2505)
  mainForm.rtbEditor(Index).SetFocus
End Sub

Private Sub EditorBox_Click(Index As Integer)
  EditorBox(Index).Value = 0
  
  SaveEditorFile Index
  
End Sub

Private Sub AddTab_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  AddTab.Value = 0
  
  ProfileCount = ProfileCount + 1
  ReDim Preserve ProfileStructs(ProfileCount - 1)
  
  NewTab
  
  CurrentProfIndx = ProfileCount - 1
  ShowConfig (True)
  txtConfUsername.SetFocus          ' Quick conf access
  FocusTab CurrentProfIndx    ' Draws shape around new tab
  BringChannelToFront               ' Keeps new tabs under channel frame
  
End Sub

Private Sub NewTab()
' Creates new tab
  Dim B As Integer
  Dim Index As Integer
  If bNewTabOffset = True Then B = 1 Else B = 0
  Index = ProfileCount - 1
  
  'Slide right
  AddTab.Left = AddTab.Left + 2160
  
  ' Load controls
  Load rtbChat(Index)
  Load lvChannel(Index)
  Load ProfileTabs(Index)
  Load CloseTabButton(Index)
  Load tabFrame(Index)
  Load wsTCP(Index)
  Load lblChannel(Index)
  Load RejoinTimer(Index)
  Load RetryTimer(Index)
  ReDim Preserve pbuffer(Index)
  
  ' Init Advanced settings
  With ProfileStructs(ProfileCount - 1)
    If .KeyInUseMsg = vbNullString Then .KeyInUseMsg = "Flavor Bot"
    If .Home = vbNullString Then .Home = "The Flavor Hideout"
    If .Trigger = vbNullString Then .Trigger = "."
  End With

  ' Adjustments
  With rtbChat(Index)
    rtbChat(Index).Text = ""
    GreetMsg Index
  End With
  
  With lblChannel(Index)
    .ZOrder BRING_TO_FRONT         ' Bring to front
    .Visible = True
  End With
  
  With lvChannel(Index)
    .ZOrder BRING_TO_FRONT          ' Bring to front
    .Visible = True
    .ListItems.Clear
  End With
  
  With ProfileTabs(Index)
    .Left = 120 + (Index) * 2160
    .Top = 240
    .Visible = True
    .Caption = "New *       "
  End With

  With CloseTabButton(Index)
    .Left = 1680 + (Index) * 2160
    .Top = 240
    .Visible = True
    .ZOrder BRING_TO_FRONT             ' Bring to front
  End With
  
  With tabFrame(Index)
    .Visible = True
    .Left = 60 + (Index) * 2160
    .ZOrder BRING_TO_FRONT             ' Bring to front
    FocusTab Index
  End With
  
  Form_Resize               ' Further adjusments
  
End Sub

Private Sub FocusTab(Index As Integer)
  Dim i As Integer
  
' Notification Alert Graphics
  ProfileStructs(Index).bNotifySet = False
  tabFrame(Index).BorderColor = &H0
  
' Swap in tab controls
  i = 0
  While i < lblChannel.Count
    If i <> Index Then lblChannel(i).Visible = False Else lblChannel(i).Visible = True
    i = i + 1
  Wend
  
  i = 0
  While i < lvChannel.Count
    If i <> Index Then lvChannel(i).Visible = False Else lvChannel(i).Visible = True
    i = i + 1
  Wend
  
  i = 0
  While i < rtbChat.Count
    If i <> Index Then rtbChat(i).Visible = False Else rtbChat(i).Visible = True
    i = i + 1
  Wend
  
  CoverObj.Left = 150 + Index * 2160
  InputText.SetFocus
  
End Sub

Private Sub CloseTabButton_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  CloseTabButton(Index).Value = 0
  
  If bShftKeyDown Then
    Select Case ProfileStructs(Index).iConnected
      Case 0    ' Disconnected
        Connect Index
      Case 1    ' Working
        Disconnect Index
      Case 2    ' Connected
        Disconnect Index
    End Select
    
  Else
    ' Close tab
    
    '...
    
  End If
  
  
End Sub

Private Sub InputText_KeyPress(KeyAscii As Integer)
  Const ASC_CTRL_A As Integer = 1
  If KeyAscii = ASC_CTRL_A Then
    InputText.SelStart = 0
    InputText.SelLength = Len(InputText.Text)
  End If
End Sub

Private Sub InputTextFrame_GotFocus()
  InputText.SetFocus
End Sub

Private Sub KeyEditorBtn_Click()
' Open Key Editor
  KeyEditorBtn.Value = 0
  
End Sub

Private Sub lvChannel_DblClick(Index As Integer)
  Dim AccountName As String
  If lvChannel(Index).ListItems.Count = 0 Then Exit Sub
  
  AccountName = lvChannel(Index).SelectedItem.Text

  Send0x26 Index, AccountName
  
End Sub

Private Sub MassConnectTimer_Timer()
  iMassConnectIndex = iMassConnectIndex + 1
  
  If iMassConnectIndex > UBound(ProfileStructs) Then
    MassConnectTimer.Enabled = False
    ClearConnectionIcons
    bMassConnect = False
    Exit Sub
  End If
  
  If ProfileStructs(iMassConnectIndex).iConnected = 0 Then
    'Connect profile
    If ProfileStructs(iMassConnectIndex).Server = vbNullString Or _
       ProfileStructs(iMassConnectIndex).Username = vbNullString Then
         
      AddChat iMassConnectIndex, " > ", vbWhite
      AddChat iMassConnectIndex, "This profile needs to be configured before connecting.", fcOrange, , NO_LINE
      'MassConnectTimer.Enabled = False
      Exit Sub
    End If
       
    Connect iMassConnectIndex
  Else
    'Immediately skip to next
    MassConnectTimer_Timer
  End If
  
End Sub

Private Sub ProfileTabs_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  ProfileTabs(Index).Value = 0
  
  CurrentProfIndx = Index
  LoadConf CurrentProfIndx
  LoadBotConf CurrentProfIndx

  FocusTab Index
  
End Sub

Private Sub confCloseBtn_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  confCloseBtn.Value = 0
  ShowConfig (False)
End Sub

Private Sub confBotCloseBtn_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
' Button reset
  confBotCloseBtn.Value = 0
  ShowBotConfig (False)
End Sub

'--- UI Functions ---------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Public Sub LoadConnectionIcons()
  Dim i As Integer
  
  ' For each tab, change close button to connection icon
  i = 0
  While i < CloseTabButton.Count
    CloseTabButton(i).Caption = ""
    Select Case ProfileStructs(i).iConnected
      Case 0:        CloseTabButton(i).Picture = UI_Icons.ListImages(3).Picture
      Case 1:        CloseTabButton(i).Picture = UI_Icons.ListImages(2).Picture
      Case 2:        CloseTabButton(i).Picture = UI_Icons.ListImages(1).Picture
    End Select
    
    i = i + 1
  Wend
    
End Sub

Public Sub ClearConnectionIcons()
  Dim i As Integer
  
  'Unload connection icons
  i = 0
  While i < CloseTabButton.Count
    CloseTabButton(i).Caption = "X"
    CloseTabButton(i).Picture = Nothing
    i = i + 1
  Wend
    
End Sub

Public Sub AddToChannel(Index As Integer, Username As String, ClientID As String)
On Error GoTo Err

  Dim itm As ListItem
  Dim i As Integer
  Dim ChIndx As Integer
  
  ' For unique channel key (Kind of hacky)
  uniqChID = uniqChID + 1
  
  ' If already present, exit
  While i < lvChannel(Index).ListItems.Count
    If lvChannel(Index).ListItems(i + 1).Text = Username Then Exit Sub
    i = i + 1
  Wend
  
  Select Case ClientID
    Case "MOD", "BLIZ", "SYSOP"
      ChIndx = 1
    Case Else
      ChIndx = lvChannel(Index).ListItems.Count + 1
  End Select
  
  ' Add user
  Set itm = lvChannel(Index).ListItems.Add(ChIndx, ClientID & uniqChID, Username, , imgListChannel.ListImages(ClientID).Index)
  UpdateChannel Index
  Exit Sub
Err:
  Set itm = lvChannel(Index).ListItems.Add(ChIndx, "UNKNOWN" & uniqChID, Username, , imgListChannel.ListImages("UNKNOWN").Index)
  UpdateChannel Index
End Sub

Public Sub RemoveFromCh(Index As Integer, Username As String)
  Dim i As Integer
  
  While i < lvChannel(Index).ListItems.Count
    If lvChannel(Index).ListItems(i + 1).Text = Username Then
      lvChannel(Index).ListItems.Remove (i + 1)
    End If
    i = i + 1
  Wend
  
  UpdateChannel Index
End Sub

Public Function ChLookup(Index As Integer, Username As String) As String
  Dim i As Integer
  Dim k As Integer
  
  k = -1
  i = 1
  While i < lvChannel(Index).ListItems.Count + 1
    If lvChannel(Index).ListItems.Item(i).Text = Username Then
      k = i
    End If
    i = i + 1
  Wend
  
  If k = -1 Then
    ChLookup = "NOTFOUND"
    Exit Function
  End If
  
  ChLookup = lvChannel(Index).ListItems.Item(k).Key
  
End Function

Public Sub ChUpdate(Index As Integer, Username As String, ClientID As String)
  Dim OldClientID As String
  
  OldClientID = Left$(ChLookup(Index, Username), Len(ClientID))
  
  If OldClientID <> ClientID Then
    RemoveFromCh Index, Username
    AddToChannel Index, Username, ClientID
  End If
End Sub

Private Sub ShowConfig(Toggle As Boolean)
  bShowConfig = Toggle            ' Used only during window resizing
  configFrame.Visible = Toggle
  configFrame.Left = -99999        ' Cheap way to hide flickering
  configBotFrame.Visible = False
  
  Form_Resize
  
  LoadConf CurrentProfIndx
  
End Sub

Private Sub ShowBotConfig(Toggle As Boolean)
  bShowConfig = Toggle            ' Used only during window resizing
  configBotFrame.Visible = Toggle
  configFrame.Visible = False
  
  Form_Resize
  
  LoadBotConf CurrentProfIndx
  
End Sub

Private Function BringChannelToFront()
  Dim i As Integer
  
  channelBack.ZOrder BRING_TO_FRONT
  i = 0
  While i < lvChannel.Count
    lvChannel(i).ZOrder BRING_TO_FRONT
    i = i + 1
  Wend
End Function

' ===== UI :: Editor ================================================

Private Sub EditorCloseAll_Click()
  EditorCloseAll.Value = 0
  UnloadEditor
End Sub

Private Sub EditorSaveAll_Click()
  EditorSaveAll.Value = 0
  
  If FileStatus(0) = EDITOR_FILE_EDITED Then SaveEditorFile 0
  If FileStatus(1) = EDITOR_FILE_EDITED Then SaveEditorFile 1
  If FileStatus(2) = EDITOR_FILE_EDITED Then SaveEditorFile 2
  
End Sub


' ===== Helper functions ============================================

Private Sub GreetMsg(Index As Integer)
  Dim NameFit As String
  Dim printLine As String
  Dim FileNumb As Integer
  Dim Buffer As String
  
  Buffer = ""
  
  ' Print Flavor Bot Signature
  FileNumb = FreeFile
  Open App.Path & "\txt\welcome msg.txt" For Input As #FileNumb
  AddChat Index, "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯", &H8E1DFF, rtfCenter
  Do While Not EOF(FileNumb)          ' Loop until end of file
    Line Input #FileNumb, printLine$  ' Read line into variable
    AddChat Index, printLine & Buffer, &HB469FF, rtfCenter
    Buffer = Buffer + Chr(144)        ' Tricks rtfalignment with space-like character
  Loop
  AddChat Index, "", &HB469FF, rtfCenter
  AddChat Index, "_____________________________________________________________________", &H8E1DFF, rtfCenter
  AddChat Index, "", &HB469FF, rtfCenter
  Close #FileNumb
  
  AddChat Index, VersionMacro & vbCrLf & vbCrLf & vbCrLf, vbWhite, rtfCenter
  
  ' Print profile info
  If ProfileStructs(Index).Username <> "" Then
    
    AddChat Index, " > [System Ready] ", vbWhite
    AddChat Index, " > ", vbWhite
    
    NameFit = ProfileStructs(Index).Username
    While Len(NameFit) < 15
      NameFit = NameFit & " "
    Wend
    
    AddChat Index, " " & NameFit & "   " & ProfileStructs(Index).Server, fcGray, , NO_LINE
    AddChat Index, " > ", vbWhite
    
    NameFit = ProfileStructs(Index).Client
    While Len(NameFit) < 15
      NameFit = NameFit & " "
    Wend
    
    AddChat Index, " " & NameFit & "   " & ProfileStructs(Index).Home, fcGray, , NO_LINE
    AddChat Index, " > ", vbWhite
  Else
    AddChat Index, " > [New profile] ", vbWhite
    AddChat Index, " > ", vbWhite
  End If
  
End Sub

Private Sub RejoinTimer_Timer(Index As Integer)
  Send0x0C Index, ProfileStructs(Index).LastChannel
  RejoinTimer(Index).Enabled = False
End Sub

Private Sub RetryTimer_Timer(Index As Integer)
  AddChat Index, " > ", vbWhite
  AddChat Index, "Connection seems slow, retrying.", fcYellow, , NO_LINE
  Disconnect Index
  Connect Index
End Sub

Private Sub rtbEditor_Change(Index As Integer)
  
  If FileStatus(Index) = 0 Then Exit Sub
  SetFileStatus Index, EDITOR_FILE_EDITED
  
  '-----------------------------------------------------------------
  ' Update syntax highlighting : defunct due to erasing undo history
  '-----------------------------------------------------------------
  
  'Dim SelStartOld As Long
  'SelStartOld = rtbEditor(Index).SelStart
  
  'SplitString = Split(rtbEditor(Index).Text, vbCrLf)
  'If (UBound(SplitString) = -1) Then Exit Sub
  
  'Dim SelStart As Integer, SelLength As Integer
  'SelStart = 1
  'SelLength = InStr(Mid$(rtbEditor(Index).Text, SelStart), vbCrLf)
  'While SelLength <> 0
  '  EdtrHighlight Index, SelStart, SelLength
  '  SelStart = SelStart + SelLength
  '  If InStr(Mid$(rtbEditor(Index).Text, SelStart), vbCrLf) <> 0 Then
  '    SelLength = InStr(Mid$(rtbEditor(Index).Text, SelStart), vbCrLf)
  '  Else
  '    SelLength = 0
  '  End If
  'Wend
  
  'rtbEditor(Index).SelStart = SelStartOld
  
End Sub

Private Sub rtbEditor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If rtbEditor(Index).SelColor <> fcGray Then rtbEditor(Index).SelColor = fcGray
End Sub

Private Sub UIFlash_Timer()
  Dim i As Integer
  i = 0
  bFlashOrange = Not bFlashOrange
  While i < tabFrame.Count
    If ProfileStructs(i).bNotifySet = True Then
      If bFlashOrange = True Then
        tabFrame(i).BorderColor = fcOrange
      Else
        tabFrame(i).BorderColor = fcYellow
      End If
    End If
    i = i + 1
  Wend
End Sub

'--- Winsock ---------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------

Private Sub wsTCP_Connect(Index As Integer)
  AddChat Index, " > ", vbWhite
  AddChat Index, " Connected!", fcGreen, , NO_LINE
  
  Send0x50 Index
End Sub

Private Sub wsTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  'On Error GoTo err
  
  Dim PacketID As Byte
  Dim Data As String
  Dim strTemp As String
  Dim lngLen As Long
  Dim i As Integer

  With pbuffer(Index)
    wsTCP(Index).GetData strTemp
    .strBuffer = .strBuffer & strTemp
    
    While Len(.strBuffer) > 4
      strTemp = StrReverse$(Mid$(.strBuffer, 3, 2))
      lngLen = Asc(Right$(strTemp, 1))
      i = Asc(Left$(strTemp, 1))
      lngLen = lngLen + i * 255
      If Len(.strBuffer) < lngLen Then Exit Sub
      
      Data = Left$(.strBuffer, lngLen)
      PacketID = Asc(Mid$(Data, 2, 1))
      ParseBNCS Index, Data, PacketID
      
      .strBuffer = Mid$(.strBuffer, lngLen + 1 + i)
    Wend
  End With
  Exit Sub
Err:
  AddChat Index, "Error: " & Err.Number & " at WS_DataArrival() : " & Err.Description, fcRed
  pbuffer(Index).strBuffer = ""
End Sub

Private Sub wsTCP_Close(Index As Integer)

  AddChat Index, " > ", vbWhite
  AddChat Index, "Disconnected", fcRed, , NO_LINE
      
  SetConnectStatus Index, 0 'Offline mode
  
  lvChannel(Index).ListItems.Clear
  lblChannel(Index).Caption = "Channel"
End Sub

Private Sub wsTCP_Error(Index As Integer, ByVal Number As Integer, Description As String, _
                        ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, _
                        ByVal HelpContext As Long, CancelDisplay As Boolean)
                        
  AddChat Index, " > ", vbWhite
  AddChat Index, "Connection Error: ", fcRed, , NO_LINE
  AddChat Index, Description, fcYellow, , NO_LINE

  Disconnect Index

End Sub

Public Sub Connect(Index As Integer)
  wsTCP(Index).Close
  wsTCP(Index).Connect ProfileStructs(Index).Server, 6112
      
  AddChat Index, " > ", vbWhite
  AddChat Index, "Connecting to " & ProfileStructs(Index).Server, fcGreen, , NO_LINE
  
  SetConnectStatus Index, 1 'Working mode
  
End Sub

Public Sub Disconnect(Index As Integer)
  wsTCP(Index).Close
  wsTCP_Close Index
End Sub





















