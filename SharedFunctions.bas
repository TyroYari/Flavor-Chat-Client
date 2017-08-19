Attribute VB_Name = "SharedFunctions"
'==============================================================================================
' Shared Functions
'
' Author: Ben Pogrund
'
'==============================================================================================
Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Reading and writing to files
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

''----------------------------------------------------------------------------------------------
' File functions
''----------------------------------------------------------------------------------------------

Public Function ReadINI(Filename As String, Section As String, Key As String) As String
  Dim RetVal As String * 255  ' Passed by reference
  Dim Length As Long
  Length = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
  ReadINI = Left(RetVal, Length)
End Function

Public Function BoolINI(Filename As String, Section As String, Key As String) As Boolean
  Dim RetVal As String * 255  ' Passed by reference
  Dim Length As Long
  Length = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
  BoolINI = (Left(RetVal, Length) = True)
End Function

Public Function ReadIntINI(Filename As String, Section As String, Key As String) As Integer
  Dim RetVal As String * 255  ' Passed by reference
  Dim Length As Long
  Length = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
  ReadIntINI = Val(Left(RetVal, Length))
End Function

Public Sub WriteINI(Section As String, Key As String, Value As String, FilePath As String)
  WritePrivateProfileString Section, Key, Value, FilePath
End Sub

''----------------------------------------------------------------------------------------------
' UI functions
''----------------------------------------------------------------------------------------------

Public Sub AddChat(Index As Integer, Message As String, Color As Long, Optional Alignment As Long, Optional Noline As Long)
  If Alignment = 0 Then Alignment = rtfLeft
  With mainForm.rtbChat(Index)
    .SelAlignment = Alignment
    .SelStart = 99999
    .SelLength = 0
    .SelColor = Color
    If Noline = 0 Then .SelText = vbCrLf
    .SelText = Message
  End With
End Sub

Public Sub EdtrLoadLine(Index As Integer, Line As String)
  Select Case Index
    Case 0: AccessHighlighter Index, Line   'Access.txt
    Case 1: CDKeyHighlighter Index, Line    'CDKeys.txt
    Case 2: ServerHighlighter Index, Line   'Servers.txt
  End Select
End Sub

Public Sub SetProfileData(Data As String, rtb As RichTextBox)
  With rtb
    .SelAlignment = rtfLeft
    .SelStart = 0
    .SelLength = 99999
    .SelColor = RGB(192, 192, 192)
    .SelText = Data
  End With
End Sub

Public Sub SetConnectStatus(Index As Integer, Value As Integer)
  ProfileStructs(Index).iConnected = Value
  
  With mainForm
    If mainForm.bShftKeyDown = True Or bMassConnect Then
      mainForm.CloseTabButton(Index).Caption = ""
      Select Case Value
        Case 0:        .CloseTabButton(Index).Picture = .UI_Icons.ListImages(3).Picture
        Case 1:        .CloseTabButton(Index).Picture = .UI_Icons.ListImages(2).Picture
        Case 2:        .CloseTabButton(Index).Picture = .UI_Icons.ListImages(1).Picture
      End Select
    End If
    
    Select Case Value
      Case 0:        .ProfileTabs(Index).ForeColor = fcDarkGray
      Case 1:        .ProfileTabs(Index).ForeColor = fcYellow
      Case 2:        .ProfileTabs(Index).ForeColor = fcBrightGreen
    End Select
    
    ' False if connected, or disconnected, True if connecting
    .RetryTimer(Index).Enabled = (Value = 1)
    
  End With
  
  'Update window status
  Dim i As Integer
  Dim Count As Integer
  Dim bConnecting As Boolean
  bConnecting = False
  While i <= UBound(ProfileStructs)
    If ProfileStructs(i).iConnected = 2 Then
      Count = Count + 1
    ElseIf ProfileStructs(i).iConnected = 1 Then
      bConnecting = True
    End If
    i = i + 1
  Wend
  
  If bConnecting Then
    mainForm.Caption = "Flavor Chat (Connecting)"
  ElseIf Count = 0 Then
    mainForm.Caption = "Flavor Chat (Offline)"
  Else
    mainForm.Caption = "Flavor Chat (" & Count & " Online)"
  End If
  
End Sub

Public Sub LoadEditor(Index As Integer, Filename As String)
On Error GoTo Err
  CurrentEdtrIndx = Index
  CurrentEdtrFile = Filename
  
  With mainForm.EditorFrame
    .Visible = True
    .ZOrder BRING_TO_FRONT
  End With
  
  With mainForm.rtbEditor(Index)
    .SelStart = 0
    .SelLength = 99999
    .SelText = ""
    .SelColor = fcGray
    .ZOrder BRING_TO_FRONT
  End With
  
  Dim FileNumb As Integer
  Dim sInput As String
  FileNumb = FreeFile
  
  Open App.Path & "\txt\" & Filename For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Input #FileNumb, sInput
    EdtrLoadLine Index, sInput
    sInput = ""
  Loop
Err:  ' Input can get picky about how EOF works or something
  Close #FileNumb
  
  SetFileStatus Index, EDITOR_FILE_OPENED
  mainForm.EditorCoverObj.ZOrder BRING_TO_FRONT
  mainForm.EditorCoverObj.Left = 150 + (CurrentEdtrIndx * 2505)
  mainForm.rtbEditor(Index).SetFocus
End Sub

Public Sub SetFileStatus(Index As Integer, Status As Integer)
  
  Select Case Status
    Case 0:  mainForm.EditorTab(Index).ForeColor = fcDarkGray
    Case 1:  mainForm.EditorTab(Index).ForeColor = fcYellow
    Case 2:  mainForm.EditorTab(Index).ForeColor = fcBrightGreen
  End Select
  
  FileStatus(Index) = Status
  
End Sub

Public Sub SaveEditorFile(Index As Integer)
  Dim Filename As String
  Select Case Index
    Case 0:   Filename = "Access.txt"
    Case 1:   Filename = "CDKeys.txt"
    Case 2:   Filename = "Servers.txt"
  End Select
  
  Dim EditorText() As String
  EditorText = Split(mainForm.rtbEditor(Index).Text, vbCrLf)
  
  Dim FileNumb As Integer
  FileNumb = FreeFile
  Open App.Path & "\txt\" & Filename For Output As #FileNumb
  
  Dim i As Integer
  i = 0
  While i < UBound(EditorText)
    Print #FileNumb, EditorText(i)
    i = i + 1
  Wend
  Print #FileNumb, EditorText(i);
  
  Close #FileNumb
  
  ' Reload Access List
  If Index = 0 Then
    ReDim Preserve AccessList(0)
    AddChat MAIN_TAB, " > Loading Access List", vbWhite
    LoadAccessList
  End If
  
  SetFileStatus Index, EDITOR_FILE_OPENED
  
  ' Update highlighting
  Dim OldSelect As Integer
  OldSelect = mainForm.rtbEditor(Index).SelStart
  LoadEditor Index, Filename
  mainForm.rtbEditor(Index).SelStart = OldSelect
  
End Sub

Public Sub UnloadEditor()
  mainForm.EditorFrame.Visible = False
  SetFileStatus 0, EDITOR_FILE_CLOSED
  SetFileStatus 1, EDITOR_FILE_CLOSED
  SetFileStatus 2, EDITOR_FILE_CLOSED
End Sub

Public Sub UpdateChannel(Index As Integer)
  Dim Channel As String
  Dim Count As Integer
  
  Channel = ProfileStructs(Index).CurrChannel
  Count = mainForm.lvChannel(Index).ListItems.Count
  
  mainForm.lblChannel(Index).Caption = Channel & " (" & Count & ")"
End Sub

''----------------------------------------------------------------------------------------------
' Formatting functions
''----------------------------------------------------------------------------------------------

Public Sub CDKeyHighlighter(Index As Integer, Line As String)
  Dim Color As Long
  
  Select Case Len(Line)
    Case 13: Color = &HFF8080     ' Starcraft CDKey
    Case 16: Color = &H8080FF     ' Diablo II CDKey (or Warcraft 2)
    Case 26: Color = &H80C0FF     ' Warcraft3 CDKey
    Case Else
      Color = fcGray
  End Select
  
  'Should probably use IsCharAlphaNumericA function, but feeling lazy :P
  If InStr(1, Line, "-") <> 0 Then Color = fcGray
  If InStr(1, Line, "_") <> 0 Then Color = fcGray
  If InStr(1, Line, " ") <> 0 Then Color = fcGray
  
  With mainForm.rtbEditor(Index)
    .SelAlignment = rtfLeft
    .SelStart = 99999
    .SelLength = 0
    .SelColor = Color
    .SelText = Line
    .SelText = vbCrLf
  End With
End Sub

Public Sub CDKeyHighlighter_new(Index As Integer, SelStart As Integer, SelLength As Integer)
  Dim Color As Long
  Dim Line As String
  
  If SelLength < 3 Then Exit Sub
  Line = Mid(mainForm.rtbEditor(Index).Text, SelStart, SelLength - 2)
  
  AddChat MAIN_TAB, Line, fcPurple
  
  Select Case Len(Line)
    Case 13: Color = &HFF8080     ' Starcraft CDKey
    Case 16: Color = &H8080FF     ' Diablo II CDKey (or Warcraft 2)
    Case 26: Color = &H80C0FF     ' Warcraft3 CDKey
    Case Else
      Color = fcGray
  End Select
  
  'Should probably use IsCharAlphaNumericA function, but feeling lazy :P
  If InStr(1, Line, "-") <> 0 Then Color = fcGray
  If InStr(1, Line, "_") <> 0 Then Color = fcGray
  If InStr(1, Line, " ") <> 0 Then Color = fcGray
  
  With mainForm.rtbEditor(Index)
    .SelStart = SelStart
    .SelLength = SelLength
    .SelColor = Color
  End With
End Sub

Public Sub AccessHighlighter(Index As Integer, Line As String)
  Dim SplitString() As String
  
  SplitString = Split(Line, " ")
  
  With mainForm.rtbEditor(Index)
    .SelAlignment = rtfLeft
    .SelStart = 99999
    .SelLength = 0
    If UBound(SplitString) = 1 Then
      If IsNumeric(SplitString(1)) = True Then  'Would have liked to use short circuit here, but its not supported in vb6
        .SelColor = &H80C0FF
        .SelText = SplitString(0)
      
        If SplitString(1) >= 100 Then
          .SelColor = &H80FF80          ' Super user
        ElseIf SplitString(1) >= 0 Then
          .SelColor = fcGray            ' Normal user
        Else
          .SelColor = &H8080FF          ' Blacklist user
        End If
        
        .SelText = " " & SplitString(1)  ' Have to reinsert our delimiter here
        .SelText = vbCrLf
        Exit Sub
      End If
    End If
    .SelColor = fcGray
    .SelText = Line
    .SelText = vbCrLf
  End With
End Sub

Public Sub ServerHighlighter(Index As Integer, Line As String)
  With mainForm.rtbEditor(Index)
    .SelAlignment = rtfLeft
    .SelStart = 99999
    .SelLength = 0
    
    If InStr(1, Line, "battle.net") <> 0 Then
      .SelColor = &HFF8080
    ElseIf CheckValidIP(Line) = True Then
      .SelColor = &H8080FF
    ElseIf InStr(1, Line, " ") = 0 Then
      .SelColor = &H80C0FF
    Else
      .SelColor = fcGray
    End If
    
    .SelText = Line
    .SelText = vbCrLf
  End With
End Sub

Public Function CheckValidIP(ServerText As String) As Boolean
  Dim SplitString() As String
  CheckValidIP = False
  
  SplitString = Split(ServerText, ".")
  
  If UBound(SplitString) <> 3 Then Exit Function
  
  If IsNumeric(SplitString(0)) = False Then Exit Function
  If IsNumeric(SplitString(1)) = False Then Exit Function
  If IsNumeric(SplitString(2)) = False Then Exit Function
  If IsNumeric(SplitString(3)) = False Then Exit Function
  
  CheckValidIP = True
  
End Function

Public Function TabFormat(str As String) As String
' Formats tab captions to fit 12 character limit using  ellipsis (...)
  TabFormat = str
  While Len(TabFormat) < 12
    TabFormat = TabFormat & " "
  Wend
  If Len(TabFormat) = 12 Then Exit Function
  TabFormat = Left(TabFormat, 9) & "..."
End Function

''----------------------------------------------------------------------------------------------
' Bot Functions
''----------------------------------------------------------------------------------------------

Public Function ParseChatCommand(Index As Integer, Command As String, Args As String, Optional Instigator As String = "*") As Boolean
' Returns "True" if an internal bot command is recognized
' Returns "False" to send command as message to server
  Dim bCaught As Boolean
  bCaught = True
  ParseChatCommand = True
  
  ' Check for simple commands
  Select Case Command
    Case "j", "join":
      If CheckPermission(Instigator, AL_JOIN) = False Then Exit Function
      Send0x0E Index, "/join " & Args
      
    Case "home":
      If CheckPermission(Instigator, AL_HOME) = False Then Exit Function
      Send0x0E Index, "/join " & ProfileStructs(Index).Home
     
    Case "sethome":
      If CheckPermission(Instigator, AL_SETHOME) = False Then Exit Function
      ProfileStructs(Index).Home = Args
      Send0x0E Index, "Home set to " & ProfileStructs(Index).Home
    
    Case "say":
      Dim AccessLevel As Integer
      If Left(Args, 1) = "/" Then AccessLevel = AL_SAY2 Else AccessLevel = AL_SAY1
      
      If CheckPermission(Instigator, AccessLevel) = False Then Exit Function
      Send0x0E Index, Args
    
    Case "ver":
      If CheckPermission(Instigator, AL_VER) = False Then Exit Function
      Send0x0E Index, "/me " & VersionMacro
    
    Case "add", "set":
      If CheckPermission(Instigator, AL_ADD) = False Then Exit Function
      AddUserAccess Index, Args
      
    Case "whitelist", "wl":
      If CheckPermission(Instigator, AL_WHITELIST) = False Then Exit Function
      SpecifyUserAccess Index, Args, 0
      
    Case "remove", "rem":
      If CheckPermission(Instigator, AL_REMOVE) = False Then Exit Function
      RemoveUserAccess Index, Args
    
    Case "shitlist", "sl", "blacklist", "bl":
      If CheckPermission(Instigator, AL_BLACKLIST) = False Then Exit Function
      SpecifyUserAccess Index, Args, -1
      
    Case "rj", "rejoin":
      If CheckPermission(Instigator, AL_REJOIN) = False Then Exit Function
      Send0x10 Index
      Send0x0C Index, ProfileStructs(Index).LastChannel
      
    Case "rc", "reconnect":
      If CheckPermission(Instigator, AL_RECONNECT) = False Then Exit Function
      mainForm.Disconnect Index
      mainForm.Connect Index
      
    Case "spoof":
      If CheckPermission(Instigator, AL_SPOOF) = False Then Exit Function
      ProfileStructs(Index).SpoofClient = Left(Args, 4)
      AddChat CInt(Index), " > ", vbWhite
      AddChat CInt(Index), "Client Spoofing: " & Left(Args, 4), fcBlue, , NO_LINE
    
    Case Else
      bCaught = False
  End Select
  If bCaught Then Exit Function
  
  ' Check for Kleene star processing
  ProcessKleene Index, Args
  
  ' Commands with Kleene support
  Select Case Command
    Case "op":
      If CheckPermission(Instigator, AL_OP) = False Then Exit Function
      Send0x0E Index, "/op " & Args
      
    Case "k", "kick":
      If CheckPermission(Instigator, AL_KICK) = False Then Exit Function
      Send0x0E Index, "/kick " & Args
      
    Case "b", "ban":
      If CheckPermission(Instigator, AL_BAN) = False Then Exit Function
      Send0x0E Index, "/ban " & Args
      
    Case Else
      ParseChatCommand = False
  End Select
End Function

Public Sub ProcessKleene(Index As Integer, ByRef Args As String)
  Dim SplitString() As String
  If Args = "" Then Exit Sub
  
  SplitString = Split(Args, " ")
  
  Dim k As Integer, Length As Integer
  
  If Left(SplitString(0), 1) = "*" Then
    Length = Len(SplitString(0)) - 1
    k = 1
    While k <= mainForm.lvChannel(Index).ListItems.Count
      If LCase(Right(mainForm.lvChannel(Index).ListItems(k).Text, Length)) = LCase(Right(SplitString(0), Length)) Then
        SplitString(0) = mainForm.lvChannel(Index).ListItems(k).Text
        '... others?
      End If
      k = k + 1
    Wend
  End If
  
  If Right(SplitString(0), 1) = "*" Then
    Length = Len(SplitString(0)) - 1
    k = 1
    While k <= mainForm.lvChannel(Index).ListItems.Count
      If LCase(Left(mainForm.lvChannel(Index).ListItems(k).Text, Length)) = LCase(Left(SplitString(0), Length)) Then
        SplitString(0) = mainForm.lvChannel(Index).ListItems(k).Text
        '... others?
      End If
      k = k + 1
    Wend
  End If
  
  Dim i As Integer
  Args = ""
  While i < UBound(SplitString)
    Args = Args & SplitString(i) & " "
    i = i + 1
  Wend
  Args = Args & SplitString(i)
  
End Sub

Public Sub LoadAccessList()
On Error GoTo Err
  Dim SUserCount, NUserCount, WUserCount, BUserCount
  Dim FileNumb As Integer
  Dim FileLine As String
  Dim SplitLine() As String
  FileNumb = FreeFile
  
  SUserCount = 0
  NUserCount = 0
  WUserCount = 0
  BUserCount = 0
  
  Open App.Path & "\txt\access.txt" For Input As #FileNumb
  Do Until EOF(FileNumb)
    Input #FileNumb, FileLine
    
    SplitLine = Split(FileLine, " ")
    If UBound(SplitLine) = 1 Then
      If IsNumeric(SplitLine(1)) = True Then
        AccessList(UBound(AccessList)).Name = SplitLine(0)
        AccessList(UBound(AccessList)).Level = Val(SplitLine(1))
        ReDim Preserve AccessList(UBound(AccessList) + 1)
        
        If Val(SplitLine(1)) >= 100 Then
          SUserCount = SUserCount + 1
        ElseIf Val(SplitLine(1)) >= 1 Then
          NUserCount = NUserCount + 1
        ElseIf Val(SplitLine(1)) = 0 Then
          WUserCount = WUserCount + 1
        ElseIf Val(SplitLine(1)) < 0 Then
          BUserCount = BUserCount + 1
        End If
        
      End If
    End If
    FileLine = ""
  Loop
Err:
  Close #FileNumb

  AddChat MAIN_TAB, "    ", vbWhite
  AddChat MAIN_TAB, SUserCount & " Super users, ", &H80FF80, , NO_LINE
  'AddChat MAIN_TAB, "    ", vbWhite, , NO_LINE
  AddChat MAIN_TAB, NUserCount + WUserCount & " Normal users, ", fcGray, , NO_LINE
  'AddChat MAIN_TAB, "    ", vbWhite, , NO_LINE
  AddChat MAIN_TAB, BUserCount & " Blacklisted users", &H8080FF, , NO_LINE
  
End Sub

Public Sub PrintPermissions(Index As Integer)
  AddChat Index, "", fcBlue
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "  Permissions Table                        ", fcGray
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "   Add            " & AL_ADD & " Access      ", fcBlue
  AddChat Index, "   Remove         " & AL_REMOVE & " Access   ", fcBlue
  AddChat Index, "   Op             " & AL_OP & " Access       ", fcBlue
  AddChat Index, "   Blacklist      " & AL_BLACKLIST & " Access", fcBlue
  AddChat Index, "   Ban            " & AL_BAN & " Access      ", fcBlue
  AddChat Index, "   Reconnect      " & AL_RECONNECT & " Access", fcBlue
  AddChat Index, "   Say            " & AL_SAY2 & " Access     ", fcBlue
  AddChat Index, "   Join           " & AL_JOIN & " Access     ", fcBlue
  AddChat Index, "   Spoof          " & AL_SPOOF & " Access    ", fcBlue
  AddChat Index, "   Whitelist      " & AL_WHITELIST & " Access", fcBlue
  AddChat Index, "   Kick           " & AL_KICK & " Access     ", fcBlue
  AddChat Index, "   Rejoin         " & AL_REJOIN & " Access   ", fcBlue
  AddChat Index, "   Sethome        " & AL_SETHOME & " Access  ", fcBlue
  AddChat Index, "   Home           " & AL_HOME & " Access     ", fcBlue
  AddChat Index, "   Say (No slash) " & AL_SAY1 & " Access     ", fcBlue
  AddChat Index, "   Ver            " & AL_VER & " Access      ", fcBlue
  AddChat Index, "   ?Trigger       " & AL_TRIGGER & " Access       ", fcBlue
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "", fcBlue
  
End Sub

Public Sub PrintSpoofs(Index As Integer)
  AddChat Index, "", fcBlue
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "  Spoof Table (For PvPGN)                  ", fcGray
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "   CHAT            (Telnet)                ", fcBlue
  AddChat Index, "   SSHR            (Starcraft Shareware)   ", fcBlue
  AddChat Index, "   JSTR            (Starcraft Japan)       ", fcBlue
  AddChat Index, "   STAR            (Starcraft)             ", fcBlue
  AddChat Index, "   SEXP            (Starcraft Exp)         ", fcBlue
  AddChat Index, "   DSHR            (Diablo Shareware)      ", fcBlue
  AddChat Index, "   DIAB            (Diablo Beta)           ", fcBlue
  AddChat Index, "   DTST            (Diablo Stress Test)    ", fcBlue
  AddChat Index, "   DRTL            (Diablo)                ", fcBlue
  AddChat Index, "   D2ST            (Diablo 2 Stress Test)  ", fcBlue
  AddChat Index, "   D2DV            (Diablo 2)              ", fcBlue
  AddChat Index, "   D2XP            (Diablo 2 Exp)          ", fcBlue
  AddChat Index, "   W2BN            (Warcraft 2 BNE)        ", fcBlue
  AddChat Index, "   W3DM            (Warcraft III Demo)     ", fcBlue
  AddChat Index, "   WAR3            (Warcraft III)          ", fcBlue
  AddChat Index, "   W3XP            (Warcraft III Exp)      ", fcBlue
  AddChat Index, " ------------------------------------------", fcBlue
  AddChat Index, "", fcBlue
  
End Sub

Public Function CheckPermission(Invoker As String, Level As Integer) As Boolean
' If invoked from the client window, skip permission check
  If Invoker = "*" Then
    CheckPermission = True
    Exit Function
  End If
  
  ' User lookup
  Dim i As Integer
  While i < UBound(AccessList)
    If LCase(AccessList(i).Name) = LCase(Invoker) Then GoTo CheckLevel
    i = i + 1
  Wend
  
  ' Not Found...
  CheckPermission = False
  Exit Function
  
CheckLevel:
  If AccessList(i).Level >= Level Then CheckPermission = True Else CheckPermission = False
  
End Function

Public Function ReplaceCodes(GreetMsg As String, Username As String, UniqName As String, CurrChannel As String) As String
  ReplaceCodes = GreetMsg
  ReplaceCodes = Replace(ReplaceCodes, "%name", Username)
  ReplaceCodes = Replace(ReplaceCodes, "%n", Username)
  
  ReplaceCodes = Replace(ReplaceCodes, "%myself", UniqName)
  ReplaceCodes = Replace(ReplaceCodes, "%m", UniqName)
  
  ReplaceCodes = Replace(ReplaceCodes, "%time", Time())
  ReplaceCodes = Replace(ReplaceCodes, "%t", Time())
  
  ReplaceCodes = Replace(ReplaceCodes, "%ver", VersionMacro)
  ReplaceCodes = Replace(ReplaceCodes, "%v", VersionMacro)
  
  ReplaceCodes = Replace(ReplaceCodes, "%channel", CurrChannel)
  ReplaceCodes = Replace(ReplaceCodes, "%c", CurrChannel)
  
End Function

''----------------------------------------------------------------------------------------------
' Packet manipulation
''----------------------------------------------------------------------------------------------

Public Function MakeLong(Data As String) As Long
  If Len(Data) < 4 Then Exit Function
  CopyMemory MakeLong, ByVal Data, 4
End Function

Public Function MakeByte(Data As String) As Byte
  If Len(Data) < 1 Then Exit Function
  CopyMemory MakeByte, ByVal Data, 1
End Function

Public Function KillNull(Data As String) As String
  If InStrB(Data, vbNullChar) Then
    KillNull = Left$(Data, InStr(Data, vbNullChar) - 1)
  Else
    KillNull = Data
  End If
End Function

''----------------------------------------------------------------------------------------------
' Access List Commands
''----------------------------------------------------------------------------------------------

Public Sub AddUserAccess(Index As Integer, Args As String)
  Dim SplitString() As String
  
  If Args = "" Then
    Send0x0E Index, "Instructions: add <username> <access level>"
    Exit Sub
  End If
  
  SplitString = Split(Args, " ")
  If UBound(SplitString) < 1 Then
    Send0x0E Index, "Bad syntax."
    Exit Sub
  ElseIf IsNumeric(SplitString(1)) = False Then
    Send0x0E Index, "Bad syntax."
    Exit Sub
  End If
  
  ' SplitString(0) holds username
  ' SplitString(1) holds access level
  SpecifyUserAccess Index, SplitString(0), Val(SplitString(1))
End Sub

Public Sub SpecifyUserAccess(Index As Integer, Args As String, Level As Integer)
  If Args = "" Then Exit Sub
  
  Dim SplitString() As String
  Dim Username As String
  SplitString = Split(Args, " ")
  Username = SplitString(0)
  
  ' look for user in accesslist
  Dim i As Integer, Found As Boolean
  Found = False
  i = 0
  While i < UBound(AccessList)
    ' if found change level
    If LCase(AccessList(i).Name) = LCase(Username) Then
      AccessList(i).Level = Level
      Found = True
      Send0x0E Index, "Set " & Username & " to " & Level & "."
    End If
    i = i + 1
  Wend
  
  ' else add to list
  If Not Found Then
    AccessList(UBound(AccessList)).Name = Username
    AccessList(UBound(AccessList)).Level = Level
    ReDim Preserve AccessList(UBound(AccessList) + 1)
    Send0x0E Index, "Added " & Username & " with " & Level & "."
  End If
  
  ' look for user in access.txt
  ' if found, change it
  Dim FileNumb As Integer
  Dim FileLine As String, SplitLine() As String
  Dim FileOutput As String
  Dim FoundUser As Boolean
  
  FoundUser = False
  FileNumb = FreeFile
  
  Open App.Path & "\txt\access.txt" For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Input #FileNumb, FileLine
    SplitLine = Split(FileLine, " ")
    If UBound(SplitLine) = 1 Then
      If IsNumeric(SplitLine(1)) = True Then
        If LCase(SplitLine(0)) = LCase(Username) Then
          SplitLine(1) = CStr(Level)
          FileLine = SplitLine(0) & " " & SplitLine(1)
          FoundUser = True
        End If
      End If
    End If
    FileOutput = FileOutput & FileLine & vbCrLf
    
    FileLine = ""
  Loop
  Close #FileNumb
  
  FileNumb = FreeFile
  
  Open App.Path & "\txt\access.txt" For Output As #FileNumb
  Write #FileNumb, FileOutput
  If FoundUser = False Then Write #FileNumb, Username & " " & Level
  Close #FileNumb
  
  ' if access.txt is "opened" status then ... welp
  ' ...
  ' meh
  
End Sub

Public Sub RemoveUserAccess(Index As Integer, Args As String)
  If Args = "" Then
    Send0x0E Index, "Instructions: rem <username>"
    Exit Sub
  End If
  
  ' look for user in accesslist
  ' if found remove
  Dim SplitString() As String, Username As String
  SplitString = Split(Args, " ")
  Username = SplitString(0)
  
  Dim i As Integer, Found As Boolean
  Found = False
  i = 0
  While i < UBound(AccessList)
    If LCase(AccessList(i).Name) = LCase(Username) Then
      GoTo RemoveStep
    End If
    i = i + 1
  Wend
  Send0x0E Index, "Could not find user."
  Exit Sub
  
RemoveStep:
  Send0x0E Index, "Access for " & LCase(Username) & " removed."
  
  While i < UBound(AccessList)
    AccessList(i) = AccessList(i + 1)
    i = i + 1
  Wend
  ReDim Preserve AccessList(UBound(AccessList) - 1)
  
  ' Rewrite Access.txt without specified user
  
  
  
  Dim FileNumb As Integer   ' Stores index for file descriptor
  Dim FileLine As String    ' Stores a line read from file
  FileNumb = FreeFile
  
  Dim FullFile As String
  
  Open App.Path & "\txt\Access.txt" For Input As #FileNumb
  Do While Not EOF(FileNumb)
    Input #FileNumb, FileLine
    FullFile = FullFile & FileLine & vbCrLf
    FileLine = ""
  Loop
  Close #FileNumb
  
  AddChat Index, FullFile, fcPurple
  
  Dim OutLines() As String
  OutLines = Split(FullFile, vbCrLf)
  
  FileNumb = FreeFile
  Open App.Path & "\txt\Access.txt" For Output As #FileNumb
  
  i = 0
  While i < UBound(OutLines)
    If LCase(Left(OutLines(i), Len(Username))) <> LCase(Username) Then
      Print #FileNumb, OutLines(i)
    End If
    i = i + 1
  Wend
  Print #FileNumb, OutLines(i);
  
  Close #FileNumb
  '...f
  
End Sub

''----------------------------------------------------------------------------------------------
' Misc Commands
''----------------------------------------------------------------------------------------------

Public Function GetKeyCount(Client As String) As Integer
  Select Case Client
    Case "Diablo II LOD", "Warcraft 3XP":
      GetKeyCount = 2
    Case "Warcraft 2", "Diablo II", "Starcraft", "Starcraft BW", "Warcraft 3":
      GetKeyCount = 1
    Case Else
      GetKeyCount = 0
  End Select
End Function














































