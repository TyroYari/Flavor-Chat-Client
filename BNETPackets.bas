Attribute VB_Name = "BNETPackets"
'==============================================================================================
' BNET Packets
'
' Author: Ben Pogrund
'
'==============================================================================================
Option Explicit

Public pbuffer() As New PacketBuffer

Const pOFFSET As Integer = 4

' APIs
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numBytes As Long)
'

Private Function OffsetDWORD(offset As Integer) As Integer
  OffsetDWORD = 1 + offset * 4
End Function

Public Sub ParseBNCS(Index As Integer, Data As String, PacketID As Byte)

'Dim Name As String
'Dim Text As String
'Dim Ping As String
'Dim Flags As Long
'Dim EID As Long

  Select Case PacketID
    Case &H25, &H59, &HA, &HF
    ' Hide these packets from output protocol
    
    Case &HA
      If ProfileStructs(Index).iConnected = 1 Then
        AddChat Index, "     Received <- 0x0" & Hex(PacketID), fcOrange
      End If
      
    Case Else
      If ProfileStructs(Index).iConnected = 1 Then
        AddChat Index, "     Received <- 0x" & Hex(PacketID), fcOrange
      End If
  End Select
  
  Select Case PacketID
    'Case &H4C:           rtbChat vbGreen, "<< 0x4C - Required Work: " & KillNull(Mid(Data, 5))
    Case &H0:            pbuffer(Index).sendPacket Index, &H0
    Case &HA:            Parse0x0A Index, Data
    Case &HF:            Parse0x0F Index, Data
    Case &H19:           Parse0x19 Index, Data
    Case &H25:
      If ProfileStructs(Index).iConnected <> 0 Then mainForm.wsTCP(Index).SendData Data
    Case &H26:           Parse0x26 Data
    Case &H50:           Parse0x50 Index, Data
    Case &H51:           Parse0x51 Index, Data
    Case &H3A:           Parse0x3A Index, Data
    Case &H3D:           Parse0x3D Index, Data
    Case &H59:           Parse0x59 Index
    Case &H67:           Parse0x67 Index, Data
    Case &H68:           Parse0x68 Index, Data
    Case Else:           AddChat Index, "   * Received unhandled packet [ 0x" & Hex(PacketID) & " ]", fcYellow
  End Select
End Sub


Private Sub Send0x0A(Index As Integer, Username As String)
'SID_ENTERCHAT
  With pbuffer(Index)
    .InsertNTString Username    ' Username
    .InsertBYTE 0               ' Statstring
    .sendPacket Index, &HA
  End With
End Sub

Private Sub Parse0x0A(Index As Integer, Data As String)
'SID_ENTERCHAT

' (STRING) Unique name
' (STRING) Statstring
' (STRING) Account name

  Dim str() As String
  
  str = Split(Data, vbNullChar)
  
  ProfileStructs(Index).UniqName = str(1)

  AddChat Index, " > ", vbWhite
  AddChat Index, "Logged in as " & ProfileStructs(Index).UniqName & ", joining '" & ProfileStructs(Index).Home & "'", fcGreen, , NO_LINE
  
  mainForm.ProfileTabs(Index).Caption = TabFormat(ProfileStructs(Index).UniqName)
  
  SetConnectStatus Index, 2 'Online mode
  
End Sub


Public Sub Send0x0C(Index As Integer, Channel As String)
'SID_JOINCHANNEL
  If Channel = vbNullString Then Channel = "The Flavor Hideout"
  
  With pbuffer(Index)
    .InsertDWORD &H2            ' Flags (0x01: First join, 0x02: Forced join)
    .InsertNTString Channel     ' Channel
    .sendPacket Index, &HC
  End With
End Sub
      
Public Sub Send0x0E(Index As Integer, Message As String)
'SID_CHATCOMMAND
  Dim sMsg As String
  
  ProfileStructs(Index).IdleElapsed = 0
          
  sMsg = Mid$(Message, 1, 223)
  If Len(sMsg) = 0 Then Exit Sub
  
  With pbuffer(Index)
    .InsertNTString sMsg
    .sendPacket Index, &HE
  End With
  
  If Left(Message, 1) <> "/" Then
    AddChat Index, " > ", vbWhite
    AddChat Index, "<" & ProfileStructs(Index).Username & "> ", fcBlue, , NO_LINE
    AddChat Index, Message, fcGray, , NO_LINE
  End If
  
End Sub

Public Sub Parse0x0F(Index As Integer, Data As String)
'(UINT32) Event ID
'(UINT32) User's Flags
'(UINT32) Ping
'(UINT32) IP Address *
'(UINT32) Account number *
'(UINT32) Registration Authority *
'(STRING) Username
'(STRING) Text **

'---------
'SID_CHATEVENT
'0x01 EID_SHOWUSER:             User in channel
'0x02 EID_JOIN:                 User joined channel
'0x03 EID_LEAVE:                User left channel
'0x04 EID_WHISPER:              Recieved whisper
'0x05 EID_TALK:                 Chat Text
'0x06 EID_BROADCAST:            Server broadcast
'0x07 EID_CHANNEL:              Channel Information
'0x09 EID_USERFLAGS:            Flags Update
'0x0A EID_WHISPERSENT:          Sent whisper
'0x0D EID_CHANNELFULL:          Channel full
'0x0E EID_CHANNELDOESNOTEXIST:  Channel doesn 't exist
'0x0F EID_CHANNELRESTRICTED:    Channel is restricted
'0x12 EID_INFO:                 Information
'0x13 EID_ERROR:                Error Message
'0x15 EID_IGNORE:               Notifies that a user has been ignored (DEFUNCT)
'0x16 EID_ACCEPT:               Notifies that a user has been unignored (DEFUNCT)
'0x17 EID_EMOTE:                Emote

  Dim EventID As Long
  Dim UserFlags As String
  Dim Ping As Long
  Dim sUsername As String
  Dim sMessage As String

  EventID = MakeLong(Mid$(Data, 5, 4))
  UserFlags = MakeLong(Mid$(Data, 9, 4))
  Ping = MakeLong(Mid$(Data, 13, 4))
  sUsername = KillNull(Mid$(Data, 29))
  sMessage = KillNull(Mid$(Data, Len(sUsername) + 30))

  Dim Length As Long
  Dim Command As String
  Dim Args As String
  
  Select Case EventID
    Case &H1            'EID_SHOWUSER
      Select Case UserFlags
        Case 1:         sMessage = "BLIZ"
        Case 2:         sMessage = "MOD"
        Case 4:         sMessage = "SPKR"
        Case 8:         sMessage = "SYSOP"
        Case Else
          sMessage = StrReverse$(Left(sMessage, 4))
      End Select
      mainForm.AddToChannel Index, sUsername, sMessage
      Exit Sub
      
    Case &H2            'EID_JOIN
      AddChat Index, " > ", vbWhite
      AddChat Index, sUsername & " has joined the channel.  -", fcGreen, , NO_LINE
      AddChat Index, StrReverse$(Left(sMessage, 4)), fcYellow, , NO_LINE
      AddChat Index, "-", fcGreen, , NO_LINE
      mainForm.AddToChannel Index, sUsername, StrReverse$(Left(sMessage, 4))
      
      ' Greeting
      With ProfileStructs(Index)
        If .GreetOn Then
          Dim ModifiedMsg As String
          ModifiedMsg = ReplaceCodes(.GreetMsg, sUsername, .UniqName, .CurrChannel)
          Send0x0E Index, ModifiedMsg
        End If
      End With
      Exit Sub
      
    Case &H3            'EID_LEAVE
      AddChat Index, " > ", vbWhite
      AddChat Index, sUsername & " has left the channel.", fcGreen, , NO_LINE
      mainForm.RemoveFromCh Index, sUsername
      Exit Sub
      
    Case &H4            'EID_WHISPER
      AddChat Index, " > ", vbWhite
      AddChat Index, "<From: " & sUsername & "> ", fcYellow, , NO_LINE
      AddChat Index, sMessage, fcDarkGray, , NO_LINE
      
      'Process Chat Commands
      Length = InStr(1, sMessage, " ") - 2
      If Length <= 0 Then Length = Len(sMessage) - 1
      Command = LCase(Mid(sMessage, 2, Length))
      Args = Mid(sMessage, Len(Command) + 3)
      
      If Left(sMessage, 1) = ProfileStructs(Index).Trigger Then
        ParseChatCommand Index, Command, Args, sUsername
      ElseIf LCase(Left(sMessage, 8)) = "?trigger" Then
        If CheckPermission(sUsername, AL_TRIGGER) Then
          Send0x0E CInt(Index), "/w " & sUsername & " My trigger is '" & ProfileStructs(Index).Trigger & "'"
        End If
      End If
      Exit Sub
      
    Case &H5            'EID_TALK
      Dim Color1 As Long
      Dim Color2 As Long
      If sUsername = ProfileStructs(Index).UniqName Then
        ' Self
        Color1 = fcBlue
        Color2 = fcGray
      ElseIf Left$(mainForm.ChLookup(Index, sUsername), 4) = "BLIZ" Then
        ' Blizzard
        Color1 = fcBlue
        Color2 = fcBlue
      ElseIf Left$(mainForm.ChLookup(Index, sUsername), 3) = "MOD" Then
        ' Operator
        Color1 = vbWhite
        Color2 = fcGray
      ElseIf Left$(mainForm.ChLookup(Index, sUsername), 5) = "SYSOP" Then
        ' System operator
        Color1 = &H46F758
        Color2 = fcYellow
      ElseIf Left$(mainForm.ChLookup(Index, sUsername), 4) = "SPKR" Then
        ' System operator
        Color1 = fcOrange
        Color2 = fcYellow
      Else
        'Normal users
        Color1 = fcOrange
        Color2 = fcGray
      End If
      AddChat Index, " > ", vbWhite
      AddChat Index, "<" & sUsername & "> ", Color1, , NO_LINE
      AddChat Index, sMessage, Color2, , NO_LINE
      
      'Process Chat Commands
      Length = InStr(1, sMessage, " ") - 2
      If Length <= 0 Then Length = Len(sMessage) - 1
      Command = LCase(Mid(sMessage, 2, Length))
      Args = Mid(sMessage, Len(Command) + 3)
      
      If Left(sMessage, 1) = ProfileStructs(Index).Trigger Then
        ParseChatCommand Index, Command, Args, sUsername
      ElseIf LCase(Left(sMessage, 8)) = "?trigger" Then
        If CheckPermission(sUsername, AL_TRIGGER) Then
          Send0x0E CInt(Index), "My trigger is '" & ProfileStructs(Index).Trigger & "'"
        End If
      End If
      
      'Name Notification
      If InStr(1, LCase(sMessage), LCase(ProfileStructs(Index).Username)) <> 0 _
      Or InStr(1, LCase(sMessage), LCase(ProfileStructs(Index).UniqName)) <> 0 Then
        If CurrentProfIndx <> Index Then
          ProfileStructs(Index).bNotifySet = True
        End If
      End If
      
      Exit Sub
    
    Case &H6            'EID_BROADCAST
      AddChat Index, " > ", vbWhite
      AddChat Index, sMessage, fcBlue, , NO_LINE
      Exit Sub
      
    Case &H7            'EID_CHANNEL
      ProfileStructs(Index).CurrChannel = sMessage
      mainForm.lvChannel(Index).ListItems.Clear
      AddChat Index, "", fcGreen
      AddChat Index, " --- Joined Channel: " & sMessage & " --- ", fcGreen
      If LCase(sMessage) <> "the void" Then
        ProfileStructs(Index).LastChannel = sMessage
      Else
        mainForm.RejoinTimer(Index).Enabled = True
      End If
      Exit Sub
      
    Case &H9            'EID_USERFLAGS (Update Flags)
      Select Case UserFlags
        Case 1:  mainForm.ChUpdate Index, sUsername, "BLIZ"
        Case 2:  mainForm.ChUpdate Index, sUsername, "MOD"
        Case 4:  mainForm.ChUpdate Index, sUsername, "SPKR"
        Case 8:  mainForm.ChUpdate Index, sUsername, "SYSOP"
        Case Else
          mainForm.ChUpdate Index, sUsername, StrReverse$(Left(sMessage, 4))
      End Select
      Exit Sub
      
    Case &HA            'EID_WHISPERSENT
      AddChat Index, " > ", vbWhite
      AddChat Index, "<To: " & sUsername & "> ", fcBlue, , NO_LINE
      AddChat Index, sMessage, fcDarkGray, , NO_LINE
      Exit Sub
      
    Case &H12           'EID_INFO
      AddChat Index, " > ", vbWhite
      AddChat Index, CStr(sMessage), fcOrange, , NO_LINE
      Exit Sub
      
    Case &H13           'EID_ERROR
      AddChat Index, " > ", vbWhite
      AddChat Index, CStr(sMessage), fcRed, , NO_LINE
      Exit Sub
      
    Case &H17           'EID_EMOTE
      If Left$(mainForm.ChLookup(Index, sUsername), 3) = "MOD" Then
        Color1 = vbWhite
      Else
        Color1 = fcOrange
      End If
        
      AddChat Index, " > ", vbWhite
      AddChat Index, "<" & sUsername & " " & sMessage & "> ", Color1, , NO_LINE
      Exit Sub
      
    Case Else
      AddChat Index, " > ", vbWhite
      AddChat Index, "Unhandled 0x0F (&H" & Hex(EventID) & ")", fcRed, , NO_LINE
      AddChat Index, " [ " & CStr(EventID) & ", " & CStr(UserFlags) & ", " & CStr(Ping), fcPurple, , NO_LINE
      AddChat Index, ", " & CStr(sUsername) & ", " & CStr(sMessage) & " ] ", fcPurple, , NO_LINE
      
  End Select
  
End Sub

Public Sub Send0x10(Index As Integer)
'SID_LEAVECHAT
  With pbuffer(Index)
    '[blank]
    .sendPacket Index, &H10
  End With
End Sub

Private Sub Parse0x19(Index As Integer, Data As String)
'SID_MESSAGEBOX
  Dim Message As String
  
  Message = Mid$(Data, OffsetDWORD(2))
  AddChat Index, "", vbWhite
  AddChat Index, "[Message Box]", fcOrange
  AddChat Index, "   " & Message, fcYellow, rtfCenter
  AddChat Index, "", vbWhite

End Sub

Public Sub Send0x26(Index As Integer, AccountName As String)
'SID_READUSERDATA

'(UINT32) Number of Accounts
'(UINT32) Number of Keys
'(UINT32) Request ID
'(STRING)[] Requested Accounts
'(STRING)[] Requested Keys

  Load ProfileForm
  ProfileForm.Show
  ProfileForm.ViewIndex = Index
  ProfileForm.Caption = "Profile Viewer (VIEWING AS: " & ProfileStructs(Index).UniqName & ")"
  
  With pbuffer(Index)
    .InsertDWORD &H1
    .InsertDWORD &H10    ' Number of entries (16)
    .InsertDWORD &H1
    .InsertNTString AccountName
    .InsertNTString "profile\sex"
    .InsertNTString "profile\age"
    .InsertNTString "profile\location"
    .InsertNTString "profile\description"
    
    .InsertNTString "Record\W2BN\0\wins"
    .InsertNTString "Record\W2BN\0\losses"
    .InsertNTString "Record\W2BN\0\disconnects"
    
    .InsertNTString "Record\STAR\0\wins"
    .InsertNTString "Record\STAR\0\losses"
    .InsertNTString "Record\STAR\0\disconnects"
    
    .InsertNTString "Record\SEXP\0\wins"
    .InsertNTString "Record\SEXP\0\losses"
    .InsertNTString "Record\SEXP\0\disconnects"
    
    .InsertNTString "Record\WAR3\0\wins"
    .InsertNTString "Record\WAR3\0\losses"
    .InsertNTString "Record\WAR3\0\disconnects"
    
    .sendPacket Index, &H26
  End With
  
  ProfileForm.lblName.Caption = AccountName
  
End Sub

Public Sub Parse0x26(Data As String)
'SID_READUSERDATA

'(UINT32) Number of accounts
'(UINT32) Number of keys
'(UINT32) Request ID
'(STRING) [] Requested Key Values
  
  Dim DisplayData() As String
  
  Data = Mid$(Data, 17) 'Eat thru header
  DisplayData = Split(Data, vbNullChar)
  
  SetProfileData DisplayData(0), ProfileForm.rtbProfSex
  SetProfileData DisplayData(1), ProfileForm.rtbProfAge
  SetProfileData DisplayData(2), ProfileForm.rtbProfLoc
  SetProfileData DisplayData(3), ProfileForm.rtbProfDesc
  
  Dim i As Integer
  i = 4
  While i < 16
    If Len(DisplayData(i)) = 0 Then DisplayData(i) = "0"
    i = i + 1
  Wend
  
  ProfileForm.lblW2BNStats.Caption = DisplayData(4) & " - " & DisplayData(5) & " - " & DisplayData(6)
  ProfileForm.lblSTARStats.Caption = DisplayData(7) & " - " & DisplayData(8) & " - " & DisplayData(9)
  ProfileForm.lblSEXPStats.Caption = DisplayData(10) & " - " & DisplayData(11) & " - " & DisplayData(12)
  ProfileForm.lblWAR3Stats.Caption = DisplayData(13) & " - " & DisplayData(14) & " - " & DisplayData(15)
  
End Sub

Public Sub Send0x27(Index As Integer)
'SID_WRITEUSERDATA

'(UINT32) Number of accounts
'(UINT32) Number of keys
'(STRING) [] Accounts to update
'(STRING) [] Keys to update
'(STRING) [] New values

  With pbuffer(Index)
    .InsertDWORD &H1
    .InsertDWORD &H4    ' 4 Entries (Age,Sex,Location,Description)
    
    .InsertNTString ProfileStructs(Index).Username
    .InsertNTString "profile\sex"
    .InsertNTString "profile\age"
    .InsertNTString "profile\location"
    .InsertNTString "profile\description"
    
    .InsertNTString ProfileForm.rtbProfSex.Text
    .InsertNTString ProfileForm.rtbProfAge.Text
    .InsertNTString ProfileForm.rtbProfLoc.Text
    .InsertNTString ProfileForm.rtbProfDesc.Text
    
    .sendPacket Index, &H27
  End With

End Sub

Private Sub Send0x3A(Index As Integer)
'SID_LOGONRESPONSE2
 
 '(UINT32)     Client Token
 '(UINT32)     Server Token
 '(UINT8)[20]  Password Hash
 '(STRING)     Username
 
  Dim HashedPW As String * 20
  HashedPW = doubleHashPassword(ProfileStructs(Index).Password, _
                                ProfileStructs(Index).ClientToken, _
                                ProfileStructs(Index).ServerToken _
                                )

  With pbuffer(Index)
    .InsertDWORD ProfileStructs(Index).ClientToken
    .InsertDWORD ProfileStructs(Index).ServerToken
    .InsertNonNTString HashedPW
    .InsertNTString ProfileStructs(Index).Username
    .sendPacket Index, &H3A
  End With
End Sub

Private Sub Parse0x3A(Index As Integer, Data As String)
'SID_LOGONRESPONSE2
  Select Case Asc(Mid$(Data, OffsetDWORD(1), 1))
    Case &H0    '0x00: Success.
      Send0x0A Index, ProfileStructs(Index).Username
      Send0x0C Index, ProfileStructs(Index).Home
      Exit Sub
      
    Case &H1    '0x01: Account does not exist.
      AddChat Index, "   * Account does not exist.", fcGreen
      Send0x3D Index
      Exit Sub
      
    Case &H2    '0x02: Invalid password.
      AddChat Index, " > ", vbWhite
      AddChat Index, "Invalid Password.", fcRed, , NO_LINE
    
    Case &H6    '0x06: Account closed.
      AddChat Index, " > ", vbWhite
      AddChat Index, "Account Closed.", fcRed, , NO_LINE
    
  End Select
  mainForm.Disconnect Index
End Sub

Private Sub Send0x3D(Index As Integer)
'SID_CREATEACCOUNT2
  With pbuffer(Index)
    .InsertNonNTString hashPassword(ProfileStructs(Index).Password)
    .InsertNTString ProfileStructs(Index).Username
    .sendPacket Index, &H3D
  End With
  
End Sub

Private Sub Parse0x3D(Index As Integer, Data As String)
'SID_CREATEACCOUNT2
  '(UINT32) Status
  '(STRING) Account name suggestion
  
  '0x00: Account created
  '0x01: Name is too short
  '0x02: Name contained invalid characters
  '0x03: Name contained a banned word
  '0x04: Account already exists
  '0x05: Account is still being created
  '0x06: Name did not contain enough alphanumeric characters
  '0x07: Name contained adjacent punctuation characters
  '0x08: Name contained too many punctuation characters
  Select Case Asc(Mid$(Data, OffsetDWORD(1), 1))
    Case &H0
      AddChat Index, "   * Creation Successful.", fcGreen
      Send0x3A Index
      Exit Sub
    Case &H1
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username is too short.", fcRed, , NO_LINE
      Exit Sub
    Case &H2
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username contains invalid characters.", fcRed, , NO_LINE
      Exit Sub
    Case &H3
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username contains a banned word.", fcRed, , NO_LINE
      Exit Sub
    Case &H4
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username is taken.", fcRed, , NO_LINE
      Exit Sub
    Case &H5
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username is still being created?", fcYellow, , NO_LINE
      Exit Sub
    Case &H6
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username did not contain enough alphanumeric characters.", fcRed, , NO_LINE
      Exit Sub
    Case &H7
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username contained adjacent punctuation characters.", fcRed, , NO_LINE
      Exit Sub
    Case &H8
      AddChat Index, " > ", vbWhite
      AddChat Index, "Username contained too many punctuation characters.", fcRed, , NO_LINE
      Exit Sub
    
  End Select
  mainForm.Disconnect Index
End Sub

Public Sub Send0x50(Index As Integer)
'SID_AUTH_INFO
  Dim ProductCode As String
  
  If ProfileStructs(Index).SpoofClient = vbNullString Then
    Select Case ProfileStructs(Index).Client
      Case "Warcraft 2":          ProductCode = StrReverse("W2BN")
      Case "Starcraft":           ProductCode = StrReverse("STAR")
      Case "Starcraft BW":        ProductCode = StrReverse("SEXP")
      Case "Diablo":              ProductCode = StrReverse("DRTL")
      Case "Diablo II":           ProductCode = StrReverse("D2DV")
      Case "Diablo II LOD":       ProductCode = StrReverse("D2XP")
      Case "Warcraft 3":          ProductCode = StrReverse("WAR3")
      Case "Warcraft 3XP":        ProductCode = StrReverse("W3XP")
      Case "Starcraft Japan":     ProductCode = StrReverse("JSTR")
    End Select
  Else
    ProductCode = StrReverse(ProfileStructs(Index).SpoofClient)
  End If
  
  mainForm.wsTCP(Index).SendData ChrW$(1)
  With pbuffer(Index)
    .InsertDWORD &H0                ' Protocol ID
    .InsertNonNTString "68XI"       ' Platform Code
    .InsertNonNTString ProductCode  ' Product Code
    .InsertDWORD "&H4F"             ' Version Byte
    .InsertDWORD &H0                ' Language Code
    .InsertDWORD &H0                ' Local IP
    .InsertDWORD &H0                ' Time Zone Bias
    .InsertDWORD &H0                ' MPQ Local ID
    .InsertDWORD &H0                ' User Language ID
    .InsertNTString "USA"           ' Country Abbreviation
    .InsertNTString "United States" ' Country
    .sendPacket Index, &H50
  End With
End Sub

Private Sub Parse0x50(Index As Integer, Data As String)
'SID_AUTH_INFO

'On Error GoTo Err

' (UINT32)   Logon type
' (UINT32)   Server token
' (UINT32)   UDP value
' (FILETIME) MPQ filetime
' (STRING)   MPQ filename
' (STRING)   ValueString
'
'  WAR3/W3XP Only:
' (VOID)     128-byte Server signature

' xxxx|xxxx|xxxx|filetime|filename|valuestring

  Dim LogonType As Long
  'Dim ServerToken As Long
  Dim UDPValue As Long
  Dim MPQFileTime As String 'struct of 8 bytes ... double?
  Dim MPQFileName As String
  Dim ValueString As String

  Dim Splitter() As String

  LogonType = pbuffer(Index).GetDWORD(Mid$(Data, OffsetDWORD(1), 4))
  ProfileStructs(Index).ServerToken = pbuffer(Index).GetDWORD(Mid$(Data, OffsetDWORD(2), 4))
  UDPValue = pbuffer(Index).GetDWORD(Mid$(Data, OffsetDWORD(3), 4))
  MPQFileTime = Mid$(Data, OffsetDWORD(4), 8)
  
  Splitter = Split(Mid$(Data, OffsetDWORD(6)), vbNullChar)
  
  MPQFileName = Splitter(0)
  ValueString = Splitter(1)
  MPQFileName = Mid$(MPQFileName, InStr(1, MPQFileName, "-ix86-") + 6)
  
  'AddChat index, " * " & LogonType, fcPurple
  'AddChat Index, " * " & ProfileStructs(Index).ServerToken, fcPurple
  'AddChat index, " * " & UDPValue, fcPurple
  'AddChat Index, " * " & MPQFileTime, fcPurple
  'AddChat Index, " * " & MPQFileName, fcPurple
  'AddChat Index, " * " & ValueString, fcPurple
  'AddChat Index, " * " & Mid$(Data, InStr(1, Data, "-ix86-") + 6), fcPurple
  
  ' process these things?...
  
  '...
  
  Send0x51 Index, ValueString, MPQFileName
  
End Sub


Private Sub Send0x51(Index As Integer, ValueString As String, MPQFileName As String)
'SID_AUTH_CHECK
  Dim EXEVersion As Long
  'Dim ClientToken As Long
  Dim EXEInfo As String
  Dim sFile(2) As String
  Dim Checksum As Long
  Dim KeyCount As Long
  Dim ProductValue As Long
  Dim PublicValue As Long
  Dim KeyHash As String * 20
  Dim CDKey As String
  Dim keyLength As Long
  
  Dim mpqNumber As Long
  
  ' Get Hashes
  Select Case ProfileStructs(Index).Client
    Case "Warcraft 2"
      sFile(0) = "\Hashes\W2BN\Warcraft II BNE.exe"
      sFile(1) = "\Hashes\W2BN\Storm.dll"
      sFile(2) = "\Hashes\W2BN\Battle.snp"
    Case "Starcraft", "Starcraft BW"
      sFile(0) = "\Hashes\STAR\Starcraft.exe"
      sFile(1) = "\Hashes\STAR\Storm.dll"
      sFile(2) = "\Hashes\STAR\Battle.snp"
    Case "Diablo"
      sFile(0) = "\Hashes\DRTL\Diablo.exe"
      sFile(1) = "\Hashes\DRTL\Storm.dll"
      sFile(2) = "\Hashes\DRTL\Battle.snp"
    Case "Diablo II"
      sFile(0) = "\Hashes\D2DV\Game.exe"
      
    Case "Diablo II LOD"
      sFile(0) = "\Hashes\D2XP\Game.exe"
      
    Case "Warcraft 3", "Warcraft 3XP"
      sFile(0) = "\Hashes\war3\War3.exe"
      sFile(1) = "\Hashes\war3\Storm.dll"
      sFile(2) = "\Hashes\war3\Game.dll"
  End Select
  
  ' Get KeyCount
  KeyCount = GetKeyCount(ProfileStructs(Index).Client)
  
  ' Get KeyLength
  keyLength = Len(ProfileStructs(Index).CDKey)
  
  ' Get CDKey
  CDKey = ProfileStructs(Index).CDKey
  
  ' Client Token
  ProfileStructs(Index).ClientToken = GetTickCount + 7000
  
  ' Versioning
  EXEVersion = getExeInfo(App.Path & sFile(0), EXEInfo)
  Select Case ProfileStructs(Index).Client
    Case "Warcraft 2", _
         "Starcraft", "Starcraft BW", _
         "Diablo", "Diablo II", "Diablo II LOD", _
         "Warcrtaft 3", "Warcraft 3XP":
         
      If EXEVersion = 0 Then
        AddChat Index, " > ", vbWhite
        AddChat Index, "EXE versioning Failed.", fcRed, , NO_LINE
        mainForm.Disconnect Index
        Exit Sub
      End If
      
    Case Else
      '...
      
  End Select
  
  
  ' EXE Info Format (Separated by spaces)
  '  EXE Name (ex. war3.exe)
  '  Last Modified Date (ex. 08/16/09)
  '  Last Modified Time (ex. 19:21:59)
  '  Filesize in bytes (ex. 471040)
  
  If ProfileStructs(Index).Client <> "Diablo" Then    'hacky
  
  If kd_quick(CDKey, ProfileStructs(Index).ClientToken, ProfileStructs(Index).ServerToken, PublicValue, ProductValue, KeyHash, 20) = 0 Then
    AddChat Index, " > ", vbWhite
    AddChat Index, "Failed to decode CDKey. (" & CDKey & ")", fcRed, , NO_LINE
    
    mainForm.Disconnect Index
    Exit Sub
  End If
  
  End If
  
  
  ' Crunch exe hash
  mpqNumber = Val#(Left$(MPQFileName, InStr(1, LCase$(MPQFileName), ".mpq") - 1))
  checkRevision ValueString, App.Path & sFile(0), App.Path & sFile(1), App.Path & sFile(2), mpqNumber, Checksum
  
  With pbuffer(Index)
    .InsertDWORD ProfileStructs(Index).ClientToken      ' Client Token
    .InsertDWORD EXEVersion                             ' EXE Version
    .InsertDWORD Checksum                               ' EXE Hash                                x
    .InsertDWORD KeyCount                               ' Number of CDKeys
    .InsertDWORD &H0                                    ' Spawn Key (1 is TRUE, 0 is FALSE)
    
    ' Original Key
    .InsertDWORD keyLength                              ' Key Length
    .InsertDWORD ProductValue                           ' Key Product value                       x
    .InsertDWORD PublicValue                            ' Key Public value                        x
    .InsertDWORD 0                                      ' Unknown
    .InsertNonNTString KeyHash                          ' Hashed Key Data                         x
    
    If KeyCount = 2 Then
      ' Expansion Key
      .InsertDWORD keyLength                              ' Key Length
      .InsertDWORD ProductValue                           ' Key Product value                       x
      .InsertDWORD PublicValue                            ' Key Public value                        x
      .InsertDWORD 0                                      ' Unknown
      .InsertNonNTString KeyHash                          ' Hashed Key Data                         x
    End If
    
    .InsertNTString EXEInfo                             ' Exe Information
    .InsertNTString ProfileStructs(Index).KeyInUseMsg   ' Key Owner Name
    .sendPacket Index, &H51
  End With
  
End Sub

Private Sub Parse0x51(Index As Integer, Data As String)
'SID_AUTH_CHECK
  Select Case pbuffer(Index).GetDWORD(Mid$(Data, 5, 4))
    '0x000:    Passed challenge
    '0x100:    Old game version (Additional info field supplies patch MPQ filename)
    '0x101:    Invalid Version
    '0x102:    Game version must be downgraded (Additional info field supplies patch MPQ filename)
    '0x0NN:    (where NN is the version code supplied in SID_AUTH_INFO): Invalid version code (note that 0x100 is not set in this case).
    '0x200:    Invalid CD key *
    '0x201:    CD key in use (Additional info field supplies name of user)
    '0x202:    Banned Key
    '0x203:    Wrong Product
    Case &H0
      pbuffer(Index).InsertNonNTString "tenb"
      pbuffer(Index).sendPacket Index, &H14
      Send0x3A Index
      Exit Sub
      
    Case &H100
      AddChat Index, " > ", vbWhite
      AddChat Index, "Old game version.", fcRed, , NO_LINE
      
    Case &H101
      AddChat Index, " > ", vbWhite
      AddChat Index, "Invalid game version.", fcRed, , NO_LINE
      
    Case &H200
      AddChat Index, " > ", vbWhite
      AddChat Index, "Invalid CDKey. (Expect a temporary ban for this)", fcRed, , NO_LINE
      
    Case &H201
      Dim SS1() As String, sKeyInUseBy As String
      SS1() = Split(Data, vbNullChar)
      If UBound(SS1) > 2 Then sKeyInUseBy = SS1(3)
      AddChat Index, "   * CDKey in use by " & sKeyInUseBy, fcYellow
      
    Case &H202
      AddChat Index, " > ", vbWhite
      AddChat Index, "Banned CDKey.", fcRed, , NO_LINE
      
    Case &H203
      AddChat Index, " > ", vbWhite
      AddChat Index, "Wrong product CDKey.", fcRed, , NO_LINE
      
  End Select
  mainForm.Disconnect Index
End Sub

Private Sub Parse0x59(Index As Integer)
'SID_SETEMAIL
  AddChat Index, "   * Skipping Email registration.", fcYellow
End Sub

Public Sub Parse0x67(Index As Integer, Data As String)
'SID_FRIENDSADD

' (STRING) Account
' (UINT8)  Status         (0x01: Mutual, 0x02: DND, 0x04: Away)
' (UINT8)  Location id
' (UINT32) Product id
' (STRING) Location name

'Location ID
'0x00: Offline
'0x01: Not in chat
'0x02: In chat
'0x03: In public game
'0x04: In private game not on your friends list
'0x05: In private game is on your friends list
  
  'Chop header
  Data = Mid(Data, 5)
  
  'Account
  Dim Account As String
  Account = Left(Data, InStr(1, Data, vbNullChar) - 1)
  Data = Mid(Data, Len(Account) + 2)
  
  'Status
  Dim Status As Byte
  Dim Location As Byte
  Dim Product As String
  Dim LocName As String
  Status = MakeByte(Left$(Data, 1))
  Location = MakeByte(Mid$(Data, 2, 1))
  Product = Mid$(Data, 3, 4)
  LocName = Mid$(Data, 7)
  
  'For now just notify mutuality
  If Status = 1 Then
    AddChat Index, " > ", vbWhite
    AddChat Index, "Your friendship with " & Account & " is mutual!", fcOrange, , NO_LINE
  End If
  
End Sub

Public Sub Parse0x68(Index As Integer, Data As String)
'SID_FRIENDSREMOVE

' (UINT8) Entry number

' Meh
End Sub
















