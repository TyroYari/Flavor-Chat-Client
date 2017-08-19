Attribute VB_Name = "SharedData"
'==============================================================================================
' Shared Data
'
' Author: Ben Pogrund
'
'==============================================================================================
Option Explicit

' --- Client info --- '
Public Type ProfileStruct
  'Configured Fields
  Username As String
  Password As String
  Client As String
  Server As String
  CDKey As String
  XPKey As String
  Role As String
  Home As String
  
  'Runtime variables
  UniqName As String
  iConnected As Integer
  CurrChannel As String
  bNotifySet As Boolean
  LastChannel As String
  SpoofClient As String
  'Adv Config
  bBNLS As Boolean
  BNLSServ As String
  PingResponse As Byte
  EmailReg As String
  ProfileTitle As String
  ClientSpoof As String
  
  'Bot Properties
  GreetMsg As String
  GreetOn As Boolean
  IdleMsg As String
  IdleOn As Boolean
  IdleTimer As Long
  IdleElapsed As Long
  KeyInUseMsg As String
  Trigger As String
  
  'Authentication Variables
  ClientToken As Long
  ServerToken As Long
  
End Type

Public ProfileStructs() As ProfileStruct
Public ProfileCount As Long
Public CurrentProfIndx As Integer
Public CurrentEdtrIndx As Long
Public CurrentEdtrFile As String

' --- Functionality --- '
Public PreviousInputs() As String
Public PrevInputIndex As Integer

Public AccessList() As UserAccess
Public Type UserAccess
  Name As String
  Level As Integer
End Type

' --- UI Variables --- '
Public Const NO_LINE = 1
Public Const BRING_TO_FRONT = 0
Public Const SEND_TO_BACK = 1
Public Const MAIN_TAB = 0
Public uniqChID As Long
Public bFlashOrange As Boolean

' --- Colors --- '
Public Const BB = 65536
Public Const GG = 256!  ' Type as single to prevent overflow
Public Const RR = 1
Public Const fcGreen As Long = (42 * BB) + (178 * GG) + (20 * RR) '        RGB(42, 178, 20)
Public Const fcBrightGreen As Long = (88 * BB) + (247 * GG) + (70 * RR) '  RGB(70,247,88)
Public Const fcBlue As Long = (231 * BB) + (145 * GG) + (43 * RR) '        RGB(43, 145, 231)
Public Const fcYellow As Long = (123 * BB) + (247 * GG) + (255 * RR) '     RGB(255, 247, 123)
Public Const fcOrange As Long = (0 * BB) + (185 * GG) + (255 * RR) '       RGB(255, 185, 0)
Public Const fcRed As Long = (83 * BB) + (20 * GG) + (255 * RR) '          RGB(255, 20, 83)
Public Const fcPurple As Long = (184 * BB) + (67 * GG) + (184 * RR) '      RGB(184, 67, 225)
Public Const fcGray As Long = (210 * BB) + (210 * GG) + (210 * RR) '       RGB(210, 210, 210)
Public Const fcDarkGray As Long = (128 * BB) + (128 * GG) + (128 * RR) '   RGB(128, 128, 128)

' --- Connection Vars --- '
Public bMassConnect As Boolean
Public iMassConnectIndex As Integer

' --- Editor Vars --- '
Public FileStatus(3) As Integer
Public Const EDITOR_FILE_OPENED = 2
Public Const EDITOR_FILE_EDITED = 1
Public Const EDITOR_FILE_CLOSED = 0

' --- Version Macro --- '
Public VersionMacro As String
' "/me " & "Version: " & App.Major & "." & App.Minor & "." & App.Revision

' --- Permissions: Access Level Constants --- '
Public Const AL_ADD = 100
Public Const AL_REMOVE = 100
Public Const AL_OP = 95
Public Const AL_BLACKLIST = 90
Public Const AL_BAN = 85
Public Const AL_RECONNECT = 85
Public Const AL_SAY2 = 75
Public Const AL_JOIN = 60
Public Const AL_SPOOF = 55
Public Const AL_WHITELIST = 50
Public Const AL_KICK = 40
Public Const AL_REJOIN = 35
Public Const AL_SETHOME = 35
Public Const AL_HOME = 30
Public Const AL_SAY1 = 20
Public Const AL_VER = 10
Public Const AL_TRIGGER = 5





























