VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================
' Packet Buffer
'==============================================================================================

Public strBuffer As String

Private RawData As String
Public Position As Long
Private Buffer As String

Public Sub sendPacket(Index As Integer, PPacketID As Byte)
On Error GoTo err
  Dim sPacketString As String
  
  ' Format packet names
  If Len(Hex(PPacketID)) = 1 Then
    sPacketString = "0" & Hex(PPacketID)
  Else
    sPacketString = Hex(PPacketID)
  End If
  
  ' Output control
  Select Case PPacketID
    Case &H14, &HA
      'Omit output
    Case Else
      If ProfileStructs(Index).iConnected = 1 And LenB(sPacketString) > 0 Then
        AddChat Index, "     Sent     -> 0x" & sPacketString, fcOrange
      End If
  End Select
   
  ' Packet Send
  mainForm.wsTCP(Index).SendData Chr$(&HFF) & Chr$(PPacketID) & MakeWORD(Len(Buffer) + 4) & Buffer
  Buffer = vbNullString
  strBuffer = vbNullString
  Exit Sub
  
err:
  AddChat Index, " > ", vbWhite
  AddChat Index, err.Number & " : " & err.Description, fcRed, NO_LINE
  
End Sub

Public Function InsertDWORD(Data As Long)
  Buffer = Buffer & MakeDWORD(Data)               ' Concatenates data as dword
End Function

Public Function InsertBYTE(Data As Integer)
  Buffer = Buffer & Chr$(Data)                    ' Concatenates as character (one byte)
End Function

Public Function InsertNTString(Data As String)
  Buffer = Buffer & Data & vbNullChar             ' Concatenate with null terminator
End Function

Public Function InsertNonNTString(Data As String)
  Buffer = Buffer & Data                          ' Concatenate w/o null terminator
End Function

Private Function MakeDWORD(Value As Long) As String
  Dim Result As String * 4                        ' Converts value to DWORD as string
  CopyMemory ByVal Result, Value, 4
  MakeDWORD = Result
End Function

Private Function MakeWORD(Value As Integer) As String
  Dim Result As String * 2                        ' Converts value to WORD as string
  CopyMemory ByVal Result, Value, 2
  MakeWORD = Result
End Function

Public Function GetDWORD(Data As String) As Long
  Dim lReturn As Long                             ' Converts string to DWORD as long
  Call CopyMemory(lReturn, ByVal Data, 4)
  GetDWORD = lReturn
End Function

