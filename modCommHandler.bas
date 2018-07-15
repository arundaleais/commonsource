Attribute VB_Name = "modCommHandler"
'Only Comm
Option Explicit

Public Const DEFAULT_HANDLER_INDEX = 0

Private Type CommDef    'must be private
    BaudRate As String
    Name As String      'Short name COM6 (to see if weve changed it)
    AutoBaud As Boolean
    VCP As String       'Associated VCP if any
End Type

Public Type SocketDef   'AisDecoder version
    Fragment As String  'Partial senetence received on this socket
                        'CFLF not received
'    Buffer(1 To Qmax) As AdrDef  'Full when 1 spare slot
'    Qrear As Long   'added to rear
'    Qfront As Long  'Points to 1 AHEAD of next one to be removed
'    QLost As Long
    DevName As String   'Determined by the User
                        'PORT_number
                        'COMMnumber
                        'Name
                        'TEXTBOX_number
    Handler As Long     'The Type of handler
'These are the same is the selected cboHandler.ListIndex
                        '-1 = not determined
                        '0 = Winsock
                        '1 = COMM
                        '2 = File
                        '3 = TTY
                        '4 = Loopback
    Hidx As Long         'Handler Array index comms(Hidx), Winsock(Hidx)
    State As Integer    'The State returned by the Handler
                        '-1 =  Handler Object (inc DCB) Index = Nothing
                        '0 = Handler Object Index exists (but Port Closed)
                        '11 = Handler Opening (or trying to) Open Port
                        '1  = Port Open
                        '18 = Handler Closing (or trying to) Close Port
                        '21 = Serial Data Loss
                        '22 = Serial Pending Output buffer not empty
    errmsg As String    'Normally the error reported by the handler
    Enabled As Boolean  'Set by the user to suspend this socket
    MsgCount As Double
    LastMsgCount As Long    'used by speed timer
'    Graph As Boolean    'Set by user to enable/disable Graph series
    LostMsgCount As Double
    TryCount As Long
    Chrs As Long        'Since last Speed time
    ResetCount As Long      'No of times no chrs received on each reconnect
'    Routes() As Routedef      'Route index
                        'Only created if Idx less than Ridx, as its
                        'included for one Idx irrespective of Direction
'    Forwards() As Long    'Socket number of destination
                        'Allocated sequentially
    Direction As Long   '0=both Input,Output
                        '1=Input Only
                        '2=Output Only
'    OutputFormat As OutputFormatDef
'    IEC As IecFormatDef
'    Winsock As WinsockDef
    Comm As CommDef
'    File As FileDef
'    Recorder As RecorderDef
'    VDO As VdoDef
'    AisFilter As clsAisFilter
End Type

'from modRouter
Public CurrentSocket As Long    'Used if Commcfg is defined
Public sockets() As SocketDef   'Array of data

'Comm only
'Opens the handler for a socket if it is enabled
Public Sub OpenHandler(Idx As Long)
Dim kb As String
    If sockets(Idx).Enabled = False Then
        sockets(Idx).State = 0
        Exit Sub
    End If

'When a call is made to Open a handler Incr the TryCount
    sockets(Idx).TryCount = sockets(Idx).TryCount + 1

'Only log if not a re-open by
    Select Case sockets(Idx).TryCount
    Case Is = 1
        WriteLog "Opening Handler for " & sockets(Idx).DevName _
        & ", Current state is " & aState(sockets(Idx).State), LogForm
    Case Is = 2
        WriteLog "Trying to open " & sockets(Idx).DevName _
        & ", Current state is " & aState(sockets(Idx).State), LogForm
    Case Else       'Supress if Try
'        WriteLog "to open " & Sockets(Idx).DevName _
'        & ", Current state is " _
'        & aState(Sockets(Idx).State), Idx
    End Select


'    Call DisplayTries("OpenHandler - Enter", Idx)
'State must be set to 11 to prevent File reading first record until timer is started
    sockets(Idx).State = 11 'Trying to open handler
    
    Select Case sockets(Idx).Handler
'Sockets().Hidx is set by the handler (-1) if cant open
    Case Is = 0            '0 = Winsock
'        Call CreateWinsock(CInt(Idx))
'        If sockets(Idx).Hidx > 0 Then
'            sockets(Idx).State = frmRouter.Winsock(sockets(Idx).Hidx).State
'            If sockets(Idx).State = sckListening Then
'            End If
'        Else
'            sockets(Idx).State = 0  'Closed
'        End If
    Case Is = 1           '1 = COMM
'Hidx needs to be -1 to create a new Comm or Open and existing one
        Call CreateComm(Idx)
        If sockets(Idx).Hidx > 0 Then
            sockets(Idx).State = Comms(sockets(Idx).Hidx).State
'        Else
'            Sockets(Idx).State = 0  'Closed
        End If
    Case Is = 2            '2 = File
'        Call CreateFile(Idx)
'        If sockets(Idx).Hidx > 0 Then
'            sockets(Idx).State = Files(sockets(Idx).Hidx).State
'        Else
'Sockets(idx).state should be set in create file
'        End If
    Case Is = 3            '3 = TTY
'        Call CreateTTY(Idx)
'Hidx is left as 11 if trying to open TTYs
'        If sockets(Idx).Hidx > 0 Then
'            sockets(Idx).State = TTYs(sockets(Idx).Hidx).State
'        End If
    Case Is = 4            '4 = LoopBack
'        Call CreateLoopBack(Idx)
'        If sockets(Idx).Hidx > 0 Then
'            sockets(Idx).State = LoopBacks(sockets(Idx).Hidx).State
'        End If
    Case Else
        MsgBox "Handler " & sockets(Idx).Handler & " not found", , "Close Profile"
    End Select
    
    WriteLog sockets(Idx).DevName & " [" _
    & aHandler(sockets(Idx).Handler) _
    & " handler] is " & aState(sockets(Idx).State)

'If at the end of the Open Handler call the when the Sockets(Idx).Status
'is 1 (Open) decr the Try Count
'MsgBox WinsockState(Idx)
    
'    Select Case sockets(Idx).State
'    Case Is = sckOpen, sckListening, sckConnected
'        sockets(Idx).TryCount = sockets(Idx).TryCount - 1
'        If sockets(Idx).TryCount > 0 Then
'            Call frmRouter.ResetTries(Idx)
'        End If
'    End Select
    
'end of trying to open handler
End Sub


'Only Comm
Public Sub CloseHandler(ReqIdx As Long)
Dim Idx As Long
Dim Hidx As Long
Dim kb As String
'MsgBox "CloseHandler " & ReqIdx

    For Idx = 1 To UBound(sockets)
        If Idx = ReqIdx Or ReqIdx = 0 Then
            sockets(Idx).errmsg = ""
            Hidx = sockets(Idx).Hidx
            If sockets(Idx).State <> -1 Then
                If Hidx > 0 Then
'Handler has been allocated

'Dont remove the Forwards - but dont re-create when OpenHandler is called
'                    kb = Routecfg.SocketForwards(ReqIdx, "Remove")
'If jnasetup = True Then
'    MsgBox kb, , "CloseHandler-Remove (" & Idx & ")"
'End If
                    Select Case sockets(Idx).Handler
                    Case Is = 0            '0 = Winsock
'                        Call CloseWinsock(CInt(sockets(Idx).Hidx))
                    Case Is = 1           '1 = COMM
'Stop
'                        Set Comms(sockets(Idx).Hidx).DCB = Nothing
                        Set Comms(sockets(Idx).Hidx) = Nothing
                    Case Is = 2            '2 = File
'                        Call Files(sockets(Idx).Hidx).CloseFile
'Added v26 (same as when profile is closed)
'                        Set Files(sockets(Idx).Hidx) = Nothing
                    Case Is = 3            '3 = TTY
'                        Call TTYs(sockets(Idx).Hidx).TTYClose
                    Case Is = 4            '4 = LoopBack
'                        Set LoopBacks(sockets(Idx).Hidx) = Nothing
                    Case Else
                        MsgBox "Handler " & sockets(Idx).Handler & " not found", , "Close Profile"
                    End Select
                                        
                    WriteLog "Closed " & sockets(Idx).DevName _
                    & " " & aHandler(sockets(Idx).Handler) _
                    & " handler"
'Remove link to handler
                    sockets(Idx).Hidx = 0
                End If
                    
'set state before ResetTries
'                sockets(Idx).State = sckClosed
'                If sockets(Idx).TryCount > 0 Then
'                    Call frmRouter.ResetTries(Idx)
'                End If
            End If
'Reset Clears all values from Sockets() so values
'must be recreated (as New) or re-loaded from Registry
            If ReqIdx = 0 Then
                Call ClearSocket(Idx)
            End If
        End If

    Next Idx
Exit Sub

Reset_error:
    Select Case err.Number
    Case Is = 10
        Exit Sub
    Case Else
MsgBox "Error " & err.Number & " - " & err.Description, , "CloseHandler"
    End Select
'Stop
End Sub

'Only Comm
Public Sub ClearSocket(Idx As Long)
Dim i As Long
            
            With sockets(Idx)
'                If .TryCount > 0 Then
'                    Call frmRouter.ResetTries(Idx)
'                End If
'                Call RemoveRecorder(Idx)
                .Fragment = ""
'                For i = 1 To Qmax
'                   .Buffer(i).Data = ""
'                   .Buffer(i).Source = 0
'                    .Buffer(i).UtcUnix = 0
'                Next i
                'Full when 1 spare slot
'                .Qrear = 0
'                .Qfront = 0
'                .QLost = 0
                .DevName = "Connection " & Idx
                .Handler = DEFAULT_HANDLER_INDEX
                .Hidx = 0
                .State = -1
                .errmsg = ""
                .Enabled = False
                .MsgCount = 0
'                .Graph = False
                .LostMsgCount = 0
                .Chrs = 0
                .ResetCount = 0
'                ReDim .Routes(1 To 1)
'                ReDim .Forwards(1 To 1)
                .Direction = -1      'undefined
'                .Winsock.Server = -1   '-1=undefined
'                .Winsock.Protocol = sckUDPProtocol
'                .Winsock.RemoteHost = ""
'                .Winsock.RemotePort = ""
'                .Winsock.LocalPort = ""
'                .Winsock.RemoteHostIP = ""
'                ReDim .Winsock.Streams(1 To 1)
'                .Winsock.Oidx = -1
'                .Winsock.Sidx = -1
'                .Winsock.PermittedStreams = 1
'                .Winsock.PermittedIPStreams = 1
                .Comm.BaudRate = ""
                .Comm.AutoBaud = False
                .Comm.Name = ""
                .Comm.VCP = ""
'                .File.SocketFileName = ""
'                .File.ReadRate = 0
'                .File.RollOver = False
'                .Recorder.Enabled = False
'                .VDO.SequenceNo = 0
'                .VDO.Source = 0
'                .VDO.UtcUnix = 0
'                .VDO.Destination = 0
'                .VDO.Data = ""
'                .VDO.LastVdoUpdate = 0
'                If Not .Recorder.Output Is Nothing Then
'                    Unload .Recorder.Output
'                End If
            End With
End Sub



