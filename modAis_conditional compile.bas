Attribute VB_Name = "modAis"
'These are routines common to AisDecoder and NmeaRouter
'In particular they include the IO routines and functions
Option Explicit
#Const IO = -1
#Const Serial = -1

'#If Serial Then
Private Type CommDef    'must be private
    BaudRate As String
    Name As String      'Short name COM6 (to see if weve changed it)
    VCP As String       'Associated VCP if any
End Type
'#End If

'#If IO Then
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
    
'#If Serial Then
    Comm As CommDef
'#End If
'    File As FileDef
'    Recorder As RecorderDef
'    VDO As VdoDef
'    AisFilter As clsAisFilter
End Type
'from modRouter
Public CurrentSocket As Long    'Used if Commcfg is defined
Public sockets() As SocketDef   'Array of data
'#End If


Public MAX_SOCKETS As Long '= 10

Public Function aHandler(Handler As Long) As String
    Select Case Handler
    Case Is = -1
        aHandler = "none"
    Case Is = 0
        aHandler = "TCP/IP"
    Case Is = 1
        aHandler = "Serial"
    Case Is = 2
        aHandler = "File"
    Case Is = 3
        aHandler = "TTY"
    Case Is = 4
        aHandler = "Loop back"
    Case Else
        aHandler = "Invalid"
    End Select
End Function
 
Public Function aDirection(Direction As Long) As String
    Select Case Direction
    Case Is = 0
        aDirection = "Input,Output"
    Case Is = 1
        aDirection = "Input"
    Case Is = 2
        aDirection = "Output"
    Case Is = -2
        aDirection = "undefined"
    End Select
End Function

Public Function aServer(Server As Long) As String
    Select Case Server
    Case Is = -1
        aServer = "Undefined"
    Case Is = 0
        aServer = "Client"
    Case Is = 1
        aServer = "Server"
    Case Else
        aServer = "Invalid"
    End Select
End Function

Public Function aReadRate(ReadRate As Long) As String
    Select Case ReadRate
    Case Is = 0     'unlimited
        aReadRate = "Unlimited"
    Case Is = 1
        aReadRate = "1 Sentence/Minute"
    Case Is = 2
        aReadRate = "10 Sentences/Minute"
    Case Is = 3
        aReadRate = "50 Sentences/Minute"
    Case Is = 4
        aReadRate = "100 Sentences/Minute"
    Case Is = 5
        aReadRate = "500 Sentences/Minute"
    Case Is = 6
        aReadRate = "1000 Sentences/Minute"
    Case Is = 7
        aReadRate = "5000 Sentences/Minute"
    Case Is = 8
        aReadRate = "50000 Sentences/Minute"
    Case Else
        aReadRate = "Undefined"
    End Select
End Function

Public Function aProtocol(Protocol As Integer)
    Select Case Protocol
    Case Is = 0     'sckTCPProtocol - not defined if winsock not in use
        aProtocol = "TCP"
    Case Is = 1     'sckUDPProtocol - not defined if winsock not in use
        aProtocol = "UDP"
    End Select
End Function

Public Function aState(State As Integer) As String
    Select Case State
    Case Is = -1
        aState = "Nothing"
    Case Is = 0
        aState = "Closed"
    Case Is = 1
        aState = "Open"
    Case Is = 2
        aState = "Listening"
    Case Is = 3
        aState = "Connection pending"
    Case Is = 4
        aState = "Resolving host"
    Case Is = 5
        aState = "Host resolved"
    Case Is = 6
        aState = "Connecting"
    Case Is = 7
        aState = "Connected"
    Case Is = 8
        aState = "Peer is closing connection"
    Case Is = 9
        aState = "Error"
    Case Is = 11
        aState = "Opening"
    Case Is = 18
        aState = "Closing"
    Case Is = 21
        aState = "Data loss"
    Case Is = 22
        aState = "Data in buffer"
    Case Else
        aState = "Invalid"
    End Select
End Function

Public Function aEnabled(Enabled As Boolean) As String
    If Enabled = True Then
        aEnabled = "Enabled"
    Else
        aEnabled = "Disabled"
    End If
End Function

Public Sub CloseHandler(Idx As Long)
    MsgBox "CloseHandler (" & Idx & ")"
End Sub

