Attribute VB_Name = "modAis"
'These are routines common to AisDecoder and NmeaRouter
'In particular they include the IO routines and functions
Option Explicit
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

