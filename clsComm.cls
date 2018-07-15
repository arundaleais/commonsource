VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'22Feb   Added CommBuffer
' clsDCB - Device Communication Block utility class
' Part of the Desaware API Class Library
' Copyright (c) 1996 by Desaware Inc.
' All Rights Reserved
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Private Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type ADDINFO    'additional info
'    Index As Long
    hName As String
'    ErrMsg As String
    State As Long      '0 = closed, 1=Opem (Only polled when state is 1)
    Hidx As Long    'Comm Array Index(So we know the Socket when data rcvd)
    Idx As Long      'Sockets index (So we know the Socket when data rcvd)
    AutoBaudRate As Boolean 'Turn on NMEA baudRate detection
End Type

' Private members
Private timeouts As COMMTIMEOUTS
Private handle As Long  ' Comm handle
Private DevName$ ' Com1, com2 or other compatible comm device

' Public members
Public DCB As clsDCB

'JNA Private members
Public PartSentence As String
'Will need to be public to access with a Forwarder module

Private Info As ADDINFO

'JNA Public members
'Public StatusBlock As COMMSTATUS
'Public Name As String

' Current state indicators
' Holds output data that arrives while an output transfer is in progress
Private PendingOutput$
Private CurrentEventMask&   ' Non zero if events are being watched for
' Buffers for overlapped input and output
' Must take this approach due to VB's ability to move strings
Private CurrentInputBuffer&
Private CurrentOutputBuffer&
Private TriggeredEvents&    ' Variable to load with event results

' Three overlapped structures,
' 0 = read, 1 = write, 2 = waitevent
Private overlaps(2) As OVERLAPPED
' Indicates background operation is in progress
Private inprogress(2) As Boolean
' Amount of data transferred on write
Private closeinprogress As Boolean
Private DataWritten&
Private DataRead&
Private EventResults&   'New

' This object must have two functions
' CommInput(dev As clsComm, info As String)
' CommEvent(dev As clsComm, event as long)
Private CallbackObject As Object

' Declarations
Private Declare Function apiSetCommTimeouts Lib "kernel32" Alias "SetCommTimeouts" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function apiGetCommTimeouts Lib "kernel32" Alias "GetCommTimeouts" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Private Declare Function GetCommModemStatus Lib "kernel32" (ByVal hFile As Long, ByRef lpModemStat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyFromBuffer Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal Buffer As Long, ByVal iMaxLength As Long) As Long
Private Declare Function lstrcpyToBuffer Lib "kernel32" Alias "lstrcpynA" (ByVal Buffer As Long, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
'new
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, ByVal l As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
'Not used added back by jna to handle embedded nulls
'Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal dwBytes As Long)
'My mod - byval
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal dwBytes As Long)
'Added for strin manipulation routines
Private Declare Function lstrlenA Lib "kernel32" Alias "lstrlen" (ByVal lpString As String) As Long
Private Declare Function lstrlenW Lib "kernel32" Alias "lstrlen" (ByVal lpString As String) As Long
'added to get system messages when failed
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const INVALID_HANDLE_VALUE = -1
Private Const GMEM_FIXED = &H0
Private Const ClassBufferSizes% = 1024
Private Const ERROR_IO_PENDING = 997 '  dderror
Private Const ERROR_OPERATION_ABORTED = 995
Private Const WAIT_TIMEOUT = &H102&

' GetCommModemStatus flags
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&

' Error values
Private Const CLASS_NAME$ = "clsComm"
Private Const ERR_NOCOMMACCESS = 31010
Private Const ERR_UNINITIALIZED = 31011
Private Const ERR_MODEMSTATUS = 31012
Private Const ERR_READFAIL = 31013
'new
Private Const ERR_EVENTFAIL = 31014

'New
Private Const EV_RXCHAR = &H1
Private Const EV_RXFLAG = &H2
Private Const EV_TXEMPTY = &H4
Private Const EV_CTS = &H8
Private Const EV_DSR = &H10
Private Const EV_RLSD = &H20
Private Const EV_BREAK = &H40
Private Const EV_ERR = &H80
Private Const EV_RING = &H100
Private Const EV_PERR = &H200
Private Const EV_RX80FULL = &H400
Private Const EV_EVENT1 = &H800
Private Const EV_EVENT2 = &H1000

'New
Private Const CE_RXOVER = &H1
Private Const CE_OVERRUN = &H2
Private Const CE_RXPARITY = &H4
Private Const CE_FRAME = &H8
Private Const CE_BREAK = &H10
Private Const CE_TXFULL = &H100


' An empty string with a single null character
Private EmptyString As String * 1

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Private Sub Class_Initialize()
Dim olnum%

'debugserial function is defined in ModRouter but not used in AisDecoder
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsComm; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    
    Set DCB = New clsDCB
    CurrentInputBuffer = GlobalAlloc(GMEM_FIXED, ClassBufferSizes + 1)
    CurrentOutputBuffer = GlobalAlloc(GMEM_FIXED, ClassBufferSizes + 1)
    CurrentEventMask = EV_ERR   'new
    EmptyString = Chr$(0)
    ' Create event objects for the background transfer
    For olnum = 0 To 2
        overlaps(olnum).hEvent = CreateEvent(0, True, False, vbNullString)
    Next olnum
End Sub

Private Sub Class_Terminate()
    Dim olnum
' Close existing comm device
'Set DCB = Nothing   ' Be sure DCB is free (Must close before CloseComm)
    Call CloseComm
    ' Dump the event objects
'Debug.Print "Terminate " & inprogress(0) & inprogress(1) & inprogress(2)
    For olnum = 0 To 2
        Call CloseHandle(overlaps(olnum).hEvent)
        inprogress(olnum) = False
    Next olnum
    Set DCB = Nothing   ' Be sure DCB is free
    Call GlobalFree(CurrentInputBuffer)
    Call GlobalFree(CurrentOutputBuffer)
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub

' Useful error routines
Private Sub DeviceNotOpenedError()
    Call CloseComm
'Debug.Print "DeviceNotOpenedError"
Exit Sub
    err.Raise vbObjectError + ERR_UNINITIALIZED, CLASS_NAME, "Communication Device is not open"
End Sub

Private Sub ModemStatusError()
    err.Raise vbObjectError + ERR_MODEMSTATUS, CLASS_NAME, "GetCommModemStatus Failed"
End Sub


'-----------------------------------------------
' Timeout property access follows
'-----------------------------------------------
Public Property Get ReadIntervalTimeout() As Long
    ReadIntervalTimeout = timeouts.ReadIntervalTimeout
End Property

Public Property Let ReadIntervalTimeout(vNewValue As Long)
    timeouts.ReadIntervalTimeout = vNewValue
End Property

Public Property Get ReadTotalTimeoutMultiplier() As Long
    ReadTotalTimeoutMultiplier = timeouts.ReadTotalTimeoutMultiplier
End Property

Public Property Let ReadTotalTimeoutMultiplier(vNewValue As Long)
    timeouts.ReadTotalTimeoutMultiplier = vNewValue
End Property


Public Property Get ReadTotalTimeoutConstant() As Long
    ReadTotalTimeoutConstant = timeouts.ReadTotalTimeoutConstant
End Property

Public Property Let ReadTotalTimeoutConstant(vNewValue As Long)
    timeouts.ReadTotalTimeoutConstant = ReadTotalTimeoutConstant
End Property

Public Property Get WriteTotalTimeoutMultiplier() As Long
    WriteTotalTimeoutMultiplier = timeouts.WriteTotalTimeoutMultiplier
End Property

Public Property Let WriteTotalTimeoutMultiplier(vNewValue As Long)
    timeouts.WriteTotalTimeoutMultiplier = WriteTotalTimeoutMultiplier
End Property

Public Property Get WriteTotalTimeoutConstant() As Long
    WriteTotalTimeoutConstant = timeouts.WriteTotalTimeoutConstant
End Property

Public Property Let WriteTotalTimeoutConstant(vNewValue As Long)
    timeouts.WriteTotalTimeoutConstant = WriteTotalTimeoutConstant
End Property

' The device handle is read only
Public Property Get hCommDev() As Long
    hCommDev = handle
End Property

' This property is read only
Public Property Get DeviceName() As String
    DeviceName = DevName
End Property

'JNA Properties
Public Property Get State() As Long
    State = Info.State
End Property

Public Property Let State(vNewValue As Long)
    Info.State = vNewValue
End Property

Public Property Get hIndex() As Long
    hIndex = Info.Hidx
End Property

Public Property Let hIndex(vNewValue As Long)
    Info.Hidx = vNewValue
End Property

Public Property Get sIndex() As Long
    sIndex = Info.Idx
End Property

Public Property Let sIndex(vNewValue As Long)
    Info.Idx = vNewValue
End Property

Public Property Get Name() As String
    Name = Info.hName
End Property

Public Property Let Name(vNewValue As String)
    Info.hName = vNewValue
End Property

Public Property Get AutoBaudRate() As Boolean
    AutoBaudRate = Info.AutoBaudRate
End Property

Public Property Let AutoBaudRate(vNewValue As Boolean)
    Info.AutoBaudRate = vNewValue
End Property

Public Property Get PendingOutputLen() As Double
    PendingOutputLen = Len(PendingOutput)
End Property

#If False Then
'complie error
Public Property Get errmsg() As String
    errmsg = StatusBlock.Name
End Property

Public Property Let errmsg(vNewValue As String)
    StatusBlock.errmsg = vNewValue
End Property
#End If

Public Property Get Closing() As Boolean
    Closing = closeinprogress
End Property

Public Property Let Closing(vNewValue As Boolean)
    Closing = vNewValue
End Property


Public Sub GetCommTimeouts()
    ' Is there any real need to report errors here?
    If handle = 0 Then Exit Sub
    Call apiGetCommTimeouts(handle, timeouts)
End Sub

Public Function SetCommTimeouts() As Long
    If handle = 0 Then Exit Function ' Returns false
    SetCommTimeouts = apiSetCommTimeouts(handle, timeouts) <> 0
End Function

' The main function for opening a comm device
' Receives device name (com?) and optionally the size of the internal input and output queues
Public Function OpenComm(CommDeviceName As String, Notify As Object, Optional cbInQueue, Optional cbOutQueue) As Long
Dim Idx As Long

    ' Close an existing port when reopening
    WriteLog "Opening " & Mid$(CommDeviceName, 5), LogForm
    On Error GoTo OpenComm_error
    If handle <> 0 Then CloseComm
    DevName = CommDeviceName
    Set CallbackObject = Notify
    handle = CreateFile(DevName, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
    If handle = INVALID_HANDLE_VALUE Then
'Exit Function
'MsgBox GetLastSystemError
        err.Raise ERR_NOCOMMACCESS, CLASS_NAME, "Unable to open communications device"
    End If
    ' If the input and output queue size is specified, set it now
    If Not (IsMissing(cbInQueue) Or IsMissing(cbOutQueue)) Then
        Call SetupComm(handle, cbInQueue, cbOutQueue)
    Else
        Call SetupComm(handle, 8192, 1024)  'old 4096
    End If
    ' Ok, we've got the comm port. Initialize the timeouts
    GetCommTimeouts
    ' Set some default timeouts
    timeouts.ReadIntervalTimeout = 1
    timeouts.ReadTotalTimeoutMultiplier = 0 'old 1
    timeouts.ReadTotalTimeoutConstant = 10  'old 1
    timeouts.WriteTotalTimeoutMultiplier = 1
    timeouts.WriteTotalTimeoutConstant = 1
    SetCommTimeouts
    ' Initialize the DCB to the current device parameters
    Call DCB.GetCommState(Me)
    Call SetCommMask(handle, CurrentEventMask)  'new
'Now set in StartInput
'    StartInput '(Poll calls StartInput)
    Exit Function
OpenComm_error:
'The error number is passed back to Commcfg & reported by it
    OpenComm = err.Number   'return error
'Idx = CurrentSocket changed to sindex
    Idx = sIndex
    sockets(Idx).State = -1
    sockets(Idx).errmsg = err.Description
    If sockets(Idx).Hidx > 0 Then    'Its possible for the handler to have closed the socket
        Set Comms(sockets(Idx).Hidx) = Nothing   'Remove the Comm Socket if closed
        sockets(Idx).Hidx = -1
    End If
'calling another function clears the err. messages
    WriteLog "Open failed with error " & err.Number & " " & err.Description, LogForm
End Function    'will clear err.


' Close the comm port (Previous settings are retained)
Public Function CloseComm() As Long
'    frmRouter.PollTimer.Enabled = False
    Me.State = 0
'    Me.sIndex = 0
'Clear the Pending buffer
    PendingOutput = ""
'If Already closed, just exit
'Debug.Print "CloseComm Handle=" & handle
    If handle > 0 Then
'Debug.Print "CloseComm " & inprogress(0) & inprogress(1) & inprogress(2)
        Set CallbackObject = Nothing
        Call CloseHandle(handle)
        handle = 0
'Must set the socket state to 0 so that the reconnect timer '
'tries to reconnect, which it will only do if closed
'V33 SocketState should only be changed on return from this call
'There is a possibility the socket has been closed and this
'CloseComm has been called by the poll_timer
        On Error Resume Next
            sockets(Me.sIndex).State = 0    'v19
        On Error GoTo 0
    End If
'    frmRouter.PollTimer.Enabled = True
    WriteLog Mid$(DevName, 5) & " closed", LogForm
End Function

Public Function DisableComm()
    Call CloseComm
End Function

#If False Then
Public Function EnableComm_notused()
Dim ret As Long
Dim Hidx As Long
    On Error GoTo EnableComm_error
    Call DisplayComms("EnableComm Start")
    If handle <= 0 Then
        ret = OpenComm("\\.\" & Name, frmRouter)
        If ret <> 0 Then
            err.Raise ret, "EnableComm", sockets(CurrentSocket).errmsg
        End If
    End If
    Me.State = 1
    Call DisplayComms("EnableComm Finish")
Exit Function

EnableComm_error:
'Stop
'This will clear the error     On Error GoTo 0
    Select Case err.Number
    Case Is = 31010
        err.Description = err.Description & vbCrLf & "Unable to open " & Name
'        Set Comms(hIndex) = Nothing
    Case Is = 380
        err.Description = err.Description & vbCrLf & "Invalid Baud Rate " & DCB.BaudRate
'        Set Comms(hIndex) = Nothing
    Case Else
    End Select
    MsgBox err.Number & " " & err.Description, , "Enable Comm Error"
'Dont unload the form so that user has to cancel or enter a valid port
    
Stop
End Function
#End If

' This is another entry to retreive the comm state
' Note how it handles the problem of DCB needing the
' clsComm object parameter
Public Function GetCommState() As Long
    If handle = 0 Then DeviceNotOpenedError
    GetCommState = DCB.GetCommState(Me)
End Function

' This is another entry to retreive the comm state
' Note how it handles the problem of DCB needing the
' clsComm object parameter
Public Function SetCommState() As Long
    If handle = 0 Then DeviceNotOpenedError
    SetCommState = DCB.SetCommState(Me)
End Function

' Here are some easy functions to determine the current
' modem status

Public Property Get CTS_ON()
    Dim modemstatus&
    Dim Res&
    If handle = 0 Then DeviceNotOpenedError
    Res = GetCommModemStatus(handle, modemstatus)
    If Res = 0 Then ModemStatusError
    CTS_ON = (modemstatus And MS_CTS_ON) <> 0
End Property

Public Property Get DSR_ON()
    Dim modemstatus&
    Dim Res&
    If handle = 0 Then DeviceNotOpenedError
    Res = GetCommModemStatus(handle, modemstatus)
    If Res = 0 Then ModemStatusError
    DSR_ON = (modemstatus And MS_DSR_ON) <> 0
End Property

Public Property Get RING_ON()
    Dim modemstatus&
    Dim Res&
    If handle = 0 Then DeviceNotOpenedError
    Res = GetCommModemStatus(handle, modemstatus)
    If Res = 0 Then ModemStatusError
    RING_ON = (modemstatus And MS_RING_ON) <> 0

End Property

Public Property Get RLSD_ON()
    Dim modemstatus&
    Dim Res&
    If handle = 0 Then DeviceNotOpenedError
    Res = GetCommModemStatus(handle, modemstatus)
    If Res = 0 Then ModemStatusError
    RLSD_ON = (modemstatus And MS_RLSD_ON) <> 0
End Property

'To deal with the Serial Data Loss a better way would be to set a larger
'Output queue for serial comms and buffer the data in this queue
'rather than moving the data to comm output and then having to have
'a big stream buffer (PendingOutput) here
'There qould not be a need to poll for output
'as this would be governed by the Deque, which would alter its
'rate of output dependant on how quickly commoutput cleared the data
'If there was any poutput pending, it would not deque any more data
'If any data had to be discarded it could be in complete sentences
'by not queueing them in the first place.
'If commoutput returns false (if unsuccesfull) to Deque then the item would
'not be removed from the queue
Public Function CommOutput(outputdata As String) As Long
    Dim bytestosend&
    Dim Res&
    Dim kb As String
    Dim b() As Byte
    Dim lpsz As Long
    Dim ret As Long
    Static lostcount As Long
    
    State = 1       'Must be to be here
    If handle = 0 Then DeviceNotOpenedError
'Output data can be blant because WriteComplete calls CommOutput
'with a null string to get CommOutput set to true
'Debug.Print Left$(outputdata, 3) & " " & Len(outputdata)
    If outputdata <> "" Then
        If PendingOutputLen + Len(outputdata) < MAX_COMM_OUTPUT_BUFFER_SIZE Then
            PendingOutput = PendingOutput & outputdata
            State = 1
        Else
            lostcount = lostcount + 1
'frmRouter.StatusBar.Panels(1).Text = lostcount
'            State = 21
        End If
        
    End If
    
    If inprogress(1) Then    ' Write operation is in progress
        CommOutput = True
        Exit Function
    End If

'Debug.Print Len(PendingOutput)
   
'With a reserve of global memory of 20 it the output gets truncated
'so ive allowed 40 (I think is still can at 40)
DoEvents
    If Len(PendingOutput) < ClassBufferSizes - 100 Then
        bytestosend = Len(PendingOutput)
        Call lstrcpyToBuffer(CurrentOutputBuffer&, PendingOutput$ & vbNull, bytestosend + 1)
        PendingOutput = ""
    Else
        bytestosend = ClassBufferSizes - 1 - 100
        Call lstrcpyToBuffer(CurrentOutputBuffer&, Left$(PendingOutput$, bytestosend) & vbNull, bytestosend + 1)
        PendingOutput = Mid$(PendingOutput, bytestosend + 1)
    End If
    
    If bytestosend > 0 Then
        Res = WriteFile(handle, CurrentOutputBuffer, bytestosend, DataWritten, overlaps(1))
' Debug.Print bytestosend
        If Res <> 0 Then
            ProcessWriteComplete
            CommOutput = True
        Else
            If err.LastDllError = ERROR_IO_PENDING Then
                inprogress(1) = True
                CommOutput = True
            End If
        End If
    End If
End Function

' Restart the next output operation if necessary
Public Sub ProcessWriteComplete()
    inprogress(1) = False
    Call CommOutput("")
End Sub

' Called periodically
Public Sub PollWrite()
    Dim Res&
    If Not inprogress(1) Then Exit Sub
    ' Check the event
    Res = WaitForSingleObject(overlaps(1).hEvent, 0)
    ' If not yet signaled, just exit
    If Res = WAIT_TIMEOUT Then Exit Sub
    ' Data was written - Try writing any pending data
    ProcessWriteComplete
End Sub

' This function enables or disables data transfer
Private Sub StartInput()
    Dim Res&
    Dim errors&
    ' Read already in progress
    State = 1
    If inprogress(0) Then
        Exit Sub
    End If
    If handle = 0 Then
'Exit Sub
        DeviceNotOpenedError
    End If
    Res = ReadFile(handle, CurrentInputBuffer, ClassBufferSizes, DataRead, overlaps(0))
    If Res <> 0 Then
        ProcessReadComplete
    Else
        Select Case err.LastDllError
        Case Is = ERROR_IO_PENDING
            inprogress(0) = True
            'Debug.Print "pended read"
        Case Is = ERROR_OPERATION_ABORTED
'This is normally caused by closing the socket but a
'read operation has become pending (from the polling)
 '           Debug.Print "read aborted"
'Clear the error and continue
            Call ClearCommError(handle, errors, 0)
'Read will never be completed
            inprogress(0) = False
        Case Is > 0
'v19 MsgBox Err.LastDllError & vbCrLf & GetLastSystemError
'we appear to come here when weve stopped the input for a time
'We come here when we pull out a USB device
'            Call CloseComm
'Exit Sub
'            Err.Raise vbObjectError + ERR_READFAIL, CLASS_NAME, "Failure on Comm device read operation"
Call CloseComm
'Call OpenComm
'Set Comms(Me.Index) = Nothing
        End Select
    End If
End Sub

Public Sub PollRead()
    Dim Res&
    If Not inprogress(0) Then
        StartInput
        Exit Sub
    End If
        
    ' Check the event
    Res = WaitForSingleObject(overlaps(0).hEvent, 0)
    ' If not yet signaled, just exit
    If Res = WAIT_TIMEOUT Then Exit Sub
    ' Data was written - Try writing any pending data
    ProcessReadComplete
End Sub

Public Sub ProcessReadComplete()
Dim resstring$
Dim copied&

    If inprogress(0) Then ' Was overlapped
        DataRead = overlaps(0).InternalHigh
        inprogress(0) = False
    End If
    If DataRead <> 0 Then
'Debug.Print "Read " & "(" & Me.Name & ") " & DataRead & " bytes"
'Debug.Print "Read " & inprogress(0) & inprogress(1) & inprogress(2)
'Create Buffer (resstring) of reqired length full of nulls
        resstring$ = String$(DataRead + 1, 0)
'The return value is the address of a temp buffer
'which is no loner valid with vb
   '     copied = lstrcpyFromBuffer(resstring, CurrentInputBuffer, DataRead + 1)
        Call lstrcpyFromBuffer(resstring, CurrentInputBuffer, DataRead + 1)
        Call CommBuffer(resstring)
    End If
End Sub

'This is event driven when data is received
'The Data MUST be buffered in this socket before being
'sent to be displayed or forwarded
Private Sub CommBuffer(commdata As String)
Dim cpos%                            '
Dim Is8BitAscii As Boolean  'NMEA should only be 7 bit ascii

If Left$(commdata$, 1) = Chr$(10) Then
'Debug.Print "LF commdata"
End If
If Left$(PartSentence, 1) = Chr$(10) Then
'Debug.Print "LF PartSentence"
End If
'Add on any previously received partial sentence
    commdata = PartSentence & commdata
    PartSentence = ""
    
    If commdata <> "" Then
'Debug.Print thiscomm.DeviceName
'commdata$ always has a NULL appended by CommRead
'MsgBox thiscomm.ChrCtrl(commdata$), , "CommInput (commdata$)"
'        txtTerm.SelStart = Len(txtTerm.Text)
        ' Substitute the CR with a CRLF pair, dump the LF
        Do Until Len(commdata$) = 0
            cpos% = InStr(commdata$, Chr$(13))
            If cpos% > 0 Then   'Complete sentence
     '           txtTerm.SelText = Left$(commdata$, cpos% - 1) & vbCrLf
'chrctrl added to try and see if any non ascii characters are being output
'If Left$(commdata$, 1) = Chr$(10) Then
'                Call CommRcv(thiscomm.ChrCtrl(Left$(commdata$, cpos% - 1)) & vbCrLf, thiscomm.Index)
'End If
        If Not (CallbackObject Is Nothing) Then
            Call CallbackObject.CommRcv(Left$(commdata$, cpos% - 1) & vbCrLf, hIndex)
        End If
'Call back                Call CommRcv(Left$(commdata$, cpos% - 1) & vbCrLf, Index)
                commdata$ = Mid$(commdata$, cpos% + 1)
                cpos% = InStr(commdata$, Chr$(10))
                If cpos% > 0 Then
                    commdata$ = Mid$(commdata$, cpos% + 1)
                End If
             Else           'No CR
                cpos% = InStr(commdata$, Chr$(10))
'We probably ought to replace a LF Null with CRLF
                If cpos% > 0 Then   'But has LF, Keep LF + NULL
                    commdata$ = Mid$(commdata$, cpos% + 1)
                Else                'No CR or LF but will have last NULL
'Output the whole buffer when a terminating Null
'                    txtTerm.SelText = commdata$ 'includes last NULL
'except the Null
'save the partial sentence for this socket
                    PartSentence = PartSentence _
                    & Left$(commdata$, Len(commdata$) - 1)
'                    Call CommRcv(Left$(commdata$, Len(commdata$) - 1), thiscomm.Index)
'Code to AutoBaudRate
                    If AutoBaudRate = True Then
'Check to see if bit(7) is 1 (chrno >= 128) - NMEA is only 7 bit
                        For cpos = 1 To Len(PartSentence)
                            If Asc(Mid$(PartSentence, cpos, 1)) >= 128 Then
                                Is8BitAscii = True
                            End If
                        Next cpos
                        If Is8BitAscii Then
                            Call RetryBaudRate
                        End If
                    End If
                    commdata$ = ""
                End If
            End If
        Loop
'        If Len(txtTerm.Text) > 4096 Then
'            txtTerm.Text = Right$(txtTerm.Text, 2048)
'        End If
    End If
End Sub


#If False Then
Private Sub StartEventWatch()   'new
    Dim Res&
    If inprogress(2) Then Exit Sub
    If handle = 0 Then DeviceNotOpenedError
    EventResults = 0
    Res = WaitCommEvent(handle, EventResults, overlaps(2))
    If Res <> 0 Then
        ProcessEventComplete
    Else
        If GetLastError() = ERROR_IO_PENDING Then
            inprogress(2) = True
            #If DEBUGMODE Then
'               Debug.Print "pended event"
            #End If
        Else
            err.Raise vbObjectError + ERR_EVENTFAIL, CLASS_NAME, "Failure on Comm device event test operation"
        End If
    End If
End Sub

Private Sub ProcessEventComplete()  'new
    Dim errors&
    If inprogress(2) Then
        inprogress(2) = False
    End If
    
    If EventResults <> 0 Then
        #If DEBUGMODE Then
'         Debug.Print "Event value " & Hex$(EventResults)
         #End If
        If Not (CallbackObject Is Nothing) Then
            Call ClearCommError(handle, errors, 0)
                
'            If (errors And CE_RXOVER) <> 0 Then Call CallbackObject.CommEvent(Me, "Receive Queue Full Error")
'            If (errors And CE_OVERRUN) <> 0 Then Call CallbackObject.CommEvent(Me, "Receive Overrun Error")
'            If (errors And CE_RXPARITY) <> 0 Then Call CallbackObject.CommEvent(Me, "Receive Parity Error")
'            If (errors And CE_FRAME) <> 0 Then Call CallbackObject.CommEvent(Me, "Frame Error")
'            If (errors And CE_BREAK) <> 0 Then Call CallbackObject.CommEvent(Me, "Break Detected")
'            If (errors And CE_TXFULL) <> 0 Then Call CallbackObject.CommEvent(Me, "Transmit Queue Full")
        End If
    End If
End Sub

Private Sub PollEvent() 'new
    Dim Res&
    If Not inprogress(2) Then
        StartEventWatch
        Exit Sub
    End If
        
    Res = WaitForSingleObject(overlaps(2).hEvent, 0)
    If Res = WAIT_TIMEOUT Then Exit Sub
    ProcessEventComplete
End Sub
#End If

Public Sub PollClose()
    Call CloseComm
End Sub
' Test results on all background processes
Public Function Poll() As Boolean
    If Not Closing Then
        If handle > 0 Then
'frmRouter.StatusBar.Panels(2).Text = PendingOutputLen
            PollWrite
            PollRead
            Poll = True
        Else
'stop
        End If
'       PollEvent   'new
    Else
        CloseComm
    End If
'frmRouter.StatusBar.Panels(3).Text = State
End Function

Public Function ChrCtrl(ByRef Line As String) As String
Dim kb As String
Dim Outbuf As String
Dim i As Long
Dim Chrno As Long
Dim b() As Byte

    b = StrConv(Line, vbFromUnicode)
'    If Not b Is Nothing Then
        For i = 0 To UBound(b)
            If b(i) >= 32 And b(i) <= 127 Then
            Outbuf = Outbuf & Chr$(b(i))
        Else
            Outbuf = Outbuf & "<" & b(i) & ">"
        End If
        Next i
'    Else
'        DpyBuf = "<empty>"
'    End If
    ChrCtrl = Outbuf
End Function

    Function BSTRtoLPWSTR(sBSTR As String, b() As Byte, lpwsz As Long) As Long
     
    ' Input: a nonempty BSTR string
    ' Input: **undimensioned** byte array b()
    ' Output: Fills byte array b() with Unicode char string from sBSTR
    ' Output: Fills lpwsz with a pointer to b() array
    ' Returns byte count, not including terminating 2-byte Unicode null character
    ' Original BSTR is not affected
     
    Dim cBytes As Long
     
    cBytes = LenB(sBSTR)
     
    ' ReDim array, with space for terminating null
    ReDim b(1 To cBytes + 2) As Byte
     
    ' Point to BSTR char array
    lpwsz = StrPtr(sBSTR)
     
    ' Copy the array
    CopyMemory b(1), ByVal lpwsz, cBytes + 2
     
    ' Point lpsz to new array
    lpwsz = VarPtr(b(1))
     
    ' Return byte count
    BSTRtoLPWSTR = cBytes
     
    End Function

'The function to convert a BSTR to an LPSTR is similar, but requires a translation from Unicode to ANSI first:
    Function BSTRtoLPSTR(sBSTR As String, b() As Byte, lpsz As Long) As Long
     
    ' Input: a nonempty BSTR string
    ' Input: **undimensioned** byte array b()
    ' Output: Fills byte array b() with ANSI char string
    ' Output: Fills lpsz with a pointer to b() array
    ' Returns byte count, not including terminating null
    ' Original BSTR is not affected
     
    Dim cBytes As Long
    Dim sABSTR As String
     
    cBytes = LenB(sBSTR)
     
    ' ReDim array, with space for terminating null
    ReDim b(1 To cBytes + 2) As Byte
     
    ' Convert to ANSI
    sABSTR = StrConv(sBSTR, vbFromUnicode)
     
    ' Point to BSTR char array
    lpsz = StrPtr(sABSTR)
     
    ' Copy the array
    CopyMemory b(1), ByVal lpsz, cBytes + 2
     
    ' Point lpsz to new array
    lpsz = VarPtr(b(1))
     
    ' Return byte count
    BSTRtoLPSTR = cBytes
     
    End Function

    Function LPWSTRtoBSTR(ByVal lpwsz As Long) As String
     
    ' Input: a valid LPWSTR pointer lpwsz
    ' Return: a sBSTR with the same character array
     
    Dim cChars As Long
     
    ' Get number of characters in lpwsz
    cChars = lstrlenW(lpwsz)
     
    ' Initialize string
    LPWSTRtoBSTR = String$(cChars, 0)
     
    ' Copy string
    CopyMemory ByVal StrPtr(LPWSTRtoBSTR), ByVal lpwsz, cChars * 2
     
    End Function

    Function LPSTRtoBSTR(ByVal lpsz As Long) As String
     
    ' Input: a valid LPSTR pointer lpsz
    ' Output: a sBSTR with the same character array
     
    Dim cChars As Long
     
    ' Get number of characters in lpsz
    cChars = lstrlenA(lpsz)
     
    ' Initialize string
    LPSTRtoBSTR = String$(cChars, 0)
     
    ' Copy string
    CopyMemory ByVal StrPtr(LPSTRtoBSTR), ByVal lpsz, cChars
     
    ' Convert to Unicode
    LPSTRtoBSTR = Trim(StrConv(LPSTRtoBSTR, vbUnicode))
     
    End Function

Public Function GetLastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Dim sError As String * 500 '\\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long
Dim lCount As Long

lErrNum = err.LastDllError
lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lErrMsg, 0&, sError, Len(sError), ByVal 0)
    If lCount Then
        sError = Left$(sError, lCount - 2)    'Remove line feeds
    End If

GetLastSystemError = Trim(sError)

End Function

Private Function RetryBaudRate()
Dim ret As Long
Dim NextBaudRate As Long

    On Error GoTo BaudRate_error
        
    Select Case sockets(Me.sIndex).Comm.BaudRate
    Case Is = 1200
        NextBaudRate = 2400
    Case Is = 2400
        NextBaudRate = 4800
    Case Is = 4800
        NextBaudRate = 9600
    Case Is = 9600
        NextBaudRate = 19200
    Case Is = 19200
        NextBaudRate = 38400
    Case Else
        NextBaudRate = 4800
    End Select
    
    WriteLog sockets(Me.sIndex).DevName & " trying " & NextBaudRate & " baud next", LogForm
    sockets(Me.sIndex).Comm.BaudRate = NextBaudRate
    With Comms(sockets(Me.sIndex).Hidx)
        .DCB.BaudRate = NextBaudRate
        If .DCB.BaudRate <> sockets(Me.sIndex).Comm.BaudRate Then
            err.Raise 380, "RetryBaudRate"
        End If
    
        ret = .SetCommState
'Open was successful
        If ret = -1 Then
            .State = 1
        End If
    End With
Exit Function

BaudRate_error:
'v45 changed from Sockets(Comms(Hidx).sIndex) to Hidx as Sockets() may not exist
    MsgBox "Retry Baud Rate - error " & Str(err.Number) & ", " & err.Description & vbCrLf _
    & "Next Baud Rate = " & NextBaudRate, , "RetryBaudRate"
End Function

