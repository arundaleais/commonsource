Attribute VB_Name = "modComm"
Option Explicit
' Communication Constants
Global Const NOPARITY = 0
Global Const ODDPARITY = 1
Global Const EVENPARITY = 2
Global Const MARKPARITY = 3
Global Const SPACEPARITY = 4
Global Const ONESTOPBIT = 0
Global Const ONE5STOPBITS = 1
Global Const TWOSTOPBITS = 2
Global Const IGNORE = 0
Global Const INFINITE = &HFFFF
Global Const CE_RXOVER = &H1
Global Const CE_OVERRUN = &H2
Global Const CE_RXPARITY = &H4
Global Const CE_FRAME = &H8
Global Const CE_BREAK = &H10
Global Const CE_CTSTO = &H20
Global Const CE_DSRTO = &H40
Global Const CE_RLSDTO = &H80
Global Const CE_TXFULL = &H100
Global Const CE_PTO = &H200
Global Const CE_IOE = &H400
Global Const CE_DNS = &H800
Global Const CE_OOP = &H1000
Global Const CE_MODE = &H8000
Global Const IE_BADID = (-1)
Global Const IE_OPEN = (-2)
Global Const IE_NOPEN = (-3)
Global Const IE_MEMORY = (-4)
Global Const IE_DEFAULT = (-5)
Global Const IE_HARDWARE = (-10)
Global Const IE_BYTESIZE = (-11)
Global Const IE_BAUDRATE = (-12)
Global Const EV_RXCHAR = &H1
Global Const EV_RXFLAG = &H2
Global Const EV_TXEMPTY = &H4
Global Const EV_CTS = &H8
Global Const EV_DSR = &H10
Global Const EV_RLSD = &H20
Global Const EV_BREAK = &H40
Global Const EV_ERR = &H80
Global Const EV_RING = &H100
Global Const EV_PERR = &H200
Global Const EV_CTSS = &H400
Global Const EV_DSRS = &H800
Global Const EV_RLSDS = &H1000
Global Const SETXOFF = 1
Global Const SETXON = 2
Global Const SETRTS = 3
Global Const CLRRTS = 4
Global Const SETDTR = 5
Global Const CLRDTR = 6
Global Const RESETDEV = 7
Global Const GETMAXLPT = 8
Global Const GETMAXCOM = 9
Global Const GETBASEIRQ = 10
Global Const CBR_110 = &HFF10
Global Const CBR_300 = &HFF11
Global Const CBR_600 = &HFF12
Global Const CBR_1200 = &HFF13
Global Const CBR_2400 = &HFF14
Global Const CBR_4800 = &HFF15
Global Const CBR_9600 = &HFF16
Global Const CBR_14400 = &HFF17
Global Const CBR_19200 = &HFF18
Global Const CBR_38400 = &HFF1B
Global Const CBR_56000 = &HFF1F
Global Const CBR_57600 = &HFF20
Global Const CBR_128000 = &HFF23
Global Const CBR_256000 = &HFF27
Global Const CN_RECEIVE = &H1
Global Const CN_TRANSMIT = &H2
Global Const CN_EVENT = &H4
Global Const CSTF_CTSHOLD = &H1
Global Const CSTF_DSRHOLD = &H2
Global Const CSTF_RLSDHOLD = &H4
Global Const CSTF_XOFFHOLD = &H8
Global Const CSTF_XOFFSENT = &H10
Global Const CSTF_EOF = &H20
Global Const CSTF_TXIM = &H40
Global Const LPTx = &H80

' Application constants
' The size of the input and output buffers we will use
Public Const MAX_COMMS = 10
Public Const MAX_COMM_OUTPUT_BUFFER_SIZE = 50000
Public Const BufferSize% = 2048
Public Comms() As clsComm    'Array of Objects
Private PollTimerForm As Object 'this is passed to clsComm when the class is created
                            'It is used to Notify the form the call back is to
                            '(form has the poll timer)
Public CommPollTimer As Control 'This must exist on frmRouter,NmeaRcv or Form1
'as PollTimer. Must be able to be enabled in Commcfg and
Public BaudRateForm As Object
Public BaudRateComboBox As Control

' The port configuration for the demo
'jna not used Public Dialing% ' Currently dialing

'If Hidx is 0 then a new handler is created
'Else set the existing one
'Puts the Hidx into Sockets(Idx).Hidx if successful
'Else -1
Public Sub CreateComm(Idx As Long)
Dim ret As Long
Dim ctrl As Control

    WriteLog "Creating Serial Socket " & sockets(Idx).DevName & " [" & sockets(Idx).Comm.Name & "]", LogForm
    On Error GoTo CreateComm_error
'On Error GoTo 0     'debug
    
'If the handler is not closed - close it
'We have to do it first as Close will set Hidx to 0
    If sockets(Idx).Hidx > 0 Then
        Call CloseHandler(Idx)
    End If

    If sockets(Idx).Hidx <= 0 Then
'When first opened HIDX = 0
        sockets(Idx).Hidx = FreeComm
        If sockets(Idx).Hidx = -1 Then
            MsgBox "Cant create Comm handler"
            Exit Sub
        End If
    End If
            
'If first time, find location of Poll Timer and BaudRateComboBox (if any)
    If PollTimerForm Is Nothing Then
        For Each PollTimerForm In Forms
            For Each ctrl In PollTimerForm
                If TypeOf ctrl Is Timer And ctrl.Name = "PollTimer" Then
                    Set CommPollTimer = ctrl
                    Exit For
                End If
            Next ctrl
            If Not CommPollTimer Is Nothing Then Exit For
        Next PollTimerForm
    End If
    
'Must always have defined a PollTimer in some form
    If PollTimerForm Is Nothing Then
        MsgBox "Poll Timer Control not Found", , "modComm.CreateComm"
        Exit Sub
    End If
        
    Call DisableCommPollTimer   'stop all comm polling during create
    
    sockets(Idx).errmsg = ""

'Create the Comms(index) control if required
    If UBound(Comms) < sockets(Idx).Hidx Then ReDim Preserve Comms(sockets(Idx).Hidx)
    
    Call EnableCommPollTimer   'start all comm polling
    
'Are we opening an existing Comm
    If Not Comms(sockets(Idx).Hidx) Is Nothing Then
        
            WriteLog "Using " & aHandler(sockets(Idx).Handler) & " Handler " & sockets(Idx).Hidx, LogForm
'Stop polling this socket
        
'        Comms(Sockets(Idx).Hidx).State = 0
'        Call Comms(Sockets(Idx).Hidx).CloseComm

' Comm is already valid. Have we changed ports?
'Name is my name added to Comm() - the Handler name
'Sockets()Name is the New name
        If sockets(Idx).Comm.Name <> Comms(sockets(Idx).Hidx).Name Then
                    
            ' We're changing device
            ' Note that this also serves to close and
            ' release the previous comm object (with dwClass_terminate)
            Set Comms(sockets(Idx).Hidx) = New clsComm
'State will be 0 when first created
'Must be set immediately after creation to to report errors to Router(Idx)
'Note the index in Comms is the Cmms index
'            Comms(Hidx).hIndex = Hidx
'This must be set before the Comm port is opened
'Otherwise data can be received and another comm port will be opened
'We need to keep Idx on the hanler so that when data is received we know which
'Socket its from
            Comms(sockets(Idx).Hidx).sIndex = Idx
            ' This demo doesn't use buffer sizes
'            Ret = Comms(Sockets(Idx).Hidx).OpenComm("\\.\" & Sockets(Idx).Comm.Name, frmRouter)
            ret = Comms(sockets(Idx).Hidx).OpenComm("\\.\" & sockets(Idx).Comm.Name, PollTimerForm)
        End If
        ' If device is unchanged, SetCommState is all that is needed
    Else
        WriteLog aHandler(sockets(Idx).Handler) & " Handler " _
& sockets(Idx).Hidx & " allocated to " & sockets(Idx).Comm.Name, LogForm
        Set Comms(sockets(Idx).Hidx) = New clsComm

'State will be 0 when first created
'Must be set immediately after creation to to report errors to Router(Idx)
        Comms(sockets(Idx).Hidx).hIndex = sockets(Idx).Hidx
'This must be set before the Comm port is opened
'Otherwise data can be received and another comm port will be opened
        Comms(sockets(Idx).Hidx).sIndex = Idx
'        Ret = Comms(Sockets(Idx).Hidx).OpenComm("\\.\" & Sockets(Idx).Comm.Name, frmRouter)
        ret = Comms(sockets(Idx).Hidx).OpenComm("\\.\" & sockets(Idx).Comm.Name, PollTimerForm)
    End If
    
    If ret <> 0 Then
        err.Raise ret, "CreateComm", sockets(Idx).errmsg
    End If
    
    sockets(Idx).Comm.AutoBaud = True
    
'Required baud rate
    Comms(sockets(Idx).Hidx).DCB.BaudRate = sockets(Idx).Comm.BaudRate
    If Comms(sockets(Idx).Hidx).DCB.BaudRate <> sockets(Idx).Comm.BaudRate Then
        err.Raise 380, "DCB"
    End If
    Comms(sockets(Idx).Hidx).DCB.fNull = True
    Comms(sockets(Idx).Hidx).DCB.fErrorChar = False     'Dont replace parity errors with ErrorChar
    Comms(sockets(Idx).Hidx).DCB.ErrorChar = "~"
    Comms(sockets(Idx).Hidx).DCB.ByteSize = 8
    'Comm.DCB.Parity = 1   'odd, for example (0 = none)
    ' Perform any other DCB setting here

'My settings
'We have to set up the Handler index in Sockets here as this
'is used on the return to Socketcfg to see if weve been successful
'in setting up the Comm array
'Now set in Socketcfg    Sockets(CurrentSocket).Hidx = txtHidx
    Comms(sockets(Idx).Hidx).sIndex = Idx
    Comms(sockets(Idx).Hidx).Name = sockets(Idx).Comm.Name 'Short device name
    Comms(sockets(Idx).Hidx).AutoBaudRate = sockets(Idx).Comm.AutoBaud
'Only 1 destination set at the moment
'    Comms(txtHidx).Destination(1) = CLng(txtForward.Text)
    ' Now record the configuration changes
    
'Comm is now not opened until Sockets is exiting (may be disabled)
    ret = Comms(sockets(Idx).Hidx).SetCommState
'Open was successful
    If ret = -1 Then
        Comms(sockets(Idx).Hidx).State = 1
    End If
#If False Then
'trying to debug comms open
    If ret = 0 Then
    Call Comms(sockets(Idx).Hidx).GetLastSystemError
        Comms(sockets(Idx).Hidx).State = 1
    Else
Stop
        Comms(sockets(Idx).Hidx).State = 1
    Call Comms(sockets(Idx).Hidx).GetLastSystemError
    End If
#End If
'I'm not sure how Commcfg can be open but as CreateComm is used
'by both AisDecoder and NmeaRouter but Commcfg is only used
'in NmeaRouter (AisDecoder sets BaudRate in NmeaRcv) vwe callot
'just unload Commcfg as the program will not compile
    If Not IsFormLoaded("Commcfg") Is Nothing Then Unload IsFormLoaded("Commcfg")
    Call EnableCommPollTimer
    Exit Sub

CreateComm_error:
    sockets(Idx).State = 9  'Error
    sockets(Idx).errmsg = err.Description
'Stop
'This will clear the error     On Error GoTo 0
    Select Case err.Number
    Case Is = 31010
'already removed when terminating        Set Comms(Sockets(Idx).Hidx) = Nothing
        err.Description = err.Description & vbCrLf & "Unable to open " & sockets(Idx).Comm.Name
    Case Is = 380
'ditto        Set Comms(Sockets(Idx).Hidx) = Nothing
        err.Description = err.Description & vbCrLf & "Invalid Baud Rate " & sockets(Idx).Comm.BaudRate
    Case Else
    End Select
    WriteLog "Create Comm Error " & err.Number & " " & err.Description, LogForm
'Dont unload the form so that user has to cancel or enter a valid port
    
'restart polling as there may be other Comm ports in use (NmeaRouter)
    Call EnableCommPollTimer

End Sub

Public Sub DisplayComms(Optional Caption As String)
Dim result As Boolean
Dim Idx As Long
Dim Hidx As Variant
Dim kb As String
Dim Count As Long
Dim cPorts As New Collection
Dim Ports() As String
Dim k As Long
Dim myPort As Variant
Dim PortName As String
Dim NoPCComm As Boolean
Dim NoPCCommPort As Boolean
Dim PortCount As Long
Dim VCPPort As String
Dim VCPInstalledDir As String

    Ports = GetSerialPorts      'registry method
    On Error Resume Next
    PortCount = UBound(Ports) + 1
    On Error GoTo 0
    
    If PortCount > 0 Then
        For k = 0 To UBound(Ports)  'should be same as NameValueCount
            cPorts.Add Ports(k), Ports(k)
        Next k
    
        VCPInstalledDir = QueryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\com0com", "Install_Dir")
        If VCPInstalledDir <> "" Then
            kb = "VCP Driver is installed" & vbCrLf
        Else
            kb = "VCP Driver is not installed" & vbCrLf
        End If
        
        For Hidx = 1 To UBound(Comms)
            If Not Comms(Hidx) Is Nothing Then
                If sockets(Comms(Hidx).sIndex).State <> -1 Then
                    If Count = 0 Then
                        kb = kb & "PC Comm Ports in use by " & App.EXEName & vbCrLf
                    End If
                    kb = kb & Comms(Hidx).Name & " (Socket " & Hidx & ")" & vbCrLf
'Must be Comms name NOT Device Name                     VCPPort = GetVCP(Sockets(Comms(Hidx).sIndex).DevName)
                    VCPPort = GetVCP(Comms(Hidx).Name)
                    kb = kb & vbTab & FriendlyName(Comms(Hidx).Name) & vbCrLf
                    If VCPPort <> "" Then
                        kb = kb & vbTab & "Linked Virtual Comm Port " & VCPPort & vbCrLf
                    End If
                    kb = kb & vbTab & "Connection = " & sockets(Comms(Hidx).sIndex).DevName & vbCrLf
                    kb = kb & vbTab & "Baud Rate = " & Comms(Hidx).DCB.BaudRate & vbCrLf
                    kb = kb & vbTab & "State = " & aState(sockets(Comms(Hidx).sIndex).State) & vbCrLf
'                   kb = kb & vbTab & "Error = " & Comms(Hidx).ErrMsg & vbCrLf
                    On Error GoTo NoPort
                    cPorts.Remove Comms(Hidx).Name
                    On Error GoTo 0
                    Count = Count + 1
                End If
            End If
        Next Hidx
                
        If Count = 0 Then
            kb = kb & "There are no PC Comm Ports in use by " & App.EXEName & vbCrLf
        End If
    Else
        kb = "There are no PC Serial Ports" & vbCrLf
    End If
    
    kb = kb & vbCrLf
'Display any Ports on  the PC which do not have a Serial Handler
'allocated
    If cPorts.Count > 0 Then
        If cPorts.Count = 1 Then
            kb = kb & "There is " & cPorts.Count & " PC Comm Port"
        Else
            kb = kb & "There are " & cPorts.Count & " PC Comm Ports"
        End If
        kb = kb & " not in use by " & App.EXEName & vbCrLf
        For Each myPort In cPorts
            kb = kb & myPort
            If FriendlyName(CStr(myPort)) <> "" Then
                kb = kb & vbTab & FriendlyName(CStr(myPort))
            End If
            kb = kb & vbCrLf
'Must be Comms name NOT Device Name                     VCPPort = GetVCP(Sockets(Comms(Hidx).sIndex).DevName)
            VCPPort = GetVCP(CStr(myPort))
            If VCPPort <> "" Then
                kb = kb & vbTab & "Linked Virtual Comm Port " & VCPPort & vbCrLf
            End If
        Next
    End If
    
'Display Serial sockets with no handler allocated
'because the Serial Socket is not on this PC
    For Idx = 1 To UBound(sockets)
        If sockets(Idx).State <> -1 Then
            If sockets(Idx).Handler = 1 Then
                If sockets(Idx).Hidx = -1 Then
                    If NoPCCommPort = False Then
                        kb = kb & vbCrLf & "PC Comm Ports which cannot be opened by " & App.EXEName & vbCrLf
                        NoPCCommPort = True
                    End If
                    If sockets(Idx).Comm.Name <> "" Then
                        kb = kb & sockets(Idx).Comm.Name & " (No socket)" & vbCrLf
                        If FriendlyName(sockets(Idx).Comm.Name) <> "" Then
                            kb = kb & vbTab & FriendlyName(sockets(Idx).Comm.Name) & vbCrLf
'Must be Comms name NOT Device Name                     VCPPort = GetVCP(Sockets(Comms(Hidx).sIndex).DevName)
                            VCPPort = GetVCP(sockets(Idx).Comm.Name)
                            If VCPPort <> "" Then
                                kb = kb & vbTab & "Linked Virtual Comm Port " & VCPPort & vbCrLf
                            End If
                        End If
                        kb = kb & vbTab & "Connection = " & sockets(Idx).DevName & vbCrLf
                        On Error GoTo NoPort
                        cPorts.Remove sockets(Idx).Comm.Name
                        If sockets(Idx).errmsg <> "" Then
                            kb = kb & vbTab & sockets(Idx).errmsg & vbCrLf
                        End If
retNoPort:
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next Idx
    
    If Not CommPollTimer Is Nothing Then
        If PollTimerForm.PollTimer.Enabled Then
            kb = kb & "Comm Timer is enabled"
        Else
            kb = kb & "Comm Timer is disabled"
        End If
    End If
    
    If Caption = "" Then
        Caption = "PC Comm Ports & Serial Sockets"
    End If
    MsgBox kb, , Caption
'frmRouter.PollTimer.Enabled = True
Exit Sub

NoPort:
    kb = kb & vbTab & "PC Comm Port not found" & vbCrLf
    Resume retNoPort
End Sub

'Must be done like this because the Timer or form may not exist, if not wouldnt compile
Public Sub EnableCommPollTimer()
    If Not CommPollTimer Is Nothing Then CommPollTimer.Enabled = True
End Sub

Public Sub DisableCommPollTimer()
    If Not CommPollTimer Is Nothing Then CommPollTimer.Enabled = False
End Sub

'Not the same as FreeSockets because of is nothing check and initial values
Public Function FreeComm() As Long
Dim i As Long

'Try & allocate a released socket (if any)
    For i = 1 To UBound(Comms)
        If Not Comms(i) Is Nothing Then
            If Comms(i).State = -1 Then
                Exit For
            End If
        Else
            Exit For
        End If
    Next i

'If no released Comms, i will be the next one available
    FreeComm = i
    
    If FreeComm > MAX_COMMS Then
'no free Comms
        WriteLog "No free Serial handlers, limit is " & MAX_COMMS, LogForm
        FreeComm = -1
    Else
        If FreeComm > UBound(Comms) Then
'We can still allocate more sockets
            ReDim Preserve Comms(1 To FreeComm)
        End If
'Reset any initial values
    End If
End Function

'Used to check if PollTimer requires enabling
Public Function IsSerialHandlerInUse() As Boolean
Dim i As Long

    For i = 1 To UBound(Comms)
        If Not Comms(i) Is Nothing Then
            If Comms(i).State <> -1 Then
                IsSerialHandlerInUse = True
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
End Function

'Moved from modRouter 9/10/15
'Comm name MUST be Comms(Hidx).Name or Sockets(Idx).Comm.Name
'NOT the DeviceName
Public Function GetVCP(CommName As String) As String
Dim Key As String
Dim KeyCount As Long
Dim Keys() As Variant
Dim k As Long
Dim SubKey As String
Dim Words() As String
Dim FriendlyName As String
Dim ReComName As String
Dim ReqCNCName As String

    ReqCNCName = Replace(CommName, "CNCA", "CNCB")
    Key = "SYSTEM\CurrentControlSet\Enum\com0com\port\" _
    & ReqCNCName & "\Device Parameters"
    GetVCP = QueryValue(HKEY_LOCAL_MACHINE, Key, "PortName")
    
#If False Then
    Key = "SYSTEM\CurrentControlSet\Enum\com0com\port\" & CommName
    If FriendlyName <> "" Then
    FriendlyName = QueryValue(HKEY_LOCAL_MACHINE, SubKey, "FriendlyName")
'    KeyCount = ReadKeys(HKEY_LOCAL_MACHINE, Key, Keys)
'    If KeyCount > 0 Then
'        For k = 0 To KeyCount - 1
'            SubKey = Key & "\" & Keys(k)
'            FriendlyName = QueryValue(HKEY_LOCAL_MACHINE, SubKey, "FriendlyName")
            Words = Split(FriendlyName, " ")
            If UBound(Words) = 6 Then
                If Left$(Words(5), 4) = "CNCB" Then
                    GetVCP = Mid$(Words(6), 2, Len(Words(6)) - 2)
                End If
            End If
'        Next k
'    End If
    End If
#End If
End Function

'Moved from modRouter 9/10/15
Public Function GetVCPs() As String()
Dim Key As String
Dim KeyCount As Long
Dim Keys() As Variant
Dim k As Long
Dim SubKey As String
Dim Words() As String
Dim FriendlyName As String
Dim ReComName As String
Dim ReqCNCName As String
Dim List As String
Dim ListCount As Long


    Key = "SYSTEM\CurrentControlSet\Enum\com0com\port"
    KeyCount = ReadKeys(HKEY_LOCAL_MACHINE, Key, Keys)
    If KeyCount > 0 Then
        For k = 0 To KeyCount - 1
            SubKey = Key & "\" & Keys(k)
            FriendlyName = QueryValue(HKEY_LOCAL_MACHINE, SubKey, "FriendlyName")
            Words = Split(FriendlyName, " ")
            If UBound(Words) = 6 Then
                If Left$(Words(5), 4) = "CNCB" Then
                    If ListCount > 0 Then List = List + ","
                    List = List + Mid$(Words(6), 2, Len(Words(6)) - 2)
                    ListCount = ListCount + 1
                End If
            End If
        Next k
    End If
    GetVCPs = Split(List, ",")
End Function

'Sets both Comm Combo box defaults to Port,Baud in Sockets(CurrentSocket).Comm
'If the required defaults are not in the list they are ignored
Public Function SetCommComboBoxDefault(myPort As ComboBox, myBaud As ComboBox)
Dim i As Long

'WriteLog "SetCommboBoxDefault", LogForm

'No Currentsocket
    If CurrentSocket <= 0 Then
        Exit Function
    End If
'Ensure CurrentSocket has been defined (Wont be on initial load of treefilter)
'The handler will be nothing if the socket WAS disabled
'So we need to check if handler info is on Sockets()
    If sockets(CurrentSocket).Handler = 1 Then  'Comms handler
'WriteLog "SetCommboBoxDefault-myBaud", LogForm
        With myBaud
            If .ListCount > 0 Then
                For i = 0 To .ListCount - 1
                    If .List(i) = sockets(CurrentSocket).Comm.BaudRate Then
                        .ListIndex = i
                        Exit For
                    End If
                Next i
            End If
        End With
            
        With myPort
'WriteLog "SetCommboBoxDefault-myPort", LogForm
            
            If .ListCount > 0 Then
                For i = 0 To .ListCount - 1
                    If .List(i) = sockets(CurrentSocket).Comm.Name Then
                        .ListIndex = i
'Dont allow user to change Comm name (if we have one)
'v31 allow                cboCommName.Enabled = False
                        Exit For
                    End If
                Next i
            End If
        End With
    End If
End Function

