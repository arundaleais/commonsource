VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDCB - Device Communication Block utility class
' Part of the Desaware API Class Library
' Copyright (c) 1996 by Desaware Inc.
' All Rights Reserved
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Private Type dcbType    ' Win32API.TXT is incorrect here.
        DCBlength As Long
        BaudRate As Long
        Bits1 As Long
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved2 As Integer
End Type

Private DCB As dcbType
Private BufferSize As Integer

Private Const ERR_INVALIDPROPERTY = 31000
Private Const CLASS_NAME$ = "clsDCB"

Private Const FLAG_fBinary& = &H1
Private Const FLAG_fParity& = &H2
Private Const FLAG_fOutxCtsFlow = &H4
Private Const FLAG_fOutxDsrFlow = &H8
Private Const FLAG_fDtrControl = &H30
Private Const FLAG_fDsrSensitivity = &H40
Private Const FLAG_fTXContinueOnXoff = &H80
Private Const FLAG_fOutX = &H100
Private Const FLAG_fInX = &H200
Private Const FLAG_fErrorChar = &H400
Private Const FLAG_fNull = &H800
Private Const FLAG_fRtsControl = &H3000
Private Const FLAG_fAbortOnError = &H4000

Private Declare Function apiSetCommState Lib "kernel32" Alias "SetCommState" (ByVal hCommDev As Long, lpDCB As dcbType) As Long
Private Declare Function apiGetCommState Lib "kernel32" Alias "GetCommState" (ByVal nCid As Long, lpDCB As dcbType) As Long

Property Get DebugID() As Long
   DebugID = mlngDebugID
End Property

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsDCB; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    
    ' The structure length must always be set
    DCB.DCBlength = Len(DCB)
    ' Set some default values
    BufferSize = 2048
    fParity = False
    fOutxCtsFlow = True
    fOutxDsrFlow = True
    fDtrControl = 1
    fDsrSensitivity = True
    fTXContinueOnXoff = True
    fOutX = True
    fInX = True
    fErrorChar = True
    fNull = True
    fRtsControl = 1
    fAbortOnError = True
    DCB.XonLim = 100
    DCB.XoffLim = BufferSize - 100
    DCB.ByteSize = 8
    DCB.Parity = 0
    DCB.StopBits = 0
    DCB.XonChar = 17
    DCB.XoffChar = 19
    DCB.ErrorChar = Asc("~")
    DCB.EofChar = 26 ' ^Z
    DCB.EvtChar = 255
    ' Set some default value
    DCB.BaudRate = 4800
End Sub

Public Property Get BaudRate() As Long
    BaudRate = DCB.BaudRate
End Property

Public Property Let BaudRate(vNewValue As Long)
Dim i As Integer
'    On Error GoTo InvalidBaudRate
    Select Case vNewValue
        Case 110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 56000, 57600, 115200, 128000, 256000
            DCB.BaudRate = vNewValue
        Case Else
'On return check if baudrate has been set
'This forces the InvalidBaudRate code to execute
'            DCB.BaudRate = 0
'            Err.Raise ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid baud rate"
    End Select
#If False Then
    Exit Property
InvalidBaudRate:
    Select Case err.Number
        Case Is = ERR_INVALIDPROPERTY
    MsgBox "Error " & err.Number & " " & err.Description & vbCrLf _
    & vNewValue & " Baud is not permitted" & vbCrLf _
    & "Baud Rate set to " & DCB.BaudRate, vbOKOnly, "Device Control Block Error"
'    Resume Next 'clears err.
        Case Else
    End Select
    Sockets(clsComm.sIndex).errmsg = err.Description
#End If
End Property

Public Property Get fParity() As Boolean
    If DCB.Bits1 And FLAG_fParity Then
        fParity = True
    End If
End Property

Public Property Let fParity(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fParity)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fParity
End Property

Public Property Get fOutxCtsFlow() As Boolean
    If DCB.Bits1 And FLAG_fOutxCtsFlow Then
        fOutxCtsFlow = True
    End If
End Property

Public Property Let fOutxCtsFlow(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fOutxCtsFlow)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fOutxCtsFlow
End Property

Public Property Get fOutxDsrFlow() As Boolean
    If DCB.Bits1 And FLAG_fOutxDsrFlow Then
        fOutxDsrFlow = True
    End If
End Property

Public Property Let fOutxDsrFlow(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fOutxDsrFlow)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fOutxDsrFlow
End Property

Public Property Get fDtrControl() As Integer
    Dim ival&
    ival = DCB.Bits1 And FLAG_fDtrControl
    fDtrControl = ival \ 16    ' Shift right 4 bits
End Property

' 0 to disable, 1 to enable, 2 for handshake mode
Public Property Let fDtrControl(vNewValue As Integer)
    If vNewValue < 0 Or vNewValue > 2 Then
        err.Raise vbObjectError + ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid fDtrControl setting"
    End If
    DCB.Bits1 = DCB.Bits1 And FLAG_fDtrControl
    DCB.Bits1 = DCB.Bits1 Or (vNewValue * 16)
End Property

Public Property Get fDsrSensitivity() As Boolean
    If DCB.Bits1 And FLAG_fDsrSensitivity Then
        fDsrSensitivity = True
    End If
End Property

Public Property Let fDsrSensitivity(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fDsrSensitivity)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fDsrSensitivity
End Property

Public Property Get fTXContinueOnXoff() As Boolean
    If DCB.Bits1 And FLAG_fTXContinueOnXoff Then
        fTXContinueOnXoff = True
    End If
End Property

Public Property Let fTXContinueOnXoff(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fTXContinueOnXoff)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fTXContinueOnXoff
End Property

Public Property Get fOutX() As Boolean
    If DCB.Bits1 And FLAG_fOutX Then
        fOutX = True
    End If
End Property

Public Property Let fOutX(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fOutX)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fOutX
End Property

Public Property Get fInX() As Boolean
    If DCB.Bits1 And FLAG_fInX Then
        fInX = True
    End If
End Property

Public Property Let fInX(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fInX)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fInX
End Property

Public Property Get fErrorChar() As Boolean
    If DCB.Bits1 And FLAG_fErrorChar Then
        fErrorChar = True
    End If
End Property

Public Property Let fErrorChar(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fErrorChar)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fErrorChar
End Property

Public Property Get fNull() As Boolean
    If DCB.Bits1 And FLAG_fNull Then
        fNull = True
    End If
End Property

Public Property Let fNull(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fNull)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fNull
End Property

Public Property Get fRtsControl() As Integer
    Dim ival&
    ival = DCB.Bits1 And FLAG_fRtsControl
    fRtsControl = ival \ &H1000    ' Shift right 4 bits
End Property

Public Property Let fRtsControl(vNewValue As Integer)
    If vNewValue < 0 Or vNewValue > 3 Then
        err.Raise vbObjectError + ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid fRtsControl setting"
    End If
    DCB.Bits1 = DCB.Bits1 And FLAG_fRtsControl
    DCB.Bits1 = DCB.Bits1 Or (vNewValue * &H1000)
End Property

Public Property Get fAbortOnError() As Boolean
    If DCB.Bits1 And FLAG_fAbortOnError Then
        fAbortOnError = True
    End If
End Property

Public Property Let fAbortOnError(vNewValue As Boolean)
    DCB.Bits1 = DCB.Bits1 And (Not FLAG_fAbortOnError)
    If vNewValue Then DCB.Bits1 = DCB.Bits1 Or FLAG_fAbortOnError
End Property

Public Property Get XonLim() As Integer
    XonLim = DCB.XonLim
End Property

Public Property Let XonLim(vNewValue As Integer)
    DCB.XonLim = vNewValue
End Property

Public Property Get XoffLim() As Integer
    XoffLim = DCB.XoffLim
End Property

Public Property Let XoffLim(vNewValue As Integer)
    DCB.XoffLim = vNewValue
End Property

Public Property Get ByteSize() As Byte
    ByteSize = DCB.ByteSize
End Property

Public Property Let ByteSize(vNewValue As Byte)
    If vNewValue < 4 Or vNewValue > 8 Then
        err.Raise vbObjectError + ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid Byte size setting"
    End If
    DCB.ByteSize = vNewValue
End Property

Public Property Get Parity() As Byte
    Parity = DCB.Parity
End Property

' 0 - 4 = No, odd, even, mark, space
Public Property Let Parity(vNewValue As Byte)
    If vNewValue < 0 Or vNewValue > 4 Then
        err.Raise vbObjectError + ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid Parity setting"
    End If
    DCB.Parity = vNewValue
End Property

Public Property Get StopBits() As Byte
    StopBits = DCB.StopBits
End Property

' 0 = 1, 1 = 1.5, 2 = 2
Public Property Let StopBits(vNewValue As Byte)
    If vNewValue < 0 Or vNewValue > 4 Then
        err.Raise vbObjectError + ERR_INVALIDPROPERTY, CLASS_NAME, "Invalid Stop bits"
    End If
    DCB.StopBits = vNewValue
End Property

Public Property Get XonChar() As String
    XonChar = Chr$(DCB.XonChar)
End Property

Public Property Let XonChar(vNewValue As String)
    DCB.XonChar = Asc(vNewValue)
End Property

Public Property Get XoffChar() As String
    XoffChar = Chr$(DCB.XoffChar)
End Property

Public Property Let XoffChar(vNewValue As String)
    DCB.XoffChar = Asc(vNewValue)
End Property

Public Property Get ErrorChar() As String
    ErrorChar = Chr$(DCB.ErrorChar)
End Property

Public Property Let ErrorChar(vNewValue As String)
    DCB.ErrorChar = Asc(vNewValue)
End Property

Public Property Get EofChar() As String
    EofChar = Chr$(DCB.EofChar)
End Property

Public Property Let EofChar(vNewValue As String)
    DCB.EofChar = Asc(vNewValue)
End Property

Public Property Get EvtChar() As String
    EvtChar = Chr$(DCB.EvtChar)
End Property

Public Property Let EvtChar(vNewValue As String)
    DCB.EvtChar = Asc(vNewValue)
End Property

' Load the current dcb with the state of the comm device
Public Function GetCommState(Comm As clsComm) As Boolean
    Dim Res&
    ' Make sure comm device is initialized
    If Comm.hCommDev = 0 Then Exit Function
    Res = apiGetCommState(Comm.hCommDev, DCB)
    GetCommState = Res <> 0
End Function

' Set the comm device from the current dcb
Public Function SetCommState(Comm As clsComm) As Boolean
    Dim Res&
    ' Make sure comm device is initialized
    If Comm.hCommDev = 0 Then Exit Function
    Res = apiSetCommState(Comm.hCommDev, DCB)
    SetCommState = Res <> 0
End Function

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
