Attribute VB_Name = "modComPorts"
Option Explicit

'API calls
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Sub CopyMem Lib "kernel32.dll" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

'API Structures
Private Type PORT_INFO_2
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type
Dim Ports() As PORT_INFO_2

Private Type API_PORT_INFO_2
    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long
End Type


Public Function TrimStr(strName As String) As String
'====================================================
Dim X As Integer

    X = InStr(strName, vbNullChar)
    If X > 0 Then TrimStr = Left(strName, X - 1) Else TrimStr = strName

End Function

Private Function LPstrToString(ByVal lngPointer As Long) As String
'==================================================================
Dim lngLength As Long

    'Get number of characters in string
    lngLength = lstrlenW(lngPointer) * 2
    'Initialize string so we have something to copy the string into
    LPstrToString = String(lngLength, 0)
    'Copy the string
    CopyMem ByVal StrPtr(LPstrToString), ByVal lngPointer, lngLength
    'Convert to Unicode
    LPstrToString = TrimStr(StrConv(LPstrToString, vbUnicode))

End Function

'Uses the API (but didnt spot MarineGadgets)
Public Function GetAvailablePorts(ServerName As String) As String()
'================================================================
Dim ret As Long
Dim PortsStruct(0 To 100) As API_PORT_INFO_2
Dim pcbNeeded As Long
Dim pcReturned As Long
Dim TempBuff As Long
Dim i As Integer
Dim ComPorts() As String
Dim j As Integer

    'Get the amount of bytes needed to contain the data returned by the API call
    ret = EnumPorts(ServerName, 2, TempBuff, 0, pcbNeeded, pcReturned)
    'Allocate the Buffer
    TempBuff = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    ret = EnumPorts(ServerName, 2, TempBuff, pcbNeeded, pcbNeeded, pcReturned)
    If ret Then
        ReDim Ports(pcReturned - 1)
        'Convert the returned String Pointer Values to VB String Type
        CopyMem PortsStruct(0), ByVal TempBuff, pcbNeeded
        For i = 0 To pcReturned - 1
            Ports(i).pDescription = LPstrToString(PortsStruct(i).pDescription)
            Ports(i).pPortName = LPstrToString(PortsStruct(i).pPortName)
            Ports(i).pMonitorName = LPstrToString(PortsStruct(i).pMonitorName)
            Ports(i).fPortType = PortsStruct(i).fPortType
If Left$(LPstrToString(PortsStruct(i).pPortName), 3) = "COM" Then
    ReDim Preserve ComPorts(j)
'remove ; at end
    ComPorts(j) = Left$(LPstrToString(PortsStruct(i).pPortName), Len(LPstrToString(PortsStruct(i).pPortName)) - 1)
    j = j + 1
End If
        Next i
    End If
'    GetAvailablePorts = pcReturned
    'Free the Heap Space allocated for the Buffer
    If TempBuff Then HeapFree GetProcessHeap(), 0, TempBuff
    GetAvailablePorts = ComPorts
End Function

'Uses the Registry
Public Function GetSerialPorts() As String()
Dim Key As String
Dim NameValueCount As Long
Dim Names() As Variant  'must be for passing as array argument
Dim Values() As Variant
Dim i As Long
Dim arry() As String
Dim ComPorts() As String
Dim PortCount As Long
Dim j As Long

        Key = "HARDWARE\DEVICEMAP\SERIALCOMM"
'WriteLog "Getting Comm Ports key =" & Key
        
        NameValueCount = ReadNameValues(HKEY_LOCAL_MACHINE, Key, Names, Values)
'WriteLog "NameValueCount=" & NameValueCount
'May be no Serial Ports
'Test NameValueCount = 0
        If NameValueCount > 0 Then
'WriteLog "Names Size Returned=" & UBound(Names)
            For i = 0 To UBound(Names)
'WriteLog "Name(" & i & ") is " & Names(i) & "=" & Values(i)
                arry = Split(Names(i), "\")
'If arry = "" Ubound(arry)=-1
'arry(0)="",(1)="Device",(2)="Serial1" or com0com11
'WriteLog "Token array size (Zero based)=" & UBound(arry)
For j = 0 To UBound(arry)
'    WriteLog "Token(" & j & ")=" & arry(j)
Next j

'v48 additional check
                If UBound(arry) > 1 Then
                    If UCase$(arry(1)) = UCase$("Device") Then
'                        If Not Left$(arry(2), 7) = "vComDrv" Then
                        ReDim Preserve ComPorts(PortCount)
                        ComPorts(PortCount) = Values(i)
'WriteLog "Device " & arry(2) & "=" & Values(i) & " added to ComPorts(" & PortCount & ")", LogForm
                        PortCount = PortCount + 1
'                       End If
                    Else
'WriteLog "Device Identifier not found"
                    End If
                Else
'WriteLog "Device Identifier not found - insufficient tokens"
                End If
            Next i
        End If
    GetSerialPorts = ComPorts
End Function

'Check is a com port exists
'Used to ensure a Com port in a profile exists
Public Function IsSerialPort(SerialPortName As String) As Boolean
Dim Ports() As String
Dim i As Long
Dim SerialPortCount As Long
Dim j As Long   'Lowest port index (-1 = no ports)

    j = -1
    Ports = GetSerialPorts
'May be no ports
    On Error GoTo GotaPort
    SerialPortCount = 1 + UBound(Ports) 'zero based
    j = LBound(Ports)
GotaPort: On Error GoTo 0
    If j = -1 Then Exit Function 'no ports
    For i = LBound(Ports) To UBound(Ports)
        If UCase(SerialPortName) = UCase(Ports(i)) Then
            IsSerialPort = True
            Exit Function
        End If
    Next i
End Function

Public Function GetSerialPortCount(Ports() As String) As Long
    On Error GoTo noPorts
    GetSerialPortCount = UBound(Ports) + 1
noPorts:
End Function

