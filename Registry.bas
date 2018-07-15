Attribute VB_Name = "Registry"
'see http://support.microsoft.com/kb/145679
'also http://vbnet.mvps.org/index.html?code/reg/reguninstall.htm
Option Explicit

   Public Const REG_SZ As Long = 1
   Public Const REG_EXPAND_SZ As Long = 2
   Public Const REG_DWORD As Long = 4
Private Const ERROR_MORE_DATA = 234

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   
      Const ERROR_SUCCESS = 0&
      Const SYNCHRONIZE = &H100000
      Const STANDARD_RIGHTS_READ = &H20000
      Const STANDARD_RIGHTS_WRITE = &H20000
      Const STANDARD_RIGHTS_EXECUTE = &H20000
      Const STANDARD_RIGHTS_REQUIRED = &HF0000
      Const STANDARD_RIGHTS_ALL = &H1F0000
   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F
      Const KEY_CREATE_SUB_KEY = &H4
      Const KEY_ENUMERATE_SUB_KEYS = &H8
      Const KEY_NOTIFY = &H10
      Const KEY_CREATE_LINK = &H20
      Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY) And _
                        (Not SYNCHRONIZE))


   Public Const REG_OPTION_NON_VOLATILE = 0

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000


Private Type FILETIME 'ft
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

   Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hkey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long

   Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
   Declare Function RegEnumValue Lib "advapi32.dll" _
          Alias "RegEnumValueA" _
          (ByVal hkey As Long, _
          ByVal dwIndex As Long, _
          ByVal lpValueName As String, _
          lpcbValueName As Long, _
          ByVal lpReserved As Long, _
          lpType As Long, _
          lpData As Any, _
          lpcbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" _
   Alias "RegEnumKeyExA" _
   (ByVal hkey As Long, _
   ByVal dwIndex As Long, _
   ByVal lpName As String, _
   lpcbName As Long, _
   ByVal lpReserved As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   lpftLastWriteTime As FILETIME) As Long

    Declare Function RegQueryInfoKey Lib "advapi32.dll" _
   Alias "RegQueryInfoKeyA" _
  (ByVal hkey As Long, _
   ByVal lpClass As String, _
   lpcbClass As Long, _
   ByVal lpReserved As Long, _
   lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, _
   lpcbMaxClassLen As Long, _
   lpcValues As Long, _
   lpcbMaxValueNameLen As Long, _
   lpcbMaxValueLen As Long, _
   lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As FILETIME) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias _
        "FormatMessageA" (ByVal dwFlags As Long, _
        lpSource As Long, ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
        ByVal nSize As Long, Arguments As Any) As Long
        
Public RegistryDisplayControl As Control   'Text box to use if RegistryDisplay is called
                                            'Public to allow for main form to unload on exit
                                            
Public Function SetValueEx(ByVal hkey As Long, sValueName As String, _
   lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       Select Case lType
           Case REG_SZ, REG_EXPAND_SZ
               sValue = vValue & Chr$(0)
               SetValueEx = RegSetValueExString(hkey, sValueName, 0&, _
                                              lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValueEx = RegSetValueExLong(hkey, sValueName, 0&, _
   lType, lValue, 4)
           End Select
   End Function

   Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As _
   String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String
        Dim ExpandedValue As String
       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
           ' For strings
           Case REG_SZ, REG_EXPAND_SZ:
               sValue = String(cch, 0)

   lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
   sValue, cch)
               If lrc = ERROR_NONE Then
                    If lType = REG_EXPAND_SZ Then
'must create a buffer null filled large enough for the output
                        ExpandedValue = String(cch + 400, 0)
                        cch = ExpandEnvironmentStrings(sValue, ExpandedValue, Len(ExpandedValue))
'trim the output to the actual length output
'removing the trailing null
                        vValue = Left$(ExpandedValue, cch - 1)
                    Else
                        vValue = Left$(sValue, cch - 1)
                    End If
               Else
                   vValue = Empty
               End If
           ' For DWORDS
           Case REG_DWORD:
   lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
   lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
   End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Variant
       Dim lRetVal As Long         'result of the API functions
       Dim hkey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, _
   KEY_QUERY_VALUE, hkey)
       lRetVal = QueryValueEx(hkey, sValueName, vValue)
       QueryValue = vValue
       RegCloseKey (hkey)
   End Function

Public Sub SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, _
   vValueSetting As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hkey As Long         'handle of open key

       'open the specified key
       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, _
                                 KEY_SET_VALUE, hkey)
       lRetVal = SetValueEx(hkey, sValueName, lValueType, vValueSetting)
       RegCloseKey (hkey)
   End Sub

Public Sub CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
       Dim hNewKey As Long         'handle to the new key
       Dim lRetVal As Long         'result of the RegCreateKeyEx function

       lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
       RegCloseKey (hNewKey)
   End Sub
 
Public Sub DeleteKey(lPredefinedKey As Long, sNewKeyName As String)
Dim lngResult As Long
    lngResult = RegDeleteKey(lPredefinedKey, sNewKeyName)
    If lngResult = 2 Then   'Key not found (happens on New key)
        Exit Sub
    End If
    If lngResult <> ERROR_SUCCESS Then 'no more data
        MsgBox "System Error " & lngResult & vbCrLf _
        & "sNewKeyName=" & sNewKeyName & vbCrLf _
        & OLEError(lngResult), , "DeleteKey"
    End If
End Sub

Public Sub testRegistry()
Dim retvalue As Variant
'CreateNewKey "TestKey\SubKey1\SubKey2", HKEY_LOCAL_MACHINE
CreateNewKey HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings"
SetKeyValue HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "LastSetupFileDownloadDate", "Hello", REG_SZ
retvalue = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "LastSetupFileDownloadDate")

End Sub
'Root: HKLM; Subkey: "Software\Arundale\Ais Decoder\Settings"; ValueType: string; ValueName: "Path"; ValueData: "{app}"
'LastSetupFileDownloadDate = GetSetting(appname:="AisDecoder", section:="Startup", _
                       Key:="LastSetupFileDownloadDate")

Public Function ReadKeys(HiveKey As Long, Key As String, Keys()) As Long
    ReadKeys = EnumKeys(HiveKey, Key, Keys)
End Function

Public Function ReadNameValues(HiveKey As Long, Key As String, Names(), Values()) As Long
    ReadNameValues = EnumKeyNameAndData(HiveKey, Key, Names, Values)
End Function

Public Function EnumKeyNameAndData(lPredefinedKey As Long, sKeyName As String, retName(), retData()) As Long

         Dim lngKeyHandle As Long
         Dim lngResult As Long
         Dim lngCurIdx As Long
         Dim strValue As String
         Dim lngValueLen As Long
         Dim strData As String     'pointer to data buffer
         Dim lngDataLen As Long     'length of data buffer
         Dim strResult As String

                lngResult = RegOpenKeyEx(lPredefinedKey, _
                 sKeyName, _
                  0&, _
                  KEY_READ, _
                  lngKeyHandle)
         
         If lngResult <> ERROR_SUCCESS Then
'this will fail on initial load because there will not be a key
'             MsgBox "Cannot open key"
             Exit Function
         End If

         lngCurIdx = 0
         Do
            lngValueLen = 4000
'create buffers for the data (RegEnumValue expects a pointer to the value)
'The variable must be given enough size - fill with nulls
            strValue = String(lngValueLen, 0)
            lngDataLen = 4000
            strData = String(lngDataLen, 0)
            lngResult = RegEnumValue(lngKeyHandle, _
                                     lngCurIdx, _
                                     ByVal strValue, _
                                     lngValueLen, _
                                     0&, _
                                     REG_SZ, _
                                     ByVal strData, _
                                     lngDataLen)
            
        If lngResult = ERROR_SUCCESS Then
            ReDim Preserve retName(lngCurIdx)
            ReDim Preserve retData(lngCurIdx)
            retName(lngCurIdx) = Left(strValue, lngValueLen)
            retData(lngCurIdx) = StripNulls(strData) 'Left(strData, lngDataLen - 1) 'remove trailing null
            lngCurIdx = lngCurIdx + 1
         Else
            If lngResult = 259 Or lngResult = 234 Then
'no more data or More Data availaable (Win10 enumerating ports)
            Else
                MsgBox "System Error " & lngResult & vbCrLf _
                & "sKeyName=" & sKeyName & vbCrLf _
                & OLEError(lngResult), , "EnumKeyNameAndData"
            End If
         End If

         Loop While lngResult = ERROR_SUCCESS
         Call RegCloseKey(lngKeyHandle)
        EnumKeyNameAndData = lngCurIdx
End Function

Public Function EnumKeys(lPredefinedKey As Long, sKey As String, retKeys()) As Long

   Dim hkey As Long
   Dim dwIndex As Long
   Dim dwSubKeys As Long
   Dim dwMaxSubKeyLen As Long
   Dim ft As FILETIME
   Dim Success As Long
   Dim sName As String
   Dim cbName As Long
    Dim lngResult As Long
    
  'obtain a handle to the uninstall key
     lngResult = RegOpenKeyEx(lPredefinedKey, _
                 sKey, _
                  0&, _
                  KEY_READ, _
                  hkey)
         
         If lngResult <> ERROR_SUCCESS Then
'this will fail on initial load because there will not be a key
'             MsgBox "Cannot open key"
             Exit Function
         End If
   
  'if valid
   If hkey <> 0 Then
        
     'query registry for the number of
     'entries under that key
      If RegQueryInfoKey(hkey, _
                         0&, _
                         0&, _
                         0, _
                         dwSubKeys, _
                         dwMaxSubKeyLen&, _
                         0&, _
                         0&, _
                         0&, _
                         0&, _
                         0&, _
                         ft) = ERROR_SUCCESS Then

 
        'enumerate each item
         For dwIndex = 0 To dwSubKeys - 1
         
            sName = Space$(dwMaxSubKeyLen + 1)
            cbName = Len(sName)
            
            Success = RegEnumKeyEx(hkey, _
                                   dwIndex, _
                                   sName, _
                                   cbName, _
                                   0, _
                                   0, _
                                   0, _
                                   ft)
            
            If Success = ERROR_SUCCESS Or _
               Success = ERROR_MORE_DATA Then
            ReDim Preserve retKeys(dwIndex)
            retKeys(dwIndex) = StripNulls(sName) 'Left(sName, Len(sName) - 1)
            End If
         
         Next  'For dwIndex

      End If  'If RegQueryInfoKey
      
      Call RegCloseKey(hkey)
  
   End If  'If hKey <> 0
EnumKeys = dwIndex
End Function

Public Function OLEError(lError As Long) As String
Dim sReturn As String
Dim lReturn As Long
sReturn = Space$(256)
lReturn = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lError, 0&, sReturn, 256&, 0&)

If lReturn > 0 Then
    OLEError = Left$(sReturn, lReturn)
Else
    OLEError = "Error not found."
End If
End Function

Function StripNulls(ByVal S As String) As String
  Dim i As Integer
  i = InStr(S, Chr$(0))
  If i > 0 Then
    StripNulls = Left$(S, i - 1)
  Else
    StripNulls = S
  End If
End Function

'TTYs or Outputs to a file the Registry Keys/Values
'For output to a form, then form must have txtRegistryDisplay defined as a text box
'AND the Form must have been loaded before OutputAllKeys is called
Public Function OutputAllKeys(RootKey As String, Optional RegFile As String) As String
Dim Key As String
Dim SubKeyCount As Long
Dim SubKeys() As Variant
Dim k As Long
Static Level As Long
Static Down As Boolean
Dim NameValueCount As Long
Dim Names() As Variant  'must be for passing as array argument
Dim Values() As Variant
Dim kb As String
Static RegFileCh As Integer
Dim RegistryDisplayForm As Form
Dim ctrl As Control

    If Level = 0 Then
        Down = True
        If RegFile <> "" Then   'FILE Output
            RegFileCh = FreeFile
            kb = RegFile
            Open kb For Output As #RegFileCh
        Else                    'Display Output
 'If first time, find location of RegistryDisplayControl
            If RegistryDisplayControl Is Nothing Then   'Module level declaration
                For Each RegistryDisplayForm In Forms
                  For Each ctrl In RegistryDisplayForm
                    If TypeOf ctrl Is Textbox And ctrl.Name = "txtRegistryDisplay" Then
                        Set RegistryDisplayControl = ctrl
                        Exit For
                    End If
                    Next ctrl
                    If Not RegistryDisplayControl Is Nothing Then Exit For
                Next RegistryDisplayForm
            End If
            
'Must always have defined a txtRegistryDisplay control in some form (if displaying)
            If RegistryDisplayControl Is Nothing Then
                If RegistryDisplayForm Is Nothing Then
                    MsgBox "RegistryDisplay Form and Control not Found", , "Registry.OutputAllKeys"
                Else
                    MsgBox "txtRegistryDisplay Control not Found", , "Registry.OutputAllKeys"
                End If
                Exit Function   'Cant display as we have no form+control
            End If
        End If

'Insert the header
        If RegFileCh = 0 Then
'            Call frmRegistry.DisplayOutput(RootKey & vbCrLf)
            Call RegistryDisplayControl.Parent.DisplayOutput(RootKey & vbCrLf)
        Else
            Print #RegFileCh, "Windows Registry Editor Version 5.00"
        End If
    End If
    Level = Level + 1
    Key = RootKey
    If Down Then
        kb = "[HKEY_CURRENT_USER\" & Key & "]"
        If RegFileCh = 0 Then
'            Call frmRegistry.DisplayOutput(vbCrLf & kb & vbCrLf)
            Call RegistryDisplayControl.Parent.DisplayOutput(vbCrLf & kb & vbCrLf)
        Else
            Print #RegFileCh, kb
        End If
        NameValueCount = ReadNameValues(HKEY_CURRENT_USER, Key, Names, Values)
        If NameValueCount > 0 Then
            For k = 0 To UBound(Names)  'should be same as NameValueCount
                kb = """" & Names(k) & """" & "=" & """" & Values(k) & """"
                If RegFileCh = 0 Then
'                    Call frmRegistry.DisplayOutput(kb & vbCrLf)
                    Call RegistryDisplayControl.Parent.DisplayOutput(kb & vbCrLf)
                Else
                    Print #RegFileCh, kb
                End If
            Next k
        End If
    End If
    SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
    If SubKeyCount > 0 Then
        For k = 0 To SubKeyCount - 1
            Down = True
            Call OutputAllKeys(Key & "\" & SubKeys(k), TempPath & RegFile)
        Next k
    End If
    Level = Level - 1
    Down = False
    If Level = 0 Then
        Close #RegFileCh
        If RegFileCh = 0 Then
 'may already be visible
            If RegistryDisplayControl.Parent.Visible = False Then
                RegistryDisplayControl.Parent.Show  'make vbModal if main form does not force close
            End If
        End If
        RegFileCh = 0
    End If
End Function

'Recursively goes through the registry
Public Function EnumAllKeys(lPredefinedKey As Long, RootKey As String) As String
Dim Key As String
Dim SubKeyCount As Long
Dim SubKeys() As Variant
Dim k As Long
Static Level As Long
Static Down As Boolean
Dim NameValueCount As Long
Dim Names() As Variant  'must be for passing as array argument
Dim Values() As Variant
Dim kb As String
Dim PortName As String
Static FriendlyName As String

    If Level = 0 Then
        Down = True
    End If
    Level = Level + 1
    lPredefinedKey = HKEY_LOCAL_MACHINE
    Key = RootKey
    If Down Then
        NameValueCount = ReadNameValues(lPredefinedKey, Key, Names, Values)
        
        If NameValueCount > 0 Then
            For k = 0 To UBound(Names)  'should be same as NameValueCount
                kb = Names(k) & ":" & Values(k)
If Names(k) = "FriendlyName" Then
    FriendlyName = Values(k)
    PortName = ""
'Stop
End If
If Names(k) = "PortName" Then
    PortName = Values(k)
kb = PortName & ":" & FriendlyName
'Stop
End If
            Next k
        End If

    End If
    SubKeyCount = ReadKeys(lPredefinedKey, Key, SubKeys)
    If SubKeyCount > 0 Then
        For k = 0 To SubKeyCount - 1
            Down = True

If SubKeys(k) = "Device Parameters" Then
'Stop
End If

            Call EnumAllKeys(lPredefinedKey, Key & "\" & SubKeys(k))
        Next k
    End If
    Level = Level - 1
    Down = False
End Function

'Recursively goes through the registry
Public Function FriendlyName(CommName As String, Optional RootKey As String) As String
Dim Key As String
Dim SubKeyCount As Long
Dim SubKeys() As Variant
Dim k As Long
Static Level As Long
Static Down As Boolean
Dim NameValueCount As Long
Dim Names() As Variant  'must be for passing as array argument
Dim Values() As Variant
Dim kb As String
Dim PortName As String
Static LastFriendlyName As String
Static Found As Boolean
    
    If Level = 0 Then
        Down = True
        If RootKey = "" Then RootKey = "SYSTEM\CurrentControlSet\Enum"
        LastFriendlyName = ""
        Found = False
    End If
    Level = Level + 1
    Key = RootKey
    If Down Then
        NameValueCount = ReadNameValues(HKEY_LOCAL_MACHINE, Key, Names, Values)
        
        If NameValueCount > 0 Then
            For k = 0 To UBound(Names)  'should be same as NameValueCount
                kb = Names(k) & ":" & Values(k)
                If Names(k) = "FriendlyName" Then
                    LastFriendlyName = Values(k)
'Stop
                End If
                If Names(k) = "PortName" Then
                    If Values(k) = CommName Then
                        FriendlyName = LastFriendlyName
                        Found = True
                        GoTo Back
                    End If
                End If
            Next k
        End If

    End If
    SubKeyCount = ReadKeys(HKEY_LOCAL_MACHINE, Key, SubKeys)
    If SubKeyCount > 0 Then
        For k = 0 To SubKeyCount - 1
            Down = True

If SubKeys(k) = "Device Parameters" Then
'Stop
End If

            Call FriendlyName(CommName, Key & "\" & SubKeys(k))
            If Found Then
                FriendlyName = LastFriendlyName
                GoTo Back
            End If
        Next k
    End If
    
Back:
    Level = Level - 1
    Down = False
End Function


'Deletes the key and all subkeys
Public Function DeleteKeys(RootKey As String)
Dim Key As String
Dim SubKeyCount As Long
Dim SubKeys() As Variant
Dim k As Long

    Key = RootKey
    SubKeyCount = ReadKeys(HKEY_CURRENT_USER, Key, SubKeys)
    If SubKeyCount > 0 Then
        For k = 0 To SubKeyCount - 1
            Call DeleteKeys(Key & "\" & SubKeys(k))
            Call DeleteKey(HKEY_CURRENT_USER, Key & "\" & SubKeys(k))
'            Debug.Print key & "\" & SubKeys(k)
            Next k
    End If
    Call DeleteKey(HKEY_CURRENT_USER, Key)
'    Debug.Print key
End Function

'used to output key values to StartupLogFile
Public Function PrintKey(Key As String, ExtKey As String, SubKey As String) As String
Dim KeyConstant As Long
    Select Case Key
    Case Is = "HKCU"
        KeyConstant = HKEY_CURRENT_USER
    Case Is = "HKLM"
        KeyConstant = HKEY_LOCAL_MACHINE
    End Select
    PrintKey = Key & "\Software\Arundale\" & App.EXEName & "\Settings" & ExtKey _
    & SubKey & " = " _
    & QueryValue(KeyConstant, "Software\Arundale\" & App.EXEName & "\Settings" & ExtKey, SubKey)

End Function

Public Sub PrintRegistry()
Dim i As Long

    Call WriteStartUpLog(PrintKey("HKCU", "\", "InitialisationFile"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "InstalledAppName"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "InstalledAppDateTime"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "InstalledAppVersion"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "InstallDateTime"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "AllUsersPath"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "NewUserInitialisationFile"))
    Call WriteStartUpLog(PrintKey("HKLM", "\", "FallBackInitialisationFile"))
    For i = 1 To 15
        Call WriteStartUpLog(PrintKey("HKLM", "\ReservedFiles", CStr(i)))
    Next i
    Call WriteStartUpLog(PrintKey("HKCU", "\", "InstalledIniFileDateTime"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "InitialisationFile"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "CurrentUserPath"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "DisableNmeaFillBitsError"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "TestDac"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "DacMap"))
    Call WriteStartUpLog(PrintKey("HKCU", "\", "Licence"))
End Sub


