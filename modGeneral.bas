Attribute VB_Name = "modGeneral"
'These are General routines that can in priciple be used by any program
'In the main they will include formatting and usedful VB routines

'Requires Project > References "Windows Script Host Object Model" to use FSO (FileSystemObject)
Option Explicit

Public gcolDebug As New Collection

Public TempPath As String   'StartupLog

Public StartupLogFile As String
Public StartupLogFileCh As Long

Public LogFileName As String
Public LogFileCh As Integer
Public LogForm As Form      'Form we are also displaying log messages on status bar (NmeaRouter, Form1)
                            'Requires ClearStatusBarTimer control on it, and Sub ClearStatusBarTimer_Timer()
                            
Public Function FormExists(FormName As String) As Boolean
Dim i As Long

    For i = 0 To Forms.Count - 1
        If Forms(i).Name = FormName Then
            FormExists = True
        End If
    Next i
End Function

' Global function to give each object a unique ID.
' Storage for the debug ID.
'Private mlngDebugID As Long needs defining in each class module

Public Function DebugSerial() As Long
Static lngSerial As Long
Dim mycolDebug As Variant
Dim ScannedObjects As Long

'check not overflow
    If lngSerial = (2 ^ 31 - 1) Then lngSerial = 0
    lngSerial = lngSerial + 1
'if reset to 0 must check does not already exist
    On Error GoTo NotExisting
    Do Until ScannedObjects > 1000
        mycolDebug = gcolDebug.Item(lngSerial)
        Set mycolDebug = Nothing
        lngSerial = lngSerial + 1
        ScannedObjects = ScannedObjects + 1
    Loop
NotExisting:
    DebugSerial = lngSerial
End Function

Public Function LogDebugCls()
Dim mycolDebug As Variant
Dim kb As String
Dim Count As Long

    kb = kb & "Entry to LogDebugCls.gcolDebug.Count = " & gcolDebug.Count & vbCrLf
    For Each mycolDebug In gcolDebug
        kb = kb & vbTab & mycolDebug + vbCrLf
        Count = Count + 1
        If Count = 20 Then
            kb = kb & "there are also another " & gcolDebug.Count - Count & " classes open" & vbCrLf
            Exit For
        End If
    Next
    If kb = "" Then
        kb = "All classes have been terminated" & vbCrLf
    End If
    kb = kb & "Exit from LogDebugCls.gcolDebug.Count = " & gcolDebug.Count & vbCrLf
    WriteStartUpLog ("DebugCls:" & vbCrLf & kb)
    Call CloseStartupLogFile    'My be called after log file has been closed on program termination
End Function

Public Function DisplayDebugSerial()
Dim mycolDebug As Variant
Dim kb As String

    For Each mycolDebug In gcolDebug
        kb = mycolDebug + vbCrLf
    Next
    If kb <> "" Then Call frmDpyBox.DpyBox(kb, 5, "DebugSerial")
End Function

'Returns true if exists and is loaded
'Optionally unloads it
'Required if the form is not defined in this project (can happen with a common sub)
Public Function IsFormLoaded(frmName As String) As Form
Dim frm As Form
   For Each frm In Forms
      If (frm.Name = frmName) Then
        Set IsFormLoaded = frm
        Exit Function
      End If
   Next
   Set IsFormLoaded = Nothing
End Function


Public Function LongFileName(ByVal short_name As String) As _
    String
Dim pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
'            long_name = Dir$(short_name, vbNormal + _
'                vbHidden + vbSystem + vbDirectory)
'a blank name (above returns a ".")
            long_name = ""
        Else
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

Public Sub WriteStartUpLog(kb As String)
If StartupLogFileCh = 0 Then
    StartupLogFileCh = FreeFile
    If StartupLogFile = "" Then
        StartupLogFile = Environ("TEMP") & "\" & App.EXEName & ".log"
    End If
    Open StartupLogFile For Append As #StartupLogFileCh
    Call WriteStartUpLog("")
    Call WriteStartUpLog("Log Appended at " & Now())
End If
Print #StartupLogFileCh, kb
End Sub

Public Sub CloseStartupLogFile()
If StartupLogFileCh <> 0 Then
    Call WriteStartUpLog("Log Closed at " & Now())
    Close #StartupLogFileCh
    StartupLogFileCh = 0
End If
End Sub

'Check to see if any file is left open
Public Function IsAnyFileOpen(Optional CloseOpenFiles As Boolean) As Boolean
Dim ch As Integer
On Error GoTo FileOpen
For ch = 1 To 255
    Open StartupLogFile For Append As #ch
    Close #ch
Next ch
Exit Function
FileOpen:
    If CloseOpenFiles = True Then
        Close #ch
        Resume Next
    End If
    IsAnyFileOpen = True
End Function

Public Sub UnloadAllForms(Optional Caller As Form)
Dim f As Form
Dim CallerName As String
Dim kb As String
Dim i As Long

'Call DisplayForms
If Not Caller Is Nothing Then CallerName = Caller.Name

'Call DisplayForms
'remove in reverse order to load
For i = Forms.Count To 1 Step -1
    Set f = Forms(i - 1)
    If f.Name <> CallerName Then     'unload all but calling form
        Unload f
        Set f = Nothing 'clears all references Programmers Guide P 428
    End If
Next i

'For Each f In Forms
'    If f.Name <> CallerName Then     'unload all but calling form
'kb = f.Name & ":" & f.Caption
'        f.Visible = False
'        If f.Name <> "frmSplash" Then
'f.Visible = True
'        Unload f
'        End If
'    End If
'Next f
'Call DisplayForms
If CallerName <> "" Then
    Unload Caller
    Set Caller = Nothing
End If
'Call DisplayForms
End Sub

'If StatusBarForm is defined then Log is also written to this forms status bar
Public Function WriteLog(kb As String, Optional StatusBarForm As Form)
Static LastTime As String

'If no previously name defined
    If TempPath = "" Then TempPath = LongFileName(Environ("TEMP") & "\")

'If InStr(1, kb, "Profile Loaded") > 1 Then Stop
    If LogFileCh = -1 Then  'First time its opened
        LogFileCh = FreeFile
Try:
'Get a unique file name AisDecoderLog_xxx.log
'If this is the only instance of this program, all existing
'files in the above format will be deleted (if in TempPath)
        LogFileName = UniqueFileName(TempPath & App.EXEName & "Log.log")
        On Error GoTo BadFile
        If FileExists(LogFileName) Then
            err.Raise 31000
        End If
        Open LogFileName For Output As #LogFileCh
        On Error GoTo 0
        WriteLog "Open Event Log [" & LogFileName & "]", LogForm
    End If
    If LogFileCh = 0 Then   'Re-open in append
        LogFileCh = FreeFile
        If LogFileName = "" Then LogFileName = TempPath & App.EXEName & "Log.log"
        Open LogFileName For Append As #LogFileCh
    End If
    
    If Now() <> LastTime Then
        Print #LogFileCh, Now() & vbTab & kb
    Else
        Print #LogFileCh, Space$(Len(LastTime)) & vbTab & kb
    End If
    LastTime = Now()

'Reset clear message timer (30 secs)
    If Not StatusBarForm Is Nothing Then
        With StatusBarForm
            .ClearStatusBarTimer.Enabled = False
            .StatusBar.Panels(1).Text = kb
            .ClearStatusBarTimer.Enabled = True
        End With
    End If
    
'When running ensure log file is written out
'    If frmRouter.cmdStart.Enabled = False Then
        Close LogFileCh
        LogFileCh = 0
'    End If

Exit Function

BadFile:
    Resume Try
End Function

'Given FileName.ext returns FileName-seqno.ext
'Does not add a Seqno if there is only one file
'Which will be unique to this application
'If no other copies of the application are running when it is called
'It will delete all Filename-*.ext first
Public Function UniqueFileName(FullFileName As String) As String
Dim fso As New FileSystemObject
Dim fol As Folder
Dim i As Long
Dim Path As String
Dim Name As String
Dim ExistingFileName As String
Dim FullExistingFileName As String
Dim MatchedFiles() As String
Dim myFile As File
Dim MatchLen As Long
Dim Seqno As Long
Dim MaxSeqno   As Long
Dim j As Long
Dim FileDeleted As Boolean

    MaxSeqno = -1   'No file found
    Path = PathFromFullName(FullFileName)
    If Len(Path) = 0 Then Exit Function
'MsgBox "UniqueFileName - argument" & vbCrLf & FullFileName
    Name = NameFromFullPath(FullFileName, , False)   'dont remove rollover date
'we may have to remove the version but leave the rollover, if we want to keep previous versions
    If Name = "" Then Exit Function
    MatchLen = InStr(Name, ".") - 1
    If MatchLen <= 0 Then
           MatchLen = Len(Name)
    End If
    Set fol = fso.GetFolder(Path)
'    MatchedFiles = Filter(fol.Files, "NmeaRouter")
    For Each myFile In fol.Files
        If Left$(myFile.Name, MatchLen) = Left$(Name, MatchLen) Then
'So that I can test funny file names
            ExistingFileName = myFile.Name
'This does not detect VBE is running
            If App.PrevInstance = False Then
                FullExistingFileName = myFile
'If the delete fails, clock the seq no
                On Error GoTo BadDelete
                Kill FullExistingFileName
            Else
'ExistingFileName = "a"
'Stop
DeleteSkipped:  On Error GoTo 0
'At least one file is found and is left
                If MaxSeqno = -1 Then MaxSeqno = 0
                j = InStr(ExistingFileName, ".") 'get the first dot
                If j = 0 Then j = Len(ExistingFileName) 'no dot so all the string
                i = InStrRev(Left$(ExistingFileName, j), ";")
                If i > 0 And j > i + 1 Then 'filename i j ext
                    If IsNumeric(Mid$(ExistingFileName, i + 1, j - i - 1)) Then
                        Seqno = Mid$(ExistingFileName, i + 1, j - i - 1)
                        If MaxSeqno < Seqno Then MaxSeqno = Seqno
                    End If
                End If
            End If  'Get next highest Seqno
        End If
    Next
'if there is no other file, just return the passed file name
    If MaxSeqno = -1 Then
        UniqueFileName = FullFileName
    Else
        UniqueFileName = ExtendFullName(FullFileName, ";" & MaxSeqno + 1)
    End If
'MsgBox "UniqueFileName - return" & vbCrLf & UniqueFileName
    Exit Function
    
BadDelete:
    Resume DeleteSkipped
End Function

' Return True if a file exists
Public Function FileExists(FileName As String) As Boolean
    FileExists = False
'MsgBox FileName & ":" & GetAttr(FileName)
    On Error GoTo errorhandler
    If NameFromFullPath(FileName) <> "" Then  'directory
'does file exists
        If (GetAttr(FileName) And vbNormal) = vbNormal Then FileExists = True
    End If
'MsgBox Filename & vbCrLf & FileExists
errorhandler:
    ' if an error occurs, this function returns False
End Function


Public Function NameFromFullPath(FullPath As String, Optional Delimiter As String, Optional RemoveRollover As Boolean) As String
'Input: Name/Full Path of a file
'Returns: Name of file

    Dim sPath As String
    Dim sList() As String
    Dim sAns As String
    Dim iArrayLen As Integer
    Dim i As Integer
    Dim j As Integer
    Dim kb As String
    
'MsgBox "NameFromFullPath -arguments" & vbCrLf & FullPath & "," & Delimiter & "," & RemoveRollover
    If Delimiter = "" Then Delimiter = "\"
    If Len(FullPath) = 0 Then Exit Function
    sList = Split(FullPath, Delimiter)
    iArrayLen = UBound(sList)
'if arraylen = 0 the sans="" else sans=last element in array (ie Filename)
    sAns = IIf(iArrayLen = 0, "", sList(iArrayLen))
'only filename
    If sAns = "" And iArrayLen = 0 Then sAns = FullPath 'no \ in full path
    If RemoveRollover And sAns <> "" Then
'remove the version (if any)
        j = InStr(sAns, ".") 'get the first dot
        j = j - 1   'j is last chr before dot (if any)
        If j <= 0 Then j = Len(sAns) 'no dot so all the string
        i = InStrRev(Left$(sAns, j), ";")
        If i > 0 And j - i > 0 Then  'at least -n (1 version number)
            If IsNumeric(Mid$(sAns, i + 1, j - i)) Then
                sAns = Replace(sAns, Mid$(sAns, i, j - i + 1), "")
            End If
            
        End If
        
'Remove the rollover date (if any)
        j = InStr(sAns, ".") 'get the first dot
        j = j - 1   'j is last chr before dot (if any)
        If j <= 0 Then j = Len(sAns) 'no dot so all the string
        i = InStrRev(Left$(sAns, j), "_")
'v144        If j = i + 9 Then 'must be _yyyymmdd.
        Do
'Check if date (must be _yyyymmdd)
            If i > 0 And j = i + 8 Then 'v144 must be _yyyymmdd. fix for all numeric user defined file name
                If IsNumeric(Mid$(sAns, i + 1, 8)) Then
                    sAns = Replace(sAns, Mid$(sAns, i, 9), "")  '_yyyymmdd
                    Exit Do
                End If
            End If
            j = i - 1   'Length of filename left to scan for date
            If j < 8 Then Exit Do   'must be at least 9 characters _yyyymmdd
            i = InStrRev(Left$(sAns, j), "_")
        Loop
    End If
    
    NameFromFullPath = sAns
'MsgBox "NameFromFullPath - return " & vbCrLf & NameFromFullPath
End Function

Function PathFromFullName(FullPath As String, Optional Delimiter As String) As String
Dim i As Long
If Delimiter = "" Then Delimiter = "\"
i = InStrRev(FullPath, Delimiter)
If i > 0 Then
    PathFromFullName = Left$(FullPath, i - 1)
Else
    PathFromFullName = ""
End If
End Function

Public Function ExtendFullName(FullPath As String, AddIn As String) As String
Dim i As Long
Dim PathAndName As String
Dim Ext As String
i = InStrRev(FullPath, ".")
If i > 0 Then
    PathAndName = Left$(FullPath, i - 1)
    Ext = Mid$(FullPath, i)
Else
    PathAndName = FullPath
End If
ExtendFullName = PathAndName & AddIn & Ext
End Function

Sub DisplayForms()
Dim f As Form
Dim kb As String

    For Each f In Forms
        kb = kb & f.Name & ":" & f.Caption
        kb = kb & vbCrLf
    Next f
    MsgBox kb
End Sub

Sub LogForms(Optional Title As String)  'v3.4.143
Dim f As Form
Dim kb As String
Dim Visible As String

    kb = Title & vbCrLf
    For Each f In Forms
        If f.Visible = False Then
            Visible = " (Hidden)"
        Else
            Visible = ""
        End If
        kb = kb & f.Name & ":" & f.Caption & Visible & vbCrLf
    Next f
    kb = kb & "Total of " & Forms.Count & " forms loaded" & vbCrLf
    Call WriteStartUpLog(kb)
End Sub

'v149
Public Function IsFileInUse(FileName As String) As Boolean
Dim ch As Long

    ch = FreeFile
    On Error Resume Next
    Open FileName For Input Lock Read Write As #ch
    If err.Number = 70 Or err.Number = 55 Then IsFileInUse = True
    Close #ch
End Function

#If False Then      'V142
Sub DisplayQueryUnload(FormName As String, Cancel As Integer, UnloadMode As Integer)
Dim Reason As String

Select Case UnloadMode
Case Is = vbFormControlMenu
    Reason = "User clicked close(X)"
    QueryQuit = True
Case Is = vbFormCode
    Reason = "Unload invoked from code"
Case Is = vbAppWindows
    Reason = "Operating environment session ending"
Case Is = vbAppTaskManager
    Reason = "Task Manager"
Case Is = vbFormMDIForm
    Reason = "MDI parent closing child"
Case Is = vbFormOwner
    Reason = "Owner closing"
Case Else
    Reason = "Unknown"
End Select
Call WriteStartUpLog(FormName & ".Form_QueryUnload (UnloadMode=" & Reason & ", Cancel=" & Cancel & ")")

End Sub
#End If

#If True Then
'Used to terminate main program form to avoid Not Quitting program cleanly
'Called by frmRouter_QueryUnload
'and frm_QueryUnload
'Returns vbOK if user clicks X, otherwise 0
Public Function QueryQuit(ByRef UnloadMode As Integer) As Integer
Dim Reason As String
Dim YesNo As String
    QueryQuit = vbOK      'default is Quit
Select Case UnloadMode
Case Is = vbFormControlMenu
    Reason = "User clicked close(X)"
    QueryQuit = MsgBox("Do you wish to Quit " & App.EXEName, vbOKCancel, App.EXEName)
Case Is = vbFormCode
    Reason = "Unload invoked from code"
Case Is = vbAppWindows
    Reason = "Operationg environment session ending"
Case Is = vbAppTaskManager
    Reason = "Task Manager"
Case Is = vbFormMDIForm
    Reason = "MDI parent closing child"
Case Is = vbFormOwner
    Reason = "Owner closing"
Case Else
    Reason = "Unknown"
End Select
    If QueryQuit = vbOK Then
        YesNo = "OK"
    Else
        YesNo = "Cancel"
    End If
    WriteStartUpLog ("Query Quit = " & YesNo & ", (UnloadMode=" & Reason & ")")
'MsgBox Reason
End Function
#End If

'This is not suitable for sorting large collections
'http://www.freevbcode.com/ShowCode.asp?ID=4522
Public Sub SortCollection(col As Collection, psSortPropertyName As String, pbAscending As Boolean, Optional psKeyPropertyName As String)
'This routine is designed to re-arrange OBJECTS (not values!)
'inside a collection by any numeric/string/date property
'which name is supplied as an argument.
Dim obj As Object
Dim i As Integer
Dim j As Integer
Dim iMinMaxIndex As Integer
Dim vMinMax As Variant
Dim vValue As Variant
Dim bSortCondition As Boolean
Dim bUseKey As Boolean
Dim sKey As String
    
    bUseKey = (psKeyPropertyName <> "")
    
    For i = 1 To col.Count - 1
        Set obj = col(i)
        vMinMax = CallByName(obj, psSortPropertyName, VbGet)
        iMinMaxIndex = i
        
        For j = i + 1 To col.Count
            Set obj = col(j)
            vValue = CallByName(obj, psSortPropertyName, VbGet)
            
            If (pbAscending) Then
                bSortCondition = (vValue < vMinMax)
            Else
                bSortCondition = (vValue > vMinMax)
            End If
            
            If (bSortCondition) Then
                vMinMax = vValue
                iMinMaxIndex = j
            End If
            
            Set obj = Nothing
        Next j
        
        If (iMinMaxIndex <> i) Then
            Set obj = col(iMinMaxIndex)
            
            col.Remove iMinMaxIndex
            If (bUseKey) Then
                sKey = CStr(CallByName(obj, psKeyPropertyName, VbGet))
                col.Add obj, sKey, i
            Else
                col.Add obj, , i
            End If
            
            Set obj = Nothing
        End If
        
        Set obj = Nothing
    Next i
        
End Sub
