VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Major.Minor.Revision = Version.Revision
'The program checks the latest version and lowest revision on the website
'The Latest version the program will download latest version with the lowest revision
'When the last revision of of a version is ready to make live
'   1.  Make and Save a copy of the source with the next version no (eg save
'   2.  Compile and upload it (This is the Source of Latest version users can download)
'   3.  Change the version to the next revision
'   Revision 1,2,3 etc are now the "Testing" versions which a user can be requested to download
'until we consolidate the changes into a New Version
'The lastest LIVE version is the HIGHEST version with the LOWEST revision.
'If there is a later revision of the Highest Version, it is a test revision

'Check registry to see if we are updating if new version
'These need inserting in the Installation File
'[Registry]
'Root: HKLM; Subkey: "Software\Arundale"; Flags: uninsdeletekeyifempty
'Root: HKLM; Subkey: "Software\Arundale\ProgramUpdate"; Flags: uninsdeletekey
'Root: HKLM; Subkey: "Software\Arundale\ProgramUpdate\Settings"; ValueType: string; ValueName: "LastVersion"; ValueData: "{#MyAppVersion}"

Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Private fEOF As Boolean
'Private fDebugFlg As Boolean   'if true display debugging info on exit
Private fThisVersion As String  'Version currently running
Private fNewVersion As String   'Latest Live version on the website
Private fLastVersion As String  'Most recent version we have checked or Skipped (LastVersion in registry)
Private fWebHighestVersion As String    'Highest found on website
Private fDownloaded As Boolean  'True if file has been downloaded
Private fDownloadedUrl As String
Private fSetupExecuted As Boolean

Private Myform As New frmNewVersion

Public Property Get DownloadedURL() As String
    DownloadedURL = fDownloadedUrl
End Property

Public Property Get Downloaded() As Boolean
    Downloaded = fDownloaded
End Property

Public Property Get SetupExecuted() As Boolean
    SetupExecuted = fSetupExecuted
End Property

Public Property Get NewVersion() As String
    NewVersion = fNewVersion
End Property

Public Property Get ThisVersion() As String
    ThisVersion = fThisVersion
End Property

Public Property Get LastVersion() As String
    LastVersion = fLastVersion
End Property

Public Property Get WebHighestVersion() As String
    WebHighestVersion = fWebHighestVersion
End Property

'Public Property Let DebugFlg(ByVal NewVal As Boolean)
'    fDebugFlg = NewVal
'End Property

Public Property Let EOF(ByVal NewVal As Boolean)
    fEOF = NewVal
End Property

Public Property Let SetupExecuted(ByVal NewVal As Boolean)
    fSetupExecuted = NewVal
End Property

'Returns False if new version not downloaded, else returns true
Public Sub CheckNewVersion(Server As String, ReloadVersion As String, Optional DebugFlg As Boolean)
Dim Website As String
'Dim ThisSetupUrl As String
Dim DirHtml As String
Dim P1 As Long  'Position of <a anchor tag
Dim P2 As Long  'Position of > closing anchor tag
Dim Link As String
Dim Prefix As String
Dim V1 As Long  'Start position of version info
Dim arry() As String    'Array to split version
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewMajor As Long
Dim NewMinor As Long
Dim LastMajor As Long   'Last = Lowest that must NOT be updated (the one kept in registry)
Dim LastMinor As Long
Dim NewRevision As Long
Dim NewVersion As String    'Latest version user may download
Dim NewUrl As String
Dim LiveMajor As Long   'Current live download
Dim LiveMinor As Long
Dim LiveRevision As Long
Dim LiveVersion As String
Dim kb As String
Dim i As Long
    
    arry = Split(Server, "/")
    Website = "http://" & arry(0)
    Prefix = App.EXEName & "_setup_"
'    ThisSetupUrl = Prefix & App.Major & "_" & App.Minor & "_" & App.Revision & ".exe"
'ensure key exists
        CreateNewKey HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings"
'Get the Last version that was OK from registry
    fLastVersion = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "LastVersion")
    If fLastVersion <> "" Then
        arry = Split(fLastVersion, ".")
        LastMajor = arry(0)
        LastMinor = arry(1)
    End If
'Initially set up the New Version as the existing version
    NewMajor = App.Major
    NewMinor = App.Minor
    NewRevision = App.Revision
    fThisVersion = NewMajor & "." & NewMinor & "." & NewRevision
    fNewVersion = fThisVersion

'Get directory listing into DirHTML
    DirHtml = Myform.HttpGetDocument(Server)
If DirHtml = "" Then
'MsgBox "Blank DirHtml"
End If
    
    P2 = 0
    Do
        P1 = InStr(P2 + 1, DirHtml, "<a href=""", vbTextCompare)
        If P1 = 0 Then Exit Do
        P2 = InStr(P1 + 1, DirHtml, """>", vbTextCompare)
'Check individual line
        Link = Mid$(DirHtml, P1 + 9, P2 - P1 - 9)
'Get position of Version info
        V1 = InStr(1, Link, Prefix, vbTextCompare)
        If V1 > 0 Then
'Split version into 1.0.0.exe or 1.0.0.0.exe
Debug.Print Link
            arry = Split(Mid$(Link, V1 + Len(Prefix)), ".")
            
'Skip non-numeric versions v3.4.143  allows 143A etc for testing
            For i = 0 To UBound(arry) - 1
                If IsNumeric(arry(i)) = False Then
Debug.Print "version " & Mid$(Link, V1 + Len(Prefix)) & " skipped"
                GoTo skip_version
                End If
            Next i
            
            Major = arry(0)
            Minor = arry(1)
            Revision = arry(UBound(arry) - 1)  'Last but one
            fWebHighestVersion = Major & "." & Minor & "." & Revision
 'The Latest (alpha) version to upload is the Lowest Revision of the Highest Major + Minor
            
            If Major > LiveMajor Or (Major = LiveMajor And Minor > LiveMinor) Then
                LiveMajor = Major
                LiveMinor = Minor
                LiveRevision = Revision
                LiveVersion = Major & "." & Minor & "." & Revision
            End If
                        
            If Major > NewMajor Or (Major = NewMajor And Minor > NewMinor) Then
'It must also be a higher version than the last (registry-beta) version
                If Major > LastMajor Or (Major = LastMajor And Minor > LastMinor) Then
                    NewMajor = Major
                    NewMinor = Minor
'Always reset revision and URL to the first New version found on the website
                    NewRevision = Revision
                    NewVersion = NewMajor & "." & NewMinor & "." & NewRevision
                    NewUrl = Server & Link
'Keep 1.0.1 for registry
                    fNewVersion = NewMajor & "." & NewMinor & "." & NewRevision
                End If
           End If
'If we have a later revision ensure we pick the earliest revision for this version
'A User must upgrade to the lowest revision, higher (beta) revisions are only for testing or bugs.
            If NewUrl <> "" Then
                If Revision < NewRevision Then
                    NewRevision = Revision
                    NewUrl = Server & Link
                End If
            End If
        End If
skip_version:
    Loop
    
    If ReloadVersion <> "" Then
        NewUrl = Server & Prefix & ReloadVersion & ".exe"
    
    End If

'MsgBox "ThisSetupUrl=" & ThisSetupUrl & vbCrLf & "newurl=" & NewUrl
'There is a valid New upload
    If NewUrl <> "" Then
'        frmDpyBox.DpyBox "New Version " & NewMajor & "." & NewMinor & vbCrLf & "Available for download at " & Website & vbCrLf & "..." & vbCrLf & "..." & vbCrLf, 30, App.EXEName
'MsgBox "New Version Available" & vbCrLf & NewUrl
        fDownloaded = Myform.DownloadBox(NewUrl, 30)   'Returns True if downloaded
        If fDownloaded = True Then fDownloadedUrl = NewUrl
    Else
'MsgBox "skipping"
    End If

    kb = "Live version is " & LiveVersion & vbCrLf
    kb = kb & "Latest test version is " & MyNewVersion.WebHighestVersion & vbCrLf
    kb = kb & "This executable version is " & MyNewVersion.ThisVersion & " will be replaced when setup is run" & vbCrLf
    kb = kb & "Downloads will be skipped until after " & MyNewVersion.LastVersion & vbCrLf
    kb = kb & "(This is the last version we have asked the user if they wish to download" & vbCrLf
    kb = kb & "and they have downloaded or Skipped the version.  It's the LastVersion in registry)" & vbCrLf
    
    If MyNewVersion.Downloaded = True Then
        kb = kb & MyNewVersion.DownloadedURL & " downloaded OK" & vbCrLf
        kb = kb & "Version " & MyNewVersion.ThisVersion & " will be replaced by version " & MyNewVersion.NewVersion & " when setup is run" & vbCrLf
        kb = kb & App.EXEName & " is terminating to allow install of version " & MyNewVersion.NewVersion & vbCrLf
    Else
        If NewUrl = "" Then
            kb = kb & "No new download available" & vbCrLf
        Else
            kb = kb & "New download skipped" & vbCrLf
        End If
        kb = kb & "Downloads will be skipped until after " & MyNewVersion.LastVersion & vbCrLf
        kb = kb & "The Latest Downloadable Version is " & MyNewVersion.NewVersion & vbCrLf
        kb = kb & "Current " & App.EXEName & " will continue"
'set new version to this version so program will not terminate
        fNewVersion = fThisVersion
    
    End If
    
    Call WriteStartUpLog(kb)
    
End Sub

Private Sub Class_Initialize()
'Incase class is called twice(should not happen with this class but demonstrates the principle)
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsNewVersion; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    Set Myform = New frmNewVersion
End Sub

Private Sub Class_Terminate()
    Set Myform = Nothing
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub


