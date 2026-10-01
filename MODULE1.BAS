Attribute VB_Name = "MODULE11"
Option Explicit

'**********************************************************
' Declare Windows NT API Calls
'
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long '2.35
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'/* *************************************/
'/* declare WIN32 API constants. *
'/* *************************************/
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_APPLMODAL = &H0&
Public Const MB_COMPOSITE = &H2
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&
Public Const MB_DEFMASK = &HF00&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_MISCMASK = &HC000&
Public Const MB_MODEMASK = &H3000&
Public Const MB_NOFOCUS = &H8000&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_PRECOMPOSED = &H1
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_TYPEMASK = &HF&
Public Const MB_USEGLYPHCHARS = &H4
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&
Public Const IDABORT = 3
Public Const IDCANCEL = 2
Public Const IDIGNORE = 5
Public Const IDNO = 7
Public Const IDOK = 1
Public Const IDRETRY = 4
Public Const IDYES = 6
'/* **************************************/
'/* declare WIN32 API structures. *
'/* **************************************/
Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public t_NotifyStructure As NOTIFYICONDATA

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


'/* **************************************/
'/* declare WIN 32 API functions. */
'/* **************************************/
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
lpData As Any) As Long


'/* **********************************/
'/* declare personal constants. *
'/* **********************************/
Public Const MBC_OKEXCLAMATION = MB_OK Or MB_ICONEXCLAMATION Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const MBC_OKSTOP = MB_OK Or MB_ICONSTOP Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const MBC_OKINFORMATION = MB_OK Or MB_ICONINFORMATION Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const MBC_YESNOCANCELEXCLAMATION = MB_YESNOCANCEL Or MB_ICONEXCLAMATION Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const MBC_YESNOICONINFORMATION = MB_YESNO Or MB_ICONINFORMATION Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const MBC_YESNOICONEXCLAMATION = MB_YESNO Or MB_ICONEXCLAMATION Or MB_DEFBUTTON1 'Or MB_TASKMODAL
Public Const GENERIC_READ = &H80000000
Public Const OPEN_EXISTING = 3
Public Const FILE_SHARE_READ = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000

'/* **************************************/
'/* declare WIN 32 API functions. */
'/* **************************************/
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, _
ByVal lpCaption As String, ByVal wType As Long) As Long

Global cfg_DB_FILE As String
Global cfg_DB_PATH As String
Global cfg_DB_WEMNAME As String
Global Taskname As String
Global MyWorkspace As Workspace
Global db As Database
Global dbArchive As Database

Global Rdb As Database
Global olddate As String
Global nottime As Integer
Global DB_FILE As String
Global cfg_UseSQL As Long           '2.00   0 means use Access, 1 means use SQL
Global cfg_SQLMAIN_Mode As Long     '2.00
Global oSQLTK As SQLToolkit             '2.00
Global cfg_SQLMAIN_DSN As String            '2.00
Global cfg_SQLMAIN_ARCHIVE_DSN As String    '2.32
Global cfg_SQLMAIN_OverrideSvrName As String            '2.00
Global cfg_SQLMAIN_WEMAdminUserID As String            '2.00
Global cfg_SQLMAIN_WEMAdminPWD As String            '2.00
Global cfg_SQLMAIN_OverrideDBName As String            '2.00
Global sWEMAdminPWD_Encrypted As String         '2.00
Global sWEMAdminPWD2_Encrypted As String        ' 2.17
Global Language As String       ' 2.17
Global gbl_jobname As String

Global cfile(400) As String
Global Sourcefile(500) As String
Global Destfile(500) As String
Global Deletefile(500) As String
Global tofile(500) As String
Global Maxcopys As Integer
Global othercopys As Integer
Global deletecopys As Integer                 '2.01.0
Global btn_copy As String
Global gbl_dontchecktime As Boolean
Global sql_BaseFolder As String                 '2.15
Global PCNAME(21) As String                     '2.15
Global backup_ctr As Integer                    ''2.18

'**********************************************************
' Declare Windows NT API Calls
'
'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nsize As Long) As Long
'Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long
'Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long '2.35
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global BackupHour(24) As Integer
Global CFG_BACKUPTYPE As String         '2.01.1
Global CFG_BACKUPTYPE_REDUNDANCY As String  '2.01.1
Global Install_Path As String
Global MDBDirectory As String
Global MDBFile As String
Global DB_Name As String
Public g_oAppLog As New AppLog          '2.01.0
Global cfg_destdrive As String
Global CFG_BACKUP_AT_STARTUP As Boolean     '2.23
Global cfg_resetonceperday As Boolean       '2.29
'V2.20
Global cfg_System As sysType
Public Enum sysType
    stUnknown = 0
    stFeed = 1
    stAsphalt = 2
End Enum

Function GetWindowsDir() As String
Dim temp$
Dim X

    temp$ = String$(145, 0)              ' Size Buffer
    X = GetWindowsDirectory(temp$, 145)  ' Make API Call
    temp$ = Left$(temp$, X)              ' Trim Buffer

    If Right$(temp$, 1) <> "\" Then      ' Add \ if necessary
        GetWindowsDir$ = temp$ + "\"
    Else
        GetWindowsDir$ = temp$
    End If
End Function
'************************************************************************************************
' LoadCIniGlobals - Pick up information from *.INI file.
'
'       Calls GetPrivateProfileString
'
'       Parameters
'             Application group name  - WEMFLOSTALK
'             Profile Name            - PASSDATAPATH
'             Default                 - ""
'             Return value string [WemReturn = Space$(129)]
'             Number of characters allowed to return
'             Name of ini file - PHOENIX.INI
'
Sub LoadCIniGlobals()

Dim FileIniName$
Dim fs%
Dim tplen%
Dim tp$
Dim WemApp$
Dim WemProfile$
Dim temp$
Dim windir$
Dim nSize As Long
Dim WEMIniName$         '2.00
Dim WEMDefault
Dim X

    On Error Resume Next

    windir$ = UCase$(GetWindowsDir$())
    '2.00
    WEMIniName$ = windir$ & "WEM.ini"
    FileIniName$ = windir$ + "AUTOBACK.INI"
   '-----------------------------
    'SQL Server INI settings
    'get initial default values from the wem.ini.  Included hard coded standard defaults that will be functional.
    cfg_SQLMAIN_Mode = Val(GetINISetting("SQL_MAIN", "SQLMode", "0", WEMIniName$))
    cfg_SQLMAIN_DSN = GetINISetting("SQL_MAIN", "DSN", "WEMSQL_DSN", WEMIniName$)
    cfg_SQLMAIN_ARCHIVE_DSN = GetINISetting("SQL_MAIN_ARCHIVE", "DSNArchive", "WEMSQL_ARCHIVE_DSN", WEMIniName$)
    cfg_SQLMAIN_OverrideSvrName = GetINISetting("SQL_MAIN", "SvrNameSkipDSN", "", WEMIniName$)
    cfg_SQLMAIN_OverrideDBName = GetINISetting("SQL_MAIN", "DBName", "", WEMIniName$)
    sWEMAdminPWD_Encrypted = GetINISetting("SQL_MAIN", "WEMAdminPWD", "8AE699801DC00F09280A037130D00ABD", WEMIniName$)
    sWEMAdminPWD2_Encrypted = GetINISetting("SQL_MAIN", "WEMAdminPWD2", "", WEMIniName$)        ' 2.17
    
    
    
    'get app specific overrides (if present).  use defaults from the wem.ini
    cfg_SQLMAIN_Mode = Val(GetINISetting("SQL_MAIN", "SQLMode", CStr(cfg_SQLMAIN_Mode), FileIniName$))
    cfg_SQLMAIN_DSN = GetINISetting("SQL_MAIN", "DSN", cfg_SQLMAIN_DSN, FileIniName$)
    cfg_SQLMAIN_ARCHIVE_DSN = GetINISetting("SQL_MAIN_ARCHIVE", "DSNArchive", "WEMSQL_ARCHIVE_DSN", FileIniName$)
    cfg_SQLMAIN_OverrideSvrName = GetINISetting("SQL_MAIN", "SvrNameSkipDSN", cfg_SQLMAIN_OverrideSvrName, FileIniName$)
    cfg_SQLMAIN_OverrideDBName = GetINISetting("SQL_MAIN", "DBName", cfg_SQLMAIN_OverrideDBName, FileIniName$)
    sWEMAdminPWD_Encrypted = GetINISetting("SQL_MAIN", "WEMAdminPWD", sWEMAdminPWD_Encrypted, FileIniName$)
    
    Language = GetINISetting("Language", "Language", "", WEMIniName$)       ' 2.17
    If cfg_SQLMAIN_Mode > 0 Then
        cfg_UseSQL = cfg_SQLMAIN_Mode
        'set user id
        cfg_SQLMAIN_WEMAdminUserID = "WEMAdmin"
        'decrypt the password
        Set oSQLTK = New SQLToolkit
        If UCase(Language) = "CHINESE" Then '2.17
            cfg_SQLMAIN_WEMAdminPWD = oSQLTK.DecryptText2(sWEMAdminPWD2_Encrypted)
        Else
            cfg_SQLMAIN_WEMAdminPWD = oSQLTK.DecryptText(sWEMAdminPWD_Encrypted)
        End If
    End If
'start 2.08.0
    WemApp$ = "WEMBACKUP"
    
    WEMDefault = "wem.mdb"
    WemProfile$ = "DB_FILE"            ' Get name of mdb directory
    DB_FILE = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    WemProfile$ = "MDBDIRECTORY"            ' Get name of mdb directory
    WEMDefault = "F:\WEM\DATA"
    MDBDirectory = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    DB_Name = MDBDirectory & "\" & DB_FILE
   
    WemProfile$ = "INSTALLDIRECTORY"            ' Get name of directory where we are installed
    WEMDefault = "C:\WEM\AUTOBACK"
    Install_Path = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    Dim thehour As Integer
    thehour = Left$(Time$, 2)       '2.18
    If thehour = 0 Then thehour = 24    '2.29
    WemProfile$ = "BACKUPHOUR"
    WEMDefault = 1
    Dim tmp_backuphour As Variant   '2.18
    tmp_backuphour = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable) '2.18
    tmp_backuphour = Split(tmp_backuphour, ",", , vbTextCompare)
    If UBound(tmp_backuphour) = 0 Then BackupHour(0) = tmp_backuphour(0)
    For X = 0 To UBound(tmp_backuphour)
        BackupHour(X) = tmp_backuphour(X)
    Next
    
    backup_ctr = 0
    For X = 0 To UBound(tmp_backuphour) - 1
        If thehour < BackupHour(X) Then            '2.18
            backup_ctr = X
            Exit For
        End If
    Next
    
    WEMDefault = "TRUE"           '2.01.1
    WemProfile$ = "BACKUP_AT_STARTUP"
    CFG_BACKUP_AT_STARTUP = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    WEMDefault = "DCOLLECT"           '2.01.1
    WemProfile$ = "BACKUPTYPE"
    CFG_BACKUPTYPE = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    WEMDefault = "f:\"           '2.15
    WemProfile$ = "sql_BaseFolder"
    sql_BaseFolder = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    WEMDefault = ""           '2.24
    WemProfile$ = "WEMJOBNUMBER"
    gbl_jobname = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    
    WEMDefault = "False"           '2.01.1
    WemProfile$ = "resetonceperday"
    cfg_resetonceperday = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    
    WemApp$ = "PCNAMES"
    WEMDefault = ""           '2.15
    For X = 1 To 20     '2.31
        WemProfile$ = "PC" & X
        PCNAME(X) = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    Next
    
    FileIniName$ = windir$ + "WEM_REDUNDANCY.INI"
    WemApp$ = "General"
    WEMDefault = "REDUNDANCY"           '2.01.1
    WemProfile$ = "BACKUPTYPE"
    CFG_BACKUPTYPE_REDUNDANCY = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable
    
    If UCase(CFG_BACKUPTYPE) <> UCase(CFG_BACKUPTYPE_REDUNDANCY) Then
        MsgBox ("AUTOBACK.INI BACKUPTYPE DOES NOT MATCH WEM_REDUNDANCY.INI BACKUPTYPE." & vbCrLf & "PLEASE CORRECT")
    End If
    
    
    FileIniName$ = windir$ + "DCOLLECTION.INI"          '2.07
    WemApp$ = "General"
    WemProfile$ = "DESTDRIVE"            'Get name of customer where we are installed
    WEMDefault = "F:"
    cfg_destdrive = ReturnIniValue(FileIniName$, WemApp$, WemProfile$, WEMDefault)    'Copy string into variable    'Copy string into variable
    'end 2.08.0
    
    nSize = 128
    tp$ = Space$(128)
    fs% = GetComputerName(tp$, nSize)
    If fs% > 0 Then
        Taskname = Left$(tp$, nSize) 'just get the computer name
    Else
        Taskname = "UnKnown"
    End If
'    Taskname = "w16830client7"
'    Form1.RegInUse.RegDirectory = MDBDirectory    '2.00
'    Form1.RegInUse.Taskname = Taskname + "-AutoBack-S"
'    Form1.RegInUse.UnRegisterTask
    On Error GoTo 0
End Sub
'For INI setting that exists, but is set equal to blank (eg. MYKEY=) GetPrivateProfileString will return a blank
'even if default is specified.  This proc tests if any INI value is returned.  if nothing or blank, returns the default.
Private Function GetINISetting(sSection As String, sKey As String, sDefault As String, sINIFilePath As String) As String
On Error Resume Next
Dim sBuffer As String * 512
Dim lResult As Long
Dim lNullPos As Long
Dim sValue As String

lResult = GetPrivateProfileString(sSection, sKey, sDefault, sBuffer, Len(sBuffer), sINIFilePath)
If lResult > 0 Then
    lNullPos = InStr(sBuffer, Chr$(0))
    If lNullPos > 0 Then
         sValue = Trim(Left$(sBuffer, lNullPos - 1))  'strip off terminating null.  GetPrivateProfileString probably does a Trim()
                                                      'so this part may be redundant.
    End If
End If

'if blank is returned
If sValue = "" Then
    'make sure is converted to default
    sValue = sDefault
End If
GetINISetting = sValue
Err.Clear
End Function

Public Sub Main()
    g_oAppLog.AddEntry "Program Started"
    Load Form1
End Sub

Public Sub AddItem(o_FormObject As Form, s_ToolTipText As String)

    '/* turn On Local error handling. */
    On Local Error GoTo ErrorHandler

    '/*  set the individual values of the NOTIFYICONDATA data type. */
    t_NotifyStructure.cbSize = Len(t_NotifyStructure)
    t_NotifyStructure.hwnd = o_FormObject.hwnd
    t_NotifyStructure.uID = vbNull
    t_NotifyStructure.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t_NotifyStructure.uCallbackMessage = WM_MOUSEMOVE
    t_NotifyStructure.hIcon = o_FormObject.Icon
    t_NotifyStructure.szTip = s_ToolTipText

    '/* call the Shell_NotifyIcon function to add the icon to the taskbar status area. */
    Shell_NotifyIcon NIM_ADD, t_NotifyStructure
    
    Exit Sub
    
ErrorHandler:

    '/* set mouse pointer to default. */
    Screen.MousePointer = 0

    '/* raise error message. */
    Err.Raise Err.Number, "autoback.SystemTray", Err.Description

    Exit Sub

End Sub

Public Sub RemoveItem()

    '/* turn On Local error handling. */
    On Local Error GoTo ErrorHandler

    '/* click this button to delete the added icon from the taskbar status area by calling the Shell_NotifyIcon function. */
    Shell_NotifyIcon NIM_DELETE, t_NotifyStructure
    
    Exit Sub
    
ErrorHandler:

    '/* set mouse pointer to default. */
    Screen.MousePointer = 0

    '/* raise error message. */
    Err.Raise Err.Number, "Collect.SystemTray", Err.Description

    Exit Sub

End Sub

Public Function copy_update_folder(job)         '2.26
Dim fs As New FileSystemObject
On Error Resume Next
fs.CreateFolder ("f:\update")
fs.CreateFolder ("f:\update\gbl_apps")
fs.CreateFolder ("f:\update\gbl_ww")
fs.CreateFolder ("f:\update\gbl_system32")
fs.CopyFolder "\\" & job & "fserver\wemdata\update\gbl_apps", "f:\update\gbl_apps"
fs.CopyFolder "\\" & job & "fserver\wemdata\update\gbl_system32", "f:\update\gbl_system32"
If fs.FolderExists("\\" & job & "fserver\wemdata\update\gbl_ww\" & job) = True Then
    fs.MoveFolder "F:\update\gbl_ww", "f:\update\gbl_ww-old"
    fs.CopyFolder "\\" & job & "fserver\wemdata\update\gbl_ww", "f:\update\"
    If fs.FolderExists("f:\update\gbl_ww\" & job) = True Then
        fs.DeleteFolder "f:\update\gbl_ww-old"
    Else
        fs.MoveFolder "f:\update\gbl_ww-old", "f:\update\gbl_ww"
    End If
End If
End Function

Public Function GetDCollectFiles(job As String, comp As String)
Dim locrec As Recordset
Dim table As String
Dim Index As Integer
Dim filefound As String
Dim temp As String
Dim sp As Integer
Dim srcpath As String
Dim destpath As String
Dim newdir As String
Dim ATTR As Integer
'Dim cfg_destdrive As String
Dim fielddatadate As String
Dim sfile As File
Dim dfile As File
                                    
    On Error Resume Next
    fielddatadate = ""
'    cfg_destdrive = "F:"
'    MsgBox comp
    CreateDESTDir cfg_destdrive, comp
    Form1.MousePointer = vbHourglass
    On Error Resume Next
    Dim WS As Workspace
    Dim dcollect_db As dao.Database    '2.00
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    Set WS = Workspaces(0)    '2.00
    Set dcollect_db = WS.OpenDatabase("c:\wem\dcollection\dcollection.mdb")
    For Index = 0 To 3
        Select Case Index
        Case 0
            table = "WEM"
        Case 1
            table = "WEMDATA"
        Case 2
            table = "WINNT"
        Case 3
            table = "PROGRAM FILES"
        End Select
        
'        MsgBox "DATABASE OPENED"

        Set locrec = dcollect_db.OpenRecordset(table, dbOpenSnapshot)
        locrec.MoveLast
        locrec.MoveFirst
        Do Until locrec.EOF
            DoEvents
'            MsgBox locrec.Fields("PATH")
            If IsNull(locrec.Fields("computer")) = True Or InStr(1, comp, locrec.Fields("computer"), vbTextCompare) > 0 Or Len(locrec.Fields("computer")) = 0 Or Len(gbl_jobname) > 0 Or UCase(Mid(job, 1, 1)) = "I" Then
'                MsgBox "TAG3"
                If locrec.Fields("Copy File") = True Then
'                    MsgBox "TAG4"
                    sp = InStr(1, UCase(locrec.Fields("path")), "JOBNUMBER")
'                    g_oAppLog.AddEntry "Path-" & UCase(locrec.Fields("path"))   '2.30
                    If sp > 0 Then
'                        MsgBox "TAG5"
                        If Len(gbl_jobname) = 0 Then
                            srcpath = "\\" & job & comp & "\" & table
                        Else
                            srcpath = "\\" & comp & "\" & table
                        End If
                        If Len(gbl_jobname) = 0 Then
                            srcpath = fs.BuildPath(srcpath, job)
                            srcpath = fs.BuildPath(srcpath, Mid(locrec.Fields("path"), sp + 10, Len(locrec.Fields("path"))))
                        Else
                            srcpath = fs.BuildPath(srcpath, gbl_jobname)
                            srcpath = fs.BuildPath(srcpath, Mid(locrec.Fields("path"), sp + 10, Len(locrec.Fields("path"))))
                        End If
                        Form1.Text1.Text = "Looking " & srcpath '2.30
                        DoEvents
                        filefound = Dir(srcpath)
                    Else
'                    MsgBox "TAG6"
                        If Len(gbl_jobname) = 0 Then
                            Form1.Text1.Text = "Dir of " & "\\" & job & comp & "\" & table & "\" & locrec.Fields("path") '2.30
                            DoEvents
                            filefound = Dir("\\" & job & comp & "\" & table & "\" & locrec.Fields("path"))
                        Else
                            Form1.Text1.Text = "Dir of " & "\\" & comp & "\" & table & "\" & locrec.Fields("path") '2.30
                            DoEvents
                            filefound = Dir("\\" & comp & "\" & table & "\" & locrec.Fields("path"))
                        End If
                        
                    End If
                    While Len(filefound) > 0
                        '2.30
                        Dim jk As Integer
                        Dim rTime As Long
                        Randomize
                        rTime = Rnd(10) * 10
                        For jk = 1 To rTime
                            Sleep 10 'up to .1 sec wait
                            DoEvents
                        Next jk
                        'g_oAppLog.AddEntry "Filefound-" & filefound & " random " & rTime   '2.30
                        temp = locrec.Fields("path")
                        destpath = cfg_destdrive & "\REDUNDANCY\" & comp & "\" & table
                        If Len(gbl_jobname) = 0 Then
                            srcpath = "\\" & job & comp & "\" & table
                        Else
                            srcpath = "\\" & comp & "\" & table
                        End If
                        If InStr(1, temp, "jobnumber", vbTextCompare) > 0 Then          '2.07
                            '2.25 Dim wwfolder(20) As String
                            Dim wwfolder(1500) As String
                            Dim X
                            X = 0
                            If Len(job) > 0 Then
                                Form1.Text1 = "Dir " & srcpath & "\" & job   '2.30
                                wwfolder(X) = Dir(srcpath & "\" & job & "*", vbDirectory)
                            Else
                                Form1.Text1 = "Dir " & srcpath & "\" & gbl_jobname
                                wwfolder(X) = Dir(srcpath & "\" & gbl_jobname & "*", vbDirectory)
                            End If
                            Do Until Len(wwfolder(X)) = 0
                                DoEvents
                                X = X + 1
                                wwfolder(X) = Dir()
                                
                                '2.25
                                If X >= 1499 Then
                                    g_oAppLog.AddEntry "Incomplete folder backup. More than 1500 files in " & srcpath & "\" & job & "*"
                                    Exit Do
                                End If
                                
                            Loop
                            
                            For X = X - 1 To 0 Step -1
                                Form1.Text1 = "Copywwfolder " & srcpath & "\" & wwfolder(X)   '2.30
                                DoEvents
                                If InStr(Form1.Text1, "_") <= 0 Then    '2.30 don't back up ww old folders with _ in them
                                    Randomize
                                    rTime = Rnd(10) * 100 + 2  'wait different times to back up this large folder
                                    For jk = 1 To rTime
                                        Sleep 100 'up to 10 seconds sec wait
                                        DoEvents
                                    Next jk
                                    fs.CopyFolder srcpath & "\" & wwfolder(X), destpath & "\"
                                End If
                                filefound = ""
                            Next
                        Else
                            Form1.Text1 = "Dest " & destpath & " Source " & srcpath   '2.30
                            Do
                                DoEvents
                                sp = InStr(1, temp, "\")
                                If sp > 0 Then
                                    newdir = Mid(temp, 1, sp - 1)
                                Else
                                    newdir = ""
                                End If
                                If UCase(newdir) = "JOBNUMBER" Then newdir = job
                                destpath = fs.BuildPath(destpath, newdir)
                                srcpath = fs.BuildPath(srcpath, newdir)
                                fs.CreateFolder destpath
                                fs.CreateFolder srcpath
                                temp = Mid(temp, sp + 1, Len(temp))
                                sp = InStr(1, temp, "\")
                            Loop Until sp = 0
                            Form1.Text1 = "Copying " & srcpath & "\" & filefound    '2.30
                            ATTR = GetAttr(srcpath & "\" & filefound) And vbReadOnly    ' Returns 2.
    '                        If ATTR > 0 Then
    '                            MsgBox ("Warning!!  The file " & srcpath & "\" & filefound & " has Read Only Attributes.")
    '                        End If
                            '2.19 If UCase(filefound) = "WEMSQL.MDF" And cfg_SQLMAIN_Mode > 0 Then   '2.00
                            If (UCase(filefound) = "WEMSQL.MDF" Or UCase(filefound) = "WEMASQL.MDF") And cfg_SQLMAIN_Mode > 0 Then
'                            MsgBox "tag0"
                                'Dim WS As Workspace    '2.00
                                Dim db As dao.Database    '2.00
                                Dim sFinalFullPath_RelApp As String    '2.00
                                 'Set WS = Workspaces(0)    '2.00
                                Dim existingbackupfile As String        '2.22
                                
                                fs.CreateFolder destpath & "\temp"
                                existingbackupfile = Dir(destpath & "\*.SQB")
                                If fs.FileExists(destpath & "\" & existingbackupfile) = True Then
                                    fs.Deletefile (destpath & "\temp\*.*")
                                    fs.MoveFile destpath & "\" & "*.sqb", destpath & "\temp\"
                                End If
                                'V2.20
                                If cfg_System = stFeed Or cfg_System = stUnknown Then
                                    Err.Clear
                                    Set db = oSQLTK.OpenDBSvr("", job & comp & "\WEMSQLENGINE", cfg_SQLMAIN_WEMAdminUserID, cfg_SQLMAIN_WEMAdminPWD, "WEMSQL")    '2.00
                                    If db Is Nothing Then
                                        cfg_System = stUnknown
                                    Else
                                        cfg_System = stFeed
                                    End If
                                    If Len(cfg_SQLMAIN_ARCHIVE_DSN) > 0 And cfg_System = stFeed Then
'                                        MsgBox "tag1"
                                        Set dbArchive = oSQLTK.OpenDBSvr(cfg_SQLMAIN_ARCHIVE_DSN, cfg_SQLMAIN_OverrideSvrName, _
                                        cfg_SQLMAIN_WEMAdminUserID, cfg_SQLMAIN_WEMAdminPWD, cfg_SQLMAIN_OverrideDBName)    '2.46
                                    End If
'
                                    
                                End If
                                If cfg_System = stAsphalt Or cfg_System = stUnknown Then
                                    Err.Clear
                                    Set db = oSQLTK.OpenDBSvr("", job & comp & "\WEMSQLENGINE", cfg_SQLMAIN_WEMAdminUserID, cfg_SQLMAIN_WEMAdminPWD, "WEMASQL")    '2.00
                                    If Err.Number = 0 Then
                                        cfg_System = stAsphalt
                                    Else
                                        cfg_System = stUnknown
                                    End If
                                End If
                                'MsgBox "Backup" & cfg_System & " DB" & destpath & " " & db.Name & job & comp & "\WEMSQLENGINE"
                                
                                g_oAppLog.AddEntry "making wemsql backup"

                                sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(db, , , destpath, BACKUP_STATUS_Consistent, , sql_BaseFolder)  '2.00 '2.21
                                If Len(cfg_SQLMAIN_ARCHIVE_DSN) > 0 And cfg_System = stFeed Then
'                                    MsgBox "tag2"
g_oAppLog.AddEntry "checking wemsqlarchive backup"
Dim rslt
'                                    rslt = Form1.makeorcopyarchive(srcpath & "\sql", destpath)
'MsgBox srcpath
'MsgBox destpath

                                    rslt = Form1.makeorcopyarchive(srcpath, destpath)

                                    
'                                    sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(dbArchive, , , destpath, BACKUP_STATUS_Consistent, , sql_BaseFolder)  '2.00 '2.21
                                    
'                                    MsgBox "fs=" & sFinalFullPath_RelApp
'                                    MsgBox "bf=" & sql_BaseFolder
                                End If
                                Sleep 2000          '2.22
                                If Len(Dir(destpath & "\*.SQB")) > 0 Then
                                    fs.Deletefile destpath & "\temp\*.SQB", True
                                Else
                                    fs.MoveFile destpath & "\temp\" & "*.sqb", destpath & "\"
                                    g_oAppLog.AddEntry "database backup failed - keeping old one"
                                End If
                            Else    '2.00
                                If fs.FileExists(destpath & "\" & filefound) = True Then
                                    Set dfile = fs.GetFile(destpath & "\" & filefound)
                                    Set sfile = fs.GetFile(srcpath & "\" & filefound)
                                    If dfile.DateLastModified <> sfile.DateLastModified Then
                                        fs.CopyFile srcpath & "\" & filefound, destpath & "\" & filefound
                                    End If
                                ElseIf InStr(1, filefound, ".sqb", vbTextCompare) > 0 Then
                                    If InStr(1, filefound, "__archive_", vbTextCompare) > 0 Then
'                                        MsgBox "destpath=" & destpath
                                        del_old_archive_backups destpath, "__archive"
                                        fs.CopyFile srcpath & "\" & filefound, destpath & "\" & filefound
                                    Else
                                        fs.CopyFile srcpath & "\" & filefound, destpath & "\" & filefound
                                    
                                    End If
                                    
                                Else
                                    fs.CopyFile srcpath & "\" & filefound, destpath & "\" & filefound
                                End If
                            End If
                            filefound = ""
                            filefound = Dir()
                        End If
                    Wend
                End If
'                MsgBox "TAG15"
                Form1.Text1.Text = ""
'                MsgBox "BOOM"
            End If
            locrec.MoveNext
        Loop
        locrec.Close
    Next
    Form1.MousePointer = vbNormal
End Function
Private Sub del_old_archive_backups(pathtosearch As String, criteria As String)
Dim fname As String
Dim fso As New FileSystemObject
g_oAppLog.AddEntry "deleting archived sqb file - there is a new one"
fname = Dir(pathtosearch & "\*.sqb")
Do Until Len(fname) < 3
    If InStr(1, fname, criteria, vbTextCompare) > 0 Then
        fso.Deletefile pathtosearch & "\" & fname, True
    End If
    fname = Dir()
Loop
End Sub
Sub del_old_sql_backups(pathtosearch As String)
Dim fname As String
Dim fso As New FileSystemObject
g_oAppLog.AddEntry "deleting wem main db sqb file going to make a new one"
fname = Dir(pathtosearch & "\*.sqb")
Do Until Len(fname) < 3
    If InStr(1, fname, "__archive", vbTextCompare) = 0 Then
        fso.Deletefile pathtosearch & "\" & fname, True
    End If
    fname = Dir()
Loop
End Sub


Private Sub CreateDESTDir(cfg_destdrive As String, COMPUTER As String)
Dim ListCount As Long
Dim Index As Integer
Dim fp As Integer
Dim jobname2 As String
Dim fielddatadate As String
On Error Resume Next
jobname2 = "REDUNDANCY"
fielddatadate = ""
'MsgBox "CREATE DIRECTORY"
'MsgBox "COMPUTER"
MkDir cfg_destdrive & "\" & jobname2
MkDir cfg_destdrive & "\" & jobname2 & fielddatadate & "\" & COMPUTER
MkDir cfg_destdrive & "\" & jobname2 & fielddatadate & "\" & COMPUTER & "\WEM"
MkDir cfg_destdrive & "\" & jobname2 & fielddatadate & "\" & COMPUTER & "\WINNT"
MkDir cfg_destdrive & "\" & jobname2 & fielddatadate & "\" & COMPUTER & "\WEMDATA"
MkDir cfg_destdrive & "\" & jobname2 & fielddatadate & "\" & COMPUTER & "\PROGRAM FILES"
'MsgBox "DIRECTORY CREATED"
    
    
End Sub

Function ReturnIniValue(FileIniName As String, WemApp As String, WemProfile As String, Default As Variant, Optional do_not_writedefault As Boolean) As Variant  '2.08.0
Dim fs As Integer
Dim tp As String
 
    On Error Resume Next
    tp = Space$(400)
    fs = GetPrivateProfileString(WemApp, WemProfile, "", tp, 400, FileIniName)
    If fs <= 0 Then
        If do_not_writedefault = False Then WritePrivateProfileString WemApp, WemProfile, CStr(Default), FileIniName
        ReturnIniValue = Default
    Else
        ReturnIniValue = Left(tp, fs)
    End If
End Function
  
