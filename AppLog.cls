VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Version History:
'1.00.00 ~03/01/01 REA
'         - Initial Release (first used in TaskServer component of Asphalt4000)
'2.00.00  10/20/03 REA
'         - Added separate logs per calendar day with purging, INI configuration options.
'         - Improved the wait during retry of disk operations.
'2.01.00  8/19/04  REA
'         - Removed use of VB Dir() command and replaced with API calls (else, if client is using
'           VB Dir(), the client code can become confused)
'         - Reduced number of disk retries and duration to 3 retries of no more than total of 60 ms
'           (To reduce possibly locking up an app due to client passing bad path, etc.)

'======================================================================================
'Description:  Self contained application logging module.
'Sample use:
'   (1)  In a module, create a global instance of clsAppLog
'        Public g_oAppLog As New AppLog
'
'   (2)  Optionally set the INI file name(else will default to <VB App.EXEName>.INI)
'        g_oAppLog.INIFileName = "Loadout.INI"

'   (3) When desired, send a string to the log
'       g_oAppLog.AddEntry "Some information to be logged."
'
'       This will append the following to the log file:
'       05/29/02 12:31:40  Some information to be logged.
'
'   (4) By default, log files are placed in a "Log" subfolder of the executing application.
'
'   (5) If "module level" mode flags are in the INI, AddEntry2 will check these before
'       deciding to log the information.  AddEntry2 will also (by default) generate a header
'       for the entry:
'       g_oAppLog.AddEntry2 "Some information to be logged, "MyModuleName"
'
'       This will append the following to the log file:
'
'       05/29/03 12:31:40  <From APP='MyExeName',MODULE='MyModuleName':
'       05/29/03 12:31:40  Some information to be logged
'       05/29/03 12:31:40  >
'
'       If a procedure name is passed, this will also be added to the header:
'       g_oAppLog.AddEntry2 "Some information to be logged, "MyModuleName", "MyProcedureName"
'
'       This will append the following to the log file:
'
'       05/29/03 12:31:40  <From APP=MyExeName, MODULE=MyModuleName, PROC=MyProcedureName:
'       05/29/03 12:31:40  Some information to be logged
'       05/29/03 12:31:40  >
'
'
'   (6) Optional flags for calls to AddEntry and AddEntry2:
'          bForceOutput:
'              For calls to AddEntry:  True:  forces output to file even if LogMode is Off (default is false)
'              For calls to AddEntry2:  True:  forces output to file even if LogMode and MODULE_MODE flags are off.
'                  (default is False)
'          bNewLine:  For calls to AddEntry and AddEntry2, if true (default), the new entry starts on a new line.
'          bSourceHeader:  For calls to AddEntry2, if true (default), new entry is prefixed by a
'              source header of the form:  "<From APP=App.EXEName, MODULE=ModuleName, PROC=ProcedureName:"
'
'
'   (7) Log file name format is similar to:  LOG 2003-10-20 FileName.Ext
'       where FileName defaults to the AppEXEName and Ext defaults to .txt.
'       These can be overridden using INI settings.
'
'   (8) Log files for prior calendar days are purged once per day (with the first call to AddEntry or AddEntry2).
'       By default, 60 days of logs are maintained.
'
'   (9) Retry mechanism allows multiple apps to hit the same log file with limited chance for loss of data.
'
'
'INI Configuration:
'   (1) INI file is not read until AddEntry or AddEntry2 are called.  This
'       allows specifying the location of the INI file.  For example, the INI file can be redirected to
'       be the same as the existing application INI file (e.g., Loadout.INI, Batching.INI or similar)
'
'   (2) INI settings are cached in memory after being read the first time.
'
'   (3) Code has been added to convert "empty" INI settings to their default values (normally, blank
'       INI settings become empty strings).
'
'   (4) Default configuration section.  (The following will work as an actual INI file if remove the
'       leading left margin):
'
'        [AppLog]
'
'        ; master mode setting for logging.
'        ; -1 = suppresses all logging.
'        ;  0 = suppress logging (except where bForceOutput flag has been used)
'        ;  1 = logging is on (except where MODULE_MODE flags are explicitly set to 0) (default)
'        LogMode=1
'
'        ; # of days of log files to keep.  older ones are purged on first call to AddEntry
'        ; or AddEntry2 for each calendar day.  (Default=60)
'        KeepOldLogsDays=60
'
'        ; maximum log file size before archiving (1 level of archiving is maintained per day)
'        ; (Default=500000)
'        MaxLogFileBytes=500000

'        ; Folder where all log files will be kept (best to *leave blank* so will default to <AppFolder>\Log)
'        LogFolder=
'
'        ; default file name without extension (best to *leave blank* so will default to VB App.EXEName
'        ; which will be unique for different apps.  File name format will be LOG YYYY-MM-DD + name.extension)
'        LogFileName=
'        ; default file extension (default is .txt)
'        LogFileExt=
'
'        ; Optional module level flags have the form "MODULE_MODE.modulename".  These are
'        ; used by AddEntry2 to decide if debugging should be on/off for a specific module.
'        ; Module names can be any convenient identifier for a section of code.
'        ; 0 = Don't log this module
'        ; 1 = log this module (only if master LogMode=1)
'
'        ; set logging for a module named "MyModule" to ON
'        MODULE_MODE.MyModule1=1
'        ;set logging for a module named "AnotherModule" to OFF
'        MODULE_MODE.AnotherModule=0

'======================================================================================
Private Const INI_SECTION_NAME As String = "AppLog"
Private Const MODULE_MODE_INI_PREFIX As String = "MODULE_MODE"
Private m_sINIFileNameAndExt As String
Private m_bInitCompleted As Boolean
Private m_dateLastPurge As Date  'date of last purge of old log files
Private m_colCachedINISettings As Collection
Private m_sLogFolder As String
Private m_lMaxLogFileBytes As Long
Private m_lMaxNumDiskRetries As Long
Private m_lKeepOldLogsDays As Long
Private m_lLogMode As Long
Private m_sLogFileName As String
Private m_sLogFileExt As String
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'Note:  this routine is used in error logging.  As result, it can't use generic
'  error display mechanism or recursive loop could result.
'  MessageBox API is used for displaying messages.
'  Errors are unlikely.
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Const MB_OK = &H0&
Private Const MB_TASKMODAL = &H2000&
Private Const MB_TOPMOST = &H40000       'from winuser.h
'========================================================================================
'Declarations used in Dir_API to emulate VB Dir() command

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'========================================================================================

'For future use (do not remove)
Public Event GenericEvent(ArrayInVariant As Variant)

Public Function GenericFn(ParamArray Params() As Variant) As Variant
'For future use (do not remove)
End Function

'========================================================================================
Public Property Get INIFileNameAndExt() As String
INIFileNameAndExt = m_sINIFileNameAndExt
End Property

Public Property Let INIFileNameAndExt(ByVal vNewValue As String)
m_sINIFileNameAndExt = vNewValue
End Property
'========================================================================================
Private Sub Class_Initialize()
On Error GoTo Class_Initialize_Err
'NOTE: Don't do anything here that requires the INI file as client must be able to redirect this.

'following not set to true until AddEntry or AddEntry2 is called
m_bInitCompleted = False

Set m_colCachedINISettings = New Collection

'set default INI location (client can override before logging actually occurs)
m_sINIFileNameAndExt = App.EXEName & ".INI"

'set max retries for disk operations
'(based on testing, this value appears to be adequate for normal use.)
'V2.01:  Reduced to 3 (previously 5)
m_lMaxNumDiskRetries = 3

Exit Sub
Class_Initialize_Err:
'(must handle errors in class init else VB degrades msg to 'automation error')
MessageBox 0, "Error in logging routine (AppLog.Class_Initialize), #" & Err.Number & ":  " _
    & vbCrLf & vbCrLf & Err.Description, "Error", MB_OK Or MB_TASKMODAL Or MB_TOPMOST
End Sub

'reads data from INI.  Only called from AddEntry/AddEntry2 so clients can redirect INI file location before
'any logging/purgins is done.
Private Function DoInitialization()
On Error Resume Next

'only do this once to conserve on processing time
If m_bInitCompleted Then
    'exit early
    GoTo Proc_Exit
End If

'initialize necessary values
m_lLogMode = Val(GetAppINISetting(INI_SECTION_NAME, "LogMode", "1"))

m_lKeepOldLogsDays = Val(GetAppINISetting(INI_SECTION_NAME, "KeepOldLogsDays", "60"))

'get folder for logging (use <AppFolder>\Log as default)
m_sLogFolder = GetAppINISetting(INI_SECTION_NAME, "LogFolder", App.Path & "\" & "Log")

'remove any trailing slash from the log folder (in case user specified)
If Right(m_sLogFolder, 1) = "\" Then
     m_sLogFolder = Left(m_sLogFolder, Len(m_sLogFolder) - 1)
End If

'create the log folder (if necessary)
CreateFolderIfNec m_sLogFolder
'-------------

m_lMaxLogFileBytes = Val(GetAppINISetting(INI_SECTION_NAME, "MaxLogFileBytes", "500000"))

m_sLogFileName = GetAppINISetting(INI_SECTION_NAME, "LogFileName", App.EXEName)

m_sLogFileExt = GetAppINISetting(INI_SECTION_NAME, "LogFileExt", ".txt")

'prevent re-execution
m_bInitCompleted = True

Proc_Exit:
Err.Clear
End Function

'get the log file path (based on the current INI setting)
Public Function GetLogFilePath() As String
On Error Resume Next
'make sure values have been initialized
DoInitialization
GetLogFilePath = CreateLogFullPath(m_sLogFolder, m_sLogFileName, m_sLogFileExt, False)
Err.Clear
End Function

Public Function AddEntry(ByVal sMessage As String, Optional ByVal bForceOutput As Boolean = False, _
                         Optional ByVal bNewLine As Boolean = True)
On Error Resume Next

'get INI values (if necessary)
DoInitialization

'if master flag suppresses all logging
If (m_lLogMode = -1) Then
    'exit early
    GoTo Proc_Exit
End If

'if logging is active OR forced output
If (m_lLogMode > 0) Or bForceOutput Then
    AddEntry = AddEntry_Generic(m_sLogFolder, m_sLogFileName, m_sLogFileExt, _
                            sMessage, bNewLine, True)
End If  'skip everything

Proc_Exit:
Err.Clear
End Function

Public Function AddEntry2(ByVal sMessage As String, _
                          Optional ByVal sModuleName As String = "", _
                          Optional ByVal sProcedureName As String = "", _
                          Optional ByVal bForceOutput As String = False, _
                          Optional ByVal bSourceHeader As Boolean = True, _
                          Optional ByVal bNewLine As Boolean = True) As String
On Error Resume Next
Dim sModuleINIKey As String

'get INI values (only if necessary)
DoInitialization

'if master flag suppresses all logging
If (m_lLogMode = -1) Then
    'exit early
    GoTo Proc_Exit
End If

'if not forced output
If Not bForceOutput Then
    'we need to ck the logging mode flags

    'if master active flag is off
    If (m_lLogMode = 0) Then
        'exit early
        GoTo Proc_Exit
    End If

    'create custom INI section name for this module
    sModuleINIKey = MODULE_MODE_INI_PREFIX & "." & sModuleName

    'if INI has 'active' setting for this module turned off
    '(if setting not present or blank, default to log)
    If (Val(GetAppINISetting(INI_SECTION_NAME, sModuleINIKey, "1")) = 0) Then
         'exit early
         GoTo Proc_Exit
    End If
    
End If  'else, forced output--ignore other mode flags

If bSourceHeader Then
    sMessage = "<From APP=" & App.EXEName _
          & IIf(sModuleName = "", "", ", MODULE=" & sModuleName) _
          & IIf(sProcedureName = "", "", ", PROC=" & sProcedureName) & ":" & vbCrLf _
          & sMessage & vbCrLf _
          & ">"
End If
'log the value and return what was actually added
AddEntry2 = AddEntry_Generic(m_sLogFolder, m_sLogFileName, m_sLogFileExt, sMessage, bNewLine, True)

'don't return error state to caller
Proc_Exit:
Err.Clear
End Function

Private Function AddEntry_Generic(sFolder As String, sFileName As String, sFileExt As String, _
                                  sMessage As String, bNewLine As Boolean, _
                                  bTimeStampPrefix As Boolean) As String
On Error GoTo AddEntry_Err
Dim sLogPath As String
Dim sArchiveLogPath As String
Dim lLogFileNum As Long
Dim lLastNumRetries As Long
Dim lPos As Long
Dim lLenTimeStamp As Long
Dim sTimeStamp As String

'create the main log and archive log file paths
sLogPath = CreateLogFullPath(sFolder, sFileName, sFileExt, False)
sArchiveLogPath = CreateLogFullPath(sFolder, sFileName, sFileExt, True)

'truncate the log file if necessary (includes transfer to 'Archive' log file)
TruncateLogFile sLogPath, sArchiveLogPath

lLogFileNum = FreeFile

'(this will create the file if it doesn't exist).  based on testing, get best performance
'if specify Shared.
Open sLogPath For Append Shared As lLogFileNum

If bNewLine Then
    sMessage = vbCrLf & sMessage
End If

If bTimeStampPrefix Then
    sTimeStamp = Format(Now, "mm/dd/yy hh:nn:ss") & "  "
    lLenTimeStamp = Len(sTimeStamp)
    'insert the timestamp after every CRLF (haven't used Replace as want vb5 compatibility)
    lPos = 1
    Do
       lPos = InStr(lPos, sMessage, vbCrLf, vbTextCompare)
       If lPos > 0 Then
            sMessage = Left(sMessage, lPos + 1) & sTimeStamp & Mid(sMessage, lPos + 2)
            'set next search position beyond the found CRLF (2 chars) and the inserted string(lLenTimeStamp chars)
            lPos = lPos + 2 + lLenTimeStamp
       End If
    Loop While lPos > 0
End If

Print #lLogFileNum, sMessage;

Close lLogFileNum

'return the string actually added
AddEntry_Generic = sMessage

Proc_Exit:
'purge old files (this is done *last* so timestamp of the log entry more closely matches when entry was requested)
PurgeOldLogFiles sFolder, sFileName, sFileExt

'make sure no VB error code is returned to caller
Err.Clear

Exit Function
AddEntry_Err:
'execute a retry
If RetryDiskOp(lLastNumRetries) Then
   Resume
End If
'make sure the file is closed on exit
CloseFileWithoutError lLogFileNum
End Function

Private Function CloseFileWithoutError(lFileNum As Long)
On Error Resume Next
Close lFileNum
Err.Clear
End Function

Private Function TruncateLogFile(sLogPath As String, sArchiveLogPath As String)
On Error Resume Next
Dim lFileLength As Long

lFileLength = FileLen(sLogPath)

'if the log file does not exist (or is empty or cannot be read)
If lFileLength = 0 Then
   'file is empty or cannot be read--no need to truncate it.  exit early
   GoTo Proc_Exit
End If

'check if file length exceeds max length
If lFileLength > m_lMaxLogFileBytes Then
   'move the current log file to the 'Archive' log file
   MoveToArchiveLogFile sLogPath, sArchiveLogPath
End If

Proc_Exit:
'make sure no error code is returned to caller
Err.Clear
End Function

'NOTE:  File name format must be same as that used in proc PurgeOldLogFiles
Private Function CreateLogFullPath(sFolder As String, sFileName As String, sFileExt As String, bIsArchive As Boolean) As String
On Error Resume Next
CreateLogFullPath = sFolder & IIf(Right(sFolder, 1) = "\", "", "\") & "LOG " _
                & Format(Date, "YYYY-MM-DD") & " " & CreateLogFileNameExtSuffix(sFileName, sFileExt, bIsArchive)
Err.Clear
End Function

Private Function CreateLogFileNameExtSuffix(sFileName As String, sFileExt As String, bIsArchive As Boolean) As String
On Error Resume Next
CreateLogFileNameExtSuffix = sFileName & IIf(bIsArchive, " OLD", "") & IIf(Left(sFileExt, 1) = ".", "", ".") & sFileExt
Err.Clear
End Function

'returns number of files actually purged
Private Function PurgeOldLogFiles(sFolder As String, sFileName As String, sFileExt As String) As Long
On Error Resume Next
Dim sFileNameExt As String
Dim sFilePath As String
Dim sOldestDatePath As String
Dim iAttr As Integer

'if purge has already been done today
If Date = m_dateLastPurge Then
    'exit early--don't purge more than 1 time per day
    GoTo Proc_Exit
End If

'calc path prefix for oldest allowable file
sOldestDatePath = UCase(sFolder & IIf(Right(sFolder, 1) = "\", "", "\") _
                     & "LOG " & Format(DateAdd("d", -1 * m_lKeepOldLogsDays, Date), "YYYY-MM-DD"))

'cycle through all files in the folder that have the required format.
'(1) To reduce number of files tested, only look at log files that have the same suffix as the file we're currently
'    logging to.  This will limit the number of files tested to 60 (or whatever the purge time is).
'    Include the "*" to catch files with the "OLD" suffix.
'(2) Note that Read Only files are returned by Dir_API (so these will be purged)
'V2.01:  Replaced VB Dir with API calls.
sFileNameExt = Dir_API(m_sLogFolder & "\" & "LOG ????-??-?? " & CreateLogFileNameExtSuffix(sFileName & "*", sFileExt, False))

Do While sFileNameExt <> ""
    'create the full path to this file
    sFilePath = UCase(m_sLogFolder & "\" & sFileNameExt)

    'if file name with this pattern is older than oldest allowed date
    If sFilePath < sOldestDatePath Then
        'remove readonly flag (if present)
        iAttr = GetAttr(sFilePath)
        If iAttr And vbReadOnly Then
            SetAttr sFilePath, (iAttr - vbReadOnly)
        End If
        'delete the file
        Kill sFilePath
    End If
    'get next file
    sFileNameExt = Dir_API
Loop

'update timestamp of last purge (so purge not repeated again today)
m_dateLastPurge = Date

Proc_Exit:
'don't return error state to caller
Err.Clear
End Function

Private Sub Class_Terminate()
On Error GoTo Class_Terminate_Err
'currently does nothing

Exit Sub
Class_Terminate_Err:
'(must handle errors in terminate event as it is event driven)
MessageBox 0, "Error in logging routine (AppLog.Class_Terminate), #" & Err.Number & ":  " _
    & vbCrLf & vbCrLf & Err.Description, "Error", MB_OK Or MB_TASKMODAL Or MB_TOPMOST
End Sub

'works with files or folders
Private Function FileExists(sFileFullPath As String) As Boolean
On Error Resume Next
GetAttr sFileFullPath
FileExists = (Err.Number = 0)
'make sure no error is returned to caller
Err.Clear
End Function

Private Sub MoveToArchiveLogFile(sLogPath As String, sArchiveLogPath As String)
Dim lLastNumRetries As Long
'--------------
'delete the 'Archive' log file (may or may not exist)
On Error Resume Next
Kill sArchiveLogPath
'--------------
'turn error handler on so can get retry
On Error GoTo EH
'rename the current log file to the 'Archive' file name
Name sLogPath As sArchiveLogPath
Exit Sub
EH:
'if get error that file exists
If Err.Number = 58 Then
    'assume that another app has executed the archive.  exit without retrying as
    'rename will not succeed.
    Exit Sub
End If
'execute a retry
If RetryDiskOp(lLastNumRetries) Then
   Resume
End If  'else, just fall through and ignore the error
End Sub

'used in multiple places to allow retry of a disk operation. (based on testing, code in this
'module may be re-entered or multiple apps may attempt to access the same file.  for this reason,
'use a retry loop.  if retry fails, code exits without errors to prevent user confusion.
Private Function RetryDiskOp(ByRef lLastNumRetries As Long, Optional ByVal lMaxNumRetries As Long = -1) As Boolean
On Error GoTo EH
Dim lWaitMs As Long
Dim lTimeStamp As Long

'for testing only.
'Debug.Print "RetryDiskOp executed with lLastNumRetries = " & lLastNumRetries & " at " & Now

'set default return value => no more retries left
RetryDiskOp = False

'initialize max number of retries
If lMaxNumRetries < 0 Then
   lMaxNumRetries = m_lMaxNumDiskRetries
End If

lLastNumRetries = lLastNumRetries + 1
If lLastNumRetries > lMaxNumRetries Then
   'exit early with false
   'for testing only:
   'debug.Print "Max number of retries exceeded in clsAppLog.RetryDiskOp"
   Exit Function
End If
'if this is the first try
If lLastNumRetries = 1 Then
   'initialize the random number generator
   Randomize
End If
'--------------
'calc random number of Ms to wait--wait longer with each retry.  with max retries of 5, max wait will
'be about 110ms on the last try.  absolute maximum delay in any case (5 tries) will be 30+50+70+90+110=350msec.
'average delay on failure with 5 tries will be about 175msec.
'V2.01:  Reduced retry time to max of 60 msec out of concern that a client app could be locked if bad path passed and
'        retry mechanism executes repeatedly.
lWaitMs = Int(20 * Rnd())
lTimeStamp = timeGetTime
Do While GetElapsedMs(lTimeStamp) < lWaitMs
    'note:  use of DoEvents/Sleep in this loop could confuse client code and probably is best left out.
Loop
'--------------
'return true => can try again
RetryDiskOp = True

Exit Function
EH:
MessageBox 0, "Error in logging routine (AppLog.RetryDiskOp), #" & Err.Number & ":  " _
    & vbCrLf & vbCrLf & Err.Description, "Error", MB_OK Or MB_TASKMODAL Or MB_TOPMOST
End Function

'NOTE:  (1) THIS CACHES INI SETTINGS IN MEMORY AFTER RETRIEVING THE FIRST TIME
'           (to avoid perfmance impact if called frequently)
'       (2) If empty string is returned, returned value is converted to the default.
'           (this is to handle situations where INI contains "Key="--which will be
'           returned as blank string even if sDefault is specified.
Private Function GetAppINISetting(sSection As String, sKey As String, sDefault As String) As String
On Error Resume Next
Dim sINIPath As String
Dim sValue As String
Err.Clear
'attempt to get value from cache
sValue = m_colCachedINISettings(sSection & vbTab & sKey)
'if not found in cache
If Err.Number <> 0 Then
    'attempt to get from INI file
    sINIPath = GetWinDirWSlash & m_sINIFileNameAndExt
    GetPrivateIniSetting sSection, sKey, sDefault, sValue, sINIPath
    'if blank value is returned
    If sValue = "" Then
        'convert to the default (API won't do this if the Key=<blank> is present)
        sValue = sDefault
    End If
    'add this value to the collection
    m_colCachedINISettings.Add sValue, sSection & vbTab & sKey
Else
    'if blank value is returned
    If sValue = "" Then
        'convert to the default
        sValue = sDefault
    End If
End If
'return value to caller
GetAppINISetting = sValue
Err.Clear
End Function

Private Function StripAtFirstNull(sNullTerminated As String) As String
On Error Resume Next
Dim lNullPos As Long
lNullPos = InStr(sNullTerminated, Chr$(0))
If lNullPos > 0 Then
   StripAtFirstNull = Left$(sNullTerminated, lNullPos - 1)  'strip off terminating null
Else
   StripAtFirstNull = sNullTerminated
End If
Err.Clear
End Function

'buffer size has been made very large (5K) to accomodate saving large amounts of data to ini file.
Private Function GetPrivateIniSetting(sSection As String, sKey As String, sDefault As String, _
   ByRef sReturnValWNullStripped As String, sIniFullPath As String) As Boolean
On Error Resume Next
Dim sWinReturnValue As String * 512, lResult As Long
lResult = GetPrivateProfileString(sSection, sKey, sDefault, sWinReturnValue, 512 - 1, sIniFullPath)
If (lResult = 0) Then
    sReturnValWNullStripped = ""
    GetPrivateIniSetting = False
    Exit Function
End If
sReturnValWNullStripped = StripAtFirstNull(sWinReturnValue)
GetPrivateIniSetting = True
Err.Clear
End Function

Private Function WritePrivateIniSetting(sSection As String, sKey As String, _
   sWriteValue As String, sIniFullPath As String) As Boolean
On Error Resume Next
Dim lResult As Long
lResult = WritePrivateProfileString(sSection, sKey, sWriteValue, sIniFullPath)
WritePrivateIniSetting = (lResult <> 0)
Err.Clear
End Function

Private Function GetWinDirWSlash() As String
On Error Resume Next
Dim lReturn As Long, sReturn As String * 255, sPath As String
GetWinDirWSlash = ""    'default => error
lReturn = GetWindowsDirectory(sReturn, 255 - 1)
If (lReturn = 0) Or (lReturn > (255 - 1)) Then
    Return  'if error, return empty string to calling routine.  Anticipate that this will be extremely rare.
End If
sPath = StripAtFirstNull(sReturn)   'strip off terminating null
If Right$(sPath, 1) <> "\" Then
   sPath = sPath & "\"  'add backslash if necessary
End If
GetWinDirWSlash = sPath
Err.Clear
End Function

'will create nested folder path.  returns true if created or if exists
Private Function CreateFolderIfNec(sFolderPath As String) As Boolean
On Error Resume Next
Dim lPos As Long
Dim sSubFolder As String

'do an initial ck for folder exists
If FileExists(sFolderPath) Then
    'exit early with true
    CreateFolderIfNec = True
    GoTo Proc_Exit
End If
sSubFolder = sFolderPath
lPos = 0
Do
    'find the first or next backslash
    lPos = InStr(lPos + 1, sFolderPath, "\", vbTextCompare)
    'if found a backslash
    If lPos > 0 Then
        'get everything up to the backslash
        sSubFolder = Left(sFolderPath, lPos - 1)
    Else
        'use the whole passed string
        sSubFolder = sFolderPath
    End If
    
    'if doesn't exist
    If Not FileExists(sSubFolder) Then
        'create this folder
        Err.Clear
        MkDir sSubFolder
        'if unsuccessful
        If Err.Number <> 0 Then
             'exit with false
             GoTo Proc_Exit
        End If
    End If 'else, exists => don't need to create
'continue looping while backslashes are still present
Loop While lPos > 0

'if make it to here, return success
CreateFolderIfNec = True
Proc_Exit:
'don't return error state to caller
Err.Clear
End Function

'returns value of timeGetTime as a positive currency value (see preceeding fns for documentation)
Private Function timeGetTimeAsCurrency() As Currency
timeGetTimeAsCurrency = LongToUnsignedCurrency(timeGetTime())
End Function

'returns the number of millisec that have elapsed since a previous timestamp.  Note that this
'function can takes a raw value from timeGetTime as the input and does all necessary
'neg/wrap adjustments
Private Function GetElapsedMs(ByVal lStart_RawTimeGetTimeMs As Long)
Dim curStartMs As Currency
Dim curNowMs As Currency
'take care of timer wrap into negative range
curStartMs = LongToUnsignedCurrency(lStart_RawTimeGetTimeMs)
curNowMs = timeGetTimeAsCurrency
'adjust for wrap of timer around zero
AdjustForTimerWrap curNowMs, curStartMs
'return the elapsed time
GetElapsedMs = curNowMs - curStartMs
End Function

'function checks if the second millisec reading is greater than the first
'(timerGetTime wraps around to zero every 49 days).
'if not, then assume that the counter has wrapped around zero.
'return the adjusted second counter value in this case.
'method should work as long as the time difference between curMSecOlder and curMSecNewer is
' < 49 days.
Private Sub AdjustForTimerWrap(curMSecNewer As Currency, curMSecOlder As Currency)
'if impossible situation has resulted
If curMSecNewer < curMSecOlder Then
   'then assume that the counter has looped around zero.
   'Add 2^32 (@is currency datatype) to the "newer" counter to correct for this.
   curMSecNewer = curMSecNewer + 4294967296@
End If
'adjusted curMSecNewer is returned by reference
End Sub

'return values in Win32 that are unsigned longs are returned in VB as
'signed longs.  must convert these to unsigned currency to do valid
' < and > comparisons.  Applies to return value of GetTickCount() and
'timeGetTime() (both count up to (2^32 -1) millisec or 4,294,967,295 millisec = 49.7 days.  This
'is twice the max for + long =  2^31 -1 = 2,147,483,647.  The values beyond the + long range
'end up being returned in VB as a negative long.  Following fn converts these to
'unsigned longs stored in a currency data type (which has a + max of 922,337,203,685,477).
Private Function LongToUnsignedCurrency(lSigned As Long) As Currency
Dim curUnsigned As Currency
curUnsigned = CCur(lSigned)   'save input value as currency data type
If curUnsigned < 0 Then
   'if input value is negative, then convert to unsigned form by adding 2^32 (@ => currency)
   curUnsigned = curUnsigned + 4294967296@
End If
LongToUnsignedCurrency = curUnsigned
End Function

'==================================================================================================
'V2.01:  Function that emulates the VB Dir() fn.
'NOTES:  1.  Will *not* conflict with VB code that actually uses the Dir() function.  This has been tested.
'            However, if you call Dir_API from more than one place, it can become confused (just like the VB Dir).
'            If need to make overlapping calls, make another copy of the routine (e.g., Dir_API2).
'        2.  VB Dir() function will ignore Attributes if a PathName is not provided.  Code below does the same.
'        3.  VB Dir() appears to use separate call to GetVolumeInformation() if the vbVolume attribute is requested.
'            As this is rarely used and significantly complicates the code, this has been omitted.
Private Function Dir_API(Optional ByVal PathName As Variant, Optional ByVal Attributes As VbFileAttribute = vbNormal) As String
On Error Resume Next
Static lFileHdl As Long                 'holds handle for find
Static lAttrib_Target As Long           'holds desired attributes for the find
Dim lReturnFindNext As Long
Dim lpFindFileData As WIN32_FIND_DATA   'holds found file information
Dim lZeroPos As Long
Dim lAttrib_XOR As Long

'if no path is provided and no handle is available
If IsMissing(PathName) And (lFileHdl = 0) Then
    'clear everything out
    lFileHdl = 0
    lAttrib_Target = 0
    Dir_API = ""
    'immitate VB behavior by raising an error
    On Error GoTo 0
    Err.Raise 5
    '(will not reach here)
End If

'if caller provided a path (note that a blank string passed to VB Dir() causes the current
'Windows directory to be used and the code below does the same)
If Not IsMissing(PathName) Then
    'if we have an old handle
    If lFileHdl > 0 Then
        'make sure it's closed
        FindClose lFileHdl
    End If
    'save the passed attributes
    lAttrib_Target = Attributes
    'based on testing with VB dir, if provided path is blank string, then the current Windows path is used
    If PathName = "" Then
       'add backslash so next block will append *.*
       PathName = CurDir & "\"
    End If
    'based on testing with VB dir, if provided path ends in a "\" or ":", a "*.*" is added
    If (Right(PathName, 1) = "\") Or (Right(PathName, 1) = ":") Then
        PathName = PathName & "*.*"
    End If
    'try to get a new file handle
    lFileHdl = FindFirstFile(PathName, lpFindFileData)
    'if call fails
    If lFileHdl = -1 Then
        'reset everything and return empty string
        GoTo Proc_Exit
    End If
    'set fake return value so loop below executes
    lReturnFindNext = 1
Else
    'otherwise, use the existing handle to read in the next file
    lReturnFindNext = FindNextFile(lFileHdl, lpFindFileData)
End If

'calc the inverse of the 1st 5 bits and strip off the low bit (VB Dir ignores the read only flag, archive, and
'some other obscure attribute flags)
lAttrib_XOR = ((lAttrib_Target Xor 31) And 30)

Do While (lReturnFindNext <> 0) And (lReturnFindNext <> ERROR_NO_MORE_FILES)

    'get the 1st 5 bits of the file attribute.  AND with the XOR of the target flags
    'to create the following rule:  if the bit is ON in the file and is OFF in the
    'target, then exclude the file.  (This method does not emulate VB Dir() handling
    'of the vbVolume flag--which is so rarely used it has been omitted here)
    
    'File  Target  (XOR Target)  File AND (XOR Target)
    ' 0       0         1              0  => include the file
    ' 0       1         0              0  => include the file
    ' 1       0         1              1  => exclude the file
    ' 1       1         0              0  => include the file

    If (((lpFindFileData.dwFileAttributes And 30) And lAttrib_XOR) = 0) Then
        'return this file name.  Keep the file handle and attributes for future cycles.
        lZeroPos = InStr(lpFindFileData.cFileName, Chr$(0))
        If lZeroPos > 0 Then
            Dir_API = Left$(lpFindFileData.cFileName, lZeroPos - 1)
        Else
            Dir_API = lpFindFileData.cFileName
        End If
        'exit early
        Err.Clear
        Exit Function
    End If  'we don't want this file
        
    'read in the next file
    lReturnFindNext = FindNextFile(lFileHdl, lpFindFileData)
Loop

Proc_Exit:

'if we have a handle
If (lFileHdl > 0) Then
   'be sure to close it
   FindClose lFileHdl
End If

'reset everything and return empty string
lFileHdl = 0
lAttrib_Target = 0
Dir_API = ""
Err.Clear
End Function
'==================================================================================================




