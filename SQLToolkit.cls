VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLToolkit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Description:  Class with library of functions for working with SQL Server 2000/2005 databases.
'The "master" copy is generally kept with WEMAdmin project.

'Version history:
'1.00.00  12/06/04 REA
'      - Initial release
'1.01.00  12/07/04 REA
'      - Removed some redundant code from ErrHandlerPrivate
'      - Changed default folder behavior for CreateDB()
'      - Added functions for encrypting/decrypting passwords
'1.02.00  03/01/05 REA
'      - Signficant revision.  Server name param for OpenDBSvr has been moved.
'      - Added functions for killing connections to a database, attaching/detaching database files.
'      - Corrected error in Set AppLogObj--wasn't using the passed AppLog object instance.
'1.03.00  03/02/05 REA
'      - Added functions for resolving PCName into IP address
'1.04.00  03/08/05 REA
'      - Optional automatic resolution of server PCName into IP addresses.
'      - Additional network diagnostics before connecting (ping server IP)
'1.05.00  03/18/05 REA
'      - Changes to support DAO 3.60 (casting sql_variants, remove use of .Filter)
'        NOTE:  DAO 3.6 CAN CRASH ENTIRE APP WITH *FATAL* EXCEPTION IF VARBINARY
'        DATATYPE IS RETURNED FROM SQL SERVER STORED PROC.  SQL_VARIANT RESULTS IN
'        'FIELD IS TOO LONG' ERROR.
'1.06.00  03/22/05 REA
'      - Changed OpenDBSvr2 so only uses DSN-less connection if need to resolve the server to an IP.
'        (There's an advantage to using DSN in the connection:  it's considered a "new" connection by the
'        JET cache and allows specifying a different user name/password/security settings).
'1.07.00  04/14/05 REA
'      - Added function for updating the 'Auto Start' property of the SQL Server service.
'1.08.00  06/21/05 REA
'      - Fixed issue:  DMO cmds that don't open a connection cannot recognize server name containing "(local)"
'      - Changed encryption key to static string (previously Format() could get confused if user changes Windows date format)
'1.09.00  09/16/05 REA
'      - Improved default value handling for creation of SQL Svr bit fields from Access boolean fields
'      - Added another option for copying of PK indexes to SQL Server.
'      - Fixed errors related to spaces in table names.
'      - Added more status logging options
'      - Added fns for creating GUIDs in VB code
'      - Replaced VB5 string manip fns with built in VB6 functions.
'1.10.00  09/23/05 REA
'      - Changed accidental references to "g_oAppLog" to "m_oAppLog"
'      - Added workaround for VB6 IDE bug that causes error #3228:  Selected collating sequence not supported..."
'1.11.00  09/26/05 REA
'      - Added SQL "rowversion" fields as field creation option (known in older versions of SQL Server as 'timestamp' fields)
'      - Enhanced clustered/nonclustered index choices.
'1.12.00  10/03/05 REA
'      - Added GetSQLSvrVersionString():  Returns full SQL Server version string
'      - Added GetDBSchemaModDTime(): For SQL2005, returns dtime of db last table schema change.
'                                     For Pre-SQL2005, returns psuedo dtime that behaves similarly.
'      - Added IsSQLSvr2000 and IsSQLSvr2005 for detecting server versions.
'1.13.00  10/04/05 REA
'      - Added GetFV/SetFV utility functions to this module
'1.14.00  10/06/05 REA
'      - Added setting JET ODBC ConnectionTimeout registry setting to 1 second in class initialize event (with error suppression)
'1.15.00  10/19/05
'      - Removed table schema manip code to separate classes.
'      - Added fix for orphaned connections created by SQLDMO (found that must explicitly close these)
'      - Added fix for JET connection cache confusion if don't specify database name in ODBC connection string.
'      - Added fix for read of registry DSN info if power user on Win2003Server.
'      - Fixed issue with MapDBUsersToSvrLogins ("database id not found" error).
'      - Added SQL2005 compatibility changes
'1.16.00  12/12/05
'      - Additional support functions for backup/restore
'1.17.00  12/14/05
'      - Fixed bug in proc DropDB:  previously would fail if space in database name
'      - Caught some additional places where database names containing single quotes could fail.
'1.18.00  02/10/05
'      - Added GetTableVersion function for detecting changes in data of a table.
'      - Improved stripping of bracketed info from error messages.

'NOTES:
'1.  Current options are:
'
'    LogErrors:
'         True:   errors are logged to file using contained AppLog instance.
'         False:  error logging to file is disabled
'
'    AppLogObj:
'         This is the instance of AppLog used to optionally log errors from this module.
'         Can manipulate default log file locations (etc.) by manipulating this instance
'         or can set to own AppLog instance.
'
'2.  Sample call for creating a SQL Server database. (By default, the .mdf/.ldf files are placed on the server instance PC at
'    F:\WEM\Data\SQL--this path is always with respect to SQL Server location.)

'Dim dbSvr As DAO.Database
'Dim oSQLTK As SQLToolkit
'Set oSQLTK = New SQLToolkit
'Set dbSvr = oSQLTK.OpendbSvr("MyDSNName","", "MyUserID", "MyPassword")
'oSQLTK.CreateDB dbSvr, "MyNewDatabaseName"
'
'3.  If main application has its own instance of AppLog, this can also be used by SQLToolkit as follows
'(otherwise, oSQLTK will use its default instance of AppLog with default settings).
'
'Set oSQLTK.AppLogObj = oMyAppLogInstance
'
'3.  Dependencies:
'    --Standard WEM VB5 modules:  AppLog.cls
'    --Microsoft DAO 3.6 Object Library
'    --Microsoft SQLDMO Object Library
'    --Microsoft Scripting Runtime
'4.  ExecTSQL requires SQL Server Transact SQL syntax.  It can be used to open "ordinary" recordsets for read-only
'    access.  Efficiency for typical data retrieval tasks has not been tested.
'5.  Class currently does *not* maintain any internal database or server state information.

'==========================================================================================
'All API declarations used by this module placed here so class is self sufficient.
'==========================================================================================
'General purpose API functions
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'=============================================================================
'For reading list of system DSN's from the registry
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                       ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, _
                       ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                       ByVal lpSubKey As String, ByVal ulOptions As Long, _
                       ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                       ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, _
                       lpData As Any, lpcbData As Long) As Long     ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Const MB_OK = &H0&
Private Const MB_APPLMODAL = &H0&
Private Const MB_TOPMOST = &H40000    'from winuser.h
Private Const MB_TASKMODAL = &H2000&
'=============================================================================
'additions to read/write JET ODBC ConnectionTimeout setting
Private Const REG_DWORD As Long = 4
Private Const ERROR_NONE = 0
Private Const REG_OPTION_NON_VOLATILE = 0
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
  "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
  ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
  As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
  As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
  "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
  String, ByVal lpReserved As Long, lpType As Long, lpData As _
  Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
  "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
  String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
  As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
  "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
  ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
  ByVal cbData As Long) As Long

'=============================================================================
'API's for adding/removing system DSN's
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long
Private Const ODBC_ADD_DSN = 1        ' Add user DSN
Private Const ODBC_ADD_SYS_DSN = 4    ' Add system DSN
Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) user DSN
Private Const ODBC_CONFIG_SYS_DSN = 5 ' Modify an existing system DSN
Private Const ODBC_REMOVE_DSN = 3     ' Remove user DSN
Private Const ODBC_REMOVE_SYS_DSN = 6  'Remove system DSN
Private Const vbAPINull As Long = 0&  ' NULL Pointer
'=============================================================================
'Declarations used to encrypt/decrypt passwords stored in .INI files
Public Enum EncryptionType_Enum
    ENCRYPTION_TYPE_RC2 = 0   'symmetric private key "block" algorithm.  based on testing, encrypted output is longer than the input.
                              'this is considered a stronger encryption method han RC4.  "Symmetric" means that the same key is used
                              'to both encrypt and decrypt the data.
    ENCRYPTION_TYPE_RC4 = 1   'symmetric private key "stream" algorithm.  based on testing, encrypted output is same length as input.
                              'this is considered a weaker encryption method than RC2.   "Symmetric" means that the same key is used
                              'to both encrypt and decrypt the data.
End Enum
Private Const CRYPTO_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const PROV_RSA_FULL         As Long = 1
Private Const CALG_MD5 = &H8003&    '     (ALG_CLASS_HASH | ALG_TYPE_ANY | ALG_SID_MD5)
Private Const CRYPT_NO_SALT As Long = &H10
Private Const CALG_RC2 = &H6602&    '     (ALG_CLASS_DATA_ENCRYPT|ALG_TYPE_BLOCK|ALG_SID_RC2)
Private Const CALG_RC4 = &H6801&    '     (ALG_CLASS_DATA_ENCRYPT|ALG_TYPE_STREAM|ALG_SID_RC4)
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal algID As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal algID As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, _
    ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, _
    ByRef pdwDataLen As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'=============================================================================
'For executing DOS command line:
Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
'=============================================================================
'Declares related to resolving a PCName to an IP address.
'Private Const IP_SUCCESS As Long = 0  '(dup declaration with next section)
Private Const MAX_WSADescription As Long = 256  '(dup declaration with next section)
 Private Const MAX_WSASYSStatus As Long = 128   '(dup declaration with next section)
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
'Private Const MIN_SOCKETS_REQD As Long = 1  '(dup declaration with next section)
'Private Const SOCKET_ERROR As Long = -1     '(dup declaration with next section)
'Private Const ERROR_SUCCESS As Long = 0  '(duplicate declaration in this module)

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32.dll" _
  (ByVal hostname As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)

Private Declare Function lstrlenA Lib "kernel32" _
  (lpString As Any) As Long

Private Declare Function WSAStartup Lib "wsock32.dll" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
   
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

Private Declare Function inet_ntoa Lib "wsock32.dll" _
  (ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, _
   ByVal Ptr As Long) As Long
   
Private Declare Function gethostname Lib "wsock32.dll" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long

'=============================================================================
'additional declares need to ping an IP address (reuses some declarations from preceding section)
Private Const IP_SUCCESS As Long = 0  'dup declaration with above section
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
'Private Const WS_VERSION_REQD As Long = &H101    '(dup declaration with previous section)
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = &HFFFFFFFF
'Private Const MAX_WSADescription As Long = 256  '(dup declaration with previous section)
'Private Const MAX_WSASYSStatus As Long = 128    '(dup declaration with previous section)

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
    
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
   
Private Declare Function inet_addr Lib "wsock32" _
   (ByVal s As String) As Long
'============================================================================
'additional declares for converting an IP address into a host name
'(reuses some functions from preceding sections)
Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
'Private Const WS_VERSION_REQD As Long = &H101  '(dup declaration with above sections)
'Private Const IP_SUCCESS As Long = 0           '(dup declaration with above sections)
'Private Const SOCKET_ERROR As Long = -1        '(dup declaration with above sections)
Private Const AF_INET As Long = 2

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long
   
Private Declare Function lstrlen Lib "kernel32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long
'=============================================================================
'declares for getting new GUIDs and converting to string
Private Type GUID
   Data1 As Long
   Data2 As Long
   Data3 As Long
   Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, _
                                                           ByVal cbMax As Long) As Long
    
'=============================================================================
'=============================================================================
'UDT for saving/restoring values of the Err object.
Private Type ErrClone
    Number As Long
    Source As String
    Description As String
    HelpFile As String
    HelpContext As Long
    LastDllError As String
End Type

'flags for backup status
Public Enum BackupStatus_Enum
    BACKUP_STATUS_Consistent = 0
    BACKUP_STATUS_Inconsistent = 1
End Enum

Private m_oFSO As FileSystemObject
Private m_oAppLog As AppLog
Private m_bLogErrors As Boolean
'saves last ODBC connection string passed to OpenDBSvr2
'(this may be altered if server name is resolved to IP address)
Private m_sOpenDBSvr2_OrigODBCConnStr As String

Public Property Get AppLogObj() As AppLog
Set AppLogObj = m_oAppLog
End Property
Public Property Set AppLogObj(ByVal vNewValue As AppLog)
Set m_oAppLog = vNewValue
End Property

Public Property Get LogErrors() As Boolean
LogErrors = m_bLogErrors
End Property
Public Property Let LogErrors(ByVal vNewValue As Boolean)
m_bLogErrors = vNewValue
End Property

'if ODBC connection string is changed so that IP is used for the server name,
'this holds the original passed connection string
Public Property Get OpenDBSvr2_OrigODBCConnStr() As String
OpenDBSvr2_OrigODBCConnStr = m_sOpenDBSvr2_OrigODBCConnStr
End Property
Public Property Let OpenDBSvr2_OrigODBCConnStr(ByVal vNewValue As String)
m_sOpenDBSvr2_OrigODBCConnStr = vNewValue
End Property

Private Sub Class_Initialize()
On Error GoTo EH

Set m_oAppLog = New AppLog
Set m_oFSO = New FileSystemObject

'----------------
'V1.14:  Lower connection timeout property to 1 second so that old connections can be quickly cleared from the
'JET connection cache.  Suppress errors (as user may or may not have registry editing rights).
m_bLogErrors = False
On Error Resume Next
If GetODBCConnTimeoutSec_Jet40 <> 1 Then
    SetODBCConnTimeoutSec_Jet40 1
End If
On Error GoTo EH
m_bLogErrors = True
'----------------

Exit Sub
EH:
ErrHandlerPrivate True, "SQLToolkit", "Class_Initialize"
End Sub

'creates a DAO db object suitable for use by ExecTSQL.  Assumes that target is a
'SQL Server database.
'checks for server existence before execution.
Public Function OpenDBSvr(ByVal sDSN As String, _
                          ByVal sSvrNameSkipDSN As String, _
                          ByVal sLoginName As String, _
                          ByVal sPassword As String, _
                          Optional ByVal sDBName As String = "", _
                          Optional ByVal sOverrideAppName As String = "", _
                          Optional ByVal bResolveSvrToIP As Boolean = True) As dao.Database
On Error GoTo EH
Dim sConn As String

'create the connection string.  NOTE:  DAO will allow creating a database object without
'specifying the database name.  If skip the database name here, it helps with suppressing
'display of the ODBC/SQL connection dialogs.  That is, it appears best to specify the
'database when execute the TSQL using ExecTSQL below.
sConn = BuildODBCConnStr(sDSN, sSvrNameSkipDSN, sLoginName, sPassword, sDBName, sOverrideAppName)

'open the connection
Set OpenDBSvr = OpenDBSvr2(sConn, bResolveSvrToIP)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "OpenDBSvr"
End Function

'opens a database object using a completed ODBC connection string.
'does not check for server existence before execution.
Public Function OpenDBSvr2(ByVal sODBCConnStr As String, Optional ByVal bResolveSvrToIP As Boolean = True) As dao.Database
On Error GoTo EH
Dim dbSvr As dao.Database
Dim oSvrDMO As SQLDMO.SQLServer
Dim qdefTemp As dao.QueryDef
Dim rsTemp As dao.Recordset
Dim sDSN As String
Dim sServer As String
Dim sOrigSvrName As String
Dim sDriver As String
Dim sLoginName As String
Dim sPassword As String
Dim sDBName As String
Dim sApp As String
Dim sWSID As String
Dim sSvrAsIP As String
Dim sDriverForIP As String
Dim vServerList As Variant
Dim vDatabaseList As Variant
Dim bLogErrorsFlag_OLD As Boolean
Dim bDBExists As Boolean
Dim bLoginHasDBAccess As Boolean
Dim lErrNumber As Long
Dim sErrDescr As String
Dim ErrClone1 As ErrClone

m_oAppLog.AddEntry "<Procedure SQLToolkit.OpenDBSvr2:  Start connecting to server." & vbCrLf _
                 & " (Entry Conn= '" & HidePWDInConnStr(sODBCConnStr) & "')..." & vbCrLf _
                 & ">"

'save the original passed ODBC connection string (this will be modified)
m_sOpenDBSvr2_OrigODBCConnStr = sODBCConnStr

'extract the needed connection info
ExtractODBCParams sODBCConnStr, sDSN, sServer, sDriver, sLoginName, sPassword, sDBName, sApp, sWSID

'if DSN was used
If sDSN <> "" Then
    'if DSN exists on local PC
    If GetSystemDSNList(sDSN, , , , vServerList, vDatabaseList) Then
        'if override db name not provided in the connection string
        If sDBName = "" Then
            'use the one from the DSN
            sDBName = vDatabaseList(0)
        End If
        'if override server name not provided in the connection string
        If sServer = "" Then
            'use the one from the DSN
            sServer = vServerList(0)
        End If
    Else
        Err.Raise 1001, "SQLToolkit.OpenDBSvr2", _
                  "Procedure SQLToolkit.OpenDBSvr2 was passed a DSN" _
                 & " that does not exist or is not a 'System' DSN." & vbCrLf & vbCrLf _
                  & "[DSN = '" & sDSN & "']"
   End If
'if server name also not provided
ElseIf sServer = "" Then
    Err.Raise 1001, "SQLToolkit.OpenDBSvr2", _
        "Procedure SQLToolkit.OpenDBSvr2 has been passed an ODBC connection string without a DSN or server name." _
        & "  One of these must be specified."
End If

'if database name is still blank
If sDBName = "" Then
    'force to "master".  based on testing, bizarre JET connection caching issues occur if no db name is specified.
    sDBName = "master"
End If

'if flag is on
If bResolveSvrToIP Then
    'convert PC portion of server name to an IP address to avoid network timeouts. includes ping test to make
    'sure server PC can be reached.  sOrigSvrName is returned by reference
    sServer = ResolveSvrEngineToIP(sDSN, sServer, sOrigSvrName)
End If

'assemble a DSN-less connection string (this ensures that we use the same connection method whether or
'not the server name is resolved to an IP address.  it also verrides any bad config settings in the DSN)
sODBCConnStr = BuildODBCConnStr("", sServer, sLoginName, sPassword, sDBName, sApp)

'to prevent SQL driver dialog, make sure the server instance exists.  this uses DMO and produces a trappable error.
'oSvrDMO is returned by ref.
ServerIsReachable "", sServer, sLoginName, sPassword, bResolveSvrToIP, oSvrDMO, lErrNumber, sErrDescr
If lErrNumber <> 0 Then
     'raise error back to client
     Err.Raise lErrNumber, "SQLToolkit.OpenDBSvr2", sErrDescr
End If

'do screening tests on whether can access the requested database. (trap errors here to prevent ODBC connection
'dialog from appearing)
If Not ScreenDBForAccess(oSvrDMO, sDBName, bDBExists, bLoginHasDBAccess, lErrNumber, sErrDescr) Then
    If Not bDBExists Then
        Err.Raise 1001, "SQLToolkit.OpenDBSvr2", "The database named '" & sDBName & "' does not exist." & vbCrLf & vbCrLf _
              & "[DSN = '" & sDSN & "']" & vbCrLf _
              & "[Server = '" & sServer & "']"
    End If
    If Not bLoginHasDBAccess Then
        Err.Raise 1001, "SQLToolkit.OpenDBSvr2", "The login name '" & sLoginName & "' does not have access to the" _
            & " '" & sDBName & "' database." & vbCrLf & vbCrLf _
              & "[DSN = '" & sDSN & "']" & vbCrLf _
              & "[Server = '" & sServer & "']"
    End If
    'if make it to here, an unexpected error occurred
    Err.Raise 1001, "SQLToolkit.OpenDBSvr2", "An unexpected error occurred when attempting to connect to the" _
         & " database named '" & sDBName & "':" & vbCrLf _
         & "[Err.Number = " & lErrNumber & "]" & vbCrLf _
         & "[Err.Description = '" & sErrDescr & "']" & vbCrLf _
         & "[DSN = '" & sDSN & "']" & vbCrLf _
         & "[Server = '" & sServer & "']"
End If

'make sure DMO server reference is closed or this proc will generate 'orphaned' SQL Server connections
CloseDMOServerRef oSvrDMO

'make absolutely sure database is specified in the ODBC connection string.  otherwise, based on testing, JET may
'consider this connection to be a match for connections to other databases and cause very confusing errors.
sODBCConnStr = UpdateDBInODBCConnStr(sODBCConnStr, sDBName)

'clear out any old connections from the JET ODBC connection cache.  (this helps clear out old orphaned connections after network disconnect)
DBEngine.Idle

'open the connection
'Notes:  (1) dbDriverNoPrompt only works with ODBCDirect workspaces and will not stop
'            the SQL driver dialog from popping up if there are connection problems
'        (2) with older versions of DAO, there was a problem with the .LoginTimeout property not working
'            unless prefixed the OpenDatabase with DBEngine.  Based on testing with DAO 3.6/JET4.0, this bug has been
'            fixed and the DBEngine prefix is not necesssary.
Set dbSvr = OpenDatabase("", 0, 0, sODBCConnStr)

'-----------------
'V1.10:  A bug in the VB6 IDE may cause the following error the very first time a TSQL statement is executed or OpenRecordset
'is executed against a newly opened connection (see MSDN article #246167, old ref #Q246167).
'#3228:  Selected collating sequence not supported by the operating system."
'Based on documentation and testing, the error does not occur if run a compiled .exe.
'Following executes a trivial dummy call against the connection which forces the error and seems to prevent it from occurring again.
'Based on testing, the recordset must be executed against a table that actually exists in the database--sysusers works.
'If Microsoft ever "hides" the sysusers table, this code may need to be modified.
On Error Resume Next
dbSvr.OpenRecordset "SELECT name FROM sysusers WHERE 1=0", dbOpenForwardOnly
On Error GoTo EH
'-----------------
'due to past issues with JET connection caching, make absolutely sure the 'database context' of the connection matches
'the database we requested.  Don't use ExecTSQL to avoid error logging here.
Set qdefTemp = dbSvr.CreateQueryDef("")
qdefTemp.Connect = dbSvr.Connect
qdefTemp.ReturnsRecords = True
qdefTemp.SQL = "Exec sp_who @@spid"
Set rsTemp = qdefTemp.OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)
If (UCase(rsTemp!DBName) <> UCase(sDBName)) Then
    'raise error here (else caller will get very confusing errors when unsuccessfully tries to use the db ref)
    Err.Raise 1001, "SQLToolkit.OpenDBSvr2", "Procedure SQLToolkit.OpenDBSvr2 has detected a 'database context' that does" _
                & " not match the name of the requested database.  The database reference is thus invalid and cannot be used." & vbCrLf _
                & "[requested database name = '" & sDBName & "']" & vbCrLf _
                & "[actual database context = '" & rsTemp!DBName & "']"
End If
'-----------------

'return the database reference
Set OpenDBSvr2 = dbSvr

m_oAppLog.AddEntry "<Procedure SQLToolkit.OpenDBSvr2:  ...Done connecting to server." & vbCrLf _
                 & " (Final Conn='" & HidePWDInConnStr(sODBCConnStr) & "')" & vbCrLf _
                 & ">"

Exit Function
EH:
ErrClone1 = MakeErrClone
'make sure DMO server reference is closed or will get orphaned SQL Server connections
CloseDMOServerRef oSvrDMO
RestoreErrObject ErrClone1
'PWD portion of ODBC connection string is hidden prior to logging
ErrHandlerPrivate False, "SQLToolkit", "OpenDBSvr2", _
         "[ODBC Conn Str = '" & HidePWDInConnStr(sODBCConnStr) & "']" & vbCrLf _
         & "[Orig ODBC Conn Str = '" & HidePWDInConnStr(m_sOpenDBSvr2_OrigODBCConnStr) & "']"
End Function

'returns the name of the current database on the current connection
Public Function GetDBContext(dbSvr As dao.Database) As String
On Error GoTo EH
Dim rsTemp As dao.Recordset

Set rsTemp = ExecTSQL(dbSvr, "Exec sp_who @@spid", True, , True)

If Not (rsTemp Is Nothing) Then
   If Not rsTemp.EOF Then
       GetDBContext = GetFV(rsTemp, "DBName")
   End If
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetDBContext"
End Function

'takes server name of the form PCName\Instance and breaks up into two parts.
'or, if not using a named instance, returns the PCName.
'NOTE:  sSvrPCName_Rtn may be "(local)" unless set flag to resolve into local PC name
Public Sub ParseSvrName(sFullSvrName As String, Optional sSvrPCName_Rtn As String, _
                                                Optional sSvrInstanceName_Rtn As String, _
                                                Optional bResolveLOCALToPCName As Boolean = False)
On Error GoTo EH
Dim lPos As Long

'extract the PCName portion of the SQL Server instance name.
lPos = InStr(1, sFullSvrName, "\")

If lPos > 0 Then
    sSvrPCName_Rtn = Left(sFullSvrName, lPos - 1)
    sSvrInstanceName_Rtn = Mid(sFullSvrName, lPos + 1)
Else
    sSvrPCName_Rtn = sFullSvrName
    sSvrInstanceName_Rtn = ""
End If

If UCase(sSvrPCName_Rtn) = "(LOCAL)" And bResolveLOCALToPCName Then
    sSvrPCName_Rtn = GetLocalPCName
End If

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "ParseSvrName"
End Sub

'converts the PC name portion of the SQL server instance name to an IP address (if not the '(local)' instance).
'raises error if target machine cannot be successfully ping'ed
Public Function ResolveSvrEngineToIP(ByVal sDSN As String, _
                                     ByVal sSvrNameSkipDSN As String, _
                                     Optional ByRef sOrigSvrName_Rtn As String) As String
On Error GoTo EH
Dim sServer As String
Dim sSvrPCName As String
Dim sSvrInstanceName As String
Dim sIPAddr As String
Dim sPingRtnStatus As String

'set default return values
ResolveSvrEngineToIP = ""
sOrigSvrName_Rtn = ""

sServer = GetServerName(sDSN, sSvrNameSkipDSN)

'V1.06:  Return original server name to caller
sOrigSvrName_Rtn = sServer

'extract the PCName and instance name (if any) from the full server name
ParseSvrName sServer, sSvrPCName, sSvrInstanceName

'if reference to local machine
'V1.06:  added test for local PC name
If (UCase(sSvrPCName) = "(LOCAL)") Or (UCase(sSvrPCName) = UCase(GetLocalPCName)) Then
    'exit early without converting to IP (based on testing, no advantage to converting "(local)" or
    'local PC name to IP address
    ResolveSvrEngineToIP = sServer
    Exit Function
End If

'if this is not an IP address already
If Not IsValidIPFormat(sSvrPCName) Then
     'attempt to resolve into IP address (this can take 15-20 sec on a slow network)
     sIPAddr = GetIPFromHostName(sSvrPCName)
     If sIPAddr = "" Then
         Err.Raise 1001, "SQLToolkit.ResolveSvrEngineToIP", "Procedure SQLToolkit.ResolveSvrEngineToIP is unable" _
            & " to resolve the SQL Server PC name into an IP address.  Please check the following:" & vbCrLf _
            & " (1) DSN/Server points to current server machine name and instance name." & vbCrLf _
            & " (2) Target server machine has a compatible/reachable IP address." & vbCrLf _
            & " (3) Target server machine is turned on." & vbCrLf _
            & " (4) SQL Server Service is running on the target server machine." _
            & vbCrLf & vbCrLf _
            & "[DSN = '" & sDSN & "']" & vbCrLf _
            & "[Server name = '" & sServer & "']" & vbCrLf _
            & "[PCName portion of server name = '" & sSvrPCName & "']"
     End If
    'assemble server name using an IP address for the server PC and return to caller
    ResolveSvrEngineToIP = sIPAddr & IIf(sSvrInstanceName = "", "", "\" & sSvrInstanceName)
Else
    sIPAddr = sSvrPCName
    'okay to return the original server name as this already is an IP
    ResolveSvrEngineToIP = sOrigSvrName_Rtn
End If

'attempt to PING this address
If Not PingIP(sIPAddr, , sPingRtnStatus) Then
    Err.Raise 1001, "SQLToolkit.ResolveSvrEngineToIP", "Procedure SQLToolkit.ResolveSvrEngineToIP is unable" _
       & " to ping the IP address of" _
       & " the target SQL Server PC." & vbCrLf & vbCrLf _
       & "Please check the following:" & vbCrLf & vbCrLf _
       & "(1) the target server PC is turned on" & vbCrLf _
       & "(2) target PC has a compatible IP address" & vbCrLf _
       & "(3) network communication is functioning" & vbCrLf _
       & "(4) target SQL Server instance is named correctly" & vbCrLf & vbCrLf _
       & "[DSN = '" & sDSN & "']" & vbCrLf _
       & "[SQL Server instance name = '" & sServer & "']" & vbCrLf _
       & "[Server PC name = '" & sSvrPCName & "']" & vbCrLf _
       & "[IP address of server PC = '" & sIPAddr & "']" & vbCrLf _
       & "[Ping return status = '" & sPingRtnStatus & "']"
End If  'else, ping successful => server exists

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "ResolveSvrEngineToIP"
End Function

Public Function IsValidIPFormat(sIPAddr_Test As String) As Boolean
On Error GoTo EH
Dim vItems As Variant
Dim vItem As Variant

'set default
IsValidIPFormat = False

'break up at the periods
vItems = Split(sIPAddr_Test, ".")
'if don't get exactly 3 items
If UBound(vItems) <> 3 Then
    Exit Function
End If
For Each vItem In vItems
    'if find non-numeric value
    If Not IsNumeric(vItem) Then
        Exit Function
    End If
Next vItem

'if make it to here, return success
IsValidIPFormat = True

Exit Function
EH:
'PWD portion of ODBC connection string is hidden prior to logging
ErrHandlerPrivate False, "SQLToolkit", "IsValidIPFormat"
End Function

'Function executes TSQL as pass-through query.  sTQSL must be Transact SQL syntax (not JET syntax).
'sDBName is only optional if a database was specified when calling OpendbSvr AND this database is the
'target for the TSQL.
'If bIgnoreNoRowsError is true, if error 3325 occurs (pass through query with returns records true returned no rows)
'this error is ignored and Nothing is returned.
Public Function ExecTSQL(dbSvr As dao.Database, _
                         ByVal sTSQL As String, _
                         ByVal bReturnsRecords As Boolean, _
                         Optional ByVal sDBName As String = "", _
                         Optional ByVal bIgnoreNoRecsError As Boolean) As dao.Recordset
On Error GoTo EH
Dim qd As dao.QueryDef
Dim sConn As String
Dim lPosDB As Long
Dim lPosEnd As Long

'set default returns
Set ExecTSQL = Nothing

'get the connection string from the database object
sConn = dbSvr.Connect

'if provided with a DB name
If sDBName <> "" Then
   'make sure the connection string includes this db name
   sConn = UpdateDBInODBCConnStr(sConn, sDBName)
Else
    'make sure a database was specified in the original conn string (case insensitive compare)
    lPosDB = InStr(1, sConn, "DATABASE=", vbTextCompare)
    If lPosDB > 0 Then
        'find the first semicolon after this
        lPosEnd = InStr(lPosDB, sConn, ";", vbTextCompare)
        'if no semicolon
        If lPosEnd < 1 Then
            'use end of string plus 1
            lPosEnd = Len(sConn) + 1
        End If
    End If
    'if no database specified
    If (lPosDB <= 0) Or ((lPosEnd - lPosDB) <= Len("DATABASE=")) Then
        'raise an error--caller should specify a database (if none apply, then can always use 'master')
        Err.Raise 1001, App.EXEName & ".SQLToolkit.ExecTSQL", _
                   "Procedure SQLToolkit.ExecTSQL cannot execute.  No database name is specified."
    End If
End If

Set qd = dbSvr.CreateQueryDef("")
'use a connect string on the query def that *does* include the database
qd.Connect = sConn  'specifying connection string tells JET this is pass through SQL
qd.ReturnsRecords = bReturnsRecords  'tell JET whether SQL returns records or not (if choose wrong, will get errors)
qd.SQL = sTSQL             'set the .SQL property to the Transact SQL statement
If bReturnsRecords Then
    Set ExecTSQL = qd.OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)   'Neither flag needs to be specified--these are the
                                     'defaults for SQL Server.  Based on testing, results are always returned as
                                     'read-only (so, dbSeeChanges is not necessary).
Else
    qd.Execute dbSQLPassThrough         'execute the TSQL.  Based on testing, do not need dbFailOnError to see errors.
                                        'the dbPassThrough is not essential as this is set automatically if the .Connect
                                        'property is set.
End If

Exit Function
EH:

'if need to suppress "no rows returned" error
If (Err.Number = 3325) And bIgnoreNoRecsError Then
    Err.Clear
    Exit Function
End If

'NOTE:  following hides the PWD= arg of the ODBC conn string and suppresses logging of the TSQL if it
'contains the word 'password'.  intent is to prevent logging of password info.  may need to add
'other tests in the future.
ErrHandlerPrivate False, "SQLToolkit", "ExecTSQL", _
       "[ODBC Conn Str = '" & HidePWDInConnStr(sConn) & "']" & vbCrLf _
       & "[Orig ODBC Conn Str = '" & HidePWDInConnStr(m_sOpenDBSvr2_OrigODBCConnStr) & "']" _
       & IIf(InStr(1, sTSQL, "password", vbTextCompare) > 0, "", vbCrLf & "[TSQL = '" & sTSQL & "']")
End Function

'NOTE:  This procedure adds the App.EXEName and local PCName to help with troubleshooting.
Public Function BuildODBCConnStr(ByVal sDSN As String, _
                                 ByVal sSvrNameSkipDSN As String, _
                                 ByVal sLoginName As String, _
                                 ByVal sPassword As String, _
                                 Optional ByVal sDBName As String = "", _
                                 Optional ByVal sOverrideAppName As String = "") As String
On Error GoTo EH
Dim sConn As String

'check if blank DSN and server name (otherwise, ODBC driver will bring up a dialog box even if
'specify dbDriverNoPrompt--which only works if use ODBCDirect)
If (sDSN = "") And (sSvrNameSkipDSN = "") Then
   Err.Raise 1001, App.EXEName & ".SQLToolkit." & "ExecTSQLRecUsingDSN", _
              "Function SQLToolkit.BuildODBCConnStr was passed an empty DSN and empty optional server name." _
            & "  One of these must be specified."
End If

'if server name is provided.
'("ODBC" prefix is required by JET or connection string is not recognized as an ODBC connection.)
If sSvrNameSkipDSN <> "" Then
    sConn = "ODBC;DRIVER={SQL Server};" _
          & "SERVER=" & sSvrNameSkipDSN & ";"
Else
    sConn = "ODBC;DSN=" & sDSN & ";"
End If

'Build rest of conn string.  If database is not specified, leave it off (can be useful for trapping ODBC errors without getting user ODBC dialog)
'Include the name of the application and PC to help with SQL logging
sConn = sConn _
      & "UID=" & sLoginName & ";" _
      & "PWD=" & sPassword & ";" _
      & IIf(sDBName = "", "", "DATABASE=" & sDBName & ";") _
      & "APP=" & IIf(sOverrideAppName = "", App.EXEName, sOverrideAppName) & ";" _
      & "WSID=" & GetLocalPCName & ";"
    
BuildODBCConnStr = sConn
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "BuildODBCConnStr"
End Function

'this updates the "DATABASE=" portion of the connection string
'(or adds one if not present).  if sDBName is blank, will remove the DATABASE= portion
Public Function UpdateDBInODBCConnStr(ByVal sODBCConnStr As String, ByVal sDBName As String) As String
On Error GoTo EH
Dim lPosStart As Long
Dim lPosEnd As Long
Dim sLeft As String
Dim sRight As String
Dim sConnNew As String

'if the database connection string included a DATABASE= param.  use case insensitive search
lPosStart = InStr(1, sODBCConnStr, "DATABASE=", vbTextCompare)
If lPosStart > 0 Then
    'find the first semicolon after this
    lPosEnd = InStr(lPosStart, sODBCConnStr, ";", vbTextCompare)
    'if no semicolon
    If lPosEnd < 1 Then
        'read to end of string
        lPosEnd = Len(sODBCConnStr)
    End If
    'if no chars to the left of DATABASE
    If lPosStart < 2 Then
        'the DATABASE was the first param so nothing to the left
        sLeft = ""
    Else
        'else, get the preceeding chars
        sLeft = Left(sODBCConnStr, lPosStart - 1)
    End If
    'skip over any existing semicolon--we'll insert our own below
    sRight = Mid(sODBCConnStr, lPosEnd + 1)
Else
    'set up for appending the database name
    sLeft = sODBCConnStr
    'code below assumes that any existing string ends in a semicolon so add one now if necessary
    If Right(sODBCConnStr, 1) <> ";" Then
        sLeft = sLeft & ";"
    End If
    sRight = ""
End If

'if have a database name
If sDBName <> "" Then
    'add it to the conn string
    sConnNew = sLeft & "DATABASE=" & sDBName & ";" & sRight
Else
    'remove the DATABASE=
    sConnNew = sLeft & sRight
End If

UpdateDBInODBCConnStr = sConnNew

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "UpdateDBInODBCConnStr"
End Function

'generates backup file in standard WEM format:
'<PCName>_<EngineName>_<DBName>_YYYYMMDD_HHNNSS.sqb
'
'PCNames of the form W12345FSERVER are abbreviated as 12345FS
'PCNames of the form W12345ClientXX are abbreviated as 12345CXX
'WEMSQLEngine is abbreviated as WE
'WEMSQL is abbreviated as WSQ
'WEMSQL_IO is abbreviated as WIO

Public Function CreateStdBackupFNameExt(ByVal sDSN As String, _
                                        ByVal sSvrNameSkipDSN As String, _
                                        Optional ByVal sDBName As String = "", _
                                        Optional ByVal dtDateTime As Date, _
                                        Optional ByVal lBackupStatus As BackupStatus_Enum = BACKUP_STATUS_Inconsistent, _
                                        Optional ByVal sDefaultDBCmptLevel As String) _
                                        As String
On Error GoTo EH
Dim vServerList As Variant
Dim vDatabaseList As Variant
Dim sSvrName As String
Dim sDBName_DSN As String
Dim sSvrPCName As String
Dim sSvrInstanceName As String
Dim sPCName_Out As String
Dim sSvrInstance_Out As String
Dim sDBName_Out As String
Dim sBackupStatus_Out As String
Dim sDefaultDBCmptLevel_Out As String

'if override server name not provided
If sSvrNameSkipDSN = "" Then
    'if DSN provided
    If sDSN <> "" Then
        'look this up from the provided DSN
        If GetSystemDSNList(sDSN, , , , vServerList, vDatabaseList) Then
            sSvrName = vServerList(0)
            sDBName_DSN = vDatabaseList(0)
        End If
    End If  'else, will default to blanks below
    'if override database name not provided
    If sDBName = "" Then
        'use the default one from the DSN
        sDBName = sDBName_DSN
    End If
Else
    sSvrName = sSvrNameSkipDSN
End If

'get server PC name and instance name.  Resolve "(local)" to PC name
ParseSvrName sSvrName, sSvrPCName, sSvrInstanceName, True

'if SvrPCName has recognized format
Select Case True
    Case UCase(sSvrPCName) Like UCase("W#####FServer")
         'use std abbreviated format
         sPCName_Out = Mid(sSvrPCName, 2, 5) & "FS"
    Case UCase(sSvrPCName) Like UCase("W#####Client*")
         sPCName_Out = Mid(sSvrPCName, 2, 5) & "C" & Mid(sSvrPCName, 13)
    Case Else
         sPCName_Out = sSvrPCName
End Select

'if SvrInstance name has recognized format
If UCase(sSvrInstanceName) = "WEMSQLENGINE" Then
    'use abbreviation
    sSvrInstance_Out = "WE"
Else
    sSvrInstance_Out = sSvrInstanceName
End If

'if database name has recognized format, use 2 letter abbreviation
Select Case UCase(sDBName)
    Case "WEMSQL"
        sDBName_Out = "SQ"
    Case "WEMSQL_IO"
        sDBName_Out = "IO"
    Case Else
        sDBName_Out = sDBName
End Select

'if date not provided
If dtDateTime = CDate(0) Then
    'default to current date/time
    dtDateTime = Now
End If

'translate the backup status flag
Select Case lBackupStatus
     Case BACKUP_STATUS_Inconsistent
        sBackupStatus_Out = "i"
     Case BACKUP_STATUS_Consistent
        sBackupStatus_Out = "c"
     Case Else
        'default to inconsistent
        sBackupStatus_Out = "i"
End Select

'convert compatibility level from "80" or "90" to "08" or "09".
'if not provided, will default to "00"
'(this is in anticipation of level "100" for next version)

sDefaultDBCmptLevel = Right("000" & sDefaultDBCmptLevel, 3)
sDefaultDBCmptLevel_Out = Mid(sDefaultDBCmptLevel, Len(sDefaultDBCmptLevel) - 2, 2)

'to reduce size when displayed with proportional fonts, convert all to lower case.
sPCName_Out = StrConv(sPCName_Out, vbLowerCase)
sSvrInstance_Out = StrConv(sSvrInstance_Out, vbLowerCase)
sDBName_Out = StrConv(sDBName_Out, vbLowerCase)
sDefaultDBCmptLevel_Out = StrConv(sDefaultDBCmptLevel_Out, vbLowerCase)

' to help with parsing file name in future, convert any existing underscores to
'double underscores. (This will fail if server or database name has
'leading underscore(s), but these should be rare.
'Embedded underscores should still allow parsing.)
sPCName_Out = Replace(sPCName_Out, "_", "__", , , vbTextCompare)
sSvrInstance_Out = Replace(sSvrInstance_Out, "_", "__", , , vbTextCompare)
sDBName_Out = Replace(sDBName_Out, "_", "__", , , vbTextCompare)
sDefaultDBCmptLevel_Out = Replace(sDefaultDBCmptLevel_Out, "_", "__", , , vbTextCompare)

'build output file name with extension
CreateStdBackupFNameExt = sPCName_Out _
    & "_" & IIf(sSvrInstance_Out = "", "", sSvrInstance_Out & "_") _
    & sDBName_Out _
    & "_" & Format(dtDateTime, "yyyymmdd") _
    & "_" & Format(dtDateTime, "hhnnss") _
    & sBackupStatus_Out _
    & sDefaultDBCmptLevel_Out _
    & ".sqb"

Exit Function
EH:
ErrHandlerPrivate True, "SQLToolkit", "CreateStdBackupFNameExt"
End Function

'returns False if unable to recognize the format.
Public Function ParseStdBackupFNameExt(ByVal sStdBkupFNameExt As String, _
                                       Optional sSvrPCName_Rtn As String, _
                                       Optional sSvrInstanceName_Rtn As String, _
                                       Optional sDBName_Rtn As String, _
                                       Optional dtDateTime_Rtn As String, _
                                       Optional lBackupStatus_Rtn As BackupStatus_Enum, _
                                       Optional sDefaultDBCmptLevel_Rtn) As Boolean
On Error GoTo EH
Dim sStdBkupFName As String
Dim sSvrPCName_Temp As String
Dim sSvrInstanceName_Temp As String
Dim sDBName_Temp As String
Dim sDate_Temp As String
Dim sTime_Temp As String
Dim sDateTime_Temp As String
Dim dtDateTime_Temp As Date
Dim sBackupStatus_Temp As String
Dim lBackupStatus_Temp As BackupStatus_Enum
Dim sDefaultDBCmptLevel_Temp As String
Dim lPos1 As Long
Dim lPos2 As Long
Dim lPos3 As Long
Dim lPos_DTimeStart As Long
Dim sPart2 As String

'set defaults => couldn't parse
ParseStdBackupFNameExt = False
sSvrPCName_Rtn = ""
sSvrInstanceName_Rtn = ""
sDBName_Rtn = ""
dtDateTime_Rtn = CDate(0)
lBackupStatus_Rtn = BACKUP_STATUS_Inconsistent
sDefaultDBCmptLevel_Rtn = ""

'typical backup fname will look similar to:
'1234FS_we_sq_20051207_103703c09.sqb OR
'1234C3_we_sq_20051207_103703c09.sqb

'the _we_ portion may not be present if the (local) server instance is used. So, these are also valid:
'1234FS_sq_20051207_103703c09.sqb OR
'1234C3_sq_20051207_103703c09.sqb
      
'if doesn't have right format (allow extra trailing chars for upward compatibility)
If Not (UCase(sStdBkupFNameExt) Like UCase("*_########_######[ci]##*.sqb")) Then
    Exit Function
End If

'remove extension portion
sStdBkupFName = m_oFSO.GetBaseName(sStdBkupFNameExt)

'convert double underscores to custom char (double underscore is marker for true embedded underscores)
sStdBkupFName = Replace(sStdBkupFName, "__", Chr$(1), , , vbTextCompare)

'get first underscore
lPos1 = InStr(1, sStdBkupFName, "_")
If lPos1 < 2 Then
    Exit Function
End If
sSvrPCName_Temp = Mid(sStdBkupFName, 1, lPos1 - 1)

'get second underscore
lPos2 = InStr(lPos1 + 1, sStdBkupFName, "_")
If lPos2 < (lPos1 + 2) Then
    Exit Function
End If

'get what may be instance name or db name
sPart2 = Mid(sStdBkupFName, lPos1 + 1, lPos2 - lPos1 - 1)

'if portion beyond second underscore looks like the end
If Mid(sStdBkupFName, lPos2 + 1) Like "########_######[ci]##*" Then
    'then no server instance name
    sSvrInstanceName_Temp = ""
    'the second part is the database name
    sDBName_Temp = sPart2
    'date/time start after the 2nd underscore
    lPos_DTimeStart = lPos2 + 1
Else
    'second part is the instance name
    sSvrInstanceName_Temp = sPart2
    
    'get third underscore
    lPos3 = InStr(lPos2 + 1, sStdBkupFName, "_")
    If lPos3 < (lPos2 + 2) Then
        Exit Function
    End If
    
    'if portion beyond the third underscore doesn't look like the end
    If Not (Mid(sStdBkupFName, lPos3 + 1) Like "########_######[ci]##*") Then
        Exit Function
    End If
    
    'get the database name
    sDBName_Temp = Mid(sStdBkupFName, lPos2 + 1, lPos3 - lPos2 - 1)
    
    'date/time starts after the 3rd underscore
    lPos_DTimeStart = lPos3 + 1
End If

sDate_Temp = Mid(sStdBkupFName, lPos_DTimeStart, 8)
sTime_Temp = Mid(sStdBkupFName, lPos_DTimeStart + Len(sDate_Temp) + 1, 6)
sBackupStatus_Temp = Mid(sStdBkupFName, lPos_DTimeStart + Len(sDate_Temp) + 1 + Len(sTime_Temp), 1)
sDefaultDBCmptLevel_Temp = Mid(sStdBkupFName, _
           lPos_DTimeStart + Len(sDate_Temp) + 1 + Len(sTime_Temp) + Len(sBackupStatus_Temp), 2)

'process the backup status code
Select Case sBackupStatus_Temp
    Case "i"
       lBackupStatus_Temp = BACKUP_STATUS_Inconsistent
    Case "c"
       lBackupStatus_Temp = BACKUP_STATUS_Consistent
    Case Else
       'return with false--can't parse properly (shouldn't get here)
       Exit Function
End Select

'process the date and time
sTime_Temp = Mid(sTime_Temp, 1, 2) & ":" & Mid(sTime_Temp, 3, 2) & ":" & Mid(sTime_Temp, 5, 2)
sDate_Temp = Mid(sDate_Temp, 5, 2) & "/" & Mid(sDate_Temp, 7, 2) & "/" & Mid(sDate_Temp, 1, 4)
sDateTime_Temp = sDate_Temp & " " & sTime_Temp
If Not IsDate(sDateTime_Temp) Then
    Exit Function
End If
dtDateTime_Temp = CDate(sDateTime_Temp)

'for items that could conceivable include an underscore, convert the placeholder Chr(1) back to an underscore.
sSvrPCName_Temp = Replace(sSvrPCName_Temp, Chr(1), "_")
sSvrInstanceName_Temp = Replace(sSvrInstanceName_Temp, Chr(1), "_")
sDBName_Temp = Replace(sDBName_Temp, Chr(1), "_")
sDefaultDBCmptLevel_Temp = Replace(sDefaultDBCmptLevel_Temp, Chr(1), "_")

'translate PC name
Select Case True
    'if looks like a std abbreviated FSERVER PC name
    Case UCase(sSvrPCName_Temp) Like "#####FS"
        'convert back to full form
        sSvrPCName_Temp = "W" & Left(sSvrPCName_Temp, 5) & "FSERVER"
    'if looks like std abbreviated Client PC name
    Case UCase(sSvrPCName_Temp) Like "#####C#*"
        'convert back to full form
        sSvrPCName_Temp = "W" & Left(sSvrPCName_Temp, 5) & "CLIENT" & Mid(sSvrPCName_Temp, 7)
    Case Else
        'else, assume no translation
End Select

'process the instance name (if any)
Select Case UCase(sSvrInstanceName_Temp)
   Case "WE"
       sSvrInstanceName_Temp = "WEMSQLENGINE"
   Case Else
       'assume no translation
End Select

'process the database name
Select Case UCase(sDBName_Temp)
    Case "SQ"
        sDBName_Temp = "WEMSQL"
    Case "IO"
        sDBName_Temp = "WEMSQL_IO"
    Case Else
End Select

'return everything to the caller.
sSvrPCName_Rtn = sSvrPCName_Temp
sSvrInstanceName_Rtn = sSvrInstanceName_Temp
sDBName_Rtn = sDBName_Temp
dtDateTime_Rtn = dtDateTime_Temp
lBackupStatus_Rtn = lBackupStatus_Temp
sDefaultDBCmptLevel_Rtn = sDefaultDBCmptLevel_Temp

'return success
ParseStdBackupFNameExt = True

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "ParseStdBackupFNameExt"
End Function

'Version of the above with standard WEM assumptions regarding file names, folder paths, etc:

'1.  Creates backup file to to folder sDestFolder_RelApp.  If sCustomFileName is blank(default), then
'    standard "long" file name is generated (which embeds time stamp, database name, server name,
'    and other info)
'2.  Final full path to backup file is returned (relative to local running application).
'3.  If sDBName is not provided, the database context of the current connection is backed up
'    (which is usually the desired behavior)
'4.  If server is remote, the server PC must have the F: drive shared. (By default, as "WEMDATA".  Can
'    use sBaseFolder_LocalToSvr and sBaseFolder_ShareName to specify a different share name.)
'5.  If target file already exists the existing file is deleted.  It's up to caller to warn user that an
'    overwrite will occur.
Public Function BackupDBToFile2(dbSvr As dao.Database, _
                                Optional ByVal sDBName As String = "", _
                                Optional ByVal sBkupPassword As String = "", _
                                Optional ByVal sDestFolder_RelApp As String = "", _
                                Optional ByVal lBackupStatus As BackupStatus_Enum = BACKUP_STATUS_Inconsistent, _
                                Optional ByVal sBkupDescr As String = "", _
                                Optional ByVal sBaseFolder_LocalToSvr As String = "F:\", _
                                Optional ByVal sBaseFolder_ShareName As String = "WEMDATA", _
                                Optional ByVal sCustomFileName As String = "") _
                                As String
On Error GoTo EH
Dim sServerName As String
Dim sSvrPCName As String
Dim sInstanceName As String
Dim sDefaultDBCmptLevel As String
Dim sBackupFileName As String
Dim sTempFolder_RelApp As String
Dim sBackupFNameOnly As String
Dim sBackupExtOnly As String
Dim sCount As String
Dim sTempFileName As String
Dim sTempFilePath_RelApp As String
Dim sTempFilePath_RelSvr As String
Dim sCustomTempFileName As String
Dim sCustomTempFilePath_RelApp As String
Dim sCustomFilePath_RelApp As String
Dim sFilePath_RelApp As String
Dim lErrNumber As Long
Dim sErrDescription As String

'make sure base folders end in backslashes (else, issues with x: versus x:\ when build paths)
If sBaseFolder_LocalToSvr <> "" Then
    sBaseFolder_LocalToSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "\")
End If
'g_oAppLog.AddEntry "TAG1"
If sBaseFolder_ShareName <> "" Then
    sBaseFolder_ShareName = m_oFSO.BuildPath(sBaseFolder_ShareName, "\")
End If
'g_oAppLog.AddEntry "TAG2"

'if custom target folder specified
If sDestFolder_RelApp <> "" Then
    'make sure ends in backslash (else, can have x: versus x:\ issues)
    sDestFolder_RelApp = m_oFSO.BuildPath(sDestFolder_RelApp, "\")
    'make sure output folder exists (do this early so less likely to strand .tmp files)
    If Not CreateFolderIfNec(sDestFolder_RelApp, lErrNumber, sErrDescription) Then
         Err.Raise lErrNumber, "SQLToolkit.BackupDBToFile2", "An error occurred when creating a destination folder" _
           & " for a SQL Server database backup.  (File paths are relative to current local application.)" & vbCrLf & vbCrLf _
           & sErrDescription
    End If
End If
'g_oAppLog.AddEntry "TAG3"

'if no database provided
If sDBName = "" Then
    'get the database context of the connection
    sDBName = GetDBContext(dbSvr)
End If
'g_oAppLog.AddEntry "TAG4"

'get server details
GetServerProperties dbSvr, sServerName, sSvrPCName, sInstanceName, , , sDefaultDBCmptLevel
'g_oAppLog.AddEntry "TAG5"

'test if we need (and have) access to shared folders on the server PC (error is raised if there's a problem)
ValidateSvrNWShare dbSvr, sSvrPCName, sBaseFolder_LocalToSvr, sBaseFolder_ShareName, sDestFolder_RelApp
'g_oAppLog.AddEntry "TAG6"

'if no custom file name provided
If sCustomFileName = "" Then
    'generate a name for the output file.
    sBackupFileName = CreateStdBackupFNameExt("", sServerName, sDBName, Now, lBackupStatus, sDefaultDBCmptLevel)
Else 'use the custom name
    sBackupFileName = sCustomFileName
End If
'g_oAppLog.AddEntry "TAG7"
    
'if server is local to current app
If UCase(sSvrPCName) = UCase(GetLocalPCName) Then
    'use local \temp folder
    sTempFolder_RelApp = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\Temp")
Else
    'use UNC \temp folder (use standard shared folder name on the server--validated above)
    sTempFolder_RelApp = m_oFSO.BuildPath("\\" & sSvrPCName, sBaseFolder_ShareName)
    sTempFolder_RelApp = m_oFSO.BuildPath(sTempFolder_RelApp, "WEM\Data\Temp")
End If
'g_oAppLog.AddEntry "TAG8"

'create \Temp folder if it doesn't exist
If Not CreateFolderIfNec(sTempFolder_RelApp, lErrNumber, sErrDescription) Then
    Err.Raise lErrNumber, "SQLToolkit.BackupDBToFile2", "An error occurred when creating the '\temp' folder" _
       & " for a SQL Server database backup.  (File paths are relative to current local application.)" & vbCrLf & vbCrLf _
       & sErrDescription
End If
'g_oAppLog.AddEntry "TAG9"

'create unique temp file name
sBackupFNameOnly = m_oFSO.GetBaseName(sBackupFileName)
sBackupExtOnly = m_oFSO.GetExtensionName(sBackupFileName)
sCount = "0"
'g_oAppLog.AddEntry "TAG10"
Do
    sCount = Val(sCount) + 1
    sTempFileName = sBackupFNameOnly & "#" & sCount & "." & sBackupExtOnly & ".tmp"
    sTempFilePath_RelApp = m_oFSO.BuildPath(sTempFolder_RelApp, sTempFileName)
    sTempFilePath_RelSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\Temp") & "\" & sTempFileName
    'continue while file exists
Loop While m_oFSO.FileExists(sTempFilePath_RelApp)
'g_oAppLog.AddEntry "TAG11"

'create the backup to temp file name.

BackupDBToFile dbSvr, sDBName, sTempFilePath_RelSvr, sBkupPassword, sBkupDescr
'g_oAppLog.AddEntry "TAG12"

'if custom target folder specified
If sDestFolder_RelApp <> "" Then
'g_oAppLog.AddEntry "TAG12.1"
    'make sure it ends in a backslash
    sDestFolder_RelApp = m_oFSO.BuildPath(Trim(sDestFolder_RelApp), "\")
'g_oAppLog.AddEntry "TAG12.2"
    '(know folder exists from validation above.)
    
    'if temp file path relative to app is local
    If Left(sTempFilePath_RelApp, 2) <> "\\" Then
        'if possible, convert dest folder to local path (this is to help reduce situations where Name()
        'will be passed a network path and a local path to the same local physical drive.  in these cases, Name()
        'will execute a file copy--which can take a long time.  however, if convert both to local paths or both to network
        'paths, Name() will just update the drive's FAT with the new entry--which is very fast)
        sDestFolder_RelApp = GetLocalPath_RelPC(GetLocalPCName, sBaseFolder_LocalToSvr, sBaseFolder_ShareName, sDestFolder_RelApp)
    End If
'g_oAppLog.AddEntry "TAG12.3"
    'build temp output file path.  make sure that target file name does not already exist
    '(if does, most likely due to custom folder being same as the \Temp folder).
    'start with same sCount from above as duplicates in the final folder are unlikely.
    Do
        sCustomTempFileName = sBackupFNameOnly & "#" & sCount & "." & sBackupExtOnly & ".tmp"
        sCustomTempFilePath_RelApp = m_oFSO.BuildPath(sDestFolder_RelApp, sCustomTempFileName)
        sCount = Val(sCount) + 1
        'continue while file exists
    Loop While m_oFSO.FileExists(sCustomTempFilePath_RelApp)
'g_oAppLog.AddEntry "TAG12.4"
    If InStr(1, sTempFilePath_RelApp, "\\", vbTextCompare) > 0 Then Sleep 5000 '1.11
    'move file to final folder (this may actually be a copy which could take some time.  if on same
    'physical hard drive and both are local paths or both are network paths, no copy is involved)
    
    'Name sTempFilePath_RelApp As sCustomTempFilePath_RelApp
    m_oFSO.MoveFile sTempFilePath_RelApp, sCustomTempFilePath_RelApp
    'm_oFSO.CopyFile sTempFilePath_RelApp, sCustomTempFilePath_RelApp
    
    'build final output file path
'g_oAppLog.AddEntry "TAG12.5"
    sCustomFilePath_RelApp = m_oFSO.BuildPath(sDestFolder_RelApp, sBackupFileName)
    'make sure it doesn't already exist
'g_oAppLog.AddEntry "TAG12.6"
    DeleteFile_NoError sCustomFilePath_RelApp
    'rename temp file to final name
'g_oAppLog.AddEntry "TAG12.7"
    Name sCustomTempFilePath_RelApp As sCustomFilePath_RelApp
    'compress the .sqb backup into a .zip file (same base name), then remove the original .sqb
    BackupDBToFile2 = CompressBackupFileToZip(sCustomFilePath_RelApp)
'g_oAppLog.AddEntry "TAG12.8"
Else
'g_oAppLog.AddEntry "TAG12.9"
    'create final path within the \Temp folder
    sFilePath_RelApp = m_oFSO.BuildPath(sTempFolder_RelApp, sBackupFileName)
'g_oAppLog.AddEntry "TAG12.10"
    'make sure target file name doesn't already exist
    DeleteFile_NoError sFilePath_RelApp
'g_oAppLog.AddEntry "TAG12.11"
    'rename temp file to final name
    Name sTempFilePath_RelApp As sFilePath_RelApp
'g_oAppLog.AddEntry "TAG12.12"
    'compress the .sqb backup into a .zip file (same base name), then remove the original .sqb
    BackupDBToFile2 = CompressBackupFileToZip(sFilePath_RelApp)
'g_oAppLog.AddEntry "TAG12.13"
End If
'g_oAppLog.AddEntry "TAG13"

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "BackupDBToFile2"
End Function

'Compresses sSourceFilePath_RelApp (typically a ".sqb" backup file) into a new ".zip" file with the
'same base name, in the same folder.  Uses PowerShell's Compress-Archive, run synchronously and
'hidden via WScript.Shell, so this works reliably from a Windows service or scheduled task even when
'there is no interactive desktop session (unlike the Shell.Application/Explorer "send to zip" trick,
'which requires an interactive shell and can hang or silently fail when run unattended).
'If compression succeeds, the original file is deleted and the path to the new .zip file is
'returned.  If compression fails for any reason, the original file is left alone (never deleted)
'and its path is returned unchanged - a compression failure must never prevent the backup file
'from existing.
Private Function CompressBackupFileToZip(ByVal sSourceFilePath_RelApp As String) As String
On Error GoTo EH
Dim sZipFilePath_RelApp As String
Dim oShell As Object
Dim oExec As Object
Dim sCmd As String
Dim lExitCode As Long
Dim sOutput As String
Dim lWaitCount As Long
Dim lLastSize As Currency
Dim lCurSize As Currency
Dim bReady As Boolean

'default return value is the original (uncompressed) path, in case compression fails
CompressBackupFileToZip = sSourceFilePath_RelApp

'the source .sqb may have just been created/renamed over a network share - wait until the file
'both (a) exists and (b) has a stable size across two checks, so we never hand PowerShell a path
'that isn't fully visible/settled yet (safety timeout ~30 seconds)
lLastSize = -1
bReady = False
For lWaitCount = 1 To 60
    If m_oFSO.FileExists(sSourceFilePath_RelApp) Then
        lCurSize = m_oFSO.GetFile(sSourceFilePath_RelApp).Size
        If lCurSize = lLastSize Then
            bReady = True
            Exit For
        End If
        lLastSize = lCurSize
    Else
        lLastSize = -1
    End If
    Sleep 500
Next lWaitCount

'if, after waiting, the source file still doesn't actually exist, there's nothing to compress
If Not m_oFSO.FileExists(sSourceFilePath_RelApp) Then
    m_oAppLog.AddEntry "CompressBackupFileToZip:  source file '" & sSourceFilePath_RelApp & "' does not exist (even after waiting) - nothing to compress"
    Exit Function
End If

If Not bReady Then
    m_oAppLog.AddEntry "CompressBackupFileToZip:  source file '" & sSourceFilePath_RelApp & "' never reached a stable size - attempting compression anyway"
End If

'build the target .zip path (same folder/base name as the source file)
sZipFilePath_RelApp = m_oFSO.BuildPath(m_oFSO.GetParentFolderName(sSourceFilePath_RelApp), _
                                       m_oFSO.GetBaseName(sSourceFilePath_RelApp) & ".zip")

'make sure it doesn't already exist
DeleteFile_NoError sZipFilePath_RelApp

'write a temp .ps1 script file and run it - this avoids command-line quoting issues and lets us
'use LoadWithPartialName (works on all PowerShell versions including v2, unlike Add-Type -AssemblyName)
Dim sScriptPath As String
Dim iScript As Integer
sScriptPath = m_oFSO.BuildPath(Environ$("TEMP"), "zip_backup.ps1")
iScript = FreeFile
Open sScriptPath For Output As #iScript
Print #iScript, "[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem') | Out-Null"
Print #iScript, "[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression') | Out-Null"
Print #iScript, "$src = '" & Replace(sSourceFilePath_RelApp, "'", "''") & "'"
Print #iScript, "$dst = '" & Replace(sZipFilePath_RelApp, "'", "''") & "'"
Print #iScript, "if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }"
Print #iScript, "$fs = [System.IO.File]::Open($dst, [System.IO.FileMode]::Create)"
Print #iScript, "$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)"
Print #iScript, "try { [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $src, (Split-Path -Leaf $src)) | Out-Null }"
Print #iScript, "finally { $zip.Dispose(); $fs.Dispose() }"
Close #iScript

'run the script hidden and wait for completion
Set oShell = CreateObject("WScript.Shell")
oShell.Run "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & sScriptPath & """", 0, True

'clean up the temp script
DeleteFile_NoError sScriptPath

'if the zip was created successfully, remove the original uncompressed file and return the zip path
If m_oFSO.FileExists(sZipFilePath_RelApp) Then
    DeleteFile_NoError sSourceFilePath_RelApp
    CompressBackupFileToZip = sZipFilePath_RelApp
Else
    m_oAppLog.AddEntry "CompressBackupFileToZip:  failed to compress '" & sSourceFilePath_RelApp & "' - keeping original .sqb file"
    DeleteFile_NoError sZipFilePath_RelApp
End If

Exit Function
EH:
'don't let a compression failure abort the backup - keep the original .sqb file and log the problem
m_oAppLog.AddEntry "CompressBackupFileToZip:  error compressing '" & sSourceFilePath_RelApp & "' - " & Err.Description
CompressBackupFileToZip = sSourceFilePath_RelApp
End Function

'--delete if not used.
''checks if file exists.  if yes, appends "#n" suffix until unique.
'Public Function GetUniqueFilePath(ByVal sFilePath_RelApp As String) As String
'On Error GoTo EH
'Dim lCounter As Long
'Dim sFNameOnly As String
'Dim sExtOnly As String
'Dim sFolder As String
'
'Do While m_oFSO.FileExists(sFilePath_RelApp)
'    lCounter = lCounter + 1
'
'    'parse the path
'    sFNameOnly = m_oFSO.GetBaseName(sFilePath_RelApp)
'    sExtOnly = m_oFSO.GetExtensionName(sFilePath_RelApp)
'    sFolder = m_oFSO.GetParentFolderName(sFilePath_RelApp)
'
'    sFilePath_RelApp = m_oFSO.BuildPath(sFolder, sFNameOnly & "#" & Format(lCounter) & "." & sExtOnly)
'
'Loop
'
'GetUniqueFilePath = sFilePath_RelApp
'
'Exit Function
'EH:
'ErrHandlerPrivate False, "SQLToolkit", "GetUniqueFilePath"
'End Function

'if possible, converts sFileOrFolderPath to a local path on PC named 'sPCName'.
'assumes that folder sBaseFolder_LocalToPCName is shared as sBaseFolder_ShareName on sPCName.
'If sFileOrFolderPath is a folder, must end with "\"--else, will be assumed to be path to a file.
Public Function GetLocalPath_RelPC(ByVal sPCName As String, _
                                   ByVal sBaseFolder_LocalToPCName As String, _
                                   ByVal sBaseFolder_ShareName As String, _
                                   ByVal sFileOrFolderPath As String) As String
On Error GoTo EH
Dim sDriveName As String
Dim oDrive As Drive
Dim sBaseFolder_UNC As String

'set default = what was passed
GetLocalPath_RelPC = sFileOrFolderPath

sFileOrFolderPath = Trim(sFileOrFolderPath)

'if no target path
If sFileOrFolderPath = "" Then
    Exit Function
End If

'if source path is not a UNC
If Left(sFileOrFolderPath, 2) <> "\\" Then
    'handle mapped drive paths by converting to a UNC
    sDriveName = m_oFSO.GetDriveName(sFileOrFolderPath)
    Set oDrive = m_oFSO.GetDrive(sDriveName)
    'if the path still doesn't look like a UNC
    If Left(oDrive.ShareName, 2) <> "\\" Then
        Exit Function
    End If
    'convert the input file path to UNC form
    '(sDrivename will look like X:, so start beyond the X:\)
     sFileOrFolderPath = m_oFSO.BuildPath(oDrive.ShareName, Mid(sFileOrFolderPath, Len(sDriveName) + 2))
End If

'create UNC to the 'base' folder
sBaseFolder_UNC = m_oFSO.BuildPath("\\" & sPCName, sBaseFolder_ShareName)
'make sure it ends in a backslash
sBaseFolder_UNC = m_oFSO.BuildPath(sBaseFolder_UNC, "\")

'if prefix doesn't match 'base' UNC folder
If Left(UCase(sFileOrFolderPath), Len(sBaseFolder_UNC)) <> UCase(sBaseFolder_UNC) Then
    'can't convert to local path
    Exit Function
End If

'Combine the portion of the file UNC that matches the 'base' local folder
'plus any extra subfolders plus the file name.
GetLocalPath_RelPC = m_oFSO.BuildPath(sBaseFolder_LocalToPCName, Mid(sFileOrFolderPath, Len(sBaseFolder_UNC) + 1))

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetLocalPath_RelPC"
End Function

'if possible, converts sFileOrFolderPath to a shared network path on PC named 'sPCName'.
'assumes that folder sBaseFolder_LocalToPCName is shared as sBaseFolder_ShareName on sPCName.
'If sFileOrFolderPath is a folder, must end with "\"--else, will be assumed to be path to a file.
Public Function GetNWPath_RelPC(ByVal sPCName As String, _
                                ByVal sBaseFolder_LocalToPCName As String, _
                                ByVal sBaseFolder_ShareName As String, _
                                ByVal sFileOrFolderPath As String) As String
On Error GoTo EH
Dim sDriveName_Base As String
Dim sDriveName_Test As String
Dim sFileOrFolderPath_UNC As String

'set default = what was passed
GetNWPath_RelPC = sFileOrFolderPath

sFileOrFolderPath = Trim(sFileOrFolderPath)

'if no target path
If sFileOrFolderPath = "" Then
    Exit Function
End If

'if not a local path
If Left(sFileOrFolderPath, 2) = "\\" Then
   Exit Function
End If

'comparison of local paths are complicated by the two acceptable syntaxes:  x:FolderOrFile and x:\FolderOrFile.
'it's safest to force the x:\ format as some drive commands will *fail* if use x:folder form.
'GetDriveName() *always* returns x: *without* the backslash.
sDriveName_Base = m_oFSO.GetDriveName(sBaseFolder_LocalToPCName)
'recombine and force a backslash (BuildPath eliminates any duplicate slashes)
sBaseFolder_LocalToPCName = m_oFSO.BuildPath(sDriveName_Base & "\", Mid(sBaseFolder_LocalToPCName, Len(sDriveName_Base) + 1))
'also make sure that the local folder ends in backslash (assumed below)
sBaseFolder_LocalToPCName = m_oFSO.BuildPath(sBaseFolder_LocalToPCName, "\")

'repeat for the test path
sDriveName_Test = m_oFSO.GetDriveName(sFileOrFolderPath)
sFileOrFolderPath = m_oFSO.BuildPath(sDriveName_Test & "\", Mid(sFileOrFolderPath, Len(sDriveName_Test) + 1))

'if the prefix of the test path is not the same as the local folder path
If UCase(Left(sFileOrFolderPath, Len(sBaseFolder_LocalToPCName))) <> UCase(sBaseFolder_LocalToPCName) Then
     'no translation is possible
     Exit Function
End If

'convert the local path prefix to a UNC and add on any subfolders or file name
sFileOrFolderPath_UNC = m_oFSO.BuildPath("\\" & sPCName, sBaseFolder_ShareName)
sFileOrFolderPath_UNC = m_oFSO.BuildPath(sFileOrFolderPath_UNC, Mid(sFileOrFolderPath, Len(sBaseFolder_LocalToPCName) + 1))

GetNWPath_RelPC = sFileOrFolderPath_UNC

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetNWPath_RelPC"
End Function

'works with files or folders
Private Function FileExists(sFileFullPath As String) As Boolean
On Error Resume Next
GetAttr sFileFullPath
FileExists = (Err.Number = 0)
'make sure no error is returned to caller
Err.Clear
End Function

'will create nested folder path.  returns true if created or if exists
Private Function CreateFolderIfNec(sFolderPath As String, Optional lErrNumber_Rtn As Long, _
                                   Optional sErrDescription_Rtn As String) As Boolean
On Error Resume Next
Dim lPos As Long
Dim sSubFolder As String

'set default returns
CreateFolderIfNec = False
lErrNumber_Rtn = 0
sErrDescription_Rtn = ""

'do an initial ck for folder exists
If FileExists(sFolderPath) Then
    'exit early with true
    CreateFolderIfNec = True
    GoTo Proc_Exit
End If

'if path starts with UNC
If Left(sFolderPath, 2) = "\\" Then
   'start search for folder after the \\PCName\ prefix
   lPos = InStr(3, sFolderPath, "\")
Else
   lPos = 0
End If

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
             lErrNumber_Rtn = Err.Number
             sErrDescription_Rtn = Err.Description & vbCrLf & vbCrLf _
                             & "[Failed folder = '" & sSubFolder & "]" & vbCrLf _
                             & "[Full folder path = '" & sFolderPath & "']" & vbCrLf _
                             & "[If this is a network folder, check that you are logged into the remote PC.]"
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

Public Function BackupDBToFile(dbSvr As dao.Database, ByVal sDBName As String, ByVal sDestFilePath_RelSvr As String, _
                               Optional ByVal sBkupPassword As String = "", _
                               Optional ByVal sBkupDescr As String = "")
On Error GoTo EH
Dim sTSQL

'Notes:
'1.  Based on testing, if don't specify "WITH FORMAT" or "WITH INIT", the new backup will be appended to any existing file.
'    SQL Server can store more than one backup in a single disk file.  If fail to use WITH INIT, then the restore will require
'    specifying the number or name of the backup set in the disk file--which is probably not worthwhile for WEM use.
'    FORMAT is more severe about clearing the header information and destroys any existing backup.
'    For WEM use, have elected to use FORMAT.
'2.  The backup file extension is user definable  There are no requirements regarding the target file extension.
'3.  The source database is specified in the TSQL, but use the "master" for execution of the stored procedure.
'4.  SQL Server Desktop/Personal edition uses the "Simple" recovery model which does not save inactive transactions.  So,
'    truncate log related options are not necessary.
'5.  The BACKUP/RESTORE will recreate the database with the same file size as the original.  If want to force a shrink, need
'    to execute this explicitly.
'6.  Blank database name causes a TSQL syntax error, so must trap this explicitly
'7.  NON-UNC FILE PATHS ARE ALWAYS INTERPRETED RELATIVE TO THE SERVER'S PC.

If sDBName = "" Then
    'extract the database context of the connection
    sDBName = GetDBContext(dbSvr)
End If

sTSQL = "BACKUP DATABASE [" & sDBName & "]" & vbCrLf _
      & "TO DISK = '" & TSQLQuoteFix(sDestFilePath_RelSvr) & "'" & vbCrLf _
      & "WITH FORMAT"

'if non-blank password specified
If sBkupPassword <> "" Then
    'add on a password for the "backup set" (there is also an option for a "media" password--probably not as useful here)
    sTSQL = sTSQL & "," & vbCrLf _
          & "PASSWORD = '" & TSQLQuoteFix(sBkupPassword) & "'"
End If

'if non-blank description
If sBkupDescr <> "" Then
    sTSQL = sTSQL & "," & vbCrLf _
          & "     DESCRIPTION = '" & TSQLQuoteFix(Left(sBkupDescr, 255)) & "'"
End If

dbSvr.QueryTimeout = 880                   '1.13

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "BackupDBToFile"
End Function

'NOTES:
'If database already exists, exits without error.
'If pass blank string for sDBFileFolder (or if this arg is missing), then the default WEM file folder
'is used = "F:\WEM\Data\SQL".
'To have the database created in the SQL Engine default folder with default file names, pass the string "ENGINE_DEFAULT" as
'the sDBFileFolder argument.

'V1.15:  As JET/ODBC appear to open/close many connections to the database, seems safest to leave AutoClose off by default.
'Also, to prevent issues with large database size, setting AutoShrink to True appears reasonable.  Most likely, this will
'only execute rarely with WEM use.  Based on testing, shrinking does not stop backups from occurring
'(although it may slow backup down if significant shrinking is required).
Public Sub CreateDB(dbSvr As dao.Database, _
                    ByVal sDBName As String, _
                    Optional ByVal sDBFileFolder As String, _
                    Optional ByVal bAutoCloseOption As Boolean = False, _
                    Optional ByVal bAutoShrinkOption As Boolean = True)
On Error GoTo EH
Dim sTSQL As String
Const ENGINE_DEFAULT_FLAG = "ENGINE_DEFAULT"

'exit early if database already exists
If IsDBInSvr(dbSvr, sDBName) Then
    Exit Sub
End If

If (sDBFileFolder = "") Or IsMissing(sDBFileFolder) Then
    sDBFileFolder = "F:\WEM\Data\SQL\"
End If

If UCase(sDBFileFolder) <> ENGINE_DEFAULT_FLAG Then
    If Right(sDBFileFolder, 1) <> "\" Then
       sDBFileFolder = sDBFileFolder & "\"
    End If
End If

'build up the TSQL
sTSQL = ""
sTSQL = sTSQL & "CREATE DATABASE [" & sDBName & "]" & vbCrLf

'if provided a file folder
If UCase(sDBFileFolder) <> ENGINE_DEFAULT_FLAG Then
    'we need to specify everthing
    sTSQL = sTSQL & "ON PRIMARY" & vbCrLf
    '(_DATA is how SQL assigns logical file name to primary file by default)
    'V1.15:  SQL 2005 uses different default naming convention:
    '  data logical file name:  same as database
    '  data file name:  databasename.mdf
    '  log logical file name:  databasename_log
    '  log file name:  databasename_log.ldf
    sTSQL = sTSQL & "   (NAME = [" & sDBName & "]," & vbCrLf
    sTSQL = sTSQL & "   FILENAME = '" & TSQLQuoteFix(sDBFileFolder & sDBName & ".mdf") & "')" & vbCrLf
    sTSQL = sTSQL & "LOG ON" & vbCrLf
    sTSQL = sTSQL & "   (NAME = [" & sDBName & "_log]," & vbCrLf
    sTSQL = sTSQL & "   FILENAME = '" & TSQLQuoteFix(sDBFileFolder & sDBName & "_log.ldf") & "')" & vbCrLf
End If 'if no file folder, let sql svr use default logical file name and physical file names

ExecTSQL dbSvr, sTSQL, False, "master"

'V1.15:  Set the database options (same syntax for SQL2K and SQL2005)
sTSQL = ""
sTSQL = sTSQL & "ALTER DATABASE [" & sDBName & "]" & vbCrLf
sTSQL = sTSQL & "    SET AUTO_CLOSE " & IIf(bAutoCloseOption, "ON", "OFF") & "," & vbCrLf
sTSQL = sTSQL & "        AUTO_SHRINK " & IIf(bAutoShrinkOption, "ON", "OFF") & ";" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "CreateDB"
End Sub

'Notes:
'1.  Dropping a database *deletes* all data files.  (See 'DetachDBFiles' if want to preserve the data files.)
'2.  If database does not exist, no error is raised.
'3.  Existing connections are optionally killed prior to dropping the database
Public Sub DropDB(dbSvr As dao.Database, ByVal sDBName As String, Optional ByVal bKillConnections As Boolean = True)
On Error GoTo EH
Dim sTSQL As String

If bKillConnections Then
     'delete existing connections (note:  this won't prevent clients from attempting to reconnect.
     'could debate putting database into single user mode)
     KillConnectionsToDB dbSvr, sDBName
End If

'if try to drop database that does not exist, will get error
'--this TSQL code exits early without error if database doesn't exist
sTSQL = ""
sTSQL = sTSQL & "IF EXISTS (SELECT *" & vbCrLf
sTSQL = sTSQL & "           FROM master..sysdatabases" & vbCrLf
'V1.17.00:  added quote fix
sTSQL = sTSQL & "           WHERE name = '" & TSQLQuoteFix(sDBName) & "')" & vbCrLf
'V1.17.00:  added brackets
sTSQL = sTSQL & "   DROP DATABASE [" & sDBName & "]" & vbCrLf & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DropDB"
End Sub

'==============================================================================================
'Sample code for starting SQL Server service using "net start" on command line.  Dropped this
'method after got the SQLDMO commands to work
'Public Function StartLocalSQLSvrSvc(ByVal sDSN As String, _
'                                   Optional ByVal sSvrNameSkipDSN As String = "") As Boolean
'On Error GoTo EH
'Dim sServerName As String
'Dim lPos As Long
'Dim sService As String
'Dim lReturn As Long
'
'sServerName = GetServerName(sDSN, sSvrNameSkipDSN)
'
'lPos = InStrRev(sServerName, "\", , vbTextCompare)
'If lPos Then
'    'assume this is a named instance.  We need the portion after the PCname.
'    sServerName = Mid(sServerName, lPos + 1)
'    sService = "MSSQL$" & sServerName
'Else
'    'use the default instance name
'    sService = "MSSQLSERVER"
'End If
'
'lReturn = ExecuteCmdLine("net start " & sService, True)
'
''return success/failure
'StartLocalSQLSvrSvc = (lReturn = 0)
'
'Exit Function
'EH:
'ErrHandlerPrivate False, "SQLToolkit", "StartLocalSQLSvrSvc"
'End Function
'==============================================================================================

'Starts a SQLServer service.  Does *not* need SQL user name/password; however, the code
'must run under a Windows user that has Windows administrative rights on the target machine.
'If not, will fail with error.
'Exits early if service already running.
'V1.08:  since conversion of WEM network to a domain, start times may exceed 10 seconds. Changed
'default for lMaxWaitMs to 30 seconds.
Public Sub StartSQLSvrSvc(ByVal sDSN As String, _
                          Optional ByVal sSvrNameSkipDSN As String = "", _
                          Optional lMaxWaitMs As Long = 30000)
On Error GoTo EH
Dim oServer As SQLDMO.SQLServer
Dim sServerName As String
Dim lTimeStamp As Long
Dim ErrClone1 As ErrClone

Set oServer = New SQLDMO.SQLServer
'V1.08:  Added replacement of "(local)" with PC name (else call will fail)
sServerName = ReplaceLocalWithPCName(GetServerName(sDSN, sSvrNameSkipDSN))

oServer.Name = sServerName

'if server already running
If oServer.status = SQLDMOSvc_Running Then
    'exit early
    Exit Sub
End If

'attempt to start the service without connecting
oServer.start False, sServerName

lTimeStamp = timeGetTime

'wait for the service to fully start before returning.  based on testing, takes about 5-6
'seconds on a remote machine.
Do While (oServer.status <> SQLDMOSvc_Running) And (GetElapsedMs(lTimeStamp) < lMaxWaitMs)
    'pause a short while so not hitting the server object too often.  Don't do DoEvents
    'so clients are not confused by events.
    WaitXMsec 500, False
Loop

If (oServer.status <> SQLDMOSvc_Running) Then
    Err.Raise 1001, App.EXEName & ".SQLToolkit.StartSQLSvrSvc", "Procedure SQLToolkit.StartSQLSvrSvc" _
         & " was unable to start the SQL Server named '" & sServerName & "' after a maximum of" _
         & " " & Format(lMaxWaitMs / 1000#, "0.00") & " seconds."
End If

'close DMO server object or can get orphaned SQL Server connection
CloseDMOServerRef oServer
         
Exit Sub
EH:
ErrClone1 = MakeErrClone
'have to destroy DMO server ref else VB Err.Raise gets confused (reason unknown)
CloseDMOServerRef oServer
'if error due to not having permissions on the target machine.
If ErrClone1.Number = -2147024891 Then
   'add some additional information to the message
   ErrClone1.Description = ErrClone1.Description & vbCrLf _
        & "This error can be due to not having Windows 'Administrator' rights on the target machine or not being logged in." & vbCrLf _
        & "[SQLDMO Error #" & ErrClone1.Number & "]"
   'Note:  the Err.Raise in ErrHandlerPrivate will translate this error number to #70 "permission denied"
   '--unlear why as many other negative values are propagated without translation.
   'To avoid confusion in the logging code, translate to #70 here.
   ErrClone1.Number = 70
End If
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "StartSQLSvrSvc"
End Sub

'This is considered an more "abrupt' shutdown of the SqlSvrSvc than the "Shutdown" method.
'No SQL connection or SQL user name/password is required.  However, you ***must*** be a
'Windows administrator on the target machine.
'Will exit early (without error) if service is already stopped.)
'V1.08:  Changed lMaxWaitMs to 30 seconds (to match default time for StartSQLSvrSvc)
Public Sub StopSQLSvrSvc(ByVal sDSN As String, _
                         Optional ByVal sSvrNameSkipDSN As String = "", _
                         Optional lMaxWaitMs As Long = 30000)
On Error GoTo EH
Dim oServer As SQLDMO.SQLServer
Dim sServerName As String
Dim lTimeStamp As Long
Dim ErrClone1 As ErrClone

Set oServer = New SQLDMO.SQLServer
'V1.08:  Added replacement of "(local)" with PC name (else call will fail)
oServer.Name = ReplaceLocalWithPCName(GetServerName(sDSN, sSvrNameSkipDSN))

'if server already stopped
If oServer.status = SQLDMOSvc_Stopped Then
    'exit early
    Exit Sub
End If

'attempt to stop the service (no connection required)
oServer.Stop

lTimeStamp = timeGetTime

'wait for the service to fully stop before returning.  based on testing, takes about 3-4 seconds on remote machine.
Do While (oServer.status <> SQLDMOSvc_Stopped) And (GetElapsedMs(lTimeStamp) < lMaxWaitMs)
    'pause a short while so not hitting the server object too often.  Don't do DoEvents
    'so clients are not confused by events.
    WaitXMsec 500, False
Loop

If (oServer.status <> SQLDMOSvc_Stopped) Then
    Err.Raise 1001, App.EXEName & ".SQLToolkit.StartSQLSvrSvc", "Procedure SQLToolkit.StartSQLSvrSvc" _
         & " was unable to stop the SQL Server named '" & sServerName & "' after a maximum of" _
         & " " & Format(lMaxWaitMs / 1000#, "0.00") & " seconds."
End If

'close DMO server object or can get orphaned SQL Server connection
CloseDMOServerRef oServer
         
Exit Sub
EH:
ErrClone1 = MakeErrClone
'have to destroy DMO server ref else VB Err.Raise gets confused (reason unknown)
CloseDMOServerRef oServer
'if error due to not having permissions on the target machine.
If ErrClone1.Number = -2147024891 Then
   'add some additional information to the message
   ErrClone1.Description = ErrClone1.Description & vbCrLf _
        & "This error can be due to not having Windows 'Administrator' rights on the target machine." & vbCrLf _
        & "[SQLDMO Error #" & ErrClone1.Number & "]"
   'Note:  the Err.Raise in ErrHandlerPrivate will translate this error number to #70 "permission denied"
   '--unclear why as other DMO error numbers are propagated without translation.
   'To avoid confusion in the logging code, translate to #70 here.
   ErrClone1.Number = 70
End If
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "StopSQLSvrSvc"
End Sub

'V1.07:  (Use ConnectSvrDMO to create/connect oSvrDMO)
Public Function SetSQLSvrSvcAutoStart(oSvrDMO As SQLDMO.SQLServer, _
                                      ByVal bNewAutoStartState As Boolean) As Boolean
On Error GoTo EH

oSvrDMO.Registry.AutostartServer = bNewAutoStartState

'return true if successful
SetSQLSvrSvcAutoStart = True

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "SetSQLSvrSvcAutoStart"
End Function
                                      
'V1.07:  (Use ConnectSvrDMO to create/connect oSvrDMO)
Public Function GetSQLSvrSvcAutoStart(oSvrDMO As SQLDMO.SQLServer) As Boolean
On Error GoTo EH

GetSQLSvrSvcAutoStart = oSvrDMO.Registry.AutostartServer

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetSQLSvrSvcAutoStart"
End Function

'used in various places to convert DSN/optional server name into a server name.
'raises error if unable to find server name.
Public Function GetServerName(ByVal sDSN As String, _
                              Optional ByVal sSvrNameSkipDSN As String = "") As String
On Error GoTo EH
Dim lDSNCount As Long
Dim vDSNNameList As Variant
Dim vServerList As Variant

'if override server name not provided
If sSvrNameSkipDSN = "" Then
    'if DSN not provided
    If sDSN = "" Then
        Err.Raise 1001, App.EXEName & ".SQLToolkit.GetServerName", _
              "Procedure SQLToolkit.GetServerName was passed an empty DSN and empty optional server name."
    End If
    'get the server name for this DSN
    lDSNCount = GetSystemDSNList(sDSN, vDSNNameList, , , vServerList)
    If lDSNCount <> 1 Then
        Err.Raise 1001, App.EXEName & ".SQLToolkit.GetServerName", _
               "Procedure SQLToolkit.GetServerName has been passed" _
             & " a DSN (Data Source Name) that is not a system DSN or is invalid = '" & sDSN & "'."
    End If
    GetServerName = vServerList(0)
Else
    GetServerName = sSvrNameSkipDSN
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetServerName"
End Function

'values returned by reference.
'sServerName_Rtn is the full server name (include machine name and possibly an instance name)
'sMachineName_Rtn in the PC name of the server.
'sInstanceName_Rtn is instance name (or blank string if this is a default instance)
'sServerNameVar_Rtn is the @@SERVERNAME variable.  This is related to connecting to this server using OLEDB. It
'  won't always reflect changes in the local PC name.  For the "standard" server name, sServerName_Rtn is better.
'sProductVersion is the server product version similar to the following:
'   SQL Svr 2000:  "8.00.76.0"
'   SQL Svr 2005:  "9.00.76.0"
'sDefaultDBCmptLevel is the compatibility level of the 'model' database.  This is the compatibility level that
'   a new database will have on this server by default.  This value is in a format that can be compared to other
'   compatibility values (e.g., the "CompatibilityLevel" stored in a db backup file header.)
'   80 = SQL Server 2000
'   90 = SQL Server 2005
Public Function GetServerProperties(dbSvr As dao.Database, _
                                    Optional sServerName_Rtn As String, _
                                    Optional sMachineName_Rtn As String, _
                                    Optional sInstanceName_Rtn As String, _
                                    Optional sServerNameVar_Rtn As String, _
                                    Optional sProductVersion_Rtn As String, _
                                    Optional sDefaultDBCmptLevel_Rtn As String)
On Error GoTo EH
Dim sTSQL As String
Dim rs As dao.Recordset

'set default return values
sServerName_Rtn = ""
sMachineName_Rtn = ""
sInstanceName_Rtn = ""
sServerNameVar_Rtn = ""
sProductVersion_Rtn = ""
sDefaultDBCmptLevel_Rtn = ""
                                    
sTSQL = "SELECT CONVERT(varchar,SERVERPROPERTY('ServerName')) AS ServerName," & vbCrLf _
        & "CONVERT(varchar,SERVERPROPERTY('MachineName')) AS MachineName," & vbCrLf _
        & "CONVERT(varchar,SERVERPROPERTY('InstanceName')) AS InstanceName," & vbCrLf _
        & "@@SERVERNAME AS ServerNameVar," & vbCrLf _
        & "CONVERT(varchar,SERVERPROPERTY('ProductVersion')) AS ProductVersion," & vbCrLf _
        & "(SELECT cmptlevel FROM master..sysdatabases WHERE name = 'model') AS DefaultDBCmptLevel"

Set rs = ExecTSQL(dbSvr, sTSQL, True, "master")

If Not rs.EOF Then
    sServerName_Rtn = GetFV(rs, "ServerName")
    sMachineName_Rtn = GetFV(rs, "MachineName")
    sInstanceName_Rtn = GetFV(rs, "InstanceName")
    sServerNameVar_Rtn = GetFV(rs, "ServerNameVar")
    sProductVersion_Rtn = GetFV(rs, "ProductVersion")
    sDefaultDBCmptLevel_Rtn = GetFV(rs, "DefaultDBCmptLevel")
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetServerProperties"
End Function

'V1.08:  For DMO cmds that do not open a connection, must replace "(local)" with the local PC
'name.  Else, DMO call will fail.
Public Function ReplaceLocalWithPCName(ByVal sServerName As String) As String
On Error GoTo EH
Dim sServerName2 As String
Dim sInitialBackslashes As String
Const LOCAL_NAME_UCASE As String = "(LOCAL)"

'set default return = entry value
ReplaceLocalWithPCName = sServerName

'remove initial space if present
sServerName2 = LTrim(sServerName)

'remove initial back slashes if present
If Left(sServerName2, 2) = "\\" Then
    sInitialBackslashes = "\\"
    sServerName2 = LTrim(Mid(sServerName2, 3))
End If

'if "(local)" prefix
If UCase(Left(sServerName2, Len(LOCAL_NAME_UCASE))) = LOCAL_NAME_UCASE Then
    'remove the (local) portion and any additional spaces
    sServerName2 = LTrim(Mid(sServerName2, Len(LOCAL_NAME_UCASE) + 1))
    'if no remaining chars or if next char is "\"
     If (sServerName2 = "") Or (Left(sServerName2, 1) = "\") Then
         'add back the leading backslashes (if present) plus the PCName plus any sub path.  return to caller
         ReplaceLocalWithPCName = sInitialBackslashes & GetLocalPCName & sServerName2
     End If  'else, (local) is part of the server name (unlikely) => no change
End If  'else, no change

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "ReplaceLocalWithPCName"
End Function

'Notes:
'1.  Function return is status enum value.  Enum also returned in lStatus_RTN and
'test description is in sStatusAsText_RTN.  Additional error information is
'returned in lSQLDMOErrorNumber that is lost (for unknown reasons) when error is propagated by error handler.
'2.  Based on testing, this code must run under security context of Windows user that has
'at least "User" Windows rights on the target PC (no SQL Server rights are nec).
'(NOTE:  Server.PingSQLServerVersion requires making a connection to the SQL Server and is not as useful)
'3.  Based on testing, executing this function DOES use a network connection on the target PC
'(as indicated by the "net sessions" command line function) even if server is on local PC.
'As a result, this function may not be
'a good choice if want to keep # network connections below 10 max for non-server Windows.
'4.  This proc can get the status of a SQL Server even if it cannot be ping'd using TCP/IP.  This can
'provide misleading results if having TCP/IP connectivity problems.
Public Function GetServerStatus(ByVal sDSN As String, _
                           Optional ByVal sSvrNameSkipDSN As String = "", _
                           Optional ByRef lStatus_RTN As SQLDMO_SVCSTATUS_TYPE, _
                           Optional ByRef sStatusAsText_RTN As String, _
                           Optional ByRef lSQLDMOErrorNumber As Long) As SQLDMO_SVCSTATUS_TYPE
On Error GoTo EH
Dim oServer As SQLDMO.SQLServer
Dim vDSNNameList As Variant
Dim vServerList As Variant
Dim sServerName As String
Dim lDSNCount As Long
Dim lSvrVersion As Long
Dim lErrNumber As Long
Dim sErrDescription As String
Dim ErrClone1 As ErrClone

'set default return values
lStatus_RTN = SQLDMOSvc_Unknown
GetServerStatus = lStatus_RTN
sStatusAsText_RTN = "Unknown"
lSQLDMOErrorNumber = 0

Set oServer = New SQLDMO.SQLServer
'V1.08:  Added conversion of "(local)" to local PC name (else, call will fail)
oServer.Name = ReplaceLocalWithPCName(GetServerName(sDSN, sSvrNameSkipDSN))

On Error Resume Next
lStatus_RTN = oServer.status
lSQLDMOErrorNumber = Err.Number
ErrClone1 = MakeErrClone
GetServerStatus = lStatus_RTN
On Error GoTo EH

Select Case ErrClone1.Number
    Case 0  'no error--use status as retrieved
    Case -2147023174   'Server not available => Server doesn't exist/ is unreachable
        lStatus_RTN = SQLDMOSvc_Unknown
        sStatusAsText_RTN = "Unknown:  Server unreachable."
        GoTo Proc_Exit
    Case -2147024891  'Access denied
        lStatus_RTN = SQLDMOSvc_Unknown
        sStatusAsText_RTN = "Unknown:  No Windows permissions on target PC (or not currently logged in)."
        GoTo Proc_Exit
    Case Else
        'for anything else, reraise the error.  be sure to destroy the DMO server object or description is overwritten
        CloseDMOServerRef oServer
        RaiseError ErrClone1
End Select

'translate standard status situations
Select Case lStatus_RTN
    Case SQLDMOSvc_Continuing ' = 6   Service execution state in transition from paused to running.
        sStatusAsText_RTN = "Continuing"
    Case SQLDMOSvc_Paused     ' = 2   Service execution is paused.
        sStatusAsText_RTN = "Paused"
    Case SQLDMOSvc_Pausing    ' = 7   Service execution state in transition from running to paused.
        sStatusAsText_RTN = "Pausing"
    Case SQLDMOSvc_Running    ' = 1   Service is running.
        sStatusAsText_RTN = "Running"
    Case SQLDMOSvc_Starting   ' = 4   Service execution state in transition from stopped to running.
        sStatusAsText_RTN = "Starting"
    Case SQLDMOSvc_Stopped    ' = 3   Service is stopped.
        sStatusAsText_RTN = "Stopped"
    Case SQLDMOSvc_Stopping   ' = 5   Service execution state in transition from running to stopped.
        sStatusAsText_RTN = "Stopping"
    Case SQLDMOSvc_Unknown    ' = 0   Unable to determine service execution state.   (Based on testing, have not
                              '       seen this alternative come up.  If can't access the server, DMO raises an
                              '       error.)
        sStatusAsText_RTN = "Unknown"
    Case Else  'anything else will be returned as "unknown"
End Select

Proc_Exit:
GetServerStatus = lStatus_RTN

'close DMO server object or can get orphaned SQL Server connection
CloseDMOServerRef oServer

Exit Function
EH:
'SQLDMO errors are negative numbers and are translated to VB errors by Err.Raise.  Save the
'original error number in case caller needs this.
lSQLDMOErrorNumber = Err.Number
'must destroy DMO server object or error description is overwritten--unknown whye
ErrClone1 = MakeErrClone
CloseDMOServerRef oServer
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "GetServerStatus"
End Function

'Note:  this creates a temporary connection to SQL Server, but does not use any additional network connections
'to the server (as indicated by "net sessions").  Also, if this fails, it does NOT bring up the SQL ODBC driver
'dialog.  If server exists, returns very quickly.  Otherwise, will take a while to time out.
'This can also be done using API's, but involves a lot of code. See MSDN article Q155304.
Public Function ServerIsReachable(ByVal sDSN As String, _
                                  ByVal sSvrNameSkipDSN As String, _
                                  ByVal sLoginName As String, _
                                  ByVal sPassword As String, _
                                  Optional ByVal bResolveSvrToIP As Boolean = True, _
                                  Optional ByRef oSvrDMO_Rtn As SQLDMO.SQLServer, _
                                  Optional ByRef lErrNumber_Rtn As Long, _
                                  Optional ByRef sErrDescr_RTN As String) As Boolean
On Error Resume Next

'set default return values
ServerIsReachable = False
Set oSvrDMO_Rtn = Nothing
lErrNumber_Rtn = 0
sErrDescr_RTN = ""

Err.Clear
Set oSvrDMO_Rtn = ConnectSvrDMO(sDSN, sSvrNameSkipDSN, sLoginName, sPassword, bResolveSvrToIP)
'can get different errors depending on where failure occured (with server or login/password)
lErrNumber_Rtn = Err.Number
sErrDescr_RTN = Err.Description
ServerIsReachable = (lErrNumber_Rtn = 0)
Err.Clear
End Function

'Shuts down the SQL Server instance.  If leave bAllowSvrCleanup as default, will do more "orderly" shutdown
'of SQL Server service than StopServer() method.  Does *not* require Windows user access to the target PC, but
'*does* require a valid SQL login/password.
'Supresses error if service is already off or unreachable.
'Based on testing, service will not be in 'stopped' status until ~3 seconds after Server.Shutdown.
'To prevent need for Windows user access to target PC, routine does *NOT* wait for service to
'stop before returning.  If desired, caller can use GetServerStatus()--but this does require Windows login.
'See StopServer for more alternate method.
'(Could not get the TSQL "SHUTDOWN" to work consistently without errors, so resorted to DMO)
'(Based on documentation, use of bAllowSvrCleanup=False requires SysAdmin login.)
Public Sub ShutdownServer(ByVal sDSN As String, _
                          ByVal sSvrNameSkipDSN As String, _
                          ByVal sLoginName As String, _
                          ByVal sPassword As String, _
                          Optional ByVal bAllowSvrCleanup As Boolean = True, _
                          Optional lMaxWaitMs As Long = 10000, _
                          Optional ByVal bResolveSvrToIP As Boolean = True)
On Error GoTo EH
Dim oServer As SQLDMO.SQLServer
Dim sServerName As String
Dim lServerStatus As SQLDMO_SVCSTATUS_TYPE
Dim ErrClone1 As ErrClone

If bResolveSvrToIP Then
    sServerName = ResolveSvrEngineToIP(sDSN, sSvrNameSkipDSN)
Else
    sServerName = GetServerName(sDSN, sSvrNameSkipDSN)
End If

Set oServer = New SQLDMO.SQLServer

'based on testing, can't execute unless first connect. (see StopSvr for method that does not require
'a connection)  If server is unreachable, this will **NOT** bring up any driver dialogs.
On Error Resume Next
oServer.Connect sServerName, sLoginName, sPassword
ErrClone1 = MakeErrClone
On Error GoTo EH
Select Case ErrClone1.Number
    'able to connect.  drop through.
    Case 0
    Case -2147221504
        'server is already off or is unreachable.  exit without error
        Err.Clear
        Exit Sub
    Case Else
        'something unexpected--reraise the error.  destroy the server object to prevent issues with Err.Raise
        CloseDMOServerRef oServer
        RaiseError ErrClone1
End Select

'attempt to shutdown the server
On Error Resume Next
oServer.Shutdown bAllowSvrCleanup
ErrClone1 = MakeErrClone
On Error GoTo EH

Select Case ErrClone1.Number
   'if no error (based on testing, this doesn't occur--if successful, get #-2147215802 trapped below)
   Case 0
       Exit Sub
   'if this is informational error message indicating shutdown (note:  based on testing, service actually takes
   'another 3 seconds or so to complete shutdown)
   Case -2147215802
       Exit Sub
   Case Else
       'something unexpected--reraise the error.  destroy the server object to prevent issues with Err.Raise
       CloseDMOServerRef oServer
       RaiseError ErrClone1
End Select

'close DMO server object or can get orphaned SQL Server connection
CloseDMOServerRef oServer

Exit Sub
EH:
'make sure no issues with oServer object intefering with Err.Raise by destroying it here
ErrClone1 = MakeErrClone
CloseDMOServerRef oServer
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "ShutdownServer"
End Sub

'on used in ShutdownServer.
'based on testing, there's no good way to use the oServer object to test for shutdown.  However, the
'oServer.Shutdown method can take several seconds to complete, so must wait until done.
'For this routine to work, current Windows login must have at least 'user' rights on the target machine.
Private Sub WaitForShutdownToComplete(ByVal sServerName As String, ByVal lMaxWaitMSec As Long)
On Error Resume Next
Dim lTimeStamp As Long

lTimeStamp = timeGetTime

'wait for the service to fully stop before returning.  based on testing, takes about 3-4 seconds on remote machine.
Do While (GetServerStatus("", sServerName) <> SQLDMOSvc_Stopped) And (GetElapsedMs(lTimeStamp) < lMaxWaitMSec)
    'pause a short while so not hitting the server object too often.  Don't do DoEvents
    'so clients are not confused by events.
    WaitXMsec 500, False
    Debug.Print GetServerStatus("", sServerName), GetElapsedMs(lTimeStamp)
Loop
Debug.Print GetServerStatus("", sServerName), GetElapsedMs(lTimeStamp)
Err.Clear

End Sub

Public Sub CloseDMOServerRef(oServer As SQLDMO.SQLServer)
On Error Resume Next
If Not (oServer Is Nothing) Then
    'must execute close or, based on testing, will always get an orphaned SQL Server connection
    oServer.Close
End If
Set oServer = Nothing
Err.Clear
End Sub

'Kills all connections between the specified server and the specified database.
Public Sub KillConnectionsToDB(dbSvr As dao.Database, ByVal sDBName As String)
On Error GoTo EH
Dim sTSQL As String

sTSQL = ""
sTSQL = sTSQL & "BEGIN" & vbCrLf
sTSQL = sTSQL & "    DECLARE @dbname sysname" & vbCrLf
'V1.17:  added quote fix
sTSQL = sTSQL & "    SET @dbname = '" & TSQLQuoteFix(sDBName) & "'" & vbCrLf
sTSQL = sTSQL & "    DECLARE loop_name INSENSITIVE CURSOR FOR" & vbCrLf
sTSQL = sTSQL & "        SELECT spid" & vbCrLf
sTSQL = sTSQL & "        FROM master..sysprocesses" & vbCrLf
sTSQL = sTSQL & "        WHERE dbid = DB_ID(@dbname)" & vbCrLf
sTSQL = sTSQL & "    OPEN loop_name" & vbCrLf
sTSQL = sTSQL & "    DECLARE @conn_id SMALLINT" & vbCrLf
sTSQL = sTSQL & "    DECLARE @exec_str VARCHAR(255)" & vbCrLf
sTSQL = sTSQL & "    FETCH NEXT FROM loop_name INTO @conn_id" & vbCrLf
sTSQL = sTSQL & "    WHILE (@@fetch_status = 0)" & vbCrLf
sTSQL = sTSQL & "        BEGIN" & vbCrLf
sTSQL = sTSQL & "            SET @exec_str = 'KILL ' + CONVERT(VARCHAR(10), @conn_id)" & vbCrLf
sTSQL = sTSQL & "            EXECUTE( @exec_str )" & vbCrLf
sTSQL = sTSQL & "            FETCH NEXT FROM loop_name INTO @conn_id" & vbCrLf
sTSQL = sTSQL & "        END" & vbCrLf
sTSQL = sTSQL & "    DEALLOCATE loop_name" & vbCrLf
sTSQL = sTSQL & "END" & vbCrLf

'must be sure to specify 'master'.  if attempt to drop connections to current database context, SQL
'will give an error.
ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "KillConnectionsToDB"
End Sub

'Input file path MUST BE RELATIVE TO THE INSTANCE OF SQL SERVER.  This can be confusing if attaching files over a network.
Public Sub AttachDBFiles(dbSvr As dao.Database, sDBName As String, _
                         sRelPathToMDFFile As String, _
                         sRelPathToLDFFile As String, _
                         Optional bMapDBUsersToSvrLogins As Boolean = True)
On Error GoTo EH
Dim sTSQL As String

sTSQL = ""
'V1.17:  added quote fix
sTSQL = sTSQL & "sp_attach_db '" & TSQLQuoteFix(sDBName) & "','" & TSQLQuoteFix(sRelPathToMDFFile) & "','" & TSQLQuoteFix(sRelPathToLDFFile) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

If bMapDBUsersToSvrLogins Then
    'make an attempt to fix orphaned users
     MapDBUsersToSvrLogins dbSvr, sDBName
End If

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "AttachDBFiles"
End Sub

Public Sub DetachDBFiles(dbSvr As dao.Database, ByVal sDBName As String, Optional ByVal bKillConnections As Boolean = True)
On Error GoTo EH
Dim sTSQL As String

'no one can be connected to the specified database or the detach will fail.  so, in general, killing the connections
'is desirable.  (we could also put the database into single user mode to prevent any new connections, but this mode
'would then end up stored with the database and would have to be handled if/when the database is reattached).  It's unclear
'what the effect of the "skipchecks" option is, so have left this off so default will be used.
If bKillConnections Then
     KillConnectionsToDB dbSvr, sDBName
End If

sTSQL = ""
'V1.17:  added quote fix
sTSQL = sTSQL & "sp_detach_db '" & TSQLQuoteFix(sDBName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DetachDBFiles"
End Sub

'file paths returned are RELATIVE TO THE SQL SERVER INSTANCE.  This can be confusing if accessing files over a network.
'Can only handle situations where there is an mdf file and a single ldf file--does not support file groups.
Public Sub GetDBFileInfo(dbSvr As dao.Database, ByVal sDBName As String, _
                         sMDFRelPath_RTN As String, _
                         sMDFLogicalName_RTN As String, _
                         sLDFRelPath_RTN As String, _
                         sLDFLogicalName_RTN As String)
On Error GoTo EH
Dim rsResult As dao.Recordset
Dim sTSQL As String

'set default return values (in case of error)
sMDFRelPath_RTN = ""
sMDFLogicalName_RTN = ""
sLDFRelPath_RTN = ""
sLDFLogicalName_RTN = ""

sTSQL = ""
'need RTRIM as SQL adds trailing filler spaces.  also, "filename" is 260 chars and DAO appears to fail and return
'blank field if leave the trailing spaces.
sTSQL = sTSQL & "SELECT RTRIM(sysfiles.name) AS NAME," & vbCrLf
sTSQL = sTSQL & "    RTRIM(sysfiles.filename) AS FILENAME," & vbCrLf
sTSQL = sTSQL & "    RTRIM(sysfilegroups.groupname) AS GROUPNAME" & vbCrLf
sTSQL = sTSQL & "FROM sysfiles" & vbCrLf
sTSQL = sTSQL & "LEFT JOIN sysfilegroups" & vbCrLf
sTSQL = sTSQL & "ON sysfiles.groupid = sysfilegroups.groupid" & vbCrLf

Set rsResult = ExecTSQL(dbSvr, sTSQL, True, sDBName)

With rsResult
    Do While Not .EOF
        If UCase(rsResult!groupname & "") = "PRIMARY" Then
             sMDFRelPath_RTN = !filename & ""
             sMDFLogicalName_RTN = !Name & ""
        Else
             sLDFRelPath_RTN = !filename & ""
             sLDFLogicalName_RTN = !Name & ""
        End If
        .MoveNext
    Loop
End With
rsResult.Close

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetDBFileInfo"
End Sub

'NOTES:
'1.  This function will ***overwrite*** any existing database.  It will also allow restoring a database
'to a new name without warning.  It's up to the caller to test ahead of time if these possibilities are
'not desirable.
'2.  If restoring over an existing database, existing connections to the database will be broken.
'If using JET/ODBC, then "stranded" ODBC connections can be left in the JET ODBC connection pool.  These
'bad connections may stay in the pool for a long time (600 sec or indefinitely if DAO db object remains open) and
'may cause errors.  The only definitive/global way to clear these is to close any open applications and restart
'them.
'3.  NON-UNC FILE PATHS ARE ALWAYS INTERPRETED RELATIVE TO THE SERVER'S PC.
'    BEST TO ONLY EXECUTE THIS PROC ON THE SERVER MACHINE.
Public Function RestoreDBFromFile(dbSvr As dao.Database, _
                                  ByVal sSrceFilePath_RelSvr As String, _
                                  ByVal sDestDBName As String, _
                                  Optional ByVal sDestMDFFilePath_RelSvr As String = "", _
                                  Optional ByVal sDestLDFFilePath_RelSvr As String = "", _
                                  Optional ByVal bMapDBUsersToSvrLogins As Boolean = True, _
                                  Optional ByVal sBkupPassword As String = "")
On Error GoTo EH
Dim rsTemp As dao.Recordset
Dim sTSQL As String
Dim sTSQL_Temp As String
Dim bIsDBInSvr As Boolean
Dim sBackupMDFPath As String
Dim sBackupMDFLogicalName As String
Dim sBackupLDFPath As String
Dim sBackupLDFLogicalName As String
Dim vTemp As Variant
Dim sODBCConnStr As String
Dim lTimeStamp As Long

If sDestDBName = "" Then
    Err.Raise 1001, "SQLToolkit.RestoreDBFromFile", "Procedure SQLToolkit.BackupDBToFile has been passed a" _
        & " blank target database name."
End If

'if target file path not provided
If sDestMDFFilePath_RelSvr = "" Then
    'use standard WEM default file path/names
    'NOTE:  If \SQL folder does not exist on the server PC, the restore will fail.
    sDestMDFFilePath_RelSvr = "F:\WEM\Data\SQL\" & sDestDBName & ".mdf"
    sDestLDFFilePath_RelSvr = "F:\WEM\Data\SQL\" & sDestDBName & ".ldf"
End If

bIsDBInSvr = IsDBInSvr(dbSvr, sDestDBName)

If bIsDBInSvr Then
    'we have to make sure that everyone else is out of the existing database.  (we could also just DROP the
    'database, but this requires 'killing' all processes IDs which is more involved.
    'this syntax forcefully disconnects any other users without any delay.  It has the advantage that
    'it also prevents anyone else from connecting to the database in the meantime.)
    sTSQL = sTSQL & vbCrLf _
          & "ALTER DATABASE [" & sDestDBName & "]" & vbCrLf _
          & "   SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
End If

'to explicitly set the output locations of the .mdf and .ldf files, we need the logical file names.
'(there are several conventions for naming these, so can't calc from db name)
GetBackupFileInfo_FileList dbSvr, sSrceFilePath_RelSvr, sBkupPassword, sBackupMDFPath, sBackupMDFLogicalName, _
                           sBackupLDFPath, sBackupLDFLogicalName
           
'create combined string.  The WITH REPLACE will force an overwrite of an existing database in the server even
'if the database name in the backup file is different than the target database name.
'By default, SQL Server executes WITH RECOVERY
'which rolls back any uncommitted transactions and leaves the database in a "usable" state.  (In future, if want to apply
'transaction log backups to the restored database, then will need option to add WITH NORECOVERY or WITH STANDBY).
sTSQL = sTSQL & vbCrLf _
     & "RESTORE DATABASE [" & sDestDBName & "]" & vbCrLf _
     & "   FROM DISK = '" & TSQLQuoteFix(sSrceFilePath_RelSvr) & "'" & vbCrLf _
     & "   WITH" & vbCrLf _
     & "      REPLACE"
     
'if non-blank password provided
If sBkupPassword <> "" Then
    sTSQL = sTSQL & vbCrLf _
        & "      ,PASSWORD='" & TSQLQuoteFix(sBkupPassword) & "'"
End If
        
'give an error if target for database and trans log are the same--not sure what SQL Svr would do in that case
If UCase(sDestMDFFilePath_RelSvr) = UCase(sDestLDFFilePath_RelSvr) Then
    Err.Raise 1001, App.EXEName & ".SQLToolkit.RestoreDBFromFile", "Procedure SQLToolkit.RestoreDBFromFile has been passed" _
        & " target file paths for the SQL Server .mdf and .ldf files that are the same.  This is not allowed." & vbCrLf & vbCrLf _
        & "Target file path = '" & sDestMDFFilePath_RelSvr & "'"
End If

'tell SQL where to put the mdf/ldf files
sTSQL = sTSQL & vbCrLf _
     & "      ,MOVE '" & TSQLQuoteFix(sBackupMDFLogicalName) & "' TO '" & TSQLQuoteFix(sDestMDFFilePath_RelSvr) & "'" & vbCrLf _
     & "      ,MOVE '" & TSQLQuoteFix(sBackupLDFLogicalName) & "' TO '" & TSQLQuoteFix(sDestLDFFilePath_RelSvr) & "'"

'if we previously set to single user mode
If bIsDBInSvr Then
    'add syntax for returning database back to multi-user mode (ideally, should really test if this was the mode before
    'executing the restore--see UserAccess property of DATABASEPROPERTYEX)
    'Based on testing, even if the RESTORE fails with an error, the ALTER DATABASE command that follows is executed.
    'If database gets stuck in single user mode, "(Single User)" is added to the database name in Enterprise Man.
    sTSQL = sTSQL & vbCrLf _
         & "ALTER DATABASE [" & sDestDBName & "]" & vbCrLf _
         & "   SET MULTI_USER" & vbCrLf
End If

dbSvr.QueryTimeout = 880                   'RBH

ExecTSQL dbSvr, sTSQL, False, "master"

'if need to re-map the db user ID's to server logins
' 1.08 Chinese
'If bMapDBUsersToSvrLogins And g_oAppSettings.MainDB_Language <> "Chinese" Then
'     MapDBUsersToSvrLogins dbSvr, sDestDBName
'End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "RestoreDBFromFile"
End Function

'version of preceding functions with hard coded defaults for standard WEM systems.
'if successful, returns name of the restored database
Public Function RestoreDBFromFile2(dbSvr As dao.Database, _
                                  ByVal sBkupFilePath_RelApp As String, _
                                  Optional ByVal sDestDBName As String = "", _
                                  Optional ByVal sBkupPassword As String = "", _
                                  Optional ByVal bShrinkDB As Boolean = True, _
                                  Optional ByVal sBaseFolder_LocalToSvr As String = "F:\", _
                                  Optional ByVal sBaseFolder_ShareName As String = "WEMDATA") As String
On Error GoTo EH
Dim sSvrPCName As String
Dim sDBName_FromHeader As String
Dim sMdfLdfFolder_RelApp As String
Dim sMDFFilePath_RelSvr As String
Dim sLDFFilePath_RelSvr As String
Dim sCount As String
Dim bTempCopyNeeded As Boolean
Dim sSrceFilePath_RelSvr As String
Dim sTempFilePath_RelApp As String
Dim sTempFilePath_RelSvr As String
Dim lErrNumber As Long
Dim sErrDescription As String
Dim ErrClone1 As ErrClone

'if source file doesn't exist
If Not m_oFSO.FileExists(sBkupFilePath_RelApp) Then
    'error out now or will get confusing error messages from server below
    Err.Raise 1001, "SQLToolkit.RestoreDBFromFile2", _
           "Procedure SQLToolkit.RestoreDBFromFile2 has been passed" _
         & " a source backup file path that does not exist or is unreachable." & vbCrLf & vbCrLf _
         & "[Source backup file path relative to application = '" & sBkupFilePath_RelApp & "']"
End If

'get server's machine name
GetServerProperties dbSvr, , sSvrPCName

'test if we need (and have) access to shared folders on the server PC (error is raised if there's a problem)
ValidateSvrNWShare dbSvr, sSvrPCName, sBaseFolder_LocalToSvr, sBaseFolder_ShareName, sBkupFilePath_RelApp

'determine if server is likely to have access to this file
'get path to this file that the server can use
bTempCopyNeeded = GetFilePath_RelSvr(dbSvr, sBkupFilePath_RelApp, _
                                     sSrceFilePath_RelSvr, _
                                     sTempFilePath_RelApp, _
                                     sTempFilePath_RelSvr, _
                                     sBaseFolder_LocalToSvr, _
                                     sBaseFolder_ShareName)
'if copy is needed
If bTempCopyNeeded Then
    'copy backup file to the server's \Temp folder.
    m_oFSO.CopyFile sBkupFilePath_RelApp, sTempFilePath_RelApp, True
End If

'build path to the mdf/ldf folder.  if server is local
If UCase(sSvrPCName) = UCase(GetLocalPCName) Then
    sMdfLdfFolder_RelApp = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\SQL")
Else 'else, server is remote
    sMdfLdfFolder_RelApp = m_oFSO.BuildPath("\\" & sSvrPCName, sBaseFolder_ShareName)
    sMdfLdfFolder_RelApp = m_oFSO.BuildPath(sMdfLdfFolder_RelApp, "WEM\Data\SQL")
End If

'if dest db name not given
If sDestDBName = "" Then
    'assume we need to restore to original name.  extract this from the backup file
    GetBackupFileInfo_Header dbSvr, IIf(bTempCopyNeeded, sTempFilePath_RelSvr, sSrceFilePath_RelSvr), _
                                    sBkupPassword, sDBName_FromHeader
    sDestDBName = sDBName_FromHeader
End If

'make sure that the target folder for the ldf/mdf file exists
If Not CreateFolderIfNec(sMdfLdfFolder_RelApp, lErrNumber, sErrDescription) Then
     Err.Raise lErrNumber, "SQLToolkit.RestoreDBFromFile2", "An error occurred when creating a destination folder" _
       & " for the SQL Server .mdf/.ldf files.  (File paths are relative to current local application.)" & vbCrLf & vbCrLf _
       & sErrDescription
End If

'create standard MDF/LDF file names for server--these must be local to svr.
sMDFFilePath_RelSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\SQL") & "\" & sDestDBName & ".mdf"
sLDFFilePath_RelSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\SQL") & "\" & sDestDBName & "_log.ldf"

'execute the restore (with mapping of db users to server logins)
RestoreDBFromFile dbSvr, IIf(bTempCopyNeeded, sTempFilePath_RelSvr, sSrceFilePath_RelSvr), _
                         sDestDBName, sMDFFilePath_RelSvr, sLDFFilePath_RelSvr, True, sBkupPassword

'if we had to make a copy of source file in \Temp folder
If bTempCopyNeeded Then
    'delete the temporary file
    DeleteFile_NoError sTempFilePath_RelSvr
End If  'else, leave the source backup file alone

'if need to shrink the database
If bShrinkDB Then
    ShrinkDatabase dbSvr, sDestDBName
End If

'return final database name
RestoreDBFromFile2 = sDestDBName

Exit Function
EH:
ErrClone1 = MakeErrClone
If bTempCopyNeeded Then
    'clean up the temp file on the server
    DeleteFile_NoError sTempFilePath_RelSvr
End If
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "RestoreDBFromFile2"
End Function

Public Function DeleteFile_NoError(sFilePath As String)
On Error Resume Next
'delete the file
m_oFSO.Deletefile sFilePath, True
Err.Clear
End Function

'makes sure specified file is accessible to the server.  intended use is backup files.
'Params:
'    sFilePath_RelApp:  file path relative to current app.  File does *not* have to exist.

'Returns:
'     False:  file is accessible to server in current location.
'             sFilePath_RelSvr_Rtn:  path to the original file, relative to the server.
'             sTempFilePath_RelApp_Rtn:  is blank
'             sTempFilePath_RelSvr_Rtn:  is blank
'     True:   file may not be accessible to server in current location.
'             A copy to the server \Temp folder is needed.
'             sFilePath_RelSvr_Rtn is blank
'             sTempFilePath_RelSvr_Rtn:  temp file location relative to the server.
'             sTempFilePath_RelApp_Rtn:  temp file location relative to the app
'             The temp file paths are guaranteed unique at the time of execution (number is
'             appended to file names to force uniqueness).
'

'It's up to the caller to execute the file copy with code similar to the following:
'bTempFileCopy = GetFilePath_RelSvr(dbSvr, sFilePath_RelApp, sFilePath_RelSvr, sTempFilePath_RelSvr)
'If bTempFileCopy Then
'    m_oFSO.CopyFile sFilePath_RelApp, sTempFilePath_RelSvr
'End if
'
'then, at end of procedure (and in error handler with preservation of error info), use
'If bTempFileCopy Then
'    DeleteFile_NoError sTempFilePath_RelApp
'end if

Public Function GetFilePath_RelSvr(dbSvr As dao.Database, _
                                   ByVal sFilePath_RelApp As String, _
                                   sFilePath_RelSvr_Rtn As String, _
                                   sTempFilePath_RelApp_Rtn As String, _
                                   sTempFilePath_RelSvr_Rtn As String, _
                                   Optional ByVal sBaseFolder_LocalToSvr As String = "F:\", _
                                   Optional ByVal sBaseFolder_ShareName As String = "WEMDATA") As Boolean
On Error GoTo EH
Dim oDrive As Drive
Dim sDriveName As String
Dim sServerName As String
Dim sSvrPCName As String
Dim sInstanceName As String
Dim sSvrPCName_UNC As String
Dim sBaseFolder_UNC As String
Dim sTempFolder_RelApp As String
Dim sTempFolder_RelSvr As String
Dim sTempFileName As String
Dim sTempFilePath_RelApp As String
Dim sTempFilePath_RelSvr As String
Dim bSrceIsNetworkPath As Boolean
Dim sFNameOnly As String
Dim sExtOnly As String
Dim sCount As String
Dim lErrNumber As Long
Dim sErrDescription As String

'set default => no copy made
GetFilePath_RelSvr = False
sFilePath_RelSvr_Rtn = ""
sTempFilePath_RelApp_Rtn = ""
sTempFilePath_RelSvr_Rtn = ""

sFilePath_RelApp = Trim(sFilePath_RelApp)

'get the location of the server
GetServerProperties dbSvr, sServerName, sSvrPCName, sInstanceName

'create UNC's for comparison
sSvrPCName_UNC = "\\" & sSvrPCName & "\"

'create UNC to the 'base' folder
sBaseFolder_UNC = m_oFSO.BuildPath(sSvrPCName_UNC, sBaseFolder_ShareName)
'make sure it ends in a backslash
sBaseFolder_UNC = m_oFSO.BuildPath(sBaseFolder_UNC, "\")

'-----------------
'if source path is not a UNC
If Left(sFilePath_RelApp, 2) <> "\\" Then
    'handle mapped drive paths by converting to a UNC
    sDriveName = m_oFSO.GetDriveName(sFilePath_RelApp)
    Set oDrive = m_oFSO.GetDrive(sDriveName)
    'if the drive is shared and looks like a UNC
    If Left(oDrive.ShareName, 2) = "\\" Then
        'convert the input file path to UNC form
        '(sDrivename will look like X:, so start beyond the X:\)
         sFilePath_RelApp = m_oFSO.BuildPath(oDrive.ShareName, Mid(sFilePath_RelApp, Len(sDriveName) + 2))
    End If
End If
'-----------------

'if the source path has the form of a UNC
If Left(sFilePath_RelApp, 2) = "\\" Then
    bSrceIsNetworkPath = True
    Select Case True
        'if the network folder of the source file matches the known server base drive share
        Case Left(UCase(sFilePath_RelApp), Len(sBaseFolder_UNC)) = UCase(sBaseFolder_UNC)
            'then we can use a local path on the server.  Combine the portion of the file UNC that matches the 'base' folder
            'plus any extra subfolders plus the file name.
            sFilePath_RelSvr_Rtn = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, Mid(sFilePath_RelApp, Len(sBaseFolder_UNC) + 1))
            Exit Function
        'if the machine name matches the server
        Case Left(UCase(sFilePath_RelApp), Len(sSvrPCName_UNC)) = UCase(sSvrPCName_UNC)
             'then the UNC to the file is not to the 'base' folder that we can translate into a local path,
             'but it will be accessible from the server.
             '(note:  this may not be true if SQL Server is running under a very restricted account--but this should be rare).
             sFilePath_RelSvr_Rtn = sFilePath_RelApp
             Exit Function
    End Select
End If

'if server is local AND source file is local
If (UCase(sSvrPCName) = UCase(GetLocalPCName)) And (Not bSrceIsNetworkPath) Then
    'provided path is local and is always accessible
    '(this will be the most common case when routine is run from the FSERVER)
    sFilePath_RelSvr_Rtn = sFilePath_RelApp
    Exit Function
End If

'Either:
'(1) the provided path is a UNC to something other than the server machine.  based on testing, can
'    have issues with SQL Server not having access to the shared folder => need to copy to server
'    OR
'(2) we've been passed a local path (relative to this app), but this app is not local to the server
'    => path is not local to the server => need to copy to server.

'if the server is local
If UCase(sSvrPCName) = UCase(GetLocalPCName) Then
    'path to the temp file relative to this app can be made local
    sTempFolder_RelApp = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\Temp")
Else
    'path to the temp file rel to this app must be a UNC
    sTempFolder_RelApp = m_oFSO.BuildPath("\\" & sSvrPCName, sBaseFolder_ShareName)
    sTempFolder_RelApp = m_oFSO.BuildPath(sTempFolder_RelApp, "WEM\Data\Temp")
End If

'server path to the temp folder must be local to the server
sTempFolder_RelSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "WEM\Data\Temp")

'make sure the temp folder exists or can be created
If Not CreateFolderIfNec(sTempFolder_RelApp, lErrNumber, sErrDescription) Then
    Err.Raise lErrNumber, "SQLToolkit.GetFilePath_RelSvr", "An error occurred when creating the '\temp' folder" _
       & " for a file copy to the SQL Server PC.  (File paths are relative to current local application.)" & vbCrLf & vbCrLf _
       & sErrDescription
End If

sFNameOnly = m_oFSO.GetBaseName(sFilePath_RelApp)
sExtOnly = m_oFSO.GetExtensionName(sFilePath_RelApp)
'create unique temp file name
sCount = "1"
Do
    sTempFileName = sFNameOnly & "#" & sCount & "." & sExtOnly & ".tmp"
    sTempFilePath_RelApp = m_oFSO.BuildPath(sTempFolder_RelApp, sTempFileName)
    sTempFilePath_RelSvr = m_oFSO.BuildPath(sTempFolder_RelSvr, sTempFileName)
    sCount = Val(sCount) + 1
Loop While m_oFSO.FileExists(sTempFilePath_RelApp)

'if make it to here, return copy flag and results
sTempFilePath_RelApp_Rtn = sTempFilePath_RelApp
sTempFilePath_RelSvr_Rtn = sTempFilePath_RelSvr

'indicate that a copy to the temp folder is required
GetFilePath_RelSvr = True

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetFilePath_RelSvr"
End Function

'specialized routine used by backup/restore procs to verify that F: on the server PC is the
'same as \\<ServerPC>\WEMData.  Only executes the test if necessary based on test for
'either:  (1) app not local to the server OR
'         (2) specified (non blank) target folder is a network path that has \\<server pc.\SharedFolder prefix.
'--these are the only two cases in which may need to assume that the local path and network share are the same.
'Raises error if shared folder test is needed and test does not succeed.
'It's assumed caller has obtained sSvrPCName from GetServerProperties().
Public Function ValidateSvrNWShare(dbSvr As dao.Database, _
                                   ByVal sSvrPCName As String, _
                                   ByVal sBaseFolder_LocalToSvr As String, _
                                   ByVal sBaseFolder_ShareName As String, _
                                   Optional ByVal sTargetFolder As String = "") As Boolean
On Error GoTo EH
Dim bTestNeeded As Boolean
Dim sLogErrors_OldFlag As String
Dim sFolderPath_UNC As String
Dim sTempFileName As String
Dim sTempFilePath_UNC As String
Dim sTempFilePath_LocalToSvr As String
Dim oTextFile As TextStream
Dim oDrive As Drive
Dim sDriveName As String
Dim sTSQL As String
Dim ErrClone1 As ErrClone
Dim bEmptyFileError As Boolean
Dim bFileNotExistError As Boolean
Dim lIndex As Long
Dim sMsg As String

'set default
bTestNeeded = False

'make sure provided folders end in "\" (else can have x: versus x:\ issues with local paths)
If sBaseFolder_LocalToSvr <> "" Then
    sBaseFolder_LocalToSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, "\")
End If
If sBaseFolder_ShareName <> "" Then
    sBaseFolder_ShareName = m_oFSO.BuildPath(sBaseFolder_ShareName, "\")
End If

'get folder path into the server PC as UNC
sFolderPath_UNC = m_oFSO.BuildPath("\\" & sSvrPCName, sBaseFolder_ShareName)
'make sure ends in backslash
sFolderPath_UNC = m_oFSO.BuildPath(sFolderPath_UNC, "\")

'if current app is not running on the server PC
If UCase(GetLocalPCName) <> UCase(sSvrPCName) Then
    bTestNeeded = True
'else, if a test folder is specified
ElseIf sTargetFolder <> "" Then
    'if not a UNC
    If Left(sTargetFolder, 2) <> "\\" Then
        'handle mapped drive paths by converting to a UNC
        sDriveName = m_oFSO.GetDriveName(sTargetFolder)
        Set oDrive = m_oFSO.GetDrive(sDriveName)
        'if the drive is shared and looks like a UNC
        If Left(oDrive.ShareName, 2) = "\\" Then
            'convert the input file path to UNC form
            '(sDrivename will look like X:, so start beyond the X:\)
             sTargetFolder = m_oFSO.BuildPath(oDrive.ShareName, Mid(sTargetFolder, Len(sDriveName) + 2))
        End If
    End If
    'make sure folder ends in backslash
    sTargetFolder = m_oFSO.BuildPath(sTargetFolder, "\")
    
    'if target folder path uses the std shared folder on the server
    If Left(UCase(sTargetFolder), Len(sFolderPath_UNC)) = UCase(sFolderPath_UNC) Then
        'then need to verify the folder share = local path
        bTestNeeded = True
    End If
End If

If Not bTestNeeded Then
    'exit early with success
    ValidateSvrNWShare = True
    Exit Function
End If

If Not m_oFSO.FolderExists(sFolderPath_UNC) Then
    Err.Raise 1001, "SQLToolkit.ValidateSvrNWShare", "The application needs access to the" _
         & " following shared network folder.  This folder either does not exist or is not shared or is not" _
         & " accessible.  Please check the network shares, the network communication, and" _
         & " whether a logon to the PC is required." & vbCrLf & vbCrLf _
         & "[Network share = '" & sFolderPath_UNC & "']"
End If

'create temp file name and local/UNC paths to it
sTempFileName = "TestSQLSvrNWShare_" & m_oFSO.GetTempName
sTempFilePath_UNC = m_oFSO.BuildPath(sFolderPath_UNC, sTempFileName)
sTempFilePath_LocalToSvr = m_oFSO.BuildPath(sBaseFolder_LocalToSvr, sTempFileName)

'create zero byte file using the network share (overwrite if exists)
Set oTextFile = m_oFSO.CreateTextFile(sTempFilePath_UNC, True)
oTextFile.Close

'attempt to read header info from this file as if it is a SQL Server backup.
'This will generate one of two errors:
'(1) If the file can be located from the server:    #3254: The volume on device <path> is empty"
'(2) If the file cannot be located from the server: #3201: Cannot open backup device...
'                                                          The system cannot find the file specified.
sTSQL = "RESTORE HEADERONLY" & vbCrLf _
      & "    FROM DISK = '" & TSQLQuoteFix(sTempFilePath_LocalToSvr) & "'"
'save old error mode and turn of logging of this error
sLogErrors_OldFlag = CStr(m_bLogErrors)
m_bLogErrors = False
'execute the TSQL
On Error Resume Next
ExecTSQL dbSvr, sTSQL, False
ErrClone1 = MakeErrClone
On Error GoTo EH
m_bLogErrors = CBool(sLogErrors_OldFlag)
sLogErrors_OldFlag = ""

'delete the temp file
m_oFSO.Deletefile sTempFilePath_UNC, True

'test if either expected error message is returned
For lIndex = 0 To DBEngine.Errors.Count - 1
   Select Case DBEngine.Errors(lIndex).Number
       Case 3254  'empty backup file found
           bEmptyFileError = True
       Case 3201  'backup file not exist
           bFileNotExistError = True
       Case Else  'something else
   End Select
Next lIndex

'if no empty file error
If Not bEmptyFileError Then
    'if didn't get expected file not found error
    If Not bFileNotExistError Then
        sMsg = "An unexpected error occurred when attempting to confirm that folder" _
              & "'" & sBaseFolder_LocalToSvr & "' and '" & sBaseFolder_ShareName & "'" _
              & " are the same on the PC named '" & sSvrPCName & "'." & vbCrLf & vbCrLf _
              & ErrClone1.Number & "#:  " & ErrClone1.Description
        Err.Raise 1001, "SQLToolkit.ValidateSvrNWShare", sMsg
    Else
        sMsg = "The application is unable to confirm that the" _
              & " following folders on machine '" & sSvrPCName & "' are the same.  Please check the" _
              & " network shares and network communication." & vbCrLf & vbCrLf _
              & "Local path:  '" & sBaseFolder_LocalToSvr & "'" & vbCrLf _
              & "Network share:  '" & sBaseFolder_ShareName & "'"
        Err.Raise 1001, "SQLToolkit.ValidateSvrNWShare", sMsg
    End If
End If 'else, got expected result

'return success
ValidateSvrNWShare = True

Exit Function
EH:
'restore log errors flag for the module
If sLogErrors_OldFlag <> "" Then
   m_bLogErrors = CBool(sLogErrors_OldFlag)
End If
ErrHandlerPrivate False, "SQLToolkit", "ValidateSvrNWShare"
End Function


'Usage/Notes (Russ A. 11-29-04):
'
'Summary of the problem:
'If restoring or attaching a database to a different server (or the same server at a later date),
'the Security IDs (SIDs) of the UserNames stored in the database file may not map to any ServerEngine LoginName SIDs
'(stored in the master database).  This situation makes the UserName information in the database non functional.
'
'The procedure below is a "first pass" attempt to correct this situation.  It will remap
'the database UserNames to matching LoginNames based on a text match of the 'name'.
'
'Usage:
'To link a specific UserName and LoginName, specify a sUserName and sLoginName.
'To link all matching UserNames and LoginNames, specify a blank sUserName.
'
'Issues:
'
'(1) If there is no matching LoginName for a UserName, then the UserName will be skipped and
'will remain 'orphaned'.  It's up to the caller to check for this and decide what to do.
'For all orphans to be mapped, the caller must add all the necessary logins *before* running
'this procedure.  (Sample code from Microsoft (article 298897) automatically adds the LoginName
'with a Null password, but this opens a serious security risk and has been left out.)
'
'(2) For the dbo alias, the UserName/LoginName is *not* stored in the database--only the SID.  If the SID
'doesn't match an existing login, there's no way to determine the original LoginName for the dbo.
'To allow the dbo alias to work, it must be mapped to some known login.  The
'code below maps the dbo to 'sa' as the 'sa' login is always present.  I don't know if this has
'any negative consequences.  This technique was used in various samples on the web.

'(3) This procedure does NOT attempt to fix Windows authentication users.  The following
'discuss ways to deal with Windows authentication issues with a restored database:
'MSDN article 240872 "HOW TO: Resolve Permissions Issues When You Move A Database"
'                     Between Servers That Are Running SQL Server"
'MSDN article 298897:"SAMPLE: Mapsids.exe Helps Map SIDs Between User and Master"
'                     Databases When Database Is Moved"
'SQL Svr Magazine Sept 2002, InstantDoc #25983, http://www.winnetmag.com/SQLServer/Articles/ArticleID/25983/pg/2/2.html
'                     "Safe Transit" By Kalen Delaney
'
'(4) More sophisticated approaches may exist that may be worth pursing in the future.  The fundamental
'problem with SQL Server is that the
'database files do NOT store the SQL Server login information.  This might be solved by 'manually'
'saving the SQL Server logins in the database.  The following page discusses a technique that
'might be adapted for doing this:  "Creating Logins For Databases Restored To A Replacement SQL Server" at
'http://www.sql-server-performance.com/dd_creating_logins.asp
'For example, before doing the backup/detach, save all relevant user logins, encrypted passwords, and SIDs in a custom table
'in the database being backed up/detached.  Then, do the backup/detach.  Later, restore/attach the database. Then use
'the stored login information to "fix up" any missing logins.  There are some issues with avoiding modification of
'existing Logins (that may already be in use in other existing databases)--if not done carefully, the changes could
'have effects on existing databases.  Time constraints prevented pursing this method.
'
'(5) NOTE: Prior to V1.15, the TSQL used a "USE <db name>" statement at the start and "USE <master>" at end and
'executed the ExecTSQL with "master".  When run against a newly created database (or newly restored database),
'this routinely caused error #3146--ODBC call failed and
''#615 "Could not find database table ID <number>, name '<name of database>'".
'Removing the "USE" statements and setting the database context in the ExecTSQL call resolves this.  Source of problem unclear.
'There may be notes in older code re: executing trivial TSQL against the new database before running this code.  With
'removal of 'USE' statements, this is no longer necessary.

Public Function MapDBUsersToSvrLogins(dbSvr As dao.Database, ByVal sDBName As String, _
                                      Optional ByVal sUserName As String = "", _
                                      Optional ByVal sLoginName As String = "")
On Error GoTo EH
Dim sTSQL As String

If sUserName = "" Then

    'this script adapted from http://www.sqlteam.com/forums/topic.asp?TOPIC_ID=12615.  it contains a fix for an issue
    'with the 'dbo' user.  what apears to be an earlier version without this fix is located at:
    'http://www.sql-server-performance.com/sql_server_maintenance.asp.
    'I added (a) indentation and (b) the LEFT JOIN to the syslogins table (c) a test for whether or not a matching
    'LoginName actually exists and (d) a fix for a strange problem with the IMPLICIT_TRANSACTIONS flag when executing
    'in Query Analyzer in debug mode.
    sTSQL = ""
    sTSQL = sTSQL & "BEGIN" & vbCrLf
    sTSQL = sTSQL & "    DECLARE @implicit_flag integer" & vbCrLf
    sTSQL = sTSQL & "    DECLARE @username SYSNAME" & vbCrLf
    sTSQL = sTSQL & "    DECLARE @loginname SYSNAME" & vbCrLf
    sTSQL = sTSQL & "    DECLARE fixusers CURSOR" & vbCrLf
    sTSQL = sTSQL & "        FOR" & vbCrLf
    sTSQL = sTSQL & "            SELECT T1.name AS UserName, T2.name AS LoginName" & vbCrLf
    sTSQL = sTSQL & "            FROM sysusers AS T1" & vbCrLf
    sTSQL = sTSQL & "            -- REA:  I added this join to the syslogins table (similar to TSQL in MSDN article 305711) so we" & vbCrLf
    sTSQL = sTSQL & "            -- can determine below if a matching LoginName exists or not.  I didn't like the idea of executing" & vbCrLf
    sTSQL = sTSQL & "            -- 'sp_change_users_login' when it was sure to error." & vbCrLf
    sTSQL = sTSQL & "            LEFT JOIN master.dbo.syslogins AS T2" & vbCrLf
    sTSQL = sTSQL & "            ON T1.name = T2.name" & vbCrLf
    sTSQL = sTSQL & "            -- 'isSQLuser = 1':  makes sure this is a user (and not a role)" & vbCrLf
    sTSQL = sTSQL & "            -- 'sid is not null': makes sure this is not a system user (e.g., 'public')" & vbCrLf
    sTSQL = sTSQL & "            -- 'sid <> 0':  again, makes sure this is a real user" & vbCrLf
    sTSQL = sTSQL & "            -- 'suser_sname(sid) is null':  makes sure there is no matching login for this user" & vbCrLf
    sTSQL = sTSQL & "            WHERE (T1.isSQLuser = 1) AND (T1.sid IS NOT NULL) AND (T1.sid <> 0x0)" & vbCrLf
    sTSQL = sTSQL & "                AND (suser_sname(T1.sid) IS NULL)" & vbCrLf
    sTSQL = sTSQL & "            ORDER BY T1.name" & vbCrLf
    sTSQL = sTSQL & "    -- REA:  save the implicit transaction state" & vbCrLf
    sTSQL = sTSQL & "    SELECT @implicit_flag = (@@OPTIONS & 2)" & vbCrLf
    sTSQL = sTSQL & "    -- REA:  turn off implicit transactions (else 'OPEN', 'FETCH' will start" & vbCrLf
    sTSQL = sTSQL & "    -- a transaction and will cause sp_change_users_login to fail.  For" & vbCrLf
    sTSQL = sTSQL & "    -- unknown reasons, Query Analyzer overrides the default value of 'OFF')" & vbCrLf
    sTSQL = sTSQL & "    IF @implicit_flag <> 0 SET IMPLICIT_TRANSACTIONS OFF   " & vbCrLf
    sTSQL = sTSQL & "    OPEN fixusers" & vbCrLf
    sTSQL = sTSQL & "    FETCH NEXT FROM fixusers INTO @username, @loginname" & vbCrLf
    sTSQL = sTSQL & "    WHILE @@FETCH_STATUS = 0" & vbCrLf
    sTSQL = sTSQL & "        BEGIN" & vbCrLf
    sTSQL = sTSQL & "            IF @username='dbo'" & vbCrLf
    sTSQL = sTSQL & "                BEGIN" & vbCrLf
    sTSQL = sTSQL & "                    -- rea:  the 'dbo' alias can end up linked to a SID that no longer matches any login SIDs." & vbCrLf
    sTSQL = sTSQL & "                    -- we can't backtrack to figure out what the name of the original login name was so link 'dbo' to the" & vbCrLf
    sTSQL = sTSQL & "                    -- only LoginName we can be sure exists = 'sa'.  sp_change_users_login fails if use 'sa'." & vbCrLf
    sTSQL = sTSQL & "                    EXEC sp_changedbowner 'sa'" & vbCrLf
    sTSQL = sTSQL & "                END" & vbCrLf
    sTSQL = sTSQL & "            ELSE" & vbCrLf
    sTSQL = sTSQL & "                -- rea:  if there really exists a matching loginname" & vbCrLf
    sTSQL = sTSQL & "                IF @username = @loginname " & vbCrLf
    sTSQL = sTSQL & "                    BEGIN" & vbCrLf
    sTSQL = sTSQL & "                        -- Sync the SID of the UserName to that of the LoginName" & vbCrLf
    sTSQL = sTSQL & "                        EXEC sp_change_users_login 'update_one', @username, @loginname " & vbCrLf
    sTSQL = sTSQL & "                    END" & vbCrLf
    sTSQL = sTSQL & "            FETCH NEXT FROM fixusers INTO @username, @loginname" & vbCrLf
    sTSQL = sTSQL & "        END" & vbCrLf
    sTSQL = sTSQL & "    CLOSE fixusers" & vbCrLf
    sTSQL = sTSQL & "    DEALLOCATE fixusers" & vbCrLf
    sTSQL = sTSQL & "    -- REA: reset state of implicit transactions if nec" & vbCrLf
    sTSQL = sTSQL & "    IF @implicit_flag <> 0 SET IMPLICIT_TRANSACTIONS ON " & vbCrLf
    sTSQL = sTSQL & "END" & vbCrLf

Else

    'only update the single specified user.
    
    'handle embedded single quotes (vars passed byval). sUserName must exist in the database
    'and sLoginName must exist in the SQL Server instance or the call will fail.
    sUserName = TSQLQuoteFix(sUserName)
    If sLoginName = "" Then
        sLoginName = sUserName
    Else
        sLoginName = TSQLQuoteFix(sLoginName)
    End If
    
    'if special case of changing the database owner
    If UCase(sUserName) = "DBO" Then
        'Note:  This will fail if the sLoginName is already a user in this database.  Ideally, the existing
        'user should be dropped first.  Currently, it's left up the caller to test for this before executing.
        sTSQL = ""
        sTSQL = sTSQL & "EXEC sp_changedbowner '" & sLoginName & " '" & vbCrLf
    Else
        'otherwise, ordinary user change
        sTSQL = ""
        sTSQL = sTSQL & "EXEC sp_change_users_login 'Update_one','" & sUserName & "','" & sLoginName & "'" & vbCrLf
    End If
    
End If

'V1.15:  Changed to set the database context here (rather than in the TSQL).  Resolves "database id not found" errors
'with new databases.
ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "MapDBUsersToSvrLogins"
End Function

'replaces all single apostrophes with two single apostrophes
'--this allows TSQL to handle literal strings that contain single or double quotes
Public Function TSQLQuoteFix(ByVal sSourceString As String) As String
TSQLQuoteFix = Replace(sSourceString, "'", "''")
End Function

Public Function DAOQuoteFix(ByVal sSourceString As String, Optional ByVal sQuoteDelimiter As String = "'") As String
DAOQuoteFix = Replace(sSourceString, sQuoteDelimiter, sQuoteDelimiter & sQuoteDelimiter)
End Function

'Following only intended for use with SQL Server DSN's (not tested with other cases)
'Returns True if succeeds.
'If DSN does not exist, it is created.  If exists, it is deleted and recreated.
'Based on documentation, DSN's connect faster if don't provide a default db name.  So, probably best to leave off
'the database name.
Public Function CreateOrUpdateSystemDSN(sDSN As String, _
                                        sSQLSvrName As String, _
                                        Optional sDescription As String = "", _
                                        Optional sDriver As String = "SQL Server", _
                                        Optional sDefaultDBName As String = "") As Boolean
On Error GoTo EH
Dim sAttributes As String
Dim intRet As Long

'set default return => failure
CreateOrUpdateSystemDSN = False

'Set the attributes delimited by null.
'"See driver documentation for a complete list of supported attributes."
sAttributes = "SERVER=" & sSQLSvrName & Chr$(0) _
            & "DESCRIPTION=" & sDescription & Chr$(0) _
            & "DSN=" & sDSN & Chr$(0) _
            & "DATABASE=" & sDefaultDBName & Chr$(0)
                                      
'(use of 0& suppresses user dialog if information is missing)
intRet = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, _
                             sDriver, sAttributes)
If intRet Then
    CreateOrUpdateSystemDSN = True
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "CreateOrUpdateSystemDSN"
End Function

'returns true if succeeds
Public Function DeleteSystemDSN(sDSN As String, Optional sDriver As String = "SQL Server") As Boolean
On Error GoTo EH
Dim sAttributes As String
Dim intRet As Integer

'set default return => failure
DeleteSystemDSN = False

'Set the attributes delimited by null.
'See driver documentation for a complete list of attributes.
sAttributes = "DSN=" & sDSN & Chr$(0)
'(use of 0& prevents suppresses user dialog if information is missing)
intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_SYS_DSN, _
          sDriver, sAttributes)
         
If intRet Then
    DeleteSystemDSN = True
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "DeleteSystemDSN"
End Function

'returns list of all system DSNs as zero based arrays in variant.  Function return is count of system DSNs.
'Can optionally provide a name filter to get data for DSNs with specific name.
'adapted from http://www.vb-helper.com/howto_list_dsns.html.
'==================
'Note:  VB6 also comes with a tool named ODBCTOOL.DLL that has a .GetDataSourceList function.  However,
'this seems to list everything in the ODBC.ini registry entry--including drivers, etc.
'Dim odbcTool As odbcTool.Dsn
'Dim DSNList() As String
'Dim vDSN As Variant
'Set odbcTool = New odbcTool.Dsn
'
'If odbcTool.GetDataSourceList(DSNList()) Then
'    cboExistingDSNs.Clear
'    For Each vDSN In DSNList
'        cboExistingDSNs.AddItem vDSN
'    Next vDSN
'Else
'    MBox_OnTop "Unable to read DSN list"
'End If
'==================
'(Note: sNameLIKEFilter results in case insenstive filtering of the results)
Public Function GetSystemDSNList(Optional ByVal sNameLIKEFilter As String = "*", _
                                 Optional vDSNNameList_Rtn As Variant, _
                                 Optional vDescriptionList_Rtn As Variant, _
                                 Optional vDriverList_Rtn As Variant, _
                                 Optional vServerList_Rtn As Variant, _
                                 Optional vDatabaseList_Rtn As Variant _
                                 ) As Long
On Error GoTo EH
Dim sNames() As String
Dim sServers() As String
Dim sDescriptions() As String
Dim sDrivers() As String
Dim sDatabases() As String
Dim lDSNCount As Long
Dim hKey As Long
Dim value As String
Dim length As Long
Dim value_type As Long
Dim key_num As Long
Dim subkey_name As String
Dim subkey_length As Long
Dim file_time As FILETIME
Dim lReturn As Long

'set default return value => none found
GetSystemDSNList = 0

ReDim sNames(20), sServers(20), sDescriptions(20), sDrivers(20), sDatabases(20)

'V1.15:  changed from KEY_ALL_ACCESS to KEY_READ due to issues with power users on Win2003Server
' Open HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI.
If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
    "Software\ODBC\ODBC.INI", _
    0&, KEY_READ, hKey) <> ERROR_SUCCESS _
Then
    GoTo Proc_Exit
End If

' Enumerate the subkeys, skipping "ODBC Data Sources."
key_num = 0
lDSNCount = 0
Do
    ' Get the subkey's information.
    subkey_length = 256
    subkey_name = Space$(subkey_length)
    lReturn = RegEnumKeyEx( _
        hKey, key_num, subkey_name, subkey_length, _
        ByVal 0&, ByVal 0&, 0, file_time)
    'Note:  had to add some additional error checks to this call. REA
    If (lReturn = ERROR_NO_MORE_ITEMS) Or ((lReturn <> ERROR_SUCCESS) And (lReturn <> ERROR_MORE_DATA)) Then Exit Do
    subkey_name = Left$(subkey_name, subkey_length)

    ' Display the information.
    'Note:  added exclusion of File DSN as an item.
    If (subkey_name <> "ODBC Data Sources") And (subkey_name <> "ODBC File DSN") And (UCase(subkey_name) Like UCase(sNameLIKEFilter)) Then
    
        lDSNCount = lDSNCount + 1
        If (lDSNCount - 1) > UBound(sNames) Then
            ReDim Preserve sNames(20 + lDSNCount), sServers(20 + lDSNCount), sDescriptions(20 + lDSNCount), _
                           sDrivers(20 + lDSNCount), sDatabases(20 + lDSNCount)
        End If
        
        'add the new DSN to the list
        sNames(lDSNCount - 1) = subkey_name
        
        sDescriptions(lDSNCount - 1) = GetKeyValue(HKEY_LOCAL_MACHINE, _
                "Software\ODBC\ODBC.INI\" & _
                    subkey_name, _
                "Description")
                
        sDrivers(lDSNCount - 1) = GetKeyValue(HKEY_LOCAL_MACHINE, _
                "Software\ODBC\ODBC.INI\" & _
                    subkey_name, _
                "Driver")
                
        sServers(lDSNCount - 1) = GetKeyValue(HKEY_LOCAL_MACHINE, _
                "Software\ODBC\ODBC.INI\" & _
                    subkey_name, _
                "Server")
                                
        sDatabases(lDSNCount - 1) = GetKeyValue(HKEY_LOCAL_MACHINE, _
                "Software\ODBC\ODBC.INI\" & _
                    subkey_name, _
                "Database")
                                                        
    End If

    key_num = key_num + 1
Loop

' Close the key.
If RegCloseKey(hKey) <> ERROR_SUCCESS Then
    GoTo Proc_Exit
End If

Proc_Exit:

'if no elements found
If lDSNCount = 0 Then
    'return arrays that have a UBound of -1
    vDSNNameList_Rtn = Array()
    vDescriptionList_Rtn = Array()
    vDriverList_Rtn = Array()
    vServerList_Rtn = Array()
    vDatabaseList_Rtn = Array()
Else
    'reduce array size to minimum
    ReDim Preserve sNames(lDSNCount - 1)
    ReDim Preserve sDescriptions(lDSNCount - 1)
    ReDim Preserve sDrivers(lDSNCount - 1)
    ReDim Preserve sServers(lDSNCount - 1)
    ReDim Preserve sDatabases(lDSNCount - 1)
    vDSNNameList_Rtn = sNames
    vDescriptionList_Rtn = sDescriptions
    vDriverList_Rtn = sDrivers
    vServerList_Rtn = sServers
    vDatabaseList_Rtn = sDatabases
End If

GetSystemDSNList = lDSNCount

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetSystemDSNList"
End Function

' Return a subkey's value as string
Private Function GetKeyValue(ByVal key_root As Long, ByVal _
                             key_name As String, ByVal subkey_name As String) As String
Dim hKey As Long
Dim value As String
Dim length As Long
Dim value_type As Long
Dim key_num As Long
Dim file_time As FILETIME

    ' Open the key.
    If RegOpenKeyEx(key_root, key_name, _
            0&, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS _
        Then Exit Function

    ' Get the subkey's value.
    length = 256
    value = Space$(length)
    If RegQueryValueEx(hKey, subkey_name, _
        0&, value_type, ByVal value, length) _
            <> ERROR_SUCCESS _
        Then Exit Function

    ' Remove the trailing null character.
    GetKeyValue = Left$(value, length - 1)

    ' Close the key.
    RegCloseKey hKey
End Function

'return key value as long
Private Function GetKeyValueLong(ByVal key_root As Long, ByVal _
                                 key_name As String, ByVal subkey_name As String) As Long
Dim hKey As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

' Open the key.
If RegOpenKeyEx(key_root, key_name, _
                0&, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS _
Then Exit Function

lrc = RegQueryValueExNULL(hKey, subkey_name, 0&, lType, 0&, cch)
If lrc <> ERROR_NONE Then Exit Function

lrc = RegQueryValueExLong(hKey, subkey_name, 0&, lType, lValue, cch)
If lrc = ERROR_NONE Then GetKeyValueLong = lValue

' Close the key.
RegCloseKey hKey

End Function

Private Function SetKeyValueLong(ByVal key_root As Long, ByVal key_name As String, ByVal subkey_name As String, _
                            ByVal lNewValue As Long) As Boolean
Dim lRetVal As Long         'result of the SetValueExLong function
Dim hKey As Long            'handle of open key

lRetVal = RegOpenKeyEx(key_root, key_name, 0, _
                       KEY_ALL_ACCESS, hKey)
lRetVal = RegSetValueExLong(hKey, subkey_name, 0&, _
                       REG_DWORD, lNewValue, 4)
RegCloseKey hKey
SetKeyValueLong = (lRetVal = ERROR_NONE)
End Function

Public Function GetODBCConnTimeoutSec_Jet35()
GetODBCConnTimeoutSec_Jet35 = GetKeyValueLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Jet\3.5\Engines\ODBC", _
                                "ConnectionTimeout")
End Function

Public Function GetODBCConnTimeoutSec_Jet40()
GetODBCConnTimeoutSec_Jet40 = GetKeyValueLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Jet\4.0\Engines\ODBC", _
                                "ConnectionTimeout")
End Function

'NOTE:  THIS REGISTRY CHANGE IS NOT EFFECTIVE UNTIL APPLICATION IS RESTARTED
Public Function SetODBCConnTimeoutSec_Jet35(ByVal lNewTimeoutSec As Long) As Boolean
'don't allow setting to zero or connections will never be closed
If lNewTimeoutSec < 1 Then
   lNewTimeoutSec = 1
End If

SetODBCConnTimeoutSec_Jet35 = SetKeyValueLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Jet\3.5\Engines\ODBC", _
                                              "ConnectionTimeout", lNewTimeoutSec)
End Function

'NOTE:  THIS REGISTRY CHANGE IS NOT EFFECTIVE UNTIL APPLICATION IS RESTARTED
Public Function SetODBCConnTimeoutSec_Jet40(ByVal lNewTimeoutSec As Long) As Boolean
'don't allow setting to zero or connections will never be closed
If lNewTimeoutSec < 1 Then
   lNewTimeoutSec = 1
End If

SetODBCConnTimeoutSec_Jet40 = SetKeyValueLong(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Jet\4.0\Engines\ODBC", _
                                              "ConnectionTimeout", lNewTimeoutSec)
End Function

'NOTE:  1.  This only returns the ***first*** database in the backup set.  Backup sets can actually contain many individual
'backups.
'       2.  NON-UNC FILE PATHS ARE ALWAYS INTERPRETED RELATIVE TO THE SERVER'S PC.
Public Sub GetBackupFileInfo_Header(dbSvr As dao.Database, _
                             ByVal sBackupFilePath_RelSvr As String, _
                             Optional ByVal sBkupPassword As String = "", _
                             Optional sBackupDBName_RTN As String, _
                             Optional dBackupDate_RTN As Date, _
                             Optional sBackupServerName_RTN As String, _
                             Optional sBackupPCName_RTN As String, _
                             Optional sBackupDescription_RTN As String, _
                             Optional sBackupDBCmptLevel_RTN As String)
On Error GoTo EH
Dim sTSQL_Temp As String
Dim rsTemp As dao.Recordset
Dim ErrClone1 As ErrClone

'set default returns
sBackupDBName_RTN = ""
dBackupDate_RTN = CDate(0)
sBackupServerName_RTN = ""
sBackupPCName_RTN = ""
sBackupDescription_RTN = ""
sBackupDBCmptLevel_RTN = ""

'get the database name(s) (there can be more than one, but we're only interested in the first one)
sTSQL_Temp = "RESTORE HEADERONLY" & vbCrLf _
           & "   FROM DISK = '" & TSQLQuoteFix(sBackupFilePath_RelSvr) & "'"
           
If sBkupPassword <> "" Then
     sTSQL_Temp = sTSQL_Temp & vbCrLf _
                  & "   WITH PASSWORD='" & TSQLQuoteFix(sBkupPassword) & "'"
End If

Set rsTemp = ExecTSQL(dbSvr, sTSQL_Temp, True, "master")
If Not rsTemp.EOF Then
    
    '(there are many more values that can be retrieved, but these appear most useful)
    sBackupDBName_RTN = GetFV(rsTemp, "DatabaseName")
    dBackupDate_RTN = GetFV(rsTemp, "BackupFinishDate")
    sBackupServerName_RTN = GetFV(rsTemp, "ServerName")
    sBackupPCName_RTN = GetFV(rsTemp, "MachineName")
    sBackupDescription_RTN = GetFV(rsTemp, "BackupDescription")
    sBackupDBCmptLevel_RTN = GetFV(rsTemp, "CompatibilityLevel")
End If

rsTemp.Close

Exit Sub
EH:
'save error state
ErrClone1 = MakeErrClone
'if "ODBC Call failed error"
If ErrClone1.Number = 3146 Then
   'if 'Too many backup devices specified for backup or restore; only 64 are allowed.' error
   If DBEngineErrorExists(3205) Then
       'add additional information.  the raw error description doesn't indicate the actual problem.
       ErrClone1.Description = ErrClone1.Description & vbCrLf & vbCrLf _
            & "THIS ERROR MAY OCCUR IF THE SERVER VERSION IS OLDER THAN THE BACKUP VERSION."
   End If
End If
RestoreErrObject ErrClone1
'propagate errors to caller
ErrHandlerPrivate False, "SQLToolkit", "GetBackupFileInfo_Header"
End Sub

'NOTE:  NON-UNC FILE PATHS ARE ALWAYS INTERPRETED RELATIVE TO THE SERVER'S PC.
Public Sub GetBackupFileInfo_FileList(dbSvr As dao.Database, _
                             ByVal sBackupFilePath_RelSvr As String, _
                             Optional ByVal sBkupPassword As String = "", _
                             Optional sBackupMDFPath_RTN As String, _
                             Optional sBackupMDFLogicalName_RTN As String, _
                             Optional sBackupLDFPath_RTN As String, _
                             Optional sBackupLDFLogicalName_RTN As String)
On Error GoTo EH
Dim sTSQL_Temp As String
Dim rsTemp As dao.Recordset
Dim ErrClone1 As ErrClone

'set default returns
sBackupMDFPath_RTN = ""
sBackupMDFLogicalName_RTN = ""
sBackupLDFPath_RTN = ""
sBackupLDFLogicalName_RTN = ""

sTSQL_Temp = "RESTORE FILELISTONLY" & vbCrLf _
           & "   FROM DISK = '" & TSQLQuoteFix(sBackupFilePath_RelSvr) & "'"
           
If sBkupPassword <> "" Then
     sTSQL_Temp = sTSQL_Temp & vbCrLf _
                  & "   WITH PASSWORD='" & TSQLQuoteFix(sBkupPassword) & "'"
End If

Set rsTemp = ExecTSQL(dbSvr, sTSQL_Temp, True, "master")

Do While Not rsTemp.EOF
    Select Case rsTemp.Fields("Type").value
        Case "D"
            sBackupMDFPath_RTN = GetFV(rsTemp, "PhysicalName")
            sBackupMDFLogicalName_RTN = GetFV(rsTemp, "LogicalName")
        Case "L"
            sBackupLDFPath_RTN = GetFV(rsTemp, "PhysicalName")
            sBackupLDFLogicalName_RTN = GetFV(rsTemp, "LogicalName")
        Case Else
            'we're not prepared to handle other file types (e.g., file groups or "N")--just skip these
    End Select
    rsTemp.MoveNext
Loop

rsTemp.Close
Exit Sub
EH:
'save error state
ErrClone1 = MakeErrClone
'if "ODBC Call failed error"
If ErrClone1.Number = 3146 Then
   'if 'Too many backup devices specified for backup or restore; only 64 are allowed.' error
   If DBEngineErrorExists(3205) Then
       'add additional information.  the raw error description doesn't indicate the actual problem.
       ErrClone1.Description = ErrClone1.Description & vbCrLf & vbCrLf _
            & "THIS ERROR MAY OCCUR IF THE SERVER VERSION IS OLDER THAN THE BACKUP VERSION."
   End If
End If
RestoreErrObject ErrClone1
'propagate errors to caller
ErrHandlerPrivate False, "SQLToolkit", "GetBackupFileInfo_FileList"
End Sub

'Returns true if specified error number currently exists in the DBEngine.Errors collection
Public Function DBEngineErrorExists(lDBEngineErrNumber As Long) As Boolean
On Error Resume Next
Dim lIndex As Long
DBEngineErrorExists = False
For lIndex = 0 To DBEngine.Errors.Count - 1
    If DBEngine.Errors(lIndex).Number = lDBEngineErrNumber Then
        DBEngineErrorExists = True
        Exit Function
    End If
Next lIndex
Err.Clear
End Function

Public Function GetLocalPCName() As String
Dim sComputerName As String * 255
GetComputerName sComputerName, 255 - 1
GetLocalPCName = StripAtFirstNull(sComputerName)
'GetLocalPCName = "w16830client7"

End Function

Private Function StripAtFirstNull(sNullTerminated As String) As String
Dim lNullPos As Long
lNullPos = InStr(sNullTerminated, Chr$(0))
If lNullPos > 0 Then
   StripAtFirstNull = Left$(sNullTerminated, lNullPos - 1)  'strip off terminating null
Else
   StripAtFirstNull = sNullTerminated
End If
End Function

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

'function checks if the second millisec reading is greater than the first
'(timerGetTime wraps around to zero every 49 days).
'if not, then assume that the counter has wrapped around zero.
'return the adjusted second counter value in this case.
'method should work as long as the time difference between curMSecEarly and curMSecLater is
' < 49 days.
Private Sub AdjustForTimerWrap(curMSecLater As Currency, curMSecEarly As Currency)
If curMSecLater < curMSecEarly Then
   'then assume that the counter has looped around zero.
   'Add 2^32 (@is currency datatype) to the "later" counter to correct for this.
   curMSecLater = curMSecLater + 4294967296@
End If
'adjusted curMSecLater is returned by reference
End Sub

'returns value of timeGetTime as a positive currency value (see preceeding fns for documentation)
Private Function timeGetTimeAsCurrency() As Currency
timeGetTimeAsCurrency = LongToUnsignedCurrency(timeGetTime())
End Function

'returns the number of millisec that have elapsed since a previous timestamp.  Note that this
'function takes a raw value from timeGetTime as the input and does all necessary
'neg/wrap adjustments for timeGetTime.  Will still fail (by rolling over to zero)
'if elapsed time exceeds (2^32 -1) milliseconds = 49 days.
Private Function GetElapsedMs(ByVal lStart_RawTimeGetTimeMs As Long) As Currency
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

Private Sub WaitXMsec(lWaitMsec As Long, Optional bWithDoEvents As Boolean = True)
On Error GoTo EH
'function provides delay for lWaitMSec.  Based on documentation in Access97 Developer's Handbook,
'timerGetTime is accurate to within 1 millisec and executes more quickly than GetTickCount() and
'should be available on any Win95/NT4 PC.
Dim curMSecStart As Currency
Dim curMSecEnd As Currency
Dim curMSecNow As Currency
curMSecStart = timeGetTimeAsCurrency()
curMSecEnd = curMSecStart + CCur(lWaitMsec)
Do
   If bWithDoEvents Then
       'let windows do processing
       DoEvents
   End If
   'get the time for current cycle (should be a "later" value than curMSecStart)
   curMSecNow = timeGetTimeAsCurrency()
   'check if counter has wrapped around to 0 and adjust curMSecNow if necessary
   AdjustForTimerWrap curMSecNow, curMSecStart
Loop While curMSecNow < curMSecEnd
Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "WaitXMSec"
End Sub

'formats ODBC connection string for user display.
Public Function FormatODBC(ByVal sODBCConnStr As String, Optional bHidePassword As Boolean = True) As String
'No error handler here as this is called from error handlers.  errors should be unlikely.
Dim sReturn As String

If bHidePassword Then
    sReturn = HidePWDInConnStr(sODBCConnStr)
Else
    sReturn = sODBCConnStr
End If

sReturn = Replace(sReturn, ";", ";" & vbCrLf)

'if ending CRLF
If Right(sReturn, 2) = vbCrLf Then
    'strip it off
    sReturn = Left(sReturn, Len(sReturn) - 2)
End If

FormatODBC = sReturn

End Function

'Can be used to hide the PWD value of a connection string (or any other argument)
'If connection string is badly formed, this could still return a connection string
'with a password, so the resulting values should still only be logged to file.
Public Function HidePWDInConnStr(ByVal sConnStr As String) As String
'No error handler here as this is called from error handlers (based on testing, errors
'should not be possible here)
Dim lLenPrefix As Long
Dim lPosPWD As Long
Dim lPosSC As Long
Dim sLeft As String
Dim sRight As String

lLenPrefix = Len("PWD=")
lPosPWD = InStr(1, sConnStr, "PWD=", vbTextCompare)

'if no prefix found
If lPosPWD < 1 Then
   'assume the connection string is not well formed and return blank string
   Exit Function
End If

Do While lPosPWD > 0
    'find ending semicolon
    lPosSC = InStr(lPosPWD, sConnStr, ";", vbTextCompare)
    'if not found
    If lPosSC < 1 Then
        'use end of string plus 1
        lPosSC = Len(sConnStr) + 1
    End If
    'splice out the password
    sLeft = Left(sConnStr, lPosPWD + lLenPrefix - 1)
    sRight = Mid(sConnStr, lPosSC)
    'replace with 5 asterisks
    sConnStr = sLeft & "*****" & sRight
    'test if any others
    lPosPWD = InStr(lPosPWD + lLenPrefix + 5, sConnStr, "PWD=", vbTextCompare)
Loop

'return resulting string
HidePWDInConnStr = sConnStr

End Function

Private Function BuildDBEngineErrStr(Optional bRemoveBracketedInfo As Boolean = False) As String
Dim sDBErr As String
Dim lErrIndex As Long

'assemble a DAO errors collection string (includes DBEngine/ODBC/SQL Server errors)
For lErrIndex = 0 To DBEngine.Errors.Count - 1
    If bRemoveBracketedInfo Then
        sDBErr = IIf(sDBErr = "", "", sDBErr & vbCrLf) _
                 & "#" & DBEngine.Errors(lErrIndex).Number & ":  " & ReplaceBracketedInfo(DBEngine.Errors(lErrIndex).Description)
    Else
        sDBErr = IIf(sDBErr = "", "", sDBErr & vbCrLf) _
                 & "#" & DBEngine.Errors(lErrIndex).Number & ":  " & DBEngine.Errors(lErrIndex).Description
    End If
Next lErrIndex

'return result or blank
BuildDBEngineErrStr = sDBErr

End Function

'Only called by preceding routine to strip bracketed info from the DBEngine error string
'V1.18:  improved code so won't lose data if more than one set of brackets (never actually occurs with
'DBEngine errors, but does occur with VB errors)
Public Function ReplaceBracketedInfo(ByVal sDBEngineErrStr As String, _
                                    Optional ByVal sReplaceWith As String = "", _
                                    Optional ByVal lThruPosNum As Long = 0) As String
Dim lPosLeft As Long
Dim lPosRight As Long
Dim lLenReplace As Long

'set default return
ReplaceBracketedInfo = sDBEngineErrStr

If (lThruPosNum < 1) Or (lThruPosNum > Len(sDBEngineErrStr)) Then
    lThruPosNum = Len(sDBEngineErrStr)
End If

lLenReplace = Len(sReplaceWith)

lPosLeft = 1
Do
    'get position of next starting left bracket
    lPosLeft = InStr(lPosLeft, sDBEngineErrStr, "[", vbTextCompare)
    'if no additional left bracket
    If (lPosLeft < 1) Then
        'exit loop
        Exit Do
    End If
    'get pos of right bracket just beyond the left one
    lPosRight = InStr(lPosLeft + 1, sDBEngineErrStr, "]", vbTextCompare)
    'if no additional right brackets Or if beyon the thru pos
    If (lPosRight < 1) Or (lPosRight > lThruPosNum) Then
        'exit loop
        Exit Do
    End If
    'delete the brackets and everything in between
    sDBEngineErrStr = Left(sDBEngineErrStr, lPosLeft - 1) & sReplaceWith & Mid(sDBEngineErrStr, lPosRight + 1)
    'update the 'through' position to reflect change in length of the original string.
    'due to tests above, know that the entire replacement ends at or before the thru position.
    ' = (old thru position) + (  (size of inserted string) - (size of the bracketed string)  )
    lThruPosNum = lThruPosNum + lLenReplace - (lPosRight - lPosLeft + 1)
    'get ready for next cycle
    lPosLeft = lPosLeft + lLenReplace
Loop

ReplaceBracketedInfo = sDBEngineErrStr

End Function

'Put error handling fn here so class is more self-sufficient.
Private Sub ErrHandlerPrivate(bHandleError As Boolean, sModuleName As String, sProcedureName As String, _
                              Optional sOtherAdvErrInfo As String = "")
 
Dim lErrNumber As Long
Dim sErrDescription As String
Dim sErrDescLog As String
Dim sErrSource As String
Dim sErrHelpFile As String
Dim sErrHelpContext As String
Dim sDBErr As String
Dim bFirstTime As Boolean
Dim bODBCError As Boolean

'get stable copy of the error information
lErrNumber = Err.Number
sErrSource = Err.Source
sErrDescription = Err.Description
sErrDescLog = sErrDescription
sErrHelpFile = Err.HelpFile
sErrHelpContext = Err.HelpContext
Err.Clear

'remove any hourglass cursor
Screen.MousePointer = vbDefault

'determine if this is the first time error raised (vs. is propagation)
bFirstTime = (InStr(1, sErrDescription, "[Module.Procedure = ", vbTextCompare) < 1)
'determine if ODBC error (does case sensitive test for "ODBC" in message
bODBCError = (InStr(1, sErrDescription, "ODBC") > 0)

'if this proc was entered when there was really no error
If lErrNumber = 0 Then
    sErrDescription = "The error handling routine of " & App.EXEName & ".SQLToolkit has been entered when no error state exists." _
         & "  The available data is:" & vbCrLf & vbCrLf _
         & "Error #" & lErrNumber & ":  " & sErrDescription
    sErrDescLog = sErrDescription
End If

If bFirstTime Then

    If lErrNumber <> 0 Then
        sErrDescription = "Error #" & lErrNumber & ":  " & sErrDescription
        sErrDescLog = sErrDescription
    End If
    
    sDBErr = BuildDBEngineErrStr(True)
    
    If sDBErr <> "" Then
        
        'if ODBC error
        If bODBCError Then
    
            'append database errors to the error message
            sErrDescription = sErrDescription & vbCrLf & vbCrLf _
               & "DBEngine.Errors:" & vbCrLf _
               & sDBErr
               
        End If
          
       'also create advanced message that always includes the full DB Engine error info (even though this may include some
       'old/irrelevant error info)
       sErrDescLog = sErrDescLog & vbCrLf & vbCrLf _
          & "FULL DBEngine ERRORS LIST" & IIf(Not bODBCError, " (may include old errors)", "") & ":" & vbCrLf _
          & BuildDBEngineErrStr(False)
          
    End If
    
    'add the originating Err.Source property and the current module/procedure to the error description
    sErrDescription = sErrDescription & vbCrLf & vbCrLf _
          & "[Original Source = " & sErrSource & "]" & vbCrLf _
          & "[Module.Procedure = " & sModuleName & "." & sProcedureName & "]"
          
    'start buildup of call stack in Err.Source property
    sErrSource = "[Original Source = " & sErrSource & "]"
          
End If
    
'always append the current module and procedure name to the call stack in the Err.Source property
sErrSource = sErrSource & vbCrLf _
      & "[Module.Procedure = " & sModuleName & "." & sProcedureName & "]"
    
'if need to propagate and we have a valid error number
If (Not bHandleError) And (lErrNumber <> 0) Then
    
    'if first time and logging is desired
    If bFirstTime And m_bLogErrors And (Not (m_oAppLog Is Nothing)) Then
    
        'log error with 'propagating first time' message and advanced error format (logging here guarantees that error will
        'be logged even if never handled by original caller)
        m_oAppLog.AddEntry String(60, "-") & vbCrLf _
                         & "<" & App.EXEName & "." & "SQLToolkit.ErrHandlerPrivate is logging initial error propagation>" & vbCrLf _
                         & sErrDescLog & vbCrLf & vbCrLf _
                         & "ERROR SOURCE/CALL STACK AT THIS TIME:" & vbCrLf _
                         & sErrSource & vbCrLf _
                         & IIf(sOtherAdvErrInfo <> "", vbCrLf & "OTHER ERROR INFO:" & vbCrLf & sOtherAdvErrInfo & vbCrLf, "") _
                         & String(60, "-")

    End If  'else, don't log with subsequent propagations (just fills up the log)

    'propagate error back to client and to next higher handler
    'NOTE:  If lErrNumber is a negative value from SQLDMO, it may be translated to a more generic
    'VB error number when the error is raised here.  Reason unclear.
    Err.Raise lErrNumber, sErrSource, sErrDescription, sErrHelpFile, sErrHelpContext
        
End If

'log errors if desired
If m_bLogErrors And (Not (m_oAppLog Is Nothing)) Then
    m_oAppLog.AddEntry String(60, "-") & vbCrLf _
                     & "<" & App.EXEName & "." & "SQLToolkit.ErrHandlerPrivate is logging user error message>" & vbCrLf _
                     & sErrDescLog & vbCrLf & vbCrLf _
                     & "ERROR SOURCE/CALL STACK:" & vbCrLf _
                     & sErrSource & vbCrLf _
                     & IIf(sOtherAdvErrInfo <> "", vbCrLf & "OTHER ERROR INFO:" & vbCrLf & sOtherAdvErrInfo & vbCrLf, "") _
                     & String(60, "-")
    
End If

MBox_OnTop sErrDescription, vbOKOnly, "Error From '" & App.EXEName & "'"

'make sure no error is returned back to caller
Err.Clear
End Sub

'uses message box API to force message box to be on top of all other forms
Private Function MBox_OnTop(Prompt As String, Optional Buttons As VbMsgBoxStyle = -1, Optional Title As String = "") As VbMsgBoxResult
Dim lButtons As Long
Dim sTitle As String
If Buttons = -1 Then
   lButtons = vbOKOnly
Else
   lButtons = Buttons
End If
If Title = "" Then
   sTitle = App.Title
Else
   sTitle = Title
End If
MBox_OnTop = MessageBox(0&, Prompt, sTitle, lButtons Or MB_TASKMODAL Or MB_TOPMOST)
End Function

Private Function MakeErrClone() As ErrClone
Dim TempErrClone As ErrClone
With TempErrClone
   .Description = Err.Description
   .HelpContext = Err.HelpContext
   .HelpFile = Err.HelpFile
   .LastDllError = Err.LastDllError
   .Number = Err.Number
   .Source = Err.Source
End With
MakeErrClone = TempErrClone
End Function

Private Function RestoreErrObject(ByRef SavedErrClone As ErrClone)
With Err
   .Description = SavedErrClone.Description
   .HelpContext = SavedErrClone.HelpContext
   .HelpFile = SavedErrClone.HelpFile
   'can't restore the LastDllError property, so leave it with entry value
   '.LastDllError = SavedErrClone.LastDllError
   .Number = SavedErrClone.Number
   .Source = SavedErrClone.Source
End With
End Function

Private Function RaiseError(ByRef SavedErrClone As ErrClone)
With SavedErrClone
    Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
End With
End Function

'=============================================================================
'START PASSWORD ENCRYPTION/DESCRYPTION FUNCTIONS
'=============================================================================
' EncryptDecrypt (FUNCTION)
'
' PARAMETERS:
' (In) - sText            - String                   - Text to encrypt or decrypt
' (In) - sKeyRoot         - String                   - Text to use in generating key
' (In) - bEncrypt         - Boolean                  - True to encrypt, false to
'                                                      decrypt
' (In) - lEncryptionType - EncryptionType_Enum       - Block or stream encryption
'
' RETURN VALUE:
' String - The encrypted or decrypted string
'
' DESCRIPTION:
' Takes the input string and encrypts or decrypts it.
'REA:
'(1) Use of "Microsoft Base Cryptographic Provider v1.0" probably means that this uses
'"only" 40 bit encryption.  This is probably the safest method as there are issues
'with using 128 bit encryption with older systems (e.g., Win2000, WinNT).  See MSDN #841715
'for details on issues with 128 bit encryption when move encrypted data between Win2K and WinXPSP1.
'(2) This code is similar to examples provided by Microsoft in MSDN article #841715.
'(3) If attempt to decrypt a blank string using RC2 algorithm, will exit with message showing
'    API error -2146893820 "Bad Length".  (With RC4 algorithm, blanks encode to blanks, so no
'    error is generated)
'(4) If attempt to decrypt a string using different key than used to encrypt, RC2 will exit with
'    message showing API error -2146893819 "Bad data."  With RC4 algorithm, no error is generated--you just
'    get back nonsense data.
'(5) If attempt to decrypt a string using RC2 that has not actually been encoded with RC2, routine
'    will exit with message showing API error -2146893819 "Bad data."  RC4 is less fussy and will just
'    return nonsense data.
'*******************************************************************************
Private Function EncryptDecrypt(ByVal sText As String, _
                                ByVal sKeyRoot As String, _
                                ByVal bEncrypt As Boolean, _
                                Optional ByVal lEncryptionType As EncryptionType_Enum _
                                                        = ENCRYPTION_TYPE_RC2) As String
On Error GoTo EH
Dim lProvider           As Long
Dim lResult             As Long
Dim lHash               As Long
Dim lKey                As Long
Dim lHashPassword       As Long
Dim lFlags              As Long
Dim lData               As Long
Dim lClear              As Long

'Get a cryptographic provider context handle
lResult = CryptAcquireContext(lProvider, vbNullString, _
                     CRYPTO_PROVIDER, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
If lResult = 0 Then
     Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptAcquireContext()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
End If

' Create a handle to a hash object  using the MD5 algorithm
lResult = CryptCreateHash(lProvider, CALG_MD5, 0, 0, lHash)
If lResult = 0 Then
     Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptCreateHash()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
End If

' Add custom key text to the hash object for use in generating our key.
' sKeyRoot is, in effect, our key and should be fairly complex and
' not easily guessed
lHashPassword = Len(sKeyRoot)
lResult = CryptHashData(lHash, sKeyRoot, lHashPassword, 0)
If lResult = 0 Then
     Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptHashData()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
End If

' Select appropriate encryption method
If lEncryptionType = ENCRYPTION_TYPE_RC2 Then
    lEncryptionType = CALG_RC2
Else
    lEncryptionType = CALG_RC4
End If

' Generate a session key for use with the cypher
lFlags = CRYPT_NO_SALT
lResult = CryptDeriveKey(lProvider, lEncryptionType, lHash, lFlags, lKey)
If lResult = 0 Then
     Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptDeriveKey()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
End If

' Encrypt or decrypt data as required
lData = Len(sText)
If bEncrypt Then
    lClear = lData
    ' Call with a null first to see how long the string needs to be
    lResult = CryptEncrypt(lKey, 0, 1, 0, vbNullString, lData, lClear)
    sText = sText & String(lData - lClear, " ")
    ' Encrypt some text
    lResult = CryptEncrypt(lKey, 0, 1, 0, sText, lClear, lData)
    If lResult = 0 Then
        Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptEncrypt()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
    End If
Else
    ' Decrypt the text
    lResult = CryptDecrypt(lKey, 0, 1, 0, sText, lData)
    sText = Left(sText, lData)
    If lResult = 0 Then
        'Note:  if data to decrypt is blank, get LastDllError = -2146893820 with both RC2 and RC4 algorithm => the
        'decryption routine always expects some data to decrypt.
        'RC2 will also reject data that was not created using RC2 or data that was encrypted with a different key.
        'RC4 will "succeed" in these cases, but will return nonsense data.
        Err.Raise 1001, App.EXEName & ".SQLToolkit.EncryptDecrypt", "Windows API 'CryptDecrypt()' failed:  " & vbCrLf & vbCrLf _
                            & "API Error #" & Err.LastDllError & ":  '" & GetAPIErrMsg(Err.LastDllError) & "'."
    End If
End If

EncryptDecrypt = sText
If lHash <> 0 Then
    lResult = CryptDestroyHash(lHash)
End If
If lKey <> 0 Then
    lResult = CryptDestroyKey(lKey)
End If
    
Exit Function
EH:
If lHash <> 0 Then
    lResult = CryptDestroyHash(lHash)
End If
If lKey <> 0 Then
    lResult = CryptDestroyKey(lKey)
End If
ErrHandlerPrivate False, "SQLToolkit", "EncryptDecrypt"
End Function

'this has been optimized for use with WEM passwords.  The RC2 "block" algorithm is considered more secure
'and also produces an encrypted output that is longer than the input--which appears to be useful
'for obscuring the length of short passwords.  If need an output that is the same length as the
'input (except for doubling due to hex conversion), then use RC4.
'Routine assumes user wants encrypted data returned as hexidecimal text.
'Both RC2 and RC4 are implemented here as symmetric "private key" algorithms, so the same key
'must be used to encrypt and decrypt.
Public Function EncryptText(ByVal sTextToEncrypt As String) As String
On Error GoTo EH
Dim sPK

'use arbitrary standard string as private key for encryption.  Use of date data type helps to
'obscure the key in the .exe file more than embedded string literal.
'(NOTE:  will need same exact key to decrypt the data so in
'general, not a good idea to change this)
'V1.08:  Changed to hard coded string.  Previously, used Format() function which can give
'different results if user changes default Windows date format.
sPK = "WEM Automation 12/15/2004 01:36:43 PM"
EncryptText = StringToHex(EncryptDecrypt(sTextToEncrypt, sPK, True, ENCRYPTION_TYPE_RC2))

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "EncryptText"
End Function

'see comments above
Public Function DecryptText(ByVal sTextAsHexToDecrypt As String) As String
On Error GoTo EH
Dim sPK

'if passed blank encrypted string
If sTextAsHexToDecrypt = "" Then
    'RC4 algorithm encrypts a blank as blank.  However, RC2 will encrypt a blank as
    'non-blank.  RC2 decryption will fail if try to decrypt a blank string.  Caller probably expects a
    'blank encryption string to be decoded as blank in any case, so do that here (even though not technically accurate).
    DecryptText = ""
    Exit Function
End If

'use arbitrary standard string as the private key.  Use of date data type helps to
'obscure the key in the .exe file more than embedded string literal.
'(NOTE:  MUST match key used in encrypting the data or call will fail with
'API error.  So, in general, not a good idea to change the key.)
'V1.08:  Changed to hard coded string (previously, Format() command could change if user changes Windows date format)
sPK = "WEM Automation 12/15/2004 01:36:43 PM"
                        
DecryptText = EncryptDecrypt(StringFromHex(sTextAsHexToDecrypt), sPK, False, ENCRYPTION_TYPE_RC2)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "DecryptText"
End Function
'1.07 for Chinese
Public Function DecryptText2(ByVal sTextAsHexToDecrypt As String) As String
On Error GoTo EH
Dim sPK
Dim ii As Integer

'if passed blank encrypted string
If sTextAsHexToDecrypt = "" Then
    DecryptText2 = ""
    Exit Function
End If

sPK = "WEM Automation 12/15/2004 01:36:43 PM"
DecryptText2 = ""
For ii = 1 To Len(sTextAsHexToDecrypt)
    DecryptText2 = DecryptText2 & Chr(Asc(Mid(sTextAsHexToDecrypt, ii, 1)) - 3)
Next ii
                        
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "DecryptText2"
End Function

Private Function StringToHex(ByVal sText As String) As String
    Dim lCount  As Long
    Dim sHex    As String
    Dim sResult As String
    
    For lCount = 1 To Len(sText)
        sHex = Hex(Asc(Mid(sText, lCount, 1)))
        If Len(sHex) = 1 Then
            sHex = "0" & sHex
        End If
        sResult = sResult & sHex
    Next
    
    StringToHex = sResult
End Function

Private Function StringFromHex(ByVal sText As String) As String
    Dim lCount  As Long
    Dim sChar   As String
    Dim sResult As String
    Dim lLength As Long
    
    lLength = Len(sText)
    If lLength Mod 2 <> 0 Then
         Err.Raise 1001, App.EXEName & ".SQLToolkit.StringFromHex", "Procedure StringFromHex() was passed a string hex value that has an odd number of characters."
    End If
        
    For lCount = 1 To lLength
        sChar = Mid(sText, lCount, 1)
        If sChar < "0" Or sChar > "9" Then
            If sChar < "A" Or sChar > "F" Then
                Err.Raise 1001, App.EXEName & ".SQLToolkit.StringFromHex", "Procedure StringFromHex() was passed a string hex value that contains a non-hex character = '" & sChar & "'."
            End If
        End If
    Next
    
    For lCount = 1 To lLength Step 2
        sResult = sResult & Chr("&H" & Mid(sText, lCount, 2))
    Next
    
    StringFromHex = sResult
End Function

'following converts a Windows API error number (typically as API return value) to text string
Private Function GetAPIErrMsg(rtn As Long) As String
   Dim strError As String, i As Long
   Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
   strError = String(512, 0)

   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, rtn, 0&, strError, 512, ByVal 0&

   i = InStr(strError, Chr$(0))
   If i > 1 Then strError = Left$(strError, i - 1)
   
   'based on testing, above code appends a terminal CRLF to the string.  strip this off as it confuses
   'formatting elsewhere
   If Right(strError, 2) = vbCrLf Then
       strError = Left(strError, Len(strError) - 2)
   End If
   GetAPIErrMsg = strError
End Function
'=============================================================================
'END PASSWORD ENCRYPTION/DESCRYPTION FUNCTIONS
'=============================================================================

'(note:  Be careful about putting any code here that could affect error information--or be sure to restore
'the error info when done.  class instance may be destroyed while client error handler is active.)
Private Sub Class_Terminate()

End Sub

'Executes a command line function and returns result.  could be used for "net start" and similar
'commands.
Public Function ExecuteCmdLine(cmdline As String, Optional bHideWindow As Boolean = True) As Long
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   Dim ret As Long
   Const STARTF_USESHOWWINDOW = &H1
   Const SW_HIDE = 0

   If bHideWindow Then
        'suppress display of the DOS window (based on testing, this works for net start and net stop)
        start.dwFlags = STARTF_USESHOWWINDOW
        start.wShowWindow = SW_HIDE
   End If

   ' Initialize the STARTUPINFO structure:
   start.cb = Len(start)

   ' Start the shelled application:
   ret = CreateProcessA(0&, cmdline, 0&, 0&, 1&, _
      NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)


   ' Wait for the shelled application to finish:
      ret = WaitForSingleObject(proc.hProcess, INFINITE)
      Call GetExitCodeProcess(proc.hProcess, ret)
      Call CloseHandle(proc.hThread)
      Call CloseHandle(proc.hProcess)
      ExecuteCmdLine = ret
End Function

'Note:  this only handles '*', '%', and "_" situations.
Public Function VBLikePatternToTSQL(ByVal sVBLikePattern As String) As String
Dim sResult As String

'the underscore is a SQL Like special char, so must isolate it
sResult = Replace(sVBLikePattern, "_", "[_]")

'replace any "%" with TSQL underscore
sResult = Replace(sResult, "%", "_")

'replace any * with TSQL %
sResult = Replace(sResult, "*", "%")

VBLikePatternToTSQL = sResult

End Function


'Notes:
'1.  Return value is message result from "DBCCConcurrencyViolation".
'2.  Uses SQLDMO as results are returned as a SQL msg (not as a result set).
'    (Can get messages with DAO, but have to use LogMessages property on the
'    query def object and a temp .mdb file.  Another alternative is to call
'    the DMO objects directly from a TSQL script.  See pp. 176-178 of
'    "Real-World SQL-DMO for SQL Server" for sample code.  This involves using
'    TSQL's OLE commands.  Advantage would be that clients hitting the database
'    would not need DMO installed locally.  For WEM, likely that DMO will be part of
'    the install, so no advantage.
'3.  Alternatives for the sDBCCArgument are:  DISPLAY (the default), RESET (which resets count totals,
'    STARTLOG (logs nonzero concurrency counts to SQL log--which is the default), and
'    STOPLOG (turns off logging)
Public Function DBCCConcurrencyViolation(oSvrDMO As SQLDMO.SQLServer, _
                                         Optional ByVal sDBCCArgument = "DISPLAY") As String
On Error GoTo EH
Dim sTSQL As String
Dim sMsg As String

If sDBCCArgument = "" Then
    sDBCCArgument = "DISPLAY"
End If

sTSQL = "DBCC CONCURRENCYVIOLATION(" & sDBCCArgument & ")"

'get the message (Note:  SQLDMO Help recommends using
'ExecuteWithResultsAndMessages2 with VB as [Length] param is optional.
'However, this method does not exist.  Following method also has
'optional [Length] param, so documentation is probably incorrect.
oSvrDMO.ExecuteWithResultsAndMessages sTSQL, , sMsg

'if this server does not have the workload governor
If sMsg = "[Microsoft][ODBC SQL Server Driver][SQL Server]DBCC execution completed." _
     & " If DBCC printed error messages, contact your system administrator." Then
     'append some additional info to the message
     sMsg = sMsg & vbCrLf & vbCrLf _
            & "( => THIS SERVER DOES *NOT* HAVE THE WORKLOAD GOVERNOR)"
End If

'return message to caller
DBCCConcurrencyViolation = sMsg

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "DBCCConcurrencyViolation"
End Function

'if sDBName is blank, uses current database context.
'shrinks both .mdf and .ldf files.
'(Default percent of 25 is same action executed by AutoShrink property of database.)
'In most cases, returns a result set with info about the file sizes and contained pages.
'--see BOL for details.
Public Function ShrinkDatabase(dbSvr As dao.Database, _
                               Optional ByVal sDBName As String = "", _
                               Optional ByVal lTargetPercent As Long = 25) As dao.Recordset
On Error GoTo EH
Dim sTSQL As String

If sDBName = "" Then
   'use current database for the connection
   sTSQL = "DBCC SHRINKDATABASE(0," & lTargetPercent & ")"
Else
   sTSQL = "DBCC SHRINKDATABASE('" & TSQLQuoteFix(sDBName) & "'," & lTargetPercent & ")"
End If

'return result set (or Nothing if database doesn't need to shrink).
'Leave database context with default
'(so that use of "0" above will work).
Set ShrinkDatabase = ExecTSQL(dbSvr, sTSQL, True, , True)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "ShrinkDatabase"
End Function

'Fills vServerNameList_RTN with string array in variant of all SQL Server names.
'Function return is count of # of array elements
'Uses SQL-DMO as the API calls required are complex
Public Function GetSQLEngineList(vServerNameList_RTN As Variant) As Long
On Error GoTo EH
Dim sNames() As String
Dim lIndex As Long
Dim lCount As Long
Dim oSQLServer As SQLDMO.SQLServer
Dim oServerList As SQLDMO.NameList
Dim vServer As Variant

ReDim sNames(20)

Set oSQLServer = New SQLDMO.SQLServer
Set oServerList = oSQLServer.Application.ListAvailableSQLServers

'(NOTE:  Use of For Each syntax in VB5 here seems to cause lockups)
lCount = 0
For lIndex = 1 To oServerList.Count
    lCount = lCount + 1
    If (lCount - 1) > UBound(sNames) Then
        ReDim Preserve sNames(lCount + 20)
    End If
    sNames(lCount - 1) = oServerList(lIndex)
Next

If lCount = 0 Then
    vServerNameList_RTN = Array()
Else
    ReDim Preserve sNames(lCount - 1)
    vServerNameList_RTN = sNames
End If

GetSQLEngineList = lCount

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetSQLEngineList"
End Function

'NOTE:  You can create dbSvr without specifying a database.  So, this function can be called
'without any knowledge of the databases a server contains (or could create dbSvr using 'master')
'The database names are in the 'name' field.

'if specify sDBName, then results are filtered.
Public Function GetDBNamesRecSet(dbSvr As dao.Database, Optional ByVal sDBName As String = "") As dao.Recordset
On Error GoTo EH
Dim sTSQL As String

'NOTE:  The documented way to do this is by using "sp_databases", but this takes 15-20 seconds to execute
'(possibly because it calculates the size of each database).  Hitting the sysdatabases table directly
'returns almost immediately.  It may return databases that the current user does not actually have access to.

sTSQL = "SELECT name FROM master.dbo.sysdatabases"

If sDBName <> "" Then
   sTSQL = sTSQL & " WHERE name = '" & TSQLQuoteFix(sDBName) & "'"
End If

Set GetDBNamesRecSet = ExecTSQL(dbSvr, sTSQL, True, "master")

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetDBNamesRecSet"
End Function



'returns object ID within the database (or 0 if not found).
'based on contents of sysobjects, appears object IDs are non-zero.
Public Function GetOBJECT_ID(dbSvr As dao.Database, sObjectName As String, Optional sDBName As String = "") As Long
On Error GoTo EH
Dim rs As dao.Recordset
Dim sTSQL As String

'default => object doesn't exist
GetOBJECT_ID = 0

sTSQL = "SELECT OBJECT_ID('" & TSQLQuoteFix(sObjectName) & "') AS ObjID"

Set rs = ExecTSQL(dbSvr, sTSQL, True, sDBName)
If Not rs.EOF Then
    GetOBJECT_ID = GetFV(rs, "ObjID")
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetOBJECT_ID"
End Function

'NOTE:  (1) For SQL 2000, function returns a generated date/time that is the newest table creation date plus
'           the "schema_ver" divided by 16 and converted to millisec.  Differences are detectable using the
'           <> operators, but will not be visible in printout unless use CDbl() on returned result.
'           Method could fail if modify a table, then drop, and recreate it--all in less than 5 seconds or so.
'           In actual practice this is not expected.  Expect to move to SQL 2005 shortly, so no need for "perfect"
'           method.
'       (2) For SQL 2005, function returns the newest modify_date for all tables in the database.  This will
'           be a true modified date and should be reliable.  (Will not detect table deletions without
'           recreation.)
'       (3) Intent is to give caller overall value that can be compared to see if SQL db structure has changed from past.
Public Function GetDBSchemaModDTime(dbSvr As dao.Database, Optional ByVal sDBName As String = "") As Date
On Error GoTo EH
Dim sTSQL As String
Dim rs As dao.Recordset
Dim dMaxDateAsDbl As Double
Dim dVersionSum As Double
Dim dPsuedoModDTimeAsDbl As Double

'set default = current dtime
GetDBSchemaModDTime = Now

Select Case True
     Case IsSqlSvr2012(dbSvr)   'V1.12
        
        '(type = 'U' restricts tests to 'user' tables)
        sTSQL = "SELECT TOP 1 name, create_date, modify_date" & vbCrLf _
            & " FROM sys.objects" & vbCrLf _
            & " WHERE type = 'U'" & vbCrLf _
            & " ORDER BY modify_date DESC"
            
        Set rs = ExecTSQL(dbSvr, sTSQL, True, sDBName)
        If Not rs.EOF Then
            GetDBSchemaModDTime = GetFV(rs, "modify_date")
        Else
            'based on testing, no records will be returned if no tables in the database.
            'in this case, return the creation date/time of the database
            sTSQL = "SELECT crdate FROM master.dbo.sysdatabases" & vbCrLf _
                 & " WHERE name = '" & TSQLQuoteFix(sDBName) & "'"
            Set rs = ExecTSQL(dbSvr, sTSQL, True, "master")
            If Not rs.EOF Then
                GetDBSchemaModDTime = GetFV(rs, "crdate")
            Else
                'we have no valid date to return--return default => very recent modification
                GetDBSchemaModDTime = Now
            End If
        End If
     Case IsSqlSvr2008(dbSvr)   'V1.06
        
        '(type = 'U' restricts tests to 'user' tables)
        sTSQL = "SELECT TOP 1 name, create_date, modify_date" & vbCrLf _
            & " FROM sys.objects" & vbCrLf _
            & " WHERE type = 'U'" & vbCrLf _
            & " ORDER BY modify_date DESC"
            
        Set rs = ExecTSQL(dbSvr, sTSQL, True, sDBName)
        If Not rs.EOF Then
            GetDBSchemaModDTime = GetFV(rs, "modify_date")
        Else
            'based on testing, no records will be returned if no tables in the database.
            'in this case, return the creation date/time of the database
            sTSQL = "SELECT crdate FROM master.dbo.sysdatabases" & vbCrLf _
                 & " WHERE name = '" & TSQLQuoteFix(sDBName) & "'"
            Set rs = ExecTSQL(dbSvr, sTSQL, True, "master")
            If Not rs.EOF Then
                GetDBSchemaModDTime = GetFV(rs, "crdate")
            Else
                'we have no valid date to return--return default => very recent modification
                GetDBSchemaModDTime = Now
            End If
        End If
        
     Case IsSqlSvr2005(dbSvr)
        
        '(type = 'U' restricts tests to 'user' tables)
        sTSQL = "SELECT TOP 1 name, create_date, modify_date" & vbCrLf _
            & " FROM sys.objects" & vbCrLf _
            & " WHERE type = 'U'" & vbCrLf _
            & " ORDER BY modify_date DESC"
            
        Set rs = ExecTSQL(dbSvr, sTSQL, True, sDBName)
        If Not rs.EOF Then
            GetDBSchemaModDTime = GetFV(rs, "modify_date")
        Else
            'based on testing, no records will be returned if no tables in the database.
            'in this case, return the creation date/time of the database
            sTSQL = "SELECT crdate FROM master.dbo.sysdatabases" & vbCrLf _
                 & " WHERE name = '" & TSQLQuoteFix(sDBName) & "'"
            Set rs = ExecTSQL(dbSvr, sTSQL, True, "master")
            If Not rs.EOF Then
                GetDBSchemaModDTime = GetFV(rs, "crdate")
            Else
                'we have no valid date to return--return default => very recent modification
                GetDBSchemaModDTime = Now
            End If
        End If
    
    Case IsSqlSvr2000(dbSvr)
        'no modify date exists for SQL 2000.  create a psuedo mod date using most recent table creation date
        'plus the sum of all table version counters.
        '(type = 'U' restricts tests to 'user' tables)
        sTSQL = "SELECT MAX(crdate) AS MAX_DATE, SUM(schema_ver) AS VERSION_SUM" & vbCrLf _
              & "FROM sysobjects" & vbCrLf _
              & "WHERE type = 'U'"
            
        Set rs = ExecTSQL(dbSvr, sTSQL, True, sDBName)
        If Not rs.EOF Then
            dMaxDateAsDbl = CDbl(GetFV(rs, "MAX_DATE"))
            '"schema_ver" increase by 16 each time a table structure is modified, so divide by this factor.
            'Typical sum(schema_ver) for all tables in the WEMSQL database is then around 4700.
            'Divide by another factor of 1000*86400 to convert this to millisec in Date datatype format (86400 sec in one day).
            '--which results in about 4.7 seconds total added to the MAX(crdate).
            'A single version increase will cause a change of 1 millisec in the returned date.
            'This shows up in the returned CDbl(date) as an increase of of ~1.2 E-8 as a fraction of the day.
            'Access can store/retrieve this in a Date data type.
            'Difference will affect < > compares, but will not be visible unless do CDbl() on the date.
            dVersionSum = CDbl(GetFV(rs, "VERSION_SUM")) / (16# * 1000# * 86400#)
            dPsuedoModDTimeAsDbl = dMaxDateAsDbl + dVersionSum
            'return as date.  millisec will only be visible if do CDbl() on the result.  However,
            'Access will save/recall the millisec to/from a Date database field.
            'Round value to nearest E-10 to help avoid round off errors when compare the results with stored value.
            GetDBSchemaModDTime = CDate(CDbl(Format(dPsuedoModDTimeAsDbl, "0.0000000000")))
        End If
    
    Case Else
        'return the default = current date/time.  assumes worse case--that db has always been recently modified.
End Select

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetDBSchemaModDTime"
End Function

Public Function GetSQLSvrVersionString(dbSvr As dao.Database) As String
On Error GoTo EH
Dim rs As dao.Recordset
Dim sTSQL As String

sTSQL = "SELECT @@VERSION AS SQLVersionString"

Set rs = ExecTSQL(dbSvr, sTSQL, True, "master")
If Not rs.EOF Then
    GetSQLSvrVersionString = GetFV(rs, "SQLVersionString")
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetSQLSvrVersionString"
End Function

Public Function IsSqlSvr2000(dbSvr As dao.Database) As Boolean
On Error GoTo EH
IsSqlSvr2000 = (Left(Trim(GetSQLSvrVersionString(dbSvr)), 26) = "Microsoft SQL Server  2000")
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSqlSvr2000"
End Function

Public Function IsSqlSvr2005(dbSvr As dao.Database) As Boolean
On Error GoTo EH

'V1.15:  Updated to work properly against SQL Server 2005
IsSqlSvr2005 = (Left(Trim(GetSQLSvrVersionString(dbSvr)), 25) = "Microsoft SQL Server 2005")
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSqlSvr2005"
End Function
Public Function IsSqlSvr2008(dbSvr As dao.Database) As Boolean  'V1.06
On Error GoTo EH

'V1.15:  Updated to work properly against SQL Server 2008
IsSqlSvr2008 = (Left(Trim(GetSQLSvrVersionString(dbSvr)), 25) = "Microsoft SQL Server 2008")
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSqlSvr2008"
End Function
Public Function IsSqlSvr2012(dbSvr As dao.Database) As Boolean  'V1.12
On Error GoTo EH

'V1.15:  Updated to work properly against SQL Server 2008
IsSqlSvr2012 = (Left(Trim(GetSQLSvrVersionString(dbSvr)), 25) = "Microsoft SQL Server 2012")
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSqlSvr2012"
End Function

'(also see ScreenDBForAccess below)
Public Function IsDBInSvr(dbSvr As dao.Database, ByVal sDBName As String) As Boolean
On Error GoTo EH
Dim rs As dao.Recordset

'must capture blanks or GetDBNamesRecSet will return all databases
If sDBName = "" Then
    IsDBInSvr = False
    Exit Function
End If

Set rs = GetDBNamesRecSet(dbSvr, sDBName)

IsDBInSvr = (Not rs.EOF)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsDBInSvr"
End Function

Public Function IsSystemTableOrView(sTableOrViewName As String) As Boolean
On Error GoTo EH

IsSystemTableOrView = ((sTableOrViewName Like "sys.*") Or _
                      (sTableOrViewName Like "dbo.sys*") Or _
                      (sTableOrViewName Like "INFORMATION_SCHEMA.*"))

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSystemTableOrView"
End Function

Public Function IsSystemDatabase(sDBName As String) As Boolean
On Error GoTo EH

Select Case UCase(sDBName)
    Case "MASTER", "MODEL", "MSDB", "TEMPDB"
        IsSystemDatabase = True
End Select
    
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsSystemDatabase"
End Function

'Does test similar to those that cause the ODBC connection dialog to show up.
'Will detect whether database exists and whether oSvrDMO login can access the specified
'database.  Does not require a DAO database object.  Can test existence of
'any database; however, result of bLoginHasDBAccess_Rtn
'only applies to the login used to connect the oSvrDMO object.
'Possible return combinations:
' Function   bDBExists_Rtn  bLoginHasDBAccess_Rtn  Meaning
'   True       n/a (True)      n/a (True)          Access okay
'   False      False           n/a (False)         Database doesn't exist
'   False      True            False               Database exists, but Login doesn't have access
'   False      True            True                Unknown error--ck lErrNumber_Rtn/sErrDescription_Rtn

Public Function ScreenDBForAccess(oSvrDMO As SQLDMO.SQLServer, ByVal sDBName As String, _
                             Optional bDBExists_Rtn As Boolean, _
                             Optional bLoginHasDBAccess_Rtn As Boolean, _
                             Optional lErrNumber_Rtn As Long, _
                             Optional sErrDescription_Rtn As String) As Boolean
On Error Resume Next
'set default return => tests failed
ScreenDBForAccess = False
bDBExists_Rtn = False
bLoginHasDBAccess_Rtn = False
lErrNumber_Rtn = 0
sErrDescription_Rtn = ""

Err.Clear
'(Note:  based on testing, DMO always resets the database context to 'master', so executing
'this statement has no effect beyond this batch)
oSvrDMO.ExecuteImmediate "USE [" & sDBName & "]"
lErrNumber_Rtn = Err.Number
sErrDescription_Rtn = Err.Description

'return true if everything okay.
ScreenDBForAccess = (lErrNumber_Rtn = 0)

'process any error messages
bDBExists_Rtn = (lErrNumber_Rtn <> -2147220593)               '  "no entry in sysdatabases"
If bDBExists_Rtn Then
    bLoginHasDBAccess_Rtn = (lErrNumber_Rtn <> -2147220588)   '  "not a valid user in database"
End If  'else, default to False as user access is N/A if database doesn't exist
'don't return error to caller
Err.Clear
End Function
                           
Public Function ConnectSvrDMO(ByVal sDSN As String, _
                              ByVal sSvrNameSkipDSN As String, _
                              ByVal sLoginName As String, _
                              ByVal sPassword As String, _
                              Optional ByVal bResolveSvrToIP As Boolean = True) As SQLDMO.SQLServer
On Error GoTo EH
Dim oSvrDMO As SQLDMO.SQLServer
Dim sServerName As String
Dim ErrClone1 As ErrClone

If bResolveSvrToIP Then
    sServerName = ResolveSvrEngineToIP(sDSN, sSvrNameSkipDSN)
Else
    sServerName = GetServerName(sDSN, sSvrNameSkipDSN)
End If

Set oSvrDMO = New SQLDMO.SQLServer
oSvrDMO.Connect sServerName, sLoginName, sPassword
Set ConnectSvrDMO = oSvrDMO

Exit Function
EH:
'must destroy DMO server object or error description is overwritten--unknown why
ErrClone1 = MakeErrClone
CloseDMOServerRef oSvrDMO
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "ConnectSvrDMO"
End Function

Public Function ConnectSvrDMO2(ByVal sODBCConnStr As String, _
                              Optional ByVal bResolveSvrToIP As Boolean = True) As SQLDMO.SQLServer
On Error GoTo EH
Dim oSvrDMO As SQLDMO.SQLServer
Dim sDSN As String
Dim sSvrNameSkipDSN As String
Dim sDriver As String
Dim sUID As String
Dim sPWD As String
Dim sServerName As String
Dim ErrClone1 As ErrClone

ExtractODBCParams sODBCConnStr, sDSN, sSvrNameSkipDSN, sDriver, sUID, sPWD

If bResolveSvrToIP Then
    sServerName = ResolveSvrEngineToIP(sDSN, sSvrNameSkipDSN)
Else
    sServerName = GetServerName(sDSN, sSvrNameSkipDSN)
End If

Set oSvrDMO = New SQLDMO.SQLServer
oSvrDMO.Connect sServerName, sUID, sPWD
Set ConnectSvrDMO2 = oSvrDMO

Exit Function
EH:
'must destroy DMO server object or error description is overwritten--unknown why
ErrClone1 = MakeErrClone
CloseDMOServerRef oSvrDMO
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "ConnectSvrDMO2"
End Function

'could find no way to execute sp_helplogins using DAO.  This sp returns an varbinary type which
'will cause DAO 3.60 to crash.  Can't use INSERT-EXEC to store into temp table as sp returns
'two results sets--which causes INSERT-EXEC to fail. sp_helplogins is a large routine and is
'too large to reasonably incorporate into this class.  Resorted to returning as DMO QueryResult.
'This requires making an additional connection to the server, so not a good idea to execute this
'too often.  If specify sLoginName, results are returned only for specified login.
'(If really need results in DAO recordset, could create a temporary MDB and return recordset
'based on it.)
''this will return the two SQL resultset which returns:
'#1:  LoginName (string), SID (binary), DefDBName (string), DefLangName (string), AUser (string), ARemote (string)
'#2:  LoginName (string), DBName (string), UserName (string), UserOrAlias (string)
'DMO row/cols and resultsets are use 1 based counting.
Public Function GetLoginNamesQResult(oSvrDMO As SQLDMO.SQLServer, Optional ByVal sLoginName As String = "") As SQLDMO.QueryResults2
On Error GoTo EH
Dim sTSQL As String
Dim ErrClone1 As ErrClone

sTSQL = "sp_helplogins"

If sLoginName <> "" Then
    '(BOL states that the login name must exist if specified.  however, based on testing, no error is
    ' generated if loginname does not exist--merely get back no rows)
    sTSQL = sTSQL & " '" & TSQLQuoteFix(sLoginName) & "'"
End If

Set GetLoginNamesQResult = oSvrDMO.ExecuteWithResults(sTSQL)

Exit Function
EH:
'must destroy DMO server object or error description is overwritten--unknown why
ErrClone1 = MakeErrClone
CloseDMOServerRef oSvrDMO
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "ServerIsReachable"
End Function

Public Function IsLoginInSvr(dbSvr As dao.Database, ByVal sLoginName As String) As Boolean
On Error GoTo EH
Dim oSvrDMO As SQLDMO.SQLServer
Dim oDMOQryResult As SQLDMO.QueryResults2
Dim ErrClone1 As ErrClone

If sLoginName = "" Then
    IsLoginInSvr = False
    Exit Function
End If

Set oSvrDMO = ConnectSvrDMO2(dbSvr.Connect)

Set oDMOQryResult = GetLoginNamesQResult(oSvrDMO, sLoginName)
     
IsLoginInSvr = (oDMOQryResult.Rows > 0)

'close the DMO server object, else get orphaned connections (note:  this also kills the QryResult object)
CloseDMOServerRef oSvrDMO

Exit Function
EH:
ErrClone1 = MakeErrClone
CloseDMOServerRef oSvrDMO
RestoreErrObject ErrClone1
ErrHandlerPrivate False, "SQLToolkit", "IsLoginInSvr"
End Function

'adapted/simplified from sp_helpuser.  If user is an "orphan" (that is, there is no corresponding loginname), then
'loginname will be blank.  Returned fields are:
'username  (NOTE:  This may be 'dbo' if the name is the database owner.)
'uid = user ID
'sid = login security ID  (cast to sql_variant so result is in readable hex in DAO). for dao 3.6, must cast to varchar
'       which removes ability to read these.  in future, consider adding user defined fn to do the conversion.

'if specify non-blank sUserName, results are filtered for this users.  (if user name matches database owner login,
'then match will also occur with the 'dbo' user)
Public Function GetUserNamesRecSet(dbSvr As dao.Database, _
                                   ByVal sDBName As String, _
                                   Optional ByVal sUserName As String = "") As dao.Recordset
On Error GoTo EH
Dim sTSQL As String

sTSQL = ""
sTSQL = sTSQL & "SELECT" & vbCrLf
sTSQL = sTSQL & "   usu.name AS username" & vbCrLf
sTSQL = sTSQL & "  ,lo.loginname" & vbCrLf
sTSQL = sTSQL & "  ,usu.uid" & vbCrLf
sTSQL = sTSQL & "  ,CAST(usu.sid  AS varchar) AS sid" & vbCrLf
sTSQL = sTSQL & "FROM" & vbCrLf
sTSQL = sTSQL & "   sysusers usu left outer join master.dbo.syslogins lo " & vbCrLf
sTSQL = sTSQL & "   on usu.sid = lo.sid" & vbCrLf
sTSQL = sTSQL & "WHERE" & vbCrLf
sTSQL = sTSQL & "   (usu.islogin = 1 and usu.isaliased = 0 and usu.hasdbaccess = 1)" & vbCrLf

If sUserName <> "" Then
    sTSQL = sTSQL & "   AND (" & vbCrLf
    sTSQL = sTSQL & "            (usu.name = '" & TSQLQuoteFix(sUserName) & "')" & vbCrLf
    sTSQL = sTSQL & "         OR (" & vbCrLf
    sTSQL = sTSQL & "                (usu.name = 'dbo')" & vbCrLf
    sTSQL = sTSQL & "                AND" & vbCrLf
    sTSQL = sTSQL & "                (lo.loginname = '" & TSQLQuoteFix(sUserName) & "')" & vbCrLf
    sTSQL = sTSQL & "             )" & vbCrLf
    sTSQL = sTSQL & "        )" & vbCrLf
End If

Set GetUserNamesRecSet = ExecTSQL(dbSvr, sTSQL, True, sDBName)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetUserNamesRecSet"
End Function

'NOTE:  For screening test if user has any ability to access a database, ScreenDBForAccess() is preferred.
'For example, IsUserInDB will fail if user is a SysAdmin, but is not the database owner--even though a
'SysAdmin has full access to the database.
Public Function IsUserInDB(dbSvr As dao.Database, ByVal sDBName As String, ByVal sUserName As String) As Boolean
On Error GoTo EH
Dim rs As dao.Recordset

If sUserName = "" Then
   IsUserInDB = False
   Exit Function
End If

Set rs = GetUserNamesRecSet(dbSvr, sDBName, sUserName)
     
IsUserInDB = (Not rs.EOF)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsUserInDB"
End Function

'if pass blank for sUserName (or leave off), then all users are returned.
'if pass specific user for sUserName, then only records for that user are returned.
'--similar reasoning for optional sDBRoleName param.

'returned fields are:
'UserName = user name in the database.  NOTE:  database "owner" user will always be listed as 'dbo'.
'    This proc handles filtering based on sUserName by also testing for dbo user and login name match.
'GroupName = database role that this user belongs to.
'LoginName  = login name corresponding to this user (may be blank if orphan)
'DefDBName
'UserID (numerical identifier for this user)
'SID (security ID)

Public Function GetUserRolesRecSet(dbSvr As dao.Database, _
                                   ByVal sDBName As String, _
                                   Optional ByVal sUserName As String = "", _
                                   Optional ByVal sDBRoleName As String = "") As dao.Recordset
On Error GoTo EH
Dim sTSQL As String
Dim rsTemp As dao.Recordset
Dim sFilter As String
Dim bSQL2000 As Boolean

'had to output results of sp_helpuser to temp table so can filter results and
'convert varbinary sid to varchar (else DAO 3.60 has problems)
'V1.15:  Added logic to support both SQL Server 2000 and 2005.  Had to use different temp table
'name (? previous one is now used by SQL Server itself).
bSQL2000 = IsSqlSvr2000(dbSvr)

sTSQL = ""
sTSQL = sTSQL & "--turn off msgs to reduce network traffic" & vbCrLf
sTSQL = sTSQL & "Set nocount on" & vbCrLf
sTSQL = sTSQL & "--remove any existing temp table" & vbCrLf
sTSQL = sTSQL & "IF OBJECT_ID('tempdb.dbo.#TempSPHelpUser','U') IS NOT NULL" & vbCrLf
sTSQL = sTSQL & "    DROP TABLE #TempSPHelpUser" & vbCrLf
sTSQL = sTSQL & "--create temp table to hold output of stored procedure" & vbCrLf
sTSQL = sTSQL & "Create Table #TempSPHelpUser" & vbCrLf
sTSQL = sTSQL & "   (" & vbCrLf
sTSQL = sTSQL & "    UserName        sysname        collate database_default Null" & vbCrLf
sTSQL = sTSQL & "   ,GroupName       sysname        collate database_default Null" & vbCrLf
sTSQL = sTSQL & "   ,LoginName       sysname        collate database_default Null" & vbCrLf
sTSQL = sTSQL & "   ,DefDBName       sysname        collate database_default Null" & vbCrLf
If Not bSQL2000 Then
    'V1.15: added for SQL2005 (and above) compatibility
    sTSQL = sTSQL & "   ,DefSchemaName   sysname        collate database_default Null" & vbCrLf
End If
sTSQL = sTSQL & "   ,UserID             char(6)        Null" & vbCrLf
sTSQL = sTSQL & "   ,SID             varbinary(85)  Null" & vbCrLf
sTSQL = sTSQL & "   )" & vbCrLf
sTSQL = sTSQL & "--execute stored procedure and store output into temp table" & vbCrLf
sTSQL = sTSQL & "INSERT INTO #TempSPHelpUser EXEC ('sp_helpuser')" & vbCrLf
sTSQL = sTSQL & "--return results with optional filtering and conversion of" & vbCrLf
sTSQL = sTSQL & "--varbinary to varchar (or DAO 3.6 will have problems)" & vbCrLf
sTSQL = sTSQL & "SELECT" & vbCrLf
sTSQL = sTSQL & "UserName," & vbCrLf
sTSQL = sTSQL & "GroupName," & vbCrLf
sTSQL = sTSQL & "LoginName," & vbCrLf
sTSQL = sTSQL & "DefDBName," & vbCrLf
If Not bSQL2000 Then
    'V1.15: added for SQL2005 (and above) compatibility
    sTSQL = sTSQL & "DefSchemaName," & vbCrLf  'V1.15: added for SQL2005 compatibility
End If
sTSQL = sTSQL & "UserID," & vbCrLf
sTSQL = sTSQL & "CAST(SID AS varchar) AS  'SID'" & vbCrLf
sTSQL = sTSQL & "from  #TempSPHelpUser" & vbCrLf

If sUserName <> "" Or sDBRoleName <> "" Then
    sTSQL = sTSQL & "WHERE" & vbCrLf
End If

If sUserName <> "" Then
    'filter on match of (1) user name or (2) if user name is dbo, test for match with login name
    sTSQL = sTSQL & "(     (UserName = '" & TSQLQuoteFix(sUserName) & "')" & vbCrLf
    sTSQL = sTSQL & "   OR (  (UserName = 'dbo') AND (LoginName = '" & TSQLQuoteFix(sUserName) & "')  )      )" & vbCrLf
End If

If sDBRoleName <> "" Then
    sTSQL = sTSQL & IIf(sUserName = "", "", "AND ") & " (GroupName = '" & TSQLQuoteFix(sDBRoleName) & "')" & vbCrLf
End If

'add cleanup
sTSQL = sTSQL & "--clean up" & vbCrLf
sTSQL = sTSQL & "IF OBJECT_ID('tempdb.dbo.#TempSPHelpUser','U') IS NOT NULL" & vbCrLf
sTSQL = sTSQL & "    DROP TABLE #TempSPHelpUser" & vbCrLf

'return results
Set GetUserRolesRecSet = ExecTSQL(dbSvr, sTSQL, True, sDBName)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetUserRolesRecSet"
End Function

Public Function IsUserInDBRole(dbSvr As dao.Database, _
                               ByVal sDBName As String, _
                               ByVal sUserName As String, _
                               ByVal sDBRoleName As String) As Boolean
On Error GoTo EH
Dim rs As dao.Recordset

If (sUserName = "") Or (sDBRoleName = "") Then
    IsUserInDBRole = False
    Exit Function
End If

Set rs = GetUserRolesRecSet(dbSvr, sDBName, sUserName, sDBRoleName)
     
IsUserInDBRole = (Not rs.EOF)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsUserInDBRole"
End Function

'if login already exists, exits without error
Public Sub AddLoginToSvr(dbSvr As dao.Database, _
                          ByVal sLoginName As String, _
                          ByVal sPassword As String, _
                          Optional ByVal sDefaultDBName As String)
On Error GoTo EH
Dim sTSQL As String

If IsLoginInSvr(dbSvr, sLoginName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_addlogin '" & TSQLQuoteFix(sLoginName) & "','" & TSQLQuoteFix(sPassword) & "'" & vbCrLf

'if a default database is specified
If sDefaultDBName <> "" Then
    'add this to the TSQL request
    sTSQL = sTSQL & ",'" & TSQLQuoteFix(sDefaultDBName) & "'"
End If

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "AddLoginToSvr"
End Sub

'Notes:
'(1) handling logins that are object owners is complex
'and is not addressed here.  may add this in future.
'here's a sample script for a more robust approach.
'http://www.experts-exchange.com/Databases/Q_20821575.html
'same script also available here:
'http://www.sqlservercentral.com/scripts/contributions/413.asp
'(2) if login is not in this server, then exits early without error.
Public Sub DropLoginFromSvr(dbSvr As dao.Database, _
                            ByVal sLoginName As String)
On Error GoTo EH
Dim sTSQL As String

If Not IsLoginInSvr(dbSvr, sLoginName) Then
   Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_droplogin '" & sLoginName & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DropLoginFromSvr"
End Sub

'if pass bIgnoreOldPassword as True, then caller must have rights to overwrite
'password without the old password (e.g., SysAdmin rights)
Public Sub ChangeLoginPassword(dbSvr As dao.Database, _
                              ByVal sLoginName As String, _
                              ByVal sNewPassword As String, _
                              ByVal sOldPassword As String, _
                              Optional ByVal bIgnoreOldPassword As Boolean = False)
On Error GoTo EH
Dim sTSQL As String

sTSQL = ""
'(replacement of single quotes with double single quotes allows single quotes in the password and
' will also prevent SQL injection)
sTSQL = sTSQL & "sp_password" _
      & " " & IIf(bIgnoreOldPassword, "NULL", "'" & TSQLQuoteFix(sOldPassword) & "'") & "," _
      & "'" & TSQLQuoteFix(sNewPassword) & "'," _
      & "'" & TSQLQuoteFix(sLoginName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "ChangeLoginPassword"
End Sub

'returned fields are:
'ServerRole = one of the fixed server roles
'MemberName = Login that is a member of this fixed server role
'MemberSID = security ID of this login.

'if don't specify a login name or server role name, then get all logins and all server roles.
'if specify login and/or server role name, results are filtered.
Public Function GetSvrRolesRecSet(dbSvr As dao.Database, Optional ByVal sLoginName As String = "", _
                                  Optional ByVal sServerRoleName As String = "") As dao.Recordset
On Error GoTo EH
Dim rsTemp As dao.Recordset
Dim sTSQL As String
Dim bWhereClause As Boolean

'V1.15:  Updated to be SQL2005 compliant (previously, accessed several system tables that are now hidden)
sTSQL = ""
'(for reasons that are unclear, this TSQL will fail unless execute SET NOCOUNT ON at the start
' --appears to have something to do with the code that creates the temporary tables)
sTSQL = sTSQL & "SET NOCOUNT ON" & vbCrLf
sTSQL = sTSQL & "IF OBJECT_ID('tempdb.dbo.#RoleMembers','U') IS NOT NULL " & vbCrLf
sTSQL = sTSQL & "    DROP TABLE #RoleMembers" & vbCrLf
sTSQL = sTSQL & "CREATE TABLE #RoleMembers (" & vbCrLf
sTSQL = sTSQL & "       ServerRole sysname," & vbCrLf
sTSQL = sTSQL & "       MemberName sysname," & vbCrLf
sTSQL = sTSQL & "       MemberSID varbinary(85)  )" & vbCrLf
sTSQL = sTSQL & "INSERT INTO #RoleMembers EXEC sp_helpsrvrolemember" & vbCrLf
sTSQL = sTSQL & "SELECT * FROM #RoleMembers" & vbCrLf

If sServerRoleName <> "" Then
    'filter results to the specified server role name
    sTSQL = sTSQL & "WHERE (ServerRole = '" & TSQLQuoteFix(sServerRoleName) & "')" & vbCrLf
    bWhereClause = True
End If

If sLoginName <> "" Then
    'filter results to specified login name
    sTSQL = sTSQL _
          & IIf(bWhereClause, "AND ", "WHERE ") & "(MemberName = '" & TSQLQuoteFix(sLoginName) & "')" & vbCrLf
End If

'always delete the temporary table
sTSQL = sTSQL & "DROP TABLE #RoleMembers" & vbCrLf

'must execute against master database
Set GetSvrRolesRecSet = ExecTSQL(dbSvr, sTSQL, True, "master")

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetSvrRolesRecSet"
End Function

Public Function IsLoginInSvrRole(dbSvr As dao.Database, _
                            ByVal sLoginName As String, _
                            ByVal sServerRoleName As String) As Boolean
On Error GoTo EH
Dim rsTemp As dao.Recordset

If sLoginName = "" Then
   IsLoginInSvrRole = False
   Exit Function
End If

Set rsTemp = GetSvrRolesRecSet(dbSvr, sLoginName, sServerRoleName)

IsLoginInSvrRole = (Not rsTemp.EOF)

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "IsLoginInSvrRole"
End Function

'Note:  sServerRoleName must be one of the built in server roles:
'SysAdmin
'SecurityAdmin
'ServerAdmin
'SetupAdmin
'ProcessAdmin
'DiskAdmin
'DBCreator
'BulkAdmin
'If Login is already a member of the server role, exits without error
Public Sub AddLoginToSvrRole(dbSvr As dao.Database, _
                              ByVal sLoginName As String, _
                              ByVal sServerRoleName As String)
On Error GoTo EH
Dim sTSQL As String

If IsLoginInSvrRole(dbSvr, sLoginName, sServerRoleName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_addsrvrolemember '" & TSQLQuoteFix(sLoginName) & "','" & sServerRoleName & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "AddLoginToSvrRole"
End Sub

'Notes:  sServerRoleName must be built in server role (see AddLoginToSvrRole for list).
'If login is already server role member, exits early without error.
Public Sub DropLoginFromSvrRole(dbSvr As dao.Database, _
                                ByVal sLoginName As String, _
                                ByVal sServerRoleName As String)
On Error GoTo EH
Dim sTSQL As String

If Not IsLoginInSvrRole(dbSvr, sLoginName, sServerRoleName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_dropsrvrolemember '" & TSQLQuoteFix(sLoginName) & "','" & TSQLQuoteFix(sServerRoleName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, "master"

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DropLoginFromSvrRole"
End Sub

'NOTES:
'(1) Does not test if corresponding login exists
'(2) If user already in database, exits early
'(2) By default, assumes that UserID should be made the same as an existing LoginName.  Under ordinary
'WEM circumstances, this always should be true.
'(3) If matching login does not already exist, will error.
Public Sub AddUserToDB(dbSvr As dao.Database, _
                            ByVal sDBName As String, _
                            ByVal sUserName As String, _
                            Optional ByVal sLoginName As String = "")
On Error GoTo EH
Dim sTSQL As String

If sLoginName = "" Then
    sLoginName = sUserName
End If

If IsUserInDB(dbSvr, sDBName, sUserName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_grantdbaccess '" & TSQLQuoteFix(sLoginName) & "','" & TSQLQuoteFix(sUserName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "AddUserToDB"
End Sub

'NOTES:
'(1) Will currently fail if this user is the 'owner' of the database or any objects in the database.
'(in which case, should probably re-assign ownership to the dbo account.  may add this in the future.)
Public Sub DropUserFromDB(dbSvr As dao.Database, _
                               ByVal sDBName As String, _
                               ByVal sUserName As String)
On Error GoTo EH
Dim sTSQL As String

If Not IsUserInDB(dbSvr, sDBName, sUserName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_revokedbaccess '" & TSQLQuoteFix(sUserName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DropUserFromDB"
End Sub

'Note:  sDBRoleName can be any valid database role (fixed/built in or
'custom).  If user is already a member of role, exits early without error
Public Sub AddUserToDBRole(dbSvr As dao.Database, _
                           ByVal sDBName As String, _
                           ByVal sUserName As String, _
                           ByVal sDBRoleName As String)
On Error GoTo EH
Dim sTSQL As String

If IsUserInDBRole(dbSvr, sDBName, sUserName, sDBRoleName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_addrolemember '" & TSQLQuoteFix(sDBRoleName) & "','" & TSQLQuoteFix(sUserName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "AddUserToDBRole"
End Sub

'Exits without error is user is not member of DB role
Public Sub DropUserFromDBRole(dbSvr As dao.Database, _
                              ByVal sDBName As String, _
                              ByVal sUserName As String, _
                              ByVal sDBRoleName As String)
On Error GoTo EH
Dim sTSQL As String

If Not IsUserInDBRole(dbSvr, sDBName, sUserName, sDBRoleName) Then
    Exit Sub
End If

sTSQL = ""
sTSQL = sTSQL & "sp_droprolemember '" & TSQLQuoteFix(sDBRoleName) & "','" & TSQLQuoteFix(sUserName) & "'" & vbCrLf

ExecTSQL dbSvr, sTSQL, False, sDBName

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "DropUserFromDBRole"
End Sub

'returns UBound for an array.  If error occurs (either b'c arrTest is not an array or is empty), then
'returns -1.
Private Function UBoundOrMinus1(arrTest As Variant, Optional lDimensionPosition As Long = 1) As Long
On Error Resume Next
UBoundOrMinus1 = UBound(arrTest, lDimensionPosition)
If Err.Number <> 0 Then
  UBoundOrMinus1 = -1
End If
'don't return error to caller
Err.Clear
End Function


Public Function ExtractODBCParam(ByVal sODBCConnStr, ByVal sParamName As String) As String
On Error GoTo EH
Dim vPairs As Variant
Dim vPair As Variant
Dim vItems As Variant

'set default => param not found
ExtractODBCParam = ""

'break up the string at ";" delimiters
vPairs = Split(sODBCConnStr, ";")
'cycle through these
For Each vPair In vPairs
    'split again at the "=" sign
    vItems = Split(vPair, "=")
    'if single "=" is found, there will always be 2 items (vItems always has a valid UBound)
    If UBound(vItems) = 1 Then
        'if match with name of param
        '(have assumed ODBC driver trims whitespace from param names--not sure if this is true)
        If Trim(UCase(vItems(0))) = Trim(UCase(sParamName)) Then
            'return the second array item
            ExtractODBCParam = vItems(1)
            Exit For
        End If 'no match with param name => skip this one
    End If  'not a param=value structure => skip this one
Next vPair

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "ExtractODBCParam"
End Function

Public Sub ExtractODBCParams(ByVal sODBCConnStr As String, _
                             Optional sDSN_Rtn As String, _
                             Optional sSERVER_Rtn As String, _
                             Optional sDriver_Rtn As String, _
                             Optional sUID_Rtn As String, _
                             Optional sPWD_Rtn As String, _
                             Optional sDATABASE_Rtn As String, _
                             Optional sAPP_Rtn As String, _
                             Optional sWSID_Rtn As String)
On Error GoTo EH
Dim lIndex As Long
Dim vPairs As Variant
Dim vPair As Variant
Dim vItems As Variant

'set defaults => param not found
sDSN_Rtn = ""
sSERVER_Rtn = ""
sDriver_Rtn = ""
sUID_Rtn = ""
sPWD_Rtn = ""
sDATABASE_Rtn = ""
sAPP_Rtn = ""
sWSID_Rtn = ""

'break up the string at ";" delimiters
vPairs = Split(sODBCConnStr, ";")

'if no values
If UBound(vPairs) < 0 Then
    Exit Sub
End If

'cycle through these.  do in reverse order so 1st instance of any param is used (vPairs always has valid UBound)
For lIndex = UBound(vPairs) To 0 Step -1
    vPair = vPairs(lIndex)
    'split again at the "=" sign
    vItems = Split(vPair, "=")
    'if single "=" is found, there will be exactly 2 items (vItems always has a valid UBound)
    If UBound(vItems) = 1 Then
        'if match with name of param
        '(have assumed ODBC driver trims whitespace from param names--not sure if this is true)
        Select Case Trim(UCase(vItems(0)))
            Case "DSN"
                sDSN_Rtn = vItems(1)
            Case "SERVER"
                sSERVER_Rtn = vItems(1)
            Case "DRIVER"
                sDriver_Rtn = vItems(1)
            Case "UID"
                sUID_Rtn = vItems(1)
            Case "PWD"
                sPWD_Rtn = vItems(1)
            Case "DATABASE"
                sDATABASE_Rtn = vItems(1)
            Case "APP"
                sAPP_Rtn = vItems(1)
            Case "WSID"
                sWSID_Rtn = vItems(1)
        End Select
    End If  'not a param=value structure => skip this one
Next lIndex

Exit Sub
EH:
ErrHandlerPrivate False, "SQLToolkit", "ExtractODBCParams"
End Sub

'=============================================================================
'START functions related to resolving PC name into an IP address
'=============================================================================
'returns the IP address of the local PC
Public Function GetIPOfLocalPC() As String
   Dim sHostName As String * 256
   Dim sHostNameTrimmed As String
   'REA:  Based on testing, gethostname never succeeds and sHostName remains a null string.
   'This is apparently a valid flag for the local PC as GetIPFromHostName correctly returns the
   'IP of the local PC.  Passing the PC name from GetComputerName also works.  Not sure which is better,
   'so left original code as is.
   If gethostname(sHostName, 256) = ERROR_SUCCESS Then
      sHostNameTrimmed = Trim$(sHostName)
   End If

GetIPOfLocalPC = GetIPFromHostName(sHostNameTrimmed)
End Function

Private Function SocketsInitialize() As Boolean
   Dim WSAD As WSADATA
   Dim success As Long
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

Private Sub SocketsCleanup()
   If WSACleanup() <> 0 Then
       'MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
End Sub

Public Function GetIPFromHostName(ByVal sHostName As String) As String

  'converts a host name to an IP address

   Dim nbytes As Long
   Dim ptrHosent As Long  'address of HOSENT structure
   Dim ptrName As Long    'address of name pointer
   Dim ptrAddress As Long 'address of address pointer
   Dim ptrIPAddress As Long
   Dim ptrIPAddress2 As Long

   If Not SocketsInitialize() Then
       Exit Function
   End If

   ptrHosent = gethostbyname(sHostName & vbNullChar)

   If ptrHosent <> 0 Then

     'assign pointer addresses and offset

     'Null-terminated list of addresses for the host.
     'The Address is offset 12 bytes from the start of
     'the HOSENT structure. Note: Here we are retrieving
     'only the first address returned. To return more than
     'one, define sAddress as a string array and loop through
     'the 4-byte ptrIPAddress members returned. The last
     'item is a terminating null. All addresses are returned
     'in network byte order.
      ptrAddress = ptrHosent + 12
     
     'get the IP address
      CopyMemory ptrAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress, ByVal ptrAddress, 4
      CopyMemory ptrIPAddress2, ByVal ptrIPAddress, 4

      GetIPFromHostName = GetInetStrFromPtr(ptrIPAddress2)
   End If
   SocketsCleanup
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function GetInetStrFromPtr(Address As Long) As String
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))
End Function
'=============================================================================
'END functions related to resolving PC name into an IP address
'=============================================================================
'START functions for pinging IP adress from VB
'=============================================================================
'returns True if target exists.  Other params can optionally be returned.
Public Function PingIP(sAddress As String, Optional sMsgToSend As String = "", _
                 Optional sReturnStatus_Rtn As String, _
                 Optional lReturnAddrDec_Rtn As String, _
                 Optional lRoundTripMs_Rtn As Long, _
                 Optional lDataPacketBytes_Rtn As Long, _
                 Optional sMsgReturned_Rtn As String, _
                 Optional lDataPointer_Rtn As Long) As Boolean
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   'set default return values
   PingIP = False
   sReturnStatus_Rtn = ""
   lReturnAddrDec_Rtn = 0
   lRoundTripMs_Rtn = 0
   lDataPacketBytes_Rtn = 0
   sMsgReturned_Rtn = ""
   lDataPointer_Rtn = 0
   
   If sMsgToSend = "" Then
       'assign 32 chars of arbitrary data
       sMsgToSend = String(32, "T")
   End If
   
   If SocketsInitialize() Then
   
     'ping the IP by passing the address,
     'text to send, and the ECHO structure.
     
     'return results to caller.  if ping is successful, returns status = 0
      success = Ping_API(sAddress, sMsgToSend, ECHO)
      PingIP = (success = IP_SUCCESS)

      sReturnStatus_Rtn = GetStatusCode(success)
      lReturnAddrDec_Rtn = ECHO.Address
      lRoundTripMs_Rtn = ECHO.RoundTripTime
      lDataPacketBytes_Rtn = ECHO.DataSize
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         sMsgReturned_Rtn = Left$(ECHO.Data, pos - 1)
      End If
   
      lDataPointer_Rtn = ECHO.DataPointer
      
      SocketsCleanup
   
   End If
   
End Function

Private Function Ping_API(sAddress As String, _
                     sDataToSend As String, _
                     ECHO As ICMP_ECHO_REPLY) As Long

  'If Ping_API succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code
   
   Dim hPort As Long
   Dim dwAddress As Long
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)
   
  'if a valid address..
   If dwAddress <> INADDR_NONE Then
   
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
      
        'ping it.
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

        'return the status as ping succes and close
         Ping_API = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else
        'the address format was probably invalid
         Ping_API = INADDR_NONE
         
   End If
  
End Function

Private Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(status) & "   ( " & msg & " )"
   
End Function
'=============================================================================
'END functions related to pinging an IP address
'=============================================================================
'START functions related to converting an IP address into a host name
'=============================================================================
Public Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
   
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               GetHostNameFromIP = sAddress
            End If
         
         Else
            'MsgBox "Call to gethostbyaddr failed."
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else
         'MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   Else
      'MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function

'========================================================================================================
'GUID creation functions.  Adapted from http://www.vbaccelerator.com/home/VB/Tips/Create_GUIDs/article.asp.
Public Function GetNewGuidAsString(Optional bDropBraces As Boolean = True) As String
'On error resume next
Dim tG As GUID
Dim b(79) As Byte   'size based on storage for UNICODE = (40*2)-1
Dim lR As Long

'get a new GUID
CoCreateGuid tG

'convert to string.  return value is length of string including null terminator.
lR = StringFromGUID2(tG, VarPtr(b(0)), 40)

If bDropBraces Then
    '(-3 due to trimming termating zero)
    GetNewGuidAsString = Mid$(b, 2, lR - 3)
Else
    GetNewGuidAsString = Left$(b, lR - 1)
End If

Err.Clear
End Function
'========================================================================================================
'local versions of these functions so SQL Toolkit not dependent on outside utility modules.
Public Function GetFV(DAORecSetObj As dao.Recordset, sFieldName As String, _
                      Optional bConvertNulls As Boolean = True) As Variant
Dim vFieldValue As Variant
Dim ErrClone1 As ErrClone
Dim sSQL As String
'disable error handling--use line by line testing
On Error Resume Next

'get the field value
vFieldValue = DAORecSetObj.Fields(sFieldName).value
If Err.Number = 0 Then
   'if no error, return the field value (with Null conversion if necessary)
   GetFV = IIf(bConvertNulls And IsNull(vFieldValue), Empty, vFieldValue)
Else
   'save the error info
   ErrClone1 = MakeErrClone
   ' return "empty"
   GetFV = Empty
   'get the SQL used to create this recordset
   sSQL = DAORecSetObj.Name
   'turn VB error handling back on
   On Error GoTo EH
   'regenerate the error supplying additional info about the problem.  will return control to calling routine
   With ErrClone1
       Err.Raise .Number, .Source, .Description & vbCrLf & vbCrLf _
              & "[Field Name= '" & sFieldName & "']" & vbCrLf _
              & "[RecSet SQL (1st 256 chars) = '" & sSQL & "']", _
              .HelpFile, _
              .HelpContext
   End With
End If

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetFV"
End Function

Public Function SetFV(DAORecSetObj As dao.Recordset, sFieldName As String, vNewValue As Variant) As Boolean
Dim ErrClone1 As ErrClone
Dim sSQL As String

On Error Resume Next
'set default return value
SetFV = False
'attempt to set the field value
DAORecSetObj.Fields(sFieldName).value = vNewValue
'if error
If Err.Number <> 0 Then
   'save the error state
   ErrClone1 = MakeErrClone
   'get the SQL used to create this recordset
   sSQL = DAORecSetObj.Name
   'turn error handling on
   On Error GoTo EH
   'regenerate the error supplying additional info about the problem.  will return control to calling routine.
   With ErrClone1
       Err.Raise .Number, .Source, .Description & vbCrLf & vbCrLf _
              & "[Field Name= '" & sFieldName & "']" & vbCrLf _
              & "[RecSet SQL (1st 256 chars) = '" & sSQL & "']", _
              .HelpFile, _
              .HelpContext
   End With
End If
'return true
SetFV = True
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "SetFV"
End Function

'result returned by reference in sTableName.  Returns True if prefix found.
Public Function StripDboPrefix(sTableName As String) As Boolean
On Error GoTo EH
If UCase(Left(sTableName, 4)) = "DBO." Then
    sTableName = Right(sTableName, Len(sTableName) - 4)
    StripDboPrefix = True
End If
Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "StripDboPrefix"
End Function

'V1.18:
'Generates a "table version" string that will change if any data in a table is altered.
'Assumes that the table has a SQL Server RowVersion field named "SQLRowVersion".
'Returned string is the #TableRows:Max(SQLRowVersion) (as readable hex).
'The SQLRowVersion field detects any edits/adds.
'The #Rows detects any record deletions without an edit.
'Table structure changes are *not* detected.
Public Function GetTableVersionStr(dbSvr As dao.Database, sTableName As String) As String
On Error GoTo EH
Dim rs As dao.Recordset
Dim i1 As Long
Dim sTSQL As String
Dim sHex As String
Dim sRowVersion As String
Dim vBytes As Variant

sTSQL = "SELECT COUNT(*) AS NumRows, MAX(SQLRowVersion) AS MaxRowVer FROM [" & sTableName & "]"
Set rs = ExecTSQL(dbSvr, sTSQL, True)

'DAO will return the SQLRowVersion binary(8) field as an array of 8 bytes.
vBytes = rs.Fields("MaxRowVer").value

'To help with debugging, convert to readable hex string
For i1 = 0 To UBound(vBytes)
    sHex = sHex & Hex(vBytes(i1))
Next i1

'add row count prefix (essential as the SQLRowVersion is *not* changed if a row is deleted)
GetTableVersionStr = rs.Fields("NumRows").value & ":" & sHex

Exit Function
EH:
ErrHandlerPrivate False, "SQLToolkit", "GetTableVersionStr"
End Function


