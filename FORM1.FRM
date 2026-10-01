VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Auto Backup"
   ClientHeight    =   1335
   ClientLeft      =   1095
   ClientTop       =   1770
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1335
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   8175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8760
      Top             =   720
   End
   Begin VB.Label label_version 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu_execute 
      Caption         =   "Execute"
      Index           =   0
      Begin VB.Menu mnu_backup 
         Caption         =   "Backup"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get Version() As String
    '/* get the property value. */
    Version = label_version.Caption
End Property

Public Property Let Version(ByVal vNewValue As String)
    '/* set the property value. */
    label_version.Caption = vNewValue
End Property

Private Sub Form_Load()
Dim fnum
Dim i As Integer
Dim fs
On Error Resume Next
    
    MODULE11.AddItem Me, "WEM - Autoback" & vbNullChar
    Caption = "Autoback " & Format(App.Major, "0") & "." & Format(App.Minor, "0") & "." & Format(App.Revision, "0")
    LoadCIniGlobals
    olddate = ""
    Err = 0
    i = 0
    Maxcopys = -1
    fnum = FreeFile                                                 ' Determine file number.
    Open Install_Path & "\" & "copyfile.txt" For Input As fnum      ' Open file
    If Err = 0 Then
        Do While Not EOF(fnum)
            Input #fnum, cfile(i), tofile(i)
            If cfile(i) = "" Then Exit Do
            Maxcopys = i
            i = i + 1
            If i > 250 Then
                Exit Do
            End If
        Loop
        Close fnum  ' Close all files
    End If
    i = 0
    othercopys = -1
    fnum = FreeFile ' Determine file number.
    Err = 0
    If UCase(CFG_BACKUPTYPE) = "REDUNDANCY" Then
        Open Install_Path & "\" & "otherfile.txt" For Input As fnum ' Open file
        If Err = 0 Then
            Do While Not EOF(fnum)
                Input #fnum, Sourcefile(i), Destfile(i)
                If Sourcefile(i) = "" Then Exit Do
                othercopys = i
                i = i + 1
                If i > 750 Then
                    Exit Do
                End If
            Loop
            Close fnum  ' Close all files
        End If
    End If
    i = 0
    deletecopys = -1
    fnum = FreeFile ' Determine file number.
    Err = 0
    Open Install_Path & "\" & "deletefile.txt" For Input As fnum ' Open file
    If Err = 0 Then
        Do While Not EOF(fnum)
            Input #fnum, Deletefile(i)
            If Deletefile(i) = "" Then Exit Do
            deletecopys = i
            i = i + 1
            If i > 500 Then
                Exit Do
            End If
        Loop
        Close fnum  ' Close all files
    End If
    If CFG_BACKUP_AT_STARTUP = True Then
        gbl_dontchecktime = True
        backup_ctr = backup_ctr - 1
        If backup_ctr < 0 Then backup_ctr = 0
        nottime = 0
    End If
'    If Val(Left(Time$, 2)) > Val(BackupHour(0)) Then       '2.10 '2.18
'        gbl_dontchecktime = True
'        backup_ctr = backup_ctr - 1
'
'        'V2.20
'        If backup_ctr < 0 Then
'            backup_ctr = 0
'        End If
'
'        nottime = 0
'    End If
    Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

    '/* turn On Local error handling. */
    On Local Error GoTo ErrorHandler

    '/* declare local variables. */
    Dim l_ReturnedMessage As Long
    
    '/* define returned message. */
    l_ReturnedMessage = X / Screen.TwipsPerPixelX
    
    '/* determine which message was returned. */
    Select Case l_ReturnedMessage
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
            '/* display "Configuration" dialog. */
            'Configuration
            Form1.WindowState = 0
            Form1.Show
            'SysTray_Timer.Enabled = True

       Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
            '/* enable timer. */
       Case WM_RBUTTONDBLCLK
    End Select

    Exit Sub
    
ErrorHandler:

    '/* set mouse pointer to default. */
    Screen.MousePointer = 0

    '/* raise error message. */
    Err.Raise Err.Number, "autoback.Hidden", Err.Description

    Exit Sub


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)  '1.27
Dim ret As Integer
Dim WEMFile As Integer

    On Error Resume Next
    If UnloadMode = 0 Then  'user requested to shutdown
        ret = MsgBox("Are you sure?", vbYesNo + vbQuestion)
    ElseIf UnloadMode = 1 Then  'user requested to shutdown
'        RegInUse.UnRegisterTask
    '/* remove icon in system tray. */
        MODULE11.RemoveItem
        Exit Sub
    Else
        ret = vbYes
    End If
    If ret = vbYes Then
'        RegInUse.UnRegisterTask
    '/* remove icon in system tray. */
        MODULE11.RemoveItem
        End
    Else
        Cancel = True 'don't unload
    End If
End Sub

Private Sub Form_Resize()
    If Form1.WindowState = 1 Then
        Form1.Visible = False
    End If
End Sub


Sub mnu_backup_Click()
    Form1.Timer1.Enabled = False
    gbl_dontchecktime = True
    nottime = 0
    Call timer1_timer
End Sub

Private Sub mnu_exit_Click()
    Unload Form1
End Sub

Public Sub timer1_timer()
Dim i As Integer
Dim thehour As Integer
Dim daynum As Integer
Dim lResult As Long
Dim fs As FileSystemObject
Dim reply, itemp, temp$
Dim filetocopy As String
Dim sfile As Variant
Dim dfile As Variant

Dim src_created As FILETIME
Dim src_lastaccessed As FILETIME
Dim src_lastwritten As FILETIME
Dim dst_created As FILETIME
Dim dst_lastaccessed As FILETIME
Dim dst_lastwritten As FILETIME
Dim lpsecurity As SECURITY_ATTRIBUTES
Dim result As Long
Dim result2 As Long
Dim result3 As Long
Dim fh As Long
Dim localpath As String
Dim filename As String
Dim sFinalFullPath_RelApp As String '2.00
Dim Source As String
Dim Dest As String
Dim jobnumber As String
Dim cname As String
Dim fp As Integer
Dim COMPUTER As String
Dim jobname As String
Dim y
Dim X
Dim xMDBDirectory As String 'V2.20
Dim f_drive_space As Double     '2.27
Dim BackupAll As Boolean '2.30

    On Error Resume Next
    'On Error GoTo 0
    Timer1.Enabled = False
    BackupAll = False '2.30
    
    jobnumber = "Redundancy"
    cname = Taskname
    
    fp = InStr(1, cname, "FSERVER", vbTextCompare)
    If fp = 0 Then fp = InStr(1, cname, "WSERVER", vbTextCompare)
    If fp = 0 Then fp = InStr(1, cname, "CLIENT", vbTextCompare)
    If fp = 0 Then fp = InStr(1, cname, "PROMPT", vbTextCompare)
    If fp = 0 Then fp = InStr(1, cname, "UNRCV", vbTextCompare)
    If fp = 0 Then fp = InStr(1, cname, "BUS", vbTextCompare)       '2.26
    If fp = 0 Then
        COMPUTER = cname
        jobname = ""
    Else
        jobname = UCase(Mid(cname, 1, fp - 1))
        COMPUTER = UCase(Mid(cname, fp, Len(cname) - fp + 1))
    End If
    
    If gbl_dontchecktime = True Then
        g_oAppLog.AddEntry "Force Backup Execution"
        thehour = BackupHour(backup_ctr)        '2.21
        Form1.Timer1.Interval = 60000
        BackupAll = True '2.30 do all computers if forced or on power up
    Else
        Form1.Timer1.Interval = 60000
        thehour = Left$(Time$, 2)
        If thehour = 0 Then thehour = 24
    End If
    
    If thehour = BackupHour(backup_ctr) Then    '2.18
        If nottime = 0 Then 'back it up at appointed time
            f_drive_space = check_fix_destdrivespace("F:")      '2.27
            If f_drive_space < 1000000000 Then
                g_oAppLog.AddEntry "F Drive space < 1 GB.  Program will be terminated"
                End
            End If
            g_oAppLog.AddEntry "F Drive space = " & f_drive_space       '2.27
            g_oAppLog.AddEntry "Backup started at " & Now()
            nottime = 1     '2.16
            'backup_ctr = backup_ctr + 1
            If gbl_dontchecktime = True Then
                gbl_dontchecktime = False       '2.16
                thehour = Left$(Time$, 2)
                If thehour = 0 Then thehour = 24        '2.29
            End If
            Set fs = New FileSystemObject
            If InStr(1, MDBDirectory, COMPUTER, vbTextCompare) > 0 Then  '2.14 only the fserver will do make the backup dirs.
                
            'If InStr(1, COMPUTER, "fserver", vbTextCompare) > 0 Then
                g_oAppLog.AddEntry "Fserver Backing up data"
                daynum = Weekday(Now)
                Err = 0
                ChDir MDBDirectory & "\" & Format(daynum, "#")
                If Err <> 0 Then
                    MkDir MDBDirectory & "\" & Format(daynum, "#")
                    ChDir MDBDirectory & "\" & Format(daynum, "#")
                End If
            ElseIf InStr(1, COMPUTER, "BUS", vbTextCompare) > 0 Then        '2.26
                g_oAppLog.AddEntry "BUS Backing up data"
                daynum = Weekday(Now)
                Err = 0
                ChDir "F:\WEM\DATA\" & Format(daynum, "#")
                If Err <> 0 Then
                    MkDir MDBDirectory & "\" & Format(daynum, "#")
                    ChDir MDBDirectory & "\" & Format(daynum, "#")
                End If
            End If
            
            
            '2.00
            g_oAppLog.AddEntry "Deleting data based on deletecopy.txt"
            For i = 0 To deletecopys            '2.01.0
                
                DoEvents
                Form1.Text1 = "Deleting " & Deletefile(i)
                Form1.Refresh
                If InStr(1, Deletefile(i), "*", vbTextCompare) > 0 Then     '2.11
                    Dim fname$
                    fname$ = Dir(Deletefile(i))
                    If Len(fname$) > 0 Then
                        fs.Deletefile Deletefile(i), True
                    Else
                        If Len(fname$) = 0 Then fname$ = Dir(Deletefile(i), vbDirectory)
                        If Len(fname$) = 0 Then fname$ = Dir(Deletefile(i), vbDirectory Or vbHidden)
                        If Len(fname$) = 0 Then fname$ = Dir(Deletefile(i), vbDirectory Or vbHidden Or vbReadOnly)
                    End If
                    If Len(fname$) > 0 Then
                        fs.DeleteFolder Deletefile(i), True
                    End If
                Else
                    If fs.FileExists(Deletefile(i)) = True Then
                        fs.Deletefile Deletefile(i), True
                    ElseIf fs.FolderExists(Deletefile(i)) = True Then
                        fs.DeleteFolder Deletefile(i), True
                    End If
                End If                                                      'end 2.11
            Next i
            
            'On Error GoTo 0
    
            
            If cfg_UseSQL > 0 Then
                If InStr(1, MDBDirectory, COMPUTER, vbTextCompare) > 0 Or InStr(1, COMPUTER, "BUS", vbTextCompare) > 0 Then '2.26 only the fserver will do the SQL .sqb file creation
                    g_oAppLog.AddEntry "backing up sql database on fserver or bus"
            
                    'V2.20
                    If cfg_System = stUnknown Then
                        Form1.Text1 = "Backing up SQL db"
                    ElseIf cfg_System = stAsphalt Then
                        Form1.Text1 = "Backing up WEMASQL"
                    Else
                        Form1.Text1 = "Backing up WEMSQL"
                    End If
                    
                    DoEvents
                    Set db = oSQLTK.OpenDBSvr(cfg_SQLMAIN_DSN, cfg_SQLMAIN_OverrideSvrName, _
                        cfg_SQLMAIN_WEMAdminUserID, cfg_SQLMAIN_WEMAdminPWD, cfg_SQLMAIN_OverrideDBName)
                    If Len(cfg_SQLMAIN_ARCHIVE_DSN) > 0 Then
                        Set dbArchive = oSQLTK.OpenDBSvr(cfg_SQLMAIN_ARCHIVE_DSN, cfg_SQLMAIN_OverrideSvrName, _
                                        cfg_SQLMAIN_WEMAdminUserID, cfg_SQLMAIN_WEMAdminPWD, cfg_SQLMAIN_OverrideDBName)    '2.46
                    End If
'                    RegInUse.table = "*UPDATE*"  '2.00
'                    RegInUse.LockAll = True
'                    Do While True
'                        RegInUse.TableLock
'                        If RegInUse.TableStatus = 0 Then
'                            Exit Do
'                        End If
'                        Form1.Text1 = "Waiting for " & RegInUse.LockedTaskname
'                        Form1.Refresh
'                        DoEvents
'                        g_oAppLog.AddEntry "Waiting for " & RegInUse.LockedTaskname
'                        Sleep 1000
'                    Loop
                    
                    'V2.20 Asphalt SQL correction for SQL DB location (F:\WEM\DATA\SQL) not F:|WEM\DATA\DATABASE\SQL)
                    If Right(MDBDirectory, 9) = "\DATABASE" Then
                        xMDBDirectory = Left(MDBDirectory, Len(MDBDirectory) - 9)
                    ElseIf InStr(1, COMPUTER, "BUS", vbTextCompare) > 0 Then
                        xMDBDirectory = "F:\WEM\DATA\"
                    Else
                        xMDBDirectory = MDBDirectory
                    End If
                    '2.00 remove any old backup zip files from the directory we are backing up to
                    fname = Dir(xMDBDirectory & "\" & Format(daynum, "#") & "\*.zip")
                    Do Until Len(fname) = 0
                        Dim tf As File
                        Set tf = fs.GetFile(xMDBDirectory & "\" & Format(daynum, "#") & "\" & fname)
                        If InStr(1, fname, "__archive", vbTextCompare) > 0 Then
                        Else
                        
                            If DateDiff("d", Now(), tf.DateCreated) <> 0 Then
                                g_oAppLog.AddEntry "Deleting all backup zip files from " & Format(daynum, "#")
                                Kill xMDBDirectory & "\" & Format(daynum, "#") & "\*.zip"
                            End If
                        End If
'                        If InStr(1, tf.Name, "ARCHIVE", vbTextCompare) > 0 Then
'
'                            fs.Deletefile (tf.Path)
'                        End If
                        fname = Dir()
                    Loop
                    'Kill xMDBDirectory & "\" & Format(daynum, "#") & "\*.zip"
                    g_oAppLog.AddEntry "making wemsql backup"

                    sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(db, , , xMDBDirectory & "\" & Format(daynum, "#"), BACKUP_STATUS_Consistent, , sql_BaseFolder)
                    'make sure the compressed backup also lands in the day folder
                    If Len(sFinalFullPath_RelApp) > 0 Then
                        If fs.FileExists(sFinalFullPath_RelApp) Then
                            If UCase(fs.GetParentFolderName(sFinalFullPath_RelApp)) <> UCase(xMDBDirectory & "\" & Format(daynum, "#")) Then
                                g_oAppLog.AddEntry "copying backup " & sFinalFullPath_RelApp & " to " & xMDBDirectory & "\" & Format(daynum, "#")
                                fs.CopyFile sFinalFullPath_RelApp, xMDBDirectory & "\" & Format(daynum, "#") & "\" & fs.GetFileName(sFinalFullPath_RelApp), True
                            End If
                        End If
                    End If
                    If UCase(CFG_BACKUPTYPE) = "DCOLLECT" Then              '2.02
                        'Kill xMDBDirectory & "\SQL\*.zip"
                        MODULE11.del_old_sql_backups (xMDBDirectory & "\SQL")
                        sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(db, , , xMDBDirectory & "\SQL", BACKUP_STATUS_Consistent, , sql_BaseFolder)
                        
                    End If
'''                    If Len(cfg_SQLMAIN_ARCHIVE_DSN) > 0 Then
'''                        Me.Text1.Text = "BACKING UP SQL ARCHIVE"
'''                        Me.Refresh
'''                        sFinalFullPath_RelApp = ""
'''                        g_oAppLog.AddEntry "checking wemsqlarchive backup"
'''                        Dim rslt As Boolean
'''                        rslt = makeorcopyarchive(xMDBDirectory & "\sql", xMDBDirectory & "\" & Format(daynum, "#"))
'''
''''                        sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(dbArchive, , , xMDBDirectory & "\" & Format(daynum, "#"), BACKUP_STATUS_Consistent, , sql_BaseFolder)
''''                        If UCase(CFG_BACKUPTYPE) = "DCOLLECT" Then              '2.02
''''                            Me.Text1.Text = "BACKING UP SQL ARCHIVE"
''''                            Me.Refresh
''''                            sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(dbArchive, , , xMDBDirectory & "\SQL", BACKUP_STATUS_Consistent, , sql_BaseFolder)
''''                        End If
'''                    End If
                    
'                    RegInUse.table = ""     '2.10
'                    RegInUse.TableUnlock    '2.00
'                    RegInUse.UnRegisterTask '2.10
                    db.Close
                    Sleep 1000             '2.15
                Else
                    'if not fserver, delay before attempting to copy all the files from the fserver
                    Form1.Text1 = "Waiting to start copy"
                    Form1.Refresh
                    Sleep (5000)
                End If
                Form1.Text1 = "Preparing next file to copy"
                Form1.Refresh
                Sleep (2000)
            Else
                g_oAppLog.AddEntry "copying data based on copyfile.txt"
                
                'This is only used with copyfile.txt for feed, no need to change for asphalt.
                For i = 0 To Maxcopys
                    Form1.Text1 = "Backing up WEM.MDB"
                    DoEvents
'                    RegInUse.table = "*UPDATE*"  '2.00
'                    RegInUse.LockAll = True
'                    Do While True
'                        RegInUse.TableLock
'                        If RegInUse.TableStatus = 0 Then
'                            Exit Do
'                        End If
'                        Form1.Text1 = "Waiting for " & RegInUse.LockedTaskname
'                        Form1.Refresh
'                        DoEvents
'                        g_oAppLog.AddEntry "Waiting for " & RegInUse.LockedTaskname
'                        Sleep 1000
'                    Loop
                    Form1.Text1 = "Backing up WEM.MDB"
                    Form1.Refresh
                    DoEvents
    
                    If InStr(1, COMPUTER, "fserver", vbTextCompare) > 0 Then        '2.12.0
                        If InStr(cfile(i), "\") > 0 Then 'if we specified the whole path and file name, use it
'                            g_oAppLog.AddEntry cfile(i) & " File Copy Started"
                            lResult = CopyFile(cfile(i), MDBDirectory & "\" & Format(daynum, "#") & "\" & tofile(i), 0&)  'make a copy of the file
'                            g_oAppLog.AddEntry cfile(i) & " File Copy Finished"
                            Debug.Print "hello"
                        Else
'                            g_oAppLog.AddEntry cfile(i) & " File Copy Started"
                            lResult = CopyFile(MDBDirectory & "\" & cfile(i), MDBDirectory & "\" & Format(daynum, "#") & "\" & tofile(i), 0&) 'make a copy of the file
'                            g_oAppLog.AddEntry cfile(i) & " File Copy Finished"
                            Debug.Print "hello"
                        End If
                    
                        If lResult = 0 Then
                            Text1.Text = "'" & cfile(i) & "' copy error."
                            g_oAppLog.AddEntry cfile(i) & " error"
                            Form1.Refresh
                        End If
                    End If
'                    RegInUse.table = ""     '2.10
'                    RegInUse.TableUnlock    '2.00
'                    RegInUse.UnRegisterTask '2.10
                    Form1.Text1 = "Preparing next file to copy"
                    Form1.Refresh
                    Sleep (5000)
                Next i
            End If
            
            '2.30
            Dim fp2 As Integer
            Dim DoBackup As Boolean
            DoBackup = True
'            If BackupAll = False Then   'if periodic update, not forced or first time
'                fp2 = InStr(1, cname, "CLIENT", vbTextCompare)
'                If fp2 > 0 Then
'                    fp2 = Val(Right(cname, Len(cname) - (fp2 + 5))) 'get client number
'                    If fp2 >= 2 Then 'don't backup clients 2-n every hour
'                        If (fp2 + 5) <> thehour Then
'                            DoBackup = False 'backup client2 at 7:00, 3 at 8:00  etc.
'                        Else
'                            g_oAppLog.AddEntry "backup " & cname
'                        End If
'                    End If
'                End If
'            End If
            
            If UCase(CFG_BACKUPTYPE) = "DCOLLECT" Then  '2.07
                If DoBackup = True Then '2.30
                    g_oAppLog.AddEntry "Copy config based on dcollection.mdb"
                    If Len(jobname) > 0 Then
                    
                        Form1.Text1.Text = "Checking " & "\\" & jobname & "FSERVER\WINNT\"  '2.30
                        DoEvents
                        filename$ = Dir("\\" & jobname & "FSERVER\WINNT\", vbDirectory)
                        If Len(filename$) > 0 Then
                            MODULE11.GetDCollectFiles jobname, "FSERVER"
                            If InStr(1, cname, "BUS", vbTextCompare) > 0 Then copy_update_folder jobname        '2.26
                        End If
                        Form1.Text1.Text = "Checking " & "\\" & jobname & "WSERVER\WINNT\"  '2.30
                        DoEvents
                        filename$ = ""
                        filename$ = Dir("\\" & jobname & "WSERVER\WINNT\", vbDirectory)
                        If Len(filename$) > 0 Then
                            MODULE11.GetDCollectFiles jobname, "WSERVER"
                        End If
                        Form1.Text1.Text = "Checking " & "\\" & jobname & "UNRCV\WINNT\"  '2.30
                        DoEvents
                        DoEvents                        '2.15
                        filename$ = ""
                        filename$ = Dir("\\" & jobname & "UNRCV\WINNT\", vbDirectory)
                        If Len(filename$) > 0 Then
                            MODULE11.GetDCollectFiles jobname, "UNRCV"
                        End If
                        DoEvents                        '2.15
                        For y = 1 To 15
                            Form1.Text1.Text = "Checking " & "\\" & jobname & "CLIENT" & y & "\WINNT"  '2.30
                            DoEvents
                            filename$ = ""
                            If fs.FolderExists("\\" & jobname & "CLIENT" & y & "\WINNT") = True Then
                                MODULE11.GetDCollectFiles jobname, "CLIENT" & y
                            End If
                            DoEvents                    '2.15
                        Next
                    End If
                    For X = 1 To 20                 '2.15 2.31
                        If Len(PCNAME(X)) > 0 Then
                            Form1.Text1.Text = "Checking " & jobname & " " & PCNAME(X)    '2.30
                            DoEvents
                            MODULE11.GetDCollectFiles jobname, PCNAME(X)
                        End If
                        DoEvents                    '2.15
                    Next
                End If
            Else
                g_oAppLog.AddEntry "copying data based on otherfile.txt"
            
                For i = 0 To othercopys
                    Form1.Text1 = "Looking for Files to Backup"
                    Form1.Refresh
                    Source = Sourcefile(i)      '2.00
                    Dest = Destfile(i)
                    If (GetAttr(Source) And vbReadOnly) = vbReadOnly Then SetAttr Source, vbArchive
                    If (GetAttr(Dest) And vbReadOnly) = vbReadOnly Then SetAttr Dest, vbArchive
                    DoEvents
                    itemp = InStr(1, Source, "*.", 1)
                    If itemp > 0 Then
                        filetocopy = ""                     '1.7.0
                        filetocopy = Dir(Source)
                        temp$ = Mid(Source, 1, itemp - 1)
                        While Len(filetocopy) > 0
                            g_oAppLog.AddEntry Dest & " File Compare"
                            Form1.Text1.Text = "Analyzing files " & Dest
                            Form1.Refresh
                            DoEvents
                            fh = CreateFile(temp$ & filetocopy, 0, FILE_SHARE_READ, lpsecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, FILE_FLAG_RANDOM_ACCESS)
                            If fh > 0 Then
                                result = GetFileTime(fh, src_created, src_lastaccessed, src_lastwritten)
                                If result > 0 Then
                                    result = CloseHandle(fh)
                                    fh = CreateFile(Dest & "\" & filetocopy, 0, FILE_SHARE_READ, lpsecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, FILE_FLAG_RANDOM_ACCESS)
                                    If fh > 0 Then
                                        result = GetFileTime(fh, dst_created, dst_lastaccessed, dst_lastwritten)
                                        result = CloseHandle(fh)
                                        result = CompareFileTime(src_created, dst_created)
                                        result3 = CompareFileTime(src_lastwritten, dst_lastwritten)
                                        If (result3 > 0) Or (result2 > 0) And (result > 0) Then
                                            Form1.Text1.Text = "Copying files to " & Dest
                                            Form1.Refresh
'                                            g_oAppLog.AddEntry filetocopy & " File Copy Started"
                                            fs.CopyFile temp$ & filetocopy, Dest & "\" & filetocopy
'                                            g_oAppLog.AddEntry filetocopy & " File Copy Completed"
                                            Form1.Text1 = "Preparing next file to copy"
                                            Form1.Refresh
'                                            g_oAppLog.AddEntry filetocopy & " File Copy Completed"
                                            Sleep (5000)
                                        End If
                                    Else
                                        Form1.Text1.Text = "Copying files to " & Dest
                                        Form1.Refresh
'                                        g_oAppLog.AddEntry filetocopy & " File Copy Started"
                                        fs.CopyFile temp$ & filetocopy, Dest & "\" & filetocopy
'                                        g_oAppLog.AddEntry filetocopy & " File Copy Completed"
                                        Form1.Text1 = "Preparing next file to copy"
                                        Form1.Refresh
                                        Sleep (5000)
                
                                    End If
                                Else
                                    result = CloseHandle(fh)
                                End If
                            Else
                                result = CloseHandle(fh)
                            End If
                            filetocopy = "" '1.26 if the dir fails for some reason, this will make sure we jsut kick out and we'll do it again next hour
                            filetocopy = Dir
                        Wend
                    Else
                        Source = Sourcefile(i)  '2.00
                        Dest = Destfile(i)
                        filetocopy = dfile = sfile = ""                     '1.7.0
                        filetocopy = Dir(Source)
                        '2.00
                        If cfg_UseSQL > 0 Then  'doing the SQL shuffle?
                            'V2.20 If InStr(UCase(Source), "WEM.MDB") > 0 Then 'want to copy the data?
                            If InStr(UCase(Source), "WEM.MDB") > 0 Or InStr(UCase(Source), "WEMA.MDB") > 0 Then 'want to copy the data?
                                If COMPUTER <> "FSERVER" Then   'if we are not the fserver, then we want to copy the backup zip file
                                                                'but only when it is ready, so we will check the register to make sure we can get in
'                                    RegInUse.table = "*UPDATE*"  '2.00
'                                    RegInUse.LockAll = True
'                                    Do While True
'                                        RegInUse.TableLock
'                                        If RegInUse.TableStatus = 0 Then
'                                            Exit Do
'                                        End If
'                                        Form1.Text1 = "Waiting for " & RegInUse.LockedTaskname
'                                        Form1.Refresh
'                                        DoEvents
'                                        g_oAppLog.AddEntry "Waiting for " & RegInUse.LockedTaskname
'                                        Sleep 1000
'                                    Loop
'                                    RegInUse.table = ""     '2.10
'                                    RegInUse.TableUnlock    '2.00
'                                    RegInUse.UnRegisterTask '2.10
                                End If
                            
                                'V2.20
                                If Right(MDBDirectory, 9) = "\DATABASE" Then
                                    xMDBDirectory = Left(MDBDirectory, Len(MDBDirectory) - 9)
                                    Source = xMDBDirectory & "\" & Format(daynum, "#") & "\*.zip"
                                    filetocopy = Dir(Source)
                                    Source = xMDBDirectory & "\" & Format(daynum, "#") & "\" & filetocopy
                                    Dest = Replace(UCase(Destfile(i)), "DATABASE\WEMA.MDB", "*.zip")
                                    Kill Dest       'get rid of any other wem backup file; we only want the last one
                                    Dest = Replace(UCase(Destfile(i)), "DATABASE\WEMA.MDB", filetocopy)
                                Else
                                
                                    Source = MDBDirectory & "\" & Format(daynum, "#") & "\*.zip"
                                    filetocopy = Dir(Source)
                                    Source = MDBDirectory & "\" & Format(daynum, "#") & "\" & filetocopy
                                    Dest = Replace(UCase(Destfile(i)), "WEM.MDB", "*.zip")
                                    Kill Dest       'get rid of any other wem backup file; we only want the last one
                                    Dest = Replace(UCase(Destfile(i)), "WEM.MDB", filetocopy)
                                
                                End If
                            
                            End If
                        End If
                        If Len(filetocopy) > 0 Then
                            Form1.Text1.Text = "Analyzing file " & filetocopy
                            Form1.Refresh
                            fh = CreateFile(Source, 0, FILE_SHARE_READ, lpsecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, FILE_FLAG_RANDOM_ACCESS)
                            If fh > 0 Then
                                result = GetFileTime(fh, src_created, src_lastaccessed, src_lastwritten)
                                If result > 0 Then
                                    result = CloseHandle(fh)
                                    fh = CreateFile(Dest, 0, FILE_SHARE_READ, lpsecurity, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, FILE_FLAG_RANDOM_ACCESS)
                                    If fh > 0 Then
                                        result = GetFileTime(fh, dst_created, dst_lastaccessed, dst_lastwritten)
                                        result = CloseHandle(fh)
                                        result = CompareFileTime(src_created, dst_created)
                                        result3 = CompareFileTime(src_lastwritten, dst_lastwritten)
                                        If (result3 > 0) Or (result2 > 0) And (result > 0) Then
                                            Form1.Text1.Text = "Copying files to " & Dest
                                            Form1.Refresh
'                                            g_oAppLog.AddEntry filetocopy & " File Copy Started"
                                            fs.CopyFile Source, Dest
'                                            g_oAppLog.AddEntry filetocopy & " File Copy Completed"
                                            Form1.Text1 = "Preparing next file to copy"
                                            Form1.Refresh
                                            Sleep (5000)
                                        End If
                                    Else
                                        Form1.Text1.Text = "Copying files to " & Dest
                                        Form1.Refresh
'                                        g_oAppLog.AddEntry Source & " File Copy Started"
                                        fs.CopyFile Source, Dest
'                                        g_oAppLog.AddEntry Source & " File Copy Completed"
                                        Form1.Text1 = "Preparing next file to copy"
                                        Form1.Refresh
                                        Sleep (5000)
                                    End If
                                Else
                                    result = CloseHandle(fh)
                                End If
                            Else
                                result = CloseHandle(fh)
                            End If
                            
                        End If
                    End If
                Next i
            End If
            nottime = 1
            Form1.Text1.Text = " "
            Form1.Timer1.Interval = 1000
            find_next_backup_time
'            If BackupHour(backup_ctr) = 0 Then         '2.29
'                If cfg_resetonceperday = True Then          '2.29
'                    Unload Form1
'                    Main
'                End If
'            End If
'
'            If BackupHour(backup_ctr) < 12 Then     '2.21
'                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " A.M."
'            ElseIf BackupHour(backup_ctr) = 12 Then '2.29
'                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " P.M."
'            ElseIf BackupHour(backup_ctr) = 24 Then '2.29
'                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " A.M."
'            Else
'                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " P.M."
'            End If
        Else
            g_oAppLog.AddEntry "Skipping Backup! We've already done it."
            backup_ctr = backup_ctr + 1
            If Val(BackupHour(backup_ctr)) = 0 Then         '2.29
                If cfg_resetonceperday = True Then          '2.29
                    Unload Form1
                    Main
                End If
            End If
            
            If BackupHour(backup_ctr) = 0 Then         '2.29
                If cfg_resetonceperday = True Then          '2.29
                    Unload Form1
                    Main
                End If
            End If
            
           
            chk_valid_backup_time (backup_ctr)         '2.32
            If BackupHour(backup_ctr) < 12 Then     '2.21
                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " A.M."
            ElseIf BackupHour(backup_ctr) = 12 Then '2.29
                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " P.M."
            ElseIf BackupHour(backup_ctr) = 24 Then '2.29
                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " A.M."
            Else
                Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " P.M."
            End If
            nottime = 0
            Form1.Timer1.Interval = 60000
        End If
    Else
        If Val(BackupHour(backup_ctr)) = 0 Then
            backup_ctr = 0      '2.21
        End If
        nottime = 0             '2.16
        'chk_valid_backup_time (backup_ctr)         '2.32
            
        If BackupHour(backup_ctr) < 12 Then     '2.21
            Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " A.M."
        ElseIf BackupHour(backup_ctr) = 12 Then '2.29
            Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) & " P.M."
        ElseIf BackupHour(backup_ctr) = 24 Then '2.29
            Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " A.M."
        Else
            Form1.Text1 = "Next Scheduled Backup is " & BackupHour(backup_ctr) - 12 & " P.M."
        End If
'        If thehour < 1 Or thehour > 24 Then     '2.29
'            backup_ctr = 0       '2.18
'
'        End If
        Form1.Timer1.Interval = 60000
    End If
    Timer1.Enabled = True   '2.16
End Sub
Function makeorcopyarchive(src As String, dst As String) As Boolean
Dim fso As New FileSystemObject
Dim fname As String
Dim sFinalFullPath_RelApp
Dim sf As File
fname = Dir(src & "\*.zip")
'MsgBox "src = " & src & ",dst=" & dst

Do Until Len(fname) < 3
    If InStr(1, fname, "__archive") > 0 Then
        Set sf = fso.GetFile(src & "\" & fname)
        If DateDiff("d", sf.DateCreated, Now()) > 7 Then
            
            g_oAppLog.AddEntry "deleting archive file because older than 7 days"
            fso.Deletefile (src & "\" & fname)
        Else
            If src <> dst Then
                g_oAppLog.AddEntry "copy " & fname & " to " & dst
                fso.CopyFile src & "\" & fname, dst & "\"
            End If
            makeorcopyarchive = True
            Exit Function
        End If
    End If
    fname = Dir()
Loop
g_oAppLog.AddEntry "call sql backup for archive"
sFinalFullPath_RelApp = oSQLTK.BackupDBToFile2(dbArchive, , , src, BACKUP_STATUS_Consistent, , sql_BaseFolder)
'MsgBox "just made archive"
If src <> dst Then
    g_oAppLog.AddEntry "copy " & fname & " to " & dst
    fso.CopyFile sFinalFullPath_RelApp, dst
End If

End Function
Function check_fix_destdrivespace(driveleter As String) As Double       '2.27
Dim fso As New FileSystemObject
Dim tmp_drive As Drive

g_oAppLog.AddEntry "chk_fix_destdrivespeac"
Set tmp_drive = fso.GetDrive(driveleter)
check_fix_destdrivespace = tmp_drive.AvailableSpace
If check_fix_destdrivespace < 1000000000 Then
    Call cleanup_sql_tempfiles
    Call cleanup_udl_files
    Call cleanup_old_files
    check_fix_destdrivespace = tmp_drive.AvailableSpace
End If
End Function

Sub cleanup_sql_tempfiles()     '2.27
Dim fso As New FileSystemObject
On Error Resume Next
    g_oAppLog.AddEntry "cleanup_sql_tempfiles"
    fso.Deletefile ("f:\wem\data\temp\*.*")
    fso.Deletefile MDBDirectory & "\temp\*.*"
    fso.Deletefile ("f:\redundancy\fserver\wemdata\wem\data\temp\*.*")

End Sub

Sub cleanup_udl_files()         '2.27
Dim fso As New FileSystemObject
On Error Resume Next
    g_oAppLog.AddEntry "cleanup_udl_files"
    fso.Deletefile ("f:\wem\data\udl\*.*")
    fso.Deletefile ("f:\redundancy\fserver\wemdata\wem\data\udl\*.*")

End Sub

Sub cleanup_old_files()             '2.27
Dim tmp_files As Collection
Dim colDataPaths As New Collection
Dim colDataParentFolders As New Collection
Dim colHoldPaths As New Collection
Dim lngTopIndex As Long
Dim lngPathIndex As Long
Dim lngIndex As Long
ReDim strpaths(0) As String

Dim fsoFSO As New FileSystemObject
Dim fldFolder As Folder, fldFolders As Folders
Dim filFile As File, filFiles As Files
On Error Resume Next

    g_oAppLog.AddEntry "cleanup_old_files"

    lngTopIndex = 0
    lngPathIndex = 0
    strpaths(0) = "f:\redundancy\"
    If fsoFSO.FolderExists("f:\redundancy") = False Then
        strpaths(0) = "C:\WEMDATA\redundancy\"
        If fsoFSO.FolderExists("C:\WEMDATA\redundancy") = False Then
            Exit Sub
        End If
    End If
    
    Do
        Set fldFolders = fsoFSO.GetFolder(strpaths(lngPathIndex)).SubFolders
        For Each fldFolder In fldFolders
            lngTopIndex = lngTopIndex + 1
            ReDim Preserve strpaths(lngTopIndex)
            strpaths(lngTopIndex) = strpaths(lngPathIndex) & fldFolder.Name & "\"
            Set filFiles = fsoFSO.GetFolder(strpaths(lngTopIndex)).Files
            For Each filFile In filFiles
                colDataPaths.Add filFile.Path
                colDataParentFolders.Add fldFolder.ParentFolder
            Next
        Next
        lngPathIndex = lngPathIndex + 1
    Loop Until lngPathIndex > lngTopIndex
    
    If colDataPaths.Count > 0 Then
        For lngIndex = 1 To colDataPaths.Count
            If InStr(1, colDataPaths.Item(lngIndex), "_OLD_", vbTextCompare) > 0 Then
                fsoFSO.Deletefile colDataPaths.Item(lngIndex)
            End If
        Next
        For lngIndex = 1 To lngPathIndex - 1
            If InStr(1, strpaths(lngIndex), "_OLD", vbTextCompare) > 0 Then
                fsoFSO.DeleteFolder Mid(strpaths(lngIndex), 1, Len(strpaths(lngIndex)) - 1), True
            End If
        Next
    End If
    

End Sub
Sub find_next_backup_time()
Dim X
Dim tmp_backuphour
Dim thehour As Integer
thehour = Left$(Time$, 2)
    
backup_ctr = 0
X = 0
    Do Until BackupHour(X) = 0
        If thehour < BackupHour(X) Then            '2.18
            backup_ctr = X
            Exit Sub
        End If
        X = X + 1
    Loop
End Sub

Sub chk_valid_backup_time(ByVal backupindex As Integer)
Dim thehour As Integer
     On Error Resume Next
    'BackupHour (backup_ctr)
    thehour = Left$(Time$, 2)
    
    Do Until BackupHour(backupindex + 1) = 0
        If thehour <= BackupHour(backupindex) And thehour > BackupHour(backupindex + 1) Then
        
        ElseIf thehour > BackupHour(backupindex) And thehour < BackupHour(backupindex + 1) Then
            backup_ctr = backupindex + 1
            Exit Sub
        Else
            backupindex = backupindex + 1
        End If
    Loop
    backup_ctr = 0
    Exit Sub
'    If thehour = 0 Then thehour = 24
'    If backupindex > 0 Then
'        If BackupHour(backupindex - 1) <= thehour And BackupHour(backupindex) > thehour Then        'this is correct we are waiting at the correct spot
'            Exit Sub
'        End If
'    Else
'        If BackupHour(0) <= thehour And BackupHour(1) > thehour Then       'this is correct we are waiting at the correct spot
'            Exit Sub
'        End If
'    End If
'    backupindex = 1
'    Do Until (BackupHour(backupindex - 1) <= thehour And BackupHour(backupindex) > thehour) Or BackupHour(backupindex) = 0
'        backupindex = backupindex + 1
'    Loop
'    backup_ctr = backupindex
'    If BackupHour(backup_ctr) = 0 Then backup_ctr = 0
        
        
    
End Sub
            
        
