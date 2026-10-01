Attribute VB_Name = "Revisions"
'Revisions
'Version    Date        Description
'1.0.1      03/01/01    Added other file copying using source and destination
'1.1.0      03/01/01    Made into standard
'1.2.0      06/12/01    Corrected redundant copying once/minute to once/hour
'1.3.0      12/04/01    Added file date checking and backward copying for otherfile functionality.
'1.6.0      01/14/02    Corrected Otherfile functionality and data checking
'1.7.0      01/18/02    Fixed bug which would cause the system to lock up if the computer copying is not on.
'1.9.0      02/04/02    Modified date/time format to truncate off seconds which causes a false negative alarm (microsoft bug)
'1.10.0     02/12/02    Fixed so if destination file does not exist it does not generate a file descrpency
'1.14.0     06/12/02    Corrected for time zone mess embedded in files
'1.16.0     09/12/02    Performed read only file checking and correction
'1.17.0     10/02/02    Performed MORE read only file checking and correction
'1.18.0     10/28/02    Used api calls to perform file date/time checking
'1.20.0     02/20/03    Removed date/time checking and updated file regardless
'1.21.0     02/27/03    Added a menu item for forcing the program to perform it's backup
'1.22.0     02/27/03    Added a deletefile.txt reading and execution to delete ww log files
'1.23.0     02/27/03    Implemented a routine to look once an hour to make sure that the database has been copied within the current day
'                        Implemented a file compare routine before each copy
'1.24.0     07/17/04    Implemented a routine to look once an hour to make sure that the database has been copied within the current day
'1.25.0     08/09/04    Added more logging to help find autoback not responding errors.
'1.26.0     08/18/04    Set the filename to "" before the dir command so that if logging resets the dir(), we won't get
'                       stuck in a loop. Logging should be changed in the future to NOT use the dir command.
'1.27.0     08/23/04    New applog class module
'1.28.0     03/03/05    Change from unload to unloadquery so it shutsdown nicely
'2.00.0     08/03/05    VB 6 and SQL
'2.01.0     04/13/06    Moved deletefile before otherfile functionality to support sql server redundancy
'2.02.0     06/06/06    Modified DCOLLECT MODE TO PERFORM SQB BACKUP
'2.03.0     01/31/07    Modified to not perform a dir of *.* (only happened once a day)
'2.04.0     05/10/07    Added destdrive to allow redundancy folder to be written to drive other that f:
'2.05.0     02/27/08    Code to not check wem.mdb if sql mode
'2.06.0     08/05/08    Fixes per false log entrys.
'2.07.0     10/01/08    Added functionality for multiple system ww apps
'2.08.0     09/02/09    RBH - Changed the .ini routine to write all .ini files settings.
'2.09.0     10/08/09    rbh - Changed the way autoback starts.  (Now it starts if we are after the backup time on startup)
'2.10.0     10/10/09    rbh - Now that autoback starts, it sometimes does not clear the "update" lock.  Added register.unregistertask to clear the update lock.
'2.11.0     12/13/09    rbh - Changed code to delete items in deletefile.txt better
'2.12.0     03/28/10    rbh - CONDITIONAL CHECK FOR FSERVER BACKING UP WEM.MDB OVER ITSELF.  NOT A PROBLEM WITH SQL
'2.13.0     06/06/10    rbh - Settings/Code change for non wem standard computer names.
'2.14.0     07/07/10    rbh - Settings/Code to make the 1 - 7 backup folders.
'2.15.0     02/07/11    rbh - Delay to wait for backup so all programs can start.
'2.15.1     06/23/11    rbh - Change to work with chinese characters
'2.16.0     06/28/11    rbh - Autoback would make a backup on start but not fire on specified time
'2.17.0     09/29/11    yd -  work with chinese characters
'2.18.0     06/04/12    rbh - made to do multiple backups up to 1 / hour
'2.19.0     07/13/12    kjj - Support for Asphalt WEMASQL.mdf
'2.20.0     01/23/13    kjj - Check if backup_ctr < 0.
'                             Additional WEMASQL corrections.
'                             When moving a sql back up from Temp folder on FServer to a client sometimes the copy was attmpted before the file was ready.
'2.21.0     09/17/13    RBH - FIX TO BACKUP SQL DATABASE IF ON SERVER 2008
'2.25.0     11/21/14    kjj - Fix for backing up more than 20 files in a WEM\W12345\*.* folder. Program would lock in infinite loop
'                             Increased dim to 1500 and will exit loop and log if exceeded.
'2.26.0     07/08/16    rbh - Code to support the "BUS"
'2.27.0     10/12/16    rbh - code to cleanup the local f drive so db backups can be made
'2.27.0     12/10/16    rbh - Logging added for debug
'2.28.0     12/12/16    rbh - Code to act as follows, when autoback runs at startup and there is an array of times (1,2,3 for backup hour, it will schedule it's backup for the hour it is on.  Code to force it to look for the next hour.
'2.29.0     12/29/16    rbh - Added code to allow a backup at midnight - 24 entered in the backup hour entry
'                             Reset after last backup - resetonceperday=TRUE in the .ini file - done for Wengers.
'2.30.0     09/27/18    adf - Put random delays in before copying files. Don't copy WW folders with '_' in the name. Delay before copying WW.
'2.31.0     04/16/19    rbh - increased number of other computers from 10 to 20 for hi-pro and others
'                           - CODE TO BACKUP THE SQL_ARCHIVE
'2.32.0     10/16/21    rbh - Remove register reference for when we do a backup.
'2.33.0     01/31/22    rbh - Changed odbc timeout to 200 seconds for wemsql_archive backup
'2.34.0     04/15/22    rbh - Modified for asphalt - custom computer names
'2.35.0     10/19/22    rbh - used sqltoolkit frm wemsqlbackup - updated for large dbs - wemsql_archive


