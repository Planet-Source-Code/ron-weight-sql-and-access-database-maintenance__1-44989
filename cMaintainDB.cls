VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMaintainDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************
' Class Module:  cMaintainDB
' File: cMaintainDB.cls
' Author: Ronald Weight
' Date: 4/10/2003
' Purpose:  This DLL has functions to help maintain and backup databases
'           It will backup, compact and repair Access databases,
'           and Attach, Detach, backup and restore SQL server databases
' References:
'              MS Jet & Replication Objects (MSJRO)
'              MS Scripting (SCRRUN.DLL)
'              MS SQL DMO Library (SQLDMO.DLL - SQLDMO.RLL)
'              MS Active X Data objects 2.7 or later
'********************************************************************************
Option Explicit

Private WithEvents m_objSQLBackup As SQLDMO.Backup
Attribute m_objSQLBackup.VB_VarHelpID = -1
Private WithEvents m_objSQLRestore As SQLDMO.Restore
Attribute m_objSQLRestore.VB_VarHelpID = -1

' events - use to communicate to client
Event CompactError(ByVal OriginalDB As String, ByVal DBCopy As String, ByVal FailMessage As String)
Event CompactFinished(ByVal OriginalDB As String, ByVal DBCopy As String)
Event CopyError(ByVal OriginalPath As String, ByVal CopyPath As String, ByVal FailMessage As String)
Event CreateAFolder(ByVal FolderPath As String, ByRef Cancel As Boolean)
Event CopyCancelled()
Event BackupFinished(ByVal BackupFolder As String)
Event BackupError(ByVal FailMessage As String, ByVal BackupFolder As String)
Event BackupLeftOnServer(ByVal FailMessage As String, ByVal FileOnServer As String)
Event PercentComplete(ByVal ProgressMessage As String, ByVal PercentDone As Long)
Event RestoreError(ByVal FailMessage As String)
Event RestoreFinished(ByVal Message As String, Database As String)
Event ModuleError(ByVal ErrMessage As String)

' constants
Private Const m_TEMP_FOLDER As Integer = 2

' member vars
Private m_strDBName As String
Private m_strDBPath As String
Private m_strErrMessage As String
Private m_lngErrNum As Long
Private m_strErrDesc As String
Private m_strErrSource As String

' ***************************** Read Only Properties ****************************

Public Property Get DBName() As String

    DBName = m_strDBName

End Property

Public Property Get DBPath() As String

    DBPath = m_strDBPath

End Property

Public Property Get LastError() As String

    Dim strErrString As String

    If m_lngErrNum = 0 Then    ' no error
        LastError = "None"
        Exit Sub
    End If

    ' return the last error as 3 parts - Source, Error Number and Description
    strErrString = "Error #: " & m_lngErrNum
    strErrString = strErrString & vbCrLf & "Description: " & m_strErrDesc
    strErrString = strErrString & vbCrLf & "Source: " & m_strErrSource

    LastError = strErrString

End Property

' ******************************** Public Methods *******************************

Public Function CompactAccessDB(ByVal OriginalDBPath As String) As Boolean
    '********************************************************************************
    ' Procedure CompactAccessDB
    ' Author: Ronald Weight
    ' Date: 4/10/2003
    ' Purpose: This function compresses the Access database represented by OriginalDBPath. The compress
    '                  method will also repair a corrupted database.
    '
    ' Input parameters:  The path and file name of the Access database to be repaired/compacted (OriginalDBPath)
    ' Output parameters: True or False - returns True if successful
    '********************************************************************************

    On Error GoTo Compact_Err

    Dim objFileSystem As FileSystemObject, fsoFile As File
    Dim strDrive As String, strErrMsg As String
    Dim bSuccess As Boolean, strConnect As String
    Dim strTempCopy As String, strCompactCopy As String
    Dim strFolder As String, bExclusive As Boolean

    ' check if we can have exclusive access to copy the database - if not, the copy won't work properly
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & OriginalDBPath
    bExclusive = CheckExclusive(strConnect)
    If Not bExclusive Then
        CompactAccessDB = False
        strErrMsg = "The database is in use. The compact/repair process "
        strErrMsg = strErrMsg & vbCrLf & "can not be done if the database is in use."
        strErrMsg = strErrMsg & vbCrLf & "Please check that nobody is using the database and try again."
        RaiseEvent CompactError(OriginalDBPath, "", strErrMsg)
        Exit Function
    End If

    Set objFileSystem = New FileSystemObject

    strFolder = objFileSystem.GetSpecialFolder(m_TEMP_FOLDER)

    If Right$(strFolder, 1) = "\" Then
        strTempCopy = strFolder & GetFormattedDateString
    Else
        strTempCopy = strFolder & "\" & GetFormattedDateString
    End If

    strCompactCopy = strTempCopy & "Compacted.mdb"
    strTempCopy = strTempCopy & "TempDB.mdb"

    ' copy the original file to local drive
    objFileSystem.CopyFile OriginalDBPath, strTempCopy

    ' by doing it this way, we are working with copies of the original DB rather than the original. If something
    ' goes wrong with the compact/repair process, then we still have the orginal file to work with
    bSuccess = DoCompact(strTempCopy, strCompactCopy)
    If Not bSuccess Then
        CompactAccessDB = False
        RaiseEvent CompactError(OriginalDBPath, strTempCopy, m_strErrMessage)
        Exit Function
    End If

    ' ONLY if it successfully completes, delete the original DB file
    Set fsoFile = objFileSystem.GetFile(OriginalDBPath)
    fsoFile.Delete True

    ' Move & rename the new one to the original database name - we get the compacted copy first
    Set fsoFile = objFileSystem.GetFile(strCompactCopy)
    fsoFile.Move OriginalDBPath

    ' once we have moved the copied, compacted DB back to the original place,
    ' delete the copy to clean up
    objFileSystem.DeleteFile strTempCopy, True

    CompactAccessDB = True
    RaiseEvent CompactFinished(OriginalDBPath, strTempCopy)

Compact_Exit:
    On Error Resume Next
    Set fsoFile = Nothing
    Set objFileSystem = Nothing
    Exit Function

Compact_Err:
    CompactAccessDB = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.CompactAccessDB"
    strErrMsg = "A copy of the database was made before compacting. The copy is: "
    strErrMsg = strErrMsg & vbCrLf & strTempCopy
    strErrMsg = strErrMsg & vbCrLf & "If the compact process succeeded, but the copy did not, there "
    strErrMsg = strErrMsg & vbCrLf & "should be another file called: "
    strErrMsg = strErrMsg & vbCrLf & strCompactCopy
    RaiseEvent CompactError(OriginalDBPath, strTempCopy, strErrMsg)
    Resume Compact_Exit

End Function

Public Function BackupAccessDB(ByVal SourceDBPath As String, ByVal DestFolder As String) As Boolean

    '********************************************************************************
    ' Procedure BackupAccessDB
    ' Author: Ronald Weight
    ' Date: 4/10/2003
    ' Purpose:  Backs up an Access DB by copying it to the designated folder. This function assumes that nobody
    '                   is using the database. If someone is using the database, the copy process may produce an error
    '
    ' Input parameters:  Source database and path (SourceDBPath) - Destination folder (DestFolder)
    ' Output parameters: True or False - True if successful at copying the file
    '********************************************************************************

    Dim objFileSystem As FileSystemObject
    Dim bExists As Boolean, strConnect As String
    Dim strDestFilePath As String, bCancel As Boolean
    Dim strFolder As String, bExclusive As Boolean
    Dim strMsg As String, strDBName As String

    On Error GoTo Copy_Err

    If DestFolder = "" Then DestFolder = "C:\Backup"

    ' check if we can have exclusive access to copy the database - if not, the copy won't work properly
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SourceDBPath
    bExclusive = CheckExclusive(strConnect)
    If Not bExclusive Then
        BackupAccessDB = False
        strMsg = "The database is in use. The backup process cannot be "
        strMsg = strMsg & "done reliably if the database is open."
        strMsg = strMsg & vbCrLf & "Please check that nobody is using the database and try again."
        RaiseEvent CopyError(SourceDBPath, DestFolder, strMsg)
        Exit Function
    End If

    Set objFileSystem = New FileSystemObject

    ' check if the destination folder exists
    bExists = objFileSystem.FolderExists(DestFolder)
    If Not bExists Then
        ' send event with Cancel param to see if user wants to create the folder that doesn't exist - if True, don't create destination
        ' folder and exit function
        bCancel = False
        RaiseEvent CreateAFolder(DestFolder, bCancel)
        If bCancel Then
            BackupAccessDB = False
            RaiseEvent CopyCancelled
            Exit Function
        End If
        objFileSystem.CreateFolder DestFolder
    End If

    ' check if source file exists before continuing
    bExists = objFileSystem.FileExists(SourceDBPath)
    If Not bExists Then
        ' file doesn't exist, so we can't copy it
        BackupAccessDB = False
        RaiseEvent CopyError(SourceDBPath, DestFolder, "Source File does not exist in the given path!")
        Exit Function
    End If

    ' now create the destination filename, and append it to the DestFolder given
    If Right$(DestFolder, 1) = "\" Then
        strDestFilePath = DestFolder & GetFormattedDateString
    Else
        strDestFilePath = DestFolder & "\" & GetFormattedDateString
    End If

    ' get DB Name to use as part of the copy
    strDBName = GetFileNameFromPath(SourceDBPath)
    ' finish setting up final copy file and path name
    strDestFilePath = strDestFilePath & strDBName

    ' copy file,  but do NOT allow it to overwrite an existing one
    objFileSystem.CopyFile SourceDBPath, strDestFilePath, False

    BackupAccessDB = True

Copy_Exit:
    On Error Resume Next
    Set objFileSystem = Nothing
    Exit Function

Copy_Err:
    BackupAccessDB = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.BackupAccessDB"
    RaiseEvent CopyError(SourceDBPath, strDestFilePath, LastError)
    Resume Copy_Exit

End Function

Public Function AttachSQLDatabase(ByVal SQLServerName As String, _
        ByVal DBName As String, ByVal SQLUserLogin As String, ByVal PWD As String, _
        ByVal DBFile As String, ByVal LogFile As String, _
        Optional ByVal ConnectUsingWindowsAuth As Boolean = True) As Boolean
    
    Dim objSQLServer As SQLServer
    Dim strFiles As String
    
    On Error GoTo Attach_Err
    
    Set objSQLServer = New SQLServer
    If ConnectUsingWindowsAuth Then
        objSQLServer.Connect SQLServerName
    Else
        objSQLServer.Connect SQLServerName, SQLUserLogin, PWD
    End If
    
    strFiles = "[" & DBFile & "]," & "[" & LogFile & "]"
    
'    oBackup.RelocateFiles = "[Northwind1]" + "," + "[D:\Data\North_1.mdf]" _
 '  + "," + "[Northwind2]" + "," + "[D:\Data\North_2.mdf]"

    objSQLServer.AttachDB DBName, strFiles
    
    AttachSQLDatabase = True
    
Attach_Exit:
    On Error Resume Next
    objSQLServer.DisConnect
    objSQLServer.Close
    Set objSQLServer = Nothing
    Exit Function
    
Attach_Err:
    AttachSQLDatabase = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.AttachSQLDatabase"
    Debug.Print LastError & vbCrLf & "File String: " & strFiles
    RaiseEvent ModuleError(LastError)
    Resume Attach_Exit
    
End Function

Public Function DetachSQLDatabase(ByVal SQLServerName As String, _
        ByVal DBName As String, ByVal SQLUserLogin As String, ByVal PWD As String, _
        Optional ByVal ConnectUsingWindowsAuth As Boolean = True) As Boolean

    ' This function allows the client to detach a DB from SQL server - then the
    ' physical files for the DB can be moved, copied or backed up
    ' An Attach function could then put the DB back in SQL server
    Dim objSQLServer As SQLServer
    
    On Error GoTo Detach_Err
    
    Set objSQLServer = New SQLServer
    If ConnectUsingWindowsAuth Then
        objSQLServer.Connect SQLServerName
    Else
        objSQLServer.Connect SQLServerName, SQLUserLogin, PWD
    End If
    
    objSQLServer.DetachDB DBName, True
    
    DetachSQLDatabase = True
    
Detach_Exit:
    On Error Resume Next
    objSQLServer.DisConnect
    objSQLServer.Close
    Set objSQLServer = Nothing
    Exit Function
    
Detach_Err:
    DetachSQLDatabase = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.DetachSQLDatabase"
    Debug.Print LastError
    RaiseEvent ModuleError(LastError)
    Resume Detach_Exit
    
End Function
        

Public Function BackupSQLDatabase(ByVal ServerFolder As String, ByVal SQLServerName As String, _
            ByVal DBName As String, ByVal SQLUserLogin As String, _
            ByVal PWD As String, Optional ByVal LocalCopyPath As String = "", _
            Optional ByVal ConnectUsingWindowsAuth As Boolean = True) As Boolean
    '********************************************************************************
    ' Procedure BackupSQLDatabase
    ' Author: Ronald Weight
    ' Date: 4/11/2003
    ' Purpose: Performs a full database backup of the database represented by DBName on the SQLServerName server
    '                   The file created can then be restored later
    ' Input parameters:  Destination folder for backup, SQL Server name, Database name, UserID, Password and an optional
    '                                  param to determine if we should just use Windows auth or SQL auth
    ' Output parameters:  True if successful, false if not - It also
    '********************************************************************************

    Dim strBackupFile As String
    Dim bSuccess As Boolean
    Dim strLocalFile As String
    Dim objSQLServer As SQLServer

    On Error GoTo Backup_Err

    ' format the Local path and file name
    If LocalCopyPath <> "" Then
        If Right$(LocalCopyPath, 1) = "\" Then
            strLocalFile = LocalCopyPath & GetFormattedDateString
        Else
            strLocalFile = LocalCopyPath & "\" & GetFormattedDateString
        End If
        strLocalFile = strLocalFile & DBName & ".bak"
    End If
    
    ' format the filename for the server
    If Right$(ServerFolder, 1) = "\" Then
        strBackupFile = ServerFolder & GetFormattedDateString
    Else
        strBackupFile = ServerFolder & "\" & GetFormattedDateString
    End If
    strBackupFile = strBackupFile & DBName & ".bak"

    ' instantiate SQL objects to do the backup
    Set m_objSQLBackup = New SQLDMO.Backup
    Set objSQLServer = New SQLServer

    ' check if we connect using Windows Authorization or SQL server
    ' NOTE: The password is passed in plain text - this could be changed
    ' to accept an encrypted password, and un-encrypt it just before passing
    ' it to the connect method. I HIGHLY recommend using Windows Authorization to
    ' avoid passing any passwords and for better security
    If ConnectUsingWindowsAuth Then
        objSQLServer.Connect SQLServerName
    Else
        objSQLServer.Connect SQLServerName, SQLUserLogin, PWD
    End If

    ' set the properties needed to perform the backup and then do it
    With m_objSQLBackup
        .Database = DBName
        .Files = strBackupFile
        .Action = SQLDMOBackup_Database
        .SQLBackup objSQLServer
    End With

    ' This step is needed for servers that can't backup to a network location
    ' otherwise, you could change this so SQL server just backs up to where you
    ' wanted it in the first place. In my situation, the SQL servers I work with
    ' only have access to drives on the server - so I first back up the SQL DB to
    ' a server folder, then move that file to a local or mapped network drive
    ' that's available on the computer running this process. Of course, this assumes
    ' that the client computer has access to a folder on the SQL server - if not
    ' then the server would need to be configured to be able to backup to a network
    ' drive - if so, then this process could be skipped.

    ' copy (or move) the DB backup file to local folder
    bSuccess = MoveBackupToLocal(strBackupFile, strLocalFile)
    ' if it didn't succeed, then we still have a valid backup file, it's just
    ' in the original location - let user know via an appropriate event
    If Not bSuccess Then
        BackupSQLDatabase = True
        RaiseEvent BackupLeftOnServer(LastError, strBackupFile)
    Else
        BackupSQLDatabase = True
        RaiseEvent BackupFinished(strLocalFile)
    End If

Backup_Exit:
    On Error Resume Next
    ' release objects
    Set m_objSQLBackup = Nothing
    objSQLServer.DisConnect
    Set objSQLServer = Nothing
    Exit Function

Backup_Err:
    BackupSQLDatabase = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.BackupSQLDatabase"
    RaiseEvent BackupError(LastError, ServerFolder)
    Resume Backup_Exit

End Function

Public Function RestoreSQLDatabase(ByVal SQLServerName As String, _
        ByVal SQLUserLogin As String, ByVal PWD As String, ByVal DBName As String, _
        ByVal BackupFile As String, Optional ByVal ConnectUsingWindowsAuth As Boolean = True, _
        Optional ByVal CreateDatabase As Boolean = True) As Boolean

    Dim objSQLServer As SQLServer

    On Error GoTo Restore_Err

    ' instantiate SQL objects to do the backup
    Set m_objSQLRestore = New SQLDMO.Restore
    Set objSQLServer = New SQLServer

    ' check if we connect using Windows Authorization or SQL server
    ' NOTE: The password is passed in plain text - this could be changed
    ' to accept an encrypted password, and un-encrypt it just before passing
    ' it to the connect method. I HIGHLY recommend using Windows Authorization to
    ' avoid passing any passwords
    If ConnectUsingWindowsAuth Then
        objSQLServer.Connect SQLServerName
    Else
        objSQLServer.Connect SQLServerName, SQLUserLogin, PWD
    End If

    With m_objSQLRestore
        .Database = DBName
        If CreateDatabase Then
            ' this will force creation of the database - this might be a problem
            ' if the database already exists on SQL Server
            .ReplaceDatabase = True
        Else
            .ReplaceDatabase = False
        End If
        .DatabaseFiles = ""
        .Files = BackupFile
        .Action = SQLDMORestore_Database
        .SQLRestore objSQLServer
    End With

    RestoreSQLDatabase = True
    RaiseEvent RestoreFinished("Restore completed", DBName)

Restore_Exit:
    On Error Resume Next
    ' release objects
    Set m_objSQLBackup = Nothing
    objSQLServer.DisConnect
    Set objSQLServer = Nothing
    Exit Function

Restore_Err:
    RestoreSQLDatabase = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.RestoreSQLDatabase"
    Debug.Print LastError
    RaiseEvent RestoreError(LastError)
    Resume Restore_Exit

End Function

Public Function RestoreAccessDB(ByVal BackupPath As String, ByVal RestorePath As String, _
                Optional ByVal CreateRestorePath As Boolean = False) As Boolean

    Dim objFileSys As FileSystemObject
    Dim strRestoreFile As String, strErrMsg As String
    Dim bExists As Boolean, strFileName As String, bCancel As Boolean

    On Error GoTo Restore_Err

    Set objFileSys = New FileSystemObject

    ' check if the file represented by BackupPath exists
    bExists = objFileSys.FileExists(BackupPath)
    If Not bExists Then
        ' backup file doesn't exist
        Err.Raise 1001, "cMaintainDB.RestoreAccessDB", "The file selected to restore does not exist: " & BackupPath
    End If

    ' check to see if the path represented by RestorePath exists - optionally ask user to create it
    strFileName = GetFileNameFromPath(RestorePath)
    ' now, m_strDBPath  holds the path
    bExists = objFileSys.FolderExists(m_strDBPath)
    If Not bExists Then
        bCancel = False
        RaiseEvent CreateAFolder(m_strDBPath, bCancel)
        If bCancel Then
            ' user said not to create the restore folder
            RestoreAccessDB = False
            RaiseEvent CopyCancelled
            GoTo Restore_Exit
        Else
            objFileSys.CreateFolder (m_strDBPath)
        End If
    End If
    
    strRestoreFile = RestorePath
    
    ' now we can copy the backup file to the restore folder - DO NOT ALLOW Overwriting of an existing file!
    objFileSys.CopyFile BackupPath, strRestoreFile, False

    RestoreAccessDB = True
    RaiseEvent RestoreFinished("The database was successfully backed up from: " & BackupPath & " to: " & strRestoreFile, m_strDBName)

Restore_Exit:
    On Error Resume Next
    ' release objects
    Set objFileSys = Nothing
    Exit Function

Restore_Err:
    RestoreAccessDB = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.RestoreAccessDB"
    If m_lngErrNum = 58 Then
        ' the file already exists - append the word COPY to the restore file name
        ' and notify client of the problem
        strFileName = "CopyOf" & strFileName
        strRestoreFile = m_strDBPath & strFileName
        strErrMsg = "The file you are trying to restore already exists."
        strErrMsg = strErrMsg & vbCrLf & RestorePath
        strErrMsg = strErrMsg & vbCrLf & "It has been renamed and copied to: " & strRestoreFile
        RaiseEvent RestoreError(strErrMsg)
        Resume
    End If
    Debug.Print LastError
    RaiseEvent RestoreError(LastError)
    Resume Restore_Exit


End Function

Public Sub About()

    On Error Resume Next

    frmAbout.Show vbModal

End Sub

' ******************************** Private Methods *******************************

Private Function DoCompact(ByVal strSourceDB As String, ByVal strDestDB As String) As Boolean
    '********************************************************************************
    ' Procedure DoCompact
    ' Author: Ronald Weight
    ' Date: 4/10/2003
    ' Purpose: This is the function that actually compacts the DB
    ' Input parameters:  Source Database and Destination database
    ' Output parameters: True or False - True if successful
    '********************************************************************************

    Dim objJetEngine          As JRO.JetEngine
    Dim strSourceConnect   As String
    Dim strDestConnect     As String
    Dim lngEngineType      As Long
    Dim objConn As ADODB.Connection, strConnect As String

    On Error GoTo Compact_Err

    ' open a connection to the source DB so we can determine the version of
    ' this Access DB - if version 5, we compress it with Engine type = 5
    ' if version 5, we compress it with Engine type = 4
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSourceDB
    Set objConn = New ADODB.Connection
    With objConn
        .ConnectionString = strConnect
        .Mode = adModeRead
        .CursorLocation = adUseClient
        .Open
    End With

    ' determine the engine type / version so we can compress it properly
    lngEngineType = objConn.Properties.Item("Jet OLEDB:Engine Type").Value

    ' Build connection strings for SourceConnection and
    ' DestConnection arguments.
    strSourceConnect = "Data Source=" & strSourceDB
    ' give a different string for verion 5 and 4 DB's
    If lngEngineType = 5 Then
        strDestConnect = "Data Source=" & strDestDB & ";" & _
                "Jet OLEDB:Engine Type=5"
    ElseIf lngEngineType = 4 Then
        strDestConnect = "Data Source=" & strDestDB & ";" & _
                "Jet OLEDB:Engine Type=4"
    Else
        strDestConnect = "Data Source=" & strDestDB & ";" & _
                "Jet OLEDB:Engine Type=4"
    End If

    On Error Resume Next
    ' close the connection
    objConn.Close
    Set objConn = Nothing

    On Error GoTo Compact_Err
    ' open the jet object to compact the db
    Set objJetEngine = New JRO.JetEngine

    ' Compact and repair the database specified by strSourceDB
    ' to the name and path specified by strDestDB.
    objJetEngine.CompactDatabase strSourceConnect, strDestConnect

    DoCompact = True

Compact_Exit:
    On Error Resume Next
    Set objJetEngine = Nothing
    Exit Function

Compact_Err:
    DoCompact = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.DoCompact"
    Debug.Print LastError
    m_strErrMessage = LastError
    Resume Compact_Exit

End Function

Private Function GetFormattedDateString() As String

    On Error GoTo Format_Err

    Dim strDay As String, strMonth As String, strYear As String
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim strDateString As String

    intYear = Year(Date)
    intMonth = Month(Date)
    intDay = Day(Date)

    strMonth = Format$(intMonth, "00")
    strDay = Format$(intDay, "00")
    strYear = Format$(intYear, "0000")

    strDateString = strMonth & strDay & strYear

    GetFormattedDateString = strDateString

Format_Err_Exit:
    Exit Function

Format_Err:
    ' default value if procedure fails
    GetFormattedDateString = "010101"
    Resume Format_Err_Exit

End Function

Private Function CheckExclusive(ByVal ConnectString As String) As Boolean

    '************************* Procedure Header ****************************
    ' Procedure CheckExclusive
    ' Author: Ronald  Weight
    ' Date: 4/18/2003
    ' Purpose: Checks and Access DB to see if we have exclusive access
    ' Input parameters:  Connection string to connect to the Access DB
    ' Output parameters:  Returns True if we can get exclusive access - False if it's already opened
    '**********************************************************************

    Dim objConn As ADODB.Connection

    On Error GoTo Check_Err

    Set objConn = New ADODB.Connection
    With objConn
        .ConnectionString = ConnectString
        .Mode = adModeShareExclusive
        .CursorLocation = adUseClient
        .Open
    End With

    CheckExclusive = True

Check_Exit:
    On Error Resume Next
    objConn.Close
    Set objConn = Nothing
    Exit Function

Check_Err:
    CheckExclusive = False
    Resume Check_Exit

End Function

Private Function MoveBackupToLocal(ByVal RemoteFile As String, _
                ByVal LocalFile As String) As Boolean

    ' This function moves a file from RemoteFile to LocalFile. It's meant to be used as a function to move a remote backup
    ' file to a local drive, but can be used to move a file from anywhere to anywhere else - based on the parameters passed in

    Dim objFileSys As FileSystemObject
    Dim strPath As String, strFileName As String
    Dim bExists As Boolean, bCancel As Boolean

    On Error GoTo Move_Err

    Set objFileSys = New FileSystemObject

    ' check if local path exists - this will separate the Path and the File name so we can check either or both
    strFileName = GetFileNameFromPath(LocalFile)
    ' check the path
    bExists = objFileSys.FolderExists(m_strDBPath)
    If Not bExists Then
        ' send event with Cancel param to see if user wants to create the folder that doesn't exist - if True, don't create destination
        ' folder and exit function
        bCancel = False
        RaiseEvent CreateAFolder(m_strDBPath, bCancel)
        If bCancel Then
            MoveBackupToLocal = False
            RaiseEvent CopyCancelled
            Exit Function
        End If
        objFileSys.CreateFolder m_strDBPath
    End If

    objFileSys.MoveFile RemoteFile, LocalFile

    MoveBackupToLocal = True

Move_Exit:
    On Error Resume Next
    Set objFileSys = Nothing
    Exit Function

Move_Err:
    MoveBackupToLocal = False
    m_lngErrNum = Err.Number
    m_strErrDesc = Err.Description
    m_strErrSource = "cMaintainDB.MoveBackupToLocal"
    m_strErrMessage = LastError
    Resume Move_Exit

End Function


'     strDeviceName = "Temp_Bak1"
'
'    Set objSQLDevice = New SQLDMO.BackupDevice
'    objSQLDevice.Name = strDeviceName
'    objSQLDevice.Type = SQLDMODevice_DiskDump
'    objSQLDevice.PhysicalLocation = strBackupFile
'
'    objSQLServer.BackupDevices.Add objSQLDevice
'
'
'   Set objConn = New ADODB.Connection
'   With objConn
'        .ConnectionString = ConnString
'        .CursorLocation = adUseClient
'        .Mode = adModeRead
'        .Open
'    End With
'
'    Set objCat = New ADOX.Catalog
'    Set objCat.ActiveConnection = objConn
'    lngTableCount = objCat.Tables.Count
'    For lngLoop = 0 To lngTableCount - 1
'        If objCat.Tables.Item(lngLoop).Type = "TABLE" Then
'            strTable = objCat.Tables.Item(lngLoop).Name
'            Debug.Print "Table name: " & strTable
'            strSQL = "SELECT * FROM " & strTable
'            Debug.Print "SQL: " & strSQL
'        End If
'    Next lngLoop

Friend Function GetFileNameFromPath(ByVal PathAndFileName As String) As String

    ' *******************************************************************************
    ' Routine Name : Function GetFileNameFromPath
    ' Company      :
    ' Programmer   : Ronald Weight
    ' Date Written : 4/17/03
    ' Inputs       :  Path and File name string
    ' Outputs    :  Returns the File name only as a string
    ' *******************************************************************************

    Dim intTemp As Integer, intLen As Integer
    Dim intIndex As Integer

    On Error GoTo Sep_Err

    ' parse the path and name to get the path and name separate...
    intTemp = 0
    intLen = Len(PathAndFileName)
    For intIndex = intLen To 1 Step -1
        If Mid(PathAndFileName, intIndex, 1) <> "\" Then
            intTemp = intTemp + 1
        Else
            ' when we find the first \ it's time to get out
            Exit For
        End If
    Next

    ' save filename and path values
    m_strDBName = Trim(Right(PathAndFileName, intTemp))
    m_strDBPath = Trim(Left(PathAndFileName, (intLen - intTemp)))

    GetFileNameFromPath = m_strDBName

Sep_Err_Exit:
    Exit Function

Sep_Err:
    GetFileNameFromPath = ""
    m_strDBName = ""
    m_strDBPath = ""
    Resume Sep_Err_Exit

End Function

Private Sub Class_Initialize()

    m_strDBName = ""
    m_strDBPath = ""
    m_strErrMessage = ""
    m_lngErrNum = 0
    m_strErrDesc = ""
    m_strErrSource = ""

End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    Unload frmAbout
    Set frmAbout = Nothing

    Set m_objSQLBackup = Nothing
    Set m_objSQLRestore = Nothing

End Sub

' ******************************** SQL Object Events **********************************************

' Because of a bug in SQLDMO, if we have any events being used, we must have some code or even just comments
' in EVERY event - this prevents a weird error - so that's why there isn't much here in some events.
Private Sub m_objSQLBackup_Complete(ByVal Message As String)

    'MsgBox Message

End Sub

Private Sub m_objSQLBackup_NextMedia(ByVal Message As String)
    
    ' just comments - do nothing
    
End Sub

Private Sub m_objSQLBackup_PercentComplete(ByVal Message As String, ByVal Percent As Long)

    ' track the progress of the backup - this would allow you to
    ' implement a progress bar in client that calls this class
    RaiseEvent PercentComplete(Message, Percent)

End Sub

Private Sub m_objSQLRestore_Complete(ByVal Message As String)

    ' track when the restore is finished
    RaiseEvent RestoreFinished(Message, m_strDBName)

End Sub

Private Sub m_objSQLRestore_NextMedia(ByVal Message As String)

    ' just a comment

End Sub

Private Sub m_objSQLRestore_PercentComplete(ByVal Message As String, ByVal Percent As Long)

    ' track the progress of the restore - this would allow you to
    ' implement a progress bar in client that calls this class
    RaiseEvent PercentComplete(Message, Percent)

End Sub
