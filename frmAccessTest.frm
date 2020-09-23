VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAccessTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test DB Compact and Repair"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmAccessTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Form"
      Height          =   345
      Left            =   180
      TabIndex        =   19
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   5715
      TabIndex        =   18
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Show About"
      Height          =   345
      Left            =   3480
      TabIndex        =   15
      Top             =   5040
      Width           =   1200
   End
   Begin VB.Frame fraTestAccessDB 
      Caption         =   "Testing Access DB Functions "
      Height          =   4755
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtDBRestore 
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   3720
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetRestorePath 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3690
         Width           =   435
      End
      Begin VB.CommandButton cmdLookup2 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2950
         Width           =   435
      End
      Begin VB.TextBox txtDBBackup 
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   2980
         Width           =   5415
      End
      Begin VB.CommandButton cmdRestoreAccess 
         Caption         =   "Test Restore Access"
         Height          =   375
         Left            =   3780
         TabIndex        =   9
         Top             =   4140
         Width           =   1935
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Compact"
         Height          =   375
         Left            =   3780
         TabIndex        =   8
         Top             =   1080
         Width           =   1875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy DB (Backup)"
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   1980
         Width           =   1875
      End
      Begin VB.TextBox txtCopyDB 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetFolder 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1530
         Width           =   435
      End
      Begin VB.TextBox txtDB 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   5415
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   375
         Left            =   5820
         Picture         =   "frmAccessTest.frx":0820
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   630
         Width           =   435
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   435
         Left            =   180
         TabIndex        =   14
         Top             =   4140
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path to Restore DB To (only path - no file name)"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   3380
         Width           =   5415
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   180
         X2              =   6540
         Y1              =   2480
         Y2              =   2480
      End
      Begin VB.Line linSep 
         Index           =   0
         X1              =   180
         X2              =   6540
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path of Backup DB to Restore (path and file name)"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   2640
         Width           =   5415
      End
      Begin VB.Label llblTest 
         Caption         =   "Enter Path to Backup DB To"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   5415
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter Path to Access DB"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
   End
   Begin MSComDlg.CommonDialog dlgTest 
      Left            =   2760
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAccessTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_objMaintain As cMaintainDB
Attribute m_objMaintain.VB_VarHelpID = -1

Private Const m_MESSAGE As String = "Please enter required data in fields."

Private Sub cmdAbout_Click()

    m_objMaintain.About

End Sub

Private Sub cmdClear_Click()
    
    txtCopyDB.Text = ""
    txtDB.Text = ""
    txtDBBackup.Text = ""
    txtDBRestore.Text = ""
    lblMsg.Visible = False
    
End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdCopy_Click()

    Dim strOriginalDB As String
    Dim strCopyPath As String
    Dim bSuccess As Boolean
    Dim objMouse As cMouseCursor

    Set objMouse = New cMouseCursor
    objMouse.SetCursor vbHourglass

    If txtDB.Text = "" Then
        lblMsg.Visible = True
        txtDB.SetFocus
        Exit Sub
    Else
        lblMsg.Visible = False
    End If
    
    If txtCopyDB.Text = "" Then
        txtCopyDB.SetFocus
        lblMsg.Visible = True
        Exit Sub
    Else
        lblMsg.Visible = False
    End If

    strOriginalDB = txtDB.Text
    strCopyPath = txtCopyDB.Text
    
    ' now we have the info needed to do the backup
    bSuccess = m_objMaintain.BackupAccessDB(strOriginalDB, strCopyPath)
    If bSuccess Then
        MsgBox "DB copied to: " & strCopyPath
    End If

    Debug.Print "DB Name: " & m_objMaintain.DBName
    Debug.Print "DB Path: " & m_objMaintain.DBPath

End Sub

Private Sub cmdLookup_Click()

    On Error Resume Next
    
    Dim strFile As String

    With dlgTest
        .CancelError = True
        .DialogTitle = "Select Access DB to Compress"
        .Filter = "Access Database(*.mdb) | *.mdb"
        .InitDir = App.Path
        .ShowOpen
    End With

    If Err.Number <> 0 Then
        ' user cancelled
        strFile = ""
        Me.txtDB.Text = ""
        txtDB.SetFocus
        Exit Sub
    End If

    strFile = dlgTest.FileName

    txtDB.Text = strFile

End Sub

Private Sub cmdLookup2_Click()
    
    On Error Resume Next
    
    Dim strFile As String

    With dlgTest
        .CancelError = True
        .DialogTitle = "Select Database to restore"
        .Filter = "Access Database(*.mdb) | *.mdb"
        .InitDir = App.Path
        .ShowOpen
    End With

    If Err.Number <> 0 Then
        ' user cancelled
        strFile = ""
        txtDBBackup.Text = ""
        txtDBBackup.SetFocus
        Exit Sub
    End If

    strFile = dlgTest.FileName

    txtDBBackup.Text = strFile

End Sub

Private Sub cmdRestoreAccess_Click()

    Dim strBackupPath As String
    Dim strRestorePath As String

    strBackupPath = txtDBBackup.Text
    strRestorePath = txtDBRestore.Text

    If strBackupPath = "" Then
        txtDBBackup.SetFocus
        lblMsg.Visible = True
        Exit Sub
    Else
        lblMsg.Visible = False
    End If

    If strRestorePath = "" Then
        txtDBRestore.SetFocus
        lblMsg.Visible = True
        Exit Sub
    Else
        lblMsg.Visible = False
    End If

    m_objMaintain.RestoreAccessDB strBackupPath, strRestorePath, True

End Sub

Private Sub cmdTest_Click()

    Dim strOriginalDB As String
    Dim bSuccess As Boolean
    Dim objMouse As cMouseCursor

    Set objMouse = New cMouseCursor
    objMouse.SetCursor vbHourglass

    strOriginalDB = txtDB.Text

    If strOriginalDB = "" Then
        lblMsg.Visible = True
        txtDB.SetFocus
        Exit Sub
    Else
        lblMsg.Visible = False
    End If

    bSuccess = m_objMaintain.CompactAccessDB(strOriginalDB)

End Sub


Private Sub cmdGetFolder_Click()

    Dim strFile As String
    Dim strPrompt As String

    strPrompt = "Select the folder to copy your backup file to."

    strFile = BrowseForFolder(Me.hWnd, strPrompt)

    txtCopyDB.Text = strFile

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Set m_objMaintain = New cMaintainDB
    lblMsg.Caption = m_MESSAGE
    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Violations;PWD="";Data Source=(Local)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    Set m_objMaintain = Nothing

End Sub

Private Sub m_objMaintain_BackupError(ByVal FailMessage As String, ByVal BackupFolder As String)

    Dim strMsg As String

    strMsg = FailMessage
    strMsg = strMsg & vbCrLf & "Backup folder = " & BackupFolder

    MsgBox strMsg, vbOKOnly + vbInformation, "Backup Failed"


End Sub

Private Sub m_objMaintain_BackupFinished(ByVal BackupFolder As String)

    Dim strMsg As String

    strMsg = "The backup completed successfully. The files are located at: " & BackupFolder

    MsgBox strMsg

End Sub

Private Sub m_objMaintain_BackupLeftOnServer(ByVal FailMessage As String, ByVal FileOnServer As String)

    MsgBox FailMessage & vbCrLf & "File is: " & FileOnServer, vbOKOnly + vbInformation, "Cannot copy backup file"

End Sub

Private Sub m_objMaintain_CompactError(ByVal OriginalDB As String, ByVal DBCopy As String, ByVal FailMessage As String)

    Dim strMsg As String

    strMsg = "There was a problem with compacting the database."
    strMsg = strMsg & vbCrLf & "For your information, the original path is: "
    strMsg = strMsg & vbCrLf & OriginalDB
    strMsg = strMsg & vbCrLf & "The DB was first copied to the following location: "
    strMsg = strMsg & vbCrLf & DBCopy
    strMsg = strMsg & vbCrLf & "It is possible that the original DB is missing or corrupted. If so, "
    strMsg = strMsg & vbCrLf & "The copy should be good. In this case, replace the original with the copy."
    strMsg = strMsg & vbCrLf & "The error is reported as: "
    strMsg = strMsg & vbCrLf & FailMessage

    MsgBox strMsg, vbOKOnly + vbCritical, "Error in compacting DB"

End Sub

Private Sub m_objMaintain_CompactFinished(ByVal OriginalDB As String, ByVal DBCopy As String)

    Dim strMsg As String

    strMsg = "The database was successfully compacted/repaired."
    strMsg = strMsg & vbCrLf & "The compacted DB = " & OriginalDB

    MsgBox strMsg

End Sub

Private Sub m_objMaintain_CopyCancelled()

    MsgBox "User cancelled the database copy process."

End Sub

Private Sub m_objMaintain_CopyError(ByVal OriginalPath As String, ByVal CopyPath As String, ByVal FailMessage As String)

    Dim strMsg As String

    strMsg = "There was a problem with copying the database for backup or maintenance purposes."
    strMsg = strMsg & vbCrLf & "For your information, the original path is: "
    strMsg = strMsg & vbCrLf & OriginalPath
    strMsg = strMsg & vbCrLf & "The path to copy the DB to is: "
    strMsg = strMsg & vbCrLf & CopyPath
    strMsg = strMsg & vbCrLf & "The error is reported as: "
    strMsg = strMsg & vbCrLf & FailMessage

    MsgBox strMsg, vbOKOnly + vbCritical, "Error in copying DB"

End Sub

Private Sub m_objMaintain_CreateAFolder(ByVal FolderPath As String, Cancel As Boolean)

    Dim intResponse As Integer
    Dim strMsg As String

    strMsg = "The folder " & FolderPath & " does not exist. Do you want to create it?"
    strMsg = strMsg & vbCrLf & "Click on Yes to create the folder, No if you don't want to create it."
    strMsg = strMsg & vbCrLf & "If you click on No, the database will NOT be copied."
    intResponse = MsgBox(strMsg, vbYesNo + vbQuestion, "Create Folder?")
    If intResponse = vbNo Then
        Cancel = True
    End If

End Sub

Private Sub m_objMaintain_RestoreError(ByVal FailMessage As String)

    MsgBox FailMessage, vbOKOnly + vbCritical, "Unable to restore DB"

End Sub

Private Sub m_objMaintain_RestoreFinished(ByVal Message As String, Database As String)

    MsgBox Message, vbOKOnly + vbInformation, "Restore Successful!"

End Sub
