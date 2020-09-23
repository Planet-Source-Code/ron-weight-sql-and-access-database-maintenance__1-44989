VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSQLTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test SQL DB Maintenance"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSQLTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "Show About"
      Height          =   345
      Left            =   4620
      TabIndex        =   30
      Top             =   5940
      Width           =   1200
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   6675
      TabIndex        =   6
      Top             =   5940
      Width           =   1200
   End
   Begin VB.Frame fraSQL 
      Caption         =   "Test SQL DB Functions "
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7755
      Begin VB.CommandButton cmdDetach 
         Caption         =   "Detach"
         Height          =   375
         Left            =   6720
         TabIndex        =   32
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "Attach"
         Height          =   375
         Left            =   6720
         TabIndex        =   31
         Top             =   420
         Width           =   795
      End
      Begin VB.CheckBox chkCreate 
         Alignment       =   1  'Right Justify
         Caption         =   "Create DB?"
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRestoreFile 
         Height          =   375
         Left            =   5700
         Picture         =   "frmSQLTest.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4598
         Width           =   435
      End
      Begin VB.TextBox txtRestoreFile 
         Height          =   330
         Left            =   180
         TabIndex        =   27
         Top             =   4620
         Width           =   5415
      End
      Begin VB.Frame fraFunction 
         Caption         =   "Select function type "
         Height          =   975
         Left            =   4500
         TabIndex        =   23
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optType 
            Caption         =   "Restore SQL DB"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optType 
            Caption         =   "Backup SQL DB"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   300
            Width           =   1755
         End
      End
      Begin VB.Frame fraAuth 
         Caption         =   "Select authorization mode "
         Height          =   975
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   4095
         Begin VB.OptionButton optAuth 
            Caption         =   "SQL Server Authorization"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   600
            Width           =   3555
         End
         Begin VB.OptionButton optAuth 
            Caption         =   "Windows Authorization (recommended)"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   300
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmdCopyTo 
         Height          =   375
         Left            =   5700
         Picture         =   "frmSQLTest.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3645
         Width           =   435
      End
      Begin VB.TextBox txtSQLCopy 
         Height          =   330
         Left            =   180
         TabIndex        =   18
         Top             =   3660
         Width           =   5415
      End
      Begin VB.TextBox txtDBName 
         Height          =   330
         Left            =   1800
         TabIndex        =   16
         Top             =   1920
         Width           =   2115
      End
      Begin VB.TextBox txtPWD 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   5460
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   1920
         Width           =   2115
      End
      Begin VB.TextBox txtUID 
         Height          =   330
         Left            =   5460
         TabIndex        =   12
         Top             =   1500
         Width           =   2115
      End
      Begin VB.TextBox txtSQLServer 
         Height          =   330
         Left            =   1800
         TabIndex        =   10
         Top             =   1500
         Width           =   2115
      End
      Begin VB.CommandButton cmdRestoreSQL 
         Caption         =   "Test SQL Restore"
         Height          =   555
         Left            =   6360
         TabIndex        =   4
         Top             =   4508
         Width           =   1095
      End
      Begin VB.CommandButton cmdLookupSQL 
         Height          =   375
         Left            =   5700
         Picture         =   "frmSQLTest.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2850
         Width           =   435
      End
      Begin VB.TextBox txtSQLBackup 
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   2865
         Width           =   5415
      End
      Begin VB.CommandButton cmdTestSQL 
         Caption         =   "Test SQL Backup"
         Height          =   555
         Left            =   6360
         TabIndex        =   1
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter path to existing BACKUP file to use to restore SQL DB"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   4320
         Width           =   5295
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   180
         X2              =   7520
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line linSep 
         Index           =   0
         X1              =   180
         X2              =   7520
         Y1              =   2410
         Y2              =   2410
      End
      Begin VB.Label lblTest 
         Caption         =   "Path to copy backup to on LOCAL drive (path only)"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   17
         Top             =   3360
         Width           =   5295
      End
      Begin VB.Label lblTest 
         Caption         =   "Database Name"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   15
         Top             =   1958
         Width           =   1395
      End
      Begin VB.Label lblTest 
         Caption         =   "Password"
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   13
         Top             =   1958
         Width           =   975
      End
      Begin VB.Label lblTest 
         Caption         =   "User ID (login)"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   11
         Top             =   1538
         Width           =   1275
      End
      Begin VB.Label lblTest 
         Caption         =   "SQL Server Name"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1538
         Width           =   1455
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Err"
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   5100
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Caption         =   " %"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3780
         TabIndex        =   7
         Top             =   5160
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblTest 
         Caption         =   "Enter path  to save SQL backup to on server (path only)"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   2580
         Width           =   5295
      End
   End
   Begin MSComDlg.CommonDialog dlgTest 
      Left            =   5820
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSQLTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_objMaintain As cMaintainDB
Attribute m_objMaintain.VB_VarHelpID = -1

Private Const m_MESSAGE As String = "Please enter required data in fields."

Private Const m_WINDOWS_AUTH As Integer = 0
Private Const m_SQL_AUTH As Integer = 1

Private Const m_TYPE_BACKUP As Integer = 0
Private Const m_TYPE_RESTORE As Integer = 1

Private Const m_PWD_LBL As Integer = 4

Private Sub cmdAbout_Click()
    
    m_objMaintain.About
    
End Sub

Private Sub cmdCopyTo_Click()
    
    Dim strFile As String
    Dim strPrompt As String

    strPrompt = "Select the folder on your Local drive for backup file"

    strFile = BrowseForFolder(Me.hWnd, strPrompt)

    txtSQLCopy.Text = strFile
    
    If txtSQLCopy.Text = "" Then
        lblMsg.Visible = True
    Else
        lblMsg.Visible = False
    End If
    
End Sub

Private Sub cmdAttach_Click()
    
    Dim bWindowsAuth As Boolean, bSuccess As Boolean
    Dim strSQLServer As String
    Dim strDB As String, strUser As String, strPWD As String
    Dim strDBFile As String, strLogFile As String
    
    bWindowsAuth = optAuth(m_WINDOWS_AUTH).Value
    
    strSQLServer = txtSQLServer
    strDB = txtDBName
    strUser = txtUID
    If bWindowsAuth Then
        strPWD = ""
    Else
        strPWD = txtPWD
    End If
    strDBFile = "C:\MSSQL7\data\UPCManager_data.mdf"
    strLogFile = "C:\MSSQL7\data\UPCManager_log.ldf"
    
    bSuccess = m_objMaintain.AttachSQLDatabase(strSQLServer, strDB, strUser, strPWD, strDBFile, strLogFile, bWindowsAuth)
    
    If bSuccess Then
        MsgBox "The DB: " & strDB & " was Attached successfully!"
    End If
    
End Sub

Private Sub cmdDetach_Click()
    
    Dim bWindowsAuth As Boolean, bSuccess As Boolean
    Dim strSQLServer As String
    Dim strDB As String, strUser As String, strPWD As String
    
    bWindowsAuth = optAuth(m_WINDOWS_AUTH).Value
    
    strSQLServer = txtSQLServer
    strDB = txtDBName
    strUser = txtUID
    If bWindowsAuth Then
        strPWD = ""
    Else
        strPWD = txtPWD
    End If
    
    bSuccess = m_objMaintain.DetachSQLDatabase(strSQLServer, strDB, strUser, strPWD, bWindowsAuth)
    If bSuccess Then
        MsgBox "The DB: " & strDB & " was detached successfully!"
    End If
    
End Sub

Private Sub cmdRestoreFile_Click()
    
    On Error Resume Next
    
    Dim strFile As String

    With dlgTest
        .CancelError = True
        .DialogTitle = "Select SQL Backup file"
        .Filter = "SQL Backup File(*.bak) | *.bak"
        .InitDir = App.Path
        .ShowOpen
    End With

    If Err.Number <> 0 Then
        ' user cancelled
        strFile = ""
        txtRestoreFile.Text = ""
        lblMsg.Visible = True
        txtRestoreFile.SetFocus
        Exit Sub
    Else
        lblMsg.Visible = False
    End If

    strFile = dlgTest.FileName

    txtRestoreFile.Text = strFile
    
    If strFile = "" Then
        lblMsg.Visible = True
    Else
        lblMsg.Visible = False
    End If
End Sub

Private Sub m_objMaintain_ModuleError(ByVal ErrMessage As String)
    
    MsgBox ErrMessage, vbOKOnly + vbInformation, "Error"
    
End Sub

Private Sub optType_Click(Index As Integer)

    Select Case Index
        Case m_TYPE_BACKUP
            txtSQLBackup.Enabled = True
            txtSQLCopy.Enabled = True
            cmdLookupSQL.Enabled = True
            cmdCopyTo.Enabled = True
            cmdTestSQL.Enabled = True
            cmdRestoreFile.Enabled = False
            txtRestoreFile.Enabled = False
            cmdRestoreSQL.Enabled = False
        Case m_TYPE_RESTORE
            txtRestoreFile.Enabled = True
            cmdRestoreFile.Enabled = True
            cmdRestoreSQL.Enabled = True
            cmdTestSQL.Enabled = False
            txtSQLBackup.Enabled = False
            txtSQLCopy.Enabled = False
            cmdLookupSQL.Enabled = False
            cmdCopyTo.Enabled = False
    End Select

End Sub

Private Sub optAuth_Click(Index As Integer)

    Select Case Index
        Case m_SQL_AUTH
            txtPWD.Visible = True
            Me.lblTest(m_PWD_LBL).Visible = True
        Case m_WINDOWS_AUTH
            txtPWD.Visible = False
            Me.lblTest(m_PWD_LBL).Visible = False
    End Select

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdLookupSQL_Click()

    Dim strFile As String
    Dim strPrompt As String

    strPrompt = "Select the folder on the SERVER to save backup file"

    strFile = BrowseForFolder(Me.hWnd, strPrompt)

    txtSQLBackup.Text = strFile
    
    If txtSQLBackup.Text = "" Then
        lblMsg.Visible = True
    Else
        lblMsg.Visible = False
    End If
    
End Sub

Private Sub Form_Load()

    Set m_objMaintain = New cMaintainDB
    lblMsg.Caption = m_MESSAGE

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmSQLTest = Nothing

End Sub

Private Sub cmdTestSQL_Click()

    Dim bSuccess As Boolean
    Dim objMouse As cMouseCursor
    Dim strSQLServer As String, strUser As String, strPWD As String
    Dim strDB As String, strDestPath As String
    Dim strLocalCopy As String, bWindowsAuth As Boolean

    Set objMouse = New cMouseCursor
    objMouse.SetCursor vbHourglass

    If txtSQLBackup.Text = "" Then
        txtSQLBackup.SetFocus
        lblMsg.Visible = True
        Exit Sub
    Else
        lblMsg.Visible = False
    End If
    
    If txtSQLCopy.Text = "" Then
        txtSQLCopy.SetFocus
        lblMsg.Visible = True
        Exit Sub
    Else
        lblMsg.Visible = False
    End If
    
    bWindowsAuth = optAuth(m_WINDOWS_AUTH).Value
    
    strSQLServer = txtSQLServer
    strDB = txtDBName
    strUser = txtUID
    If bWindowsAuth Then
        strPWD = ""
    Else
        strPWD = txtPWD
    End If
    strDestPath = Me.txtSQLBackup
    strLocalCopy = Me.txtSQLCopy

    bSuccess = m_objMaintain.BackupSQLDatabase(strDestPath, strSQLServer, strDB, strUser, strPWD, strLocalCopy, bWindowsAuth)
    
End Sub

Private Sub cmdRestoreSQL_Click()

    Dim strSQLServer As String, strDB As String, strUser As String
    Dim strPWD As String, strBackupFile As String
    Dim objMouse As cMouseCursor, bWindowsAuth As Boolean
    Dim bCreateDB As Boolean
    
    Set objMouse = New cMouseCursor
    objMouse.SetCursor vbArrowHourglass

    If txtRestoreFile.Text = "" Then
        lblMsg.Visible = True
        txtRestoreFile.SetFocus
        Exit Sub
    Else
        lblMsg.Visible = False
    End If
    
    bWindowsAuth = optAuth(m_WINDOWS_AUTH).Value
    
    strSQLServer = txtSQLServer
    strDB = txtDBName
    strUser = txtUID
    If bWindowsAuth Then
        strPWD = ""
    Else
        strPWD = txtPWD
    End If
    strBackupFile = txtRestoreFile
    
    bCreateDB = chkCreate.Value
    
    m_objMaintain.RestoreSQLDatabase strSQLServer, strUser, strPWD, strDB, strBackupFile, bWindowsAuth, bCreateDB

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
    
    lblProgress.Visible = False
    
End Sub

Private Sub m_objMaintain_BackupLeftOnServer(ByVal FailMessage As String, ByVal FileOnServer As String)

    MsgBox FailMessage & vbCrLf & "File is: " & FileOnServer, vbOKOnly + vbInformation, "Cannot copy backup file"

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

Private Sub m_objMaintain_PercentComplete(ByVal ProgressMessage As String, ByVal PercentDone As Long)
    
    lblProgress.Visible = True
    lblProgress.Refresh
    lblProgress.Caption = PercentDone & "%"

End Sub

Private Sub m_objMaintain_RestoreError(ByVal FailMessage As String)

    MsgBox FailMessage, vbOKOnly + vbCritical, "Unable to restore DB"

End Sub

Private Sub m_objMaintain_RestoreFinished(ByVal Message As String, Database As String)

    lblProgress.Visible = False
    MsgBox Message, vbOKOnly + vbInformation, "Restore Successful!"
    
End Sub

