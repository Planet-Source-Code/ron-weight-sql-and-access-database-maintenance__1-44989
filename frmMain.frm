VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Testing DB Maintenance"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2715
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "End Testing"
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   2355
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test SQL Functions"
      Height          =   375
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   2355
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Access Functions"
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   2355
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click(Index As Integer)

    Select Case Index
        Case 0    ' test access
            frmAccessTest.Show vbModal
        Case 1    ' test sQL
            frmSQLTest.Show vbModal
        Case 2    ' end
            Unload Me
        Case Else
            MsgBox "Unexpected button pressed - index = " & Index
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmMain = Nothing

End Sub
