Attribute VB_Name = "modMain"
Option Explicit

' api's for Browse folders ...
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Private Const SPI_GETWORKAREA = 48

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

' UDTs needed for system info and rez functions and to browse folders ...
Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Function BrowseForFolder(hwndOwner As Long, PromptString As String) As String

    ' *********************************************************************
    ' Purpose:  Opens a Browse Folder dialog box
    ' Inputs:   Form handle and a Prompt to be shown
    ' Outputs:  Returns the folder selected (ONLY folders can be selected)
    ' *********************************************************************

    'declare variables to be used
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo

    'initialize variables
    With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(PromptString, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Call the browse for folder API
    lpIDList = SHBrowseForFolder(udtBI)

    'get the resulting string path
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
    End If

    'If cancel was pressed, sPath = ""
    BrowseForFolder = sPath

End Function

