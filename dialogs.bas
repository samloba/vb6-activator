Attribute VB_Name = "dialogs"
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
Private Const MAX_PATH = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_EDITBOX = &H10

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 3

Private Const WM_USER = &H400

Private Const BFFM_SETSTATUSTEXT As Long = (WM_USER + 100)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTION As Long = (WM_USER + 102)
   
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                                 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                                   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long


Public Function EcrireDansFichierINI(ByVal Section As String, ByVal cle As String, _
                                     ByVal Valeur As String, ByVal Fichier As String) As Long
    EcrireDansFichierINI = WritePrivateProfileString(Section, cle, Valeur, Fichier)
End Function

Public Function LireDansFichierINI(ByVal Section As String, ByVal cle As String, ByVal Fichier As String, _
                                   Optional ByVal ValeurParDefaut As String = "") As String
    Dim strReturn As String
    strReturn = String(255, 0)
    GetPrivateProfileString Section, cle, ValeurParDefaut, strReturn, Len(strReturn), Fichier
    LireDansFichierINI = Left(strReturn, InStr(strReturn, Chr(0)) - 1)
End Function

Public Function ChoisirUnDossier(ByVal MsgEntete As String, Optional ByVal newfolder As Boolean = 0) As String
    Dim BI As BROWSEINFO, pidl As Long, lpSelPath As Long
    Dim spath As String * MAX_PATH
    Dim Start As String
    
    Start = CurDir

    'fill in the info it needs
    With BI
        .hOwner = GetForegroundWindow
        .pidlRoot = 0
        .lpszTitle = MsgEntete
        .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
        .ulFlags = BIF_RETURNONLYFSDIRS '+ BIF_EDITBOX
        If newfolder = True Then .ulFlags = BIF_RETURNONLYFSDIRS + BIF_EDITBOX + BIF_NEWDIALOGSTYLE
        lpSelPath = LocalAlloc(lPtr, Len(Start) + 1)
        CopyMemory ByVal lpSelPath, ByVal Start, Len(Start) + 1
        .lParam = lpSelPath
    End With

    'get the idlist long from the returned folder
    pidl = SHBrowseForFolder(BI)

    'do then if they clicked ok
    If pidl Then
        If SHGetPathFromIDList(pidl, spath) Then
            'next line is the returned folder
            ChoisirUnDossier = Left$(spath, InStr(spath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(pidl)
    Else
        'user clicked cancel
    End If

    Call LocalFree(lpSelPath)
End Function
'this seems to happen before the box comes up and when a folder is clicked on within it
Public Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Dim spath As String, bFlag As Long
    spath = Space$(MAX_PATH)
    Select Case uMsg
        Case BFFM_INITIALIZED
            'browse has been initialized, set the start folder
            Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal lpData)
        Case BFFM_SELCHANGED
            If SHGetPathFromIDList(lParam, spath) Then
                spath = Left(spath, InStr(1, spath, Chr(0)) - 1)
            End If
    End Select
End Function
Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

Sub gggggg()
    'Debug.Print ChoisirUnDossier("Select a folder", -1)
End Sub


