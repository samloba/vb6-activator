VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                                                 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                                                   (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long


Public Function EcrireDansFichierINI(Section As String, cle As String, _
                                     Valeur As String, Fichier As String) As Long
    EcrireDansFichierINI = WritePrivateProfileString(Section, cle, Valeur, Fichier)
End Function

Public Function LireDansFichierINI(Section As String, cle As String, Fichier As String, _
                                   Optional ValeurParDefaut As String = "") As String
    Dim strReturn As String
    strReturn = String(255, 0)
    GetPrivateProfileString Section, cle, ValeurParDefaut, strReturn, Len(strReturn), Fichier
    LireDansFichierINI = Left(strReturn, InStr(strReturn, Chr(0)) - 1)
End Function



