Attribute VB_Name = "Mod_UTF8"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_ACP = 0
Private Const CP_UTF8 = 65001

Public Function UTF8_Encode(ByVal Text As String) As String

Dim sBuffer As String
Dim lLength As Long

lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
sBuffer = Space$(lLength)
lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
sBuffer = StrConv(sBuffer, vbUnicode)
UTF8_Encode = Left$(sBuffer, lLength - 1)

End Function

Public Function UTF8_Decode(ByVal Text As String) As String

Dim lLength As Long
Dim sBuffer As String

Text = StrConv(Text, vbFromUnicode)
lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, 0, 0)
sBuffer = Space$(lLength)
lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer))
UTF8_Decode = Left$(sBuffer, lLength - 1)

End Function

