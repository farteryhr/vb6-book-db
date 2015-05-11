Attribute VB_Name = "Encode"
Option Explicit

Private Declare Function MultiByteToWideChar Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32 " (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_ACP = 0 ' default to ANSI code page
Private Const CP_UTF8 = 65001 ' default to UTF-8 code page
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'×Ö·û×ª UTF8
Public Function EncodeToBytes(ByVal sdata As String) As Byte() ' Note: Len(sData) > 0
Dim aRetn() As Byte
Dim nSize As Long
nSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sdata), Len(sdata), 0, 0, 0, 0)
If nSize = 0 Then Exit Function
ReDim aRetn(0 To nSize - 1) As Byte
WideCharToMultiByte CP_UTF8, 0, StrPtr(sdata), Len(sdata), VarPtr(aRetn(0)), nSize, 0, 0
EncodeToBytes = aRetn
Erase aRetn
End Function

' UTF8 ×ª×Ö·û
Public Function DecodeToBytes(ByVal sdata As String) As String ' Note: Len(sData) > 0
Dim aRetn() As Byte
Dim nSize As Long
nSize = MultiByteToWideChar(CP_UTF8, 0, StrPtr(sdata), LenB(sdata), 0, 0)
If nSize = 0 Then Exit Function
ReDim aRetn(0 To 2 * nSize - 1) As Byte
MultiByteToWideChar CP_UTF8, 0, StrPtr(sdata), LenB(sdata), VarPtr(aRetn(0)), nSize
DecodeToBytes = aRetn
Erase aRetn
End Function

Public Function SqlQuote(ByVal sdata As String, Optional ByVal asterisk As Boolean = False) As String
Dim ret As String
Dim nSize As Long
Dim xhn As Long, l As Long, char As Integer
sdata = Trim(sdata)
l = Len(sdata)
For xhn = 0 To l - 1
    char = AscW(Mid(sdata, xhn + 1, 1))
    If char = AscW("'") Then
        nSize = nSize + 2
    Else
        nSize = nSize + 1
    End If
Next xhn
ret = IIf(asterisk, "'%", "'") & String(nSize, " ") & IIf(asterisk, "%'", "'")
nSize = IIf(asterisk, 2, 1)
For xhn = 0 To l - 1
    char = AscW(Mid(sdata, xhn + 1, 1))
    If char = AscW("'") Then
        Mid(ret, nSize + 1, 2) = "''"
        nSize = nSize + 2
    ElseIf char = AscW(" ") And asterisk Then
        Mid(ret, nSize + 1, 1) = "%"
        nSize = nSize + 1
    Else
        Mid(ret, nSize + 1, 1) = Mid(sdata, xhn + 1, 1)
        nSize = nSize + 1
    End If
Next xhn
SqlQuote = ret
End Function

Public Function CNStr(f As Variant) As String
    If IsNull(f) Then
        CNStr = ""
    Else
        CNStr = f
    End If
End Function

Public Function CNumNStr(f As String) As String
    If f = "" Then
        CNumNStr = "null"
    Else
        CNumNStr = f
    End If
End Function

