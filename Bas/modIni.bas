Attribute VB_Name = "modIni"
'********************************************************
'**
'**模 块 名：modIni
'**
'**说    明：配置文件操作模块
'**
'********************************************************
Option Explicit

'APIs to access INI files and retrieve data
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Private Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)

Public Function GetIniParam(NomFichier As String, NomSection As String, NomVariable As String) As String
    Dim ReadString As String * 255
    Dim Returnv    As String
    Dim mResultLen As Integer
    
    mResultLen = GetPrivateProfileString(NomSection, NomVariable, "(Unassigned)", ReadString, Len(ReadString) - 1, NomFichier)
    If IsNull(ReadString) Or Left(ReadString, 12) = "(Unassigned)" Then
        Dim Tempvalue As Variant
        Dim Message As String
        Message = "配置文件 " & NomFichier & " 不存在！"
        Returnv = ""
    Else
        Returnv = Left(ReadString, InStr(ReadString, Chr$(0)) - 1)
    End If
    GetIniParam = Returnv
End Function

Public Function WriteIniParam(NomDuIni As String, sLaSection As String, sNouvelleCle As String, sNouvelleValeur As String)
    Dim iSucccess As Integer
    
    iSucccess = WritePrivateProfileStringByKeyName(sLaSection, sNouvelleCle, sNouvelleValeur, NomDuIni)
    If iSucccess = 0 Then
        WriteIniParam = False
    Else
        WriteIniParam = True
    End If
End Function
