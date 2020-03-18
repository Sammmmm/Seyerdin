Attribute VB_Name = "modRegistry"
Option Explicit
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1
    Public Const REG_DWORD = 4
    Dim r As Long, lValueType As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Public Sub SaveKey(Hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Public Function GetString(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function


Public Sub SaveString(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Function GetDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function


Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Function


Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Public Sub CheckForBan()
On Error Resume Next
Dim Check As String, Check2 As String
Check = GetString(HKEY_CLASSES_ROOT, "Folders\", "Blarg")
Check2 = GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\hdProcess\", "Mhz")

If Exists("C:/Windows/readm.dat") = True Then
    MsgBox "Permabanned"
    End
End If

If Check = "Yep" Then
    MsgBox "Permabanned"
    End
End If

If Check2 = "0x0001(0)" Then
    MsgBox "Permabanned"
    End
End If
End Sub

'Permaban sucks, easily broken ..
Public Sub WriteBan()
On Error Resume Next
    Open "C:/Windows/readm.dat" For Output As #1
    Write #1, , "Bad."
    Close #1
    
    Call SaveString(HKEY_CLASSES_ROOT, "Folders\", "Blarg", "Yep")
    Call SaveString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\hdProcess\", "Mhz", "0x0001(0)")
    End
End Sub

Public Sub RemoveBan()
On Error Resume Next
    Kill App.Path & "/script.dat"
    Kill "C:/Windows/bant.dat"
    WriteString "Options", "Away", "0"
    Call DeleteKey(HKEY_CLASSES_ROOT, "Folders\")
    Call DeleteKey(HKEY_LOCAL_MACHINE, "Hardware\Description\System\hrdProcess\")
End Sub
