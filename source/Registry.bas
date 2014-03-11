Attribute VB_Name = "modRegistry"
Option Explicit

Public Const REG_NONE = 0 ' No value Type

Public Const REG_SZ = 1 ' Unicode nul terminated String

Public Const REG_EXPAND_SZ = 2 ' Unicode nul terminated String

Public Const REG_BINARY = 3 ' Free form binary

Public Const REG_DWORD = 4 ' 32-bit number

Public Const REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)

Public Const REG_DWORD_BIG_ENDIAN = 5 ' 32-bit number

Public Const REG_LINK = 6 ' Symbolic Link (unicode)

Public Const REG_MULTI_SZ = 7 ' Multiple Unicode strings

Public Const REG_RESOURCE_LIST = 8 ' Resource list In the resource map

Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9 ' Resource list In the hardware description

Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10



Public Enum hKeyNames

    HKEY_CLASSES_ROOT = &H80000000

    HKEY_CURRENT_USER = &H80000001

    HKEY_LOCAL_MACHINE = &H80000002

    HKEY_USERS = &H80000003

End Enum


Public Const ERROR_SUCCESS = 0&

Public Const ERROR_NONE = 0

Public Const ERROR_BADDB = 1

Public Const ERROR_BADKEY = 2

Public Const ERROR_CANTOPEN = 3

Public Const ERROR_CANTREAD = 4

Public Const ERROR_CANTWRITE = 5

Public Const ERROR_OUTOFMEMORY = 6

Public Const ERROR_ARENA_TRASHED = 7

Public Const ERROR_ACCESS_DENIED = 8

Public Const ERROR_INVALID_PARAMETERS = 87

Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0



Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long



Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long



Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long



Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long



Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long



Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long



Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long



Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long



Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long



Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long



Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long



Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long



Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long



Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long


    Dim lValue As Long

    Dim sValue As String



    Select Case lType

        Case REG_SZ

        sValue = vValue & Chr$(0)

        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))

        Case REG_DWORD

        lValue = vValue

        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)

    End Select


End Function




Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long


    Dim cch As Long

    Dim lrc As Long

    Dim lType As Long

    Dim lValue As Long

    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to

    '     be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)

    If lrc <> ERROR_NONE Then Error 5



    Select Case lType

        ' For strings

        Case REG_SZ, REG_EXPAND_SZ:

        sValue = String(cch, 0)

        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)



        If lrc = ERROR_NONE Then

            vValue = Left$(sValue, cch - 1)

        Else

            vValue = Empty

        End If


        ' For DWORDS

        Case REG_DWORD:

        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)

        If lrc = ERROR_NONE Then vValue = lValue

        Case Else

        'all other data types not supported

        lrc = -1

    End Select


QueryValueExExit:

QueryValueEx = lrc

Exit Function

QueryValueExError:

Resume QueryValueExExit

End Function




Public Function GetSettingEx(AppName As String, Section As String, Key As String, Optional default As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE")


    Dim lRetVal As Long 'result of the API functions

    Dim hKey As Long 'handle of opened key

    Dim vValue As Variant 'setting of queried value

    Dim keyString As String

    On Error GoTo e_Trap

    keyString = ""



    If AppNameHeader <> "" Then

        keyString = keyString + AppNameHeader

    End If




    If AppName <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & AppName

    End If




    If Section <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & Section

    End If


    lRetVal = RegOpenKeyEx(hKeyName, keyString, 0, KEY_ALL_ACCESS, hKey)

    lRetVal = QueryValueEx(hKey, Key, vValue)



    If IsEmpty(vValue) Then

        vValue = default

    End If


    GetSettingEx = vValue

    RegCloseKey (hKey)

    Exit Function

e_Trap:

    vValue = default

    Exit Function

End Function

Public Sub SetRegKey(ByVal hKeyName As hKeyNames, ByVal vsPath As String, ByVal vsKey As String, ByVal vsKeyValue As String)
    Dim lRetVal As Long 'result of the SetValueEx Function
    Dim hKey As Long 'handle of open key
    
    lRetVal = RegCreateKeyEx(hKeyName, vsPath, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
    'Todo: check return value
    lRetVal = SetValueEx(hKey, vsKey, REG_SZ, vsKeyValue)
    'Todo: check return value
    RegCloseKey (hKey)
End Sub

Public Function SaveSettingEx(AppName As String, Section As String, Key As String, Setting As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE") As Boolean


    Dim lRetVal As Long 'result of the SetValueEx Function

    Dim hKey As Long 'handle of open key

    Dim keyString As String

    On Error GoTo e_Trap

    keyString = ""



    If AppNameHeader <> "" Then

        keyString = keyString + AppNameHeader

    End If




    If AppName <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & AppName

    End If




    If Section <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & Section

    End If


    lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

    lRetVal = SetValueEx(hKey, Key, REG_SZ, Setting)

    RegCloseKey (hKey)

    SaveSettingEx = True

    Exit Function

e_Trap:

    SaveSettingEx = False

    Exit Function

End Function




Public Function DeleteSettingEx(AppName As String, Optional Section As String, Optional Key As String, Optional hKeyName As hKeyNames = HKEY_LOCAL_MACHINE, Optional AppNameHeader = "SOFTWARE") As Boolean


    Dim hNewKey As Long 'handle To the new key

    Dim lRetVal As Long 'result of the SetValueEx Function

    Dim hKey As Long 'handle of open key

    Dim keyString As String

    On Error GoTo e_Trap

    keyString = ""



    If AppNameHeader <> "" Then

        keyString = keyString + AppNameHeader

    End If




    If AppName <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & AppName

    End If




    If Section <> "" Then



        If keyString <> "" Then

            keyString = keyString & "\"

        End If


        keyString = keyString & Section

    End If




    If Key <> "" Then

        lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

        lRetVal = RegDeleteValue(hKey, Key)

        RegCloseKey (hKey)

    Else

        lRetVal = RegDeleteKey(hKeyName, keyString)

    End If


    DeleteSettingEx = True

    Exit Function

e_Trap:

    DeleteSettingEx = False

    Exit Function

End Function




Public Function AssociateFileType(extension As String, Optional useNotepadToEdit As Boolean = True) As Boolean


    Dim lRetVal As Long 'result of the SetValueEx Function

    Dim hKey As Long 'handle of open key

    Dim appPath As String

    On Error GoTo e_Trap

    



    If Mid(App.Path, Len(App.Path) - 1, 1) = "\" Then

        appPath = App.Path & App.EXEName & ".exe"

    Else

        appPath = App.Path & "\" & App.EXEName & ".exe"

    End If


    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

    lRetVal = SetValueEx(hKey, "", REG_SZ, App.Title & " App")

    RegCloseKey (hKey)

    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, "." & LCase(extension), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

    lRetVal = SetValueEx(hKey, "", REG_SZ, App.Title)

    RegCloseKey (hKey)

    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\shell\open\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

    lRetVal = SetValueEx(hKey, "", REG_SZ, appPath & " %1")

    RegCloseKey (hKey)



    If useNotepadToEdit = True Then

        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\shell\edit\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

        lRetVal = SetValueEx(hKey, "", REG_SZ, "notepad.exe %1")

        RegCloseKey (hKey)

    End If


    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, App.Title & "\DefaultIcon", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)

    lRetVal = SetValueEx(hKey, "", REG_SZ, appPath)

    RegCloseKey (hKey)

    AssociateFileType = True

    Exit Function

e_Trap:

    AssociateFileType = False

    Exit Function

End Function




Public Sub CreateRunOnStartup(Optional commandLine As String)




    If commandLine <> "" Then

        commandLine = " " & commandLine

    End If


    Call SaveSettingEx("CurrentVersion", "Run", App.Title, App.Path & "\" & App.EXEName & ".exe" & commandLine, HKEY_CURRENT_USER, "Software\Microsoft\Windows")

End Sub




Public Sub DeleteRunOnStartup()


    Call DeleteSettingEx("CurrentVersion", "Run", App.Title, HKEY_CURRENT_USER, "Software\Microsoft\Windows")

End Sub






Public Function GetIniInt(Section As String, Key As String, IniLocation As String, Optional default As Long) As Long


    GetIniInt = GetPrivateProfileInt(Section, Key, default, IniLocation)

End Function




Public Function GetIniString(Section As String, Key As String, IniLocation As String, Optional default As String) As String


    Dim ReturnValue As String * 128

    Dim i, sLet

    Dim iLen As Long

    Dim length As Long

    length = GetPrivateProfileString(Section, Key, default, ReturnValue, 128, IniLocation)

    i = InStr(1, Trim(ReturnValue), Chr(0))

    iLen = Len(Trim(ReturnValue))

    GetIniString = CStr(Left(Trim(ReturnValue), (i - 1)))

End Function




Public Function SaveIniString(Section As String, Key As String, Setting As String, IniLocation As String) As Long


    SaveIniString = WritePrivateProfileString(Section, Key, Setting, IniLocation)

End Function




Public Property Get Environ(variableName As String) As String


    Environ = GetSettingEx("Session Manager", "Environment", variableName, "", HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control")

End Property




Public Property Let Environ(variableName As String, Setting As String)


    Call SaveSettingEx("Session Manager", "Environment", variableName, Setting, HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control")

    Call SetEnvironmentVariable(variableName, Setting)

End Property




Public Sub VerifyPath(pathString As String)


    Dim CurrentPath As String

    pathString = Trim(pathString)

    If pathString = "" Then Exit Sub

    CurrentPath = Environ("PATH")



    If Mid(pathString, 1, 1) = ";" Then

        pathString = Mid(pathString, 2)

    End If




    If Mid(pathString, Len(pathString), 1) = ";" Then

        pathString = Mid(pathString, 1, Len(pathString) - 1)

    End If




    If InStr(1, UCase(CurrentPath), UCase(pathString), vbTextCompare) = 0 Then



        If Mid(CurrentPath, Len(CurrentPath), 1) = ";" Then

            Environ("PATH") = CurrentPath & pathString

        Else

            Environ("PATH") = CurrentPath & ";" & pathString

        End If


    End If


End Sub

