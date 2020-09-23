Attribute VB_Name = "DNS"
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4
Private Const STANDARD_RIGHTS_ALL      As Long = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE  As Long = &H20000
Private Const STANDARD_RIGHTS_READ     As Long = &H20000
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_WRITE    As Long = &H20000
Private Const SYNCHRONIZE              As Long = &H100000
Private Const KEY_CREATE_LINK          As Long = &H20&
Private Const KEY_CREATE_SUB_KEY       As Long = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS   As Long = &H8&
Private Const KEY_EVENT                As Long = &H1&
Private Const KEY_NOTIFY               As Long = &H10&
Private Const KEY_QUERY_VALUE          As Long = &H1&
Private Const KEY_SET_VALUE            As Long = &H2&
Private Const KEY_WRITE                As Long = (STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And Not SYNCHRONIZE
Private Const KEY_READ                 As Long = (STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And Not SYNCHRONIZE
Private Const KEY_EXECUTE              As Long = KEY_READ&
Private Const KEY_ALL_ACCESS           As Long = (STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And Not SYNCHRONIZE



Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long ' Note that If you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Public Function isSZKeyExist(szKeyPath As String, szKeyName As String, _
            ByRef szKeyValue As String) As Boolean

    Dim bRes As Boolean
    Dim lRes As Long
    Dim hKey As Long

    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyPath, 0&, KEY_QUERY_VALUE, hKey)

    If lRes <> ERROR_SUCCESS Then

        isSZKeyExist = False
        Exit Function

    End If

    lRes = RegQueryValueEx(hKey, szKeyName, 0&, REG_SZ, ByVal szKeyValue, Len(szKeyValue))
    RegCloseKey (hKey)

    If lRes <> ERROR_SUCCESS Then

        isSZKeyExist = False
        Exit Function

    End If

    isSZKeyExist = True

End Function

Public Function checkAccessDriver(ByRef Path As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean
    bRes = False

    szKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\Microsoft Access Driver (*.mdb)"

    szKeyName = "Driver"

    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then

        Path = szKeyValue
        bRes = True

    Else

        bRes = False

    End If

    checkAccessDriver = bRes

End Function

Public Function checkWantedAccessDSN(szWantedDSN As String) As Boolean

    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim bRes As Boolean

    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"

    szKeyName = szWantedDSN

    szKeyValue = String(255, Chr(32))

    If isSZKeyExist(szKeyPath, szKeyName, szKeyValue) Then

        bRes = True

    Else
        bRes = False

    End If

    checkWantedAccessDSN = bRes

End Function

Public Function createAccessDSN(Path As String, _
            szWantedDSN As String) As Boolean

    Dim hKey As Long
    Dim szKeyPath As String
    Dim szKeyName As String
    Dim szKeyValue As String
    Dim lKeyValue As Long
    Dim lRes As Long
    Dim lSize As Long
    Dim szEmpty As String

    szEmpty = Chr(0)

    lSize = 4

    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, hKey)

    If lRes <> ERROR_SUCCESS Then

        createAccessDSN = False

        Exit Function

    End If

    lRes = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
            szEmpty, Len(szEmpty))

    szKeyValue = Path

    lRes = RegSetValueExString(hKey, "DBQ", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    szKeyValue = Path

    lRes = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    szKeyValue = "MS Access;"

    lRes = RegSetValueExString(hKey, "FIL", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lKeyValue = 25

    lRes = RegSetValueExLong(hKey, "DriverId", 0&, REG_DWORD, _
            lKeyValue, 4)

    lKeyValue = 0

    lRes = RegSetValueExLong(hKey, "SafeTransactions", 0&, REG_DWORD, _
            lKeyValue, 4)

    lRes = RegCloseKey(hKey)

    szKeyPath = "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines\Jet"
    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, szKeyPath, hKey)

    If lRes <> ERROR_SUCCESS Then

        createAccessDSN = False
        Exit Function

    End If

    lRes = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
            szEmpty, Len(szEmpty))

    szKeyValue = "Yes"

    lRes = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lKeyValue = 2048

    lRes = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, lKeyValue, 4)

    lKeyValue = 5

    lRes = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, lKeyValue, 4)

    lKeyValue = 3

    lRes = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, lKeyValue, 4)

    lRes = RegCloseKey(hKey)

    lRes = RegCreateKey(HKEY_LOCAL_MACHINE, _
            "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKey)

    If lRes <> ERROR_SUCCESS Then
        createAccessDSN = False
        Exit Function
    End If

    szKeyValue = "Microsoft Access Driver (*.mdb)"

    lRes = RegSetValueExString(hKey, szWantedDSN, 0&, REG_SZ, _
            szKeyValue, Len(szKeyValue))

    lRes = RegCloseKey(hKey)
    createAccessDSN = True

End Function


Public Function DeleteAccessDSN(szWantedDSN As String) As Boolean
    Dim lRes As Long
    Dim hKey As Long

    
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources\", 0&, KEY_ALL_ACCESS, hKey)
    If lRes <> ERROR_SUCCESS Then

      DeleteAccessDSN = False
      Exit Function

    End If
    
    lRes = RegDeleteValue(hKey, szWantedDSN)
    RegCloseKey (hKey)

    
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN & "\Engines", 0&, KEY_ALL_ACCESS, hKey)
    If lRes <> ERROR_SUCCESS Then

        DeleteAccessDSN = False
        Exit Function

    End If

    lRes = RegDeleteKey(hKey, "Jet")
    RegCloseKey (hKey)

    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & szWantedDSN, 0&, KEY_ALL_ACCESS, hKey)
    If lRes <> ERROR_SUCCESS Then

        DeleteAccessDSN = False
        Exit Function

    End If

    lRes = RegDeleteKey(hKey, "Engines")
    RegCloseKey (hKey)
     
    lRes = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI", 0&, KEY_ALL_ACCESS, hKey)
    If lRes <> ERROR_SUCCESS Then

        DeleteAccessDSN = False
        Exit Function

    End If

    lRes = RegDeleteKey(hKey, szWantedDSN)
    RegCloseKey (hKey)
End Function



