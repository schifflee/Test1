Option Compare Text


Module modGubbsDSN

    '   '***************************************************************
    '   '*****
    '   '*****	This script creates a DSN for connecting to a
    '   '*****	SQL Server database. To view errors comment out line 16
    '   '*****
    '   '*****	Script Name: AutoDSN.vbs
    '   '*****	Author: Darron Nesbitt
    '   '*****	Depends: VBScript, WScript Host
    '   '*****	Created: 10/2/2001
    '   '*****
    '   '***************************************************************

    '   'Values for variables on lines 25 - 29, 32, and 36
    '   'must be set prior to running this script.

    'On Error Resume Next

    '   Dim RegObj As Object
    '   Dim SysEnv

    'Set RegObj = WScript.CreateObject("WScript.Shell")

    '   '***** Specify the DSN parameters *****

    '   DataSourceName = "Name_of_Connection"
    '   DatabaseName = "Name_of_DB"
    '   Description = "Description of connection"
    '   LastUser = "Default_Username"
    '   Server = "Put_server_name_here"

    '   'if you use SQL Server the driver name would be "SQL Server"
    '   DriverName = "SQL Server"

    '   'Set this to True if Windows Authentication is used
    '   'else set to False or comment out
    '   WindowsAuthentication = True

    '   'point to DSN in registry
    '   REG_KEY_PATH = "HKLM\SOFTWARE\ODBC\ODBC.INI\" & DataSourceName

    '   ' Open the DSN key and check for Server entry
    ' 	lResult = RegObj.RegRead (REG_KEY_PATH & "\Server")

    '   'if lResult is nothing, DSN does not exist; create it
    ' 	if lResult = "" then

    '   'get os version through WSCript Enviroment object
    '	  Set SysEnv = RegObj.Environment("SYSTEM")
    '  OSVer = UCase(SysEnv("OS"))

    '   'check which os is running so correct driver path can be set
    '  Select Case OSVer
    '    Case "WINDOWS_NT"
    '        DrvrPath = "C:\WinNT\System32"
    '    Case Else
    '        DrvrPath = "C:\Windows\System"
    '  End Select

    '   'create entries in registry
    '  RegObj.RegWrite REG_KEY_PATH & "\DataBase",DatabaseName,"REG_SZ"
    '  RegObj.RegWrite REG_KEY_PATH & "\Description",Description,"REG_SZ"
    '  RegObj.RegWrite REG_KEY_PATH & "\LastUser",LastUser,"REG_SZ"
    '  RegObj.RegWrite REG_KEY_PATH & "\Server",Server,"REG_SZ"
    '  RegObj.RegWrite REG_KEY_PATH & "\Driver",DrvrPath,"REG_SZ"

    '   'if WindowsAuthentication set to True,
    '   'a trusted connection entry is added to registry
    '   'else, SQL Authentication is used.
    '  if WindowsAuthentication = True then
    '    RegObj.RegWrite REG_KEY_PATH & "\Trusted_Connection","Yes","REG_SZ"
    '  end if

    '   'point to data sources key
    '  REG_KEY_PATH = "HKLM\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources\" &
    '      DataSourceName

    '   'and add the name of the new dsn and the driver to use with it
    '  RegObj.RegWrite REG_KEY_PATH,DriverName,"REG_SZ"

    '  MsgBox DataSourceName & " DSN Created!"

    'else
    '	MsgBox DataSourceName & " DSN already exists!"
    'end if

    'Set RegObj = Nothing
    'Set SysEnv = Nothing



    '   '***************************************************************
    '   '  END AutoDSN.txt
    '   '***************************************************************




    '   '***************************************************************
    '   '*****
    '   '*****	VB_AutoDSN.txt
    '   '*****
    '   '***************************************************************

    '   Private Const REG_SZ = 1    'Constant for a string variable type.
    '   Private Const HKEY_LOCAL_MACHINE = &H80000002

    '   'Registry action types.
    '   Private Const ERROR_SUCCESS = 0&
    '   Private Const ERROR_NO_MORE_ITEMS = 259&
    '   Private Const REG_OPTION_NON_VOLATILE = 0
    '   Private Const KEY_QUERY_VALUE = &H1
    '   Private Const KEY_SET_VALUE = &H2
    '   Private Const KEY_CREATE_SUB_KEY = &H4
    '   Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    '   Private Const KEY_NOTIFY = &H10
    '   Private Const KEY_CREATE_LINK = &H20
    '   Private Const SYNCHRONIZE = &H100000
    '   Private Const STANDARD_RIGHTS_ALL = &H1F0000
    '   Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
    '       KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or _
    '       KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

    '   Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
    '      (ByVal hKey As Long, _
    '       ByVal lpSubKey As String, _
    '       ByVal ulOptions As Long, _
    '       ByVal samDesired As Long, _
    '   ByVal phkResult As Long) _
    '       As Long

    '   Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
    '      "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    '   ByVal phkResult As Long) As Long

    '   Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    '      "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    '      ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal _
    '      cbData As Long) As Long

    '   Private Declare Function RegCloseKey Lib "advapi32.dll" _
    '      (ByVal hKey As Long) As Long

    '   Public Declare Function GetVersionExA Lib "kernel32" _
    '       (ByVal lpVersionInformation As OSVERSIONINFO) As Integer

    '   Public Type OSVERSIONINFO
    '      dwOSVersionInfoSize As Long
    '      dwMajorVersion As Long
    '      dwMinorVersion As Long
    '      dwBuildNumber As Long
    '      dwPlatformId As Long
    '      szCSDVersion As String * 128
    '   End Type

    '   Public Function Chk_for_DSN(ByVal DataSourceName As String)

    '       ' ***********************************************
    '       ' Declare local usage variables.
    '       ' ***********************************************
    '       Dim dwResult As Long
    '       Dim dwType As Long, cbData As Long
    '       Dim REG_APP_KEYS_PATH As String

    '       Dim DataSourceName As String
    '       Dim DatabaseName As String
    '       Dim Description As String
    '       Dim DriverPath As String
    '       Dim DriverName As String
    '       Dim LastUser As String
    '       Dim Regional As String
    '       Dim Server As String
    '       Dim DrvrPath As String
    '       Dim OSVer As String
    '       Dim WindowsAuthentication As Boolean

    '       Dim lResult As Long
    '       Dim hKeyHandle As Long

    '       'Specify the DSN parameters.

    '       DatabaseName = "APR"
    '       Description = "CTS APR DB"
    '       DriverPath = DrvrPath
    '       LastUser = "ctsapr"
    '       Server = "CIITD010"
    '       DriverName = "SQL Server"
    '       WindowsAuthentication = True

    '       REG_APP_KEYS_PATH = "SOFTWARE\ODBC\ODBC.INI\" & DataSourceName

    '       ' ***********************************************
    '       ' Open the key for application's path.
    '       ' ***********************************************
    '  lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
    '                         REG_APP_KEYS_PATH, _
    '                         ByVal 0&, KEY_ALL_ACCESS, dwResult)
    '       If Not (lResult = ERROR_SUCCESS) Then


    '           OSVer = getVersion()

    '           Select Case OSVer
    '               Case "W2K"
    '                   DrvrPath = "C:\WinNT\System32"
    '               Case "NT4"
    '                   DrvrPath = "C:\WinNT\System32"
    '               Case "W95"
    '                   DrvrPath = "C:\Windows\System"
    '               Case "W98"
    '                   DrvrPath = "C:\Windows\System"
    '               Case "Failed"
    '                   MsgBox("Failed to get OS Version")
    '                   Exit Function
    '               Case Else
    '                   DrvrPath = "C:\WinNT\System32"
    '           End Select


    '           'Create the new DSN key.

    '           lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
    '                DataSourceName, hKeyHandle)

    '           'Set the values of the new DSN key.

    '      lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
    '         ByVal DatabaseName, Len(DatabaseName))
    '      lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
    '         ByVal Description, Len(Description))
    '      lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
    '         ByVal DrvrPath, Len(DrvrPath))
    '      lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
    '         ByVal LastUser, Len(LastUser))
    '      lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
    '         ByVal Server, Len(Server))

    '           If WindowsAuthentication = True Then
    '           lResult = RegSetValueEx(hKeyHandle, "Trusted_Connection", 0&, REG_SZ, _
    '                                   ByVal "Yes", Len("Yes"))
    '           End If

    '           'Close the new DSN key.

    '           lResult = RegCloseKey(hKeyHandle)

    '           'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
    '           'Specify the new value.
    '           'Close the key.

    '           lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
    '              "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
    '      lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
    '         ByVal DriverName, Len(DriverName))
    '           lResult = RegCloseKey(hKeyHandle)

    '           MsgBox(DataSourceName & " DSN created!")
    '       End If

    '   End Function

    '   Public Function getVersion() As String
    '       Dim osinfo As OSVERSIONINFO
    '       Dim retvalue As Integer

    '       osinfo.dwOSVersionInfoSize = 148
    '       osinfo.szCSDVersion = Space$(128)
    '       retvalue = GetVersionExA(osinfo)

    '       With osinfo
    '           Select Case .dwPlatformId
    '               Case 1
    '                   If .dwMinorVersion = 0 Then
    '                       getVersion = "W95"
    '                   ElseIf .dwMinorVersion = 10 Then
    '                       getVersion = "W98"
    '                   End If
    '               Case 2
    '                   If .dwMajorVersion = 3 Then
    '                       getVersion = "NT3"
    '                   ElseIf .dwMajorVersion = 4 Then
    '                       getVersion = "NT4"
    '                   ElseIf .dwMajorVersion = 5 Then
    '                       getVersion = "W2K"
    '                   End If
    '               Case Else
    '                   getVersion = "Failed"
    '           End Select
    '       End With
    '   End Function


End Module
