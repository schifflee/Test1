Public Class Security

    ' the default should be used.
    Private Const RPC_C_AUTHN_NONE As Long = 0
    Private Const RPC_C_AUTHN_WINNT As Long = 10
    Private Const RPC_C_AUTHN_DEFAULT As Long = &HFFFFFFFF

    ' Authentication level constants
    Private Const RPC_C_AUTHN_LEVEL_DEFAULT As Long = 0
    Private Const RPC_C_AUTHN_LEVEL_NONE As Long = 1
    Private Const RPC_C_AUTHN_LEVEL_CONNECT As Long = 2
    Private Const RPC_C_AUTHN_LEVEL_CALL As Long = 3
    Private Const RPC_C_AUTHN_LEVEL_PKT As Long = 4
    Private Const RPC_C_AUTHN_LEVEL_PKT_INTEGRITY As Long = 5
    Private Const RPC_C_AUTHN_LEVEL_PKT_PRIVACY As Long = 6

    ' Impersonation level constants
    Private Const RPC_C_IMP_LEVEL_ANONYMOUS As Long = 1
    Private Const RPC_C_IMP_LEVEL_IDENTIFY As Long = 2
    Private Const RPC_C_IMP_LEVEL_IMPERSONATE As Long = 3
    Private Const RPC_C_IMP_LEVEL_DELEGATE As Long = 4

    ' Constants for the capabilities
    Private Const API_NULL As Long = 0
    Private Const S_OK As Long = 0
    Private Const EOAC_NONE As Long = &H0
    Private Const EOAC_MUTUAL_AUTH As Long = &H1
    Private Const EOAC_CLOAKING As Long = &H10
    Private Const EOAC_SECURE_REFS As Long = &H2
    Private Const EOAC_ACCESS_CONTROL As Long = &H4
    Private Const EOAC_APPID As Long = &H8

    Public Sub New()
        Dim HRESULT As Long

        HRESULT = Security.CoInitializeSecurity(IntPtr.Zero, _
                                        -1, _
                                        IntPtr.Zero, _
                                        IntPtr.Zero, _
                                        RPC_C_AUTHN_LEVEL_NONE, _
                                        RPC_C_IMP_LEVEL_IMPERSONATE, _
                                        IntPtr.Zero, _
                                        EOAC_NONE, _
                                        IntPtr.Zero)
        If HRESULT <> S_OK Then
            'MsgBox(HRESULT)
            '''''''console.writeline(HRESULT.ToString)
            'MsgBox("CoInitializeSecurity failed with error code: 0x" & Trim$(Str$(Hex(HRESULT))), vbCritical, "Application Initialization Failure")
        End If

    End Sub

    Declare Function CoInitializeSecurity Lib "ole32.dll" (ByVal pVoid As IntPtr, _
    ByVal cAuthSvc As Integer, ByVal asAuthSvc As IntPtr, _
    ByVal pReserved1 As IntPtr, ByVal dwAuthnLevel As Integer, ByVal dwImpLevel As Integer, _
    ByVal pAuthList As IntPtr, ByVal dwCapabilities As Integer, ByVal pReserved3 As IntPtr) As Integer


End Class
