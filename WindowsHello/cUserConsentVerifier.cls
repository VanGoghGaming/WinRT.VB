VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUserConsentVerifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncOperationCompletedHandlerUserConsentVerificationResult
Implements IAsyncOperationCompletedHandlerICredentialPickerResults

Private Const LOGON32_LOGON_INTERACTIVE As Long = &H2, LOGON32_PROVIDER_DEFAULT As Long = 0

Private Declare Function LogonUserW Lib "advapi32" (ByVal lpszUsername As LongPtr, ByVal lpszDomain As LongPtr, ByVal lpszPassword As LongPtr, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As LongPtr) As BOOL

Private m_CredentialPickerOptions As ICredentialPickerOptions, m_cAwait As cAwait, m_hWndParent As LongPtr, m_bCredentialPicker As Boolean

Public Event UserConsentVerificationResult(UserConsentVerificationResult As UserConsentVerificationResult)
Public Event CredentialPickerResult(UserConsentVerificationResult As UserConsentVerificationResult)

Friend Function CheckAvailability() As UserConsentVerifierAvailability
    CheckAvailability = -1
    If m_cAwait.Await(UserConsentVerifierStatics.CheckAvailabilityAsync) = AsyncStatus_Completed Then m_cAwait.GetResults VarPtr(CheckAvailability)
End Function

Friend Property Get ParentWindow() As LongPtr
    ParentWindow = m_hWndParent
End Property

Friend Property Let ParentWindow(hWndParent As LongPtr)
    m_hWndParent = hWndParent
End Property

Friend Sub ShowWindowsHello(sMessage As String)
Dim UserConsentVerifierInterop As IUserConsentVerifierInterop, vArray(0) As IAsyncOperationUserConsentVerificationResult
    If CheckAvailability = UserConsentVerifierAvailability_Available Then
        Set UserConsentVerifierInterop = UserConsentVerifierStatics
        Set UserConsentVerifierInterop.RequestVerificationForWindowAsync(m_hWndParent, StrRef(sMessage), GuidFromArray(vArray)).Completed = Me
    Else
        If Not m_bCredentialPicker Then m_bCredentialPicker = True: Set CredentialPickerStatics.PickWithOptionsAsync(m_CredentialPickerOptions).Completed = Me
    End If
End Sub

Private Sub Class_Initialize()
    Set m_cAwait = New cAwait: Set m_CredentialPickerOptions = NewObject("CredentialPickerOptions")
    With m_CredentialPickerOptions
        .Caption = StrRef("Authentication Required"): .Message = StrRef("Please enter your credentials"): .TargetName = StrRef(App.Title)
        .CredentialSaveOption = CredentialSaveOption_Hidden: .CallerSavesCredential = APITRUE: .AuthenticationProtocol = AuthenticationProtocol_Basic
    End With
End Sub

Private Sub IAsyncOperationCompletedHandlerUserConsentVerificationResult_Invoke(ByVal AsyncOperation As IAsyncOperationUserConsentVerificationResult, ByVal AsyncStatus As AsyncStatus)
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            RaiseEvent UserConsentVerificationResult(AsyncOperation.GetResults)
    End Select
End Sub

Private Sub IAsyncOperationCompletedHandlerICredentialPickerResults_Invoke(ByVal AsyncOperation As IAsyncOperationICredentialPickerResults, ByVal AsyncStatus As AsyncStatus)
Dim hToken As LongPtr
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            With AsyncOperation.GetResults
                If .ErrorCode = ERROR_CANCELLED Then
                    RaiseEvent CredentialPickerResult(UserConsentVerificationResult_Canceled)
                Else
                    If LogonUserW(StrPtr(GetStr(.CredentialUserName)), vbNullPtr, StrPtr(GetStr(.CredentialPassword)), LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken) Then
                        RaiseEvent CredentialPickerResult(UserConsentVerificationResult_Verified)
                    Else
                        RaiseEvent CredentialPickerResult(UserConsentVerificationResult_NotConfiguredForUser)
                    End If
                    If hToken Then CloseHandle hToken
                End If
            End With
    End Select
    m_bCredentialPicker = False
End Sub

