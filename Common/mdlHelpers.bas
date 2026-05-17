Attribute VB_Name = "mdlHelpers"
Option Explicit

Private Type VTable
    pVTable As LongPtr
    VTable() As LongPtr
End Type

Public Type HSTRING_HEADER
    Reserved1 As Long
    Reserved2 As Long
    Reserved3 As Long
    Reserved4 As Long
    Reserved5 As Long
    Reserved6 As Long
End Type

Private Type ASYNC_HANDLER_OBJECT
    pVTable As LongPtr
    cRefs As Long
    IID_AsyncOperationCompletedHandler As UUID
    hEvent As LongPtr
End Type

Private Type HSTRING_OBJECT
    pVTable As LongPtr
    cRefs As Long
    hStringPtr As HSTRING
End Type

Private Type SAFE_ARRAY_OBJECT
    pVTable As LongPtr
    cRefs As Long
    pSA As LongPtr
    bFreeData As Boolean
    SafeArray As SAFE_ARRAY
End Type

Private Type WEAKREF_OBJECT
    pVTable As LongPtr
    cRefs As Long
    WeakRef As IUnknown
End Type

Private Enum UNSUPPORTED_INTERFACE_IDENTIFIERS
    IID_IApplicationFrame = &H40EAA015 ' {143715D9-A015-40EA-B695-D5CC267E36EE}
    IID_IApplicationFrameEventHandler = &H4DA0770D ' {EA5D0DE4-770D-4DA0-A9F8-D7F9A140FF79}
    IID_IApplicationFrameManager = &H4413DBB9 ' {D6DEFAB3-DBB9-4413-8AF9-554586FDFF94}
    IID_IAudioDeviceGraph = &H484C37B2 ' {3C169FF7-37B2-484C-B199-C3155590F316}
    IID_ICallFactory = &H11CE2A1C ' {1C733A30-2A1C-11CE-ADE5-00AA0044773D}
    IID_IEUserBroker = &H4A69E6BB ' {1AC7516E-E6BB-4A69-B63F-E841904DC5A6}
    IID_IFrameTaskManager = &H49271B35 ' {35BD3360-1B35-4927-BAE4-B10E70D99EFF}
    IID_IMarshalOptions = &H4296E3E3 ' {4C1E39E1-E3E3-4296-AA86-EC938D896E92}
    IID_INoMarshal = &H4DC0C1DB ' {ECC8691B-C1DB-4DC0-855E-65F6C551AF49}
    IID_IPimcContext2 = &H415FAB5A ' {1868091E-AB5A-415F-A02F-5C4DD0CF901D}
    IID_IPrivateCallbackReceiver2 = &H490909D1 ' {11456F96-09D1-4909-8F36-4EB74E42B93E}
    IID_IProvideClassInfo = &H101ABAB4 ' {B196B283-BAB4-101A-B69C-00AA00341D07}
    IID_IRemUnknownTestHook = &H49FF50DC ' {2C258AE7-50DC-49FF-9D1D-2ECB9A52CDD7}
    IID_IStreamGroup = &H4EFC5523 ' {816E5B3E-5523-4EFC-9223-98EC4214C3A0}
    IID_Undocumented1 = &H3B150E79 ' {334D391F-0E79-3B15-C9FF-EAC65DD07C42}
    IID_Undocumented2 = &H2BC3139C ' {77DD1250-139C-2BC3-BD95-900ACED61BE5}
    IID_Undocumented3 = &H4E415A1F ' {BFD60505-5A1F-4E41-88BA-A6FB07202DA9}
    IID_Undocumented4 = &H4BB52289 ' {9BC79C93-2289-4BB5-ABF4-3287FD9CAE39}
    IID_Undocumented5 = &H45F5D534 ' {03FB5C57-D534-45F5-A1F4-D39556983875}
    IID_IVpoContext = &H4E9B6DED ' {4F4F92B5-6DED-4E9B-A93F-013891B3A8B7}
    IID_IVerbStateTaskCallBack = &H4474232E ' {F2153260-232E-4474-9D0A-9F2AB153441D}
End Enum

Private Const MaxConcurrentStringRefs As Long = 1024

Private m_VTableHSTRING As VTable, StringHeaders(0 To MaxConcurrentStringRefs) As HSTRING_HEADER, lCountStringRefs As Long, m_IID_IUnknown As UUID, m_IID_IHSTRING As UUID
Private m_VTableSafeArray As VTable, m_IID_ISafeArray As UUID
Private m_VTableWeakRef As VTable, m_IID_IWeakRef As UUID
Private m_VTableAsyncHandler As VTable, m_IID_IAgileObject As UUID
Private m_RuntimeClasses As IMap_HSTRING, m_IID_IMetaDataImport As UUID, m_IID_IActivationFactory As UUID, qpcFrequency As Currency, qpcStart As Currency

' ---------------------------------------------- IHSTRING ----------------------------------------------

Public Function NewString(Optional vString As Variant) As IHSTRING
Dim This As HSTRING_OBJECT
    With This
        If Not IsMissing(vString) Then If VarType(vString) = vbString Then WindowsCreateString StrPtr(vString), Len(vString), .hStringPtr Else .hStringPtr = vString
        .pVTable = m_VTableHSTRING.pVTable: .cRefs = 1
        If .pVTable = vbNullPtr Then
            If IsEmptyUUID(m_IID_IUnknown) Then m_IID_IUnknown = MakeUUID(, , &HC0, &H46000000)
            m_IID_IHSTRING = MakeUUID(&HF97000F6, &H4D03A6FC, &HF9722F9B, &H481D1392)
            .pVTable = GetVTablePointer(m_VTableHSTRING, AddressOf QueryInterfaceString, AddressOf AddRefString, AddressOf ReleaseString, AddressOf hStringGet, AddressOf hStringLet, _
                       AddressOf GetLength, AddressOf GetString, AddressOf GetStringReplace, AddressOf Compare, AddressOf Concat, AddressOf Detach, AddressOf Duplicate, AddressOf Substring, _
                       AddressOf SubstringLen)
        End If
        PutMemPtr ByVal VarPtr(NewString), vbaCopyBytes(LenB(This), ByVal CoTaskMemAlloc(LenB(This)), ByVal VarPtr(.pVTable))
    End With
End Function

Public Function GetStr(ByVal hStringPtr As HSTRING, Optional ByVal bPersistString As Boolean) As String
    SysReAllocString VarPtr(GetStr), WindowsGetStringRawBuffer(hStringPtr)
    If Not bPersistString Then WindowsDeleteString hStringPtr
End Function

Public Function StrRef(vString As Variant, Optional ByVal StringHeaderPtr As LongPtr) As HSTRING
    If StringHeaderPtr = vbNullPtr Then
        If lCountStringRefs > MaxConcurrentStringRefs Then lCountStringRefs = 0
        StringHeaderPtr = VarPtr(StringHeaders(lCountStringRefs)): lCountStringRefs = lCountStringRefs + 1
    End If
    WindowsCreateStringReference StrPtr(vString), Len(vString), StringHeaderPtr, StrRef
End Function

Public Function StrRefPtr(ByVal StringPtr As LongPtr, ByVal StringLength As Long, Optional ByVal StringHeaderPtr As LongPtr) As HSTRING
    If StringHeaderPtr = vbNullPtr Then
        If lCountStringRefs > MaxConcurrentStringRefs Then lCountStringRefs = 0
        StringHeaderPtr = VarPtr(StringHeaders(lCountStringRefs)): lCountStringRefs = lCountStringRefs + 1
    End If
    WindowsCreateStringReference StringPtr, StringLength, StringHeaderPtr, StrRefPtr
End Function

Public Function GetVTablePointer(VTable As VTable, ParamArray vaVTableAddresses() As Variant) As LongPtr
    GetVTablePointer = UBound(vaVTableAddresses) + 1
    If GetVTablePointer > 0 Then
        With VTable
            ReDim .VTable(0 To CLng(GetVTablePointer) - 1)
            For GetVTablePointer = 0 To GetVTablePointer - 1: .VTable(GetVTablePointer) = vaVTableAddresses(GetVTablePointer): Next GetVTablePointer
            .pVTable = VarPtr(.VTable(0)): GetVTablePointer = .pVTable
        End With
    End If
End Function

Public Function IsEmptyUUID(Guid As UUID, Optional ByVal bEmpty As Boolean) As Boolean
    With Guid
        If .Data1 = 0 Then If .Data2 = 0 Then If .Data3 = 0 Then IsEmptyUUID = .Data4 = 0
        If bEmpty Then If Not IsEmptyUUID Then .Data1 = 0: .Data2 = 0: .Data3 = 0: .Data4 = 0
    End With
End Function

Public Function IsEqualUUID(Guid1 As UUID, Guid2 As UUID) As Boolean
    With Guid1
        If .Data1 = Guid2.Data1 Then If .Data2 = Guid2.Data2 Then If .Data3 = Guid2.Data3 Then IsEqualUUID = .Data4 = Guid2.Data4
    End With
End Function

Public Function MakeUUID(Optional ByVal Data1 As Long, Optional ByVal Data2 As Long, Optional ByVal Data3 As Long, Optional ByVal Data4 As Long) As UUID
    With MakeUUID
        .Data1 = Data1: .Data2 = Data2: .Data3 = Data3: .Data4 = Data4
    End With
End Function

Private Function QueryInterfaceString(This As HSTRING_OBJECT, Guid As UUID, ppvObj As LongPtr) As HRESULT
    With This
        If IsEqualUUID(Guid, m_IID_IHSTRING) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        ElseIf IsEqualUUID(Guid, m_IID_IUnknown) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        Else
            ppvObj = vbNullPtr: QueryInterfaceString = E_NOINTERFACE
        End If
    End With
End Function

Private Function AddRefString(This As HSTRING_OBJECT) As Long
    AddRefString = This.cRefs + 1: This.cRefs = AddRefString
End Function

Private Function ReleaseString(This As HSTRING_OBJECT) As Long
    With This
        .cRefs = .cRefs - 1: ReleaseString = .cRefs
        If .cRefs = 0 Then WindowsDeleteString .hStringPtr: CoTaskMemFree VarPtr(.pVTable)
    End With
End Function

Private Function hStringGet(This As HSTRING_OBJECT, ByVal bReplace As Boolean, hStringPtr As HSTRING) As HRESULT
    hStringPtr = This.hStringPtr
End Function

Private Function hStringLet(This As HSTRING_OBJECT, ByVal bReplace As Boolean, ByVal hStringPtr As HSTRING) As HRESULT
    With This
        If bReplace Then WindowsDeleteString .hStringPtr
        .hStringPtr = hStringPtr
    End With
End Function

Private Function GetLength(This As HSTRING_OBJECT, Length As Long) As HRESULT
    Length = WindowsGetStringLen(This.hStringPtr)
End Function

Private Function GetString(This As HSTRING_OBJECT, pStr As LongPtr) As HRESULT
    pStr = SysAllocString(WindowsGetStringRawBuffer(This.hStringPtr))
End Function

Private Function GetStringReplace(This As HSTRING_OBJECT, ByVal hNewString As HSTRING, pStr As LongPtr) As HRESULT
    With This
        WindowsDeleteString .hStringPtr: .hStringPtr = hNewString: pStr = SysAllocString(WindowsGetStringRawBuffer(.hStringPtr))
    End With
End Function

Private Function Compare(This As HSTRING_OBJECT, ByVal hCompareString As HSTRING, lResult As Long) As HRESULT
    WindowsCompareStringOrdinal This.hStringPtr, hCompareString, lResult: WindowsDeleteString hCompareString
End Function

Private Function Concat(This As HSTRING_OBJECT, ByVal hConcatString As LongPtr, ByVal bReverse As Boolean, hNewString As HSTRING) As HRESULT
    With This
        If Not bReverse Then WindowsConcatString .hStringPtr, hConcatString, hNewString Else WindowsConcatString hConcatString, .hStringPtr, hNewString
        WindowsDeleteString .hStringPtr: WindowsDeleteString hConcatString: .hStringPtr = hNewString
    End With
End Function

Private Function Detach(This As HSTRING_OBJECT, hStringPtr As HSTRING) As HRESULT
    With This
        hStringPtr = .hStringPtr: .hStringPtr = vbNullPtr
    End With
End Function

Private Function Duplicate(This As HSTRING_OBJECT, hNewString As HSTRING) As HRESULT
    WindowsDuplicateString This.hStringPtr, hNewString
End Function

Private Function Substring(This As HSTRING_OBJECT, ByVal StartIndex As Long, ByVal bReplace As Boolean, hNewString As HSTRING) As HRESULT
    With This
        WindowsSubstring .hStringPtr, StartIndex, hNewString
        If bReplace Then WindowsDeleteString .hStringPtr: .hStringPtr = hNewString
    End With
End Function

Private Function SubstringLen(This As HSTRING_OBJECT, ByVal StartIndex As Long, ByVal Length As Long, ByVal bReplace As Boolean, hNewString As HSTRING) As HRESULT
    With This
        WindowsSubstringWithSpecifiedLength .hStringPtr, StartIndex, Length, hNewString
        If bReplace Then WindowsDeleteString .hStringPtr: .hStringPtr = hNewString
    End With
End Function

' ---------------------------------------------- ISafeArray ----------------------------------------------

Public Function NewSA(Optional ByVal pSA As LongPtr, Optional ByVal cbElements As Long = 1, Optional ByVal pvData As LongPtr, Optional ByVal cElements1 As Long = 1, Optional ByVal cElements2 As Long, Optional ByVal lLBound1 As Long, Optional ByVal lLBound2 As Long, Optional ByVal bFreeData As Boolean) As ISafeArray
Dim tSA As SAFE_ARRAY, tSAOBJ As SAFE_ARRAY_OBJECT, SafeArrayObject() As SAFE_ARRAY_OBJECT
    InitSA ArrPtr(SafeArrayObject), tSA, cbElements:=LenB(tSAOBJ), pvData:=CoTaskMemAlloc(LenB(tSAOBJ))
    With SafeArrayObject(0)
        With .SafeArray
            .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE: .cLocks = 1: If cElements2 = 0 Then .cDims = 1 Else .cDims = 2
            .pvData = pvData: .cbElements = cbElements: .cElements1 = cElements1: .cElements2 = cElements2: .lLBound1 = lLBound1: .lLBound2 = lLBound2
            If pSA Then PutMemPtr ByVal pSA, VarPtr(.cDims)
        End With
        .pVTable = m_VTableSafeArray.pVTable: .cRefs = 1: .pSA = pSA: .bFreeData = bFreeData
        If .pVTable = vbNullPtr Then
            If IsEmptyUUID(m_IID_IUnknown) Then m_IID_IUnknown = MakeUUID(, , &HC0, &H46000000)
            m_IID_ISafeArray = MakeUUID(&H7763C956, &H41F9CBC4, &H5D58439C, &H54D02C78)
            .pVTable = GetVTablePointer(m_VTableSafeArray, AddressOf QueryInterfaceSafeArray, AddressOf AddRefSafeArray, AddressOf ReleaseSafeArray, AddressOf BindToArray, AddressOf SafeArrayGet, _
                       AddressOf ToByteString, AddressOf ToHexString, AddressOf LengthGet, AddressOf pvDataGet, AddressOf pvDataLet, AddressOf cbElementsGet, AddressOf cbElementsLet, _
                       AddressOf cElements1Get, AddressOf cElements1Let, AddressOf cElements2Get, AddressOf cElements2Let)
        End If
        PutMemPtr ByVal VarPtr(NewSA), VarPtr(.pVTable)
    End With
End Function

Public Sub InitSA(ByVal pSA As LongPtr, tSA As SAFE_ARRAY, Optional ByVal cbElements As Long = 1, Optional ByVal pvData As LongPtr, Optional ByVal cElements1 As Long = 1, Optional ByVal cElements2 As Long, Optional ByVal lLBound1 As Long, Optional ByVal lLBound2 As Long)
    With tSA
        If .fFeatures = 0 Then PutMemPtr ByVal pSA, VarPtr(.cDims): .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE: .cLocks = 1: If cElements2 = 0 Then .cDims = 1 Else .cDims = 2
        .pvData = pvData: .cbElements = cbElements: .cElements1 = cElements1: .lLBound1 = lLBound1: .cElements2 = cElements2: .lLBound2 = lLBound2
    End With
End Sub

Public Property Get SafeArray(ByVal pSA As LongPtr) As SAFE_ARRAY
    GetMemPtr ByVal pSA, pSA
    If pSA Then
        pSA = pSA - LenB(SafeArray.Guid): vbaCopyBytes LenB(SafeArray) - 8, ByVal VarPtr(SafeArray), ByVal pSA
        If SafeArray.cDims = 2 Then vbaCopyBytes LenB(SafeArray), ByVal VarPtr(SafeArray), ByVal pSA
    End If
End Property

Private Function QueryInterfaceSafeArray(This As SAFE_ARRAY_OBJECT, Guid As UUID, ppvObj As LongPtr) As HRESULT
    With This
        If IsEqualUUID(Guid, m_IID_ISafeArray) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        ElseIf IsEqualUUID(Guid, m_IID_IUnknown) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        Else
            ppvObj = vbNullPtr: QueryInterfaceSafeArray = E_NOINTERFACE
        End If
    End With
End Function

Private Function AddRefSafeArray(This As SAFE_ARRAY_OBJECT) As Long
    AddRefSafeArray = This.cRefs + 1: This.cRefs = AddRefSafeArray
End Function

Private Function ReleaseSafeArray(This As SAFE_ARRAY_OBJECT) As Long
    With This
        ReleaseSafeArray = .cRefs - 1: .cRefs = ReleaseSafeArray
        If .cRefs = 0 Then
            If .pSA Then PutMemPtr ByVal .pSA, vbNullPtr
            If .bFreeData Then If .SafeArray.pvData Then CoTaskMemFree .SafeArray.pvData
            CoTaskMemFree VarPtr(.pVTable)
        End If
    End With
End Function

Private Function BindToArray(This As SAFE_ARRAY_OBJECT, ByVal pSA As LongPtr) As HRESULT
    If pSA Then
        With This
            If .pSA Then PutMemPtr ByVal .pSA, vbNullPtr
            .pSA = pSA: PutMemPtr ByVal pSA, VarPtr(.SafeArray.cDims)
        End With
    End If
End Function

Private Function SafeArrayGet(This As SAFE_ARRAY_OBJECT, SafeArray As SAFE_ARRAY) As HRESULT
    SafeArray = This.SafeArray
End Function

Private Function ToByteString(This As SAFE_ARRAY_OBJECT, sDelimiter As String, sToByteString As String) As HRESULT
Dim i As Long, cBytes As Long, CurrentByte As Byte
    With This.SafeArray
        If .pvData Then
            LengthGet This, cBytes: cBytes = cBytes * .cbElements
            While i < cBytes
                GetMem1 ByVal .pvData + i, CurrentByte
                If i < cBytes - 1 Then sToByteString = sToByteString & CurrentByte & sDelimiter Else sToByteString = sToByteString & CurrentByte
                i = i + 1
            Wend
        End If
    End With
End Function

Private Function ToHexString(This As SAFE_ARRAY_OBJECT, sDelimiter As String, ByVal bUpperCase As Boolean, sToHexString As String) As HRESULT
Dim i As Long, cBytes As Long, CurrentByte As Byte, sCurrentByte As String
    With This.SafeArray
        If .pvData Then
            LengthGet This, cBytes: cBytes = cBytes * .cbElements
            While i < cBytes
                GetMem1 ByVal .pvData + i, CurrentByte: If CurrentByte > &HF Then sCurrentByte = Hex$(CurrentByte) Else sCurrentByte = ChrW$(48) & Hex$(CurrentByte)
                If i < cBytes - 1 Then sToHexString = sToHexString & sCurrentByte & sDelimiter Else sToHexString = sToHexString & sCurrentByte
                i = i + 1
            Wend
            If Not bUpperCase Then sToHexString = LCase$(sToHexString)
        End If
    End With
End Function

Private Function LengthGet(This As SAFE_ARRAY_OBJECT, lLength As Long) As HRESULT
    With This.SafeArray
        lLength = .cElements1: If .cElements2 Then lLength = lLength * .cElements2
    End With
End Function

Private Function pvDataGet(This As SAFE_ARRAY_OBJECT, pvData As LongPtr) As HRESULT
    pvData = This.SafeArray.pvData
End Function

Private Function pvDataLet(This As SAFE_ARRAY_OBJECT, ByVal pvData As LongPtr) As HRESULT
    This.SafeArray.pvData = pvData
End Function

Private Function cbElementsGet(This As SAFE_ARRAY_OBJECT, cbElements As Long) As HRESULT
    cbElements = This.SafeArray.cbElements
End Function

Private Function cbElementsLet(This As SAFE_ARRAY_OBJECT, ByVal cbElements As Long) As HRESULT
    This.SafeArray.cbElements = cbElements
End Function

Private Function cElements1Get(This As SAFE_ARRAY_OBJECT, lLBound1 As Long, cElements1 As Long) As HRESULT
    With This.SafeArray
        lLBound1 = .lLBound1: cElements1 = .cElements1
    End With
End Function

Private Function cElements1Let(This As SAFE_ARRAY_OBJECT, lLBound1 As Long, ByVal cElements1 As Long) As HRESULT
    With This.SafeArray
        .lLBound1 = lLBound1: .cElements1 = cElements1
    End With
End Function

Private Function cElements2Get(This As SAFE_ARRAY_OBJECT, lLBound2 As Long, cElements2 As Long) As HRESULT
    With This.SafeArray
        lLBound2 = .lLBound2: cElements2 = .cElements2
    End With
End Function

Private Function cElements2Let(This As SAFE_ARRAY_OBJECT, lLBound2 As Long, ByVal cElements2 As Long) As HRESULT
    With This.SafeArray
        .lLBound2 = lLBound2: .cElements2 = cElements2
    End With
End Function

' ---------------------------------------------- IWeakRef ----------------------------------------------

Public Function NewWeakRef(ByVal pWeakReference As LongPtr) As IWeakRef
Dim This As WEAKREF_OBJECT
    With This
        PutMemPtr ByVal VarPtr(.WeakRef), pWeakReference: .pVTable = m_VTableWeakRef.pVTable: .cRefs = 1
        If .pVTable = vbNullPtr Then
            If IsEmptyUUID(m_IID_IUnknown) Then m_IID_IUnknown = MakeUUID(, , &HC0, &H46000000)
            m_IID_IWeakRef = MakeUUID(&HD70947F7, &H4EFC46E5, &H68B118B8, &H380D0683)
            .pVTable = GetVTablePointer(m_VTableWeakRef, AddressOf QueryInterfaceWeakRef, AddressOf AddRefWeakRef, AddressOf ReleaseWeakRef, AddressOf StillAliveGet, AddressOf StillAliveLet, _
                       AddressOf Resolve)
        End If
        PutMemPtr ByVal VarPtr(NewWeakRef), vbaCopyBytesZero(LenB(This), ByVal CoTaskMemAlloc(LenB(This)), ByVal VarPtr(.pVTable))
    End With
End Function

Private Function QueryInterfaceWeakRef(This As WEAKREF_OBJECT, Guid As UUID, ppvObj As LongPtr) As HRESULT
    With This
        If IsEqualUUID(Guid, m_IID_IWeakRef) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        ElseIf IsEqualUUID(Guid, m_IID_IUnknown) Then
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        Else
            ppvObj = vbNullPtr: QueryInterfaceWeakRef = E_NOINTERFACE
        End If
    End With
End Function

Private Function AddRefWeakRef(This As WEAKREF_OBJECT) As Long
    AddRefWeakRef = This.cRefs + 1: This.cRefs = AddRefWeakRef
End Function

Private Function ReleaseWeakRef(This As WEAKREF_OBJECT) As Long
    With This
        ReleaseWeakRef = .cRefs - 1: .cRefs = ReleaseWeakRef
        If .cRefs = 0 Then
            If Not .WeakRef Is Nothing Then PutMemPtr ByVal VarPtr(.WeakRef), vbNullPtr
            CoTaskMemFree VarPtr(.pVTable)
        End If
    End With
End Function

Private Function StillAliveGet(This As WEAKREF_OBJECT, bStillAlive As Boolean) As HRESULT
    bStillAlive = Not This.WeakRef Is Nothing
End Function

Private Function StillAliveLet(This As WEAKREF_OBJECT, ByVal bStillAlive As Boolean) As HRESULT
    If Not bStillAlive Then PutMemPtr ByVal VarPtr(This.WeakRef), vbNullPtr
End Function

Private Function Resolve(This As WEAKREF_OBJECT, WeakRef As IUnknown) As HRESULT
    Set WeakRef = This.WeakRef
End Function

' ---------------------------------------------- AsyncOperationCompletedHandler ----------------------------------------------

Public Function NewAsyncOperationCompletedHandler(ByVal hEvent As LongPtr) As IUnknown
Dim This As ASYNC_HANDLER_OBJECT
    With This
        .pVTable = m_VTableAsyncHandler.pVTable: .cRefs = 1: .hEvent = hEvent
        If .pVTable = vbNullPtr Then
            If IsEmptyUUID(m_IID_IUnknown) Then m_IID_IUnknown = MakeUUID(, , &HC0, &H46000000)
            m_IID_IAgileObject = MakeUUID(&H94EA2B94, &H49E0E9CC, &H64EEFFC0, &H905B8FCA)
            .pVTable = GetVTablePointer(m_VTableAsyncHandler, AddressOf QueryInterfaceAsyncHandler, AddressOf AddRefAsyncHandler, AddressOf ReleaseAsyncHandler, AddressOf InvokeAsyncHandler)
        End If
        PutMemPtr ByVal VarPtr(NewAsyncOperationCompletedHandler), vbaCopyBytes(LenB(This), ByVal CoTaskMemAlloc(LenB(This)), ByVal VarPtr(.pVTable))
    End With
End Function

Private Function QueryInterfaceAsyncHandler(This As ASYNC_HANDLER_OBJECT, Guid As UUID, ppvObj As LongPtr) As HRESULT
    With This
        If IsEqualUUID(Guid, .IID_AsyncOperationCompletedHandler) Then ' We don't know the correct IID of this Completed Handler just yet, but we'll find it below soon enough! :)
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        ElseIf IsEqualUUID(Guid, m_IID_IUnknown) Then ' All objects must support the ubiquitous IUnknown interface
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        ElseIf IsEqualUUID(Guid, m_IID_IAgileObject) Then ' Mark this object as "Agile" so its methods can be called from any thread
            .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable)
        Else
            Select Case Guid.Data2
                Case 0, IID_IApplicationFrame, IID_IApplicationFrameEventHandler, IID_IApplicationFrameManager, IID_IAudioDeviceGraph, IID_ICallFactory, IID_IEUserBroker, IID_IFrameTaskManager, _
                     IID_IMarshalOptions, IID_INoMarshal, IID_IPimcContext2, IID_IPrivateCallbackReceiver2, IID_IRemUnknownTestHook, IID_IStreamGroup, IID_Undocumented1, IID_Undocumented2, _
                     IID_Undocumented3, IID_Undocumented4, IID_Undocumented5, IID_IVpoContext, IID_IVerbStateTaskCallBack: ppvObj = vbNullPtr: QueryInterfaceAsyncHandler = E_NOINTERFACE
                Case IID_IProvideClassInfo: ppvObj = vbNullPtr: QueryInterfaceAsyncHandler = E_NOTIMPL ' returning E_NOINTERFACE would crash VB6 in debugging mode when hovering the mouse over the variable's name
                Case Else
                    If IsEmptyUUID(.IID_AsyncOperationCompletedHandler) Then
                        .IID_AsyncOperationCompletedHandler = Guid: .cRefs = .cRefs + 1: ppvObj = VarPtr(.pVTable) ' This is the correct IID that corresponds to the Completed Handler of this AsyncOperation
                    Else
                        ppvObj = vbNullPtr: QueryInterfaceAsyncHandler = E_NOINTERFACE
                    End If
            End Select
        End If
    End With
End Function

Private Function AddRefAsyncHandler(This As ASYNC_HANDLER_OBJECT) As Long
    AddRefAsyncHandler = This.cRefs + 1: This.cRefs = AddRefAsyncHandler
End Function

Private Function ReleaseAsyncHandler(This As ASYNC_HANDLER_OBJECT) As Long
    ReleaseAsyncHandler = This.cRefs - 1: This.cRefs = ReleaseAsyncHandler: If ReleaseAsyncHandler = 0 Then CoTaskMemFree VarPtr(This)
End Function

Private Function InvokeAsyncHandler(This As ASYNC_HANDLER_OBJECT, ByVal AsyncOperation As IInspectable, ByVal AsyncStatus As AsyncStatus) As HRESULT
    SetEvent This.hEvent ' SetEvent is being called from another thread and must be declared in a TypeLib in VB6. This signals the completion of the AsyncOperation and exits the waiting loop
End Function

' ---------------------------------------------- ISubclass ----------------------------------------------

Public Function GetSubclassPointer(ByVal ObjectToSubclass As Object) As LongPtr
Dim Subclass As ISubclass
    If Not ObjectToSubclass Is Nothing Then If TypeOf ObjectToSubclass Is ISubclass Then Set Subclass = ObjectToSubclass: GetSubclassPointer = ObjPtr(Subclass)
End Function

Public Function IsWndSubclassed(ByVal hWnd As LongPtr, ObjectToSubclass As Variant, Optional dwRefData As LongPtr) As Boolean
Dim uIdSubclass As LongPtr
    If IsObject(ObjectToSubclass) Then uIdSubclass = GetSubclassPointer(ObjectToSubclass) Else uIdSubclass = ObjectToSubclass
    IsWndSubclassed = GetWindowSubclass(hWnd, AddressOf WndProc, uIdSubclass, dwRefData) = APITRUE
End Function

Public Function SubclassWnd(ByVal hWnd As LongPtr, ObjectToSubclass As Variant, Optional ByVal dwRefData As LongPtr, Optional SubclassDLL As Object) As Boolean
Dim uIdSubclass As LongPtr, dwOldRefData As LongPtr
    If SubclassDLL Is Nothing Then
        If IsObject(ObjectToSubclass) Then uIdSubclass = GetSubclassPointer(ObjectToSubclass) Else uIdSubclass = ObjectToSubclass
        If uIdSubclass Then
            If Not IsWndSubclassed(hWnd, uIdSubclass, dwOldRefData) Then
                SubclassWnd = SetWindowSubclass(hWnd, AddressOf WndProc, uIdSubclass, dwRefData) = APITRUE
            Else
                If dwOldRefData <> dwRefData Then SubclassWnd = SetWindowSubclass(hWnd, AddressOf WndProc, uIdSubclass, dwRefData) = APITRUE
            End If
        End If
    Else
        SubclassWnd = SubclassDLL.SubclassWnd(hWnd, dwRefData) ' Forward subclassing to an ActiveX DLL for IDE protection
    End If
End Function

Public Function UnSubclassWnd(ByVal hWnd As LongPtr, ObjectToSubclass As Variant, Optional SubclassDLL As Object) As Boolean
Dim uIdSubclass As LongPtr
    If SubclassDLL Is Nothing Then
        If IsObject(ObjectToSubclass) Then uIdSubclass = GetSubclassPointer(ObjectToSubclass) Else uIdSubclass = ObjectToSubclass
        If uIdSubclass Then If IsWndSubclassed(hWnd, uIdSubclass) Then UnSubclassWnd = RemoveWindowSubclass(hWnd, AddressOf WndProc, uIdSubclass) = APITRUE
    Else
        UnSubclassWnd = SubclassDLL.UnSubclassWnd(hWnd)
    End If
End Function

Private Function WndProc(ByVal hWnd As LongPtr, ByVal uMsg As WINDOWS_MESSAGES, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal Subclass As ISubclass, ByVal dwRefData As LongPtr) As LongPtr
Dim bDiscardMessage As Boolean
    Select Case uMsg
        Case WM_NCDESTROY ' Remove subclassing as the window is about to be destroyed
            RemoveWindowSubclass hWnd, AddressOf mdlHelpers.WndProc, ObjPtr(Subclass)
        Case Else
            Subclass.MessageReceived hWnd, uMsg, wParam, lParam, dwRefData, bDiscardMessage, WndProc ' bDiscardMessage can be toggled as required by each local ISubclass_MessageReceived
    End Select
    If Not bDiscardMessage Then WndProc = DefSubclassProc(hWnd, uMsg, wParam, lParam) ' Choose whether to pass along this message or discard it
End Function

' ---------------------------------------------- Helper Functions ----------------------------------------------

Public Function ActivateInstance(ActivationFactory As IActivationFactory) As IInspectable: ActivationFactory.ActivateInstance ActivateInstance: End Function

Public Function AppPath() As String
Static sAppPath As String
    If Len(sAppPath) = 0 Then
        #If (VBA7 <> False Or VBA6 <> False) And TWINBASIC = False Then
            sAppPath = Application.VBE.ActiveVBProject.FileName
            If InStr(sAppPath, ChrW$(92)) > 0 Then sAppPath = Left$(sAppPath, InStrRev(sAppPath, ChrW$(92)) - 1)
        #Else
            sAppPath = App.Path
        #End If
    End If
    AppPath = sAppPath
End Function

Public Function ArrPtrVar(vArray As Variant) As LongPtr
Dim vt As Integer
    If IsArray(vArray) Then
        GetMem2 ByVal VarPtr(vArray), vt: ArrPtrVar = VarPtr(vArray) + 8
        If (vt And VT_BYREF) = VT_BYREF Then GetMemPtr ByVal ArrPtrVar, ArrPtrVar
    End If
End Function

Public Function Base64Url(Buffer As IBuffer) As String
Dim waBase64() As Integer, tSA As SAFE_ARRAY, i As Long
    If Not Buffer Is Nothing Then
        Base64Url = GetStr(CryptographicBufferStatics.EncodeToBase64String(Buffer))
        i = Len(Base64Url)
        If i > 0 Then
            InitSA ArrPtr(waBase64), tSA, cbElements:=2, pvData:=StrPtr(Base64Url), cElements1:=i
            For i = 0 To i - 1
                Select Case waBase64(i)
                    Case 43: waBase64(i) = 45 ' "+" -> "-"
                    Case 47: waBase64(i) = 95 ' "/" -> "_"
                    Case 61: waBase64(i) = 0 ' "=" -> ""
                End Select
            Next i
            SysReAllocString VarPtr(Base64Url), VarPtr(waBase64(0))
        End If
    End If
End Function

Public Function Box(Value As Variant) As IInspectable
Dim lLength As Long, pvData As LongPtr, ByteArray() As Byte, StringArray() As LongPtr, StringHeaders() As HSTRING_HEADER
    With PropertyValueStatics
        If IsArray(Value) Then
            With SafeArray(ArrPtrVar(Value))
                pvData = .pvData: lLength = .cElements1: If .cElements2 Then lLength = lLength * .cElements2
            End With
        End If
        Select Case VarType(Value)
            Case vbString: Set Box = .CreateString(StrRef(Value))
            Case vbByte: Set Box = .CreateUInt8(Value)
            Case vbBoolean: Set Box = .CreateBoolean(Abs(Value))
            Case vbInteger: Set Box = .CreateInt16(Value)
            Case vbLong: Set Box = .CreateInt32(Value)
            Case vbCurrency, vbLongLong: Set Box = .CreateInt64(Value)
            Case vbSingle: Set Box = .CreateSingle(Value)
            Case vbDouble: Set Box = .CreateDouble(Value)
            Case vbEmpty: Set Box = .CreateEmpty
            Case vbObject, vbDataObject: If TypeOf Value Is IInspectable Then Set Box = Value
            Case vbArray Or vbByte
                If lLength Then Set Box = .CreateUInt8Array(lLength, ByVal pvData)
            Case vbArray Or vbInteger
                If lLength Then Set Box = .CreateInt16Array(lLength, ByVal pvData)
            Case vbArray Or vbLong
                If lLength Then Set Box = .CreateInt32Array(lLength, ByVal pvData)
            Case vbArray Or vbCurrency
                If lLength Then Set Box = .CreateInt64Array(lLength, ByVal pvData)
            Case vbArray Or vbSingle
                If lLength Then Set Box = .CreateSingleArray(lLength, ByVal pvData)
            Case vbArray Or vbDouble
                If lLength Then Set Box = .CreateDoubleArray(lLength, ByVal pvData)
            Case vbArray Or vbBoolean
                If lLength Then
                    ReDim ByteArray(0 To lLength - 1): For lLength = 0 To lLength - 1: ByteArray(lLength) = Abs(Value(lLength)): Next lLength
                    Set Box = .CreateBooleanArray(lLength, ByteArray(0))
                End If
            Case vbArray Or vbString, vbArray Or vbVariant
                If lLength Then
                    ReDim StringArray(0 To lLength - 1): ReDim StringHeaders(0 To lLength - 1)
                    For lLength = 0 To lLength - 1: StringArray(lLength) = StrRef(Value(lLength), VarPtr(StringHeaders(lLength))): Next lLength
                    Set Box = .CreateStringArray(lLength, StringArray(0))
                End If
        End Select
    End With
End Function

Public Function Clamp(ByVal lValue As Long, ByVal lMin As Long, ByVal lMax As Long) As Long
    Select Case lValue
        Case Is < lMin: Clamp = lMin
        Case Is > lMax: Clamp = lMax
        Case Else: Clamp = lValue
    End Select
End Function

Private Function CreateMetaDataDispenser() As IMetaDataDispenser
Dim CLSID_CorMetaDataDispenser As UUID, CLSID_CorMetaDataDispenserRuntime As UUID, CLSID_CLRMetaHost As UUID, IID_IMetaDataDispenser As UUID, IID_ICLRMetaHost As UUID, _
    IID_ICLRRuntimeInfo As UUID, CLRMetaHost As ICLRMetaHost, CLRRuntimeInfo As ICLRRuntimeInfo, vArray As Variant
    CLSID_CorMetaDataDispenser = MakeUUID(&HE5CB7A31, &H11D27512, &H8000CE89, &HD8E592C7): ReDim vArray(0) As IMetaDataDispenser: IID_IMetaDataDispenser = GuidFromArray(vArray)
    If MetaDataGetDispenser(CLSID_CorMetaDataDispenser, IID_IMetaDataDispenser, CreateMetaDataDispenser) <> S_OK Then
        CLSID_CorMetaDataDispenserRuntime = MakeUUID(&H1EC2DE53, &H11D275CC, &HA0007597, &HCD5B4C9)
        If CoCreateInstance(CLSID_CorMetaDataDispenserRuntime, Nothing, CLSCTX_INPROC_SERVER, IID_IMetaDataDispenser, CreateMetaDataDispenser) <> S_OK Then
            CLSID_CLRMetaHost = MakeUUID(&H9280188D, &H48670E8E, &HA87F0CB3, &HDEE88438)
            ReDim vArray(0) As ICLRMetaHost: IID_ICLRMetaHost = GuidFromArray(vArray): ReDim vArray(0) As ICLRRuntimeInfo: IID_ICLRRuntimeInfo = GuidFromArray(vArray)
            If CLRCreateInstance(CLSID_CLRMetaHost, IID_ICLRMetaHost, CLRMetaHost) = S_OK Then
                If CLRMetaHost.GetRuntime(StrPtr(FrameworkVersion(4, 0, 30319)), IID_ICLRRuntimeInfo, CLRRuntimeInfo) = S_OK Then
                    CLRRuntimeInfo.GetInterface CLSID_CorMetaDataDispenser, IID_IMetaDataDispenser, CreateMetaDataDispenser
                End If
            End If
        End If
    End If
End Function

Public Function Dispose(IClosable As IClosable, Optional ByVal bRelease As Boolean = True) As Boolean
    If Not IClosable Is Nothing Then Dispose = IClosable.Close = S_OK: If bRelease Then Set IClosable = Nothing
End Function

Public Function Elapsed(Optional ByVal cyStart As Currency = INFINITE) As Currency
    If cyStart = INFINITE Then cyStart = qpcStart
    QueryPerformanceCounter Elapsed: Elapsed = (Elapsed - cyStart) * 1000 / qpcFrequency
End Function

Private Function FrameworkVersion(ParamArray VersionNumbers() As Variant) As String
    FrameworkVersion = ChrW$(118) & Join(VersionNumbers, ChrW$(46))
End Function

Public Function FromDate(dtDate As Date, Optional ByVal bUniversalTimeToLocalTime As Boolean) As Currency
Dim SysTime As SYSTEMTIME
    If VariantTimeToSystemTime(dtDate, SysTime) Then
        If bUniversalTimeToLocalTime Then SystemTimeToTzSpecificLocalTimeEx ByVal vbNullPtr, SysTime, SysTime
        SystemTimeToFileTime SysTime, FromDate
    End If
End Function

Public Function GetErrorMessage(ByVal lLastError As Long) As String
Dim lpBuffer As LongPtr
    If FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS Or FORMAT_MESSAGE_MAX_WIDTH_MASK, 0, lLastError, 0, lpBuffer, 0, 0) Then
        If lpBuffer Then SysReAllocString VarPtr(GetErrorMessage), lpBuffer: LocalFree lpBuffer
    End If
End Function

Public Function GuidFromArray(vArray As Variant) As UUID
    GuidFromArray = SafeArray(ArrPtrVar(vArray)).Guid
End Function

Public Function GuidFromHash(sStringToHash As String, Optional sAlgorithm As String = "SHA1") As UUID
Dim BufferByteAccess As IBufferByteAccess
    With HashAlgorithmProviderStatics.OpenAlgorithm(StrRef(sAlgorithm))
        Set BufferByteAccess = .HashData(CryptographicBufferStatics.CreateFromByteArray(LenB(sStringToHash), ByVal StrPtr(sStringToHash)))
        vbaCopyBytes LenB(GuidFromHash), VarPtr(GuidFromHash), BufferByteAccess.Buffer
    End With
End Function

Public Function GuidFromString(sGuid As String) As UUID
    IIDFromString StrPtr(sGuid), GuidFromString
End Function

Public Function IInspectableGetIIDs(Interface As IInspectable, SupportedIIDs() As UUID) As Long
Dim ParamTypes(0 To 1) As Integer, ParamValues(0 To 1) As LongPtr, vParams As Variant, lParamCount As Long, hRes As Variant, IIDsPtr As LongPtr
Const IInspectable_GetIIDs As Long = 3 * PTR_SIZE
    If Not Interface Is Nothing Then
        vParams = Array(VarPtr(IInspectableGetIIDs), VarPtr(IIDsPtr))
        For lParamCount = 0 To UBound(vParams): ParamTypes(lParamCount) = VarType(vParams(lParamCount)): ParamValues(lParamCount) = VarPtr(vParams(lParamCount)): Next lParamCount
        DispCallFunc ObjPtr(Interface), IInspectable_GetIIDs, CC_STDCALL, vbLong, lParamCount, ParamTypes(0), ParamValues(0), hRes
        If IInspectableGetIIDs > 0 Then
            ReDim SupportedIIDs(0 To IInspectableGetIIDs - 1)
            vbaCopyBytes LenB(SupportedIIDs(0)) * IInspectableGetIIDs, ByVal VarPtr(SupportedIIDs(0)), ByVal IIDsPtr: CoTaskMemFree IIDsPtr
        End If
    End If
End Function

Public Function IInspectableGetRuntimeClassName(Interface As IInspectable) As String
Dim vParam As Variant, hStringPtr As HSTRING
Const IInspectable_GetRuntimeClassName As Long = 4 * PTR_SIZE
    If Not Interface Is Nothing Then
        vParam = VarPtr(hStringPtr)
        DispCallFunc ObjPtr(Interface), IInspectable_GetRuntimeClassName, CC_STDCALL, vbLong, 1, VarType(vParam), VarPtr(vParam), vParam
        If vParam = S_OK Then IInspectableGetRuntimeClassName = GetStr(hStringPtr)
    End If
End Function

Public Function InitializeWithWindow(ByVal IInitializeWithWindow As IInitializeWithWindow, ByVal hWnd As LongPtr) As Boolean
    InitializeWithWindow = IInitializeWithWindow.Initialize(hWnd) = S_OK
End Function

Private Sub InitNamespaces()
Dim TempObject As IInspectable, vArray As Variant, StringHeader As HSTRING_HEADER
    If RoActivateInstance(StrRef("Windows.Foundation.Collections.StringMap", VarPtr(StringHeader)), TempObject) = S_OK Then
        ReDim vArray(0) As IMetaDataImport: m_IID_IMetaDataImport = GuidFromArray(vArray): ReDim vArray(0) As IActivationFactory: m_IID_IActivationFactory = GuidFromArray(vArray)
        Set m_RuntimeClasses = TempObject: ResolveNamespaces CreateMetaDataDispenser
    End If
End Sub

Private Function IsClassActivatable(ByVal ptkTypeDef As mdTypeDef, MetaDataImport As IMetaDataImport) As Boolean
Dim hEnum As HCORENUM, CustomAttributesBuffer(0 To 255) As mdCustomAttribute, sAttributeNameBuffer As String, ptkMethodRef As mdMemberRef, i As Long
    With MetaDataImport
        If .EnumCustomAttributes(hEnum, ptkTypeDef, 0&, CustomAttributesBuffer(0), UBound(CustomAttributesBuffer) + 1, i) = S_OK Then
            sAttributeNameBuffer = String$(256, vbNullChar)
            For i = 0 To i - 1
                If .GetCustomAttributeProps(CustomAttributesBuffer(i), 0&, ptkMethodRef, ByVal vbNullPtr, 0&) = S_OK Then
                    If .GetMemberRefProps(ptkMethodRef, ptkTypeDef, vbNullPtr, 0&, 0&, ByVal vbNullPtr, 0&) = S_OK Then
                        If .GetTypeRefProps(ptkTypeDef, 0&, StrPtr(sAttributeNameBuffer), Len(sAttributeNameBuffer), 0&) = S_OK Then
                            Select Case True
                                Case InStr(sAttributeNameBuffer, ".StaticAttribute") > 0: IsClassActivatable = True: Exit For
                                Case InStr(sAttributeNameBuffer, ".ActivatableAttribute") > 0: IsClassActivatable = True: Exit For
                                Case InStr(sAttributeNameBuffer, ".ComposableAttribute") > 0: IsClassActivatable = True: Exit For
                            End Select
                        End If
                    End If
                End If
            Next i
            .CloseEnum hEnum
        End If
    End With
End Function

Public Function NewObject(ByVal sRuntimeClass As String, Optional ByVal bDefaultConstructor As Boolean = True) As IInspectable
Dim StringHeader As HSTRING_HEADER, ActivationFactory As IActivationFactory
    If m_RuntimeClasses Is Nothing Then InitNamespaces
    If m_RuntimeClasses.HasKey(StrRef(sRuntimeClass, VarPtr(StringHeader))) = APITRUE Then
        sRuntimeClass = GetStr(m_RuntimeClasses.Lookup(StrRef(sRuntimeClass, VarPtr(StringHeader))))
        If RoGetActivationFactory(StrRef(sRuntimeClass, VarPtr(StringHeader)), m_IID_IActivationFactory, ActivationFactory) = S_OK Then
            If bDefaultConstructor Then
                If ActivationFactory.ActivateInstance(NewObject) <> S_OK Then
                    #If bInIDE Then
                        Debug.Print sRuntimeClass, "Class does not implement a parameterless constructor!": Debug.Assert False
                    #Else
                        MsgBox sRuntimeClass & vbNewLine & "Class does not implement a parameterless constructor!", vbOKOnly + vbExclamation, App.Title
                    #End If
                End If
            Else
                Set NewObject = ActivationFactory
            End If
        End If
    Else
        #If bInIDE Then
            Debug.Print sRuntimeClass, "Invalid class name or class does not implement an activation factory or statics interface!": Debug.Assert False
        #Else
            MsgBox sRuntimeClass & " Invalid class name or class does not implement an activation factory or statics interface!", vbOKOnly + vbExclamation, App.Title
        #End If
    End If
End Function

Private Sub ParseNamespace(MetaDataImport As IMetaDataImport)
Dim hEnum As HCORENUM, TypeDefsBuffer(0 To 255) As mdTypeDef, i As Long, sTypeNameBuffer As String, lBufferLength As Long, pdwTypeDefFlags As Long, ptkExtends As mdToken, _
    sTypeName As String, sBaseTypeBuffer As String, StringHeaderKey As HSTRING_HEADER, StringHeaderValue As HSTRING_HEADER
    sTypeNameBuffer = String$(256, vbNullChar): sBaseTypeBuffer = String$(256, vbNullChar)
    With MetaDataImport
        Do
            .EnumTypeDefs hEnum, TypeDefsBuffer(0), UBound(TypeDefsBuffer) + 1, i
            For i = 0 To i - 1
                If .GetTypeDefProps(TypeDefsBuffer(i), StrPtr(sTypeNameBuffer), Len(sTypeNameBuffer), lBufferLength, pdwTypeDefFlags, ptkExtends) = S_OK Then
                    If (pdwTypeDefFlags And tdClassSemanticsMask) = tdClass Then ' Exclude Interfaces
                        If .GetTypeRefProps(ptkExtends, 0&, StrPtr(sBaseTypeBuffer), Len(sBaseTypeBuffer), 0&) = S_OK Then
                            If InStr(sBaseTypeBuffer, "System.Enum") = 0 And InStr(sBaseTypeBuffer, "System.ValueType") = 0 Then ' Exclude ValueTypes and Enums
                                If IsClassActivatable(TypeDefsBuffer(i), MetaDataImport) Then
                                    sTypeName = Left$(sTypeNameBuffer, lBufferLength - 1)
                                    m_RuntimeClasses.Insert StrRef(Mid$(sTypeName, InStrRev(sTypeName, ChrW$(46)) + 1), VarPtr(StringHeaderKey)), StrRef(sTypeName, VarPtr(StringHeaderValue))
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        Loop While i > 0
        .CloseEnum hEnum
    End With
End Sub

Public Function QueryObj(ByVal lpInterface As LongPtr, ParamArray ParamsArray() As Variant) As Variant
Dim ParamTypes(0 To 1) As Integer, ParamValues(0 To 1) As LongPtr, lParamCount As Long
    If lpInterface Then
        For lParamCount = 0 To UBound(ParamsArray): ParamTypes(lParamCount) = VarType(ParamsArray(lParamCount)): ParamValues(lParamCount) = VarPtr(ParamsArray(lParamCount)): Next lParamCount
        DispCallFunc lpInterface, vbNullPtr, CC_STDCALL, vbLong, lParamCount, ParamTypes(0), ParamValues(0), QueryObj
    Else
        QueryObj = E_POINTER
    End If
End Function

Public Property Get RegValStr(Optional ByVal hPredefinedKey As REGISTRY_PREDEFINED_KEYS = HKEY_CURRENT_USER, Optional sSubKey As String, Optional sValueName As String) As Variant
Dim cbData As Long, pValue As LongPtr
    RegValStr = vbNullString
    If RegGetValueW(hPredefinedKey, StrPtr(sSubKey), StrPtr(sValueName), RRF_RT_REG_SZ, REG_NONE, vbNullPtr, cbData) = ERROR_SUCCESS Then
        pValue = CoTaskMemAlloc(cbData)
        If RegGetValueW(hPredefinedKey, StrPtr(sSubKey), StrPtr(sValueName), RRF_RT_REG_SZ, REG_NONE, pValue, cbData) = ERROR_SUCCESS Then SysReAllocString VarPtr(RegValStr) + 8, pValue
        CoTaskMemFree pValue
    End If
End Property

Public Property Let RegValStr(Optional ByVal hPredefinedKey As REGISTRY_PREDEFINED_KEYS = HKEY_CURRENT_USER, Optional sSubKey As String, Optional sValueName As String, sValueData As Variant)
    RegSetKeyValueW hPredefinedKey, StrPtr(sSubKey), StrPtr(sValueName), REG_SZ, StrPtr(sValueData), LenB(sValueData) + 2
End Property

Private Sub ResolveNamespaces(MetaDataDispenser As IMetaDataDispenser, Optional sNameSpace As String = "Windows")
Dim lFileCount As Long, pFiles As LongPtr, lSubNamespaceCount As Long, pSubNamespaces As LongPtr, i As Long, hName As HSTRING, MetaDataImport As IMetaDataImport, StringHeader As HSTRING_HEADER
    If RoResolveNamespace(StrRef(sNameSpace, VarPtr(StringHeader)), vbNullPtr, 0, vbNullPtr, lFileCount, pFiles, lSubNamespaceCount, pSubNamespaces) = S_OK Then
        For i = 0 To lFileCount - 1
            GetMemPtr ByVal pFiles + i * PTR_SIZE, hName
            If MetaDataDispenser.OpenScope(StrPtr(GetStr(hName)), CorOpenFlags_ofRead, m_IID_IMetaDataImport, MetaDataImport) = S_OK Then ParseNamespace MetaDataImport: Set MetaDataImport = Nothing
        Next i
        CoTaskMemFree pFiles
        For i = 0 To lSubNamespaceCount - 1
            GetMemPtr ByVal pSubNamespaces + i * PTR_SIZE, hName: ResolveNamespaces MetaDataDispenser, sNameSpace & ChrW$(46) & GetStr(hName)
        Next i
        CoTaskMemFree pSubNamespaces
    End If
End Sub

Public Sub StartTiming(Optional cyStart As Currency = INFINITE)
    If qpcFrequency = 0 Then QueryPerformanceFrequency qpcFrequency
    If cyStart = INFINITE Then QueryPerformanceCounter qpcStart Else QueryPerformanceCounter cyStart
End Sub

Public Function StringFromGuid(ByVal rIID As LongPtr) As String
    If rIID Then If StringFromIID(ByVal rIID, rIID) = S_OK Then SysReAllocString VarPtr(StringFromGuid), rIID: CoTaskMemFree rIID
End Function

Public Function ToDate(cyDate As Currency, Optional ByVal bUniversalTimeToLocalTime As Boolean) As Date
Dim SysTime As SYSTEMTIME
    If FileTimeToSystemTime(cyDate, SysTime) Then
        If bUniversalTimeToLocalTime Then SystemTimeToTzSpecificLocalTimeEx ByVal vbNullPtr, SysTime, SysTime
        SystemTimeToVariantTime SysTime, ToDate
    End If
End Function

Public Function ToString(ByVal IStringable As IStringable) As String: ToString = GetStr(IStringable.ToString): End Function

Public Function Unbox(ByVal Value As IInspectable) As Variant
Dim PropertyValue As IPropertyValue, lLength As Long, ValuePtr As LongPtr, ValueArray As Variant, PtrArray() As LongPtr, tSA As SAFE_ARRAY
    If Value Is Nothing Then
        Exit Function
    ElseIf TypeOf Value Is IPropertyValue Then
        Set PropertyValue = Value
        With PropertyValue
            Select Case .Type
                Case PropertyType_String: Unbox = GetStr(.GetString)
                Case PropertyType_UInt8: Unbox = .GetUInt8
                Case PropertyType_Int16: Unbox = .GetInt16
                Case PropertyType_Int32: Unbox = .GetInt32
                Case PropertyType_Int64: Unbox = .GetInt64
                Case PropertyType_Single: Unbox = .GetSingle
                Case PropertyType_Double: Unbox = .GetDouble
                Case PropertyType_Boolean: Unbox = CBool(.GetBoolean)
                Case PropertyType_DateTime: Unbox = .GetDateTime
                Case PropertyType_TimeSpan: Unbox = .GetTimeSpan
                Case PropertyType_Char16: Unbox = .GetChar16
                Case PropertyType_UInt16: Unbox = .GetUInt16
                Case PropertyType_UInt32: Unbox = .GetUInt32
                Case PropertyType_UInt64: Unbox = .GetUInt64
                Case PropertyType_StringArray
                    .GetStringArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As String
                    InitSA ArrPtr(PtrArray), tSA, cbElements:=LenB(ValuePtr), pvData:=ValuePtr, cElements1:=lLength
                    For lLength = 0 To lLength - 1: ValueArray(lLength) = GetStr(PtrArray(lLength)): Next lLength
                    CoTaskMemFree ValuePtr
                Case PropertyType_UInt8Array To PropertyType_TimeSpanArray
                    Select Case .Type
                        Case PropertyType_UInt8Array: .GetUInt8Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Byte
                        Case PropertyType_Int16Array: .GetInt16Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Integer
                        Case PropertyType_Int32Array: .GetInt32Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Long
                        Case PropertyType_Int64Array: .GetInt64Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Currency
                        Case PropertyType_SingleArray: .GetSingleArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Single
                        Case PropertyType_DoubleArray: .GetDoubleArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Double
                        Case PropertyType_BooleanArray: .GetBooleanArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Byte
                        Case PropertyType_DateTimeArray: .GetDateTimeArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Currency
                        Case PropertyType_TimeSpanArray: .GetTimeSpanArray lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Currency
                        Case PropertyType_Char16Array: .GetChar16Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Integer
                        Case PropertyType_UInt16Array: .GetUInt16Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Integer
                        Case PropertyType_UInt32Array: .GetUInt32Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Long
                        Case PropertyType_UInt64Array: .GetUInt64Array lLength, ValuePtr: ReDim ValueArray(0 To lLength - 1) As Currency
                    End Select
                    If IsArray(ValueArray) Then With SafeArray(ArrPtrVar(ValueArray)): vbaCopyBytes .cElements1 * .cbElements, .pvData, ValuePtr: End With
                    CoTaskMemFree ValuePtr
            End Select
        End With
    Else
        Set Unbox = Value
    End If
    If Not IsEmpty(Unbox) Then Exit Function
    Unbox = ValueArray
End Function

' ----------------------------------------------------------------------------------------------------------------------------

#If (VBA7 <> False Or VBA6 <> False) And TWINBASIC = False Then

Public Function vbaCopyBytes(ByVal Length As Long, ByVal Destination As LongPtr, ByVal Source As LongPtr) As LongPtr
    CopyMemory ByVal Destination, ByVal Source, Length: vbaCopyBytes = Destination
End Function

Public Function vbaCopyBytesZero(ByVal Length As Long, ByVal Destination As LongPtr, ByVal Source As LongPtr) As LongPtr
    CopyMemory ByVal Destination, ByVal Source, Length: ZeroMemory ByVal Source, Length: vbaCopyBytesZero = Destination
End Function

Public Sub GetMem1(ByVal Ptr As LongPtr, RetVal As Byte)
    CopyMemory RetVal, ByVal Ptr, LenB(RetVal)
End Sub

Public Sub GetMem2(ByVal Ptr As LongPtr, RetVal As Integer)
    CopyMemory RetVal, ByVal Ptr, LenB(RetVal)
End Sub

Public Sub GetMem4(ByVal Ptr As LongPtr, RetVal As Long)
    CopyMemory RetVal, ByVal Ptr, LenB(RetVal)
End Sub

Public Sub GetMem8(ByVal Ptr As LongPtr, RetVal As Currency)
    CopyMemory RetVal, ByVal Ptr, LenB(RetVal)
End Sub

Public Sub PutMem1(ByVal Ptr As LongPtr, ByVal NewVal As Byte)
    CopyMemory ByVal Ptr, NewVal, LenB(NewVal)
End Sub

Public Sub PutMem2(ByVal Ptr As LongPtr, ByVal NewVal As Integer)
    CopyMemory ByVal Ptr, NewVal, LenB(NewVal)
End Sub

Public Sub PutMem4(ByVal Ptr As LongPtr, ByVal NewVal As Long)
    CopyMemory ByVal Ptr, NewVal, LenB(NewVal)
End Sub

Public Sub PutMem8(ByVal Ptr As LongPtr, ByVal NewVal As Currency)
    CopyMemory ByVal Ptr, NewVal, LenB(NewVal)
End Sub

Public Sub GetMemPtr(ByVal Ptr As LongPtr, RetVal As LongPtr)
    CopyMemory RetVal, ByVal Ptr, LenB(RetVal)
End Sub

Public Sub PutMemPtr(ByVal Ptr As LongPtr, ByVal NewVal As LongPtr)
    CopyMemory ByVal Ptr, NewVal, LenB(NewVal)
End Sub

#End If
