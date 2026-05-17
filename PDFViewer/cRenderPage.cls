VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRenderPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncActionCompletedHandler
Implements IAsyncOperationCompletedHandlerIBitmapDecoder
Implements IAsyncOperationCompletedHandlerIBitmapEncoder
Implements Unrestricted.IInspectable

Private m_PdfPage As IPdfPage, m_RandomAccessStream As IRandomAccessStream, m_RotatedRandomAccessStream As IRandomAccessStream, m_BitmapSource As IBitmapSource, m_AsyncInfo As IAsyncInfo, _
        m_PdfPageRotation As PdfPageRotation, m_bIsRenderingFinished As Boolean

Friend Property Get BitmapSource() As IBitmapSource
    Set BitmapSource = m_BitmapSource
End Property

Friend Sub CancelRendering()
    If IsBusy Then m_AsyncInfo.Cancel
End Sub

Friend Property Get IsBusy() As Boolean
    If Not m_AsyncInfo Is Nothing Then IsBusy = m_AsyncInfo.Status = AsyncStatus_Started
End Property

Friend Property Get IsRenderingFinished() As Boolean
    IsRenderingFinished = m_bIsRenderingFinished
End Property

Friend Function StartRendering(PdfPage As IPdfPage, PdfPageRenderOptions As IPdfPageRenderOptions, PdfPageRotation As PdfPageRotation) As cRenderPage
Dim AsyncAction As IAsyncAction
    Set m_PdfPage = PdfPage: m_PdfPageRotation = PdfPageRotation: Set StartRendering = Me
    With m_PdfPage.Size ' Render at double the original size to improve quality while zooming in/out:
        PdfPageRenderOptions.DestinationWidth = .Width * 2: PdfPageRenderOptions.DestinationHeight = .Height * 2
    End With
    Set AsyncAction = m_PdfPage.RenderWithOptionsToStreamAsync(m_RandomAccessStream, PdfPageRenderOptions)
    Set m_AsyncInfo = AsyncAction: Set AsyncAction.Completed = Me
End Function

Private Sub IAsyncActionCompletedHandler_Invoke(ByVal AsyncAction As IAsyncAction, ByVal AsyncStatus As AsyncStatus) ' Handles two AsyncActions: PdfPage.RenderWithOptionsToStreamAsync and BitmapEncoder.FlushAsync
Dim AsyncOperationBitmapDecoder As IAsyncOperationIBitmapDecoder
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            If m_PdfPageRotation = PdfPageRotation_Normal Then ' Now the stream contains the finished rendering of the PdfPage (normal or rotated) so we can use it as the source for the Image control
                If m_RotatedRandomAccessStream Is Nothing Then m_BitmapSource.SetSource m_RandomAccessStream Else m_BitmapSource.SetSource m_RotatedRandomAccessStream
                m_bIsRenderingFinished = True: Class_Terminate ' Dispose of the PdfPage and allocated streams early since they are no longer needed
            Else ' At this point m_RandomAccessStream contains the finished rendering of the PdfPage but now we need to apply a rotation transform (takes a few async steps):
                Set AsyncOperationBitmapDecoder = BitmapDecoderStatics.CreateAsync(m_RandomAccessStream) ' First we create a BitmapDecoder from m_RandomAccessStream
                Set m_AsyncInfo = AsyncOperationBitmapDecoder: Set AsyncOperationBitmapDecoder.Completed = Me ' This can be canceled at any time so update m_AsyncInfo with the new async operation
            End If
    End Select
End Sub

Private Sub IAsyncOperationCompletedHandlerIBitmapDecoder_Invoke(ByVal AsyncOperation As IAsyncOperationIBitmapDecoder, ByVal AsyncStatus As AsyncStatus)
Dim AsyncOperationBitmapEncoder As IAsyncOperationIBitmapEncoder
    Select Case AsyncStatus
        Case AsyncStatus_Completed ' Now the BitmapDecoder is ready so we need to create a BitmapEncoder to apply the rotation transform:
            Set m_RotatedRandomAccessStream = NewObject("InMemoryRandomAccessStream")
            Set AsyncOperationBitmapEncoder = BitmapEncoderStatics.CreateForTranscodingAsync(m_RotatedRandomAccessStream, AsyncOperation.GetResults)
            Set m_AsyncInfo = AsyncOperationBitmapEncoder: Set AsyncOperationBitmapEncoder.Completed = Me ' This can be canceled at any time so update m_AsyncInfo with the new async operation
    End Select
End Sub

Private Sub IAsyncOperationCompletedHandlerIBitmapEncoder_Invoke(ByVal AsyncOperation As IAsyncOperationIBitmapEncoder, ByVal AsyncStatus As AsyncStatus)
Dim AsyncAction As IAsyncAction
    Select Case AsyncStatus
        Case AsyncStatus_Completed
            With AsyncOperation.GetResults ' The BitmapEncoder is now ready to apply the rotation transform:
                .BitmapTransform.Rotation = m_PdfPageRotation: m_PdfPageRotation = PdfPageRotation_Normal ' Restore the rotation value to normal for the IAsyncActionCompletedHandler above
                Set AsyncAction = .FlushAsync: Set m_AsyncInfo = AsyncAction: Set AsyncAction.Completed = Me ' FlushAsync is the second AsyncAction handled by IAsyncActionCompletedHandler above
            End With
    End Select
End Sub

Private Sub Class_Initialize()
    Set m_RandomAccessStream = NewObject("InMemoryRandomAccessStream"): Set m_BitmapSource = NewObject("BitmapSource")
End Sub

Private Sub Class_Terminate()
    Dispose m_RandomAccessStream: Dispose m_RotatedRandomAccessStream: Dispose m_PdfPage ' These objects implement IClosable so we need to free their allocated resources in the absence of an automated garbage collection mechanism
End Sub

' Boilerplate IInspectable implementation (doesn't get called but our object needs to respond with S_OK to IInspectable queries)

Private Sub IInspectable_GetIIDs(iidCount As Long, iidPtr As LongPtr)
    #If bInIDE Then
        Debug.Print "IInspectable_GetIIDs"
    #End If
End Sub

Private Function IInspectable_GetRuntimeClassName() As HSTRING
    #If bInIDE Then
        Debug.Print "IInspectable_GetRuntimeClassName"
    #End If
End Function

Private Function IInspectable_GetTrustLevel() As TrustLevel
    #If bInIDE Then
        Debug.Print "IInspectable_GetTrustLevel"
    #End If
End Function
