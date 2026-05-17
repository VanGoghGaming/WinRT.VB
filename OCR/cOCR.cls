VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_OcrEngine As IOcrEngine, m_BitmapTransform As IBitmapTransform, m_SoftwareBitmapSource As ISoftwareBitmapSource, m_ScaleFactor As Single, m_ToGrayscale As Boolean, m_Await As cAwait

Public Event RecognitionResults(OcrResult As IOcrResult, SoftwareBitmapSource As ISoftwareBitmapSource, ByVal PixelWidth As Long, ByVal PixelHeight As Long)

Friend Sub GetBitmapFromStream(RandomAccessStreamReference As IRandomAccessStreamReference)
Dim RandomAccessStream As IRandomAccessStream, SoftwareBitmap As ISoftwareBitmap, BitmapFrame As IBitmapFrame, BitmapFrameWithSoftwareBitmap As IBitmapFrameWithSoftwareBitmap
    With m_Await
        If .Await(RandomAccessStreamReference.OpenReadAsync) = AsyncStatus_Completed Then
            Set RandomAccessStream = .GetResults
            If .Await(BitmapDecoderStatics.CreateAsync(RandomAccessStream)) = AsyncStatus_Completed Then
                Set BitmapFrame = .GetResults: Set BitmapFrameWithSoftwareBitmap = BitmapFrame
                m_BitmapTransform.ScaledWidth = BitmapFrame.PixelWidth * m_ScaleFactor: m_BitmapTransform.ScaledHeight = BitmapFrame.PixelHeight * m_ScaleFactor
                If .Await(BitmapFrameWithSoftwareBitmap.GetSoftwareBitmapTransformedAsync(BitmapPixelFormat_Bgra8, BitmapAlphaMode_Premultiplied, m_BitmapTransform, ExifOrientationMode_IgnoreExifOrientation, ColorManagementMode_DoNotColorManage)) = AsyncStatus_Completed Then
                    Set SoftwareBitmap = .GetResults
                    If m_ToGrayscale Then ConvertToGrayscale SoftwareBitmap.LockBuffer(BitmapBufferAccessMode_ReadWrite)
                    StartRecognition SoftwareBitmap
                End If
            End If
        End If
    End With
    Dispose SoftwareBitmap: Dispose RandomAccessStream
End Sub

Friend Property Get ScaleFactor() As Single
    ScaleFactor = m_ScaleFactor
End Property

Friend Property Let ScaleFactor(ByVal sngScaleFactor As Single)
    m_ScaleFactor = sngScaleFactor
End Property

Friend Property Get ToGrayscale() As Boolean
    ToGrayscale = m_ToGrayscale
End Property

Friend Property Let ToGrayscale(ByVal bToGrayscale As Boolean)
    m_ToGrayscale = bToGrayscale
End Property

Private Sub ConvertToGrayscale(BitmapBuffer As IBitmapBuffer)
Dim MemoryBuffer As IMemoryBuffer, MemoryBufferByteAccess As IMemoryBufferByteAccess, lCapacity As Long, BufferPtr As LongPtr, Pixels() As RGBQUAD, i As Long, j As Long, lGray As Long
    Set MemoryBuffer = BitmapBuffer: Set MemoryBufferByteAccess = MemoryBuffer.CreateReference
    If MemoryBufferByteAccess.GetBuffer(BufferPtr, lCapacity) = S_OK Then
        With BitmapBuffer.GetPlaneDescription(0)
            With NewSA(ArrPtr(Pixels), 4, BufferPtr, .Height, .Width)
                For j = LBound(Pixels, 2) To UBound(Pixels, 2)
                    For i = LBound(Pixels, 1) To UBound(Pixels, 1)
                        With Pixels(i, j)
                            lGray = (.rgbRed * 299& + .rgbGreen * 587& + .rgbBlue * 114&) \ 1000
                            .rgbBlue = lGray: .rgbGreen = lGray: .rgbRed = lGray
                        End With
                    Next i
                Next j
            End With
        End With
    End If
    Dispose MemoryBufferByteAccess: Dispose BitmapBuffer
End Sub

Private Sub StartRecognition(SoftwareBitmap As ISoftwareBitmap)
    If Not m_OcrEngine Is Nothing Then
        With m_Await
            If .Await(m_SoftwareBitmapSource.SetBitmapAsync(SoftwareBitmap)) = AsyncStatus_Completed Then
                If .Await(m_OcrEngine.RecognizeAsync(SoftwareBitmap)) = AsyncStatus_Completed Then
                    RaiseEvent RecognitionResults(.GetResults, m_SoftwareBitmapSource, SoftwareBitmap.PixelWidth, SoftwareBitmap.PixelHeight)
                End If
            End If
        End With
    End If
End Sub

Private Sub Class_Initialize()
    Set m_Await = New cAwait: m_ScaleFactor = 1.5: m_ToGrayscale = True
    Set m_SoftwareBitmapSource = NewObject("SoftwareBitmapSource"): Set m_BitmapTransform = NewObject("BitmapTransform"): m_BitmapTransform.InterpolationMode = BitmapInterpolationMode_Fant
    Set m_OcrEngine = OcrEngineStatics.TryCreateFromUserProfileLanguages
    #If bInIDE Then
        If Not m_OcrEngine Is Nothing Then
            With m_OcrEngine.RecognizerLanguage
                Debug.Print "OCR Engine Ready, language:", GetStr(.DisplayName), GetStr(.LanguageTag)
            End With
        Else
            Debug.Print "Failed to create OCR Engine from user profile languages!"
        End If
    #End If
End Sub

Private Sub Class_Terminate()
    Dispose m_SoftwareBitmapSource
End Sub
