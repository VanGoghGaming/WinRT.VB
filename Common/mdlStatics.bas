Attribute VB_Name = "mdlStatics"
Option Explicit

Public Function BitmapDecoderStatics() As IBitmapDecoderStatics
Static IBitmapDecoderStatics As IBitmapDecoderStatics
    If IBitmapDecoderStatics Is Nothing Then Set IBitmapDecoderStatics = NewObject("BitmapDecoder", False)
    Set BitmapDecoderStatics = IBitmapDecoderStatics
End Function

Public Function BitmapEncoderStatics() As IBitmapEncoderStatics
Static IBitmapEncoderStatics As IBitmapEncoderStatics
    If IBitmapEncoderStatics Is Nothing Then Set IBitmapEncoderStatics = NewObject("BitmapEncoder", False)
    Set BitmapEncoderStatics = IBitmapEncoderStatics
End Function

Public Function BufferFactory() As IBufferFactory
Static IBufferFactory As IBufferFactory
    If IBufferFactory Is Nothing Then Set IBufferFactory = NewObject("Buffer", False)
    Set BufferFactory = IBufferFactory
End Function

Public Function BufferStatics() As IBufferStatics
Static IBufferStatics As IBufferStatics
    If IBufferStatics Is Nothing Then Set IBufferStatics = NewObject("Buffer", False)
    Set BufferStatics = IBufferStatics
End Function

Public Function CanvasStatics() As ICanvasStatics
Static ICanvasStatics As ICanvasStatics
    If ICanvasStatics Is Nothing Then Set ICanvasStatics = NewObject("Canvas", False)
    Set CanvasStatics = ICanvasStatics
End Function

Public Function ClipboardStatics() As IClipboardStatics
Static IClipboardStatics As IClipboardStatics
    If IClipboardStatics Is Nothing Then Set IClipboardStatics = NewObject("Clipboard", False)
    Set ClipboardStatics = IClipboardStatics
End Function

Public Function ColorHelperStatics() As IColorHelperStatics
Static IColorHelperStatics As IColorHelperStatics
    If IColorHelperStatics Is Nothing Then Set IColorHelperStatics = NewObject("ColorHelper", False)
    Set ColorHelperStatics = IColorHelperStatics
End Function

Public Function ColorsStatics() As IColorsStatics
Static IColorsStatics As IColorsStatics
    If IColorsStatics Is Nothing Then Set IColorsStatics = NewObject("Colors", False)
    Set ColorsStatics = IColorsStatics
End Function

Public Function CoreCursorFactory() As ICoreCursorFactory
Static ICoreCursorFactory As ICoreCursorFactory
    If ICoreCursorFactory Is Nothing Then Set ICoreCursorFactory = NewObject("CoreCursor", False)
    Set CoreCursorFactory = ICoreCursorFactory
End Function

Public Function CornerRadiusHelperStatics() As ICornerRadiusHelperStatics
Static ICornerRadiusHelperStatics As ICornerRadiusHelperStatics
    If ICornerRadiusHelperStatics Is Nothing Then Set ICornerRadiusHelperStatics = NewObject("CornerRadiusHelper", False)
    Set CornerRadiusHelperStatics = ICornerRadiusHelperStatics
End Function

Public Function CredentialPickerStatics() As ICredentialPickerStatics
Static ICredentialPickerStatics As ICredentialPickerStatics
    If ICredentialPickerStatics Is Nothing Then Set ICredentialPickerStatics = NewObject("CredentialPicker", False)
    Set CredentialPickerStatics = ICredentialPickerStatics
End Function

Public Function CryptographicBufferStatics() As ICryptographicBufferStatics
Static ICryptographicBufferStatics As ICryptographicBufferStatics
    If ICryptographicBufferStatics Is Nothing Then Set ICryptographicBufferStatics = NewObject("CryptographicBuffer", False)
    Set CryptographicBufferStatics = ICryptographicBufferStatics
End Function

Public Function DataReaderFactory() As IDataReaderFactory
Static IDataReaderFactory As IDataReaderFactory
    If IDataReaderFactory Is Nothing Then Set IDataReaderFactory = NewObject("DataReader", False)
    Set DataReaderFactory = IDataReaderFactory
End Function

Public Function DataTransferManagerStatics() As IDataTransferManagerStatics
Static IDataTransferManagerStatics As IDataTransferManagerStatics
    If IDataTransferManagerStatics Is Nothing Then Set IDataTransferManagerStatics = NewObject("DataTransferManager", False)
    Set DataTransferManagerStatics = IDataTransferManagerStatics
End Function

Public Function DataWriterFactory() As IDataWriterFactory
Static IDataWriterFactory As IDataWriterFactory
    If IDataWriterFactory Is Nothing Then Set IDataWriterFactory = NewObject("DataWriter", False)
    Set DataWriterFactory = IDataWriterFactory
End Function

Public Function DeviceInformationStatics() As IDeviceInformationStatics
Static IDeviceInformationStatics As IDeviceInformationStatics
    If IDeviceInformationStatics Is Nothing Then Set IDeviceInformationStatics = NewObject("DeviceInformation", False)
    Set DeviceInformationStatics = IDeviceInformationStatics
End Function

Public Function Direct3D11CaptureFramePoolStatics() As IDirect3D11CaptureFramePoolStatics
Static IDirect3D11CaptureFramePoolStatics As IDirect3D11CaptureFramePoolStatics
    If IDirect3D11CaptureFramePoolStatics Is Nothing Then Set IDirect3D11CaptureFramePoolStatics = NewObject("Direct3D11CaptureFramePool", False)
    Set Direct3D11CaptureFramePoolStatics = IDirect3D11CaptureFramePoolStatics
End Function

Public Function DisplayInformationStatics() As IDisplayInformationStatics
Static IDisplayInformationStatics As IDisplayInformationStatics
    If IDisplayInformationStatics Is Nothing Then Set IDisplayInformationStatics = NewObject("DisplayInformation", False)
    Set DisplayInformationStatics = IDisplayInformationStatics
End Function

Public Function DurationHelperStatics() As IDurationHelperStatics
Static IDurationHelperStatics As IDurationHelperStatics
    If IDurationHelperStatics Is Nothing Then Set IDurationHelperStatics = NewObject("DurationHelper", False)
    Set DurationHelperStatics = IDurationHelperStatics
End Function

Public Function ElementCompositionPreviewStatics() As IElementCompositionPreviewStatics
Static IElementCompositionPreviewStatics As IElementCompositionPreviewStatics
    If IElementCompositionPreviewStatics Is Nothing Then Set IElementCompositionPreviewStatics = NewObject("ElementCompositionPreview", False)
    Set ElementCompositionPreviewStatics = IElementCompositionPreviewStatics
End Function

Public Function FaceDetectorStatics() As IFaceDetectorStatics
Static IFaceDetectorStatics As IFaceDetectorStatics
    If IFaceDetectorStatics Is Nothing Then Set IFaceDetectorStatics = NewObject("FaceDetector", False)
    Set FaceDetectorStatics = IFaceDetectorStatics
End Function

Public Function FileIOStatics() As IFileIOStatics
Static IFileIOStatics As IFileIOStatics
    If IFileIOStatics Is Nothing Then Set IFileIOStatics = NewObject("FileIO", False)
    Set FileIOStatics = IFileIOStatics
End Function

Public Function FileRandomAccessStreamStatics() As IFileRandomAccessStreamStatics
Static IFileRandomAccessStreamStatics As IFileRandomAccessStreamStatics
    If IFileRandomAccessStreamStatics Is Nothing Then Set IFileRandomAccessStreamStatics = NewObject("FileRandomAccessStream", False)
    Set FileRandomAccessStreamStatics = IFileRandomAccessStreamStatics
End Function

Public Function FontWeightsStatics() As IFontWeightsStatics
Static IFontWeightsStatics As IFontWeightsStatics
    If IFontWeightsStatics Is Nothing Then Set IFontWeightsStatics = NewObject("FontWeights", False)
    Set FontWeightsStatics = IFontWeightsStatics
End Function

Public Function FrameworkElementStatics() As IFrameworkElementStatics
Static IFrameworkElementStatics As IFrameworkElementStatics
    If IFrameworkElementStatics Is Nothing Then Set IFrameworkElementStatics = NewObject("FrameworkElement", False)
    Set FrameworkElementStatics = IFrameworkElementStatics
End Function

Public Function GeneratorPositionHelperStatics() As IGeneratorPositionHelperStatics
Static IGeneratorPositionHelperStatics As IGeneratorPositionHelperStatics
    If IGeneratorPositionHelperStatics Is Nothing Then Set IGeneratorPositionHelperStatics = NewObject("GeneratorPositionHelper", False)
    Set GeneratorPositionHelperStatics = IGeneratorPositionHelperStatics
End Function

Public Function GraphicsCaptureItemStatics() As IGraphicsCaptureItemStatics
Static IGraphicsCaptureItemStatics As IGraphicsCaptureItemStatics
    If IGraphicsCaptureItemStatics Is Nothing Then Set IGraphicsCaptureItemStatics = NewObject("GraphicsCaptureItem", False)
    Set GraphicsCaptureItemStatics = IGraphicsCaptureItemStatics
End Function

Public Function GraphicsCaptureSessionStatics() As IGraphicsCaptureSessionStatics
Static IGraphicsCaptureSessionStatics As IGraphicsCaptureSessionStatics
    If IGraphicsCaptureSessionStatics Is Nothing Then Set IGraphicsCaptureSessionStatics = NewObject("GraphicsCaptureSession", False)
    Set GraphicsCaptureSessionStatics = IGraphicsCaptureSessionStatics
End Function

Public Function GridLengthHelperStatics() As IGridLengthHelperStatics
Static IGridLengthHelperStatics As IGridLengthHelperStatics
    If IGridLengthHelperStatics Is Nothing Then Set IGridLengthHelperStatics = NewObject("GridLengthHelper", False)
    Set GridLengthHelperStatics = IGridLengthHelperStatics
End Function

Public Function GuidHelperStatics() As IGuidHelperStatics
Static IGuidHelperStatics As IGuidHelperStatics
    If IGuidHelperStatics Is Nothing Then Set IGuidHelperStatics = NewObject("GuidHelper", False)
    Set GuidHelperStatics = IGuidHelperStatics
End Function

Public Function HashAlgorithmNamesStatics() As IHashAlgorithmNamesStatics
Static IHashAlgorithmNamesStatics As IHashAlgorithmNamesStatics
    If IHashAlgorithmNamesStatics Is Nothing Then Set IHashAlgorithmNamesStatics = NewObject("HashAlgorithmNames", False)
    Set HashAlgorithmNamesStatics = IHashAlgorithmNamesStatics
End Function

Public Function HashAlgorithmProviderStatics() As IHashAlgorithmProviderStatics
Static IHashAlgorithmProviderStatics As IHashAlgorithmProviderStatics
    If IHashAlgorithmProviderStatics Is Nothing Then Set IHashAlgorithmProviderStatics = NewObject("HashAlgorithmProvider", False)
    Set HashAlgorithmProviderStatics = IHashAlgorithmProviderStatics
End Function

Public Function HostNameFactory() As IHostNameFactory
Static IHostNameFactory As IHostNameFactory
    If IHostNameFactory Is Nothing Then Set IHostNameFactory = NewObject("HostName", False)
    Set HostNameFactory = IHostNameFactory
End Function

Public Function HtmlUtilities() As IHtmlUtilities
Static IHtmlUtilities As IHtmlUtilities
    If IHtmlUtilities Is Nothing Then Set IHtmlUtilities = NewObject("HtmlUtilities", False)
    Set HtmlUtilities = IHtmlUtilities
End Function

Public Function HttpBufferContentFactory() As IHttpBufferContentFactory
Static IHttpBufferContentFactory As IHttpBufferContentFactory
    If IHttpBufferContentFactory Is Nothing Then Set IHttpBufferContentFactory = NewObject("HttpBufferContent", False)
    Set HttpBufferContentFactory = IHttpBufferContentFactory
End Function

Public Function HttpCredentialsHeaderValueFactory() As IHttpCredentialsHeaderValueFactory
Static IHttpCredentialsHeaderValueFactory As IHttpCredentialsHeaderValueFactory
    If IHttpCredentialsHeaderValueFactory Is Nothing Then Set IHttpCredentialsHeaderValueFactory = NewObject("HttpCredentialsHeaderValue", False)
    Set HttpCredentialsHeaderValueFactory = IHttpCredentialsHeaderValueFactory
End Function

Public Function HttpFormUrlEncodedContentFactory() As IHttpFormUrlEncodedContentFactory
Static IHttpFormUrlEncodedContentFactory As IHttpFormUrlEncodedContentFactory
    If IHttpFormUrlEncodedContentFactory Is Nothing Then Set IHttpFormUrlEncodedContentFactory = NewObject("HttpFormUrlEncodedContent", False)
    Set HttpFormUrlEncodedContentFactory = IHttpFormUrlEncodedContentFactory
End Function

Public Function HttpMediaTypeHeaderValueFactory() As IHttpMediaTypeHeaderValueFactory
Static IHttpMediaTypeHeaderValueFactory As IHttpMediaTypeHeaderValueFactory
    If IHttpMediaTypeHeaderValueFactory Is Nothing Then Set IHttpMediaTypeHeaderValueFactory = NewObject("HttpMediaTypeHeaderValue", False)
    Set HttpMediaTypeHeaderValueFactory = IHttpMediaTypeHeaderValueFactory
End Function

Public Function HttpMultipartContentFactory() As IHttpMultipartContentFactory
Static IHttpMultipartContentFactory As IHttpMultipartContentFactory
    If IHttpMultipartContentFactory Is Nothing Then Set IHttpMultipartContentFactory = NewObject("HttpMultipartContent", False)
    Set HttpMultipartContentFactory = IHttpMultipartContentFactory
End Function

Public Function HttpMultipartFormDataContentFactory() As IHttpMultipartFormDataContentFactory
Static IHttpMultipartFormDataContentFactory As IHttpMultipartFormDataContentFactory
    If IHttpMultipartFormDataContentFactory Is Nothing Then Set IHttpMultipartFormDataContentFactory = NewObject("HttpMultipartFormDataContent", False)
    Set HttpMultipartFormDataContentFactory = IHttpMultipartFormDataContentFactory
End Function

Public Function HttpNameValueHeaderValueFactory() As IHttpNameValueHeaderValueFactory
Static IHttpNameValueHeaderValueFactory As IHttpNameValueHeaderValueFactory
    If IHttpNameValueHeaderValueFactory Is Nothing Then Set IHttpNameValueHeaderValueFactory = NewObject("HttpNameValueHeaderValue", False)
    Set HttpNameValueHeaderValueFactory = IHttpNameValueHeaderValueFactory
End Function

Public Function HttpStreamContentFactory() As IHttpStreamContentFactory
Static IHttpStreamContentFactory As IHttpStreamContentFactory
    If IHttpStreamContentFactory Is Nothing Then Set IHttpStreamContentFactory = NewObject("HttpStreamContent", False)
    Set HttpStreamContentFactory = IHttpStreamContentFactory
End Function

Public Function HttpStringContentFactory() As IHttpStringContentFactory
Static IHttpStringContentFactory As IHttpStringContentFactory
    If IHttpStringContentFactory Is Nothing Then Set IHttpStringContentFactory = NewObject("HttpStringContent", False)
    Set HttpStringContentFactory = IHttpStringContentFactory
End Function

Public Function ImageStatics() As IImageStatics
Static IImageStatics As IImageStatics
    If IImageStatics Is Nothing Then Set IImageStatics = NewObject("Image", False)
    Set ImageStatics = IImageStatics
End Function

Public Function JsonArrayStatics() As IJsonArrayStatics
Static IJsonArrayStatics As IJsonArrayStatics
    If IJsonArrayStatics Is Nothing Then Set IJsonArrayStatics = NewObject("JsonArray", False)
    Set JsonArrayStatics = IJsonArrayStatics
End Function

Public Function JsonObjectStatics() As IJsonObjectStatics
Static IJsonObjectStatics As IJsonObjectStatics
    If IJsonObjectStatics Is Nothing Then Set IJsonObjectStatics = NewObject("JsonObject", False)
    Set JsonObjectStatics = IJsonObjectStatics
End Function

Public Function JsonValueStatics() As IJsonValueStatics
Static IJsonValueStatics As IJsonValueStatics
    If IJsonValueStatics Is Nothing Then Set IJsonValueStatics = NewObject("JsonValue", False)
    Set JsonValueStatics = IJsonValueStatics
End Function

Public Function KeyTimeHelperStatics() As IKeyTimeHelperStatics
Static IKeyTimeHelperStatics As IKeyTimeHelperStatics
    If IKeyTimeHelperStatics Is Nothing Then Set IKeyTimeHelperStatics = NewObject("KeyTimeHelper", False)
    Set KeyTimeHelperStatics = IKeyTimeHelperStatics
End Function

Public Function KnownUserPropertiesStatics() As IKnownUserPropertiesStatics
Static IKnownUserPropertiesStatics As IKnownUserPropertiesStatics
    If IKnownUserPropertiesStatics Is Nothing Then Set IKnownUserPropertiesStatics = NewObject("KnownUserProperties", False)
    Set KnownUserPropertiesStatics = IKnownUserPropertiesStatics
End Function

Public Function LanguageFactory() As ILanguageFactory
Static ILanguageFactory As ILanguageFactory
    If ILanguageFactory Is Nothing Then Set ILanguageFactory = NewObject("Language", False)
    Set LanguageFactory = ILanguageFactory
End Function

Public Function LauncherStatics() As ILauncherStatics
Static ILauncherStatics As ILauncherStatics
    If ILauncherStatics Is Nothing Then Set ILauncherStatics = NewObject("Launcher", False)
    Set LauncherStatics = ILauncherStatics
End Function

Public Function Matrix3DHelperStatics() As IMatrix3DHelperStatics
Static IMatrix3DHelperStatics As IMatrix3DHelperStatics
    If IMatrix3DHelperStatics Is Nothing Then Set IMatrix3DHelperStatics = NewObject("Matrix3DHelper", False)
    Set Matrix3DHelperStatics = IMatrix3DHelperStatics
End Function

Public Function MatrixHelperStatics() As IMatrixHelperStatics
Static IMatrixHelperStatics As IMatrixHelperStatics
    If IMatrixHelperStatics Is Nothing Then Set IMatrixHelperStatics = NewObject("MatrixHelper", False)
    Set MatrixHelperStatics = IMatrixHelperStatics
End Function

Public Function MediaSourceStatics() As IMediaSourceStatics
Static IMediaSourceStatics As IMediaSourceStatics
    If IMediaSourceStatics Is Nothing Then Set IMediaSourceStatics = NewObject("MediaSource", False)
    Set MediaSourceStatics = IMediaSourceStatics
End Function

Public Function MessageDialogFactory() As IMessageDialogFactory
Static IMessageDialogFactory As IMessageDialogFactory
    If IMessageDialogFactory Is Nothing Then Set IMessageDialogFactory = NewObject("MessageDialog", False)
    Set MessageDialogFactory = IMessageDialogFactory
End Function

Public Function OcrEngineStatics() As IOcrEngineStatics
Static IOcrEngineStatics As IOcrEngineStatics
    If IOcrEngineStatics Is Nothing Then Set IOcrEngineStatics = NewObject("OcrEngine", False)
    Set OcrEngineStatics = IOcrEngineStatics
End Function

Public Function PathIOStatics() As IPathIOStatics
Static IPathIOStatics As IPathIOStatics
    If IPathIOStatics Is Nothing Then Set IPathIOStatics = NewObject("PathIO", False)
    Set PathIOStatics = IPathIOStatics
End Function

Public Function PdfDocumentStatics() As IPdfDocumentStatics
Static IPdfDocumentStatics As IPdfDocumentStatics
    If IPdfDocumentStatics Is Nothing Then Set IPdfDocumentStatics = NewObject("PdfDocument", False)
    Set PdfDocumentStatics = IPdfDocumentStatics
End Function

Public Function PlaylistStatics() As IPlaylistStatics
Static IPlaylistStatics As IPlaylistStatics
    If IPlaylistStatics Is Nothing Then Set IPlaylistStatics = NewObject("Playlist", False)
    Set PlaylistStatics = IPlaylistStatics
End Function

Public Function PointHelperStatics() As IPointHelperStatics
Static IPointHelperStatics As IPointHelperStatics
    If IPointHelperStatics Is Nothing Then Set IPointHelperStatics = NewObject("PointHelper", False)
    Set PointHelperStatics = IPointHelperStatics
End Function

Public Function PropertyPathFactory() As IPropertyPathFactory
Static IPropertyPathFactory As IPropertyPathFactory
    If IPropertyPathFactory Is Nothing Then Set IPropertyPathFactory = NewObject("PropertyPath", False)
    Set PropertyPathFactory = IPropertyPathFactory
End Function

Public Function PropertyValueStatics() As IPropertyValueStatics
Static IPropertyValueStatics As IPropertyValueStatics
    If IPropertyValueStatics Is Nothing Then Set IPropertyValueStatics = NewObject("PropertyValue", False)
    Set PropertyValueStatics = IPropertyValueStatics
End Function

Public Function RandomAccessStreamReferenceStatics() As IRandomAccessStreamReferenceStatics
    Static IRandomAccessStreamReferenceStatics As IRandomAccessStreamReferenceStatics
    If IRandomAccessStreamReferenceStatics Is Nothing Then Set IRandomAccessStreamReferenceStatics = NewObject("RandomAccessStreamReference", False)
    Set RandomAccessStreamReferenceStatics = IRandomAccessStreamReferenceStatics
End Function

Public Function RandomAccessStreamStatics() As IRandomAccessStreamStatics
    Static IRandomAccessStreamStatics As IRandomAccessStreamStatics
    If IRandomAccessStreamStatics Is Nothing Then Set IRandomAccessStreamStatics = NewObject("RandomAccessStream", False)
    Set RandomAccessStreamStatics = IRandomAccessStreamStatics
End Function

Public Function RectHelperStatics() As IRectHelperStatics
    Static IRectHelperStatics As IRectHelperStatics
    If IRectHelperStatics Is Nothing Then Set IRectHelperStatics = NewObject("RectHelper", False)
    Set RectHelperStatics = IRectHelperStatics
End Function

Public Function RepeatBehaviorHelperStatics() As IRepeatBehaviorHelperStatics
Static IRepeatBehaviorHelperStatics As IRepeatBehaviorHelperStatics
    If IRepeatBehaviorHelperStatics Is Nothing Then Set IRepeatBehaviorHelperStatics = NewObject("RepeatBehaviorHelper", False)
    Set RepeatBehaviorHelperStatics = IRepeatBehaviorHelperStatics
End Function

Public Function SerialDeviceStatics() As ISerialDeviceStatics
Static ISerialDeviceStatics As ISerialDeviceStatics
    If ISerialDeviceStatics Is Nothing Then Set ISerialDeviceStatics = NewObject("SerialDevice", False)
    Set SerialDeviceStatics = ISerialDeviceStatics
End Function

Public Function SignalNotifierStatics() As ISignalNotifierStatics
Static ISignalNotifierStatics As ISignalNotifierStatics
    If ISignalNotifierStatics Is Nothing Then Set ISignalNotifierStatics = NewObject("SignalNotifier", False)
    Set SignalNotifierStatics = ISignalNotifierStatics
End Function

Public Function SizeHelperStatics() As ISizeHelperStatics
Static ISizeHelperStatics As ISizeHelperStatics
    If ISizeHelperStatics Is Nothing Then Set ISizeHelperStatics = NewObject("SizeHelper", False)
    Set SizeHelperStatics = ISizeHelperStatics
End Function

Public Function SoftwareBitmapFactory() As ISoftwareBitmapFactory
Static ISoftwareBitmapFactory As ISoftwareBitmapFactory
    If ISoftwareBitmapFactory Is Nothing Then Set ISoftwareBitmapFactory = NewObject("SoftwareBitmap", False)
    Set SoftwareBitmapFactory = ISoftwareBitmapFactory
End Function

Public Function SoftwareBitmapStatics() As ISoftwareBitmapStatics
Static ISoftwareBitmapStatics As ISoftwareBitmapStatics
    If ISoftwareBitmapStatics Is Nothing Then Set ISoftwareBitmapStatics = NewObject("SoftwareBitmap", False)
    Set SoftwareBitmapStatics = ISoftwareBitmapStatics
End Function

Public Function SolidColorBrushFactory() As ISolidColorBrushFactory
Static ISolidColorBrushFactory As ISolidColorBrushFactory
    If ISolidColorBrushFactory Is Nothing Then Set ISolidColorBrushFactory = NewObject("SolidColorBrush", False)
    Set SolidColorBrushFactory = ISolidColorBrushFactory
End Function

Public Function StandardDataFormatsStatics() As IStandardDataFormatsStatics
Static IStandardDataFormatsStatics As IStandardDataFormatsStatics
    If IStandardDataFormatsStatics Is Nothing Then Set IStandardDataFormatsStatics = NewObject("StandardDataFormats", False)
    Set StandardDataFormatsStatics = IStandardDataFormatsStatics
End Function

Public Function StorageFileStatics() As IStorageFileStatics
Static IStorageFileStatics As IStorageFileStatics
    If IStorageFileStatics Is Nothing Then Set IStorageFileStatics = NewObject("StorageFile", False)
    Set StorageFileStatics = IStorageFileStatics
End Function

Public Function StorageFolderStatics() As IStorageFolderStatics
Static IStorageFolderStatics As IStorageFolderStatics
    If IStorageFolderStatics Is Nothing Then Set IStorageFolderStatics = NewObject("StorageFolder", False)
    Set StorageFolderStatics = IStorageFolderStatics
End Function

Public Function StoryboardStatics() As IStoryboardStatics
Static IStoryboardStatics As IStoryboardStatics
    If IStoryboardStatics Is Nothing Then Set IStoryboardStatics = NewObject("Storyboard", False)
    Set StoryboardStatics = IStoryboardStatics
End Function

Public Function SystemIdentificationStatics() As ISystemIdentificationStatics
Static ISystemIdentificationStatics As ISystemIdentificationStatics
    If ISystemIdentificationStatics Is Nothing Then Set ISystemIdentificationStatics = NewObject("SystemIdentification", False)
    Set SystemIdentificationStatics = ISystemIdentificationStatics
End Function

Public Function TextBoxStatics() As ITextBoxStatics
Static ITextBoxStatics As ITextBoxStatics
    If ITextBoxStatics Is Nothing Then Set ITextBoxStatics = NewObject("TextBox", False)
    Set TextBoxStatics = ITextBoxStatics
End Function

Public Function ThicknessHelperStatics() As IThicknessHelperStatics
    Static IThicknessHelperStatics As IThicknessHelperStatics
    If IThicknessHelperStatics Is Nothing Then Set IThicknessHelperStatics = NewObject("ThicknessHelper", False)
    Set ThicknessHelperStatics = IThicknessHelperStatics
End Function

Public Function ThreadPoolStatics() As IThreadPoolStatics
Static IThreadPoolStatics As IThreadPoolStatics
    If IThreadPoolStatics Is Nothing Then Set IThreadPoolStatics = NewObject("ThreadPool", False)
    Set ThreadPoolStatics = IThreadPoolStatics
End Function

Public Function ThreadPoolTimerStatics() As IThreadPoolTimerStatics
Static IThreadPoolTimerStatics As IThreadPoolTimerStatics
    If IThreadPoolTimerStatics Is Nothing Then Set IThreadPoolTimerStatics = NewObject("ThreadPoolTimer", False)
    Set ThreadPoolTimerStatics = IThreadPoolTimerStatics
End Function

Public Function ToolTipServiceStatics() As IToolTipServiceStatics
Static IToolTipServiceStatics As IToolTipServiceStatics
    If IToolTipServiceStatics Is Nothing Then Set IToolTipServiceStatics = NewObject("ToolTipService", False)
    Set ToolTipServiceStatics = IToolTipServiceStatics
End Function

Public Function UICommandFactory() As IUICommandFactory
Static IUICommandFactory As IUICommandFactory
    If IUICommandFactory Is Nothing Then Set IUICommandFactory = NewObject("UICommand", False)
    Set UICommandFactory = IUICommandFactory
End Function

Public Function UriEscapeStatics() As IUriEscapeStatics
Static IUriEscapeStatics As IUriEscapeStatics
    If IUriEscapeStatics Is Nothing Then Set IUriEscapeStatics = NewObject("Uri", False)
    Set UriEscapeStatics = IUriEscapeStatics
End Function

Public Function UriFactory() As IUriFactory
Static IUriFactory As IUriFactory
    If IUriFactory Is Nothing Then Set IUriFactory = NewObject("Uri", False)
    Set UriFactory = IUriFactory
End Function

Public Function UserConsentVerifierStatics() As IUserConsentVerifierStatics
Static IUserConsentVerifierStatics As IUserConsentVerifierStatics
    If IUserConsentVerifierStatics Is Nothing Then Set IUserConsentVerifierStatics = NewObject("UserConsentVerifier", False)
    Set UserConsentVerifierStatics = IUserConsentVerifierStatics
End Function

Public Function UserDataPathsStatics() As IUserDataPathsStatics
Static IUserDataPathsStatics As IUserDataPathsStatics
    If IUserDataPathsStatics Is Nothing Then Set IUserDataPathsStatics = NewObject("UserDataPaths", False)
    Set UserDataPathsStatics = IUserDataPathsStatics
End Function

Public Function UserStatics() As IUserStatics
Static IUserStatics As IUserStatics
    If IUserStatics Is Nothing Then Set IUserStatics = NewObject("User", False)
    Set UserStatics = IUserStatics
End Function

Public Function VisualTreeHelperStatics() As IVisualTreeHelperStatics
Static IVisualTreeHelperStatics As IVisualTreeHelperStatics
    If IVisualTreeHelperStatics Is Nothing Then Set IVisualTreeHelperStatics = NewObject("VisualTreeHelper", False)
    Set VisualTreeHelperStatics = IVisualTreeHelperStatics
End Function

Public Function WindowStatics() As IWindowStatics
Static IWindowStatics As IWindowStatics
    If IWindowStatics Is Nothing Then Set IWindowStatics = NewObject("Window", False)
    Set WindowStatics = IWindowStatics
End Function

Public Function WindowsXamlManagerStatics() As IWindowsXamlManagerStatics
Static IWindowsXamlManagerStatics As IWindowsXamlManagerStatics
    If IWindowsXamlManagerStatics Is Nothing Then Set IWindowsXamlManagerStatics = NewObject("WindowsXamlManager", False)
    Set WindowsXamlManagerStatics = IWindowsXamlManagerStatics
End Function

Public Function WriteableBitmapFactory() As IWriteableBitmapFactory
Static IWriteableBitmapFactory As IWriteableBitmapFactory
    If IWriteableBitmapFactory Is Nothing Then Set IWriteableBitmapFactory = NewObject("WriteableBitmap", False)
    Set WriteableBitmapFactory = IWriteableBitmapFactory
End Function

Public Function XamlReaderStatics() As IXamlReaderStatics
Static IXamlReaderStatics As IXamlReaderStatics
    If IXamlReaderStatics Is Nothing Then Set IXamlReaderStatics = NewObject("XamlReader", False)
    Set XamlReaderStatics = IXamlReaderStatics
End Function

Public Function XamlSourceFocusNavigationRequestFactory() As IXamlSourceFocusNavigationRequestFactory
Static IXamlSourceFocusNavigationRequestFactory As IXamlSourceFocusNavigationRequestFactory
    If IXamlSourceFocusNavigationRequestFactory Is Nothing Then Set IXamlSourceFocusNavigationRequestFactory = NewObject("XamlSourceFocusNavigationRequest", False)
    Set XamlSourceFocusNavigationRequestFactory = IXamlSourceFocusNavigationRequestFactory
End Function
