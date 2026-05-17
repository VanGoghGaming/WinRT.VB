VERSION 5.00
Begin VB.Form frmPDFViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDF Viewer"
   ClientHeight    =   14565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17910
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPDFViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14565
   ScaleWidth      =   17910
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmPDFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IAsyncOperationCompletedHandlerIStorageFile
Implements IRoutedEventHandler
Implements ITypedEventHandlerListViewBaseContainerContentChanging
Implements ITypedEventHandlerXamlUICommandExecuteRequested

Private m_DesktopWindowXamlSource As IDesktopWindowXamlSource, m_hWndIsland As LongPtr, m_PdfListView As IItemsControl, m_PdfDocument As IPdfDocument, m_PdfPageRenderOptions As IPdfPageRenderOptions, _
        m_XamlUICommand As IXamlUICommand, m_RotationMenu As IMenuBarItem, m_RandomAccessStream As IRandomAccessStream, m_PdfPageRotation As PdfPageRotation, m_FileOpenPicker As IFileOpenPicker

Private Sub Form_Load()
Dim ClientRect As RECTL, DesktopWindowXamlSourceNative As IDesktopWindowXamlSourceNative
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    If GetClientRect(hWnd, ClientRect) = APITRUE Then
        Set m_DesktopWindowXamlSource = NewObject("DesktopWindowXamlSource") ' The DesktopWindowXamlSource object allows hosting of XAML Islands on a regular Win32 hWnd
        Set DesktopWindowXamlSourceNative = m_DesktopWindowXamlSource
        With DesktopWindowXamlSourceNative
            If .AttachToWindow(hWnd) = S_OK Then ' Use the Form itself as a container for this XAML Island
                m_hWndIsland = .WindowHandle
                If SetWindowPos(m_hWndIsland, HWND_TOP, 0, 0, ClientRect.Right, ClientRect.Bottom, SWP_SHOWWINDOW Or SWP_NOMOVE) = APITRUE Then ' The XAML Island window is hidden by default so here we resize and make it visible
                    Set m_DesktopWindowXamlSource.Content = LayoutRoot(StrConv(LoadResData("XAML", "CUSTOM"), vbUnicode))
                End If
            End If
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dispose m_RandomAccessStream: Dispose m_DesktopWindowXamlSource ' Destroy the XAML Island gracefully
End Sub

Private Function GetColorFromMenuItemTag(SolidColorBrush As ISolidColorBrush) As ColorLong
    GetColorFromMenuItemTag = SolidColorBrush.Color
End Function

Private Function GetPageImage(Container As IContentControl2) As IFrameworkElement
Dim ContentTemplateRoot As IFrameworkElement
    Set ContentTemplateRoot = Container.ContentTemplateRoot ' Grab the root element of the template for this ListViewItem
    Set GetPageImage = ContentTemplateRoot.FindName(StrRef("PageImage")) ' and find the name of the Image control in the template so that we can access its properties
End Function

Private Sub IAsyncOperationCompletedHandlerIStorageFile_Invoke(ByVal AsyncOperation As IAsyncOperationIStorageFile, ByVal AsyncStatus As AsyncStatus) ' Handler for the FileOpenPicker
Dim StorageFile As IStorageFile
    Select Case AsyncStatus
        Case AsyncStatus_Completed: Set StorageFile = AsyncOperation.GetResults: If Not StorageFile Is Nothing Then LoadNewPdfFile StorageFile
        Case AsyncStatus_Error: MsgBox "FileOpenPicker can't run elevated (as administrator)!", vbOKOnly + vbExclamation, App.Title
    End Select
End Sub

Private Sub IRoutedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IRoutedEventArgs) ' This the Loaded event (fires once when the XAML content has been successfully parsed and instantiated)
    LoadNewPdfFile ' Load the default PDF sample file from a URI
End Sub

Private Sub ITypedEventHandlerListViewBaseContainerContentChanging_Invoke(ByVal Sender As IListViewBase, ByVal Args As IContainerContentChangingEventArgs)
Dim Image As IFrameworkElement, RenderPage As cRenderPage
    With Args
        Set Image = GetPageImage(.ItemContainer): .Handled = APITRUE ' Retrieve the Image control from this ListViewItem and mark the event as handled
        If .InRecycleQueue Then ' This ListViewItem has scrolled outside the cached ViewPort area so we can free its resources
            Set RenderPage = Image.Tag ' Retrieve the instance of cRenderPage saved in the Tag property
            RenderPage.CancelRendering ' If the PdfPage is still rendering this image we cancel the operation since it's useless at this point
            Set Image.Tag = Nothing ' This triggers the Class_Terminate event of cRenderPage and releases allocated resources
        Else
            Select Case .Phase ' At phase 0 we start rendering, subsequent phases check for completion
                Case 0
                    With New cRenderPage ' Save an instance of our render class in the Tag property and start rendering this page:
                        Set Image.Tag = .StartRendering(m_PdfDocument.GetPage(Args.ItemIndex), m_PdfPageRenderOptions, m_PdfPageRotation)
                    End With
                    .RegisterUpdateCallback Me ' Page rendering is an async operation so we queue this event again until it completes (also increases the phase number)
                Case Else ' Subsequent callbacks of this event increase the phase number
                    Set RenderPage = Image.Tag
                    If RenderPage.IsRenderingFinished Then
                        Set Image.DataContext = RenderPage.BitmapSource ' The Image control has its Source property bound in XAML so it will pick up the image from this BitmapSource
                    Else
                        .RegisterUpdateCallback Me ' Rendering hasn't completed yet, queue the event again
                    End If
            End Select
        End If
    End With
End Sub

Private Sub ITypedEventHandlerXamlUICommandExecuteRequested_Invoke(ByVal Sender As IXamlUICommand, ByVal Args As IExecuteRequestedEventArgs) ' This handler is invoked for every menu item
Dim MenuFlyoutItem As IFrameworkElement, sMenuName As String, PdfPageRotation As PdfPageRotation
    If TypeOf Args.Parameter Is IMenuFlyoutItem Then
        Set MenuFlyoutItem = Args.Parameter: sMenuName = GetStr(MenuFlyoutItem.Name)
        Select Case True ' Decide which menu item has been interacted with:
            Case sMenuName = "OpenFile" ' Open the FileOpenPicker dialog to select a new PDF file to display:
                Set m_FileOpenPicker.PickSingleFileAsync.Completed = Me
            Case sMenuName = "Quit" ' Close the form asynchronously:
                PostMessageW hWnd, WM_CLOSE
            Case InStr(sMenuName, "Rotation") > 0
                SelectRotationMenuItem MenuFlyoutItem: PdfPageRotation = Unbox(MenuFlyoutItem.Tag) ' Get the page rotation value from the Tag property
                If m_PdfPageRotation <> PdfPageRotation Then m_PdfPageRotation = PdfPageRotation: UpdateListViewContents
            Case TypeOf MenuFlyoutItem.Tag Is ISolidColorBrush
                m_PdfPageRenderOptions.BackgroundColor = GetColorFromMenuItemTag(MenuFlyoutItem.Tag): UpdateListViewContents
        End Select
    End If
End Sub

Private Function LayoutRoot(sXamlString As String) As IFrameworkElement
Dim PdfListView As IListViewBase2
    Set m_FileOpenPicker = NewObject("FileOpenPicker"): InitializeWithWindow m_FileOpenPicker, m_hWndIsland: m_FileOpenPicker.FileTypeFilter.Append StrRef(".pdf")
    Set m_PdfPageRenderOptions = NewObject("PdfPageRenderOptions"): Set LayoutRoot = XamlReaderStatics.Load(StrRef(sXamlString))
    With LayoutRoot ' Find these controls by their name from the XAML source
        Set m_RotationMenu = .FindName(StrRef("RotationMenu"))
        Set PdfListView = .FindName(StrRef("PdfListView")): Set m_PdfListView = PdfListView
        Set m_XamlUICommand = .FindName(StrRef("MenuCommand")) ' This XamlUICommand is bound to all menu items to aggregate their functionality in one place instead of a separate click event for each
        .AddLoaded Me ' Subscribe to the Loaded event
    End With
    m_XamlUICommand.AddExecuteRequested Me ' Subscribe to the ExecuteRequested event (fires when a menu item is invoked)
    PdfListView.AddContainerContentChanging Me ' Subscribe to the ContainerContentChanging event (fires when the ListView control loads a new container (ListViewItem) or recycles an old one)
End Function

Private Sub LoadNewPdfFile(Optional StorageFile As IStorageFile, Optional sUri As String = "https://www.princexml.com/samples/icelandic/dictionary.pdf")
Dim UriStream As IRandomAccessStream
    With New cAwait
        If StorageFile Is Nothing Then
            If .Await(RandomAccessStreamReferenceStatics.CreateFromUri(UriFactory.CreateUri(StrRef(sUri))).OpenReadAsync) = AsyncStatus_Completed Then
                Set UriStream = .GetResults: Dispose m_RandomAccessStream: Set m_RandomAccessStream = NewObject("InMemoryRandomAccessStream")
                If .Await(RandomAccessStreamStatics.CopyAsync(UriStream, m_RandomAccessStream)) = AsyncStatus_Completed Then ' Uri streams are not seekable so make a local copy in memory
                    If .Await(PdfDocumentStatics.LoadFromStreamAsync(m_RandomAccessStream)) = AsyncStatus_Completed Then Set m_PdfDocument = .GetResults
                End If
            End If
            Dispose UriStream
        Else
            If .Await(PdfDocumentStatics.LoadFromFileAsync(StorageFile)) = AsyncStatus_Completed Then Set m_PdfDocument = .GetResults: Dispose m_RandomAccessStream
        End If
    End With
    UpdateListViewContents
End Sub

Private Sub SelectRotationMenuItem(MenuFlyoutItem As IMenuFlyoutItem)
Dim ToggleMenuFlyoutItem As IToggleMenuFlyoutItem, i As Long
    With m_RotationMenu.Items
        For i = 0 To .Size - 1
            Set ToggleMenuFlyoutItem = .GetAt(i)
            ToggleMenuFlyoutItem.IsChecked = Abs(ToggleMenuFlyoutItem Is MenuFlyoutItem) ' Make sure only the currently selected menu item has a check mark
        Next i
    End With
End Sub

Private Sub SetFocusProgrammatic(Control As IControl)
    Control.Focus FocusState_Programmatic
End Sub

Private Sub UpdateListViewContents()
    With New cBindableVector
        Set m_PdfListView.ItemsSource = .This(m_PdfDocument.PageCount)
    End With
    SetFocusProgrammatic m_PdfListView
End Sub
