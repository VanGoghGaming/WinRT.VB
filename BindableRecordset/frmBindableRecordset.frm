VERSION 5.00
Begin VB.Form frmBindableRecordset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bindable Recordset Demo"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBindableRecordset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   19110
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmBindableRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IRoutedEventHandler

Private m_DesktopWindowXamlSource As IDesktopWindowXamlSource, m_ListViewNorthwind As IItemsControl, m_FadeAnimation As IStoryboard

Private Sub LoadXaml(sXamlString As String)
Dim Connection As New ADODB.Connection, RecordSet As New ADODB.RecordSet, UIElement5 As IUIElement5, LayoutRoot As IFrameworkElement
    Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Northwind.mdb;" ' This Microsoft Jet driver is only available in 32-bit. You need an ACE driver for 64-bit.
    RecordSet.Open "SELECT * FROM Customers", Connection, adOpenKeyset, adLockOptimistic, adCmdText ' Open the RecordSet in Read/Write mode so that we can update records
    Set LayoutRoot = XamlReaderStatics.Load(StrRef(sXamlString)) ' Parse the XAML string and instantiate all controls
    Set UIElement5 = LayoutRoot: UIElement5.TabFocusNavigation = KeyboardNavigationMode_Cycle ' Keyboard focus (Tab) will cycle back to the first element when reaching the last one
    With LayoutRoot ' Search controls by their name (defined in the XAML string) and assign them to variables
        Set m_FadeAnimation = .FindName(StrRef("FadeAnimation")) ' This is our Storyboard Opacity/Scale animation
        Set m_ListViewNorthwind = .FindName(StrRef("lvNorthwind")) ' This is our ListView control
        With New cBindableRecordSet
            Set m_ListViewNorthwind.ItemsSource = .This(RecordSet) ' Bind the ItemsSource property to our implementation of IBindableVector (basically a zero-based array of rows where each row is an observable collection of objects)
        End With
        .AddLoaded Me ' Subscribe to the Loaded event (fires when the all XAML content has been successfully parsed and instantiated)
    End With
    Set m_DesktopWindowXamlSource.Content = LayoutRoot
End Sub

Private Sub Form_Load()
Dim ClientRect As RECTL, DesktopWindowXamlSourceNative As IDesktopWindowXamlSourceNative
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    Set m_DesktopWindowXamlSource = NewObject("DesktopWindowXamlSource") ' Instantiate our XAML Island container
    If GetClientRect(hWnd, ClientRect) = APITRUE Then ' Get the size of the main form that hosts the XAML Island
        Set DesktopWindowXamlSourceNative = m_DesktopWindowXamlSource
        With DesktopWindowXamlSourceNative ' This interface provides access to the window handle (hWnd) of the XAML Island
            If .AttachToWindow(hWnd) = S_OK Then ' Here we attach the XAML Island to our form and then resize it to fit the entire form and make it visible (it's hidden by default)
                SetWindowPos .WindowHandle, HWND_TOP, 0, 0, ClientRect.Right, ClientRect.Bottom, SWP_SHOWWINDOW Or SWP_NOMOVE
                LoadXaml StrConv(LoadResData("XAML", "CUSTOM"), vbUnicode) ' Everything has been set up to start loading and parsing the XAML code from resources
            End If
        End With
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_ListViewNorthwind.ItemsSource = Nothing ' Remove the binding to our cBindableRecordSet object to stop receiving notifications and prepare for closing the database connection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dispose m_DesktopWindowXamlSource ' This object implements the IClosable interface (equivalent to .NET's IDisposable) that handles graceful cleanup of allocated resources. The form is closing anyway so cleanup is redundant at this point.
End Sub

Private Sub IRoutedEventHandler_Invoke(ByVal Sender As IInspectable, ByVal Args As IRoutedEventArgs) ' This handles the Loaded event that we subscribed to earlier
    Select Case True
        Case Sender Is m_DesktopWindowXamlSource.Content ' Check whether this event comes from the XAML Island itself
            Visible = True ' Everything has been loaded at this point so Show the main form and set the keyboard focus to the first focusable element
            m_DesktopWindowXamlSource.NavigateFocus XamlSourceFocusNavigationRequestFactory.CreateInstance(XamlSourceFocusNavigationReason_First)
            m_FadeAnimation.Begin ' Start our cheesy Opacity/Scale animation of the Rainbow TextBlock
    End Select
End Sub
