VERSION 5.00
Begin VB.Form frmGIFPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIF Player"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGIFPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picImage 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuGIF 
         Caption         =   "Play Data.gif"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuGIF 
         Caption         =   "Play Travolta.gif"
         Index           =   1
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuGIF 
         Caption         =   "Play Cheers.gif"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuGIF 
         Caption         =   "Play MrBean.gif"
         Index           =   3
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuGIF 
         Caption         =   "Play Matrix.gif"
         Index           =   4
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuGIF 
         Caption         =   "Play Kramer.gif"
         Index           =   5
         Shortcut        =   ^{F6}
      End
   End
End
Attribute VB_Name = "frmGIFPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements cEventsSink

Private m_DesktopWindowXamlSource As IDesktopWindowXamlSource, m_hWndIsland As LongPtr, m_BitmapImage As IBitmapImage

Private Sub cEventsSink_EventHandler(ByVal EventIdentifier As EVENT_IDENTIFIERS, ByVal Sender As IInspectable, Optional ByVal PointerRoutedEventArgs As IPointerRoutedEventArgs, Optional ByVal RoutedEventArgs As IRoutedEventArgs, Optional ByVal ExceptionRoutedEventArgs As IExceptionRoutedEventArgs, Optional ByVal TappedRoutedEventArgs As ITappedRoutedEventArgs)
Dim BitmapSource As IBitmapSource, BitmapImage3 As IBitmapImage3
    Select Case EventIdentifier
        Case Event_Loaded ' Fires once when the XAML Island is first loaded
            mnuGIF_Click 0 ' Play the first menu item
        Case Event_ImageOpened ' Fires every time a new image is successfully opened
            Set BitmapSource = Sender
            With picImage ' Set the size of the PictureBox container to the actual size of the image just opened
                .Move .Left, .Top, .ScaleX(BitmapSource.PixelWidth, vbPixels, .ScaleMode), .ScaleY(BitmapSource.PixelHeight, vbPixels, .ScaleMode)
            End With
        Case Event_ImageFailed ' Images may fail loading if the URL is incorrect or the file is an unsupported format
            MsgBox GetStr(ExceptionRoutedEventArgs.ErrorMessage), vbOKOnly + vbExclamation, App.Title
        Case Event_Tapped ' Check whether the loaded image is an animated GIF and if so then toggle playing on/off
            If TypeOf m_BitmapImage Is IBitmapImage3 Then ' This IBitmapImage3 interface requires as least Windows 10, version 1607, so here we check for availability before calling any methods
                Set BitmapImage3 = m_BitmapImage
                With BitmapImage3
                    If .IsAnimatedBitmap Then If .IsPlaying Then .Stop Else .Play
                End With
            End If
    End Select
End Sub

Private Sub Form_Load()
Dim ClientRect As RECTL, DesktopWindowXamlSourceNative As IDesktopWindowXamlSourceNative
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    If GetClientRect(picImage.hWnd, ClientRect) = APITRUE Then
        Set m_DesktopWindowXamlSource = NewObject("DesktopWindowXamlSource") ' The DesktopWindowXamlSource object allows hosting of XAML Islands on a regular Win32 hWnd
        Set DesktopWindowXamlSourceNative = m_DesktopWindowXamlSource
        With DesktopWindowXamlSourceNative
            If .AttachToWindow(picImage.hWnd) = S_OK Then ' Use the PictureBox as a container for this XAML Island
                m_hWndIsland = .WindowHandle
                If SetWindowPos(m_hWndIsland, HWND_TOP, 0, 0, ClientRect.Right, ClientRect.Bottom, SWP_SHOWWINDOW Or SWP_NOMOVE) = APITRUE Then ' The XAML Island window is hidden by default so here we resize and make it visible
                    Set m_BitmapImage = NewObject("BitmapImage") ' Create a new BitmapImage object (empty for now) to serve as the source for the Image control
                    With m_BitmapImage
                        .AddImageOpened NewEventHandler(Event_ImageOpened): .AddImageFailed NewEventHandler(Event_ImageFailed) ' Subscribe to the ImageOpened and ImageFailed events
                    End With
                    Set m_DesktopWindowXamlSource.Content = LayoutRoot(NewObject("Grid"), NewObject("Image")) ' Populate the content of the XAML Island with a Grid container and an Image control
                    ' Make a list of some random animated GIFs from the web:
                    mnuGIF(0).Tag = "https://www.vbforums.com/images/ieimages/2025/03/2.gif"
                    mnuGIF(1).Tag = "https://www.vbforums.com/images/ieimages/2025/03/3.gif"
                    mnuGIF(2).Tag = "https://media.githubusercontent.com/media/SixLabors/ImageSharp/main/tests/Images/Input/Gif/cheers.gif"
                    mnuGIF(3).Tag = "https://www.photofunky.net/output/image/a/2/f/0/a2f029/photofunky.gif"
                    mnuGIF(4).Tag = "https://i.pinimg.com/originals/05/cf/b9/05cfb96cdc8931cb75a5c6b12fae43b2.gif"
                    mnuGIF(5).Tag = "https://i.makeagif.com/media/7-27-2018/Rw8uJr.gif"
                End If
            End If
        End With
    End If
End Sub

Private Function LayoutRoot(Grid As IPanel, Image As IImage) As IUIElement
Dim UIElementCollection As IVector_IUIElement, UIElement As IUIElement, FrameworkElement As IFrameworkElement, lBackColor As Long
    ToolTipServiceStatics.SetToolTip Image, Box("Click to Stop/Start animation") ' Assign a tooltip for the Image control
    With Image
        .Stretch = Stretch_None ' Make sure images loaded in the Image control keep their original size
        Set .Source = m_BitmapImage ' Set the BitmapImage object (empty for now) as the source for this Image control
    End With
    Set UIElement = Image: UIElement.AddTapped NewEventHandler(Event_Tapped): Set FrameworkElement = Grid: FrameworkElement.AddLoaded NewEventHandler(Event_Loaded)
    OleTranslateColor BackColor, vbNullPtr, lBackColor: lBackColor = "&H" & Hex$(lBackColor) & "FF" ' Add transparent Alpha channel
    Set Grid.Background = SolidColorBrushFactory.CreateInstanceWithColor(lBackColor) ' Set the Grid's background color the same as the Form's
    Set UIElementCollection = Grid.Children: UIElementCollection.Append Image ' Add the Image control to the Grid's children collection
    Set LayoutRoot = Grid
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dispose m_DesktopWindowXamlSource ' Destroy the XAML Island gracefully
End Sub

Private Sub mnuGIF_Click(Index As Integer)
Dim i As Long
    If Not mnuGIF(Index).Checked Then
        For i = mnuGIF.LBound To mnuGIF.UBound: mnuGIF(i).Checked = i = Index: Next i
        picImage.Visible = False: Set m_BitmapImage.UriSource = UriFactory.CreateUri(StrRef(mnuGIF(Index).Tag)) ' Load a new GIF image asynchronously
    End If
End Sub

Private Sub picImage_Resize()
    With picImage ' Resize the Form to match the size of the PictureBox and its loaded image
        Width = .Width + (Width - ScaleWidth) + .Left * 2: Height = .Height + (Height - ScaleHeight) + .Top * 2
        SetWindowPos m_hWndIsland, HWND_TOP, 0, 0, .ScaleX(.Width, .ScaleMode, vbPixels), .ScaleY(.Height, .ScaleMode, vbPixels), SWP_SHOWWINDOW Or SWP_NOMOVE
        .Visible = True
    End With
End Sub

Private Function NewEventHandler(ByVal EventIdentifier As EVENT_IDENTIFIERS) As cEventsSink
    With New cEventsSink
        Set NewEventHandler = .This(Me, EventIdentifier)
    End With
End Function

