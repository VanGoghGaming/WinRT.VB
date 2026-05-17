VERSION 5.00
Begin VB.Form frmOCR 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCR Demo - Copy an image to the Clipboard for instant recognition!"
   ClientHeight    =   12180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOCR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12180
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOCR 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   9210
      Width           =   12000
   End
   Begin VB.PictureBox picOCR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   120
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements cEventsSink

Private WithEvents m_OcrEngine As cOCR, m_DesktopWindowXamlSource As IDesktopWindowXamlSource, hWndIsland As LongPtr, m_ScrollViewer As IContentControl
Attribute m_OcrEngine.VB_VarHelpID = -1

Private Sub AddNewBoundingRectangle(Canvas As IPanel, BoundingRectangle As IUIElement, Left As Double, Top As Double)
Dim UIElementCollection As IVector_IUIElement
    Set UIElementCollection = Canvas.Children
    If TypeOf BoundingRectangle Is IRectangle Then CanvasStatics.SetLeft BoundingRectangle, Left: CanvasStatics.SetTop BoundingRectangle, Top: UIElementCollection.Append BoundingRectangle
End Sub

Private Sub ClearBoundingRectangles(Canvas As IPanel)
Dim UIElementCollection As IVector_IUIElement, i As Long
    Set UIElementCollection = Canvas.Children
    With UIElementCollection
        For i = .Size - 1 To 0 Step -1
            If TypeOf .GetAt(i) Is IRectangle Then .RemoveAt i
        Next i
    End With
End Sub

Private Sub cEventsSink_EventHandler(ByVal EventIdentifier As EVENT_IDENTIFIERS, ByVal Sender As IInspectable, Optional ByVal PointerRoutedEventArgs As IPointerRoutedEventArgs, Optional ByVal RoutedEventArgs As IRoutedEventArgs, Optional ByVal ExceptionRoutedEventArgs As IExceptionRoutedEventArgs, Optional ByVal TappedRoutedEventArgs As ITappedRoutedEventArgs)
Dim cAwait As New cAwait, ScrollViewer As IScrollViewer, ScrollViewer2 As IScrollViewer2, ZoomFactor As Single
    Select Case EventIdentifier
        Case Event_ClipboardContentChanged
            With ClipboardStatics.GetContent
                If .Contains(StrRef(GetStr(StandardDataFormatsStatics.Bitmap))) Then
                    picOCR.Visible = False
                    If cAwait.Await(.GetBitmapAsync) = AsyncStatus_Completed Then m_OcrEngine.GetBitmapFromStream cAwait.GetResults
                End If
            End With
        Case Event_PointerWheelChanged
            If PointerRoutedEventArgs.KeyModifiers = VirtualKeyModifiers_Control Then
                Set ScrollViewer = Sender
                With ScrollViewer
                    ZoomFactor = .ZoomFactor
                    If PointerRoutedEventArgs.GetCurrentPoint(ScrollViewer).Properties.MouseWheelDelta < 0 Then
                        ZoomFactor = ZoomFactor + .MinZoomFactor
                    Else
                        ZoomFactor = ZoomFactor - .MinZoomFactor
                    End If
                    Select Case ZoomFactor
                        Case Is < .MinZoomFactor: ZoomFactor = .MinZoomFactor
                        Case Is > .MaxZoomFactor: ZoomFactor = .MaxZoomFactor
                    End Select
                End With
                Set ScrollViewer2 = ScrollViewer: ScrollViewer2.ChangeView Nothing, Nothing, Box(ZoomFactor)
            End If
    End Select
End Sub

Private Sub Form_Load()
Dim DesktopWindowXamlSourceNative As IDesktopWindowXamlSourceNative, ClientRect As RECTL
    If RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") <> APITRUE Then
        RegValStr(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Visual Basic\6.0", "AllowUnsafeObjectPassing") = APITRUE ' This registry key is needed so that VB6 classes can receive COM callbacks
    End If
    If GetClientRect(picOCR.hWnd, ClientRect) = APITRUE Then
        Set m_DesktopWindowXamlSource = NewObject("DesktopWindowXamlSource"): Set DesktopWindowXamlSourceNative = m_DesktopWindowXamlSource
        With DesktopWindowXamlSourceNative
            If .AttachToWindow(picOCR.hWnd) = S_OK Then
                hWndIsland = .WindowHandle
                If SetWindowPos(hWndIsland, HWND_TOP, 0, 0, ClientRect.Right, ClientRect.Bottom, SWP_SHOWWINDOW Or SWP_NOMOVE) = APITRUE Then
                    Set m_DesktopWindowXamlSource.Content = LayoutRoot(NewObject("ScrollViewer"), NewObject("Canvas"))
                    Set m_OcrEngine = New cOCR: ClipboardStatics.AddContentChanged NewEventHandler(Event_ClipboardContentChanged)
                End If
            End If
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dispose m_DesktopWindowXamlSource
End Sub

Private Function LayoutRoot(ScrollViewer As IScrollViewer, Canvas As IPanel) As IUIElement
    ToolTipServiceStatics.SetToolTip ScrollViewer, Box("Ctrl+MouseWheel to Zoom In/Out")
    With ScrollViewer
        .HorizontalScrollBarVisibility = ScrollBarVisibility_Auto: .VerticalScrollBarVisibility = ScrollBarVisibility_Auto: .ZoomMode = ZoomMode_Enabled
    End With
    Set m_ScrollViewer = ScrollViewer: Set m_ScrollViewer.Content = Canvas: Set Canvas.Background = NewObject("ImageBrush")
    Set LayoutRoot = ScrollViewer ': LayoutRoot.AddPointerWheelChanged NewEventHandler(Event_PointerWheelChanged)
End Function

Private Sub m_OcrEngine_RecognitionResults(OcrResult As IOcrResult, SoftwareBitmapSource As ISoftwareBitmapSource, ByVal PixelWidth As Long, ByVal PixelHeight As Long)
Dim BoundingRectangle As IFrameworkElement, RotateTransform As IRotateTransform, i As Long, j As Long, CenterX As Single, CenterY As Single, AngleRadians As Double, RotatedX As Double, RotatedY As Double
    ClearBoundingRectangles m_ScrollViewer.Content: CenterX = PixelWidth / 2: CenterY = PixelHeight / 2: Set RotateTransform = NewObject("RotateTransform")
    With OcrResult
        txtOCR = GetStr(.Text): If Not .TextAngle Is Nothing Then RotateTransform.Angle = .TextAngle.Value: AngleRadians = .TextAngle.Value * Atn(1) / 45
        With .Lines
            For i = 0 To .Size - 1
                With .GetAt(i).Words
                    For j = 0 To .Size - 1
                        With .GetAt(j)
                            Set BoundingRectangle = NewObject("Rectangle"): ToolTipServiceStatics.SetToolTip BoundingRectangle, Box(GetStr(.Text))
                            With .BoundingRect
                                RotatedX = (.X - CenterX) * Cos(AngleRadians) - (.Y - CenterY) * Sin(AngleRadians) + CenterX
                                RotatedY = (.X - CenterX) * Sin(AngleRadians) + (.Y - CenterY) * Cos(AngleRadians) + CenterY
                                BoundingRectangle.Width = .Width: BoundingRectangle.Height = .Height
                            End With
                        End With
                        AddNewBoundingRectangle m_ScrollViewer.Content, SetRectangleProperties(BoundingRectangle, RotateTransform), RotatedX, RotatedY
                    Next j
                End With
            Next i
        End With
    End With
    SetCanvasImageBrushSource m_ScrollViewer.Content, SoftwareBitmapSource, PixelWidth, PixelHeight: picOCR.Visible = True: WindowState = vbNormal: Me.SetFocus
End Sub

Private Function NewEventHandler(ByVal EventIdentifier As EVENT_IDENTIFIERS) As cEventsSink
    With New cEventsSink
        Set NewEventHandler = .This(Me, EventIdentifier)
    End With
End Function

Private Sub SetCanvasImageBrushSource(Canvas As IPanel, SoftwareBitmapSource As ISoftwareBitmapSource, ByVal PixelWidth As Long, ByVal PixelHeight As Long)
Dim FrameworkElement As IFrameworkElement, ImageBrush As IImageBrush
    Set FrameworkElement = Canvas: FrameworkElement.Width = PixelWidth: FrameworkElement.Height = PixelHeight
    Set ImageBrush = Canvas.Background: Set ImageBrush.ImageSource = SoftwareBitmapSource
End Sub

Private Function SetRectangleProperties(Shape As IShape, RotateTransform As IRotateTransform) As IUIElement
    Set Shape.Stroke = SolidColorBrushFactory.CreateInstanceWithColor(ColorsStatics.Red)
    Set Shape.Fill = SolidColorBrushFactory.CreateInstanceWithColor(ColorsStatics.Transparent)
    Set SetRectangleProperties = Shape: Set SetRectangleProperties.RenderTransform = RotateTransform
End Function
