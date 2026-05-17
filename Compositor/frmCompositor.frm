VERSION 5.00
Begin VB.Form frmCompositor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compositor Demo"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompositor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShadow 
      Caption         =   "Button with DropShadow"
      Height          =   810
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmCompositor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_PI As Single = 3.14159265, PerspDepth As Single = 1 / 900, CubeSide As Single = 200, HalfSide As Single = CubeSide / 2

Private Compositor As ICompositor, DesktopWindowTarget As ICompositionTarget, RootVisual As IContainerVisual, SpinAnimation As IScalarKeyFrameAnimation, _
        PointLight As IPointLight, AmbientLight As IAmbientLight, ShadowOn As IVector3KeyFrameAnimation, ShadowOff As IVector3KeyFrameAnimation, ButtonShadow As IDropShadow

Private Sub cmdShadow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CompositionObject As ICompositionObject
    Set CompositionObject = ButtonShadow: CompositionObject.StartAnimation StrRef("Offset"), ShadowOff
End Sub

Private Sub cmdShadow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CompositionObject As ICompositionObject
    Set CompositionObject = ButtonShadow: CompositionObject.StartAnimation StrRef("Offset"), ShadowOn
End Sub

Private Sub Form_Load()
Dim Options As DispatcherQueueOptions, DispatcherQueueController As IDispatcherQueueController, CompositorDesktopInterop As ICompositorDesktopInterop, Cube As IContainerVisual, _
    Visual2 As IVisual2, Visual As IVisual, Compositor2 As ICompositor2, AmbientLight2 As IAmbientLight2, CompositionLight As ICompositionLight, PointLight2 As IPointLight2, CompositionObject As ICompositionObject
    If CreateDispatcherQueueController(LenB(Options), DQTAT_COM_STA, DQTAT_COM_STA, DispatcherQueueController) = S_OK Then
        Set Compositor = NewObject("Compositor"): Set CompositorDesktopInterop = Compositor
        Set DesktopWindowTarget = CompositorDesktopInterop.CreateDesktopWindowTarget(hWnd, APIFALSE)
        Set RootVisual = Compositor.CreateContainerVisual: Set Visual = RootVisual: Set Visual2 = RootVisual: Visual2.RelativeSizeAdjustmentLet 1, 1
        Set DesktopWindowTarget.Root = RootVisual
        Visual.TransformMatrixLet 1, 0, 0, 0, _
                                  0, 1, 0, 0, _
                                  0, 0, 1, PerspDepth, _
                                  0, 0, 0, 1
        Set Cube = Compositor.CreateContainerVisual: Set Visual = Cube
        With Visual
            .SizeLet CubeSide, CubeSide
            .OffsetLet ScaleX(ScaleWidth, ScaleMode, vbPixels) / 2 - HalfSide, ScaleY(ScaleHeight, ScaleMode, vbPixels) / 2 - HalfSide, 0
            .RotationAxisLet 1, 1, 0
        End With
        
        With ColorsStatics
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(0, 1, 0), 0, V3(0, 0, -HalfSide), .Red, .Orange)
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(0, 1, 0), m_PI, V3(0, 0, HalfSide), .Indigo, .Thistle)
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(0, 1, 0), -m_PI / 2, V3(HalfSide, 0, 0), .Goldenrod, .Yellow)
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(0, 1, 0), m_PI / 2, V3(-HalfSide, 0, 0), .Navy, .Cyan)
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(1, 0, 0), -m_PI / 2, V3(0, -HalfSide, 0), .DarkGreen, .Chartreuse)
            Cube.Children.InsertAtTop CreateCubeFace(Compositor, V3(1, 0, 0), m_PI / 2, V3(0, HalfSide, 0), .Peru, .Thistle)
        End With
        
        RootVisual.Children.InsertAtTop Cube: Set Compositor2 = Compositor
        
        Set AmbientLight = Compositor2.CreateAmbientLight: Set AmbientLight2 = AmbientLight: Set CompositionLight = AmbientLight
        AmbientLight.Color = ColorsStatics.White: AmbientLight2.Intensity = 0.2: CompositionLight.Targets.Add Cube
        
        Set PointLight = Compositor2.CreatePointLight: Set PointLight2 = PointLight: Set CompositionLight = PointLight
        With PointLight
            .Color = ColorsStatics.White: Set .CoordinateSpace = RootVisual: .OffsetLet CubeSide, CubeSide, 80
            .ConstantAttenuation = 1: .QuadraticAttenuation = 0.005
        End With
        PointLight2.Intensity = 2: CompositionLight.Targets.Add Cube
        
        Set SpinAnimation = CreateSpinAnimation(Compositor, 6000, 0, 2 * m_PI): Set CompositionObject = Cube
        CompositionObject.StartAnimation StrRef("RotationAngle"), SpinAnimation
        
        RootVisual.Children.InsertAtTop CreateDropShadow(Compositor, cmdShadow)
        Set ShadowOn = CreateOffsetAnimation(Compositor, 300, NewV3(0, 0), NewV3(5, 5))
        Set ShadowOff = CreateOffsetAnimation(Compositor, 300, NewV3(5, 5), NewV3(0, 0))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dispose AmbientLight: Dispose PointLight: Dispose SpinAnimation: Dispose RootVisual: Dispose DesktopWindowTarget: Dispose Compositor
End Sub

Private Function CreateCubeFace(Compositor As ICompositor, RotationAxis As Vector3, ByVal RotationAngle As Single, Offset As Vector3, ByVal C1 As ColorLong, ByVal C2 As ColorLong) As ISpriteVisual
Dim Visual As IVisual
    Set CreateCubeFace = Compositor.CreateSpriteVisual: Set Visual = CreateCubeFace
    With Visual
        .SizeLet CubeSide, CubeSide: .CenterPointLet HalfSide, HalfSide, 0: .RotationAngle = RotationAngle: .BackfaceVisibility = CompositionBackfaceVisibility_Hidden
    End With
    With RotationAxis: Visual.RotationAxisLet .X, .Y, .Z: End With
    With Offset: Visual.OffsetLet .X, .Y, .Z: End With
    Set CreateCubeFace.Brush = CreateLinearGradientBrush(Compositor, C1, C2)
End Function

Private Function CreateColorGradientStop(Compositor As ICompositor4, ByVal Offset As Single, ByVal Color As ColorLong) As ICompositionColorGradientStop
    Set CreateColorGradientStop = Compositor.CreateColorGradientStop: CreateColorGradientStop.Color = Color: CreateColorGradientStop.Offset = Offset
End Function

Private Function CreateDropShadow(Compositor As ICompositor, Control As Control) As ISpriteVisual2
Dim Compositor2 As ICompositor2, Visual As IVisual
    Set CreateDropShadow = Compositor.CreateSpriteVisual: Set Visual = CreateDropShadow
    Set Compositor2 = Compositor: Set ButtonShadow = Compositor2.CreateDropShadow
    With Visual
        .SizeLet ScaleX(Control.Width, ScaleMode, vbPixels), ScaleY(Control.Height, ScaleMode, vbPixels)
        .OffsetLet ScaleX(Control.Left, ScaleMode, vbPixels), ScaleY(Control.Top, ScaleMode, vbPixels), 0
    End With
    With ButtonShadow: .BlurRadius = 10: .OffsetLet 5, 5, 0: .Opacity = 0.75: End With
    Set CreateDropShadow.Shadow = ButtonShadow
End Function

Private Function CreateLinearGradientBrush(Compositor As ICompositor4, ParamArray GradientStops() As Variant) As ICompositionLinearGradientBrush
Dim i As Long, CompositionGradientBrush As ICompositionGradientBrush, ColorStops As IVector_ICompositionColorGradientStop
    Set CreateLinearGradientBrush = Compositor.CreateLinearGradientBrush: Set CompositionGradientBrush = CreateLinearGradientBrush
    With CreateLinearGradientBrush
        .StartPointLet 0, 0: .EndPointLet 1, 1
    End With
    Set ColorStops = CompositionGradientBrush.ColorStops
    With ColorStops
        For i = LBound(GradientStops) To UBound(GradientStops)
            .Append CreateColorGradientStop(Compositor, i / UBound(GradientStops), GradientStops(i))
        Next i
    End With
End Function

Private Function CreateOffsetAnimation(Compositor As ICompositor, ByVal Duration As TimeSpan, ParamArray Values() As Variant) As IVector3KeyFrameAnimation
Dim i As Long, KeyFrameAnimation As IKeyFrameAnimation, CompositionAnimation2 As ICompositionAnimation2, cV3 As cVector3
    Set CreateOffsetAnimation = Compositor.CreateVector3KeyFrameAnimation: Set KeyFrameAnimation = CreateOffsetAnimation: Set CompositionAnimation2 = CreateOffsetAnimation
    For i = LBound(Values) To UBound(Values)
        Set cV3 = Values(i): With cV3.V3: CreateOffsetAnimation.InsertKeyFrame i / UBound(Values), .X, .Y, .Z: End With
    Next
    KeyFrameAnimation.Duration = Duration: CompositionAnimation2.Target = StrRef("Offset")
End Function

Private Function CreateSpinAnimation(Compositor As ICompositor, ByVal Duration As TimeSpan, ParamArray Values() As Variant) As IScalarKeyFrameAnimation
Dim i As Long, LinearEasingFunction As ILinearEasingFunction, KeyFrameAnimation As IKeyFrameAnimation, CompositionAnimation2 As ICompositionAnimation2
    Set CreateSpinAnimation = Compositor.CreateScalarKeyFrameAnimation: Set LinearEasingFunction = Compositor.CreateLinearEasingFunction
    Set KeyFrameAnimation = CreateSpinAnimation: Set CompositionAnimation2 = CreateSpinAnimation
    For i = LBound(Values) To UBound(Values)
        CreateSpinAnimation.InsertKeyFrameWithEasingFunction i / UBound(Values), Values(i), LinearEasingFunction
    Next i
    With KeyFrameAnimation: .Duration = Duration: .IterationBehavior = AnimationIterationBehavior_Forever: End With
    CompositionAnimation2.Target = StrRef("RotationAngle")
End Function

Private Function V3(Optional ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single) As Vector3
    With V3: .X = X: .Y = Y: .Z = Z: End With
End Function

Private Function NewV3(Optional ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single) As cVector3
    With New cVector3
        Set NewV3 = .This(X, Y, Z)
    End With
End Function
