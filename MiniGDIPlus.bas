Attribute VB_Name = "MiniGDIPlus"
'MiniGDIPlus
'制作：马云爱逛京东
'版本：0.1beta

Option Explicit

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuff As GdiplusStartupInput, Optional ByVal OutputBuff As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDeviceContext As Long, ByRef hGraphics As Long) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal SmoothingMethod As SmoothingMode) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal pFileName As Long, hImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef Height As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePointRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As GpUnit, ByRef hPen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal Brush As Long, ByVal Width As Single, ByVal unit As GpUnit, ByRef hPen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GpStatus
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, hBrush As Long) As GpStatus
Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (Point1 As PointL, Point2 As PointL, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, ByRef hLineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (rect As RectL, ByVal color1 As Long, ByVal color2 As Long, ByVal Mode As LinearGradientMode, ByVal WrapMd As WrapMode, ByVal hLineGradient As Long) As GpStatus
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As GpStatus
Public Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GpStatus
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus
Public Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal AddingPath As Long, ByVal bConnect As Long) As GpStatus
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As Long, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As FillMode, path As Long) As GpStatus
Public Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal path As Long, ByVal str As Long, ByVal Length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RectL, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCreateLineBrush Lib "gdiplus" (Point1 As PointF, Point2 As PointF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As Long, ByVal unit As GpUnit, createdfont As Long) As GpStatus
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RectF, ByVal StringFormat As Long, boundingBox As RectF, codepointsFitted As Long, linesFilled As Long) As GpStatus
Public Declare Function GdipCreateRegionRect Lib "gdiplus" (rect As RectF, region As Long) As GpStatus
Public Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal Pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As GpStatus
Public Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As GpStatus
Public Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As GpStatus

Private Type GdiplusStartupInput
    GdiplusVersion As Long                                                      '版本
    DebugEventCallback As Long                                                  '除错事件回调
    SuppressBackgroundThread As Long                                            '抑制背景线程
    SuppressExternalCodecs As Long                                              '抑制外部编解码器
End Type

Public Type tImage
    imgHandle As Long                                                           '图片的句柄。
    imgIndex As Long                                                            '图片的索引，可空。
    imgName As String                                                           '图片的名称，通常来自于图片指向的文件名主名。
End Type

Public Type PointL
    X As Long
    Y As Long
End Type

Public Type PointF
    X As Single
    Y As Single
End Type
 
Public Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
 
Public Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

Public Enum StringTrimming
    StringTrimmingNone = 0
    StringTrimmingCharacter = 1
    StringTrimmingWord = 2
    StringTrimmingEllipsisCharacter = 3
    StringTrimmingEllipsisWord = 4
    StringTrimmingEllipsisPath = 5
End Enum

Public Enum BasicShape
    Rectangle = 0
    Ellipse = 1
End Enum

Public Enum ShapeStyle
    BorderOnly = 0
    FillOnly = 1
    Both = 2
End Enum

Public Enum eImageSuffix
    sfxJpg = &H0
    sfxBmp = &H1
    sfxPng = &H2
    sfxGif = &H3
End Enum

Public Enum FillMode
    FillModeAlternate
    FillModeWinding
End Enum
 
Public Enum WrapMode
    WrapModeTile
    WrapModeTileFlipX
    WrapModeTileFlipY
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

Public Enum CombineMode
    CombineModeReplace                                                          ' 0
    CombineModeIntersect                                                        ' 1
    CombineModeUnion                                                            ' 2
    CombineModeXor                                                              ' 3
    CombineModeExclude                                                          ' 4
    CombineModeComplement                                                       ' 5
End Enum

Public Enum LinearGradientMode
    LinearGradientModeHorizontal
    LinearGradientModeVertical
    LinearGradientModeForwardDiagonal
    LinearGradientModeBackwardDiagonal
End Enum

Public Enum GpUnit
    UnitWorld                                                                   '全局
    UnitDisplay                                                                 '显示
    UnitPixel                                                                   '像素
    UnitPoint                                                                   '点
    UnitInch                                                                    '英寸
    UnitDocument                                                                '文档
    UnitMillimeter                                                              '毫米
End Enum
 
Public Enum GpStatus
    Ok = 0                                                                      '没问题
    GenericError = 1                                                            '通用错误
    InvalidParameter = 2                                                        '无效参数
    OutOfMemory = 3                                                             '内存溢出
    ObjectBusy = 4                                                              '对象忙
    InsufficientBuffer = 5                                                      '缓冲区容量不足
    NotImplemented = 6                                                          '未执行
    Win32Error = 7                                                              '引发Win32的错误
    WrongState = 8                                                              '状态不正确
    Aborted = 9                                                                 '被终止
    FileNotFound = 10                                                           '文件未找到
    ValueOverflow = 11                                                          '值溢出
    AccessDenied = 12                                                           '访问被拒绝
    UnknownImageFormat = 13                                                     '未知的图像格式
    FontFamilyNotFound = 14                                                     '未找到字体族
    FontStyleNotFound = 15                                                      '未找到字体样式
    NotTrueTypeFont = 16                                                        '不是TrueType字体
    UnsupportedGdiplusVersion = 17                                              '不支持的GDIPlus版本
    GdiplusNotInitialized = 18                                                  'GDIPlus未被初始化
    PropertyNotFound = 19                                                       '属性未找到
    PropertyNotSupported = 20                                                   '属性不支持
End Enum

Public Enum SmoothingMode
    SmoothingModeInvalid = -1                                                   '无效/不采用
    SmoothingModeDefault = 0                                                    '默认
    SmoothingModeHighSpeed = 1                                                  '高速
    SmoothingModeHighQuality = 2                                                '高质量
    SmoothingModeNone = 3                                                       '不采用
    SmoothingModeAntiAlias = 4                                                  '抗锯齿
End Enum

Private mToken As Long

'InitializeGDIPlus：初始化GDIPlus，一般放于启动窗口的Form_Load事件，或Sub Main中
Public Sub InitializeGDIPlus()
    If mToken <> 0 Then Exit Sub
    Dim Retn As GpStatus, uInput As GdiplusStartupInput
    uInput.GdiplusVersion = 1
    Retn = GdiplusStartup(mToken, uInput)
    If Retn <> 0 Then Debug.Print "GDIPlus未能初始化。"
End Sub

'TerminateGDIPlus：终止化GDIPlus，一般放于启动窗口的Form_Unload事件中
Public Sub TerminateGDIPlus()
    If mToken = 0 Then Exit Sub
    GdiplusShutdown mToken
    mToken = 0
End Sub

'DrawImageFromFile：从文件中创建图像对象，并输出在容器上
'Container：指定的容器，需要有hDC，可以是窗体或者图片框
'Path：图像绝对路径
'Left/Top：图像左上角相对于容器的位置（像素）
'Zoom：缩放比例，小于1时为缩小，大于1时为放大
Public Sub DrawImageFromFile(ByVal Container As Object, ByVal path As String, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, Optional ByVal Zoom As Double = 1#)
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    If Zoom <= 0 Then
        MsgBox "出现参数错误：Zoom的值应大于0。", vbCritical, "错误"
        Exit Sub
    End If
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    If Container.AutoRedraw = False Then Container.AutoRedraw = True
    Dim tGraphic As Long, tImage As Long, imgWidth As Long, imgHeight As Long
    Call GdipCreateFromHDC(Container.hDC, tGraphic)
    Call GdipSetSmoothingMode(tGraphic, SmoothingModeAntiAlias)
    Call GdipLoadImageFromFile(StrPtr(path), tImage)
    Call GdipGetImageWidth(tImage, imgWidth)
    Call GdipGetImageHeight(tImage, imgHeight)
    Call GdipDrawImageRectI(tGraphic, tImage, Left, Top, imgWidth * Zoom, imgHeight * Zoom)
    Call GdipDisposeImage(tImage)
    Call GdipDeleteGraphics(tGraphic)
    Container.Refresh
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageOnPictureBoxFromFile：在图片框上绘制采用文件绝对路径指定的图像
'AutoSize：是否使图片框适应图像尺寸。为True时自动调节图片框尺寸以适应图像；为False时自动调节图像以适应图像框
Public Sub DrawImageOnPictureBoxFromFile(ByRef DestPictureBox As PictureBox, ByVal path As String, Optional ByVal AutoSize As Boolean = True)
    On Error GoTo ExitSub
    Dim tGraphic As Long, tImage As Long, imgWidth As Long, imgHeight As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Call GdipLoadImageFromFile(StrPtr(path), tImage)
    Call GdipGetImageWidth(tImage, imgWidth)
    Call GdipGetImageHeight(tImage, imgHeight)
    If AutoSize Then
        DestPictureBox.Width = imgWidth
        DestPictureBox.Height = imgHeight
    End If
    Call GdipCreateFromHDC(DestPictureBox.hDC, tGraphic)
    Call GdipSetSmoothingMode(tGraphic, SmoothingModeAntiAlias)
    If AutoSize Then
        Call GdipDrawImageRectI(tGraphic, tImage, 0, 0, imgWidth, imgHeight)
    Else
        Call GdipDrawImageRect(tGraphic, tImage, 0, 0, DestPictureBox.ScaleWidth, DestPictureBox.ScaleHeight)
    End If
    Call GdipDisposeImage(tImage)
    Call GdipDeleteGraphics(tGraphic)
    DestPictureBox.Refresh
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageFromHandle：从句柄中创建图像对象，并输出在容器上
'Container：指定的容器，需要有hDC，可以是窗体或者图片框
'Handle：图像句柄
'Left/Top：图像左上角相对于容器的位置（像素）
'Zoom：缩放比例，小于1时为缩小，大于1时为放大
Public Sub DrawImageFromHandle(ByVal Container As Object, ByVal Handle As Long, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, Optional ByVal Zoom As Single = 1#)
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    If Zoom <= 0 Then
        MsgBox "出现参数错误：Zoom的值应大于0。", vbCritical, "GDIPlus2"
        Exit Sub
    End If
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    If Container.AutoRedraw = False Then Container.AutoRedraw = True
    Dim Graphic As Long, Image As Long, imgWidth As Long, imgHeight As Long
    Call GdipCreateFromHDC(Container.hDC, Graphic)
    Call GdipSetSmoothingMode(Graphic, SmoothingModeAntiAlias)
    Image = Handle
    Call GdipGetImageWidth(Image, imgWidth)
    Call GdipGetImageHeight(Image, imgHeight)
    Call GdipDrawImageRect(Graphic, Image, Left, Top, imgWidth * Zoom, imgHeight * Zoom)
    Call GdipDeleteGraphics(Graphic)
    Container.Refresh
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'【注】此函数有bug：可能导致卡死。
'DrawImageFromHandleRectToRect：从句柄中创建图像对象的指定区域，并输出在容器上。
'dstLeft/dstTop/dstWidth/dstHeight：目标的位置和尺寸。
'srcLeft/srcTop/srcWidth/srcHeight：源的位置和尺寸。
Public Sub DrawImageFromHandleRectToRect(ByVal Container As Object, ByVal Handle As Long, ByVal dstLeft As Long, ByVal dstTop As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, _
    ByVal srcLeft As Long, ByVal srcTop As Long, ByVal srcWidth As Long, ByVal srcHeight As Long)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True
    InitializeGDIPlus
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    Dim Graphic As Long
    If Container.AutoRedraw = False Then Container.AutoRedraw = True
    GdipCreateFromHDC Container.hDC, Graphic
    If mToken = 0 Then GoTo ExitSub
    GdipDrawImageRectRectI Graphic, Handle, dstLeft, dstTop, dstWidth, dstHeight, srcLeft, srcTop, srcWidth, srcHeight, UnitPixel
    Container.Refresh
    Call GdipDeleteGraphics(Graphic)
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'【注】此函数有bug：可能导致卡死。
Public Sub DrawImageFromHandleRectToRect2(ByVal Container As Object, ByVal Handle As Long, ByRef DestRect As RectL, ByRef SourceRect As RectL)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True
    InitializeGDIPlus
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    Dim Graphic As Long
    If Container.AutoRedraw = False Then Container.AutoRedraw = True
    GdipCreateFromHDC Container.hDC, Graphic
    If mToken = 0 Then GoTo ExitSub
    GdipDrawImageRectRectI Graphic, Handle, DestRect.Left, DestRect.Top, DestRect.Right, DestRect.Bottom, SourceRect.Left, SourceRect.Top, SourceRect.Right, SourceRect.Bottom, UnitPixel
    Container.Refresh
    Call GdipDeleteGraphics(Graphic)
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageOnPictureBoxFromHandle：在图片框上绘制采用图像句柄引用的图像
'AutoSize：是否使图片框适应图像尺寸。为True时自动调节图片框尺寸以适应图像；为False时自动调节图像以适应图像框
Public Sub DrawImageOnPictureBoxFromHandle(ByRef DestPictureBox As PictureBox, ByVal Handle As Long, Optional ByVal AutoSize As Boolean = True)
    On Error GoTo ExitSub
    Dim tGraphic As Long, tImage As Long, imgWidth As Long, imgHeight As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    tImage = Handle
    Call GdipGetImageWidth(tImage, imgWidth)
    Call GdipGetImageHeight(tImage, imgHeight)
    If AutoSize Then
        DestPictureBox.Width = imgWidth
        DestPictureBox.Height = imgHeight
    End If
    Call GdipCreateFromHDC(DestPictureBox.hDC, tGraphic)
    Call GdipSetSmoothingMode(tGraphic, SmoothingModeAntiAlias)
    If AutoSize Then
        Call GdipDrawImageRectI(tGraphic, tImage, 0, 0, imgWidth, imgHeight)
    Else
        Call GdipDrawImageRect(tGraphic, tImage, 0, 0, DestPictureBox.ScaleWidth, DestPictureBox.ScaleHeight)
    End If
    Call GdipDisposeImage(tImage)
    Call GdipDeleteGraphics(tGraphic)
    DestPictureBox.Refresh
    If tFlag Then TerminateGDIPlus
    Exit Sub
ExitSub:
    MsgBox "出现绘制错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'LoadImageHandleFromFile：从图像文件中生成一个图像句柄
Public Function LoadImageHandleFromFile(ByVal path As String) As Long
    Dim Image As Long
    On Error GoTo ExitSub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    If Dir(path) = "" Then
        MsgBox "文件 " & path & " 不存在。", vbCritical, "错误"
        Exit Function
    End If
    Call GdipLoadImageFromFile(StrPtr(path), Image)
    LoadImageHandleFromFile = Image
    If tFlag Then TerminateGDIPlus
    Exit Function
ExitSub:
    MsgBox "出现未知错误。", vbCritical, "错误"
    If tFlag Then TerminateGDIPlus
End Function

'DelImageHandle：从内存中删除一个图像句柄
Public Sub DelImageHandle(ByRef Handle As Long)
    If Handle = 0 Then Exit Sub
    On Error Resume Next
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Call GdipDisposeImage(Handle)
    If tFlag Then TerminateGDIPlus
End Sub

'LoadImageFromFolderToImageArray：将一个指定文件夹的所有图片（指定同一个后缀）加载到图片数组里。
'ImageArray：要添加到的图片数组。
'FolderPath：文件夹路径。
'Suffix：文件名后缀。
Public Sub LoadImageFromFolderToImageArray(ByRef ImageArray() As tImage, ByVal FolderPath As String, ByVal Suffix As eImageSuffix)
    Dim FSO As Object, FolderObj As Object, FileObj As Object, nPath As String, nCount As Long, t As Boolean, tHandle As Long, tName As String
    nPath = FolderPath
    If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(nPath) Then
        Set FolderObj = FSO.GetFolder(nPath)
        On Error GoTo InitArray
        nCount = UBound(ImageArray)
InitArray:
        If Err.Number = 9 Then nCount = -1
        For Each FileObj In FolderObj.Files
            t = False
            tName = FileObj.name
            Select Case Suffix
            Case sfxBmp
                If GetFileSuffix(tName) = "bmp" Then t = True
            Case sfxPng
                If GetFileSuffix(tName) = "png" Then t = True
            Case sfxGif
                If GetFileSuffix(tName) = "gif" Then t = True
            Case sfxJpg
                If GetFileSuffix(tName) = "jpg" Or GetFileSuffix(tName) = "jpeg" Then t = True
            End Select
            If t Then
                nCount = nCount + 1
                ReDim Preserve ImageArray(nCount) As tImage
                tHandle = LoadImageHandleFromFile(nPath & tName)
                If tHandle <> 0 Then ImageArray(nCount).imgHandle = tHandle
                ImageArray(nCount).imgIndex = nCount
                ImageArray(nCount).imgName = Left(tName, Len(tName) - Len(GetFileSuffix(tName)))
                If Right(ImageArray(nCount).imgName, 1) = "." Then ImageArray(nCount).imgName = Left(ImageArray(nCount).imgName, Len(ImageArray(nCount).imgName) - 1)
            End If
        Next FileObj
    Else
        MsgBox "文件夹不存在。", vbCritical, "错误"
    End If
    Set FSO = Nothing: Set FolderObj = Nothing: Set FileObj = Nothing
End Sub

'LoadImageFromFolderToHandleArray：将一个指定文件夹的所有图片加载到一个数组里。
'HandleArray：要添加到的图像句柄数组。
'FolderPath：文件夹路径。
Public Sub LoadImageFromFolderToHandleArray(ByRef HandleArray() As Long, ByVal FolderPath As String)
    Dim FSO As Object, FolderObj As Object, FileObj As Object, nPath As String, nCount As Long
    nPath = FolderPath
    If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(nPath) Then
        Set FolderObj = FSO.GetFolder(nPath)
        On Error GoTo InitArray
        nCount = UBound(HandleArray)
InitArray:
        If Err.Number = 9 Then nCount = -1
        For Each FileObj In FolderObj.Files
            If GetFileSuffix(FileObj.name) = "png" Or GetFileSuffix(FileObj.name) = "gif" Or GetFileSuffix(FileObj.name) = "bmp" Or GetFileSuffix(FileObj.name) = "jpg" Then
                nCount = nCount + 1
                ReDim Preserve HandleArray(nCount) As Long
                HandleArray(nCount) = LoadImageHandleFromFile(nPath & FileObj.name)
            End If
        Next FileObj
    Else
        MsgBox "文件夹不存在。", vbCritical, "错误"
    End If
    Set FSO = Nothing: Set FolderObj = Nothing: Set FileObj = Nothing
End Sub

'LoadImageFromFileStringSuffixToHandleArray：将指定文件夹下包含文件名字符串的所有特定图片文件加载到一个数组里。
'HandleArray：要添加到的图像句柄数组。
'FolderPath：文件夹路径。
'FileString：多个文件名和分隔符组成的字符串（格式：文件名 分隔符 文件名 分隔符 文件名……）。
'Delimiter：分隔符，缺省为半角逗号。不可以是空串（""）。
'Suffix：文件名后缀。当指定了文件名后缀时，FileString的所有文件都不用再添加文件名后缀；否则需要单独为每一个文件指定后缀。
Public Sub LoadImageFromFileStringSuffixToHandleArray(ByRef HandleArray() As Long, ByVal FolderPath As String, ByVal FileString As String, _
    Optional ByVal Delimiter As String = ",", Optional ByVal Suffix As String = "")
    If Delimiter = "" Then MsgBox "分隔符（Delimiter）无效。", vbCritical, "错误": Exit Sub
    If Suffix <> "" And LCase(Suffix) <> "png" And LCase(Suffix) <> "gif" And LCase(Suffix) <> "bmp" And LCase(Suffix) <> "jpg" _
        Then MsgBox "指定的文件后缀无效。" & vbCrLf & "文件后缀可以是png、gif、bmp、jpg其中一种或省略。", vbCritical, "错误"
        Dim FileArray As Variant, i As Long, nCount As Long, nPath As String
        nPath = FolderPath
        If Right(nPath, 1) <> "\" Then nPath = nPath & "\"
        FileArray = Split(FileString, Delimiter)
        If Suffix <> "" Then
            For i = 0 To UBound(FileArray)
                FileArray(i) = FileArray(i) & "." & Suffix
            Next i
        End If
        On Error GoTo InitArray
        nCount = UBound(HandleArray)
InitArray:
        If Err.Number = 9 Then nCount = -1
        For i = 0 To UBound(FileArray)
            nCount = nCount + 1
            ReDim Preserve HandleArray(nCount) As Long
            HandleArray(nCount) = LoadImageHandleFromFile(nPath & FileArray(i))
            Debug.Print "添加文件：" & nPath & FileArray(i)
        Next i
End Sub

'ReleaseImgHandleArray：释放一个图像句柄数组里所有的图像句柄。此函数通常用于GDIPlus被终结之前。
'HandleArray：图像句柄数组。
Public Sub ReleaseImgHandleArray(ByRef HandleArray() As Long)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim i As Long
    For i = 0 To UBound(HandleArray)
        GdipDisposeImage HandleArray(i)
    Next i
    If tFlag Then TerminateGDIPlus
End Sub

'ReleaseImgImageArray：释放一个图片数组里所有的图像句柄。此函数通常用于GDIPlus被终结之前。
'ImageArray：图片数组。
Public Sub ReleaseImgImageArray(ByRef ImageArray() As tImage)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim i As Long
    For i = 0 To UBound(ImageArray)
        GdipDisposeImage ImageArray(i).imgHandle
    Next i
    On Error Resume Next
    Erase ImageArray
    If tFlag Then TerminateGDIPlus
End Sub

'GetFileSuffix：获得文件的后缀名。
'FilePath：文件路径。
Public Function GetFileSuffix(ByVal FilePath As String) As String
    Dim nPath As String
    nPath = FilePath
    GetFileSuffix = LCase(Right(nPath, Len(nPath) - InStrRev(nPath, ".")))
End Function

'CombinePathFromPathArray：将一个路径数组中的所有路径合并到指定的目标路径中，合并方式为或（并）运算。目标路径需要先被创建。
'PathArray：路径句柄数组。
'TargetPath：指定的目标路径。
Public Sub CombinePathFromPathArray(ByRef TargetPath As Long, ByRef PathArray() As Long)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim i As Long
    For i = 0 To UBound(PathArray)
        GdipAddPathPath TargetPath, PathArray(i), 1
    Next i
    If tFlag Then TerminateGDIPlus
End Sub

'DrawPath：填充和描边指定路径。
'ContainerHDC：容器的设备场景（Device Context）句柄。只允许使用窗体或图片框的hDC。
'Path：绘图路径。
'FillColor：填充色。
'BorderColor：描边色。
'ShapeMode：填充和描边样式。
'BorderWidth：描边粗细。
Public Sub DrawPath(ByRef ContainerHDC As Long, ByVal path As Long, ByVal FillColor As Long, ByVal BorderColor As Long, ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If BorderWidth <= 0 Then MsgBox "指定的描边粗细BorderWidth无效。该参数值必须大于0。", vbCritical, "错误": Exit Sub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim tGraphics As Long, tPen As Long, tBrush As Long
    GdipCreateFromHDC ContainerHDC, tGraphics
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateSolidFill FillColor, tBrush
    Select Case ShapeMode
    Case BorderOnly
        GdipDrawPath tGraphics, tPen, path
    Case FillOnly
        GdipFillPath tGraphics, tBrush, path
    Case Both
        GdipFillPath tGraphics, tBrush, path
        GdipDrawPath tGraphics, tPen, path
    End Select
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'DrawPathWithGradientColors：渐变填充和描边指定路径。
'ContainerHDC：容器的设备场景（Device Context）句柄。只允许使用窗体或图片框的hDC。
'Path：绘图路径。
'FillColor1(2)：填充色1(2)。
'BorderColor：描边色。
'Point1(2)：起(终)点。
'ShapeMode：填充和描边样式。
'BorderWidth：描边粗细。
Public Sub DrawPathWithGradientColors(ByRef ContainerHDC As Long, ByVal path As Long, ByVal FillColor1 As Long, ByVal FillColor2 As Long, ByVal BorderColor As Long, ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If BorderWidth <= 0 Then MsgBox "指定的描边粗细BorderWidth无效。该参数值必须大于0。", vbCritical, "错误"
    Dim tGraphics As Long, tPen As Long, tBrush As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateLineBrushI Point1, Point2, FillColor1, FillColor2, WrapModeTile, tBrush
    Select Case ShapeMode
    Case BorderOnly
        GdipDrawPath tGraphics, tPen, path
    Case FillOnly
        GdipFillPath tGraphics, tBrush, path
    Case Both
        GdipFillPath tGraphics, tBrush, path
        GdipDrawPath tGraphics, tPen, path
    End Select
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'DrawImageByName：根据图片名称绘制图片。
'Container：容器。只允许使用窗体或图片框。
'ImageArray：指定的图片数组。
'imgName：图片的名称，通常为图片的文件名称。
'Left/Top：绘制图片的左边距/上边距。
Public Sub DrawImageByName(ByRef Container As Object, ByRef ImageArray() As tImage, ByVal imgName As String, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0)
    Dim i As Long, pHandle As Long, pName As String
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    pName = imgName
    pHandle = 0
    For i = 0 To UBound(ImageArray)
        If ImageArray(i).imgName = pName Then
            pHandle = ImageArray(i).imgHandle
            Exit For
        End If
    Next i
    If pHandle = 0 Then                                                         '找不到对应的图片
        Debug.Print "DrawImageByName：找不到对应的图片。"
        Exit Sub
    End If
    DrawImageFromHandle Container, pHandle, Left, Top, 1
    If tFlag Then TerminateGDIPlus
End Sub

'DrawSimpleString：绘制简单字符串
'ContainerHDC：容器的设备场景（Device Context）句柄。只允许使用窗体或图片框的hDC。
'StringText：字符串文本内容。
'StringFontFamily：字体。
'StringSize：字号。
'StringStyle：字型
'FillColor/BorderColor：填充色/描边色。采用ARGB。
'BorderWidth：描边宽度。
'DrawBorder：是否描边。
Public Sub DrawSimpleString(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor As Long, Optional ByVal BorderColor As Long = 0, _
    Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       '根据指定的字体系列创建tFontFamily对象
    GdipCreateStringFormat 0, 0, tStringFormat                                  '基于字符串格式标志和语言创建tStringFormat对象
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 '设置此tStringFormat对象相对于布局矩形的原点的字符对齐，使用布局矩形来定位显示的字符串
    Dim tPen As Long, tBrush As Long
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateSolidFill FillColor, tBrush
    With tRectLayout
        .Left = Left
        .Top = Top
        .Bottom = Top + Height
        .Right = Left + Width
    End With
    GdipCreatePath FillModeAlternate, tStringPath                               '创建字符串路径
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '添加字符串路径
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '描边
    GdipFillPath tGraphics, tBrush, tStringPath                                 '填充路径
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    If tFlag Then TerminateGDIPlus
End Sub

'DrawGradientString：绘制渐变字符串（渐变方向：左→右）
Public Sub DrawGradientString(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor1 As Long, ByVal FillColor2 As Long, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' 根据指定的字体系列创建tFontFamily对象
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' 基于字符串格式标志和语言创建tStringFormat对象
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' 设置此tStringFormat对象相对于布局矩形的原点的字符对齐，使用布局矩形来定位显示的字符串
    Dim tPen As Long, tBrush As Long
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    With tRectLayout
        .Left = Left
        .Top = Top
        .Bottom = Top + Height
        .Right = Left + Width
    End With
    Dim tPoint1 As PointF, tPoint2 As PointF
    tPoint1.X = tRectLayout.Left: tPoint1.Y = (tRectLayout.Top + tRectLayout.Bottom) / 2
    tPoint2.X = tRectLayout.Right: tPoint2.Y = (tRectLayout.Top + tRectLayout.Bottom) / 2
    GdipCreateLineBrush tPoint1, tPoint2, FillColor1, FillColor2, WrapModeTileFlipX, tBrush
    GdipCreatePath FillModeAlternate, tStringPath                               '创建字符串路径
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '添加字符串路径
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '描边
    GdipFillPath tGraphics, tBrush, tStringPath                                 '填充路径
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeletePath tStringPath
    GdipDeleteBrush tBrush
    GdipDeletePen tPen
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    If tFlag Then TerminateGDIPlus
End Sub

'【注】此函数有bug：量测字符串布局矩形tRectF宽度有错误，但遮罩百分比已做数值补正。
'DrawDoubleColorsStringWithMaskByStringPercentage：以字符串内容百分比形式绘制带遮罩的字符串
'MaskColor：遮罩颜色。采用ARGB。
'StringPercentage：遮罩百分比。介于0-100之间。
Public Sub DrawDoubleColorsStringWithMaskByStringPercentage(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, _
    ByVal FillColor As Long, ByVal MaskColor As Long, ByVal StringPercentage As Single, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    If StringPercentage < 0 Or StringPercentage > 100 Then MsgBox "参数错误！StringPercentage的值介于[0,100]。", vbCritical, "参数错误": Exit Sub
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tFont As Long
    Dim tGraphics As Long
    Dim tCodePointsFitted As Long, tLinesFilled As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' 根据指定的字体系列创建tFontFamily对象
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' 基于字符串格式标志和语言创建tStringFormat对象
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' 设置此tStringFormat对象相对于布局矩形的原点的字符对齐，使用布局矩形来定位显示的字符串
    GdipCreateFont tFontFamily, StringSize, StringStyle, UnitPixel, tFont       '创建tFont对象
    Dim tPen As Long, tBrush1 As Long, tBrush2 As Long, tRect As Long, fRectF As RectF, tRectL As RectL, tRectF As RectF, tMask As Long, mRectF As RectF, mRect As Long
    Dim MaskWidth As Single                                                     '遮罩宽
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen                    '创建画笔
    GdipSetStringFormatTrimming tStringFormat, StringTrimmingEllipsisCharacter  '对字符串格式进行修整获得等宽字符（中文）的矩形fRectL
    With fRectF
        .Left = Left
        .Top = Top
        .Right = .Left + StringSize * (Len(StringText) + 0.5)
        .Bottom = .Top + StringSize
    End With
    GdipMeasureString tGraphics, StrPtr(StringText), Len(StringText) + 1, tFont, fRectF, tStringFormat, tRectF, tCodePointsFitted, tLinesFilled
    'tRectL
    With tRectL
        .Left = tRectF.Left
        .Top = tRectF.Top
        .Bottom = tRectF.Bottom
        .Right = tRectF.Right
    End With
    'mRectF
    With mRectF
        .Left = tRectF.Left
        .Top = tRectF.Top
        .Right = (tRectF.Right - StringSize * 0.5) * StringPercentage
        .Bottom = tRectF.Bottom
    End With
    GdipCreateRegionRect mRectF, mRect                                          '创建矩形区域
    GdipCreateSolidFill FillColor, tBrush1                                      '填充画刷
    GdipCreatePath FillModeAlternate, tStringPath                               '创建字符串路径
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectL, tStringFormat '添加字符串路径
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '描边
    GdipFillPath tGraphics, tBrush1, tStringPath                                '填充底层路径
    GdipDeletePath tStringPath
    GdipCreateSolidFill MaskColor, tBrush2
    GdipCreatePath FillModeAlternate, tStringPath
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectL, tStringFormat '添加字符串路径
    MaskWidth = mRectF.Right / 100
    GdipSetClipRectI tGraphics, tRectF.Left, tRectF.Top, MaskWidth, tRectF.Bottom, CombineModeReplace
    GdipFillPath tGraphics, tBrush2, tStringPath                                '填充遮罩层路径
    GdipResetClip tGraphics
    GdipDeleteBrush tBrush1
    GdipDeleteBrush tBrush2
    GdipDeletePen tPen
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    GdipDeleteFont tFont
    GdipDeletePath tStringPath
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'DrawDoubleColorsStringWithMaskByPercentage：以绘制矩形实际宽度百分比形式绘制带遮罩的字符串
'Percentage：遮罩百分比。介于0-100之间。
Public Sub DrawDoubleColorsStringWithMaskByPercentage(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor As Long, ByVal MaskColor As Long, ByVal Percentage As Single, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    If Percentage < 0 Or Percentage > 100 Then MsgBox "参数错误！Percentage的值介于[0,100]。", vbCritical, "参数错误": Exit Sub
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' 根据指定的字体系列创建tFontFamily对象
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' 基于字符串格式标志和语言创建tStringFormat对象
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' 设置此tStringFormat对象相对于布局矩形的原点的字符对齐，使用布局矩形来定位显示的字符串
    Dim tPen As Long, tBrush1 As Long, tBrush2 As Long, tRect As Long, tRectF As RectF, tRectL As RectL, tMask As Long
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    With tRectLayout
        .Left = Left
        .Top = Top
        .Bottom = Top + Height
        .Right = Left + Width
    End With
    With tRectF
        .Left = tRectLayout.Left
        .Top = tRectLayout.Top
        .Bottom = tRectLayout.Bottom
        .Right = .Left + Width * Percentage
    End With
    With tRectL
        .Left = tRectLayout.Left
        .Top = tRectLayout.Top
        .Bottom = tRectLayout.Bottom
        .Right = .Left + Width * Percentage
    End With
    GdipCreateRegionRect tRectF, tRect                                          '创建矩形区域
    GdipCreateSolidFill FillColor, tBrush1
    GdipCreatePath FillModeAlternate, tStringPath                               '创建字符串路径
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '添加字符串路径
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '描边
    GdipFillPath tGraphics, tBrush1, tStringPath                                '填充底层路径
    GdipDeletePath tStringPath
    GdipCreateSolidFill MaskColor, tBrush2
    GdipCreatePath FillModeAlternate, tStringPath
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '添加字符串路径
    Dim MaskWidth As Single
    MaskWidth = Width * Percentage / 100
    GdipSetClipRectI tGraphics, Left, Top, MaskWidth, Height, CombineModeReplace
    GdipFillPath tGraphics, tBrush2, tStringPath                                '填充遮罩层路径
    GdipResetClip tGraphics
    GdipDeleteBrush tBrush1
    GdipDeleteBrush tBrush2
    GdipDeletePen tPen
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    GdipDeletePath tStringPath
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'DrawLine：画直线段。
'X1/Y1/X2/Y2：点1/的横/纵坐标。
'Width：线宽。
'Color：线颜色。采用ARGB。
Public Sub DrawLine(ByVal ContainerHDC As Long, _
    ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, _
    ByVal Width As Long, Optional ByVal Color As Long = &HFF000000)
    Dim tPen As Long, tGraphics As Long
    If Width <= 0 Then MsgBox "参数无效。粗细应大于0。": Exit Sub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreatePen1 Color, Width, UnitPixel, tPen
    GdipDrawLine tGraphics, tPen, X1, Y1, X2, Y2
    GdipDeleteGraphics tGraphics
    GdipDeletePen tPen
    If tFlag Then TerminateGDIPlus
End Sub

'DrawArc：画正圆的弧线。
'CircleCenterX/CircleCenterY：圆心坐标。
'Radius：半径。
'StartDeg/FinishDeg：起始角度/终止角度。采用弧度制。
Public Sub DrawArc(ByVal ContainerHDC As Long, ByVal CircleCenterX As Long, ByVal CircleCenterY As Long, ByVal Radius As Long _
    , ByVal StartDeg As Single, ByVal FinishDeg As Single, _
    ByVal Width As Long, Optional ByVal Color As Long = &HFF000000)
    If Width <= 0 Then MsgBox "参数无效。粗细应大于0。": Exit Sub
    If Radius <= 0 Then MsgBox "参数无效。半径应大于0。": Exit Sub
    Dim tPen As Long, tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 Color, Width, UnitPixel, tPen
    GdipDrawArcI tGraphics, tPen, CircleCenterX - Radius, CircleCenterY - Radius, 2 * Radius, 2 * Radius, StartDeg, FinishDeg
    GdipDeletePen tPen
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'DrawBasicShape：画椭圆/矩形
'ImageShape：形状。
'ShapeMode：绘制方式。
Public Sub DrawBasicShape(ByVal ContainerHDC As Long, _
    ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, _
    ByVal FillColor As Long, ByVal BorderColor As Long, _
    ByVal ImageShape As BasicShape, _
    ByVal ShapeMode As ShapeStyle, _
    Optional ByVal BorderWidth As Long = 1)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim tPen As Long, tGraphics As Long, tBrush As Long
    If Width <= 0 Or Height <= 0 Then
        MsgBox "参数错误！", vbCritical, "错误"
        Exit Sub
    End If
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '获得指定设备场景的HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen                    '创建画笔
    GdipCreateSolidFill FillColor, tBrush                                       '创建画刷
    Select Case ShapeMode
    Case 0                                                                      '仅绘制边框
        If ImageShape = 0 Then
            GdipDrawRectangleI tGraphics, tPen, Left, Top, Width, Height        '矩形
        Else
            GdipDrawEllipseI tGraphics, tPen, Left, Top, Width, Height          '椭圆
        End If
    Case 1                                                                      '仅填充内部
        If ImageShape = 0 Then
            GdipFillRectangleI tGraphics, tBrush, Left, Top, Width, Height      '矩形
        Else
            GdipFillEllipseI tGraphics, tBrush, Left, Top, Width, Height        '椭圆
        End If
    Case 2                                                                      '先填充，后绘制
        If ImageShape = 0 Then
            GdipFillRectangleI tGraphics, tBrush, Left, Top, Width, Height      '矩形
            GdipDrawRectangleI tGraphics, tPen, Left, Top, Width, Height
        Else
            GdipFillEllipseI tGraphics, tBrush, Left, Top, Width, Height        '椭圆
            GdipDrawEllipseI tGraphics, tPen, Left, Top, Width, Height
        End If
    End Select
    GdipDeleteBrush tBrush
    GdipDeletePen tPen
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'SetPointsArray：设置点数组。
'PointsArray()：数组名称。此数组必须为动态数组。
'PointsX/PointsY：横/纵坐标字符串。如“3,4,5,2”。
'Delimiter：分隔符，默认是半角逗号。
Public Sub SetPointsArray(ByRef PointsArray() As PointL, ByVal PointsX As String, ByVal PointsY As String, _
    Optional ByVal Delimiter As String = ",")
    Dim pX As Variant, pY As Variant, Count As Long, t As Long
    If Delimiter = "" Then MsgBox "无效的分隔符。分隔符为一个字符。", vbCritical, "无效分隔符": Exit Sub
    pX = Split(PointsX, Delimiter): pY = Split(PointsY, Delimiter)
    If UBound(pX) <> UBound(pY) Then MsgBox "点集横坐标X的数量和纵坐标Y的数量不一致。", vbCritical, "数据缺失或错误": Exit Sub
    Count = UBound(pX)
    ReDim PointsArray(Count) As PointL
    For t = 0 To Count
        PointsArray(t).X = Val(pX(t))
        PointsArray(t).Y = Val(pY(t))
    Next t
End Sub

'DrawPolygonByPointsArray：绘制闭合N边形。至少为3条边。
'PointsArray：端点的位置数组。
Public Sub DrawPolygonByPointsArray(ByVal ContainerHDC As Long, ByRef PointsArray() As PointL, _
    ByVal FillColor As Long, ByVal BorderColor As Long, _
    ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    Dim tGraphics As Long, tPen As Long, tBrush As Long, iPoint As Long, tPath As Long, lPath() As Long
    If UBound(PointsArray) < 2 Then MsgBox "点集数组数据不足。至少需要3个点。", vbCritical, "数据不足": Exit Sub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateSolidFill FillColor, tBrush
    GdipCreatePath FillModeWinding, tPath
    ReDim lPath(UBound(PointsArray)) As Long
    For iPoint = 0 To UBound(PointsArray) - 1
        GdipCreatePath FillModeWinding, lPath(iPoint)
        GdipAddPathLineI lPath(iPoint), PointsArray(iPoint).X, PointsArray(iPoint).Y, PointsArray(iPoint + 1).X, PointsArray(iPoint + 1).Y
        GdipAddPathPath tPath, lPath(iPoint), 1
    Next iPoint
    GdipCreatePath FillModeWinding, lPath(UBound(PointsArray))
    GdipAddPathLineI lPath(UBound(PointsArray)), PointsArray(UBound(PointsArray)).X, PointsArray(UBound(PointsArray)).Y, _
    PointsArray(0).X, PointsArray(0).Y
    GdipAddPathPath tPath, lPath(UBound(PointsArray)), 1
    Select Case ShapeMode
    Case BorderOnly
        GdipDrawPath tGraphics, tPen, tPath
    Case FillOnly
        GdipFillPath tGraphics, tBrush, tPath
    Case Both
        GdipDrawPath tGraphics, tPen, tPath
        GdipFillPath tGraphics, tBrush, tPath
    End Select
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeletePath tPath
    For iPoint = 0 To UBound(PointsArray)
        GdipDeletePath lPath(iPoint)
    Next iPoint
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    If tFlag Then TerminateGDIPlus
End Sub

'DrawRoundRectangle：绘制圆角矩形（含跑道形）。
'RoundSize：圆角的半径。
Public Sub DrawRoundRectangle(ByVal ContainerHDC As Long, ByVal Left As Long, ByVal Top As Long, _
    ByVal Width As Long, ByVal Height As Long, _
    ByVal RoundSize As Long, ByVal FillColor As Long, ByVal BorderColor As Long, _
    ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If Width <= 0 Then MsgBox "参数错误！Width的值必须大于0。", vbCritical, "参数错误": Exit Sub
    If Height <= 0 Then MsgBox "参数错误！Height的值必须大于0。", vbCritical, "参数错误": Exit Sub
    If RoundSize < 0 Or RoundSize > Width / 2 Or RoundSize > Height / 2 Then MsgBox _
    "参数错误！RoundSize的值必须大于等于0且小于" & Width / 2 & "和" & Height / 2 & "。", vbCritical, "参数错误": Exit Sub
    Dim tGraphics As Long, tPath(7) As Long, tPen As Long, tBrush As Long
    Dim nPath As Long, t As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias
    GdipCreatePath FillModeWinding, nPath
    For t = 0 To 7
        GdipCreatePath FillModeWinding, tPath(t)
    Next t
    GdipAddPathArcI tPath(0), Left, Top, RoundSize * 2, RoundSize * 2, 180, 90
    GdipAddPathArcI tPath(2), Left + Width - 2 * RoundSize, Top, RoundSize * 2, RoundSize * 2, 270, 90
    GdipAddPathArcI tPath(4), Left + Width - 2 * RoundSize, Top + Height - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 0, 90
    GdipAddPathArcI tPath(6), Left, Top + Height - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 90, 90
    GdipAddPathLineI tPath(1), Left + RoundSize, Top, Left + Width - RoundSize, Top
    GdipAddPathLineI tPath(3), Left + Width, Top + RoundSize, Left + Width, Top + Height - RoundSize
    GdipAddPathLineI tPath(5), Left + Width - RoundSize, Top + Height, Left + RoundSize, Top + Height
    GdipAddPathLineI tPath(7), Left, Top + RoundSize, Left, Top + Height - RoundSize
    For t = 0 To 7
        GdipAddPathPath nPath, tPath(t), 1
    Next t
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateSolidFill FillColor, tBrush
    Select Case ShapeMode
    Case BorderOnly
        GdipDrawPath tGraphics, tPen, nPath
    Case FillOnly
        GdipFillPath tGraphics, tBrush, nPath
    Case Both
        GdipDrawPath tGraphics, tPen, nPath
        GdipFillPath tGraphics, tBrush, nPath
    End Select
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeletePath nPath
    For t = 0 To 7
        GdipDeletePath tPath(t)
    Next t
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    If tFlag Then TerminateGDIPlus
End Sub

'NewRectL：设置或者创建一个RectL结构体
Public Function NewRectL(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RectL
    Dim Retn As RectL
    With Retn
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    NewRectL = Retn
End Function

'NewPointL：设置或者创建一个PointL结构体
Public Function NewPointL(ByVal X As Long, ByVal Y As Long) As PointL
    Dim Retn As PointL
    With Retn
        .X = X
        .Y = Y
    End With
    NewPointL = Retn
End Function
