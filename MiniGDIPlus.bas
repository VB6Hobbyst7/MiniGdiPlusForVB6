Attribute VB_Name = "MiniGDIPlus"
'MiniGDIPlus
'���������ư��侩��
'�汾��0.1beta

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
    GdiplusVersion As Long                                                      '�汾
    DebugEventCallback As Long                                                  '�����¼��ص�
    SuppressBackgroundThread As Long                                            '���Ʊ����߳�
    SuppressExternalCodecs As Long                                              '�����ⲿ�������
End Type

Public Type tImage
    imgHandle As Long                                                           'ͼƬ�ľ����
    imgIndex As Long                                                            'ͼƬ���������ɿա�
    imgName As String                                                           'ͼƬ�����ƣ�ͨ��������ͼƬָ����ļ���������
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
    UnitWorld                                                                   'ȫ��
    UnitDisplay                                                                 '��ʾ
    UnitPixel                                                                   '����
    UnitPoint                                                                   '��
    UnitInch                                                                    'Ӣ��
    UnitDocument                                                                '�ĵ�
    UnitMillimeter                                                              '����
End Enum
 
Public Enum GpStatus
    Ok = 0                                                                      'û����
    GenericError = 1                                                            'ͨ�ô���
    InvalidParameter = 2                                                        '��Ч����
    OutOfMemory = 3                                                             '�ڴ����
    ObjectBusy = 4                                                              '����æ
    InsufficientBuffer = 5                                                      '��������������
    NotImplemented = 6                                                          'δִ��
    Win32Error = 7                                                              '����Win32�Ĵ���
    WrongState = 8                                                              '״̬����ȷ
    Aborted = 9                                                                 '����ֹ
    FileNotFound = 10                                                           '�ļ�δ�ҵ�
    ValueOverflow = 11                                                          'ֵ���
    AccessDenied = 12                                                           '���ʱ��ܾ�
    UnknownImageFormat = 13                                                     'δ֪��ͼ���ʽ
    FontFamilyNotFound = 14                                                     'δ�ҵ�������
    FontStyleNotFound = 15                                                      'δ�ҵ�������ʽ
    NotTrueTypeFont = 16                                                        '����TrueType����
    UnsupportedGdiplusVersion = 17                                              '��֧�ֵ�GDIPlus�汾
    GdiplusNotInitialized = 18                                                  'GDIPlusδ����ʼ��
    PropertyNotFound = 19                                                       '����δ�ҵ�
    PropertyNotSupported = 20                                                   '���Բ�֧��
End Enum

Public Enum SmoothingMode
    SmoothingModeInvalid = -1                                                   '��Ч/������
    SmoothingModeDefault = 0                                                    'Ĭ��
    SmoothingModeHighSpeed = 1                                                  '����
    SmoothingModeHighQuality = 2                                                '������
    SmoothingModeNone = 3                                                       '������
    SmoothingModeAntiAlias = 4                                                  '�����
End Enum

Private mToken As Long

'InitializeGDIPlus����ʼ��GDIPlus��һ������������ڵ�Form_Load�¼�����Sub Main��
Public Sub InitializeGDIPlus()
    If mToken <> 0 Then Exit Sub
    Dim Retn As GpStatus, uInput As GdiplusStartupInput
    uInput.GdiplusVersion = 1
    Retn = GdiplusStartup(mToken, uInput)
    If Retn <> 0 Then Debug.Print "GDIPlusδ�ܳ�ʼ����"
End Sub

'TerminateGDIPlus����ֹ��GDIPlus��һ������������ڵ�Form_Unload�¼���
Public Sub TerminateGDIPlus()
    If mToken = 0 Then Exit Sub
    GdiplusShutdown mToken
    mToken = 0
End Sub

'DrawImageFromFile�����ļ��д���ͼ����󣬲������������
'Container��ָ������������Ҫ��hDC�������Ǵ������ͼƬ��
'Path��ͼ�����·��
'Left/Top��ͼ�����Ͻ������������λ�ã����أ�
'Zoom�����ű�����С��1ʱΪ��С������1ʱΪ�Ŵ�
Public Sub DrawImageFromFile(ByVal Container As Object, ByVal path As String, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, Optional ByVal Zoom As Double = 1#)
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    If Zoom <= 0 Then
        MsgBox "���ֲ�������Zoom��ֵӦ����0��", vbCritical, "����"
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageOnPictureBoxFromFile����ͼƬ���ϻ��Ʋ����ļ�����·��ָ����ͼ��
'AutoSize���Ƿ�ʹͼƬ����Ӧͼ��ߴ硣ΪTrueʱ�Զ�����ͼƬ��ߴ�����Ӧͼ��ΪFalseʱ�Զ�����ͼ������Ӧͼ���
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageFromHandle���Ӿ���д���ͼ����󣬲������������
'Container��ָ������������Ҫ��hDC�������Ǵ������ͼƬ��
'Handle��ͼ����
'Left/Top��ͼ�����Ͻ������������λ�ã����أ�
'Zoom�����ű�����С��1ʱΪ��С������1ʱΪ�Ŵ�
Public Sub DrawImageFromHandle(ByVal Container As Object, ByVal Handle As Long, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, Optional ByVal Zoom As Single = 1#)
    On Error GoTo ExitSub
    If Not Container.HasDC Then Exit Sub
    If Zoom <= 0 Then
        MsgBox "���ֲ�������Zoom��ֵӦ����0��", vbCritical, "GDIPlus2"
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'��ע���˺�����bug�����ܵ��¿�����
'DrawImageFromHandleRectToRect���Ӿ���д���ͼ������ָ�����򣬲�����������ϡ�
'dstLeft/dstTop/dstWidth/dstHeight��Ŀ���λ�úͳߴ硣
'srcLeft/srcTop/srcWidth/srcHeight��Դ��λ�úͳߴ硣
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'��ע���˺�����bug�����ܵ��¿�����
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'DrawImageOnPictureBoxFromHandle����ͼƬ���ϻ��Ʋ���ͼ�������õ�ͼ��
'AutoSize���Ƿ�ʹͼƬ����Ӧͼ��ߴ硣ΪTrueʱ�Զ�����ͼƬ��ߴ�����Ӧͼ��ΪFalseʱ�Զ�����ͼ������Ӧͼ���
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
    MsgBox "���ֻ��ƴ���", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
    Exit Sub
End Sub

'LoadImageHandleFromFile����ͼ���ļ�������һ��ͼ����
Public Function LoadImageHandleFromFile(ByVal path As String) As Long
    Dim Image As Long
    On Error GoTo ExitSub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    If Dir(path) = "" Then
        MsgBox "�ļ� " & path & " �����ڡ�", vbCritical, "����"
        Exit Function
    End If
    Call GdipLoadImageFromFile(StrPtr(path), Image)
    LoadImageHandleFromFile = Image
    If tFlag Then TerminateGDIPlus
    Exit Function
ExitSub:
    MsgBox "����δ֪����", vbCritical, "����"
    If tFlag Then TerminateGDIPlus
End Function

'DelImageHandle�����ڴ���ɾ��һ��ͼ����
Public Sub DelImageHandle(ByRef Handle As Long)
    If Handle = 0 Then Exit Sub
    On Error Resume Next
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Call GdipDisposeImage(Handle)
    If tFlag Then TerminateGDIPlus
End Sub

'LoadImageFromFolderToImageArray����һ��ָ���ļ��е�����ͼƬ��ָ��ͬһ����׺�����ص�ͼƬ�����
'ImageArray��Ҫ��ӵ���ͼƬ���顣
'FolderPath���ļ���·����
'Suffix���ļ�����׺��
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
        MsgBox "�ļ��в����ڡ�", vbCritical, "����"
    End If
    Set FSO = Nothing: Set FolderObj = Nothing: Set FileObj = Nothing
End Sub

'LoadImageFromFolderToHandleArray����һ��ָ���ļ��е�����ͼƬ���ص�һ�������
'HandleArray��Ҫ��ӵ���ͼ�������顣
'FolderPath���ļ���·����
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
        MsgBox "�ļ��в����ڡ�", vbCritical, "����"
    End If
    Set FSO = Nothing: Set FolderObj = Nothing: Set FileObj = Nothing
End Sub

'LoadImageFromFileStringSuffixToHandleArray����ָ���ļ����°����ļ����ַ����������ض�ͼƬ�ļ����ص�һ�������
'HandleArray��Ҫ��ӵ���ͼ�������顣
'FolderPath���ļ���·����
'FileString������ļ����ͷָ�����ɵ��ַ�������ʽ���ļ��� �ָ��� �ļ��� �ָ��� �ļ�����������
'Delimiter���ָ�����ȱʡΪ��Ƕ��š��������ǿմ���""����
'Suffix���ļ�����׺����ָ�����ļ�����׺ʱ��FileString�������ļ�������������ļ�����׺��������Ҫ����Ϊÿһ���ļ�ָ����׺��
Public Sub LoadImageFromFileStringSuffixToHandleArray(ByRef HandleArray() As Long, ByVal FolderPath As String, ByVal FileString As String, _
    Optional ByVal Delimiter As String = ",", Optional ByVal Suffix As String = "")
    If Delimiter = "" Then MsgBox "�ָ�����Delimiter����Ч��", vbCritical, "����": Exit Sub
    If Suffix <> "" And LCase(Suffix) <> "png" And LCase(Suffix) <> "gif" And LCase(Suffix) <> "bmp" And LCase(Suffix) <> "jpg" _
        Then MsgBox "ָ�����ļ���׺��Ч��" & vbCrLf & "�ļ���׺������png��gif��bmp��jpg����һ�ֻ�ʡ�ԡ�", vbCritical, "����"
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
            Debug.Print "����ļ���" & nPath & FileArray(i)
        Next i
End Sub

'ReleaseImgHandleArray���ͷ�һ��ͼ�������������е�ͼ�������˺���ͨ������GDIPlus���ս�֮ǰ��
'HandleArray��ͼ�������顣
Public Sub ReleaseImgHandleArray(ByRef HandleArray() As Long)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim i As Long
    For i = 0 To UBound(HandleArray)
        GdipDisposeImage HandleArray(i)
    Next i
    If tFlag Then TerminateGDIPlus
End Sub

'ReleaseImgImageArray���ͷ�һ��ͼƬ���������е�ͼ�������˺���ͨ������GDIPlus���ս�֮ǰ��
'ImageArray��ͼƬ���顣
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

'GetFileSuffix������ļ��ĺ�׺����
'FilePath���ļ�·����
Public Function GetFileSuffix(ByVal FilePath As String) As String
    Dim nPath As String
    nPath = FilePath
    GetFileSuffix = LCase(Right(nPath, Len(nPath) - InStrRev(nPath, ".")))
End Function

'CombinePathFromPathArray����һ��·�������е�����·���ϲ���ָ����Ŀ��·���У��ϲ���ʽΪ�򣨲������㡣Ŀ��·����Ҫ�ȱ�������
'PathArray��·��������顣
'TargetPath��ָ����Ŀ��·����
Public Sub CombinePathFromPathArray(ByRef TargetPath As Long, ByRef PathArray() As Long)
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    Dim i As Long
    For i = 0 To UBound(PathArray)
        GdipAddPathPath TargetPath, PathArray(i), 1
    Next i
    If tFlag Then TerminateGDIPlus
End Sub

'DrawPath���������ָ��·����
'ContainerHDC���������豸������Device Context�������ֻ����ʹ�ô����ͼƬ���hDC��
'Path����ͼ·����
'FillColor�����ɫ��
'BorderColor�����ɫ��
'ShapeMode�����������ʽ��
'BorderWidth����ߴ�ϸ��
Public Sub DrawPath(ByRef ContainerHDC As Long, ByVal path As Long, ByVal FillColor As Long, ByVal BorderColor As Long, ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If BorderWidth <= 0 Then MsgBox "ָ������ߴ�ϸBorderWidth��Ч���ò���ֵ�������0��", vbCritical, "����": Exit Sub
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

'DrawPathWithGradientColors�������������ָ��·����
'ContainerHDC���������豸������Device Context�������ֻ����ʹ�ô����ͼƬ���hDC��
'Path����ͼ·����
'FillColor1(2)�����ɫ1(2)��
'BorderColor�����ɫ��
'Point1(2)����(��)�㡣
'ShapeMode�����������ʽ��
'BorderWidth����ߴ�ϸ��
Public Sub DrawPathWithGradientColors(ByRef ContainerHDC As Long, ByVal path As Long, ByVal FillColor1 As Long, ByVal FillColor2 As Long, ByVal BorderColor As Long, ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If BorderWidth <= 0 Then MsgBox "ָ������ߴ�ϸBorderWidth��Ч���ò���ֵ�������0��", vbCritical, "����"
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

'DrawImageByName������ͼƬ���ƻ���ͼƬ��
'Container��������ֻ����ʹ�ô����ͼƬ��
'ImageArray��ָ����ͼƬ���顣
'imgName��ͼƬ�����ƣ�ͨ��ΪͼƬ���ļ����ơ�
'Left/Top������ͼƬ����߾�/�ϱ߾ࡣ
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
    If pHandle = 0 Then                                                         '�Ҳ�����Ӧ��ͼƬ
        Debug.Print "DrawImageByName���Ҳ�����Ӧ��ͼƬ��"
        Exit Sub
    End If
    DrawImageFromHandle Container, pHandle, Left, Top, 1
    If tFlag Then TerminateGDIPlus
End Sub

'DrawSimpleString�����Ƽ��ַ���
'ContainerHDC���������豸������Device Context�������ֻ����ʹ�ô����ͼƬ���hDC��
'StringText���ַ����ı����ݡ�
'StringFontFamily�����塣
'StringSize���ֺš�
'StringStyle������
'FillColor/BorderColor�����ɫ/���ɫ������ARGB��
'BorderWidth����߿�ȡ�
'DrawBorder���Ƿ���ߡ�
Public Sub DrawSimpleString(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor As Long, Optional ByVal BorderColor As Long = 0, _
    Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       '����ָ��������ϵ�д���tFontFamily����
    GdipCreateStringFormat 0, 0, tStringFormat                                  '�����ַ�����ʽ��־�����Դ���tStringFormat����
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 '���ô�tStringFormat��������ڲ��־��ε�ԭ����ַ����룬ʹ�ò��־�������λ��ʾ���ַ���
    Dim tPen As Long, tBrush As Long
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen
    GdipCreateSolidFill FillColor, tBrush
    With tRectLayout
        .Left = Left
        .Top = Top
        .Bottom = Top + Height
        .Right = Left + Width
    End With
    GdipCreatePath FillModeAlternate, tStringPath                               '�����ַ���·��
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '����ַ���·��
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '���
    GdipFillPath tGraphics, tBrush, tStringPath                                 '���·��
    GdipDeletePen tPen
    GdipDeleteBrush tBrush
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    If tFlag Then TerminateGDIPlus
End Sub

'DrawGradientString�����ƽ����ַ��������䷽������ң�
Public Sub DrawGradientString(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor1 As Long, ByVal FillColor2 As Long, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' ����ָ��������ϵ�д���tFontFamily����
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' �����ַ�����ʽ��־�����Դ���tStringFormat����
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' ���ô�tStringFormat��������ڲ��־��ε�ԭ����ַ����룬ʹ�ò��־�������λ��ʾ���ַ���
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
    GdipCreatePath FillModeAlternate, tStringPath                               '�����ַ���·��
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '����ַ���·��
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '���
    GdipFillPath tGraphics, tBrush, tStringPath                                 '���·��
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    GdipDeletePath tStringPath
    GdipDeleteBrush tBrush
    GdipDeletePen tPen
    GdipDeleteFontFamily tFontFamily
    GdipDeleteStringFormat tStringFormat
    If tFlag Then TerminateGDIPlus
End Sub

'��ע���˺�����bug�������ַ������־���tRectF����д��󣬵����ְٷֱ�������ֵ������
'DrawDoubleColorsStringWithMaskByStringPercentage�����ַ������ݰٷֱ���ʽ���ƴ����ֵ��ַ���
'MaskColor��������ɫ������ARGB��
'StringPercentage�����ְٷֱȡ�����0-100֮�䡣
Public Sub DrawDoubleColorsStringWithMaskByStringPercentage(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, _
    ByVal FillColor As Long, ByVal MaskColor As Long, ByVal StringPercentage As Single, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    If StringPercentage < 0 Or StringPercentage > 100 Then MsgBox "��������StringPercentage��ֵ����[0,100]��", vbCritical, "��������": Exit Sub
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tFont As Long
    Dim tGraphics As Long
    Dim tCodePointsFitted As Long, tLinesFilled As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' ����ָ��������ϵ�д���tFontFamily����
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' �����ַ�����ʽ��־�����Դ���tStringFormat����
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' ���ô�tStringFormat��������ڲ��־��ε�ԭ����ַ����룬ʹ�ò��־�������λ��ʾ���ַ���
    GdipCreateFont tFontFamily, StringSize, StringStyle, UnitPixel, tFont       '����tFont����
    Dim tPen As Long, tBrush1 As Long, tBrush2 As Long, tRect As Long, fRectF As RectF, tRectL As RectL, tRectF As RectF, tMask As Long, mRectF As RectF, mRect As Long
    Dim MaskWidth As Single                                                     '���ֿ�
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen                    '��������
    GdipSetStringFormatTrimming tStringFormat, StringTrimmingEllipsisCharacter  '���ַ�����ʽ����������õȿ��ַ������ģ��ľ���fRectL
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
    GdipCreateRegionRect mRectF, mRect                                          '������������
    GdipCreateSolidFill FillColor, tBrush1                                      '��仭ˢ
    GdipCreatePath FillModeAlternate, tStringPath                               '�����ַ���·��
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectL, tStringFormat '����ַ���·��
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '���
    GdipFillPath tGraphics, tBrush1, tStringPath                                '���ײ�·��
    GdipDeletePath tStringPath
    GdipCreateSolidFill MaskColor, tBrush2
    GdipCreatePath FillModeAlternate, tStringPath
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectL, tStringFormat '����ַ���·��
    MaskWidth = mRectF.Right / 100
    GdipSetClipRectI tGraphics, tRectF.Left, tRectF.Top, MaskWidth, tRectF.Bottom, CombineModeReplace
    GdipFillPath tGraphics, tBrush2, tStringPath                                '������ֲ�·��
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

'DrawDoubleColorsStringWithMaskByPercentage���Ի��ƾ���ʵ�ʿ�Ȱٷֱ���ʽ���ƴ����ֵ��ַ���
'Percentage�����ְٷֱȡ�����0-100֮�䡣
Public Sub DrawDoubleColorsStringWithMaskByPercentage(ByVal ContainerHDC As Long, _
    ByVal StringText As String, ByVal StringFontFamily As String, ByVal StringSize As Single, ByVal StringStyle As FontStyle, _
    ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, _
    ByVal FillColor As Long, ByVal MaskColor As Long, ByVal Percentage As Single, _
    Optional ByVal BorderColor As Long = 0, Optional ByVal BorderWidth As Long = 1, Optional ByVal DrawBorder As Boolean = False)
    If Percentage < 0 Or Percentage > 100 Then MsgBox "��������Percentage��ֵ����[0,100]��", vbCritical, "��������": Exit Sub
    Dim tFontFamily As Long, tStringFormat As Long, tStringPath As Long, tRectLayout As RectL
    Dim tGraphics As Long
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateFontFamilyFromName StrPtr(StringFontFamily), 0, tFontFamily       ' ����ָ��������ϵ�д���tFontFamily����
    GdipCreateStringFormat 0, 0, tStringFormat                                  ' �����ַ�����ʽ��־�����Դ���tStringFormat����
    GdipSetStringFormatAlign tStringFormat, StringAlignmentNear                 ' ���ô�tStringFormat��������ڲ��־��ε�ԭ����ַ����룬ʹ�ò��־�������λ��ʾ���ַ���
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
    GdipCreateRegionRect tRectF, tRect                                          '������������
    GdipCreateSolidFill FillColor, tBrush1
    GdipCreatePath FillModeAlternate, tStringPath                               '�����ַ���·��
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '����ַ���·��
    If DrawBorder Then GdipDrawPath tGraphics, tPen, tStringPath                '���
    GdipFillPath tGraphics, tBrush1, tStringPath                                '���ײ�·��
    GdipDeletePath tStringPath
    GdipCreateSolidFill MaskColor, tBrush2
    GdipCreatePath FillModeAlternate, tStringPath
    GdipAddPathStringI tStringPath, StrPtr(StringText), -1, tFontFamily, StringStyle, StringSize, tRectLayout, tStringFormat '����ַ���·��
    Dim MaskWidth As Single
    MaskWidth = Width * Percentage / 100
    GdipSetClipRectI tGraphics, Left, Top, MaskWidth, Height, CombineModeReplace
    GdipFillPath tGraphics, tBrush2, tStringPath                                '������ֲ�·��
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

'DrawLine����ֱ�߶Ρ�
'X1/Y1/X2/Y2����1/�ĺ�/�����ꡣ
'Width���߿�
'Color������ɫ������ARGB��
Public Sub DrawLine(ByVal ContainerHDC As Long, _
    ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, _
    ByVal Width As Long, Optional ByVal Color As Long = &HFF000000)
    Dim tPen As Long, tGraphics As Long
    If Width <= 0 Then MsgBox "������Ч����ϸӦ����0��": Exit Sub
    Dim tFlag As Boolean
    If mToken = 0 Then tFlag = True: InitializeGDIPlus
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreatePen1 Color, Width, UnitPixel, tPen
    GdipDrawLine tGraphics, tPen, X1, Y1, X2, Y2
    GdipDeleteGraphics tGraphics
    GdipDeletePen tPen
    If tFlag Then TerminateGDIPlus
End Sub

'DrawArc������Բ�Ļ��ߡ�
'CircleCenterX/CircleCenterY��Բ�����ꡣ
'Radius���뾶��
'StartDeg/FinishDeg����ʼ�Ƕ�/��ֹ�Ƕȡ����û����ơ�
Public Sub DrawArc(ByVal ContainerHDC As Long, ByVal CircleCenterX As Long, ByVal CircleCenterY As Long, ByVal Radius As Long _
    , ByVal StartDeg As Single, ByVal FinishDeg As Single, _
    ByVal Width As Long, Optional ByVal Color As Long = &HFF000000)
    If Width <= 0 Then MsgBox "������Ч����ϸӦ����0��": Exit Sub
    If Radius <= 0 Then MsgBox "������Ч���뾶Ӧ����0��": Exit Sub
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

'DrawBasicShape������Բ/����
'ImageShape����״��
'ShapeMode�����Ʒ�ʽ��
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
        MsgBox "��������", vbCritical, "����"
        Exit Sub
    End If
    GdipCreateFromHDC ContainerHDC, tGraphics                                   '���ָ���豸������HDC
    GdipSetSmoothingMode tGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreatePen1 BorderColor, BorderWidth, UnitPixel, tPen                    '��������
    GdipCreateSolidFill FillColor, tBrush                                       '������ˢ
    Select Case ShapeMode
    Case 0                                                                      '�����Ʊ߿�
        If ImageShape = 0 Then
            GdipDrawRectangleI tGraphics, tPen, Left, Top, Width, Height        '����
        Else
            GdipDrawEllipseI tGraphics, tPen, Left, Top, Width, Height          '��Բ
        End If
    Case 1                                                                      '������ڲ�
        If ImageShape = 0 Then
            GdipFillRectangleI tGraphics, tBrush, Left, Top, Width, Height      '����
        Else
            GdipFillEllipseI tGraphics, tBrush, Left, Top, Width, Height        '��Բ
        End If
    Case 2                                                                      '����䣬�����
        If ImageShape = 0 Then
            GdipFillRectangleI tGraphics, tBrush, Left, Top, Width, Height      '����
            GdipDrawRectangleI tGraphics, tPen, Left, Top, Width, Height
        Else
            GdipFillEllipseI tGraphics, tBrush, Left, Top, Width, Height        '��Բ
            GdipDrawEllipseI tGraphics, tPen, Left, Top, Width, Height
        End If
    End Select
    GdipDeleteBrush tBrush
    GdipDeletePen tPen
    GdipResetClip tGraphics
    GdipDeleteGraphics tGraphics
    If tFlag Then TerminateGDIPlus
End Sub

'SetPointsArray�����õ����顣
'PointsArray()���������ơ����������Ϊ��̬���顣
'PointsX/PointsY����/�������ַ������硰3,4,5,2����
'Delimiter���ָ�����Ĭ���ǰ�Ƕ��š�
Public Sub SetPointsArray(ByRef PointsArray() As PointL, ByVal PointsX As String, ByVal PointsY As String, _
    Optional ByVal Delimiter As String = ",")
    Dim pX As Variant, pY As Variant, Count As Long, t As Long
    If Delimiter = "" Then MsgBox "��Ч�ķָ������ָ���Ϊһ���ַ���", vbCritical, "��Ч�ָ���": Exit Sub
    pX = Split(PointsX, Delimiter): pY = Split(PointsY, Delimiter)
    If UBound(pX) <> UBound(pY) Then MsgBox "�㼯������X��������������Y��������һ�¡�", vbCritical, "����ȱʧ�����": Exit Sub
    Count = UBound(pX)
    ReDim PointsArray(Count) As PointL
    For t = 0 To Count
        PointsArray(t).X = Val(pX(t))
        PointsArray(t).Y = Val(pY(t))
    Next t
End Sub

'DrawPolygonByPointsArray�����Ʊպ�N���Ρ�����Ϊ3���ߡ�
'PointsArray���˵��λ�����顣
Public Sub DrawPolygonByPointsArray(ByVal ContainerHDC As Long, ByRef PointsArray() As PointL, _
    ByVal FillColor As Long, ByVal BorderColor As Long, _
    ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    Dim tGraphics As Long, tPen As Long, tBrush As Long, iPoint As Long, tPath As Long, lPath() As Long
    If UBound(PointsArray) < 2 Then MsgBox "�㼯�������ݲ��㡣������Ҫ3���㡣", vbCritical, "���ݲ���": Exit Sub
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

'DrawRoundRectangle������Բ�Ǿ��Σ����ܵ��Σ���
'RoundSize��Բ�ǵİ뾶��
Public Sub DrawRoundRectangle(ByVal ContainerHDC As Long, ByVal Left As Long, ByVal Top As Long, _
    ByVal Width As Long, ByVal Height As Long, _
    ByVal RoundSize As Long, ByVal FillColor As Long, ByVal BorderColor As Long, _
    ByVal ShapeMode As ShapeStyle, Optional ByVal BorderWidth As Long = 1)
    If Width <= 0 Then MsgBox "��������Width��ֵ�������0��", vbCritical, "��������": Exit Sub
    If Height <= 0 Then MsgBox "��������Height��ֵ�������0��", vbCritical, "��������": Exit Sub
    If RoundSize < 0 Or RoundSize > Width / 2 Or RoundSize > Height / 2 Then MsgBox _
    "��������RoundSize��ֵ������ڵ���0��С��" & Width / 2 & "��" & Height / 2 & "��", vbCritical, "��������": Exit Sub
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

'NewRectL�����û��ߴ���һ��RectL�ṹ��
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

'NewPointL�����û��ߴ���һ��PointL�ṹ��
Public Function NewPointL(ByVal X As Long, ByVal Y As Long) As PointL
    Dim Retn As PointL
    With Retn
        .X = X
        .Y = Y
    End With
    NewPointL = Retn
End Function
