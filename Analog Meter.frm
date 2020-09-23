VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analog Meter"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "Analog Meter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   750
      Max             =   120
      TabIndex        =   0
      Top             =   2520
      Value           =   1
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Const GdiPlusVersion As Long = 1
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private GdipToken As Long

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As Long

' Quality mode constants
Private Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       ' Best performance
   QualityModeHigh = 2       ' Best rendering quality
End Enum

Private Enum SmoothingMode
   SmoothingModeInvalid = -1
   SmoothingModeDefault = 0
   SmoothingModeHighSpeed = 1
   SmoothingModeHighQuality = 2
   SmoothingModeNone = 3
   SmoothingModeAntiAlias = 4
End Enum

Private Enum GpUnit
   UnitWorld = 0
   ' World coordinate (non-physical unit)
   UnitDisplay = 1
   ' Variable -- for PageTransform only
   UnitPixel = 2
   ' Each unit is one device pixel.
   UnitPoint = 3
   ' Each unit is a printer's point, or 1/72 inch.
   UnitInch = 4
   ' Each unit is 1 inch.
   UnitDocument = 5
   ' Each unit is 1/300 inch.
   UnitMillimeter = 6
   ' Each unit is 1 millimeter.
End Enum

Private Enum Colors
   AliceBlue = &HFFF0F8FF
   AntiqueWhite = &HFFFAEBD7
   Aqua = &HFF00FFFF
   Aquamarine = &HFF7FFFD4
   Azure = &HFFF0FFFF
   Background = &HFFD4D0C8
   Beige = &HFFF5F5DC
   Bisque = &HFFFFE4C4
   Black = &HFF000000
   BlanchedAlmond = &HFFFFEBCD
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   BurlyWood = &HFFDEB887
   CadetBlue = &HFF5F9EA0
   Chartreuse = &HFF7FFF00
   Chocolate = &HFFD2691E
   Coral = &HFFFF7F50
   CornflowerBlue = &HFF6495ED
   Cornsilk = &HFFFFF8DC
   Crimson = &HFFDC143C
   Cyan = &HFF00FFFF
   DarkBlue = &HFF00008B
   DarkCyan = &HFF008B8B
   DarkGoldenrod = &HFFB8860B
   DarkGray = &HFFA9A9A9
   DarkGreen = &HFF006400
   DarkKhaki = &HFFBDB76B
   DarkMagenta = &HFF8B008B
   DarkOliveGreen = &HFF556B2F
   DarkOrange = &HFFFF8C00
   DarkOrchid = &HFF9932CC
   DarkRed = &HFF8B0000
   DarkSalmon = &HFFE9967A
   DarkSeaGreen = &HFF8FBC8B
   DarkSlateBlue = &HFF483D8B
   DarkSlateGray = &HFF2F4F4F
   DarkTurquoise = &HFF00CED1
   DarkViolet = &HFF9400D3
   DeepPink = &HFFFF1493
   DeepSkyBlue = &HFF00BFFF
   Desktop = &HFF3A6EA5
   DimGray = &HFF696969
   DodgerBlue = &HFF1E90FF
   Firebrick = &HFFB22222
   FloralWhite = &HFFFFFAF0
   ForestGreen = &HFF228B22
   Fuchsia = &HFFFF00FF
   Gainsboro = &HFFDCDCDC
   GhostWhite = &HFFF8F8FF
   Gold = &HFFFFD700
   Goldenrod = &HFFDAA520
   Gray = &HFF808080
   Green = &HFF008000
   GreenYellow = &HFFADFF2F
   Honeydew = &HFFF0FFF0
   HotPink = &HFFFF69B4
   IndianRed = &HFFCD5C5C
   Indigo = &HFF4B0082
   Ivory = &HFFFFFFF0
   Khaki = &HFFF0E68C
   Lavender = &HFFE6E6FA
   LavenderBlush = &HFFFFF0F5
   LawnGreen = &HFF7CFC00
   LemonChiffon = &HFFFFFACD
   LightBlue = &HFFADD8E6
   LightCoral = &HFFF08080
   LightCyan = &HFFE0FFFF
   LightGoldenrodYellow = &HFFFAFAD2
   LightGray = &HFFD3D3D3
   LightGreen = &HFF90EE90
   LightPink = &HFFFFB6C1
   LightSalmon = &HFFFFA07A
   LightSeaGreen = &HFF20B2AA
   LightSkyBlue = &HFF87CEFA
   LightSlateGray = &HFF778899
   LightSteelBlue = &HFFB0C4DE
   LightYellow = &HFFFFFFE0
   Lime = &HFF00FF00
   LimeGreen = &HFF32CD32
   Linen = &HFFFAF0E6
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   MediumAquamarine = &HFF66CDAA
   MediumBlue = &HFF0000CD
   MediumOrchid = &HFFBA55D3
   MediumPurple = &HFF9370DB
   MediumSeaGreen = &HFF3CB371
   MediumSlateBlue = &HFF7B68EE
   MediumSpringGreen = &HFF00FA9A
   MediumTurquoise = &HFF48D1CC
   MediumVioletRed = &HFFC71585
   MidnightBlue = &HFF191970
   MintCream = &HFFF5FFFA
   MistyRose = &HFFFFE4E1
   Moccasin = &HFFFFE4B5
   NavajoWhite = &HFFFFDEAD
   Navy = &HFF000080
   OldLace = &HFFFDF5E6
   Olive = &HFF808000
   OliveDrab = &HFF6B8E23
   Orange = &HFFFFA500
   OrangeRed = &HFFFF4500
   Orchid = &HFFDA70D6
   PaleGoldenrod = &HFFEEE8AA
   PaleGreen = &HFF98FB98
   PaleTurquoise = &HFFAFEEEE
   PaleVioletRed = &HFFDB7093
   PapayaWhip = &HFFFFEFD5
   PeachPuff = &HFFFFDAB9
   Peru = &HFFCD853F
   Pink = &HFFFFC0CB
   Plum = &HFFDDA0DD
   PowderBlue = &HFFB0E0E6
   Purple = &HFF800080
   Red = &HFFFF0000
   RosyBrown = &HFFBC8F8F
   RoyalBlue = &HFF4169E1
   SaddleBrown = &HFF8B4513
   Salmon = &HFFFA8072
   SandyBrown = &HFFF4A460
   SeaGreen = &HFF2E8B57
   SeaShell = &HFFFFF5EE
   Sienna = &HFFA0522D
   Silver = &HFFC0C0C0
   SkyBlue = &HFF87CEEB
   SlateBlue = &HFF6A5ACD
   SlateGray = &HFF708090
   Snow = &HFFFFFAFA
   SpringGreen = &HFF00FF7F
   SteelBlue = &HFF4682B4
   Tan = &HFFD2B48C
   Teal = &HFF008080
   Thistle = &HFFD8BFD8
   Tomato = &HFFFF6347
   Tooltip = &HFFFFFFE1
   Transparent = &HFFFFFF
   Turquoise = &HFF40E0D0
   Violet = &HFFEE82EE
   Wheat = &HFFF5DEB3
   White = &HFFFFFFFF
   WhiteSmoke = &HFFF5F5F5
   Yellow = &HFFFFFF00
   YellowGreen = &HFF9ACD32
End Enum

Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Colors, ByVal Width As Single, ByVal unit As GpUnit, ByRef pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal color As Colors, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long

Private Function ShutdownGDIPlus() As Long
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function
Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Long
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Function



Private Sub Form_Load()

    Call StartUpGDIPlus(GdiPlusVersion)

    Form1.Left = (Screen.Width / 2) - (Form1.Width / 2)
    Form1.Top = (Screen.Height / 2) - (Form1.Height / 2)
    
    HScroll1.Value = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ShutdownGDIPlus
End Sub


Private Sub HScroll1_Change()

    Dim x_Count As Integer
    Dim m_Externo As Integer
    Dim m_Interno As Integer
    Dim m_Graphics As Long
    Dim m_Pen As Long
    Dim m_Brush As Long
    Dim m_Count As Integer
    Dim m_Sin As Integer
    Dim m_Cos As Integer
    Const PI = 3.14159265358979
    Const m_Number = PI / 180
    Dim m_X As Integer
    Dim m_Y As Integer
    Dim m_Radio As Single
    
    m_X = Me.ScaleWidth / 2
    m_Y = Me.ScaleHeight / 2
    
    Me.Cls
    Call GdipCreateFromHDC(Me.hdc, m_Graphics)
    Call GdipSetSmoothingMode(m_Graphics, SmoothingModeAntiAlias)
    
    Call GdipCreatePen1(Black, 4, UnitPixel, m_Pen)
    Call GdipDrawArc(m_Graphics, m_Pen, Me.ScaleWidth / 2 - 75, Me.ScaleHeight / 2 - 75, 150, 150, -225, 270)
    
    Call GdipCreatePen1(Gold, 2, UnitPixel, m_Pen)
    Call GdipDrawArc(m_Graphics, m_Pen, Me.ScaleWidth / 2 - 75, Me.ScaleHeight / 2 - 75, 150, 150, -225, 270)
    
    Call GdipCreatePen1(Black, 4, UnitPixel, m_Pen)
    Call GdipDrawArc(m_Graphics, m_Pen, Me.ScaleWidth / 2 - 50, Me.ScaleHeight / 2 - 50, 100, 100, -225, 270)
   
    Call GdipCreatePen1(Gold, 2, UnitPixel, m_Pen)
    Call GdipDrawArc(m_Graphics, m_Pen, Me.ScaleWidth / 2 - 50, Me.ScaleHeight / 2 - 50, 100, 100, -225, 270)
    
    m_Interno = 53
    m_Externo = 72
    For m_Count = 135 To 405 Step 5
    Call GdipCreatePen1(DarkSlateGray, 2, UnitPixel, m_Pen)
    Call GdipDrawLine(m_Graphics, m_Pen, m_X + m_Interno * Cos(PI / 180 * m_Count), m_Y + m_Interno * Sin(PI / 180 * m_Count), m_Externo * Cos(PI / 180 * m_Count) + m_X, m_Externo * Sin(PI / 180 * m_Count) + m_Y)
    Next m_Count
   
    For m_Count = 135 To 405 Step 30
    x_Count = x_Count + 1
    Call GdipCreatePen1(Black, 4, UnitPixel, m_Pen)
    Call GdipDrawLine(m_Graphics, m_Pen, m_X + m_Interno * Cos(PI / 180 * m_Count), m_Y + m_Interno * Sin(PI / 180 * m_Count), m_Externo * Cos(PI / 180 * m_Count) + m_X, m_Externo * Sin(PI / 180 * m_Count) + m_Y)
    Me.CurrentX = m_X + 85 * Cos(PI / 180 * m_Count) - Me.TextWidth(x_Count) / 2 - 2
    Me.CurrentY = m_Y + 85 * Sin(PI / 180 * m_Count) - Me.TextHeight(x_Count) / 2
    Me.ForeColor = RGB(52, 52, 52)
    Me.FontName = "Arial"
    Me.FontSize = 9
    Me.FontBold = True
    Me.Print x_Count
    Next m_Count
    
    For m_Count = 135 To 405 Step 30
    x_Count = x_Count + 1
    Call GdipCreatePen1(Gold, 2, UnitPixel, m_Pen)
    Call GdipDrawLine(m_Graphics, m_Pen, m_X + m_Interno * Cos(PI / 180 * m_Count), m_Y + m_Interno * Sin(PI / 180 * m_Count), m_Externo * Cos(PI / 180 * m_Count) + m_X, m_Externo * Sin(PI / 180 * m_Count) + m_Y)
    Next m_Count
    
    m_Radio = 75
    m_Sin = m_X + m_Radio * Sin((0 - (((45 / 120) * HScroll1.Value) + 7.5) * 6) * m_Number)
    m_Cos = m_Y + m_Radio * Cos((0 - (((45 / 120) * HScroll1.Value) + 7.5) * 6) * m_Number)
    
    Call GdipCreatePen1(Black, 4, UnitPixel, m_Pen)
    Call GdipDrawLine(m_Graphics, m_Pen, m_X, m_Y, m_Sin, m_Cos)
    
    Call GdipCreatePen1(Gold, 2, UnitPixel, m_Pen)
    Call GdipDrawLine(m_Graphics, m_Pen, m_X, m_Y, m_Sin, m_Cos)
    
    Call GdipCreateSolidFill(Black, m_Brush)
    Call GdipFillEllipse(m_Graphics, m_Brush, m_X - 21, m_Y - 21, 42, 42)
    
    Call GdipCreateSolidFill(Gold, m_Brush)
    Call GdipFillEllipse(m_Graphics, m_Brush, m_X - 20, m_Y - 20, 40, 40)
    
    Call GdipDeleteBrush(m_Brush)
    Call GdipDeletePen(m_Pen)
    Call GdipDeleteGraphics(m_Graphics)

End Sub



