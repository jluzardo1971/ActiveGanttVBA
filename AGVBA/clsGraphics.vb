Option Explicit On 


Imports System.Drawing

Friend Class clsGraphics

    Private Structure T_PRECT
        Public lLeft As Integer
        Public lTop As Integer
        Public lRight As Integer
        Public lBottom As Integer
    End Structure

    Private mp_oControl As ActiveGanttVBACtl
    Private mp_udtPreviousClipRegion As T_PRECT
    Private mp_audtActiveReversibleFrames As System.Collections.ArrayList
    Private mp_audtActiveReversibleLinesStart As System.Collections.ArrayList
    Private mp_audtActiveReversibleLinesEnd As System.Collections.ArrayList
    Private mp_bCustomPrinting As Boolean
    Private mp_lCustomDC As Graphics
    Private mp_lPWidth As Integer
    Private mp_lPHeight As Integer
    Private mp_lFocusLeft As Integer
    Private mp_lFocusTop As Integer
    Private mp_lFocusRight As Integer
    Private mp_lFocusBottom As Integer
    Private mp_bEnableClipRegions As Boolean
    Friend mp_oToolTipGraphics As Graphics
    Friend bToolTipGraphics As Boolean

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Construction/Destruction & Initialization
    '// ---------------------------------------------------------------------------------------------------------------------

    Friend Sub New(ByVal Value As ActiveGanttVBACtl)
        mp_oControl = Value
        mp_audtActiveReversibleFrames = New System.Collections.ArrayList()
        mp_audtActiveReversibleLinesStart = New System.Collections.ArrayList()
        mp_audtActiveReversibleLinesEnd = New System.Collections.ArrayList()
        mp_bCustomPrinting = False
        mp_bEnableClipRegions = True
        bToolTipGraphics = False
    End Sub

    Friend Property EnableClipRegions() As Boolean
        Get
            Return mp_bEnableClipRegions
        End Get
        Set(ByVal Value As Boolean)
            mp_bEnableClipRegions = Value
        End Set
    End Property

    Friend Property f_FocusLeft() As Integer
        Get
            Return mp_lFocusLeft
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusLeft = Value
        End Set
    End Property

    Friend Property f_FocusTop() As Integer
        Get
            Return mp_lFocusTop
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusTop = Value
        End Set
    End Property

    Friend Property f_FocusRight() As Integer
        Get
            Return mp_lFocusRight
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusRight = Value
        End Set
    End Property

    Friend Property f_FocusBottom() As Integer
        Get
            Return mp_lFocusBottom
        End Get
        Set(ByVal Value As Integer)
            mp_lFocusBottom = Value
        End Set
    End Property

    Public ReadOnly Property oGraphics() As Graphics
        Get
            If mp_bCustomPrinting = False Then
                If bToolTipGraphics = False Then
                    Return mp_oControl.f_HDC()
                Else
                    Return mp_oToolTipGraphics
                End If
            Else
                Return mp_lCustomDC
            End If
        End Get
    End Property

    Public Property CustomPrinting() As Boolean
        Get
            Return mp_bCustomPrinting
        End Get
        Set(ByVal Value As Boolean)
            mp_bCustomPrinting = Value
        End Set
    End Property

    Public WriteOnly Property CustomDC() As Graphics
        Set(ByVal Value As Graphics)
            mp_lCustomDC = Value
        End Set
    End Property

    Public Function Width() As Integer
        If mp_bCustomPrinting = False Then
            Return mp_oControl.f_Width
        Else
            Return mp_lPWidth
        End If
    End Function

    Public Function Height() As Integer
        If mp_bCustomPrinting = False Then
            Return mp_oControl.f_Height
        Else
            Return mp_lPHeight
        End If
    End Function

    Public Property CustomWidth() As Integer
        Get
            Return mp_lPWidth
        End Get
        Set(ByVal Value As Integer)
            mp_lPWidth = Value
        End Set
    End Property

    Public Property CustomHeight() As Integer
        Get
            Return mp_lPHeight
        End Get
        Set(ByVal Value As Integer)
            mp_lPHeight = Value
        End Set
    End Property

    Public Sub DrawPolygon(ByVal v_lColor As Color, ByRef r_oPoints() As Point, ByVal v_Len As Integer)
        Dim hPen As Pen
        hPen = New Pen(v_lColor)
        oGraphics.DrawPolygon(hPen, r_oPoints)
    End Sub

    Public Function ConvertColor(ByVal dwOleColour As Color) As Integer
        '        Dim clrref As Integer
        '        OleTranslateColor(dwOleColour, 0, clrref)
        '        ConvertColor = clrref
    End Function

    Public Sub DrawEdge(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrBackColor As Color, ByVal v_yButtonStyle As GRE_BUTTONSTYLE, ByVal v_lEdgeType As GRE_EDGETYPE, ByVal v_bFilled As Boolean, ByVal oStyle As clsStyle)
        Dim lExteriorLeftTopColor As Color
        Dim lInteriorLeftTopColor As Color
        Dim lExteriorRightBottomColor As Color
        Dim lInteriorRightBottomColor As Color
        If v_yButtonStyle = GRE_BUTTONSTYLE.BT_NORMALWINDOWS Then
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.FromArgb(255, 240, 240, 240)
                        lInteriorLeftTopColor = Color.FromArgb(255, 192, 192, 192)
                        lInteriorRightBottomColor = Color.Gray
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.Gray
                        lInteriorLeftTopColor = Color.FromArgb(255, 64, 64, 64)
                        lInteriorRightBottomColor = Color.FromArgb(255, 192, 192, 192)
                        lExteriorRightBottomColor = Color.FromArgb(255, 240, 240, 240)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lInteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenInteriorLeftTopColor
                        lInteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenInteriorRightBottomColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            '// Exterior Left
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Top
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Right
            DrawLine(v_X2, v_Y2, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Exterior Bottom
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Left
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X1 + 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Top
            DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Right
            DrawLine(v_X2 - 1, v_Y2 - 1, v_X2 - 1, v_Y1 + 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            '// Interior Bottom
            DrawLine(v_X1 + 1, v_Y2 - 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_NORMAL, lInteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 2, v_Y1 + 2, v_X2 - 2, v_Y2 - 2, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        Else
            Select Case v_lEdgeType
                Case GRE_EDGETYPE.ET_RAISED
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.White
                        lExteriorRightBottomColor = Color.FromArgb(255, 64, 64, 64)
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.RaisedExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.RaisedExteriorRightBottomColor
                    End If
                Case GRE_EDGETYPE.ET_SUNKEN
                    If oStyle Is Nothing Then
                        lExteriorLeftTopColor = Color.Gray
                        lExteriorRightBottomColor = Color.WhiteSmoke
                    Else
                        lExteriorLeftTopColor = oStyle.ButtonBorderStyle.SunkenExteriorLeftTopColor
                        lExteriorRightBottomColor = oStyle.ButtonBorderStyle.SunkenExteriorRightBottomColor
                    End If
            End Select
            DrawLine(v_X1, v_Y1, v_X2, v_Y1, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y1, v_X1, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorLeftTopColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X1, v_Y2, v_X2, v_Y2, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            DrawLine(v_X2, v_Y2, v_X2, v_Y1 - 1, GRE_LINETYPE.LT_NORMAL, lExteriorRightBottomColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            If v_bFilled = True Then
                DrawLine(v_X1 + 1, v_Y1 + 1, v_X2 - 1, v_Y2 - 1, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            End If
        End If
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, v_lColor, v_lDrawStyle, 1, True)
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer)
        DrawLine(v_X1, v_Y1, v_X2, v_Y2, v_yStyle, v_lColor, v_lDrawStyle, v_lWidth, True)
    End Sub

    Public Sub DrawLine(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_yStyle As GRE_LINETYPE, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByVal v_lWidth As Integer, ByVal v_bCreatePens As Boolean)
        Dim mp_ucPen As New Pen(v_lColor, v_lWidth)
        Dim mp_ucBrush As New SolidBrush(v_lColor)
        Select Case v_lDrawStyle
            Case GRE_LINEDRAWSTYLE.LDS_SOLID
                mp_ucPen.DashStyle = Drawing.Drawing2D.DashStyle.Solid
            Case GRE_LINEDRAWSTYLE.LDS_DOT
                mp_ucPen.DashStyle = Drawing.Drawing2D.DashStyle.Dot
        End Select
        Select Case v_yStyle
            Case GRE_LINETYPE.LT_NORMAL
                Dim Points(1) As Point
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2
                Points(1).Y = v_Y2
                oGraphics.DrawPolygon(mp_ucPen, Points)
            Case GRE_LINETYPE.LT_BORDER
                Dim Points(4) As Point
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2
                Points(1).Y = v_Y1
                Points(2).X = v_X2
                Points(2).Y = v_Y2
                Points(3).X = v_X1
                Points(3).Y = v_Y2
                Points(4).X = v_X1
                Points(4).Y = v_Y1
                oGraphics.DrawPolygon(mp_ucPen, Points)
            Case GRE_LINETYPE.LT_FILLED
                Dim Points(4) As Point
                Points(0).X = v_X1
                Points(0).Y = v_Y1
                Points(1).X = v_X2 + 1
                Points(1).Y = v_Y1
                Points(2).X = v_X2 + 1
                Points(2).Y = v_Y2 + 1
                Points(3).X = v_X1
                Points(3).Y = v_Y2 + 1
                Points(4).X = v_X1
                Points(4).Y = v_Y1
                oGraphics.FillPolygon(mp_ucBrush, Points)
        End Select
    End Sub

    Public Sub DrawFigure(ByVal v_X As Integer, ByVal v_Y As Integer, ByVal v_dx As Integer, ByVal v_dy As Integer, ByVal v_yFigureType As GRE_FIGURETYPE, ByVal v_lBorderColor As Color, ByVal v_lFillColor As Color, ByVal v_yBorderStyle As GRE_LINEDRAWSTYLE)
        Dim oPen As New Pen(v_lBorderColor)
        Dim oBrush As New SolidBrush(v_lFillColor)
        If v_dx Mod 2 <> 0 Then
            v_dx = v_dx + 1
            v_dy = v_dy + 1
        End If
        Select Case v_yFigureType
            Case GRE_FIGURETYPE.FT_PROJECTUP
                Dim Points(4) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 2
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 2
                Points(4).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_PROJECTDOWN
                Dim Points(4) As PointF
                Points(0).X = v_X + v_dx / 2
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X - v_dx / 2
                Points(4).Y = v_Y
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_DIAMOND
                Dim Points(3) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEDIAMOND
                Dim Points(3) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X
                Points(2).Y = v_Y + (3 * v_dy) / 4
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy / 2
                oGraphics.DrawEllipse(oPen, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEUP
                Dim Points(2) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X - v_dx / 2
                Points(2).Y = v_Y + v_dy
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLEDOWN
                Dim Points(2) As PointF
                Points(0).X = v_X + v_dx / 2
                Points(0).Y = v_Y
                Points(1).X = v_X - v_dx / 2
                Points(1).Y = v_Y
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLERIGHT
                Dim Points(2) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X + v_dx
                Points(2).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_TRIANGLELEFT
                Dim Points(2) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X
                Points(1).Y = v_Y + v_dy
                Points(2).X = v_X - v_dx
                Points(2).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEUP
                Dim Points(2) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + (3 * v_dy) / 4
                Points(2).X = v_X - v_dx / 4
                Points(2).Y = v_Y + (3 * v_dy) / 4
                oGraphics.DrawEllipse(oPen, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLETRIANGLEDOWN
                Dim Points(2) As PointF
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X
                Points(2).Y = v_Y + (3 * v_dy) / 4
                oGraphics.DrawEllipse(oPen, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_ARROWUP
                Dim Points(6) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 2
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 4
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y + v_dy
                Points(5).X = v_X - v_dx / 4
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 2
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_ARROWDOWN
                Dim Points(6) As PointF
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 2
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X
                Points(4).Y = v_Y + v_dy
                Points(5).X = v_X - v_dx / 2
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 4
                Points(6).Y = v_Y + v_dy / 2
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWUP
                Dim Points(6) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 2
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 8
                Points(3).Y = v_Y + (3 * v_dy) / 4
                Points(4).X = v_X - v_dx / 8
                Points(4).Y = v_Y + (3 * v_dy) / 4
                Points(5).X = v_X - v_dx / 8
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 4
                Points(6).Y = v_Y + v_dy / 2
                oGraphics.DrawEllipse(oPen, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLEARROWDOWN
                Dim Points(6) As PointF
                Points(0).X = v_X - v_dx / 8
                Points(0).Y = v_Y + v_dy / 4
                Points(1).X = v_X + v_dx / 8
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X + v_dx / 4
                Points(3).Y = v_Y + v_dy / 2
                Points(4).X = v_X
                Points(4).Y = v_Y + (3 * v_dy) / 4
                Points(5).X = v_X - v_dx / 4
                Points(5).Y = v_Y + v_dy / 2
                Points(6).X = v_X - v_dx / 8
                Points(6).Y = v_Y + v_dy / 2
                oGraphics.DrawEllipse(oPen, mp_oControl.MathLib.RoundDouble(v_X - v_dx / 2), v_Y, v_dx, v_dy)
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTUP
                Dim Points(4) As PointF
                Points(0).X = v_X
                Points(0).Y = v_Y + v_dy / 2
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + (3 * v_dy) / 4
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y + (3 * v_dy) / 4
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_SMALLPROJECTDOWN
                Dim Points(4) As PointF
                Points(0).X = v_X + v_dx / 4
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dy / 4
                Points(2).X = v_X
                Points(2).Y = v_Y + v_dy / 2
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + v_dy / 4
                Points(4).X = v_X - v_dx / 4
                Points(4).Y = v_Y
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_RECTANGLE
                Dim Points(3) As PointF
                Points(0).X = v_X - v_dx / 8
                Points(0).Y = v_Y
                Points(1).X = v_X + v_dx / 8
                Points(1).Y = v_Y
                Points(2).X = v_X + v_dx / 8
                Points(2).Y = v_Y + v_dy
                Points(3).X = v_X - v_dx / 8
                Points(3).Y = v_Y + v_dy
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_SQUARE
                Dim Points(3) As PointF
                Points(0).X = v_X - v_dx / 4
                Points(0).Y = v_Y + v_dx / 4
                Points(1).X = v_X + v_dx / 4
                Points(1).Y = v_Y + v_dx / 4
                Points(2).X = v_X + v_dx / 4
                Points(2).Y = v_Y + (3 * v_dy) / 4
                Points(3).X = v_X - v_dx / 4
                Points(3).Y = v_Y + (3 * v_dy) / 4
                mp_DrawFigureAux(oBrush, oPen, Points)
            Case GRE_FIGURETYPE.FT_CIRCLE
                oGraphics.FillEllipse(oBrush, CSng(v_X - v_dx / 2), CSng(v_Y), CSng(v_dx), CSng(v_dy))
            Case Else
                Return
        End Select

    End Sub

    Private Sub mp_DrawFigureAux(ByRef oBrush As SolidBrush, ByRef oPen As Pen, ByRef oPoints() As PointF)
        oGraphics.FillPolygon(oBrush, oPoints)
        oGraphics.DrawPolygon(oPen, oPoints)
    End Sub

    Public Sub DrawPattern(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_lColor As Color, ByVal v_lDrawStyle As GRE_PATTERN, ByVal v_iPatternFactor As Integer)
        Dim tmp As Integer
        Dim c As Integer
        Dim c1 As Integer
        Dim c2 As Integer
        Dim i1 As Integer
        Dim j1 As Integer
        Dim i2 As Integer
        Dim j2 As Integer
        If v_X1 > v_X2 Then
            tmp = v_X1
            v_X1 = v_X2
            v_X2 = tmp
        End If
        If v_Y1 > v_Y2 Then
            tmp = v_Y1
            v_Y1 = v_Y2
            v_Y2 = tmp
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_HORIZONTALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_Y1 + v_iPatternFactor) To v_Y2 Step v_iPatternFactor
                DrawLine(v_X1, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_VERTICALLINE Or v_lDrawStyle = GRE_PATTERN.FP_CROSS Then
            For j1 = (v_X1 + v_iPatternFactor) To v_X2 Step v_iPatternFactor
                DrawLine(j1, v_Y1, j1, v_Y2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_UPWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = Int((v_Y1 + v_X1) / v_iPatternFactor + 1)
            c2 = Int((v_Y2 + v_X2) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = c * v_iPatternFactor - i1
                j2 = c * v_iPatternFactor - i2
                If j2 < v_Y1 Then
                    i2 = c * v_iPatternFactor - v_Y1
                    j2 = c * v_iPatternFactor - i2
                End If
                If j1 > v_Y2 Then
                    i1 = c * v_iPatternFactor - v_Y2
                    j1 = c * v_iPatternFactor - i1
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DOWNWARDDIAGONAL Or v_lDrawStyle = GRE_PATTERN.FP_DIAGONALCROSS Then
            c1 = Int((v_Y1 - v_X2) / v_iPatternFactor + 1)
            c2 = Int((v_Y2 - v_X1) / v_iPatternFactor)
            For c = c1 To c2
                i1 = v_X1
                i2 = v_X2
                j1 = i1 + c * v_iPatternFactor
                j2 = i2 + c * v_iPatternFactor
                If j1 < v_Y1 Then
                    i1 = v_Y1 - c * v_iPatternFactor
                    j1 = i1 + c * v_iPatternFactor
                End If
                If j2 > v_Y2 Then
                    i2 = v_Y2 - c * v_iPatternFactor
                    j2 = i2 + c * v_iPatternFactor
                End If
                DrawLine(i1, j1, i2, j2, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID, 1, False)
            Next c
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_LIGHT Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_MEDIUM Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                Else
                    For j2 = (v_X1 + 2) To (v_X2 - 1) Step 2
                        DrawLine(j2, j1, j2 + 1, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    Next j2
                End If
            Next j1
        End If
        If v_lDrawStyle = GRE_PATTERN.FP_DARK Then
            For j1 = (v_Y1 + 1) To (v_Y2 - 1)
                If j1 Mod 2 = 0 Then
                    For j2 = (v_X1 + 1) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                Else
                    DrawLine(v_X1, j1, v_X1 + 2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                    For j2 = (v_X1 + 3) To (v_X2 - 1) Step 4
                        If j2 + 3 < v_X2 Then
                            DrawLine(j2, j1, j2 + 3, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        Else
                            DrawLine(j2, j1, v_X2, j1, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                        End If
                    Next j2
                End If
            Next j1
        End If
    End Sub

    Public Sub DrawPolyLine(ByVal v_lColor As Color, ByVal v_lWidth As Integer, ByVal v_lDrawStyle As GRE_LINEDRAWSTYLE, ByRef r_oPoints() As Point, ByVal v_Len As Integer)
        Dim hPen As Pen
        hPen = New Pen(v_lColor, v_lWidth)
        oGraphics.DrawLines(hPen, r_oPoints)
    End Sub

    Public Sub DrawTextEx(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_sParam As String, ByVal v_lFlags As StringFormat, ByVal v_lColor As Color, ByVal v_oFont As Font, Optional ByVal v_bClip As Boolean = True)
        Dim mp_ucBrush As New SolidBrush(v_lColor)
        Dim oRect As New RectangleF()
        oRect.X = v_X1
        oRect.Y = v_Y1
        oRect.Width = v_X2 - v_X1
        oRect.Height = v_Y2 - v_Y1
        oGraphics.DrawString(v_sParam, v_oFont, mp_ucBrush, oRect, v_lFlags)
    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal v_lColor As Color, ByVal v_oFont As Font)
        DrawAlignedText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_sParam, v_yHPos, v_yVPos, v_lColor, v_oFont, True)
    End Sub

    Public Sub DrawAlignedText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_sParam As String, ByVal v_yHPos As GRE_HORIZONTALALIGNMENT, ByVal v_yVPos As GRE_VERTICALALIGNMENT, ByVal v_lColor As Color, ByVal v_oFont As Font, ByVal v_bClip As Boolean)
        Dim lHeight As Integer
        Dim lWidth As Integer
        Dim mp_ucBrush As New SolidBrush(v_lColor)
        lHeight = mp_oControl.mp_lStrHeight(v_sParam, v_oFont)
        lWidth = mp_oControl.mp_lStrWidth(v_sParam, v_oFont)
        Dim oRect As New RectangleF()
        Dim oFormat As New StringFormat()
        If v_bClip = False And lWidth > (v_lRight - v_lLeft) Then
            Return
        End If
        If v_bClip = False And lHeight > (v_lBottom - v_lTop) Then
            Return
        End If
        oRect.X = v_lLeft
        oRect.Y = v_lTop
        oRect.Width = v_lRight - v_lLeft
        oRect.Height = v_lBottom - v_lTop
        Select Case v_yHPos
            Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                oFormat.Alignment = StringAlignment.Near
            Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                oFormat.Alignment = StringAlignment.Center
            Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                oFormat.Alignment = StringAlignment.Far
        End Select
        Select Case v_yVPos
            Case GRE_VERTICALALIGNMENT.VAL_TOP
                oFormat.LineAlignment = StringAlignment.Near
            Case GRE_VERTICALALIGNMENT.VAL_CENTER
                oFormat.LineAlignment = StringAlignment.Center
            Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                oFormat.LineAlignment = StringAlignment.Far
        End Select
        oGraphics.DrawString(v_sParam, v_oFont, mp_ucBrush, oRect, oFormat)
    End Sub

    Public Sub ClipRegion(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bStore As Boolean)
        If (mp_bEnableClipRegions = False) Then
            Return
        End If
        Dim oRect As New Rectangle(v_X1, v_Y1, (v_X2 - v_X1 + 1), (v_Y2 - v_Y1 + 1))
        Dim oRegion As New Region(oRect)
        If v_bStore = True Then
            mp_udtPreviousClipRegion.lLeft = v_X1
            mp_udtPreviousClipRegion.lRight = v_X2
            mp_udtPreviousClipRegion.lTop = v_Y1
            mp_udtPreviousClipRegion.lBottom = v_Y2
        End If
        oGraphics.Clip = oRegion
    End Sub

    Public Sub RestorePreviousClipRegion()
        If (mp_bEnableClipRegions = False) Then
            Return
        End If
        ClipRegion(mp_udtPreviousClipRegion.lLeft, mp_udtPreviousClipRegion.lTop, mp_udtPreviousClipRegion.lRight, mp_udtPreviousClipRegion.lBottom, False)
    End Sub

    Public Sub ClearClipRegion()
        oGraphics.ResetClip()
    End Sub

    Public Sub TileImageHorizontal(ByVal ImageHandle As Image, ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal v_bTransparent As Boolean)
        Dim X As Integer
        Dim lImageWidth As Integer
        Dim lImageHeight As Integer
        lImageHeight = ImageHandle.Height
        lImageWidth = ImageHandle.Width
        Do While X < (v_X2 - v_X1)
            If (X + lImageWidth) > (v_X2 - v_X1) Then
                PaintImage(ImageHandle, v_X2 - lImageWidth, v_Y1, v_X2, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            Else
                PaintImage(ImageHandle, v_X1 + X, v_Y1, v_X1 + X + lImageWidth, v_Y1 + lImageHeight, 0, 0, v_bTransparent)
            End If
            X = X + lImageWidth
        Loop
    End Sub

    Public Sub PaintImage(ByVal Image As Image, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal xOrigin As Integer, ByVal yOrigin As Integer, ByVal bUseMask As Boolean)
        Dim DestRect As RectangleF
        Dim OriginRect As RectangleF
        OriginRect.X = xOrigin
        OriginRect.Y = yOrigin
        OriginRect.Width = Image.Width - xOrigin
        OriginRect.Height = Image.Height - yOrigin
        DestRect.X = X1
        DestRect.Y = Y1
        DestRect.Width = X2 - X1
        DestRect.Height = Y2 - Y1
        If bUseMask = False Then
            oGraphics.DrawImage(Image, DestRect, OriginRect, Drawing.GraphicsUnit.Pixel)
        Else
            Dim oBitmap As Bitmap
            oBitmap = Image.Clone()
            oBitmap.MakeTransparent(Color.White)
            oGraphics.DrawImage(oBitmap, DestRect, OriginRect, Drawing.GraphicsUnit.Pixel)
        End If
    End Sub

    Public Sub DrawImage(ByRef v_oImage As Image, ByRef v_yHorizontalAlignment As GRE_HORIZONTALALIGNMENT, ByRef v_yVerticalAlignment As GRE_VERTICALALIGNMENT, ByVal v_lImageXMargin As Integer, ByVal v_lImageYMargin As Integer, ByRef v_lLeft As Integer, ByRef v_lRight As Integer, ByRef v_lTop As Integer, ByRef v_lBottom As Integer, ByVal v_bTransparent As Boolean)
        Dim bDrawImage As Boolean
        Dim bHorizontalSmall As Boolean
        Dim bVerticalSmall As Boolean
        Dim XOrigin As Integer
        Dim YOrigin As Integer
        Dim xDest As Integer
        Dim yDest As Integer
        Dim lxWidth As Integer
        Dim lyHeight As Integer
        Dim lImageHeight As Integer
        Dim lImageWidth As Integer
        If (v_oImage Is Nothing) Then
            Return
        End If
        lImageHeight = v_oImage.Height
        lImageWidth = v_oImage.Width
        If v_yHorizontalAlignment = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
            v_lImageXMargin = 0
        End If
        If v_yVerticalAlignment = GRE_VERTICALALIGNMENT.VAL_CENTER Then
            v_lImageYMargin = 0
        End If
        bDrawImage = True
        If (v_lRight - v_lLeft) < (lImageWidth + v_lImageXMargin) Then
            lxWidth = v_lRight - v_lLeft - v_lImageXMargin
            If lxWidth <= 0 Then bDrawImage = False
            bHorizontalSmall = True
        Else
            lxWidth = lImageWidth
            bHorizontalSmall = False
        End If
        If (v_lBottom - v_lTop) < (lImageHeight + v_lImageYMargin) Then
            lyHeight = v_lBottom - v_lTop - v_lImageYMargin
            If lyHeight <= 0 Then bDrawImage = False
            bVerticalSmall = True
        Else
            lyHeight = lImageHeight
            bVerticalSmall = False
        End If
        If bHorizontalSmall = False Then
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    xDest = ((v_lRight - v_lLeft) - lImageWidth) / 2 + v_lLeft
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    xDest = v_lRight - lImageWidth - v_lImageXMargin
            End Select
            XOrigin = 0
        Else
            Select Case v_yHorizontalAlignment
                Case GRE_HORIZONTALALIGNMENT.HAL_LEFT
                    XOrigin = 0
                    xDest = v_lLeft + v_lImageXMargin
                Case GRE_HORIZONTALALIGNMENT.HAL_CENTER
                    XOrigin = (lImageWidth - lxWidth) / 2
                    xDest = v_lLeft
                Case GRE_HORIZONTALALIGNMENT.HAL_RIGHT
                    XOrigin = lImageWidth - lxWidth
                    xDest = v_lRight - lxWidth - v_lImageXMargin
            End Select
        End If
        If bVerticalSmall = False Then
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    yDest = ((v_lBottom - v_lTop) - lImageHeight) / 2 + v_lTop
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    yDest = v_lBottom - lImageHeight - v_lImageYMargin
            End Select
            YOrigin = 0
        Else
            Select Case v_yVerticalAlignment
                Case GRE_VERTICALALIGNMENT.VAL_TOP
                    YOrigin = 0
                    yDest = v_lTop + v_lImageYMargin
                Case GRE_VERTICALALIGNMENT.VAL_CENTER
                    YOrigin = (lImageHeight - lyHeight) / 2
                    yDest = v_lTop
                Case GRE_VERTICALALIGNMENT.VAL_BOTTOM
                    YOrigin = lImageHeight - lyHeight
                    yDest = v_lBottom - lyHeight - v_lImageYMargin
            End Select
        End If
        If bDrawImage = True Then
            PaintImage(v_oImage, xDest, yDest, xDest + lxWidth, yDest + lyHeight, XOrigin, YOrigin, v_bTransparent)
        End If
    End Sub

    Public Sub DrawFocusRectangle(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer)
        Dim myrect As New Rectangle(v_X1, v_Y1, v_X2 - v_X1, v_Y2 - v_Y1)
        System.Windows.Forms.ControlPaint.DrawFocusRectangle(oGraphics, myrect)
    End Sub

    Public Sub GradientFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrStartColor As Color, ByVal clrEndColor As Color, ByVal iGradientType As GRE_GRADIENTFILLMODE)
        If (v_X2 - v_X1) <= 0 Then
            Return
        End If
        If (v_Y2 - v_Y1) <= 0 Then
            Return
        End If
        Dim oRect As New RectangleF(v_X1, v_Y1, v_X2 - v_X1, v_Y2 - v_Y1)
        Dim mp_ucBrush As Drawing2D.LinearGradientBrush = Nothing
        If (iGradientType = GRE_GRADIENTFILLMODE.GDT_VERTICAL) Then
            mp_ucBrush = New Drawing2D.LinearGradientBrush(oRect, clrStartColor, clrEndColor, Drawing2D.LinearGradientMode.Vertical)
        ElseIf (iGradientType = GRE_GRADIENTFILLMODE.GDT_HORIZONTAL) Then
            mp_ucBrush = New Drawing2D.LinearGradientBrush(oRect, clrStartColor, clrEndColor, Drawing2D.LinearGradientMode.Horizontal)
        End If
        Dim Points(4) As Point
        Points(0).X = v_X1
        Points(0).Y = v_Y1
        Points(1).X = v_X2 + 1
        Points(1).Y = v_Y1
        Points(2).X = v_X2 + 1
        Points(2).Y = v_Y2 + 1
        Points(3).X = v_X1
        Points(3).Y = v_Y2 + 1
        Points(4).X = v_X1
        Points(4).Y = v_Y1
        oGraphics.FillPolygon(mp_ucBrush, Points)
    End Sub

    Public Sub HatchFill(ByVal v_X1 As Integer, ByVal v_Y1 As Integer, ByVal v_X2 As Integer, ByVal v_Y2 As Integer, ByVal clrForeColor As Color, ByVal clrBackColor As Color, ByVal yHatchStyle As GRE_HATCHSTYLE)
        Dim mp_ucBrush As Drawing2D.HatchBrush
        mp_ucBrush = New Drawing2D.HatchBrush(yHatchStyle, clrForeColor, clrBackColor)
        Dim Points(4) As Point
        Points(0).X = v_X1
        Points(0).Y = v_Y1
        Points(1).X = v_X2 + 1
        Points(1).Y = v_Y1
        Points(2).X = v_X2 + 1
        Points(2).Y = v_Y2 + 1
        Points(3).X = v_X1
        Points(3).Y = v_Y2 + 1
        Points(4).X = v_X1
        Points(4).Y = v_Y1
        oGraphics.FillPolygon(mp_ucBrush, Points)
    End Sub


    Public Function ConvertColor(ByVal dwOleColor As Integer) As Color
        Return ColorTranslator.FromOle(dwOleColor)
    End Function

    Public Function LineIntersection(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Boolean
        Return True
    End Function

    Friend Sub mp_DrawItemI(ByRef oTask As clsTask, ByVal sStyleIndex As String, ByVal Selected As Boolean, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        Dim oMilestoneStyle As clsMilestoneStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItemI")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItemI")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_FLAT
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oStyle.MilestoneStyle.Image Is Nothing Then

                Else
                    DrawImage(oStyle.MilestoneStyle.Image, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, oTask.Left, oTask.Right, oTask.Top, oTask.Bottom, oStyle.UseMask)
                End If
            Case Else
                oMilestoneStyle = oStyle.MilestoneStyle
                DrawFigure(mp_oControl.MathLib.GetXCoordinateFromDate(oTask.StartDate), oTask.Top, oTask.Bottom - oTask.Top, oTask.Bottom - oTask.Top, oMilestoneStyle.ShapeIndex, oMilestoneStyle.BorderColor, oMilestoneStyle.FillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
        End Select
        mp_DrawItemText(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, oTask.LeftTrim, oTask.RightTrim, oStyle, oTask.Text)
        If oStyle.SelectionRectangleStyle.Visible = True And Selected Then
            If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
                DrawFocusRectangle(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom)
            ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
                DrawLine(oTask.Left, oTask.Top, oTask.Right, oTask.Bottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
            End If
        End If
    End Sub

    Friend Sub mp_DrawItemEx(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle, ByVal clrBackColor As Color, ByVal clrForeColor As Color, ByVal clrStartGradientColor As Color, ByVal clrEndGradientColor As Color, ByVal clrHatchBackColor As Color, ByVal clrHatchForeColor As Color)
        Dim oStyle As clsStyle
        Dim oTaskStyle As clsTaskStyle
        If (v_oStyle Is Nothing) Then
            mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_NULL, "Style object is null when preparing to draw.", "mp_DrawItemEx")
            Return
        Else
            oStyle = v_oStyle
        End If
        oTaskStyle = oStyle.TaskStyle
        Select Case oStyle.Appearance
            Case E_STYLEAPPEARANCE.SA_RAISED
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_RAISED, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_SUNKEN
                DrawEdge(v_lLeft, v_lTop, v_lRight, v_lBottom, clrBackColor, oStyle.ButtonStyle, GRE_EDGETYPE.ET_SUNKEN, True, v_oStyle)
            Case E_STYLEAPPEARANCE.SA_FLAT
                Dim lTop As Integer
                Dim lBottom As Integer
                lTop = v_lTop
                lBottom = v_lBottom
                Select Case oStyle.FillMode
                    Case GRE_FILLMODE.FM_COMPLETELYFILLED
                    Case GRE_FILLMODE.FM_UPPERHALFFILLED
                        lBottom = v_lTop + ((v_lBottom - v_lTop) / 2)
                    Case GRE_FILLMODE.FM_LOWERHALFFILLED
                        lTop = v_lBottom - ((v_lBottom - v_lTop) / 2)
                End Select
                If (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_SOLID) Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_GRADIENT) Then
                    GradientFill(v_lLeft, lTop, v_lRight, lBottom, clrStartGradientColor, clrEndGradientColor, oStyle.GradientFillMode)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_PATTERN) Then
                    DrawPattern(v_lLeft, lTop, v_lRight, lBottom, clrBackColor, oStyle.Pattern, oStyle.PatternFactor)
                ElseIf (oStyle.BackgroundMode = GRE_BACKGROUNDMODE.FP_HATCH) Then
                    HatchFill(v_lLeft, lTop, v_lRight, lBottom, clrHatchForeColor, clrHatchBackColor, oStyle.HatchStyle)
                End If
                If oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_SINGLE Then
                    DrawLine(v_lLeft, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_BORDER, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                ElseIf oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_CUSTOM Then
                    If oStyle.CustomBorderStyle.Left = True Then
                        DrawLine(v_lLeft, lTop, v_lLeft, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Top = True Then
                        DrawLine(v_lLeft, lTop, v_lRight, lTop, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Right = True Then
                        DrawLine(v_lRight, lTop, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                    If oStyle.CustomBorderStyle.Bottom = True Then
                        DrawLine(v_lLeft, lBottom, v_lRight, lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
                    End If
                End If
                DrawFigure(v_lRight, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.EndShapeIndex, oTaskStyle.EndBorderColor, oTaskStyle.EndFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawFigure(v_lLeft, v_lTop, v_lBottom - v_lTop, v_lBottom - v_lTop, oTaskStyle.StartShapeIndex, oTaskStyle.StartBorderColor, oTaskStyle.StartFillColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
            Case E_STYLEAPPEARANCE.SA_CELL
                DrawLine(v_lLeft, v_lTop, v_lRight, v_lBottom, GRE_LINETYPE.LT_FILLED, clrBackColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                DrawLine(v_lLeft, v_lBottom, v_lRight, v_lBottom, GRE_LINETYPE.LT_NORMAL, oStyle.BorderColor, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.BorderWidth)
            Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                If oTaskStyle.MiddleImage Is Nothing Or oTaskStyle.StartImage Is Nothing Or oTaskStyle.EndImage Is Nothing Then

                Else
                    Dim lImageHeight As Integer
                    Dim lImageWidth As Integer
                    lImageHeight = oTaskStyle.MiddleImage.Height
                    lImageWidth = oTaskStyle.MiddleImage.Width
                    TileImageHorizontal(oTaskStyle.MiddleImage, v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle.UseMask)
                    '// Exit if the start and end sections don't fit
                    If (v_lRight - v_lLeft) > (lImageWidth * 2) Then
                        '// Left Section
                        PaintImage(oTaskStyle.StartImage, v_lLeft, v_lTop, v_lLeft + lImageWidth, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                        '// Right Section
                        PaintImage(oTaskStyle.EndImage, v_lRight - lImageWidth, v_lTop, v_lRight, v_lTop + lImageHeight, 0, 0, oStyle.UseMask)
                    End If
                End If
        End Select
        If Not (v_oImage Is Nothing) Then
            DrawImage(v_oImage, oStyle.ImageAlignmentHorizontal, oStyle.ImageAlignmentVertical, oStyle.ImageXMargin, oStyle.ImageYMargin, v_lLeft, v_lRight, v_lTop, v_lBottom, oStyle.UseMask)
        End If
        mp_DrawItemText(v_lLeft, v_lTop, v_lRight, v_lBottom, v_lLeftTrim, v_lRightTrim, oStyle, sText)
        If oStyle.SelectionRectangleStyle.Visible = True And v_bIsSelected Then
            mp_DrawSelectionRectangle(v_lLeft, v_lTop, v_lRight, v_lBottom, oStyle)
        End If
    End Sub

    Friend Sub mp_DrawSelectionRectangle(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal oStyle As clsStyle)
        If oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_DOTTED Then
            DrawFocusRectangle(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom)
        ElseIf oStyle.SelectionRectangleStyle.Mode = E_SELECTIONRECTANGLEMODE.SRM_COLOR Then
            DrawLine(v_lLeft + oStyle.SelectionRectangleStyle.OffsetLeft, v_lTop + oStyle.SelectionRectangleStyle.OffsetTop, v_lRight - oStyle.SelectionRectangleStyle.OffsetRight, v_lBottom - oStyle.SelectionRectangleStyle.OffsetBottom, GRE_LINETYPE.LT_BORDER, oStyle.SelectionRectangleStyle.Color, GRE_LINEDRAWSTYLE.LDS_SOLID, oStyle.SelectionRectangleStyle.BorderWidth)
        End If
    End Sub

    Friend Sub mp_DrawItem(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal sStyleIndex As String, ByVal sText As String, ByVal v_bIsSelected As Boolean, ByRef v_oImage As Image, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef v_oStyle As clsStyle)
        Dim oStyle As clsStyle
        If (v_oStyle Is Nothing) Then
            If mp_oControl.StrLib.StrIsNumeric(sStyleIndex) Then
                If mp_oControl.StrLib.StrCLng(sStyleIndex) < 0 Or mp_oControl.StrLib.StrCLng(sStyleIndex) > mp_oControl.Styles.Count Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_INDEX, "Style object element not found when preparing to draw, invalid index", "mp_DrawItem")
                    Return
                End If
            Else
                If mp_oControl.Styles.oCollection.m_bDoesKeyExist(sStyleIndex) = False Then
                    mp_oControl.mp_ErrorReport(SYS_ERRORS.STYLE_INVALID_KEY, "Style object element not found when preparing to draw, invalid key (""" & sStyleIndex & """)", "mp_DrawItem")
                    Return
                End If
            End If
            oStyle = mp_oControl.Styles.FItem(sStyleIndex)
        Else
            oStyle = v_oStyle
        End If
        mp_DrawItemEx(v_lLeft, v_lTop, v_lRight, v_lBottom, sText, v_bIsSelected, v_oImage, v_lLeftTrim, v_lRightTrim, oStyle, oStyle.BackColor, oStyle.ForeColor, oStyle.StartGradientColor, oStyle.EndGradientColor, oStyle.HatchBackColor, oStyle.HatchForeColor)
    End Sub

    Private Sub mp_DrawItemText(ByVal v_lLeft As Integer, ByVal v_lTop As Integer, ByVal v_lRight As Integer, ByVal v_lBottom As Integer, ByVal v_lLeftTrim As Integer, ByVal v_lRightTrim As Integer, ByRef oStyle As clsStyle, ByVal sText As String)
        Dim lTextLeft As Integer
        Dim lTextRight As Integer
        Dim lTextTop As Integer
        Dim lTextBottom As Integer
        If oStyle.TextVisible = False Then
            Return
        End If
        If sText = "" Then
            Return
        End If
        Select Case oStyle.TextPlacement
            Case E_TEXTPLACEMENT.SCP_OBJECTEXTENTSPLACEMENT
                If (oStyle.DrawTextInVisibleArea = False) Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight
                Else
                    lTextLeft = v_lLeftTrim
                    lTextRight = v_lRightTrim
                End If
                lTextTop = v_lTop
                lTextBottom = v_lBottom
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft + oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextRight = v_lRight - oStyle.TextXMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop + oStyle.TextYMargin
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextBottom = v_lBottom - oStyle.TextYMargin
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, oStyle.TextAlignmentHorizontal, oStyle.TextAlignmentVertical, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_OFFSETPLACEMENT
                DrawTextEx(v_lLeft + oStyle.TextFlags.OffsetLeft, v_lTop + oStyle.TextFlags.OffsetTop, v_lRight - oStyle.TextFlags.OffsetRight, v_lBottom - oStyle.TextFlags.OffsetBottom, sText, oStyle.TextFlags.GetFlags, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
            Case E_TEXTPLACEMENT.SCP_EXTERIORPLACEMENT
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_LEFT Then
                    lTextLeft = v_lLeft - mp_oControl.mp_lStrWidth(sText, oStyle.Font) - oStyle.TextXMargin
                    lTextRight = v_lLeft - oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_RIGHT Then
                    lTextLeft = v_lRight + oStyle.TextXMargin
                    lTextRight = v_lRight + mp_oControl.mp_lStrWidth(sText, oStyle.Font) + oStyle.TextXMargin + 1
                End If
                If oStyle.TextAlignmentHorizontal = GRE_HORIZONTALALIGNMENT.HAL_CENTER Then
                    lTextLeft = v_lLeft
                    lTextRight = v_lRight + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_TOP Then
                    lTextTop = v_lTop - mp_oControl.mp_lStrHeight(sText, oStyle.Font) - oStyle.TextYMargin
                    lTextBottom = v_lTop - oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_BOTTOM Then
                    lTextTop = v_lBottom + oStyle.TextYMargin
                    lTextBottom = v_lBottom + mp_oControl.mp_lStrHeight(sText, oStyle.Font) + oStyle.TextYMargin + 1
                End If
                If oStyle.TextAlignmentVertical = GRE_VERTICALALIGNMENT.VAL_CENTER Then
                    lTextTop = v_lTop
                    lTextBottom = v_lBottom + 1
                End If
                DrawAlignedText(lTextLeft, lTextTop, lTextRight, lTextBottom, sText, GRE_HORIZONTALALIGNMENT.HAL_LEFT, GRE_VERTICALALIGNMENT.VAL_TOP, oStyle.ForeColor, oStyle.Font, oStyle.ClipText)
        End Select
    End Sub

    Friend Sub DrawPoint(ByVal X As Integer, ByVal Y As Integer, ByVal clrColor As Color)
        Dim mp_ucBrush As New SolidBrush(clrColor)
        oGraphics.FillRectangle(mp_ucBrush, X, Y, 1, 1)
    End Sub

    Friend Sub mp_DrawArrow(ByVal v_X As Integer, ByVal v_Y As Integer, ByVal v_ArrowDirection As GRE_ARROWDIRECTION, ByVal v_ArrowSize As Integer, ByVal v_lColor As Color)
        Dim i As Integer
        Select Case v_ArrowDirection
            Case GRE_ARROWDIRECTION.AWD_LEFT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X + 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_RIGHT
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_X = v_X - 1
                    DrawLine(v_X, v_Y - i, v_X, v_Y + i, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_UP
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y + 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
            Case GRE_ARROWDIRECTION.AWD_DOWN
                DrawPoint(v_X, v_Y, v_lColor)
                For i = 1 To v_ArrowSize
                    v_Y = v_Y - 1
                    DrawLine(v_X - i, v_Y, v_X + i, v_Y, GRE_LINETYPE.LT_NORMAL, v_lColor, GRE_LINEDRAWSTYLE.LDS_SOLID)
                Next
        End Select
    End Sub

End Class

