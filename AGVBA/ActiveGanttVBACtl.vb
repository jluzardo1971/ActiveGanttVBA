Option Explicit On

Imports System.ComponentModel
Imports System.Drawing
Imports System.Web.UI

< _
Designer("AGVBA.clsDesigner, AGVBA"), _
ToolboxData("<{0}:ActiveGanttVBACtl runat=server></{0}:ActiveGanttVBACtl>"), _
LicenseProviderAttribute(GetType(RegistryLicenseProvider)), _
System.Runtime.InteropServices.GuidAttribute("DE2C6F1B-CA45-4E47-A9BE-78ADBCDA260C") _
> _
Public Class ActiveGanttVBACtl
    Inherits System.Web.UI.WebControls.WebControl
    Implements IPostBackDataHandler, IPostBackEventHandler

#Region "IPostBackDataHandler"

    Public Function LoadPostData(ByVal postDataKey As String, ByVal postCollection As System.Collections.Specialized.NameValueCollection) As Boolean Implements System.Web.UI.IPostBackDataHandler.LoadPostData

    End Function

    Public Sub RaisePostDataChangedEvent() Implements System.Web.UI.IPostBackDataHandler.RaisePostDataChangedEvent

    End Sub

#End Region

#Region "IPostBackEventHandler"

    Public Sub RaisePostBackEvent(ByVal eventArgument As String) Implements IPostBackEventHandler.RaisePostBackEvent
        Dim X As Integer
        Dim Y As Integer
        X = Page.Request.Params("__CLICKCOORD_X")
        Y = Page.Request.Params("__CLICKCOORD_Y")
        OnClick(New System.Web.UI.ImageClickEventArgs(X, Y))
    End Sub

#End Region

    Private Sub WebCustomControl1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Page.RegisterRequiresPostBack(Me)

        Dim formName As String = Nothing
        Dim oControl As Control
        For Each oControl In Page.Controls
            If (oControl.GetType().ToString() = "System.Web.UI.HtmlControls.HtmlForm") Then
                formName = oControl.UniqueID
                Exit For
            End If
        Next
        If mp_sFormID = "" Then
            If formName Is Nothing Then
                Throw New Exception("The page has no form.")
            End If
        Else
            formName = mp_sFormID
        End If

        Page.ClientScript.RegisterHiddenField("__EVENTTARGET", "")
        Page.ClientScript.RegisterHiddenField("__EVENTARGUMENT", "")
        Page.ClientScript.RegisterHiddenField("__CLICKCOORD_X", "-1")
        Page.ClientScript.RegisterHiddenField("__CLICKCOORD_Y", "-1")
        'VS2002 & VS2003 equivalent:
        'Page.RegisterHiddenField("__EVENTTARGET", "")
        'Page.RegisterHiddenField("__EVENTARGUMENT", "")
        'Page.RegisterHiddenField("__CLICKCOORD_X", "-1")
        'Page.RegisterHiddenField("__CLICKCOORD_Y", "-1")

        Dim script As System.Text.StringBuilder = New System.Text.StringBuilder()
        script.Append(Environment.NewLine)
        script.AppendFormat("<script language=""javascript"">{0}", Environment.NewLine)
        script.AppendFormat("function __SCSIMG_CLICK(sCaller) {0}{1}", "{", Environment.NewLine)
        script.AppendFormat("var theform = document.{0};{1}", formName, Environment.NewLine)
        script.AppendFormat("theform.__EVENTTARGET.value = sCaller;{0}", Environment.NewLine)
        script.AppendFormat("theform.__EVENTARGUMENT.value = sCaller.split(""$"").join("":"");{0}", Environment.NewLine)
        script.AppendFormat("theform.__CLICKCOORD_X.value = event.offsetX;{0}", Environment.NewLine)
        script.AppendFormat("theform.__CLICKCOORD_Y.value = event.offsetY;{0}", Environment.NewLine)
        script.AppendFormat("theform.submit();{0}", Environment.NewLine)
        script.AppendFormat("{0}{1}", "}", Environment.NewLine)
        script.Append("</script>")

        If Page.ClientScript.IsClientScriptBlockRegistered("__SCSIMG_CLICK") = False Then
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "__SCSIMG_CLICK", script.ToString())
        End If
        ''VS2002 & VS2003 equivalent:
        'If Page.IsClientScriptBlockRegistered("__SCSIMG_CLICK") = False Then
        '    Page.RegisterClientScriptBlock("__SCSIMG_CLICK", script.ToString())
        'End If
    End Sub

    Protected Overrides Sub Render(ByVal output As System.Web.UI.HtmlTextWriter)
        Dim uniqueName As String = GenerateUniqueName()
        'Dim toolTip As String = mp_oToolTip
        Page.Application(clsAGVBARenderStream.ImageNamePrefix & uniqueName) = Me
        Dim sOutput As String
        sOutput = "<img src='" & clsAGVBARenderStream.ImageHandlerRequestFilename & "?id=" & uniqueName & "' id='" & Me.UniqueID & "' name='" & Me.UniqueID & "' border='0' height='" & Me.Height.Value & "' width='" & Me.Width.Value & "' alt='" & ToolTip & "' onclick='javascript:__SCSIMG_CLICK(""" & Me.UniqueID & """)'>"
        output.Write(sOutput)
    End Sub

    Private Function GenerateUniqueName() As String
        Dim sControlName As String = System.Guid.NewGuid().ToString()
        Return sControlName
    End Function

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Private Enumerations ActiveGanttCtl
    '// ---------------------------------------------------------------------------------------------------------------------

    Private Enum E_SCROLLSTATE
        SS_CANTDISPLAY = 0
        SS_NOTNEEDED = 1
        SS_NEEDED = 2
        SS_SHOWN = 3
        SS_HIDDEN = 4
    End Enum

    Private Enum E_DRAWOPTYPE
        DOT_ALL = 0
        DOT_ROWSANDCLIENTAREA = 1
        DOT_TABLEAREA = 2
        DOT_TIMELINEANDCLIENTAREA = 3
    End Enum

    '// ---------------------------------------------------------------------------------------------------------------------
    '// Member Variables
    '// ---------------------------------------------------------------------------------------------------------------------

    Private mp_oLicense As License = Nothing

    '// Public Classes
    <System.ComponentModel.Browsable(False)> Public Rows As clsRows
    <System.ComponentModel.Browsable(False)> Public Tasks As clsTasks
    <System.ComponentModel.Browsable(False)> Public Columns As clsColumns
    <System.ComponentModel.Browsable(False)> Public Styles As clsStyles
    <System.ComponentModel.Browsable(False)> Public Layers As clsLayers
    <System.ComponentModel.Browsable(False)> Public Percentages As clsPercentages
    <System.ComponentModel.Browsable(False)> Public TimeBlocks As clsTimeBlocks
    <System.ComponentModel.Browsable(False)> Public Predecessors As clsPredecessors
    <System.ComponentModel.Browsable(False)> Public Views As clsViews
    <System.ComponentModel.Browsable(False)> Public Splitter As clsSplitter
    <System.ComponentModel.Browsable(False)> Public Treeview As clsTreeview
    <System.ComponentModel.Browsable(False)> Public Drawing As clsDrawing
    <System.ComponentModel.Browsable(False)> Public MathLib As clsMath
    <System.ComponentModel.Browsable(False)> Public StrLib As clsString
    <System.ComponentModel.Browsable(False)> Public VerticalScrollBar As clsVerticalScrollBar
    <System.ComponentModel.Browsable(False)> Public HorizontalScrollBar As clsHorizontalScrollBar
    <System.ComponentModel.Browsable(False)> Public TierAppearance As clsTierAppearance
    <System.ComponentModel.Browsable(False)> Public TierFormat As clsTierFormat
    <System.ComponentModel.Browsable(False)> Public ScrollBarSeparator As clsScrollBarSeparator

    Friend oViewState As New clsViewState()

    Private tmpTimeBlocks As clsTimeBlocks
    Friend MouseKeyboardEvents As clsMouseKeyboardEvents
    Private mp_oCurrentView As clsView
    Friend clsG As clsGraphics
    Private mp_bAllowAdd As Boolean = True
    Private mp_bAllowEdit As Boolean = True
    Private mp_bAllowSplitterMove As Boolean = True
    Private mp_bAllowRowSize As Boolean = True
    Private mp_bAllowRowMove As Boolean = True
    Private mp_bAllowColumnSize As Boolean = True
    Private mp_bAllowColumnMove As Boolean = True
    Private mp_bAllowTimeLineScroll As Boolean = True
    Private mp_bAllowPredecessorAdd As Boolean = True
    Private mp_bDoubleBuffering As Boolean = True
    Private mp_bPropertiesRead As Boolean = False
    Private mp_bEnforcePredecessors As Boolean = False
    Private mp_lMinColumnWidth As Integer = 5
    Private mp_lMinRowHeight As Integer = 5
    Private mp_lSelectedTaskIndex As Integer = 0
    Private mp_lSelectedColumnIndex As Integer = 0
    Private mp_lSelectedRowIndex As Integer = 0
    Private mp_lSelectedCellIndex As Integer = 0
    Private mp_lSelectedPercentageIndex As Integer = 0
    Private mp_lSelectedPredecessorIndex As Integer = 0
    Private mp_lTreeviewColumnIndex As Integer = 0
    Private mp_sCurrentLayer As String = "0"
    Private mp_sCurrentView As String = ""
    Private mp_yAddMode As E_ADDMODE = E_ADDMODE.AT_TASKADD
    Private mp_yAddDurationInterval As E_INTERVAL = E_INTERVAL.IL_SECOND
    Private mp_yScrollBarBehaviour As E_SCROLLBEHAVIOUR = E_SCROLLBEHAVIOUR.SB_HIDE
    Private mp_yTimeBlockBehaviour As E_TIMEBLOCKBEHAVIOUR = E_TIMEBLOCKBEHAVIOUR.TBB_ROWEXTENTS
    Private mp_yLayerEnableObjects As E_LAYEROBJECTENABLE = E_LAYEROBJECTENABLE.EC_INCURRENTLAYERONLY
    Private mp_yErrorReports As E_REPORTERRORS = E_REPORTERRORS.RE_MSGBOX
    Private mp_yTierAppearanceScope As E_TIERAPPEARANCESCOPE = E_TIERAPPEARANCESCOPE.TAS_CONTROL
    Private mp_yTierFormatScope As E_TIERFORMATSCOPE = E_TIERFORMATSCOPE.TFS_CONTROL
    Private mp_yPredecessorMode As E_PREDECESSORMODE = E_PREDECESSORMODE.PM_CREATEWARNINGFLAG
    Private mp_sControlTag As String = ""
    Private mp_oGraphics As Graphics
    Private mp_oBitmap As Bitmap
    Private mp_oCulture As System.Globalization.CultureInfo
    Private mp_sStyleIndex As String
    Private mp_oStyle As clsStyle
    Private mp_oImage As Image
    Private mp_sImageTag As String
    Public ToolTipEventArgs As ToolTipEventArgs = New ToolTipEventArgs()
    Public ObjectAddedEventArgs As ObjectAddedEventArgs = New ObjectAddedEventArgs()
    Public CustomTierDrawEventArgs As CustomTierDrawEventArgs = New CustomTierDrawEventArgs()
    Public MouseEventArgs As MouseEventArgs = New MouseEventArgs()
    Public KeyEventArgs As KeyEventArgs = New KeyEventArgs()
    Public ScrollEventArgs As ScrollEventArgs = New ScrollEventArgs()
    Public DrawEventArgs As DrawEventArgs = New DrawEventArgs()
    Public PredecessorDrawEventArgs As PredecessorDrawEventArgs = New PredecessorDrawEventArgs()
    Public ObjectSelectedEventArgs As ObjectSelectedEventArgs = New ObjectSelectedEventArgs()
    Public ObjectStateChangedEventArgs As ObjectStateChangedEventArgs = New ObjectStateChangedEventArgs()
    Public ErrorEventArgs As ErrorEventArgs = New ErrorEventArgs()
    Public NodeEventArgs As NodeEventArgs = New NodeEventArgs()
    Public PredecessorExceptionEventArgs As PredecessorExceptionEventArgs = New PredecessorExceptionEventArgs()
    Private mp_oMSPIDataSet As DataSet

    Private mp_sFormID As String = ""

    Public Event ControlClick(ByVal sender As Object, ByVal e As MouseEventArgs)

    Public Event Draw(ByVal sender As Object, ByVal e As DrawEventArgs)
    Public Event PredecessorDraw(ByVal sender As Object, ByVal e As PredecessorDrawEventArgs)
    Public Event CustomTierDraw(ByVal sender As Object, ByVal e As CustomTierDrawEventArgs)
    Public Event TierTextDraw(ByVal sender As Object, ByVal e As CustomTierDrawEventArgs)

    Public Event ObjectSelected(ByVal sender As Object, ByVal e As ObjectSelectedEventArgs)

    Public Event ActiveGanttError(ByVal sender As Object, ByVal e As ErrorEventArgs)
    Public Event PredecessorException(ByVal sender As Object, ByVal e As PredecessorExceptionEventArgs)
    Public Event ControlScroll(ByVal sender As Object, ByVal e As ScrollEventArgs)
    Public Event ControlRedrawn(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event ControlDraw(ByVal sender As Object, ByVal e As System.EventArgs)
    Public Event TimeLineChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    Public Event NodeExpanded(ByVal sender As Object, ByVal e As NodeEventArgs)
    Public Event NodeChecked(ByVal sender As Object, ByVal e As NodeEventArgs)

    Public Event ViewStateRefreshed(ByVal sender As Object, ByVal e As System.EventArgs)

    Friend Sub FirePredecessorException()
        RaiseEvent PredecessorException(Me, PredecessorExceptionEventArgs)
    End Sub

    Friend Sub FireControlClick()
        RaiseEvent ControlClick(Me, MouseEventArgs)
    End Sub

    Friend Sub FireDraw()
        RaiseEvent Draw(Me, DrawEventArgs)
    End Sub

    Friend Sub FirePredecessorDraw()
        RaiseEvent PredecessorDraw(Me, PredecessorDrawEventArgs)
    End Sub

    Friend Sub FireCustomTierDraw()
        RaiseEvent CustomTierDraw(Me, CustomTierDrawEventArgs)
    End Sub

    Friend Sub FireTierTextDraw()
        RaiseEvent TierTextDraw(Me, CustomTierDrawEventArgs)
    End Sub

    Friend Sub FireObjectSelected()
        RaiseEvent ObjectSelected(Me, ObjectSelectedEventArgs)
    End Sub

    Friend Sub FireActiveGanttError()
        RaiseEvent ActiveGanttError(Me, ErrorEventArgs)
    End Sub

    Friend Sub FireControlScroll()
        RaiseEvent ControlScroll(Me, ScrollEventArgs)
    End Sub

    Friend Sub FireNodeChecked()
        RaiseEvent NodeChecked(Me, NodeEventArgs)
    End Sub

    Friend Sub FireNodeExpanded()
        RaiseEvent NodeExpanded(Me, NodeEventArgs)
    End Sub

    Friend Sub FireControlDraw()
        RaiseEvent ControlDraw(Me, New System.EventArgs())
    End Sub

    Friend Sub FireControlRedrawn()
        RaiseEvent ControlRedrawn(Me, New System.EventArgs())
    End Sub

    Friend Sub FireTimeLineChanged()
        RaiseEvent TimeLineChanged(Me, New System.EventArgs())
    End Sub

    Friend Function TempTimeBlocks() As clsTimeBlocks
        Return tmpTimeBlocks
    End Function

    Public Sub New()
        Me.mp_oLicense = LicenseManager.Validate(GetType(ActiveGanttVBACtl), Me)

        clsG = New clsGraphics(Me)
        MathLib = New clsMath(Me)
        StrLib = New clsString(Me)
        Styles = New clsStyles(Me)
        mp_sStyleIndex = "DS_CONTROL"
        mp_oStyle = Styles.FItem("DS_CONTROL")
        VerticalScrollBar = New clsVerticalScrollBar(Me)
        HorizontalScrollBar = New clsHorizontalScrollBar(Me)
        Rows = New clsRows(Me)
        Tasks = New clsTasks(Me)
        Columns = New clsColumns(Me)
        Layers = New clsLayers(Me)
        Percentages = New clsPercentages(Me)
        TimeBlocks = New clsTimeBlocks(Me)
        Predecessors = New clsPredecessors(Me)
        tmpTimeBlocks = New clsTimeBlocks(Me)
        Splitter = New clsSplitter(Me)
        Views = New clsViews(Me)
        Treeview = New clsTreeview(Me)
        mp_oCurrentView = Views.FItem("0")
        MouseKeyboardEvents = New clsMouseKeyboardEvents(Me)
        Drawing = New clsDrawing(Me)
        mp_oCulture = System.Globalization.CultureInfo.CurrentCulture.Clone()
        TierAppearance = New clsTierAppearance(Me)
        TierFormat = New clsTierFormat(Me)
        ScrollBarSeparator = New clsScrollBarSeparator(Me)

        mp_oImage = Nothing
        mp_sImageTag = ""

    End Sub

    Public Overrides Sub Dispose()
        If Not (mp_oLicense Is Nothing) Then
            mp_oLicense.Dispose()
            mp_oLicense = Nothing
        End If
    End Sub

    Friend Function OnPaint() As System.IO.MemoryStream 'All Drawing Here
        'On Error GoTo ErrorHandler
        Dim memStream As System.IO.MemoryStream = New System.IO.MemoryStream()
        Dim b As New Bitmap(clsG.Width, clsG.Height, Imaging.PixelFormat.Format24bppRgb)
        mp_oGraphics = Graphics.FromImage(b)
        mp_Draw()
        mp_oGraphics.Save()
        Dim imgformat As Imaging.ImageFormat = Imaging.ImageFormat.Png
        b.Save(memStream, imgformat)
        Return memStream

        'ErrorHandler:
        '        mp_ErrorReport(Err.Number, Err.Description, "Draw")
    End Function

    Private Sub mp_Draw()
        FireControlDraw()
        clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        clsG.mp_DrawItem(0, 0, clsG.Width - 1, clsG.Height - 1, "", "", False, Me.Image, 0, 0, Me.Style)
        mp_oCurrentView.TimeLine.Calculate()
        mp_PositionScrollBars()
        Columns.Position()
        Rows.InitializePosition()
        Rows.PositionRows()
        Columns.Draw()
        Rows.Draw()
        Treeview.Draw()
        mp_oCurrentView.TimeLine.Draw()
        mp_oCurrentView.TimeLine.ProgressLine.Draw()
        TimeBlocks.CreateTemporaryTimeBlocks()
        TimeBlocks.Draw()
        mp_oCurrentView.ClientArea.Grid.Draw()
        mp_oCurrentView.ClientArea.Draw()
        Predecessors.Draw()
        Tasks.Draw()
        Percentages.Draw()
        mp_oCurrentView.TimeLine.ProgressLine.Draw()
        Splitter.Draw()
        clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            clsG.mp_DrawItem(VerticalScrollBar.Left, VerticalScrollBar.Top + VerticalScrollBar.Height, VerticalScrollBar.Left + 16, VerticalScrollBar.Top + VerticalScrollBar.Height + 16, "", "", False, Nothing, 0, 0, ScrollBarSeparator.Style)
            clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        ElseIf mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            clsG.mp_DrawItem(mp_oCurrentView.TimeLine.TimeLineScrollBar.Left + mp_oCurrentView.TimeLine.TimeLineScrollBar.Width, mp_oCurrentView.TimeLine.TimeLineScrollBar.Top, mp_oCurrentView.TimeLine.TimeLineScrollBar.Left + mp_oCurrentView.TimeLine.TimeLineScrollBar.Width + 16, mp_oCurrentView.TimeLine.TimeLineScrollBar.Top + 16, "", "", False, Nothing, 0, 0, ScrollBarSeparator.Style)
            clsG.ClipRegion(0, 0, clsG.Width, clsG.Height, False)
        End If
        mp_DrawDebugMetrics()
        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            VerticalScrollBar.ScrollBar.Draw()
        End If
        If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            HorizontalScrollBar.ScrollBar.Draw()
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.ScrollBar.Draw()
        End If
#If DemoVersion Then
        Dim oFont As New Font("Arial", 12, FontStyle.Bold)
        Dim rnd As System.Random
        rnd = New System.Random()
        Dim oColor As Color = New Color()
        oColor = Color.FromArgb(255, rnd.Next(0, 255), rnd.Next(0, 255), rnd.Next(0, 255))
        clsG.DrawAlignedText(20, 20, clsG.Width() - 20, clsG.Height() - 20, "ActiveGanttVBA Scheduler Component" & vbCrLf & "Trial Version: " & Version & vbCrLf & "For evaluation purposes only" & vbCrLf & "Purchase the full version through: " & vbCrLf & "http://www.sourcecodestore.com", GRE_HORIZONTALALIGNMENT.HAL_RIGHT, GRE_VERTICALALIGNMENT.VAL_BOTTOM, oColor, oFont, True)
#End If
        FireControlRedrawn()
    End Sub

    Private Sub mp_DrawDebugMetrics()

    End Sub

    Friend Function f_HDC() As Graphics
        Return mp_oGraphics
    End Function

    Friend Function f_Width() As Integer
        Return Me.Width.Value
    End Function

    Friend Function f_Height() As Integer
        Return Me.Height.Value
    End Function

    Friend Function mp_lStrWidth(ByRef sString As String, ByRef r_oFont As Font) As Integer
        Return MathLib.RoundDouble(mp_oGraphics.MeasureString(sString, r_oFont).Width)
    End Function

    Friend Function mp_lStrHeight(ByRef sString As String, ByRef r_oFont As Font) As Integer
        Return MathLib.RoundDouble(mp_oGraphics.MeasureString(sString, r_oFont).Height)
    End Function

    Friend Sub f_Draw()
        mp_Draw()
    End Sub

    Friend ReadOnly Property f_UserMode() As Boolean
        Get
            Return True
        End Get
    End Property

    Friend ReadOnly Property mt_BorderThickness() As Integer
        Get
            Select Case mp_oStyle.Appearance
                Case E_STYLEAPPEARANCE.SA_RAISED
                    Return 2
                Case E_STYLEAPPEARANCE.SA_SUNKEN
                    Return 2
                Case E_STYLEAPPEARANCE.SA_FLAT
                    If mp_oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE Then
                        Return 0
                    Else
                        Return mp_oStyle.BorderWidth
                    End If
                Case E_STYLEAPPEARANCE.SA_CELL
                    If mp_oStyle.BorderStyle = GRE_BORDERSTYLE.SBR_NONE Then
                        Return 0
                    Else
                        Return mp_oStyle.BorderWidth
                    End If
                Case E_STYLEAPPEARANCE.SA_GRAPHICAL
                    Return 0
            End Select
            Return 0
        End Get
    End Property

    Friend ReadOnly Property mt_TableBottom() As Integer
        Get
            If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
                Return clsG.Height - mt_BorderThickness - 1 - HorizontalScrollBar.Height
            Else
                Return clsG.Height - mt_BorderThickness - 1
            End If
        End Get
    End Property

    Friend ReadOnly Property mt_TopMargin() As Integer
        Get
            Return mt_BorderThickness
        End Get
    End Property

    Friend ReadOnly Property mt_LeftMargin() As Integer
        Get
            Return mt_BorderThickness
        End Get
    End Property

    Friend ReadOnly Property mt_RightMargin() As Integer
        Get
            If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
                Return clsG.Width - mt_BorderThickness - 1 - VerticalScrollBar.Width
            Else
                Return clsG.Width - mt_BorderThickness - 1
            End If
        End Get
    End Property

    Friend ReadOnly Property mt_BottomMargin() As Integer
        Get
            Return clsG.Height - mt_BorderThickness - 1
        End Get
    End Property

    Friend Function GrphLib() As clsGraphics
        Return clsG
    End Function

    Protected Overridable Sub OnClick(ByVal e As System.Web.UI.ImageClickEventArgs)
        RefreshViewState()
        mp_oCurrentView.TimeLine.Calculate()
        mp_PositionScrollBars()
        Columns.Position()
        Rows.InitializePosition()
        Rows.PositionRows()
        MouseKeyboardEvents.OnMouseClick(e.X, e.Y)
    End Sub

    Friend Sub VerticalScrollBar_ValueChanged(ByVal Offset As Integer)
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_VERTICAL
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub HorizontalScrollBar_ValueChanged(ByVal Offset As Integer)
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_HORIZONTAL1
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub TimeLineScrollBar_ValueChanged(ByVal Offset As Integer)
        ScrollEventArgs.Clear()
        ScrollEventArgs.ScrollBarType = E_SCROLLBAR.SCR_HORIZONTAL2
        ScrollEventArgs.Offset = Offset
        FireControlScroll()
    End Sub

    Friend Sub mp_PositionScrollBars()
        If clsG.Height <= mp_oCurrentView.ClientArea.Top Then
            VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            Return
        End If

        '// Determine need for HorizontalScrollBar
        Dim lWidth As Integer = 0
        lWidth = Columns.Width
        If lWidth > Splitter.Right Then
            If HorizontalScrollBar.mf_Visible = True Then
                HorizontalScrollBar.State = E_SCROLLSTATE.SS_NEEDED
            Else
                HorizontalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
        Else
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
        End If
        If Splitter.Right < 5 Then
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        '// Determine need for mp_oCurrentView.TimeLine.TimeLineScrollBar
        If Splitter.Right < clsG.Width - (18 + mt_BorderThickness) Then
            If mp_oCurrentView.TimeLine.TimeLineScrollBar.Enabled = True Then
                If mp_oCurrentView.TimeLine.TimeLineScrollBar.mf_Visible = True Then
                    mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NEEDED
                Else
                    mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
                End If
            Else
                mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
            End If
        Else
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        '// Determine need for VerticalScrollBar
        If ((Rows.Height() + mp_oCurrentView.ClientArea.Top + HorizontalScrollBar.Height + mt_BorderThickness) > clsG.Height) Or (Rows.RealFirstVisibleRow > 1) Then
            If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY Then
                VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
            Else
                VerticalScrollBar.State = E_SCROLLSTATE.SS_NEEDED
            End If
        Else
            VerticalScrollBar.State = E_SCROLLSTATE.SS_NOTNEEDED
        End If

        If VerticalScrollBar.mf_Visible = False Then
            VerticalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If
        If HorizontalScrollBar.mf_Visible = False Then
            HorizontalScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.mf_Visible = False Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_CANTDISPLAY
        End If

        If VerticalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            VerticalScrollBar.Position()
        End If
        If HorizontalScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            HorizontalScrollBar.Position()
        End If
        If mp_oCurrentView.TimeLine.TimeLineScrollBar.State = E_SCROLLSTATE.SS_SHOWN Then
            mp_oCurrentView.TimeLine.TimeLineScrollBar.Position()
        End If
    End Sub


    Public Sub WriteXML(ByVal url As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        mp_WriteXML(oXML)
        oXML.WriteXML(url)
    End Sub

    Public Sub ReadXML(ByVal url As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        oXML.ReadXML(url)
        mp_ReadXML(oXML)
    End Sub

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        oXML.SetXML(sXML)
        mp_ReadXML(oXML)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(Me, "ActiveGanttCtl")
        mp_WriteXML(oXML)
        Return oXML.GetXML
    End Function

    Private Sub mp_WriteXML(ByRef oXML As clsXML)
        oXML.InitializeWriter()
        oXML.WriteProperty("Version", "AGVBA")
        oXML.WriteProperty("ControlTag", mp_sControlTag)
        oXML.WriteProperty("AddMode", mp_yAddMode)
        oXML.WriteProperty("AddDurationInterval", mp_yAddDurationInterval)
        oXML.WriteProperty("AllowAdd", mp_bAllowAdd)
        oXML.WriteProperty("AllowColumnMove", mp_bAllowColumnMove)
        oXML.WriteProperty("AllowColumnSize", mp_bAllowColumnSize)
        oXML.WriteProperty("AllowEdit", mp_bAllowEdit)
        oXML.WriteProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd)
        oXML.WriteProperty("AllowRowMove", mp_bAllowRowMove)
        oXML.WriteProperty("AllowRowSize", mp_bAllowRowSize)
        oXML.WriteProperty("AllowSplitterMove", mp_bAllowSplitterMove)
        oXML.WriteProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll)
        oXML.WriteProperty("EnforcePredecessors", mp_bEnforcePredecessors)
        oXML.WriteProperty("CurrentLayer", mp_sCurrentLayer)
        oXML.WriteProperty("CurrentView", mp_sCurrentView)
        oXML.WriteProperty("DoubleBuffering", mp_bDoubleBuffering)
        oXML.WriteProperty("ErrorReports", mp_yErrorReports)
        oXML.WriteProperty("LayerEnableObjects", mp_yLayerEnableObjects)
        oXML.WriteProperty("MinColumnWidth", mp_lMinColumnWidth)
        oXML.WriteProperty("MinRowHeight", mp_lMinRowHeight)
        oXML.WriteProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour)
        oXML.WriteProperty("SelectedCellIndex", mp_lSelectedCellIndex)
        oXML.WriteProperty("SelectedColumnIndex", mp_lSelectedColumnIndex)
        oXML.WriteProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex)
        oXML.WriteProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex)
        oXML.WriteProperty("SelectedRowIndex", mp_lSelectedRowIndex)
        oXML.WriteProperty("SelectedTaskIndex", mp_lSelectedTaskIndex)
        oXML.WriteProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex)
        oXML.WriteProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour)
        oXML.WriteProperty("TierAppearanceScope", mp_yTierAppearanceScope)
        oXML.WriteProperty("TierFormatScope", mp_yTierFormatScope)
        oXML.WriteProperty("PredecessorMode", mp_yPredecessorMode)
        oXML.WriteProperty("StyleIndex", mp_sStyleIndex)
        oXML.WriteProperty("Image", mp_oImage)
        oXML.WriteProperty("ImageTag", mp_sImageTag)
        oXML.WriteObject(Styles.GetXML())
        oXML.WriteObject(Rows.GetXML())
        oXML.WriteObject(Columns.GetXML())
        oXML.WriteObject(Layers.GetXML())
        oXML.WriteObject(Tasks.GetXML())
        oXML.WriteObject(Predecessors.GetXML())
        oXML.WriteObject(Views.GetXML())
        oXML.WriteObject(TimeBlocks.GetXML())
        oXML.WriteObject(TimeBlocks.CP_GetXML())
        oXML.WriteObject(Percentages.GetXML())
        oXML.WriteObject(Splitter.GetXML())
        oXML.WriteObject(Treeview.GetXML())
        oXML.WriteObject(TierAppearance.GetXML())
        oXML.WriteObject(TierFormat.GetXML())
        oXML.WriteObject(ScrollBarSeparator.GetXML())
        oXML.WriteObject(VerticalScrollBar.GetXML())
        oXML.WriteObject(HorizontalScrollBar.GetXML())
    End Sub

    Private Sub mp_ReadXML(ByRef oXML As clsXML)
        Dim sVersion As String = ""
        Dim sCurrentView As String = ""
        Clear()
        oXML.InitializeReader()
        oXML.ReadProperty("Version", sVersion)
        oXML.ReadProperty("ControlTag", mp_sControlTag)
        oXML.ReadProperty("AddMode", mp_yAddMode)
        oXML.ReadProperty("AddDurationInterval", mp_yAddDurationInterval)
        oXML.ReadProperty("AllowAdd", mp_bAllowAdd)
        oXML.ReadProperty("AllowColumnMove", mp_bAllowColumnMove)
        oXML.ReadProperty("AllowColumnSize", mp_bAllowColumnSize)
        oXML.ReadProperty("AllowEdit", mp_bAllowEdit)
        oXML.ReadProperty("AllowPredecessorAdd", mp_bAllowPredecessorAdd)
        oXML.ReadProperty("AllowRowMove", mp_bAllowRowMove)
        oXML.ReadProperty("AllowRowSize", mp_bAllowRowSize)
        oXML.ReadProperty("AllowSplitterMove", mp_bAllowSplitterMove)
        oXML.ReadProperty("AllowTimeLineScroll", mp_bAllowTimeLineScroll)
        oXML.ReadProperty("EnforcePredecessors", mp_bEnforcePredecessors)
        oXML.ReadProperty("CurrentLayer", mp_sCurrentLayer)
        oXML.ReadProperty("CurrentView", mp_sCurrentView)
        oXML.ReadProperty("DoubleBuffering", mp_bDoubleBuffering)
        oXML.ReadProperty("ErrorReports", mp_yErrorReports)
        oXML.ReadProperty("LayerEnableObjects", mp_yLayerEnableObjects)
        oXML.ReadProperty("MinColumnWidth", mp_lMinColumnWidth)
        oXML.ReadProperty("MinRowHeight", mp_lMinRowHeight)
        oXML.ReadProperty("ScrollBarBehaviour", mp_yScrollBarBehaviour)
        oXML.ReadProperty("SelectedCellIndex", mp_lSelectedCellIndex)
        oXML.ReadProperty("SelectedColumnIndex", mp_lSelectedColumnIndex)
        oXML.ReadProperty("SelectedPercentageIndex", mp_lSelectedPercentageIndex)
        oXML.ReadProperty("SelectedPredecessorIndex", mp_lSelectedPredecessorIndex)
        oXML.ReadProperty("SelectedRowIndex", mp_lSelectedRowIndex)
        oXML.ReadProperty("SelectedTaskIndex", mp_lSelectedTaskIndex)
        oXML.ReadProperty("TreeviewColumnIndex", mp_lTreeviewColumnIndex)
        oXML.ReadProperty("TimeBlockBehaviour", mp_yTimeBlockBehaviour)
        oXML.ReadProperty("TierAppearanceScope", mp_yTierAppearanceScope)
        oXML.ReadProperty("TierFormatScope", mp_yTierFormatScope)
        oXML.ReadProperty("PredecessorMode", mp_yPredecessorMode)
        oXML.ReadProperty("StyleIndex", mp_sStyleIndex)
        oXML.ReadProperty("Image", mp_oImage)
        oXML.ReadProperty("ImageTag", mp_sImageTag)
        Styles.SetXML(oXML.ReadObject("Styles"))
        Rows.SetXML(oXML.ReadObject("Rows"))
        Columns.SetXML(oXML.ReadObject("Columns"))
        Layers.SetXML(oXML.ReadObject("Layers"))
        Tasks.SetXML(oXML.ReadObject("Tasks"))
        Predecessors.SetXML(oXML.ReadObject("Predecessors"))
        Views.SetXML(oXML.ReadObject("Views"))
        TimeBlocks.SetXML(oXML.ReadObject("TimeBlocks"))
        TimeBlocks.CP_SetXML(oXML.ReadObject("CP_TimeBlocks"))
        Percentages.SetXML(oXML.ReadObject("Percentages"))
        Splitter.SetXML(oXML.ReadObject("Splitter"))
        Treeview.SetXML(oXML.ReadObject("Treeview"))
        TierAppearance.SetXML(oXML.ReadObject("TierAppearance"))
        TierFormat.SetXML(oXML.ReadObject("TierFormat"))
        ScrollBarSeparator.SetXML(oXML.ReadObject("ScrollBarSeparator"))
        VerticalScrollBar.SetXML(oXML.ReadObject("VerticalScrollBar"))
        HorizontalScrollBar.SetXML(oXML.ReadObject("HorizontalScrollBar"))
        StyleIndex = mp_sStyleIndex
        Rows.UpdateTree()
        CurrentView = mp_sCurrentView
        mp_oCurrentView.TimeLine.Position(mp_oCurrentView.TimeLine.StartDate)
    End Sub

    Friend Sub mp_ErrorReport(ByVal ErrNumber As Integer, ByVal ErrDescription As String, ByVal ErrSource As String)
        If mp_yErrorReports = E_REPORTERRORS.RE_MSGBOX Then
            ShowMessageBox(System.Convert.ToString(ErrNumber) & ": " & ErrDescription & " (" & ErrSource & ")")
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_HIDE Then
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_RAISE Then
            Dim ex As AGError = New AGError(ErrNumber.ToString() & ": " & ErrDescription + " - " & ErrSource)
            ex.ErrNumber = ErrNumber
            ex.ErrDescription = ErrDescription
            ex.ErrSource = ErrSource
            Throw ex
        ElseIf mp_yErrorReports = E_REPORTERRORS.RE_RAISEEVENT Then
            ErrorEventArgs.Clear()
            ErrorEventArgs.Number = ErrNumber
            ErrorEventArgs.Description = ErrDescription
            ErrorEventArgs.Source = ErrSource
            FireActiveGanttError()
        End If
    End Sub

    Public Property ErrorReports() As E_REPORTERRORS
        Get
            Return mp_yErrorReports
        End Get
        Set(ByVal Value As E_REPORTERRORS)
            mp_yErrorReports = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property CurrentLayer() As String
        Get
            Return mp_sCurrentLayer
        End Get
        Set(ByVal Value As String)
            mp_sCurrentLayer = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property CurrentView() As String
        Get
            Return mp_sCurrentView
        End Get
        Set(ByVal Value As String)
            If Value = "" Then
                Value = "0"
            End If
            mp_oCurrentView = Views.FItem(Value)
            mp_sCurrentView = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property CurrentViewObject() As clsView
        Get
            Return mp_oCurrentView
        End Get
    End Property

    Public Property LayerEnableObjects() As E_LAYEROBJECTENABLE
        Get
            Return mp_yLayerEnableObjects
        End Get
        Set(ByVal Value As E_LAYEROBJECTENABLE)
            mp_yLayerEnableObjects = Value
        End Set
    End Property

    Public Property ScrollBarBehaviour() As E_SCROLLBEHAVIOUR
        Get
            Return mp_yScrollBarBehaviour
        End Get
        Set(ByVal Value As E_SCROLLBEHAVIOUR)
            mp_yScrollBarBehaviour = Value
        End Set
    End Property

    Public Property TierAppearanceScope() As E_TIERAPPEARANCESCOPE
        Get
            Return mp_yTierAppearanceScope
        End Get
        Set(ByVal Value As E_TIERAPPEARANCESCOPE)
            mp_yTierAppearanceScope = Value
        End Set
    End Property

    Public Property TierFormatScope() As E_TIERFORMATSCOPE
        Get
            Return mp_yTierFormatScope
        End Get
        Set(ByVal Value As E_TIERFORMATSCOPE)
            mp_yTierFormatScope = Value
        End Set
    End Property

    Public Property TimeBlockBehaviour() As E_TIMEBLOCKBEHAVIOUR
        Get
            Return mp_yTimeBlockBehaviour
        End Get
        Set(ByVal Value As E_TIMEBLOCKBEHAVIOUR)
            mp_yTimeBlockBehaviour = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property SelectedTaskIndex() As Integer
        Get
            Return mp_lSelectedTaskIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Tasks.Count Then
                Value = Tasks.Count
            End If
            mp_lSelectedTaskIndex = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property SelectedColumnIndex() As Integer
        Get
            Return mp_lSelectedColumnIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Columns.Count Then
                Value = Columns.Count
            End If
            mp_lSelectedColumnIndex = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property SelectedRowIndex() As Integer
        Get
            Return mp_lSelectedRowIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Rows.Count Then
                Value = Rows.Count
            End If
            mp_lSelectedRowIndex = Value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property SelectedCellIndex() As Integer
        Get
            Return mp_lSelectedCellIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Columns.Count Then
                Value = Columns.Count
            End If
            mp_lSelectedCellIndex = Value
        End Set
    End Property

    Public Property SelectedPercentageIndex() As Integer
        Get
            Return mp_lSelectedPercentageIndex
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                Value = 0
            ElseIf Value > Percentages.Count Then
                Value = Percentages.Count
            End If
            mp_lSelectedPercentageIndex = Value
        End Set
    End Property

    Public Property SelectedPredecessorIndex() As Integer
        Get
            Return mp_lSelectedPredecessorIndex
        End Get
        Set(ByVal value As Integer)
            If value <= 0 Then
                value = 0
            ElseIf value > Percentages.Count Then
                value = Percentages.Count
            End If
            mp_lSelectedPredecessorIndex = value
        End Set
    End Property

    Public Property TreeviewColumnIndex() As Integer
        Get
            If Columns.Count = 0 Then
                Return 0
            ElseIf mp_lTreeviewColumnIndex > Columns.Count Then
                Return 0
            ElseIf mp_lTreeviewColumnIndex < 0 Then
                Return 0
            Else
                Return mp_lTreeviewColumnIndex
            End If
        End Get
        Set(ByVal value As Integer)
            If value <= 0 Then
                value = 0
            ElseIf value > Columns.Count Then
                value = Columns.Count
            End If
            mp_lTreeviewColumnIndex = value
        End Set
    End Property

    Public Property StyleIndex() As String
        Get
            If mp_sStyleIndex = "DS_CONTROL" Then
                Return ""
            Else
                Return mp_sStyleIndex
            End If
        End Get
        Set(ByVal Value As String)
            Value = Value.Trim()
            If Value.Length = 0 Then Value = "DS_CONTROL"
            mp_sStyleIndex = Value
            mp_oStyle = Styles.FItem(Value)
        End Set
    End Property

    Public Shadows ReadOnly Property Style() As clsStyle
        Get
            Return mp_oStyle
        End Get
    End Property

    Public Property Image() As Image
        Get
            Return mp_oImage
        End Get
        Set(ByVal Value As Image)
            mp_oImage = Value
        End Set
    End Property

    Public Property ImageTag() As String
        Get
            Return mp_sImageTag
        End Get
        Set(ByVal Value As String)
            mp_sImageTag = Value
        End Set
    End Property

    Public Property Culture() As System.Globalization.CultureInfo
        Get
            Return mp_oCulture
        End Get
        Set(ByVal Value As System.Globalization.CultureInfo)
            mp_oCulture = Value
        End Set
    End Property

    Public Property ControlTag() As String
        Get
            Return mp_sControlTag
        End Get
        Set(ByVal Value As String)
            mp_sControlTag = Value
        End Set
    End Property

    Public Sub ClearSelections()
        mp_lSelectedTaskIndex = 0
        mp_lSelectedColumnIndex = 0
        mp_lSelectedRowIndex = 0
        mp_lSelectedCellIndex = 0
        mp_lSelectedPercentageIndex = 0
        mp_lSelectedPredecessorIndex = 0
    End Sub

    Public Sub Clear()
        Tasks.Clear()
        Rows.Clear()
        Styles.Clear()
        Layers.Clear()
        Columns.Clear()
        TimeBlocks.Clear()
        Views.Clear()
    End Sub

    Public Sub CheckPredecessors()
        Dim i As Integer
        Dim oTask As clsTask
        For i = 1 To Tasks.Count
            oTask = Tasks.oCollection.m_oReturnArrayElement(i)
            oTask.mp_bWarning = False
        Next
        If Predecessors.Count = 0 Then
            Return
        End If
        Dim oPredecessor As clsPredecessor
        For i = 1 To Predecessors.Count
            oPredecessor = Predecessors.oCollection.m_oReturnArrayElement(i)
            oPredecessor.Check(mp_yPredecessorMode)
        Next
    End Sub

    Public Property EnforcePredecessors() As Boolean
        Get
            Return mp_bEnforcePredecessors
        End Get
        Set(ByVal value As Boolean)
            mp_bEnforcePredecessors = value
        End Set
    End Property

    Public Property PredecessorMode() As E_PREDECESSORMODE
        Get
            Return mp_yPredecessorMode
        End Get
        Set(ByVal value As E_PREDECESSORMODE)
            mp_yPredecessorMode = value
        End Set
    End Property

    Public ReadOnly Property ModuleCompletePath() As String
        Get
            Return System.Reflection.Assembly.GetExecutingAssembly.Location
        End Get
    End Property

    Public ReadOnly Property Version() As String
        Get
            Dim ai As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
            Return ai.GetName.Version().ToString()
        End Get
    End Property

    Protected Overrides Function SaveViewState() As Object
        Dim oState(5) As Object
        oState(0) = VerticalScrollBar.ScrollBar.SaveViewState()
        oState(1) = HorizontalScrollBar.ScrollBar.SaveViewState()
        oState(2) = CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.SaveViewState()
        'CheckBoxes
        Dim lIndex As Integer
        Dim bChecked As Boolean
        Dim bExpanded As Boolean
        Dim sKey As String
        Dim sParam As String = ""
        Dim oRow As clsRow
        For lIndex = 1 To Rows.Count - 1
            oRow = Rows.Item(lIndex)
            bChecked = oRow.Node.Checked
            sKey = oRow.Key
            sParam = sParam & sKey & "," & System.Convert.ToInt16(bChecked).ToString & ";"
        Next
        oState(3) = sParam
        sParam = ""
        For lIndex = 1 To Rows.Count - 1
            oRow = Rows.Item(lIndex)
            bExpanded = oRow.Node.Expanded
            sKey = oRow.Key
            sParam = sParam & sKey & "," & System.Convert.ToInt16(bExpanded).ToString & ";"
        Next
        oState(4) = sParam
        Return oState
    End Function

    Protected Overrides Sub LoadViewState(ByVal savedState As Object)
        If Not (savedState Is Nothing) Then
            VerticalScrollBar.ScrollBar.LoadViewState(savedState(0))
            HorizontalScrollBar.ScrollBar.LoadViewState(savedState(1))
            CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.LoadViewState(savedState(2))
            oViewState.VerticalScrollBar_Value = VerticalScrollBar.ScrollBar.Value
            oViewState.HorizontalScrollBar_Value = HorizontalScrollBar.ScrollBar.Value
            oViewState.TimeLineScrollBar_Value = mp_oCurrentView.TimeLine.TimeLineScrollBar.ScrollBar.Value
            oViewState.sCheckedNodes = savedState(3)
            oViewState.sExpandedNodes = savedState(4)
        End If
    End Sub

    Private Sub RefreshViewState()
        Dim sParam As String
        Dim aParam() As String
        Dim lIndex As Long
        Dim sRow As String
        Dim bChecked As Boolean
        Dim bExpanded As Boolean
        RaiseEvent ViewStateRefreshed(Me, New System.EventArgs())
        VerticalScrollBar.ScrollBar.Value = oViewState.VerticalScrollBar_Value
        HorizontalScrollBar.ScrollBar.Value = oViewState.HorizontalScrollBar_Value
        mp_oCurrentView.TimeLine.TimeLineScrollBar.ScrollBar.Value = oViewState.TimeLineScrollBar_Value

        sParam = oViewState.sCheckedNodes
        aParam = sParam.Split(";")
        For lIndex = 0 To aParam.GetUpperBound(0)
            Dim aRow() As String
            sRow = aParam(lIndex)
            If sRow.Length > 0 Then
                aRow = sRow.Split(",")
                If Rows.oCollection.m_bDoesKeyExist(aRow(0)) = True Then
                    If aRow(1) = "0" Then
                        bChecked = False
                    Else
                        bChecked = True
                    End If
                    Rows.Item(aRow(0)).Node.Checked = bChecked
                End If
            End If
        Next

        sParam = oViewState.sExpandedNodes
        aParam = sParam.Split(";")
        For lIndex = 0 To aParam.GetUpperBound(0)
            Dim aRow() As String
            sRow = aParam(lIndex)
            If sRow.Length > 0 Then
                aRow = sRow.Split(",")
                If Rows.oCollection.m_bDoesKeyExist(aRow(0)) = True Then
                    If aRow(1) = "0" Then
                        bExpanded = False
                    Else
                        bExpanded = True
                    End If
                    Rows.Item(aRow(0)).Node.Expanded = bExpanded
                End If
            End If
        Next

    End Sub

    Public Property FormID() As String
        Get
            Return mp_sFormID
        End Get
        Set(ByVal Value As String)
            mp_sFormID = Value
        End Set
    End Property

    Private Function FindColumn(ByVal oDataTable As DataTable, ByVal sColumnName As String) As Boolean
        For Each oColumn As DataColumn In oDataTable.Columns
            If oColumn.ColumnName.ToLower() = sColumnName.ToLower() Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Sub ShowMessageBox(ByVal sMessage As String)
        sMessage = sMessage.Replace("'", "\\'")
        Dim sScript As String = "<script type=""text/javascript"">alert('" & sMessage & "')</script>"
        Dim oPage As System.Web.UI.Page
        oPage = CType(System.Web.HttpContext.Current.CurrentHandler, System.Web.UI.Page)
        If ((Not oPage Is Nothing) And (Not oPage.ClientScript.IsClientScriptBlockRegistered("alert"))) Then
            oPage.ClientScript.RegisterClientScriptBlock(oPage.GetType(), "alert", sScript)
        End If
    End Sub

End Class

Public Class AGError
    Inherits Exception

    Private mp_sErrDescription As String
    Private mp_lErrNumber As Integer
    Private mp_sErrSource As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal s As String)
        MyBase.New(s)
    End Sub

    Public Sub New(ByVal s As String, ByVal ex As Exception)
        MyBase.New(s, ex)
    End Sub

    Public Property ErrDescription() As String
        Get
            Return mp_sErrDescription
        End Get
        Set(ByVal value As String)
            mp_sErrDescription = value
        End Set
    End Property

    Public Property ErrNumber() As Integer
        Get
            Return mp_lErrNumber
        End Get
        Set(ByVal value As Integer)
            mp_lErrNumber = value
        End Set
    End Property

    Public Property ErrSource() As String
        Get
            Return mp_sErrSource
        End Get
        Set(ByVal value As String)
            mp_sErrSource = value
        End Set
    End Property
End Class
