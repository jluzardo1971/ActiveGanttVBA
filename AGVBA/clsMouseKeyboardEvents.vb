Option Explicit On 

Friend Class clsMouseKeyboardEvents

    Private mp_oControl As ActiveGanttVBACtl

    Friend Sub New(ByVal Value As ActiveGanttVBACtl)
        mp_oControl = Value
    End Sub

    Friend Sub OnMouseClick(ByVal X As Integer, ByVal Y As Integer)
        Dim yEventTarget As E_EVENTTARGET = E_EVENTTARGET.EVT_NONE
        yEventTarget = CursorPosition(X, Y)
        mp_oControl.MouseEventArgs.X = X
        mp_oControl.MouseEventArgs.Y = Y
        mp_oControl.MouseEventArgs.EventTarget = yEventTarget
        mp_oControl.MouseEventArgs.Cancel = False
        mp_oControl.FireControlClick()
        If mp_oControl.MouseEventArgs.Cancel = True Then
            Return
        End If
        Select Case yEventTarget
            Case E_EVENTTARGET.EVT_VSCROLLBAR
                mp_oControl.VerticalScrollBar.ScrollBar.ScrollBarClick(X, Y)
            Case E_EVENTTARGET.EVT_HSCROLLBAR
                mp_oControl.HorizontalScrollBar.ScrollBar.ScrollBarClick(X, Y)
            Case E_EVENTTARGET.EVT_TIMELINESCROLLBAR
                'X = X - mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.Left
                mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.ScrollBarClick(X, Y)
            Case E_EVENTTARGET.EVT_TREEVIEW
                mp_oControl.Treeview.TreeviewClick(X, Y)
            Case E_EVENTTARGET.EVT_TASK
                mp_oControl.SelectedTaskIndex = mp_oControl.MathLib.GetTaskIndexByPosition(X, Y)
            Case E_EVENTTARGET.EVT_PREDECESSOR
                mp_oControl.SelectedPredecessorIndex = mp_oControl.MathLib.GetPredecessorIndexByPosition(X, Y)
            Case E_EVENTTARGET.EVT_ROW
                mp_oControl.SelectedRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
            Case E_EVENTTARGET.EVT_CELL
                mp_oControl.SelectedCellIndex = mp_oControl.MathLib.GetCellIndexByPosition(X)
                mp_oControl.SelectedRowIndex = mp_oControl.MathLib.GetRowIndexByPosition(Y)
            Case E_EVENTTARGET.EVT_COLUMN
                mp_oControl.SelectedColumnIndex = mp_oControl.MathLib.GetColumnIndexByPosition(X, Y)
            Case E_EVENTTARGET.EVT_PERCENTAGE
                mp_oControl.SelectedPercentageIndex = mp_oControl.MathLib.GetPercentageIndexByPosition(X, Y)
        End Select
    End Sub

    Private Function CursorPosition(ByVal X As Integer, ByVal Y As Integer) As E_EVENTTARGET
        If mp_oControl.VerticalScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_VSCROLLBAR
        ElseIf mp_oControl.HorizontalScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_HSCROLLBAR
        ElseIf mp_oControl.CurrentViewObject.TimeLine.TimeLineScrollBar.ScrollBar.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TIMELINESCROLLBAR
        ElseIf mp_oControl.Treeview.OverControl(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TREEVIEW
        ElseIf mp_bOverSplitter(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SPLITTER
        ElseIf mp_bOverEmptySpace(Y) = True Then
            Return E_EVENTTARGET.EVT_NONE
        ElseIf mp_bOverTimeLine(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TIMELINE
        ElseIf mp_bOverSelectedColumn(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDCOLUMN
        ElseIf mp_bOverColumn(X, Y) = True Then
            Return E_EVENTTARGET.EVT_COLUMN
        ElseIf mp_bOverSelectedRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDROW
        ElseIf mp_bOverCell(X, Y) = True Then
            Return E_EVENTTARGET.EVT_CELL
        ElseIf mp_bOverRow(X, Y) = True Then
            Return E_EVENTTARGET.EVT_ROW
        ElseIf mp_bOverSelectedPercentage(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDPERCENTAGE
        ElseIf mp_bOverPercentage(X, Y) = True Then
            Return E_EVENTTARGET.EVT_PERCENTAGE
        ElseIf mp_bOverSelectedTask(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDTASK
        ElseIf mp_bOverTask(X, Y) = True Then
            Return E_EVENTTARGET.EVT_TASK
        ElseIf mp_bOverPredecessor(X, Y) = True Then
            Return E_EVENTTARGET.EVT_SELECTEDPREDECESSOR
        ElseIf mp_bOverClientArea(X, Y) = True Then
            Return E_EVENTTARGET.EVT_CLIENTAREA
        Else
            Return E_EVENTTARGET.EVT_NONE
        End If
    End Function

    Private Function mp_bOverSplitter(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If mp_oControl.Splitter.Width = 0 Then
            Return False
        End If
        If X >= (mp_oControl.Splitter.Right - mp_oControl.Splitter.Width) And X <= mp_oControl.Splitter.Right And Y < mp_oControl.clsG.Height() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverEmptySpace(ByVal Y As Integer) As Boolean
        If Y > mp_oControl.Rows.TopOffset Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverTimeLine(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom And Y >= mp_oControl.mt_TopMargin Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn = Nothing
        If mp_oControl.SelectedColumnIndex = 0 Or mp_oControl.Columns.Count = 0 Then
            Return False
        End If
        oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverColumn(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oColumn As clsColumn = Nothing
        Dim lIndex As Integer
        If Not (X <= mp_oControl.Splitter.Left And Y <= mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Columns.Count
            oColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(lIndex), clsColumn)
            If oColumn.Visible = True Then
                If X >= oColumn.Left And X <= oColumn.Right And Y >= oColumn.Top And Y <= oColumn.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Friend Function mp_bOverSelectedRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        If mp_oControl.SelectedRowIndex = 0 Or mp_oControl.Rows.Count = 0 Then
            Return False
        End If
        oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If oRow.MergeCells = True Then
            If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                Return True
            Else
                Return False
            End If
        Else
            If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                If mp_oControl.SelectedCellIndex = mp_oControl.MathLib.GetCellIndexByPosition(X) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End If
    End Function

    Friend Function mp_bOverRow(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        Dim lIndex As Integer
        If Not (X <= mp_oControl.CurrentViewObject.TimeLine.f_lStart And Y > mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Visible = True Then
                If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Friend Function mp_bOverCell(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oRow As clsRow = Nothing
        Dim lIndex As Integer
        If Not (X <= mp_oControl.CurrentViewObject.TimeLine.f_lStart And Y > mp_oControl.CurrentViewObject.TimeLine.Bottom) Then
            Return False
        End If
        For lIndex = 1 To mp_oControl.Rows.Count
            oRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(lIndex), clsRow)
            If oRow.Visible = True Then
                If X >= oRow.Left And X <= oRow.Right And Y >= oRow.Top And Y <= oRow.Bottom And oRow.MergeCells = False Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverSelectedTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedTask As clsTask = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedTaskIndex = 0 Then
            Return False
        End If
        oSelectedTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If X >= oSelectedTask.Left And X <= oSelectedTask.Right And Y >= oSelectedTask.Top And Y <= oSelectedTask.Bottom And mp_oControl.MathLib.InCurrentLayer(oSelectedTask.LayerIndex) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedPredecessor(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedPredecessor As clsPredecessor = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedPredecessorIndex = 0 Then
            Return False
        End If
        oSelectedPredecessor = DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(mp_oControl.SelectedPredecessorIndex), clsPredecessor)
        If oSelectedPredecessor.HitTest(X, Y) = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_bOverSelectedPercentage(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oSelectedPercentage As clsPercentage = Nothing
        If X < mp_oControl.CurrentViewObject.TimeLine.f_lStart Then
            Return False
        End If
        If X > mp_oControl.CurrentViewObject.TimeLine.f_lEnd Then
            Return False
        End If
        If mp_oControl.SelectedPercentageIndex = 0 Then
            Return False
        End If
        oSelectedPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(mp_oControl.SelectedPercentageIndex), clsPercentage)
        If X >= oSelectedPercentage.Left And X <= oSelectedPercentage.RightSel And Y >= oSelectedPercentage.Top And Y <= oSelectedPercentage.Bottom Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_yTaskArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedTask As clsTask = Nothing
        oSelectedTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(mp_oControl.SelectedTaskIndex), clsTask)
        If X >= oSelectedTask.Left And X <= oSelectedTask.Right And Y >= oSelectedTask.Top And Y <= oSelectedTask.Bottom And mp_oControl.MathLib.InCurrentLayer(oSelectedTask.LayerIndex) Then
            If X >= oSelectedTask.Left And X <= oSelectedTask.Left + 2 Then
                If oSelectedTask.f_bLeftVisible = True Then
                    Return E_AREA.EA_LEFT
                Else
                    Return E_AREA.EA_CENTER
                End If
            End If
            If X >= oSelectedTask.Right - 2 And X <= oSelectedTask.Right Then
                If oSelectedTask.f_bRightVisible = True Then
                    Return E_AREA.EA_RIGHT
                Else
                    Return E_AREA.EA_CENTER
                End If
            End If
            Return E_AREA.EA_CENTER
        End If
        Return E_AREA.EA_NONE
    End Function

    Friend Function mp_yRowArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedRow As clsRow = Nothing
        oSelectedRow = DirectCast(mp_oControl.Rows.oCollection.m_oReturnArrayElement(mp_oControl.SelectedRowIndex), clsRow)
        If Y >= oSelectedRow.Bottom And Y <= oSelectedRow.Bottom + 3 Then
            Return E_AREA.EA_BOTTOM
        Else
            Return E_AREA.EA_CENTER
        End If
    End Function

    Private Function mp_yColumnArea(ByVal X As Integer, ByVal Y As Integer) As E_AREA
        Dim oSelectedColumn As clsColumn = Nothing
        oSelectedColumn = DirectCast(mp_oControl.Columns.oCollection.m_oReturnArrayElement(mp_oControl.SelectedColumnIndex), clsColumn)
        If X >= (oSelectedColumn.Right - 3) And X <= oSelectedColumn.Right Then
            Return E_AREA.EA_RIGHT
        Else
            Return E_AREA.EA_CENTER
        End If
    End Function

    Private Function mp_bOverTask(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oTask As clsTask = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Tasks.Count To 1 Step -1
            oTask = DirectCast(mp_oControl.Tasks.oCollection.m_oReturnArrayElement(lIndex), clsTask)
            If oTask.Visible = True And mp_oControl.MathLib.InCurrentLayer(oTask.LayerIndex) Then
                If X >= oTask.Left And X <= oTask.Right And Y >= oTask.Top And Y <= oTask.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverPredecessor(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oPredecessor As clsPredecessor = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Predecessors.Count To 1 Step -1
            oPredecessor = (DirectCast(mp_oControl.Predecessors.oCollection.m_oReturnArrayElement(lIndex), clsPredecessor))
            If oPredecessor.Visible = True Then
                If oPredecessor.HitTest(X, Y) = True Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Function mp_bOverPercentage(ByVal X As Integer, ByVal Y As Integer) As Boolean
        Dim oPercentage As clsPercentage = Nothing
        Dim lIndex As Integer
        For lIndex = mp_oControl.Percentages.Count To 1 Step -1
            oPercentage = DirectCast(mp_oControl.Percentages.oCollection.m_oReturnArrayElement(lIndex), clsPercentage)
            If oPercentage.Visible = True Then
                If X >= oPercentage.Left And X <= oPercentage.RightSel And Y >= oPercentage.Top And Y <= oPercentage.Bottom Then
                    Return True
                End If
            End If
        Next lIndex
        Return False
    End Function

    Private Function mp_bOverClientArea(ByVal X As Integer, ByVal Y As Integer) As Boolean
        If X >= mp_oControl.CurrentViewObject.TimeLine.f_lStart And X <= mp_oControl.CurrentViewObject.TimeLine.f_lEnd And Y >= mp_oControl.CurrentViewObject.ClientArea.Top Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mp_fSnapX(ByVal X As Integer) As Integer
        Dim dtDate As AGVBA.DateTime = New AGVBA.DateTime()
        If mp_oControl.CurrentViewObject.ClientArea.Grid.SnapToGrid = False Then
            Return X
        End If
        dtDate = mp_oControl.MathLib.GetDateFromXCoordinate(X)
        dtDate = mp_oControl.MathLib.RoundDate(mp_oControl.CurrentViewObject.ClientArea.Grid.Interval, mp_oControl.CurrentViewObject.ClientArea.Grid.Factor, dtDate)
        Return mp_oControl.MathLib.GetXCoordinateFromDate(dtDate)
    End Function

End Class

