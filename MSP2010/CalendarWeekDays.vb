Option Explicit On

Public Class CalendarWeekDays

	Private mp_oCollection As clsCollectionBase

	Public Sub New()
		mp_oCollection = New clsCollectionBase("CalendarWeekDay")
	End Sub

	Public ReadOnly Property Count() As Integer
		Get
			Return mp_oCollection.m_lCount
		End Get
	End Property

	Public Function Item(ByVal Index As String) As CalendarWeekDay
		Return mp_oCollection.m_oItem(Index, SYS_ERRORS.MP_ITEM_1, SYS_ERRORS.MP_ITEM_2, SYS_ERRORS.MP_ITEM_3, SYS_ERRORS.MP_ITEM_4)
	End Function

	Public Function Add() As CalendarWeekDay
		mp_oCollection.AddMode = True
		Dim oCalendarWeekDay As New CalendarWeekDay()
		oCalendarWeekDay.mp_oCollection = mp_oCollection
		mp_oCollection.m_Add(oCalendarWeekDay, "", SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
		Return oCalendarWeekDay
	End Function

	Public Sub Clear()
		mp_oCollection.m_Clear()
	End Sub

	Public Sub Remove(ByVal Index As String)
		mp_oCollection.m_Remove(Index, SYS_ERRORS.MP_REMOVE_1, SYS_ERRORS.MP_REMOVE_2, SYS_ERRORS.MP_REMOVE_3, SYS_ERRORS.MP_REMOVE_4)
	End Sub

	Public Function IsNull() As Boolean
		Dim bReturn As Boolean = True
		If Count > 0 Then
			bReturn = False
		End If
		Return bReturn
	End Function

	Public Function GetXML() As String
		If IsNull() = True Then
			Return "<WeekDays/>"
		End if
		Dim lIndex As Integer
		Dim oCalendarWeekDay As CalendarWeekDay
		Dim oXML As New clsXML("WeekDays")
		oXML.BoolsAreNumeric = True
		oXML.InitializeWriter()
		For lIndex = 1 To Count
			oCalendarWeekDay = mp_oCollection.m_oReturnArrayElement(lIndex)
			oXML.WriteObject(oCalendarWeekDay.GetXML)
		Next lIndex
		Return oXML.GetXML
	End Function

	Public Sub SetXML(ByVal sXML As String)
		Dim lIndex As Integer
		Dim oXML As New clsXML("WeekDays")
		oXML.SupportOptional = True
		oXML.SetXML(sXML)
		oXML.InitializeReader()
		mp_oCollection.m_Clear()
		If oXML.ReadCollectionCount = 0 Then
			Return
		End If
		For lIndex = 1 To oXML.ReadCollectionCount
			Dim oCalendarWeekDay As New CalendarWeekDay()
			oCalendarWeekDay.SetXML(oXML.ReadCollectionObject(lIndex))
			mp_oCollection.AddMode = True
			Dim sKey As String = ""
			oCalendarWeekDay.mp_oCollection = mp_oCollection
			mp_oCollection.m_Add(oCalendarWeekDay, sKey, SYS_ERRORS.MP_ADD_1, SYS_ERRORS.MP_ADD_2, False, SYS_ERRORS.MP_ADD_3)
			oCalendarWeekDay = Nothing
		Next lIndex
	End Sub

End Class
