Option Explicit On
Imports System.Drawing

Public Class clsTierAppearance

    Private mp_oControl As ActiveGanttVBACtl
    Public MicrosecondColors As clsTierColors
    Public MillisecondColors As clsTierColors
    Public SecondColors As clsTierColors
    Public MinuteColors As clsTierColors
    Public HourColors As clsTierColors
    Public DayColors As clsTierColors
    Public DayOfWeekColors As clsTierColors
    Public DayOfYearColors As clsTierColors
    Public WeekColors As clsTierColors
    Public MonthColors As clsTierColors
    Public QuarterColors As clsTierColors
    Public YearColors As clsTierColors

    Friend Sub New(ByVal Value As ActiveGanttVBACtl)
        mp_oControl = Value
        MicrosecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MICROSECOND)
        MicrosecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MicrosecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MicrosecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MicrosecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MicrosecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MicrosecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MicrosecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MicrosecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MicrosecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MicrosecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MillisecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MILLISECOND)
        MillisecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MillisecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MillisecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MillisecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MillisecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MillisecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MillisecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MillisecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MillisecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MillisecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        SecondColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_SECOND)
        SecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        SecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        SecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        SecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        SecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        SecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        SecondColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        SecondColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        SecondColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        SecondColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MinuteColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MINUTE)
        MinuteColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MinuteColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MinuteColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MinuteColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MinuteColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MinuteColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MinuteColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        MinuteColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        MinuteColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        MinuteColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        HourColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_HOUR)
        HourColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        HourColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        HourColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        HourColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        HourColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        HourColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        HourColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        HourColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        HourColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        HourColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAY)
        DayColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        DayColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        DayColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        DayColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        DayColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayOfWeekColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAYOFWEEK)
        DayOfWeekColors.Add(Color.CornflowerBlue, Color.Black, Color.CornflowerBlue, Color.Black, Color.CornflowerBlue, Color.Black, "Sunday")
        DayOfWeekColors.Add(Color.MediumSlateBlue, Color.Black, Color.MediumSlateBlue, Color.Black, Color.MediumSlateBlue, Color.Black, "Monday")
        DayOfWeekColors.Add(Color.SlateBlue, Color.Black, Color.SlateBlue, Color.Black, Color.SlateBlue, Color.Black, "Tuesday")
        DayOfWeekColors.Add(Color.RoyalBlue, Color.Black, Color.RoyalBlue, Color.Black, Color.RoyalBlue, Color.Black, "Wednesday")
        DayOfWeekColors.Add(Color.SteelBlue, Color.Black, Color.SteelBlue, Color.Black, Color.SteelBlue, Color.Black, "Thursday")
        DayOfWeekColors.Add(Color.CadetBlue, Color.Black, Color.CadetBlue, Color.Black, Color.CadetBlue, Color.Black, "Friday")
        DayOfWeekColors.Add(Color.DodgerBlue, Color.Black, Color.DodgerBlue, Color.Black, Color.DodgerBlue, Color.Black, "Saturday")
        DayOfYearColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_DAYOFYEAR)
        DayOfYearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayOfYearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayOfYearColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        DayOfYearColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        DayOfYearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayOfYearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        DayOfYearColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        DayOfYearColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        DayOfYearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        DayOfYearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        WeekColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_WEEK)
        WeekColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        WeekColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        WeekColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        WeekColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        WeekColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        WeekColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        WeekColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        WeekColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        WeekColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        WeekColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        MonthColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_MONTH)
        MonthColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black, "January")
        MonthColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black, "February")
        MonthColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, "March")
        MonthColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black, "April")
        MonthColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black, "May")
        MonthColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black, "June")
        MonthColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, "July")
        MonthColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black, "August")
        MonthColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black, "September")
        MonthColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black, "October")
        MonthColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, "November")
        MonthColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black, "December")
        QuarterColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_QUARTER)
        QuarterColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        QuarterColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        QuarterColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        QuarterColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        YearColors = New clsTierColors(mp_oControl, E_TIERTYPE.ST_YEAR)
        YearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        YearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        YearColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        YearColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        YearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        YearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
        YearColors.Add(Color.DarkGray, Color.Black, Color.DarkGray, Color.Black, Color.DarkGray, Color.Black)
        YearColors.Add(Color.Silver, Color.Black, Color.Silver, Color.Black, Color.Silver, Color.Black)
        YearColors.Add(Color.DimGray, Color.Black, Color.DimGray, Color.Black, Color.DimGray, Color.Black)
        YearColors.Add(Color.Gray, Color.Black, Color.Gray, Color.Black, Color.Gray, Color.Black)
    End Sub

    Public Function GetXML() As String
        Dim oXML As New clsXML(mp_oControl, "TierAppearance")
        oXML.InitializeWriter()
        oXML.WriteObject(DayColors.GetXML())
        oXML.WriteObject(DayOfWeekColors.GetXML())
        oXML.WriteObject(DayOfYearColors.GetXML())
        oXML.WriteObject(HourColors.GetXML())
        oXML.WriteObject(MinuteColors.GetXML())
        oXML.WriteObject(SecondColors.GetXML())
        oXML.WriteObject(MillisecondColors.GetXML())
        oXML.WriteObject(MicrosecondColors.GetXML())
        oXML.WriteObject(MonthColors.GetXML())
        oXML.WriteObject(QuarterColors.GetXML())
        oXML.WriteObject(WeekColors.GetXML())
        oXML.WriteObject(YearColors.GetXML())
        Return oXML.GetXML()
    End Function

    Public Sub SetXML(ByVal sXML As String)
        Dim oXML As New clsXML(mp_oControl, "TierAppearance")
        oXML.SetXML(sXML)
        oXML.InitializeReader()
        DayColors.SetXML(oXML.ReadObject("DayColors"))
        DayOfWeekColors.SetXML(oXML.ReadObject("DayOfWeekColors"))
        DayOfYearColors.SetXML(oXML.ReadObject("DayOfYearColors"))
        HourColors.SetXML(oXML.ReadObject("HourColors"))
        MinuteColors.SetXML(oXML.ReadObject("MinuteColors"))
        SecondColors.SetXML(oXML.ReadObject("SecondColors"))
        MillisecondColors.SetXML(oXML.ReadObject("MillisecondColors"))
        MicrosecondColors.SetXML(oXML.ReadObject("MicrosecondColors"))
        MonthColors.SetXML(oXML.ReadObject("MonthColors"))
        QuarterColors.SetXML(oXML.ReadObject("QuarterColors"))
        WeekColors.SetXML(oXML.ReadObject("WeekColors"))
        YearColors.SetXML(oXML.ReadObject("YearColors"))
    End Sub

End Class

