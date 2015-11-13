'' shouldn't need any imported namespaces
'' just manipulating a date time variable


Public Class DTPManipulation
    '' make as generic as possible to expose to a host of other forms
    ''
    '' make a function? return to values to pipe back to the date time pickers?
    '' make a sub with properties?
    ''
    '' general idea:
    '' -------------------------- 
    '' new(byval range as string)
    '' //if function
    '' return d1 and d2 as dates
    '' // if sub
    '' return two properties apptdatefrom as date, apptdateto as date
    ''
    '' General Notes:
    '' --------------------------
    '' date manipuations that state "Previous" or "Last" are subtracted from the today's date.
    '' Ranges that state "Next" begin on the subsequent Monday.
    '' IE: "Next Week" begins on next "Monday"
    '' 
    Public retDateFrom As String = ""
    Public retDateTo As String = ""
    Private retFrom As String = ""
    Private retTo As String = ""

    Public Property ReturDateFrom() As String
        Get
            Return retDateFrom
        End Get
        Set(ByVal value As String)
            retDateFrom = value
        End Set
    End Property

    Public Property ReturnDateTo() As String
        Get
            Return retTo
        End Get
        Set(ByVal value As String)
            retTo = value
        End Set
    End Property

    Public Sub New(ByVal RangeChoice As String)
        '' notes:
        ''
        '' what is All? return all dates?
        '' 
        Select Case RangeChoice
            Case Is = "All"
                'MsgBox("Reset function here.", MsgBoxStyle.Information)
                Exit Select
            Case Is = "Today"
                '' today i assume would be the this date value of today on both date time pickers...
                Dim d1 As Date = Date.Today.Date.ToShortDateString
                Dim d2 As Date = Date.Today.Date.ToShortDateString
                Me.retDateFrom = d1
                Me.retDateTo = d2
                Exit Select
            Case Is = "Yesterday"
                '' yesterday is going to manipulate a timespan and minus one day.
                ''
                '' Notes: must look up timespans for accurate manipulation again.
                '' from is going to be todays date, to is going to be yesterdays date.
                '' in order to manipulate a timespan, you must break apart the date being dealt with
                '' here, it is going to be today's date // from date
                Dim day As String = ""
                Dim month As String = ""
                Dim year As String = ""
                Dim ts As New TimeSpan(1)
                Dim d1 As Date = Date.Today.Date.ToShortDateString
                Dim d2 As Date = d1.Subtract(ts).ToShortDateString
                Me.retDateFrom = d1.ToShortDateString

                Me.retDateFrom = d2
                Me.retDateTo = d2
                '' andy said to just have the date time pickers reflect yesterday's date.

                Exit Select
            Case Is = "This Week"
                '' first determine what day it is
                '' second determine the difference between what day it is and "Monday"
                '' IE: today is thurs, how many days between thurs and monday
                ''     
                Dim DayOfWeek As Integer = Date.Today.DayOfWeek
                Select Case DayOfWeek
                    Case Is = 0
                        '' sunday
                        Dim ts As New TimeSpan(6, 0, 0, 0, 0)
                        Dim d2 As Date = Date.Today.Add(ts) '' previous monday
                        Me.retDateFrom = d2.ToShortDateString
                        Me.retDateTo = Date.Today.Date.ToShortDateString
                        Exit Select
                    Case Is = 1
                        '' monday
                        Dim tsadd As New TimeSpan(5, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(1, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.Date.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 2
                        '' tuesday
                        Dim tsadd As New TimeSpan(4, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(2, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 3
                        '' wednesday
                        Dim tsadd As New TimeSpan(3, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(3, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 4
                        '' thurs
                        Dim tsadd As New TimeSpan(2, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(4, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 5
                        '' fri
                        Dim tsadd As New TimeSpan(1, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(5, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 6
                        '' sat
                        Dim tsadd As New TimeSpan(0, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(6, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                End Select
                Exit Select
            Case Is = "This Week - to date"
                '' i cant remember if this one is seven days back or if it is the one above it....
                '' easily swappable
                Dim DayOfWeek As Integer = Date.Today.DayOfWeek
                Select Case DayOfWeek
                    Case Is = 0
                        '' sunday
                        Dim ts As New TimeSpan(0, 0, 0, 0, 0)
                        Dim d2 As Date = Date.Today.Add(ts) '' previous monday
                        Me.retDateFrom = d2.ToShortDateString
                        Me.retDateTo = Date.Today.Date.ToShortDateString
                        Exit Select
                    Case Is = 1
                        '' monday
                        Dim tsadd As New TimeSpan(5, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(1, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.Date.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 2
                        '' tuesday
                        Dim tsadd As New TimeSpan(4, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(2, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 3
                        '' wednesday
                        Dim tsadd As New TimeSpan(3, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(3, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 4
                        '' thurs
                        Dim tsadd As New TimeSpan(2, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(4, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 5
                        '' fri
                        Dim tsadd As New TimeSpan(1, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(5, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 6
                        '' sat
                        Dim tsadd As New TimeSpan(0, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(6, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Subtract(tssub)
                        Dim d2 As Date = Date.Today
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                End Select
                Exit Select
                Exit Select
            Case Is = "This Month"
                '' dunno just going to gather the days of the month and use the final date as d2
                ''
                '' this month is going to need the beginning of the month
                '' and the end date of the month.
                Dim DaysInMonth As Integer = 0
                Dim StartDay As Date = Date.Today.Month & "/01/" & Date.Today.Year
                DaysInMonth = Date.DaysInMonth(Date.Today.Year, Date.Today.Month)
                Dim EndDay As Date = Date.Today.Month & "/" & DaysInMonth & "/" & Date.Today.Year
                Me.retDateFrom = StartDay.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "This Month - to date"
                '' relatively the same thing here.
                '' but stop at todays date.
                Dim StartDay As Date = Date.Today.Month & "/01/" & Date.Today.Year
                Dim EndDay As Date = Date.Today.Date
                Me.retDateFrom = StartDay.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "This Year"
                '' easy stuff still. starts at 1/01/ & date.today.year
                Dim StartDay As Date = "1/01/" & Date.Today.Year
                '' How Many days in december?"
                Dim DaysInDecember As Integer = Date.DaysInMonth(Date.Today.Year, 12)
                Dim EndDay As Date = "12/" & DaysInDecember & "/" & Date.Today.Year
                Me.retDateFrom = StartDay.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "This Year - to date"
                '' same thing just stop at today's date.
                ''
                Dim StartDay As Date = "1/01/" & Date.Today.Year
                Dim EndDay As Date = Date.Today.Date
                Me.retDateFrom = StartDay.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "Next Week"
                '' first determine next mondays date
                '' then move seven days out from that.
                ''
                Dim DayOfWeek As Integer = Date.Today.DayOfWeek
                Select Case DayOfWeek
                    Case Is = 0
                        '' sunday
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(7, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.Date.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 1
                        '' monday
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(6, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.Date.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 2
                        '' tuesday
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(5, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 3
                        '' wednesday
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(4, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 4
                        '' thurs
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(3, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 5
                        '' fri
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(2, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                    Case Is = 6
                        '' sat
                        Dim tsadd As New TimeSpan(6, 0, 0, 0, 0)
                        Dim tssub As New TimeSpan(1, 0, 0, 0, 0)
                        Dim d1 As Date = Date.Today.Add(tssub)
                        Dim d2 As Date = Date.Today.Add(tsadd)
                        Me.retDateFrom = d1.ToShortDateString
                        Me.retDateTo = d2.Date.ToShortDateString
                        Exit Select
                End Select
            Case Is = "Next Month"
                '' this one is pretty easy
                '' just get what month it is, and add 1 to it.
                '' need to modify this for a scenario where the month is december.
                Dim Month As Integer = Date.Today.Month
                Dim nextmonth As Integer = 0
                Dim year As Integer = Date.Today.Year
                Select Case Month
                    Case Is = 1 '' january
                        nextmonth = 2
                        Exit Select
                    Case Is = 2 '' feb
                        nextmonth = 3
                        Exit Select
                    Case Is = 3 '' march
                        nextmonth = 4
                        Exit Select
                    Case Is = 4 '' april
                        nextmonth = 5
                        Exit Select
                    Case Is = 5 '' may
                        nextmonth = 6
                        Exit Select
                    Case Is = 6 '' june
                        nextmonth = 7
                        Exit Select
                    Case Is = 7 '' july
                        nextmonth = 8
                        Exit Select
                    Case Is = 8 '' august
                        nextmonth = 9
                        Exit Select
                    Case Is = 9 '' sept
                        nextmonth = 10
                        Exit Select
                    Case Is = 10 '' oct
                        nextmonth = 11
                        Exit Select
                    Case Is = 11 '' nov
                        nextmonth = 12
                        Exit Select
                    Case Is = 12 '' dec
                        nextmonth = 1
                        year = year + 1
                        Exit Select
                End Select
                Dim StartMonth As Integer = nextmonth
                Dim StartDate As Date = nextmonth & "/01/" & year
                Dim DaysInNextMonth As Integer = Date.DaysInMonth(year, StartMonth)
                Dim ts As New TimeSpan(DaysInNextMonth, 0, 0, 0, 0)
                Dim EndDate As Date = nextmonth & "/" & DaysInNextMonth & "/" & year
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDate.ToShortDateString
                Exit Select
            Case Is = "Last Week"
                '' first determine the date of last monday
                '' in order to do that, determine what day it is now and subtract the correct amount of days from it.
                Dim DayItIs As Integer = Date.Today.DayOfWeek
                Dim StartDate As Date = Date.Today.Date
                Select Case DayItIs
                    Case Is = 0 '' sunday subtract 7
                        Dim ts As New TimeSpan(7, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 1 '' Monday Subtract 8
                        Dim ts As New TimeSpan(8, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 2 '' Tuesday Subtract 9
                        Dim ts As New TimeSpan(9, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 3 '' Wednesday subtract 10
                        Dim ts As New TimeSpan(10, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 4 '' Thursday subtract 11
                        Dim ts As New TimeSpan(11, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 5 '' Friday subtract 12
                        Dim ts As New TimeSpan(12, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 6 '' Saturday subtract 13
                        Dim ts As New TimeSpan(13, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                End Select
                Dim ts2 As New TimeSpan(6, 0, 0, 0, 0)
                Dim EndDay As Date = StartDate.Add(ts2)
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "Last Week - to date"
                Dim DayItIs As Integer = Date.Today.DayOfWeek
                Dim StartDate As Date = Date.Today.Date
                Select Case DayItIs
                    Case Is = 0 '' sunday subtract 7
                        Dim ts As New TimeSpan(7, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 1 '' Monday Subtract 8
                        Dim ts As New TimeSpan(8, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 2 '' Tuesday Subtract 9
                        Dim ts As New TimeSpan(9, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 3 '' Wednesday subtract 10
                        Dim ts As New TimeSpan(10, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 4 '' Thursday subtract 11
                        Dim ts As New TimeSpan(11, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 5 '' Friday subtract 12
                        Dim ts As New TimeSpan(12, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                    Case Is = 6 '' Saturday subtract 13
                        Dim ts As New TimeSpan(13, 0, 0, 0, 0)
                        StartDate = StartDate.Subtract(ts)
                        Exit Select
                End Select
                Dim ts2 As New TimeSpan(7, 0, 0, 0, 0)
                Dim EndDay As Date = Date.Today.Subtract(ts2)
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDay.ToShortDateString
                Exit Select
            Case Is = "Last Month"
                Dim ThisMonth As Integer = Date.Today.Month
                Dim PreviousMonth As Integer = 0
                Dim Year As Integer = Date.Today.Year
                Select Case ThisMonth
                    Case Is = 1
                        PreviousMonth = 12
                        Year = Year - 1
                        Exit Select
                    Case Is = 2
                        PreviousMonth = 1
                        Exit Select
                    Case Is = 3
                        PreviousMonth = 2
                        Exit Select
                    Case Is = 4
                        PreviousMonth = 3
                        Exit Select
                    Case Is = 5
                        PreviousMonth = 4
                        Exit Select
                    Case Is = 6
                        PreviousMonth = 5
                        Exit Select
                    Case Is = 7
                        PreviousMonth = 6
                        Exit Select
                    Case Is = 8
                        PreviousMonth = 7
                        Exit Select
                    Case Is = 9
                        PreviousMonth = 8
                        Exit Select
                    Case Is = 10
                        PreviousMonth = 9
                        Exit Select
                    Case Is = 11
                        PreviousMonth = 10
                        Exit Select
                    Case Is = 12
                        PreviousMonth = 11
                        Exit Select
                End Select
                Dim StartDate As Date = PreviousMonth & "/01/" & Year
                Dim DaysInMonth As Integer = Date.DaysInMonth(Year, PreviousMonth)
                Dim EndDate As Date = PreviousMonth & "/" & DaysInMonth & "/" & Year
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDate.ToShortDateString
                Exit Select
            Case Is = "Last Month - to date"
                Dim ThisMonth As Integer = Date.Today.Month
                Dim PreviousMonth As Integer = 0
                Dim Year As Integer = Date.Today.Year
                Select Case ThisMonth
                    Case Is = 1
                        PreviousMonth = 12
                        Year = Year - 1
                        Exit Select
                    Case Is = 2
                        PreviousMonth = 1
                        Exit Select
                    Case Is = 3
                        PreviousMonth = 2
                        Exit Select
                    Case Is = 4
                        PreviousMonth = 3
                        Exit Select
                    Case Is = 5
                        PreviousMonth = 4
                        Exit Select
                    Case Is = 6
                        PreviousMonth = 5
                        Exit Select
                    Case Is = 7
                        PreviousMonth = 6
                        Exit Select
                    Case Is = 8
                        PreviousMonth = 7
                        Exit Select
                    Case Is = 9
                        PreviousMonth = 8
                        Exit Select
                    Case Is = 10
                        PreviousMonth = 9
                        Exit Select
                    Case Is = 11
                        PreviousMonth = 10
                        Exit Select
                    Case Is = 12
                        PreviousMonth = 11
                        Exit Select
                End Select
                Dim StartDate As Date = PreviousMonth & "/01/" & Year
                Dim DaysInMonth As Integer = Date.DaysInMonth(Year, PreviousMonth)
                Dim ts As New TimeSpan(1, 0, 0, 0, 0)
                Dim EndDate As Date = PreviousMonth & "/" & Date.Today.Day.ToString & "/" & Year
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDate.ToShortDateString
                Exit Select
            Case Is = "Last Year"
                Dim StartYear As Integer = Date.Today.Year
                Dim LastYear As Integer = (Date.Today.Year - 1)
                Dim StartDate As Date = "1/01/" & LastYear
                Dim DaysInDecember As Integer = Date.DaysInMonth(LastYear, 12)
                Dim EndDate As Date = "12/" & DaysInDecember & "/" & LastYear
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDate.ToShortDateString
                Exit Select
            Case Is = "Last Year - to date"
                Dim StartYear As Integer = Date.Today.Year
                Dim LastYear As Integer = (Date.Today.Year - 1)
                Dim StartDate As Date = "1/01/" & LastYear
                Dim EndDate As Date = Date.Today.Month & "/" & Date.Today.Day & "/" & LastYear
                Me.retDateFrom = StartDate.ToShortDateString
                Me.retDateTo = EndDate.ToShortDateString
                Exit Select
            Case Is = "Custom"
                MsgBox("Need Custom Logic here.", MsgBoxStyle.Information)
                Exit Select
        End Select
    End Sub
End Class
