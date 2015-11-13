Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic
Public Class TimeFormat
    Public RetTime As String = ""
    Private CorrectTime As String = ""

    Public Function CheckTimeFormat(ByVal time As String)
        Try
            Dim x As String = ""
            x = time
            If x = "" Then
                Exit Function
            End If

            Select Case time.ToString
                Case Is = "1p", "1 p", "1 P", "1 PM", "1PM", "1 Pm", "1 pm", "1pm", "1:00 PM", "1:00 P", "01:00 P", "01:00 PM", "01:00  PM", "1:00  PM"
                    RetTime = "01:00 PM"
                    Exit Select
                Case Is = "115p", "115 p", "115 P", "115 PM", "115PM", "115 Pm", "115 pm", "115pm", "1:15 PM", "1:15 P", "01:15 P", "01:15 PM", "01:15  PM", "1:15  PM"
                    RetTime = "01:15 PM"
                    Exit Select
                Case Is = "130p", "130 p", "130 P", "130 PM", "130PM", "130 Pm", "130 pm", "130pm", "1:30 PM", "1:30 P", "01:30 P", "01:30 PM", "01:30  PM", "1:30  PM"
                    RetTime = "01:30 PM"
                    Exit Select
                Case Is = "145p", "145 p", "145 P", "145 PM", "145PM", "145 Pm", "145 pm", "145pm", "1:45 PM", "1:45 P", "01:45 P", "01:45 PM", "01:45  PM", "1:45  PM"
                    RetTime = "01:45 PM"
                    Exit Select
                Case Is = "2p", "2 p", "2 P", "2 PM", "2PM", "2 Pm", "2 pm", "2pm", "2:00 PM", "2:00 P", "02:00 P", "02:00 PM", "02:00  PM", "2:00  PM"
                    RetTime = "02:00 PM"
                    Exit Select
                Case Is = "215p", "215 p", "215 P", "215 PM", "215PM", "215 Pm", "215 pm", "215pm", "2:15 PM", "2:15 P", "02:15 P", "02:15 PM", "02:15  PM", "2:15  PM"
                    RetTime = "02:15 PM"
                    Exit Select
                Case Is = "230p", "230 p", "230 P", "230 PM", "230PM", "230 Pm", "230 pm", "230pm", "2:30 PM", "2:30 P", "02:30 P", "02:30 PM", "02:30  PM", "2:30  PM"
                    RetTime = "02:30 PM"
                    Exit Select
                Case Is = "245p", "245 p", "245 P", "245 PM", "245PM", "245 Pm", "245 pm", "245pm", "2:45 PM", "2:45 P", "02:45 P", "02:45 PM", "02:45  PM", "2:45  PM"
                    RetTime = "02:45 PM"
                    Exit Select
                Case Is = "3p", "3 p", "3 P", "3 PM", "3PM", "3 Pm", "3 pm", "3pm", "3:00 PM", "3:00 P", "03:00 P", "03:00 PM", "03:00  PM", "3:00  PM"
                    RetTime = "03:00 PM"
                    Exit Select
                Case Is = "315p", "315 p", "315 P", "315 PM", "315PM", "315 Pm", "315 pm", "315pm", "3:15 PM", "3:15 P", "03:15 P", "03:15 PM", "03:15  PM", "3:15  PM"
                    RetTime = "03:15 PM"
                    Exit Select
                Case Is = "330p", "330 p", "330 P", "330 PM", "330PM", "330 Pm", "330 pm", "330pm", "3:30 PM", "3:30 P", "03:30 P", "03:30 PM", "03:30  PM", "3:30  PM"
                    RetTime = "03:30 PM"
                    Exit Select
                Case Is = "345p", "345 p", "345 P", "345 PM", "345PM", "345 Pm", "345 pm", "345pm", "3:45 PM", "3:45 P", "03:45 P", "03:45 PM", "03:45  PM", "3:45  PM"
                    RetTime = "03:45 PM"
                    Exit Select
                Case Is = "4p", "4 p", "4 P", "4 PM", "4PM", "4 Pm", "4 pm", "4pm", "4:00 PM", "4:00 P", "04:00 P", "04:00 PM", "04:00  PM", "4:00  PM"
                    RetTime = "04:00 PM"
                    Exit Select
                Case Is = "415p", "415 p", "415 P", "415 PM", "415PM", "415 Pm", "415 pm", "415pm", "4:15 PM", "4:15 P", "04:15 P", "04:15 PM", "04:15  PM", "4:15  PM"
                    RetTime = "04:15 PM"
                    Exit Select
                Case Is = "430p", "430 p", "430 P", "430 PM", "430PM", "430 Pm", "430 pm", "430pm", "4:30 PM", "4:30 P", "04:30 P", "04:30 PM", "04:30  PM", "4:30  PM"
                    RetTime = "04:30 PM"
                    Exit Select
                Case Is = "445p", "445 p", "445 P", "445 PM", "445PM", "445 Pm", "445 pm", "445pm", "4:45 PM", "4:45 P", "04:45 P", "04:45 PM", "04:45  PM", "4:45  PM"
                    RetTime = "04:45 PM"
                    Exit Select
                Case Is = "5p", "5 p", "5 P", "5 PM", "5PM", "5 Pm", "5 pm", "5pm", "5:00 PM", "5:00 P", "05:00 P", "05:00 PM", "05:00  PM", "5:00  PM"
                    RetTime = "05:00 PM"
                    Exit Select
                Case Is = "515p", "515 p", "515 P", "515 PM", "515PM", "515 Pm", "515 pm", "515pm", "5:15 PM", "5:15 P", "05:15 P", "05:15 PM", "05:15  PM", "5:15  PM"
                    RetTime = "05:15 PM"
                    Exit Select
                Case Is = "530p", "530 p", "530 P", "530 PM", "530PM", "530 Pm", "530 pm", "530pm", "5:30 PM", "5:30 P", "05:30 P", "05:30 PM", "05:30  PM", "5:30  PM"
                    RetTime = "05:30 PM"
                    Exit Select
                Case Is = "545p", "545 p", "545 P", "545 PM", "545PM", "545 Pm", "545 pm", "545pm", "5:45 PM", "5:45 P", "05:45 P", "05:45 PM", "05:45  PM", "5:45  PM"
                    RetTime = "05:45 PM"
                    Exit Select
                Case Is = "6p", "6 p", "6 P", "6 PM", "6PM", "6 Pm", "6 pm", "6pm", "6:00 PM", "6:00 P", "06:00 P", "06:00 PM", "06:00  PM", "6:00  PM"
                    RetTime = "06:00 PM"
                    Exit Select
                Case Is = "615p", "615 p", "615 P", "615 PM", "615PM", "615 Pm", "615 pm", "615pm", "6:15 PM", "6:15 P", "06:15 P", "06:15 PM", "06:15  PM", "6:15  PM"
                    RetTime = "06:15 PM"
                    Exit Select
                Case Is = "630p", "630 p", "630 P", "630 PM", "630PM", "630 Pm", "630 pm", "630pm", "6:30 PM", "6:30 P", "06:30 P", "06:30 PM", "06:30  PM", "6:30  PM"
                    RetTime = "06:30 PM"
                    Exit Select
                Case Is = "645p", "645 p", "645 P", "645 PM", "645PM", "645 Pm", "645 pm", "645pm", "6:45 PM", "6:45 P", "06:45 P", "06:45 PM", "06:45  PM", "6:45  PM"
                    RetTime = "06:45 PM"
                    Exit Select
                Case Is = "7p", "7 p", "7 P", "7 PM", "7PM", "7 Pm", "7 pm", "7pm", "7:00 PM", "7:00 P", "07:00 P", "07:00 PM", "07:00  PM", "7:00  PM"
                    RetTime = "07:00 PM"
                    Exit Select
                Case Is = "715p", "715 p", "715 P", "715 PM", "715PM", "715 Pm", "715 pm", "715pm", "7:15 PM", "7:15 P", "07:15 P", "07:15 PM", "07:15  PM", "7:15  PM"
                    RetTime = "07:15 PM"
                    Exit Select
                Case Is = "730p", "730 p", "730 P", "730 PM", "730PM", "730 Pm", "730 pm", "730pm", "7:30 PM", "7:30 P", "07:30 P", "07:30 PM", "07:30  PM", "7:30  PM"
                    RetTime = "07:30 PM"
                    Exit Select
                Case Is = "745p", "745 p", "745 P", "745 PM", "745PM", "745 Pm", "745 pm", "745pm", "7:45 PM", "7:45 P", "07:45 P", "07:45 PM", "07:45  PM", "7:45  PM"
                    RetTime = "07:45 PM"
                    Exit Select
                Case Is = "8p", "8 p", "8 P", "8 PM", "8PM", "8 Pm", "8 pm", "8pm", "8:00 PM", "8:00 P", "08:00 P", "08:00 PM", "08:00  PM", "8:00  PM"
                    RetTime = "08:00 PM"
                    Exit Select
                Case Is = "815p", "815 p", "815 P", "815 PM", "815PM", "815 Pm", "815 pm", "815pm", "8:15 PM", "8:15 P", "08:15 P", "08:15 PM", "08:15  PM", "8:15  PM"
                    RetTime = "08:15 PM"
                    Exit Select
                Case Is = "830p", "830 p", "830 P", "830 PM", "830PM", "830 Pm", "830 pm", "830pm", "8:30 PM", "8:30 P", "08:30 P", "08:30 PM", "08:45  PM", "8:45  PM"
                    RetTime = "08:30 PM"
                    Exit Select
                Case Is = "845p", "845 p", "845 P", "845 PM", "845PM", "845 Pm", "845 pm", "845pm", "8:45 PM", "8:45 P", "08:45 P", "08:45 PM", "08:45  PM", "8:45  PM"
                    RetTime = "08:45 PM"
                    Exit Select
                Case Is = "9p", "9 p", "9 P", "9 PM", "9PM", "9 Pm", "9 pm", "9pm", "9:00 PM", "9:00 P", "09:00 P", "09:00 PM", "09:00  PM", "9:00  PM"
                    RetTime = "09:00 PM"
                    Exit Select
                Case Is = "915p", "915 p", "915 P", "915 PM", "915PM", "915 Pm", "915 pm", "915pm", "9:15 PM", "9:15 P", "09:15 P", "09:15 PM", "09:15  PM", "9:15  PM"
                    RetTime = "09:15 PM"
                    Exit Select
                Case Is = "930p", "930 p", "930 P", "930 PM", "930PM", "930 Pm", "930 pm", "930pm", "9:30 PM", "9:30 P", "09:30 P", "09:30 PM", "09:30  PM", "9:30  PM"
                    RetTime = "09:30 PM"
                    Exit Select
                Case Is = "945p", "945 p", "945 P", "945 PM", "945PM", "945 Pm", "945 pm", "945pm", "9:45 PM", "9:45 P", "09:45 P", "09:45 PM", "09:45  PM", "9:45  PM"
                    RetTime = "09:45 PM"
                    Exit Select
                Case Is = "10p", "10 p", "10 P", "10 PM", "10PM", "10 Pm", "10 pm", "10pm", "10:00 PM", "10:00 P", "10:00 P", "10:00 PM", "10:00  PM", "10:00  PM"
                    RetTime = "10:00 PM"
                    Exit Select
                Case Is = "1015p", "1015 p", "1015 P", "1015 PM", "1015PM", "1015 Pm", "1015 pm", "1015pm", "10:15 PM", "10:15 P", "10:15 P", "10:15 PM", "10:15  PM", "10:15  PM"
                    RetTime = "10:15 PM"
                    Exit Select
                Case Is = "1030p", "1030 p", "1030 P", "1030 PM", "1030PM", "1030 Pm", "1030 pm", "1030pm", "10:30 PM", "10:30 P", "10:30 P", "10:30 PM", "10:30  PM", "10:30  PM"
                    RetTime = "10:30 PM"
                    Exit Select
                Case Is = "1045p", "1045 p", "1045 P", "1045 PM", "1045PM", "1045 Pm", "1045 pm", "1045pm", "10:45 PM", "10:45 P", "10:45 P", "10:45 PM", "10:45  PM", "10:45  PM"
                    RetTime = "10:45 PM"
                    Exit Select
                Case Is = "11p", "11 p", "11 P", "11 PM", "11PM", "11 Pm", "11 pm", "11pm", "11:00 PM", "11:00 P", "11:00 P", "11:00 PM", "11:00  PM", "11:00  PM"
                    RetTime = "11:00 PM"
                    Exit Select
                Case Is = "1115p", "1115 p", "1115 P", "1115 PM", "1115PM", "1115 Pm", "1115 pm", "1115pm", "11:15 PM", "11:15 P", "11:15 P", "11:15 PM", "11:15  PM", "11:15  PM"
                    RetTime = "11:15 PM"
                    Exit Select
                Case Is = "1130p", "1130 p", "1130 P", "1130 PM", "1130PM", "1130 Pm", "1130 pm", "1130pm", "11:30 PM", "11:30 P", "11:30 P", "11:30 PM", "11:30  PM", "11:30  PM"
                    RetTime = "11:30 PM"
                    Exit Select
                Case Is = "1145p", "1145 p", "1145 P", "1145 PM", "1145PM", "1145 Pm", "1145 pm", "1145pm", "11:45 PM", "11:45 P", "11:45 P", "11:45 PM", "11:45  PM", "11:45  PM"
                    RetTime = "11:45 PM"
                    Exit Select
                Case Is = "12p", "12 p", "12 P", "12 PM", "12PM", "12 Pm", "12 pm", "12pm", "12:00 PM", "12:00 P", "12:00 P", "12:00 PM", "12:00  PM", "12:00  PM"
                    RetTime = "12:00 PM"
                    Exit Select
                Case Is = "1215p", "1215 p", "1215 P", "1215 PM", "1215PM", "1215 Pm", "1215 pm", "1215pm", "12:15 PM", "12:15 P", "12:15 P", "12:15 PM", "12:15  PM", "12:15  PM"
                    RetTime = "12:15 PM"
                    Exit Select
                Case Is = "1230p", "1230 p", "1230 P", "1230 PM", "1230PM", "1230 Pm", "1230 pm", "1230pm", "12:30 PM", "12:30 P", "12:30 P", "12:30 PM", "12:30  PM", "12:30  PM"
                    RetTime = "12:30 PM"
                    Exit Select
                Case Is = "1245p", "1245 p", "1245 P", "1245 PM", "1245PM", "1245 Pm", "1245 pm", "1245pm", "12:45 PM", "12:45 P", "12:45 P", "12:45 PM", "12:45  PM", "12:45  PM"
                    RetTime = "12:45 PM"
                    Exit Select

                Case Is = "1a", "1 A", "1 a", "1 AM", "1AM", "1 Am", "1 am", "1am", "1:00 AM", "1:00 A", "01:00 A", "01:00 AM", "01:00  AM", "1:00  AM"
                    RetTime = "01:00 AM"
                    Exit Select
                Case Is = "115a", "115 A", "115 a", "115 AM", "115AM", "115 Am", "115 am", "115am", "1:15 AM", "1:15 A", "01:15 A", "01:15 AM", "01:15  AM", "1:15  AM"
                    RetTime = "01:15 AM"
                    Exit Select
                Case Is = "130a", "130 A", "130 a", "130 AM", "130AM", "130 Am", "130 am", "130am", "1:30 AM", "1:30 A", "01:30 A", "01:30 AM", "01:30  AM", "1:30  AM"
                    RetTime = "01:30 AM"
                    Exit Select
                Case Is = "145a", "145 A", "145 a", "145 AM", "145AM", "145 Am", "145 am", "145am", "1:45 AM", "1:45 A", "01:45 A", "01:45 AM", "01:45  AM", "1:45  AM"
                    RetTime = "01:45 AM"
                    Exit Select
                Case Is = "2a", "2 A", "2 a", "2 AM", "2AM", "2 Am", "2 am", "2am", "2:00 AM", "2:00 A", "02:00 A", "02:00 AM", "02:00  AM", "2:00  AM"
                    RetTime = "02:00 AM"
                    Exit Select
                Case Is = "215a", "215 A", "215 a", "215 AM", "215AM", "215 Am", "215 am", "215am", "2:15 AM", "2:15 A", "02:15 A", "02:15 AM", "02:15  AM", "2:15  AM"
                    RetTime = "02:15 AM"
                    Exit Select
                Case Is = "230a", "230 A", "230 a", "230 AM", "230AM", "230 Am", "230 am", "230am", "2:30 AM", "2:30 A", "02:30 A", "02:30 AM", "02:30  AM", "2:30  AM"
                    RetTime = "02:30 AM"
                    Exit Select
                Case Is = "245a", "245 A", "245 a", "245 AM", "245AM", "245 Am", "245 am", "245am", "2:45 AM", "2:45 A", "02:45 A", "02:45 AM", "02:45  AM", "2:45  AM"
                    RetTime = "02:45 AM"
                    Exit Select
                Case Is = "3a", "3 A", "3 a", "3 AM", "3AM", "3 Am", "3 am", "3am", "3:00 AM", "3:00 A", "03:00 A", "03:00 AM", "03:00  AM", "3:00  AM"
                    RetTime = "03:00 AM"
                    Exit Select
                Case Is = "315a", "315 A", "315 a", "315 AM", "315AM", "315 Am", "315 am", "315am", "3:15 AM", "3:15 A", "03:15 A", "03:15 AM", "03:15  AM", "3:15  AM"
                    RetTime = "03:15 AM"
                    Exit Select
                Case Is = "330a", "330 A", "330 a", "330 AM", "330AM", "330 Am", "330 am", "330am", "3:30 AM", "3:30 A", "03:30 A", "03:30 AM", "03:30  AM", "3:30  AM"
                    RetTime = "03:30 AM"
                    Exit Select
                Case Is = "345a", "345 A", "345 a", "345 AM", "345AM", "345 Am", "345 am", "345am", "3:45 AM", "3:45 A", "03:45 A", "03:45 AM", "03:45  AM", "3:45  AM"
                    RetTime = "03:45 AM"
                    Exit Select
                Case Is = "4a", "4 A", "4 a", "4 AM", "4AM", "4 Am", "4 am", "4am", "4:00 AM", "4:00 A", "04:00 A", "04:00 AM", "04:00  AM", "4:00  AM"
                    RetTime = "04:00 AM"
                    Exit Select
                Case Is = "415a", "415 A", "415 a", "415 AM", "415AM", "415 Am", "415 am", "415am", "4:15 AM", "4:15 A", "04:15 A", "04:15 AM", "04:15  AM", "4:15  AM"
                    RetTime = "04:15 AM"
                    Exit Select
                Case Is = "430a", "430 A", "430 a", "430 AM", "430AM", "430 Am", "430 am", "430am", "4:30 AM", "4:30 A", "04:30 A", "04:30 AM", "04:30  AM", "4:30  AM"
                    RetTime = "04:30 AM"
                    Exit Select
                Case Is = "445a", "445 A", "445 a", "445 AM", "445AM", "445 Am", "445 am", "445am", "4:45 AM", "4:45 A", "04:45 A", "04:45 AM", "04:45  AM", "4:45  AM"
                    RetTime = "04:45 AM"
                    Exit Select
                Case Is = "5a", "5 A", "5 a", "5 AM", "5AM", "5 Am", "5 am", "5am", "5:00 AM", "5:00 A", "05:00 A", "05:00 AM", "05:00  AM", "5:00  AM"
                    RetTime = "05:00 AM"
                    Exit Select
                Case Is = "515a", "515 A", "515 a", "515 AM", "515AM", "515 Am", "515 am", "515am", "5:15 AM", "5:15 A", "05:15 A", "05:15 AM", "05:15  AM", "5:15  AM"
                    RetTime = "05:15 AM"
                    Exit Select
                Case Is = "530a", "530 A", "530 a", "530 AM", "530AM", "530 Am", "530 am", "530am", "5:30 AM", "5:30 A", "05:30 A", "05:30 AM", "05:30  AM", "5:30  AM"
                    RetTime = "05:30 AM"
                    Exit Select
                Case Is = "545a", "545 A", "545 a", "545 AM", "545AM", "545 Am", "545 am", "545am", "5:45 AM", "5:45 A", "05:45 A", "05:45 AM", "05:45  AM", "5:45  AM"
                    RetTime = "05:45 AM"
                    Exit Select
                Case Is = "6a", "6 A", "6 a", "6 AM", "6AM", "6 Am", "6 am", "6am", "6:00 AM", "6:00 A", "06:00 A", "06:00 AM", "06:00  AM", "6:00  AM"
                    RetTime = "06:00 AM"
                    Exit Select
                Case Is = "615a", "615 A", "615 a", "615 AM", "615AM", "615 Am", "615 am", "615am", "6:15 AM", "6:15 A", "06:15 A", "06:15 AM", "06:15  AM", "6:15  AM"
                    RetTime = "06:15 AM"
                    Exit Select
                Case Is = "630a", "630 A", "630 a", "630 AM", "630AM", "630 Am", "630 am", "630am", "6:30 AM", "6:30 A", "06:30 A", "06:30 AM", "06:30  AM", "6:30  AM"
                    RetTime = "06:30 AM"
                    Exit Select
                Case Is = "645a", "645 A", "645 a", "645 AM", "645AM", "645 Am", "645 am", "645am", "6:45 AM", "6:45 A", "06:45 A", "06:45 AM", "06:45  AM", "6:45  AM"
                    RetTime = "06:45 AM"
                    Exit Select
                Case Is = "7a", "7 A", "7 a", "7 AM", "7AM", "7 Am", "7 am", "7am", "7:00 AM", "7:00 A", "07:00 A", "07:00 AM", "07:00  AM", "7:00  AM"
                    RetTime = "07:00 AM"
                    Exit Select
                Case Is = "715a", "715 A", "715 a", "715 AM", "715AM", "715 Am", "715 am", "715am", "7:15 AM", "7:15 A", "07:15 A", "07:15 AM", "07:15  AM", "7:15  AM"
                    RetTime = "07:15 AM"
                    Exit Select
                Case Is = "730a", "730 A", "730 a", "730 AM", "730AM", "730 Am", "730 am", "730am", "7:30 AM", "7:30 A", "07:30 A", "07:30 AM", "07:30  AM", "7:30  AM"
                    RetTime = "07:30 AM"
                    Exit Select
                Case Is = "745a", "745 A", "745 a", "745 AM", "745AM", "745 Am", "745 am", "745am", "7:45 AM", "7:45 A", "07:45 A", "07:45 AM", "07:45  AM", "7:45  AM"
                    RetTime = "07:45 AM"
                    Exit Select
                Case Is = "8a", "8 A", "8 a", "8 AM", "8AM", "8 Am", "8 am", "8am", "8 AM", "8 A", "08 A", "08 AM", "08:00  AM", "8:00  AM"
                    RetTime = "08:00 AM"
                    Exit Select
                Case Is = "815a", "815 A", "815 a", "815 AM", "815AM", "815 Am", "815 am", "815am", "815 AM", "815 A", "0815 A", "0815 AM", "08:15  AM", "8:15  AM"
                    RetTime = "08:15 AM"
                    Exit Select
                Case Is = "830a", "830 A", "830 a", "830 AM", "830AM", "830 Am", "830 am", "830am", "8:30 AM", "8:30 A", "08:30 A", "08:30 AM", "08:30  AM", "8:30  AM"
                    RetTime = "08:30 AM"
                    Exit Select
                Case Is = "845a", "845 A", "845 a", "845 AM", "845AM", "845 Am", "845 am", "845am", "8:45 AM", "8:45 A", "08:45 A", "08:45 AM", "08:45  AM", "8:45  AM"
                    RetTime = "08:45 AM"
                    Exit Select
                Case Is = "915a", "915 A", "915 a", "915 AM", "915AM", "915 Am", "915 am", "915am", "9:15 AM", "9:15 A", "09:15 A", "09:15 AM", "09:15  AM", "9:15  AM"
                    RetTime = "09:15 AM"
                    Exit Select
                Case Is = "930a", "930 A", "930 a", "930 AM", "930AM", "930 Am", "930 am", "930am", "9:30 AM", "9:30 A", "09:30 A", "09:30 AM", "09:30  AM", "9:30  AM"
                    RetTime = "09:30 AM"
                    Exit Select
                Case Is = "945a", "945 A", "945 a", "945 AM", "945AM", "945 Am", "945 am", "945am", "9:45 AM", "9:45 A", "09:45 A", "09:45 AM", "09:45  AM", "9:45  AM"
                    RetTime = "09:45 AM"
                    Exit Select
                Case Is = "9a", "9 A", "9 a", "9 AM", "9AM", "9 Am", "9 am", "9am", "9:00 AM", "9:00 A", "09:00 A", "09:00 AM", "09:00  AM", "9:00  AM"
                    RetTime = "09:30 AM"
                    Exit Select

                Case Is = "10a", "10 A", "10 a", "10 AM", "10AM", "10 Am", "10 am", "10am", "10:00 AM", "10:00 A", "10:00 A", "10:00 AM", "10:00  AM", "10:00  AM"
                    RetTime = "10:00 AM"
                    Exit Select
                Case Is = "1015a", "1015 A", "1015 a", "1015 AM", "1015AM", "1015 Am", "1015 am", "1015am", "10:15 AM", "10:15 A", "10:15 A", "10:15 AM", "10:15  AM", "10:15  AM"
                    RetTime = "10:15 AM"
                    Exit Select
                Case Is = "1030a", "1030 A", "1030 a", "1030 AM", "1030AM", "1030 Am", "1030 am", "1030am", "10:30 AM", "10:30 A", "10:30 A", "10:30 AM", "10:30  AM", "10:30  AM"
                    RetTime = "10:30 AM"
                    Exit Select
                Case Is = "1045a", "1045 A", "1045 a", "1045 AM", "1045AM", "1045 Am", "1045 am", "1045am", "10:45 AM", "10:45 A", "10:45 A", "10:45 AM", "10:45  AM", "10:45  AM"
                    RetTime = "10:45 AM"
                    Exit Select
                Case Is = "1130a", "1130 A", "1130 a", "1130 AM", "1130AM", "1130 Am", "1130 am", "1130am", "11:30 AM", "11:30 A", "11:30 A", "11:30 AM", "11:30  AM", "11:30  AM"
                    RetTime = "11:30 AM"
                    Exit Select
                Case Is = "1145a", "1145 A", "1145 a", "1145 AM", "1145AM", "1145 Am", "1145 am", "1145am", "11:45 AM", "11:45 A", "11:45 A", "11:45 AM", "11:45  AM", "11:45  AM"
                    RetTime = "11:45 AM"
                    Exit Select
                Case Is = "11a", "11 A", "11 a", "11 AM", "11AM", "11 Am", "11 am", "11am", "11:00 AM", "11:00 A", "11:00 A", "11:00 AM", "11:00  AM", "11:00  AM"
                    RetTime = "11:00 AM"
                    Exit Select
                Case Is = "1115a", "1115 A", "1115 a", "1115 AM", "1115AM", "1115 Am", "1115 am", "1115am", "11:15 AM", "11:15 A", "11:15 A", "11:15 AM", "11:15  AM", "11:15  AM"
                    RetTime = "11:15 AM"
                    Exit Select
                Case Is = "12a", "12 A", "12 a", "12 AM", "12AM", "12 Am", "12 am", "12am", "12:00 AM", "12:00 A", "12:00 A", "12:00 AM", "12:00  AM", "12:00  AM"
                    RetTime = "12:00 AM"
                    Exit Select
                Case Is = "1215a", "1215 A", "1215 a", "1215 AM", "1215AM", "1215 Am", "1215 am", "1215am", "12:15 AM", "12:15 A", "12:15 A", "12:15 AM", "12:15  AM", "12:15  AM"
                    RetTime = "12:15 AM"
                    Exit Select
                Case Is = "1230a", "1230 A", "1230 a", "1230 AM", "1230AM", "1230 Am", "1230 am", "1230am", "12:30 AM", "12:30 A", "12:30 A", "12:30 AM", "12:30  AM", "12:30  AM"
                    RetTime = "12:30 AM"
                    Exit Select
                Case Is = "1245a", "1245 A", "1245 a", "1245 AM", "1245AM", "1245 Am", "1245 am", "1245am", "12:45 AM", "12:45 A", "12:45 A", "12:45 AM", "12:45  AM", "12:45  AM"
                    RetTime = "12:45 AM"
                    Exit Select
                Case Else
                    MsgBox("Incorrect time format. Time format example: '1p', '1PM', '115pm', '1:45 PM', '130PM', '130pm'", MsgBoxStyle.Exclamation, "Format Time Error")
                    RetTime = ""
                    Exit Select
            End Select
            Return RetTime
        Catch ex As Exception
            Return RetTime
            Dim err As New ErrorLogFlatFile
            err.WriteLog("TimeFormat", "ByVal time As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CheckTimeFormat")

        End Try

    End Function
    Public Sub ConvertTimeFromSQL(ByVal Time As String)
        Try
            Dim d1
            d1 = Split(Time, " ", 2)
            Dim FirstPart As String = ""
            Dim SecondPart As String = ""
            FirstPart = d1(0).ToString  '' Date part IE  1/1/1900
            SecondPart = d1(1).ToString  '' Time part iE 2:15:00 PM

            '' now break apart time even more to get just hour and mintues with AM or PM
            ''

            Dim t3
            t3 = Split(SecondPart, ":", 3)
            Dim Hour As String = ""
            Hour = t3(0).ToString

            Dim minute As String = ""
            minute = t3(1).ToString

            Dim AMPM
            AMPM = Split(t3(2), " ", 2)
            Dim Designator As String = ""
            Designator = AMPM(1).ToString
            CorrectTime = Hour & ":" & minute & " " & Designator
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("TimeFormat", "ByVal Time As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "ConvertTimeFromSQL")

        End Try


    End Sub
   
    Public Property TimeFromSQLTable()
        Get
            Return correcttime
        End Get
        Set(ByVal value)
            correcttime = value
        End Set
    End Property
End Class
