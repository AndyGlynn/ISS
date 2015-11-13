Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class WarmCalling
    Dim PLS As String = WCaller.cboPLSWarmCalling.Text
    Dim SLS As String = WCaller.cboSLSWarmCalling.Text
    Dim Sat As String
    Dim Sun As String
    Dim D1 As String
    Dim D2 As String
    Dim T1 As String
    Dim T2 As String
    Dim R1 As String
    Dim R2 As String
    Dim R3 As String
    Dim R4 As String
    Dim R5 As String
    Dim Z1 As String
    Dim Z2 As String
    Dim Z3 As String
    Dim Z4 As String
    Dim Z5 As String
    Dim Z6 As String
    Dim Z7 As String
    Dim Z8 As String
    Dim Z9 As String
    Dim Z10 As String
    Dim C1 As String
    Dim C2 As String
    Dim C3 As String
    Dim C4 As String
    Dim C5 As String
    Dim C6 As String
    Dim C7 As String
    Dim C8 As String
    Dim C9 As String
    Dim C10 As String
    Dim GB As String = WCaller.cboGroupBy.Text
    Public Sub GroupBy()

        WCaller.lvWarmCalling.Groups.Clear()
        If PLS = "" Then
            PLS = "%"
        End If
        If SLS = "" Then
            SLS = "%"
        End If
        If WCaller.cbWeekdays.Checked = True Then
            Sat = "Saturday"
            Sun = "Sunday"
        Else
            Sat = "Widget"
            Sun = "Widget"
        End If
        If WCaller.txtDate1.Text = "" And WCaller.txtDate2.Text = "" Then
            D1 = "1/1/1900 12:00 AM"
            D2 = "1/1/2100 12:00 AM"
        ElseIf WCaller.txtDate1.Text <> "" And WCaller.txtDate2.Text = "" Then
            D1 = WCaller.txtDate1.Text & " 12:00 AM"
            D2 = "1/1/2100 12:00 AM"
        ElseIf WCaller.txtDate1.Text = "" And WCaller.txtDate2.Text <> "" Then
            D1 = "1/1/1900 12:00 AM"
            D2 = WCaller.txtDate2.Text & " 12:00 AM"
        Else
            D1 = WCaller.txtDate1.Text & " 12:00 AM"
            D2 = WCaller.txtDate2.Text & " 12:00 AM"
        End If
        If WCaller.txtTime1.Text = "" And WCaller.txtTime2.Text = "" Then
            T1 = "1/1/1900 12:00 AM"
            T2 = "1/1/1900 11:59 PM"
        Else
            T1 = "1/1/1900 " & WCaller.dtpTime1.Value.ToShortTimeString
            T2 = "1/1/1900 " & WCaller.dptTime2.Value.ToShortTimeString
        End If

        Dim a = WCaller.chlstResults.CheckedItems
        Dim i As Integer = WCaller.chlstResults.CheckedItems.Count
        Select Case i
            Case Is = 1
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(0).ToString()
                R3 = WCaller.chlstResults.CheckedItems(0).ToString()
                R4 = WCaller.chlstResults.CheckedItems(0).ToString()
                R5 = WCaller.chlstResults.CheckedItems(0).ToString()
            Case Is = 2
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(1).ToString()
                R4 = WCaller.chlstResults.CheckedItems(1).ToString()
                R5 = WCaller.chlstResults.CheckedItems(1).ToString()
            Case Is = 3
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(2).ToString()
                R5 = WCaller.chlstResults.CheckedItems(2).ToString()
            Case Is = 4
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(3).ToString()
                R5 = WCaller.chlstResults.CheckedItems(3).ToString()
            Case Is = 5
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(3).ToString()
                R5 = WCaller.chlstResults.CheckedItems(4).ToString()
        End Select
        Dim z As Integer = WCaller.lbZipCity.CheckedItems.Count
        If z <> 0 Then
            If InStr(WCaller.btnZipCity.Text, "Zip Codes") <> 0 Then
                Select Case z
                    Case Is = 1
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 2
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 3
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 4
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 5
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 6
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 7
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 8
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 9
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 10
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Else
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                End Select
            Else
                Select Case z
                    Case Is = 1
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 2
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 3
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 4
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 5
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 6
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 7
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 8
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 9
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 10
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Else
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                End Select
            End If
        Else
            Z1 = "%"
            Z2 = "%"
            Z3 = "%"
            Z4 = "%"
            Z5 = "%"
            Z6 = "%"
            Z7 = "%"
            Z8 = "%"
            Z9 = "%"
            Z10 = "%"
            C1 = "%"
            C2 = "%"
            C3 = "%"
            C4 = "%"
            C5 = "%"
            C6 = "%"
            C7 = "%"
            C8 = "%"
            C9 = "%"
            C10 = "%"
        End If
        Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
        Dim param2 As SqlParameter = New SqlParameter("@SLS", SLS)
        Dim param3 As SqlParameter = New SqlParameter("@Date1", D1)
        Dim param4 As SqlParameter = New SqlParameter("@Date2", D2)
        Dim param5 As SqlParameter = New SqlParameter("@Time1", T1)
        Dim param6 As SqlParameter = New SqlParameter("@Time2", T2)
        Dim param7 As SqlParameter = New SqlParameter("@R1", R1)
        Dim param8 As SqlParameter = New SqlParameter("@R2", R2)
        Dim param9 As SqlParameter = New SqlParameter("@R3", R3)
        Dim param10 As SqlParameter = New SqlParameter("@R4", R4)
        Dim param11 As SqlParameter = New SqlParameter("@R5", R5)
        Dim param12 As SqlParameter = New SqlParameter("@Sat", Sat)
        Dim param13 As SqlParameter = New SqlParameter("@Sun", Sun)
        Dim param14 As SqlParameter = New SqlParameter("@Z1", Z1)
        Dim param15 As SqlParameter = New SqlParameter("@Z2", Z2)
        Dim param16 As SqlParameter = New SqlParameter("@Z3", Z3)
        Dim param17 As SqlParameter = New SqlParameter("@Z4", Z4)
        Dim param18 As SqlParameter = New SqlParameter("@Z5", Z5)
        Dim param19 As SqlParameter = New SqlParameter("@Z6", Z6)
        Dim param20 As SqlParameter = New SqlParameter("@Z7", Z7)
        Dim param21 As SqlParameter = New SqlParameter("@Z8", Z8)
        Dim param22 As SqlParameter = New SqlParameter("@Z9", Z9)
        Dim param23 As SqlParameter = New SqlParameter("@Z10", Z10)
        Dim param24 As SqlParameter = New SqlParameter("@C1", C1)
        Dim param25 As SqlParameter = New SqlParameter("@C2", C2)
        Dim param26 As SqlParameter = New SqlParameter("@C3", C3)
        Dim param27 As SqlParameter = New SqlParameter("@C4", C4)
        Dim param28 As SqlParameter = New SqlParameter("@C5", C5)
        Dim param29 As SqlParameter = New SqlParameter("@C6", C6)
        Dim param30 As SqlParameter = New SqlParameter("@C7", C7)
        Dim param31 As SqlParameter = New SqlParameter("@C8", C8)
        Dim param32 As SqlParameter = New SqlParameter("@C9", C9)
        Dim param33 As SqlParameter = New SqlParameter("@C10", C10)
        Dim param34 As SqlParameter = New SqlParameter("@GB", WCaller.cboGroupBy.Text)

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.WarmCallingOrderBy", cnn)
        Dim r As SqlDataReader
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.Parameters.Add(param6)
        cmdGet.Parameters.Add(param7)
        cmdGet.Parameters.Add(param8)
        cmdGet.Parameters.Add(param9)
        cmdGet.Parameters.Add(param10)
        cmdGet.Parameters.Add(param11)
        cmdGet.Parameters.Add(param12)
        cmdGet.Parameters.Add(param13)
        cmdGet.Parameters.Add(param14)
        cmdGet.Parameters.Add(param15)
        cmdGet.Parameters.Add(param16)
        cmdGet.Parameters.Add(param17)
        cmdGet.Parameters.Add(param18)
        cmdGet.Parameters.Add(param19)
        cmdGet.Parameters.Add(param20)
        cmdGet.Parameters.Add(param21)
        cmdGet.Parameters.Add(param22)
        cmdGet.Parameters.Add(param23)
        cmdGet.Parameters.Add(param24)
        cmdGet.Parameters.Add(param25)
        cmdGet.Parameters.Add(param26)
        cmdGet.Parameters.Add(param27)
        cmdGet.Parameters.Add(param28)
        cmdGet.Parameters.Add(param29)
        cmdGet.Parameters.Add(param30)
        cmdGet.Parameters.Add(param31)
        cmdGet.Parameters.Add(param32)
        cmdGet.Parameters.Add(param33)
        cmdGet.Parameters.Add(param34)


        Try
            cnn.Open()
            r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            Dim cnt As Integer = 0
            While r.Read
                Dim g As New ListViewGroup
                If WCaller.cboGroupBy.Text = "Zip Code" Then
                    g.Name = r.Item(0)
                    g.Header = r.Item(0)
                    WCaller.lvWarmCalling.Groups.Add(g)
                ElseIf WCaller.cboGroupBy.Text = "City, State" Then
                    g.Name = r.Item(0)
                    g.Header = r.Item(0) & ", " & r.Item(1)
                    WCaller.lvWarmCalling.Groups.Add(g)
                ElseIf WCaller.cboGroupBy.Text = "Primary Product" Then
                    g.Name = r.Item(0)
                    g.Header = r.Item(0)
                    WCaller.lvWarmCalling.Groups.Add(g)
                ElseIf WCaller.cboGroupBy.Text = "Primary Lead Source" Then
                    g.Name = r.Item(0)
                    g.Header = r.Item(0)
                    WCaller.lvWarmCalling.Groups.Add(g)
                ElseIf WCaller.cboGroupBy.Text = "Marketing Result" Then
                    g.Name = r.Item(0)
                    g.Header = r.Item(0)
                    WCaller.lvWarmCalling.Groups.Add(g)
                End If
            End While
            r.Close()
            cnn.Close()
            WCaller.txtRecordsMatching.Text = CStr(cnt)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Me.Populate()

    End Sub
    Dim LastID As String
    Public Sub Populate()

        If WCaller.lvWarmCalling.SelectedItems.Count = 1 Then
            LastId = WCaller.lvWarmCalling.SelectedItems(0).Text
        End If

        WCaller.lvWarmCalling.Items.Clear()

        If PLS = "" Then
            PLS = "%"
        End If
        If SLS = "" Then
            SLS = "%"
        End If
        If WCaller.cbWeekdays.Checked = True Then
            Sat = "Saturday"
            Sun = "Sunday"
        Else
            Sat = "Widget"
            Sun = "Widget"
        End If
        If WCaller.txtDate1.Text = "" And WCaller.txtDate2.Text = "" Then
            D1 = "1/1/1900 12:00 AM"
            D2 = "1/1/2100 12:00 AM"
        ElseIf WCaller.txtDate1.Text <> "" And WCaller.txtDate2.Text = "" Then
            D1 = WCaller.txtDate1.Text & " 12:00 AM"
            D2 = "1/1/2100 12:00 AM"
        ElseIf WCaller.txtDate1.Text = "" And WCaller.txtDate2.Text <> "" Then
            D1 = "1/1/1900 12:00 AM"
            D2 = WCaller.txtDate2.Text & " 12:00 AM"
        Else
            D1 = WCaller.txtDate1.Text & " 12:00 AM"
            D2 = WCaller.txtDate2.Text & " 12:00 AM"
        End If
        If WCaller.txtTime1.Text = "" And WCaller.txtTime2.Text = "" Then
            T1 = "1/1/1900 12:00 AM"
            T2 = "1/1/1900 11:59 PM"
        ElseIf WCaller.txtTime1.Text <> "" And WCaller.txtTime2.Text = "" Then
            T1 = "1/1/1900 " & WCaller.txtTime1.Text
            T2 = "1/1/1900 11:59 PM"
        ElseIf WCaller.txtTime1.Text = "" And WCaller.txtTime2.Text <> "" Then
            T1 = "1/1/1900 12:00 AM"
            T2 = "1/1/1900 " & WCaller.txtTime2.Text
        Else
            T1 = "1/1/1900 " & WCaller.txtTime1.Text
            T2 = "1/1/1900 " & WCaller.txtTime2.Text
        End If

        Dim a = WCaller.chlstResults.CheckedItems
        Dim i As Integer = WCaller.chlstResults.CheckedItems.Count
        Select Case i
            Case Is = 1
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(0).ToString()
                R3 = WCaller.chlstResults.CheckedItems(0).ToString()
                R4 = WCaller.chlstResults.CheckedItems(0).ToString()
                R5 = WCaller.chlstResults.CheckedItems(0).ToString()
            Case Is = 2
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(1).ToString()
                R4 = WCaller.chlstResults.CheckedItems(1).ToString()
                R5 = WCaller.chlstResults.CheckedItems(1).ToString()
            Case Is = 3
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(2).ToString()
                R5 = WCaller.chlstResults.CheckedItems(2).ToString()
            Case Is = 4
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(3).ToString()
                R5 = WCaller.chlstResults.CheckedItems(3).ToString()
            Case Is = 5
                R1 = WCaller.chlstResults.CheckedItems(0).ToString()
                R2 = WCaller.chlstResults.CheckedItems(1).ToString()
                R3 = WCaller.chlstResults.CheckedItems(2).ToString()
                R4 = WCaller.chlstResults.CheckedItems(3).ToString()
                R5 = WCaller.chlstResults.CheckedItems(4).ToString()
        End Select
        Dim z As Integer = WCaller.lbZipCity.CheckedItems.Count
        If z <> 0 Then
            If InStr(WCaller.btnZipCity.Text, "Zip Codes") <> 0 Then
                Select Case z
                    Case Is = 1
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 2
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 3
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 4
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 5
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 6
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 7
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 8
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 9
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Is = 10
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"
                    Case Else
                        Z1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        Z3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        Z4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        Z5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        Z6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        Z7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        Z8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        Z9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        Z10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        C1 = "%"
                        C2 = "%"
                        C3 = "%"
                        C4 = "%"
                        C5 = "%"
                        C6 = "%"
                        C7 = "%"
                        C8 = "%"
                        C9 = "%"
                        C10 = "%"

                        Dim cnt As Integer = WCaller.lbZipCity.CheckedItems.Count
                        WCaller.cntZip = cnt
                        Dim o As Integer = 10
                        For o = 11 To cnt
                            WCaller.lbZipCity.SetItemChecked(o - 1, False)
                        Next


                End Select
            Else
                Select Case z
                    Case Is = 1
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 2
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 3
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 4
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 5
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 6
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 7
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 8
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 9
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(0).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Is = 10
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                    Case Else
                        C1 = WCaller.lbZipCity.CheckedItems(0).ToString
                        C2 = WCaller.lbZipCity.CheckedItems(1).ToString
                        C3 = WCaller.lbZipCity.CheckedItems(2).ToString
                        C4 = WCaller.lbZipCity.CheckedItems(3).ToString
                        C5 = WCaller.lbZipCity.CheckedItems(4).ToString
                        C6 = WCaller.lbZipCity.CheckedItems(5).ToString
                        C7 = WCaller.lbZipCity.CheckedItems(6).ToString
                        C8 = WCaller.lbZipCity.CheckedItems(7).ToString
                        C9 = WCaller.lbZipCity.CheckedItems(8).ToString
                        C10 = WCaller.lbZipCity.CheckedItems(9).ToString
                        Z1 = "%"
                        Z2 = "%"
                        Z3 = "%"
                        Z4 = "%"
                        Z5 = "%"
                        Z6 = "%"
                        Z7 = "%"
                        Z8 = "%"
                        Z9 = "%"
                        Z10 = "%"
                        Dim cnt As Integer = WCaller.lbZipCity.CheckedItems.Count
                        WCaller.cntCity = cnt
                        Dim o As Integer = 10
                        For o = 11 To cnt
                            WCaller.lbZipCity.SetItemChecked(o - 1, False)
                        Next

                End Select
            End If
        Else
            Z1 = "%"
            Z2 = "%"
            Z3 = "%"
            Z4 = "%"
            Z5 = "%"
            Z6 = "%"
            Z7 = "%"
            Z8 = "%"
            Z9 = "%"
            Z10 = "%"
            C1 = "%"
            C2 = "%"
            C3 = "%"
            C4 = "%"
            C5 = "%"
            C6 = "%"
            C7 = "%"
            C8 = "%"
            C9 = "%"
            C10 = "%"
        End If
        Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
        Dim param2 As SqlParameter = New SqlParameter("@SLS", SLS)
        Dim param3 As SqlParameter = New SqlParameter("@Date1", D1)
        Dim param4 As SqlParameter = New SqlParameter("@Date2", D2)
        Dim param5 As SqlParameter = New SqlParameter("@Time1", T1)
        Dim param6 As SqlParameter = New SqlParameter("@Time2", T2)
        Dim param7 As SqlParameter = New SqlParameter("@R1", R1)
        Dim param8 As SqlParameter = New SqlParameter("@R2", R2)
        Dim param9 As SqlParameter = New SqlParameter("@R3", R3)
        Dim param10 As SqlParameter = New SqlParameter("@R4", R4)
        Dim param11 As SqlParameter = New SqlParameter("@R5", R5)
        Dim param12 As SqlParameter = New SqlParameter("@Sat", Sat)
        Dim param13 As SqlParameter = New SqlParameter("@Sun", Sun)
        Dim param14 As SqlParameter = New SqlParameter("@Z1", Z1)
        Dim param15 As SqlParameter = New SqlParameter("@Z2", Z2)
        Dim param16 As SqlParameter = New SqlParameter("@Z3", Z3)
        Dim param17 As SqlParameter = New SqlParameter("@Z4", Z4)
        Dim param18 As SqlParameter = New SqlParameter("@Z5", Z5)
        Dim param19 As SqlParameter = New SqlParameter("@Z6", Z6)
        Dim param20 As SqlParameter = New SqlParameter("@Z7", Z7)
        Dim param21 As SqlParameter = New SqlParameter("@Z8", Z8)
        Dim param22 As SqlParameter = New SqlParameter("@Z9", Z9)
        Dim param23 As SqlParameter = New SqlParameter("@Z10", Z10)
        Dim param24 As SqlParameter = New SqlParameter("@C1", C1)
        Dim param25 As SqlParameter = New SqlParameter("@C2", C2)
        Dim param26 As SqlParameter = New SqlParameter("@C3", C3)
        Dim param27 As SqlParameter = New SqlParameter("@C4", C4)
        Dim param28 As SqlParameter = New SqlParameter("@C5", C5)
        Dim param29 As SqlParameter = New SqlParameter("@C6", C6)
        Dim param30 As SqlParameter = New SqlParameter("@C7", C7)
        Dim param31 As SqlParameter = New SqlParameter("@C8", C8)
        Dim param32 As SqlParameter = New SqlParameter("@C9", C9)
        Dim param33 As SqlParameter = New SqlParameter("@C10", C10)

        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("dbo.WarmCalling", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.Parameters.Add(param6)
        cmdGet.Parameters.Add(param7)
        cmdGet.Parameters.Add(param8)
        cmdGet.Parameters.Add(param9)
        cmdGet.Parameters.Add(param10)
        cmdGet.Parameters.Add(param11)
        cmdGet.Parameters.Add(param12)
        cmdGet.Parameters.Add(param13)
        cmdGet.Parameters.Add(param14)
        cmdGet.Parameters.Add(param15)
        cmdGet.Parameters.Add(param16)
        cmdGet.Parameters.Add(param17)
        cmdGet.Parameters.Add(param18)
        cmdGet.Parameters.Add(param19)
        cmdGet.Parameters.Add(param20)
        cmdGet.Parameters.Add(param21)
        cmdGet.Parameters.Add(param22)
        cmdGet.Parameters.Add(param23)
        cmdGet.Parameters.Add(param24)
        cmdGet.Parameters.Add(param25)
        cmdGet.Parameters.Add(param26)
        cmdGet.Parameters.Add(param27)
        cmdGet.Parameters.Add(param28)
        cmdGet.Parameters.Add(param29)
        cmdGet.Parameters.Add(param30)
        cmdGet.Parameters.Add(param31)
        cmdGet.Parameters.Add(param32)
        cmdGet.Parameters.Add(param33)


        Try
            cnn.Open()
            r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            Dim cnt As Integer = 0
            While r.Read
                Dim lv As New ListViewItem
                lv.Name = r.Item(0)
                lv.Text = r.Item(0)
                Dim u As String = r.Item(2).ToString
                Dim w = InStr(u, " ")
                u = Microsoft.VisualBasic.Right(u, w + 2)
                Trim(u)
                Dim u2 As String
                Dim u3 As String
                If u.Length = 11 Then
                    u2 = Microsoft.VisualBasic.Left(u, 5)
                    u3 = Microsoft.VisualBasic.Right(u, 3)
                    u = u2 & u3
                Else
                    u2 = Microsoft.VisualBasic.Left(u, 4)
                    u3 = Microsoft.VisualBasic.Right(u, 3)
                    u = u2 & u3
                End If
                lv.SubItems.Add(r.Item(1))
                lv.SubItems.Add(u)
                If r.Item(5) = "" And r.Item(6) = "" Then
                    lv.SubItems.Add(r.Item(4) & ", " & r.Item(3))
                ElseIf r.Item(5) <> "" And r.Item(6) <> "" And r.Item(4) = r.Item(6) Then
                    lv.SubItems.Add(r.Item(4) & ", " & r.Item(3) & " & " & r.Item(5))
                ElseIf r.Item(5) <> "" And r.Item(6) <> "" And r.Item(4) <> r.Item(6) Then
                    lv.SubItems.Add(r.Item(4) & ", " & r.Item(3) & " & " & r.Item(6) & ", " & r.Item(5))
                End If
                lv.SubItems.Add(r.Item(7) & " " & r.Item(8) & ", " & r.Item(9) & " " & r.Item(10))
                If r.Item(12) = "" And r.Item(13) = "" Then
                    lv.SubItems.Add(r.Item(11))
                ElseIf r.Item(12) <> "" And r.Item(13) = "" Then
                    lv.SubItems.Add(r.Item(11) & " - " & r.Item(12))
                ElseIf r.Item(12) = "" And r.Item(13) <> "" Then
                    lv.SubItems.Add(r.Item(11) & " - " & r.Item(13))
                ElseIf r.Item(12) <> "" And r.Item(13) <> "" Then
                    lv.SubItems.Add(r.Item(15) & " - " & r.Item(16) & " - " & r.Item(17))
                End If
                lv.SubItems.Add(r.Item(14))
                If WCaller.cboGroupBy.Text <> "" Then
                    If WCaller.cboGroupBy.Text = "Zip Code" Then
                        lv.Group = WCaller.lvWarmCalling.Groups(r.Item(10))
                    ElseIf WCaller.cboGroupBy.Text = "City, State" Then
                        lv.Group = WCaller.lvWarmCalling.Groups(r.Item(8))
                    ElseIf WCaller.cboGroupBy.Text = "Marketing Result" Then
                        lv.Group = WCaller.lvWarmCalling.Groups(r.Item(14))
                    ElseIf WCaller.cboGroupBy.Text = "Primary Lead Source" Then
                        lv.Group = WCaller.lvWarmCalling.Groups(r.Item(18))
                    ElseIf WCaller.cboGroupBy.Text = "Primary Product" Then
                        lv.Group = WCaller.lvWarmCalling.Groups(r.Item(11))
                    End If
                End If
                WCaller.lvWarmCalling.Items.Add(lv)
                If r.Item(0) = LastID Then
                    lv.Selected = True
                End If
                cnt = cnt + 1

            End While
            r.Close()
            cnn.Close()
            WCaller.txtRecordsMatching.Text = CStr(cnt)
            WCaller.LastD1 = WCaller.txtDate1.Text
            WCaller.LastD2 = WCaller.txtDate2.Text
            If WCaller.lvWarmCalling.SelectedItems.Count = 0 And WCaller.lvWarmCalling.Items.Count > 0 Then
                WCaller.lvWarmCalling.TopItem.Selected = True
                LastID = WCaller.lvWarmCalling.SelectedItems(0).Text
            End If
            If WCaller.lvWarmCalling.Items.Count = 0 Then
                Me.PullCustomerINFO("")
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub ManagerCriteria()

        Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("dbo.WarmCallingCriteria", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        Try
            r.Read()





            If r.Item(1) <> "" And r.Item(2) <> "" Then ''bug

                WCaller.txtTime1.Visible = False
                WCaller.txtTime1.Text = r.Item(1)
                WCaller.dtpTime1.Value = CDate(r.Item(1))
                WCaller.txtTime2.Visible = False
                WCaller.txtTime2.Text = r.Item(2)
                WCaller.dptTime2.Value = CDate(r.Item(2))
                If r.Item(18) = True Then
                    WCaller.GroupBox4.Enabled = False
                    WCaller.ErrorProvider1.SetError(WCaller.dptTime2, "This field has been locked by a Marketing Manager")
                    WCaller.ErrorProvider1.SetError(WCaller.dtpTime1, "This field has been locked by a Marketing Manager")

                End If
            End If
            WCaller.cboDateRange.Text = r.Item(19)
            If r.Item(12) <> "" And r.Item(13) <> "" Then

                WCaller.txtDate1.Visible = False
                WCaller.txtDate2.Visible = False
                WCaller.dtp1.Value = CDate(r.Item(12))
                WCaller.dtp2.Value = CDate(r.Item(13))
                WCaller.txtDate1.Text = r.Item(12)
                WCaller.txtDate2.Text = r.Item(13)
                WCaller.cboDateRange.Text = r.Item(19)
                If r.Item(18) = True Then
                    WCaller.cboDateRange.Enabled = False
                    WCaller.dtp1.Enabled = False
                    WCaller.ErrorProvider1.SetError(WCaller.dtp1, "This field has been locked by a Marketing Manager")
                    WCaller.dtp2.Enabled = False
                    WCaller.ErrorProvider1.SetError(WCaller.dtp2, "This field has been locked by a Marketing Manager")
                End If
            End If

            WCaller.cboPLSWarmCalling.Text = r.Item(9)

            If r.Item(9) <> "" Then
                WCaller.lblPLSWarmCalling.Visible = False
                Me.GetSLS(r.Item(9))
                WCaller.cboSLSWarmCalling.Text = r.Item(10)
                SLS = r.Item(10)
            End If

            If r.Item(10) <> "" Then
                WCaller.lblSLSWarmCalling.Visible = False
            End If
            If r.Item(9) <> "" And r.Item(18) = True Then
                WCaller.cboPLSWarmCalling.Enabled = False
                WCaller.lblPLSWarmCalling.Visible = False
                WCaller.ErrorProvider1.SetError(WCaller.cboPLSWarmCalling, "This field has been locked by a Marketing Manager")
            End If
            If r.Item(10) <> "" And r.Item(18) = True Then
                WCaller.cboSLSWarmCalling.Enabled = False
                WCaller.lblSLSWarmCalling.Visible = False
                WCaller.ErrorProvider1.SetError(WCaller.cboSLSWarmCalling, "This field has been locked by a Marketing Manager")
            End If
            If r.Item(16) = True Then
                WCaller.rdoZip.Checked = True
                WCaller.rdoCity.Checked = False
            Else
                WCaller.rdoZip.Checked = False
                WCaller.rdoCity.Checked = True
            End If
            If r.Item(15) <> "" Then
                WCaller.PBar = True

                Progress.Show()
                My.Application.DoEvents()
                WCaller.txtZipCode.Text = r.Item(15)
                WCaller.nupMiles.Value = r.Item(14)
                WCaller.btnZipCity_Click(Nothing, Nothing)
                WCaller.lblCheckAll_Click(Nothing, Nothing)

            End If
            If r.Item(15) <> "" And r.Item(18) = True Then
                WCaller.txtZipCode.Enabled = False
                WCaller.rdoZip.Enabled = False
                WCaller.rdoCity.Enabled = False
                WCaller.btnZipCity.Enabled = False
                WCaller.ErrorProvider1.SetError(WCaller.txtZipCode, "This field has been locked by a Marketing Manager")
                WCaller.ErrorProvider1.SetError(WCaller.nupMiles, "This field has been locked by a Marketing Manager")
                WCaller.ErrorProvider1.SetError(WCaller.btnZipCity, "This field has been locked by a Marketing Manager")
                If WCaller.rdoZip.Checked = True Then
                    WCaller.ErrorProvider1.SetError(WCaller.rdoZip, "This field has been locked by a Marketing Manager")
                Else
                    WCaller.ErrorProvider1.SetError(WCaller.rdoCity, "This field has been locked by a Marketing Manager")
                End If
            End If
            If r.Item(3) = True Then
                WCaller.chlstResults.SetItemChecked(0, True)
            Else
                WCaller.chlstResults.SetItemChecked(0, False)
            End If
            If r.Item(4) = True Then
                WCaller.chlstResults.SetItemChecked(1, True)
            Else
                WCaller.chlstResults.SetItemChecked(1, False)
            End If
            If r.Item(5) = True Then
                WCaller.chlstResults.SetItemChecked(2, True)
            Else
                WCaller.chlstResults.SetItemChecked(2, False)
            End If
            If r.Item(6) = True Then
                WCaller.chlstResults.SetItemChecked(3, True)
            Else
                WCaller.chlstResults.SetItemChecked(3, False)
            End If
            If r.Item(7) = True Then
                WCaller.chlstResults.SetItemChecked(4, True)
            Else
                WCaller.chlstResults.SetItemChecked(4, False)
            End If
            If WCaller.chlstResults.CheckedItems.Count <> 5 And r.Item(18) = True Then
                WCaller.chlstResults.Enabled = False
                WCaller.ErrorProvider1.SetError(WCaller.chlstResults, "This field has been locked by a Marketing Manager")

            End If
            If r.Item(8) = True Then
                WCaller.cbWeekdays.Checked = True
                Sat = "Saturday"
                Sun = "Sunday"
            End If
            If r.Item(8) = True And r.Item(18) = True Then
                WCaller.cbWeekdays.Enabled = False
                WCaller.ErrorProvider1.SetError(WCaller.cbWeekdays, "This field has been locked by a Marketing Manager")
            End If
            If r.Item(11) <> "" Then
                WCaller.cboGroupBy.Text = r.Item(11)
                WCaller.lblGroupBy.Visible = False
                If r.Item(18) = True Then
                    WCaller.cboGroupBy.Enabled = False
                    WCaller.ErrorProvider1.SetError(WCaller.cboGroupBy, "This field has been locked by a Marketing Manager")
                End If
            End If
            r.Close()
            cnn.Close()
        Catch ex As Exception
            r.Close()
            cnn.Close()
        End Try




      
    End Sub
    Public dset_PriLS As Data.DataSet = New Data.DataSet("PLS")
    Public da_PRI As SqlDataAdapter = New SqlDataAdapter
    Public Sub GetPrimaryLeadSource()
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetPLS", Cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Try
            cnn.Open()
            da_PRI.SelectCommand = cmdGet
            da_PRI.Fill(dset_PriLS, "PrimaryLeadSource")
            Dim cnt As Integer = 0
            cnt = dset_PriLS.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    WCaller.cboPLSWarmCalling.Items.Clear()
                    WCaller.cboPLSWarmCalling.Items.Add("")


                    Exit Select
                Case Is >= 1
                    WCaller.cboPLSWarmCalling.Items.Clear()

                    WCaller.cboPLSWarmCalling.Items.Add("")

                    Dim b
                    For b = 0 To dset_PriLS.Tables(0).Rows.Count - 1
                        WCaller.cboPLSWarmCalling.Items.Add(dset_PriLS.Tables(0).Rows(b).Item(1))

                    Next
                    Exit Select
            End Select
            cnn.Close()

        Catch ex As Exception

        End Try
    End Sub
    Public dset_SLS As Data.DataSet = New Data.DataSet("SLS")
    Public da_SLS As SqlDataAdapter = New SqlDataAdapter
    Public Sub GetSLS(ByVal PrimaryLS As String)
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetSLS", Cnn)
        Dim param1 As SqlParameter = New SqlParameter("@PRILS", PrimaryLS)
        cmdGet.Parameters.Add(param1)

        cnn.Open()
        cmdGet.CommandType = CommandType.StoredProcedure
        da_SLS.SelectCommand = cmdGet
        da_SLS.Fill(dset_SLS, "SecondaryLeadSource")
        Dim cnt As Integer = 0
        cnt = dset_SLS.Tables(0).Rows.Count

        Select Case cnt
            Case Is <= 0
                WCaller.cboSLSWarmCalling.Items.Clear()
                WCaller.cboSLSWarmCalling.Items.Add("")


                Exit Select
            Case Is >= 1
                WCaller.cboSLSWarmCalling.Items.Clear()

                WCaller.cboSLSWarmCalling.Items.Add("")

                Dim b As Integer = 0
                For b = 0 To dset_SLS.Tables(0).Rows.Count - 1
                    WCaller.cboSLSWarmCalling.Items.Add(dset_SLS.Tables(0).Rows(b).Item(0))

                Next
                Exit Select
        End Select
        cnn.Close()

    End Sub
    Public Sub PullCustomerINFO(ByVal ID As String)
        If WCaller.ID = ID Then
            Exit Sub
        End If
        WCaller.ID = ID

        If ID = "" Then
            WCaller.txtContact1.Text = ""
            WCaller.txtContact2.Text = ""
            WCaller.txtAddress.Text = ""
            WCaller.txtWorkHours.Text = ""
            WCaller.txtHousePhone.Text = ""
            WCaller.txtaltphone1.Text = ""
            WCaller.txtaltphone2.Text = ""
            WCaller.txtAlt1Type.Text = ""
            WCaller.txtAlt2Type.Text = ""
            WCaller.lnkEmail.Text = ""
            WCaller.txtApptDate.Text = ""
            WCaller.txtApptDay.Text = ""
            WCaller.txtApptTime.Text = ""
            WCaller.txtProducts.Text = ""
            WCaller.txtColor.Text = ""
            WCaller.txtQty.Text = ""
            WCaller.txtYrBuilt.Text = ""
            WCaller.txtYrsOwned.Text = ""
            WCaller.txtHomeValue.Text = ""
            WCaller.rtbSpecialInstructions.Text = ""
            WCaller.pnlCustomerHistory.Controls.Clear()
            Exit Sub
        End If
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetCustomerINFO", Cnn)

        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        Try
            Cnn.Open()


            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read
                WCaller.txtContact1.Text = r1.Item(5) & " " & r1.Item(6)
                WCaller.txtContact2.Text = r1.Item(7) & " " & r1.Item(8)
                WCaller.txtAddress.Text = r1.Item(9) & " " & vbCrLf & r1.Item(10) & ", " & r1.Item(11) & " " & r1.Item(12)
                If r1.Item(7) = "" Then
                    WCaller.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19)
                Else
                    WCaller.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19) & vbNewLine & r1.Item(7) & ": " & r1.Item(20)
                End If
                WCaller.txtHousePhone.Text = r1.Item(13)
                WCaller.txtaltphone1.Text = r1.Item(14)
                WCaller.txtaltphone2.Text = r1.Item(16)
                WCaller.txtAlt1Type.Text = r1.Item(15)
                WCaller.txtAlt2Type.Text = r1.Item(17)
                WCaller.lnkEmail.Text = r1.Item(52)
                Dim d
                d = Split(r1.Item(29), " ", 2)
                Trim(d(0))
                WCaller.txtApptDate.Text = d(0)
                WCaller.txtApptDay.Text = r1.Item(30)
                '' r1.item(31) = 'ApptTime' Field | (DateTime, null) Type on table [ EnterLead ] 
                '' sample data from 68336:> '1900-01-01 15:00:00.000'
                '' im pretty sure what is trying to happen here is just to pull out the 
                '' the time like "5:00 PM" and its getting lost in translation. 



                'Dim t = Split(r1.Item(31), " ", 2)
                'Dim u = t(1)
                'Dim u2 As String
                'Dim u3 As String
                'If u.Length = 11 Then
                '    u2 = Microsoft.VisualBasic.Left(u, 5)
                '    u3 = Microsoft.VisualBasic.Right(u, 3)
                '    u = u2 & u3
                'Else
                '    u2 = Microsoft.VisualBasic.Left(u, 4)
                '    u3 = Microsoft.VisualBasic.Right(u, 3)
                '    u = u2 & u3
                'End If

                Dim _Hour As Object = r1.Item("ApptTime").ToString
                Dim dateTime() = Split(_Hour, " ", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim _date As String = ""
                Dim _time As String = ""
                Dim _AmPM As String = ""
                _date = dateTime(0) '' 1900-01-01
                _time = dateTime(1) '' 12:00:00 
                _AmPM = dateTime(2) '' AM/PM
                Dim splitTime() = Split(_time, ":", -1, Microsoft.VisualBasic.CompareMethod.Text)
                Dim hour As String = splitTime(0) '' 12
                Dim minute As String = splitTime(1) '' 00-59 
                Dim correctedTime As String = ((hour & ":" & minute) & " " & _AmPM)
                WCaller.txtApptTime.Text = correctedTime.ToString
                WCaller.txtProducts.Text = r1.Item(21) & vbCrLf & r1.Item(22) & vbCrLf & r1.Item(23)
                WCaller.txtColor.Text = r1.Item(24)
                WCaller.txtQty.Text = r1.Item(25)
                WCaller.txtYrBuilt.Text = r1.Item(27)
                WCaller.txtYrsOwned.Text = r1.Item(26)
                WCaller.txtHomeValue.Text = r1.Item(28)
                WCaller.rtbSpecialInstructions.Text = r1.Item(32)
            End While
            r1.Close()
            Cnn.Close()

        Catch ex As Exception
            'Cnn.Close()
            'Me.PullCustomerINFO(ID)
            MsgBox("Lost Network Connection! Pull Customer Info" & ex.ToString, MsgBoxStyle.Critical, "Server not Available")
        End Try
        Dim c As New CustomerHistory
        If ID <> "" Then
            c.SetUp(WCaller, ID, WCaller.TScboCustomerHistory)
        End If

    End Sub
    Public Class MyApptsTab
        Public Class DisplayColumn
            Public Sub New()
                If WCaller.cboDisplayColumn.Text = "Contact(s) (Default)" Or WCaller.cboDisplayColumn.Text = "" Then
                    WCaller.lvMyAppts.Columns(1).DisplayIndex = 1
                    WCaller.lvMyAppts.Columns(2).DisplayIndex = 2
                    WCaller.lvMyAppts.Columns(5).DisplayIndex = 5
                    WCaller.cboDisplayColumn.Text = ""
                ElseIf WCaller.cboDisplayColumn.Text = "Appt. Date/Time" Then
                    WCaller.lvMyAppts.Columns(1).DisplayIndex = 2
                    WCaller.lvMyAppts.Columns(2).DisplayIndex = 1
                    WCaller.lvMyAppts.Columns(5).DisplayIndex = 5
                ElseIf WCaller.cboDisplayColumn.Text = "Product(s)" Then
                    WCaller.lvMyAppts.Columns(5).DisplayIndex = 1
                    WCaller.lvMyAppts.Columns(1).DisplayIndex = 2
                    WCaller.lvMyAppts.Columns(2).DisplayIndex = 3
                End If
                WCaller.lvMyAppts.Refresh()
            End Sub
        End Class
        Public Class Populate
            Public Sub New(ByVal Filter As String)
                If WCaller.lvMyAppts.SelectedItems.Count <> 0 Then
                    WCaller.lvMyAppts.Tag = WCaller.lvMyAppts.SelectedItems(0).Tag
                Else
                    WCaller.lvMyAppts.Tag = ""
                End If
                WCaller.lvMyAppts.Groups.Clear()
                WCaller.lvMyAppts.Items.Clear()
                If WCaller.cboGroupSetAppt.Text = "" Then
                    WCaller.lvMyAppts.Groups.Add("grpSet", "My Set Appointments")
                    WCaller.lvMyAppts.Groups.Add("grpMemorized", "My Memorized Appointments")
                Else
                    Me.Groups()
                End If
                Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.SetApptsWarmCall", cnn)
                Dim r1 As SqlDataReader
                Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
                Dim param2 As SqlParameter = New SqlParameter("@Filter", WCaller.cboFilter.Text)
                cmdGet.Parameters.Add(param1)
                cmdGet.Parameters.Add(param2)
                cmdGet.CommandType = CommandType.StoredProcedure
                cnn.Open()
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                While r1.Read
                    Dim lv As New ListViewItem
                    lv.Tag = r1.Item(15)
                    If r1.Item(0) <> "" Then
                        lv.ToolTipText = r1.Item(0)
                        lv.ImageIndex = 8
                    End If
                    If r1.Item(4) = "" Then
                        lv.SubItems.Add(r1.Item(2) & ", " & r1.Item(1))
                    ElseIf r1.Item(2).ToString = r1.Item(4).ToString Then
                        lv.SubItems.Add(r1.Item(2) & ", " & r1.Item(1) & " & " & r1.Item(3))
                    Else
                        lv.SubItems.Add(r1.Item(2) & ", " & r1.Item(1) & " & " & r1.Item(4) & ", " & r1.Item(3))
                    End If
                    Dim u As String = r1.Item(5).ToString
                    Dim w = InStr(u, " ")
                    u = Microsoft.VisualBasic.Right(u, w + 2)
                    Trim(u)
                    Dim u2 As String
                    Dim u3 As String
                    If u.Length = 11 Then
                        u2 = Microsoft.VisualBasic.Left(u, 5)
                        u3 = Microsoft.VisualBasic.Right(u, 3)
                        u = u2 & u3
                    Else
                        u2 = Microsoft.VisualBasic.Left(u, 4)
                        u3 = Microsoft.VisualBasic.Right(u, 3)
                        u = u2 & u3
                    End If
                    Dim ApptTime As String = u
                    Dim ApptDate As Date = r1.Item(6)
                    ApptDate.ToShortDateString()
                    lv.SubItems.Add(ApptDate & " - " & ApptTime)
                    lv.SubItems.Add(r1.Item(7))
                    lv.SubItems.Add(r1.Item(8))
                    If r1.Item(11) = "" And r1.Item(13) = "" Then
                        lv.SubItems.Add(r1.Item(9))
                    ElseIf r1.Item(11) <> "" And r1.Item(13) = "" Then
                        lv.SubItems.Add(r1.Item(9) & "-" & r1.Item(11))
                    ElseIf r1.Item(11) <> "" And r1.Item(13) <> "" Then
                        lv.SubItems.Add(r1.Item(9) & "-" & r1.Item(11) & "-" & r1.Item(13))
                    ElseIf r1.Item(11) = "" And r1.Item(13) <> "" Then
                        lv.SubItems.Add(r1.Item(9) & "-" & r1.Item(13))
                    End If
                    If r1.Item(16) = True Then
                        If WCaller.cboGroupSetAppt.Text = "" Then
                            lv.Group = WCaller.lvMyAppts.Groups("grpSet")
                        ElseIf WCaller.cboGroupSetAppt.Text = "Date" Then
                            lv.Group = WCaller.lvMyAppts.Groups(ApptDate)
                        ElseIf WCaller.cboGroupSetAppt.Text = "Primary Product" Then
                            lv.Group = WCaller.lvMyAppts.Groups(r1.Item(9).ToString)
                        ElseIf WCaller.cboGroupSetAppt.Text = "City, State" Then
                            lv.Group = WCaller.lvMyAppts.Groups(r1.Item(17) & ", " & r1.Item(18))
                        ElseIf WCaller.cboGroupSetAppt.Text = "Zip Code" And r1.Item(19) <> "" Then
                            lv.Group = WCaller.lvMyAppts.Groups(r1.Item(19))
                        ElseIf WCaller.cboGroupSetAppt.Text = "Zip Code" And r1.Item(19) = "" Then
                            lv.Group = WCaller.lvMyAppts.Groups("No Zip Code")
                        End If
                    Else
                        lv.Group = WCaller.lvMyAppts.Groups("grpMemorized")
                    End If
                    WCaller.lvMyAppts.Items.Add(lv)
                    If lv.Tag = WCaller.lvMyAppts.Tag Then
                        lv.Selected = True
                    End If
                End While
                If WCaller.lvMyAppts.SelectedItems.Count = 0 And WCaller.lvMyAppts.Items.Count <> 0 Then
                    WCaller.lvMyAppts.TopItem.Selected = True
                End If
                If WCaller.lvMyAppts.Items.Count = 0 Then
                    Dim c As New WarmCalling
                    c.PullCustomerINFO("")
                End If
                r1.Close()
                cnn.Close()
            End Sub
            Private Sub Groups()
                Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.MyApptsGroups", cnn)
                Dim r1 As SqlDataReader
                Dim param1 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
                Dim param2 As SqlParameter = New SqlParameter("@Group", WCaller.cboGroupSetAppt.Text)
                cmdGet.Parameters.Add(param1)
                cmdGet.Parameters.Add(param2)
                cmdGet.CommandType = CommandType.StoredProcedure
                cnn.Open()
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                While r1.Read
                    Dim x = r1.Item(0)
                    If WCaller.cboGroupSetAppt.Text = "Date" Then
                        x = CDate(x)
                        x = x.toshortdatestring()
                    ElseIf WCaller.cboGroupSetAppt.Text = "Zip Code" And r1.Item(0) = "" Then
                        WCaller.lvMyAppts.Groups.Add("No Zip Code", "No Zip Code")
                    ElseIf WCaller.cboGroupSetAppt.Text = "City, State" Then
                        x = x & ", " & r1.Item(1)
                    End If
                    WCaller.lvMyAppts.Groups.Add(x, x)
                End While
                r1.Close()
                cnn.Close()
                WCaller.lvMyAppts.Groups.Add("grpMemorized", "My Memorized Appointments")
            End Sub
        End Class
    End Class
    Public Class LoadProcedure
        Public Sub New()
            ' WCaller.MdiParent = Main
            ''  Set up default times and dates to dispay on click or focus of control
            WCaller.dtp1.Value = CDate("1/1/" & Today.Year)
            WCaller.dtp2.Value = Today
            WCaller.dtpTime1.Value = CDate("1/1/1900 7:00 AM")
            WCaller.dptTime2.Value = CDate("1/1/1900 7:00 PM")
            '' Checks all Marketing results in the search criteria side bar
            For i As Integer = 0 To WCaller.chlstResults.Items.Count - 1
                WCaller.chlstResults.SetItemChecked(i, True)
            Next
            ''      Populates Primary Lead Source
            Dim c As New WarmCalling
            c.GetPrimaryLeadSource()
            ''      Clears all text boxes overlaying all date and time pickers to give it 
            ''      the appearance of blank values in all these fields indicating that no time or 
            ''      date range is in effect
            WCaller.txtTime1.Text = ""
            WCaller.txtTime2.Text = ""
            WCaller.txtDate2.Text = ""
            WCaller.txtDate1.Text = ""

            ''Sets values for all filters and controls concerning the criteria builder
            ''and locks them to Marketing Manager's Liking to pre determine the users list at logon
            c.ManagerCriteria()
            '           Set Customer History Filter
            ''Locks Customer History cbo filter to Marketing only to show only marketing history
            '' as the rest of dept data is of no concern 
            WCaller.TScboCustomerHistory.SelectedIndex = (1)
            WCaller.TScboCustomerHistory.Enabled = False
            '' Stalls Form Layout to give Mappoint a chance to do its thing
            If WCaller.PBar = True Then
                WCaller.SuspendLayout()
            End If




            '' Decides whether or not to apply groups & then populate list or just 
            ''populate list with no grouping
            If WCaller.cboGroupBy.Text = "" Then
                c.Populate()
                Dim x As New WarmCalling.MyApptsTab.Populate("")
            Else
                c.GroupBy()
            End If


            'WCaller.cboDateRange.SelectedIndex = 0
            'Set Expand Characters
            WCaller.btnExpandWarmCalling.Text = Chr(187)
            WCaller.btnExpandMyAppts.Text = Chr(187)

            '           Set Panels
            WCaller.SplitContainer1.SplitterDistance = 225
            WCaller.SplitContainer1.IsSplitterFixed = True

            '           Select First Listview Item Both Tabs
            If WCaller.lvWarmCalling.Items.Count > 0 Then
                WCaller.lvWarmCalling.TopItem.Selected = True
            End If
            'If WCaller.lvMyAppts.Items.Count > 0 Then
            '    WCaller.lvMyAppts.TopItem.Selected = True
            'End If



            '           Setup Dynamic Buttons
            WCaller.btnUndoSet.Text = "Undo Set Appointment"
            WCaller.btnUndoSet.Image = WCaller.ilToolStripIcons.Images(0)

            WCaller.btnUndoMemorize.Text = "Remove From Memorized List"
            WCaller.btnUndoMemorize.Image = WCaller.ilToolStripIcons.Images(1)

            WCaller.SplitContainer1.Panel2.Controls.Remove(WCaller.tsAutoDial)
            WCaller.Controls.Add(WCaller.tsAutoDial)
            WCaller.tsAutoDial.Location = New System.Drawing.Point(335, 25)

            If WCaller.lvWarmCalling.SelectedItems.Count > 0 Then
                WCaller.btnAutoDialer.DropDownItems.Add(WCaller.separator)
                WCaller.btnAutoDialer.DropDownItems.Add(WCaller.btnMain)
                WCaller.btnMain.Text = "Call Main- " & WCaller.txtHousePhone.Text
                If WCaller.txtaltphone1.Text <> "" Then
                    WCaller.btnAutoDialer.DropDownItems.Add(WCaller.btnAlt1)
                    WCaller.btnAlt1.Text = "Call Alt 1- " & WCaller.txtaltphone1.Text
                End If
                If WCaller.txtaltphone2.Text <> "" Then
                    WCaller.btnAutoDialer.DropDownItems.Add(WCaller.btnAlt2)
                    WCaller.btnAlt2.Text = "Call Alt 2- " & WCaller.txtaltphone2.Text
                End If
            End If
            '' Removes Auto Dial Toolbar and readds it directly under form to make it topmost
            WCaller.gbContactInfo.Controls.Remove(WCaller.tsAutoDial)
            WCaller.Controls.Add(WCaller.tsAutoDial)
            WCaller.tsAutoDial.BringToFront()

            If WCaller.txtaltphone1.Text <> "" Then
                WCaller.tsbtnHousePhone.DropDownItems.Add(WCaller.tsbtnAlt1)
                WCaller.tsbtnAlt1.Text = "Call Alt 1- " & WCaller.txtaltphone1.Text
            End If
            If WCaller.txtaltphone2.Text <> "" Then
                WCaller.tsbtnHousePhone.DropDownItems.Add(WCaller.tsbtnAlt2)
                WCaller.tsbtnAlt2.Text = "Call Alt 2- " & WCaller.txtaltphone2.Text
            End If



            Dim w = WCaller.Size.Width
            WCaller.pnlSearch.Controls.Remove(WCaller.btnExpandCollapse)
            WCaller.Controls.Add(WCaller.btnExpandCollapse)
            WCaller.btnExpandCollapse.Location = New System.Drawing.Point(w - 30, 31)
            'WCaller.btnExpandCollapse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right) Or AnchorStyles.Bottom, System.Windows.Forms.AnchorStyles)
            WCaller.SplitContainer1.Panel2.Controls.Remove(WCaller.pnlSearch)
            WCaller.Controls.Add(WCaller.pnlSearch)

            WCaller.pnlSearch.Location = New System.Drawing.Point(w - 344, 28)
            WCaller.pnlSearch.BackColor = Color.White
            WCaller.pnlSearch.Dock = DockStyle.Right
            WCaller.btnExpandCollapse.BringToFront()
            WCaller.pnlSearch.BringToFront()
            WCaller.btnExpandCollapse.Text = Chr(171)
            WCaller.LoadComplete = True  '' this variable was a solution to stop datetimepicker 
            ''                              value changed events from firing on load so the DTP's 
            ''                              could recieve thier default values  
            ''  
            WCaller.ResumeLayout()

            '' This negotiates which message box to use if mappoint returns more than 10 variables of city or zipcode
            If WCaller.msgzip = True Then
                MsgBox("Map Point returned " & WCaller.cntZip & " Zip Codes" & vbCr & _
                             "Zip Code Radius Search only supports up to 10 Zip Codes" & vbCr & _
                             "The first set of 10 are checked, to use unchecked Zipcodes," & vbCr & _
                             "uncheck Zip Codes from the first set & check latter Zip Codes," & vbCr & _
                             "then click the ""Search Button"" to refresh your list", MsgBoxStyle.Exclamation, "Radius Search - Zip Code")
            End If
            If WCaller.msgcity = True Then
                MsgBox("Map Point returned " & WCaller.cntCity & " Cities" & vbCr & _
                       "City Radius Search only supports up to 10 Cities" & vbCr & _
                       "The first set of 10 are checked, to use unchecked Cities," & vbCr & _
                       "uncheck Cities from the first set & check latter Cities," & vbCr & _
                       "then click the ""Search Button"" to refresh your list", MsgBoxStyle.Exclamation, "Radius Search - City")
            End If
        End Sub
    End Class
   
End Class







