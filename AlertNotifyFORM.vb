
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System

Public Class AlertNotify
    Public ID As String
    Public i As ListViewItem
#Region "Connection String"

    Private cnn As SqlConnection = New System.Data.SqlClient.SqlConnection(STATIC_VARIABLES.Cnn)
#End Region
    Private iter As Integer = 0

    Private Sub AlertNotify_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        'Dim cmdDEL As SqlCommand = New SqlCommand("DELETE Iss.dbo.alerttable where ID = @ID", cnn)
        'Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        'cmdDEL.Parameters.Add(param1)
        'cnn.Open()
        'Dim r1 As SqlDataReader
        'r1 = cmdDEL.ExecuteReader
        'r1.Close()
        'cnn.Close()
        'Me.Close()
        'Me.Dispose()
    End Sub
 
    Private Sub AlertNotify_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT ID, Contact1FirstName,Contact1LastName,Contact2FirstName,Contact2LastName,Staddress,city,state,zip,housephone,altphone1,altphone2,phone1type,phone2type,(select notes from AlertTable where id = @id ) from iss.dbo.enterlead where ID = (select leadnum from alerttable where @ID = id)", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGET.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        While r1.Read
            Me.lblCustomerNFO.Text = r1.Item("Contact1FirstName") & " " & r1.Item("Contact1LastName") & " and " & r1.Item("Contact2FirstName") & " " & r1.Item("Contact2LastName") & vbCrLf & r1.Item("Staddress") & vbCrLf & r1.Item("City") & ", " & r1.Item("State") & ", " & r1.Item("Zip") & _
                 vbCrLf & r1.Item("HousePhone") & vbCrLf & r1.Item("AltPhone1") & "     " & r1.Item("Phone1Type") & vbCrLf & r1.Item("AltPhone2") & "     " & r1.Item("Phone2Type")
            Me.lnkID.Text = r1.Item("ID")
            Me.lblnotes.Text = r1.Item(14)
        End While


        r1.Close()
        cnn.Close()



        Me.MdiParent = Main
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.lblTime.Text = Format(DateTime.Now(), "h:mm tt").ToString
        Me.Text = "Alert Notification" & " for Record ID: " _
                & Me.lnkID.Text.ToString
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Timer1.Start()
        My.Computer.Audio.Play("C:\Users\Public\ISS Logs\samsung_whistle.wav", _
                   AudioPlayMode.Background)
        Main.Refresh()
    End Sub



    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclosewindow.Click

        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdDEL As SqlCommand = New SqlCommand("DELETE Iss.dbo.alerttable where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdDEL.Parameters.Add(param1)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdDEL.ExecuteReader
        r1.Close()
        cnn.Close()
        Me.Close()
        Me.Dispose()

    End Sub

    Private Sub btnRemind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemind.Click
        'need to push timer by 5 minutes
        Dim time = Format(DateTime.Now.AddMinutes(5), "HH:mm:00")
        Dim dt = Date.Today()
        time = "01/01/1900 " & time
        dt = dt & " 00:00:00"


        'MsgBox(dt & "   " & time)
        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE iss.dbo.alerttable SET ExecutionDate = @Date, AlertTime = @Time, completed = 0 WHERE ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Date", dt)
        Dim param2 As SqlParameter = New SqlParameter("@Time", time)
        Dim param3 As SqlParameter = New SqlParameter("@ID", ID)
        cmdUP.Parameters.Add(param1)
        cmdUP.Parameters.Add(param2)
        cmdUP.Parameters.Add(param3)
        cnn.Open()
        Dim r1 As SqlDataReader
        r1 = cmdUP.ExecuteReader
        r1.Close()
        cnn.Close()
        Me.Close()
        Me.Dispose()

    End Sub

    Private Sub lnkID_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkID.LinkClicked


        Dim found As Integer = 0
        Dim open As Integer = 0
        For Each frm As Form In Main.MdiChildren
            Select Case frm.Name
                Case "Confirming", "Sales", "Administration", "ColdCalling", "Warmer", "Finance", "MarketingManager", _
                        "Recovery", "PreviousCustomer", "Installation"
                    open += 1
                    Select Case frm.Name
                        Case Is = "Confirming"
                            'Dim lv As New ListViewItem
                            Dim tab As String = Confirming.Tab
                            'If tab = "Confirm" Then
                            Dim foundConfirming As Integer = 0
                            For Each Me.i In Confirming.lvConfirming.Items
                                If Me.i.Text = Me.lnkID.Text Then

                                    Me.i.Selected = True
                                    found += 1
                                    foundConfirming += 1
                                    If Confirming.TabControl1.SelectedIndex <> 0 Then
                                        Confirming.TabControl1.SelectedIndex = 0
                                    End If
                                    'Confirming.Focus()
                                    Me.Label5_Click(Nothing, Nothing)

                                    'Exit Sub
                                End If
                            Next
                            'End If
                            If Confirming.lvConfirming.SelectedItems.Count <> 0 Then
                                If Confirming.lvConfirming.SelectedItems(0).Text <> Me.lnkID.Text Then
                                    For Each Me.i In Confirming.lvSales.Items
                                        If Me.i.Text = Me.lnkID.Text Then
                                            Me.i.Selected = True
                                            found += 1
                                            foundConfirming += 1
                                            If Confirming.TabControl1.SelectedIndex <> 1 Then
                                                Confirming.TabControl1.SelectedIndex = 1
                                            End If
                                            'Confirming.Focus()
                                            Me.Label5_Click(Nothing, Nothing)

                                            'Exit Sub
                                        End If
                                    Next
                                End If
                            End If

                            If foundConfirming = 0 Then
                                Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                                Dim cmdGET As SqlCommand = New SqlCommand("dbo.AlertConfirm", cnn)

                                Dim param1 As SqlParameter = New SqlParameter("@ID", Me.lnkID.Text)
                                cmdGET.Parameters.Add(param1)
                                cmdGET.CommandType = CommandType.StoredProcedure
                                cnn.Open()
                                Dim r1 As SqlDataReader

                                r1 = cmdGET.ExecuteReader
                                While r1.Read
                                    'MsgBox(r1.Item(0))
                                    Dim was As Date = Confirming.dpConfirming.Value





                                    Confirming.dpConfirming.Value = r1.Item(0)
                                    'Exit Sub '' remove 
                                    'MsgBox(Confirming.lvConfirming.Items.Count)
                                    For Each Me.i In Confirming.lvConfirming.Items
                                        'MsgBox(Confirming.lvConfirming.SelectedItems(0).Text)
                                        If Me.i.Text = Me.lnkID.Text Then

                                            Me.i.Selected = True

                                            'Dim c As New ConfirmingData
                                            'c.PullCustomerINFO(Confirming.Tab, Me.i.Text)

                                            found += 1
                                            foundConfirming += 1
                                            If Confirming.TabControl1.SelectedIndex <> 0 Then
                                                Confirming.TabControl1.SelectedIndex = 0
                                            End If
                                            'Confirming.Focus()
                                            'Me.Label5_Click(Nothing, Nothing)
                                            Confirming.Create_Tooltip(Confirming.dpConfirming, "alert")


                                            'Exit Sub
                                        End If

                                    Next
                                    If Confirming.lvConfirming.SelectedItems.Count = 0 Then
                                        Confirming.dpConfirming.Value = was
                                    End If
                                    was = Confirming.dpSales.Value

                                    If foundConfirming = 0 Then

                                        Confirming.dpSales.Value = r1.Item(0)
                                        For Each Me.i In Confirming.lvSales.Items
                                            If Me.i.Text = Me.lnkID.Text Then
                                                Me.i.Selected = True
                                                found += 1
                                                foundConfirming += 1
                                                If Confirming.TabControl1.SelectedIndex <> 1 Then
                                                    Confirming.TabControl1.SelectedIndex = 1
                                                End If
                                                'Confirming.Focus()
                                                Confirming.Create_Tooltip(Confirming.dpSales, "alert")
                                                'Me.Label5_Click(Nothing, Nothing)
                                            Else

                                                'Exit Sub
                                            End If
                                        Next
                                        If foundConfirming = 0 Then
                                            Confirming.dpSales.Value = was
                                        End If
                                    End If
                                End While
                                r1.Close()
                                cnn.Close()
                            End If


                            '' run query to see if it is open to form (this case being confirming) if confirming form = true get date, 
                            ''switch date picker then rerun for each and get lead , add tool tip explaining date has been changed to find the lead 
                            '' change below message box to say lead not available for confirming form 

                        Case Is = "Sales"
                            Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                            Dim cmdGET As SqlCommand = New SqlCommand("dbo.FindSales", cnn)

                            Dim param1 As SqlParameter = New SqlParameter("@ID", Me.lnkID.Text)
                            cmdGET.Parameters.Add(param1)
                            cmdGET.CommandType = CommandType.StoredProcedure
                            cnn.Open()
                            Dim r1 As SqlDataReader

                            r1 = cmdGET.ExecuteReader
                            Try
                                While r1.Read
                                    If r1.Item(0) = True Then
                                        If Sales.tbMain.SelectedIndex <> 1 Then
                                            Sales.tbMain.SelectedIndex = 1
                                        End If
                                        Dim dt As String = r1.Item(1)
                                        Dim s = Split(dt, " ")
                                        dt = s(0)




                                        If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) <= 0 And DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) >= 0 Then
                                            If Sales.cboSalesList.SelectedText <> "Unfiltered Sales Dept. List" Then
                                                Sales.cboSalesList.SelectedText = "Unfiltered Sales Dept. List"
                                                For Each Me.i In Sales.lvSales.Items
                                                    If Me.i.Text = Me.lnkID.Text Then
                                                        Me.i.Selected = True
                                                        found += 1
                                                    End If
                                                Next
                                            End If
                                        Else
                                            If CType(dt, Date) = Today Then
                                                Sales.cboDateRangeCustomerList.SelectedItem = "Today"
                                            End If
                                            If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) > 0 Then
                                                Sales.dtp1CustomerList.Value = dt
                                                Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                            End If
                                            If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) < 0 Then
                                                Sales.dtp2CustomerList.Value = dt
                                                Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                            End If
                                            If Sales.cboSalesList.SelectedItem <> "Unfiltered Sales Dept. List" Then
                                                Sales.cboSalesList.SelectedItem = "Unfiltered Sales Dept. List"
                                            End If
                                            For Each Me.i In Sales.lvSales.Items
                                                If Me.i.Text = Me.lnkID.Text Then
                                                    Me.i.Selected = True
                                                    found += 1
                                                End If
                                            Next
                                        End If
                                        If Sales.tbMain.SelectedIndex <> 1 Then
                                            Sales.tbMain.SelectedIndex = 1
                                        End If
                                    Else
                                        Dim dt As String = r1.Item(2)
                                        Dim s = Split(dt, " ")
                                        dt = s(0)
                                        If r1.Item(1) = True Then
                                            If Sales.cboSalesList.SelectedItem <> "Unconfirmed Appts. For Today" Then
                                                Sales.cboSalesList.SelectedItem = "Unconfirmed Appts. For Today"
                                            End If

                                            'MsgBox(DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value).ToString)
                                            If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp1CustomerList.Value) > 0 Then
                                                Sales.dtp1CustomerList.Value = dt
                                                Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                            End If
                                            If DateDiff(DateInterval.Day, CType(dt, Date), Sales.dtp2CustomerList.Value) < 0 Then
                                                Sales.dtp2CustomerList.Value = dt
                                                Sales.cboDateRangeCustomerList.SelectedItem = "Custom"
                                            End If
                                            For Each Me.i In Sales.lvSales.Items
                                                If Me.i.Text = Me.lnkID.Text Then
                                                    Me.i.Selected = True
                                                    found += 1
                                                End If
                                            Next
                                        End If
                                        If Sales.tbMain.SelectedIndex <> 1 Then
                                            Sales.tbMain.SelectedIndex = 1
                                        End If
                                    End If
                                End While
                            Catch ex As Exception

                            End Try



                        Case Is = "Installation"


                        Case Is = "Administration"


                        Case Is = "WCaller"

                        Case Is = "MarketingManager"

                        Case Is = "Recovery"

                        Case Is = "PreviousCustomer"

                        Case Is = "Finance"

                        Case Is = "ColdCalling"



                    End Select
            End Select





            

        Next
        If found = 0 And open <> 0 Then
            MsgBox("Cannot find record anywhere in your" & vbCr & "open forms with customer lists!", MsgBoxStyle.Critical, "Alert Link Failed")
        ElseIf open = 0 Then

            MsgBox("You have no lists open to link to!", MsgBoxStyle.Critical, "Alert Link Failed")
        End If



        'End Select






        'Confirming.Focus()
        'focus alert id in listview
        'add customer to listview and stop confirm refresh timer
        'remove customer in listview on lost index changed and restart refresh timer 
        Me.Label5_Click(Nothing, Nothing)
     

    End Sub

  


    Private Sub AlertNotify_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
       
        'Me.Opacity = 0
        Me.WindowState = FormWindowState.Normal

    End Sub
    Dim cnt As Integer = 0
    Dim cnt2 As Integer = 0
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.cnt += 1
        'If Me.pctOn1.Location.X <> 220 Then
        '    Me.pctOn1.Location = New System.Drawing.Point((Me.pctOn1.Location.X + 22), 11)
        '    Me.pctOn1.BringToFront()
        'Else
        '    Me.pctOn1.Location = New System.Drawing.Point(66, 11)
        'End If
        Select Case cnt
            Case 1
                Me.pctOn1.Visible = True
                Me.pctOn1.BringToFront()
            Case 2
                Me.pctOn2.Visible = True
                Me.pctOn2.BringToFront()
            Case 3
                Me.pctOn3.Visible = True
                Me.pctOn3.BringToFront()
            Case 4
                Me.pctOn4.Visible = True
                Me.pctOn4.BringToFront()
            Case 5
                Me.pctOn5.Visible = True
                Me.pctOn5.BringToFront()
            Case 6
                Me.pctOn6.Visible = True
                Me.pctOn6.BringToFront()
            Case 7
                Me.pctOn7.Visible = True
                Me.pctOn7.BringToFront()
            Case 8
                Me.pctOn8.Visible = True
                Me.pctOn8.BringToFront()
            Case 9
                Me.pctOn1.Visible = False
                Me.pctOn2.Visible = False
                Me.pctOn3.Visible = False
                Me.pctOn4.Visible = False
                Me.pctOn5.Visible = False
                Me.pctOn6.Visible = False
                Me.pctOn7.Visible = False
                Me.pctOn8.Visible = False
                Me.cnt = 0
        End Select


        If cnt2 = 150 Then
            My.Computer.Audio.Play("C:\Users\Public\ISS Logs\samsung_whistle.wav", _
                AudioPlayMode.Background)
            cnt2 = 0
        End If
        cnt2 += 1
    End Sub
End Class
