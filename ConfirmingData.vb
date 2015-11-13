
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class ConfirmingData
    Public cnnS As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public cnnC As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)

    Public Sub Populate(ByVal Tab As String, ByVal PLS As String, ByVal SLS As String, ByVal ApptDate As String, ByVal PoporRefresh As String)


        If PoporRefresh = "Refresh" Then
            If Tab = "Confirm" Then
                If Confirming.lvConfirming.SelectedItems.Count > 0 Then
                    Dim i = Confirming.lvConfirming.SelectedItems(0).Text
                    Confirming.lvConfirming.Tag = i
                End If

            ElseIf Tab = "Dispatch" Then
                If Confirming.lvSales.SelectedItems.Count > 0 Then
                    Dim i = Confirming.lvSales.SelectedItems(0).Text
                    Confirming.lvSales.Tag = i
                End If
            End If
        End If
        If Tab = "Confirm" Then
            Confirming.lvConfirming.Items.Clear()
        Else
            Confirming.lvSales.Items.Clear()
            Confirming.lvSales.Groups.Clear()
        End If

        If PLS = "" Then
            PLS = "%"
        End If
        If SLS = "" Then
            SLS = "%"
        End If

        ApptDate.ToString()
        Dim x = InStr(ApptDate, " ")
        If x <> 0 Then
            ApptDate = Microsoft.VisualBasic.Left(ApptDate, x - 1)
        End If

        ApptDate = ApptDate & " 12:00:00 AM"
        Dim cmdGet As SqlCommand

        If Tab = "Confirm" Then
            cmdGet = New SqlCommand("dbo.Confirming", cnnC)

        ElseIf Tab = "Dispatch" Then
            cmdGet = New SqlCommand("dbo.ConfirmingDispatch", cnnS)

        End If
        Dim r1 As SqlDataReader
        Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
        Dim param2 As SqlParameter = New SqlParameter("@SLS", SLS)
        Dim param3 As SqlParameter = New SqlParameter("@ApptDate", ApptDate)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        Try

            If Tab = "Confirm" Then

                cnnC.Open()
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                While R1.Read
                    Dim lv As New ListViewItem
                    lv.Text = R1.Item(0)
                    Dim u As String = r1.Item(12).ToString
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
                    lv.SubItems.Add(u)
                    lv.SubItems.Add(R1.Item(1) & ", " & R1.Item(2))
                    If r1.Item(3) <> "" Then
                        lv.SubItems.Add(r1.Item(3) & ", " & r1.Item(4))
                    Else
                        lv.SubItems.Add("")
                    End If
                    lv.SubItems.Add(r1.Item(5) & " " & r1.Item(6) & ", " & r1.Item(7) & " " & r1.Item(8))
                    If r1.Item(10) = "" And r1.Item(11) = "" Then
                        lv.SubItems.Add(r1.Item(9))
                    ElseIf r1.Item(10) <> "" And r1.Item(11) = "" Then
                        lv.SubItems.Add(r1.Item(9) & "-" & r1.Item(10))
                    ElseIf r1.Item(10) <> "" And r1.Item(11) <> "" Then
                        lv.SubItems.Add(r1.Item(9) & "-" & r1.Item(10) & "-" & r1.Item(10))
                    End If
                    lv.SubItems.Add(r1.Item(13))
                    If r1.Item(14) = "Unconfirmed" Or r1.Item(14) = "Reset" Or r1.Item(14) = "Not Hit" Or r1.Item(14) = "Not Issued" Or r1.Item(14) = "Set Appointment" Then
                        lv.Group = Confirming.lvConfirming.Groups(0)
                        lv.Tag = "Unconfirmed"
                    ElseIf r1.Item(14) = "Confirmed" Then
                        lv.Group = Confirming.lvConfirming.Groups(1)
                        lv.Tag = "Confirmed"
                    ElseIf r1.Item(14) = "Called and Cancelled" Then
                        lv.Group = Confirming.lvConfirming.Groups(2)
                        lv.Tag = "Called and Cancelled"
                    End If
                    Confirming.lvConfirming.Items.Add(lv)



                    If lv.Text = Confirming.lvConfirming.Tag Then
                        lv.Selected = True
                        STATIC_VARIABLES.CurrentID = Confirming.lvConfirming.Tag

                    End If

                End While
                r1.Close()
                cnnC.Close()

                Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
                Dim param4 As SqlParameter = New SqlParameter("@PLS", PLS)
                Dim param5 As SqlParameter = New SqlParameter("@SLS", SLS)
                Dim param6 As SqlParameter = New SqlParameter("@ApptDate", ApptDate)
                cnn.Open()
                cmdGet = New SqlCommand("dbo.ConfirmingRefreshCheck", cnn)

                cmdGet.CommandType = CommandType.StoredProcedure
                cmdGet.Parameters.Add(param4)
                cmdGet.Parameters.Add(param5)
                cmdGet.Parameters.Add(param6)

                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                r1.Read()
                'MsgBox(r1.Item(0).ToString & " " & r1.Item(1).ToString & " " & r1.Item(2).ToString & " " & r1.Item(3).ToString)
                If r1.Item(0) <> Confirming.cntTotal Or r1.Item(1) <> Confirming.cntConfirmed Or r1.Item(2) <> Confirming.cntUnconfirmed Or r1.Item(3) <> Confirming.cntCandC Then
                    Confirming.cntTotal = r1.Item(0)
                    Confirming.cntConfirmed = r1.Item(1)
                    Confirming.cntUnconfirmed = r1.Item(2)
                    Confirming.cntCandC = r1.Item(3)

                End If
                r1.Close()
                cnn.Close()
                If Confirming.lvConfirming.Items.Count > 0 And Confirming.lvConfirming.SelectedItems.Count = 0 Then
                    Confirming.lvConfirming.TopItem.Selected = True
                End If

            ElseIf Tab = "Dispatch" Then

                cnnS.Open()
                r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
                Dim g As ListViewGroup = New ListViewGroup("C&C", "Called & Cancelled")

                Confirming.lvSales.Groups.Add(New ListViewGroup("NotIssued", "Appointments Not Issued"))
                Dim cnt As Integer = 0
                Dim FirstTime As Boolean = True
                While r1.Read
                    Dim lv As New ListViewItem '
                    lv.Text = r1.Item(0) '
                    If r1.Item(12) = "Called and Cancelled" Then
                        lv.Tag = "C&C"
                    End If
                    Dim u = r1.Item(1).ToString
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
                    lv.SubItems.Add(u)

                    lv.SubItems.Add(r1.Item(17) & ", " & r1.Item(18))



                    Try
                        lv.SubItems.Add(r1.Item(3))
                    Catch ex As Exception
                        lv.SubItems.Add("")
                    End Try
                    lv.SubItems.Add(r1.Item(4) & " " & r1.Item(5) & ", " & r1.Item(6) & " " & r1.Item(7))
                    Dim p2
                    Dim p3
                    Try
                        p2 = r1.Item(9)
                    Catch ex As Exception
                        p2 = ""
                    End Try
                    Try
                        p3 = r1.Item(10)
                    Catch ex As Exception
                        p3 = ""
                    End Try
                    If p2 = "" And p3 = "" Then
                        lv.SubItems.Add(r1.Item(8))
                    ElseIf p2 <> "" And p3 = "" Then
                        lv.SubItems.Add(r1.Item(8) & "-" & r1.Item(9))
                    ElseIf p2 <> "" And p3 <> "" Then
                        lv.SubItems.Add(r1.Item(8) & "-" & r1.Item(9) & "-" & r1.Item(10))
                    End If
                    Try
                        lv.SubItems.Add(r1.Item(11))
                    Catch ex As Exception
                        lv.SubItems.Add("")
                    End Try
                    Dim rep As String
                    Try
                        rep = r1.Item(2)
                        If lv.SubItems(3).Text <> "" Then
                            lv.SubItems(3).Text = rep & " & " & lv.SubItems(3).Text
                        Else
                            lv.SubItems(3).Text = rep
                        End If

                    Catch ex As Exception
                        rep = ""

                    End Try
                    If rep <> "" And lv.Tag <> "C&C" Then

                        Dim CreatedGroup As Boolean
                        Dim NextRep As String

                        If NextRep <> rep Then
                            CreatedGroup = False
                        End If

                        If CreatedGroup = False Then
                            Dim len = rep.Length
                            If rep.Chars(len - 1) = "s" Then

                                Confirming.lvSales.Groups.Add(New ListViewGroup(rep, rep & "' Appointments"))
                                lv.Group = Confirming.lvSales.Groups(rep)
                            Else

                                Confirming.lvSales.Groups.Add(New ListViewGroup(rep, rep & "'s Appointments"))
                                lv.Group = Confirming.lvSales.Groups(rep)
                            End If
                            CreatedGroup = True
                            NextRep = rep
                        Else
                            lv.Group = Confirming.lvSales.Groups(rep)
                        End If

                    ElseIf rep = "" And lv.Tag <> "C&C" Then
                        lv.Group = Confirming.lvSales.Groups("NotIssued")
                    ElseIf lv.Tag = "C&C" Then
                        'If lv.SubItems(3).Text <> "" Then
                        '    lv.SubItems(3).Text = lv.SubItems(7).Text & " & " & lv.SubItems(3).Text
                        'Else
                        '    lv.SubItems(3).Text = lv.SubItems(7).Text
                        'End If

                        lv.Group = g
                    End If
                    Confirming.lvSales.Items.Add(lv)
                    Dim y = Confirming.lvSales.Tag
                    If lv.Text = Confirming.lvSales.Tag Then
                        lv.Selected = True
                        STATIC_VARIABLES.CurrentID = Confirming.lvSales.Tag
                    End If


                End While
                r1.Close()
                cnnS.Close()
                Confirming.lvSales.Groups.Add(g)
                If Confirming.lvSales.Items.Count > 0 And Confirming.lvSales.SelectedItems.Count = 0 Then
                    Confirming.lvSales.TopItem.Selected = True
                End If
            End If

        Catch ex As Exception
            MsgBox("Lost Network Connection! Populate List" & ex.ToString, MsgBoxStyle.Critical, "Server not Available")


        End Try
        If Tab = "Confirm" Then
            If PoporRefresh = "Populate" Then
                If Confirming.lvConfirming.Items.Count > 0 Then
                    Confirming.lvConfirming.TopItem.Selected = True
                    'Confirming.LastID = Confirming.lvConfirming.SelectedItems(0).Text
                Else
                    Dim cls As New ConfirmingData
                    cls.PullCustomerINFO("Confirm", "")
                End If

            End If
        ElseIf Tab = "Dispatch" Then
            If PoporRefresh = "Populate" Then
                If Confirming.lvSales.Items.Count > 0 Then
                    Confirming.lvSales.TopItem.Selected = True


                Else
                    Dim cls As New ConfirmingData
                    cls.PullCustomerINFO("Dispatch", "")
                End If

            End If
        End If

    End Sub


    Public dset_PriLS As Data.DataSet = New Data.DataSet("PLS")
    Public da_PRI As SqlDataAdapter = New SqlDataAdapter
    Public Sub GetPrimaryLeadSource()
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetPLS", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Try
            cnn.Open()
            da_PRI.SelectCommand = cmdGet
            da_PRI.Fill(dset_PriLS, "PrimaryLeadSource")
            Dim cnt As Integer = 0
            cnt = dset_PriLS.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    Confirming.cboConfirmingPLS.Items.Clear()
                    Confirming.cboConfirmingPLS.Items.Add("")
                    Confirming.cboSalesPLS.Items.Clear()
                    Confirming.cboSalesSLS.Items.Add("")
                    Exit Select
                Case Is >= 1
                    Confirming.cboConfirmingPLS.Items.Clear()
                    Confirming.cboSalesPLS.Items.Clear()
                    Confirming.cboConfirmingPLS.Items.Add("")
                    Confirming.cboSalesPLS.Items.Add("")
                    Dim b
                    For b = 0 To dset_PriLS.Tables(0).Rows.Count - 1
                        Confirming.cboConfirmingPLS.Items.Add(dset_PriLS.Tables(0).Rows(b).Item(1))
                        Confirming.cboSalesPLS.Items.Add(dset_PriLS.Tables(0).Rows(b).Item(1))
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
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetSLS", cnn)
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
                Confirming.cboConfirmingSLS.Items.Clear()
                Confirming.cboConfirmingSLS.Items.Add("")
                Confirming.cboSalesSLS.Items.Clear()
                Confirming.cboSalesSLS.Items.Add("")
                Exit Select
            Case Is >= 1
                Confirming.cboConfirmingSLS.Items.Clear()
                Confirming.cboConfirmingSLS.Items.Add("")
                Confirming.cboSalesSLS.Items.Clear()
                Confirming.cboSalesSLS.Items.Add("")
                Dim b As Integer = 0
                For b = 0 To dset_SLS.Tables(0).Rows.Count - 1
                    Confirming.cboConfirmingSLS.Items.Add(dset_SLS.Tables(0).Rows(b).Item(0))
                    Confirming.cboSalesSLS.Items.Add(dset_SLS.Tables(0).Rows(b).Item(0))
                Next
                Exit Select
        End Select
        cnn.Close()

    End Sub
    Public Sub PullCustomerINFO(ByVal Tab As String, ByVal ID As String)

        If ID = "" Then
            Confirming.txtContact1.Text = ""
            Confirming.txtContact2.Text = ""
            Confirming.txtAddress.Text = ""
            Confirming.txtWorkHours.Text = ""
            Confirming.txtHousePhone.Text = ""
            Confirming.txtaltphone1.Text = ""
            Confirming.txtaltphone2.Text = ""
            Confirming.txtAlt1Type.Text = ""
            Confirming.txtAlt2Type.Text = ""
            Confirming.lnkEmail.Text = ""
            Confirming.txtApptDate.Text = ""
            Confirming.txtApptDay.Text = ""
            Confirming.txtApptTime.Text = ""
            Confirming.txtProducts.Text = ""
            Confirming.txtColor.Text = ""
            Confirming.txtQty.Text = ""
            Confirming.txtYrBuilt.Text = ""
            Confirming.txtYrsOwned.Text = ""
            Confirming.txtHomeValue.Text = ""
            Confirming.rtbSpecialInstructions.Text = ""
            Confirming.pnlCustomerHistory.Controls.Clear()
        End If
        If Tab = "Confirm" Then
            If ID = Confirming.LastID And Confirming.txtContact1.Text <> "" Then
                STATIC_VARIABLES.CurrentID = ID

            End If
        ElseIf Tab = "Dispatch" Then
            If ID = Confirming.LastIDS Then
                STATIC_VARIABLES.CurrentID = ID


            End If
        End If
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetCustomerINFO", Cnn)

        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        Try



            Cnn.Open()


            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            While r1.Read
                Confirming.txtContact1.Text = r1.Item(5) & " " & r1.Item(6)
                Confirming.txtContact2.Text = r1.Item(7) & " " & r1.Item(8)
                Confirming.txtAddress.Text = r1.Item(9) & " " & vbCrLf & r1.Item(10) & ", " & r1.Item(11) & " " & r1.Item(12)


                If r1.Item(7) = "" Then
                    Confirming.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19)
                Else
                    Confirming.txtWorkHours.Text = r1.Item(5) & ": " & r1.Item(19) & vbNewLine & r1.Item(7) & ": " & r1.Item(20)
                End If
                Confirming.txtHousePhone.Text = r1.Item(13)
                Confirming.txtaltphone1.Text = r1.Item(14)
                Confirming.txtaltphone2.Text = r1.Item(16)
                Confirming.txtAlt1Type.Text = r1.Item(15)
                Confirming.txtAlt2Type.Text = r1.Item(17)
                Confirming.lnkEmail.Text = r1.Item(52)
                Dim d
                d = Split(r1.Item(29), " ", 2)
                Trim(d(0))

                Confirming.txtApptDate.Text = d(0)
                Confirming.txtApptDay.Text = r1.Item(30)

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
                Confirming.txtApptTime.Text = correctedTime.ToString
                Confirming.txtProducts.Text = r1.Item(21) & vbCrLf & r1.Item(22) & vbCrLf & r1.Item(23)
                Confirming.txtColor.Text = r1.Item(24)
                Confirming.txtQty.Text = r1.Item(25)
                Confirming.txtYrBuilt.Text = r1.Item(27)
                Confirming.txtYrsOwned.Text = r1.Item(26)
                Confirming.txtHomeValue.Text = r1.Item(28)
                Confirming.rtbSpecialInstructions.Text = r1.Item(32)
            End While
            r1.Close()
            Cnn.Close()

        Catch ex As Exception
            'Cnn.Close()
            'Me.PullCustomerINFO(ID)
            MsgBox("Lost Network Connection! Pull Customer Info" & ex.ToString, MsgBoxStyle.Critical, "Server not Available")
        End Try
        If Tab = "Confirm" Then
            Confirming.LastID = ID
            STATIC_VARIABLES.CurrentID = ID
        ElseIf Tab = "Dispatch" Then
            Confirming.LastIDS = ID
            STATIC_VARIABLES.CurrentID = ID
        End If

        Dim c As New CustomerHistory
        If ID <> "" Then
            c.SetUp(Confirming, ID, Confirming.TScboCustomerHistory)
        End If

    End Sub
    Public Sub Confirm(ByVal ID As String, ByVal cmd As String, ByVal spokewith As String, ByVal User As String)
        If cmd = "Confirm Appointment" Then
            Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdIns As SqlCommand = New SqlCommand("dbo.Confirm", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            Dim param2 As SqlParameter = New SqlParameter("@User", User)
            Dim param3 As SqlParameter = New SqlParameter("@Spokewith", spokewith)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cmdIns.Parameters.Add(param2)
            cmdIns.Parameters.Add(param3)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()

            Dim c As New CustomerHistory
            c.SetUp(Confirming, ID, Confirming.TScboCustomerHistory)
        ElseIf cmd = "Undo Confirm Appointment" Then

            Dim cmdIns As SqlCommand = New SqlCommand("dbo.ConfirmUndo", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdIns.CommandType = CommandType.StoredProcedure
            cmdIns.Parameters.Add(param1)
            cnn.Open()
            Dim R1 = cmdIns.ExecuteReader
            R1.Read()
            R1.Close()
            cnn.Close()

            Dim c As New CustomerHistory
            c.SetUp(Confirming, ID, Confirming.TScboCustomerHistory)
        Else
            Exit Sub
        End If


    End Sub
    Public dset_Reps As Data.DataSet = New Data.DataSet("Reps")
    Public da_Reps As SqlDataAdapter = New SqlDataAdapter
    Public Sub PopReps()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetSalesReps", Cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        Try
            Cnn.Open()
            da_Reps.SelectCommand = cmdGet
            da_Reps.Fill(dset_Reps, "SalesReps")
            Dim cnt As Integer = 0
            cnt = dset_Reps.Tables(0).Rows.Count
            Select Case cnt
                Case Is <= 0
                    Confirming.Rep1.Items.Clear()
                    Confirming.Rep1.Items.Add("")
                    Confirming.Rep2.Items.Clear()
                    Confirming.Rep2.Items.Add("")

                    Exit Select
                Case Is >= 1
                    Confirming.Rep1.Items.Clear()

                    Confirming.Rep2.Items.Clear()
                    Confirming.Rep2.Items.Add("")
                    Dim b
                    For b = 0 To dset_Reps.Tables(0).Rows.Count - 1
                        Confirming.Rep1.Items.Add(dset_Reps.Tables(0).Rows(b).Item(0) & " " & dset_Reps.Tables(0).Rows(b).Item(1))
                        Confirming.Rep2.Items.Add(dset_Reps.Tables(0).Rows(b).Item(0) & " " & dset_Reps.Tables(0).Rows(b).Item(1))
                    Next
                    Exit Select
            End Select
            Cnn.Close()

        Catch ex As Exception

        End Try
    End Sub
    Public Sub ChangeRep(ByVal ID As String, ByVal Rep1 As String, ByVal Rep2 As String, ByVal origRep1 As String, ByVal OrigRep2 As String)
        Dim Description As String
        If origRep1 = Rep1 And OrigRep2 = "" And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " paired " & Rep2 & " up with " & Rep1
        ElseIf origRep1 <> "" And origRep1 <> Rep1 And OrigRep2 = "" And Rep2 = "" Then
            Description = STATIC_VARIABLES.CurrentUser & " changed assigned rep from " & origRep1 & " to " & Rep1
        ElseIf origRep1 = "" And OrigRep2 = "" And Rep2 = "" Then
            Description = STATIC_VARIABLES.CurrentUser & " issued Appt. to " & Rep1
        ElseIf origRep1 = "" And OrigRep2 = "" And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " issued Appt. to " & Rep1 & " and " & Rep2
        ElseIf origRep1 = "" And OrigRep2 = "" And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " issued Appt. to " & Rep1 & " and " & Rep2
        ElseIf origRep1 <> "" And OrigRep2 <> "" And Rep2 = "" And Rep1 <> origRep1 And OrigRep2 <> Rep1 Then
            Description = STATIC_VARIABLES.CurrentUser & " changed assigned rep from " & origRep1 & " and " & OrigRep2 & " to " & Rep1
        ElseIf origRep1 = Rep1 And OrigRep2 <> Rep2 And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " changed Rep 2 from " & OrigRep2 & " to " & Rep2 & ", leaving " & Rep1 & " and " & Rep2 & " paired up"
        ElseIf origRep1 <> Rep1 And OrigRep2 <> "" And OrigRep2 <> Rep2 And Rep2 <> "" And origRep1 <> Rep2 And OrigRep2 <> Rep1 Then
            Description = STATIC_VARIABLES.CurrentUser & " changed assigned pair from " & origRep1 & " and " & OrigRep2 & " to " & Rep1 & " and " & Rep2
        ElseIf origRep1 <> Rep1 And OrigRep2 <> "" And OrigRep2 = Rep2 And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " changed Rep 1 from " & origRep1 & " to " & Rep1 & ", leaving " & Rep1 & " and " & Rep2 & " paired up"
        ElseIf origRep1 <> Rep1 And OrigRep2 = "" And Rep2 <> "" Then
            Description = STATIC_VARIABLES.CurrentUser & " changed assigned rep from " & origRep1 & ", to pair: " & Rep1 & " and " & Rep2
        ElseIf origRep1 = Rep2 And OrigRep2 = Rep1 Then
            Description = STATIC_VARIABLES.CurrentUser & " changed pair around, making " & Rep1 & " the lead rep and " & Rep2 & " the secondary rep"
        ElseIf origRep1 <> Rep1 And OrigRep2 <> "" And OrigRep2 <> Rep2 And Rep2 <> "" And origRep1 <> Rep2 And OrigRep2 = Rep1 Then
            Description = STATIC_VARIABLES.CurrentUser & " changed assigned pair from " & origRep1 & " and " & OrigRep2 & " to " & Rep1 & " and " & Rep2
        ElseIf origRep1 = Rep1 And OrigRep2 <> "" And Rep2 = "" Then
            Description = STATIC_VARIABLES.CurrentUser & " removed " & OrigRep2 & " from Appt., making " & origRep1 & " the sole Sales Rep"
        ElseIf origRep1 <> "" And OrigRep2 <> "" And Rep2 = "" And Rep1 <> origRep1 And OrigRep2 = Rep1 Then
            Description = STATIC_VARIABLES.CurrentUser & " removed " & origRep1 & " from Appt., moving " & OrigRep2 & " up to Rep 1 as the sole Sales Rep"
        End If



        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdChange As SqlCommand
        cmdChange = New SqlCommand("dbo.SalesRepChanged", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim param3 As SqlParameter = New SqlParameter("@Description", Description)
        Dim param4 As SqlParameter = New SqlParameter("@Rep1", Rep1)
        Dim param5 As SqlParameter = New SqlParameter("@Rep2", Rep2)
        cmdChange.CommandType = CommandType.StoredProcedure
        cmdChange.Parameters.Add(param1)
        cmdChange.Parameters.Add(param2)
        cmdChange.Parameters.Add(param3)
        cmdChange.Parameters.Add(param4)
        cmdChange.Parameters.Add(param5)

        cnn.Open()
        Dim R1 = cmdChange.ExecuteReader
        R1.Read()
        R1.Close()
        cnn.Close()

        Confirming.OrigRep1 = Rep1
        Confirming.OrigRep2 = Rep2





        Me.Populate("Dispatch", Confirming.cboSalesPLS.Text, Confirming.cboSalesSLS.Text, Confirming.dpSales.Value.ToString, "Refresh")
        Dim c As New CustomerHistory
        c.SetUp(Confirming, ID, Confirming.TScboCustomerHistory)
    End Sub

  
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

