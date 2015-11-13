'Imports cmraude.Service1
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System.Windows.Forms

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic

Public Class FindLead
    Public SRCH As String = ""
    Public FID As String = ""
    Public PREVIOUS_ID As String = ""
    Public i As ListViewItem
    Public x As String = ""
  
    'Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.lstSearchResults.Items.Clear()
    '    Search(WhatIsChecked())
    'End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Me.Close()
    'End Sub
    'Private Sub rdoPhone_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Me.txtSearchPhone.Enabled = True
    '    Me.txtSearchLName.Enabled = False
    '    Me.txtSearchLName.Text = ""
    '    Me.txtSearchID.Enabled = False
    '    Me.txtSearchID.Text = ""
    '    Me.txtSearchAddy.Enabled = False
    '    Me.txtSearchAddress.Enabled = False
    '    Me.txtSearchAddress.Text = ""
    '    Me.txtSearchAddy.Text = ""
    'End Sub

    'Private Sub rdoLName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.txtSearchPhone.Enabled = False
    '    Me.txtSearchPhone.Text = ""
    '    Me.txtSearchLName.Enabled = True

    '    Me.txtSearchID.Enabled = False
    '    Me.txtSearchID.Text = ""
    '    Me.txtSearchAddy.Enabled = False
    '    Me.txtSearchAddress.Enabled = False
    '    Me.txtSearchAddress.Text = ""
    '    Me.txtSearchAddy.Text = ""
    'End Sub

    'Private Sub rdoID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.txtSearchPhone.Enabled = False
    '    Me.txtSearchPhone.Text = ""
    '    Me.txtSearchLName.Enabled = False
    '    Me.txtSearchLName.Text = ""


    '    Me.txtSearchID.Enabled = True

    '    Me.txtSearchAddy.Enabled = False
    '    Me.txtSearchAddress.Enabled = False
    '    Me.txtSearchAddress.Text = ""
    '    Me.txtSearchAddy.Text = ""
    'End Sub

    'Private Sub rdoAddy_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Me.txtSearchPhone.Enabled = False
    '    Me.txtSearchPhone.Text = ""
    '    Me.txtSearchLName.Enabled = False
    '    Me.txtSearchLName.Text = ""


    '    Me.txtSearchID.Enabled = False
    '    Me.txtSearchID.Text = ""

    '    Me.txtSearchAddy.Enabled = True
    '    Me.txtSearchAddress.Enabled = True

    'End Sub

    'Private Function WhatIsChecked()
    '    Dim srchWhat As String = ""
    '    Dim x As System.Windows.Forms.Control
    '    For Each x In Me.GroupBox1.Controls
    '        If TypeOf x Is System.Windows.Forms.RadioButton Then
    '            If x.Enabled = True Then
    '                Dim y As System.Windows.Forms.RadioButton
    '                y = x
    '                If y.Checked = True Then
    '                    srchWhat = y.Name.ToString
    '                End If
    '            End If
    '        End If
    '    Next
    '    SRCH = srchWhat
    '    Return SRCH
    'End Function
    
    'Private Sub Search(ByVal s As String)
    '    Select Case s
    '        Case Is = "rdoID"
    '            Dim x As New cmraude.getFindLead
    '            If Me.txtSearchID.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            x.FindLNs(Me.txtSearchID.Text)
    '            Dim b
    '            Me.lstSearchResults.Items.Clear()
    '            For b = 1 To x.CntLN
    '                Dim lvResult As New System.Windows.Forms.ListViewItem
    '                lvResult.Text = x.LNs(b).ToString
    '                lvResult.SubItems.Add(x.LNnfo(b).ToString)
    '                lvResult.SubItems.Add(x.LNAddress(b).ToString)
    '                lvResult.SubItems.Add(x.LNPhone(b).ToString)
    '                Me.lstSearchResults.Items.Add(lvResult)
    '            Next
    '            Exit Select
    '        Case Is = "rdoPhone"
    '            If Me.txtSearchPhone.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            Dim z As New cmraude.getFindLead
    '            Dim phone As String = ""
    '            phone = Me.txtSearchPhone.Text
    '            z.FindPhones(phone)
    '            Me.lstSearchResults.Items.Clear()
    '            Dim b
    '            For b = 1 To z.CntLN
    '                Dim lvResult As New ListViewItem
    '                lvResult.Text = z.LNs(b).ToString
    '                lvResult.SubItems.Add(z.LNnfo(b).ToString)
    '                lvResult.SubItems.Add(z.LNAddress(b).ToString)
    '                lvResult.SubItems.Add(z.LNPhone(b).ToString)
    '                Me.lstSearchResults.Items.Add(lvResult)
    '            Next
    '            Exit Select
    '        Case Is = "rdoLName"
    '            If Me.txtSearchLName.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            Dim x As New cmraude.getFindLead
    '            x.FindLName(Me.txtSearchLName.Text)
    '            Me.lstSearchResults.Items.Clear()
    '            Dim b
    '            For b = 1 To x.CntLN
    '                Dim lvResult As New ListViewItem
    '                lvResult.Text = x.LNs(b).ToString
    '                lvResult.SubItems.Add(x.LNnfo(b).ToString)
    '                lvResult.SubItems.Add(x.LNAddress(b).ToString)
    '                lvResult.SubItems.Add(x.LNPhone(b).ToString)
    '                Me.lstSearchResults.Items.Add(lvResult)
    '            Next
    '            Exit Select
    '        Case Is = "rdoAddy"
    '            If Me.txtSearchAddress.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            Dim x As New cmraude.getFindLead
    '            If Me.txtSearchAddress.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            If Me.txtSearchAddy.Text = "" Then
    '                MsgBox("You must supply criteria to search for.", MsgBoxStyle.Exclamation, "Error Searching")
    '                Exit Sub
    '            End If
    '            x.FindAddress(Me.txtSearchAddress.Text, Me.txtSearchAddy.Text)
    '            Me.lstSearchResults.Items.Clear()
    '            Dim b
    '            For b = 1 To x.CntLN
    '                Dim lvResult As New ListViewItem
    '                lvResult.Text = x.LNs(b).ToString
    '                lvResult.SubItems.Add(x.LNnfo(b).ToString)
    '                lvResult.SubItems.Add(x.LNAddress(b).ToString)
    '                lvResult.SubItems.Add(x.LNPhone(b).ToString)
    '                Me.lstSearchResults.Items.Add(lvResult)
    '            Next
    '            Exit Select
    '        Case Else
    '            MsgBox("You must select a criteria to search by.", MsgBoxStyle.Exclamation, "Error Searching")
    '            Exit Select
    '    End Select
    'End Sub

    'Private Sub FindLead_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    'Me.txtSearchPhone.Clear()
    '    'Me.txtSearchLName.Clear()
    '    'Me.txtSearchID.Clear()
    '    'Me.txtSearchAddy.Clear()
    '    'Me.txtSearchAddress.Clear()
    '    'Me.lstSearchResults.Items.Clear()












    '    Me.txtSearchAddress.Enabled = False
    '    Me.txtSearchID.Enabled = False
    '    Me.txtSearchLName.Enabled = False
    '    Me.txtSearchAddy.Enabled = False
    '    Me.rdoPhone.Checked = True
    '    Me.txtSearchPhone.Select()
    'End Sub

    'Private Sub txtSearchPhone_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs)

    'End Sub

    'Private Sub txtSearchPhone_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim z As New cmraude.getFindLead
    '    Me.lstSearchResults.Items.Clear()
    '    Dim x As String = ""
    '    x = Me.txtSearchPhone.Text
    '    Select Case x
    '        Case Is = "(   )"
    '            Exit Sub
    '            Exit Select
    '        Case Is = ""
    '            Exit Sub
    '            Exit Select
    '        Case Else
    '            z.FindPhones(Me.txtSearchPhone.Text)
    '            Dim b
    '            For b = 1 To z.CntLN
    '                Dim lv As New ListViewItem
    '                lv.Text = z.LNs(b).ToString
    '                lv.SubItems.Add(z.LNnfo(b).ToString)
    '                lv.SubItems.Add(z.LNAddress(b).ToString)
    '                lv.SubItems.Add(z.LNPhone(b).ToString)
    '                Me.lstSearchResults.Items.Add(lv)
    '            Next
    '            Exit Select
    '    End Select

    'End Sub

    'Private Sub txtSearchLName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim z As New cmraude.getFindLead
    '    Me.lstSearchResults.Items.Clear()
    '    Dim x As String = ""
    '    x = Me.txtSearchLName.Text
    '    'Select Case x
    '    '    Case Is = "(   )"
    '    '        Exit Sub
    '    '        Exit Select
    '    '    Case Is = ""
    '    '        Exit Sub
    '    '        Exit Select
    '    '    Case Else
    '    z.FindLName(Me.txtSearchLName.Text)
    '    Dim b
    '    For b = 1 To z.CntLN
    '        Dim lv As New ListViewItem
    '        lv.Text = z.LNs(b).ToString
    '        lv.SubItems.Add(z.LNnfo(b).ToString)
    '        lv.SubItems.Add(z.LNAddress(b).ToString)
    '        lv.SubItems.Add(z.LNPhone(b).ToString)
    '        Me.lstSearchResults.Items.Add(lv)
    '    Next
    '    'Exit Select
    '    'End Select

    'End Sub



    Private Sub txtsearch_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsearch.GotFocus
        Me.pctFind.Visible = False
        Me.Label1.Visible = False
    End Sub

    Private Sub txtsearch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtsearch.LostFocus
        If Me.txtsearch.Text = "" Then
            Me.pctFind.Visible = True
            Me.Label1.Visible = True
        End If
    End Sub

    Private Sub txtsearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtsearch.TextChanged


        x = Me.txtsearch.Text
        If x = "" Then
            Exit Sub
        End If
        Dim len As Integer = x.ToString.Length
        Select Case len
            Case Is >= 3
                Me.BackgroundWorker1_DoWork(Nothing, Nothing)

                
                'Dim ds As Data.DataSet = y.DSFind
                'Dim r As Data.DataRow
                'Me.lstSearchResults.Items.Clear()
                'For Each r In ds.Tables(0).Rows
                '    Dim lv As New ListViewItem
                '    lv.Text = r.Item("ID")
                '    lv.SubItems.Add(r.Item("Name"))
                '    lv.SubItems.Add(r.Item("Address"))
                '    lv.SubItems.Add(r.Item("Phone"))
                '    Me.lstSearchResults.Items.Add(lv)
                'Next
                Exit Select
            Case Else
                Exit Select
        End Select
    End Sub

    Private Sub FindLead_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.txtsearch.Text = ""
        Me.lstSearchResults.Items.Clear()

        'Me.Dispose()
    End Sub

    Private Sub FindLead_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.lstSearchResults.Select()
        If Me.txtsearch.Text = "" Then
            Me.pctFind.Visible = True
            Me.Label1.Visible = True

        End If
    End Sub

    Private Sub lstSearchResults_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstSearchResults.MouseClick
        Dim selitem As ListViewItem
        selitem = Me.lstSearchResults.GetItemAt(e.X, e.Y)
        If selitem Is Nothing Then
            Exit Sub
        End If

        Dim found As Integer = 0
        Dim open As Integer = 0
        Dim focused As Form
        If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
            focused = STATIC_VARIABLES.ActiveChild
        End If
        For Each frm As Form In Main.MdiChildren

            Select Case frm.Name
                Case "Confirming", "Sales", "Administration", "ColdCalling", "WCaller", "Finance", "MarketingManager", _
                        "Recovery", "PreviousCustomer", "SecondSource", "Installation"
                    open += 1
                    Select Case frm.Name
                        Case Is = "Confirming"
                            'Dim lv As New ListViewItem
                            Dim tab As String = Confirming.Tab
                            'If tab = "Confirm" Then
                            Dim foundConfirming As Integer = 0
                            For Each Me.i In Confirming.lvConfirming.Items
                                If Me.i.Text = selitem.Text Then

                                    Me.i.Selected = True
                                    found += 1
                                    foundConfirming += 1
                                    If Confirming.TabControl1.SelectedIndex <> 0 Then
                                        Confirming.TabControl1.SelectedIndex = 0
                                    End If

                                    'Confirming.Focus()
                                    'Me.Close()

                                    'Exit Sub
                                End If
                            Next
                            '''''''''''''''   'End If''
                            If Confirming.lvConfirming.SelectedItems.Count <> 0 Then
                                If Confirming.lvConfirming.SelectedItems(0).Text <> selitem.Text Then
                                    For Each Me.i In Confirming.lvSales.Items
                                        If Me.i.Text = selitem.Text Then
                                            Me.i.Selected = True
                                            found += 1
                                            foundConfirming += 1
                                            If Confirming.TabControl1.SelectedIndex <> 1 Then
                                                Confirming.TabControl1.SelectedIndex = 1
                                            End If

                                            'Confirming.Focus()
                                            'Me.Close()

                                            'Exit Sub
                                        End If
                                    Next
                                End If
                            End If

                            If foundConfirming = 0 Then
                                Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                                Dim cmdGET As SqlCommand = New SqlCommand("dbo.AlertConfirm", cnn)

                                Dim param1 As SqlParameter = New SqlParameter("@ID", selitem.Text)
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
                                        If Me.i.Text = selitem.Text Then

                                            Me.i.Selected = True

                                            'Dim c As New ConfirmingData
                                            'c.PullCustomerINFO(Confirming.Tab, Me.i.Text)

                                            found += 1
                                            foundConfirming += 1
                                            If Confirming.TabControl1.SelectedIndex <> 0 Then
                                                Confirming.TabControl1.SelectedIndex = 0
                                            End If

                                            Confirming.Focus()
                                            'Me.Label5_Click(Nothing, Nothing)
                                            Main.forcettconfirm = True



                                            'Exit Sub
                                        End If

                                    Next
                                    If Confirming.lvConfirming.SelectedItems.Count = 0 Then
                                        Confirming.dpConfirming.Value = was
                                    End If
                                    was = Confirming.dpSales.Value

                                    If found = 0 Then

                                        Confirming.dpSales.Value = r1.Item(0)
                                        For Each Me.i In Confirming.lvSales.Items
                                            If Me.i.Text = selitem.Text Then
                                                Me.i.Selected = True
                                                found += 1
                                                foundConfirming += 1
                                                If Confirming.TabControl1.SelectedIndex <> 1 Then
                                                    Confirming.TabControl1.SelectedIndex = 1
                                                End If

                                                Confirming.Focus()
                                                Main.forcettsales = True
                                                Confirming.Create_Tooltip(Confirming.dpSales, "search results")
                                                'Me.Label5_Click(Nothing, Nothing)



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

                            Dim param1 As SqlParameter = New SqlParameter("@ID", selitem.Text)
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
                                                    If Me.i.Text = selitem.Text Then
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
                                                If Me.i.Text = selitem.Text Then
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
                                                If Me.i.Text = selitem.Text Then
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

                        Case Is = "Financing"

                        Case Is = "ColdCalling"

                        Case Is = "SecondSource"



                    End Select

            End Select







        Next
        If focused IsNot Nothing Then
            focused.Focus()
        End If

        If found = 0 And open <> 0 Then
            ConfirmingSingleRecord.ID = selitem.Text
            If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
                STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
            End If
            ConfirmingSingleRecord.Show()
            'MsgBox("Cannot find record anywhere in your" & vbCr & "open forms with customer lists!", MsgBoxStyle.Critical, "Alert Link Failed")
        ElseIf open = 0 Then
            ConfirmingSingleRecord.ID = selitem.Text
            If STATIC_VARIABLES.ActiveChild IsNot Nothing Then
                STATIC_VARIABLES.ActiveChild.WindowState = FormWindowState.Normal
            End If
            ConfirmingSingleRecord.Show()
            'MsgBox("You have no lists open to link to!", MsgBoxStyle.Critical, "Alert Link Failed")
        End If
       

        Me.BackgroundWorker1.Dispose()
        'Me.Close()
        Me.Dispose()








        'MsgBox(selitem.Text & " for dummy forms to be designed.", MsgBoxStyle.Information)
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.LinkLabel1.LinkVisited = True
        Dim tt As New ToolTip
        tt.BackColor = Color.White
        tt.IsBalloon = True
        tt.ToolTipTitle = "More About Searching..."
        tt.UseAnimation = True

        tt.Show("Type text here to begin searching. Search will not begin until you type" & vbCr & _
        "at least 3 characters. You can search by first or last name, street address," & vbCr & _
        "street name, phone number, or Customer ID. " & vbCr & vbCr & _
        "When you see the record you are looking for, Click it to access this" & vbCr & _
        "Customer's information" & vbCr & vbCr & "[Click Here To Close]", _
        Me.txtsearch, 200, -170)
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim y As New FindLogic
        y.Search(x)
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Me.txtsearch.Select()
    End Sub
End Class
