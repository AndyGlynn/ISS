Public Class frmImportBulkData

    Public status As String = ""
    
    Private Sub btnOFD_Click(sender As Object, e As EventArgs) Handles btnOFD.Click
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        Me.FolderBrowserDialog1.ShowDialog()
        Dim path As String = Me.FolderBrowserDialog1.SelectedPath
        Me.txtDirectory.Text = path
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        '' reset me, reset all mem usage 
        '' close me
        ResetForm()
        Me.Close()

    End Sub

    Private Sub ResetForm()

        Me.txtDirectory.Text = ""
        Me.chkAllFiles.CheckState = CheckState.Checked
        Me.chkAutoInsert.CheckState = CheckState.Checked
        Me.chkAutoScrubAddress.CheckState = CheckState.Unchecked
        Me.chkDisplayReport.CheckState = CheckState.Unchecked
        Me.cboFiles.Text = ""
        Me.cboFiles.Enabled = True

        Me.pbRows.Value = 0
        Me.pbMemory.Value = 0
        Me.lblTable.Text = ""
        Me.tblCurCnt.Text = ""
        Me.lblCurrentRecordNum.Text = ""
        Me.rtfPreview.Text = ""
        Me.lstScrubbed.Clear()

        Me.lblTableTotal.Text = ""
        Me.lblTotalRecordNums.Text = ""
        Me.lblCnt.Text = ""

    End Sub


    Private Sub txtDirectory_TextChanged(sender As Object, e As EventArgs) Handles txtDirectory.TextChanged
        Dim path As String = Me.txtDirectory.Text
        Dim arFile As New Hashtable
        Dim cntFiles As Integer = 0
        If path <> "" Then
            If path.ToString.Length > 3 Then
                Dim x As New ImportData_V2.FileOperations
                arFile = x.Get_Files(path, "*.csv")
                x = Nothing
            End If
        End If
        Me.cboFiles.Items.Clear()
        Dim b As Object

        '' targeted tables
        '' 
        '' 
        ''- i360__Appointment__c
        ''- i360__Prospect__c
        ''- i360__Sale__c
        ''- _User
        ''- i360__Staff__c
        ''- i360__Marketing_Source__c
        ''- i360__Marketing_Task__c
        ''- i360__Lead_Source__c


        '' tables with addresses to be scrubbed..
        'Event  -  495
        ''i360__Appointment__c - 44023
        ''Prospect - 41177
        ''i360__Staff__c   (maybe on this one. broken addresses and no real structure. like four records have complete addresses.)  - 371
        ''
        If arFile.Count = 0 Then
            Form6.Label2.Text = "Couldn't Find any Data to Import! Please select another directory!"
            Exit Sub
        End If


        For Each b In arFile.Keys
            Dim filename As String = Break_FileName_Apart(b.ToString)
            Dim str_FullName As String = (path & "\" & filename)
            Select Case filename
                Case Is = "i360__Appointment__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Prospect__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Sale__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "User"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Staff__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Marketing_Source__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Marketing_Task__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "i360__Lead_Source__c.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Is = "User.csv"
                    Me.cboFiles.Items.Add(filename)
                    Exit Select
                Case Else
                    Exit Select
            End Select
            cntFiles += 1
        Next
        Me.lblCntFiles.Text = cntFiles.ToString

        If Me.chkAllFiles.CheckState = CheckState.Checked Then
            '' proceed to process all files
            '' Me.chkDisplayReport.CheckState = CheckState.Unchecked
            Me.chkAutoInsert.CheckState = CheckState.Checked
            Me.cboFiles.Enabled = False
            Dim response As Integer = MsgBox("Are you sure you want to begin?" & vbCrLf & vbCrLf & "This operation cannot be cancelled and may take some time...", MsgBoxStyle.YesNo, "Begin Operation?")
            Select Case response
                Case vbYes
                    Dim y As New ImportData_V2.FileOperations
                    Dim cnt As Integer = 0
                    Dim x As String
                    For Each x In Me.cboFiles.Items
                        cnt += 1
                    Next
                    Me.lblTableTotal.Text = cnt.ToString
                    Dim iter As Integer = 0
                    Dim xy As String

                    Dim startTime As String = Date.Now.ToString
                    Dim arTableNames As New ArrayList
                    For Each xy In Me.cboFiles.Items
                        iter += 1
                        Me.tblCurCnt.Text = iter.ToString
                        Try
                            '' create 'batch report' here

                            ProcessFile((Me.txtDirectory.Text & "\" & xy.ToString))
                            arTableNames.Add(xy.ToString)
                        Catch ex As Exception
                            'Dim production As Boolean
                            'If Me.rdoDev.Checked = True Then
                            '    production = False
                            'ElseIf Me.rdoProduction.Checked = True Then
                            '    production = True
                            'End If
                            y.WriteErrorToLog(ex.Message.ToString, Me.chkDisplayReport.Checked, False)
                            y = Nothing
                        End Try
                    Next
                    

                    Me.lblMessage.Text = "Process Files Complete. "


                    Try
                        ''UNCOMMENT FOR PRODUCTION
                        Dim imp_data As New ImportData  '' long process here.
                        imp_data = Nothing

                    Catch ex As Exception
                        Dim msg As String = ex.Message
                        MsgBox(msg, MsgBoxStyle.Critical, "DEBUG INFO - OLD IMPORT DATA CLASS")
                    End Try




                    Me.lblMessage.Text = "Import Successful"
                    Form6.lblTable.Visible = False
                    Form6.lblTotalRecordNum.Visible = False
                    Form6.lblCurrentRecordNum.Visible = False
                    Form6.lblTblCurCnt.Visible = False
                    Dim stopTime As String = Date.Now.ToString
                    Dim z As New ImportData_V2.FileOperations

                    Dim whereIsReport As String = z.CreateBatchReport(startTime, stopTime, arTableNames) '' creates file and returns file name 

                    z = Nothing

                    Exit Select

                    If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                        System.Diagnostics.Process.Start(whereIsReport)
                    End If

                Case vbNo
                    Exit Select
            End Select
        ElseIf Me.chkAllFiles.CheckState = CheckState.Unchecked Then
            Me.cboFiles.Enabled = True
            Me.lblTableTotal.Text = "1"
            Me.tblCurCnt.Text = "1"
        End If
    End Sub

    Private Sub ProcessFile(ByVal File As String)

        'MsgBox("File To Process: " & File.ToString, MsgBoxStyle.Information, "DEBUG INFO")

        Try
            Me.lblMessage.Text = ""
            Form6.Label2.Text = "Loading Temp Table Into Memory"




            Dim file_path As String = File
            Dim arRows As New ArrayList
            If file_path <> "" Then
                Me.rtfPreview.Clear()
                Dim y As New ImportData_V2.FileOperations
                Dim arLines As ArrayList
                arLines = y.Preview_Read_File(file_path)
                Dim arHeadings As New ArrayList
                arHeadings = y.CreateColumnHeadings(arLines)

                Me.lstScrubbed.Clear()
                Dim yy As String
                Dim cnt As Integer = 0
                For Each yy In arHeadings
                    Me.lstScrubbed.Columns.Add(yy.ToString)
                    cnt += 1
                Next
                Dim g As String

                Me.Cursor = Cursors.WaitCursor

                Dim cntLines As Integer = 0
                For Each g In arLines
                    cntLines += 1
                Next

                Me.pbMemory.Maximum = cntLines
                Me.pbMemory.Minimum = 0
                Me.pbMemory.Value = 0
                Form6.ProgressBar1.Maximum = cntLines
                Form6.ProgressBar1.Minimum = 0
                Form6.ProgressBar1.Value = 0



                Me.lblCnt.Text = (cntLines - 1).ToString
                Me.lblTotalRecordNums.Text = (cntLines - 1).ToString

                Me.pbRows.Maximum = (cntLines - 1)
                Me.pbRows.Minimum = 0
                Me.pbRows.Value = 0
                Form6.pbRowMax = cntLines - 1



                If cntLines <= 1 Then
                    Me.rtfPreview.Text = arLines(0).ToString

                ElseIf cntLines > 10 Then
                    Dim dd As Integer = 0
                    For dd = 0 To 10 Step 1
                        Me.rtfPreview.Text += arLines(dd)
                    Next
                    Dim i As Integer = 0
                    For i = 1 To 4
                        Try
                            Dim lv As New ListViewItem
                            Dim arpieces()
                            Dim strLine As String = ""
                            Dim line()
                            ReDim line(arHeadings.Count - 1)
                            arpieces = y.CreateLineScrubbed(arLines(i))
                            lv.Text = Replace(arpieces(0), Chr(34), "", 1, arpieces(0).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            line(0) = Replace(arpieces(0), Chr(34), "", 1, arpieces(0).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim bb As Integer = 0
                            For bb = 1 To (arpieces.Length - 1)
                                lv.SubItems.Add(Replace(arpieces(bb), Chr(34), "", 1, arpieces(bb).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text))
                                line(bb) = Replace(arpieces(bb), Chr(34), "", 1, arpieces(bb).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            Next
                            Me.lstScrubbed.Items.Add(lv)
                            arRows.Add(line)
                            arpieces = Nothing
                        Catch ex As Exception
                            Dim err As String = ex.Message

                            status = "FAIL" & "-" & err
                            Me.lblMessage.Text = status

                            Dim xy As New ImportData_V2.FileOperations

                            Dim showReport As Boolean = False
                            If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                                showReport = True
                            ElseIf Me.chkDisplayReport.CheckState = CheckState.Unchecked Then
                                showReport = False
                            End If
                            xy.WriteErrorToLog(err, showReport, False)
                            xy = Nothing

                            Me.Cursor = Cursors.Default
                        End Try

                    Next



                    ' make sure ui shows preview
                    Application.DoEvents()

                    '
                    ' now loop through them all / displaying progress
                    ' 


                    Dim cc As Integer = 0
                    Me.Cursor = Cursors.WaitCursor

                    For cc = 1 To (cntLines - 1)
                        Try


                            Dim line()
                            ReDim line(arHeadings.Count - 1)
                            Dim arpieces()
                            arpieces = y.CreateLineScrubbed(arLines(cc))
                            line(0) = Replace(arpieces(0), Chr(34), "", 1, arpieces(0).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            Dim ccc As Integer = 0
                            For ccc = 1 To (arHeadings.Count - 1) '' anything more then the 'heading count' will be dropped off. Followed MS SQL Import Feature
                                line(ccc) = Replace(arpieces(ccc), Chr(34), "", 1, arpieces(ccc).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            Next
                            arRows.Add(line)
                            Me.pbMemory.Increment(1)
                            Form6.ProgressBar1.PerformStep()

                            Application.DoEvents()

                        Catch ex As Exception
                            Dim err As String = ex.Message
                            status = "FAIL" & "-" & err
                            Me.lblMessage.Text = status

                            Dim xy As New ImportData_V2.FileOperations

                            Dim showReport As Boolean = False
                            If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                                showReport = True
                            ElseIf Me.chkDisplayReport.CheckState = CheckState.Unchecked Then
                                showReport = False
                            End If
                            xy.WriteErrorToLog(err, showReport, False)
                            xy = Nothing

                            Me.Cursor = Cursors.Default
                        End Try


                    Next
                    Me.lblMessage.Text = "Loaded Table Successfully To Memory."
                End If


                '
                ' cnt = the number of 'columns' the 'line' needs to be split into 
                ' 

                Me.Cursor = Cursors.Default
                Application.DoEvents()

                If Me.chkAutoInsert.CheckState = CheckState.Checked Then
                    Try
                        Form6.Label2.Text = "Inserting Records To Temporary Database"

                        Dim z As New ImportData_V2.sqlOperations
                        Dim name() = Split(File, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
                        Dim t_name As String = name(name.Length - 1).ToString
                        Dim name2() = Split(t_name, ".", -1, Microsoft.VisualBasic.CompareMethod.Text)
                        Dim Table_Name As String = name2(0).ToString

                        ''UNCOMMENT FOR PRODUCTION
                        Me.lblTable.Text = Table_Name
                        z.Create_Table(arHeadings, arRows, Table_Name, Me.chkDisplayReport.CheckState, False, cntLines)
                        z = Nothing

                        Me.lblMessage.Text = "Records Inserted Successfully."

                        Application.DoEvents()

                        If Me.chkAutoScrubAddress.CheckState = CheckState.Checked Then
                            If Table_Name = "i360__Prospect__c" Then
                                Me.Cursor = Cursors.WaitCursor
                                Me.tblCurCnt.Text = "1"
                                Me.lblTableTotal.Text = "1"
                                Me.lblOperation.Text = "Verifying Addresses"

                                Me.lblTable.Text = Table_Name
                                Me.lblMessage.Text = "Updating Addresses . . . "
                                Application.DoEvents()
                                Dim mpt_bulk As New mappointBulkScrub
                                Dim sql_bulk As New ImportData_V2.sqlOperations
                                Dim arAddresses As List(Of ImportData_V2.sqlOperations.Record_And_Address)
                                arAddresses = sql_bulk.Get_Addresses_tblProspect(False)
                                Me.pbRows.Minimum = 0
                                Me.pbRows.Value = 0
                                Me.pbRows.Maximum = sql_bulk.CountRecords(Table_Name, False)
                                Me.lblTotalRecordNums.Text = sql_bulk.CountRecords(Table_Name, False)
                                Dim iteration As Integer = 0
                                ''UNCOMMENT FOR PRODUCTION
                                For Each x As ImportData_V2.sqlOperations.Record_And_Address In arAddresses
                                    iteration += 1
                                    ''If iteration >= 500 Then
                                    ''    Exit For
                                    ''End If
                                    mpt_bulk.VerifyAddress(x, False)
                                    pbRows.Increment(1)
                                    Me.lblCurrentRecordNum.Text = iteration
                                    Application.DoEvents()
                                Next

                                Me.Cursor = Cursors.Default
                                Application.DoEvents()
                                '' clean up 
                                mpt_bulk = Nothing
                                sql_bulk = Nothing
                                arAddresses = Nothing

                            End If
                        End If



                        Application.DoEvents()



                    Catch ex As Exception


                        Dim err As String = ex.Message

                        status = "FAIL" & "-" & err
                        Me.lblMessage.Text = status

                        Dim xy As New ImportData_V2.FileOperations

                        Dim showReport As Boolean = False
                        If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                            showReport = True
                        ElseIf Me.chkDisplayReport.CheckState = CheckState.Unchecked Then
                            showReport = False
                        End If
                        xy.WriteErrorToLog(err, showReport, False)
                        xy = Nothing
                        Me.Cursor = Cursors.Default
                    End Try

                ElseIf Me.chkAutoInsert.CheckState = CheckState.Unchecked Then
                    Me.Cursor = Cursors.Default
                End If

            End If
        Catch ex As Exception
            Dim msg As String = ex.Message
            Me.lblMessage.Text = "FAIL - " & msg.ToString
            Dim y As New ImportData_V2.FileOperations

            Dim showReport As Boolean = False
            If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                showReport = True
            ElseIf Me.chkDisplayReport.CheckState = CheckState.Unchecked Then
                showReport = False
            End If
            y.WriteErrorToLog(msg, showReport, False)
            y = Nothing
            Me.Cursor = Cursors.Default
        End Try
        Me.Cursor = Cursors.Default
    End Sub



    Private Sub cboFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboFiles.SelectedIndexChanged
        Try
            Me.lblMessage.Text = ""




            Dim b As ComboBox = Me.cboFiles
            Dim file_path As String = (Me.txtDirectory.Text & "\" & b.Text)
            Dim arRows As New ArrayList
            If b.Text <> "" Then
                Me.rtfPreview.Clear()
                Dim y As New ImportData_V2.FileOperations
                Dim arLines As ArrayList
                arLines = y.Preview_Read_File(file_path)
                Dim arHeadings As New ArrayList
                arHeadings = y.CreateColumnHeadings(arLines)

                Me.lstScrubbed.Clear()
                Dim yy As String
                Dim cnt As Integer = 0
                For Each yy In arHeadings
                    Me.lstScrubbed.Columns.Add(yy.ToString)
                    cnt += 1
                Next
                Dim g As String

                Me.Cursor = Cursors.WaitCursor

                Dim cntLines As Integer = 0
                For Each g In arLines
                    cntLines += 1
                Next



                Me.pbMemory.Maximum = cntLines
                Me.pbMemory.Minimum = 0
                Me.pbMemory.Value = 0

                Form6.ProgressBar1.Maximum = cntLines
                Form6.ProgressBar1.Minimum = 0
                Form6.ProgressBar1.Value = 0

                Me.lblCnt.Text = (cntLines - 1).ToString
                Me.lblTotalRecordNums.Text = (cntLines - 1).ToString

                Me.pbRows.Minimum = cntLines
                Me.pbRows.Minimum = 0
                Me.pbRows.Value = 0

                If cntLines <= 1 Then
                    Me.rtfPreview.Text = arLines(0).ToString

                ElseIf cntLines > 10 Then
                    Dim dd As Integer = 0
                    For dd = 0 To 10 Step 1
                        Me.rtfPreview.Text += arLines(dd)
                    Next
                    Dim i As Integer = 0
                    For i = 1 To 4
                        ' Try
                        Dim lv As New ListViewItem
                        Dim arpieces()
                        Dim strLine As String = ""

                        arpieces = y.CreateLineScrubbed(arLines(i))
                        lv.Text = Replace(arpieces(0), Chr(34), "", 1, arpieces(0).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                        Dim bb As Integer = 0
                        For bb = 1 To (arpieces.Length - 1)
                            lv.SubItems.Add(Replace(arpieces(bb), Chr(34), "", 1, arpieces(bb).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text))

                        Next
                        Me.lstScrubbed.Items.Add(lv)


                    Next



                    '' make sure ui shows preview
                    Application.DoEvents()

                    ''
                    '' now loop through them all / displaying progress
                    '' 


                    Dim cc As Integer = 0
                    Me.Cursor = Cursors.WaitCursor
                    Me.pbRows.Maximum = cntLines
                    Me.pbRows.Minimum = 0
                    Form6.pbRowMax = cntLines
                    For cc = 1 To (cntLines - 1)



                        Dim line()
                        ReDim line(arHeadings.Count - 1)
                        Dim arpieces()
                        arpieces = y.CreateLineScrubbed(arLines(cc))
                        line(0) = Replace(arpieces(0), Chr(34), "", 1, arpieces(0).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                        Dim ccc As Integer = 0
                        For ccc = 1 To (arHeadings.Count - 1) '' anything after the ' heading count' in the array, will be dumped off:> Followed MS SQL Import Feature

                            line(ccc) = Replace(arpieces(ccc), Chr(34), "", 1, arpieces(ccc).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)

                        Next
                        arRows.Add(line)
                        Me.pbMemory.Increment(1)
                        Form6.ProgressBar1.PerformStep()
                        Application.DoEvents()

                        Me.lblMessage.Text = "Loaded Table Successfully To Memory."

                    Next

                End If

                ''
                '' cnt = the number of 'columns' the 'line' needs to be split into 
                '' 

                pbMemory.Value = cntLines
                Form6.ProgressBar1.Value = cntLines
                Application.DoEvents()

                If Me.chkAutoInsert.CheckState = CheckState.Checked Then
                    ' Try
                    Dim z As New ImportData_V2.sqlOperations
                    Dim zz As New ImportData_V2.FileOperations
                    Dim name() = Split(Me.cboFiles.Text, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
                    Dim t_name As String = name(name.Length - 1).ToString
                    Dim name2() = Split(t_name, ".", -1, Microsoft.VisualBasic.CompareMethod.Text)
                    Dim Table_Name As String = name2(0).ToString



                    Dim startTime As String = Date.Now.ToString

                    ''UNCOMMENT FOR PRODUCTION
                    Dim maxVal As Integer = cntLines
                    z.Create_Table(arHeadings, arRows, Table_Name, Me.chkDisplayReport.CheckState, True, maxVal)
                    z = Nothing


                    Me.lblMessage.Text = "Records Inserted Successfully."
                    Me.Cursor = Cursors.Default

                    Application.DoEvents()

                    ''
                    '' now from here on to steps 2 3 4 
                    '' for additional scrubbing.

                    '' mappoint -> column association -> insert into our tables.
                    '' 
                    If Me.chkAutoScrubAddress.CheckState = CheckState.Checked Then
                        If Table_Name = "i360__Prospect__c" Then
                            Me.Cursor = Cursors.WaitCursor
                            Me.tblCurCnt.Text = "1"
                            Me.lblTableTotal.Text = "1"
                            Me.lblOperation.Text = "Verifying Addresses"

                            Me.lblTable.Text = Table_Name
                            Me.lblMessage.Text = "Updating Addresses . . . "
                            Application.DoEvents()
                            Dim mpt_bulk As New mappointBulkScrub
                            Dim sql_bulk As New ImportData_V2.sqlOperations
                            Dim arAddresses As List(Of ImportData_V2.sqlOperations.Record_And_Address)
                            arAddresses = sql_bulk.Get_Addresses_tblProspect(False)
                            Me.pbRows.Minimum = 0
                            Me.pbRows.Value = 0
                            Me.pbRows.Maximum = sql_bulk.CountRecords(Table_Name, False)
                            Me.lblTotalRecordNums.Text = sql_bulk.CountRecords(Table_Name, False)
                            Dim iteration As Integer = 0
                            ''UNCOMMENT FOR PRODUCTION
                            For Each x As ImportData_V2.sqlOperations.Record_And_Address In arAddresses
                                iteration += 1
                                ''If iteration >= 500 Then
                                ''    Exit For
                                ''End If
                          
                                mpt_bulk.VerifyAddress(x, False)
                                pbRows.Increment(1)
                                Me.lblCurrentRecordNum.Text = iteration
                                Application.DoEvents()
                            Next

                            Me.Cursor = Cursors.Default
                            mpt_bulk = Nothing
                            sql_bulk = Nothing
                            arAddresses = Nothing
                            Application.DoEvents()
                        End If
                    End If

                    '' now move data over to our tables: 
                    '' 

                    Dim stopTime As String = Date.Now.ToString
                    Dim whereIsReport As String = zz.CreateSingleReport(startTime, stopTime, Table_Name, cntLines)
                    zz = Nothing

                    If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                        System.Diagnostics.Process.Start(whereIsReport)
                    End If



                ElseIf Me.chkAutoInsert.CheckState = CheckState.Unchecked Then
                    Me.Cursor = Cursors.Default
                End If

            End If
        Catch ex As Exception
            Dim msg As String = ex.Message
            Me.lblMessage.Text = "FAIL - " & msg.ToString
            Dim y As New ImportData_V2.FileOperations

            Dim showReport As Boolean = False
            If Me.chkDisplayReport.CheckState = CheckState.Checked Then
                showReport = True
            ElseIf Me.chkDisplayReport.CheckState = CheckState.Unchecked Then
                showReport = False
            End If
            y.WriteErrorToLog(msg, showReport, False)
            y = Nothing
            Me.Cursor = Cursors.Default
        End Try




    End Sub



    Private Function Break_FileName_Apart(ByVal FileName As String)
        Dim parts() = Split(FileName, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim x As String = ""
        Dim cnt As Integer = 0
        For Each x In parts
            cnt += 1
        Next
        Dim ret_filename As String = parts(cnt - 1)
        Return ret_filename
    End Function

    Private Sub frmImportBulkData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.chkAllFiles.CheckState = CheckState.Checked
        Me.chkAutoInsert.CheckState = CheckState.Checked
        Me.chkDisplayReport.CheckState = CheckState.Unchecked
        Me.chkAutoScrubAddress.CheckState = CheckState.Unchecked
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '' Try
        Dim y As New ImportData
        y = Nothing
        '' Catch ex As Exception
        '' Dim x As String = ex.Message
        '' MsgBox(x, MsgBoxStyle.Critical, "DEBUG INFO")
        ''End Try
    End Sub

    Private Sub chkAllFiles_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllFiles.CheckedChanged
        Dim chk As CheckBox = sender
        If chk.CheckState = CheckState.Checked Then
            Me.cboFiles.Enabled = False
        ElseIf chk.CheckState = CheckState.Unchecked Then
            Me.cboFiles.Enabled = True
        End If
    End Sub



    Private Sub lblMessage_TextChanged(sender As Object, e As EventArgs) Handles lblMessage.TextChanged
        Form6.Label2.Text = Me.lblMessage.Text
    End Sub

    Private Sub lblCurrentRecordNum_TextChanged(sender As Object, e As EventArgs) Handles lblCurrentRecordNum.TextChanged
        Form6.lblCurrentRecordNum.Text = Me.lblCurrentRecordNum.Text
    End Sub

    Private Sub lblTotalRecordNums_TextChanged(sender As Object, e As EventArgs) Handles lblTotalRecordNums.TextChanged
        Form6.lblTotalRecordNum.Text = Me.lblTotalRecordNums.Text
    End Sub

    Private Sub tblCurCnt_TextChanged(sender As Object, e As EventArgs) Handles tblCurCnt.TextChanged
        Form6.lblTblCurCnt.Text = Me.tblCurCnt.Text
    End Sub
End Class
