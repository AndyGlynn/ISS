'' sql operations
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Imports System.Text    '' string builder
Imports System.IO      '' file operations
Imports System.Linq    '' inline sql queries 



Public Class ImportData_V2

    Private Const pro_cnx As String = "SERVER=192.168.1.2;Database=Data_Automation;User Id=sa;Password=spoken1"
    Private Const dev_cnx As String = "SERVER=PC-101\DEVMIRROREXPRESS;Database=TestMigration;User Id=sa;Password=Legend1!"

    Private Const Production_Dir As String = "\\192.168.1.2\Company\ISS\Data Automation Reports\"
    Private Const Testing_Dir As String = "C:\Users\Clay\Desktop\Print Leads\Data Automation Reports\"

    Public Class FileOperations
        ''
        '' read, write, append, create
        '' get a list of the files needed to re-create
        '' 

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


        Public Sub WriteErrorToLog(ByVal Fail_Error As String, ByVal Display As Boolean, ByVal Production As Boolean)


            Dim day As String = Date.Now.Day.ToString
            Dim month As Integer = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString

            Dim hour As Integer = Date.Now.Hour
            Dim minute As Integer = Date.Now.Minute
            Dim second As Integer = Date.Now.Second

            Dim reconstructed As String = (month & "-" & day & "-" & year & "_" & hour.ToString & minute.ToString & second.ToString)
            Dim path As String = ""
            If Production = True Then
                path = Production_Dir
            ElseIf Production = False Then
                path = Production_Dir
            End If
            Dim sw As New StreamWriter(path & "ERROR_" & reconstructed & ".txt", True)
            sw.WriteLine(Date.Now.ToString & ":> " & Fail_Error.ToString)
            sw.Flush()
            sw.Close()

            If Display = True Then
                System.Diagnostics.Process.Start(path & "ERROR_" & reconstructed & ".txt")
            ElseIf Display = False Then
                '' do nothing 
            End If

        End Sub

        Public Function Get_Files(ByVal Directory As String, ByVal SearchPattern As String)
            Dim arFile As New Hashtable
            Dim dir_NFO As New DirectoryInfo(Directory)
            For Each x As FileInfo In dir_NFO.GetFiles(SearchPattern)
                arFile.Add(x.FullName, x.Name.ToString)
            Next
            Return arFile
        End Function


        Public Function Preview_Read_File(ByVal FilePath As String)

            Dim arLines As New ArrayList
            Dim stream_reader As New StreamReader(FilePath)
            Dim text As String = ""
            Do While stream_reader.Peek() >= 0
                text = (stream_reader.ReadLine())
                arLines.Add(text)
            Loop
            stream_reader.Close()
            Return arLines
        End Function

        Public Function CreateColumnHeadings(ByVal FullDoc As ArrayList)
            Dim firstLine As String = FullDoc(0)
            Dim headings() = firstLine.Split(",")
            Dim cnt As Integer = 0
            Dim arHeadings As New ArrayList
            For Each s As String In headings
                arHeadings.Add(Replace(s, Chr(34), "", 1, s.ToString.Length - 1))
            Next
            Return arHeadings
        End Function

        Public Function CreateLineScrubbed(ByVal Line As String)

            Dim searchSTR = (Chr(34) & Chr(44) & Chr(34)) '' "," 
            ''
            '' this pattern FAILS when dealing with 
            '' JSON objects stored in a field.
            '' will offset the column headers and array count of lines due to finding ' "," ' multiple times.
            ''
            '' solution 1) delete values in a CSV editor before sending through program
            ''             not good as is a MANUAL solution -> looking for a FULL Auto Solution
            '' solution 2) drop the column from table creation
            ''             not ambiguous enough to deal with variety of files, tailored ONLY for i360 Export / Import
            ''
            Dim arSplit = Split(Line, searchSTR, -1, Microsoft.VisualBasic.CompareMethod.Text)
            Return arSplit

        End Function


        Public Function CreateBatchReport(ByVal StartTime As String, ByVal StopTime As String, ByVal Table_Names As ArrayList)


            Dim date_ As Date = Date.Now

            Dim hour As Integer = date_.Hour
            Dim minute As Integer = date_.Minute
            Dim second As Integer = date_.Second
            Dim milSec As Integer = date_.Millisecond
            Dim fileTimeFormat As String = (hour & "_" & minute & "_" & second & "_" & milSec)

            Dim year As Integer = date_.Year
            Dim month As Integer = date_.Month
            Dim day As Integer = date_.Day
            Dim fileDateFormat As String = (year & "_" & day & "_" & month)

            Dim strFileName As String = (Production_Dir & "Batch Report_" & fileDateFormat & "_" & fileTimeFormat & ".txt")

            Dim sr As StreamWriter

            Try
                If System.IO.File.Exists(strFileName) = True Then
                    System.IO.File.Delete(strFileName)
                ElseIf System.IO.File.Exists(strFileName) = False Then
                    sr = New StreamWriter(strFileName, True)
                End If
            Catch ex As Exception
                Dim msg As String = ex.Message
                MsgBox("Batch Report Error" & vbCrLf & "ERROR:" & msg.ToString, MsgBoxStyle.Critical, "DEBUG INFO - BATCH REPORT")
            End Try

            sr.WriteLine("Batch Started @:> " & StartTime.ToString)
            For Each y As String In Table_Names
                sr.WriteLine("Table: [" & y.ToString & "] created.")
            Next
            sr.WriteLine("Batch Ended @:> " & StopTime.ToString)
            sr.Flush()
            sr.Close()

            Return strFileName

        End Function
        Public Function CreateSingleReport(ByVal StartTime As String, ByVal StopTime As String, ByVal Table_Name As String, ByVal NumRows As String)
            Dim date_ As Date = Date.Now

            Dim hour As Integer = date_.Hour
            Dim minute As Integer = date_.Minute
            Dim second As Integer = date_.Second
            Dim milSec As Integer = date_.Millisecond
            Dim fileTimeFormat As String = (hour & "_" & minute & "_" & second & "_" & milSec)

            Dim year As Integer = date_.Year
            Dim month As Integer = date_.Month
            Dim day As Integer = date_.Day
            Dim fileDateFormat As String = (year & "_" & day & "_" & month)

            Dim strFileName As String = (Production_Dir & "Report_" & fileDateFormat & "_" & fileTimeFormat & ".txt")

            Dim sr As StreamWriter

            Try
                If System.IO.File.Exists(strFileName) = True Then
                    System.IO.File.Delete(strFileName)
                ElseIf System.IO.File.Exists(strFileName) = False Then
                    sr = New StreamWriter(strFileName, True)
                End If
            Catch ex As Exception
                Dim msg As String = ex.Message
                MsgBox("Single Table Report Error" & vbCrLf & "ERROR:" & msg.ToString, MsgBoxStyle.Critical, "DEBUG INFO - Single REPORT")
            End Try




            sr.WriteLine("Job Started @:> " & StartTime.ToString)
            sr.WriteLine("Table : [ " & Table_Name & " ] created with [ " & NumRows & " ] records inserted to database :> [ Data_Automation ] .")
            sr.WriteLine("Job Ended @:> " & StopTime.ToString)
            sr.Flush()
            sr.Close()

            Return strFileName



        End Function

    End Class


    Public Class sqlOperations
        '' queries, updates, creates
        Private Const dev_cnx As String = "SERVER=PC-101\DEVMIRROREXPRESS;Database=TestMigration;User Id=sa;Password=Legend1!"
        Private Const cnx_string As String = "SERVER=192.168.1.2;Database=Data_Automation;User Id=sa;Password=spoken1"
        Private Const Testing_Dir As String = "C:\Users\Clay\Desktop\Print Leads\Data Automation Reports\"
        Private Const Production_Dir As String = "\\192.168.1.2\Company\ISS\Data Automation Reports\"


        Public Structure Record_And_Address
            Public Id_FromTable As String
            Public StreetAddress As String
            Public City As String
            Public State As String
            Public Zip As String
        End Structure

        Public Sub Update_Table(ByVal Obj As sqlOperations.Record_And_Address, ByVal Dev_Or_Production As Boolean)
            Dim cnx As SqlConnection
            If Dev_Or_Production = True Then
                cnx = New SqlConnection(pro_cnx)
            ElseIf Dev_Or_Production = False Then
                cnx = New SqlConnection(pro_cnx)
            End If
            '' i360__Home__Address__c, i360__Home_City__c, i360__Home_State__c, i360__Home_Zip_Postal_Code__c"
            cnx.Open()
            Dim strUP As String = ("UPDATE i360__Prospect__c") & vbCrLf
            strUP += ("SET i360__Home_Address__c = N'" & Obj.StreetAddress & "',") & vbCrLf
            strUP += ("    i360__Home_City__c = N'" & Obj.City & "',") & vbCrLf
            strUP += ("    i360__Home_State__c = N'" & Obj.State & "',") & vbCrLf
            strUP += ("    i360__Home_Zip_Postal_Code__c = N'" & Obj.Zip & "'") & vbCrLf
            strUP += ("WHERE Id = N'" & Obj.Id_FromTable & "';")
            Dim cmdUP As New SqlCommand(strUP, cnx)

            cmdUP.ExecuteScalar()
            cnx.Close()
            cnx = Nothing
        End Sub

        Public Sub CreateTestTable()
            Dim cnx As New SqlConnection(cnx_string)
            cnx.Open()
            Dim cmdCREATE As New SqlCommand("CREATE TABLE Test (ID INT NOT NULL PRIMARY KEY,FName varchar(max));", cnx)
            cmdCREATE.CommandType = CommandType.Text
            cmdCREATE.ExecuteNonQuery()
            cnx.Close()
            cnx = Nothing

        End Sub

        Public Sub Create_Table(ByVal colHeadings As ArrayList, ByVal ItemLines As ArrayList, ByVal TableName As String, ByVal DisplayReportWhenDone As CheckState, ByVal Production As Boolean, ByVal maxVal As Integer)

            Dim guid As Guid = guid.NewGuid
            frmImportBulkData.lblTotalRecordNums.Text = (maxVal - 1).ToString
            frmImportBulkData.lblTable.Text = TableName

            Form6.Label2.Text = "Creating Temp Table " & frmImportBulkData.tblCurCnt.ToString & " of 8"
            ''
            '' if table already exists, drop and create new one.
            '' 

            '' first check to make sure table name is not a reserved sql keyword
            '' 
            Dim t_name As String = CheckForKeywords(TableName)

            If Production = True Then
                Using cnx As New SqlConnection(cnx_string)
                    Dim dropStr As String = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & t_name & "]') AND TYPE IN (N'U')) DROP TABLE [dbo].[" & t_name & "];"
                    Dim cmdDROP As New SqlCommand(dropStr, cnx)
                    cnx.Open()
                    Dim res As Integer = cmdDROP.ExecuteNonQuery
                    cnx.Close()
                End Using

            ElseIf Production = False Then
                Using cnx As New SqlConnection(cnx_string)
                    Dim dropStr As String = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & t_name & "]') AND TYPE IN (N'U')) DROP TABLE [dbo].[" & t_name & "];"
                    Dim cmdDROP As New SqlCommand(dropStr, cnx)
                    cnx.Open()
                    Dim res As Integer = cmdDROP.ExecuteNonQuery
                    cnx.Close()
                End Using
            End If

            Dim strCreate As String = ""
            Dim cnt As Integer = 0
            Dim arRowInsert As New ArrayList
            strCreate += "CREATE TABLE " & t_name & "" & vbCrLf
            strCreate += "( " & vbCrLf
            For Each s As String In colHeadings
                cnt += 1
            Next
            Dim i As Integer = 0
            For i = 0 To cnt - 1
                If i <> (cnt - 1) Then
                    '' check for reserved keywords
                    '' 
                    Dim strCheck As String = colHeadings(i).ToString
                    Dim checkVal As String = CheckForKeywords(strCheck)
                    strCreate += "" & checkVal & " varchar(750)," & vbCrLf
                ElseIf i = (cnt - 1) Then
                    Dim strCheck As String = colHeadings(i).ToString
                    Dim checkVal As String = CheckForKeywords(strCheck)
                    strCreate += "" & checkVal & " varchar(750)" & vbCrLf
                End If
            Next
            strCreate += ");"


            ''uncomment to write to table
            '' 
            If Production = True Then
                Using cnx2 As New SqlConnection(cnx_string)
                    cnx2.Open()
                    Dim cmdINS As New SqlCommand(strCreate, cnx2)
                    cmdINS.ExecuteNonQuery()
                    cnx2.Close()
                End Using
            ElseIf Production = False Then
                Using cnx2 As New SqlConnection(cnx_string)
                    cnx2.Open()
                    Dim cmdINS As New SqlCommand(strCreate, cnx2)
                    cmdINS.ExecuteNonQuery()
                    cnx2.Close()
                End Using
            End If


            '' uncomment ONLY if you want to see the structure of the table in the report.
            ''sr.WriteLine(strCreate & vbCrLf & vbCrLf & vbCrLf)


           
            Dim z As Integer = 0
            For z = 0 To ItemLines.Count - 1
                Dim strInsert As String = "INSERT " & t_name & "("
                For i = 0 To cnt - 1
                    If i <> (cnt - 1) Then
                        strInsert += colHeadings(i) & ","
                    ElseIf i = (cnt - 1) Then
                        strInsert += colHeadings(i)
                    End If
                Next
                strInsert += ") values ('"

                Dim arVals As Array = ItemLines(z)
                Dim aa As Integer = 0
                For aa = 0 To arVals.Length - 1
                    If aa <> (arVals.Length - 1) Then
                        If arVals(aa) Is Nothing Then
                            arVals(aa) = ""
                            strInsert += arVals(aa).ToString & "','"
                        ElseIf arVals(aa) IsNot Nothing Then
                            Dim strINS As String = Replace(arVals(aa).ToString, Chr(39), " ", 1, arVals(aa).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            strINS = CheckForKeywords(strINS)
                            '' check here for keywords that will fail statement
                            '' ie: 'Key' for a column name in 'insert' statement
                            strInsert += (strINS & "','")
                        End If
                    ElseIf aa = (arVals.Length - 1) Then
                        If arVals(aa) Is Nothing Then
                            arVals(aa) = ""
                            strInsert += arVals(aa).ToString & "');"
                        ElseIf arVals(aa) IsNot Nothing Then
                            Dim strINS As String = Replace(arVals(aa).ToString, Chr(39), " ", 1, arVals(aa).ToString.Length - 1, Microsoft.VisualBasic.CompareMethod.Text)
                            '' check here for keywords that will fail statement
                            '' ie: 'Key' for a column name in 'insert' statement
                            strINS = CheckForKeywords(strINS)
                            strInsert += (strINS & "');")
                        End If
                    End If
                Next


                arRowInsert.Add(strInsert)
                strInsert = ""
            Next

            Dim testString As String = ""
            Dim cntRows As Integer = 0
            Form6.Label2.Text = "Temp Table " & frmImportBulkData.tblCurCnt.Text & " of 8 Created Successfully"

            Form6.ProgressBar1.Maximum = Form6.pbRowMax
            Form6.ProgressBar1.Value = 0
            Form6.Label2.Text = "Inserting Records To Temp Table " & frmImportBulkData.tblCurCnt.Text & " of 8"
            Form6.lblTotalRecordNum.Visible = True
            Form6.lblCurrentRecordNum.Visible = True
            Form6.lblTblCurCnt.Visible = True
            Form6.lblTable.Visible = True
            For Each zzz As String In arRowInsert
                cntRows += 1
              
                '' uncomment to write to table 
                If Production = True Then
                    Using cnx3 As New SqlConnection(cnx_string)
                        cnx3.Open()
                        Dim cmdINSR As New SqlCommand(zzz, cnx3)
                        cmdINSR.ExecuteNonQuery()
                        cnx3.Close()
                        'testString += zzz & vbCrLf
                        frmImportBulkData.pbRows.Increment(1)
                        Form6.ProgressBar1.PerformStep()

                        frmImportBulkData.lblCurrentRecordNum.Text = cntRows.ToString
                    
                        Application.DoEvents()
                    End Using
                ElseIf Production = False Then
                    Using cnx3 As New SqlConnection(cnx_string)
                        cnx3.Open()
                        Dim cmdINSR As New SqlCommand(zzz, cnx3)
                        cmdINSR.ExecuteNonQuery()
                        cnx3.Close()
                        'testString += zzz & vbCrLf
                        frmImportBulkData.pbRows.Increment(1)
                        Form6.ProgressBar1.PerformStep()
                        frmImportBulkData.lblCurrentRecordNum.Text = cntRows.ToString

                        Application.DoEvents()
                    End Using
                End If
            Next
            Form6.Label2.Text = "Insert Records Complete for Temp Table " & frmImportBulkData.tblCurCnt.Text & " of 8"
            'sr.WriteLine(testString & vbCrLf & vbCrLf) '' un-comment for verbose information only | Show insert sql string WITH values inserted.
            Form6.lblTotalRecordNum.Visible = False
            Form6.lblCurrentRecordNum.Visible = False
            Form6.lblTblCurCnt.Visible = False
            Form6.lblTable.Visible = False
            
            ItemLines.Clear()
            colHeadings.Clear()



            Form1.Cursor = Cursors.Default

            If DisplayReportWhenDone = CheckState.Checked Then

            ElseIf DisplayReportWhenDone = CheckState.Unchecked Then

            End If


        End Sub

        Public Function Get_Addresses_tblProspect(ByVal Dev_Or_Production As Boolean)

            Dim arRecords As New List(Of Record_And_Address)
            Dim cnx As SqlConnection
            ''
            '' key:=Id_FromTable|value:=Record_And_Address
            ''
            '' both the structure and the key of the hash entry will house the id from the table
            '' 
            If Dev_Or_Production = True Then
                cnx = New SqlConnection(pro_cnx)
            ElseIf Dev_Or_Production = False Then
                cnx = New SqlConnection(pro_cnx)
            End If

            cnx.Open()
            Dim cmdGET As New SqlCommand("SELECT Id,i360__Home_Address__c,i360__Home_City__c,i360__Home_State__c,i360__Home_Zip_Postal_Code__c FROM i360__Prospect__c;", cnx)
            Dim r1 As SqlDataReader = cmdGET.ExecuteReader
            While r1.Read
                Dim x As New Record_And_Address
                x.Id_FromTable = r1.Item("Id")
                x.StreetAddress = r1.Item("i360__Home_Address__c")
                x.City = r1.Item("i360__Home_City__c")
                x.State = r1.Item("i360__Home_State__c")
                x.Zip = r1.Item("i360__Home_Zip_Postal_Code__c")
                arRecords.Add(x)
                'Form1.pbGenerate.Increment(1)
                ' Application.DoEvents()
            End While
            r1.Close()
            cnx.Close()
            cnx = Nothing
            Return arRecords

        End Function

        Public Function GetTables()
            Dim cnx As New SqlConnection(cnx_string)
            cnx.Open()
            Dim arTables As New ArrayList
            Dim cmdGET As New SqlCommand("SELECT * FROM sys.tables;", cnx)
            Dim r1 As SqlDataReader = cmdGET.ExecuteReader
            While r1.Read
                arTables.Add(r1.Item("name") & " " & r1.Item("object_id"))
            End While
            r1.Close()
            cnx.Close()
            Return arTables
        End Function

        Private Function CheckForKeywords(ByVal Val_To_Check As String)
            Dim retSTR As String = Val_To_Check
            Select Case retSTR
                Case Is = "Key"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Primary_Key"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Auto_Increment"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Int"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Foreign_Key"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Date"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Coalesce"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Count"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Not"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Null"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Declare"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "User"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "For"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Select"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Identity"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Update"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Delete"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Insert"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "If"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Where"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "*"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "%"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Case"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Drop"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "In"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Join"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "And"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Or"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Between"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Alter"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Union"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Is = "Create"
                    retSTR = ("_" & retSTR)
                    Exit Select
                Case Else
                    retSTR = retSTR
                    Exit Select
            End Select
            Return retSTR

        End Function

        Public Function CountRecords(ByVal TableName As String, ByVal DEV_or_Production As Boolean)
            Dim cnt As Integer = 0
            Dim cnx As SqlConnection
            If DEV_or_Production = True Then
                cnx = New SqlConnection(pro_cnx)
            ElseIf DEV_or_Production = False Then
                cnx = New SqlConnection(pro_cnx)
            End If
            cnx.Open()
            Dim cmdCNT As New SqlCommand("SELECT COUNT(Id) FROM [dbo].[" & TableName & "];", cnx)
            cnt = cmdCNT.ExecuteScalar
            cnx.Close()
            cnx = Nothing
            Return cnt
        End Function

 
    End Class


    Public Class objectOperations
        '' convertion things to structures 
    End Class
    







End Class
