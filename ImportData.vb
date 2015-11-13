Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports MapPoint

Public Class ImportData
    Public cnnISS As New SqlConnection("SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1;")
    Public cnnDA As New SqlConnection("SERVER=192.168.1.2;Database=Data_Automation;User Id=sa;Password=spoken1;")
    Public cnnDA2 As New SqlConnection("SERVER=192.168.1.2;Database=Data_Automation;User Id=sa;Password=spoken1;")
    Public CorrectPhoneNumberResponse As String = ""

    Private Const da_cnx As String = "SERVER=192.168.1.2;Database=Data_Automation;User Id=sa;Password=spoken1;"
    Private Const iss_cnx As String = "SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1;"

#Region "Notes 11-3-15 PRE COMBINE - RE ORG"

    ''
    '' Need a way to manage sql connections better.
    ''     cnx pooling
    '' Andy's code is opening a connection 21x per 'scrub'.
    '' on a batch with 44k records, this will be approx: 924,000 times a sql cnx will be opened.
    '' PER BATCH @44K | multiple batches will be more. 
    '' 
    '' going to have to watch mem usage on this too.
    '' manually clear out , or design with IDisposable support 
    ''   for faster garbage collection / mem reclamation
    '' mappoint takes up 750mb's (+/- 50mb's) by itself when running a batch (@44k record entries).
    '' prospect__history from i360 took 1.2gb (+/- .2gb) on my machine to process (1batch @640K+ entries)
    ''
    '' ** multi class / object support (?)  **
    '' ** Typed Data / Structs **
    '' ** use key values, hash's, dictionaries, tuples, and/or string builders for repetitous string concatenation where applicable. **
    '' 
    '' small, easily disposed of pieces. if it's use is over, kill / destroy / clear it.
    ''

    '' move as much over to sql [SP's / triggers / etc.] as possible. Share the load(?) 
    '' 

    ''
    '' build proto first thing in AM then tweak once created.
    '' write 1 (one) new class. 
    ''
    '' BASIC FLOW 
    '' 
    '' 1) import data
    ''   a) browse to data source
    ''   b) process all files listed ? | Yes.
    ''   c) For each file listed |  load to mem
    ''   d) iterate through 'processing' file / inserting data (AUTO INSERT) 
    ''   
    ''   break apart CSV (only - scrap XML and JSON support | IMPORT, not necessarily on export side. To be determined.) files and create tables / columns / fields 
    ''   after tables are created, insert the data for each record 
    ''


    '' 2) scrub data - [FIRST TIME] 
    ''   a) reset counters for pb's
    ''      i) load to mem
    ''   b) wait until all leads are 'scrubbed'
    ''
    '' 3) mappoint scrub data - [SECOND TIME] 
    ''   a) reset counters fors pb's
    ''      i) load to mem
    ''   b) wait until all address as scrubbed
    ''   c) no need to write 'FAILS' to table per Andy (about 20% of leads - On the initial
    ''
    '' 4) show / don't show report. [OPTIONAL]
    ''   a) if the "show report" is checked / show the final outcome 
    ''   b) ONE (1) streamwriter for this report. No need for individual reports per table. 
    '' 

    ''
    '' double check Lambda funcs and LINq supported in 4.0 framework (target FW) 
    '' ** may or may not use / need for optimization
    '' 

    '' INITIAL
    '' 
    '' database | exist yes=drop|no=create
    '' set  context to created database
    ''
    '' build cnx string
    '' 
    '' table operations: 
    '' 
    '' tables | exist yes=drop|no=create
    ''
    '' SP / TRIGGER operations: 
    '' 
    '' SP | exist yes=use it | no=create then use
    ''
    '' bulk insert records via SP / Triggers
    ''
    '' NEXT STEP=> 

    

#End Region

    Public Sub New()
        Dim Marketer As String
        Dim PLS As String
        Dim Generated As String
        Dim PFirstName As String
        Dim PLastName As String
        Dim SFirstName As String
        Dim SLastName As String
        Dim Addy As String
        Dim City As String
        Dim State As String
        Dim Zip As String
        Dim HPhone As String
        Dim Alt1 As String
        Dim Alt1Type As String
        Dim Alt2 As String
        Dim Alt2Type As String
        Dim P1 As String
        Dim P2 As String
        Dim P3 As String
        Dim YearsOwned As String = ""
        Dim HomeAge As String = ""
        Dim HomeVal As String = ""
        Dim ApptDate As Date
        Dim ApptTime As Date
        Dim SI As String
        Dim Lat As String
        Dim Lon As String
        Dim Rep1 As String = ""
        Dim Rep2 As String = ""
        Dim Result As String
        Dim QuotePar As Double = 0
        Dim Recover As Boolean
        Dim MNotes As String = ""
        Dim Cash As Boolean
        Dim Finance As Boolean
        Dim P1Split As Double
        Dim P2Split As Double
        Dim P3Split As Double
        Dim MResult As String
        Dim Confirmer As String = ""
        Dim DNM As Boolean
        Dim DNC As Boolean
        Dim Email As String
        Dim MManager As String = "Andy Stumph"
        Dim SManager As String = ""
        Dim id As String
        Dim Confirmed As String
        Dim Kill As String
        Dim KillReason As String
        Dim USR As String
        Dim ApptType As String
        Dim Day As String
        Dim Mapped As Boolean = False
        Dim ResultDetail As String
        Dim RCorBR As String
        Dim canceldate As Date
        Dim EditMode As Boolean = False
        Dim Description As String
        Dim SkipConfirm As Boolean = False



        ''Add Your Import CVS Code Here or Use a Seperate Class, Your Preference. 
        ''Loop Through these Tables
        ''- i360__Appointment__c
        ''- i360__Prospect__c
        ''- i360__Sale__c
        ''- _User
        ''- i360__Staff__c
        ''- i360__Marketing_Source__c
        ''- i360__Marketing_Task__c
        ''- i360__Lead_Source__c

        Form6.Label2.Text = "Preparing Temp Tables For Import"




        Using cnx As New SqlConnection(da_cnx)
            Dim cmdScrub As SqlCommand = New SqlCommand("dbo.Scrub_Tables", cnx)
            cmdScrub.CommandType = CommandType.StoredProcedure
            Dim r11 As SqlDataReader
            cnx.Open()
            r11 = cmdScrub.ExecuteReader
            r11.Read()
            r11.Close()
            cnx.Close()
        End Using

        Using cnx As New SqlConnection(da_cnx)
            Dim Cnt As SqlCommand = New SqlCommand("dbo.CountProspects", cnx)
            Cnt.CommandType = CommandType.StoredProcedure
            cnx.Open()
            Dim c As SqlDataReader
            c = Cnt.ExecuteReader
            c.Read()
            '' junk 11-3-2015 AC
            '' no code behind.....
            '' write new form / front end
            'Form6.ProgressBar1.Maximum = c.Item(0)   ''Resets progress Bar for Number of Loops in my below While Statement 
            frmImportBulkData.pbRows.Minimum = 0
            frmImportBulkData.pbRows.Maximum = c.Item(0)
            frmImportBulkData.lblTotalRecordNums.Text = c.Item(0)
            frmImportBulkData.pbRows.Value = 0
            Form6.ProgressBar1.Value = 0
            Form6.ProgressBar1.Maximum = c.Item(0)
            cnx.Close()
        End Using
        Form6.Label2.Text = "Temp Tables Prepared for Import"

        Form6.Label2.Text = "Importing Data Into " & STATIC_VARIABLES.ProgramName
        Dim cnx3 As New SqlConnection(da_cnx)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetProspects", cnx3)
        cmdGet.CommandType = CommandType.StoredProcedure
        cnx3.Open()
        Dim r1 As SqlDataReader

        r1 = cmdGet.ExecuteReader
        Form6.lblTable.Text = "Importing Record                   of                   Into " & STATIC_VARIABLES.ProgramName
        Form6.lblTblCurCnt.Text = ""
        Form6.lblTable.Visible = True
        'Form6.lblTblCurCnt.Visible = True
        Form6.lblCurrentRecordNum.Visible = True
        Form6.lblTotalRecordNum.Visible = True
        Dim iteration As Integer = 0
        frmImportBulkData.lblOperation.Text = "Moving Data Around"
        While r1.Read

            iteration += 1
            ''Write Info From Prospect Table into Enterlead 
            frmImportBulkData.pbRows.Increment(1)
            Form6.ProgressBar1.PerformStep()
            'frmImportBulkData.lblMessage.Text = "Moving Lead (s) . . . Current: " & r1.Item(0).ToString
            frmImportBulkData.lblCurrentRecordNum.Text = iteration.ToString
            'Application.DoEvents()


            id = r1.Item(0)
            PFirstName = r1.Item(1)
            PLastName = r1.Item(2)
            SFirstName = r1.Item(3)
            SLastName = r1.Item(4)
            Addy = r1.Item(5)
            City = r1.Item(6)
            State = r1.Item(7)
            Zip = r1.Item(8)
            HPhone = r1.Item(9)
            Alt1 = r1.Item(10)
            Alt1Type = r1.Item(11)
            Alt2 = r1.Item(12)
            Alt2Type = r1.Item(13)
            YearsOwned = r1.Item(14)
            HomeAge = r1.Item(15)
            HomeVal = r1.Item(16)
            Lat = r1.Item(17)
            Lon = r1.Item(18)
            DNM = r1.Item(19)
            DNC = r1.Item(20)
            Email = r1.Item(21)
            Kill = r1.Item(22)
            KillReason = r1.Item(23)

            If HPhone = Alt1 And HPhone = Alt2 Then
                Alt1 = ""
                Alt2 = ""
            ElseIf HPhone <> Alt1 And Alt1 = Alt2 Then
                Alt2 = ""
            End If

            ''Query Appt Table to complete data on Enterlead and iterate through all appointments to write lead history & fire tblWhereCanLeadGo
            Dim cnx2 As New SqlConnection(da_cnx)
            Dim cmdApp As SqlCommand = New SqlCommand("dbo.GetAppts", cnx2)
            Dim s1 As SqlParameter = New SqlParameter("@id", id)
            cmdApp.CommandType = CommandType.StoredProcedure
            cnx2.Open()
            cmdApp.Parameters.Add(s1)
            Dim r2 As SqlDataReader
            r2 = cmdApp.ExecuteReader
            Dim counter As Integer = 0
            While r2.Read
                ''sql to look for appointments for this prospect 
                ''update enterlead data 


                Marketer = r2.Item(1)
                PLS = r2.Item(2)
                Generated = r2.Item(3)
                P1 = r2.Item(4)
                P2 = r2.Item(5)
                P3 = r2.Item(6)
                ApptDate = r2.Item(7)
                ApptTime = r2.Item(8)
                SI = r2.Item(9)
                Rep1 = r2.Item(10)
                Rep2 = r2.Item(11)
                Result = r2.Item(12)

                Try
                    QuotePar = CType(r2.Item(13), Double)
                Catch ex As Exception
                    QuotePar = 0
                End Try

                Confirmed = r2.Item(14)
                Confirmer = r2.Item(15)
                SManager = r2.Item(17)
                ApptType = r2.Item(18)
                ResultDetail = r2.Item(21)


                Try
                    If r2.Item(19) = "Rehash Appointment" Then
                        Recover = True
                    Else
                        Recover = False
                    End If
                Catch ex As Exception
                    Recover = False
                End Try
                Try
                    If ResultDetail <> "" Then
                        MNotes = r2.Item(20) & " - " & ResultDetail
                    Else
                        MNotes = r2.Item(20)
                    End If
                Catch ex As Exception
                    MNotes = ResultDetail
                End Try
                If ApptType = "Rehash" Then
                    PLS = "Recovery"
                End If

                Try
                    If r2.Item(23) = "Financing" Then
                        Finance = True
                        Cash = False
                    ElseIf r2.Item(23) = "Cash" Or r2.Item(23) = "" Then
                        Cash = True
                        Finance = False
                    End If
                Catch ex As Exception
                    Cash = False
                    Finance = False
                End Try
                Try
                    If r2.Item(22) = "Bank Rejected" Or r2.Item(22) = "Recission Cancel" Then
                        RCorBR = r2.Item(22)
                    Else
                        RCorBR = ""
                    End If
                Catch ex As Exception
                    RCorBR = ""
                End Try
                Try
                    canceldate = r2.Item(24)
                Catch ex As Exception
                    canceldate = "1/1/1900 12:00:00"
                End Try




                'Dim ID

                If Marketer = "" Then
                    Description = "Appt. Generated from " & PLS & ", set up with " & PFirstName
                ElseIf Marketer <> "" Then
                    Description = "Appt. Generated from " & PLS & " by " & Marketer & ", set up with " & PFirstName
                End If
                If counter = 0 Then
                    Dim cnxISS As New SqlConnection(iss_cnx)
                    Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertEnterLead", cnxISS)
                    cmdINS.CommandType = CommandType.StoredProcedure
                    Dim param1 As SqlParameter = New SqlParameter("@Marketer", Trim(Marketer))
                    Dim param2 As SqlParameter = New SqlParameter("@PLS", Trim(PLS))
                    Dim param3 As SqlParameter = New SqlParameter("@SLS", Trim(""))
                    Dim param4 As SqlParameter = New SqlParameter("@LeadGenOn", Generated)
                    Dim param5 As SqlParameter = New SqlParameter("@Contact1FirstName", Trim(PFirstName))
                    Dim param6 As SqlParameter = New SqlParameter("@Contact1LastName", Trim(PLastName))
                    Dim param7 As SqlParameter = New SqlParameter("@Contact2FirstName", Trim(SFirstName))
                    Dim param69 As SqlParameter = New SqlParameter("@Contact2LastName", Trim(SLastName))
                    Dim param8 As SqlParameter = New SqlParameter("@YearsOwned", YearsOwned)
                    Dim param9 As SqlParameter = New SqlParameter("@HomeAge", Trim(HomeAge))
                    Dim param10 As SqlParameter = New SqlParameter("@HomeValue", HomeVal)
                    Dim param11 As SqlParameter = New SqlParameter("@SpecialInstruction", SI)

                    '' strip off literals
                    '' '(' ')' '-'
                    '' -------------------------------------------------------------------      
                    DoLiteralsExist(HPhone)
                    Dim param12 As SqlParameter = New SqlParameter("@HousePhone", Me.CorrectPhoneNumberResponse.ToString)

                    DoLiteralsExist(Alt1)
                    Dim param13 As SqlParameter = New SqlParameter("@AltPhone1", Me.CorrectPhoneNumberResponse.ToString)

                    DoLiteralsExist(Alt2)
                    Dim param14 As SqlParameter = New SqlParameter("@AltPhone2", Me.CorrectPhoneNumberResponse.ToString)
                    '' -----------------------------------------------------------------------------------

                    Dim param15 As SqlParameter = New SqlParameter("@Alt1Type", Alt1Type)
                    Dim param16 As SqlParameter = New SqlParameter("@Alt2Type", Alt2Type)
                    'Dim param17 As SqlParameter = New SqlParameter("@StNumber", EnterLead.txtStNum.Text)
                    Dim stAddress As String = ""
                    Dim param18 As SqlParameter = New SqlParameter("@StAddress", Trim(Addy))
                    Dim param19 As SqlParameter = New SqlParameter("@City", Trim(City))
                    Dim param20 As SqlParameter = New SqlParameter("@State", State)
                    Dim param21 As SqlParameter = New SqlParameter("@Zip", Zip)
                    Dim param22 As SqlParameter = New SqlParameter("@SpokeWith", PFirstName)
                    Dim param23 As SqlParameter = New SqlParameter("@C1Work", "Unknown")
                    Dim param24 As SqlParameter = New SqlParameter("@C2Work", "")
                    Dim param25 As SqlParameter = New SqlParameter("@AppDate", ApptDate)

                    Dim d1 = ApptDate.DayOfWeek
                    Select Case d1
                        Case Is = 0
                            Day = "Sunday"
                            Exit Select
                        Case Is = 1
                            Day = "Monday"
                            Exit Select
                        Case Is = 2
                            Day = "Tuesday"
                            Exit Select
                        Case Is = 3
                            Day = "Wednesday"
                            Exit Select
                        Case Is = 4
                            Day = "Thursday"
                            Exit Select
                        Case Is = 5
                            Day = "Friday"
                            Exit Select
                        Case Is = 6
                            Day = "Saturday"
                            Exit Select
                    End Select

                    Dim param26 As SqlParameter = New SqlParameter("@AppDay", Day)
                    Dim param27 As SqlParameter = New SqlParameter("@AppTime", ApptTime.ToShortTimeString)
                    Dim param28 As SqlParameter = New SqlParameter("@Product1", P1)
                    Dim param29 As SqlParameter = New SqlParameter("@Product2", P2)
                    Dim param30 As SqlParameter = New SqlParameter("@Product3", P3)
                    Dim param31 As SqlParameter = New SqlParameter("@Color", "")
                    Dim param32 As SqlParameter = New SqlParameter("@QTY", "")
                    Dim param33 As SqlParameter = New SqlParameter("@Email", Email)
                    Dim param34 As SqlParameter = New SqlParameter("@ProdAcro1", "")
                    Dim param35 As SqlParameter = New SqlParameter("@ProdAcro2", "")
                    Dim param36 As SqlParameter = New SqlParameter("@prodAcro3", "")
                    Dim param37 As SqlParameter = New SqlParameter("@User", "Imported Data")
                    Dim param38 As SqlParameter = New SqlParameter("@MarketingManager", MManager)
                    Dim param39 As SqlParameter = New SqlParameter("@Description", Description)
                    Dim param70 As SqlParameter = New SqlParameter("@Mapped", Mapped)
                    cnxISS.Open()
                    cmdINS.Parameters.Add(param1)
                    cmdINS.Parameters.Add(param2)
                    cmdINS.Parameters.Add(param3)
                    cmdINS.Parameters.Add(param4)
                    cmdINS.Parameters.Add(param5)
                    cmdINS.Parameters.Add(param6)
                    cmdINS.Parameters.Add(param7)
                    cmdINS.Parameters.Add(param8)
                    cmdINS.Parameters.Add(param9)
                    cmdINS.Parameters.Add(param10)
                    cmdINS.Parameters.Add(param11)
                    cmdINS.Parameters.Add(param12)
                    cmdINS.Parameters.Add(param13)
                    cmdINS.Parameters.Add(param14)
                    cmdINS.Parameters.Add(param15)
                    cmdINS.Parameters.Add(param16)
                    'cmdINS.Parameters.Add(param17)
                    cmdINS.Parameters.Add(param18)
                    cmdINS.Parameters.Add(param19)
                    cmdINS.Parameters.Add(param20)
                    cmdINS.Parameters.Add(param21)
                    cmdINS.Parameters.Add(param22)
                    cmdINS.Parameters.Add(param23)
                    cmdINS.Parameters.Add(param24)
                    cmdINS.Parameters.Add(param25)
                    cmdINS.Parameters.Add(param26)
                    cmdINS.Parameters.Add(param27)
                    cmdINS.Parameters.Add(param28)
                    cmdINS.Parameters.Add(param29)
                    cmdINS.Parameters.Add(param30)
                    cmdINS.Parameters.Add(param31)
                    cmdINS.Parameters.Add(param32)
                    cmdINS.Parameters.Add(param69)
                    cmdINS.Parameters.Add(param33)
                    cmdINS.Parameters.Add(param34)
                    cmdINS.Parameters.Add(param35)
                    cmdINS.Parameters.Add(param36)
                    cmdINS.Parameters.Add(param37)
                    cmdINS.Parameters.Add(param38)
                    cmdINS.Parameters.Add(param39)
                    cmdINS.Parameters.Add(param70)
                    Dim INS As SqlDataReader
                    INS = cmdINS.ExecuteReader
                    INS.Close()
                    cnxISS.Close()
                    cnxISS = Nothing
                End If
                Dim EId As Integer

                Dim cnn1 As SqlConnection = New SqlConnection(iss_cnx)
                Dim cmdId As SqlCommand = New SqlCommand("select top 1 (id), marketingresults, isrecovery from enterlead order by id desc", cnn1)
                cnn1.Open()
                Dim getID As SqlDataReader
                getID = cmdId.ExecuteReader
                getID.Read()
                Dim WriteSRsult As Boolean
                If Result = "Called and Cancelled" Then
                    WriteSRsult = False
                ElseIf Result = "" Then
                    WriteSRsult = False
                ElseIf Result = "Not Confirmed" Then
                    WriteSRsult = False
                Else
                    WriteSRsult = True
                End If
                EId = getID.Item(0)
                Dim lastresult As String
                Dim isRecovery As Boolean = getID.Item(2)
                Try
                    lastresult = getID.Item(1)
                Catch ex As Exception
                    lastresult = ""
                End Try
                Dim lrboo As Boolean = False
                If lastresult = "Reset" Then
                    lrboo = True
                ElseIf lastresult = "Not Hit" Then
                    lrboo = True
                ElseIf lastresult = "Not Issued" Then
                    lrboo = True
                ElseIf lastresult = "Unconfirmed" Then
                    lrboo = True
                End If
                getID.Close()
                cnn1.Close()
                cnn1 = Nothing
                ' MsgBox(Result.ToString)
                If WriteSRsult = True Then
                    ''add sales result here
                    Dim dep As String
                    Dim subdep As String
                    If counter >= 1 Then
                        Dim WhichManager As String
                        If Confirmer = "" And Confirmed = "0" Then
                            WhichManager = SManager
                            Description = "Appt. Set from " & PLS & " by " & Marketer & ", set up with " & PFirstName & " for " & Day & ", " & ApptDate & " at " & ApptTime.ToShortTimeString & ". Confirmed by " & SManager & ". (Confirming Bypassed)"
                            SkipConfirm = True
                            dep = "Sales"
                            subdep = "Confirming"
                        Else
                            WhichManager = MManager
                            Description = "Appt. Set from " & PLS & " by " & Marketer & ", set up with " & PFirstName & " for " & Day & ", " & ApptDate & " at " & ApptTime.ToShortTimeString & "."
                            If lrboo = True Then
                                If isRecovery = False Then
                                    dep = "Warm Calling"
                                    subdep = "Set Appointment"
                                Else
                                    dep = "Recovery"
                                    subdep = "Set Appointment"
                                End If
                            End If
                        End If



                        Dim cnn = New SqlConnection(iss_cnx)
                        cnn.open()
                        Dim cmdIns As SqlCommand = New SqlCommand("dbo.LogSetApptImport", cnn)
                        Dim param1 As SqlParameter = New SqlParameter("@ID", EId)
                        Dim param2 As SqlParameter = New SqlParameter("@User", Marketer)
                        Dim param3 As SqlParameter = New SqlParameter("@Spokewith", PFirstName)
                        Dim param4 As SqlParameter = New SqlParameter("@ApptDate", ApptDate)
                        Dim param5 As SqlParameter = New SqlParameter("@Description", Description)
                        Dim param6 As SqlParameter = New SqlParameter("@ApptTime", ApptTime.ToShortTimeString)
                        Dim param7 As SqlParameter = New SqlParameter("@Notes", "")
                        Dim param8 As SqlParameter = New SqlParameter("@Manager", WhichManager)
                        Dim param9 As SqlParameter = New SqlParameter("@Confirmed", SkipConfirm)
                        Dim param10 As SqlParameter = New SqlParameter("@PLS", PLS)
                        Dim param11 As SqlParameter = New SqlParameter("@dep", dep)
                        Dim param12 As SqlParameter = New SqlParameter("@subdep", subdep)

                        cmdIns.CommandType = CommandType.StoredProcedure
                        cmdIns.Parameters.Add(param1)
                        cmdIns.Parameters.Add(param2)
                        cmdIns.Parameters.Add(param3)
                        cmdIns.Parameters.Add(param4)
                        cmdIns.Parameters.Add(param5)
                        cmdIns.Parameters.Add(param6)
                        cmdIns.Parameters.Add(param7)
                        cmdIns.Parameters.Add(param8)
                        cmdIns.Parameters.Add(param9)
                        cmdIns.Parameters.Add(param10)
                        cmdIns.Parameters.Add(param11)
                        cmdIns.Parameters.Add(param12)
                        cmdIns.ExecuteNonQuery()
                        cnn.close()
                        cnn = Nothing
                    End If


                    If SkipConfirm = False Then
                        Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                        Dim cmdCon As SqlCommand = New SqlCommand("dbo.Confirm", cnn)
                        Dim pa1 As SqlParameter = New SqlParameter("@ID", EId)
                        Dim pa2 As SqlParameter = New SqlParameter("@User", Confirmer)
                        Dim pa3 As SqlParameter = New SqlParameter("@Spokewith", PFirstName)
                        cmdCon.CommandType = CommandType.StoredProcedure
                        cmdCon.Parameters.Add(pa1)
                        cmdCon.Parameters.Add(pa2)
                        cmdCon.Parameters.Add(pa3)
                        cnn.Open()
                        Dim Con = cmdCon.ExecuteReader
                        Con.Read()
                        Con.Close()
                        cnn.Close()
                        cnn = Nothing
                    End If
                    SkipConfirm = False
                    Dim RD As Date = ApptDate.AddDays(3)


                    Dim Reps As String
                    If Rep1 <> "" And Rep2 = "" Or Rep2 = " " Then
                        Reps = Rep1
                    ElseIf Rep1 = "" And Rep2 = "" Then
                        Reps = """Unknown"""
                    ElseIf Rep1 = "" And Rep2 <> "" Then
                        Reps = Rep2
                    Else
                        Reps = Rep1 & " and " & Rep2
                    End If

                    Dim dep2 As String

                    If ApptType = "Rehash" Then
                        dep2 = "Recovery"
                    Else
                        dep2 = "Marketing"
                    End If



                    If Result = "Not Issued" Then
                        Description = "Appt. was not issued. Logged by " & SManager & " (Forwarded back to the " & dep2 & " Department to be rescheduled)"
                    ElseIf Result = "Lost Result" Then
                        If Rep1 = "" Then
                            Description = "Appt. Logged as ""Lost Result"" by " & SManager & ", and has been forwarded back to the " & dep2 & " Department"
                        Else
                            Description = "Appt. issued to " & Reps & ", and has been logged as ""Lost Result"" by " & SManager & ". Record has been forwarded back to the " & dep2 & " Department"
                        End If

                    ElseIf Result = "Demo/No Sale" Then
                        Description = "Appt. Issued to " & Reps & ", which resulted in a Demo/No Sale. Price Quoted was $" & QuotePar & " and Par Price was $" & QuotePar & ". Logged by " & SManager
                        If Recover = True Then
                            Description = Description & " (Forwarded to the Recovery Department)"
                        End If
                    ElseIf Result = "Sale" Then
                        Dim Products As String = P1
                        If P2 <> "" Then
                            If P2 <> "" And P3 = "" Then
                                Products = Products & " and " & P2
                            Else
                                Products = Products & ", " & P2
                            End If
                        End If

                        If Finance = True Then
                            Description = "Appt. Issued to " & Reps & ", which resulted in a Sale in the amount of $" & QuotePar & ". Products Sold- " & Products & ". (Forwarded to Finance Department for approval)"
                        Else
                            Description = "Appt. Issued to " & Reps & ", which resulted in a Sale in the amount of $" & QuotePar & ". Products Sold- " & Products & ". (Forwarded to Installation Department)"
                        End If


                    ElseIf Result = "No Demo" Then
                        Description = "Appt. Issued to " & Reps & ", which resulted in a No Demo. Logged by " & SManager
                    ElseIf Result = "Reset" Or Result = "Not Hit" Then
                        Description = "Appt. issued to " & Reps & ", which resulted in a " & Result & ". Logged by " & SManager & " (Forwarded back to the " & dep2 & " Department to be rescheduled)"
                    ElseIf Result = "Lost Result" Then
                        If Rep1 = "" Then
                            Description = "Appt. Logged as ""Lost Result"" by " & SManager & ", and has been forwarded back to the " & dep2 & " Department"
                        Else
                            Description = "Appt. issued to " & Reps & ", and has been logged as ""Lost Result"" by " & SManager & ". Record has been forwarded back to the " & dep2 & " Department"
                        End If
                    End If


                    Dim NR As Boolean = False
                    Dim cnnSR As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                    Dim r As SqlDataReader
                    Dim SR1 As SqlParameter = New SqlParameter("@NR", NR)
                    Dim SR2 As SqlParameter = New SqlParameter("@Editmode", EditMode)
                    Dim SR3 As SqlParameter = New SqlParameter("@Result", Result)
                    Dim SR4 As SqlParameter = New SqlParameter("@rep1", Rep1)
                    Dim SR5 As SqlParameter = New SqlParameter("@Rep2", Rep2)
                    Dim SR6 As SqlParameter = New SqlParameter("@RDate", RD)
                    Dim SR7 As SqlParameter = New SqlParameter("@ContractAmt", QuotePar)
                    Dim SR8 As SqlParameter = New SqlParameter("@cash", Cash)
                    Dim SR9 As SqlParameter = New SqlParameter("@finance", Finance)
                    Dim SR10 As SqlParameter = New SqlParameter("@contractnotes", "")
                    Dim SR11 As SqlParameter = New SqlParameter("@Product1", P1)
                    Dim SR12 As SqlParameter = New SqlParameter("@Product2", P2)
                    Dim SR13 As SqlParameter = New SqlParameter("@Product3", P3)
                    Dim SR14 As SqlParameter = New SqlParameter("@Sold1", P1)
                    Dim SR15 As SqlParameter = New SqlParameter("@Sold2", P2)
                    Dim SR16 As SqlParameter = New SqlParameter("@Sold3", P3)
                    Dim SR17 As SqlParameter = New SqlParameter("@Sold4", "")
                    Dim SR18 As SqlParameter = New SqlParameter("@Sold5", "")
                    Dim SR19 As SqlParameter = New SqlParameter("@Split1", 0)
                    Dim SR20 As SqlParameter = New SqlParameter("@Split2", 0)
                    Dim SR21 As SqlParameter = New SqlParameter("@Split3", 0)
                    Dim SR22 As SqlParameter = New SqlParameter("@Split4", 0)
                    Dim SR23 As SqlParameter = New SqlParameter("@Split5", 0)
                    Dim SR24 As SqlParameter = New SqlParameter("@Manufacturer1", "")
                    Dim SR25 As SqlParameter = New SqlParameter("@Manufacturer2", "")
                    Dim SR26 As SqlParameter = New SqlParameter("@Manufacturer3", "")
                    Dim SR27 As SqlParameter = New SqlParameter("@Manufacturer4", "")
                    Dim SR28 As SqlParameter = New SqlParameter("@Manufacturer5", "")
                    Dim SR29 As SqlParameter = New SqlParameter("@Model1", "")
                    Dim SR30 As SqlParameter = New SqlParameter("@Model2", "")
                    Dim SR31 As SqlParameter = New SqlParameter("@Model3", "")
                    Dim SR32 As SqlParameter = New SqlParameter("@Model4", "")
                    Dim SR33 As SqlParameter = New SqlParameter("@Model5", "")
                    Dim SR34 As SqlParameter = New SqlParameter("@Style1", "")
                    Dim SR35 As SqlParameter = New SqlParameter("@Style2", "")
                    Dim SR36 As SqlParameter = New SqlParameter("@Style3", "")
                    Dim SR37 As SqlParameter = New SqlParameter("@Style4", "")
                    Dim SR38 As SqlParameter = New SqlParameter("@Style5", "")
                    Dim SR39 As SqlParameter = New SqlParameter("@Color1", "")
                    Dim SR40 As SqlParameter = New SqlParameter("@Color2", "")
                    Dim SR41 As SqlParameter = New SqlParameter("@Color3", "")
                    Dim SR42 As SqlParameter = New SqlParameter("@Color4", "")
                    Dim SR43 As SqlParameter = New SqlParameter("@Color5", "")
                    Dim SR44 As SqlParameter = New SqlParameter("@Qty1", 0)
                    Dim SR45 As SqlParameter = New SqlParameter("@Qty2", 0)
                    Dim SR46 As SqlParameter = New SqlParameter("@Qty3", 0)
                    Dim SR47 As SqlParameter = New SqlParameter("@Qty4", 0)
                    Dim SR48 As SqlParameter = New SqlParameter("@Qty5", 0)
                    Dim SR49 As SqlParameter = New SqlParameter("@Unit1", "")
                    Dim SR50 As SqlParameter = New SqlParameter("@Unit2", "")
                    Dim SR51 As SqlParameter = New SqlParameter("@Unit3", "")
                    Dim SR52 As SqlParameter = New SqlParameter("@Unit4", "")
                    Dim SR53 As SqlParameter = New SqlParameter("@Unit5", "")
                    Dim SR54 As SqlParameter = New SqlParameter("@Notes", MNotes)
                    Dim SR55 As SqlParameter = New SqlParameter("@Quoted", QuotePar)
                    Dim SR56 As SqlParameter = New SqlParameter("@Par", QuotePar)
                    Dim SR57 As SqlParameter = New SqlParameter("@Recoverable", Recover)
                    Dim SR58 As SqlParameter = New SqlParameter("@LeadHistoryId", 0)
                    Dim SR59 As SqlParameter = New SqlParameter("@ID", EId)
                    Dim SR60 As SqlParameter = New SqlParameter("@description", Description)
                    Dim SR61 As SqlParameter = New SqlParameter("@User", "Imported Data")
                    Dim cmdSR As SqlCommand
                    cmdSR = New SqlCommand("dbo.SalesResult", cnnSR)
                    cmdSR.Parameters.Add(SR1)
                    cmdSR.Parameters.Add(SR2)
                    cmdSR.Parameters.Add(SR3)
                    cmdSR.Parameters.Add(SR4)
                    cmdSR.Parameters.Add(SR5)
                    cmdSR.Parameters.Add(SR6)
                    cmdSR.Parameters.Add(SR7)
                    cmdSR.Parameters.Add(SR8)
                    cmdSR.Parameters.Add(SR9)
                    cmdSR.Parameters.Add(SR10)
                    cmdSR.Parameters.Add(SR11)
                    cmdSR.Parameters.Add(SR12)
                    cmdSR.Parameters.Add(SR13)
                    cmdSR.Parameters.Add(SR14)
                    cmdSR.Parameters.Add(SR15)
                    cmdSR.Parameters.Add(SR16)
                    cmdSR.Parameters.Add(SR17)
                    cmdSR.Parameters.Add(SR18)
                    cmdSR.Parameters.Add(SR19)
                    cmdSR.Parameters.Add(SR20)
                    cmdSR.Parameters.Add(SR21)
                    cmdSR.Parameters.Add(SR22)
                    cmdSR.Parameters.Add(SR23)
                    cmdSR.Parameters.Add(SR24)
                    cmdSR.Parameters.Add(SR25)
                    cmdSR.Parameters.Add(SR26)
                    cmdSR.Parameters.Add(SR27)
                    cmdSR.Parameters.Add(SR28)
                    cmdSR.Parameters.Add(SR29)
                    cmdSR.Parameters.Add(SR30)
                    cmdSR.Parameters.Add(SR31)
                    cmdSR.Parameters.Add(SR32)
                    cmdSR.Parameters.Add(SR33)
                    cmdSR.Parameters.Add(SR34)
                    cmdSR.Parameters.Add(SR35)
                    cmdSR.Parameters.Add(SR36)
                    cmdSR.Parameters.Add(SR37)
                    cmdSR.Parameters.Add(SR38)
                    cmdSR.Parameters.Add(SR39)
                    cmdSR.Parameters.Add(SR40)
                    cmdSR.Parameters.Add(SR41)
                    cmdSR.Parameters.Add(SR42)
                    cmdSR.Parameters.Add(SR43)
                    cmdSR.Parameters.Add(SR44)
                    cmdSR.Parameters.Add(SR45)
                    cmdSR.Parameters.Add(SR46)
                    cmdSR.Parameters.Add(SR47)
                    cmdSR.Parameters.Add(SR48)
                    cmdSR.Parameters.Add(SR49)
                    cmdSR.Parameters.Add(SR50)
                    cmdSR.Parameters.Add(SR51)
                    cmdSR.Parameters.Add(SR52)
                    cmdSR.Parameters.Add(SR53)
                    cmdSR.Parameters.Add(SR54)
                    cmdSR.Parameters.Add(SR55)
                    cmdSR.Parameters.Add(SR56)
                    cmdSR.Parameters.Add(SR57)
                    cmdSR.Parameters.Add(SR58)
                    cmdSR.Parameters.Add(SR59)
                    cmdSR.Parameters.Add(SR60)
                    cmdSR.Parameters.Add(SR61)
                    cmdSR.CommandType = CommandType.StoredProcedure

                    cnnSR.Open()
                    r = cmdSR.ExecuteReader(CommandBehavior.CloseConnection)
                    r.Read()
                    r.Close()
                    cnnSR.Close()
                    cnnSR = Nothing
               
                    ''Log Recission Cancel and Bank Rejects Here
                    If RCorBR = "Recission Cancel" Or RCorBR = "Bank Rejected" Then
                        If RCorBR = "Recission Cancel" Then
                            Dim cnn5 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                            Dim cmdGet5 As SqlCommand
                            Dim r5 As SqlDataReader
                            cmdGet5 = New SqlCommand("Select Contact1FirstName, Contact2FirstName from Enterlead where id = " & EId, cnn5)
                            cmdGet5.CommandType = CommandType.Text
                            cnn5.Open()
                            r5 = cmdGet5.ExecuteReader(CommandBehavior.CloseConnection)
                            r5.Read()
                            Dim Customer As String
                            Dim hashave As String
                            Dim theirhisher As String
                            If r5.Item(1) = "" Then
                                Customer = r5.Item(0)
                                hashave = "has"
                                theirhisher = "his/her"
                            Else
                                Customer = r5.Item(0) & " and " & r5.Item(1)
                                hashave = "have"
                                theirhisher = "their"
                            End If
                            r5.Close()
                            cnn5.Close()
                            cnn5 = Nothing

                            Dim cnn6 As New SqlConnection(iss_cnx)
                            cmdGet5 = New SqlCommand("Select Financing,Installation from tblwherecanleadgo where leadnumber = " & EId, cnn6)
                            cnn6.Open()

                            r5 = cmdGet5.ExecuteReader(CommandBehavior.CloseConnection)
                            r5.Read()
                            Dim dep1 As String
                            If r5.Item(0) = True And r5.Item(1) = True Then
                                dep1 = "(Removed from Installation and Finance Departments)"
                            ElseIf r5.Item(0) = True And r5.Item(1) = False Then
                                dep1 = "(Removed from Finance Department)"
                            ElseIf r5.Item(1) = True And r5.Item(0) = False Then
                                dep1 = "(Removed from Installation Department)"
                            Else
                                dep1 = ""
                            End If
                            r5.Close()
                            cnn6.Close()
                            cnn6 = Nothing

                            If Recover = False Then
                                Description = Customer & " " & hashave & " decided to cancel " & theirhisher & " contract within the ""Right to Rescind"" date. Recission Cancel logged by " & SManager & "." & dep1
                            Else
                                Description = Customer & " " & hashave & " decided to cancel " & theirhisher & " contract within the ""Right to Rescind"" date. Recission Cancel logged by " & SManager & ", and forwarded to the Recovery Department." & dep1
                            End If
                        ElseIf RCorBR = "Bank Rejected" Then
                            Description = "Due to credit issues, this contract has been rejected by the Finance Department. Rejection logged by " & SManager

                        End If
                        cnnSR = New SqlConnection(STATIC_VARIABLES.Cnn)
                        SR1 = New SqlParameter("@NR", NR)
                        SR2 = New SqlParameter("@Editmode", EditMode)
                        SR3 = New SqlParameter("@Result", RCorBR)
                        SR4 = New SqlParameter("@rep1", Rep1)
                        SR5 = New SqlParameter("@Rep2", Rep2)
                        SR6 = New SqlParameter("@RDate", RD)
                        SR7 = New SqlParameter("@ContractAmt", QuotePar)
                        SR8 = New SqlParameter("@cash", Cash)
                        SR9 = New SqlParameter("@finance", Finance)
                        SR10 = New SqlParameter("@contractnotes", "")
                        SR11 = New SqlParameter("@Product1", P1)
                        SR12 = New SqlParameter("@Product2", P2)
                        SR13 = New SqlParameter("@Product3", P3)
                        SR14 = New SqlParameter("@Sold1", P1)
                        SR15 = New SqlParameter("@Sold2", P2)
                        SR16 = New SqlParameter("@Sold3", P3)
                        SR17 = New SqlParameter("@Sold4", "")
                        SR18 = New SqlParameter("@Sold5", "")
                        SR19 = New SqlParameter("@Split1", 0)
                        SR20 = New SqlParameter("@Split2", 0)
                        SR21 = New SqlParameter("@Split3", 0)
                        SR22 = New SqlParameter("@Split4", 0)
                        SR23 = New SqlParameter("@Split5", 0)
                        SR24 = New SqlParameter("@Manufacturer1", "")
                        SR25 = New SqlParameter("@Manufacturer2", "")
                        SR26 = New SqlParameter("@Manufacturer3", "")
                        SR27 = New SqlParameter("@Manufacturer4", "")
                        SR28 = New SqlParameter("@Manufacturer5", "")
                        SR29 = New SqlParameter("@Model1", "")
                        SR30 = New SqlParameter("@Model2", "")
                        SR31 = New SqlParameter("@Model3", "")
                        SR32 = New SqlParameter("@Model4", "")
                        SR33 = New SqlParameter("@Model5", "")
                        SR34 = New SqlParameter("@Style1", "")
                        SR35 = New SqlParameter("@Style2", "")
                        SR36 = New SqlParameter("@Style3", "")
                        SR37 = New SqlParameter("@Style4", "")
                        SR38 = New SqlParameter("@Style5", "")
                        SR39 = New SqlParameter("@Color1", "")
                        SR40 = New SqlParameter("@Color2", "")
                        SR41 = New SqlParameter("@Color3", "")
                        SR42 = New SqlParameter("@Color4", "")
                        SR43 = New SqlParameter("@Color5", "")
                        SR44 = New SqlParameter("@Qty1", "")
                        SR45 = New SqlParameter("@Qty2", "")
                        SR46 = New SqlParameter("@Qty3", "")
                        SR47 = New SqlParameter("@Qty4", "")
                        SR48 = New SqlParameter("@Qty5", "")
                        SR49 = New SqlParameter("@Unit1", "")
                        SR50 = New SqlParameter("@Unit2", "")
                        SR51 = New SqlParameter("@Unit3", "")
                        SR52 = New SqlParameter("@Unit4", "")
                        SR53 = New SqlParameter("@Unit5", "")
                        SR54 = New SqlParameter("@Notes", MNotes)
                        SR55 = New SqlParameter("@Quoted", QuotePar)
                        SR56 = New SqlParameter("@Par", QuotePar)
                        SR57 = New SqlParameter("@Recoverable", Recover)
                        SR58 = New SqlParameter("@LeadHistoryId", 0)
                        SR59 = New SqlParameter("@ID", EId)
                        SR60 = New SqlParameter("@description", Description)
                        SR61 = New SqlParameter("@User", "Imported Data")

                        cmdSR = New SqlCommand("dbo.SalesResult", cnnSR)
                        cmdSR.Parameters.Add(SR1)
                        cmdSR.Parameters.Add(SR2)
                        cmdSR.Parameters.Add(SR3)
                        cmdSR.Parameters.Add(SR4)
                        cmdSR.Parameters.Add(SR5)
                        cmdSR.Parameters.Add(SR6)
                        cmdSR.Parameters.Add(SR7)
                        cmdSR.Parameters.Add(SR8)
                        cmdSR.Parameters.Add(SR9)
                        cmdSR.Parameters.Add(SR10)
                        cmdSR.Parameters.Add(SR11)
                        cmdSR.Parameters.Add(SR12)
                        cmdSR.Parameters.Add(SR13)
                        cmdSR.Parameters.Add(SR14)
                        cmdSR.Parameters.Add(SR15)
                        cmdSR.Parameters.Add(SR16)
                        cmdSR.Parameters.Add(SR17)
                        cmdSR.Parameters.Add(SR18)
                        cmdSR.Parameters.Add(SR19)
                        cmdSR.Parameters.Add(SR20)
                        cmdSR.Parameters.Add(SR21)
                        cmdSR.Parameters.Add(SR22)
                        cmdSR.Parameters.Add(SR23)
                        cmdSR.Parameters.Add(SR24)
                        cmdSR.Parameters.Add(SR25)
                        cmdSR.Parameters.Add(SR26)
                        cmdSR.Parameters.Add(SR27)
                        cmdSR.Parameters.Add(SR28)
                        cmdSR.Parameters.Add(SR29)
                        cmdSR.Parameters.Add(SR30)
                        cmdSR.Parameters.Add(SR31)
                        cmdSR.Parameters.Add(SR32)
                        cmdSR.Parameters.Add(SR33)
                        cmdSR.Parameters.Add(SR34)
                        cmdSR.Parameters.Add(SR35)
                        cmdSR.Parameters.Add(SR36)
                        cmdSR.Parameters.Add(SR37)
                        cmdSR.Parameters.Add(SR38)
                        cmdSR.Parameters.Add(SR39)
                        cmdSR.Parameters.Add(SR40)
                        cmdSR.Parameters.Add(SR41)
                        cmdSR.Parameters.Add(SR42)
                        cmdSR.Parameters.Add(SR43)
                        cmdSR.Parameters.Add(SR44)
                        cmdSR.Parameters.Add(SR45)
                        cmdSR.Parameters.Add(SR46)
                        cmdSR.Parameters.Add(SR47)
                        cmdSR.Parameters.Add(SR48)
                        cmdSR.Parameters.Add(SR49)
                        cmdSR.Parameters.Add(SR50)
                        cmdSR.Parameters.Add(SR51)
                        cmdSR.Parameters.Add(SR52)
                        cmdSR.Parameters.Add(SR53)
                        cmdSR.Parameters.Add(SR54)
                        cmdSR.Parameters.Add(SR55)
                        cmdSR.Parameters.Add(SR56)
                        cmdSR.Parameters.Add(SR57)
                        cmdSR.Parameters.Add(SR58)
                        cmdSR.Parameters.Add(SR59)
                        cmdSR.Parameters.Add(SR60)
                        cmdSR.Parameters.Add(SR61)
                        cmdSR.CommandType = CommandType.StoredProcedure
                        Dim r6 As SqlDataReader
                        cnnSR.Open()
                        r6 = cmdSR.ExecuteReader(CommandBehavior.CloseConnection)
                        r6.Read()
                        r6.Close()
                        cnnSR.Close()
                        cnnSR = Nothing
                    End If



                Else
                    ''Called & Cancelled Here 

                    If Result = "Called and Cancelled" Then


                        Day = ApptDate.DayOfWeek
                        If Day = 0 Then
                            Day = "Sunday"
                        ElseIf Day = 1 Then
                            Day = "Monday"
                        ElseIf Day = 2 Then
                            Day = "Tuesday"
                        ElseIf Day = 3 Then
                            Day = "Wednesday"
                        ElseIf Day = 4 Then
                            Day = "Thursday"
                        ElseIf Day = 5 Then
                            Day = "Friday"
                        ElseIf Day = 6 Then
                            Day = "Saturday"
                        End If


                        If ResultDetail = "Not Interested" Then
                            Description = PFirstName & " called in at " & "Unknown Time (Data Import)" & " spoke with " & Confirmer & " to cancel Appt., and has no interest in rescheduling at this time."
                        Else
                            Description = PFirstName & " called in at " & "Unknown Time (Data Import)" & " spoke with " & Confirmer & " to cancel Appt., and would like a call back to reschedule."

                        End If



                        Dim dt As Date = ApptDate

                        Dim tm As Date = ApptTime




                        Dim cnn4 = New SqlConnection(STATIC_VARIABLES.Cnn)

                        Dim cmdCC As SqlCommand = New SqlCommand("dbo.LogCancelledAppt", cnn4)
                        Dim CC1 As SqlParameter = New SqlParameter("@ID", EId)
                        Dim CC2 As SqlParameter = New SqlParameter("@User", "Data Import")
                        Dim CC3 As SqlParameter = New SqlParameter("@Spokewith", PFirstName)
                        Dim CC4 As SqlParameter = New SqlParameter("@ApptDate", dt)
                        Dim CC5 As SqlParameter = New SqlParameter("@Description", Description)
                        Dim CC6 As SqlParameter = New SqlParameter("@ApptTime", tm.ToShortTimeString)
                        Dim CC7 As SqlParameter = New SqlParameter("@Notes", ResultDetail)
                        cmdCC.CommandType = CommandType.StoredProcedure
                        cmdCC.Parameters.Add(CC1)
                        cmdCC.Parameters.Add(CC2)
                        cmdCC.Parameters.Add(CC3)
                        cmdCC.Parameters.Add(CC4)
                        cmdCC.Parameters.Add(CC5)
                        cmdCC.Parameters.Add(CC6)
                        cmdCC.Parameters.Add(CC7)
                        cnn4.Open()
                        Dim Rc As SqlDataReader
                        Rc = cmdCC.ExecuteReader(CommandBehavior.CloseConnection)
                        Rc.Read()
                        Rc.Close()
                        cnn4.Close()
                        cnn4 = Nothing
                    End If


                End If


                counter += 1
            End While

            r2.Close()
            If counter = 0 Then
                ''Add sql to look for marketing tasks if no appointments exist
                Dim cnxDA As New SqlConnection(da_cnx)
                cmdApp = New SqlCommand("dbo.GetLeadSource", cnxDA)
                cnxDA.Open()
                s1 = New SqlParameter("@id", id)
                cmdApp.CommandType = CommandType.StoredProcedure
                cmdApp.Parameters.Add(s1)


                r2 = cmdApp.ExecuteReader
                'Dim counter2 As Integer
                While r2.Read
                    ''sql to look for appointments for this prospect 
                    ''update enterlead data 
                    'counter2 += 1
                    Marketer = r2.Item(0)
                    PLS = r2.Item(1)
                    Generated = r2.Item(2)
                    P1 = r2.Item(3)
                    ApptDate = r2.Item(2)
                    ApptTime = "1/1/1900 05:00 PM"
                End While
                r2.Close()
                cnxDA.Close()
                cnxDA = Nothing

                If Marketer = "" Then
                    Description = "Appt. Generated from " & PLS & ", set up with " & PFirstName
                ElseIf Marketer <> "" Then
                    Description = "Appt. Generated from " & PLS & " by " & Marketer & ", set up with " & PFirstName
                End If
                Dim cnxISS As New SqlConnection(iss_cnx)
                Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertEnterLead", cnxISS)
                cmdINS.CommandType = CommandType.StoredProcedure
                cnxISS.Open()
                Dim param1 As SqlParameter = New SqlParameter("@Marketer", Trim(Marketer))
                Dim param2 As SqlParameter = New SqlParameter("@PLS", Trim(PLS))
                Dim param3 As SqlParameter = New SqlParameter("@SLS", Trim(""))
                Dim param4 As SqlParameter = New SqlParameter("@LeadGenOn", Generated)
                Dim param5 As SqlParameter = New SqlParameter("@Contact1FirstName", Trim(PFirstName))
                Dim param6 As SqlParameter = New SqlParameter("@Contact1LastName", Trim(PLastName))
                Dim param7 As SqlParameter = New SqlParameter("@Contact2FirstName", Trim(SFirstName))
                Dim param69 As SqlParameter = New SqlParameter("@Contact2LastName", Trim(SLastName))
                Dim param8 As SqlParameter = New SqlParameter("@YearsOwned", YearsOwned)
                Dim param9 As SqlParameter = New SqlParameter("@HomeAge", Trim(HomeAge))
                Dim param10 As SqlParameter = New SqlParameter("@HomeValue", HomeVal)
                Dim param11 As SqlParameter = New SqlParameter("@SpecialInstruction", "")

                '' strip off literals
                '' '(' ')' '-'
                '' -------------------------------------------------------------------      
                DoLiteralsExist(HPhone)
                Dim param12 As SqlParameter = New SqlParameter("@HousePhone", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(Alt1)
                Dim param13 As SqlParameter = New SqlParameter("@AltPhone1", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(Alt2)
                Dim param14 As SqlParameter = New SqlParameter("@AltPhone2", Me.CorrectPhoneNumberResponse.ToString)
                '' -----------------------------------------------------------------------------------

                Dim param15 As SqlParameter = New SqlParameter("@Alt1Type", Alt1Type)
                Dim param16 As SqlParameter = New SqlParameter("@Alt2Type", Alt2Type)
                'Dim param17 As SqlParameter = New SqlParameter("@StNumber", EnterLead.txtStNum.Text)
                Dim stAddress As String = ""
                Dim param18 As SqlParameter = New SqlParameter("@StAddress", Trim(Addy))
                Dim param19 As SqlParameter = New SqlParameter("@City", Trim(City))
                Dim param20 As SqlParameter = New SqlParameter("@State", State)
                Dim param21 As SqlParameter = New SqlParameter("@Zip", Zip)
                Dim param22 As SqlParameter = New SqlParameter("@SpokeWith", PFirstName)
                Dim param23 As SqlParameter = New SqlParameter("@C1Work", "Unknown")
                Dim param24 As SqlParameter = New SqlParameter("@C2Work", "")
                Dim param25 As SqlParameter = New SqlParameter("@AppDate", ApptDate)

                Dim d1 = ApptDate.DayOfWeek
                Select Case d1
                    Case Is = 0
                        Day = "Sunday"
                        Exit Select
                    Case Is = 1
                        Day = "Monday"
                        Exit Select
                    Case Is = 2
                        Day = "Tuesday"
                        Exit Select
                    Case Is = 3
                        Day = "Wednesday"
                        Exit Select
                    Case Is = 4
                        Day = "Thursday"
                        Exit Select
                    Case Is = 5
                        Day = "Friday"
                        Exit Select
                    Case Is = 6
                        Day = "Saturday"
                        Exit Select
                End Select

                Dim param26 As SqlParameter = New SqlParameter("@AppDay", Day)
                Dim param27 As SqlParameter = New SqlParameter("@AppTime", ApptTime.ToShortTimeString)
                Dim param28 As SqlParameter = New SqlParameter("@Product1", P1)
                Dim param29 As SqlParameter = New SqlParameter("@Product2", "")
                Dim param30 As SqlParameter = New SqlParameter("@Product3", "")
                Dim param31 As SqlParameter = New SqlParameter("@Color", "")
                Dim param32 As SqlParameter = New SqlParameter("@QTY", 0)
                Dim param33 As SqlParameter = New SqlParameter("@Email", Email)
                Dim param34 As SqlParameter = New SqlParameter("@ProdAcro1", "")
                Dim param35 As SqlParameter = New SqlParameter("@ProdAcro2", "")
                Dim param36 As SqlParameter = New SqlParameter("@prodAcro3", "")
                Dim param37 As SqlParameter = New SqlParameter("@User", "Imported Data")
                Dim param38 As SqlParameter = New SqlParameter("@MarketingManager", "")
                Dim param39 As SqlParameter = New SqlParameter("@Description", Description)
                Dim param70 As SqlParameter = New SqlParameter("@Mapped", Mapped)

                cmdINS.Parameters.Add(param1)
                cmdINS.Parameters.Add(param2)
                cmdINS.Parameters.Add(param3)
                cmdINS.Parameters.Add(param4)
                cmdINS.Parameters.Add(param5)
                cmdINS.Parameters.Add(param6)
                cmdINS.Parameters.Add(param7)
                cmdINS.Parameters.Add(param8)
                cmdINS.Parameters.Add(param9)
                cmdINS.Parameters.Add(param10)
                cmdINS.Parameters.Add(param11)
                cmdINS.Parameters.Add(param12)
                cmdINS.Parameters.Add(param13)
                cmdINS.Parameters.Add(param14)
                cmdINS.Parameters.Add(param15)
                cmdINS.Parameters.Add(param16)
                'cmdINS.Parameters.Add(param17)
                cmdINS.Parameters.Add(param18)
                cmdINS.Parameters.Add(param19)
                cmdINS.Parameters.Add(param20)
                cmdINS.Parameters.Add(param21)
                cmdINS.Parameters.Add(param22)
                cmdINS.Parameters.Add(param23)
                cmdINS.Parameters.Add(param24)
                cmdINS.Parameters.Add(param25)
                cmdINS.Parameters.Add(param26)
                cmdINS.Parameters.Add(param27)
                cmdINS.Parameters.Add(param28)
                cmdINS.Parameters.Add(param29)
                cmdINS.Parameters.Add(param30)
                cmdINS.Parameters.Add(param31)
                cmdINS.Parameters.Add(param32)
                cmdINS.Parameters.Add(param69)
                cmdINS.Parameters.Add(param33)
                cmdINS.Parameters.Add(param34)
                cmdINS.Parameters.Add(param35)
                cmdINS.Parameters.Add(param36)
                cmdINS.Parameters.Add(param37)
                cmdINS.Parameters.Add(param38)
                cmdINS.Parameters.Add(param39)
                cmdINS.Parameters.Add(param70)
                Dim INS As SqlDataReader
                INS = cmdINS.ExecuteReader
                INS.Close()
                cnxISS.Close()
                cnxISS = Nothing

            End If
            ''kill unqualified appts here after all leadhistory has been assembled
            If Kill = "1" Then

                Description = """Imported Data"" spoke with " & PFirstName & " and killed this appt." & " Reason: Imported Data- Unknown Reason"


                Dim cnn1 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                Dim cmdId As SqlCommand = New SqlCommand("select max (id) from enterlead", cnn1)
                cnn1.Open()
                Dim getID As SqlDataReader
                getID = cmdId.ExecuteReader
                getID.Read()

                Dim EId = getID.Item(0)
                getID.Close()
                cnn1.Close()
                cnn1 = Nothing

                Dim cnn = New SqlConnection(STATIC_VARIABLES.Cnn)
                Dim cmdIns As SqlCommand = New SqlCommand("dbo.KillAppt", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@ID", EId)
                Dim param2 As SqlParameter = New SqlParameter("@User", "Imported Data")
                Dim param3 As SqlParameter = New SqlParameter("@Spokewith", PFirstName)
                Dim param4 As SqlParameter = New SqlParameter("@Description", Description)
                Dim param5 As SqlParameter = New SqlParameter("@Notes", "")
                cmdIns.CommandType = CommandType.StoredProcedure
                cmdIns.Parameters.Add(param1)
                cmdIns.Parameters.Add(param2)
                cmdIns.Parameters.Add(param3)
                cmdIns.Parameters.Add(param4)
                cmdIns.Parameters.Add(param5)
                cnn.Open()
                Dim R9 As SqlDataReader
                R9 = cmdIns.ExecuteReader(CommandBehavior.CloseConnection)
                R9.Read()
                R9.Close()
                cnn.close()
                cnn = Nothing

            End If
            '' junk 11-3-2015 AC
            '' no code behind.....
            '' write new form / front end
            'Form6.ProgressBar1.PerformStep()

            'cnnDA2.Close()
            cnx2.Close()
            cnx2 = Nothing
            Dim recID As String = GetMaxID()
            VerifyAddress(Get_Address_To_Verify(recID), False)
        End While


        Dim cnn0 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet0 As SqlCommand
        Dim r0 As SqlDataReader
        cmdGet0 = New SqlCommand("Update Enterlead Set killpending = 'False' where marketingresults = 'Kill'", cnn0)
        cmdGet0.CommandType = CommandType.Text
        cnn0.Open()
        r0 = cmdGet0.ExecuteReader(CommandBehavior.CloseConnection)
        r0.Read()
        cnn0.Close()
        cnn0 = Nothing
        ''Add code to flip all Kill pendings to False 
        r1.Close()
        cnnDA.Close()
        cnnDA = Nothing
        Form6.lblTable.Visible = False
        Form6.Label2.Text = "Import Data Successful"
        Form6.lblTblCurCnt.Text = ""
        Form6.lblTotalRecordNum.Text = ""
        Form6.lblTable.Text = "Successfully Imported " & iteration.ToString
        Form6.btnImport.Text = "Close"
    End Sub

    Private Function DoLiteralsExist(ByVal NumberToCheck As String) As String
        Dim len As Integer = 0
        len = NumberToCheck.Length
        '' length: 10 = No Number
        '' length: 14 = Has number
        Select Case len
            Case Is = 10
                CorrectPhoneNumberResponse = ""
                Exit Select
            Case Is = 14
                CorrectPhoneNumberResponse = NumberToCheck
                Exit Select
        End Select
        Return CorrectPhoneNumberResponse
    End Function
    Public Sub VerifyAddress(ByVal obj As ImportData_V2.sqlOperations.Record_And_Address, ByVal DevOrPro As Boolean)
        Try
            oApp = STATIC_VARIABLES.oApp
            Dim oMap As MapPoint.Map = oApp.ActiveMap
            oApp.Visible = False
            oMap.Saved = True
            Dim oRes As MapPoint.FindResults

            oRes = oMap.FindAddressResults(obj.StreetAddress, obj.City, , obj.State, obj.Zip, GeoCountry.geoCountryUnitedStates)
            ''find results quality reference: 
            '' https://msdn.microsoft.com/en-us/library/aa736235.aspx
            Select Case oRes.ResultsQuality
                'Case Is = GeoFindResultsQuality.geoAllResultsValid   DONT USE
                '    '' not used
                Case Is = GeoFindResultsQuality.geoAmbiguousResults
                    '' update table with FIRST address
                    'Dim loc As MapPoint.Location = oRes.Item(1) '' Results are NON ZERO BASED INDEXES
                    'Dim stAddress As String = loc.StreetAddress.Street
                    'Dim City As String = loc.StreetAddress.City
                    'Dim zip As String = loc.StreetAddress.PostalCode
                    'Dim State As String = loc.StreetAddress.Region
                    'Dim Id As String = obj.Id_FromTable
                    'Dim t_name As String = "i360__Prospect__c"

                    'Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    'v.Id_FromTable = Id
                    'v.StreetAddress = stAddress
                    'v.City = City
                    'v.Zip = zip
                    'v.State = State
                    'Dim b As New ImportData_V2.sqlOperations
                    'b.Update_Table(v, DevOrPro)
                    'b = Nothing
                    Exit Select
                Case Is = GeoFindResultsQuality.geoFirstResultGood
                    '' update table with address
                    Dim loc As MapPoint.Location = oRes.Item(1) '' Results are NON ZERO BASED INDEXES
                    Dim stAddress As String = loc.StreetAddress.Street
                    Dim City As String = loc.StreetAddress.City
                    Dim zip As String = loc.StreetAddress.PostalCode
                    Dim State As String = loc.StreetAddress.Region
                    Dim Id As String = obj.Id_FromTable
                    Dim t_name As String = "i360__Prospect__c"

                    Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    v.Id_FromTable = Id
                    v.StreetAddress = stAddress
                    v.City = City
                    v.Zip = zip
                    v.State = State
                    Dim b As New ImportData_V2.sqlOperations
                    b.Update_Table(v, DevOrPro)
                    Update_Table_Verified_Address(v.Id_FromTable)
                    b = Nothing
                    Exit Select
                Case Is = GeoFindResultsQuality.geoNoGoodResult
                    '' dump to proxy table

                    'Dim stAddress As String = obj.StreetAddress
                    'Dim City As String = obj.City
                    'Dim zip As String = obj.Zip
                    'Dim State As String = obj.State
                    'Dim Id As String = obj.Id_FromTable
                    'Dim t_name As String = "i360__Prospect__c"

                    'Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    'v.Id_FromTable = Id
                    'v.StreetAddress = stAddress
                    'v.City = City
                    'v.Zip = zip
                    'v.State = State
                    ' Dim b As New sqlOperations
                    'b.Dump_To_ProxyTable(v, DevOrPro)
                    Exit Select
                Case Is = GeoFindResultsQuality.geoNoResults
                    '' dump to proxy table 
                    'Dim stAddress As String = obj.StreetAddress
                    'Dim City As String = obj.City
                    'Dim zip As String = obj.Zip
                    'Dim State As String = obj.State
                    'Dim Id As String = obj.Id_FromTable
                    'Dim t_name As String = "i360__Prospect__c"

                    'Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    'v.Id_FromTable = Id
                    'v.StreetAddress = stAddress
                    'v.City = City
                    'v.Zip = zip
                    'v.State = State
                    ' Dim b As New sqlOperations
                    ' b.Dump_To_ProxyTable(v, DevOrPro)
                    Exit Select
            End Select
            oRes = Nothing
            oMap.ActiveRoute.Clear()



        Catch ex As Exception
            'Dim stAddress As String = obj.StreetAddress
            'Dim City As String = obj.City
            'Dim zip As String = obj.Zip
            'Dim State As String = obj.State
            'Dim Id As String = obj.Id_FromTable
            'Dim t_name As String = "i360__Prospect__c"

            'Dim v As New ImportData_V2.sqlOperations.Record_And_Address
            'v.Id_FromTable = Id
            'v.StreetAddress = stAddress
            'v.City = City
            'v.Zip = zip
            'v.State = State
            'Dim b As New sqlOperations
            'b.Dump_To_ProxyTable(v, DevOrPro)
            Dim err As String = ex.Message
            MsgBox("Error: " & vbCrLf & err, MsgBoxStyle.Critical, "DEBUG INFO - ImportData.VerifyAddress ")
        End Try
    End Sub


    Private Sub Update_Table_Verified_Address(ByVal RecID As String)
        Try
            Dim cnxUP As New SqlConnection(iss_cnx)
            Dim cmdInsert As New SqlCommand("UPDATE VerifiedAddress SET Verified = 1 WHERE LeadNum = '" & RecID & "';", cnxUP)
            cnxUP.Open()
            cmdInsert.ExecuteScalar()
            cnxUP.Close()
            cnxUP = Nothing
        Catch ex As Exception
            Dim err As String = ex.Message
            MsgBox("ERROR: " & vbCrLf & err, MsgBoxStyle.Critical, "DEBUG INFO - ImportData.Update_Table_Verified_Address(recID)")
        End Try
    End Sub

    Private Function Get_Address_To_Verify(ByVal LeadNum As String)
        Dim cnxVerify As New SqlConnection(iss_cnx)
        cnxVerify.Open()
        Dim cmdGET As New SqlCommand("select ID, StAddress, City, State, Zip from EnterLead WHERE ID = '" & LeadNum & "';", cnxVerify)
        Dim y As New ImportData_V2.sqlOperations.Record_And_Address
        y.Id_FromTable = LeadNum
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            y.StreetAddress = r1.Item("StAddress")
            y.City = r1.Item("City")
            y.State = r1.Item("State")
            y.Zip = r1.Item("Zip")
        End While
        r1.Close()
        cnxVerify.Close()
        cnxVerify = Nothing
        Return y
    End Function

    Private Function GetMaxID()
        Dim cnxVerify As New SqlConnection(iss_cnx)
        cnxVerify.Open()
        Dim cmdGET As New SqlCommand("SELECT MAX(ID) from EnterLead;", cnxVerify)
        Dim res As String = cmdGET.ExecuteScalar
        cnxVerify.Close()
        cnxVerify = Nothing
        Return res
    End Function
End Class
