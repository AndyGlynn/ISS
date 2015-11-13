
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports MapPoint

Public Class ENTER_LEAD
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Public CorrectPhoneNumberResponse As String = ""
    Public Sub Loadup()
        Dim b As New ENTER_LEAD.PopulateMarketers
        b.GetMarketers()
        Dim c As New ENTER_LEAD.PopulatePrimaryLeadSource
        c.GetPrimaryLeadSource()
        Dim cc As New ENTER_LEAD.PopulateCities
        cc.GetCities()
        Dim wh As New ENTER_LEAD.PopulateWorkHours
        wh.GetWorkHours()
        Dim pr As New ENTER_LEAD.PopulateProducts
        pr.GetProducts()
        EnterLead.dtpApptInfo.Value = Now()
    
        'GetDayOfWeek()
    End Sub
    Public Sub GetDayOfWeek()
        'Exit Sub '' just to see if this keeps it from crashing
        Dim d1 = EnterLead.dtpApptInfo.Value.DayOfWeek
        Select Case d1
            Case Is = 0
                EnterLead.txtApptday.Text = "Sunday"
                Exit Select
            Case Is = 1
                EnterLead.txtApptday.Text = "Monday"
                Exit Select
            Case Is = 2
                EnterLead.txtApptday.Text = "Tuesday"
                Exit Select
            Case Is = 3
                EnterLead.txtApptday.Text = "Wednesday"
                Exit Select
            Case Is = 4
                EnterLead.txtApptday.Text = "Thursday"
                Exit Select
            Case Is = 5
                EnterLead.txtApptday.Text = "Friday"
                Exit Select
            Case Is = 6
                EnterLead.txtApptday.Text = "Saturday"
                Exit Select
        End Select
    End Sub
    Public Sub GetAcronym(ByVal Product As String, ByVal productnum As Integer)
        Try
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetProductAcronym", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@Product", Product)
            cmdGet.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet.ExecuteReader
            r1.Read()
            If productnum = 1 Then
                EnterLead.txtP1acro.Text = r1.Item(0)
            ElseIf productnum = 2 Then
                EnterLead.txtP2acro.Text = r1.Item(0)
            ElseIf productnum = 3 Then
                EnterLead.txtP3acro.Text = r1.Item(0)
            End If
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("ENTER_LEAD", "ByVal product as string, byval Productnum as integer", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetAcronym")
        End Try


    End Sub
    Public Class PopulateMarketers
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public dset_Marketers As Data.DataSet = New Data.DataSet("MARKETERS")
        Public da_Marketers As SqlDataAdapter = New SqlDataAdapter
        Public Sub GetMarketers()
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetMarketers", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                cnn.Open()
                da_Marketers.SelectCommand = cmdGet
                da_Marketers.Fill(dset_Marketers, "MarketerPullList")
                Dim cnt As Integer = 0
                cnt = dset_Marketers.Tables(0).Rows.Count
                Select Case cnt
                    Case Is <= 0
                        EnterLead.cboMarketer.Items.Clear()
                        EnterLead.cboMarketer.Items.Add("<Add New>")
                        EnterLead.cboMarketer.Items.Add("_____________________________________________")
                        EnterLead.cboMarketer.Items.Add("")
                        Exit Select
                    Case Is >= 1
                        EnterLead.cboMarketer.Items.Clear()
                        EnterLead.cboMarketer.Items.Add("<Add New>")
                        EnterLead.cboMarketer.Items.Add("_____________________________________________")
                        EnterLead.cboMarketer.Items.Add("")
                        Dim b
                        For b = 0 To dset_Marketers.Tables(0).Rows.Count - 1
                            EnterLead.cboMarketer.Items.Add(dset_Marketers.Tables(0).Rows(b).Item(1) & " " & dset_Marketers.Tables(0).Rows(b).Item(2))
                        Next
                        Exit Select
                End Select
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulateMarketers", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetMarketers")

            End Try

        End Sub
    End Class
    Public Class PopulatePrimaryLeadSource
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public dset_PriLS As Data.DataSet = New Data.DataSet("PLS")
        Public da_PRI As SqlDataAdapter = New SqlDataAdapter
        Public Sub GetPrimaryLeadSource()
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetPLS", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                cnn.Open()
                da_PRI.SelectCommand = cmdGet
                da_PRI.Fill(dset_PriLS, "PrimaryLeadSource")
                Dim cnt As Integer = 0
                cnt = dset_PriLS.Tables(0).Rows.Count
                Select Case cnt
                    Case Is <= 0
                        EnterLead.cboPriLead.Items.Clear()
                        EnterLead.cboPriLead.Items.Add("<Add New>")
                        EnterLead.cboPriLead.Items.Add("_____________________________________________")
                        EnterLead.cboPriLead.Items.Add("")
                        Exit Select
                    Case Is >= 1
                        EnterLead.cboPriLead.Items.Clear()
                        EnterLead.cboPriLead.Items.Add("<Add New>")
                        EnterLead.cboPriLead.Items.Add("_____________________________________________")
                        EnterLead.cboPriLead.Items.Add("")
                        Dim b
                        For b = 0 To dset_PriLS.Tables(0).Rows.Count - 1
                            EnterLead.cboPriLead.Items.Add(dset_PriLS.Tables(0).Rows(b).Item(1))
                        Next
                        Exit Select
                End Select
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulatePrimaryLeadSource", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetPrimaryLeadSource")

            End Try

        End Sub
    End Class
    Public Class PopulateSecondaryLeadSource
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public dset_SLS As Data.DataSet = New Data.DataSet("SLS")
        Public da_SLS As SqlDataAdapter = New SqlDataAdapter
        Public Sub GetSLS(ByVal PrimaryLS As String)
            Try
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
                        EnterLead.cboSecLead.Items.Clear()
                        EnterLead.cboSecLead.Items.Add("<Add New>")
                        EnterLead.cboSecLead.Items.Add("_____________________________________________")
                        EnterLead.cboSecLead.Items.Add("")
                        Exit Select
                    Case Is >= 1
                        EnterLead.cboSecLead.Items.Clear()
                        EnterLead.cboSecLead.Items.Add("<Add New>")
                        EnterLead.cboSecLead.Items.Add("_____________________________________________")
                        EnterLead.cboSecLead.Items.Add("")
                        Dim b As Integer = 0
                        For b = 0 To dset_SLS.Tables(0).Rows.Count - 1
                            EnterLead.cboSecLead.Items.Add(dset_SLS.Tables(0).Rows(b).Item(0))
                        Next
                        Exit Select
                End Select
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulateSecondaryLeadSource", "ByVal PrimaryLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetSLS")
            End Try
        End Sub
    End Class
    Public Class PopulatePLSandSLSbyMarketer
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub GetMarketerLeadSources(ByVal FName As String, ByVal LName As String)
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetMarketerLeadSources", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@FName", FName)
                Dim param2 As SqlParameter = New SqlParameter("@LName", LName)
                cmdGet.Parameters.Add(param1)
                cmdGet.Parameters.Add(param2)
                cnn.Open()
                cmdGet.CommandType = CommandType.StoredProcedure
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader
                While r1.Read

                    EnterLead.txtMarketingManager.Text = r1.Item(2)
                    EnterLead.cboPriLead.Text = ""
                    EnterLead.cboPriLead.Text = r1.Item(0)
                    'EnterLead.cboPriLead.Enabled = False
                    EnterLead.cboSecLead.Text = ""
                    EnterLead.cboSecLead.Text = r1.Item(1)
                    'EnterLead.cboSecLead.Enabled = False
                    EnterLead.txtC1FName.Focus()



                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulatePLSandSLSByMarketer", "ByVal FName As String, ByVal LName As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetMarketerLeadSources")

            End Try

        End Sub
    End Class
    Public Class PopulateCities
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub GetCities()
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetCities", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader
                EnterLead.txtCity.AutoCompleteCustomSource.Clear()
                While r1.Read
                    EnterLead.txtCity.AutoCompleteCustomSource.Add(r1.Item(0))
                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetCities")
            End Try


        End Sub
    End Class
    Public Class PopulateWorkHours
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub GetWorkHours()
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetWorkHours", cnn)
                cnn.Open()
                cmdGet.CommandType = CommandType.StoredProcedure
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader
                EnterLead.cboC1Work.Items.Clear()
                EnterLead.cboC2Work.Items.Clear()
                EnterLead.cboC2Work.Items.Add("<Add New>")
                EnterLead.cboC1Work.Items.Add("<Add New>")
                EnterLead.cboC1Work.Items.Add("_________________________________")
                EnterLead.cboC2Work.Items.Add("_________________________________")
                EnterLead.cboC2Work.Items.Add("")
                EnterLead.cboC1Work.Items.Add("")
                While r1.Read
                    EnterLead.cboC1Work.Items.Add(r1.Item(0))
                    EnterLead.cboC2Work.Items.Add(r1.Item(0))
           
                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulateWorkHours", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetWorkHours")
            End Try

        End Sub
        Public Sub InsertWH(ByVal str As String, ByVal cbo As ComboBox)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.WorkHours WHERE Hour = '" & str & "'", cnn)
                'Dim param1 As SqlParameter = New SqlParameter("@str", str)

                'cmdCNT.Parameters.Add(param1)

                cnn.Open()
                Dim cnt As Integer = 0
                Dim r1 As SqlDataReader
                r1 = cmdCNT.ExecuteReader
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is >= 1
                        MsgBox("Duplicate Work Hours Exists.Please try again.", MsgBoxStyle.Exclamation, "ERROR")
                        cbo.Text = ""
                        Exit Sub
                    Case Is < 1
                        Dim cmdINS As SqlCommand = New SqlCommand("Insert WorkHours (Hour) Values ('" & str & "')", cnn)

                        cmdINS.CommandType = CommandType.Text
                        cnn.Open()

                        cmdINS.ExecuteReader()
                        cnn.Close()
                        Me.GetWorkHours()

                        cbo.SelectedItem = str
                        Exit Select
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertSLS", "ByVal PLS As String, ByVal SLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertSLS")

            End Try
        End Sub
    End Class

    Public Class PopulateProducts
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub GetProducts()
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetProducts", cnn)
                Dim r1 As SqlDataReader
                cnn.Open()
                r1 = cmdGet.ExecuteReader
                EnterLead.cboProduct1.Items.Clear()
                EnterLead.cboProduct2.Items.Clear()
                EnterLead.cboProduct3.Items.Clear()
                EnterLead.cboProduct1.Items.Add("<Add New>")
                EnterLead.cboProduct2.Items.Add("<Add New>")
                EnterLead.cboProduct3.Items.Add("<Add New>")
                EnterLead.cboProduct1.Items.Add("_________________________")
                EnterLead.cboProduct2.Items.Add("_________________________")
                EnterLead.cboProduct3.Items.Add("_________________________")
                EnterLead.cboProduct1.Items.Add("")
                EnterLead.cboProduct2.Items.Add("")
                EnterLead.cboProduct3.Items.Add("")
                While r1.Read
                    EnterLead.cboProduct1.Items.Add(r1.Item(0))
                    EnterLead.cboProduct2.Items.Add(r1.Item(0))
                    EnterLead.cboProduct3.Items.Add(r1.Item(0))
                End While
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetProducts")

            End Try

        End Sub
    End Class
    Public Class InsertProduct
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub AddNewProduct(ByVal Product As String, ByVal ProdACRO As String, ByVal CBO As String)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT count(ID) from iss.dbo.products where Product = @Product", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@Product", Product)
                cmdCNT.Parameters.Add(param1)
                cnn.Open()
                Dim cnt As Integer = 0
                Dim r1 As SqlDataReader
                r1 = cmdCNT.ExecuteReader
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is < 1
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertProduct", cnn)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        Dim param2 As SqlParameter = New SqlParameter("@Product", Product)
                        Dim param3 As SqlParameter = New SqlParameter("@ProdACRO", ProdACRO)
                        cnn.Open()
                        cmdINS.Parameters.Add(param2)
                        cmdINS.Parameters.Add(param3)
                        Dim r2 As SqlDataReader
                        r2 = cmdINS.ExecuteReader
                        r2.Close()
                        cnn.Close()
                        Dim v As New PopulateProducts
                        v.GetProducts()
                        Select Case CBO
                            Case Is = "CBO1"
                                EnterLead.cboProduct1.SelectedItem = Product
                                Exit Select
                            Case Is = "CBO2"
                                EnterLead.cboProduct2.SelectedItem = Product
                                Exit Select
                            Case Is = "CBO3"
                                EnterLead.cboProduct3.SelectedItem = Product
                                Exit Select
                        End Select
                    Case Is >= 1
                        MsgBox("Duplicate Product exists. Please enter a new one.", MsgBoxStyle.Exclamation)
                        Exit Sub
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.AddNewProduct", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "AddNewProduct")
            End Try

        End Sub
    End Class
    Public Class InsertMarketer
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertMarketer(ByVal frm As Form)
            Try
                'Dim x As String = ""
                'Dim y As String = ""
                'x = InputBox$("Enter marketer's first name.", "New Marketer First Name")
                'If x.ToString.Length < 2 Then
                '    Exit Sub
                'End If
                'y = InputBox$("Enter marketer's last name.", "New Marketer Last Name")
                'If y.ToString.Length < 2 Then
                '    Exit Sub
                ''End If
                'x = CapitalizeText(x)
                'y = CapitalizeText(y)
                'MarketerLeadSources.MFName = x
                'MarketerLeadSources.MLName = y
                MarketerLeadSources.Frm = frm
                MarketerLeadSources.ShowInTaskbar = False
                MarketerLeadSources.ShowDialog()

            Catch ex As Exception
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertMarketer", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "InsertMarketer")
            End Try

        End Sub
        Public Sub WriteMarketerToTable(ByVal MFname As String, ByVal MLName As String, ByVal PLS As String, ByVal SLS As String, ByVal MarketingMan As String)
            Try
                Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertMarkter", cnn)
                cmdINS.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@Fname", Trim(MFname))
                Dim param2 As SqlParameter = New SqlParameter("@LName", Trim(MLName))
                Dim param3 As SqlParameter = New SqlParameter("@PriLeadSource", PLS)
                Dim param4 As SqlParameter = New SqlParameter("@SecLeadSource", SLS)
                Dim param5 As SqlParameter = New SqlParameter("@MarketingManager", MarketingMan)
                cmdINS.Parameters.Add(param1)
                cmdINS.Parameters.Add(param2)
                cmdINS.Parameters.Add(param3)
                cmdINS.Parameters.Add(param4)
                cmdINS.Parameters.Add(param5)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdINS.ExecuteReader
                r1.Close()
                cnn.Close()
                Dim c As New ENTER_LEAD.PopulatePrimaryLeadSource
                c.GetPrimaryLeadSource()
                Dim b As New ENTER_LEAD.PopulateMarketers
                b.GetMarketers()

                EnterLead.cboMarketer.SelectedItem = MFname & " " & MLName
                MarketerLeadSources.MFName = ""
                MarketerLeadSources.MLName = ""
                MarketerLeadSources.cboSLS.Text = ""
                MarketerLeadSources.cboPRILS.Text = ""
                MarketerLeadSources.cboMM.Text = ""
                MarketerLeadSources.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertMarketer", "ByVal MFname As String, ByVal MLName As String, ByVal PLS As String, ByVal SLS As String, ByVal MarketingMan As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "WriteMarketerToTable")

            End Try

        End Sub
        Public Function CapitalizeText(ByVal TextToCap As String)
            Try
                Dim Text
                Text = Mid(TextToCap, 1, 1)
                Text = Text.ToString.ToUpper
                TextToCap = Text & Mid(TextToCap, 2, TextToCap.Length)
                Return TextToCap
            Catch ex As Exception
                Return TextToCap
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertMarketer", "ByVal TextToCap as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CapitalizeText")

            End Try

        End Function


     
    End Class
    Public Class Captilalize
        Public Function CapitalizeText(ByVal TextToCap As String)
            Try
                Dim Text
                Text = Mid(TextToCap, 1, 1)
                Text = Text.ToString.ToUpper
                TextToCap = Text & Mid(TextToCap, 2, TextToCap.Length)
                Return TextToCap
            Catch ex As Exception
                Return TextToCap
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.Capitalize", "ByVal TextToCap as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CapitalizeText")

            End Try

        End Function
    End Class
    Public Class InsertPLS
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertNewPLS(ByVal PLS As String)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.PrimaryLeadSource WHERE PrimaryLead = @PLS", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
                Dim g As New Captilalize
                PLS = g.CapitalizeText(PLS)
                cmdCNT.Parameters.Add(param1)
                Dim cnt As Integer = 0
                Select Case cnt
                    Case Is >= 1
                        MsgBox("Duplicate Primary Lead Source Exists. Please enter a new one.", MsgBoxStyle.Exclamation, "ERROR")
                        Exit Sub
                    Case Is < 1
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertPLS", cnn)
                        Dim param2 As SqlParameter = New SqlParameter("@PLS", PLS)
                        cmdINS.Parameters.Add(param2)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        cnn.Open()
                        Dim r1 As SqlDataReader
                        r1 = cmdINS.ExecuteReader
                        r1.Close()
                        cnn.Close()
                        'EnterLead.cboPriLead.SelectedItem = PLS
                        Dim PopPrim As New PopulatePrimaryLeadSource
                        PopPrim.GetPrimaryLeadSource()
                        EnterLead.cboPriLead.SelectedItem = PLS
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertPLS", "ByVal PLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertNewPLS")

            End Try

        End Sub
    End Class
    Public Class InsertSLS
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertSLS(ByVal PLS As String, ByVal SLS As String)
            Try
                Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from iss.dbo.SecondaryLeadSource WHERE SecondaryLead = @SLS and PrimaryLead = @PLS", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@SLS", SLS)
                Dim param2 As SqlParameter = New SqlParameter("@PLS", PLS)
                cmdCNT.Parameters.Add(param1)
                cmdCNT.Parameters.Add(param2)
                cnn.Open()
                Dim cnt As Integer = 0
                Dim r1 As SqlDataReader
                r1 = cmdCNT.ExecuteReader
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is >= 1
                        MsgBox("Duplicate Seconadary Lead Source Exists.Please try again.", MsgBoxStyle.Exclamation, "ERROR")
                        Exit Sub
                    Case Is < 1
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertSLS", cnn)
                        Dim param3 As SqlParameter = New SqlParameter("@PrimaryLead", PLS)
                        Dim param4 As SqlParameter = New SqlParameter("@SecondaryLead", SLS)
                        cmdINS.Parameters.Add(param3)
                        cmdINS.Parameters.Add(param4)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        cnn.Open()
                        Dim r2 As SqlDataReader
                        r2 = cmdINS.ExecuteReader
                        r2.Close()
                        cnn.Close()
                        Dim PopSEC As New PopulateSecondaryLeadSource
                        PopSEC.GetSLS(PLS)
                        EnterLead.cboSecLead.SelectedItem = SLS
                        Exit Select
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertSLS", "ByVal PLS As String, ByVal SLS As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertSLS")

            End Try

        End Sub
    End Class
    Public Class CheckDuplicateLead
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub New(ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal CloseMethod As String, ByVal ForceMap As Boolean)
            '' due to new process changes, verify lead is actually fired off of the front end of enterlead
            '' BEFORE duplicate lead is checked.
            '' However, I am adding an argument to checkduplicate lead to force the @Mapped variable on or off
            ''
            'Dim c As New MappointUtilities
            'c.VerifyAddress(StAddress, City, State, Zip)
            Try
                Dim dset_Customers As Data.DataSet = New Data.DataSet("CUSTOMERS")
                Dim dColLeadNumber As Data.DataColumn = New Data.DataColumn("LeadNumber")
                Dim dColContact1FName As Data.DataColumn = New Data.DataColumn("Contact")
                Dim dColContactAddress As Data.DataColumn = New Data.DataColumn("Address")
                Dim dTab As Data.DataTable = New Data.DataTable("tblCustomers")
                dset_Customers.Tables.Add(dTab)
                dTab.Columns.Add(dColContact1FName)
                dTab.Columns.Add(dColContactAddress)
                dTab.Columns.Add(dColLeadNumber)
                Dim CmdGET As SqlCommand = New SqlCommand("SELECT ID, Contact1FirstName, Contact1LastName,  StAddress, City, State, Zip from iss.dbo.enterlead WHERE StAddress = @STA and City = @CTY and State = @STN and Zip = @ZP", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@STN", State)
                Dim param2 As SqlParameter = New SqlParameter("@STA", StAddress)
                Dim param3 As SqlParameter = New SqlParameter("@CTY", City)
                Dim param4 As SqlParameter = New SqlParameter("@ST", State)
                Dim param5 As SqlParameter = New SqlParameter("@ZP", Zip)
                CmdGET.Parameters.Add(param1)
                CmdGET.Parameters.Add(param2)
                CmdGET.Parameters.Add(param3)
                CmdGET.Parameters.Add(param4)
                CmdGET.Parameters.Add(param5)
                Dim r1 As SqlDataReader
                cnn.Open()
                r1 = CmdGET.ExecuteReader
                While r1.Read
                    'Dim r As Data.DataRow = dTab.NewRow
                    'r("Contact") = r1.Item(1) & " " & r1.Item(2)
                    'r("LeadNumber") = r1.Item(0)
                    'r("Address") = r1.Item("StAddress") & " " & r1.Item("City") & " " & r1.Item("State") & " " & r1.Item("Zip")
                    'dTab.Rows.Add(r)
                End While
                r1.Close()
                cnn.Close()
                Dim cnt As Integer = 0
                cnt = dset_Customers.Tables(0).Rows.Count
                Select Case cnt
                    Case Is >= 1
                        Dim g As Integer = 0
                        For g = 0 To dset_Customers.Tables(0).Rows.Count - 1
                            Dim b As New ListViewItem
                            b.Text = dset_Customers.Tables(0).Rows(g).Item("LeadNumber")
                            b.SubItems.Add(dset_Customers.Tables(0).Rows(g).Item("Contact"))
                            b.SubItems.Add(dset_Customers.Tables(0).Rows(g).Item("Address"))
                            DuplicateLead.lstDupes.Items.Add(b)
                        Next
                        DuplicateLead.ShowInTaskbar = False
                        DuplicateLead.ShowDialog()
                        Exit Select
                    Case Is <= 0
                        '' just write the lead here.
                        Dim ins As New InsertEnterLead
                        ins.InsertLead(ForceMap) '' this variable is changed by the argument ForceMap
                        Select Case CloseMethod
                            Case Is = "SAVE"
                                EnterLead.Reset()
                                EnterLead.Close()
                                Exit Select
                            Case Is = "SAVEANDNEW"
                                EnterLead.Reset()
                                Exit Select
                        End Select
                        Exit Select
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.CheckDuplicateLead", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal CloseMethod As String, ByVal ForceMap As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "'New'")

            End Try

        End Sub
    End Class
#Region "ANDY_DUPLICATE_LEADS"
    Public Class PopulateDuplicates

        Private Department As String
        Dim ttStatus As New ToolTip
        Dim ttNotes As New ToolTip
        Private notes As String
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)


        Public Sub SetUp(ByVal C1FName As String, ByVal C1LName As String, ByVal C2FName As String, ByVal C2LName As String, ByVal MainPhone As String, ByVal AltPhone1 As String, ByVal AltPhone2 As String, ByVal STAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal CloseMethod As String, ByVal ForceMap As Boolean)
            Try
                Dim switch As String = ""
                DuplicateRecord.pnlDuplicates.Controls.Clear()
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.CHKDuplicates", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@Contact1FirstName", C1FName)
                Dim param2 As SqlParameter = New SqlParameter("@Contact1LastName", C1LName)
                Dim param3 As SqlParameter = New SqlParameter("@Contact2FirstName", C2FName)
                Dim param4 As SqlParameter = New SqlParameter("@Contact2LastName", C2LName)
                Dim param5 As SqlParameter = New SqlParameter("@HousePhone", MainPhone)
                Dim param6 As SqlParameter = New SqlParameter("@AltPhone1", AltPhone1)
                Dim param7 As SqlParameter = New SqlParameter("@AltPhone2", AltPhone2)
                Dim param8 As SqlParameter = New SqlParameter("@StAddress", STAddress)
                Dim param9 As SqlParameter = New SqlParameter("@City", City)
                Dim param10 As SqlParameter = New SqlParameter("@State", State)
                Dim param11 As SqlParameter = New SqlParameter("@Zip", Zip)
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
                cnn.Open()
                Dim R1 As SqlDataReader = cmdGet.ExecuteReader
                Dim b As Boolean = True
                switch = CloseMethod
                While R1.Read
                    If R1.Item(0) = 0 Then
                        Dim ins As New InsertEnterLead
                        ins.InsertLead(ForceMap) '' this variable is changed by the argument ForceMap

                        If CloseMethod = "SAVEANDNEW" Then
                            ''Kickback to save and new '' No duplicates
                            'MsgBox(R1.Item(0))
                            'switch = "SAVEANDNEW"
                            EnterLead.Reset()

                        ElseIf CloseMethod = "SAVE" Then
                            ''Kickback to save and Close '' No duplicates
                            'switch = "SAVE"
                            EnterLead.Reset()
                            EnterLead.Close()
                        End If
                        Exit Sub
                    End If
                    Dim p As New Panel
                    p.Dock = DockStyle.Top
                    p.BorderStyle = BorderStyle.None
                    p.Size = New System.Drawing.Size(322, 86)
                    p.Name = "pnl" & R1.Item(0)
                    p.Cursor = Cursors.Hand
                    AddHandler p.Click, AddressOf Panel
                    If b = True Then
                        p.BackColor = Color.WhiteSmoke
                        b = False
                    ElseIf b = False Then
                        p.BackColor = Color.White
                        b = True
                    End If

                    DuplicateRecord.pnlDuplicates.Controls.Add(p)

                    Dim lblID As New Label
                    lblID.Text = "ID: " & R1.Item(0)
                    lblID.Location = New System.Drawing.Point(130, 5)
                    lblID.Size = New System.Drawing.Size(80, 13)
                    lblID.Font = New System.Drawing.Font("Tahoma", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblID.Cursor = Cursors.Hand
                    AddHandler lblID.Click, AddressOf Labels

                    Dim lblContactInfo As New Label
                    If R1.Item(3) <> "" Then
                        lblContactInfo.Text = R1.Item(1) & " " & R1.Item(2) & _
                        vbCr & R1.Item(3) & " " & R1.Item(4) & _
                        vbCr & R1.Item(8) & vbCr & R1.Item(9) & ", " & R1.Item(10) & " " _
                        & R1.Item(11)
                    Else
                        lblContactInfo.Text = R1.Item(1) & " " & R1.Item(2) & _
                        vbCr & R1.Item(8) & vbCr & R1.Item(9) & ", " & R1.Item(10) & " " _
                        & R1.Item(11)
                    End If
                    lblContactInfo.Name = "lblContactInfo" & R1.Item(0)
                    lblContactInfo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblContactInfo.Location = New System.Drawing.Point(12, 25)
                    lblContactInfo.Size = New System.Drawing.Size(145, 52)
                    lblContactInfo.AutoEllipsis = True
                    lblContactInfo.Cursor = Cursors.Hand
                    AddHandler lblContactInfo.Click, AddressOf Labels

                    Dim lblPhone As New Label
                    If R1.Item(6) = "" And R1.Item(7) = "" Then
                        lblPhone.Text = "Main Phone: " & R1.Item(5)
                    ElseIf R1.Item(6) <> "" And R1.Item(7) = "" Then
                        lblPhone.Text = "Main Phone: " & R1.Item(5) & vbCr & "Alt Phone 1: " & R1.Item(6)
                    ElseIf R1.Item(6) <> "" And R1.Item(7) <> "" Then
                        lblPhone.Text = "Main Phone: " & R1.Item(5) & vbCr & "Alt Phone 1: " & R1.Item(6) _
                        & vbCr & "Alt Phone 2: " & R1.Item(7)
                    ElseIf R1.Item(6) = "" And R1.Item(7) <> "" Then
                        lblPhone.Text = "Main Phone: " & R1.Item(5) & vbCr & "Alt Phone 2: " & R1.Item(7)
                    End If
                    lblPhone.Name = "lblPhone" & R1.Item(0)
                    lblPhone.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                    lblPhone.Location = New System.Drawing.Point(177, 25)
                    lblPhone.Anchor = AnchorStyles.Top Or AnchorStyles.Right
                    lblPhone.AutoEllipsis = True
                    lblPhone.Size = New System.Drawing.Size(150, 52)
                    lblPhone.Cursor = Cursors.Hand
                    AddHandler lblPhone.Click, AddressOf Labels

                    ''Add Controls
                    p.Controls.Add(lblID)
                    p.Controls.Add(lblContactInfo)
                    p.Controls.Add(lblPhone)
                    DuplicateRecord.pnlDuplicates.Controls.Add(p)
                End While
                R1.Close()
                cnn.Close()

                DuplicateRecord.MapPointVerified = ForceMap
                DuplicateRecord.CloseMethod = switch
                DuplicateRecord.ShowInTaskbar = False
                DuplicateRecord.StartPosition = FormStartPosition.CenterScreen
                DuplicateRecord.ShowDialog()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.PopulateDuplicates", "ByVal C1FName As String, ByVal C1LName As String, ByVal C2FName As String, ByVal C2LName As String, ByVal MainPhone As String, ByVal AltPhone1 As String, ByVal AltPhone2 As String, ByVal STAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal CloseMethod As String, ByVal ForceMap As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "SetUp")

            End Try


        End Sub
        Public Sub Panel(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim who As Panel = sender

            Dim c As Integer = DuplicateRecord.pnlDuplicates.Controls.Count
            Dim i As Integer
            For i = 1 To c
                Dim all As Panel = DuplicateRecord.pnlDuplicates.Controls(i - 1)
                If all.BorderStyle = BorderStyle.FixedSingle Then
                    all.BorderStyle = BorderStyle.None

                    'Else
                    'all.BorderStyle = BorderStyle.FixedSingle
                End If
            Next
            who.BorderStyle = BorderStyle.FixedSingle

            DuplicateRecord.btnUpdate.Text = ""
            DuplicateRecord.btnUpdate.Text = "Update Record: " & Right(who.Name.ToString, who.Name.ToString.Length - 3)
            DuplicateRecord.ID = Right(who.Name.ToString, who.Name.ToString.Length - 3)
            DuplicateRecord.btnUpdate.Enabled = True
        End Sub

        Private Sub Labels(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Dim lbl As Label = sender
            Panel(lbl.Parent, Nothing)
        End Sub


    End Class
#End Region


    Public Class MappointUtilities
        Implements IDisposable
        Public MapPointVerified As Boolean
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        'Public Sub VerifyAddress(ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String)
        '    Try
        '        Dim oApp As MapPoint.Application = New MapPoint.Application
        '        Dim oMap As MapPoint.Map = oApp.ActiveMap
        '        Dim oResults As MapPoint.FindResults
        '        oResults = oMap.FindAddressResults(StAddress, City, , State, Zip)
        '        Dim CntResults As Integer = oResults.Count
        '        Select Case CntResults
        '            Case Is = 1
        '                '' No matter what here, you need to check to see if the city exists on our table
        '                '' or not, If it doesn't, add it. If it does, Skip step.
        '                Dim cty As New IsCityInTable
        '                cty.CheckCity(City, State)
        '                Dim g As New VerifyAddress(StAddress, City, State, Zip)
        '                '' return @Mapped value here to store on table
        '                ''
        '                MapPointVerified = g.MapPointVerified
        '            Case Is > 1
        '                '' same as above
        '                ''
        '                Dim cty As New IsCityInTable
        '                cty.CheckCity(City, State)
        '                Dim g As New VerifyAddress(StAddress, City, State, Zip)
        '                '' return @Mapped value here as well to store on table
        '                ''
        '                MapPointVerified = g.MapPointVerified
        '        End Select
        '        '' kill process method is overriding these. just comment them out.
        '        ''
        '        'oMap.Saved = True
        '        'oApp.Quit()
        '        Me.Dispose()
        '    Catch ex As Exception
        '        Dim err As New ErrorLogFlatFile
        '        err.WriteLog("ENTER_LEAD.MappointUtilities", "ByVal StAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Mappoint", "VerifiyAddress")
        '    End Try

        'End Sub
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free managed resources when explicitly called
                End If

                ' TODO: free shared unmanaged resources
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

   
    End Class
    Public Class InsertEnterLead
        Public CorrectPhoneNumberResponse As String = ""
        Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Public Sub InsertLead(ByVal MapPointVerified As Boolean)
            Try
                Dim y = EnterLead.dtpApptInfo.Value.ToString
                Dim x = InStr(y, " ")
                y = Microsoft.VisualBasic.Left(y, x - 1)
                y = y & " 12:00:00 AM"
                Dim a = EnterLead.dtpLeadGen.Value.ToString
                Dim b = InStr(a, " ")
                a = Microsoft.VisualBasic.Left(a, b - 1)
                a = a & " 12:00:00 AM"

                Dim Description As String
                Dim ID

                If EnterLead.cboMarketer.Text = "" And EnterLead.cboSecLead.Text = "" Then
                    Description = "Appt. Generated from " & EnterLead.cboPriLead.Text & ", set up with " & EnterLead.cboSpokeWith.Text
                ElseIf EnterLead.cboMarketer.Text <> "" And EnterLead.cboSecLead.Text = "" Then
                    Description = "Appt. Generated from " & EnterLead.cboPriLead.Text & " by " & EnterLead.cboMarketer.Text & ", set up with " & EnterLead.cboSpokeWith.Text
                ElseIf EnterLead.cboMarketer.Text <> "" And EnterLead.cboSecLead.Text <> "" Then
                    Description = "Appt. Generated from " & EnterLead.cboPriLead.Text & "-" & EnterLead.cboSecLead.Text & " by " & EnterLead.cboMarketer.Text & ", set up with " & EnterLead.cboSpokeWith.Text
                ElseIf EnterLead.cboMarketer.Text = "" And EnterLead.cboSecLead.Text <> "" Then
                    Description = "Appt. Generated From " & EnterLead.cboPriLead.Text & "-" & EnterLead.cboSecLead.Text & ", set up with " & EnterLead.cboSpokeWith.Text
                End If

                Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertEnterLead", cnn)
                cmdINS.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@Marketer", Trim(EnterLead.cboMarketer.Text))
                Dim param2 As SqlParameter = New SqlParameter("@PLS", Trim(EnterLead.cboPriLead.Text))
                Dim param3 As SqlParameter = New SqlParameter("@SLS", Trim(EnterLead.cboSecLead.Text))
                Dim param4 As SqlParameter = New SqlParameter("@LeadGenOn", a)
                Dim param5 As SqlParameter = New SqlParameter("@Contact1FirstName", Trim(EnterLead.txtC1FName.Text))
                Dim param6 As SqlParameter = New SqlParameter("@Contact1LastName", Trim(EnterLead.txtC1LName.Text))
                Dim param7 As SqlParameter = New SqlParameter("@Contact2FirstName", Trim(EnterLead.txtC2FName.Text))
                Dim param69 As SqlParameter = New SqlParameter("@Contact2LastName", Trim(EnterLead.txtC2LName.Text))
                Dim param8 As SqlParameter = New SqlParameter("@YearsOwned", EnterLead.txtYearsOwned.Text)
                Dim param9 As SqlParameter = New SqlParameter("@HomeAge", Trim(EnterLead.txtAgeOfHome.Text))
                Dim param10 As SqlParameter = New SqlParameter("@HomeValue", EnterLead.txtHomeVal.Text)
                Dim param11 As SqlParameter = New SqlParameter("@SpecialInstruction", EnterLead.rtfSpecialIns.Text)

                '' strip off literals
                '' '(' ')' '-'
                '' -------------------------------------------------------------------      
                DoLiteralsExist(EnterLead.txtHousePhone.Text)
                Dim param12 As SqlParameter = New SqlParameter("@HousePhone", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(EnterLead.txtAltPhone1.Text)
                Dim param13 As SqlParameter = New SqlParameter("@AltPhone1", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(EnterLead.txtAltPhone2.Text)
                Dim param14 As SqlParameter = New SqlParameter("@AltPhone2", Me.CorrectPhoneNumberResponse.ToString)
                '' -----------------------------------------------------------------------------------

                Dim param15 As SqlParameter = New SqlParameter("@Alt1Type", EnterLead.cboAlt1Type.Text)
                Dim param16 As SqlParameter = New SqlParameter("@Alt2Type", EnterLead.cboAltPhone2Type.Text)
                'Dim param17 As SqlParameter = New SqlParameter("@StNumber", EnterLead.txtStNum.Text)
                Dim stAddress As String = ""
                Dim param18 As SqlParameter = New SqlParameter("@StAddress", Trim(EnterLead.txtStAddy.Text))
                Dim param19 As SqlParameter = New SqlParameter("@City", Trim(EnterLead.txtCity.Text))
                Dim param20 As SqlParameter = New SqlParameter("@State", EnterLead.txtState.Text)
                Dim param21 As SqlParameter = New SqlParameter("@Zip", EnterLead.txtZip.Text)
                Dim param22 As SqlParameter = New SqlParameter("@SpokeWith", EnterLead.cboSpokeWith.Text)
                Dim param23 As SqlParameter = New SqlParameter("@C1Work", EnterLead.cboC1Work.Text)
                Dim param24 As SqlParameter = New SqlParameter("@C2Work", EnterLead.cboC2Work.Text)
                Dim param25 As SqlParameter = New SqlParameter("@AppDate", y)
                Dim param26 As SqlParameter = New SqlParameter("@AppDay", EnterLead.txtApptday.Text)
                Dim param27 As SqlParameter = New SqlParameter("@AppTime", EnterLead.txtApptTime.Text)
                Dim param28 As SqlParameter = New SqlParameter("@Product1", EnterLead.cboProduct1.Text)
                Dim param29 As SqlParameter = New SqlParameter("@Product2", EnterLead.cboProduct2.Text)
                Dim param30 As SqlParameter = New SqlParameter("@Product3", EnterLead.cboProduct3.Text)
                Dim param31 As SqlParameter = New SqlParameter("@Color", EnterLead.txtProdColor.Text)
                Dim param32 As SqlParameter = New SqlParameter("@QTY", EnterLead.txtProdQTY.Text)
                Dim param33 As SqlParameter = New SqlParameter("@Email", EnterLead.txtEmail.Text)
                Dim param34 As SqlParameter = New SqlParameter("@ProdAcro1", EnterLead.txtP1acro.Text)
                Dim param35 As SqlParameter = New SqlParameter("@ProdAcro2", EnterLead.txtP2acro.Text)
                Dim param36 As SqlParameter = New SqlParameter("@prodAcro3", EnterLead.txtP3acro.Text)
                Dim param37 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser.ToString)
                Dim param38 As SqlParameter = New SqlParameter("@MarketingManager", EnterLead.txtMarketingManager.Text)
                Dim param39 As SqlParameter = New SqlParameter("@Description", Description)
                Dim param70 As SqlParameter = New SqlParameter("@Mapped", MapPointVerified)
                cnn.Open()
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
                Dim r1 As SqlDataReader
                r1 = cmdINS.ExecuteReader
                r1.Close()
                cnn.Close()
                '' then call whatever requery method if needed here.
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.InsertEnterLead", "ByVal MapPointVerified As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertLead")

            End Try

        End Sub
        Public Sub Edit_Lead(ByVal ID As String)

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

    
    End Class
    Public Class IsCityInTable
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub CheckCity(ByVal City As String, ByVal State As String)
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("dbo.CountCity", cnn)
                cmdGet.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@CTY", City)
                cmdGet.Parameters.Add(param1)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader
                Dim cnt As Integer = 0
                While r1.Read
                    cnt = r1.Item(0)
                End While
                r1.Close()
                cnn.Close()
                Select Case cnt
                    Case Is >= 1
                        Exit Sub
                    Case Is <= 0
                        Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertNewCity", cnn)
                        cmdINS.CommandType = CommandType.StoredProcedure
                        Dim param2 As SqlParameter = New SqlParameter("@CITY", City)
                        Dim param3 As SqlParameter = New SqlParameter("@STATE", State)
                        cmdINS.Parameters.Add(param2)
                        cmdINS.Parameters.Add(param3)
                        cnn.Open()
                        Dim r2 As SqlDataReader
                        r2 = cmdINS.ExecuteReader(CommandBehavior.CloseConnection)
                        r2.Close()
                        cnn.Close()
                        EnterLead.txtCity.AutoCompleteCustomSource.Add(City)
                        Exit Select
                End Select
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.IsCityInTable", "ByVal City As String, ByVal State As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "CheckCity")
            End Try

        End Sub
    

    End Class
    Public Sub AutoFillState(ByVal city)
        Try
            Dim cmdGETState As SqlCommand = New SqlCommand("SELECT State from iss.dbo.CityPull where City = @city", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@city", city)
            cnn.Open()
            cmdGETState.Parameters.Add(param1)
            Dim r1 As SqlDataReader
            r1 = cmdGETState.ExecuteReader
            While r1.Read
                EnterLead.txtState.Text = r1.Item(0)
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Public Class UpdateEnterLead
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public CorrectPhoneNumberResponse As String = ""
        Public Sub New(ByVal ID As String, ByVal Marketer As String, ByVal PLS As String, ByVal SLS As String, ByVal LeadGenOn As Date, _
            ByVal Contact1FirstName As String, ByVal Contact1LastName As String, ByVal Contact2FirstName As String, ByVal Contact2LastName As String, _
            ByVal YearsOwned As String, ByVal HomeAge As String, ByVal HomeValue As String, ByVal SpecialInstruction As String, ByVal HousePhone As String, _
            ByVal AltPhone1 As String, ByVal AltPhone2 As String, ByVal Alt1Type As String, ByVal Alt2Type As String, ByVal StAddress As String, _
            ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal SpokeWith As String, ByVal C1Work As String, _
            ByVal C2Work As String, ByVal ApptDate As String, ByVal ApptTime As String, ByVal ApptDay As String, ByVal Product1 As String, ByVal Product2 As String, ByVal Product3 As String, _
            ByVal Color As String, ByVal QTY As String, ByVal Email As String, ByVal Prod1Acro As String, ByVal Prod2Acro As String, _
            ByVal Prod3Acro As String, ByVal User As String, ByVal MarketingManager As String, ByVal Description As String, ByVal Mapped As Boolean)
            '' Enclose these variables with Trim() statement to make sure no spaces come through
            ''
            Try
                Dim cmdUp As SqlCommand = New SqlCommand("dbo.UpdateEnterLead", cnn)
                cmdUp.CommandType = CommandType.StoredProcedure
                cnn.Open()
                Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
                Dim param2 As SqlParameter = New SqlParameter("@Marketer", Marketer)
                Dim param3 As SqlParameter = New SqlParameter("@PLS", PLS)
                Dim param4 As SqlParameter = New SqlParameter("@SLS", SLS)

                '' datetime variable 1
                Dim param5 As SqlParameter = New SqlParameter("@LeadGenOn", LeadGenOn)

                Dim param6 As SqlParameter = New SqlParameter("@Contact1FirstName", Contact1FirstName)
                Dim param7 As SqlParameter = New SqlParameter("@Contact1LastName", Contact1LastName)
                Dim param8 As SqlParameter = New SqlParameter("@Contact2FirstName", Contact2FirstName)
                Dim param50 As SqlParameter = New SqlParameter("@Contact2LastName", Contact2LastName)
                Dim param9 As SqlParameter = New SqlParameter("@YearsOwned", YearsOwned)
                Dim param10 As SqlParameter = New SqlParameter("@HomeAge", HomeAge)
                Dim param51 As SqlParameter = New SqlParameter("@HomeValue", HomeValue)
                Dim param11 As SqlParameter = New SqlParameter("@SpecialInstruction", SpecialInstruction)

                '' if no number exists strip off literal
                '' 
                '' -------------------------------------------------------------------
                ''
                DoLiteralsExist(HousePhone)
                Dim param12 As SqlParameter = New SqlParameter("@HousePhone", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(AltPhone1)
                Dim param13 As SqlParameter = New SqlParameter("@AltPhone1", Me.CorrectPhoneNumberResponse.ToString)

                DoLiteralsExist(AltPhone2)
                Dim param14 As SqlParameter = New SqlParameter("@AltPhone2", Me.CorrectPhoneNumberResponse.ToString)

                '-----------------------------------------------------------------------

                'Dim param12 As SqlParameter = New SqlParameter("@HousePhone", HousePhone)
                'Dim param13 As SqlParameter = New SqlParameter("@AltPhone1", AltPhone1)
                'Dim param14 As SqlParameter = New SqlParameter("@AltPhone2", AltPhone2)

                Dim param15 As SqlParameter = New SqlParameter("@Alt1Type", Alt1Type)
                Dim param16 As SqlParameter = New SqlParameter("@Alt2Type", Alt2Type)
                Dim param52 As SqlParameter = New SqlParameter("@StAddress", StAddress)
                Dim param17 As SqlParameter = New SqlParameter("@City", City)
                Dim param18 As SqlParameter = New SqlParameter("@State", State)
                Dim param19 As SqlParameter = New SqlParameter("@Zip", Zip)
                Dim param20 As SqlParameter = New SqlParameter("@SpokeWith", SpokeWith)
                Dim param21 As SqlParameter = New SqlParameter("@C1Work", C1Work)
                Dim param22 As SqlParameter = New SqlParameter("@C2Work", C2Work)

                '' datetime variable 2
                Dim param23 As SqlParameter = New SqlParameter("@AppDate", ApptDate)

                Dim param24 As SqlParameter = New SqlParameter("@AppDay", ApptDay)


                '' dateteim variable 3
                Dim param25 As SqlParameter = New SqlParameter("@AppTime", ApptTime) '' // FIX

                Dim param26 As SqlParameter = New SqlParameter("@Product1", Product1)
                'Dim param54 As SqlParameter = New SqlParameter("@Prod1Acro", Prod1Acro)
                Dim param27 As SqlParameter = New SqlParameter("@Product2", Product2)
                'Dim param55 As SqlParameter = New SqlParameter("@Prod2Acro", Prod2Acro)
                Dim param28 As SqlParameter = New SqlParameter("@Product3", Product3)
                'Dim param56 As SqlParameter = New SqlParameter("@Prod3Acro", Prod3Acro)
                Dim param29 As SqlParameter = New SqlParameter("@Color", Color)
                Dim param30 As SqlParameter = New SqlParameter("@Email", Email)
                Dim param31 As SqlParameter = New SqlParameter("@ProdAcro1", Prod1Acro)
                Dim param32 As SqlParameter = New SqlParameter("@ProdAcro2", Prod2Acro)
                Dim param33 As SqlParameter = New SqlParameter("@ProdAcro3", Prod3Acro)
                Dim param34 As SqlParameter = New SqlParameter("@User", User)
                Dim param53 As SqlParameter = New SqlParameter("@QTY", QTY)
                Dim param35 As SqlParameter = New SqlParameter("@MarketingManager", MarketingManager)
                Dim param36 As SqlParameter = New SqlParameter("@Description", Description)
                Dim param37 As SqlParameter = New SqlParameter("@Mapped", Mapped)
                cmdUp.Parameters.Add(param1)
                cmdUp.Parameters.Add(param2)
                cmdUp.Parameters.Add(param3)
                cmdUp.Parameters.Add(param4)
                cmdUp.Parameters.Add(param5)
                cmdUp.Parameters.Add(param6)
                cmdUp.Parameters.Add(param7)
                cmdUp.Parameters.Add(param8)
                cmdUp.Parameters.Add(param9)
                cmdUp.Parameters.Add(param10)
                cmdUp.Parameters.Add(param11)
                cmdUp.Parameters.Add(param12)
                cmdUp.Parameters.Add(param13)
                cmdUp.Parameters.Add(param14)
                cmdUp.Parameters.Add(param15)
                cmdUp.Parameters.Add(param16)
                cmdUp.Parameters.Add(param17)
                cmdUp.Parameters.Add(param18)
                cmdUp.Parameters.Add(param19)
                cmdUp.Parameters.Add(param20)
                cmdUp.Parameters.Add(param21)
                cmdUp.Parameters.Add(param22)
                cmdUp.Parameters.Add(param23)
                cmdUp.Parameters.Add(param24)
                cmdUp.Parameters.Add(param25)
                cmdUp.Parameters.Add(param26)
                cmdUp.Parameters.Add(param27)
                cmdUp.Parameters.Add(param28)
                cmdUp.Parameters.Add(param29)
                cmdUp.Parameters.Add(param30)
                cmdUp.Parameters.Add(param31)
                cmdUp.Parameters.Add(param32)
                cmdUp.Parameters.Add(param33)
                cmdUp.Parameters.Add(param34)
                cmdUp.Parameters.Add(param35)
                cmdUp.Parameters.Add(param36)
                cmdUp.Parameters.Add(param37)
                '' Fix variables // stuff i forgot and error's are catching
                ''
                cmdUp.Parameters.Add(param50)
                cmdUp.Parameters.Add(param51)
                cmdUp.Parameters.Add(param52)
                cmdUp.Parameters.Add(param53)
                'cmdUp.Parameters.Add(param54)
                'cmdUp.Parameters.Add(param55)
                'cmdUp.Parameters.Add(param56)

                Dim r1 As SqlDataReader
                r1 = cmdUp.ExecuteReader
                r1.Close()
                cnn.Close()
            Catch ex As Exception
                cnn.Close()
                Dim err As New ErrorLogFlatFile
                err.WriteLog("ENTER_LEAD.UpdateEnterLead", "ByVal ID As String, ByVal Marketer As String, ByVal PLS As String, ByVal SLS As String, ByVal LeadGenOn As Date," _
         & "   ByVal Contact1FirstName As String, ByVal Contact1LastName As String, ByVal Contact2FirstName As String, ByVal Contact2LastName As String," _
            & "  ByVal YearsOwned As String, ByVal HomeAge As String, ByVal HomeValue As String, ByVal SpecialInstruction As String, ByVal HousePhone As String," _
           & " ByVal AltPhone1 As String, ByVal AltPhone2 As String, ByVal Alt1Type As String, ByVal Alt2Type As String, ByVal StAddress As String," _
           & " ByVal City As String, ByVal State As String, ByVal Zip As String, ByVal SpokeWith As String, ByVal C1Work As String," _
           & " ByVal C2Work As String, ByVal ApptDate As String, ByVal ApptTime As String, ByVal ApptDay As String, ByVal Product1 As String, ByVal Product2 As String, ByVal Product3 As String," _
           & " ByVal Color As String, ByVal QTY As String, ByVal Email As String, ByVal Prod1Acro As String, ByVal Prod2Acro As String," _
           & " ByVal Prod3Acro As String, ByVal User As String, ByVal MarketingManager As String, ByVal Description As String, ByVal Mapped As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "'New'")
            End Try
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
    End Class
End Class
