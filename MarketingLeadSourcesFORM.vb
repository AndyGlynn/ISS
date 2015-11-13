Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System
Public Class MarketerLeadSources
    Public MFName As String = ""
    Public MLName As String = ""
    Public Frm As Form

    Private Sub MarketerLeadSources_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim c As New PopulateLeadSources
        c.PopulatePrimary()
        'Me.cboMM.Enabled = False
        Me.ErrorProvider1.Clear()
        Dim g As New PopulateMarketingManager
        g.PopulateMarketingMAN()

    End Sub

    Private Sub cboPRILS_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPRILS.LostFocus
        Dim x As String = ""
        x = Me.cboPRILS.Text
        If x.ToString.Length < 2 Then
            Me.cboSLS.Text = ""
            Me.cboSLS.Items.Clear()
            'Me.cboMM.Enabled = False
        End If
        If x = "<Add New>" Then
            Me.cboSLS.Items.Clear()
        End If
        If x = "_____________________________________________________" Then
            x = ""

        End If
        If x = "" Then
            Me.cboSLS.Text = ""
            Me.cboSLS.Items.Clear()
            'Me.cboMM.Enabled = False
        End If
    End Sub

    Private Sub cboPRILS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPRILS.SelectedValueChanged
        Dim x As String = ""
        x = Me.cboPRILS.Text
        Select Case x
            Case Is = "<Add New>"
                'Me.cboMM.Enabled = False
                Dim c As New InsertNewPLS
                Dim cap As New Captilalize
                Dim pls As String = ""
                pls = cap.CapitalizeText(InputBox$("Enter new Primary Lead Source", "New Primary Lead Source"))
                If pls.ToString.Length < 2 Then
                    Me.cboPRILS.SelectedItem = ""
                    Me.cboSLS.Items.Clear()
                    Exit Sub
                End If
                c.InsertPLS(pls)
                Me.cboPRILS.SelectedItem = pls
                Exit Select
            Case Is = "_____________________________________________________"
                Me.cboPRILS.Text = ""
                'Me.cboMM.Enabled = False
                Exit Select
            Case Is = ""
                'Me.cboMM.Enabled = False
                Exit Select
            Case Else
                Me.cboMM.Enabled = True
                Me.cboSLS.Text = ""
                Dim c As New PopulateLeadSources
                c.PopulateSecondary(x)
                
                Exit Select
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.cboMM.Text = ""
        Me.MFName = ""
        Me.MLName = ""
        'Me.cboMM.Enabled = False
        Me.cboMM.Items.Clear()
        Me.cboPRILS.Items.Clear()
        Me.cboPRILS.Text = ""
        Me.cboSLS.Items.Clear()
        Me.cboSLS.Text = ""
        Me.cboMM.Items.Clear()
        'Me.cboMM.Enabled = False
        Me.ErrorProvider1.Clear()
        Me.Close()
        Select Case Frm.Name
            Case Is = "EnterLead"
                EnterLead.cboMarketer.SelectedItem = ""
            Case "EditCustomerInfo"
                EditCustomerInfo.cboMarketer.SelectedItem = ""
        End Select
    End Sub
#Region "Private Classes"

    Private Class PopulateLeadSources
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub PopulatePrimary()
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetPLS", cnn)
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGet.ExecuteReader
            MarketerLeadSources.cboPRILS.Items.Clear()
            MarketerLeadSources.cboSLS.Items.Clear()
            MarketerLeadSources.cboSLS.Text = ""
            MarketerLeadSources.cboPRILS.Items.Add("<Add New>")
            MarketerLeadSources.cboPRILS.Items.Add("_____________________________________________________")
            MarketerLeadSources.cboPRILS.Items.Add("")
            While r1.Read
                MarketerLeadSources.cboPRILS.Items.Add(r1.Item(1))
            End While
            r1.Close()
            cnn.Close()
        End Sub
        Public Sub PopulateSecondary(ByVal Primary As String)
            Dim cmdget As SqlCommand = New SqlCommand("dbo.GetSLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@PRILS", Primary)
            cmdget.CommandType = CommandType.StoredProcedure
            cmdget.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            MarketerLeadSources.cboSLS.Items.Clear()
            MarketerLeadSources.cboSLS.Items.Add("<Add New>")
            MarketerLeadSources.cboSLS.Items.Add("________________________________________________________")
            MarketerLeadSources.cboSLS.Items.Add("")
            r1 = cmdget.ExecuteReader
            While r1.Read
                MarketerLeadSources.cboSLS.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()
        End Sub
    End Class
    Private Class PopulateMarketingManager
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub PopulateMarketingMAN()
            Dim cmdGet As SqlCommand = New SqlCommand("dbo.GetMarketingManager", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            cnn.Open()
            r1 = cmdGet.ExecuteReader
            MarketerLeadSources.cboMM.Items.Clear()
            MarketerLeadSources.cboMM.Items.Add("<Add New>")
            MarketerLeadSources.cboMM.Items.Add("________________________________________________________")
            MarketerLeadSources.cboMM.Items.Add("")
            While r1.Read
                MarketerLeadSources.cboMM.Items.Add(r1.Item(0) & " " & r1.Item(1))
            End While
            r1.Close()
            cnn.Close()
        End Sub
    End Class
    Private Class InsertNewPLS
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertPLS(ByVal PLS As String)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertPLS", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@PLS", PLS)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()
            Dim g As New PopulateLeadSources
            g.PopulatePrimary()
            MarketerLeadSources.cboPRILS.SelectedItem = PLS
        End Sub
    End Class
    Private Class InsertNewManager
        Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Public Sub InsertManager(ByVal FName As String, ByVal LName As String)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertMarketingManager", cnn)
            Dim param1 As SqlParameter = New SqlParameter("Fname", FName)
            Dim param2 As SqlParameter = New SqlParameter("Lname", LName)
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            cmdINS.CommandType = CommandType.StoredProcedure
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()
            Dim g As New PopulateMarketingManager
            g.PopulateMarketingMAN()
            MarketerLeadSources.cboMM.SelectedItem = FName & " " & LName
        End Sub
    End Class
    Private Class Captilalize
        Public Function CapitalizeText(ByVal TextToCap As String)
            Dim Text
            Text = Mid(TextToCap, 1, 1)
            Text = Text.ToString.ToUpper
            TextToCap = Text & Mid(TextToCap, 2, TextToCap.Length)
            Return TextToCap
        End Function
    End Class
    Private Class InsertNewSLS
        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Public Sub InsertSLS(ByVal PLS As String, ByVal SLS As String)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertSLS", cnn)
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@PrimaryLead", PLS)
            Dim param2 As SqlParameter = New SqlParameter("@SecondaryLead", SLS)
            cnn.Open()
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()
            Dim g As New PopulateLeadSources
            g.PopulateSecondary(PLS)
            MarketerLeadSources.cboSLS.SelectedItem = SLS
        End Sub
    End Class
#End Region

    Private Sub cboMM_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMM.SelectedValueChanged
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

        Dim z As String = ""
        z = Me.cboMM.Text
        Select Case z
            Case Is = "<Add New>"
                Dim x As String = ""
                Dim y As String = ""
                Dim c As New InsertNewManager
                Dim cap As New Captilalize

                x = cap.CapitalizeText(InputBox$("Enter Manager's First Name", "Add New Manager"))

                If x.ToString.Length < 2 Then
                    MsgBox("You must enter a First Name!" & vbCr & "(Add New Manager Aborted)", MsgBoxStyle.Critical)
                    Me.cboMM.SelectedItem = ""
                    Exit Sub
                End If
                y = cap.CapitalizeText(InputBox$("Enter Manager's Last Name", "Add New Manager"))
                If y.ToString.Length < 2 Then
                    MsgBox("You must enter a Last Name!" & vbCr & "(Add New Manager Aborted)", MsgBoxStyle.Critical)
                    Me.cboMM.SelectedItem = ""
                    Exit Sub
                End If
                c.InsertManager(x, y)
                Dim d = New PopulateMarketingManager
                d.PopulateMarketingMAN()
                Me.cboMM.SelectedItem = x & " " & y
                Exit Select
            Case Is = "________________________________________________________"
                Me.cboMM.Text = ""
                'Me.cboMM.Enabled = False
                Exit Select
            Case Is = ""
                'Me.cboMM.Enabled = False
                Exit Select
        End Select
    End Sub

    Private Sub cboSLS_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSLS.SelectedValueChanged
        Dim x As String = ""
        x = Me.cboSLS.Text
        If x.ToString.Length < 2 Then
            Exit Sub
        End If
        Select Case x
            Case Is = "<Add New>"
                Dim y As String = Me.cboPRILS.SelectedItem
                If y.ToString.Length < 2 Then
                    Me.cboSLS.SelectedItem = ""
                    Exit Sub
                End If
                Dim z As String = InputBox$("Enter new Secondary Lead Source", "New Secondary Lead Source for " & y)
                If z.ToString.Length < 2 Then
                    Me.cboSLS.SelectedItem = ""
                    Exit Sub
                End If
                Dim b As New InsertNewSLS
                Dim c As New Captilalize
                b.InsertSLS(y, c.CapitalizeText(z))
                Exit Select
            Case Is = "________________________________________________________"
                Me.cboSLS.SelectedItem = ""
                Exit Select
            Case Is = ""
                Exit Select
        End Select
    End Sub

    
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Me.ErrorProvider1.Clear()
        Dim cnt As Integer = 0
        If Me.txtFName.Text = "" Then
            Me.ErrorProvider1.SetError(Me.txtFName, "Required Field")
            cnt += 1
        End If
        If Me.txtLname.Text = "" Then
            Me.ErrorProvider1.SetError(Me.txtLname, "Required Field")
            cnt += 1
        End If
        If Me.cboPRILS.SelectedItem = "" Then
            Me.ErrorProvider1.SetError(Me.cboPRILS, "Required Field")
            cnt += 1
        End If
        If Me.cboMM.SelectedItem = "" Then
            Me.ErrorProvider1.SetError(Me.cboMM, "Required Field")
            cnt += 1
        End If
        If cnt >= 1 Then
            Exit Sub
        End If
        If cnt <= 0 Then
            Me.ErrorProvider1.Clear()
        End If

        Select Case Frm.Name
            Case "EnterLead"
                Dim c As New ENTER_LEAD.InsertMarketer
                c.WriteMarketerToTable(Me.txtFName.Text, Me.txtLname.Text, Me.cboPRILS.Text, Me.cboSLS.Text, Me.cboMM.Text)
            Case "EditCustomerInfo"
                Dim d As New EDIT_CUSTOMER_INFORMATION
                d.WriteMarketerToTable(Me.txtFName.Text, Me.txtLname.Text, Me.cboPRILS.Text, Me.cboSLS.Text, Me.cboMM.Text)
        End Select
        Me.Close()


    End Sub

    

    Private Sub txtFName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFName.LostFocus
        Dim x = New Captilalize
        Dim y As String
        y = x.CapitalizeText(Me.txtFName.Text)
        Me.txtFName.Text = y

    End Sub

    Private Sub txtLname_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLname.LostFocus
        Dim x = New Captilalize
        Dim y As String
        y = x.CapitalizeText(Me.txtLname.Text)
        Me.txtLname.Text = y
    End Sub

End Class
