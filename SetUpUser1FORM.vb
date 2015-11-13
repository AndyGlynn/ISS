Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class SetUpUser1
    Public EditMode As Boolean = False
    Public USR As String
    Public Id As Integer
    Dim pnlIndex As Integer = 1
    Private Sub SetUpUser1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.chklstForms.Enabled = True
        Me.txtUserName.Enabled = True
        lblUncheckAll_Click(Nothing, Nothing)
        Me.ToolTip1.RemoveAll()
        Me.txtUserName.Text = ""
        Me.txtFName.Text = ""
        Me.txtLName.Text = ""
        Me.txtPassword.Text = ""
        Me.txtConfirmPassword.Text = ""
        Me.cboManager.SelectedItem = Nothing
        Me.cboStartForm.SelectedItem = Nothing
        Me.chkCanEmail.CheckState = CheckState.Unchecked
        Me.txtEmail.Text = ""
        Me.txtEmailUser.Text = ""
        Me.txtEmailPassword.Text = ""
        Me.txtIncomingPort.Text = ""
        Me.txtIncoming.Text = ""
        Me.txtOutgoingPort.Text = ""
        Me.txtOutgoing.Text = ""
        Me.txtUserOut.Text = ""
        Me.txtPasswordOut.Text = ""
        Me.chkOutgoingPortSSL.CheckState = CheckState.Unchecked
        Me.chkOutgoingRequiresAuthen.CheckState = CheckState.Unchecked
        Me.chkSameSettings.CheckState = CheckState.Checked
        Me.chkIncomingPortSSL.CheckState = CheckState.Unchecked
        Me.chkOutgoingPortSSL.CheckState = CheckState.Unchecked
        Me.rdoPopAuthen.Checked = False
        Me.rdoSecurePassword.Checked = False
        Me.rdoClearText.Checked = False
        Me.pnlIndex = 1
        Me.lblFName.Visible = True
        Me.lblLName.Visible = True

        If EditMode = True Then
            Me.Text = "Edit User"
            Me.pnlUserName.Visible = True
            Me.pnlAccess.Visible = False
            Me.pnlEmail.Visible = False
            Me.btnBack.Visible = False
            Me.btnSave.Text = "Next"
            Me.cboManager.Items.Clear()
            Me.cboManager.Items.Add("<Add New>")
            Me.cboManager.Items.Add("____________________________________")
            Me.cboManager.Items.Add("")
            Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("select ManFName + ' ' + ManLName  from ManagerMaster order by Department asc, ManLName asc", cnx)
            cnx.Open()
            Dim r As SqlDataReader
            r = cmdGet.ExecuteReader
            While r.Read
                Me.cboManager.Items.Add(r.Item(0))
            End While
            r.Close()
            cnx.Close()
            Dim cnx1 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet1 As SqlCommand = New SqlCommand("select * from UserPermissionTable where username = '" & USR & "'", cnx1)
            cnx1.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet1.ExecuteReader
            While r1.Read
                Id = r1.Item(0)
                Me.txtUserName.Text = r1.Item(1)
                txtFName_GotFocus(Nothing, Nothing)
                Me.txtFName.Text = r1.Item(2)
                txtFName_LostFocus(Nothing, Nothing)
                txtLName_GotFocus(Nothing, Nothing)
                Me.txtLName.Text = r1.Item(3)
                txtLName_LostFocus(Nothing, Nothing)
                Me.txtPassword.Text = r1.Item(4)
                Me.txtConfirmPassword.Text = r1.Item(4)
                If r1.Item(5) = True Then
                    Me.chklstForms.SetItemChecked(0, True)
                End If
                If r1.Item(6) = True Then
                    Me.chklstForms.SetItemChecked(1, True)
                End If
                If r1.Item(7) = True Then
                    Me.chklstForms.SetItemChecked(2, True)
                End If
                If r1.Item(8) = True Then
                    Me.chklstForms.SetItemChecked(3, True)
                End If
                If r1.Item(9) = True Then
                    Me.chklstForms.SetItemChecked(4, True)
                End If
                If r1.Item(10) = True Then
                    Me.chklstForms.SetItemChecked(5, True)
                End If
                If r1.Item(11) = True Then
                    Me.chklstForms.SetItemChecked(6, True)
                End If
                If r1.Item(12) = True Then
                    Me.chklstForms.SetItemChecked(7, True)
                End If
                If r1.Item(13) = True Then
                    Me.chklstForms.SetItemChecked(8, True)
                End If
                If r1.Item(14) = True Then
                    Me.chklstForms.SetItemChecked(9, True)
                End If
                If r1.Item(19) <> "" Then
                    Me.cboManager.Text = r1.Item(19) & " " & r1.Item(20)
                End If
                If USR = "Admin" Then
                    lblCheckAll_Click(Nothing, Nothing)
                    Me.txtUserName.Enabled = False
                    Me.ToolTip1.SetToolTip(Me.txtUserName, "Cannot Edit Admin User Name")
                    Me.chklstForms.Enabled = False
                End If
                chklstForms_LostFocus(Nothing, Nothing)
                cboStartForm_GotFocus(Nothing, Nothing)
                Me.cboStartForm.Text = r1.Item(15)
                If r1.Item(39) = True Then
                    Me.chkCanEmail.CheckState = CheckState.Checked
                Else
                    Exit While
                End If
                Me.txtEmail.Text = r1.Item(24)
                Me.txtIncoming.Text = r1.Item(27)
                Me.txtOutgoing.Text = r1.Item(28)
                Me.txtIncomingPort.Text = r1.Item(29).ToString
                Me.txtOutgoingPort.Text = r1.Item(30).ToString
                Me.txtEmailUser.Text = r1.Item(25)
                Me.txtEmailPassword.Text = r1.Item(26)
                If r1.Item(31) = True Then
                    Me.chkIncomingPortSSL.CheckState = CheckState.Checked
                End If
                If r1.Item(32) = True Then
                    Me.chkOutgoingPortSSL.CheckState = CheckState.Checked
                End If
                Select Case r1.Item(33)
                    Case Is = "Clear"
                        Me.rdoClearText.Checked = True
                    Case Is = "Secure"
                        Me.rdoSecurePassword.Checked = True
                    Case Is = "APOP"
                        Me.rdoPopAuthen.Checked = True
                End Select
                If r1.Item(34) = True Then
                    Me.chkOutgoingRequiresAuthen.CheckState = CheckState.Checked
                End If
                If r1.Item(35) = True Then
                    Me.chkSameSettings.CheckState = CheckState.Checked
                Else
                    Me.chkSameSettings.CheckState = CheckState.Checked
                    Me.txtUserOut.Text = r1.Item(36)
                    Me.txtPasswordOut.Text = r1.Item(37)
                    If r1.Item(38) = True Then
                        Me.chkOutgoingSSL.CheckState = CheckState.Checked
                    End If
                End If
            End While
            r1.Close()
            cnx1.Close()
     
        Else
            Me.Text = "Add New User"
            Me.pnlUserName.Visible = True
            Me.pnlAccess.Visible = False
            Me.pnlEmail.Visible = False
            Me.btnBack.Visible = False
            Me.btnSave.Text = "Next"
            Me.cboManager.Items.Clear()
            Me.cboManager.Items.Add("<Add New>")
            Me.cboManager.Items.Add("____________________________________")
            Me.cboManager.Items.Add("")
            Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("select ManFName + ' ' + ManLName  from ManagerMaster order by Department asc, ManLName asc", cnx)
            cnx.Open()
            Dim r As SqlDataReader
            r = cmdGet.ExecuteReader
            While r.Read
                Me.cboManager.Items.Add(r.Item(0))
            End While
            r.Close()
            cnx.Close()
        End If


    End Sub

    Private Sub lblCheckAll_Click(sender As Object, e As EventArgs) Handles lblCheckAll.Click
        For x As Integer = 0 To Me.chklstForms.Items.Count - 1
            Me.chklstForms.SetItemChecked(x, True)
        Next
        chklstForms_LostFocus(Nothing, Nothing)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Me.TabPage1.ImageKey = Nothing
        Me.TabPage1.ToolTipText = ""
        Me.TabPage2.ImageKey = Nothing
        Me.TabPage2.ToolTipText = ""
        Me.TabPage3.ImageKey = Nothing
        Me.TabPage3.ToolTipText = ""
        Dim errorcnt As Integer = 0
        Dim tab1ec As Integer = 0
        Dim tab2ec As Integer = 0
        Dim tab3ec As Integer = 0
        Dim Doit As Boolean = False
        Me.ErrorProvider1.Clear()

        Select Case pnlIndex
            Case 1
                If Me.txtUserName.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.txtUserName, "Required Field")
                    errorcnt += 1
                End If
                If Me.txtFName.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.txtFName, "Required Field")
                    errorcnt += 1
                End If
                If Me.txtLName.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.txtLName, "Required Field")
                    errorcnt += 1
                End If
                If Me.txtPassword.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.txtPassword, "Required Field")
                    errorcnt += 1
                End If
                If Me.txtConfirmPassword.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.txtConfirmPassword, "Required Field")
                    errorcnt += 1
                End If
                If Me.txtPassword.Text <> "" And Me.txtConfirmPassword.Text <> "" Then
                    If Me.txtPassword.Text <> Me.txtConfirmPassword.Text Then
                        Me.ErrorProvider1.SetError(Me.txtConfirmPassword, "Does not match Password")
                        errorcnt += 1
                    End If
                End If
            Case 2
                If Me.chklstForms.CheckedItems.Count <= 0 Then
                    Me.ErrorProvider1.SetError(chklstForms, "You must select at least 1 form")
                    errorcnt += 1
                End If
                If Me.cboStartForm.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.cboStartForm, "Required Field")
                    errorcnt += 1
                End If
            Case 3
                If Me.chkCanEmail.CheckState = CheckState.Checked Then
                    If Me.txtIncoming.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtIncoming, "Required Field")
                        errorcnt += 1
                        tab3ec += 1
                    End If
                    If Me.txtOutgoing.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtOutgoing, "Required Field")
                        errorcnt += 1
                        tab3ec += 1
                    End If
                    If Me.txtIncomingPort.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtIncomingPort, "Required Field")
                        errorcnt += 1
                        tab3ec += 1
                    End If
                    If Me.txtOutgoingPort.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtOutgoingPort, "Required Field")
                        errorcnt += 1
                        tab3ec += 1
                    End If
                    If Me.txtEmail.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtEmail, "Required Field")
                        errorcnt += 1
                        tab3ec += 1
                    End If
                    If Me.txtEmailUser.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtEmailUser, "Required Field")
                        errorcnt += 1
                        tab1ec += 1
                    End If
                    If Me.txtEmail.Text.Contains(".") = False Or Me.txtEmail.Text.Contains("@") = False Then
                        Me.ErrorProvider1.SetError(Me.txtEmailUser, "You must enter a valid email address")
                        errorcnt += 1
                        tab1ec += 1
                    End If
                    If Me.txtEmailPassword.Text = "" Then
                        Me.ErrorProvider1.SetError(Me.txtEmailPassword, "Required Field")
                        errorcnt += 1
                        tab1ec += 1
                    End If
                    If rdoClearText.Checked = False And rdoSecurePassword.Checked = False And rdoPopAuthen.Checked = False Then
                        Me.ErrorProvider1.SetError(Me.rdoSecurePassword, "You Must Select an authentication method")
                        errorcnt += 1
                        tab1ec += 1
                    End If
                    If Me.chkOutgoingRequiresAuthen.CheckState = CheckState.Checked Then
                        If Me.chkSameSettings.CheckState = CheckState.Unchecked Then
                            If Me.txtUserOut.Text = "" Then
                                Me.ErrorProvider1.SetError(Me.txtUserOut, "Required Field")
                                errorcnt += 1
                                tab2ec += 1
                            End If
                            If Me.txtPasswordOut.Text = "" Then
                                Me.ErrorProvider1.SetError(Me.txtPasswordOut, "Required Field")
                                errorcnt += 1
                                tab2ec += 1
                            End If
                        End If
                    End If
                End If
        End Select
        If tab1ec >= 1 Then
            Me.TabPage1.ImageIndex = 0
            Me.TabPage1.ToolTipText = "Required Fields on this Tab"
        End If
        If tab2ec >= 1 Then
            Me.TabPage2.ImageIndex = 0
            Me.TabPage2.ToolTipText = "Required Fields on this Tab"
        End If
        If tab3ec >= 1 Then
            Me.TabPage3.ImageIndex = 0
            Me.TabPage3.ToolTipText = "Required Fields on this Tab"
        End If


        If errorcnt >= 1 Then
            Exit Sub
        End If

        pnlIndex += 1

        If pnlIndex > 1 Then
            btnBack.Visible = True
        Else
            btnBack.Visible = False
        End If
        If (pnlIndex = 3 And chkCanEmail.CheckState = CheckState.Checked) Or (pnlIndex = 2 And chkCanEmail.CheckState = CheckState.Unchecked) Then
            If EditMode = True Then
                btnSave.Text = "Update"
            Else
                btnSave.Text = "Save"
            End If
        Else
            Me.btnSave.Text = "Next"
        End If
        If chkCanEmail.CheckState = CheckState.Unchecked And pnlIndex = 3 Then
            pnlIndex = 2
            Doit = True
        End If
        If pnlIndex > 3 Then
            pnlIndex = 3
            Doit = True
        End If
        pnlControl()
        If Doit = True Then



            Dim CC As Boolean
            Dim WC As Boolean
            Dim PC As Boolean
            Dim RC As Boolean
            Dim CF As Boolean
            Dim SM As Boolean
            Dim MM As Boolean
            Dim F As Boolean
            Dim I As Boolean
            Dim A As Boolean
            Dim Manager As String = ""
            Dim First As String
            Dim Last As String
            Dim s
            Dim IPort As Integer
            Dim OPort As Integer
            Dim LogonUsing As String = ""
            Try
                IPort = CType(Me.txtIncomingPort.Text, Integer)
            Catch ex As Exception
                IPort = 0
            End Try
            Try
                OPort = CType(Me.txtOutgoingPort.Text, Integer)
            Catch ex As Exception
                OPort = 0
            End Try
            If rdoClearText.Checked = True Then
                LogonUsing = "Clear"
            End If
            If rdoSecurePassword.Checked = True Then
                LogonUsing = "Secure"
            End If
            If rdoPopAuthen.Checked = True Then
                LogonUsing = "APOP"
            End If
       
            Try
                Manager = Me.cboManager.Text
            Catch ex As Exception
                Manager = ""
            End Try
            If Manager <> "" Then
                s = Split(Manager, " ")
                First = s(0)
                Last = s(1)
            Else
                First = ""
                Last = ""
            End If

            For x As Integer = 0 To Me.chklstForms.Items.Count - 1
                Select Case x
                    Case Is = 0
                        CC = Me.chklstForms.GetItemChecked(x)
                    Case Is = 1
                        WC = Me.chklstForms.GetItemChecked(x)
                    Case Is = 2
                        PC = Me.chklstForms.GetItemChecked(x)
                    Case Is = 3
                        RC = Me.chklstForms.GetItemChecked(x)
                    Case Is = 4
                        CF = Me.chklstForms.GetItemChecked(x)
                    Case Is = 5
                        SM = Me.chklstForms.GetItemChecked(x)
                    Case Is = 6
                        MM = Me.chklstForms.GetItemChecked(x)
                    Case Is = 7
                        F = Me.chklstForms.GetItemChecked(x)
                    Case Is = 8
                        I = Me.chklstForms.GetItemChecked(x)
                    Case Is = 9
                        A = Me.chklstForms.GetItemChecked(x)
                End Select
            Next
            Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertUpdateUser", cnn)
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@Editmode", EditMode)
            Dim param2 As SqlParameter = New SqlParameter("@Id", Id)
            Dim param3 As SqlParameter = New SqlParameter("@UserName", txtUserName.Text)
            Dim param4 As SqlParameter = New SqlParameter("@UserFirstName", Trim(Me.txtFName.Text))
            Dim param5 As SqlParameter = New SqlParameter("@UserLastName", Trim(Me.txtLName.Text))
            Dim param6 As SqlParameter = New SqlParameter("@UserPWD", Trim(Me.txtPassword.Text))
            Dim param7 As SqlParameter = New SqlParameter("@Coldcall", CC)
            Dim param69 As SqlParameter = New SqlParameter("@WarmCall", WC)
            Dim param8 As SqlParameter = New SqlParameter("@PreviousCust", PC)
            Dim param9 As SqlParameter = New SqlParameter("@Recovery", RC)
            Dim param10 As SqlParameter = New SqlParameter("@Confirmer", CF)
            Dim param11 As SqlParameter = New SqlParameter("@SalesManager", SM)
            Dim param12 As SqlParameter = New SqlParameter("@MarketingManager", MM)
            Dim param13 As SqlParameter = New SqlParameter("@Finance", F)
            Dim param14 As SqlParameter = New SqlParameter("@Install", I)
            Dim param15 As SqlParameter = New SqlParameter("@Administration", A)
            Dim param16 As SqlParameter = New SqlParameter("@StartupForm", Me.cboStartForm.Text)
            Dim param17 As SqlParameter = New SqlParameter("@ManagerFirstName", Trim(First))
            Dim param18 As SqlParameter = New SqlParameter("@ManagerLastName", Trim(Last))
            Dim param19 As SqlParameter = New SqlParameter("@Email", Trim(Me.txtEmail.Text))
            Dim param20 As SqlParameter = New SqlParameter("@EmailLogin", Me.txtEmailUser.Text)
            Dim param21 As SqlParameter = New SqlParameter("@EmailPassword", Me.txtEmailPassword.Text)
            Dim param22 As SqlParameter = New SqlParameter("@Incoming", Me.txtIncoming.Text)
            Dim param23 As SqlParameter = New SqlParameter("@Outgoing", Me.txtOutgoing.Text)
            Dim param24 As SqlParameter = New SqlParameter("@IPort", IPort)
            Dim param25 As SqlParameter = New SqlParameter("@OPort", OPort)
            Dim param26 As SqlParameter = New SqlParameter("@LogonUsing", LogonUsing)
            Dim param27 As SqlParameter = New SqlParameter("@IPortSSL", Me.chkIncomingPortSSL.Checked)
            Dim param28 As SqlParameter = New SqlParameter("@OPortSSL", Me.chkOutgoingSSL.Checked)
            Dim param29 As SqlParameter = New SqlParameter("@OutgoAuthen", Me.chkOutgoingRequiresAuthen.Checked)
            Dim param30 As SqlParameter = New SqlParameter("@SameSettings", Me.chkSameSettings.Checked)
            Dim param31 As SqlParameter = New SqlParameter("@OutgoUSR", Me.txtUserOut.Text)
            Dim param32 As SqlParameter = New SqlParameter("@OutgoPW", Me.txtPasswordOut.Text)
            Dim param33 As SqlParameter = New SqlParameter("@OutgoSSL", Me.chkOutgoingSSL.Checked)
            Dim param34 As SqlParameter = New SqlParameter("@CanEmail", Me.chkCanEmail.Checked)
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
            cmdINS.Parameters.Add(param17)
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
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()
            If EditMode = False Then
                SetUpUser.lstUsers.Items.Clear()
                Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                Dim cmdGet As SqlCommand = New SqlCommand("Select  Distinct(username) from UserPermissionTable order by Username asc", cnx)
                cnx.Open()
                Dim r As SqlDataReader
                r = cmdGet.ExecuteReader
                While r.Read
                    SetUpUser.lstUsers.Items.Add(r.Item(0))
                End While
                r.Close()
                cnx.Close()
            End If
            Me.Close()
        End If
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        pnlIndex -= 1
        If pnlIndex = 1 Then
            btnBack.Visible = False
        End If
        If (pnlIndex < 3 And chkCanEmail.CheckState = CheckState.Checked) Or (pnlIndex < 2 And chkCanEmail.CheckState = CheckState.Unchecked) Then
            btnSave.Text = "Next"
        End If
        If pnlIndex < 1 Then
            pnlIndex = 1
        End If
        pnlControl()
    End Sub
    Private Sub pnlControl()
        Select Case pnlIndex
            Case 1
                pnlUserName.Visible = True
                pnlAccess.Visible = False
                pnlEmail.Visible = False
            Case 2
                pnlUserName.Visible = False
                pnlAccess.Visible = True
                pnlEmail.Visible = False
            Case 3
                pnlUserName.Visible = False
                pnlAccess.Visible = False
                pnlEmail.Visible = True
        End Select


    End Sub

    Private Sub chkCanEmail_CheckStateChanged(sender As Object, e As EventArgs) Handles chkCanEmail.CheckStateChanged
        If chkCanEmail.CheckState = CheckState.Checked Then
            If Me.txtIncoming.Text = "" Then
                Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
                Dim cmdGet As SqlCommand = New SqlCommand("select * from CompanyInfo", cnx)
                cnx.Open()
                Dim r As SqlDataReader
                r = cmdGet.ExecuteReader
                While r.Read
                    If r.Item(27) = False Then
                        Exit While
                    End If
                    Me.txtIncoming.Text = r.Item(15)
                    Me.txtOutgoing.Text = r.Item(16)
                    Me.txtIncomingPort.Text = CType(r.Item(17), String)
                    Me.txtOutgoingPort.Text = CType(r.Item(18), String)
                    Me.chkIncomingPortSSL.Checked = r.Item(19)
                    Me.chkOutgoingPortSSL.Checked = r.Item(20)
                    If r.Item(21) = "Clear" Then
                        Me.rdoClearText.Checked = True
                    ElseIf r.Item(21) = "Secure" Then
                        Me.rdoSecurePassword.Checked = True
                    ElseIf r.Item(21) = "APOP" Then
                        Me.rdoPopAuthen.Checked = True
                    End If
                    Me.chkOutgoingRequiresAuthen.Checked = r.Item(22)
                    If r.Item(22) = True Then
                        Me.chkSameSettings.Checked = r.Item(23)
                    End If
                    If r.Item(23) = False Then
                        Me.chkOutgoingSSL.Checked = r.Item(26)
                    End If
                End While
                r.Close()
                cnx.Close()
            End If
            btnSave.Text = "Next"
        Else
            Me.txtIncoming.Text = ""
            Me.txtOutgoing.Text = ""
            Me.txtEmail.Text = ""
            Me.txtEmailUser.Text = ""
            Me.txtEmailPassword.Text = ""
            Me.txtIncomingPort.Text = ""
            Me.txtOutgoingPort.Text = ""
            Me.chkIncomingPortSSL.CheckState = CheckState.Unchecked
            Me.chkOutgoingPortSSL.CheckState = CheckState.Unchecked
            Me.rdoSecurePassword.Checked = False
            Me.rdoClearText.Checked = False
            Me.rdoPopAuthen.Checked = False
            Me.chkOutgoingRequiresAuthen.CheckState = CheckState.Unchecked
            Me.chkSameSettings.CheckState = CheckState.Checked
            Me.txtUserOut.Text = ""
            Me.txtPasswordOut.Text = ""
            Me.chkOutgoingSSL.CheckState = CheckState.Unchecked
            If EditMode = True Then
                btnSave.Text = "Update"
            Else
                btnSave.Text = "Save"
            End If
        End If
    End Sub



    Private Sub chkSameSettings_CheckStateChanged(sender As Object, e As EventArgs) Handles chkSameSettings.CheckStateChanged
        If Me.chkSameSettings.CheckState = CheckState.Checked Then
            Me.txtPasswordOut.Text = ""
            Me.txtUserOut.Text = ""
            Me.txtPasswordOut.Enabled = False
            Me.txtUserOut.Enabled = False
            Me.chkOutgoingSSL.Enabled = False
        Else
            Me.txtPasswordOut.Enabled = True
            Me.txtUserOut.Enabled = True
            Me.chkOutgoingSSL.Enabled = True
        End If
    End Sub

    Private Sub lblUncheckAll_Click(sender As Object, e As EventArgs) Handles lblUncheckAll.Click
        For x As Integer = 0 To Me.chklstForms.Items.Count - 1
            Me.chklstForms.SetItemChecked(x, False)
        Next
        chklstForms_LostFocus(Nothing, Nothing)
    End Sub

    Private Sub cboStartForm_GotFocus(sender As Object, e As EventArgs) Handles cboStartForm.GotFocus
        Dim frm As String = Me.cboStartForm.Text
        Me.cboStartForm.Items.Clear()
        For x As Integer = 0 To Me.chklstForms.CheckedItems.Count - 1
            Me.cboStartForm.Items.Add(Me.chklstForms.CheckedItems(x))
        Next
        Me.cboStartForm.Text = frm
    End Sub

    Private Sub lblFName_Click(sender As Object, e As EventArgs) Handles lblFName.Click
        Me.txtFName.Focus()
    End Sub

    Private Sub lblFName_GotFocus(sender As Object, e As EventArgs) Handles lblFName.GotFocus
        Me.txtFName.Focus()

    End Sub

    Private Sub lblLName_Click(sender As Object, e As EventArgs) Handles lblLName.Click
        Me.txtLName.Focus()
    End Sub

    Private Sub lblLName_GotFocus(sender As Object, e As EventArgs) Handles lblLName.GotFocus
        Me.txtLName.Focus()

    End Sub

    Private Sub txtFName_GotFocus(sender As Object, e As EventArgs) Handles txtFName.GotFocus
        Me.lblFName.Visible = False
    End Sub

    Private Sub txtLName_GotFocus(sender As Object, e As EventArgs) Handles txtLName.GotFocus
        Me.lblLName.Visible = False
    End Sub

    Private Sub txtLName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLName.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not Char.IsLetter(e.KeyChar)
        End If
    End Sub

    Private Sub txtLName_LostFocus(sender As Object, e As EventArgs) Handles txtLName.LostFocus
        If Me.txtLName.Text = "" Then
            Me.lblLName.Visible = True
        Else
            Me.txtLName.Text = Me.CapitalizeText(Me.txtLName.Text)
        End If
    End Sub

    Private Sub txtFName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFName.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not Char.IsLetter(e.KeyChar)
        End If
    End Sub

    Private Sub txtFName_LostFocus(sender As Object, e As EventArgs) Handles txtFName.LostFocus
        If Me.txtFName.Text = "" Then
            Me.lblFName.Visible = True
        Else
            Me.txtFName.Text = Me.CapitalizeText(Me.txtFName.Text)
        End If
    End Sub

    Private Sub chklstForms_LostFocus(sender As Object, e As EventArgs) Handles chklstForms.LostFocus
        Dim frm As String = Me.cboStartForm.Text
        Me.cboStartForm.Items.Clear()
        For x As Integer = 0 To Me.chklstForms.CheckedItems.Count - 1
            Me.cboStartForm.Items.Add(Me.chklstForms.CheckedItems(x))
        Next
        Me.cboStartForm.Text = frm
    End Sub

    Private Sub rdoClearText_CheckedChanged(sender As Object, e As EventArgs) Handles rdoClearText.CheckedChanged
        If Me.rdoClearText.Checked = True Then
            Me.rdoPopAuthen.Checked = False
            Me.rdoSecurePassword.Checked = False
        End If
    End Sub

    Private Sub rdoSecurePassword_CheckedChanged(sender As Object, e As EventArgs) Handles rdoSecurePassword.CheckedChanged
        If Me.rdoSecurePassword.Checked = True Then
            Me.rdoClearText.Checked = False
            Me.rdoPopAuthen.Checked = False
        End If
    End Sub

    Private Sub rdoPopAuthen_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPopAuthen.CheckedChanged
        If Me.rdoPopAuthen.Checked = True Then
            Me.rdoSecurePassword.Checked = False
            Me.rdoClearText.Checked = False
        End If
    End Sub



    Private Sub txtIncomingPort_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtIncomingPort.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not Char.IsDigit(e.KeyChar)
        End If
    End Sub


    Private Sub txtOutgoingPort_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOutgoingPort.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not Char.IsDigit(e.KeyChar)
        End If
    End Sub

    Private Sub txtUserName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUserName.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            If Char.IsLetterOrDigit(e.KeyChar) = False Then
                Me.ToolTip1.Show("This field only accepts letters and numbers", Me.txtUserName, 178, -40, 4000)
            End If
            e.Handled = Not Char.IsLetterOrDigit(e.KeyChar)
        End If
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If e.KeyChar <> ControlChars.Back Then
            If Char.IsLetterOrDigit(e.KeyChar) = False Then
                Me.ToolTip1.Show("This field only accepts letters and numbers", Me.txtPassword, 178, -40, 4000)
            End If
            e.Handled = Not Char.IsLetterOrDigit(e.KeyChar)
        End If
    End Sub

    Private Sub txtConfirmPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtConfirmPassword.KeyPress
        ' Me.ToolTip1.Hide(Me.txtConfirmPassword)
        If e.KeyChar <> ControlChars.Back Then
            If Char.IsLetterOrDigit(e.KeyChar) = False Then
                Me.ToolTip1.Show("This field only accepts letters and numbers", Me.txtConfirmPassword, 178, -40, 4000)
            End If
            e.Handled = Not Char.IsLetterOrDigit(e.KeyChar)
        End If
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
            err.WriteLog("ENTER_LEAD.Capitalize", "ByVal TextToCap as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CapitalizeText")

        End Try

    End Function

    Private Sub cboManager_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboManager.SelectedIndexChanged
        If Me.cboManager.SelectedItem <> Nothing Then
            If Me.cboManager.SelectedItem.ToString = "<Add New>" Or Me.cboManager.SelectedItem.ToString = "____________________________________" Then
                If Me.cboManager.SelectedItem.ToString = "<Add New>" Then
                    Me.cboManager.SelectedItem = Nothing
                    AddManager.ShowDialog()
                Else
                    Me.cboManager.SelectedItem = Nothing
                End If
            End If
        End If
    End Sub

    Private Sub txtUserName_LostFocus(sender As Object, e As EventArgs) Handles txtUserName.LostFocus
        If EditMode = False Then
            Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("Select  Count(id) from UserPermissionTable where username = '" & Me.txtUserName.Text & "'", cnx)
            cnx.Open()
            Dim r As SqlDataReader
            r = cmdGet.ExecuteReader
            r.Read()
            If r.Item(0) >= 1 Then
                MsgBox("Login already taken!", MsgBoxStyle.Exclamation, "Duplicate Login not allowed")
                Me.txtUserName.Text = ""
                Me.txtUserName.Focus()
            End If
            r.Close()
            cnx.Close()
        End If

    End Sub
End Class
