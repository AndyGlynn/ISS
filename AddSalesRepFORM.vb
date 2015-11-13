Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class AddSalesRep
    Dim errorcount As Integer = 0
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public combo As ComboBox
    Private Sub AddSalesRep_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If combo.Name = Nothing Then
            sender = combo
        End If
        Me.ep.Clear()
        Me.txtEmail.Text = ""
        Me.txtFName.Text = ""
        Me.txtLname.Text = ""
        Me.chkEmail.Checked = False
        Me.chkExclude.Checked = False
        Me.cboSM.SelectedItem = Nothing
        Me.PopulateSalesManager()
    End Sub



    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Me.ep.Clear()
        Me.errorcount = 0
        If Me.txtFName.Text = "" Then
            Me.errorcount += 1
            ep.SetError(Me.txtFName, "Required Field")
        End If
        If Me.txtLname.Text = "" Then
            Me.errorcount += 1
            ep.SetError(Me.txtLname, "Required Field")
        End If
        If Me.cboSM.SelectedItem = "" Or Me.cboSM.SelectedItem = Nothing Then
            Me.errorcount += 1
            ep.SetError(Me.cboSM, "Required Field")
        End If
        If Me.chkEmail.Checked = True And Me.txtEmail.Text = "" Then
            Me.errorcount += 1
            Me.ep.SetError(Me.chkEmail, "Must Supply Valid Email Address")
        End If
        If txtEmail.Text <> "" Then
            If Me.txtEmail.Text.Contains("@") = False Or Me.txtEmail.Text.Contains(".") = False Then
                Me.errorcount += 1
                Me.ep.SetError(Me.txtEmail, "Not a Valid Email Address")
            End If
        End If
 
        If Me.errorcount >= 1 Then
            Exit Sub
        End If
        Try
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsertSalesRep", cnn)
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim param1 As SqlParameter = New SqlParameter("@Fname", Trim(CapitalizeText(Me.txtFName.Text)))
            Dim param2 As SqlParameter = New SqlParameter("@LName", Trim(CapitalizeText(Me.txtLname.Text)))
            Dim param3 As SqlParameter = New SqlParameter("@Can", Me.chkEmail.CheckState)
            Dim param4 As SqlParameter = New SqlParameter("@exclude", Me.chkExclude.CheckState)
            Dim param5 As SqlParameter = New SqlParameter("@SM", Me.cboSM.Text)
            Dim param6 As SqlParameter = New SqlParameter("@email", Me.txtEmail.Text)
            cmdINS.Parameters.Add(param1)
            cmdINS.Parameters.Add(param2)
            cmdINS.Parameters.Add(param3)
            cmdINS.Parameters.Add(param4)
            cmdINS.Parameters.Add(param5)
            cmdINS.Parameters.Add(param6)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdINS.ExecuteReader
            r1.Close()
            cnn.Close()
            combo.Items.Add(Trim(CapitalizeText(Me.txtFName.Text)) & " " & Trim(CapitalizeText(Me.txtLname.Text)))
            combo.SelectedItem = Trim(CapitalizeText(Me.txtFName.Text)) & " " & Trim(CapitalizeText(Me.txtLname.Text))
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("AddSalesRep", "", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "InsertSalesRep")
        End Try
        Me.Close()
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
            err.WriteLog("AddSalesRep", "ByVal TextToCap as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CapitalizeText")

        End Try
    End Function
    Public Sub PopulateSalesManager()
        Dim cmdGet As SqlCommand = New SqlCommand("Select * from salesmanagerpull", cnn)
        cmdGet.CommandType = CommandType.Text
        Dim r1 As SqlDataReader
        cnn.Open()
        r1 = cmdGet.ExecuteReader
        Me.cboSM.Items.Clear()
        Me.cboSM.Items.Add("<Add New>")
        Me.cboSM.Items.Add("________________________________________________________")
        Me.cboSM.Items.Add("")
        While r1.Read
            Me.cboSM.Items.Add(r1.Item(1) & " " & r1.Item(2))
        End While
        r1.Close()
        cnn.Close()

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
 Private Sub cboSM_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSM.SelectedValueChanged
        If Me.cboSM.SelectedItem = "<Add New>" Then
            Me.cboSM.SelectedItem = Nothing
            Dim first As String = InputBox("Please Enter Manager's First Name", "Add New Manager")
            If first = "" Or first.Contains(" ") Then
                Exit Sub
            End If
            Dim last As String = InputBox("Please Enter Manager's Last Name", "Add New Manager")
            If last = "" Or last.Contains(" ") Then
                Exit Sub
            End If
            Try
                cnn.Open()
                Dim cmdPop As SqlCommand = New SqlCommand("insert into SalesManagerPull (FName , LName ) values ('" & CapitalizeText(first) & "' , '" & CapitalizeText(last) & "')", cnn)
                Dim r As SqlDataReader = cmdPop.ExecuteReader
                r.Read()
                r.Close()
                cnn.Close()
                PopulateSalesManager()
                Me.cboSM.SelectedItem = CapitalizeText(first) & " " & CapitalizeText(last)
            Catch ex As Exception

            End Try
            Exit Sub
        End If
        If Me.cboSM.SelectedItem = "________________________________________________________" Then
            Me.cboSM.SelectedItem = Nothing
            Exit Sub
        End If
    End Sub
End Class
