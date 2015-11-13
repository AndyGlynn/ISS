Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class SetUpUser

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAddUser.Click
        SetUpUser1.EditMode = False
        SetUpUser1.ShowDialog()
    End Sub

    Private Sub SetUpUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.lstUsers.Items.Clear()
        Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("Select  Distinct(username) from UserPermissionTable order by Username asc", cnx)
        cnx.Open()
        Dim r As SqlDataReader
        r = cmdGet.ExecuteReader
        While r.Read
            Me.lstUsers.Items.Add(r.Item(0))
        End While
        r.Close()
        cnx.Close()
    End Sub

    Private Sub btnEditUser_Click(sender As Object, e As EventArgs) Handles btnEditUser.Click
        If Me.lstUsers.SelectedItems.Count = 0 Then
            MsgBox("You must select a User to edit!", MsgBoxStyle.Exclamation, "No User Selected")
            Exit Sub
        End If
        SetUpUser1.EditMode = True
        SetUpUser1.USR = Me.lstUsers.SelectedItems(0)
        SetUpUser1.ShowDialog()
    End Sub

    Private Sub btnDeleteUser_Click(sender As Object, e As EventArgs) Handles btnDeleteUser.Click
        If Me.lstUsers.SelectedItems.Count = 0 Then
            MsgBox("You must select a User to delete!", MsgBoxStyle.Exclamation, "No User Selected")
            Exit Sub
        End If
        If Me.lstUsers.SelectedItems(0) = "Admin" Then
            MsgBox("The Admin account cannot be deleted!", MsgBoxStyle.Exclamation, "Cannot Delete")
            Exit Sub
        End If
        Dim x = MsgBox("Are you sure you want to delete this user?" & vbCr & "(This action cannot be undone)", MsgBoxStyle.YesNo, "Confirm Delete")
        If x = vbYes Then
            Dim cnx As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("Delete UserPermissionTable Where UserName = '" & Me.lstUsers.SelectedItems(0) & "'", cnx)
            cnx.Open()
            Dim r As SqlDataReader
            r = cmdGet.ExecuteReader
            While r.Read
                Me.lstUsers.Items.Add(r.Item(0))
            End While
            r.Close()
            cnx.Close()
            Me.lstUsers.Items.Clear()
            Dim cnx1 As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet1 As SqlCommand = New SqlCommand("Select  Distinct(username) from UserPermissionTable order by Username asc", cnx1)
            cnx1.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGet1.ExecuteReader
            While r1.Read
                Me.lstUsers.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnx1.Close()
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
