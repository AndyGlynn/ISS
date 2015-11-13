Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class exclusions
    Dim cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)

    Private Sub exclusions_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        For x As Integer = 0 To Me.Controls.Count - 1
            If Me.Controls.Item(x).Name.Contains("chk") Then
                Dim chk As CheckBox = Me.Controls.Item(x)
                chk.Checked = False
            End If
        Next
    End Sub

    Private Sub exclusions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cmd As SqlCommand = New SqlCommand("select * from exclusions", cnn)
        cmd.CommandType = CommandType.Text
        Try
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmd.ExecuteReader
            While r1.Read
                If r1.Item(0) = True Then
                    Me.chkGenerated.Checked = True
                End If
                If r1.Item(1) = True Then
                    Me.chkMarketer.Checked = True
                End If
                If r1.Item(2) = True Then
                    Me.chkPLS.Checked = True
                End If
                If r1.Item(3) = True Then
                    Me.chkSLS.Checked = True
                End If
                If r1.Item(4) = True Then
                    Me.chkMResult.Checked = True
                End If
                If r1.Item(5) = True Then
                    Me.chkPhone.Checked = True
                End If
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
            cnn.Close()
        End Try
    End Sub

  
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim cmd As SqlCommand = New SqlCommand("update exclusions set generated = '" & Me.chkGenerated.Checked.ToString & _
        "' , marketer = '" & Me.chkMarketer.Checked.ToString & "' , PLS = '" & Me.chkPLS.Checked.ToString & "' , SLS = '" & _
        Me.chkSLS.Checked.ToString & "' , LastMResult = '" & Me.chkMResult.Checked.ToString & "' , Phone = '" & _
        Me.chkPhone.Checked.ToString & "'", cnn)
        cmd.CommandType = CommandType.Text
        Try
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmd.ExecuteReader
            r1.Close()
            cnn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
            cnn.Close()
        End Try
        Me.Close()

    End Sub

    Private Sub chkPLS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPLS.CheckedChanged
        If Me.chkPLS.Checked = True Then
            Me.chkSLS.Checked = True
            Me.chkSLS.Enabled = False
        Else
            Me.chkSLS.Checked = False
            Me.chkSLS.Enabled = True
        End If
    End Sub

  
   
End Class
