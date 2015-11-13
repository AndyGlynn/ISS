Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class SendNotesSM
    Public ID As String



    Private Sub cboautonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.SelectedValueChanged
        Dim i As String
        If Me.cboautonotes.SelectedItem = "<Add New>" Then
            'Me.cboautonotes.SelectedItem = Nothing
            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")
            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsSMNotes", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Note", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            Me.cboautonotes.Items.Clear()
            Me.cboautonotes.Items.Add("<Add New>")
            Me.cboautonotes.Items.Add("___________________________________________________")
            cnn.Open()
            r1 = cmdINS.ExecuteReader
            While r1.Read
                Me.cboautonotes.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()

            Me.cboautonotes.Text = i

        End If
        If Me.cboautonotes.Text = i Then
            Exit Sub
        End If
        Dim x
        Dim y
        y = Me.RichTextBox1.Text.Length
        x = Me.cboautonotes.Text
        If x = "" Then
            Exit Sub
        End If
        If x = "<Add New>" Then
            Exit Sub
        ElseIf x = "___________________________________________________" Then
            Exit Sub
        End If

        Select Case y
            Case Is < 1
                Me.RichTextBox1.Text = x
                Exit Select
            Case Is > 0
                Me.RichTextBox1.Text = (Me.RichTextBox1.Text & ", " & x)
        End Select
    End Sub


    Private Sub SendNotesSM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdINS As SqlCommand = New SqlCommand("dbo.GetSMNotes", cnn)

        cmdINS.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        Me.cboautonotes.Items.Clear()
        Me.cboautonotes.Items.Add("<Add New>")
        Me.cboautonotes.Items.Add("___________________________________________________")
        cnn.Open()
        r1 = cmdINS.ExecuteReader
        While r1.Read
            Me.cboautonotes.Items.Add(r1.Item(0))


        End While
        r1.Close()
        cnn.Close()

        Try
            Me.RichTextBox1.Text = STATIC_VARIABLES.SMN.Item(ID)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdINS As SqlCommand = New SqlCommand("dbo.WriteSMNotes", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        Dim param2 As SqlParameter = New SqlParameter("@Notes", Me.RichTextBox1.Text)
        cmdINS.Parameters.Add(param1)
        cmdINS.Parameters.Add(param2)
        cmdINS.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        cnn.Open()
        r1 = cmdINS.ExecuteReader
        r1.Read()
        r1.Close()
        cnn.Close()

        Try
            STATIC_VARIABLES.SMN.Add(Me.RichTextBox1.Text, ID)
        Catch ex As Exception
            STATIC_VARIABLES.SMN.Remove(ID)
            STATIC_VARIABLES.SMN.Add(Me.RichTextBox1.Text, ID)
        End Try
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub lblautonotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblautonotes.Click
        Me.cboautonotes.DroppedDown = True
    End Sub



  
End Class
