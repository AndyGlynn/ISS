Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class LogPhoneCall
    Public Contact1 As String
    Public Contact2 As String
    Public ID As String
    Public frm As Form

    Private Sub lblautonotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblautonotes.Click
        Me.cboautonote.DroppedDown = True
    End Sub




    Private Sub cboautonote_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonote.SelectedValueChanged
        Dim i As String
        If Me.cboautonote.SelectedItem = "<Add New>" Then
            Me.cboautonote.SelectedItem = Nothing

            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsPCAutoNotes", cnn)
            Dim cmdget As SqlCommand = New SqlCommand("dbo.GetPCAutoNotes", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Note", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure
            cmdget.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            Dim r2 As SqlDataReader
            Me.cboautonote.Items.Clear()
            Me.cboautonote.Items.Add("<Add New>")
            Me.cboautonote.Items.Add("___________________________________________________")
            cnn.Open()
            r2 = cmdINS.ExecuteReader
            r2.Close()
            r1 = cmdget.ExecuteReader
            While r1.Read
                Me.cboautonote.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()




            Me.cboautonote.Text = i

        End If
        If Me.cboautonote.Text = i Then
            Exit Sub
        End If
        Dim x
        Dim y
        y = Me.txtNotes.Text.Length
        x = Me.cboautonote.Text
        If x = "" Then
            Exit Sub
        End If
        If x = "<Add New>" Then
            Me.lblautonotes.ForeColor = Color.Gray
            Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblautonotes.Text = "Select Auto Note Here"
            'Me.cboautonotes.SelectedItem = Nothing

            Exit Sub
        ElseIf x = "___________________________________________________" Then
            Me.lblautonotes.ForeColor = Color.Gray
            Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblautonotes.Text = "Select Auto Note Here"
            'Me.cboautonotes.SelectedItem = Nothing

            Exit Sub
        End If

        Select Case y
            Case Is < 1
                Me.txtNotes.Text = x
                Me.lblautonotes.Text = Me.cboautonote.Text
                Me.lblautonotes.ForeColor = Color.Black
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Exit Select
            Case Is > 0
                Me.txtNotes.Text = (Me.txtNotes.Text & ", " & x)
                Me.lblautonotes.Text = Me.cboautonote.Text
                Me.lblautonotes.ForeColor = Color.Black
                Me.lblautonotes.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        End Select

    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        Me.ErrorProvider1.Clear()
        If Me.cboInboundOutbound.Text = "" And Me.cboSpokeWith.Text = "" And Me.txtNotes.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboInboundOutbound, "This field is required")
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "This field is required")
            Me.ErrorProvider1.SetError(Me.txtNotes, "This field is required")
            Exit Sub
        ElseIf Me.cboInboundOutbound.Text <> "" And Me.cboSpokeWith.Text = "" And Me.txtNotes.Text = "" Then
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "This field is required")
            Me.ErrorProvider1.SetError(Me.txtNotes, "This field is required")
            Exit Sub
        ElseIf Me.cboInboundOutbound.Text <> "" And Me.cboSpokeWith.Text <> "" And Me.txtNotes.Text = "" Then
            Me.ErrorProvider1.SetError(Me.txtNotes, "This field is required")
            Exit Sub
        ElseIf Me.cboInboundOutbound.Text <> "" And Me.cboSpokeWith.Text = "" And Me.txtNotes.Text <> "" Then
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "This field is required")
            Exit Sub
        ElseIf Me.cboInboundOutbound.Text = "" And Me.cboSpokeWith.Text <> "" And Me.txtNotes.Text <> "" Then
            Me.ErrorProvider1.SetError(Me.cboInboundOutbound, "This field is required")
            Exit Sub
        ElseIf Me.cboInboundOutbound.Text = "" And Me.cboSpokeWith.Text = "" And Me.txtNotes.Text <> "" Then
            Me.ErrorProvider1.SetError(Me.cboInboundOutbound, "This field is required")
            Me.ErrorProvider1.SetError(Me.cboSpokeWith, "This field is required")
            Exit Sub
        End If

        Dim Description As String
        If Me.cboDepartment.Text = "" Then
            If Me.cboInboundOutbound.Text = "This Customer Called In" Then
                Description = Me.cboSpokeWith.Text & " called in at " & DateTime.Now.ToShortTimeString & " and spoke with " & STATIC_VARIABLES.CurrentUser & ". (see notes for details about this conversation)"
            ElseIf Me.cboInboundOutbound.Text = "I Called This Customer" Then
                Description = STATIC_VARIABLES.CurrentUser & " called " & Me.cboSpokeWith.Text & " at " & DateTime.Now.ToShortTimeString & ". (see notes for details about this conversation)"
            ElseIf Me.cboInboundOutbound.Text = "This Customer Called In & Left a Message" Then
                Description = Me.cboSpokeWith.Text & " called in & left a message. Message logged by " & STATIC_VARIABLES.CurrentUser & ". (see notes for details about this message)"
            ElseIf Me.cboInboundOutbound.Text = "I Called This Customer & Left a Message" Then
                Description = STATIC_VARIABLES.CurrentUser & " called at " & DateTime.Now.ToShortTimeString & " & left a message. (see notes for details about this message)"

            End If
        ElseIf Me.cboDepartment.Text <> "" Then
            If Me.cboInboundOutbound.Text = "This Customer Called In" Then
                Description = Me.cboSpokeWith.Text & " called in at " & DateTime.Now.ToShortTimeString & " and spoke with " & STATIC_VARIABLES.CurrentUser & ", concerning the " & Me.cboDepartment.Text & " Department. (see notes for details about this conversation)"
            ElseIf Me.cboInboundOutbound.Text = "I Called This Customer" Then
                Description = STATIC_VARIABLES.CurrentUser & " called " & Me.cboSpokeWith.Text & " at " & DateTime.Now.ToShortTimeString & ", concerning the " & Me.cboDepartment.Text & " Department. (see notes for details about this conversation)"
            ElseIf Me.cboInboundOutbound.Text = "I Called This Customer & Left a Message" Then
                Description = STATIC_VARIABLES.CurrentUser & " called at " & DateTime.Now.ToShortTimeString & " & left a message , concerning the " & Me.cboDepartment.Text & " Department. (see notes for details about this message)"
            ElseIf Me.cboInboundOutbound.Text = "This Customer Called In & Left a Message" Then
                Description = Me.cboSpokeWith.Text & " called in & left a message, concerning the " & Me.cboDepartment.Text & " Department. Message logged by " & STATIC_VARIABLES.CurrentUser & ". (see notes for details about this message)"
            End If
        End If


        ''''''Add Left Message Strings '''DONE



        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdins As SqlCommand = New SqlCommand("dbo.LogPhoneCall", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Description", Description)
        Dim param2 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim Param3 As SqlParameter = New SqlParameter("@ID", ID)
        Dim Param4 As SqlParameter = New SqlParameter("@Notes", Me.txtNotes.Text)
        Dim Param5 As SqlParameter = New SqlParameter("@Dept", Me.cboDepartment.Text)

        cmdins.CommandType = CommandType.StoredProcedure
        cmdins.Parameters.Add(param1)
        cmdins.Parameters.Add(param2)
        cmdins.Parameters.Add(Param3)
        cmdins.Parameters.Add(Param4)
        cmdins.Parameters.Add(Param5)

        cnn.open()
        Dim r1 = cmdins.ExecuteReader
        r1.Read()
        r1.Close()
        cnn.close()
        Dim c As New CustomerHistory

        c.SetUp(frm, ID, Confirming.TScboCustomerHistory)

    
        '' Add elseif' for other forms as needed








        Me.Close()

    End Sub

    Private Sub btncancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancel.Click
        Me.Close()
    End Sub

    Private Sub LogPhoneCall_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub LogPhoneCall_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboSpokeWith.Items.Clear()
        Me.cboDepartment.SelectedItem = Nothing
        Me.cboInboundOutbound.Text = ""
        Me.txtNotes.Text = ""
        Me.cboSpokeWith.Text = ""

        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdget As SqlCommand = New SqlCommand("dbo.GetPCAutoNotes", cnn)
        cmdget.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        Me.cboautonote.Items.Clear()
        Me.cboautonote.Items.Add("<Add New>")
        Me.cboautonote.Items.Add("___________________________________________________")

        cnn.Open()
        r1 = cmdget.ExecuteReader
        While r1.Read
            Me.cboautonote.Items.Add(r1.Item(0))
        End While
        r1.Close()
        cnn.Close()
        Dim s = Split(Me.Contact1, " ")
        Dim c1 As String = s(0)
        Dim s2 = Split(Me.Contact2, " ")
        Dim c2 = s2(0)
        Me.cboSpokeWith.Items.Add(c1)
        If Contact2 <> " " Then
            Me.cboSpokeWith.Items.Add(c2)
            Me.cboSpokeWith.Items.Add(c1 & " and " & c2)
        Else
            Me.cboSpokeWith.Text = c1

        End If
    End Sub

   
End Class
