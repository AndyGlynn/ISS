Public Class frmPrint

    Private Const Test_Directory As String = "C:/Users/Clay/Desktop/Print Leads/"
    Private Const Production_Directory As String = "\\server.greenworks.local\Company\ISS\Print Leads\"

    Private exclu As Boolean = False

    Public Property Exclusions As Boolean
        Get
            Return exclu
        End Get
        Set(value As Boolean)
            exclu = value
        End Set
    End Property

    Private Sub frmPrint_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ClearCache()
        ClearListView()
        Me.wbPrint.Navigate(Production_Directory & "DEFAULT.HTM")
    End Sub


    Private Sub frmPrint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ClearCache()
        Me.wbPrint.Navigate(Production_Directory & "DEFAULT.HTM")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnPrintThisOne.Click
        Me.wbPrint.Print()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ClearCache()
        ClearListView()
        Me.Close()

    End Sub

    Private Sub ClearCache()
        Dim g As New System.IO.DirectoryInfo(Production_Directory)

        For Each y As System.IO.FileInfo In g.GetFiles("*.htm")
            If y.Name <> "DEFAULT.HTM" Then
                If y.Name <> "TESTMOCK.htm" Then
                    y.Delete()
                End If
            End If
        Next

    End Sub
    Public Sub ClearListView()
        Me.lsLeadIds.Items.Clear()

    End Sub

    Private Sub lsLeadIds_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lsLeadIds.SelectedIndexChanged
        If Me.Exclusions = False Then
            Dim a As ListViewItem

            For Each a In Me.lsLeadIds.Items
                If a.Selected = True Then
                    Dim b As New bulkPrintOperations
                    b.DoTheWork(a.Text)
                End If
            Next
        ElseIf Me.Exclusions = True Then
            Dim a As ListViewItem

            For Each a In Me.lsLeadIds.Items
                If a.Selected = True Then
                    Dim b As New bulkPrintOperations
                    Dim ex_set As bulkPrintOperations.Exclusions
                    ex_set = b.GetExclusions
                    b.DoTheWork_EXCLUSIONS(a.Text, ex_set)
                End If
            Next
        End If

    End Sub

    Private Sub btnPageLayout_Click(sender As Object, e As EventArgs) Handles btnPrinterOptions.Click
        Me.wbPrint.ShowPrintDialog()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles btnLayout.Click
        Me.wbPrint.ShowPrintPreviewDialog()
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Exclusions = False Then
            Dim arNumbers As New ArrayList
            For Each y As ListViewItem In Me.lsLeadIds.Items
                If y.Selected = True Then
                    arNumbers.Add(y.Text)
                End If
            Next


            Dim yy As New bulkPrintOperations
            'MsgBox("String Generated: " & vbCrLf & vbCrLf & yy.Generate_BULK_MSG_BODY(arNumbers), MsgBoxStyle.Information, "DEBUG STRING GENERATION BULK PRINT OBJ")
            Dim strMSG As String = yy.Generate_BULK_MSG_BODY(arNumbers)
            yy.GenerateBULK_PRINT(strMSG)



            'Dim str As String = ""
            'Dim i As Integer = 0
            'For i = 0 To arNumbers.Count - 1
            '    str += arNumbers(i) & vbCrLf
            'Next

            'MsgBox("Record IDs From Multi-Select : " & vbCrLf & str, MsgBoxStyle.Information, "DEBUG INFO")
        ElseIf Me.Exclusions = True Then

            Dim arNumbers As New ArrayList
            For Each y As ListViewItem In Me.lsLeadIds.Items
                If y.Selected = True Then
                    arNumbers.Add(y.Text)
                End If
            Next


            Dim yy As New bulkPrintOperations
            Dim ex_set As bulkPrintOperations.Exclusions
            ex_set = yy.GetExclusions
            'MsgBox("String Generated: " & vbCrLf & vbCrLf & yy.Generate_BULK_MSG_BODY(arNumbers), MsgBoxStyle.Information, "DEBUG STRING GENERATION BULK PRINT OBJ")
            Dim strMSG As String = yy.Generate_BULK_MSG_BODY_EXCLUSIONS(arNumbers, ex_set)
            yy.GenerateBULK_PRINT(strMSG)



            'Dim str As String = ""
            'Dim i As Integer = 0
            'For i = 0 To arNumbers.Count - 1
            '    str += arNumbers(i) & vbCrLf
            'Next

            'MsgBox("Record IDs From Multi-Select : " & vbCrLf & str, MsgBoxStyle.Information, "DEBUG INFO")
        End If
        
    End Sub
End Class
