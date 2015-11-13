Public Class frmViewEditProducts
    Public Pr_SelItem As ListViewItem

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim x As New ViewEditProducts
        Me.ResetDefault()
        x.PopuluateList()

        Dim d
        For d = 0 To x.ArProdAcro.Count - 1
            Dim lv As New ListViewItem
            lv.Text = x.ArProducts(d).ToString
            lv.SubItems.Add(x.ArProdAcro(d).ToString)
            Me.lstProducts.Items.Add(lv)
        Next

    End Sub
    Private Sub ResetDefault()
        Me.lstProducts.Items.Clear()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ResetDefault()
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim y As New ViewEditProducts

        Dim pr As String = ""
        Dim pr_a As String = ""
        pr = InputBox$("Enter the new product to add. IE 'Siding'.", "Enter New Product", "")
        If pr.ToString.Length <= 0 Then
            MsgBox("You cannot add a blank product. Please try again.", MsgBoxStyle.Exclamation, "Error Adding Product")
            Exit Sub
        ElseIf pr.ToString.Length >= 1 Then
            pr_a = InputBox$("Please enter  the acronym for product '" & pr & ".", "Enter Product Acronym", "")
            If pr_a.ToString.Length <= 0 Then
                MsgBox("You cannot enter a blank product acronym.", MsgBoxStyle.Exclamation, "Error Adding Product")
                Exit Sub
            ElseIf pr_a.ToString.Length >= 1 Then
                y.Add_Product(pr, pr_a)
                Me.lstProducts.Items.Clear()
                y.PopuluateList()
                Dim d
                For d = 0 To y.ArProdAcro.Count - 1
                    Dim lv As New ListViewItem
                    lv.Text = y.ArProducts(d).ToString
                    lv.SubItems.Add(y.ArProdAcro(d).ToString)
                    Me.lstProducts.Items.Add(lv)
                Next
            End If
        End If

    End Sub
End Class
