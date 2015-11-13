Public Class addProducts

    Private _IS_FORM_VALIDATED As Boolean = False

    Private Sub addProducts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ResetForm()
    End Sub

    Private Sub txtCancel_Click(sender As Object, e As EventArgs) Handles txtCancel.Click
        Me.txtProductModel.Clear()
        Me.txtProductName.Clear()
        Me.txtUnitMeasure.Clear()
        Me.txtCategory.Clear()
        Me.Close()
    End Sub
    Private Sub ResetForm()
        Me.txtProductModel.Clear()
        Me.txtProductName.Clear()
        Me.txtUnitMeasure.Clear()
        Me.txtCategory.Clear()
        Me.txtAcro.Clear()
    End Sub
    Private Function FormValidated(ByVal Category As String, ByVal Name As String, ByVal UnitMeasure As String, ByVal ProductModel As String)

        Dim strCatLen As Integer = Category.ToString.Length
        Dim strNameLen As Integer = Name.ToString.Length
        Dim strUnitLen As Integer = UnitMeasure.ToString.Length
        Dim strProdModel As Integer = ProductModel.ToString.Length

        '' no data entered
        If strCatLen <= 0 Then
            _IS_FORM_VALIDATED = False
        ElseIf strNameLen <= 0 Then
            _IS_FORM_VALIDATED = False
        ElseIf strUnitLen <= 0 Then
            _IS_FORM_VALIDATED = False
        ElseIf strProdModel <= 0 Then
            _IS_FORM_VALIDATED = False
        End If


        '' just category entered
        If strCatLen <= 1 Then
            _IS_FORM_VALIDATED = True
            If strNameLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strUnitLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strProdModel <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
        End If

        '' just name entered
        If strNameLen <= 1 Then
            _IS_FORM_VALIDATED = True
            If strCatLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strUnitLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strProdModel <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
        End If

        '' just Unit Measure entered
        If strUnitLen <= 1 Then
            _IS_FORM_VALIDATED = True
            If strCatLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strNameLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strProdModel <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
        End If

        '' just Product Model entered
        If strProdModel <= 1 Then
            _IS_FORM_VALIDATED = True
            If strCatLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strNameLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
            If strUnitLen <= 0 Then
                _IS_FORM_VALIDATED = False
            End If
        End If

        '' all criteria met for capitalization and submission to sql table
        ''
        If strCatLen >= 1 And strNameLen >= 1 And strUnitLen >= 1 And strProdModel >= 1 Then
            _IS_FORM_VALIDATED = True
        End If

        Return _IS_FORM_VALIDATED
    End Function

    Private Sub txtButtonOK_Click(sender As Object, e As EventArgs) Handles txtButtonOK.Click
        Dim frmValidated As Boolean = FormValidated(Me.txtCategory.Text, Me.txtProductName.Text, Me.txtUnitMeasure.Text, Me.txtProductModel.Text)
        If frmValidated = True Then
            '' all checked out then send off to be submitted.
            '' 
            '' we are here on submission
            '' 
            Dim y As New sqlOperations.MarketerOperations '' for capitalization function only 
            Dim strCat As String = y.Capitalize(Me.txtCategory.Text)
            Dim strName As String = y.Capitalize(Me.txtProductName.Text)
            Dim strModel As String = y.Capitalize(Me.txtProductModel.Text)
            Dim strMeasure As String = y.Capitalize(Me.txtUnitMeasure.Text)
            Dim strAcro As String = y.Capitalize(Me.txtAcro.Text)

            Dim x As New sqlOperations.Products
            x.InsertNewProduct(strCat, strName, strModel, strMeasure, strAcro)

            Me.ResetForm()
            Me.Close()
        ElseIf frmValidated = False Then
            MsgBox("Please check your entries again. Something went wrong.", MsgBoxStyle.Critical, "Error Adding Product")
            Exit Sub
        End If
    End Sub
End Class
