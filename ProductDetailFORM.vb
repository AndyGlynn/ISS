
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic


Public Class ProductDetail
    Public Product As String
    Public Prodnum As Integer
#Region "My Sub Procedures"

    Private Sub Setup()
        Me.GetItems(Me.cboUnit)
        Me.Text = "Product Detail for " & Me.Product
        Select Case Me.Prodnum
            Case Is = 1
                If SDResult.PM1 <> Nothing Then
                    Me.cboMan.Text = SDResult.PM1
                    If Me.cboMan.Text = Nothing Then
                        Me.cboMan.Items.Add(SDResult.PM1)
                        Me.cboMan.Text = SDResult.PM1
                    End If
                End If
                If SDResult.PModel1 <> Nothing Then
                    Me.cboMake.Text = SDResult.PModel1
                    If Me.cboMake.Text = Nothing Then
                        Me.cboMake.Items.Add(SDResult.PModel1)
                        Me.cboMake.Text = SDResult.PModel1
                    End If
                End If
                If SDResult.PStyle1 <> Nothing Then
                    Me.cboStyle.Text = SDResult.PStyle1
                    If Me.cboStyle.Text = Nothing Then
                        Me.cboStyle.Items.Add(SDResult.PStyle1)
                        Me.cboStyle.Text = SDResult.PStyle1
                    End If
                End If
                If SDResult.PColor1 <> Nothing Then
                    Me.cboColor.Text = SDResult.PColor1
                    If Me.cboColor.Text = Nothing Then
                        Me.cboColor.Items.Add(SDResult.PColor1)
                        Me.cboColor.Text = SDResult.PColor1
                    End If
                End If
                If SDResult.Qty1 <> 0 Then
                    Me.spnQTY.Value = SDResult.Qty1
                End If
                If SDResult.Unit1 <> Nothing Then
                    Me.cboUnit.Text = SDResult.Unit1
                    If Me.cboUnit.Text = Nothing Then
                        Me.cboUnit.Items.Add(SDResult.Unit1)
                        Me.cboUnit.Text = SDResult.Unit1
                    End If
                End If
            Case Is = 2
                If SDResult.PM2 <> Nothing Then
                    Me.cboMan.Text = SDResult.PM2
                    If Me.cboMan.Text = Nothing Then
                        Me.cboMan.Items.Add(SDResult.PM2)
                        Me.cboMan.Text = SDResult.PM2
                    End If
                End If
                If SDResult.PModel2 <> Nothing Then
                    Me.cboMake.Text = SDResult.PModel2
                    If Me.cboMake.Text = Nothing Then
                        Me.cboMake.Items.Add(SDResult.PModel2)
                        Me.cboMake.Text = SDResult.PModel2
                    End If
                End If
                If SDResult.PStyle2 <> Nothing Then
                    Me.cboStyle.Text = SDResult.PStyle2
                    If Me.cboStyle.Text = Nothing Then
                        Me.cboStyle.Items.Add(SDResult.PStyle2)
                        Me.cboStyle.Text = SDResult.PStyle2
                    End If
                End If
                If SDResult.PColor2 <> Nothing Then
                    Me.cboColor.Text = SDResult.PColor2
                    If Me.cboColor.Text = Nothing Then
                        Me.cboColor.Items.Add(SDResult.PColor2)
                        Me.cboColor.Text = SDResult.PColor2
                    End If
                End If
                If SDResult.Qty2 <> 0 Then
                    Me.spnQTY.Value = SDResult.Qty2
                End If
                If SDResult.Unit2 <> Nothing Then
                    Me.cboUnit.Text = SDResult.Unit2
                    If Me.cboUnit.Text = Nothing Then
                        Me.cboUnit.Items.Add(SDResult.Unit2)
                        Me.cboUnit.Text = SDResult.Unit2
                    End If
                End If
            Case Is = 3
                If SDResult.PM3 <> Nothing Then
                    Me.cboMan.Text = SDResult.PM3
                    If Me.cboMan.Text = Nothing Then
                        Me.cboMan.Items.Add(SDResult.PM3)
                        Me.cboMan.Text = SDResult.PM3
                    End If
                End If
                If SDResult.PModel3 <> Nothing Then
                    Me.cboMake.Text = SDResult.PModel3
                    If Me.cboMake.Text = Nothing Then
                        Me.cboMake.Items.Add(SDResult.PModel3)
                        Me.cboMake.Text = SDResult.PModel3
                    End If
                End If
                If SDResult.PStyle3 <> Nothing Then
                    Me.cboStyle.Text = SDResult.PStyle3
                    If Me.cboStyle.Text = Nothing Then
                        Me.cboStyle.Items.Add(SDResult.PStyle3)
                        Me.cboStyle.Text = SDResult.PStyle3
                    End If
                End If
                If SDResult.PColor3 <> Nothing Then
                    Me.cboColor.Text = SDResult.PColor3
                    If Me.cboColor.Text = Nothing Then
                        Me.cboColor.Items.Add(SDResult.PColor3)
                        Me.cboColor.Text = SDResult.PColor3
                    End If
                End If
                If SDResult.Qty3 <> 0 Then
                    Me.spnQTY.Value = SDResult.Qty3
                End If
                If SDResult.Unit3 <> Nothing Then
                    Me.cboUnit.Text = SDResult.Unit3
                    If Me.cboUnit.Text = Nothing Then
                        Me.cboUnit.Items.Add(SDResult.Unit3)
                        Me.cboUnit.Text = SDResult.Unit3
                    End If
                End If
            Case Is = 4
                If SDResult.PM4 <> Nothing Then
                    Me.cboMan.Text = SDResult.PM4
                    If Me.cboMan.Text = Nothing Then
                        Me.cboMan.Items.Add(SDResult.PM4)
                        Me.cboMan.Text = SDResult.PM4
                    End If
                End If
                If SDResult.PModel4 <> Nothing Then
                    Me.cboMake.Text = SDResult.PModel4
                    If Me.cboMake.Text = Nothing Then
                        Me.cboMake.Items.Add(SDResult.PModel4)
                        Me.cboMake.Text = SDResult.PModel4
                    End If
                End If
                If SDResult.PStyle4 <> Nothing Then
                    Me.cboStyle.Text = SDResult.PStyle4
                    If Me.cboStyle.Text = Nothing Then
                        Me.cboStyle.Items.Add(SDResult.PStyle4)
                        Me.cboStyle.Text = SDResult.PStyle4
                    End If
                End If
                If SDResult.PColor4 <> Nothing Then
                    Me.cboColor.Text = SDResult.PColor4
                    If Me.cboColor.Text = Nothing Then
                        Me.cboColor.Items.Add(SDResult.PColor4)
                        Me.cboColor.Text = SDResult.PColor4
                    End If
                End If
                If SDResult.Qty4 <> 0 Then
                    Me.spnQTY.Value = SDResult.Qty4
                End If
                If SDResult.Unit4 <> Nothing Then
                    Me.cboUnit.Text = SDResult.Unit4
                    If Me.cboUnit.Text = Nothing Then
                        Me.cboUnit.Items.Add(SDResult.Unit4)
                        Me.cboUnit.Text = SDResult.Unit4
                    End If
                End If
            Case Is = 5
                If SDResult.PM5 <> Nothing Then
                    Me.cboMan.Text = SDResult.PM5
                    If Me.cboMan.Text = Nothing Then
                        Me.cboMan.Items.Add(SDResult.PM5)
                        Me.cboMan.Text = SDResult.PM5
                    End If
                End If
                If SDResult.PModel5 <> Nothing Then
                    Me.cboMake.Text = SDResult.PModel5
                    If Me.cboMake.Text = Nothing Then
                        Me.cboMake.Items.Add(SDResult.PModel5)
                        Me.cboMake.Text = SDResult.PModel5
                    End If
                End If
                If SDResult.PStyle5 <> Nothing Then
                    Me.cboStyle.Text = SDResult.PStyle5
                    If Me.cboStyle.Text = Nothing Then
                        Me.cboStyle.Items.Add(SDResult.PStyle5)
                        Me.cboStyle.Text = SDResult.PStyle5
                    End If
                End If
                If SDResult.PColor5 <> Nothing Then
                    Me.cboColor.Text = SDResult.PColor5
                    If Me.cboColor.Text = Nothing Then
                        Me.cboColor.Items.Add(SDResult.PColor5)
                        Me.cboColor.Text = SDResult.PColor5
                    End If
                End If
                If SDResult.Qty5 <> 0 Then
                    Me.spnQTY.Value = SDResult.Qty5
                End If
                If SDResult.Unit5 <> Nothing Then
                    Me.cboUnit.Text = SDResult.Unit5
                    If Me.cboUnit.Text = Nothing Then
                        Me.cboUnit.Items.Add(SDResult.Unit5)
                        Me.cboUnit.Text = SDResult.Unit5
                    End If
                End If
        End Select
    End Sub

    Private Sub AddNew(ByVal sender As ComboBox)
        Dim y As String
        Select Case sender.Name
            Case Is = "cboMan"
                y = "Manufacturer"
            Case Is = "cboMake"
                y = "Model"
            Case Is = "cboStyle"
                y = "Style"
            Case Is = "cboColor"
                y = "Color"
            Case Is = "cboUnit"
                y = "Unit"
        End Select
        Dim x As String = InputBox("Enter New " & y & ".", "Add New " & y)
        x = Trim(x)
        x = StrConv(x, VbStrConv.ProperCase)
        If x = "" Then
            Exit Sub
        End If
        Dim Cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.AddNewProductDetail", Cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Control", y)
        Dim param2 As SqlParameter = New SqlParameter("@Product", Me.Product)
        Dim param3 As SqlParameter = New SqlParameter("@Manufacturer", Me.cboMan.Text)
        Dim param4 As SqlParameter = New SqlParameter("@Model", Me.cboMake.Text)
        Dim param5 As SqlParameter = New SqlParameter("@New", x)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        Cnn.Open()
        Dim R1 As SqlDataReader
        R1 = cmdGet.ExecuteReader
        R1.Read()
        sender.Items.Add(R1.Item(0))
        sender.Text = R1.Item(0)
        R1.Close()
        Cnn.Close()
    End Sub

    Private Sub GetItems(ByVal sender As ComboBox)
        Dim col As String
        Select Case sender.Name
            Case Is = "cboMan"
                col = "Manufacturer"
            Case Is = "cboMake"
                col = "Model"
            Case Is = "cboStyle"
                col = "Style"
            Case Is = "cboColor"
                col = "Color"
            Case Is = "cboUnit"
                col = "Unit"
        End Select
        Dim Cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand = New SqlCommand("dbo.PopulateProductDetail", Cnn)
        Dim param1 As SqlParameter = New SqlParameter("@Col", col)
        Dim param2 As SqlParameter = New SqlParameter("@Product", Me.Product)
        Dim param3 As SqlParameter = New SqlParameter("@Manufacturer", Me.cboMan.Text)
        Dim param4 As SqlParameter = New SqlParameter("@Model", Me.cboMake.Text)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        Cnn.Open()
        Dim R1 As SqlDataReader
        R1 = cmdGet.ExecuteReader
        sender.Items.Clear()
        sender.Items.Add("<Add New>")
        sender.Items.Add("___________________________________________________")
        sender.Items.Add("")
        While R1.Read
            sender.Items.Add(R1.Item(0))
        End While
        R1.Close()
        Cnn.Close()
    End Sub

#End Region

    Private Sub ProductDetail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.GetItems(Me.cboMan)
        Me.cboMan.Text = Nothing
        Me.cboMake.Text = Nothing
        Me.cboStyle.Text = Nothing
        Me.cboColor.Text = Nothing
        Me.spnQTY.Value = 0
        Me.Setup()
    End Sub

    Private Sub cboMake_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMake.GotFocus
        If Me.cboMan.Text = Nothing Then
            Dim tt As New ToolTip
            tt.IsBalloon = True
            tt.ToolTipIcon = ToolTipIcon.Info
            tt.ToolTipTitle = "Select Manufacturer"
            tt.Show("You must select a manufacturer" & vbCr & "before you select a model!", Me.cboMake, 180, -70)
            Me.cboMan.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub cboStyle_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboStyle.GotFocus
        If Me.cboMake.Text = Nothing Then
            Dim tt As New ToolTip
            tt.IsBalloon = True
            tt.ToolTipIcon = ToolTipIcon.Info
            tt.ToolTipTitle = "Select Model"
            tt.Show("You must select a model" & vbCr & "before you select a style!", Me.cboStyle, 180, -70)
            If Me.cboMake.Text <> Nothing Then
                Me.cboMake.Focus()
            Else
                Me.cboMan.Focus()
            End If
            Exit Sub
        End If
    End Sub

    Private Sub cboColor_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboColor.GotFocus
        If Me.cboMake.Text = Nothing Then
            Dim tt As New ToolTip
            tt.IsBalloon = True
            tt.ToolTipIcon = ToolTipIcon.Info
            tt.ToolTipTitle = "Select Model"
            tt.Show("You must select a model" & vbCr & "before you select a color!", Me.cboColor, 180, -70)
            If Me.cboMake.Text <> Nothing Then
                Me.cboMake.Focus()
            Else
                Me.cboMan.Focus()
            End If

            Exit Sub
        End If
    End Sub

    Private Sub cboMan_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMan.SelectedValueChanged
        Me.cboMake.Text = Nothing
        Me.cboStyle.Text = Nothing
        Me.cboColor.Text = Nothing
        If Me.cboMan.Text = "___________________________________________________" Then
            Me.cboMan.Text = Nothing
            Exit Sub
        End If
        If Me.cboMan.Text = "<Add New>" Then
            Me.cboMan.Text = Nothing
            Me.AddNew(Me.cboMan)
            Exit Sub
        End If
        Me.GetItems(Me.cboMake)

    End Sub

    Private Sub cboMake_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboMake.SelectedValueChanged

        If Me.cboMake.Text = "___________________________________________________" Then
            Me.cboMake.Text = Nothing
            Exit Sub
        End If
        If Me.cboMake.Text = "<Add New>" Then
            If Me.cboMan.Text = Nothing Then
                Me.cboMake.Text = Nothing
                MsgBox("You must supply a Manufacturer to add a new Model!", MsgBoxStyle.Exclamation, "No Manufacturer")
                Me.cboMan.Focus()
                Exit Sub
            End If
            Me.cboMake.Text = Nothing
            Me.AddNew(Me.cboMake)
            Exit Sub
        End If
        Me.GetItems(Me.cboStyle)
        Me.GetItems(Me.cboColor)
    End Sub

    Private Sub cboStyle_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboStyle.SelectedValueChanged
        If Me.cboStyle.Text = "___________________________________________________" Then
            Me.cboStyle.Text = Nothing
            Exit Sub
        End If
        If Me.cboStyle.Text = "<Add New>" Then
            If Me.cboMan.Text = Nothing And Me.cboMake.Text = Nothing Then
                Me.cboStyle.Text = Nothing
                MsgBox("You must supply a Manufacturer and a Model to add a new Style!", MsgBoxStyle.Exclamation, "No Manufacturer")
                Me.cboMan.Focus()
                Exit Sub
            End If
            Me.cboStyle.Text = Nothing
            Me.AddNew(Me.cboStyle)
            Exit Sub
        End If
    End Sub

    Private Sub cboColor_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboColor.SelectedValueChanged
        If Me.cboColor.Text = "___________________________________________________" Then
            Me.cboColor.Text = Nothing
            Exit Sub
        End If
        If Me.cboColor.Text = "<Add New>" Then
            If Me.cboMan.Text = Nothing And Me.cboMake.Text = Nothing Then
                Me.cboColor.Text = Nothing
                MsgBox("You must supply a Manufacturer and a Model to add a new Color!", MsgBoxStyle.Exclamation, "No Manufacturer")
                Me.cboMan.Focus()
                Exit Sub
            End If
            Me.cboColor.Text = Nothing
            Me.AddNew(Me.cboColor)
            Exit Sub
        End If
    End Sub

    Private Sub cboUnit_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUnit.SelectedValueChanged
        If Me.cboUnit.Text = "___________________________________________________" Then
            Me.cboUnit.Text = Nothing
            Exit Sub
        End If
        If Me.cboUnit.Text = "<Add New>" Then
            Me.cboUnit.Text = Nothing
            Me.AddNew(Me.cboUnit)
            Exit Sub
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btnSaveClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveClose.Click
        If Me.spnQTY.Value <> 0 And Me.cboUnit.Text = Nothing Then
            MsgBox("You must Supply a Unit to track Quantity!", MsgBoxStyle.Exclamation, "Unit Cannot Be Blank")
            Me.cboUnit.Focus()
            Exit Sub
        End If


        Select Case Me.Prodnum
            Case Is = 1
                SDResult.PM1 = Me.cboMan.Text
                SDResult.PModel1 = Me.cboMake.Text
                SDResult.PStyle1 = Me.cboStyle.Text
                SDResult.PColor1 = Me.cboColor.Text
                SDResult.Unit1 = Me.cboUnit.Text
                SDResult.Qty1 = Me.spnQTY.Value
                If Me.cboMan.Text <> Nothing Or Me.cboUnit.Text <> Nothing Then
                    SDResult.lnkP1.LinkVisited = True
                Else
                    SDResult.lnkP1.LinkVisited = False
                End If
            Case Is = 2
                SDResult.PM2 = Me.cboMan.Text
                SDResult.PModel2 = Me.cboMake.Text
                SDResult.PStyle2 = Me.cboStyle.Text
                SDResult.PColor2 = Me.cboColor.Text
                SDResult.Unit2 = Me.cboUnit.Text
                SDResult.Qty2 = Me.spnQTY.Value
                If Me.cboMan.Text <> Nothing Or Me.cboUnit.Text <> Nothing Then
                    SDResult.lnkP2.LinkVisited = True
                Else
                    SDResult.lnkP2.LinkVisited = False
                End If
            Case Is = 3
                SDResult.PM3 = Me.cboMan.Text
                SDResult.PModel3 = Me.cboMake.Text
                SDResult.PStyle3 = Me.cboStyle.Text
                SDResult.PColor3 = Me.cboColor.Text
                SDResult.Unit3 = Me.cboUnit.Text
                SDResult.Qty3 = Me.spnQTY.Value
                If Me.cboMan.Text <> Nothing Or Me.cboUnit.Text <> Nothing Then
                    SDResult.lnkP3.LinkVisited = True
                Else
                    SDResult.lnkP3.LinkVisited = False
                End If
            Case Is = 4
                SDResult.PM4 = Me.cboMan.Text
                SDResult.PModel4 = Me.cboMake.Text
                SDResult.PStyle4 = Me.cboStyle.Text
                SDResult.PColor4 = Me.cboColor.Text
                SDResult.Unit4 = Me.cboUnit.Text
                SDResult.Qty4 = Me.spnQTY.Value
                If Me.cboMan.Text <> Nothing Or Me.cboUnit.Text <> Nothing Then
                    SDResult.lnkP4.LinkVisited = True
                Else
                    SDResult.lnkP4.LinkVisited = False
                End If
            Case Is = 5
                SDResult.PM5 = Me.cboMan.Text
                SDResult.PModel5 = Me.cboMake.Text
                SDResult.PStyle5 = Me.cboStyle.Text
                SDResult.PColor5 = Me.cboColor.Text
                SDResult.Unit5 = Me.cboUnit.Text
                SDResult.Qty5 = Me.spnQTY.Value
                If Me.cboMan.Text <> Nothing Or Me.cboUnit.Text <> Nothing Then
                    SDResult.lnkP5.LinkVisited = True
                Else
                    SDResult.lnkP5.LinkVisited = False
                End If
        End Select
        Me.Close()
        Me.Dispose()

    End Sub
End Class
