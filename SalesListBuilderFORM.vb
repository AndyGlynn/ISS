Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System


Public Class SalesListBuilder
    Public Loading As Boolean = True
    Public Rollback As String

    Private Sub cboCity_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCity.SelectedIndexChanged
        Sales.CityState = Me.cboCity.Text
    End Sub

    Private Sub cboRep_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRep.SelectedIndexChanged
        Sales.Rep = Me.cboRep.Text
    End Sub

    Private Sub cboProduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProduct.SelectedIndexChanged
        Sales.Product = Me.cboProduct.Text
    End Sub

    Private Sub cboPLS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPLS.SelectedIndexChanged
        Sales.PLS = Me.cboPLS.Text
        Me.cboSLS.Items.Clear()
        Me.cboSLS.Items.Add("")
        If Me.cboPLS.Text <> "" Then
            Me.cboSLS.Items.Clear()
            Me.cboSLS.Items.Add("")
            Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet As SqlCommand
            Dim param1 As SqlParameter = New SqlParameter("@PLS", Me.cboPLS.Text)
            Dim param2 As SqlParameter = New SqlParameter("@Fork", "SLS")
            Dim r As SqlDataReader
            cmdGet = New SqlCommand("dbo.PopulateSalesListBuilder", cnn)
            cmdGet.CommandType = CommandType.StoredProcedure
            cmdGet.Parameters.Add(param1)
            cmdGet.Parameters.Add(param2)
            cnn.Open()
            r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
            While r.Read
                Me.cboSLS.Items.Add(r.Item(0))
            End While
            r.Close()
            cnn.Close()
        End If
    End Sub

    Private Sub cboSLS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSLS.SelectedIndexChanged

        Sales.SLS = Me.cboSLS.Text
    End Sub

    Private Sub chkRecovery_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecovery.CheckedChanged
        If Me.chkRecovery.Checked = True Then
            Sales.Recovery = True
        Else
            Sales.Recovery = False
        End If
    End Sub

    Private Sub lblCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCheckAll.Click

        For i As Integer = 0 To Me.chklbResult.Items.Count - 1
            Me.chklbResult.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub lblUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblUncheckAll.Click
        For i As Integer = 0 To Me.chklbResult.Items.Count - 1
            Me.chklbResult.SetItemChecked(i, False)
        Next
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If Me.chklbResult.CheckedItems.Count = 0 Then
            MsgBox("You Must Check at Least 1 Sale Result!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim i As Integer = Me.chklbResult.CheckedItems.Count
        Select Case i
            Case Is = 1
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(0).ToString()
            Case Is = 2
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(1).ToString()
            Case Is = 3
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(2).ToString()
            Case Is = 4
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(3).ToString()
            Case Is = 5
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(4).ToString()
            Case Is = 6
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(5).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(5).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(5).ToString()
            Case Is = 7
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(5).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(6).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(6).ToString()
            Case Is = 8
                Sales.R1 = Me.chklbResult.CheckedItems(0).ToString()
                Sales.R2 = Me.chklbResult.CheckedItems(1).ToString()
                Sales.R3 = Me.chklbResult.CheckedItems(2).ToString()
                Sales.R4 = Me.chklbResult.CheckedItems(3).ToString()
                Sales.R5 = Me.chklbResult.CheckedItems(4).ToString()
                Sales.R6 = Me.chklbResult.CheckedItems(5).ToString()
                Sales.R7 = Me.chklbResult.CheckedItems(6).ToString()
                Sales.R8 = Me.chklbResult.CheckedItems(7).ToString()
        End Select
        Me.Close()
        Dim c As New SalesListManager


    End Sub

    Private Sub SalesListBuilder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Dispose()
    End Sub

    Private Sub SalesListBuilder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.cboRep.Items.Clear()
        Me.cboRep.Items.Add("")
        Me.cboProduct.Items.Clear()
        Me.cboProduct.Items.Add("")
        Me.cboCity.Items.Clear()
        Me.cboCity.Items.Add("")
        Me.cboPLS.Items.Clear()
        Me.cboPLS.Items.Add("")
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        Dim param1 As SqlParameter = New SqlParameter("@PLS", "")
        Dim param2 As SqlParameter = New SqlParameter("@Fork", "Load")
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("dbo.PopulateSalesListBuilder", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        While r.Read
            If r.Item(1) = "Rep" Then
                Me.cboRep.Items.Add(r.Item(0))
            ElseIf r.Item(1) = "Product" Then
                Me.cboProduct.Items.Add(r.Item(0))
            ElseIf r.Item(1) = "PLS" Then
                Me.cboPLS.Items.Add(r.Item(0))
            ElseIf r.Item(1) = "CityState" Then
                Me.cboCity.Items.Add(r.Item(0))
            End If
        End While
        r.Close()
        cnn.Close()


        Me.Loading = True
        Me.cboRep.Text = Sales.Rep
        Me.cboProduct.Text = Sales.Product
        Me.cboCity.Text = Sales.CityState
        Me.cboPLS.Text = Sales.PLS
        Me.cboSLS.Text = Sales.SLS
        If Sales.Recovery = True Then
            Me.chkRecovery.Checked = True
        Else
            Me.chkRecovery.Checked = False
        End If
        For i As Integer = 0 To Me.chklbResult.Items.Count - 1
            Me.chklbResult.SetItemChecked(i, False)
        Next



        If Sales.R1 = "Demo/No Sale" Or Sales.R2 = "Demo/No Sale" Or Sales.R3 = "Demo/No Sale" Or Sales.R4 = "Demo/No Sale" Or Sales.R5 = "Demo/No Sale" Or Sales.R6 = "Demo/No Sale" Or Sales.R7 = "Demo/No Sale" Or Sales.R8 = "Demo/No Sale" Then
            Me.chklbResult.SetItemChecked(0, True)
        End If
        If Sales.R1 = "No Demo" Or Sales.R2 = "No Demo" Or Sales.R3 = "No Demo" Or Sales.R4 = "No Demo" Or Sales.R5 = "No Demo" Or Sales.R6 = "No Demo" Or Sales.R7 = "No Demo" Or Sales.R8 = "No Demo" Then
            Me.chklbResult.SetItemChecked(1, True)
        End If
        If Sales.R1 = "Reset" Or Sales.R2 = "Reset" Or Sales.R3 = "Reset" Or Sales.R4 = "Reset" Or Sales.R5 = "Reset" Or Sales.R6 = "Reset" Or Sales.R7 = "Reset" Or Sales.R8 = "Reset" Then
            Me.chklbResult.SetItemChecked(2, True)
        End If
        If Sales.R1 = "Not Hit" Or Sales.R2 = "Not Hit" Or Sales.R3 = "Not Hit" Or Sales.R4 = "Not Hit" Or Sales.R5 = "Not Hit" Or Sales.R6 = "Not Hit" Or Sales.R7 = "Not Hit" Or Sales.R8 = "Not Hit" Then
            Me.chklbResult.SetItemChecked(3, True)
        End If
        If Sales.R1 = "Not Issued" Or Sales.R2 = "Not Issued" Or Sales.R3 = "Not Issued" Or Sales.R4 = "Not Issued" Or Sales.R5 = "Not Issued" Or Sales.R6 = "Not Issued" Or Sales.R7 = "Not Issued" Or Sales.R8 = "Not Issued" Then
            Me.chklbResult.SetItemChecked(4, True)
        End If
        If Sales.R1 = "Sale" Or Sales.R2 = "Sale" Or Sales.R3 = "Sale" Or Sales.R4 = "Sale" Or Sales.R5 = "Sale" Or Sales.R6 = "Sale" Or Sales.R7 = "Sale" Or Sales.R8 = "Sale" Then
            Me.chklbResult.SetItemChecked(5, True)
        End If
        If Sales.R1 = "Recission Cancel" Or Sales.R2 = "Recission Cancel" Or Sales.R3 = "Recission Cancel" Or Sales.R4 = "Recission Cancel" Or Sales.R5 = "Recission Cancel" Or Sales.R6 = "Recission Cancel" Or Sales.R7 = "Recission Cancel" Or Sales.R8 = "Recission Cancel" Then
            Me.chklbResult.SetItemChecked(6, True)
        End If
        If Sales.R1 = "No Results" Or Sales.R2 = "No Results" Or Sales.R3 = "No Results" Or Sales.R4 = "No Results" Or Sales.R5 = "No Results" Or Sales.R6 = "No Results" Or Sales.R7 = "No Results" Or Sales.R8 = "No Results" Then
            Me.chklbResult.SetItemChecked(7, True)
        End If

        Me.Loading = False
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Me.cboRep.Text = ""
        Me.cboProduct.Text = ""
        Me.cboCity.Text = ""
        Me.cboPLS.Text = ""
        Me.cboSLS.Text = ""
        Me.chkRecovery.Checked = False
        For i As Integer = 0 To Me.chklbResult.Items.Count - 1
            Me.chklbResult.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
        Sales.cboSalesList.Text = Me.Rollback
    End Sub
End Class
