Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Public Class SDResult


    
#Region "Form Operators"
    Public LHID As String = "0"
    Public ID As String
    Public EditMode As Boolean = False
    Public Editmodetext As String
    Public SeqNum As Integer = 0
    Public StepCount As Integer = 1
#End Region
#Region "Get Data Variables"
    '' Data Needed for new result 
    Public Rep1 As String
    Public Rep2 As String
    Public NR As Boolean
    Public Product1 As String
    Public Product2 As String
    Public Product3 As String
    Public ApptDate As DateTime
    Public ApptDay As String
    Public Recoverable As Boolean = False
    '' Data needed for edit result

#End Region
    Private Sub ProductList()
        Me.cboProducts.Items.Clear()
        Me.cboProducts.Items.Add("<Add New>")
        Me.cboProducts.Items.Add("___________________________________________________")
        Me.cboProducts.Items.Add("")
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("dbo.GetProducts", cnn)
        cmdGet.CommandType = CommandType.StoredProcedure
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        While r.Read
            Me.cboProducts.Items.Add(r.Item(0))
        End While
        r.Close()
        cnn.Close()
    End Sub
    Private Sub GetData()
        '' Load Current Rep List 
        Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet2 As SqlCommand
        Dim r As SqlDataReader
        cmdGet2 = New SqlCommand("dbo.GetSalesReps", cnn1)
        cmdGet2.CommandType = CommandType.StoredProcedure

        cnn1.Open()
        r = cmdGet2.ExecuteReader(CommandBehavior.CloseConnection)
        Me.cboRep1.Items.Clear()
        Me.cboRep2.Items.Clear()
        Me.cboRep1.Items.Add("")
        Me.cboRep2.Items.Add("")
        While r.Read
            Me.cboRep1.Items.Add(r.Item(0) & " " & r.Item(1))
            Me.cboRep2.Items.Add(r.Item(0) & " " & r.Item(1))
        End While
        r.Close()
        cnn1.Close()
        ''Gets current needed info from enterlead 
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet1 As SqlCommand
        Dim r1 As SqlDataReader
        cmdGet1 = New SqlCommand("Select Rep1, Rep2, NeedssaleResult, Product1, Product2, Product3, ApptDate, ApptDay, isrecovery,ispreviouscustomer from enterlead where ID = " & Me.ID, cnn)
        cmdGet1.CommandType = CommandType.Text
        cnn.Open()
        r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
        r1.Read()
        '' Loads Names of Latest Rep from sales rep pull list and will also add 
        '' the name to rep combos if they are not part of the current rep list 
        Try
            Me.cboRep1.Text = r1.Item(0)
            If Me.cboRep1.Text = "" And r1.Item(0) <> "" Then
                Me.cboRep1.Items.Add(r1.Item(0))
                Me.cboRep2.Items.Add(r1.Item(0))
                Me.cboRep1.Text = r1.Item(0)
                Me.Rep1 = r1.Item(0)
            End If
        Catch ex As Exception
            Me.cboRep1.Text = Nothing
            Me.Rep1 = ""
        End Try
        Try
            Me.cboRep2.Text = r1.Item(1)
            If Me.cboRep2.Text = "" And r1.Item(1) <> "" Then
                Me.cboRep2.Items.Add(r1.Item(1))
                Me.cboRep1.Items.Add(r1.Item(1))
                Me.cboRep2.Text = r1.Item(1)
                Me.Rep2 = r1.Item(1)
            End If
        Catch ex As Exception
            Me.cboRep2.Text = Nothing
            Me.Rep2 = ""
        End Try
        Me.NR = r1.Item(2)
        Me.Product1 = r1.Item(3)
        Me.Product2 = r1.Item(4)
        Me.Product3 = r1.Item(5)
        Me.ApptDate = r1.Item(6)
        Me.ApptDay = r1.Item(7)
        If r1.Item(8) = True And r1.Item(9) = False Then
            Me.pctRecoveryPC.Image = Me.ImageList1.Images(5)
            Me.lblRecoveryPC.Text = "Recovery Result"
        ElseIf r1.Item(8) = True And r1.Item(9) = True Then
            Me.pctRecoveryPC.Image = Me.ImageList1.Images(5)
            Me.lblRecoveryPC.Text = "Recovery Result"
        ElseIf r1.Item(8) = False And r1.Item(9) = True Then
            Me.pctRecoveryPC.Image = Me.ImageList1.Images(6)
            Me.lblRecoveryPC.Text = "Previous Customer"
        Else
            Me.pctRecoveryPC.Image = Nothing
            Me.lblRecoveryPC.Text = ""
        End If
        r.Close()
        cnn.Close()

        ''   Sets Date Picker to Assumed Recission Date

        Select Case Me.ApptDay
            Case Is = "Saturday", "Friday", "Thursday"
                Me.dtpRecDate.Value = Me.ApptDate.AddDays(5.0)
            Case Is = "Sunday", "Monday", "Tuesday", "Wednesday"
                Me.dtpRecDate.Value = Me.ApptDate.AddDays(4.0)
        End Select

        '' Populates Previous Sales Results for editing purposes
        Me.lvPrevious.Items.Clear()
        Dim cnn2 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet3 As SqlCommand
        Dim r2 As SqlDataReader
        cmdGet3 = New SqlCommand("Select ApptDate,SResult,Rep1, Rep2,QuotedSold, ParPrice, Recoverable,Notes,ID , USRDefined from Leadhistory where leadnum = " & Me.ID & " and Sresult <> '' and Sresult is not null order by TriggerDate desc", cnn2)
        cmdGet3.CommandType = CommandType.Text
        cnn2.Open()
        r2 = cmdGet3.ExecuteReader(CommandBehavior.CloseConnection)
        While r2.Read()
            Dim lv As New ListViewItem
            Try
                If r2.Item(9) = "System Required" Then
                    lv.ImageIndex = 3
                    lv.Tag = "Locked"
                End If
            Catch ex As Exception
            End Try
            lv.SubItems.Add(CType(r2.Item(0), Date).ToShortDateString)
            lv.SubItems.Add(r2.Item(1))
            lv.SubItems.Add(r2.Item(2))
            Try
                lv.SubItems.Add(r2.Item(3))
            Catch ex As Exception
                lv.SubItems.Add("")
            End Try
            lv.SubItems.Add(r2.Item(4))
            lv.SubItems.Add(r2.Item(5))
            lv.SubItems.Add(r2.Item(6))
            lv.SubItems.Add(r2.Item(7))
            lv.SubItems.Add(r2.Item(8))
            Me.lvPrevious.Items.Add(lv)
        End While
        r2.Close()
        cnn2.Close()
    End Sub

#Region "Sale Detail Variables"
    Public P1 As String
    Public P2 As String
    Public P3 As String
    Public P4 As String
    Public P5 As String
    Public PM1 As String
    Public PM2 As String
    Public PM3 As String
    Public PM4 As String
    Public PM5 As String
    Public PModel1 As String
    Public PModel2 As String
    Public PModel3 As String
    Public PModel4 As String
    Public PModel5 As String
    Public PStyle1 As String
    Public PStyle2 As String
    Public PStyle3 As String
    Public PStyle4 As String
    Public PStyle5 As String
    Public PColor1 As String
    Public PColor2 As String
    Public PColor3 As String
    Public PColor4 As String
    Public PColor5 As String
    Public Qty1 As Integer = 0
    Public Qty2 As Integer = 0
    Public Qty3 As Integer = 0
    Public Qty4 As Integer = 0
    Public Qty5 As Integer = 0
    Public Unit1 As String
    Public Unit2 As String
    Public Unit3 As String
    Public Unit4 As String
    Public Unit5 As String
    Public Split1 As Double = 0
    Public Split2 As Double = 0
    Public Split3 As Double = 0
    Public Split4 As Double = 0
    Public Split5 As Double = 0
#End Region

    Private Sub SaleDetail()
        '' Populates and Manages all Sale Detail Variables 
        Dim cnt As Integer = Me.lstSold.Items.Count
        Select Case cnt
            Case Is = 1
                If Me.P1 <> Me.lstSold.Items(0) Then
                    Me.PM1 = ""
                    Me.PModel1 = ""
                    Me.PStyle1 = ""
                    Me.PColor1 = ""
                    Me.Qty1 = 0
                    Me.Unit1 = ""
                    Me.Split1 = 0
                    Me.lnkP1.LinkVisited = False
                    Me.txtP1Amnt.Text = ""
                End If
                Me.txtP1.Text = Me.lstSold.Items(0)
                Me.P1 = Me.lstSold.Items(0)
                If Me.Split1 = 0 Then
                    Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
                Else
                    Me.txtP1Amnt.Text = CStr(Me.Split1)
                End If
                Me.txtP1.Visible = True
                Me.txtP1Amnt.Visible = True
                Me.lnkP1.Visible = True
                If Me.PM1 <> "" Or Me.Unit1 <> "" Then
                    Me.lnkP1.LinkVisited = True
                End If
                Me.P2 = ""
                Me.PM2 = ""
                Me.PModel2 = ""
                Me.PStyle2 = ""
                Me.PColor2 = ""
                Me.Qty2 = 0
                Me.Unit2 = ""
                Me.Split2 = 0
                Me.txtP2Amnt.Text = 0
                Me.P3 = ""
                Me.PM3 = ""
                Me.PModel3 = ""
                Me.PStyle3 = ""
                Me.PColor3 = ""
                Me.Qty3 = 0
                Me.Unit3 = ""
                Me.Split3 = 0
                Me.txtP3Amnt.Text = 0
                Me.P4 = ""
                Me.PM4 = ""
                Me.PModel4 = ""
                Me.PStyle4 = ""
                Me.PColor4 = ""
                Me.Qty4 = 0
                Me.Unit4 = ""
                Me.Split4 = 0
                Me.txtP4Amnt.Text = 0
                Me.P5 = ""
                Me.PM5 = ""
                Me.PModel5 = ""
                Me.PStyle5 = ""
                Me.PColor5 = ""
                Me.Qty5 = 0
                Me.Unit5 = ""
                Me.Split5 = 0
                Me.txtP5Amnt.Text = 0

                Me.txtP2.Visible = False
                Me.txtP2Amnt.Visible = False
                Me.lnkP2.Visible = False
                Me.txtP3.Visible = False
                Me.txtP3Amnt.Visible = False
                Me.lnkP3.Visible = False
                Me.txtP4.Visible = False
                Me.txtP4Amnt.Visible = False
                Me.lnkP4.Visible = False
                Me.txtP5.Visible = False
                Me.txtP5Amnt.Visible = False
                Me.lnkP5.Visible = False

                Me.txtP2.Text = ""
                Me.txtP3.Text = ""
                Me.txtP4.Text = ""
                Me.txtP5.Text = ""
                Me.txtP2Amnt.Text = 0
                Me.txtP3Amnt.Text = 0
                Me.txtP4Amnt.Text = 0
                Me.txtP5Amnt.Text = 0
            Case Is = 2
                If Me.P1 <> Me.lstSold.Items(0) Then
                    Me.PM1 = ""
                    Me.PModel1 = ""
                    Me.PStyle1 = ""
                    Me.PColor1 = ""
                    Me.Qty1 = 0
                    Me.Unit1 = ""
                    Me.Split1 = 0
                    Me.lnkP1.LinkVisited = False
                    Me.txtP1Amnt.Text = ""
                End If
                Me.txtP1.Text = Me.lstSold.Items(0)
                Me.P1 = Me.lstSold.Items(0)
                If Me.Split1 = 0 Then
                    Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
                Else
                    Me.txtP1Amnt.Text = CStr(Me.Split1)
                End If
                Me.txtP1.Visible = True
                Me.txtP1Amnt.Visible = True
                Me.lnkP1.Visible = True
                If Me.PM1 <> "" Or Me.Unit1 <> "" Then
                    Me.lnkP1.LinkVisited = True
                End If
                If Me.P2 <> Me.lstSold.Items(1) Then
                    Me.PM2 = ""
                    Me.PModel2 = ""
                    Me.PStyle2 = ""
                    Me.PColor2 = ""
                    Me.Qty2 = 0
                    Me.Unit2 = ""
                    Me.Split2 = 0
                    Me.txtP2Amnt.Text = ""
                    Me.lnkP2.LinkVisited = False
                End If
                Me.txtP2.Text = Me.lstSold.Items(1)
                Me.P2 = Me.lstSold.Items(1)
                If Me.Split2 = 0 Then
                    Me.txtP2Amnt.Text = "0.00"
                Else
                    Me.txtP2Amnt.Text = CStr(Me.Split2)
                End If
                Me.txtP2.Visible = True
                Me.txtP2Amnt.Visible = True
                Me.lnkP2.Visible = True
                If Me.PM2 <> "" Or Me.Unit2 <> "" Then
                    Me.lnkP2.LinkVisited = True
                End If
                Me.P3 = ""
                Me.PM3 = ""
                Me.PModel3 = ""
                Me.PStyle3 = ""
                Me.PColor3 = ""
                Me.Qty3 = 0
                Me.Unit3 = ""
                Me.Split3 = 0
                Me.txtP3Amnt.Text = 0
                Me.P4 = ""
                Me.PM4 = ""
                Me.PModel4 = ""
                Me.PStyle4 = ""
                Me.PColor4 = ""
                Me.Qty4 = 0
                Me.Unit4 = ""
                Me.Split4 = 0
                Me.txtP4Amnt.Text = 0
                Me.P5 = ""
                Me.PM5 = ""
                Me.PModel5 = ""
                Me.PStyle5 = ""
                Me.PColor5 = ""
                Me.Qty5 = 0
                Me.Unit5 = ""
                Me.Split5 = 0
                Me.txtP5Amnt.Text = 0

                Me.txtP3.Visible = False
                Me.txtP3Amnt.Visible = False
                Me.lnkP3.Visible = False
                Me.txtP4.Visible = False
                Me.txtP4Amnt.Visible = False
                Me.lnkP4.Visible = False
                Me.txtP5.Visible = False
                Me.txtP5Amnt.Visible = False
                Me.lnkP5.Visible = False
                Me.txtP3.Text = ""
                Me.txtP4.Text = ""
                Me.txtP5.Text = ""
                Me.txtP3Amnt.Text = 0
                Me.txtP4Amnt.Text = 0
                Me.txtP5Amnt.Text = 0
            Case Is = 3
                If Me.P1 <> Me.lstSold.Items(0) Then
                    Me.PM1 = ""
                    Me.PModel1 = ""
                    Me.PStyle1 = ""
                    Me.PColor1 = ""
                    Me.Qty1 = 0
                    Me.Unit1 = ""
                    Me.Split1 = 0
                    Me.lnkP1.LinkVisited = False
                    Me.txtP1Amnt.Text = ""
                End If
                Me.txtP1.Text = Me.lstSold.Items(0)
                Me.P1 = Me.lstSold.Items(0)
                If Me.Split1 = 0 Then
                    Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
                Else
                    Me.txtP1Amnt.Text = CStr(Me.Split1)
                End If
                Me.txtP1.Visible = True
                Me.txtP1Amnt.Visible = True
                Me.lnkP1.Visible = True
                If Me.PM1 <> "" Or Me.Unit1 <> "" Then
                    Me.lnkP1.LinkVisited = True
                End If
                If Me.P2 <> Me.lstSold.Items(1) Then
                    Me.PM2 = ""
                    Me.PModel2 = ""
                    Me.PStyle2 = ""
                    Me.PColor2 = ""
                    Me.Qty2 = 0
                    Me.Unit2 = ""
                    Me.Split2 = 0
                    Me.lnkP2.LinkVisited = False
                    Me.txtP2Amnt.Text = ""
                End If
                Me.txtP2.Text = Me.lstSold.Items(1)
                Me.P2 = Me.lstSold.Items(1)
                If Me.Split2 = 0 Then
                    Me.txtP2Amnt.Text = "0.00"
                Else
                    Me.txtP2Amnt.Text = CStr(Me.Split2)
                End If
                Me.txtP2.Visible = True
                Me.txtP2Amnt.Visible = True
                Me.lnkP2.Visible = True
                If Me.PM2 <> "" Or Me.Unit2 <> "" Then
                    Me.lnkP2.LinkVisited = True
                End If
                If Me.P3 <> Me.lstSold.Items(2) Then
                    Me.PM3 = ""
                    Me.PModel3 = ""
                    Me.PStyle3 = ""
                    Me.PColor3 = ""
                    Me.Qty3 = 0
                    Me.Unit3 = ""
                    Me.Split3 = 0
                    Me.lnkP3.LinkVisited = False
                    Me.txtP3Amnt.Text = ""
                End If
                Me.txtP3.Text = Me.lstSold.Items(2)
                Me.P3 = Me.lstSold.Items(2)
                If Me.Split3 = 0 Then
                    Me.txtP3Amnt.Text = "0.00"
                Else
                    Me.txtP3Amnt.Text = CStr(Me.Split3)
                End If
                Me.txtP3.Visible = True
                Me.txtP3Amnt.Visible = True
                Me.lnkP3.Visible = True
                If Me.PM3 <> "" Or Me.Unit3 <> "" Then
                    Me.lnkP3.LinkVisited = True
                End If
                Me.P4 = ""
                Me.PM4 = ""
                Me.PModel4 = ""
                Me.PStyle4 = ""
                Me.PColor4 = ""
                Me.Qty4 = 0
                Me.Unit4 = ""
                Me.Split4 = 0
                Me.txtP4Amnt.Text = 0
                Me.P5 = ""
                Me.PM5 = ""
                Me.PModel5 = ""
                Me.PStyle5 = ""
                Me.PColor5 = ""
                Me.Qty5 = 0
                Me.Unit5 = ""
                Me.Split5 = 0
                Me.txtP5Amnt.Text = 0

                Me.txtP4.Visible = False
                Me.txtP4Amnt.Visible = False
                Me.lnkP4.Visible = False
                Me.txtP5.Visible = False
                Me.txtP5Amnt.Visible = False
                Me.lnkP5.Visible = False
                Me.txtP4.Text = ""
                Me.txtP5.Text = ""
                Me.txtP4Amnt.Text = 0
                Me.txtP5Amnt.Text = 0
            Case Is = 4
                If Me.P1 <> Me.lstSold.Items(0) Then
                    Me.PM1 = ""
                    Me.PModel1 = ""
                    Me.PStyle1 = ""
                    Me.PColor1 = ""
                    Me.Qty1 = 0
                    Me.Unit1 = ""
                    Me.Split1 = 0
                    Me.lnkP1.LinkVisited = False
                    Me.txtP1Amnt.Text = ""
                End If
                Me.txtP1.Text = Me.lstSold.Items(0)
                Me.P1 = Me.lstSold.Items(0)
                If Me.Split1 = 0 Then
                    Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
                Else
                    Me.txtP1Amnt.Text = CStr(Me.Split1)
                End If
                Me.txtP1.Visible = True
                Me.txtP1Amnt.Visible = True
                Me.lnkP1.Visible = True
                If Me.PM1 <> "" Or Me.Unit1 <> "" Then
                    Me.lnkP1.LinkVisited = True
                End If
                If Me.P2 <> Me.lstSold.Items(1) Then
                    Me.PM2 = ""
                    Me.PModel2 = ""
                    Me.PStyle2 = ""
                    Me.PColor2 = ""
                    Me.Qty2 = 0
                    Me.Unit2 = ""
                    Me.Split2 = 0
                    Me.txtP2Amnt.Text = ""
                    Me.lnkP2.LinkVisited = False
                End If
                Me.txtP2.Text = Me.lstSold.Items(1)
                Me.P2 = Me.lstSold.Items(1)
                If Me.Split2 = 0 Then
                    Me.txtP2Amnt.Text = "0.00"
                Else
                    Me.txtP2Amnt.Text = CStr(Me.Split2)
                End If
                Me.txtP2.Visible = True
                Me.txtP2Amnt.Visible = True
                Me.lnkP2.Visible = True
                If Me.PM2 <> "" Or Me.Unit2 <> "" Then
                    Me.lnkP2.LinkVisited = True
                End If
                If Me.P3 <> Me.lstSold.Items(2) Then
                    Me.PM3 = ""
                    Me.PModel3 = ""
                    Me.PStyle3 = ""
                    Me.PColor3 = ""
                    Me.Qty3 = 0
                    Me.Unit3 = ""
                    Me.Split3 = 0
                    Me.lnkP3.LinkVisited = False
                    Me.txtP3Amnt.Text = ""
                End If
                Me.txtP3.Text = Me.lstSold.Items(2)
                Me.P3 = Me.lstSold.Items(2)
                If Me.Split3 = 0 Then
                    Me.txtP3Amnt.Text = "0.00"
                Else
                    Me.txtP3Amnt.Text = CStr(Me.Split3)
                End If
                Me.txtP3.Visible = True
                Me.txtP3Amnt.Visible = True
                Me.lnkP3.Visible = True
                If Me.PM3 <> "" Or Me.Unit3 <> "" Then
                    Me.lnkP3.LinkVisited = True
                End If
                If Me.P4 <> Me.lstSold.Items(3) Then
                    Me.PM4 = ""
                    Me.PModel4 = ""
                    Me.PStyle4 = ""
                    Me.PColor4 = ""
                    Me.Qty4 = 0
                    Me.Unit4 = ""
                    Me.Split4 = 0
                    Me.lnkP4.LinkVisited = False
                    Me.txtP4Amnt.Text = ""
                End If
                Me.txtP4.Text = Me.lstSold.Items(3)
                Me.P4 = Me.lstSold.Items(3)
                If Me.Split4 = 0 Then
                    Me.txtP4Amnt.Text = "0.00"
                Else
                    Me.txtP4Amnt.Text = CStr(Me.Split4)
                End If
                Me.txtP4.Visible = True
                Me.txtP4Amnt.Visible = True
                Me.lnkP4.Visible = True
                If Me.PM4 <> "" Or Me.Unit4 <> "" Then
                    Me.lnkP4.LinkVisited = True
                End If
                Me.P5 = ""
                Me.PM5 = ""
                Me.PModel5 = ""
                Me.PStyle5 = ""
                Me.PColor5 = ""
                Me.Qty5 = 0
                Me.Unit5 = ""
                Me.Split5 = 0
                Me.txtP5Amnt.Text = 0
                Me.txtP5.Text = ""


                Me.txtP5.Visible = False
                Me.txtP5Amnt.Visible = False
                Me.lnkP5.Visible = False
            Case Is = 5
                If Me.P1 <> Me.lstSold.Items(0) Then
                    Me.PM1 = ""
                    Me.PModel1 = ""
                    Me.PStyle1 = ""
                    Me.PColor1 = ""
                    Me.Qty1 = 0
                    Me.Unit1 = ""
                    Me.Split1 = 0
                    Me.lnkP1.LinkVisited = False
                    Me.txtP1Amnt.Text = ""
                End If
                Me.txtP1.Text = Me.lstSold.Items(0)
                Me.P1 = Me.lstSold.Items(0)
                If Me.Split1 = 0 Then
                    Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
                Else
                    Me.txtP1Amnt.Text = CStr(Me.Split1)
                End If
                Me.txtP1.Visible = True
                Me.txtP1Amnt.Visible = True
                Me.lnkP1.Visible = True
                If Me.PM1 <> "" Or Me.Unit1 <> "" Then
                    Me.lnkP1.LinkVisited = True
                End If
                If Me.P2 <> Me.lstSold.Items(1) Then
                    Me.PM2 = ""
                    Me.PModel2 = ""
                    Me.PStyle2 = ""
                    Me.PColor2 = ""
                    Me.Qty2 = 0
                    Me.Unit2 = ""
                    Me.Split2 = 0
                    Me.lnkP2.LinkVisited = False
                    Me.txtP2Amnt.Text = ""
                End If
                Me.txtP2.Text = Me.lstSold.Items(1)
                Me.P2 = Me.lstSold.Items(1)
                If Me.Split2 = 0 Then
                    Me.txtP2Amnt.Text = "0.00"
                Else
                    Me.txtP2Amnt.Text = CStr(Me.Split2)
                End If
                Me.txtP2.Visible = True
                Me.txtP2Amnt.Visible = True
                Me.lnkP2.Visible = True
                If Me.PM2 <> "" Or Me.Unit2 <> "" Then
                    Me.lnkP2.LinkVisited = True
                End If
                If Me.P3 <> Me.lstSold.Items(2) Then
                    Me.PM3 = ""
                    Me.PModel3 = ""
                    Me.PStyle3 = ""
                    Me.PColor3 = ""
                    Me.Qty3 = 0
                    Me.Unit3 = ""
                    Me.Split3 = 0
                    Me.lnkP3.LinkVisited = False
                    Me.txtP3Amnt.Text = ""
                End If
                Me.txtP3.Text = Me.lstSold.Items(2)
                Me.P3 = Me.lstSold.Items(2)
                If Me.Split3 = 0 Then
                    Me.txtP3Amnt.Text = "0.00"
                Else
                    Me.txtP3Amnt.Text = CStr(Me.Split3)
                End If
                Me.txtP3.Visible = True
                Me.txtP3Amnt.Visible = True
                Me.lnkP3.Visible = True
                If Me.PM3 <> "" Or Me.Unit3 <> "" Then
                    Me.lnkP3.LinkVisited = True
                End If
                If Me.P4 <> Me.lstSold.Items(3) Then
                    Me.PM4 = ""
                    Me.PModel4 = ""
                    Me.PStyle4 = ""
                    Me.PColor4 = ""
                    Me.Qty4 = 0
                    Me.Unit4 = ""
                    Me.Split4 = 0
                    Me.lnkP4.LinkVisited = False
                    Me.txtP4Amnt.Text = ""
                End If
                Me.txtP4.Text = Me.lstSold.Items(3)
                Me.P4 = Me.lstSold.Items(3)
                If Me.Split4 = 0 Then
                    Me.txtP4Amnt.Text = "0.00"
                Else
                    Me.txtP4Amnt.Text = CStr(Me.Split4)
                End If
                Me.txtP4.Visible = True
                Me.txtP4Amnt.Visible = True
                Me.lnkP4.Visible = True
                If Me.PM4 <> "" Or Me.Unit4 <> "" Then
                    Me.lnkP4.LinkVisited = True
                End If
                If Me.P5 <> Me.lstSold.Items(4) Then
                    Me.PM5 = ""
                    Me.PModel5 = ""
                    Me.PStyle5 = ""
                    Me.PColor5 = ""
                    Me.Qty5 = 0
                    Me.Unit5 = ""
                    Me.Split5 = 0
                    Me.lnkP5.LinkVisited = False
                    Me.txtP5Amnt.Text = ""
                End If
                Me.txtP5.Text = Me.lstSold.Items(4)
                Me.P5 = Me.lstSold.Items(4)
                If Me.Split5 = 0 Then
                    Me.txtP5Amnt.Text = "0.00"
                Else
                    Me.txtP5Amnt.Text = CStr(Me.Split5)
                End If
                Me.txtP5Amnt.Text = "0.00"
                Me.txtP5.Visible = True
                Me.txtP5Amnt.Visible = True
                Me.lnkP5.Visible = True
                If Me.PM5 <> "" Or Me.Unit5 <> "" Then
                    Me.lnkP5.LinkVisited = True
                End If
        End Select
    End Sub

    Private Sub Form_Reset()
        Me.SeqNum = 0
        Me.StepCount = 1
        Me.cboRep1.Text = Nothing
        Me.cboRep2.Text = Nothing
        Me.cboSalesResults.Text = Nothing
        Me.chkRecoverable.Checked = False
        Me.rtfNote.Text = ""
        Me.txt1ContractAmt.Text = ""
        Me.txt2ContractAmt.Text = "00"
        Me.rdoFinance.Checked = False
        Me.rdoCash.Checked = False
        Me.txt1Quoted.Text = ""
        Me.txt2Quoted.Text = "00"
        Me.txt1Par.Text = ""
        Me.txt2Par.Text = "00"
        Me.lstManage.Items.Clear()
        Me.lstSold.Items.Clear()
        Me.txtP1.Text = ""
        Me.txtP2.Text = ""
        Me.txtP3.Text = ""
        Me.txtP4.Text = ""
        Me.txtP5.Text = ""
        Me.txtP1Amnt.Text = ""
        Me.txtP2Amnt.Text = ""
        Me.txtP3Amnt.Text = ""
        Me.txtP4Amnt.Text = ""
        Me.txtP5Amnt.Text = ""
        Me.lnkP1.LinkVisited = False
        Me.lnkP2.LinkVisited = False
        Me.lnkP3.LinkVisited = False
        Me.lnkP4.LinkVisited = False
        Me.lnkP5.LinkVisited = False
        Me.rtfNote.Text = ""
        Me.rtfNote_LostFocus(Nothing, Nothing)
        Me.txtContractNotes.Text = ""
        Me.txtContractNotes_LostFocus(Nothing, Nothing)
        Me.P1 = ""
        Me.P2 = ""
        Me.P3 = ""
        Me.P4 = ""
        Me.P5 = ""
        Me.PM1 = ""
        Me.PM2 = ""
        Me.PM3 = ""
        Me.PM4 = ""
        Me.PM5 = ""
        Me.PModel1 = ""
        Me.PModel2 = ""
        Me.PModel3 = ""
        Me.PModel4 = ""
        Me.PModel5 = ""
        Me.PStyle1 = ""
        Me.PStyle2 = ""
        Me.PStyle3 = ""
        Me.PStyle4 = ""
        Me.PStyle5 = ""
        Me.PColor1 = ""
        Me.PColor2 = ""
        Me.PColor3 = ""
        Me.PColor4 = ""
        Me.PColor5 = ""
        Me.Qty1 = 0
        Me.Qty2 = 0
        Me.Qty3 = 0
        Me.Qty4 = 0
        Me.Qty5 = 0
        Me.Unit1 = ""
        Me.Unit2 = ""
        Me.Unit3 = ""
        Me.Unit4 = ""
        Me.Unit5 = ""
        Me.Split1 = 0
        Me.Split2 = 0
        Me.Split3 = 0
        Me.Split4 = 0
        Me.Split5 = 0
    End Sub

    Private Sub SDResult_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Me.btnEdit.Text = "Exit Edit Mode" Then
            Me.btnEdit_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '' Form Reset for quick recycle, dispose method not gettin it done quick enough 
        Me.Form_Reset()
        Me.GetData()
        Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID
        Me.pnlFirst.Visible = True
        Me.pnlDemo.Visible = False
        Me.pnlLast.Visible = False
        Me.pnlSale1.Visible = False
        Me.pnlSale2.Visible = False
        Me.pnlSale4.Visible = False
        Me.pnlSale5.Visible = False
        Me.btnEdit.Visible = True
        Me.btnNext.Visible = True
        Me.btnSave.Visible = False
        Me.lnkCalc.Visible = False
        Me.lvPrevious.Enabled = False
        Me.cboRep1.Focus()

        Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdget As SqlCommand = New SqlCommand("dbo.GetSMAutoNotes", cnn)
        cmdget.CommandType = CommandType.StoredProcedure
        Dim r1 As SqlDataReader
        Me.cboautonotes.Items.Clear()
        Me.cboautonotes.Items.Add("<Add New>")
        Me.cboautonotes.Items.Add("________________________________________________________________")

        cnn.Open()
        r1 = cmdget.ExecuteReader
        While r1.Read
            Me.cboautonotes.Items.Add(r1.Item(0))
        End While
        r1.Close()
        cnn.Close()
    End Sub

    Private Sub Sequence()
        Select Case Me.SeqNum
            Case Is = 1
                Select Case StepCount
                    Case Is = 1
                        Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = True
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnSave.Visible = False
                        Me.btnNext.Visible = True
                        Me.btnBack.Enabled = False
                        Me.btnEdit.Visible = True
                        Me.cboRep1.Focus()
                    Case Is = 2
                        Me.Text = "Enter notes to log with this result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = True
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnNext.Visible = False
                        Me.btnSave.Visible = True
                        Me.btnBack.Enabled = True
                        Me.btnEdit.Visible = False
                        Me.rtfNote.Focus()
                End Select
            Case Is = 2
                Select Case StepCount
                    Case Is = 1
                        Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = True
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnSave.Visible = False
                        Me.btnNext.Visible = True
                        Me.btnBack.Enabled = False
                        Me.chkRecoverable.Checked = False
                        Me.btnEdit.Visible = True
                        Me.cboRep1.Focus()
                    Case Is = 2
                        Dim m = MsgBox("Would you like to send this ""Recission Cancel""" & vbCr & "to the Recovery Department?", MsgBoxStyle.YesNo, "Recoverable?...")
                        If m = vbYes Then
                            Me.chkRecoverable.Checked = True
                        Else
                            Me.chkRecoverable.Checked = False
                        End If
                        Me.Text = "Enter notes to log with this result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = True
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnNext.Visible = False
                        Me.btnSave.Visible = True
                        Me.btnBack.Enabled = True
                        Me.btnEdit.Visible = False
                        Me.rtfNote.Focus()
                End Select
            Case Is = 3
                Select Case StepCount
                    Case Is = 1
                        Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = True
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnSave.Visible = False
                        Me.btnNext.Visible = True
                        Me.btnBack.Enabled = False
                        Me.btnEdit.Visible = True
                        Me.cboRep1.Focus()

                    Case Is = 2
                        Me.Text = "Enter Demo/No Sale Details" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = True
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnBack.Enabled = True
                        Me.btnNext.Visible = True
                        Me.btnSave.Visible = False
                        Me.btnEdit.Visible = False
                        Me.txt1Quoted.Focus()

                    Case Is = 3
                        Me.Text = "Enter notes to log with this result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = True
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnNext.Visible = False

                        Me.btnSave.Visible = True
                        Me.btnBack.Enabled = True
                        Me.rtfNote.Focus()

                End Select
            Case Is = 4
                Select Case StepCount
                    Case Is = 1
                        Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = True
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnSave.Visible = False
                        Me.btnNext.Visible = True
                        Me.btnBack.Enabled = False
                        Me.btnEdit.Visible = True
                        Me.cboRep1.Focus()

                    Case Is = 2
                        Me.Text = "Enter Contract Details" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = True
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnBack.Enabled = True
                        Me.btnEdit.Visible = False
                        Me.txt1ContractAmt.Focus()
                    Case Is = 3
                        Me.Text = "Enter Contract Notes" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = True

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.txtContractNotes.Focus()
                    Case Is = 4
                        Me.Text = "Manage Products" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = True
                        Me.pnlSale5.Visible = False
                        Me.lnkCalc.Visible = False
                        Me.lstManage.Focus()
                    Case Is = 5
                        Me.SaleDetail()
                        Me.Text = "Enter Contract Splits and Product Details" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = False
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False
                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = True
                        Me.btnNext.Visible = True
                        Me.btnSave.Visible = False
                        Me.txtP1Amnt.Focus()
                        Me.lnkCalc.Visible = True
                        Me.txtP1Amnt.Focus()


                    Case Is = 6
                        Me.Text = "Enter notes to log with this result" & "- " & Me.ID & " " & Me.Editmodetext
                        Me.pnlFirst.Visible = False
                        Me.pnlDemo.Visible = False
                        Me.pnlLast.Visible = True
                        Me.pnlSale1.Visible = False
                        Me.pnlSale2.Visible = False

                        Me.pnlSale4.Visible = False
                        Me.pnlSale5.Visible = False
                        Me.btnNext.Visible = False
                        Me.btnSave.Visible = True
                        Me.btnBack.Enabled = True
                        Me.lnkCalc.Visible = False
                        Me.rtfNote.Focus()
                End Select
        End Select


    End Sub




    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        If Me.pnlFirst.Visible Then
            If Me.cboSalesResults.Text = "" Then
                Me.ErrorProvider1.SetError(Me.cboSalesResults, "You Must Supply" & vbCr & "a Sales Result!")
                Exit Sub
            End If
            If Me.cboRep1.Text = Me.cboRep2.Text And Me.cboRep1.Text <> Nothing Then
                MsgBox("Rep 1 and Rep 2 cannot be the same person!", MsgBoxStyle.Exclamation, "Must Supply 2 Unique Rep Names")
                Exit Sub
            End If
        ElseIf Me.pnlSale1.Visible Then
            If Me.txt1ContractAmt.Text = "" Then
                Me.ErrorProvider1.SetError(Me.txt1ContractAmt, "You Must Supply" & vbCr & "a Contract Amount!")
                Me.ErrorProvider1.SetIconPadding(Me.txt1ContractAmt, 30)
                Exit Sub

            End If
            If Me.rdoCash.Checked = False And Me.rdoFinance.Checked = False Then
                Me.ErrorProvider1.SetError(Me.grpPaymentOptions, "You Must Select" & vbCr & "a Payment Option")
                Exit Sub
            End If
        ElseIf Me.pnlSale4.Visible Then
            If Me.lstSold.Items.Count = 0 Then
                Me.ErrorProvider1.SetError(Me.lstSold, "You must move at least" & vbCr & "one Product to Sold!")
                Me.ErrorProvider1.SetIconPadding(Me.lstSold, -20)
                Exit Sub
            End If
            If Me.lstManage.Items.Count > 3 Then
                Dim tt As New ToolTip
                tt.IsBalloon = True
                tt.ToolTipTitle = "Future Interest"
                tt.ToolTipIcon = ToolTipIcon.Warning
                tt.Show("Future interest only accepts 3 products." & vbCr & _
                        "You can add more than 3 products for the" & vbCr & _
                        "purpose of adding products to Sold Products," & vbCr & _
                        "but you cannot proceed if Future Interest" & vbCr & _
                        "exceeds 3 products.", Me.lstManage, 110, -132)
                Exit Sub
            End If
        ElseIf Me.pnlSale5.Visible Then
            If Me.Split1 + Me.Split2 + Me.Split3 + Me.Split4 + Me.Split5 <> CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) Then
                MsgBox("Product $ Splits do not equal Contract Total!", MsgBoxStyle.Exclamation, "Transaction Not Balanced")
                Exit Sub
            End If
        ElseIf Me.pnlDemo.Visible Then
            If Me.txt1Quoted.Text = "" Then
                Me.ErrorProvider1.SetError(Me.txt2Quoted, "You must supply a Quoted Price")
                Me.ErrorProvider1.SetIconPadding(Me.txt2Quoted, 5)
                Exit Sub
            End If
            If Me.txt1Par.Text = "" Then
                Me.ErrorProvider1.SetError(Me.txt2Par, "You must supply a Par Price")
                Me.ErrorProvider1.SetIconPadding(Me.txt2Par, 5)
                Exit Sub
            End If
        End If
        If Me.EditMode = True Then
            If Me.lvPrevious.SelectedItems.Count = 0 Then
                MsgBox("You must select a Sales Result to edit or exit edit mode!", MsgBoxStyle.Exclamation, "Edit Mode")
                Exit Sub
            Else
                Me.lvPrevious.Enabled = False
                Me.lvPrevious.SelectedItems(0).ImageIndex = 4
                Me.lvPrevious.SelectedItems(0).ToolTipText = "Editing Result"

            End If
        End If
        Me.StepCount = Me.StepCount + 1
        Me.Sequence()
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        If Me.StepCount = 2 And Me.EditMode = True And Me.lvPrevious.SelectedItems.Count <> 0 Then
            Me.lvPrevious.Enabled = True
            If Me.lvPrevious.SelectedItems(0).Tag = "Locked" Then
                Me.lvPrevious.SelectedItems(0).ImageIndex = 3
                Me.lvPrevious.SelectedItems(0).ToolTipText = ""
            Else
                Me.lvPrevious.SelectedItems(0).ImageKey = "(none)"
                Me.lvPrevious.SelectedItems(0).ToolTipText = ""
            End If
        End If
        Me.StepCount = Me.StepCount - 1
        Me.Sequence()
    End Sub



    Private Sub cboSalesResults_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSalesResults.SelectedIndexChanged
        If Me.cboSalesResults.Text = "" Then
            Exit Sub
        Else
            Me.ErrorProvider1.Clear()
        End If
        If Me.NR = False And Me.EditMode = False And (Me.cboSalesResults.Text = "Reset" Or Me.cboSalesResults.Text = "Not Hit" _
        Or Me.cboSalesResults.Text = "Not Issued" Or Me.cboSalesResults.Text = "Lost Result" Or _
         Me.cboSalesResults.Text = "Demo/No Sale" Or Me.cboSalesResults.Text = "Sale" Or Me.cboSalesResults.Text = "No Demo") Then
            Dim x = MsgBox("The System is not requiring a sale result for this record. You can" & vbCr & _
                    "still enter this result, but it is not recommended due to the fact" & vbCr & _
                   "that it can affect reporting accuracy. Alternatively, You can Set an" & vbCr & _
                    "Appointment and Quick Confirm, then the result will be expected by" & vbCr & _
                    "The System, marketing information will be updated, and reporting will" & vbCr & _
                    "remain accurate or you may want to edit a previous result where" & vbCr & _
                    "appropriate, i.e. A Demo/No Sale turns into a Sale" & vbCr & vbCr & _
                    "Would you like to ignore this Warning and enter this result anyway?" & vbCr & "(Recommended)", MsgBoxStyle.YesNo, "Result Not Required")
            If x = vbNo Then
                Me.cboSalesResults.Text = Nothing
                Exit Sub
            End If
        End If

        If Me.cboRep1.Text = "" And (Me.cboSalesResults.Text = "No Demo" _
               Or Me.cboSalesResults.Text = "Reset" Or Me.cboSalesResults.Text = "Not Hit" _
               Or Me.cboSalesResults.Text = "Sale" Or Me.cboSalesResults.Text = "Recission Cancel" _
                Or Me.cboSalesResults.Text = "Bank Rejected" Or Me.cboSalesResults.Text = _
                "Bank Approved" Or Me.cboSalesResults.Text = "Demo/No Sale") Then
            MsgBox("This record must be issued to a rep to log the result """ & Me.cboSalesResults.Text & _
                    """." & vbCr & "If this Record wasn't issued you can use the result ""Not Issued"" or ""Lost Results""", MsgBoxStyle.Exclamation, "Sales Result")
            Me.cboSalesResults.Text = Nothing
            Exit Sub
        ElseIf Me.cboSalesResults.Text = "Not Issued" And (Me.cboRep1.Text <> "") Then
            MsgBox("This record has been issued to " & Me.cboRep1.Text & ", therefore cannot be logged as ""Not Issued""." & vbCr & "Please select a different result or remove " & Me.cboRep1.Text & " from Sales Rep 1.", MsgBoxStyle.Exclamation, "Sales Result")
            Me.cboSalesResults.Text = Nothing
            Exit Sub

        ElseIf Me.lvPrevious.Items.Count <> 0 Then
            If Me.cboSalesResults.Text = "Recission Cancel" Then
                If Me.cboSalesResults.Text = "Recission Cancel" And (Me.lvPrevious.TopItem.SubItems(2).Text = "Sale" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Rejected" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Approved" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Recission Cancel") Then
                    If Me.cboSalesResults.Text = "Recission Cancel" And Me.lvPrevious.TopItem.SubItems(2).Text = "Recission Cancel" And Me.EditMode = False Then
                        MsgBox("Someone has already logged a Recission Cancel" & vbCr & _
                                "for this record. If you need to change the" & vbCr & _
                                "details for this result, you can open up" & vbCr & _
                                "previous results and edit the details there.", MsgBoxStyle.Exclamation, "Result Already Logged")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                Else
                    If Me.EditMode = False Then
                        Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Recission Cancel""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                End If
            End If
            Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet As SqlCommand
            Dim r As SqlDataReader

            cmdGet = New SqlCommand("Select cash from Saledetail where leadhistoryid = " & Me.lvPrevious.TopItem.SubItems(9).Text, cnn)
            cmdGet.CommandType = CommandType.Text
            cnn.Open()
            r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)

            While r.Read
                If r.Item(0) = True And Me.cboSalesResults.Text = "Bank Rejected" Then
                    MsgBox("The System is reporting that the last sale for this record" & vbCr & _
                            "was a cash sale, therefore cannot be Bank Rejected. If this" & vbCr & _
                            "is inaccurate, then you will need to edit the sale result to" & vbCr & _
                            "reflect that it was a finance sale, and enter the new result" & vbCr & _
                            """Bank Rejected"".", MsgBoxStyle.Exclamation, "Cash Sale")
                    Me.cboSalesResults.Text = Nothing
                    Exit Sub
                ElseIf r.Item(0) = True And Me.cboSalesResults.Text = "Bank Approved" Then
                    MsgBox("The System is reporting that the last sale for this record" & vbCr & _
                            "was a cash sale, therefore cannot be Bank Rejected. If this" & vbCr & _
                            "is inaccurate, then you will need to edit the sale result to" & vbCr & _
                            "reflect that it was a finance sale, and enter the new result" & vbCr & _
                            """Bank Approved"".", MsgBoxStyle.Exclamation, "Cash Sale")
                    Me.cboSalesResults.Text = Nothing
                    Exit Sub
                End If
            End While
            r.Close()
            cnn.Close()



            If Me.cboSalesResults.Text = "Bank Rejected" And Me.EditMode = False Then
                If Me.cboSalesResults.Text = "Bank Rejected" And (Me.lvPrevious.TopItem.SubItems(2).Text = "Sale" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Recission Cancel" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Approved" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Rejected") Then
                    If Me.cboSalesResults.Text = "Bank Rejected" And Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Rejected" Then
                        MsgBox("Someone has already logged a Bank Rejection" & vbCr & _
                                "for this record. If you need to change the" & vbCr & _
                                "details for this result, you can open up" & vbCr & _
                                "previous results and edit the details there.", MsgBoxStyle.Exclamation, "Result Already Logged")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    ElseIf Me.cboSalesResults.Text = "Bank Rejected" And Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Approved" Then
                        MsgBox("The Sytem is reporting a Bank Approval for this" & vbCr & _
                                "record. Each Sale can only have 1 Bank Approval" & vbCr & _
                                "or 1 Bank Rejection. If you need to log a Bank" & vbCr & _
                                "Rejection, you must open previous results and" & vbCr & _
                                "select the Bank Approval and edit it to reflect" & vbCr & _
                                "the new financing status.", MsgBoxStyle.Exclamation, "Cannot Log Bank Rejected")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                Else
                    If Me.EditMode = False Then
                        Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Bank Rejected""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                End If
            End If
            If Me.cboSalesResults.Text = "Bank Approved" And Me.EditMode = False Then
                If Me.cboSalesResults.Text = "Bank Approved" And (Me.lvPrevious.TopItem.SubItems(2).Text = "Sale" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Recission Cancel" Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Rejected") Or Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Approved" Then
                    If Me.cboSalesResults.Text = "Bank Approved" And Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Approved" Then
                        MsgBox("Someone has already logged a Bank Approval" & vbCr & _
                                "for this record. If you need to change the" & vbCr & _
                                "details for this result, you can open up" & vbCr & _
                                "previous results and edit the details there.", MsgBoxStyle.Exclamation, "Result Already Logged")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    ElseIf Me.cboSalesResults.Text = "Bank Approved" And Me.lvPrevious.TopItem.SubItems(2).Text = "Bank Rejected" Then
                        MsgBox("The Sytem is reporting a Bank Rejection for this" & vbCr & _
                                "record. Each Sale can only have 1 Bank Approval" & vbCr & _
                                "or 1 Bank Rejection. If you need to log a Bank" & vbCr & _
                                "Rejection, you must open previous results and" & vbCr & _
                                "select the Bank Rejection and edit it to reflect" & vbCr & _
                                "the new financing status.", MsgBoxStyle.Exclamation, "Cannot Log Bank Approved")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                Else
                    If Me.EditMode = False Then
                        Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Bank Approved""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                        Me.cboSalesResults.Text = Nothing
                        Exit Sub
                    End If
                End If
            End If
        Else
            If Me.cboSalesResults.Text = "Recission Cancel" Then
                Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Recission Cancel""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                Me.cboSalesResults.Text = Nothing
                Exit Sub
            End If
            If Me.cboSalesResults.Text = "Bank Rejected" Then
                Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Bank Rejected""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                Me.cboSalesResults.Text = Nothing
                Exit Sub
            End If

            If Me.cboSalesResults.Text = "Bank Approved" Then
                Dim x = MsgBox("The system is not showing a Sale for this record." & vbCr & "You must first log a sale for this record, then log" & vbCr & "the result ""Bank Approved""!", MsgBoxStyle.Exclamation, "No Sale Reported")
                Me.cboSalesResults.Text = Nothing
                Exit Sub
            End If
        End If
        If Me.pnlFirst.Visible = True And Me.cboSalesResults.Text = "No Demo" _
                   Or Me.cboSalesResults.Text = "Reset" Or Me.cboSalesResults.Text = "Not Hit" _
                   Or Me.cboSalesResults.Text = "Not Issued" Or Me.cboSalesResults.Text = "Lost Result" _
                    Or Me.cboSalesResults.Text = "Bank Rejected" Or Me.cboSalesResults.Text = "Bank Approved" Then
            Me.SeqNum = 1
        ElseIf Me.pnlFirst.Visible = True And Me.cboSalesResults.Text = "Recission Cancel" Then
            Me.SeqNum = 2
        ElseIf Me.pnlFirst.Visible = True And Me.cboSalesResults.Text = "Demo/No Sale" Then
            Me.SeqNum = 3
            Me.chkRecoverable.Checked = False
        ElseIf Me.pnlFirst.Visible = True And Me.cboSalesResults.Text = "Sale" Then
            Me.SeqNum = 4
            Me.chkRecoverable.Checked = False
            Me.lstManage.Items.Clear()
            If Me.Product1 <> Nothing Then
                Me.lstManage.Items.Add(Me.Product1)
            End If
            If Me.Product2 <> "" Then
                Me.lstManage.Items.Add(Me.Product2)
            End If
            If Me.Product3 <> "" Then
                Me.lstManage.Items.Add(Me.Product3)
            End If
            If Me.lstManage.Items.Count <> 0 Then
                Me.lstManage.SetSelected(0, True)
            End If
            If Me.lstManage.Items.Count = 1 And Me.EditMode = False Then
                Me.btnMovetoSold_Click(Nothing, Nothing)
            End If
            Me.rdoFinance.Checked = False
            Me.rdoCash.Checked = False
            Me.ProductList()
        End If
        If Me.EditMode = True Then
            If (Me.cboSalesResults.Text = "Demo/No Sale" Or Me.cboSalesResults.Text = "Recission Cancel") Then
            Else
                Me.chkRecoverable.Checked = False
            End If
        End If

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        If Me.NR = True Then
            Dim x As New ToolTip
            x.IsBalloon = True
            x.ToolTipTitle = "System Required Result!"
            x.ToolTipIcon = ToolTipIcon.Warning
            x.Show("The System is reporting that this" & vbCr & "record has an outstanding Sales" & vbCr & "Result! If you need to edit a past" & vbCr & "result, first enter the system" & vbCr & "required result and save. Then" & vbCr & "select the record again for editing.", Me.btnEdit, New System.Drawing.Point(40, -125))
            Exit Sub
        End If
        If btnEdit.Text = "Edit Previous Result" Then
            Me.lvPrevious.Enabled = True

            Me.Size = New System.Drawing.Size(401, 354)
            Me.btnEdit.Text = "Exit Edit Mode"
            Me.btnEdit.Image = Me.ImageList1.Images(1)
            Me.btnEdit.TextAlign = ContentAlignment.MiddleCenter
            Me.EditMode = True
            Me.Editmodetext = "[Edit Mode]"
            Me.btnSave.Text = "Edit"
            Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
            Me.pctRecoveryPC.Visible = False
            Me.lblRecoveryPC.Visible = False
        Else

            Me.Size = New System.Drawing.Size(401, 232)
            Me.btnEdit.Text = "Edit Previous Result"
            Me.btnEdit.Image = Me.ImageList1.Images(0)
            Me.btnEdit.TextAlign = ContentAlignment.MiddleLeft
            Me.EditMode = False
            Me.Editmodetext = ""

            Me.btnSave.Text = "Save"
            Me.Text = "Select Rep(s) and a Sales Result" & "- " & Me.ID & " " & Me.Editmodetext
            Me.pctRecoveryPC.Visible = True
            Me.lblRecoveryPC.Visible = True

            If Me.lvPrevious.SelectedItems.Count <> 0 Then
                Me.lvPrevious.SelectedItems(0).Selected = False
            End If
            Me.lvPrevious.Enabled = False
            Me.Form_Reset()
            Me.GetData()
        End If
    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#Region "Contract Amount"
    Private Sub txt1ContractAmt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ContractAmt.GotFocus

        Dim x = Split(Me.txt1ContractAmt.Text, ",", 3)
        If InStr(Me.txt1ContractAmt.Text, ",") <> 0 Then
            Try
                Me.txt1ContractAmt.Text = x(0) & x(1) & x(2)
            Catch ex As Exception
                Me.txt1ContractAmt.Text = x(0) & x(1)
            End Try
        End If


        Me.txt1ContractAmt.SelectAll()
    End Sub

    Private Sub txt1ContractAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt1ContractAmt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        'If Me.txt1ContractAmt.Text.Length = 7 And x.ToString <> "" Or Me.txt1ContractAmt.SelectionLength = Me.txt1ContractAmt.Text.Length Then
        '    If Me.txt1ContractAmt.SelectionLength = 0 And Me.txt1ContractAmt.Text.Length <> 0 Then
        '        e.KeyChar = Nothing
        '    End If
        'End If
    End Sub

    Private Sub txt1ContractAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1ContractAmt.LostFocus
        If Me.txt1ContractAmt.Text <> "" Then
            Me.ErrorProvider1.Clear()
        End If

        Select Case Me.txt1ContractAmt.Text.Length
            Case Is = 4, 5, 6
                Dim x = Microsoft.VisualBasic.Left(Me.txt1ContractAmt.Text, Me.txt1ContractAmt.Text.Length - 3)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1ContractAmt.Text, 3)
                Me.txt1ContractAmt.Text = x & "," & y
            Case Is = 7
                Dim x = Microsoft.VisualBasic.Left(Me.txt1ContractAmt.Text, Me.txt1ContractAmt.Text.Length - 6)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1ContractAmt.Text, 3)
                Dim z = Microsoft.VisualBasic.Mid(Me.txt1ContractAmt.Text, 2, 3)
                Me.txt1ContractAmt.Text = x & "," & z & "," & y
        End Select
        Me.Split1 = 0
        Me.Split2 = 0
        Me.Split3 = 0
        Me.Split4 = 0
        Me.Split5 = 0
    End Sub

    Private Sub txt2ContractAmt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2ContractAmt.GotFocus
        Me.txt2ContractAmt.SelectAll()
    End Sub

    Private Sub txt2ContractAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt2ContractAmt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txt1ContractAmt.Text.Length = 2 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txt2ContractAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2ContractAmt.LostFocus
        If Me.txt2ContractAmt.Text = "" Then
            Me.txt2ContractAmt.Text = "00"
        End If
        Me.Split1 = 0
        Me.Split2 = 0
        Me.Split3 = 0
        Me.Split4 = 0
        Me.Split5 = 0
    End Sub
#End Region


    Private Sub rdoCash_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoCash.CheckedChanged
        Me.ErrorProvider1.Clear()
    End Sub

    Private Sub rdoFinance_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoFinance.CheckedChanged
        Me.ErrorProvider1.Clear()
    End Sub


    Private Sub lblContractNotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblContractNotes.Click
        Me.txtContractNotes.Focus()
    End Sub

    Private Sub txtContractNotes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContractNotes.GotFocus
        Me.lblContractNotes.Visible = False
    End Sub

    Private Sub txtContractNotes_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtContractNotes.LostFocus
        If Me.txtContractNotes.Text = "" Then
            Me.lblContractNotes.Visible = True
        Else
            Me.lblContractNotes.Visible = False
        End If
    End Sub


    Private Sub btnMovetoSold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMovetoSold.Click
        If Me.lstManage.SelectedItems.Count = 0 Then
            MsgBox("You must select a product to move!", MsgBoxStyle.Exclamation, "No Product Selected")
            Exit Sub
        End If
        If Me.lstSold.Items.Count = 5 Then
            Dim tt As New ToolTip
            tt.IsBalloon = True
            tt.ToolTipTitle = "Sold Products"
            tt.ToolTipIcon = ToolTipIcon.Warning
            tt.Show("You can only log five" & vbCr & _
                    "products per Sale!", Me.lstSold, 105, -95)
            Exit Sub
        End If
        Me.lstSold.Items.Add(Me.lstManage.SelectedItem)
        Me.lstManage.Items.Remove(Me.lstManage.SelectedItem)
        Me.ErrorProvider1.Clear()
        If Me.lstManage.Items.Count <> 0 Then
            Me.lstManage.SetSelected(0, True)
        End If

    End Sub

    Private Sub btnMovetoFuture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMovetoFuture.Click
        If Me.lstSold.SelectedItems.Count = 0 Then
            MsgBox("You must select a product to move!", MsgBoxStyle.Exclamation, "No Product Selected")
            Exit Sub
        End If
        Me.lstManage.Items.Add(Me.lstSold.SelectedItem)
        Me.lstSold.Items.Remove(Me.lstSold.SelectedItem)
        If Me.lstSold.Items.Count <> 0 Then
            Me.lstSold.SetSelected(0, True)
        End If
    End Sub




    Private Sub lstSold_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstSold.GotFocus
        Me.lstManage.SelectedItem = Nothing

    End Sub

    Private Sub lstManage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstManage.GotFocus
        Me.lstSold.SelectedItem = Nothing
    End Sub

    Private Sub btnDeleteProduct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteProduct.Click
        If Me.lstSold.SelectedItems.Count = 0 And Me.lstManage.SelectedItems.Count = 0 Then
            MsgBox("You must select a product to remove!", MsgBoxStyle.Exclamation, "No Product Selected")
            Exit Sub
        End If
        If Me.lstSold.SelectedItems.Count <> 0 Then
            Me.lstSold.Items.Remove(Me.lstSold.SelectedItem)
        End If
        If Me.lstManage.SelectedItems.Count <> 0 Then
            Me.lstManage.Items.Remove(Me.lstManage.SelectedItem)
        End If
    End Sub

    Private Sub txtP1Amnt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP1Amnt.GotFocus
        Me.txtP1Amnt.SelectAll()
    End Sub



    Private Sub txtP1Amnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtP1Amnt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr Or x.ToString = ".") Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtP1Amnt.Text.Length = 7 Or (Me.txtP1Amnt.Text.Length = 0 And e.KeyChar = ".") Or (InStr(Me.txtP1Amnt.Text, ".") > 0 And e.KeyChar = ".") Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtP1Amnt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP1Amnt.LostFocus
        If Me.txtP1Amnt.Text = "" Or Me.txtP1Amnt.Text = "." Then
            Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
        End If

        If Me.txtP1Amnt.Text.Chars(Me.txtP1Amnt.Text.Length - 1) = "." Then
            Me.txtP1Amnt.Text = Microsoft.VisualBasic.Left(Me.txtP1Amnt.Text, Me.txtP1Amnt.Text.Length - 1)
        End If

        Dim cnt As Integer = Me.lstSold.Items.Count
        Select Case cnt
            Case Is = 1
                Me.txtP1Amnt.Text = Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text
            Case Is = 2, 3, 4, 5
                Me.txtP2Amnt.Text = CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) - CType(Me.txtP1Amnt.Text, Double)
        End Select
        Me.Split1 = CType(Me.txtP1Amnt.Text, Double)
        Me.Split2 = CType(Me.txtP2Amnt.Text, Double)
        Me.Split3 = CType(Me.txtP3Amnt.Text, Double)
        Me.Split4 = CType(Me.txtP4Amnt.Text, Double)
        Me.Split5 = CType(Me.txtP5Amnt.Text, Double)
    End Sub

    Private Sub txtP2Amnt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP2Amnt.GotFocus
        Me.txtP2Amnt.SelectAll()
    End Sub


    Private Sub txtP2Amnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtP2Amnt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr Or x.ToString = ".") Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtP2Amnt.Text.Length = 7 Or (Me.txtP2Amnt.Text.Length = 0 And e.KeyChar = ".") Or (InStr(Me.txtP2Amnt.Text, ".") > 0 And e.KeyChar = ".") Then
            e.KeyChar = Nothing
        End If
    End Sub




    Private Sub txtP2Amnt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP2Amnt.LostFocus
        If Me.txtP2Amnt.Text = "" Or Me.txtP2Amnt.Text = "." Then
            Me.txtP2Amnt.Text = "0.00"
        End If
        If Me.txtP2Amnt.Text.Chars(Me.txtP2Amnt.Text.Length - 1) = "." Then
            Me.txtP2Amnt.Text = Microsoft.VisualBasic.Left(Me.txtP2Amnt.Text, Me.txtP2Amnt.Text.Length - 1)
        End If
        Dim cnt As Integer = Me.lstSold.Items.Count
        Select Case cnt
            Case Is = 2

                Me.txtP1Amnt.Text = CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) - CType(Me.txtP2Amnt.Text, Double)
            Case Is = 3, 4, 5

                Me.txtP3Amnt.Text = CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) - (CType(Me.txtP1Amnt.Text, Double) + CType(Me.txtP2Amnt.Text, Double))
        End Select
        Me.Split1 = CType(Me.txtP1Amnt.Text, Double)
        Me.Split2 = CType(Me.txtP2Amnt.Text, Double)
        Me.Split3 = CType(Me.txtP3Amnt.Text, Double)
        Me.Split4 = CType(Me.txtP4Amnt.Text, Double)
        Me.Split5 = CType(Me.txtP5Amnt.Text, Double)
    End Sub

    Private Sub txtP3Amnt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP3Amnt.GotFocus
        Me.txtP3Amnt.SelectAll()
    End Sub



    Private Sub txtP3Amnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtP3Amnt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr Or x.ToString = ".") Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtP3Amnt.Text.Length = 7 Or (Me.txtP3Amnt.Text.Length = 0 And e.KeyChar = ".") Or (InStr(Me.txtP3Amnt.Text, ".") > 0 And e.KeyChar = ".") Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtP3Amnt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP3Amnt.LostFocus
        If Me.txtP3Amnt.Text = "" Or Me.txtP3Amnt.Text = "." Then
            Me.txtP3Amnt.Text = "0.00"
        End If
        If Me.txtP3Amnt.Text.Chars(Me.txtP3Amnt.Text.Length - 1) = "." Then
            Me.txtP3Amnt.Text = Microsoft.VisualBasic.Left(Me.txtP3Amnt.Text, Me.txtP3Amnt.Text.Length - 1)
        End If
        Dim cnt As Integer = Me.lstSold.Items.Count
        Select Case cnt
            Case Is = 4, 5
                Me.txtP4Amnt.Text = CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) - (CType(Me.txtP1Amnt.Text, Double) + CType(Me.txtP2Amnt.Text, Double) + CType(Me.txtP3Amnt.Text, Double))
        End Select
        Me.Split1 = CType(Me.txtP1Amnt.Text, Double)
        Me.Split2 = CType(Me.txtP2Amnt.Text, Double)
        Me.Split3 = CType(Me.txtP3Amnt.Text, Double)
        Me.Split4 = CType(Me.txtP4Amnt.Text, Double)
        Me.Split5 = CType(Me.txtP5Amnt.Text, Double)
    End Sub

    Private Sub txtP4Amnt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP4Amnt.GotFocus
        Me.txtP4Amnt.SelectAll()
    End Sub



    Private Sub txtP4Amnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtP4Amnt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr Or x.ToString = ".") Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtP4Amnt.Text.Length = 7 Or (Me.txtP4Amnt.Text.Length = 0 And e.KeyChar = ".") Or (InStr(Me.txtP4Amnt.Text, ".") > 0 And e.KeyChar = ".") Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtP4Amnt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP4Amnt.LostFocus
        If Me.txtP4Amnt.Text = "" Or Me.txtP4Amnt.Text = "." Then
            Me.txtP4Amnt.Text = "0.00"
        End If
        If Me.txtP4Amnt.Text.Chars(Me.txtP4Amnt.Text.Length - 1) = "." Then
            Me.txtP4Amnt.Text = Microsoft.VisualBasic.Left(Me.txtP4Amnt.Text, Me.txtP4Amnt.Text.Length - 1)
        End If
        Dim cnt As Integer = Me.lstSold.Items.Count
        Select Case cnt
            Case Is = 5
                Me.txtP5Amnt.Text = CType(Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text, Double) - (CType(Me.txtP1Amnt.Text, Double) + CType(Me.txtP2Amnt.Text, Double) + CType(Me.txtP3Amnt.Text, Double) + CType(Me.txtP4Amnt.Text, Double))
        End Select
        Me.Split1 = CType(Me.txtP1Amnt.Text, Double)
        Me.Split2 = CType(Me.txtP2Amnt.Text, Double)
        Me.Split3 = CType(Me.txtP3Amnt.Text, Double)
        Me.Split4 = CType(Me.txtP4Amnt.Text, Double)
        Me.Split5 = CType(Me.txtP5Amnt.Text, Double)
    End Sub

    Private Sub txtP5Amnt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP5Amnt.GotFocus
        Me.txtP5Amnt.SelectAll()
    End Sub

    Private Sub lnkCalc_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkCalc.LinkClicked
        Process.Start("calc.exe")
    End Sub

    Private Sub lnkP1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkP1.LinkClicked
        ProductDetail.Product = Me.txtP1.Text
        ProductDetail.Prodnum = 1
        ProductDetail.ShowInTaskbar = False
        ProductDetail.ShowDialog()
    End Sub

    Private Sub lnkP2_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkP2.LinkClicked
        ProductDetail.Product = Me.txtP2.Text
        ProductDetail.Prodnum = 2
        ProductDetail.ShowInTaskbar = False
        ProductDetail.ShowDialog()
    End Sub

    Private Sub lnkP3_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkP3.LinkClicked
        ProductDetail.Product = Me.txtP3.Text
        ProductDetail.Prodnum = 3
        ProductDetail.ShowInTaskbar = False
        ProductDetail.ShowDialog()
    End Sub

    Private Sub lnkP4_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkP4.LinkClicked
        ProductDetail.Product = Me.txtP4.Text
        ProductDetail.Prodnum = 4
        ProductDetail.ShowInTaskbar = False
        ProductDetail.ShowDialog()
    End Sub

    Private Sub lnkP5_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkP5.LinkClicked
        ProductDetail.Product = Me.txtP5.Text
        ProductDetail.Prodnum = 5
        ProductDetail.ShowInTaskbar = False
        ProductDetail.ShowDialog()
    End Sub

    Private Sub txtP5Amnt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtP5Amnt.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr Or x.ToString = ".") Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txtP5Amnt.Text.Length = 7 Or (Me.txtP5Amnt.Text.Length = 0 And e.KeyChar = ".") Or (InStr(Me.txtP5Amnt.Text, ".") > 0 And e.KeyChar = ".") Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txtP5Amnt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtP5Amnt.LostFocus
        If Me.txtP5Amnt.Text = "" Or Me.txtP5Amnt.Text = "." Then
            Me.txtP5Amnt.Text = "0.00"
        End If
        If Me.txtP5Amnt.Text.Chars(Me.txtP5Amnt.Text.Length - 1) = "." Then
            Me.txtP5Amnt.Text = Microsoft.VisualBasic.Left(Me.txtP5Amnt.Text, Me.txtP5Amnt.Text.Length - 1)
        End If
        Me.Split1 = CType(Me.txtP1Amnt.Text, Double)
        Me.Split2 = CType(Me.txtP2Amnt.Text, Double)
        Me.Split3 = CType(Me.txtP3Amnt.Text, Double)
        Me.Split4 = CType(Me.txtP4Amnt.Text, Double)
        Me.Split5 = CType(Me.txtP5Amnt.Text, Double)
    End Sub

    Private Sub cboRep2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRep2.SelectedValueChanged
        If Me.cboRep1.Text = Nothing And Me.cboRep2.Text <> Nothing Then
            Me.cboRep1.Text = Me.cboRep2.Text
            Me.cboRep2.Text = Nothing
        End If
    End Sub

    Private Sub cboProducts_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboProducts.SelectedValueChanged

        If Me.cboProducts.Text = "<Add New>" Then
            '' Need stored procedure here 

            Exit Sub
        ElseIf Me.cboProducts.Text = Nothing Or Me.cboProducts.Text = "___________________________________________________" Then
            Me.cboProducts.Text = Nothing
            Exit Sub
        End If
        If Me.lstManage.Items.Count = 3 Then
            Dim tt As New ToolTip

        End If
        Me.lstManage.Items.Add(Me.cboProducts.Text)
        Me.cboProducts.Text = Nothing
    End Sub

    Private Sub lvPrevious_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvPrevious.SelectedIndexChanged
        If Me.lvPrevious.SelectedItems.Count = 0 Then
            Exit Sub
        End If
        Me.Form_Reset()
        Me.LHID = Me.lvPrevious.SelectedItems(0).SubItems(9).Text

        Me.cboRep1.Text = Me.lvPrevious.SelectedItems(0).SubItems(3).Text
        If cboRep1.Text = Nothing Or Me.cboRep1.Text <> Me.lvPrevious.SelectedItems(0).SubItems(3).Text And Me.lvPrevious.SelectedItems(0).SubItems(3).Text <> "" Then
            Me.cboRep1.Items.Add(Me.lvPrevious.SelectedItems(0).SubItems(3).Text)
            Me.cboRep2.Items.Add(Me.lvPrevious.SelectedItems(0).SubItems(3).Text)
            Me.cboRep1.Text = Me.lvPrevious.SelectedItems(0).SubItems(3).Text
        End If
        Me.cboRep2.Text = Me.lvPrevious.SelectedItems(0).SubItems(4).Text
        If cboRep2.Text = Nothing Or Me.cboRep1.Text <> Me.lvPrevious.SelectedItems(0).SubItems(4).Text And Me.lvPrevious.SelectedItems(0).SubItems(4).Text <> "" Then
            Me.cboRep1.Items.Add(Me.lvPrevious.SelectedItems(0).SubItems(4).Text)
            Me.cboRep2.Items.Add(Me.lvPrevious.SelectedItems(0).SubItems(4).Text)
            Me.cboRep2.Text = Me.lvPrevious.SelectedItems(0).SubItems(4).Text
        End If
        Me.cboSalesResults.Text = Me.lvPrevious.SelectedItems(0).SubItems(2).Text
        If Me.cboSalesResults.Text = "Sale" Then
            Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet1 As SqlCommand
            Dim r1 As SqlDataReader
            cmdGet1 = New SqlCommand("Select * from SaleDetail where LeadHistoryID = " & Me.LHID, cnn)
            cmdGet1.CommandType = CommandType.Text
            cnn.Open()
            r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
            r1.Read()
            Me.P1 = r1.Item(3)
            Me.P2 = r1.Item(4)
            Me.P3 = r1.Item(5)
            Me.P4 = r1.Item(6)
            Me.P5 = r1.Item(7)
            Me.PM1 = r1.Item(8)
            Me.PM2 = r1.Item(9)
            Me.PM3 = r1.Item(10)
            Me.PM4 = r1.Item(11)
            Me.PM5 = r1.Item(12)
            Me.PModel1 = r1.Item(13)
            Me.PModel2 = r1.Item(14)
            Me.PModel3 = r1.Item(15)
            Me.PModel4 = r1.Item(16)
            Me.PModel5 = r1.Item(17)
            Me.PStyle1 = r1.Item(28)
            Me.PStyle2 = r1.Item(29)
            Me.PStyle3 = r1.Item(30)
            Me.PStyle4 = r1.Item(31)
            Me.PStyle5 = r1.Item(32)
            Me.PColor1 = r1.Item(18)
            Me.PColor2 = r1.Item(19)
            Me.PColor3 = r1.Item(20)
            Me.PColor4 = r1.Item(21)
            Me.PColor5 = r1.Item(22)
            Me.Unit1 = r1.Item(23)
            Me.Unit2 = r1.Item(24)
            Me.Unit3 = r1.Item(25)
            Me.Unit4 = r1.Item(26)
            Me.Unit5 = r1.Item(27)
            Me.Qty1 = r1.Item(33)
            Me.Qty2 = r1.Item(34)
            Me.Qty3 = r1.Item(35)
            Me.Qty4 = r1.Item(36)
            Me.Qty5 = r1.Item(37)
            Me.Split1 = CType(r1.Item(38), Double)
            Me.Split2 = CType(r1.Item(39), Double)
            Me.Split3 = CType(r1.Item(40), Double)
            Me.Split4 = CType(r1.Item(41), Double)
            Me.Split5 = CType(r1.Item(42), Double)
            Me.rdoCash.Checked = r1.Item(43)
            Me.rdoFinance.Checked = r1.Item(44)
            Me.dtpRecDate.Value = r1.Item(45)
            Dim x = Split(r1.Item(46), ".")
            Me.txt1ContractAmt.Text = x(0)
            Me.txt2ContractAmt.Text = x(1)
            Me.txtContractNotes.Text = r1.Item(47)
            Me.txtContractNotes_LostFocus(Nothing, Nothing)
            r1.Close()
            cnn.Close()
            Me.lstSold.Items.Add(Me.P1)
            If Me.P2 <> "" Then
                Me.lstSold.Items.Add(Me.P2)
            End If
            If Me.P3 <> "" Then
                Me.lstSold.Items.Add(Me.P3)
            End If
            If Me.P4 <> "" Then
                Me.lstSold.Items.Add(Me.P4)
            End If
            If Me.P5 <> "" Then
                Me.lstSold.Items.Add(Me.P5)
            End If
        End If
        Me.rtfNote.Text = Me.lvPrevious.SelectedItems(0).SubItems(8).Text
        Me.rtfNote_LostFocus(Nothing, Nothing)
        Me.chkRecoverable.Checked = Me.lvPrevious.SelectedItems(0).SubItems(7).Text
        Dim z = Split(Me.lvPrevious.SelectedItems(0).SubItems(5).Text, ".")
        Dim y = Split(Me.lvPrevious.SelectedItems(0).SubItems(6).Text, ".")
        Me.txt1Quoted.Text = z(0)
        Me.txt2Quoted.Text = z(1)
        Me.txt1Par.Text = y(0)
        Me.txt2Par.Text = y(1)
    End Sub

    Private Sub rtfNote_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfNote.GotFocus
        Me.lblsalesnotes.Visible = False
    End Sub

    Private Sub rtfNote_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtfNote.LostFocus
        If Me.rtfNote.Text = "" Then
            Me.lblsalesnotes.Visible = True
        Else
            Me.lblsalesnotes.Visible = False
        End If
    End Sub

    Private Sub lblsalesnotes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblsalesnotes.Click
        Me.rtfNote.Focus()

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Me.cboautonotes.Focus()
        Me.cboautonotes.DroppedDown = True

    End Sub

    Private Sub cboautonotes_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.GotFocus
        Me.Label1.Visible = False
    End Sub

    Private Sub cboautonotes_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.LostFocus
        Me.Label1.Visible = True
    End Sub


    Private Sub cboautonotes_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboautonotes.SelectedValueChanged
        If Me.cboautonotes.Text = Nothing Or Me.cboautonotes.Text = "________________________________________________________________" Then
            Exit Sub
        ElseIf Me.cboautonotes.Text = "<Add New>" Then
            Me.cboautonotes.SelectedItem = Nothing
            Dim i As String
            i = InputBox$("Enter a new ""Auto Note"" here.", "Save Auto Note")

            If i = "" Then
                MsgBox("You must enter Text to save this Auto Note!", MsgBoxStyle.Exclamation, "No Text Supplied")
                Exit Sub
            End If
            Dim cnn = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdINS As SqlCommand = New SqlCommand("dbo.InsSalesNotes", cnn)
            Dim cmdget As SqlCommand = New SqlCommand("dbo.GetSMAutoNotes", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Notes", i)
            cmdINS.Parameters.Add(param1)
            cmdINS.CommandType = CommandType.StoredProcedure
            cmdget.CommandType = CommandType.StoredProcedure
            Dim r1 As SqlDataReader
            Dim r2 As SqlDataReader
            Me.cboautonotes.Items.Clear()
            Me.cboautonotes.Items.Add("<Add New>")
            Me.cboautonotes.Items.Add("___________________________________________________")
            cnn.Open()
            r2 = cmdINS.ExecuteReader
            r2.Close()
            r1 = cmdget.ExecuteReader
            While r1.Read
                Me.cboautonotes.Items.Add(r1.Item(0))
            End While
            r1.Close()
            cnn.Close()

            Me.cboautonotes.Text = i

        End If

        If Me.rtfNote.Text <> "" Then
            If Me.cboautonotes.Text <> "" Then
                Me.rtfNote.Text = Me.rtfNote.Text & ", " & Me.cboautonotes.Text
            End If

        Else
            Me.rtfNote.Text = Me.cboautonotes.Text
        End If
        Me.cboautonotes.Text = Nothing
    End Sub

    Private Sub txt1Quoted_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1Quoted.GotFocus
        Me.txt1Quoted.SelectAll()
    End Sub


    Private Sub txt1Quoted_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt1Quoted.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txt1Quoted.Text.Length = 7 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txt1Quoted_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1Quoted.LostFocus
        If Me.txt1Quoted.Text <> "" Then
            Me.ErrorProvider1.Clear()
        End If

        Select Case Me.txt1Quoted.Text.Length
            Case Is = 4, 5, 6
                Dim x = Microsoft.VisualBasic.Left(Me.txt1Quoted.Text, Me.txt1Quoted.Text.Length - 3)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1Quoted.Text, 3)
                Me.txt1Quoted.Text = x & "," & y
            Case Is = 7
                Dim x = Microsoft.VisualBasic.Left(Me.txt1Quoted.Text, Me.txt1Quoted.Text.Length - 6)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1Quoted.Text, 3)
                Dim z = Microsoft.VisualBasic.Mid(Me.txt1Quoted.Text, 2, 3)
                Me.txt1Quoted.Text = x & "," & z & "," & y
        End Select
    End Sub

    Private Sub txt2Quoted_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2Quoted.GotFocus
        Me.txt2Quoted.SelectAll()
    End Sub

    Private Sub txt2Quoted_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt2Quoted.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txt2Quoted.Text.Length = 2 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txt2Quoted_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2Quoted.LostFocus
        If Me.txt2Quoted.Text = "" Then
            Me.txt2Quoted.Text = "00"
        End If
    End Sub

    Private Sub txt1Par_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1Par.GotFocus
        Me.txt1Par.SelectAll()
    End Sub

    Private Sub txt1Par_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt1Par.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txt1Par.Text.Length = 7 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txt1Par_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1Par.LostFocus
        If Me.txt1Par.Text <> "" Then
            Me.ErrorProvider1.Clear()
        End If

        Select Case Me.txt1Par.Text.Length
            Case Is = 4, 5, 6
                Dim x = Microsoft.VisualBasic.Left(Me.txt1Par.Text, Me.txt1Par.Text.Length - 3)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1Par.Text, 3)
                Me.txt1Par.Text = x & "," & y
            Case Is = 7
                Dim x = Microsoft.VisualBasic.Left(Me.txt1Par.Text, Me.txt1Par.Text.Length - 6)
                Dim y = Microsoft.VisualBasic.Right(Me.txt1Par.Text, 3)
                Dim z = Microsoft.VisualBasic.Mid(Me.txt1Par.Text, 2, 3)
                Me.txt1Par.Text = x & "," & z & "," & y
        End Select
    End Sub

    Private Sub txt2Par_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2Par.GotFocus
        Me.txt2Par.SelectAll()
    End Sub

    Private Sub txt2Par_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txt2Par.KeyPress
        Dim x = e.KeyChar.ToString
        If (x.ToString = "1" Or x.ToString = "2" Or x.ToString = "3" Or x.ToString = "4" Or _
        x.ToString = "5" Or x.ToString = "6" Or x.ToString = "7" Or x.ToString = "8" Or x.ToString = "9" Or _
        x.ToString = "0" Or x.ToString = "" Or x.ToString = vbCr) Then
        Else
            MsgBox("This field only accepts numbers!", MsgBoxStyle.Exclamation, "Numbers Only")
            e.KeyChar = Nothing
        End If
        If Me.txt2Par.Text.Length = 2 Then
            e.KeyChar = Nothing
        End If
    End Sub

    Private Sub txt2Par_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt2Par.LostFocus
        If Me.txt2Par.Text = "" Then
            Me.txt2Par.Text = "00"
        End If
    End Sub

    Private Sub DeleteThisProductToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteThisProductToolStripMenuItem.Click
        If Me.lvPrevious.SelectedItems.Count = 0 Then
            MsgBox("You must select a Previous Result to Delete!", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If
        If Me.lvPrevious.SelectedItems(0).Tag = "Locked" Then
            MsgBox("This result was a ""System Required Result""," & vbCr & "and has been locked by the system!" & vbCr & "You may edit this result, but it cannot be deleted.", MsgBoxStyle.Exclamation, "Cannot Delete Result")
            Exit Sub
        End If
        Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", Me.ID)
        Dim param2 As SqlParameter = New SqlParameter("@LHID", Me.LHID)
        Dim cmdGet As SqlCommand
        Dim r As SqlDataReader
        cmdGet = New SqlCommand("DeleteSalesResult", cnn1)
        cmdGet.CommandType = CommandType.StoredProcedure
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cnn1.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r.Read()
        r.Close()
        cnn1.Close()

        Me.Form_Reset()
        Me.GetData()

        Sales.ForceRefresh = True
        If Sales.TabControl2.TabIndex = 0 And Sales.lvSales.SelectedItems.Count <> 0 Then
            Sales.lvSales_SelectedIndexChanged(Nothing, Nothing)
        ElseIf Sales.TabControl2.TabIndex = 1 And Sales.lvMemorized.SelectedItems.Count <> 0 Then
            Sales.lvMemorized_SelectedIndexChanged(Nothing, Nothing)
        End If

    End Sub



    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim RD As String = Me.dtpRecDate.Value.ToShortDateString
        RD = RD & " 12:00:00 AM"
        RD = CDate(RD)
        Dim cash As Boolean
        Dim finance As Boolean
        If Me.rdoCash.Checked = True And Me.rdoFinance.Checked = False Then
            cash = True
            finance = False
        ElseIf Me.rdoCash.Checked = False And Me.rdoFinance.Checked = True Then
            cash = False
            finance = True
        ElseIf Me.rdoCash.Checked = False And Me.rdoFinance.Checked = False Then
            cash = False
            finance = False
        End If
        If Me.txt1Quoted.Text = "" Then
            Me.txt1Quoted.Text = "0"
        End If
        If Me.txt2Quoted.Text = "" Then
            Me.txt2Quoted.Text = "00"
        End If
        If Me.txt1Par.Text = "" Then
            Me.txt1Par.Text = "0"
        End If
        If Me.txt2Par.Text = "" Then
            Me.txt2Par.Text = "00"
        End If
        If Me.txt1ContractAmt.Text = "" Then
            Me.txt1ContractAmt.Text = "0"
        End If
        If Me.txt2ContractAmt.Text = "" Then
            Me.txt2ContractAmt.Text = "00"
        End If
        Me.txt2Quoted.Text = Microsoft.VisualBasic.Left(Me.txt2Quoted.Text, 2)
        Me.txt2Par.Text = Microsoft.VisualBasic.Left(Me.txt2Par.Text, 2)
        Me.txt2ContractAmt.Text = Microsoft.VisualBasic.Left(Me.txt2Par.Text, 2)
        If Me.cboSalesResults.Text = "Sale" Then
            If Me.lstManage.Items.Count = 3 Then
                Me.Product1 = Me.lstManage.Items(0)
                Me.Product2 = Me.lstManage.Items(1)
                Me.Product3 = Me.lstManage.Items(2)
            ElseIf Me.lstManage.Items.Count = 2 Then
                Me.Product1 = Me.lstManage.Items(0)
                Me.Product2 = Me.lstManage.Items(1)
                Me.Product3 = ""
            ElseIf Me.lstManage.Items.Count = 1 Then
                Me.Product1 = Me.lstManage.Items(0)
                Me.Product2 = ""
                Me.Product3 = ""
            Else
                Me.Product1 = ""
                Me.Product2 = ""
                Me.Product3 = ""
            End If
        End If

        If Me.chkRecoverable.Checked = True Then
            Me.Recoverable = True
        Else
            Me.Recoverable = False
        End If

        Dim Reps As String
        If Me.cboRep1.Text <> "" And Me.cboRep2.Text = "" Or Me.cboRep2.Text = " " Then
            Reps = Me.cboRep1.Text
        Else
            Reps = Me.cboRep1.Text & " and " & Me.cboRep2.Text
        End If

        Dim dep2 As String
        Dim cnn2 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdGet2 As SqlCommand
        Dim r2 As SqlDataReader
        cmdGet2 = New SqlCommand("Select isrecovery from Enterlead where id = " & Me.ID, cnn2)
        cmdGet2.CommandType = CommandType.Text
        cnn2.Open()
        r2 = cmdGet2.ExecuteReader(CommandBehavior.CloseConnection)
        r2.Read()
        If r2.Item(0) = True Then
            dep2 = "Recovery"
        Else
            dep2 = "Marketing"
        End If
        If Me.EditMode = True Then
            If Me.lvPrevious.Items(0).SubItems(7).Text = True And Me.cboSalesResults.Text <> "Demo/No Sale" Then
                dep2 = "Marketing"
            ElseIf Me.lvPrevious.Items(0).SubItems(7).Text = True And Me.cboSalesResults.Text <> "Recission Cancel" Then
                dep2 = "Marketing"
            End If
        End If
        r2.Close()
        cnn2.Close()
        Dim description As String
        If Me.cboSalesResults.Text = "Not Issued" Then
            description = "Appt. was not issued. Logged by " & STATIC_VARIABLES.CurrentUser & " (Forwarded back to the " & dep2 & " Department to be rescheduled)"
        ElseIf Me.cboSalesResults.Text = "Lost Result" Then
            If Me.cboRep1.Text = "" Then
                description = "Appt. Logged as ""Lost Result"" by " & STATIC_VARIABLES.CurrentUser & ", and has been forwarded back to the " & dep2 & " Department"
            Else
                description = "Appt. issued to " & Reps & ", and has been logged as ""Lost Result"" by " & STATIC_VARIABLES.CurrentUser & ". Record has been forwarded back to the " & dep2 & " Department"
            End If
        ElseIf Me.cboSalesResults.Text = "Recission Cancel" Then
            Dim cnn1 As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
            Dim cmdGet1 As SqlCommand
            Dim r1 As SqlDataReader
            cmdGet1 = New SqlCommand("Select Contact1FirstName, Contact2FirstName from Enterlead where id = " & Me.ID, cnn1)
            cmdGet1.CommandType = CommandType.Text
            cnn1.Open()
            r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
            r1.Read()
            Dim Customer As String
            Dim hashave As String
            Dim theirhisher As String
            If r1.Item(1) = "" Then
                Customer = r1.Item(0)
                hashave = "has"
                theirhisher = "his/her"
            Else
                Customer = r1.Item(0) & " and " & r1.Item(1)
                hashave = "have"
                theirhisher = "their"
            End If
            r1.Close()
            cnn1.Close()

            cmdGet1 = New SqlCommand("Select Financing,Installation from tblwherecanleadgo where leadnumber = " & Me.ID, cnn1)
            cnn1.Open()
            r1 = cmdGet1.ExecuteReader(CommandBehavior.CloseConnection)
            r1.Read()
            Dim dep As String
            If r1.Item(0) = True And r1.Item(1) = True Then
                dep = "(Removed from Installation and Finance Departments)"
            ElseIf r1.Item(0) = True And r1.Item(1) = False Then
                dep = "(Removed from Finance Department)"
            ElseIf r1.Item(1) = True And r1.Item(0) = False Then
                dep = "(Removed from Installation Department)"
            Else
                dep = ""
            End If
            r1.Close()
            cnn1.Close()

            If Me.Recoverable = False Then
                description = Customer & " " & hashave & " decided to cancel " & theirhisher & " contract within the ""Right to Rescind"" date. Recission Cancel logged by " & STATIC_VARIABLES.CurrentUser & "." & dep
            Else
                description = Customer & " " & hashave & " decided to cancel " & theirhisher & " contract within the ""Right to Rescind"" date. Recission Cancel logged by " & STATIC_VARIABLES.CurrentUser & ", and forwarded to the Recovery Department." & dep
            End If
        ElseIf Me.cboSalesResults.Text = "Bank Rejected" Then
            description = "Due to credit issues, this contract has been rejected by the Finance Department. Rejection logged by " & STATIC_VARIABLES.CurrentUser
        ElseIf Me.cboSalesResults.Text = "Bank Approved" Then
            description = "Contract has been approved by the Finance Department. Approval logged by " & STATIC_VARIABLES.CurrentUser
        ElseIf Me.cboSalesResults.Text = "Demo/No Sale" Then
            description = "Appt. Issued to " & Reps & ", which resulted in a Demo/No Sale. Price Quoted was $" & Me.txt1Quoted.Text & "." & Me.txt2Quoted.Text & " and Par Price was $" & Me.txt1Par.Text & "." & Me.txt2Par.Text & ". Logged by " & STATIC_VARIABLES.CurrentUser
            If Recoverable = True Then
                description = description & " (Forwarded to the Recovery Department)"
            End If
        ElseIf Me.cboSalesResults.Text = "Sale" Then
            Dim Products As String = Me.P1
            If Me.P2 <> "" Then
                If Me.P2 <> "" And Me.P3 = "" Then
                    Products = Products & " and " & Me.P2
                Else
                    Products = Products & ", " & Me.P2
                    If Me.P4 = "" Then
                        Products = Products & ", and " & Me.P3
                    Else
                        Products = Products & ", " & Me.P3
                        If Me.P5 = "" Then
                            Products = Products & ", and " & Me.P4
                        Else
                            Products = Products & ", " & Me.P4 & ", and " & Me.P5
                        End If
                    End If
                End If
            End If

            If Me.rdoFinance.Checked = True Then
                description = "Appt. Issued to " & Reps & ", which resulted in a Sale in the amount of $" & Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text & ". Products Sold- " & Products & ". (Forwarded to Finance Department for approval)"
            Else
                description = "Appt. Issued to " & Reps & ", which resulted in a Sale in the amount of $" & Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text & ". Products Sold- " & Products & ". (Forwarded to Installation Department)"
            End If


        ElseIf Me.cboSalesResults.Text = "No Demo" Then
            description = "Appt. Issued to " & Reps & ", which resulted in a No Demo. Logged by " & STATIC_VARIABLES.CurrentUser
        ElseIf Me.cboSalesResults.Text = "Reset" Or Me.cboSalesResults.Text = "Not Hit" Then
            description = "Appt. issued to " & Reps & ", which resulted in a " & Me.cboSalesResults.Text & ". Logged by " & STATIC_VARIABLES.CurrentUser & " (Forwarded back to the " & dep2 & " Department to be rescheduled)"
        End If



        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim param1 As SqlParameter = New SqlParameter("@NR", Me.NR)
        Dim param2 As SqlParameter = New SqlParameter("@Editmode", Me.EditMode)
        Dim param3 As SqlParameter = New SqlParameter("@Result", Me.cboSalesResults.Text)
        Dim param4 As SqlParameter = New SqlParameter("@rep1", Me.cboRep1.Text)
        Dim param5 As SqlParameter = New SqlParameter("@Rep2", Me.cboRep2.Text)
        Dim param6 As SqlParameter = New SqlParameter("@RDate", RD)
        Dim param7 As SqlParameter = New SqlParameter("@ContractAmt", Me.txt1ContractAmt.Text & "." & Me.txt2ContractAmt.Text)
        Dim param8 As SqlParameter = New SqlParameter("@cash", cash)
        Dim param9 As SqlParameter = New SqlParameter("@finance", finance)
        Dim param10 As SqlParameter = New SqlParameter("@contractnotes", Me.txtContractNotes.Text)
        Dim param11 As SqlParameter = New SqlParameter("@Product1", Me.Product1)
        Dim param12 As SqlParameter = New SqlParameter("@Product2", Me.Product2)
        Dim param13 As SqlParameter = New SqlParameter("@Product3", Me.Product3)
        Dim param14 As SqlParameter = New SqlParameter("@Sold1", Me.P1)
        Dim param15 As SqlParameter = New SqlParameter("@Sold2", Me.P2)
        Dim param16 As SqlParameter = New SqlParameter("@Sold3", Me.P3)
        Dim param17 As SqlParameter = New SqlParameter("@Sold4", Me.P4)
        Dim param18 As SqlParameter = New SqlParameter("@Sold5", Me.P5)
        Dim param19 As SqlParameter = New SqlParameter("@Split1", Me.Split1)
        Dim param20 As SqlParameter = New SqlParameter("@Split2", Me.Split2)
        Dim param21 As SqlParameter = New SqlParameter("@Split3", Me.Split3)
        Dim param22 As SqlParameter = New SqlParameter("@Split4", Me.Split4)
        Dim param23 As SqlParameter = New SqlParameter("@Split5", Me.Split5)
        Dim param24 As SqlParameter = New SqlParameter("@Manufacturer1", Me.PM1)
        Dim param25 As SqlParameter = New SqlParameter("@Manufacturer2", Me.PM2)
        Dim param26 As SqlParameter = New SqlParameter("@Manufacturer3", Me.PM3)
        Dim param27 As SqlParameter = New SqlParameter("@Manufacturer4", Me.PM4)
        Dim param28 As SqlParameter = New SqlParameter("@Manufacturer5", Me.PM5)
        Dim param29 As SqlParameter = New SqlParameter("@Model1", Me.PModel1)
        Dim param30 As SqlParameter = New SqlParameter("@Model2", Me.PModel2)
        Dim param31 As SqlParameter = New SqlParameter("@Model3", Me.PModel3)
        Dim param32 As SqlParameter = New SqlParameter("@Model4", Me.PModel4)
        Dim param33 As SqlParameter = New SqlParameter("@Model5", Me.PModel5)
        Dim param34 As SqlParameter = New SqlParameter("@Style1", Me.PStyle1)
        Dim param35 As SqlParameter = New SqlParameter("@Style2", Me.PStyle2)
        Dim param36 As SqlParameter = New SqlParameter("@Style3", Me.PStyle3)
        Dim param37 As SqlParameter = New SqlParameter("@Style4", Me.PStyle4)
        Dim param38 As SqlParameter = New SqlParameter("@Style5", Me.PStyle5)
        Dim param39 As SqlParameter = New SqlParameter("@Color1", Me.PColor1)
        Dim param40 As SqlParameter = New SqlParameter("@Color2", Me.PColor2)
        Dim param41 As SqlParameter = New SqlParameter("@Color3", Me.PColor3)
        Dim param42 As SqlParameter = New SqlParameter("@Color4", Me.PColor4)
        Dim param43 As SqlParameter = New SqlParameter("@Color5", Me.PColor5)
        Dim param44 As SqlParameter = New SqlParameter("@Qty1", Me.Qty1)
        Dim param45 As SqlParameter = New SqlParameter("@Qty2", Me.Qty2)
        Dim param46 As SqlParameter = New SqlParameter("@Qty3", Me.Qty3)
        Dim param47 As SqlParameter = New SqlParameter("@Qty4", Me.Qty4)
        Dim param48 As SqlParameter = New SqlParameter("@Qty5", Me.Qty5)
        Dim param49 As SqlParameter = New SqlParameter("@Unit1", Me.Unit1)
        Dim param50 As SqlParameter = New SqlParameter("@Unit2", Me.Unit2)
        Dim param51 As SqlParameter = New SqlParameter("@Unit3", Me.Unit3)
        Dim param52 As SqlParameter = New SqlParameter("@Unit4", Me.Unit4)
        Dim param53 As SqlParameter = New SqlParameter("@Unit5", Me.Unit5)
        Dim param54 As SqlParameter = New SqlParameter("@Notes", Me.rtfNote.Text)
        Dim param55 As SqlParameter = New SqlParameter("@Quoted", Me.txt1Quoted.Text & "." & Me.txt2Quoted.Text)
        Dim param56 As SqlParameter = New SqlParameter("@Par", Me.txt1Par.Text & "." & Me.txt2Par.Text)
        Dim param57 As SqlParameter = New SqlParameter("@Recoverable", Me.Recoverable)
        Dim param58 As SqlParameter = New SqlParameter("@LeadHistoryId", Me.LHID)
        Dim param59 As SqlParameter = New SqlParameter("@ID", Me.ID)
        Dim param60 As SqlParameter = New SqlParameter("@description", description)
        Dim param61 As SqlParameter = New SqlParameter("@User", STATIC_VARIABLES.CurrentUser)
        Dim cmdGet As SqlCommand
        cmdGet = New SqlCommand("dbo.SalesResult", cnn)
        cmdGet.Parameters.Add(param1)
        cmdGet.Parameters.Add(param2)
        cmdGet.Parameters.Add(param3)
        cmdGet.Parameters.Add(param4)
        cmdGet.Parameters.Add(param5)
        cmdGet.Parameters.Add(param6)
        cmdGet.Parameters.Add(param7)
        cmdGet.Parameters.Add(param8)
        cmdGet.Parameters.Add(param9)
        cmdGet.Parameters.Add(param10)
        cmdGet.Parameters.Add(param11)
        cmdGet.Parameters.Add(param12)
        cmdGet.Parameters.Add(param13)
        cmdGet.Parameters.Add(param14)
        cmdGet.Parameters.Add(param15)
        cmdGet.Parameters.Add(param16)
        cmdGet.Parameters.Add(param17)
        cmdGet.Parameters.Add(param18)
        cmdGet.Parameters.Add(param19)
        cmdGet.Parameters.Add(param20)
        cmdGet.Parameters.Add(param21)
        cmdGet.Parameters.Add(param22)
        cmdGet.Parameters.Add(param23)
        cmdGet.Parameters.Add(param24)
        cmdGet.Parameters.Add(param25)
        cmdGet.Parameters.Add(param26)
        cmdGet.Parameters.Add(param27)
        cmdGet.Parameters.Add(param28)
        cmdGet.Parameters.Add(param29)
        cmdGet.Parameters.Add(param30)
        cmdGet.Parameters.Add(param31)
        cmdGet.Parameters.Add(param32)
        cmdGet.Parameters.Add(param33)
        cmdGet.Parameters.Add(param34)
        cmdGet.Parameters.Add(param35)
        cmdGet.Parameters.Add(param36)
        cmdGet.Parameters.Add(param37)
        cmdGet.Parameters.Add(param38)
        cmdGet.Parameters.Add(param39)
        cmdGet.Parameters.Add(param40)
        cmdGet.Parameters.Add(param41)
        cmdGet.Parameters.Add(param42)
        cmdGet.Parameters.Add(param43)
        cmdGet.Parameters.Add(param44)
        cmdGet.Parameters.Add(param45)
        cmdGet.Parameters.Add(param46)
        cmdGet.Parameters.Add(param47)
        cmdGet.Parameters.Add(param48)
        cmdGet.Parameters.Add(param49)
        cmdGet.Parameters.Add(param50)
        cmdGet.Parameters.Add(param51)
        cmdGet.Parameters.Add(param52)
        cmdGet.Parameters.Add(param53)
        cmdGet.Parameters.Add(param54)
        cmdGet.Parameters.Add(param55)
        cmdGet.Parameters.Add(param56)
        cmdGet.Parameters.Add(param57)
        cmdGet.Parameters.Add(param58)
        cmdGet.Parameters.Add(param59)
        cmdGet.Parameters.Add(param60)
        cmdGet.Parameters.Add(param61)
        cmdGet.CommandType = CommandType.StoredProcedure
        Dim r As SqlDataReader
        cnn.Open()
        r = cmdGet.ExecuteReader(CommandBehavior.CloseConnection)
        r.Read()
        r.Close()
        cnn.Close()
        Me.Close()
        Sales.ForceRefresh = True
        If Sales.tbMain.SelectedIndex = 1 Then
            If Sales.cboSalesList.Text <> "Unfiltered Sales Dept. List" Or (Sales.cboGroupSales.Text = "Sales Result" Or Sales.cboGroupSales.Text = "Marketing Result" Or Sales.cboGroupSales.Text = "Sales Rep") Then
                Dim c As New SalesListManager
            Else
                If Sales.TabControl2.TabIndex = 0 And Sales.lvSales.SelectedItems.Count <> 0 Then
                    Sales.lvSales_SelectedIndexChanged(Nothing, Nothing)
                ElseIf Sales.TabControl2.TabIndex = 1 And Sales.lvMemorized.SelectedItems.Count <> 0 Then
                    Sales.PopulateMemorized()
                End If
            End If
        Else
            If Sales.lvnoresults.SelectedItems(0).Index = 0 And Sales.lvnoresults.Items.Count >= 2 Then
                Sales.lvnoresults.SelectedItems(0).Remove()
                Sales.lvnoresults.Items(0).Selected = True
            ElseIf Sales.lvnoresults.SelectedItems(0).Index = Sales.lvnoresults.Items.Count - 1 And Sales.lvnoresults.Items.Count >= 2 Then
                Dim x As Integer = Sales.lvnoresults.SelectedItems(0).Index
                Sales.lvnoresults.SelectedItems(0).Remove()
                Sales.lvnoresults.Items(x - 1).Selected = True

            ElseIf Sales.lvnoresults.SelectedItems(0).Index <> 0 And Sales.lvnoresults.Items.Count >= 2 Then
                Dim x As Integer = Sales.lvnoresults.SelectedItems(0).Index
                Sales.lvnoresults.SelectedItems(0).Remove()
                Sales.lvnoresults.Items(x).Selected = True
            Else
                Sales.lvnoresults.SelectedItems(0).Remove()
            End If
            Dim z As New Sales_Performance_Report
        End If
    End Sub

    Private Sub cboRep1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboRep1.SelectedValueChanged
        If Me.cboRep1.Text = Nothing And Me.cboRep2.Text <> Nothing Then
            Me.cboRep1.Text = Me.cboRep2.Text
            Me.cboRep2.Text = Nothing
        End If
    End Sub

    Private Sub chkRecoverable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecoverable.CheckedChanged
        If Me.lblRecoveryPC.Text = "Recovery Result" And Me.chkRecoverable.Checked = True And Me.EditMode = False Then
            Dim x = MsgBox("This Record has already been through the" & vbCr & _
                    "Recovery Department. Are you sure you" _
            & vbCr & "want to send it back to Recovery again?", MsgBoxStyle.YesNo, "Recovery Result")
            If x = vbNo Then
                Me.chkRecoverable.Checked = False
            End If
        End If
    End Sub


    Private Sub SDResult_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub
End Class
