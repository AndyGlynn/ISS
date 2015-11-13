
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports System.IO
Imports System.Text

Public Class WHERE_IS_LEAD_LOGIC

    Private Cold As Boolean '' DONE
    Private Warm As Boolean '' DONE
    Private PC As Boolean '' DONE
    Private Recov As Boolean '' DONE
    Private Confirm As Boolean '' DONE
    Private MarketManager As Boolean '' DONE
    Private Sales As Boolean '' DONE
    Private Install As Boolean '' DONE
    Private FinancE As Boolean '' DONE
    Private Admin As Boolean '' DONE

    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

    Public Sub FindLead(ByVal ID As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.FindLead", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
            cmdGET.CommandType = CommandType.StoredProcedure
            cmdGET.Parameters.Add(param1)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
            'Form1.ListView1.Items.Clear()

            While r1.Read
                Me.ColdCall = r1.Item(1)
                'If Me.ColdCall = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Cold Calling"
                '    lv.SubItems.Add(ID)
                '    'Form1.ListView1.Items.Add(lv)
                'End If

                Me.WarmCall = r1.Item(2)
                'If Me.WarmCall = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Warm Calling"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.PreviousCust = r1.Item(3)
                'If Me.PreviousCust = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Previous"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Recovery = r1.Item(4)
                'If Me.Recovery = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Recovery"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Confirming = r1.Item(5)
                'If Me.Confirm = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Confirming"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.MarketingManager = r1.Item(6)
                'If Me.MarketingManager = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Marketing Manager"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Sales = r1.Item(7)
                'If Me.Sales = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Sales"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Installation = r1.Item(8)
                'If Me.Installation = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Installation"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Financing = r1.Item(9)
                'If Me.Financing = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Financing"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
                Me.Admininstration = r1.Item(10)
                'If Me.Admin = True Then
                '    Dim lv As New ListViewItem
                '    lv.Text = "Administration"
                '    lv.SubItems.Add(ID)
                '    Form1.ListView1.Items.Add(lv)
                'End If
            End While

            r1.Close()
            cnn.Close()
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("WHERE_IS_LEAD_LOGIC", "ByVal ID as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "FindLead")

        End Try

    End Sub
    Public Property Admininstration() As Boolean
        Get
            Return Admin
        End Get
        Set(ByVal value As Boolean)
            Admin = value
        End Set
    End Property
    Public Property Financing() As Boolean
        Get
            Return FinancE
        End Get
        Set(ByVal value As Boolean)
            FinancE = value
        End Set
    End Property
    Public Property Installation() As Boolean
        Get
            Return Install
        End Get
        Set(ByVal value As Boolean)
            Install = value
        End Set
    End Property
    Public Property SalesForm() As Boolean
        Get
            Return Sales
        End Get
        Set(ByVal value As Boolean)
            Sales = value
        End Set
    End Property
    Public Property MarketingManager() As Boolean
        Get
            Return MarketManager
        End Get
        Set(ByVal value As Boolean)
            MarketManager = value
        End Set
    End Property
    Public Property Confirming() As Boolean
        Get
            Return Confirm
        End Get
        Set(ByVal value As Boolean)
            Confirm = value
        End Set
    End Property
    Public Property Recovery() As Boolean
        Get
            Return Recov
        End Get
        Set(ByVal value As Boolean)
            Recov = value
        End Set
    End Property
    Public Property PreviousCust() As Boolean
        Get
            Return PC
        End Get
        Set(ByVal value As Boolean)
            PC = value
        End Set
    End Property
    Public Property WarmCall() As Boolean
        Get
            Return Warm
        End Get
        Set(ByVal value As Boolean)
            Warm = value
        End Set
    End Property
    Public Property ColdCall() As Boolean
        Get
            Return Cold
        End Get
        Set(ByVal value As Boolean)
            Cold = value
        End Set
    End Property
End Class
