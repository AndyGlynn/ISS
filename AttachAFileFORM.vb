Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System

Public Class AttachAFile
    Public tt As ToolTip
    Public ep As New ErrorProvider
    Public RemoveErrP As Boolean = False

    Private Sub AttachAFile_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Private Sub AttachAFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.ep.Clear()
        Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"
        Me.txtLeadNumber.Text = STATIC_VARIABLES.CurrentID
        Dim a As New CUSTOMER_LABEL
        a.GetINFO(Me.txtLeadNumber.Text)
        Me.lblPhoneInfo.Font = New Font("Tahoma", 12.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
        Me.lblPhoneInfo.Text = a.Contact1Name & vbCrLf & a.StAddress & vbCrLf & a.HousePhone & vbCrLf & a.AltPhone1 & "     " & a.AltPhone1Type & vbCrLf & a.AltPhone2 & "     " & a.AltPhone2Type
        Me.Text = Me.Text & " for Record ID: " & Me.txtLeadNumber.Text
    End Sub

    Private Sub txtLeadNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLeadNumber.KeyPress
        tt = New ToolTip
        tt.RemoveAll()
        Select Case e.KeyChar
            Case "/"c, "\"c, "!"c, "#"c, "$"c, "%"c, "^"c, "&"c, "("c, ")"c, "?"c, "<"c, ">"c, "@"c
                tt.Show("!@#$%^&*()?<>\/  Are illegal characters to instert.", Me.txtLeadNumber, 3000)
                Me.txtLeadNumber.Text = ""
                Exit Select
        End Select
    End Sub

    Private Sub txtLeadNumber_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLeadNumber.TextChanged
        Dim str As String = ""
        str = Me.txtLeadNumber.Text
        If str.ToString.Length <= 2 Then
            'Me.lblContactInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            'Me.lblContactInfo.Text = "No Contact Information"
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"
            Exit Sub
        End If
        If str.ToString.ToString <= 0 Then
            ''Me.lblContactInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
            ''Me.lblContactInfo.Text = "No Contact Information"
            Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"
        End If
        If str.ToString.Length >= 2 Then
            'ValidateLeadNumber(str)
            'If Me.RemoveErrP = True Then
            Dim a As New CUSTOMER_LABEL
            a.GetINFO(Me.txtLeadNumber.Text)
            Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, FontStyle.Bold, GraphicsUnit.Pixel, CType(0, Byte))
            Me.lblPhoneInfo.Text = a.Contact1Name & vbCrLf & a.StAddress & vbCrLf & vbCrLf & a.HousePhone & vbCrLf & a.AltPhone1 & "     " & a.AltPhone1Type & vbCrLf & a.AltPhone2 & "     " & a.AltPhone2Type
            Me.ep.Clear()
            'End If
        End If
    End Sub
    Private Sub ValidateLeadNumber(ByVal LeadNumber As String)
        ep.BlinkRate = 0
        ep.BlinkStyle = ErrorBlinkStyle.NeverBlink
        ep.Clear()
        Dim cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Dim cmdVAL As SqlCommand = New SqlCommand("SELECT COUNT(ID) From iss.dbo.enterlead where ID = @ID", cnn)
        Dim param1 As SqlParameter = New SqlParameter("@ID", LeadNumber)
        cnn.Open()
        cmdVAL.Parameters.Add(param1)
        Dim r1 As SqlDataReader
        r1 = cmdVAL.ExecuteReader
        Try
            While r1.Read
                If r1.Item(0) <= 0 Then
                    'MsgBox("invalid lead num.")
                    ep.SetError(Me.txtLeadNumber, "Invalid Lead Number")
                    Me.lblPhoneInfo.Font = New Font("Tahoma", 10.25!, 3, GraphicsUnit.Pixel, CType(0, Byte))
                    ' Me.lblContactInfo.Text = "No Contact Information"
                    Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)"

                    Me.RemoveErrP = False
                ElseIf r1.Item(0) >= 1 Then
                    Me.RemoveErrP = True
                    Me.ep.Clear()
                End If
            End While
            r1.Close()
            cnn.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

   
    Private Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Me.ep.Clear()
        If Me.lblPhoneInfo.Text = "No Contact Information Available" & vbCrLf & "(You must supply a valid Customer ID)" Then
            Me.ep.SetError(Me.txtLeadNumber, "Invalid Customer ID")
            Exit Sub
        End If
        If Me.lblPhoneInfo.Text = "" Or Me.lblPhoneInfo.Text = " " Then
            Me.ep.SetError(Me.txtLeadNumber, "Invalid Customer ID")
            Exit Sub
        End If


        ''-----------------------------------------------------
        '' Uncomment to attach a file 
        '' commented out to target control for a forced refresh 
        '' W.I.P 8-28-15
        '' tested 8-29-15 and works as expected.
        '' immediately refreshes control to show new file 'attached'
        '' 
        Dim z As New Attach
        z.AttachFile(Me.txtLeadNumber.Text)
        ''-----------------------------------------------------





        '' EDIT:
        '' refresh sales manager af control as soon as a file is attached
        '' 8-26-15

        ''
        '' if sales is open
        ''   if <> Looking at customer history
        ''     refresh control() 
        ''   end if 
        '' end if 
        '' 

        Dim f As Windows.Forms.Form = Sales
        Dim y As Panel
        Dim g As TabControl

        g = f.Controls("tbMain")
        Dim b As TabPage = g.TabPages("tpCustomerList")
        If g.SelectedTab Is b Then
            For Each a As Control In b.Controls
                Dim c As SplitterPanel
                For Each c In a.Controls
                    Dim d As Control
                    For Each d In c.Controls
                        If d.Name = "pnlAFPics" Then
                            Dim p As Panel = d
                            Dim lv As ListView
                            For Each lv In p.Controls
                                If lv.Name = "lsAF" Then
                                    '' refresh them if they are active
                                    p.Controls.RemoveByKey("lsAF")
                                    Dim widthOfParent As Integer = Sales.pnlAFPics.Width
                                    Dim widthOfControl As Integer = (widthOfParent / 2) - 20
                                    Dim heightOfParent As Integer = Sales.pnlAFPics.Height
                                    Dim heightOfControl As Integer = (heightOfParent - 30)
                                    Dim InitPoint As System.Drawing.Point = New System.Drawing.Point((0 + 10), (0 + 10))
                                    Dim InitPoint2 As System.Drawing.Point = New System.Drawing.Point(((widthOfParent / 2) + 10), (0 + 10))

                                    Dim pt As New System.Drawing.Point(0, 0)
                                    Dim xyz As New ReusableListViewControl
                                    xyz.GenerateListControl(Sales.pnlAFPics, (STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID).ToString, pt, "lsAF", heightOfControl, widthOfControl)
                                     
                                    'ElseIf lv.Name = "lsJP" Then
                                    '    ''refresh them if they are active 
                                    '    p.Controls.RemoveByKey("lsJP")
                                    '    Dim pt2 As New System.Drawing.Point(364, 0)
                                    '    Dim xyz2 As New ReusableListViewControl(Sales.pnlAFPics, (STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID).ToString, pt2, "lsJP")
                                     
                                End If
                            Next
                        End If
                    Next
                Next
            Next
        End If

        Me.Close()




    End Sub
End Class
