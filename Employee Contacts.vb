
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic
Public Class Employee_Contacts

    Public SelName As String = "" '' var for to hold selected name from list.
    Public SelPhone As String = "" '' var to hold phone num of selected name
    Public RecID As String = "" '' var to hold the record ID to edit on table 
    Private Sub pctClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub pctEdit_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.ToolTip1.InitialDelay = 0
        Me.ToolTip1.SetToolTip(pctEdit, "Edit Employee")
    End Sub

    Private Sub pctDelete_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.ToolTip1.InitialDelay = 0
        Me.ToolTip1.SetToolTip(pctDelete, "Delete Employee")
    End Sub

    

    
    Private Sub Employee_Contacts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.cboDepartment.Items.Add("Sales and Marketing")
        Me.cboDepartment.Items.Add("Financing")
        Me.cboDepartment.Items.Add("Administration")
        Me.cboDepartment.Items.Add("Installation")
        Me.cboDepartment.Items.Add("Office Management")
        Me.cboDepartment.SelectedItem = "Administration"
        Dim c As New ROLODEX_LOGIC.GetEmployeeByDepartment
        c.GetEMPLOYEES(Me.cboDepartment.Text)
    End Sub

    Private Sub cboDepartment_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDepartment.SelectedValueChanged
        Dim x As String = ""
        x = Me.cboDepartment.Text
        If x = "" Then
            Exit Sub
        End If
        Dim c As New ROLODEX_LOGIC.GetEmployeeByDepartment
        c.GetEMPLOYEES(x)
        'Dim Z As New cmraude.getRolo
        'Z.GetEmpRolo(x)
        'Dim b
        'Me.lstEmployees.Items.Clear()
        'For b = 1 To Z.CntEmp
        '    Dim lvEMP As New ListViewItem
        '    lvEMP.Text = Z.EmpName(b).ToString
        '    lvEMP.SubItems.Add(Z.EmpPhone(b).ToString)
        '    Me.lstEmployees.Items.Add(lvEMP)
        'Next

    End Sub

    Private Sub pctnew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctnew.Click
        Contact_Info.SelDepart = Me.cboDepartment.Text
        Contact_Info.btnAction.Text = "Save"
        Contact_Info.Show()
    End Sub

    
    Private Sub pctClose_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctClose.Click
        Me.Close()
    End Sub

    Private Sub pctDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pctDelete.Click
        If RecID = "" Then
            MsgBox("You must select an employee to Delete.", MsgBoxStyle.Exclamation, "Error Editing Employee")
            Exit Sub
        End If
        Dim response As Integer = 0
        response = MsgBox("Are you sure you wish to delete this employee from the rolodex?", MsgBoxStyle.YesNo, "Delete Employee From Rolodex")
        Select Case response
            Case Is = 6 ' delete
                Dim c As New ROLODEX_LOGIC.DeleteEmployee
                c.DelEmployee(RecID)
                Exit Select
            Case Is = 7 ' do not delete
                Exit Select
        End Select
    End Sub

    Private Sub pctEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctEdit.Click
        If SelName.ToString = "" Then
            MsgBox("You must select an employee to edit.", MsgBoxStyle.Exclamation, "Error Editing Employee")
            Exit Sub
        End If
        Contact_Info.SelDepart = Me.cboDepartment.Text
        Dim namear()
        namear = Split(SelName, ", ", 2)
        Dim fname As String = ""
        Dim lname As String = ""
        fname = namear(1)
        lname = namear(0)
        If fname = "" Then
            MsgBox("You must select an employee to edit.", MsgBoxStyle.Exclamation, "Error Editing Employee")
            Exit Sub
        End If
        If lname = "" Then
            MsgBox("You must select an employee to edit.", MsgBoxStyle.Exclamation, "Error Editing Employee")
            Exit Sub
        End If

        Contact_Info.txtFName.Text = fname.ToString
        Contact_Info.txtLName.Text = lname.ToString
        Contact_Info.maskPhone.Text = SelPhone.ToString
        Contact_Info.btnAction.Text = "EDIT"
        Contact_Info.Show()
    End Sub

    Private Sub lstEmployees_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstEmployees.MouseDown
        Dim x As ListViewItem
        Me.RecID = ""
        x = Me.lstEmployees.GetItemAt(e.X, e.Y)
        If x Is Nothing Then
            Exit Sub
        End If
        SelName = x.Text
        SelPhone = x.SubItems(1).Text
        Dim g As New ROLODEX_LOGIC.GetRecIDForEmployee
        Dim namear()
        namear = Split(SelName, ", ", 2)
        g.GetRecID(namear(1), namear(0))
        Me.RecID = g.RecID
        '' on mouse down get RecID - Write Class for it.

    End Sub
End Class
