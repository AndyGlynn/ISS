'Imports cmraude.Service1
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic
Public Class Contact_Info
    Public SelDepart As String = "" '' value to carry over to add new. --> Autoselect depart in cbo
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Employee_Contacts.SelName = ""
        Employee_Contacts.SelPhone = ""
        Me.Close()
    End Sub


    Private Sub Contact_Info_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.cboDepart.Items.Add("Sales and Marketing")
        Me.cboDepart.Items.Add("Financing")
        Me.cboDepart.Items.Add("Administration")
        Me.cboDepart.Items.Add("Installation")
        Me.cboDepart.Items.Add("Office Management")
        Me.cboDepart.SelectedItem = SelDepart
    End Sub

    Private Sub btnAction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAction.Click
        Dim actn As String = ""
        actn = Me.btnAction.Text
        Select Case actn
            Case Is = "EDIT"
                Dim g As New ROLODEX_LOGIC.EditEmployee
                Dim fname As String = ""
                Dim lname As String = ""
                Dim depart As String = ""
                Dim phone As String = ""
                Dim namear()
                namear = Split(Employee_Contacts.SelName, ", ", 2)
                fname = Trim(namear(0))
                lname = Trim(namear(1))
                depart = Employee_Contacts.cboDepartment.Text
                phone = Employee_Contacts.SelPhone
                g.EditEmployees(Employee_Contacts.RecID, Me.txtFName.Text, Me.txtLName.Text, Me.maskPhone.Text)
                Me.Close()
                Employee_Contacts.SelName = ""
                Employee_Contacts.SelPhone = ""
                Employee_Contacts.RecID = ""
                Exit Select

            Case Is = "Save"
                Dim c As New ROLODEX_LOGIC.InsetEmployee
                c.InsertNewEmployee(Me.txtFName.Text, Me.txtLName.Text, Me.cboDepart.Text, Me.maskPhone.Text)
                Employee_Contacts.SelPhone = ""
                Employee_Contacts.SelName = ""
                Employee_Contacts.RecID = ""
                Me.Close()

                '        Dim x As New cmraude.getRolo
                '        Dim fname As String = ""
                '        Dim lname As String = ""
                '        Dim depart As String = ""
                '        Dim phone As String = ""
                '        fname = Capitalise.capitalize(Me.txtFName.Text)
                '        lname = Capitalise.capitalize(Me.txtLName.Text)
                '        depart = Me.cboDepart.Text
                '        phone = Me.maskPhone.Text
                '        If fname = "" Then
                '            MsgBox("You must supply a value for employee first name.", MsgBoxStyle.Exclamation, "Error Adding Employee To Rolodex")
                '            Me.Close()
                '            Exit Sub
                '        End If
                '        If lname = "" Then
                '            MsgBox("You must supply a value for employee last name.", MsgBoxStyle.Exclamation, "Error Adding Employee To Rolodex")
                '            Me.Close()
                '            Exit Sub
                '        End If
                '        If depart = "" Then
                '            MsgBox("You must supply a value for employee department.", MsgBoxStyle.Exclamation, "Error Adding Employee to Rolodex")
                '            Me.Close()
                '            Exit Sub
                '        End If
                '        If phone = "" Then
                '            MsgBox("You must supply a value for employee phone number.", MsgBoxStyle.Exclamation, "Error Adding Employee to Rolodex")
                '            Me.Close()
                '            Exit Sub
                '        End If
                '        x.AddRolo(fname, lname, depart, phone)
                '        Dim z As New cmraude.getRolo
                '        Dim g As String = ""
                '        g = Me.cboDepart.Text
                '        If g = "" Then
                '            Exit Sub
                '        End If
                '        z.GetEmpRolo(g)
                '        Dim b
                '        Employee_Contacts.lstEmployees.Items.Clear()
                '        For b = 1 To z.CntEmp
                '            Dim lvEMP As New ListViewItem
                '            lvEMP.Text = z.EmpName(b).ToString
                '            lvEMP.SubItems.Add(z.EmpPhone(b).ToString)
                '            Employee_Contacts.lstEmployees.Items.Add(lvEMP)
                '        Next
                '        Employee_Contacts.SelName = ""
                '        Employee_Contacts.SelPhone = ""
                '        Me.Close()
        End Select

    End Sub

    Private Sub cboDepart_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDepart.SelectedIndexChanged

    End Sub
End Class
