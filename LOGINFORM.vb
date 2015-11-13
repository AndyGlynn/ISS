Imports System.Text
Imports System.Xml
Imports System
Imports System.IO


Public Class LOGIN
    Dim strPath As String = "C:\Users\Public\pref.xml"
    Public Success As Boolean = False

    Private Sub LOGIN_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Me.Dispose()
        ' Main.tmrmanagealerts.Start()
        'PastDueAlerts.Show()
     



















    End Sub

    Private Sub LOGIN_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Success = False Then
            Main.Close()
        End If
    End Sub

   

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '' to be done: user snapshotp feature trumps the local xml file
        '' edit : 8-13-2015 Aaron


        Me.txtPWD.Text = ""
        Me.Text = STATIC_VARIABLES.ProgramName.ToString & " : Login"
        Dim a As New AUTO_COMPLETE_LOGINS

        Dim doc As New XmlDocument
        doc.Load(strPath)
        Dim r1 As XmlNodeReader
        r1 = New XmlNodeReader(doc)
        While r1.Read
            Select Case r1.NodeType
                Case Is = XmlNodeType.Element
                Case Is = XmlNodeType.Text
                    Dim str As String = Nothing
                    str = r1.Value
                    Me.txtUserName.Text = str
                    Exit Select
            End Select
        End While
        r1.Close()
        doc = Nothing
        Me.Activate()
        If Me.txtUserName.Text <> "" Then

            Me.txtPWD.Select()
        Else
            Me.txtUserName.Select()
        End If

        Me.txtPWD.Text = "2527" ''Remove Later

     
        Me.btnOK_Click(Nothing, Nothing) ''Remove Later
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
        Main.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If Me.txtPWD.Text.ToString.Length < 2 Then
            MsgBox("You cannot have a blank password.", MsgBoxStyle.Exclamation, "Password Needed")
            Me.txtPWD.Select()
            Exit Sub
        End If
        If txtUserName.Text = "" Then
            MsgBox("You must supply a User Name.", MsgBoxStyle.Exclamation, "Blank Login")
            Exit Sub
        End If
        'If Me.txtUserName.Text = "Admin" Then
        Main.TSMain.Enabled = True

        'If Me.txtPWD.Text <> "1234567890!" Then
        Dim log_In As New USER_LOGIC
        log_In.Login_To_System()
        If Me.Success = False Then
            Me.txtUserName.Text = ""
            Me.txtPWD.Text = ""
            Me.txtUserName.Focus()
            'Me.Success = True
            Exit Sub
        End If
        Main.TSMain.Enabled = True
        System.IO.File.Delete("C:\Users\Public\Pref.xml")
        Dim sw As New StreamWriter("C:\Users\Public\Pref.xml")
        sw.WriteLine("<ROOT>")
        sw.WriteLine("<LastLogin>" & Me.txtUserName.Text & "</LastLogin>")
        sw.WriteLine("</ROOT>")
        sw.Flush()
        sw.Close()

        'Me.txtPWD.Text = ""
        'Dim b As New ALERT_LOGIC("Admin")
        'Dim x As New ALERT_LOGIC("Admin")
        'Select Case x.CountOfAlerts
        '    Case Is >= 1
        '        Dim response As Integer = MsgBox("You have past due alerts. Would you like to manage them now?", MsgBoxStyle.YesNo, "Past Due Alerts")
        '        If response = 6 Then
        '            ManageAlerts.Show()
        '        ElseIf response = 7 Then
        '            'Exit Sub
        '        End If
        '    Case Is <= 0
        '        Exit Select
        'End Select
        ''PastDueAlerts.Show()
        'Me.Close()
        'Me.Close()
        'End If
        ' If Me.txtPWD.Text = "1234567890!" Then
        'Main.TSMain.Enabled = True
        'System.IO.File.Delete("C:\Users\Public\Pref.xml")
        'Dim sw As New StreamWriter("C:\Users\Public\Pref.xml")
        'sw.WriteLine("<ROOT>")
        'sw.WriteLine("<LastLogin>Admin</LastLogin>")
        'sw.WriteLine("</ROOT>")
        'sw.Flush()
        'sw.Close()
        'Me.txtPWD.Text = ""
        ''Dim b As New ALERT_LOGIC("Admin")
        ''Dim x As New ALERT_LOGIC("Admin")
        ''Select Case x.CountOfAlerts
        ''    Case Is >= 1
        ''        Dim response As Integer = MsgBox("You have past due alerts. Would you like to manage them now?", MsgBoxStyle.YesNo, "Past Due Alerts")
        ''        If response = 6 Then
        ''            ManageAlerts.Show()
        ''        ElseIf response = 7 Then
        ''            Exit Sub
        ''        End If
        ''    Case Is <= 0
        ''        Exit Select
        ''End Select
        ''PastDueAlerts.Show()
        'Me.Close()
        ''Me.Close()
        ' End If

        'End If
        'If Me.txtUserName.Text <> "Admin" Then
        '    Dim log_In As New USER_LOGIC
        '    log_In.Login_To_System()
        '    Main.TSMain.Enabled = True
        '    System.IO.File.Delete("C:\Users\Public\Pref.xml")
        '    Dim sw As New StreamWriter("C:\Users\Public\Pref.xml")
        '    sw.WriteLine("<ROOT>")
        '    sw.WriteLine("<LastLogin>" & STATIC_VARIABLES.Login & "</LastLogin>")
        '    sw.WriteLine("</ROOT>")
        '    sw.Flush()
        '    sw.Close()
        '    Me.txtPWD.Text = ""
        '' now check for past due alerts
        '' 
        'Dim b As New ALERT_LOGIC(STATIC_VARIABLES.CurrentUser)
        'Dim x As New ALERT_LOGIC(STATIC_VARIABLES.CurrentUser)
        'Select Case x.CountOfAlerts
        '    Case Is >= 1
        '        Dim response As Integer = MsgBox("You have past due alerts. Would you like to manage them now?", MsgBoxStyle.YesNo, "Past Due Alerts")
        '        If response = 6 Then
        '            ManageAlerts.Show()
        '        ElseIf response = 7 Then
        '            'Exit Sub
        '        End If
        '    Case Is <= 0
        '        Exit Select
        'End Select
        ''PastDueAlerts.Show()
        Me.Close()

        'End If
        'MsgBox(STATIC_VARIABLES.CurrentUser)
    End Sub

    Private Sub lnkForgot_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkForgot.LinkClicked
        MsgBox("Contact your system administrator to reset your password!", MsgBoxStyle.Exclamation, "Forgot Password")
    End Sub
End Class
