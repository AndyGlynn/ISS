Imports System.Text
Imports System.Xml
Imports System
Imports System.IO


Public Class LOGIN_OLDV2
    Dim strPath As String = "C:\pref.xml"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = STATIC_VARIABLES.ProgramName.ToString & " : Login"
        Dim doc As New XmlDocument
        doc.Load(strPath)
        Dim r1 As XmlNodeReader
        r1 = New XmlNodeReader(doc)
        While r1.Read
            Select Case r1.NodeType
                Case Is = XmlNodeType.Element
                Case Is = XmlNodeType.Text
                    Dim str
                    str = Split(r1.Value, " ", 2)
                    Me.txtUserName.Text = str(0)
                    Me.txtLName.Text = str(1)
                    Exit Select
            End Select
        End While
        r1.Close()
        doc = Nothing
        Me.txtPWD.Text = ""

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim log_In As New USER_LOGIC
        log_In.Login_To_System()
        Me.Close()
        Main.TSMain.Enabled = True
        System.IO.File.Delete("C:\Pref.xml")
        Dim sw As New StreamWriter("C:\Pref.xml")
        sw.WriteLine("<ROOT>")
        sw.WriteLine("<LastLogin>" & STATIC_VARIABLES.CurrentUser & "</LastLogin>")
        sw.WriteLine("</ROOT>")
        sw.Flush()
        sw.Close()
    End Sub
End Class
