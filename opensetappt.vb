Public Class opensetappt

    Public Sub New(ByVal frm As Form, ByVal id As String, ByVal Odate As String, ByVal Otime As String, ByVal C1 As String, ByVal C2 As String)

        SetAppt.ID = id
        SetAppt.frm = frm

        SetAppt.OrigApptDate = Odate
        SetAppt.OrigApptTime = Otime
        Dim s = Split(C1, " ")
        Dim s2 = Split(C2, " ")
        SetAppt.Contact1 = s(0)
        SetAppt.Contact2 = s2(0)
        SetAppt.ShowDialog()


    End Sub
End Class
