Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Imports Microsoft.Win32

Public Class FindLogic
    Public Sub Search(ByVal Str As String)
        Try
            FindLead.lstSearchResults.Items.Clear()
            Dim CNN As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
            Dim cmdGet As SqlCommand = New SqlCommand("iss.dbo.Search", CNN)
            Dim param1 As SqlParameter = New SqlParameter("@String", Str)
            cmdGet.CommandType = CommandType.StoredProcedure
            cmdGet.Parameters.Add(param1)
            CNN.Open()
            Dim R1 As SqlDataReader = cmdGet.ExecuteReader
            While R1.Read
                Dim lv As New ListViewItem
                lv.Name = R1.Item(0)
                lv.Text = R1.Item(0)
                Dim s = Split(R1.Item(1), " ")
                If R1.Item(2) = " " Then
                    lv.SubItems.Add(R1.Item(1))
                ElseIf R1.Item(2) <> " " Then
                    If InStr(R1.Item(2), s(1)) = 0 Then
                        lv.SubItems.Add(R1.Item(1) & " & " & R1.Item(2))
                    Else
                        lv.SubItems.Add(s(0) & " & " & R1.Item(2))
                    End If
                End If
                lv.SubItems.Add(R1.Item(3))
                lv.SubItems.Add(R1.Item(4))
                If R1.Item(5) = "NameMatch" Then
                    lv.Group = FindLead.lstSearchResults.Groups(0)
                ElseIf R1.Item(5) = "PhoneMatch" Then
                    lv.Group = FindLead.lstSearchResults.Groups(2)
                ElseIf R1.Item(5) = "AddressMatch" Then
                    lv.Group = FindLead.lstSearchResults.Groups(1)
                ElseIf R1.Item(5) = "IDMatch" Then
                    lv.Group = FindLead.lstSearchResults.Groups(3)
                End If
                FindLead.lstSearchResults.Items.Add(lv)
            End While
            R1.Close()
            CNN.Close()
        Catch ex As Exception

            Dim err As New ErrorLogFlatFile
            err.WriteLog("FindLogic", "ByVal Str As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "Search")
        End Try

    End Sub
End Class
