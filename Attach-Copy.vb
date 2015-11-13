Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports IWshRuntimeLibrary
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text
Imports System.IO
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings


Public Class Attach2
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public Sub AttachFile(ByVal ID As Integer, ByVal Hash As String)
        Try
            Dim opfd As New Windows.Forms.OpenFileDialog
            opfd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            opfd.ShowDialog()
            Dim d As String
            Try
                d = opfd.FileName.ToString
                If d.ToString = "" Then
                    AttachAFile.txtLeadNumber.Select()
                    Exit Sub
                End If
                Dim e
                Dim cnt As Integer = 0
                Dim ch As Char
                For Each ch In d
                    If ch = "\" Then
                        cnt += 1
                    End If
                Next


                e = Split(d, "\", cnt + 1)
                Dim z
                z = e(cnt)
                Dim file As String = "\\SERVER\Company\Computer Management\ISS\Attach Files\"
                System.IO.Directory.CreateDirectory("\\SERVER\Company\Computer Management\ISS\Attach Files\" + ID.ToString)
                file = file + ID.ToString + "\" + z.ToString
                Hash = file

                Try '' need to look for duplicates first. 
                    Dim cmdAttach As SqlCommand = New SqlCommand("dbo.AttachAFile", cnn)
                    cmdAttach.CommandType = CommandType.StoredProcedure
                    Dim param1 As SqlParameter = New SqlParameter("@LeadNum", ID)
                    Dim param2 As SqlParameter = New SqlParameter("@Location", d.ToString)
                    Dim param3 As SqlParameter = New SqlParameter("@Hash", Hash)
                    cmdAttach.Parameters.Add(param1)
                    cmdAttach.Parameters.Add(param2)
                    cmdAttach.Parameters.Add(param3)
                    cnn.Open()
                    Dim r2 As SqlDataReader
                    r2 = cmdAttach.ExecuteReader(CommandBehavior.CloseConnection)
                    r2.Close()
                    cnn.Close()
                Catch ex As Exception
                    Dim errp As New ErrorLogFlatFile
                    errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
                    'MsgBox(ex.Message.ToString)
                End Try

                Dim x
                Dim icnt As Integer = 0

                For Each x In Main.ILIcons.Images
                    icnt += 1
                Next

                Try
                    System.IO.File.Move(d, "\\SERVER\Company\Computer Management\ISS\Attach Files\" + ID.ToString & "\" & z.ToString)
                Catch ex As Exception
                    Dim errp As New ErrorLogFlatFile
                    errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
                End Try
                'Dim b As New GetIcons(d.ToString)
                Dim k As Integer = InStr(z, ".")
                Dim sc As String = Microsoft.VisualBasic.Left(z, k)
                Dim shell As WshShell = New WshShellClass
                Dim shortcut As WshShortcut = shell.CreateShortcut(d.ToString & ".lnk") '' create shortcut where it was
                shortcut.TargetPath = "\\SERVER\Company\Computer Management\ISS\Attach Files\" + ID.ToString & "\" & z.ToString  '' target path of where it is after being moved.
                shortcut.WorkingDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
                shortcut.Save()
                'Main.ILIcons.Images.Add(b.MyIcon)
                'Main.ILSmall.Images.Add(b.MyIcon)
                'MainApp.lstAttachedFiles.Items.Add(z, icnt)
                'MsgBox(b.x.ToString)
                AttachAFile.txtLeadNumber.Text = ""
                AttachAFile.Close()
            Catch ex As Exception
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
                'MsgBox(ex.Message.ToString)
            End Try
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "AttachFile")
        End Try

    End Sub

    Public Sub GetFilesSM(ByVal ID As String)
        'Try
        '    Dim cmdGetFiles As SqlCommand = New SqlCommand("Select Location from iss.dbo.attachfiles where LeadNum = @ID", cnn)
        '    Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
        '    cmdGetFiles.Parameters.Add(param1)
        '    Dim r1 As SqlDataReader
        '    cnn.open()
        '    r1 = cmdGetFiles.ExecuteReader
        '    ' SalesManager.lstAttachedFiles.Items.Clear()
        '    Dim icnt As Integer = 0
        '    SalesManager.ilIcons.Images.Clear()
        '    While r1.Read
        '        Dim lv As New ListViewItem
        '        lv.Text = r1.Item(0).ToString
        '        Dim d As String = ""
        '        d = lv.Text
        '        Dim e
        '        Dim cnt As Integer = 0
        '        Dim ch As Char
        '        For Each ch In d
        '            If ch = "\" Then
        '                cnt += 1
        '            End If
        '        Next
        '        e = Split(d, "\", cnt + 1)
        '        Dim z
        '        z = e(cnt)
        '        lv.Text = z
        '        Dim shell As New WshShell
        '        Dim objShell = New WshShell
        '        Dim b As New GetIcons(d.ToString)
        '        '

        '        icnt += 1

        '        SalesManager.ilIcons.Images.Add(b.MyIcon)
        '        SalesManager.ilSmall.Images.Add(b.MyIcon)
        '        SalesManager.lstJobFiles.Items.Add(z, icnt - 1)

        '    End While

        '    r1.Close()
        '    cnn.close()

        'Catch ex As Exception
        '    MsgBox(ex.Message.ToString)
        'End Try

    End Sub
    Public Sub OpenFile(ByVal file As String, ByVal id As String)
        Try
            If file.ToString = "" Then
                Exit Sub
            End If

            If id = "" Then
                Exit Sub
            End If

            Dim cmdGetFile As SqlCommand = New SqlCommand("Select Location from iss.dbo.attachfiles where LeadNum = @ID and Location LIKE '%" & file.ToString & "'", cnn)
            Dim r1 As SqlDataReader
            Dim param1 As SqlParameter = New SqlParameter("@ID", id)
            cmdGetFile.Parameters.Add(param1)

            cnn.Open()
            r1 = cmdGetFile.ExecuteReader
            Dim loc As String = ""
            While r1.Read
                loc = r1.Item(0).ToString
            End While
            r1.Close()
            cnn.Close()
            System.Diagnostics.Process.Start(loc.ToString)
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("Attach", "UserName as string, byval HASH as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "OpenFile")
        End Try

    End Sub

End Class
