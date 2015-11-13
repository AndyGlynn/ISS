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


Public Class Attach
    Dim path As String = STATIC_VARIABLES.AttachedFilesDirectory
    Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
    Public Sub AttachFile(ByVal ID As Integer)
        Try
            Dim opfd As New Windows.Forms.OpenFileDialog
            opfd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            opfd.Multiselect = True
            opfd.ShowDialog()

            Dim str As String = ""
            Dim arNames As New ArrayList
            For Each str In opfd.FileNames
                arNames.Add(str)
            Next

            Dim cnt As Integer = 0
            For Each y As String In arNames
                cnt += 1
            Next
            Dim i As Integer = 0
            For i = 0 To arNames.Count - 1
                Dim sourcePath As String = arNames(i)
                Dim fName As String = StripOffFileName(arNames(i))
                Dim targetDirectory As String = (STATIC_VARIABLES.AttachedFilesDirectory & ID & "\")
                Dim Exist As Boolean = DoesDirectoryExist(targetDirectory)
                If Exist = False Then
                    System.IO.Directory.CreateDirectory(targetDirectory)
                End If
                Dim cntDupes As Integer = CountDuplicates(ID, sourcePath, (targetDirectory & fName))

                If cntDupes <= 0 Then
                    CreateAShortCut(fName, arNames(i), targetDirectory)
                    System.IO.File.Move(sourcePath, (targetDirectory & fName))
                    LogToSQLFileMove(ID, sourcePath, (targetDirectory & fName))
                End If
            Next

        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")

        End Try






        '' Obsolete CODE:
        '' 9-4-2015
        '' New Edits are supposed to compenstate for 
        '' multiple files selected
        '' 
        '' 
        '    Dim d As String
        '    Try
        '        d = opfd.FileName.ToString

        '        If d.ToString = "" Then
        '            AttachAFile.txtLeadNumber.Select()
        '            Exit Sub
        '        End If
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
        '        Dim sfp As String = Replace(d.ToString, "\" & z, "")

        '        Dim file As String = path
        '        System.IO.Directory.CreateDirectory(path + ID.ToString)
        '        file = file + ID.ToString + "\" + z.ToString
        '        Hash = file

        '        Try '' need to look for duplicates first. 
        '            Dim cmdAttach As SqlCommand = New SqlCommand("dbo.AttachAFile", cnn)
        '            cmdAttach.CommandType = CommandType.StoredProcedure
        '            Dim param1 As SqlParameter = New SqlParameter("@LeadNum", ID)
        '            Dim param2 As SqlParameter = New SqlParameter("@Location", d.ToString)
        '            Dim param3 As SqlParameter = New SqlParameter("@Hash", Hash)
        '            cmdAttach.Parameters.Add(param1)
        '            cmdAttach.Parameters.Add(param2)
        '            cmdAttach.Parameters.Add(param3)
        '            cnn.Open()
        '            Dim r2 As SqlDataReader
        '            r2 = cmdAttach.ExecuteReader(CommandBehavior.CloseConnection)
        '            r2.Close()
        '            cnn.Close()
        '        Catch ex As Exception
        '            Dim errp As New ErrorLogFlatFile
        '            errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
        '            'MsgBox(ex.Message.ToString)
        '        End Try

        '        Dim x
        '        Dim icnt As Integer = 0

        '        For Each x In Main.ILIcons.Images
        '            icnt += 1
        '        Next

        '        Try
        '            System.IO.File.Copy(d, path + ID.ToString & "\" & z.ToString)
        '        Catch ex As Exception
        '            Dim errp As New ErrorLogFlatFile
        '            errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
        '        End Try
        '        'Dim b As New GetIcons(d.ToString)
        '        'Dim k As Integer = InStr(z, ".")
        '        'Dim sc As String = Microsoft.VisualBasic.Left(z, k)
        '        'Dim shell As WshShell = New WshShellClass
        '        'Dim shortcut As WshShortcut = shell.CreateShortcut(d.ToString & ".lnk") '' create shortcut where it was
        '        'shortcut.TargetPath = path + ID.ToString & "\" & z.ToString  '' target path of where it is after being moved.
        '        'shortcut.WorkingDirectory = sfp
        '        ''My.Computer.FileSystem.SpecialDirectories.Desktop
        '        'shortcut.Save()
        '        'Main.ILIcons.Images.Add(b.MyIcon)
        '        'Main.ILSmall.Images.Add(b.MyIcon)
        '        'MainApp.lstAttachedFiles.Items.Add(z, icnt)
        '        'MsgBox(b.x.ToString)
        '        AttachAFile.txtLeadNumber.Text = ""
        '        AttachAFile.Close()
        '    Catch ex As Exception
        '        Dim errp As New ErrorLogFlatFile
        '        errp.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "AttachFile")
        '        'MsgBox(ex.Message.ToString)
        '    End Try
        'Catch ex As Exception
        '    Dim err As New ErrorLogFlatFile
        '    err.WriteLog("Attach", "ByVal ID As Integer, ByVal Hash As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "AttachFile")
        'End Try

    End Sub

    Private Sub LogToSQLFileMove(ByVal ID As String, ByVal SourcePath As String, ByVal ServerDestination As String)
        Dim cnx_MOVE As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        cnx_MOVE.Open()
        Dim cmdINS As SqlCommand = New SqlCommand("INSERT iss.dbo.AttachFiles (LeadNum,Location,AttachedHash) values('" & ID & "','" & SourcePath & "','" & ServerDestination & "');", cnx_MOVE)
        cmdINS.ExecuteScalar()
        cnx_MOVE.Close()
        cnx_MOVE = Nothing
    End Sub

    Private Function StripOffFileName(ByVal ItemText As String)
        Dim arName = Split(ItemText, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
        Dim cnt As Integer = 0
        For Each x As String In arName
            cnt += 1
        Next
        Dim f_name As String = arName(cnt - 1)
        Return f_name
    End Function

    Private Function CountDuplicates(ByVal ID As String, ByVal SourcePath As String, ByVal ServerPath As String)
        Dim retCNT As Integer = 0
        Dim cmd_DUP As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        cmd_DUP.Open()
        Dim cmdCNT As SqlCommand = New SqlCommand("SELECT COUNT(ID) from AttachFiles where LeadNum = '" & ID & "' and Location ='" & SourcePath & "' and AttachedHash = '" & ServerPath & "';", cmd_DUP)
        retCNT = cmdCNT.ExecuteScalar
        cmd_DUP.Close()
        cmd_DUP = Nothing
        Return retCNT
    End Function

    Private Function DoesDirectoryExist(ByVal Path As String)
        Dim Exist As Boolean
        If System.IO.Directory.Exists(Path) = True Then
            Exist = True
        ElseIf System.IO.Directory.Exists(Path) = False Then
            Exist = False
        End If
    End Function

    Private Sub CreateDirectory(ByVal ID As String, ByVal RootLocation As String)
        Dim path As String = (RootLocation & "\" & ID & "\")
        System.IO.Directory.CreateDirectory(path)
    End Sub

    Public Sub CreateAShortCut(ByVal ShortCutName As String, ByVal SourceDestination As String, ByVal WhereToCreate As String)
        Try
            Dim strPath As String = WhereToCreate
            Dim fullPath As String = (strPath)
            Dim names = Split(SourceDestination, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim cntNames As Integer = 0
            For Each yy As String In names
                cntNames += 1
            Next
            Dim reconstructed As String = ""
            Dim i As Integer
            For i = 0 To cntNames - 2
                If i = 0 Then
                    reconstructed = names(i)
                ElseIf i > 0 Then
                    reconstructed = (reconstructed & "\" & names(i))
                End If

            Next


            ''
            '' coming out \C:\

            '' and filename is 'test.txt.lnk'
            '' 
            Dim fnames = Split(ShortCutName, ".", -1, Microsoft.VisualBasic.CompareMethod.Text)
            Dim justFNamea As String = fnames(0)

            Dim shl As New WshShell
            Dim linkPath As IWshShortcut = shl.CreateShortcut(reconstructed & "\" & justFNamea & ".lnk")
            linkPath.TargetPath = WhereToCreate & ShortCutName
            linkPath.Save()
        Catch ex As Exception
            '' call to error log from here on fail
            '' 
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
