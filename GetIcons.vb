Imports System.runtime.InteropServices
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System

Public Class GetIcons
    Private Structure SHFILEINFO
        Public hIcon As IntPtr ' icon
        Public iIcon As Integer ' icondex
        Public dwAttributes As Integer ' SFGAO_flags
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String
    End Structure
    Private Declare Ansi Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr
    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_SMALLICON = &H1
    Private Const SHGFI_LARGEICON = &H0
    Private nIndex = 0
    Public x As String
    Public MyIcon As System.Drawing.Icon '' the icon return to pipe to either image list

    Public Sub New(ByVal FileName As String)
        Try
            Dim himagelarge As IntPtr
            Dim fname As String = FileName

            Dim shinfo As SHFILEINFO
            shinfo = New SHFILEINFO
            shinfo.szDisplayName = New String(Chr(0), 260)
            shinfo.szTypeName = New String(Chr(0), (80))

            himagelarge = SHGetFileInfo(fname, shinfo.dwAttributes, shinfo, (16 * Marshal.SizeOf(shinfo)), SHGFI_ICON)
            MyIcon = System.Drawing.Icon.FromHandle(shinfo.hIcon)
            'x = shinfo.hIcon.handle
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("GetIcons", "ByVal FileName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetIcons")
        End Try



    End Sub

End Class
''
''
''Example code from small app. Basic idea of how it will all work
''together. Needs a refresh script for images, probably per local form but we will see....
''Table iss.dbo.attachfiles is in play, but no values on it.
''
'Imports System.Data
'Imports System.Data.Sql
'Imports System.Data.SqlClient
'Imports System
'Imports IWshRuntimeLibrary
'Imports System.Runtime.InteropServices


'Public Class Attach
'    Private cnn = New SqlConnection("Data Source=DELLXPS\SQLEXPRESS;Persist Security Info=True;User ID=Clay; PWD=test")
'    Public Sub AttachFile(ByVal ID As Integer)
'        Dim opfd As New OpenFileDialog
'        opfd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
'        opfd.ShowDialog()
'        Dim d As String
'        Try
'            d = opfd.FileName.ToString


'            Dim e
'            Dim cnt As Integer = 0
'            Dim ch As Char
'            For Each ch In d
'                If ch = "\" Then
'                    cnt += 1
'                End If
'            Next
'            e = Split(d, "\", cnt + 1)
'            Dim z
'            z = e(cnt)

'            Try
'                Dim cmdAttach As SqlCommand = New SqlCommand("Insert Iss.dbo.attachfiles (LeadNum, Location) values (@ID, @LOC)", cnn)
'                Dim param1 As SqlParameter = New SqlParameter("@ID", ID)
'                Dim param2 As SqlParameter = New SqlParameter("@LOC", d.ToString)
'                cmdAttach.Parameters.Add(param1)
'                cmdAttach.Parameters.Add(param2)
'                cnn.open()
'                Dim r2 As SqlDataReader
'                r2 = cmdAttach.ExecuteReader
'                r2.Close()
'                cnn.close()
'            Catch ex As Exception
'                MsgBox(ex.Message.ToString)
'            End Try

'            Dim x
'            Dim icnt As Integer = 0

'            For Each x In frmAttach.ilIcons.Images
'                icnt += 1
'            Next
'            Dim b As New GetIcons(d.ToString)
'            frmAttach.ilIcons.Images.Add(b.MyIcon)
'            frmAttach.ilSmall.Images.Add(b.MyIcon)
'            frmAttach.lstAttachedFiles.Items.Add(z, icnt)

'        Catch ex As Exception
'            MsgBox(ex.Message.ToString)

'        End Try

'    End Sub
'    Public Sub GetID()
'        Try
'            Dim cmdGetID As SqlCommand = New SqlCommand("Select ID from iss.dbo.enterlead", cnn)
'            Dim r1 As SqlDataReader '
'            cnn.open()
'            frmAttach.ComboBox1.Items.Clear()
'            r1 = cmdGetID.ExecuteReader
'            While r1.Read
'                frmAttach.ComboBox1.Items.Add(r1.Item(0).ToString)
'            End While
'            r1.Close()
'            cnn.close()

'        Catch ex As Exception
'            MsgBox(ex.Message.ToString)

'        End Try
'    End Sub
'    Public Sub GetFiles(ByVal ID As String)



'        Try
'            Dim cmdGetFiles As SqlCommand = New SqlCommand("Select Location from iss.dbo.attachfiles where LeadNum = @ID", cnn)
'            Dim param1 As SqlParameter = New SqlParameter("@ID", id)
'            cmdGetFiles.Parameters.Add(param1)
'            Dim r1 As SqlDataReader
'            cnn.open()
'            r1 = cmdGetFiles.ExecuteReader
'            frmAttach.lstAttachedFiles.Items.Clear()
'            Dim icnt As Integer = 0
'            frmAttach.ilIcons.Images.Clear()
'            While r1.Read
'                Dim lv As New ListViewItem
'                lv.Text = r1.Item(0).ToString
'                Dim d As String = ""
'                d = lv.Text
'                Dim e
'                Dim cnt As Integer = 0
'                Dim ch As Char
'                For Each ch In d
'                    If ch = "\" Then
'                        cnt += 1
'                    End If
'                Next
'                e = Split(d, "\", cnt + 1)
'                Dim z
'                z = e(cnt)
'                lv.Text = z
'                Dim shell As New WshShell
'                Dim objShell = New WshShell
'                Dim b As New GetIcons(d.ToString)
'                '

'                icnt += 1

'                frmAttach.ilIcons.Images.Add(b.MyIcon)
'                frmAttach.ilSmall.Images.Add(b.MyIcon)
'                frmAttach.lstAttachedFiles.Items.Add(z, icnt - 1)

'            End While

'            r1.Close()
'            cnn.close()

'        Catch ex As Exception
'            MsgBox(ex.Message.ToString)
'        End Try

'    End Sub
'    Public Sub OpenFile(ByVal file As String, ByVal id As String)

'        If file.ToString = "" Then
'            Exit Sub
'        End If

'        If id = "" Then
'            Exit Sub
'        End If

'        Dim cmdGetFile As SqlCommand = New SqlCommand("Select Location from iss.dbo.attachfiles where LeadNum = @ID and Location LIKE '%" & file.ToString & "'", cnn)
'        Dim r1 As SqlDataReader
'        Dim param1 As SqlParameter = New SqlParameter("@ID", id)
'        cmdGetFile.Parameters.Add(param1)

'        cnn.open()
'        r1 = cmdGetFile.ExecuteReader
'        Dim loc As String = ""
'        While r1.Read
'            loc = r1.Item(0).ToString
'        End While
'        r1.Close()
'        cnn.close()
'        System.Diagnostics.Process.Start(loc.ToString)
'    End Sub

'End Class
