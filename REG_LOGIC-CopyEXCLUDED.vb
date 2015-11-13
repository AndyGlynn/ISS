Imports Microsoft.VisualBasic
Imports System
Imports System.Security.Permissions
Imports Microsoft.Win32
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient


Public Class REG_LOGIC
    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Public Sub CreateKey()
        Try
            Dim TempKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            TempKey.CreateSubKey("WRLD")
            Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x1", "AAAAA-BBBBB-CCCCC-DDDDD-FFFFF-1")
            Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x2", "000-00000-001")
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("REG_LOGIC", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CreateKey")

        End Try

    End Sub
    Public Sub ReadKey(ByVal LIC_KEY As String, ByVal LEASE_KEY As String)
        'Dim TempKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\WRLD", True)
        '' Form1.lstKeys.Items.Clear()

        'Dim lv As New ListViewItem
        'lv.Text = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x1", "None")
        'lv.SubItems.Add(Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x2", "None"))
        '' Me.Form1.lstKeys.Items.Add(lv)
        'Try
        '    Dim License_KEY As String = ""
        '    Dim LeaseKEY As String = ""
        '    License_KEY = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x1", "None")
        '    LeaseKEY = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\WRLD", "0x2", "None")
        '    If License_KEY <> STATIC_VARIABLES.LicenseKey Then
        '        MsgBox("License Keys do no match. Please call customer support at 1-800-EAT-DICKS.", MsgBoxStyle.Critical, "ERROR - INVALID SOFTWARE KEY")
        '        Application.Exit()
        '    End If
        '    If LeaseKEY <> STATIC_VARIABLES.LeaseKey Then
        '        MsgBox("Lease Keys do not match. Please call customer support at 1-800-EAT-DICKS.", MsgBoxStyle.Critical, "ERROR - INVALID LICENSE KEY")
        '        Application.Exit()
        '    End If
        'Catch ex As Exception
        '    Dim err As New ErrorLogFlatFile
        '    err.WriteLog("REG_LOGIC", "ByVal LIC_KEY As String, ByVal LEASE_KEY As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "ReadKey")

        'End Try

    End Sub
    Public Sub GetKeysFromTable(ByVal UserName As String, ByVal UserPWD As String)
        Try
            Dim cmdGET As SqlCommand = New SqlCommand("dbo.GetSoftwareKeys", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@UserName", UserName)
            Dim param2 As SqlParameter = New SqlParameter("@UserPWD", UserPWD)
            cmdGET.Parameters.Add(param1)
            cmdGET.Parameters.Add(param2)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
            Dim LC_KEY As String = ""
            Dim LS_KEY As String = ""
            While r1.Read
                LC_KEY = r1.Item(0)
                LS_KEY = r1.Item(1)
            End While
            r1.Close()
            cnn.Close()
            ReadKey(LC_KEY, LS_KEY)
        Catch ex As Exception
            cnn.Close()
            Dim err As New ErrorLogFlatFile
            err.WriteLog("REG_LOGIC", "ByVal UserName As String, ByVal UserPWD As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "SQL", "GetKeysFromTable")

        End Try

    End Sub
End Class
