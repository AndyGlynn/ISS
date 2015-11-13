Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System


Public Class ErrorLogFlatFile
#Region "Variables"
    '' Directories
    '' 
    Public RootDirectory As String = "\\SERVER\Company\Computer Management\ISS Log Files"
    Public ClientDirectory As String = "C:\Users\Public\ISS Logs" ' Maybe 'c:\program files\iss\logs'

    '' Streamwriters : Server
    '' 







#End Region
    '' client
    ''
    '' client directories: 'C:\Iss\Logs'
    ''
    '' Server
    ''
    '' Server Directories '\\EKG1\iss\Logs'
    ''

    '' Structure:
    ''
    '' Module or Class Name | Method ; Property ; Function | Arguments | System Exception | Who\What Called it? 
    '' 

    '' Types:
    '' Client-
    '' Network | SQL | Front End | TAPI | Reporting Svcs | Mappoint | File.IO 
    ''

    '' Server-
    '' Network Clients | SQL | Reporting Svcs | XFER | Chat | Commands | Service Object Errors 
    '' 

    Public Sub WriteLog(ByVal Module_Or_Class As String, ByVal Arguments As String, ByVal Exception As String, ByVal Client_Or_Server As String, ByVal Who_Called As String, ByVal Type As String, ByVal Method_Property_Or_Function_Name As String)
        '' timestamp
        ''
        Dim DT As Date = Date.Now

        Select Case Client_Or_Server
            Case Is = "Client"
                Select Case Type
                    Case Is = "Network"
                        Dim Network_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\Network.txt", True)
                        'Dim network_Err As New StreamWriter("C:\Iss\Logs\Network.txt", True)
                        Network_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Network_Err_Client.Flush()
                        Network_Err_Client.Close()
                        Exit Sub
                    Case Is = "SQL"
                        Dim SQL_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\SQL.txt", True)
                        SQL_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        SQL_Err_Client.Flush()
                        SQL_Err_Client.Close()
                        Exit Select
                    Case Is = "Front_End"
                        Dim Front_End_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\Front_End.txt", True)
                        Front_End_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Front_End_Err_Client.Flush()
                        Front_End_Err_Client.Close()
                        Exit Sub
                    Case Is = "TAPI"
                        Dim TAPI_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\Tapi_Err.txt", True)
                        TAPI_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        TAPI_Err_Client.Flush()
                        TAPI_Err_Client.Close()
                        Exit Select
                    Case Is = "ReportingSVCS"
                        Dim Reporting_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\ReportingSVCS.txt", True)
                        Reporting_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Reporting_Err_Client.Flush()
                        Reporting_Err_Client.Close()
                        Exit Select
                    Case Is = "Mappoint"
                        Dim Mappoint_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\Mappoint_Err.txt", True)
                        Mappoint_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Mappoint_Err_Client.Flush()
                        Mappoint_Err_Client.Close()
                        Exit Select
                    Case Is = "File_IO"
                        Dim FileIO_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\File_IO_Err.txt", True)
                        FileIO_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        FileIO_Err_Client.Flush()
                        FileIO_Err_Client.Close()
                        Exit Select
                    Case Is = "XFER"
                        Dim XFER_Err_Client As New StreamWriter("C:\Users\Public\ISS Logs\XFER_Err.txt")
                        XFER_Err_Client.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        XFER_Err_Client.Flush()
                        XFER_Err_Client.Close()
                        Exit Select
                End Select
            Case Is = "Server"
                Select Case Type
                    Case Is = "Network"
                        Dim Network_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\Network.txt")
                        Network_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Network_Err_Server.Flush()
                        Network_Err_Server.Close()
                        Exit Select
                    Case Is = "SQL"
                        Dim SQL_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\SQL.txt", True)
                        SQL_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        SQL_Err_Server.Flush()
                        SQL_Err_Server.Close()
                        Exit Select
                    Case Is = "ReportingSVCS"
                        Dim Reporting_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\ReportingSVCS.txt", True)
                        Reporting_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Reporting_Err_Server.Flush()
                        Reporting_Err_Server.Close()
                        Exit Select
                    Case Is = "XFER"
                        Dim XFER_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\XFER_Err.txt", True)
                        XFER_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        XFER_Err_Server.Flush()
                        XFER_Err_Server.Close()
                        Exit Select
                    Case Is = "Chat"
                        Dim Chat_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\Chat_Err.txt", True)
                        Chat_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        Chat_Err_Server.Flush()
                        Chat_Err_Server.Close()
                        Exit Select
                    Case Is = "General_Client"
                        Dim General_Client_Err_Server As New StreamWriter("\\SERVER\Company\Computer Management\ISS Log Files\Client_Err.txt", True)
                        General_Client_Err_Server.Write(DT.ToString & " | " & Module_Or_Class & " | " & Method_Property_Or_Function_Name.ToString & " | " & Exception & " | " & Who_Called & " |" & vbCrLf)
                        General_Client_Err_Server.Flush()
                        General_Client_Err_Server.Close()
                        Exit Select
                End Select
        End Select
    End Sub
    Public Sub ClearLogs(ByVal Client_Or_Server As String)
        Dim response As Integer = MsgBox("WARNING! This will erase ALL logs. Are you sure you want to proceed?", MsgBoxStyle.YesNo, "WARNING!")

        Select Case response
            Case 6 ' yes
                Select Case Client_Or_Server
                    Case Is = "Client"
                        If File.Exists("C:\Users\Public\ISS Logs\Network.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Network.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\SQL.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\SQL.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\ReportingSVCS.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\ReportingSVCS.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\XFER_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\XFER_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Front_End.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Front_End.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Tapi_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Tapi_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Mappoint_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Mappoint_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\File_IO_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\File_IO_Err.txt")
                        End If
                        Exit Select

                    Case Is = "Server"
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Network.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Network.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\SQL.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\SQL.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\ReportingSVCS.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\ReportingSVCS.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\XFER_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\XFER_Err.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Chat_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Chat_Err.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Client_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Client_Err.txt")
                        End If
                        Exit Select
                    Case Is = "Both"
                        '' client side
                        '' 
                        If File.Exists("C:\Users\Public\ISS Logs\Network.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Network.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\SQL.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\SQL.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\ReportingSVCS.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\ReportingSVCS.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\XFER_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\XFER_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Front_End.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Front_End.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Tapi_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Tapi_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\Mappoint_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\Mappoint_Err.txt")
                        End If

                        If File.Exists("C:\Users\Public\ISS Logs\File_IO_Err.txt") = True Then
                            System.IO.File.Delete("C:\Users\Public\ISS Logs\File_IO_Err.txt")
                        End If

                        '' server side
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Network.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Network.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\SQL.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\SQL.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\ReportingSVCS.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\ReportingSVCS.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\XFER_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\XFER_Err.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Chat_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Chat_Err.txt")
                        End If
                        If File.Exists("\\SERVER\Company\Computer Management\ISS Log Files\Client_Err.txt") = True Then
                            System.IO.File.Delete("\\SERVER\Company\Computer Management\ISS Log Files\Client_Err.txt")
                        End If
                        Exit Select
                End Select
            Case 7 ' no
                Exit Sub
                Exit Select
        End Select
    End Sub
    Public Sub Get_Logs(ByVal Client_Or_Server As String)
        Select Case Client_Or_Server
            Case Is = "Client"
                System.Diagnostics.Process.Start("C:\Users\Public\ISS Logs")
                Exit Select
            Case Is = "Server"
                System.Diagnostics.Process.Start("\\SERVER\Company\Computer Management\ISS Log Files")
                Exit Select
            Case Is = "Both"
                System.Diagnostics.Process.Start("C:\Users\Public\ISS Logs")
                System.Diagnostics.Process.Start("\\SERVER\Company\Computer Management\ISS Log Files")
                Exit Select
        End Select
    End Sub
End Class
