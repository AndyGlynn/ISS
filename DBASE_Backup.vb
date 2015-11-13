Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System

Public Class DBASE_Backup
    Public Sub Backup()
        Try
            If System.IO.File.Exists("\\ekg1\iss\backups\iss.dat") = True Then
                Dim mnth As String = Date.Today.Month
                Dim day As String = Date.Today.Day
                Dim yr As String = Date.Today.Year

                Dim hr As String = Date.UtcNow.Hour
                Dim minute As String = Date.UtcNow.Minute
                Dim sec As String = Date.UtcNow.Second



                'Dim minute As String = Date.Today.TimeOfDay.Minutes
                'Dim sec As String = Date.Today.TimeOfDay.Seconds

                System.IO.File.Move("\\ekg1\iss\backups\iss.dat", "\\ekg1\iss\backups\archives\iss " & mnth & "-" & day & "-" & yr & " " & hr & minute & sec & ".dat")

            End If

            Shell("sqlcmd -S EKG1\SQLEXPRESS -U Clay -P spoken1 -i " & Chr(34) & "\\ekg1\iss\Backups\backupscript.txt" & Chr(34) & "", AppWinStyle.NormalNoFocus, True, -1)
        Catch ex As Exception
            MsgBox("There was an error backing up the Data Base. Please contact your Administrator and try again.", MsgBoxStyle.Exclamation, "Error Backing Up Data Base")
            Dim err As New ErrorLogFlatFile
            err.WriteLog("DBASE_Backup", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "Backup")

            Exit Sub
        End Try
    End Sub
    Public Sub Get_Existing()
        Try
            frmBackup.ListView1.Items.Clear()
            For Each entry As String In System.IO.Directory.GetFiles("\\EKG1\Iss\Backups\Archives\")
                Dim file As New IO.FileInfo(entry)
                Dim lv As New ListViewItem
                lv.Text = file.FullName
                frmBackup.ListView1.Items.Add(lv)
            Next
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("DBASE_Backup", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "Get_Existing")
        End Try
    End Sub
End Class
