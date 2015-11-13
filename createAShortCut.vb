Imports IWshRuntimeLibrary

Public Class createAShortCut

    '' anatomy of a shortcut: (desktop)
    '' 1: add reference to Windows Script Host Object Model under "COM" references
    '' 2: imports IWshRuntimelibrary
    '' 3: Modified Snippet for example --
    '' -- Basic Usage
    ''
    ''       Dim x As New WshShell
    ''       Dim link As IWshShortcut = x.CreateShortcut("C:\Users\Clay\Desktop" & "\test.lnk")
    ''       link.TargetPath = "C:\Users\Clay\Desktop\Test Directory\Iss\10000\Attached Files\this file.txt"
    ''       link.Save()
    ''
    ''
    '' source: http://www.codeproject.com/Articles/3905/Creating-Shell-Links-Shortcuts-in-NET-Programs-Usi
    '' 
    ''
    '' this is going to expand a little because we wont always know the user logged in. 
    '' so we need to gather that then create the shortcut.
    '' 


    Public Sub New(ByVal ShortCutName As String, ByVal SourceDestination As String)

        Try
            Dim strPath As String = "C:\Users\"
            Dim usr As String = Environment.UserName.ToString
            Dim strPath2 As String = "\Desktop"

            Dim fullPath As String = (strPath & usr & strPath2)

            Dim shl As New WshShell
            Dim linkPath As IWshShortcut = shl.CreateShortcut(fullPath & "\" & ShortCutName)
            linkPath.TargetPath = SourceDestination
            linkPath.Save()
        Catch ex As Exception
            '' call to error log from here on fail
            '' 
        End Try
    End Sub
End Class
