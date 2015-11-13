Imports System.IO
Imports System.IO.FileOptions
Imports System.IO.FileStream
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Interaction
Imports Microsoft.VisualBasic.Strings
Imports System.IO.FileInfo


Module Import

    Public FileName As String
    Public ThereFlag As Boolean
    Public TargetPath As String = STATIC_VARIABLES.JobPicturesFileDirectory
    Public FileArray(99)
    Public OriginalNameArray(99)
    Public AddedVal
    Public ProgressVal
    Public LeadNum


    Public Sub IsItThere()

        ThereFlag = File.Exists(TargetPath & FileName)

        'MsgBox(ThereFlag)
    End Sub
    Public Sub StripGarbage()
        Try
            Dim file
            Dim iteration
            iteration = 0
            Dim oo
            For oo = 0 To 99
                FileArray(oo) = ""
                OriginalNameArray(oo) = ""
            Next
            For Each file In ImportPictures.OpenPics.FileNames
                iteration += 1
                FileArray(iteration) = FileName
                OriginalNameArray(iteration) = file
            Next
            Dim Index
            Dim piece1
            'Dim st As Integer
            'For st = 0 To iteration
            '    Dim ch As Char
            '    For Each ch In FileArray(st)
            '        If ch = "\" Then
            '            Index = FileArray(st).ToString.LastIndexOf("\")
            '            piece1 = Mid(FileArray(st), Index + 2, ((FileArray(st).ToString.Length)))


            '        End If
            '    Next
            '    FileArray(st) = piece1
            'Next
            Dim bb
            For bb = 0 To 99
                If FileArray(bb) <> "" Then

                    FileArray(bb) = FileArray(bb) & "-1"

                End If
            Next
            Dim aa
            Dim index2
            TargetPath = TargetPath & STATIC_VARIABLES.CurrentID & "\"
            If System.IO.Directory.Exists(TargetPath) = False Then
                System.IO.Directory.CreateDirectory(TargetPath)
            End If

            For aa = 0 To 99
                If FileArray(aa) <> "" Then
                    FileName = FileArray(aa)

                    Select Case ThereFlag

                        Case False
                            'MsgBox("Not There")
                            Select Case OriginalNameArray(aa)
                                Case "[Select a File]"
                                    Exit Sub
                                Case Is <> "[Select a File]"
                                    StripOffLastDigit()
                                    System.IO.File.Copy(OriginalNameArray(aa), TargetPath & FileName & ".jpg")
                            End Select
                    End Select
                End If
            Next
            TargetPath = STATIC_VARIABLES.JobPicturesFileDirectory
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("Import", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "StripGarbage")

        End Try
        ImportPictures.Dispose()

    End Sub
    Public Sub StripOffLastDigit()
        Try
            Dim ch As Char
            Dim aa
            Dim tempfilearray(99)
            Dim ThereNotThereArray(99)
            For aa = 0 To 99
                tempfilearray(aa) = ""
                ThereNotThereArray(aa) = ""
            Next
            Dim index
            For Each ch In FileName
                If ch = "-" Then
                    index = FileName.LastIndexOf("-")
                End If
            Next
            Dim local As String = Mid(FileName, index + 2, FileName.Length)
            local = CType(local, Integer)
            'MsgBox(FileName)
            Dim cf As Integer
            Dim ch2 As Char
            Dim NoDashFile As String = Mid(FileName, 1, FileName.Length - 1)

            For cf = 1 To 99
                tempfilearray(cf) = NoDashFile & CType(cf, String)
                ThereNotThereArray(cf) = "False"
            Next
            Dim rf As Integer
            For rf = 1 To 99
                If File.Exists(TargetPath.ToString & NoDashFile & rf & ".jpg") = True Then
                    'MsgBox("found 1 ")
                    ThereNotThereArray(rf) = "True"
                End If
            Next
            Dim x
            For x = 1 To 99
                If ThereNotThereArray(x) <> "True" Then
                    'MsgBox(tempfilearray(x))
                    FileName = tempfilearray(x)
                    Exit For
                End If
            Next
            'MsgBox(FileName)
        Catch ex As Exception
            Dim err As New ErrorLogFlatFile
            err.WriteLog("Import", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "File_IO", "StripOffLastDigit")

        End Try

    End Sub
End Module
