Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions '' regex
Imports System.IO.File
Imports System.IO.Directory
Imports System.Security.AccessControl '' needed to check and set directory permissions MAYBE
Imports System.Collections


Imports System.Windows.Forms '' for context menus MAYBE




Public Class jbPicsAttachedFiles
    Private _InitDirExists As Boolean
    Private ProfilePath As String = My.Computer.FileSystem.SpecialDirectories.Desktop  '' static vars change
    Private _InitDirPath As String = STATIC_VARIABLES.AttachedFilesDirectory  '' static vars change
    Private _currentLeadNum As String = STATIC_VARIABLES.CurrentID  '' static vars change
    Private _Form As Form = STATIC_VARIABLES.ActiveChild

    Private _ArFileNameOnly As ArrayList
    Private _FileNameOnly As String = ""
    Private _SortedFileNames As ArrayList

    '' vars for ccp operations - Attached Files
    '' 
    Public _arCopiedAttachedFiles As New ArrayList
    Public _arCopiedOperation As New ArrayList
    Private FullFillPathConverted As String = ""

#Region "Design / Theory / Notes"
    '' this class is all theory and a first draft
    ''
    ''
    '' Use on Front End Example to pull files out.
    ''--------------------------
    '' Dim a As New jbPicsAttachedFiles
    ''    g = a
    ''    Dim cnt As Integer = g.ArFileNameOnly.Count
    ''    For b = 0 To cnt - 1
    ''        Me.ListView1.Items.Add(g.ArFileNameOnly(b).ToString)
    ''    Next
    ''
    ''
    '' Items Exposed By Class  | Props = 1 | Methods = 2 | Functions = 3 | Objects = 4 
    ''
    '' 1: Array Of File Names Attached to 'leadnum'
    '' 2: Method to launch the selected File
    '' 2: Method to GetAttachedFiles(LeadNum)
    '' 2: Delete File method
    '' 1: Array Of Sorted File Names
    '' 2: Method to Sort Files (ASC | DSC)
    '' 2: Method to create a subdirectory within 'attached files'
    '' 


    '' made a mistake on the first draft, context menu contains 'Delete' and removes 'Rename'. Will leave Sub there just because. 
    '' also: cpu crash power outage about 110pm 11-24-14 - hope the saved project is the same as I as working on. 


    '' 
    '' 
    '' front end needs logic to determine which view its in current to know how to 'repopulate data' after it has been sorted. 
    '' list works fine
    '' details breaks due to column headers being cleared out. *Fixed* Add a column "File Name" after items cleared before items populated.
    ''
    '' 
    '' regex expression ignores case for pattern matching "[^a-z]\D" = any character NOT in a-z and anything that is NOT a decimal digit "0-9"
    '' 
    ''
    '' 11-25-2014 = Cut Copy Paste Operations:
    ''
    '' Cut is nothing more then a removal on the front end and placing data on the clipboard
    '' Paste is getting data from the clipboard and converting it to whatever control (text)
    '' Going to have paste 'target' the last item added to the array at all times
    '' going to put in a manual 'clear call' MSDN states that the clipboard items don't clear until application exits - could be lots of memory usage 
    '' ===Doubtful if I only target file paths and not the atual object itself.===
    '' this version will not have a 'paste all' function 
    '' will have to expose an array of 'copied items' from the class but will only paste from last index in array at all times.
    '' 
    '' Cut-
    '' Front End: Get selected List View Item -> Copy Text to Clipboard (full file path) -> Remove From list
    '' Back End: Copy File -> Place in Memory -> Don't delete until a "Paste Operation" is called.
    ''
    '' Copy-
    '' Front End: Just Copy the File -> Place 'Copy' on clipboard
    '' Back End: IF a paste operation is called -> Make a duplicate File with a "-Copy Of" prefix or suffix.
    '' 
    '' Paste-
    '' Front End: add item to respective list view -> give it an appropriate image key/index -> Refresh control
    '' Back End: Depends on where it was called from. If it was called from a 'cut' operation will have to delete source and place in target. 
    ''           if it was called from a copy, just replicate file. 
    ''
    '' cross control pasting ? 
    '' multiple arrays for ccp operations to keep data seperate. 
    '' 
    '' rules:
    '' job pictures can be attached files -> not all attached files can be a job picture
    '' scheduled task cannot be a job picture -> can be an attached file
    '' OR just isolate context menus per control and skip the cross calls. 
    '' 
    '' Will isolate for the initial building and see how it goes. 
    '' 
    ''
    ''
    '' image lists- 11-25-2014
    '' 
    '' the question is to create the image lists and expose them in the class
    '' or just have the loop on the front end. 
    '' if the class does the work all you have to do is point to the exposed image lists and assign items by 'keys' instead of index. 

    '' no cross threading calls flagged and the file icons change per associated exe per computer. 
    '' 
    '' should isolate it in its own class. 
    '' probably going to be re-used throughout the program. 
    '' 
    '' class will have 4 imgLists exposed (16,16 | 32,32 | 64, 64 | 128, 128 ) 
    '' 
    '' arguments will be [directory] && the keys will be assigned by file extension NO PERIOD just suffix. exe, txt, htm etc. . . 
    '' 



#End Region
#Region "Constructor"
    Public Sub New()
        ''
        '' upon being constructed
        '' it will take the lead number and search the directories for attached files if they exist
        '' 
        '' an array of file names is exposed, upon use of class, will have to loop from front end to populate list view controls
        '' 



        DetermineInititalDirectory()
        If Me._InitDirExists = True Then
            GetAttachedFiles(_currentLeadNum)
        End If


    End Sub
#End Region
#Region "Public Subs/Functions"

    Public Function SplitOffName(ByVal FileName As String)
        '' basically just get an array of split file names by \ delimeter
        '' return the array to feed the property
        Dim xx
        Dim pos As Integer
        Dim str1 = Split(FileName, "\")
        For Each xx In str1
            pos += 1
        Next
        Dim x As Integer = pos
        _FileNameOnly = str1(x - 1)
        Return _FileNameOnly
    End Function
    Public Sub MoveCopyFile(ByVal FullPath As String, ByVal Operation As String)
        Select Case Operation
            Case "Cut"
                Try
                    Dim fNFO As New FileInfo(FullPath)
                    Dim arFileName = Split(FullPath, "\")
                    '' last 2 of array
                    '' 

                    Dim FullFile = arFileName(arFileName.Count - 1)
                    Dim subSplit = Split(FullFile, ".")
                    Dim FileName As String = subSplit(0)
                    Dim FileExt As String = subSplit(1)
                    '' also need to make sure file doesn't already exist
                    '' 
                    fNFO.CopyTo(_InitDirPath & _currentLeadNum & "\" & FileName & "." & FileExt)
                    '' once file is copied now make a call on the front end to refresh list of files
                    '' 
                Catch ex As Exception
                    '' call to log file for error 
                End Try
                Exit Select
            Case "Copy"
                Try
                    Dim fNFO As New FileInfo(FullPath)
                    Dim arFileName = Split(FullPath, "\")
                    '' last 2 of array
                    '' 

                    Dim FullFile = arFileName(arFileName.Count - 1)
                    Dim subSplit = Split(FullFile, ".")
                    Dim FileName As String = subSplit(0)
                    Dim FileExt As String = subSplit(1)
                    '' also need to make sure file doesn't already exist
                    '' 
                    fNFO.CopyTo(_InitDirPath & _currentLeadNum & "\Copy Of-" & FileName & "." & FileExt)
                    '' once file is copied now make a call on the front end to refresh list of files
                    '' 
                Catch ex As Exception
                    '' call to log file for error 
                End Try
                Exit Select
        End Select

    End Sub

    Public Sub LaunchAttachedFile(ByVal leadNum As String, ByVal FileName As String)
        Try
            '' mine: C:\Users\Clay\Desktop\Test Directory\Iss\10000\Attached Files
            '' will have to change from static vars
            leadNum = _currentLeadNum '' wont have to use once static vars is in play 
            System.Diagnostics.Process.Start(_InitDirPath & leadNum & "\" & FileName.ToString)
        Catch ex As Exception
            '' call to log in error log on fail
        End Try
    End Sub
    Public Sub RenameFile(ByVal OldFileName As String, ByVal NewFileName As String, ByVal FileExt As String)
        Try
            File.Copy(_InitDirPath & _currentLeadNum & "\" & OldFileName & "." & FileExt, _InitDirPath & _currentLeadNum & "\" & NewFileName & "." & FileExt, True)
            File.Delete(_InitDirPath & _currentLeadNum & "\" & OldFileName & "." & FileExt)
        Catch ex As Exception
            '' error log call on fail 
            '' 

            MsgBox(ex.InnerException.ToString)
        End Try

    End Sub

    Public Sub GetAttachedFiles(ByVal LeadNum As String)
        Try
            '' generate new instance of an array list to use
            _ArFileNameOnly = New ArrayList
            '' recycle path names || will be from static vars
            Dim path As String = _InitDirPath & LeadNum

            Dim g As New DirectoryInfo(path) '' object to store directory
            Dim x As FileInfo '' object to store file
            For Each x In g.GetFiles("*.*") '' use method to feed array 
                _ArFileNameOnly.Add(SplitOffName(x.FullName).ToString) '' actual feeding of the array 
            Next
        Catch ex As Exception
            '' error log call on fail
        End Try

    End Sub

    Public Function SortFiles(ByVal Files As ArrayList, ByVal Order As String)
        '' default call is ASC order or A-Z  ||  1-9
        ''
        _SortedFileNames = Files
        Select Case Order
            Case "Ascending"
                _SortedFileNames.Sort()
                Exit Select
            Case "Descending"
                _SortedFileNames.Reverse()
                Exit Select
        End Select
        Return _SortedFileNames
    End Function
    Public Sub DeleteFile(ByVal FileName As String)
        '' first, get file name
        '' check if exsits
        '' if it does double down and ask user if they really want to delete it before making a copy or moving or whatever. 
        '' 
        Try
            Dim path As String = _InitDirPath & _currentLeadNum & "\" & FileName
            Dim b As New FileInfo(path)
            If b.Exists = True Then
                '' then carryout the double confirmation of deleting said file. 
                '' 
                Dim YesNo = MsgBox("Are you sure you want to delete the file: " & FileName & " ?", MsgBoxStyle.YesNoCancel, "Delete File?")
                Select Case YesNo
                    Case 5 '' no option 
                        '' do nothing and ...
                        '' 
                        Exit Select
                    Case 6 '' yes option 
                        '' actual place to delete the file
                        File.Delete(path)
                        Exit Select
                    Case Else '' generic catch for the rest 
                        '' just exit 
                        Exit Select
                End Select
            ElseIf b.Exists = False Then
                MsgBox("File: " & path & " does not exist. Please check the file / file name and try again.", MsgBoxStyle.Information, "Error Finding File")
                '' make a call here to log file as well to log that the file wasnt found and the sub exited prematurely
                '' 
                Exit Sub
            End If
        Catch ex As Exception
            '' call to log file for error reported.
            ''
        End Try


    End Sub

    Public Sub CreateDirectoryAttachedFiles(ByVal PathName As String, ByVal DirName As String)
        Try
            ''
            '' check to see if the directory exists
            '' if doesn't create
            '' 
            '' check to make sure the name of the directory name isn't illegal or empty
            '' ?? Front End ?? 
            '' 
            '' 
            Dim lngth As Integer = Len(DirName)
            If lngth <= 0 Then
                MsgBox("You cannot create a blank folder name. Check the name and please try again. ", MsgBoxStyle.Critical, "Error Creating Folder")
                Exit Sub
            ElseIf lngth > 0 Then
                Dim path As String = PathName
                '' regex for special chars
                '' 
                'Dim Pattern As New Regex("[^a-z]\D", RegexOptions.IgnoreCase) ASK CLAY WTF?
                'If Pattern.IsMatch(path, 0) = True Then
                '    MsgBox("You cannot use special characters in your folder names. Please rename and try again.", MsgBoxStyle.Critical, "Special Characters Found In Name")
                '    Exit Sub
                'If Pattern.IsMatch(path, 0) = False Then
                Dim b As New DirectoryInfo(PathName & "\" & DirName)
                If b.Exists = True Then
                    MsgBox("There is already a folder named: " & vbCrLf & path & vbCrLf & "Please select a different name and try again.", MsgBoxStyle.Information, "Directory Already Exists")
                    Exit Sub
                ElseIf b.Exists = False Then
                    CreateDirectory(path & DirName)
                End If
            End If
            'End If
        Catch ex As Exception
            '' catch the error and log to file if this fails. 
            '' 
        End Try
    End Sub

#End Region
#Region "Cut Copy Paste Operations"
    Public Function ConvertFilePath(ByVal FileName As String, ByVal FileExt As String)
        FullFillPathConverted = (_InitDirPath & _currentLeadNum & "\" & FileName & "." & FileExt)
        Return FullFillPathConverted
    End Function
#End Region
#Region "Properties"
    Public Property ArFileNameOnly() As ArrayList
        Get
            Return _ArFileNameOnly
        End Get
        Set(ByVal value As ArrayList)
            _ArFileNameOnly = value
        End Set
    End Property
    Private Property FileNameOnly() As String
        Get
            Return _FileNameOnly
        End Get
        Set(ByVal value As String)
            _FileNameOnly = value
        End Set
    End Property
    Public Property SortedFiles() As ArrayList
        Get
            Return _SortedFileNames

        End Get
        Set(ByVal value As ArrayList)
            _SortedFileNames = value
        End Set
    End Property
#End Region
#Region "Private Subs/Functions"




    Private Function DetermineInititalDirectory()
        '' find it
        '' is it there ?: yes->move to it: no->create it
        Try
            Dim dir As DirectoryInfo
            dir = New DirectoryInfo(_InitDirPath)
            If dir.Exists = True Then
                '' do nothing is there.
                _InitDirExists = True

            ElseIf dir.Exists = False Then
                '' create directory
                ''
                '' then create subdirectories
                ''
                'CreateDirectory(_InitDirPath)
                ''
                '' now create sub directories 
                'CreateSubDirectories(_currentLeadNum)
                Exit Function
            End If

        Catch ex As Exception
            '' log to errorlog

        End Try
        If _InitDirExists = True Then
            Return _InitDirExists '' now return if you found it or not || since the call to createdirectory exists on 'false' call, chances are this will always be true || however in the off chance it's not and you need to target it, that's why it's returned 
        End If

    End Function



    Private Sub CreateSubDirectories(ByVal LeadNum As String)
        Try
            '' left lead num as an argument to be fed from static vars 
            '' again these paths should be recycled from static vars
            '' added extra locations on end of string 
            Dim _dirAttFiles As String = _InitDirPath & LeadNum.ToString & "\"
            'Dim _dirSchedTasks As String = _InitDirPath & "\" & LeadNum.ToString & "\Scheduled Tasks"
            Dim _dirJobPictures As String = _InitDirPath & "\Job Pictures" ''comeback 
            '' method to actually create the directory 
            CreateDirectory(_dirAttFiles)
            'CreateDirectory(_dirSchedTasks)
            'CreateDirectory(_dirJobPictures)

        Catch ex As Exception
            '' call to error log
        End Try
    End Sub
    Public Sub Populate_Sub_Directories(ByVal frm As Form)
        Try
            Dim f = frm
            Dim di As New IO.DirectoryInfo(STATIC_VARIABLES.AttachedFilesDirectory & STATIC_VARIABLES.CurrentID)
            Dim Drs() As IO.DirectoryInfo = di.GetDirectories()
            For Each dr As IO.DirectoryInfo In Drs
                f.lvAttachedFiles.Items.Add(dr.Name)

            Next

        Catch ex As Exception

        End Try

    End Sub
#End Region
#Region "Destructor"
    Private Sub Destroy()
        Me.ArFileNameOnly = Nothing
        Me.ProfilePath = Nothing
        Me._ArFileNameOnly = Nothing
        Me._currentLeadNum = Nothing
        Me._FileNameOnly = Nothing
        Me._InitDirExists = Nothing
        Me._InitDirPath = Nothing
    End Sub
#End Region

End Class
