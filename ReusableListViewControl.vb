Imports System.Windows.Forms
Imports System.Drawing.Icon
Imports System.IO.FileInfo
Imports System.IO.DirectoryInfo
Imports IWshRuntimeLibrary
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Drawing.Bitmap
Imports System.Drawing.Drawing2D
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.String
Imports System.Text


Public Class ReusableListViewControl
#Region "structures"
    Public Structure FileObject
        Public FileName As String
        Public FileExt As String
        Public FullPath As String
        Public IconKey As String
        Public smIcon As Icon ''16
        Public lgIcon As Icon '' 32
        Public mdIcon As Icon '' 48-64
        Public smThumb As Image
        Public mdThumb As Image
        Public lgThumb As Image
        Public jbIcon As Icon '' 256x256
        Public FileSize As String
        Public DateCreated As String
        Public DateModified As String
    End Structure

     

    Public Structure DirObject
        Public FileName As String
        Public FullPath As String
        Public IconKey As String
        Public smIcon As Icon ''16
        Public lgIcon As Icon '' 32
        Public mdIcon As Icon '' 48-64
        Public jbIcon As Icon '' 256x256
        Public FileSize As Integer
        Public DateCreated As String
        Public DateModified As String
        Public HasSubDirsAndFiles As Boolean
    End Structure

    Public Structure SHFILEINFO
        Public hIcon As IntPtr ' : icon
        Public iIcon As Integer ' : icon index
        Public dwAttributes As Integer ' 
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String '' size must be marshaled because of com interop
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String '' size must be marshaled because of com interop
    End Structure

    Public Structure IconDir
        Dim lpszFile As String
        Dim IconIndex As Integer
        Dim LargeIcons()
        Dim SmallIcons()
        Dim NumIcons As Integer
    End Structure

    Public Structure IconDir_DEF
        Dim lpszFile As String
        Dim IconIndex As Integer
        Dim LargeIcons()
        Dim SmallIcons()
        Dim NumIcons As Integer
    End Structure
#End Region
#Region "Constants"
    Public Const SHIL_JUMBO = &H4 '' 256x256 system icons
    Public Const SHIL_LARGE = &H0 '' 48x48 
    Public Const SHIL_SMALL = &H1 '' 16x16 
    Public Const SHIL_EXTRALARGE = &H2 '' Typically 48x48 but can be customized. 
    Private Const SHGFI_ICON = &H100 '' default icon size | required for shgetfileinfo flag
#End Region
#Region "Mimic Sys Image List For Icon Extraction"
    Public Class MimicSysImgList
        <StructLayout(LayoutKind.Sequential)>
        Public Structure Rect
            Public left, top, right, bottom As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Private Structure POINT
            Dim x As Integer
            Dim y As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure IMAGELISTDRAWPARAMS
            Public cbSize As Integer
            Public himl As IntPtr
            Public x As Integer
            Public y As Integer
            Public cx As Integer
            Public cy As Integer
            Public xBitmap As Integer
            Public yBitmap As Integer
            Public rgbBk As Integer
            Public rgbFg As Integer
            Public fStyle As Integer
            Public dwRop As Integer
            Public fState As Integer
            Public frame As Integer
            Public crEffect As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure IMAGEINFO
            Public hbmImage As IntPtr
            Public hbmMask As IntPtr
            Public Unused1 As Integer
            Public Unused2 As Integer
            Public rcImage As Rect
        End Structure

        <ComImportAttribute()>
        <GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950")>
        <InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
        Interface IImageList

            <PreserveSig()>
            Function Add(ByVal hbmImage As IntPtr, ByVal hbmMask As IntPtr, ByRef pi As Integer) As Integer

            <PreserveSig()>
            Function ReplaceIcon(ByVal i As Integer, ByVal hicon As IntPtr, ByRef pi As Integer) As Integer

            <PreserveSig()>
            Function SetOverlayImage(ByVal iImage As Integer, ByVal iOverlay As Integer) As Integer

            <PreserveSig()>
            Function Replace(ByVal i As Integer, ByVal hbmImage As IntPtr, ByVal hbmMask As IntPtr) As Integer

            <PreserveSig()>
            Function AddMasked(hbmImage As IntPtr, crMask As Integer, ByRef pi As Integer) As Integer

            <PreserveSig()>
            Function Draw(ByRef pimldp As IMAGELISTDRAWPARAMS) As Integer

            <PreserveSig()>
            Function Remove(ByVal i As Integer) As Integer

            <PreserveSig()>
            Function GetIcon(ByVal i As Integer, ByVal flags As Integer, ByRef picon As IntPtr) As Integer

            <PreserveSig()>
            Function GetImageInfo(ByVal i As Integer, ByRef pImageInfo As IMAGEINFO) As Integer

        End Interface

    End Class
#End Region
#Region "PINVOKE and MISC Vars"
    Public cnt As Integer = 399 '' as returned by a  total count in imageres.dll off of a windows 8 machine
    Public LargeIcons() As IntPtr = New IntPtr(cnt - 1) {}
    Public SmallIcons() As IntPtr = New IntPtr(cnt - 1) {}

    Public Const ILD_TRANSPARENT As Integer = 1 '' for draw params / would leave on transparent other flags cause overlays to be used
    Public Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImagelist As Integer, ByRef riid As Guid, ByRef ppv As MimicSysImgList.IImageList) As IntPtr
    Public Declare Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr
    Public Declare Auto Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal libName As String, ByVal iconIndex As Integer, ByVal largeIcon() As IntPtr, ByVal smallIcon() As IntPtr, ByVal nIcons As Integer) As Integer

    Private _ReconFile As FileObject

#End Region
#Region "DLL Import for PINVOKE"
    <DllImport("comctl32.dll", SetLastError:=True)> _
    Public Shared Function ImageList_GetIcon(himl As IntPtr, i As Integer, flags As Integer) As IntPtr
    End Function
#End Region
#Region "Public Vars"
    Public imlSM As ImageList
    Public imlLG As ImageList
    Public imlXL As ImageList
    Public imlJB As ImageList
    Friend WithEvents lsControl As ListView
    Public lsCollection As ListView.ListViewItemCollection
    Public idx As Integer = -1
    Friend WithEvents ctx As ContextMenuStrip
    'Public _tmpPath As String = "C:\Users\" & Environment.UserName & "\AppData\Local\Temp\ISS\IMG" '' temp location for icon extraction and thumb generations
    Public _CurrentPath As String = ""
    Public PrevFileNameExt As String
    Public RenamePrevious As String
    Public PrevDirName As String

    '' for item over and ccp operations
    '' 9-6-2015
    Public Sel_TEXT As String = ""
    Public Right_Click_Sel_Text As String = ""
    Public Dest_TEXT As String = ""
    Public Itm_IDX As Integer = -1
    Public Item_SEL As ListViewItem   '' for dbl click operations

#End Region


#Region "Notes"

    '' edits: 9-6-2015
    '' ccp ON folder
    '' mouse down / up events
    '' item drag
    '' 4 private functions to support ccp operations
    '' 


    ''----------------------------------------------------
    '' edits: 8-27-2015
    '' 
    '' delete function not deleting files from context menu
    '' as soon as attached files added, need to refresh control to show file
    '' need to NOT show thumbs.db
    '' rename needs to operate like 'explorer' where just label is unlocked to edit the name of file
    '' assume last file extension on rename, don't ask.
    '' discard 'opens with' from context menu
    '' drag and drop functionality to folders
    '' lock "up one level button on 'sales' when 'max' directory is hit
    '' when creating a 'new' folder, unlock label to edit like explorer 
    '' auto refresh when new folder added to show new folder
    '' need a function to resize control on resize of form
    '' ---------------------------------------------------


    ''
    '' 1 generate a listview control
    '' 2 generate a new collection of objects [files/folders]
    '' 3 populate list view items from collection
    '' 4 clear->repopulate on movedir [retains selected view: view.largeicons]
    '' 5 clear items
    '' 6 open items -> double click [folder/file]
    '' 7 generate context menus
    '' 8 change sizes of icons [small/large/extralarge/jumbo]
    '' 9 edit labels [was before/after edit]
    '' 10 drag and drop [from where / to where]
    '' 11 create shortcuts [target directory / name of shortcut / location of shortcut]
    '' 12 create directories  [target path] 
    ''
#End Region
#Region "Item Hover and C.C.P. Operations"
    Private Sub RefreshUI()
        Me.lsControl.Items.Clear()
        Me.Sel_TEXT = ""

        Me.Dest_TEXT = ""
        Me.GetDirInfo(Me._CurrentPath)
        Me.lsControl.HoverSelection = False
        Me.lsControl.MultiSelect = False
        '' item 6 in ctx is "Paste" option
        Me.ctx.Items(6).Enabled = False
        'My.Computer.Clipboard.Clear()
    End Sub
#End Region
#Region "Constructor"
    Public Sub GenerateListControl(ByVal ctrl As Control, ByVal InitialDir As String, ByVal WhereToPut As System.Drawing.Point, ByVal Name_OF_Control As String, ByVal Height As Integer, ByVal Width As Integer)

        _CurrentPath = InitialDir


        lsControl = New ListView

        lsControl.Scrollable = True

        lsControl.AllowDrop = True

        lsCollection = New ListView.ListViewItemCollection(lsControl)
        lsControl.View = View.List  '' default view 
        lsControl.Name = Name_OF_Control.ToString
        lsControl.Dock = DockStyle.None
        lsControl.Size = New System.Drawing.Size(Width, Height)
        lsControl.LabelWrap = True
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.AutoArrange = True

        lsControl.HoverSelection = False
        lsControl.MultiSelect = False

        Dim pd As Padding
        With pd
            .All = 2
            .Top = 2
            .Bottom = 2
            .Left = 2
            .Right = 2
        End With

        lsControl.Margin = pd


        lsControl.Location = WhereToPut
        lsControl.LabelEdit = True

        imlSM = New ImageList
        imlSM.ImageSize = New System.Drawing.Point(16, 16)
        imlSM.ColorDepth = ColorDepth.Depth32Bit
        imlLG = New ImageList
        imlLG.ImageSize = New System.Drawing.Point(48, 48)
        imlLG.ColorDepth = ColorDepth.Depth32Bit
        imlXL = New ImageList
        imlXL.ColorDepth = ColorDepth.Depth32Bit
        imlXL.ImageSize = New System.Drawing.Point(64, 64)
        imlJB = New ImageList
        imlJB.ImageSize = New System.Drawing.Size(256, 256)
        imlJB.ColorDepth = ColorDepth.Depth32Bit


        GetDirInfo(InitialDir)

        lsControl.SmallImageList = imlLG '' 16
        lsControl.LargeImageList = imlXL '' 48

        lsControl.ContextMenuStrip = CreateContextMenu(lsControl)

        ' Form1.pnlTarget.Controls.Add(lsControl)
        ''width 363
        ''height 172

        'lsControl.Width = 363
        'lsControl.Height = 172

        ctrl.Controls.Add(lsControl)

    End Sub


    

#End Region
#Region "Methods, Subs, Functions, Props"
    Public Sub ChangeDirectory(ByVal TargetDirectory As String, ByVal ListView As ListView)
        ListView.Items.Clear()
        lsControl = ListView
        lsCollection = New ListView.ListViewItemCollection(lsControl)
        GetDirInfo(TargetDirectory)
    End Sub

    Private Function CreateContextMenu(ByVal lsControl As ListView)
        ctx = New ContextMenuStrip
        Dim mnuItem1 As ToolStripMenuItem
        mnuItem1 = New ToolStripMenuItem("Views")
        ' My.Computer.Clipboard.Clear()

        '' list  | medium 
        '' details | medium
        '' small | medium
        '' medium | large
        '' large | extra large
        '' extra large | jumbo
        ''

        mnuItem1.DropDownItems.Add("List", Nothing, AddressOf ChangeToList)
        mnuItem1.DropDownItems.Add("Details", Nothing, AddressOf ChangeToDetails)
        mnuItem1.DropDownItems.Add("Small", Nothing, AddressOf ChangeToSmall)
        mnuItem1.DropDownItems.Add("Medium", Nothing, AddressOf ChangeToLarge)
        mnuItem1.DropDownItems.Add("Large", Nothing, AddressOf ChangeToExtraLarge)
        mnuItem1.DropDownItems.Add("Extra Large", Nothing, AddressOf ChangeToJumbo)

        '' sort by
        ''
        Dim mnuitem2 As New ToolStripMenuItem("Sort By")
        mnuitem2.DropDownItems.Add("Ascending", Nothing, AddressOf SortAscending)
        mnuitem2.DropDownItems.Add("Descending", Nothing, AddressOf SortDescending)
        Dim mnuItem3 As New ToolStripMenuItem("Rename", Nothing, AddressOf RenameFileNonLabel)
        Dim mnuItem_1 As ToolStripSeparator = New ToolStripSeparator
        Dim mnuItem4 As New ToolStripMenuItem("Cut", Nothing, AddressOf CutMe)
        Dim mnuItem5 As New ToolStripMenuItem("Copy", Nothing, AddressOf CopyMe)
        Dim mnuItem6 As New ToolStripMenuItem("Paste", Nothing, AddressOf PasteMe)
        Dim mnuItem_2 As ToolStripSeparator = New ToolStripSeparator
        '' nest this into a view -> dropdown 
        '' 
        Dim mnuItem_3 As ToolStripSeparator = New ToolStripSeparator
        Dim mnuNew As New ToolStripMenuItem("New")
        mnuNew.DropDownItems.Add("Send To Desktop (Create Shortcut)", Nothing, AddressOf CreateADesktopShortCut)
        mnuNew.DropDownItems.Add("Create ShortCut Here", Nothing, AddressOf CreateShortCutHere)
        mnuNew.DropDownItems.Add("New Folder", Nothing, AddressOf CreateNewFolderHere)
        Dim mnuItem_4 As ToolStripSeparator = New ToolStripSeparator
        'Dim mnuItem10 As New ToolStripMenuItem("Send To Desktop (Create Shortcut)", Nothing, AddressOf CreateADesktopShortCut)
        'Dim mnuItem11 As New ToolStripMenuItem("Create ShortCut Here", Nothing, AddressOf CreateShortCutHere)
        'Dim mnuItem7 As New ToolStripMenuItem("New Folder", Nothing, AddressOf CreateNewFolderHere)


        Dim mnuItem8 As New ToolStripMenuItem("Delete", Nothing, AddressOf DeleteItem)
        Dim mnuItem9 As New ToolStripMenuItem("Refresh", Nothing, AddressOf RefreshMe)

        '' place context menu items in container
        ''
        ctx.Items.Add(mnuItem1)
        ctx.Items.Add(mnuitem2)
        ctx.Items.Add(mnuItem3)
        ctx.Items.Add(mnuItem_1)
        ctx.Items.Add(mnuItem4)
        ctx.Items.Add(mnuItem5)
        ctx.Items.Add(mnuItem6)
        If My.Computer.Clipboard.ContainsData(DataFormats.Text) = True Then
            mnuItem6.Enabled = True
        ElseIf My.Computer.Clipboard.ContainsData(DataFormats.Text) = False Then
            mnuItem6.Enabled = False
        End If
        ctx.Items.Add(mnuItem_2)
        ctx.Items.Add(mnuItem8)
        ctx.Items.Add(mnuItem9)
        ctx.Items.Add(mnuItem_3)
        ctx.Items.Add(mnuNew)


        ''EDIT:
        '' no need for seperator as OpensWith was Deprecated
        '' 8-26-15

        'ctx.Items.Add(mnuItem_4)

        Return ctx
    End Function
#Region "Private Functions for CCP / Item Hover"
    Private Function DropTarget_F_or_F(ByVal DropItem As String)
        Dim F_or_F As String = "neither"
        If System.IO.File.Exists(DropItem) = True Then
            F_or_F = "file"
        ElseIf System.IO.Directory.Exists(DropItem) = True Then
            F_or_F = "folder"
        End If
        Return F_or_F
    End Function

    Private Function Strip_Off_File_Or_Folder_Name(ByVal ItemText As String)
        Dim f_name = ItemText.ToString.Split("\")
        Dim x As Integer = 0
        Dim z As Object
        For Each z In f_name
            x += 1
        Next
        Dim FF_Name As String = f_name(x - 1)
        Return FF_Name
    End Function

    Private Function Is_it_a_file_or_folder(ByVal ItemText As String)
        Dim returnString As String = "neither"
        If InStr(ItemText, ".", ) >= 1 Then
            returnString = "file"
        ElseIf InStr(ItemText, ".", ) <= 0 Then
            returnString = "folder"
        End If
        Return returnString
    End Function
#End Region
    Private Sub lsControl_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lsControl.AfterLabelEdit

        '' 
        '' new TESTED code
        '' 8-31-2015
        '' 
        '' Tested Lead 11783 for file 'Musica.txt' in tandem with attach files repop
        '' 


        If e.Label Is Nothing Then

            'MsgBox("nothing recieved",critical,"DEBUG INFO ONLY")
            Return
        ElseIf e.Label <> Nothing Then
            Try
                Dim snd As ListView = sender
                Dim y As Integer = 0
                y = e.Item
                Dim lvSelected As ListViewItem = snd.Items(y)
                Dim fileName As String = lvSelected.Text
                Dim filePath As String = lvSelected.Tag
                Dim splits = lvSelected.Tag.ToString.Split("\")
                Dim tt As Object
                Dim count As Integer = 0
                For Each tt In splits
                    count += 1
                Next
                Dim fileExtPeices = splits(count - 1).ToString.Split(".")
                Dim fileExt
                fileName = fileExtPeices(0).ToString
                fileExt = "." & fileExtPeices(1).ToString
                fileName = (e.Label & fileExt)
                splits(count - 1) = fileName
                Dim g As Integer
                Dim reconstructed As String = ""
                For g = 0 To count - 1
                    reconstructed += ("\" & splits(g).ToString)
                Next
                Rename(lvSelected.Tag, reconstructed)
                lvSelected.Tag = reconstructed
            Catch ex As Exception
                '' fail it if it doesn't work
            End Try
        End If


        ''
        '' OBSOLETE CODE
        '' 8-31-15
        ''

        'Dim xx As ListViewItem
        'xx = lsControl.Items(idx)
        'Dim FullPath As String = xx.Tag
        ' '' decision here to rename file or directory
        ' ''
        'Dim snd As ListView = sender
        'Dim y As Integer = 0
        'y = e.Item

        'Dim lvSelected As ListViewItem = snd.Items(y)
        'Dim fileName As String = lvSelected.Text
        'Dim filePath As String = lvSelected.Tag
        'Dim splits = Split(lvSelected.Tag, "\", -1, Microsoft.VisualBasic.CompareMethod.Text)
        'Dim tt As Object
        'Dim count As Integer = 0
        'For Each tt In splits
        '    count += 1
        'Next
        'fileName = (splits(count - 1))
        'MsgBox("FileName: " & fileName & vbCrLf & "FilePath: " & filePath)


        'If InStr(FullPath, ".") > 0 Then '' file 
        '    '' this is a file.
        '    If e.Label <> "" Then
        '        '' valid label
        '        '' 
        '        ''perform operations
        '        '' 
        '        Dim splitNames = Split(FullPath, "\")
        '        Dim t As Object
        '        Dim cntOfPieces As Integer
        '        For Each t In splitNames
        '            cntOfPieces += 1
        '        Next
        '        Dim furtherSplit = Split(splitNames(cntOfPieces - 1), ".")
        '        Dim oldFileName As String = furtherSplit(0)
        '        Dim oldFileExt As String = furtherSplit(1)
        '        Dim gg
        '        Dim path As String = ""
        '        Dim cnt As Integer = -1
        '        For Each gg In splitNames '' old path 
        '            cnt += 1
        '            If cnt < (cntOfPieces - 1) Then
        '                path = (path & gg & "\")
        '            ElseIf cnt = (cntOfPieces - 1) Then
        '                path = (path & gg)
        '            End If

        '        Next
        '        cnt = -1
        '        Dim NewPath As String = ""
        '        Dim gh
        '        For Each gh In splitNames '' old path 
        '            cnt += 1
        '            If cnt < (cntOfPieces - 1) Then
        '                NewPath = (NewPath & gh & "\")
        '            ElseIf cnt = (cntOfPieces - 1) Then
        '                NewPath = (NewPath & e.Label & "." & oldFileExt)
        '            End If

        '        Next
        '        Try
        '            'path = (path & "\" & e.Label & "." & oldFileExt)
        '            'MsgBox(FullPath & " || " & NewPath)
        '            xx.Tag = NewPath
        '            Rename(FullPath, NewPath)
        '        Catch ex As Exception
        '            MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Error Renaming File")
        '        End Try

        '    ElseIf e.Label.Length <= 0 Then
        '        '' empty label 
        '        '' fail operations.
        '        '' 
        '        MsgBox("Fail it.")
        '    End If

        'ElseIf InStr(FullPath, ".") <= 0 Then '' folder
        '    '' this is a folder.
        '    If e.Label <> "" Then
        '        '' valid folder label
        '        ''
        '        ' FullPath = FullPath & "\" & xx.Text
        '        Dim FullPathSplit = Split(FullPath, "\")
        '        Dim cntItem As Integer = 0
        '        For Each xyz As Object In FullPathSplit
        '            cntItem += 1
        '        Next
        '        Dim cnt = (cntItem - 1)
        '        Dim oldDirName As String = FullPathSplit(cnt)
        '        Dim newDirName As String = e.Label
        '        Dim pos = cnt - 1

        '        'FullPathSplit(FullPathSplit.Count - 1) = newDirName
        '        Dim reconPath As String = ""
        '        Dim i As Integer = 0
        '        For i = 0 To (cnt - 2)
        '            If i = cntItem - 1 Then
        '                reconPath = reconPath & FullPathSplit(i)
        '            ElseIf i <> cntItem - 1 Then
        '                reconPath = reconPath & FullPathSplit(i) & "\"
        '            End If
        '        Next
        '        Dim NewPath As String = (reconPath & e.Label.ToString)
        '        Try
        '            xx.Tag = NewPath
        '            Rename(FullPath, NewPath)

        '        Catch ex As Exception
        '            MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Error Renaming Folder")
        '        End Try

        '    ElseIf e.Label.Length <= 0 Then
        '        '' non valid label
        '        '' 

        '    End If
        'End If

    End Sub

    Public Sub lsControl_BeforeLabelEdit(sender As Object, e As LabelEditEventArgs) Handles lsControl.BeforeLabelEdit

        ''
        '' OBSOLETE CODE
        '' 8-31-2015
        '' nothing needs to happen here
        '' 

        'Dim lvItem As ListViewItem
        'For Each lvItem In lsControl.Items
        '    If lvItem.Selected = True Then
        '        idx = lvItem.Index
        '        ''
        '        '' need a way here to determine if renaming a folder or file. 
        '        '' folder will be directroy name only
        '        '' file will be file name only, then replace file ext. 
        '        '' this side will only gather what the file/folder WAS called in order to pass arg for rename method. 
        '        '' 

        '        Dim FullPath As String = lvItem.Tag
        '        If InStr(FullPath, ".", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
        '            '' this is a file if the period[.] char found. 
        '            Dim subsplit = Split(FullPath, "\")
        '            Dim countOfPieces As Integer
        '            Dim t As Object
        '            For Each t In subsplit
        '                countOfPieces += 1
        '            Next
        '            Dim furthersplit = Split(subsplit(countOfPieces - 1), ".")
        '            Dim FileName As String = furthersplit(0)
        '            Dim FileExt As String = furthersplit(1)
        '            PrevFileNameExt = FileExt
        '            RenamePrevious = FullPath

        '        ElseIf InStr(FullPath, ".", Microsoft.VisualBasic.CompareMethod.Text) <= 0 Then
        '            '' this is a folder if no period[.] found. 
        '            Dim subsplit = Split(FullPath, "\")
        '            Dim itemCnt As Integer = 0
        '            For Each xy As Object In subsplit
        '                itemCnt += 1
        '            Next
        '            Dim DirToRename As String = subsplit(itemCnt - 1)
        '            PrevDirName = DirToRename
        '            RenamePrevious = FullPath

        '        End If

        '    End If
        'Next
    End Sub

    

    Private Sub lsControl_DoubleClick(sender As Object, e As EventArgs) Handles lsControl.DoubleClick
        'Dim x As ListViewItem = lsControl.Items(Itm_IDX)

        'If x IsNot Nothing Then
        '    If Item_SEL.SubItems(1).Text = "File Folder" Then
        '        _CurrentPath = Item_SEL.Tag
        '        GetDirInfo(_CurrentPath)
        '    ElseIf Item_SEL.SubItems(1).Text <> "File Folder" Then
        '        System.Diagnostics.Process.Start(Item_SEL.Tag)
        '    End If

        'End If





            'Dim x As ListViewItem
            'For Each x In lsControl.Items
            '    If x.Selected = True Then
            '        If x.SubItems(1).Text = "File Folder" Then
            '            GetDirInfo(_CurrentPath & "\" & x.SubItems(0).Text)
            '            _CurrentPath = (_CurrentPath & "\" & x.SubItems(0).Text).ToString
            '            STATIC_VARIABLES.AttachedFilesDirectory = _CurrentPath
            '            GetDirInfo(_CurrentPath)
            '        End If
            '        If x.SubItems(1).Text <> "File Folder" Then
            '            MsgBox(x.Tag.ToString)
            '            System.Diagnostics.Process.Start(x.Tag.ToString)
            '        End If
            '    End If
            'Next
    End Sub

    Private Sub lsControl_DragDrop(sender As Object, e As DragEventArgs) Handles lsControl.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim arFiles() As String
            Dim x As Integer
            arFiles = e.Data.GetData(DataFormats.FileDrop)
            Dim i As Integer = 0
            For i = 0 To arFiles.Length - 1
                Dim lvItem As New ListViewItem
                lvItem.Tag = arFiles(i)
                Dim name As String = SplitApartFileName(arFiles(i))
                Dim FileExt As String = SplitApartFileExt(arFiles(i))
                Dim OldFileName As String = (name & "." & FileExt)
                lsControl.Items.Add(name, lvItem.Tag)
                System.IO.File.Move(arFiles(i), _CurrentPath & "\" & OldFileName)
            Next
        End If
        GetDirInfo(_CurrentPath)
    End Sub

    Private Sub lsControl_DragEnter(sender As Object, e As DragEventArgs) Handles lsControl.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Move
        End If
    End Sub

    Private Sub lsControl_ItemActivate(sender As Object, e As EventArgs) Handles lsControl.ItemActivate
        Dim y As ListViewItem
        For Each y In Me.lsControl.Items
            If y.Selected = True Then
                Sel_TEXT = y.Tag
                Itm_IDX = y.Index
            End If
        Next
    End Sub

    Private Sub lsControl_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles lsControl.ItemDrag
        Dim y As ListViewItem = e.Item
        Sel_TEXT = y.Tag
        Me.lsControl.HoverSelection = True
        Dim colConv As ColorConverter = New ColorConverter
        Dim col As Color = colConv.ConvertFromString("#008AFF")
        y.ForeColor = Color.White
        y.BackColor = col
    End Sub

    Private Sub lsControl_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles lsControl.MouseDoubleClick
        Try
            Dim x As ListViewItem = lsControl.Items(Itm_IDX)

            If x IsNot Nothing Then
                If x.SubItems(1).Text = "File Folder" Then
                    _CurrentPath = x.Tag
                    GetDirInfo(_CurrentPath)
                ElseIf x.SubItems(1).Text <> "File Folder" Then
                    System.Diagnostics.Process.Start(Item_SEL.Tag)
                End If

            End If
        Catch ex As Exception

        End Try


    End Sub

     

    



    Private Sub lsControl_MouseDown(sender As Object, e As MouseEventArgs) Handles lsControl.MouseDown
        Select Case e.Button
            Case Is = MouseButtons.Left
                '' trap it but do nothing with it. 
                '' mouse double left click already assigned. 
                ''
                '' EDIT: dbl click not working right
                '' work around
                '' 9-13-2015
                ''
                Dim x As ListViewItem
                For Each x In lsControl.Items
                    If x.Selected = True Then
                        Itm_IDX = x.Index
                    End If
                Next
            Case Is = MouseButtons.Middle
                '' extra 
                '' trap it but do nothing with it as of yet. 
                '' 
                Exit Select
            Case Is = MouseButtons.Right
                '' Edit for CCP / Item Hover Selection
                '' 9-6-2015
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    Dim y As ListViewItem = Me.lsControl.GetItemAt(e.X, e.Y)
                    If y IsNot Nothing Then
                        Right_Click_Sel_Text = y.Tag
                        Dim fname As String = Strip_Off_File_Or_Folder_Name(Right_Click_Sel_Text)
                        Dim f_or_f As String = Is_it_a_file_or_folder(Right_Click_Sel_Text)
                        Itm_IDX = y.Index
                        'MsgBox("File or Folder? " & vbCrLf & f_or_f & vbCrLf & "Name? " & vbCrLf & fname, MsgBoxStyle.Information, "DEBUG INFO")
                        Me.ctx.Items(5).Enabled = True
                        Me.ctx.Items(4).Enabled = True
                        If My.Computer.Clipboard.ContainsText = True Then
                            Me.ctx.Items(6).Enabled = True
                        ElseIf My.Computer.Clipboard.ContainsText = False Then
                            Me.ctx.Items(6).Enabled = False
                        End If
                    End If
                    If y Is Nothing Then
                        Right_Click_Sel_Text = ""
                        If My.Computer.Clipboard.ContainsText = True Then
                            Me.ctx.Items(6).Enabled = True
                        ElseIf My.Computer.Clipboard.ContainsText = False Then
                            Me.ctx.Items(6).Enabled = False
                        End If
                        Me.ctx.Items(5).Enabled = False
                        Me.ctx.Items(4).Enabled = False
                    End If
                End If


                '' this is where we need to show context menu 
                '' with "open with" options. 
                '' also need a way to lock/unlock paste depending if their is data on the clipboard. 
                '' .....and naturally grab this item/text/object in question / under the right click if there is one. 
                '' 
                '' notes: need to clear clipboard 
                '' 
                '' just a generic call to clear the clipboard
                '' then show context menu 
                '' scenario: right click on the control but nothing is selected
                '' 


                '' EDIT:
                '' whole method is to be deprecated per andy
                '' 8-26-15
                '' Pulling too much garbage from Registry
                '' 


                'Dim objSelected As Boolean = False
                'Dim prxyItem As ListViewItem


                'Try

                '    Dim objInQuestion As ListViewItem
                '    For Each objInQuestion In lsControl.Items
                '        If objInQuestion.Selected = True Then '' display here.
                '            objSelected = True
                '            prxyItem = objInQuestion
                '            Exit For
                '        ElseIf objInQuestion.Selected = False Then
                '            '' do nothing then. 
                '            '' no item selected so do not display. 
                '            '' 
                '            ''
                '        End If
                '    Next

                '    If objSelected = True Then
                '        '' split the file ext off
                '        Dim itm As String = SplitApartFileName(prxyItem.Tag)
                '        Dim itm2 As String = SplitApartFileExt(prxyItem.Tag)
                '        Dim _File As String = itm
                '        Dim _EXT As String = itm2
                '        '' check registry for "opens with" of file extension
                '        Dim reg As New GetOpensWithNames(_EXT) '' "UTILITY CLASS" 
                '        '' create a list
                '        Dim openWith As ArrayList = reg.OpensWith
                '        '' ammend the context menu with "open with items"
                '        '' 
                '        '' look for previous instances first then get rid of them. 
                '        ''
                '        Dim x As Object
                '        For Each x In Me.ctx.Items
                '            If TypeOf x Is ToolStripSeparator Then
                '                '' do nothing with it just trap it
                '                '' 
                '            End If
                '            If TypeOf x Is ToolStripMenuItem Then
                '                If x.Text = "Opens With" Then
                '                    Me.ctx.Items.Remove(x)
                '                    Exit For '' found it, don't keep looping. 
                '                End If
                '            End If
                '        Next




                '        Dim mnuItemOpenWith As ToolStripMenuItem = New ToolStripMenuItem("Opens With")
                '        Dim bb
                '        For bb = 0 To openWith.Count - 1

                '            Dim b As String = openWith(bb)

                '            mnuItemOpenWith.DropDownItems.Add(b.ToLower, Nothing, AddressOf OpenWithFire) '' wire up events

                '        Next
                '        Dim ctx As ContextMenuStrip = Me.ctx
                '        ctx.Items.Add(mnuItemOpenWith)

                '    End If



                'Catch ex As Exception
                '    '' just fail it. 
                '    'Dim g As String = ex.InnerException.ToString
                '    'Dim b = g

                'End Try




                Exit Select

        End Select
    End Sub





    Private Sub RenameFileNonLabel()

        ''
        '' get old file name = lvitem.tag
        '' get user input for new file name
        '' make sure new file / filename does not exist
        '' if new file name does exist, and new filename the same, ask to overwrite
        '' yes: overwrite | no:cancel sub 
        '' if new file name does not exist, rename it
        ''  
        'Try
        '    Dim lvItem As ListViewItem
        '    For Each lvItem In lsControl.Items
        '        If lvItem.Selected = True Then
        '            Dim fileName As String = SplitApartFileName(lvItem.Tag)
        '            Dim fileExt As String = SplitApartFileExt(lvItem.Tag)
        '            Dim oldfilename As String = fileName & "." & fileExt
        '            Dim newFileName As String = InputBox("Please enter name of file.", "Rename: '" & oldfilename & "'", oldfilename)
        '            If InStr(newFileName, ".", Microsoft.VisualBasic.CompareMethod.Text) > 0 Then
        '                If newFileName = oldfilename Then
        '                    Exit Sub
        '                ElseIf newFileName = "" Then
        '                    Exit Sub
        '                End If
        '                If System.IO.File.Exists(_CurrentPath & "\" & newFileName) = True Then
        '                    Dim choice = MsgBox("Overwrite File name?", MsgBoxStyle.YesNo, "Overwrite with new file name?")
        '                    Select Case choice
        '                        Case Is = 5 'no
        '                            Exit Sub
        '                        Case Is = 6 'yes
        '                            Rename(_CurrentPath & "\" & oldfilename, _CurrentPath & "\" & newFileName)
        '                            GetDirInfo(_CurrentPath)
        '                            Exit Select
        '                    End Select
        '                ElseIf System.IO.File.Exists(_CurrentPath & "\" & newFileName) = False Then
        '                    Rename(_CurrentPath & "\" & oldfilename, _CurrentPath & "\" & newFileName)
        '                    GetDirInfo(_CurrentPath)
        '                End If
        '            ElseIf InStr(newFileName, ".", Microsoft.VisualBasic.CompareMethod.Text) <= 0 Then
        '                'Dim choice As Integer = MsgBox("You must supply a file extension when renaming a file." & vbCrLf & "Use the old one: '" & fileExt & "' ?", MsgBoxStyle.YesNo, "Use Old File Extension")
        '                'Select Case choice
        '                '    Case 5 ' no
        '                '        Exit Sub
        '                '    Case 6 ' yes
        '                newFileName = newFileName & "." & fileExt
        '                Rename(_CurrentPath & "\" & oldfilename, _CurrentPath & "\" & newFileName)
        '                GetDirInfo(_CurrentPath)
        '                'End Select
        '            End If

        '        End If
        '    Next
        For Each g As ListViewItem In lsControl.Items
            If g.Selected = True Then
                g.BeginEdit()
            End If
        Next
        'Catch ex As Exception
        '    '' fail it here 
        '    '' 
        'End Try
    End Sub

    Public Sub RefreshMe()

        lsControl.Items.Clear()
        Me.imlJB.Images.Clear()
        Me.imlLG.Images.Clear()
        Me.imlXL.Images.Clear()
        Me.imlSM.Images.Clear()
        idx = -1
        GetDirInfo(_CurrentPath)
    End Sub

#Region "Views"
    Public Sub ChangeToDetails()
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.Columns.Clear()
        lsControl.View = View.Details
        lsControl.Columns.Add("FileName", 200)
        lsControl.Columns.Add("Type", 70)
        lsControl.Columns.Add("Size", 70)
        lsControl.Columns.Add("Date Modified", 200)
        lsControl.SmallImageList = Me.imlLG
        lsControl.LargeImageList = Me.imlXL
        lsControl.Refresh()
    End Sub
    Public Sub ChangeToList()
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.View = View.List
        lsControl.SmallImageList = imlLG
        lsControl.Refresh()
    End Sub
    Public Sub ChangeToSmall()
        lsControl.Alignment = ListViewAlignment.Left
        lsControl.View = View.SmallIcon
        lsControl.SmallImageList = Me.imlSM
        lsControl.Refresh()
    End Sub
    Public Sub ChangeToLarge()
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.View = View.LargeIcon
        lsControl.LargeImageList = Me.imlLG
        lsControl.Refresh()
    End Sub
    Public Sub ChangeToExtraLarge()
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.View = View.LargeIcon
        lsControl.LargeImageList = Me.imlXL
        lsControl.Refresh()
    End Sub
    Public Sub ChangeToJumbo()
        lsControl.Alignment = ListViewAlignment.Default
        lsControl.View = View.LargeIcon
        lsControl.LargeImageList = Me.imlJB
        lsControl.Refresh()
    End Sub
#End Region

#Region "ShortCuts"
    Public Sub CreateADesktopShortCut()
        Dim defaultName As String = ""
        Dim dest As String = ""
        For Each yy As ListViewItem In lsControl.Items
            If yy.Selected = True Then
                If yy.SubItems(1).Text = "File folder" Then
                    defaultName = yy.Text
                    dest = _CurrentPath & "\" & yy.Text
                ElseIf yy.SubItems(1).Text <> "File folder" Then
                    defaultName = yy.Text
                    dest = yy.Tag
                End If
            End If
        Next
        Dim y As String = InputBox("Enter a name for your shortcut.", "Shortcut Name", "Shortcut to-" & defaultName)
        If y <> "" Then
            Dim x As New createAShortCut(y, dest)
            RefreshMe()
        End If
    End Sub
    Public Sub CreateShortCutHere()
        Dim defaultName As String = ""
        Dim dest As String = ""
        For Each yy As ListViewItem In lsControl.Items
            If yy.Selected = True Then
                If yy.SubItems(1).Text = "File folder" Then
                    defaultName = yy.Text
                    dest = _CurrentPath & "\" & yy.Text
                ElseIf yy.SubItems(1).Text <> "File folder" Then
                    defaultName = yy.Text
                    dest = yy.Tag
                End If
            End If
        Next
        Dim y As String = InputBox("Enter a name for your shortcut.", "Shortcut Name", "Shortcut to-" & defaultName)
        If y <> "" Then
            Dim x As New createAShortCut(y, dest, _CurrentPath)
            RefreshMe()
        End If
    End Sub

#End Region

#Region "Sorting"
    Public Sub SortAscending()
        Dim x As New ListSortAscending
        lsControl.ListViewItemSorter = x
    End Sub
    Public Sub SortDescending()
        Dim x As New ListSortDescending
        lsControl.ListViewItemSorter = x
    End Sub
#End Region

#Region "Cut Copy Paste"
    Public Sub CutMe(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim x As ListViewItem
        'For Each x In lsControl.Items
        '    If x.Selected = True Then
        '        My.Computer.Clipboard.SetData(DataFormats.Text, "CUT" & "|" & x.Tag)
        '        ctx.Items(6).Enabled = True '' 6 is the zero based index of paste
        '    End If
        'Next

        'My.Computer.Clipboard.Clear()
        My.Computer.Clipboard.SetText("MOVE|" & Right_Click_Sel_Text)
        'Me.ctx.Items(6).Enabled = True '' 0 based index for "Paste" option


    End Sub
    Public Sub PasteMe(ByVal sender As Object, ByVal e As System.EventArgs)
        If My.Computer.Clipboard.ContainsText = True Then

            Dim ClipData As String = My.Computer.Clipboard.GetText
            Dim cmdData = ClipData.ToString.Split("|")
            Dim cmd As String = cmdData(0)
            Dim srcText As String = cmdData(1)

            Dim destTarget As String = Right_Click_Sel_Text
            If destTarget = "" Then
                destTarget = _CurrentPath.ToString & "\" & Strip_Off_File_Or_Folder_Name(srcText)
                Select Case cmd
                    Case "COPY"
                        Dim fname As String = SplitApartFileName(srcText)
                        Dim fileExt As String = SplitApartFileExt(srcText)
                        If System.IO.File.Exists(destTarget) = True Then
                            'System.IO.File.Copy(ClipData, destTarget)
                            Try
                                System.IO.File.Copy(srcText, _CurrentPath & "\" & fname & " - Copy." & fileExt)
                                ' My.Computer.Clipboard.Clear()
                                GetDirInfo(_CurrentPath)
                            Catch ex As Exception
                                ' My.Computer.Clipboard.Clear()
                            End Try
                        ElseIf System.IO.File.Exists(destTarget) = False Then
                            Try
                                System.IO.File.Copy(srcText, _CurrentPath & "\" & fname & " - Copy." & fileExt)
                                ' My.Computer.Clipboard.Clear()
                                GetDirInfo(_CurrentPath)
                            Catch ex As Exception
                                ' My.Computer.Clipboard.Clear()
                            End Try
                        End If
                        Exit Select
                    Case "MOVE"
                        If System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = True Then
                            'System.IO.File.Copy(ClipData, destTarget)
                            Dim fname As String = SplitApartFileName(srcText)
                            Dim fileExt As String = SplitApartFileExt(srcText)
                            Try
                                System.IO.File.Move(srcText, destTarget & "\" & fname & " - Copy." & fileExt)
                                'My.Computer.Clipboard.Clear()
                                GetDirInfo(_CurrentPath)
                            Catch ex As Exception
                                ' My.Computer.Clipboard.Clear()
                            End Try
                        ElseIf System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = False Then
                            Try
                                System.IO.File.Move(srcText, destTarget)
                                ' My.Computer.Clipboard.Clear()
                                GetDirInfo(_CurrentPath)
                            Catch ex As Exception
                                ' My.Computer.Clipboard.Clear()
                            End Try
                        End If
                        Exit Select
                End Select
                

            ElseIf destTarget <> "" Then
                Dim FileOrFolderNameFromClip As String = Strip_Off_File_Or_Folder_Name(ClipData)
                Dim F_Or_F_Target As String = Is_it_a_file_or_folder(destTarget)
                Dim F_Or_F_Source As String = Is_it_a_file_or_folder(FileOrFolderNameFromClip)
                Select Case F_Or_F_Target
                    Case "file"
                        '' not valid
                        ' My.Computer.Clipboard.Clear()
                        Exit Select
                    Case "folder"
                        Select Case F_Or_F_Source
                            Case "file"
                                Select Case cmd
                                    Case "COPY"
                                        Dim fname As String = SplitApartFileName(srcText)
                                        Dim fileExt As String = SplitApartFileExt(srcText)

                                        If System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = True Then
                                            'System.IO.File.Copy(ClipData, destTarget)
                                            Try
                                                System.IO.File.Copy(srcText, destTarget & "\" & fname & " - Copy." & fileExt)
                                                ' My.Computer.Clipboard.Clear()
                                                GetDirInfo(_CurrentPath)
                                            Catch ex As Exception
                                                ' My.Computer.Clipboard.Clear()
                                            End Try
                                        ElseIf System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = False Then
                                            Try
                                                System.IO.File.Copy(srcText, destTarget & "\" & fname & "." & fileExt)
                                                ' My.Computer.Clipboard.Clear()
                                                GetDirInfo(_CurrentPath)
                                            Catch ex As Exception
                                                ' My.Computer.Clipboard.Clear()
                                            End Try
                                        End If
                                        Exit Select
                                    Case "MOVE"
                                        Dim fname As String = SplitApartFileName(srcText)
                                        Dim fileExt As String = SplitApartFileExt(srcText)

                                        If System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = True Then
                                            'System.IO.File.Copy(ClipData, destTarget)
                                            Try
                                                System.IO.File.Move(srcText, destTarget & "\" & fname & " - Copy." & fileExt)
                                                ' My.Computer.Clipboard.Clear()
                                                GetDirInfo(_CurrentPath)
                                            Catch ex As Exception
                                                ' My.Computer.Clipboard.Clear()
                                            End Try
                                        ElseIf System.IO.File.Exists(destTarget & "\" & Strip_Off_File_Or_Folder_Name(srcText)) = False Then
                                            Try
                                                System.IO.File.Move(srcText, destTarget & "\" & fname & "." & fileExt)
                                                'My.Computer.Clipboard.Clear()
                                                GetDirInfo(_CurrentPath)
                                            Catch ex As Exception
                                                ' My.Computer.Clipboard.Clear()
                                            End Try
                                        End If
                                        Exit Select
                                End Select
                            Case "folder"
                                Select Case cmd
                                    Case "COPY"
                                        '' don't copy directories 
                                        Try
                                            '  My.Computer.Clipboard.Clear()
                                            Exit Select
                                        Catch ex As Exception
                                            ' My.Computer.Clipboard.Clear()
                                        End Try
                                        Exit Select
                                    Case "MOVE"
                                        Try
                                            System.IO.Directory.Move(srcText, _CurrentPath & "\" & Strip_Off_File_Or_Folder_Name(srcText))
                                            ' My.Computer.Clipboard.Clear()
                                            GetDirInfo(_CurrentPath)
                                        Catch ex As Exception
                                            ' My.Computer.Clipboard.Clear()
                                        End Try
                                        Exit Select
                                End Select
                        End Select
                End Select

            End If
            
            ' MsgBox("Clipboard Data: " & ClipData.ToString & vbCrLf & "Destination Target: " & destTarget.ToString, MsgBoxStyle.Information, "DEBUG INFO")

        End If


        'If My.Computer.Clipboard.ContainsText = False Then
        '    Exit Sub
        'ElseIf My.Computer.Clipboard.ContainsText = True Then
        '    Dim y As ListViewItem
        '    For Each y In lsControl.Items
        '        If y.Selected = True Then
        '            Dest_TEXT = y.Tag
        '        End If
        '    Next

        '    If y IsNot Nothing Then '' they are trying to copy or paste to a ' target '
        '        Dim f_or_f_Destination As String = Is_it_a_file_or_folder(Dest_TEXT) '' are we sending it to a file or folder?
        '        Dim wholeName As String = My.Computer.Clipboard.GetText() '' the object text in question from clipboard
        '        Dim fileName As String = Strip_Off_File_Or_Folder_Name(wholeName) '' just the file name or folder name
        '        Dim f_or_f_target As String = Is_it_a_file_or_folder(wholeName) '' is the target that we are manipulating a file or folder ? 
        '        Select Case f_or_f_Destination
        '            Case "file"
        '                ' file to file paste is not valid
        '                My.Computer.Clipboard.Clear()
        '                Me.ctx.Items(6).Enabled = False
        '                Dest_TEXT = ""
        '                Right_Click_Sel_Text = ""
        '                Itm_IDX = -1
        '                lsControl.HoverSelection = False
        '                lsControl.MultiSelect = False
        '                For Each a As ListViewItem In lsControl.Items
        '                    a.Selected = False
        '                Next
        '                Exit Select
        '            Case "folder"
        '                ' this is a valid move
        '                Select Case f_or_f_target
        '                    Case "file"
        '                        If System.IO.File.Exists(Dest_TEXT & "\" & fileName) = True Then
        '                            ' make a copy of it here
        '                            System.IO.File.Copy(wholeName, (Dest_TEXT & "\" & fileName))
        '                            My.Computer.Clipboard.Clear()
        '                            Me.ctx.Items(6).Enabled = False
        '                            Dest_TEXT = ""
        '                            Right_Click_Sel_Text = ""
        '                            Itm_IDX = -1
        '                            lsControl.HoverSelection = False
        '                            lsControl.MultiSelect = False
        '                            For Each a As ListViewItem In lsControl.Items
        '                                a.Selected = False
        '                            Next
        '                            GetDirInfo(_CurrentPath)
        '                        ElseIf System.IO.File.Exists(Dest_TEXT & "\" & fileName) = False Then
        '                            System.IO.File.Move(wholeName, (Dest_TEXT & "\" & fileName))
        '                            My.Computer.Clipboard.Clear()
        '                            Me.ctx.Items(6).Enabled = False
        '                            Dest_TEXT = ""
        '                            Right_Click_Sel_Text = ""
        '                            Itm_IDX = -1
        '                            lsControl.HoverSelection = False
        '                            lsControl.MultiSelect = False
        '                            For Each a As ListViewItem In lsControl.Items
        '                                a.Selected = False
        '                            Next
        '                            GetDirInfo(_CurrentPath)
        '                        End If
        '                    Case "folder"
        '                        If System.IO.Directory.Exists(Dest_TEXT & "\" & fileName) = True Then
        '                            My.Computer.Clipboard.Clear()
        '                            Me.ctx.Items(6).Enabled = False
        '                            Dest_TEXT = ""
        '                            Right_Click_Sel_Text = ""
        '                            Itm_IDX = -1
        '                            lsControl.HoverSelection = False
        '                            lsControl.MultiSelect = False
        '                            For Each a As ListViewItem In lsControl.Items
        '                                a.Selected = False
        '                            Next
        '                            Exit Select
        '                        ElseIf System.IO.Directory.Exists(Dest_TEXT & "\" & fileName) = False Then
        '                            System.IO.Directory.Move(wholeName, (Dest_TEXT & "\" & fileName))
        '                            My.Computer.Clipboard.Clear()
        '                            Me.ctx.Items(6).Enabled = False
        '                            Dest_TEXT = ""
        '                            Right_Click_Sel_Text = ""
        '                            Itm_IDX = -1
        '                            lsControl.HoverSelection = False
        '                            lsControl.MultiSelect = False
        '                            For Each a As ListViewItem In lsControl.Items
        '                                a.Selected = False
        '                            Next
        '                            GetDirInfo(_CurrentPath)
        '                        End If
        '                End Select
        '            Case Else
        '        End Select
        '    End If

        '    If y Is Nothing Then '' they are tyring to copy or paste ' here ' 
        '        Dim wholeName As String = My.Computer.Clipboard.GetText() '' the object text in question from clipboard
        '        Dim fileName As String = Strip_Off_File_Or_Folder_Name(wholeName) '' just the file name or folder name
        '        Dim f_or_f_target As String = Is_it_a_file_or_folder(wholeName) '' is the target that we are manipulating a file or folder ? 

        '        '
        '        ' if the file exists here already, make a copy
        '        ' if it does not exist here , move it here 
        '        '

        '        Select Case f_or_f_target
        '            Case "file"
        '                If System.IO.File.Exists(_CurrentPath & "\" & fileName) = True Then
        '                    System.IO.File.Copy(wholeName, (_CurrentPath & "\" & fileName))
        '                    GetDirInfo(_CurrentPath)
        '                End If
        '                If System.IO.File.Exists(_CurrentPath & "\" & fileName) = False Then
        '                    System.IO.File.Move(wholeName, (_CurrentPath & "\" & fileName))
        '                    GetDirInfo(_CurrentPath)
        '                End If

        '                Exit Select
        '            Case "folder"
        '                If System.IO.Directory.Exists(_CurrentPath & "\" & fileName) = True Then
        '                    My.Computer.Clipboard.Clear()
        '                    Me.ctx.Items(6).Enabled = False
        '                    Dest_TEXT = ""
        '                    Right_Click_Sel_Text = ""
        '                    Itm_IDX = -1
        '                    lsControl.HoverSelection = False
        '                    lsControl.MultiSelect = False
        '                    For Each a As ListViewItem In lsControl.Items
        '                        a.Selected = False
        '                    Next
        '                End If
        '                If System.IO.Directory.Exists(_CurrentPath & "\" & fileName) = False Then
        '                    System.IO.Directory.Move(wholeName, (_CurrentPath & "\" & fileName))
        '                    GetDirInfo(_CurrentPath)
        '                End If
        '            Case Else
        '        End Select
        '    End If



        '    ' was it a cut operation or a paste operation? 
        '    ' 
        '    '
        '    'split off 'command' and then go from there. 
        '    '
        '    Dim commandAr = My.Computer.Clipboard.GetData(DataFormats.Text).ToString.Split("|")
        '    Dim CommandOp As String = commandAr(0)
        '    Dim source As String = commandAr(1)

        '    Select Case CommandOp

        '        Case Is = "COPY"
        '            Dim fileName As String = SplitApartFileName(source)
        '            Dim fileExt As String = SplitApartFileExt(source)
        '            Dim FileAndExt As String = (fileName & "." & fileExt)
        '            If System.IO.File.Exists(_CurrentPath & "\" & FileAndExt) = True Then
        '                ' produce a file with a '-Copy' appended to it. 
        '                ' 
        '                If System.IO.File.Exists(_CurrentPath & "\" & fileName & "-Copy." & fileExt) = True Then '' trying to make more then one copy. 
        '                    If System.IO.File.Exists(_CurrentPath & "\" & fileName & "-Copy(1)." & fileExt) = True Then '' why do you need more then two copies of the file. exit it. 
        '                        Exit Sub
        '                    ElseIf System.IO.File.Exists(_CurrentPath & "\" & fileName & "-Copy(1)." & FileAndExt) = False Then
        '                        System.IO.File.Copy(source, _CurrentPath & "\" & fileName & "-Copy(1)." & fileExt) '' give them one more then stop. 
        '                        GetDirInfo(_CurrentPath)
        '                    End If
        '                ElseIf System.IO.File.Exists(_CurrentPath & "\" & fileName & "-Copy." & fileExt) = False Then
        '                    System.IO.File.Copy(source, _CurrentPath & "\" & fileName & "-Copy." & fileExt)
        '                    GetDirInfo(_CurrentPath)
        '                End If

        '                Exit Sub
        '            ElseIf System.IO.File.Exists(_CurrentPath & "\" & FileAndExt) = False Then
        '                ' they are not pasting in the same directory
        '                ' 
        '                ' paste it here. or rather, just copy the file to a different directory.
        '                System.IO.File.Copy(source, _CurrentPath & "\" & FileAndExt)
        '                GetDirInfo(_CurrentPath) '' now refresh to show change.
        '            End If
        '            Exit Select
        '        Case Is = "CUT"



        '            Dim fileName As String = SplitApartFileName(source)
        '            Dim fileExt As String = SplitApartFileExt(source)
        '            Dim FileAndExt As String = (fileName & "." & fileExt)
        '            ' now check to make sure user is pasting in a different directory
        '            ' 
        '            If System.IO.File.Exists(Dest_TEXT & "\" & FileAndExt) = True Then
        '                '' do nothing. they are pasting in the same directory.
        '                '' 
        '                Exit Sub
        '            ElseIf System.IO.File.Exists(Dest_TEXT & "\" & FileAndExt) = False Then
        '                ' they are not pasting in the same directory
        '                ' 
        '                ' paste it here. or rather, just move the file.
        '                Try
        '                    System.IO.File.Move(source, Dest_TEXT & "\" & FileAndExt)
        '                    GetDirInfo(_CurrentPath) '' now refresh to show change.
        '                Catch ex As Exception

        '                End Try

        '            End If




        '            Exit Select
        '        Case "PASTE"
        '            Dim targetString As String = My.Computer.Clipboard.GetText()
        '            MsgBox("Source File: " & vbCrLf & targetString & vbCrLf & "Destination: " & vbCrLf & Dest_TEXT)
        '            Dim f_or_f_target As String = Is_it_a_file_or_folder(Dest_TEXT)
        '            Select Case f_or_f_target
        '                Case "file"
        '                    ' invalid 
        '                    RefreshUI()
        '                    Exit Select
        '                Case "folder"
        '                    ' is it a file or folder we are moving ? 
        '                    Dim ff_ As String = Is_it_a_file_or_folder(targetString)
        '                    Select Case ff_
        '                        Case "file"
        '                            Dim fname As String = Strip_Off_File_Or_Folder_Name(targetString)
        '                            Try
        '                                System.IO.File.Move(targetString, Dest_TEXT & "\" & fname)
        '                                RefreshUI()
        '                            Catch ex As Exception
        '                                RefreshUI()
        '                            End Try
        '                            Exit Select
        '                        Case "folder"
        '                            Dim fname As String = Strip_Off_File_Or_Folder_Name(targetString)
        '                            Try
        '                                System.IO.Directory.Move(targetString, Dest_TEXT & "\" & fname)
        '                                RefreshUI()
        '                            Catch ex As Exception
        '                                RefreshUI()
        '                            End Try
        '                            Exit Select
        '                    End Select
        '            End Select
        '    End Select
        '   End If
    End Sub
    Public Sub CopyMe(ByVal Sender As Object, ByVal e As System.EventArgs)
        'Dim x As ListViewItem
        'For Each x In lsControl.Items
        '    If x.Selected = True Then
        '        My.Computer.Clipboard.SetData(DataFormats.Text, x.Tag)
        '        ctx.Items(6).Enabled = True '' 6 is the zero based index of paste
        '    End If
        'Next
        'My.Computer.Clipboard.Clear()
        My.Computer.Clipboard.SetText("COPY|" & Right_Click_Sel_Text)

    End Sub

#End Region

    Public Sub OpenWithFire(ByVal sender As Object, ByVal e As System.EventArgs)

        ''
        '' this whole sub needs to stripp off the working directory to pass as a new process start info property instead of using a case by case scenario. 
        '' then strip off file name
        '' the run process.start()
        '' 
        '' truncating file names with a space seems to be a universal problem when passing large directories. 
        '' source: http://stackoverflow.com/questions/3299444/process-start-parameter-containing-spaces-has-problems-on-xp
        '' 
        Dim y As ListViewItem
        Dim exeFile As String = sender.ToString
        Dim workingDir As String = _CurrentPath


        For Each y In Me.lsControl.Items
            If y.Selected = True Then
                Try
                    Dim FileName As String = SplitApartFileName(y.Tag)
                    Dim FileExt As String = SplitApartFileExt(y.Tag)
                    Dim FilePath As String = (FileName & "." & FileExt)

                    '' works for everything but IExplore.exe as usual. 
                    '' gotta make something special for M$. Big Shocker.
                    '' 
                    Select Case sender.ToString
                        Case Is = "iexplore.exe"
                            ''
                            '' M$ is different then everyone else.
                            '' 
                            FilePath = ("file:///" & workingDir & "\" & FilePath)
                            FilePath = FilePath.Replace(" ", "%20")
                            System.Diagnostics.Process.Start(sender.ToString, FilePath)
                            Exit Select
                        Case Else
                            Dim StartInfo As ProcessStartInfo = New ProcessStartInfo
                            StartInfo.WorkingDirectory = _CurrentPath
                            StartInfo.FileName = sender.ToString
                            StartInfo.Arguments = FilePath
                            System.Diagnostics.Process.Start(StartInfo)
                            Exit Select
                    End Select




                    ''
                    '' NON WORKING ORIGINAL IDEA
                    '' 12/28/14
                    '' point of reference only
                    '' 





                    'Dim filepath As String = y.Tag.ToString
                    'Select Case exeFile
                    '    Case Is = "chrome.exe"
                    '        '' works as expected.  :)  Yay Google.
                    '        '' 
                    '        filepath = filepath.Replace(" ", "%20")
                    '        System.Diagnostics.Process.Start(sender.ToString, filepath)
                    '        Exit Select
                    '    Case Is = "firefox.exe"
                    '        ''
                    '        '' firefox double encodes even when replaced spaces.
                    '        '' will add an extra space when application launches.
                    '        ''
                    '        '' WorkAround:
                    '        '' Preface with "file:///" and then strip out spaces with %20 encoding.
                    '        '' 
                    '        filepath = ("file:///" & filepath.ToString)
                    '        filepath = filepath.Replace(" ", "%20")
                    '        System.Diagnostics.Process.Start(sender.ToString, filepath)
                    '        Exit Select
                    '    Case Is = "iexplore.exe"
                    '        ''
                    '        '' i dunno wtf is going on here. wont take ANY format.  ?????
                    '        ''

                    '        '' no matter what you pass to it, ie will open your homepage. 
                    '        '' launch it from shell [run cmd] and it works.  ??? 
                    '        '' Shell(exeFile + filepath) does not though. WTF M$ ?? 

                    '        '' WorkAround:
                    '        '' Preface with "file:///" and then strip out spaces with %20 encoding.
                    '        '' 

                    '        filepath = ("file:///" & filepath.ToString)
                    '        filepath = filepath.Replace(" ", "%20")
                    '        System.Diagnostics.Process.Start(sender.ToString, filepath)

                    '        Exit Select
                    '    Case Else
                    '        Dim startINfo As ProcessStartInfo = New ProcessStartInfo
                    '        startINfo.WorkingDirectory = "C:\Users\Clay\Desktop\New Folder\Images"
                    '        startINfo.FileName = "mspaint.exe"
                    '        startINfo.Arguments = "noRotate.jpg"
                    '        System.Diagnostics.Process.Start(startINfo)
                    '        Exit Select
                    'End Select


                Catch ex As Exception
                    MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Erorr Opening:" & y.Tag)
                End Try
            End If
        Next
    End Sub

    Public Sub CreateNewFolderHere()
        Dim path As String = _CurrentPath

        '' Obsololete code: 8-27-2015
        ''
        'If System.IO.Directory.Exists(_CurrentPath & "\" & "New Folder") <> True Then
        '    System.IO.Directory.CreateDirectory(_CurrentPath & "\" & "New Folder")
        'ElseIf System.IO.Directory.Exists(_CurrentPath & "\" & "New Folder") = True Then
        '    System.IO.Directory.CreateDirectory(_CurrentPath & "\" & "New Folder - Copy")
        'End If

        ''
        '' need an edit here to be able to put as many new folders in place as possible. 
        '' 
        '' not just a ' -copy' and thats it (essentially only 'two' new folders at once....and kaput.
        '' 
        '' 
        '' idea: UNTESTED just a thought 8-27-15
        ''
        '' for each directory in current-path
        ''    if directory name = instr("new folder")
        ''       cnt +=1 
        ''    end if
        '' next
        '' 
        ''
        '' system.io.directory.createdirectory(_currentPath & "\" & "New Folder -" & cnt + 1)
        '' 
        '' 
        '' new idea practical application
        '' -- Mimic 'Explorer'
        ''
        Dim iteration As Integer = 0
        lsCollection.Clear()
        Dim dir_ As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(_CurrentPath)
        Dim cnt = dir_.GetDirectories("New fo*", IO.SearchOption.AllDirectories)

        Dim xyz As System.IO.DirectoryInfo
        For Each xyz In cnt
            iteration += 1
        Next

        Dim next1 As Integer = (iteration + 1)
        If next1 <= 1 Then
            System.IO.Directory.CreateDirectory(_CurrentPath & "\" & "New folder")
        ElseIf next1 > 1 Then
            System.IO.Directory.CreateDirectory(_CurrentPath & "\" & "New folder (" & next1.ToString & ")")
        End If

        Dim dir_2 As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(_CurrentPath)
        Dim abc As System.IO.DirectoryInfo
        For Each abc In dir_2.GetDirectories()
            Dim lv As New ListViewItem
            lv.Text = abc.Name.ToString
            lsCollection.Add(lv)
        Next

        RefreshMe()
    End Sub

    Public Sub DeleteItem()
        Try
            Dim x As ListViewItem
            For Each x In Me.lsControl.Items
                If x.Selected = True Then

                    '' break it down here to delete file or directory or shortcut
                    '' 
                    Dim text As String = x.Tag
                    Dim TextSplit = text.ToString.Split("\")
                    Dim cntItems As Integer = 0
                    Dim zz As Object
                    For Each zz In TextSplit
                        cntItems += 1
                    Next
                    Dim filename = TextSplit(cntItems - 1)

                    ' MsgBox("File Name: " & text, MsgBoxStyle.Information, "DEBUG INFO - FILE DELETE")

                    If InStr(filename, ".") > 0 Then
                        '' its a file
                        Dim subSplit = filename.ToString.Split(".")
                        If subSplit(1) = "lnk" Then '' its a shortcut , just delete it
                            System.IO.File.Delete(x.Tag)
                            lsControl.Items.Clear()
                            RefreshMe()
                        End If
                        If System.IO.File.Exists(x.Tag) = True Then
                            Dim yesNo As Integer = MsgBox("Are you sure you want to delete file: " & vbCr & x.Tag & "?", MsgBoxStyle.YesNo, "Delete File?")
                            Select Case yesNo
                                Case 5 '' no
                                    ''MsgBox(yesNo)
                                    Exit Sub
                                Case 6 '' yes 
                                    ''MsgBox(yesNo)
                                    System.IO.File.Delete(x.Tag)
                                    lsControl.Items.Clear()
                                    RefreshMe()
                                    Exit Select
                            End Select
                        End If
                    ElseIf InStr(filename, ".") = 0 Then
                        '' its a folder
                        If System.IO.Directory.Exists(x.Tag) = True Then
                            Dim yesNo As Integer = MsgBox("Are you sure you want to delete the folder: " & vbCr & x.Tag & "?", MsgBoxStyle.YesNo, "Delete Folder?")
                            Select Case yesNo
                                Case 5 '' no
                                    ''MsgBox(yesNo)
                                    '' do nothing on no
                                    Exit Sub
                                Case 6 '' yes
                                    ''MsgBox(yesNo)

                                    System.IO.Directory.Delete(x.Tag, True)
                                    lsControl.Items.Clear()
                                    RefreshMe()
                                    Exit Select
                            End Select
                        End If
                    End If

                End If
            Next

        Catch ex As Exception
            '' fail it and error log maybe? 
            ''

            MsgBox("ERROR: " & vbCrLf & ex.InnerException.ToString, MsgBoxStyle.Critical, "DEBUG INFO - ERROR DELETING FILE")

        End Try

    End Sub

    Private Sub GetDirInfo(ByVal TargetPath As String)

        Dim dirNFO As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(TargetPath)
        If dirNFO.Exists = False Then
            dirNFO.Create()
        ElseIf dirNFO.Exists = True Then
            '' flush out the respective stuffs before re-poping
            '' 
            Dim rootDirAF As String = (STATIC_VARIABLES.AttachedFilesDirectory.ToString & STATIC_VARIABLES.CurrentID.ToString)
            Dim rootDirJP As String = (STATIC_VARIABLES.JobPicturesFileDirectory.ToString & STATIC_VARIABLES.CurrentID.ToString)
            If TargetPath = rootDirAF Then
                Sales.tsAttachedFilesNAV.Enabled = False
            ElseIf TargetPath <> rootDirAF Then
                Sales.tsAttachedFilesNAV.Enabled = True
            End If
            If TargetPath = rootDirJP Then
                Sales.tsAttachedFilesNAV.Enabled = False
            ElseIf TargetPath <> rootDirJP Then
                Sales.tsAttachedFilesNAV.Enabled = True
            End If
            lsControl.Cursor = Cursors.WaitCursor
            lsControl.Items.Clear()
            'Me.imlJB.Images.Clear()
            'Me.imlLG.Images.Clear()
            'Me.imlSM.Images.Clear()
            'Me.imlXL.Images.Clear()
            ''
            ''
            ''
            '' put progress bar logic here 
            '' 
            '' get a count of files
            '' set min and max (max = count of files)
            '' min = 1
            '' ever time it iterates through the loop, move progress bar one.
            '' step 1
            '' 
            Dim ab
            Dim cntFiles As Integer = 0
            For Each ab In dirNFO.GetFiles("*.*")
                cntFiles += 1
            Next

            Main.tsProgress.Minimum = 1
            Main.tsProgress.Value = 1

            Main.tsProgress.Maximum = cntFiles
            Main.tsProgress.Step = 1


            Dim x
            For Each x In dirNFO.GetFiles("*.*")
                Dim lvItem As New ListViewItem
                idx += 1
                Main.tsProgress.PerformStep()


                Dim y As New FileObject
                y.FullPath = x.FullName
                y.FileName = SplitApartFileName(x.FullName)
                y.FileExt = SplitApartFileExt(x.FullName)

                y.FileSize = Math.Round(x.Length / 1024, 2) & " KB"
                y.DateCreated = x.CreationTime.ToString
                y.DateModified = x.LastAccessTime.ToString '' access only, not write. 
                y.IconKey = (y.FileName & "." & y.FileExt)

                Try
                    y.smIcon = GetSmalls(y.FullPath) '' 16
                    Me.imlSM.Images.Add(y.smIcon)
                    Me.imlSM.Images.SetKeyName(idx, y.IconKey.ToString)
                Catch ex As Exception
                    '' point to default error icon here.
                    '' 
                End Try



                '' now choose here depending on file extension to pull out a thumb or a jumbo icon or what
                ''
                ''

               
                 


                If y.FileExt = "jpg" Or y.FileExt = "bmp" Or y.FileExt = "png" Or y.FileExt = "tiff" Or y.FileExt = "gif" Or y.FileExt = "jpeg" Then

                    ''
                    '' two sizes to be extracted / created for thumbs | 64-90 || 256 for jumbos
                    ''
                    '' going to change this to three and going to try for 48x48 for 'larges/mediums'
                    '' 
                    Try
                        Dim thumA As Image = Bitmap.FromFile(y.FullPath, True)
                        Dim tThumbA64 As Image = thumA.GetThumbnailImage(64, 64, AddressOf GetThumbCallBackAbort, IntPtr.Zero)
                        Dim imgC As Image = tThumbA64.Clone
                        tThumbA64.Dispose()
                        thumA.Dispose()
                        y.mdThumb = imgC
                        Me.imlXL.Images.Add(imgC)
                        Me.imlXL.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try

                    Try
                        Dim thumB As Image = Bitmap.FromFile(y.FullPath, True)
                        Dim tThumbB256 As Image = thumB.GetThumbnailImage(256, 256, AddressOf GetThumbCallBackAbort, IntPtr.Zero)
                        Dim imgD As Image = tThumbB256.Clone
                        thumB.Dispose()
                        tThumbB256.Dispose()
                        y.lgThumb = imgD
                        Me.imlJB.Images.Add(imgD)
                        Me.imlJB.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try

                    Try
                        Dim thumC As Image = Bitmap.FromFile(y.FullPath, True)
                        Dim tthumbC48 As Image = thumC.GetThumbnailImage(48, 48, AddressOf GetThumbCallBackAbort, IntPtr.Zero)
                        Dim imgE As Image = tthumbC48.Clone
                        thumC.Dispose()
                        tthumbC48.Dispose()
                        y.smThumb = imgE
                        Me.imlLG.Images.Add(imgE)
                        Me.imlLG.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try
                    lvItem.Text = y.FileName

                    lvItem.Tag = y.FullPath

                    lvItem.SubItems.Add(y.FileExt.ToString)
                    lvItem.SubItems.Add(y.FileSize.ToString)
                    lvItem.SubItems.Add(y.DateModified.ToString)
                    lvItem.ImageKey = y.IconKey.ToString

                    lsCollection.Add(lvItem)

                ElseIf y.FileExt <> "jpg" Or y.FileExt <> "bmp" Or y.FileExt <> "png" Or y.FileExt <> "tiff" Or y.FileExt <> "gif" Or y.FileExt <> "jpeg" Then
                    Try
                        y.lgIcon = GetLarges(y.FullPath) '' 48 
                        Me.imlLG.Images.Add(y.lgIcon)
                        Me.imlLG.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try

                    Try
                        y.mdIcon = GetMediums(y.FullPath) '' 64-90 || thumb / bmp depending on extension 
                        Me.imlXL.Images.Add(y.mdIcon)
                        Me.imlXL.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try

                    Try
                        y.jbIcon = GetJumbos(y.FullPath) '' 256 || thumb / bmp depending on extension
                        Me.imlJB.Images.Add(y.jbIcon)
                        Me.imlJB.Images.SetKeyName(idx, y.IconKey.ToString)
                    Catch ex As Exception
                        '' point to default error icon here.
                        '' 
                    End Try


                    '' ignore common system files 
                    ''
                    '' NOT A COMPREHENSIVE EXCLUSION LIST
                    '' JUST COMMON
                    '' 
                    ''source: 
                    '' http://www.file-extensions.org/filetype/extension/name/system-files
                    ''
                    ''


                    If y.FileExt <> "db" Then
                        lvItem.Text = y.FileName

                        lvItem.Tag = y.FullPath

                        lvItem.SubItems.Add(y.FileExt.ToString)
                        lvItem.SubItems.Add(y.FileSize.ToString)
                        lvItem.SubItems.Add(y.DateModified.ToString)
                        lvItem.ImageKey = y.IconKey.ToString

                        lsCollection.Add(lvItem)
                    End If
                End If
                    ''
                    ''
                    '' end choose file extension

                    



            Next

            Dim xx
            For Each xx In dirNFO.GetDirectories()
                Dim yy As New DirObject
                Dim lvItem2 As New ListViewItem
                idx += 1
                yy.FullPath = xx.FullName
                yy.DateCreated = xx.CreationTime
                yy.DateModified = xx.LastWriteTime
                yy.IconKey = xx.FullName ''edit
                Try
                    yy.smIcon = GetSmalls(yy.FullPath)
                    Me.imlSM.Images.Add(yy.smIcon)
                    Me.imlSM.Images.SetKeyName(idx, yy.FullPath.ToString)
                Catch ex As Exception
                    '' point to default error icon here
                    '' 
                End Try

                Try
                    yy.lgIcon = GetMediums(yy.FullPath)
                    Me.imlLG.Images.Add(yy.lgIcon)
                    Me.imlLG.Images.SetKeyName(idx, yy.FullPath.ToString)
                Catch ex As Exception
                    '' point to default error icon here
                    '' 
                End Try

                Try
                    yy.lgIcon = GetLarges(yy.FullPath)
                    Me.imlXL.Images.Add(yy.lgIcon)
                    Me.imlXL.Images.SetKeyName(idx, yy.FullPath.ToString)
                Catch ex As Exception
                    '' point to default error icon here
                    '' 
                End Try

                Try
                    yy.jbIcon = GetJumbos(yy.FullPath)
                    Me.imlJB.Images.Add(yy.jbIcon)
                    Me.imlJB.Images.SetKeyName(idx, yy.FullPath.ToString)
                Catch ex As Exception
                    '' point to default error icon here
                    '' 
                End Try

                lvItem2.Text = xx.Name.ToString   '' edit 
                lvItem2.Tag = xx.FullName '' edit 
                lvItem2.ToolTipText = xx.FullName

                lvItem2.SubItems.Add("File Folder")
                lvItem2.SubItems.Add(" ")
                lvItem2.SubItems.Add(yy.DateModified)
                lvItem2.ImageKey = yy.IconKey

                lsCollection.Add(lvItem2)
                Main.tsProgress.PerformStep()
            Next
            lsControl.Cursor = Cursors.Default
        End If


    End Sub

    Public Function GetThumbCallBackAbort() As Boolean
        Return False
    End Function

    Public Sub GenerateFileObj(ByVal TargetPath As String)

        Dim y As New FileObject
        y.FullPath = TargetPath
        y.FileName = SplitApartFileName(TargetPath)
        y.FileExt = SplitApartFileExt(TargetPath)
        y.IconKey = (y.FileName & "." & y.FileExt)

        y.smIcon = GetSmalls(y.FullPath)
        y.lgIcon = GetLarges(y.FullPath)
        y.mdIcon = GetMediums(y.FullPath)
        y.jbIcon = GetJumbos(y.FullPath)

        _ReconFile = y

        Me.FileOBJ = _ReconFile

    End Sub


    Public Property FileOBJ As FileObject
        Get
            Return _ReconFile
        End Get
        Set(value As FileObject)
            _ReconFile = value
        End Set
    End Property



    Public Function SplitFolderName(ByVal FullPath As String)
        Try
            Dim BeginAr = FullPath.ToString.Split("\")
            Dim FileNameAndExt = BeginAr(BeginAr.Count - 1) '' last index of items minus one. 
            Dim LastTwo = FileNameAndExt.ToString.Split(".")
            Dim FolderName
            FolderName = LastTwo(0) '' file name is going to be position 0 
            Return FolderName
        Catch ex As Exception
            '' fail it here
            'MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Split Folder Name Error")
        End Try

    End Function

    Public Function SplitApartFileName(ByVal FullPath As String)
        Try
            Dim BeginAr() = FullPath.ToString.Split("\")
            Dim FileNameAndExt = BeginAr(BeginAr.Length - 1) '' last index of items minus one. 
            Dim LastTwo = FileNameAndExt.ToString.Split(".")
            Dim FileNameOnly As String
            FileNameOnly = LastTwo(0) '' file name is going to be position 0 
            Return FileNameOnly
        Catch ex As Exception
            '' fail it here
            'MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Split File Name Error")
        End Try
    End Function

    Public Function SplitApartFileExt(ByVal FullPath As String)
        Try
            Dim BeginAr() = FullPath.ToString.Split("\")
            Dim FileNameAndExt = BeginAr(BeginAr.Length - 1) '' last index of items minus one. 
            Dim LastTwo = FileNameAndExt.ToString.Split(".")
            Dim FileExt As String
            FileExt = LastTwo(1) '' file name is going to be position 1
            Return FileExt
        Catch ex As Exception
            '' fail it here
            'MsgBox(ex.InnerException.ToString, MsgBoxStyle.Critical, "Split File Ext Error")
        End Try

    End Function


    Public Function GetJumbos(ByVal path As String)

        Dim iml As MimicSysImgList.IImageList
        Dim riid As Guid = New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
        Dim lresult As IntPtr
        lresult = SHGetImageList(SHIL_JUMBO, riid, iml)
        Dim hImg As IntPtr
        Dim shinfo As SHFILEINFO = New SHFILEINFO
        Dim icoIDX = shinfo.iIcon
        Dim hIcon = IntPtr.Zero
        hImg = SHGetFileInfo(path, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON)
        Dim hres = iml.GetIcon(shinfo.iIcon, 1, hIcon)
        Dim ico As Icon = System.Drawing.Icon.FromHandle(hIcon)
        Dim ico2 As Icon = ico.Clone()
        ico.Dispose()
        Return ico2
    End Function

    Public Function GetLarges(ByVal Path As String)
        Dim iml As MimicSysImgList.IImageList
        Dim riid As Guid = New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
        Dim lresult As IntPtr
        lresult = SHGetImageList(SHIL_EXTRALARGE, riid, iml)
        Dim hImg As IntPtr
        Dim shinfo As SHFILEINFO = New SHFILEINFO
        Dim icoIDX = shinfo.iIcon
        Dim hIcon = IntPtr.Zero
        hImg = SHGetFileInfo(Path, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON)
        Dim hres = iml.GetIcon(shinfo.iIcon, 1, hIcon)
        Dim ico As Icon = System.Drawing.Icon.FromHandle(hIcon)
        Dim ico2 As Icon = ico.Clone()
        ico.Dispose()
        Return ico2
    End Function

    Public Function GetSmalls(ByVal path As String)
        Dim iml As MimicSysImgList.IImageList
        Dim riid As Guid = New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
        Dim lresult As IntPtr
        lresult = SHGetImageList(SHIL_SMALL, riid, iml)
        Dim hImg As IntPtr
        Dim shinfo As SHFILEINFO = New SHFILEINFO
        Dim icoIDX = shinfo.iIcon
        Dim hIcon = IntPtr.Zero
        hImg = SHGetFileInfo(path, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON)
        Dim hres = iml.GetIcon(shinfo.iIcon, 1, hIcon)
        Dim ico As Icon = System.Drawing.Icon.FromHandle(hIcon)
        Dim ico2 As Icon = ico.Clone()
        ico.Dispose()
        Return ico2
    End Function

    Public Function GetMediums(ByVal path As String)
        Dim iml As MimicSysImgList.IImageList
        Dim riid As Guid = New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
        Dim lresult As IntPtr
        lresult = SHGetImageList(SHIL_LARGE, riid, iml)
        Dim hImg As IntPtr
        Dim shinfo As SHFILEINFO = New SHFILEINFO
        Dim icoIDX = shinfo.iIcon
        Dim hIcon = IntPtr.Zero
        hImg = SHGetFileInfo(path, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON)
        Dim hres = iml.GetIcon(shinfo.iIcon, 1, hIcon)
        Dim ico As Icon = System.Drawing.Icon.FromHandle(hIcon)
        Dim ico2 As Icon = ico.Clone()
        ico.Dispose()
        Return ico2
    End Function


    Public Class ListSortAscending
        Implements IComparer

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim ret = -1
            ret = [String].Compare(CType(x, ListViewItem).Text, (CType(y, ListViewItem).Text))
            Return ret
        End Function
    End Class

    Public Class ListSortDescending
        Implements IComparer

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim ret = -1
            ret = [String].Compare(CType(y, ListViewItem).Text, (CType(x, ListViewItem).Text))
            Return ret
        End Function
    End Class

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
                Dim linkPath As IWshShortcut = shl.CreateShortcut(fullPath & "\" & ShortCutName & ".lnk")
                linkPath.TargetPath = SourceDestination
                linkPath.Save()
            Catch ex As Exception
                '' call to error log from here on fail
                '' 
            End Try

        End Sub
        Public Sub New(ByVal ShortCutName As String, ByVal SourceDestination As String, ByVal WhereToCreate As String)
            Try
                Dim strPath As String = WhereToCreate
                Dim fullPath As String = (strPath)
                Dim shl As New WshShell
                Dim linkPath As IWshShortcut = shl.CreateShortcut(fullPath & "\" & ShortCutName & ".lnk")
                linkPath.TargetPath = SourceDestination
                linkPath.Save()
            Catch ex As Exception
                '' call to error log from here on fail
                '' 
            End Try
        End Sub

    End Class


    Public Class GetOpensWithNames '' registry operations

        Private OpensWithList As ArrayList

        Public Sub New(ByVal FileExt As String)
            Try
                OpensWithList = New ArrayList
                Dim rk As RegistryKey = Registry.CurrentUser
                Dim subKey As RegistryKey = rk.OpenSubKey("Software")
                Dim subKey2 As RegistryKey = subKey.OpenSubKey("Microsoft")
                Dim subKey3 As RegistryKey = subKey2.OpenSubKey("Windows")
                Dim subKey4 As RegistryKey = subKey3.OpenSubKey("CurrentVersion")
                Dim subKey5 As RegistryKey = subKey4.OpenSubKey("Explorer")
                Dim subKey6 As RegistryKey = subKey5.OpenSubKey("FileExts")
                Dim subKey7 As RegistryKey = subKey6.OpenSubKey("." & FileExt & "")
                Dim subKey8 As RegistryKey = subKey7.OpenSubKey("OpenWithList")

                Dim valueNames As String() = subKey8.GetValueNames()
                Dim arOBJs As New ArrayList
                Dim s As String = ""
                Dim x
                For Each x In valueNames

                    Dim obj1 As Object = subKey8.GetValue(x.ToString, "Nothing.")
                    If obj1 <> Nothing Then
                        If obj1.ToString.Length > 1 Then
                            '' also need a method here to look for "." in the string and only add them. 
                            '' 
                            Dim FoundNotFound As Integer = InStr(obj1.ToString, ".", Microsoft.VisualBasic.CompareMethod.Text)
                            If FoundNotFound <> 0 Then
                                Dim name As String = obj1.ToString
                                OpensWithList.Add(name)
                            ElseIf obj1 = Nothing Then
                                '' don't add it to the list. 
                                ''
                            End If
                        End If
                    End If

                Next
            Catch ex As Exception
                '' fail it here for registry operations
                '' 

            End Try
        End Sub

        Public Property OpensWith As ArrayList
            Get
                Return OpensWithList
            End Get
            Set(value As ArrayList)
                OpensWithList = value
            End Set
        End Property

    End Class


#End Region

    
    Private Sub lsControl_MouseUp(sender As Object, e As MouseEventArgs) Handles lsControl.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim y As ListViewItem
            y = Me.lsControl.GetItemAt(e.X, e.Y)

            If y IsNot Nothing Then
                If y.Tag = Sel_TEXT Then
                    Me.lsControl.HoverSelection = False
                    'RefreshUI()
                    Itm_IDX = y.Index
                End If
                If y.Tag <> Sel_TEXT Then
                    Dest_TEXT = y.Tag
                    Dim destTarget As String = DropTarget_F_or_F(Dest_TEXT)
                    If destTarget <> "file" Then
                        Dim targetF_or_F As String = Strip_Off_File_Or_Folder_Name(Sel_TEXT)
                        Dim targetType As String = Is_it_a_file_or_folder(Sel_TEXT)
                        Select Case targetType
                            Case "file"
                                Try
                                    System.IO.File.Move(Sel_TEXT, Dest_TEXT & "\" & targetF_or_F)
                                    'Me.lsControl.Items(Itm_IDX).Remove()
                                    Me.lsControl.HoverSelection = False
                                    'RefreshUI()
                                    GetDirInfo(_CurrentPath)
                                Catch ex As Exception
                                    'RefreshUI()
                                    'GetDirInfo(_CurrentPath)
                                    Me.lsControl.HoverSelection = False
                                End Try

                            Case "folder"
                                Try
                                    System.IO.Directory.Move(Sel_TEXT, Dest_TEXT & "\" & targetF_or_F)
                                    'Me.lsControl.Items(Itm_IDX).Remove()
                                    Me.lsControl.HoverSelection = False
                                    'RefreshUI()
                                    GetDirInfo(_CurrentPath)
                                    Exit Select
                                Catch ex As Exception
                                    'RefreshUI()
                                    'GetDirInfo(_CurrentPath)
                                    Me.lsControl.HoverSelection = False
                                End Try

                        End Select
                    End If
                End If
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            Dim y As ListViewItem = Me.lsControl.GetItemAt(e.X, e.Y)
            If y IsNot Nothing Then
                Dest_TEXT = ""
                Dest_TEXT = y.Tag
            End If
            If y Is Nothing Then
                'RefreshUI()
            End If
        End If

    End Sub

     
End Class
