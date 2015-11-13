Imports System.Drawing.Icon
Imports System.IO
Imports System.Text
Imports System.IO.File
Imports System.IO.Directory
Imports System.Drawing

Public Class imgLSTS
#Region "Notes Design"
    ''
    ''
    '' this class will expose 4 image lists as readonly properties of objects
    '' 
    '' imgList16   {depth:32bit,size:16x16,keys:3 char file extension}
    '' imgList32   {depth:32bit,size:32x32,keys:3 char file extension}
    '' imgList64   {depth:32bit,size:64x64,keys:3 char file extension}
    '' imgList128  {depth:32bit,size:128x128,keys:3 char file extension}
    ''

    '' on constructor with [Directory] argument will launch and extract icons to all four image lists
    '' 
    '' 




#End Region

#Region "Private Vars "
    Private _ImgLst16 As ImageList
    Private _ImgLst32 As ImageList
    Private _ImgLst64 As ImageList
    Private _ImgLst128 As ImageList
    Private _ImgLst256 As ImageList
    Private _FilePath As String = ""
    Private _ExtractedIcon As Icon
#End Region

#Region "ReadOnly Properties"
    Public ReadOnly Property ImgList16() As ImageList
        Get
            Return _ImgLst16
        End Get
    End Property
    Public ReadOnly Property ImgList32() As ImageList
        Get
            Return _ImgLst32
        End Get
    End Property
    Public ReadOnly Property ImgList64() As ImageList
        Get
            Return _ImgLst64
        End Get
    End Property
    Public ReadOnly Property ImgList128() As ImageList
        Get
            Return _ImgLst128
        End Get
    End Property
    Public ReadOnly Property ImgList256() As ImageList
        Get
            Return _ImgLst256
        End Get
    End Property
#End Region

#Region "Constructor"

    Public Sub New(ByVal DirectoryName As String)
        Try

            _ImgLst16 = New ImageList
            _ImgLst32 = New ImageList
            _ImgLst64 = New ImageList
            _ImgLst128 = New ImageList
            _ImgLst256 = New ImageList

            _ImgLst16.ColorDepth = ColorDepth.Depth32Bit
            Dim point As New Point(16, 16)
            _ImgLst16.ImageSize = point

            _ImgLst32.ColorDepth = ColorDepth.Depth32Bit
            Dim pnt2 As New Point(32, 32)
            _ImgLst32.ImageSize = pnt2

            _ImgLst64.ColorDepth = ColorDepth.Depth32Bit
            Dim pnt3 As New Point(64, 64)
            _ImgLst64.ImageSize = pnt3

            _ImgLst128.ColorDepth = ColorDepth.Depth32Bit
            Dim pnt4 As New Point(128, 128)
            _ImgLst128.ImageSize = pnt4

            _ImgLst256.ColorDepth = ColorDepth.Depth32Bit
            Dim pnt5 As New Point(256, 256)
            _ImgLst256.ImageSize = pnt5

            Dim dirNfo As DirectoryInfo = New DirectoryInfo(DirectoryName)


            '' the loop to pull out all icons at once.
            '' and populate lists respectively
            '' 
            '' counter for loop 
            Dim flNFO As FileInfo
            Dim POS As Integer = 0
            Dim icoCnt As Integer = 0
            Dim bb
            For Each flNFO In dirNfo.GetFiles("*.*")
                POS += 1
                Dim strSplit = Split(flNFO.FullName, "\")
                For Each bb In strSplit
                    icoCnt += 1
                Next
                Dim subSplit = Split(strSplit(icoCnt - 1), ".")
                Dim FileExt = subSplit(1)
                Dim FileName = subSplit(0)
                Dim ico As Icon = GetIcon(flNFO.FullName)

                _ImgLst16.Images.Add(FileExt, ico)
                _ImgLst16.Images.SetKeyName((POS - 1), FileExt)
                _ImgLst32.Images.Add(FileExt, ico)
                _ImgLst32.Images.SetKeyName((POS - 1), FileExt)
                _ImgLst64.Images.Add(FileExt, ico)
                _ImgLst64.Images.SetKeyName((POS - 1), FileExt)
                _ImgLst128.Images.Add(FileExt, ico)
                _ImgLst128.Images.SetKeyName((POS - 1), FileExt)
                _ImgLst256.Images.Add(FileExt, ico)
                _ImgLst256.Images.SetKeyName((POS - 1), FileExt)
            Next


        Catch ex As Exception
            '' catch and write to error log here 
            ''
        End Try

    End Sub

#End Region
#Region "Private Functions"
    Function GetIcon(ByVal FilePath As String)
        _ExtractedIcon = ExtractAssociatedIcon(FilePath)
        Return _ExtractedIcon
    End Function
#End Region

#Region "Destruction"
    Public Sub DestroyMe()
        Me._ImgLst256 = Nothing
        Me._ImgLst128 = Nothing
        Me._ImgLst64 = Nothing
        Me._ImgLst32 = Nothing
        Me._ImgLst16 = Nothing
    End Sub
#End Region
End Class

