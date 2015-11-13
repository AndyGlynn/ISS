Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Drawing
Imports System.Drawing.imaging
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Collections.Generic
Imports Microsoft.Reporting.Winforms

Public Class SALES_REPORTING_LOGIC
    Implements IDisposable
    '
    ' 1) Run the query
    ' 2) get the xml data to a flat file
    ' 3) dump the flat file to share
    ' 4) generate the rdl from the flat file
    ' 5) render report
    ' 6) select printer
    ' 7) print report / or don't print report
    '

    Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

    Private m_CurrentPageIndex As Integer
    Private m_streams As IList(Of Stream)
    '' NO RESULTS
    '' __________
    Public Sub GetSQL_TO_XML_Information_NoResults(ByVal ApptDate As String)
        Try
            Dim cmdRUN As SqlCommand = New SqlCommand("dbo.GetNoResultsReportXML", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@ApptDate", ApptDate)
            cmdRUN.Parameters.Add(param1)
            cmdRUN.CommandType = CommandType.StoredProcedure
            cnn.Open()
            cmdRUN.ExecuteNonQuery()
            cnn.Close()
            '' after this point, the xml data should be in the first row of the test.dbo.InsertXMLData table.
            '' 

            GetXML_TO_TEXT_FILE_NoResults()
            Me.Run(ApptDate)

        Catch ex As Exception
            cnn.Close()
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetSQL_TO_XML_Information_NoResults")
        End Try
    End Sub

    Public Sub GetSQL_TO_XML_Information_DailyPerf(ByVal ApptDate As String)

        Try

            Dim cmdRUN As SqlCommand = New SqlCommand("dbo.SDPReportXML", cnn)
            Dim param1 As SqlParameter = New SqlParameter("@Date", ApptDate)
            cmdRUN.Parameters.Add(param1)
            cmdRUN.CommandType = CommandType.StoredProcedure
            cnn.Open()
            cmdRUN.ExecuteNonQuery()
            cnn.Close()
            '' now get it to a flat file
            GetXML_TO_TEXT_FILE_DailyPerf()
            Me.Run_DailyPerf(ApptDate)

        Catch ex As Exception
            cnn.Close()
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetSQL_TO_XML_Information_DailyPerf")

        End Try

    End Sub

    Public Sub GetXML_TO_TEXT_FILE_NoResults()
        Try
            Dim day As String = Date.Now.Day.ToString
            Dim month As String = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString
            Dim putbacktogether As String = month & "-" & day & "-" & year
            If System.IO.File.Exists("\\ekg1\iss\Reports\DataSets\NoResults_" + putbacktogether.ToString + ".xml") = True Then
                System.IO.File.Delete("\\ekg1\iss\Reports\DataSets\NoResults_" + putbacktogether.ToString + ".xml")
            End If
            Dim fs As New StreamWriter("\\ekg1\iss\Reports\DataSets\NoResults_" + putbacktogether.ToString + ".xml") '' directory needs to point to a 'iss share'
            '\\ekg1\iss\Reports\DataSets\'Generate a Name for the xml file 
            Dim cmdGET As SqlCommand = New SqlCommand("Select XML_Val from iss.dbo.InsertXMLForReporting", cnn)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            fs.WriteLine("<ROOT>") ' had to add additional 'root' tag because SQL doesn't generate it for XML, PATH argument.
            While r1.Read
                fs.WriteLine(r1.Item("XML_Val").ToString)
            End While
            r1.Close()
            cnn.Close()
            fs.WriteLine("</ROOT>")
            fs.Flush()
            fs.Close()
        Catch ex As Exception
            cnn.Close()
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetXML_TO_TEXT_FILE_NoResults")
        End Try

        Try
            Dim cmdDEL As SqlCommand = New SqlCommand("DELETE iss.dbo.InsertXMLForReporting", cnn)

            cnn.Open()
            cmdDEL.ExecuteNonQuery()
            cnn.Close()



        Catch ex As Exception
            cnn.Close()
        End Try


    End Sub
    Public Sub GetXML_TO_TEXT_FILE_DailyPerf()
        Try
            Dim day As String = Date.Now.Day.ToString
            Dim month As String = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString
            Dim putbacktogether As String = month & "-" & day & "-" & year
            If System.IO.File.Exists("\\ekg1\iss\Reports\DataSets\report_" + putbacktogether.ToString + ".xml") = True Then
                System.IO.File.Delete("\\ekg1\iss\Reports\DataSets\report_" + putbacktogether.ToString + ".xml")
            End If
            Dim fs As New StreamWriter("\\ekg1\iss\Reports\DataSets\report_" + putbacktogether.ToString + ".xml") '' directory needs to point to a 'iss share'
            '\\ekg1\iss\Reports\DataSets\'Generate a Name for the xml file 
            Dim cmdGET As SqlCommand = New SqlCommand("Select XML_Val from iss.dbo.InsertXMLForReporting", cnn)
            cnn.Open()
            Dim r1 As SqlDataReader
            r1 = cmdGET.ExecuteReader
            fs.WriteLine("<ROOT>") ' had to add additional 'root' tag because SQL doesn't generate it for XML, PATH argument.
            While r1.Read
                fs.WriteLine(r1.Item("XML_Val").ToString)
            End While
            r1.Close()
            cnn.Close()
            fs.WriteLine("</ROOT>")
            fs.Flush()
            fs.Close()
        Catch ex As Exception
            cnn.Close()
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetXML_TO_TEXT_FILE_DailyPerf")
        End Try

        Try
            Dim cmdDEL As SqlCommand = New SqlCommand("DELETE iss.dbo.InsertXMLForReporting", cnn)

            cnn.Open()
            cmdDEL.ExecuteNonQuery()
            cnn.Close()



        Catch ex As Exception
            cnn.Close()
        End Try


    End Sub

    Public Function Generate_DATASET(ByVal FileName As String) As DataTable
        Dim dataset As New DataSet
        Try
            dataset.ReadXml(FileName)
            Return DataSet.Tables(0)
        Catch ex As Exception
            Return DataSet.Tables(0)
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal FileName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Generate_DATASET")
        End Try

    End Function


    Private Function CreateStream(ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean) As Stream
        Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)

        Try
            'Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)
            m_streams.Add(stream)
            Return stream
        Catch ex As Exception
            Return Stream
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CreateStream")
        End Try

    End Function

    'Private Sub Export(ByVal Report As LocalReport)
    '    '' this may or may not need to reset depending on device selected.
    '    '' also, output may be a selectable option, like "PDF" or "EXCEL"
    '    '' for rendering.
    '    '' Depends if client machine will accept it, AND if local rendering supports it.
    '    '' 
    '    Try
    '        Dim deviceInfo As String = _
    '        "<DeviceInfo>" + _
    '        "<OutPutFormat>EMF</OutPutFormat>" + _
    '        "<PageWidth>8.5in</PageWidth>" + _
    '        "<PageHeight>11in</PageHeight>" + _
    '        "<MarginTop>1in</MarginTop>" + _
    '        "<MarginLeft>.25in</MarginLeft>" + _
    '        "<MarginRight>.25in</MarginRight>" + _
    '        "<MarginBottom>1in</MarginBottom>" + _
    '        "</DeviceInfo>"
    '        Dim warnings() As Warning = Nothing
    '        Report.EnableExternalImages = True
    '        m_streams = New List(Of Stream)()
    '        Report.Render("Image", deviceInfo, AddressOf CreateStream, warnings)
    '        Dim stream As Stream
    '        For Each stream In m_streams
    '            stream.Position = 0
    '        Next
    '    Catch ex As Exception
    '        Dim errp As New ErrorLogFlatFile
    '        errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Report As LocalReport", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Export")
    '    End Try

    'End Sub
    Private Sub PrintPage(ByVal Sender As Object, ByVal ev As PrintPageEventArgs)
        Try
            Dim pageImage As New Metafile(m_streams(m_CurrentPageIndex))
            ev.Graphics.DrawImage(pageImage, ev.PageBounds)
            m_CurrentPageIndex += 1
            ev.HasMorePages = (m_CurrentPageIndex < m_streams.Count)


            '' clean up emf's 
            '' AFTER the print button has been called.
            '' will need another clean up call in 
            '' the select case statement as well. 
            '' 



        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Sender As Object, ByVal ev As PrintPageEventArgs", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "PrintPage")
        End Try

    End Sub

    Private Sub print()
        'Const printername As String = _
        '" "
        Try
            Dim day As String = Date.Now.Day.ToString
            Dim month As String = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString
            Dim putbacktogether As String = month & "-" & day & "-" & year

            EditAutoNotesLists.Cursor = Cursors.WaitCursor
            If m_streams Is Nothing Or m_streams.Count = 0 Then
                Return
            End If

            Dim printDoc As New PrintDocument()
            'printDoc.PrinterSettings.PrinterName = printername


            If Not printDoc.PrinterSettings.IsValid Then
                'Dim msg As String = String.Format("Can't find printer ""{0}"".", printername)
                'MsgBox(msg.ToString)
                AddHandler printDoc.PrintPage, AddressOf PrintPage
                Dim pr1 As New PrintDialog
                pr1.Document = printDoc
                pr1.Document.DefaultPageSettings.Landscape = True
                printDoc.DefaultPageSettings.Margins.Bottom = 1
                printDoc.DefaultPageSettings.Margins.Top = 1
                printDoc.DefaultPageSettings.Margins.Left = 1
                printDoc.DefaultPageSettings.Margins.Right = 1
                printDoc.PrinterSettings.PrintToFile = False
                printDoc.DefaultPageSettings.Landscape = True
                printDoc.DocumentName = "No Results" + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
                'pr1.ShowDialog()
                Dim dlg As DialogResult = pr1.ShowDialog
                Select Case dlg
                    Case Is = DialogResult.Cancel
                        'Dim p As FileInfo
                        'Dim dir As DirectoryInfo = New DirectoryInfo("C:\Documents and Settings\xxclayxx\Desktop\Current 7-19-2008 445 PM\New Revision\New Revision\bin\Debug")
                        'For Each p In dir.GetFiles("*.emf", SearchOption.TopDirectoryOnly)
                        '    System.IO.File.Delete(p.FullName)
                        'Next
                        '' cannot clean up here.
                        '' the stream still has the object locked.
                        ''

                        Exit Select
                    Case Is = DialogResult.OK
                        printDoc.Print()
                        Exit Select
                End Select
                'printDoc.Print()
                Main.Cursor = Cursors.Default
                Return
            End If

            '' pipe it to the print preview

            AddHandler printDoc.PrintPage, AddressOf PrintPage

            Dim pr As New PrintDialog
            pr.Document = printDoc
            printDoc.DefaultPageSettings.Margins.Bottom = 1
            printDoc.DefaultPageSettings.Margins.Top = 1
            printDoc.DefaultPageSettings.Margins.Left = 1
            printDoc.DefaultPageSettings.Margins.Right = 1
            printDoc.PrinterSettings.PrintToFile = False
            printDoc.DefaultPageSettings.Landscape = True
            printDoc.DocumentName = "No Results " + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
            'printDoc.Print()
            Dim dlg2 As DialogResult = pr.ShowDialog
            Select Case dlg2
                '' 3 = 'Abort'
                '' 1 = 'OK'
                '' 2 = 'Cancel'
                '' 5 = 'Ignore'
                '' 7 = 'No'
                '' 0 = 'None'
                '' 4 = 'Retry'
                '' 6 = 'Yes'
                Case Is = DialogResult.Cancel
                    Exit Select
                Case Is = DialogResult.OK
                    printDoc.Print()

                    Exit Select
            End Select
            Main.Cursor = Cursors.Default
        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Print")

        End Try

    End Sub

    Private Sub Run(ByVal ApptDate As String)
        Try
            Dim day As String = Date.Now.Day.ToString
            Dim month As String = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString
            Dim putbacktogether As String = month & "-" & day & "-" & year

            'Dim report As LocalReport = New LocalReport()
            If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\NoResults_" + putbacktogether + ".rdl") = True Then
                System.IO.File.Delete("\\ekg1\iss\reports\RDL Files\NoResults_" + putbacktogether + ".rdl")
            End If

            Dim y As New GENERATE_RDL
            y.GenerateFieldsList(ApptDate)
            y.GenerateRDL()


            '  report.ReportPath = "\\ekg1\iss\Reports\RDL Files\NoResults_" + putbacktogether + ".rdl"
            '' \\ekg1\iss\Reports\RDL Files\'Name of Generated RDL file."
            'report.DataSources.Add(New ReportDataSource("XML", Me.Generate_DATASET("\\ekg1\iss\Reports\DataSets\NoResults_" + putbacktogether + ".xml")))
            '' \\ekg1\iss\Reports\DataSets\'Name of generated XML file'
            ' Export(report)
            m_CurrentPageIndex = 0
            print()


            '' its not the stream object....
            '' its the actual thread handle of the program that is locking them down
            '' after the 'render' call is made....
            '' 

            ' Dim p As FileInfo
            ' Dim dir As DirectoryInfo = New DirectoryInfo("C:\Documents and Settings\xxclayxx\Desktop\Current 7-19-2008 445 PM\New Revision\New Revision\bin\Debug")
            ' For Each p In dir.GetFiles("*.emf", SearchOption.TopDirectoryOnly)
            '    System.IO.File.Delete(p.FullName)
            ' Next

        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Run")

        End Try


    End Sub
    Public Sub Run_DailyPerf(ByVal ApptDate As String)
        Try
            Dim day As String = Date.Now.Day.ToString
            Dim month As String = Date.Now.Month.ToString
            Dim year As String = Date.Now.Year.ToString
            Dim putbacktogether As String = month & "-" & day & "-" & year

            ' Dim report As LocalReport = New LocalReport()
            If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\RDL_" + putbacktogether + ".rdl") = True Then
                System.IO.File.Delete("\\ekg1\iss\reports\RDL Files\RDL_" + putbacktogether + ".rdl")
            End If

            Dim y As New GENERATE_RDL ' No. Generate RDL will have to reflect the new query.
            '' will need a new generate RDL for daily performance.
            '' 
            y.GenerateFieldsList(ApptDate)
            y.GenerateRDL()


            ' report.ReportPath = "\\ekg1\iss\Reports\RDL Files\RDL_" + putbacktogether + ".rdl"
            '' \\ekg1\iss\Reports\RDL Files\'Name of Generated RDL file."
            ' report.DataSources.Add(New ReportDataSource("XML", Me.Generate_DATASET("\\ekg1\iss\Reports\DataSets\report_" + putbacktogether + ".xml")))
            '' \\ekg1\iss\Reports\DataSets\'Name of generated XML file'
            '  Export(report)
            m_CurrentPageIndex = 0
            print()
        Catch ex As Exception
            Dim errp As New ErrorLogFlatFile
            errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Run_DailyPerf")

        End Try

    End Sub


    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
            End If

            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        If Not (m_streams Is Nothing) Then
            Dim stream As Stream
            For Each stream In m_streams
                stream.Close() '' close each stream and clean out memory
            Next
            m_streams = Nothing
        End If
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region


    Public Class GENERATE_RDL

        ''
        '' the end result of this class will not need a message box stating that the RDL file has been generated.
        '' also, this class and the "SQL_INFO" class will roll right into one method.
        '' IE: call class -> gererate RDL as needed -> generate XML Dataset -> populate report -> Render through local render mode -> select printer -> Print.
        '' complete background printing.
        '' Notes:
        '' when appt time comes out of sql it will have to be formatted to something more legible(?) 
        '' when value of textbox1 is added on render, the time will have to be formatted properly
        '' Field names will will have to be tweaked IE: "Contact1FirstName" -> "Contact Info."
        '' If no data is present to be dropped into the dataset, must have a catch for 'Nulls'
        ''  OR a blank report is generated with header image and datetime variable.
        '' 


        Private m_connection As SqlConnection
        Private m_connectString As String
        Private m_commandText As String = "Declare @ApptDate datetime " _
    & " set @ApptDate = '8/16/2008' " _
    & " select ID ,Contact1FirstName ,Contact1LastName ,Contact2FirstName ,Contact2LastName ,StAddress ,city ,state ,zip , " _
    & " product1 ,product2 , product3 ,appttime ,rep1 ,rep2 " _
    & " from iss.dbo.enterlead " _
    & " where Result =' ' or Result is null and ApptDate = @ApptDate " _
    & " order by ApptTime, Rep1 "
        Private m_fields As ArrayList

        Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)

        Public Sub GenerateFieldsList(ByVal ApptDate As String)
            Try
                Dim cmdGet As SqlCommand = New SqlCommand("Declare @ApptDate datetime " _
        & " set @ApptDate = '8/16/2008' " _
        & " select ID ,Contact1FirstName ,Contact1LastName ,Contact2FirstName ,Contact2LastName ,StAddress ,city ,state ,zip , " _
        & " product1 ,product2 , product3 ,appttime ,rep1 ,rep2 " _
        & " from iss.dbo.enterlead " _
        & " where Result =' ' or Result is null and ApptDate = @ApptDate " _
        & " order by ApptTime, Rep1 ", cnn)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGet.ExecuteReader(CommandBehavior.SchemaOnly)
                m_fields = New ArrayList
                Dim i As Integer
                For i = 0 To r1.FieldCount - 1
                    m_fields.Add(r1.GetName(i))
                Next
            Catch ex As Exception
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateFieldList")

            End Try

        End Sub
        Public Sub GenerateRDL()
            Try
                Main.Cursor = Cursors.WaitCursor
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\NoResults_" + putbacktogether + ".rdl") = True Then '' point this to an iss share
                    '' \\ekg1\iss\Reports\RDL Files\'Name of RDL Generated'.rdl
                    System.IO.File.Delete("\\ekg1\iss\Reports\RDL Files\NoResults_" + putbacktogether + ".rdl") '' point this to an iss share
                End If
                Dim stream As FileStream
                stream = File.OpenWrite("\\ekg1\iss\Reports\RDL Files\NoResults_" + putbacktogether + ".rdl") '' point this to an 'iss share'
                Dim writer As New XmlTextWriter(stream, Encoding.UTF8)
                writer.Formatting = Formatting.Indented
                'Report Element
                writer.WriteProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
                writer.WriteStartElement("Report")
                writer.WriteAttributeString("xmlns", Nothing, "http://schemas.microsoft.com/sqlserver/reporting/2003/10/reportdefinition")
                writer.WriteElementString("Width", "6in")

                'DataSource Element
                writer.WriteStartElement("DataSources")
                writer.WriteStartElement("DataSource")
                writer.WriteAttributeString("Name", Nothing, "XML")
                writer.WriteStartElement("ConnectionProperties")
                writer.WriteElementString("DataProvider", "SQL")
                writer.WriteElementString("ConnectString", cnn.ConnectionString.ToString)
                writer.WriteEndElement() ' connection properties
                writer.WriteEndElement() ' datasource
                writer.WriteEndElement() ' datasources
                'dataset element
                writer.WriteStartElement("DataSets")
                writer.WriteStartElement("DataSet")
                writer.WriteAttributeString("Name", Nothing, "XML")
                'query element
                writer.WriteStartElement("Query")
                writer.WriteElementString("DataSourceName", "XML")
                writer.WriteElementString("CommandType", "Text")
                writer.WriteElementString("CommandText", m_commandText)
                writer.WriteElementString("Timeout", "30")
                writer.WriteEndElement() ' query
                writer.WriteStartElement("Fields")
                Dim fieldName As String = ""
                For Each fieldName In m_fields
                    writer.WriteStartElement("Field")
                    writer.WriteAttributeString("Name", Nothing, fieldName)
                    writer.WriteElementString("DataField", Nothing, fieldName)
                    writer.WriteEndElement() ' Field
                Next

                writer.WriteEndElement() 'Fields
                writer.WriteEndElement() 'Dataset
                writer.WriteEndElement() 'Datasets
                'body element
                writer.WriteStartElement("Body")
                '' image properties
                '' 
                writer.WriteElementString("Height", "5in")
                ' Report Items
                writer.WriteStartElement("ReportItems")
                ' Table Element
                writer.WriteStartElement("Table")
                writer.WriteAttributeString("Name", Nothing, "Table1")
                writer.WriteElementString("DataSetName", "XML")
                writer.WriteElementString("Top", ".5in")
                writer.WriteElementString("Left", ".5in")
                writer.WriteElementString("Height", ".5in")

                Dim str As String = CType(m_fields.Count * 1.5, String)
                writer.WriteElementString("Width", str + "in")
                ' table columns
                writer.WriteStartElement("TableColumns")
                For Each fieldName In m_fields
                    writer.WriteStartElement("TableColumn")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteEndElement() ' table column 
                Next
                writer.WriteEndElement() ' table columns
                ' header row
                writer.WriteStartElement("Header")
                writer.WriteStartElement("TableRows")
                writer.WriteStartElement("TableRow")
                writer.WriteElementString("Height", ".25in")
                writer.WriteStartElement("TableCells")

                For Each fieldName In m_fields
                    writer.WriteStartElement("TableCell")
                    writer.WriteStartElement("ReportItems")
                    'textbox
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, "Header" + fieldName)
                    writer.WriteStartElement("Style")
                    writer.WriteElementString("TextDecoration", "Underline")
                    writer.WriteEndElement() ' style
                    writer.WriteElementString("Top", "0in")
                    writer.WriteElementString("Left", "0in")
                    writer.WriteElementString("Height", ".5in")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteElementString("Value", fieldName)
                    writer.WriteEndElement() 'Textbox
                    writer.WriteEndElement() 'Report Items
                    writer.WriteEndElement() ' TableCell
                Next

                writer.WriteEndElement() 'tablecells
                writer.WriteEndElement() 'tablerow
                writer.WriteEndElement() 'tablerows
                writer.WriteEndElement() 'header
                'Details Row
                '
                writer.WriteStartElement("Details")
                writer.WriteStartElement("TableRows")
                writer.WriteStartElement("TableRow")
                writer.WriteElementString("Height", ".25in")
                writer.WriteStartElement("TableCells")
                For Each fieldName In m_fields
                    writer.WriteStartElement("TableCell")
                    writer.WriteStartElement("ReportItems")

                    ' textbox
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, fieldName)
                    writer.WriteStartElement("Style")
                    writer.WriteEndElement() ' style
                    writer.WriteElementString("Top", "0in")
                    writer.WriteElementString("Left", "0in")
                    writer.WriteElementString("Height", ".5in")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteElementString("Value", "=Fields!" + fieldName + ".Value")
                    writer.WriteElementString("HideDuplicates", "XML")
                    writer.WriteEndElement() 'TextBox
                    writer.WriteEndElement() 'Report ItemS
                    writer.WriteEndElement() ' TableCells
                Next

                writer.WriteEndElement() 'table cells
                writer.WriteEndElement() 'table row
                writer.WriteEndElement() ' table rows
                writer.WriteEndElement() ' details
                ' end details element and children 
                ' end table element 
                writer.WriteEndElement() 'table
                writer.WriteEndElement() ' report Items
                writer.WriteEndElement() ' body


                '' xtra stuff for images and the like.
                writer.WriteStartElement("PageHeader") ' write start element
                writer.WriteStartElement("ReportItems") ' write start element

                '' textbox 2
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox2")
                writer.WriteElementString("Left", "5.125in") '<Left>5.125in</Left>") ' write element string
                writer.WriteElementString("Top", "0.25in")
                'writer.WriteElementString("rd:DefaultName", "textbox2")
                writer.WriteElementString("ZIndex", "2")
                writer.WriteElementString("Width", "1.25in")
                '' style
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' </Style>
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", Date.Today.ToString)

                writer.WriteEndElement() ' </textbox> || textbox2

                '' textbox1
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox1")
                writer.WriteElementString("Left", "2.25in")
                writer.WriteElementString("Top", "0.25in")
                'writer.WriteElementString("rd:DefaultName", "textbox1")
                writer.WriteElementString("ZIndex", "1")
                writer.WriteElementString("Width", "2.25in")
                '' style
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("FontWeight", "700")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' </Style>
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", "EKG Construction Services, Inc.")
                writer.WriteEndElement() ' </Textbox> || textbox1

                '' image1
                writer.WriteStartElement("Image")
                writer.WriteAttributeString("Name", "image1")
                writer.WriteElementString("Sizing", "Fit")
                writer.WriteElementString("Left", "0.25in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("Width", "0.875in")
                writer.WriteElementString("Source", "External")
                writer.WriteElementString("Height", "0.5in")
                writer.WriteElementString("Value", "file:///C:/Inetpub/wwwroot/copy of outlined for jenn.jpg") '' point this to an IIS webserver with image for 'EXTERNAL'
                '' may be able to switch to a 'EMBEDED' Flag
                '' 
                writer.WriteEndElement() ' image1
                writer.WriteEndElement() ' </reportitems>
                writer.WriteElementString("Height", "0.75in")
                writer.WriteElementString("PrintOnLastPage", "true")
                writer.WriteElementString("PrintOnFirstPage", "true")
                writer.WriteEndElement() ' </PageHeader> 

                writer.WriteEndElement() ' report
                '' extra
                ' flush writer and close stream
                writer.Flush()
                stream.Close()

                ''
                '' THE IDEA HERE:
                '' create a class to dynamically make a report file
                '' pass the class as an object (report.rdl) to the background print class (SQL_INFO)
                '' send right to the printer.
                ''

                Main.Cursor = Cursors.Default
                'MsgBox("Report1.rdl Generated.")
            Catch ex As Exception
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateRDL")
            End Try
        End Sub
    End Class

#Region "Sales Summary Logic"


    Public Class Summary_RDL_Gen

        Implements IDisposable

        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
        Private m_CurrentPageIndex As Integer
        Private m_streams As IList(Of Stream)


        Private m_connection As SqlConnection
        Private m_connectString As String
        Private m_commandText As String = "select * from temp"
        Private m_fields As ArrayList


        Public Sub GetXML_TO_TEXT_FILE_DailyPerf()
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year
                If System.IO.File.Exists("\\ekg1\iss\Reports\DataSets\SalesSummary_" + putbacktogether.ToString + ".xml") = True Then
                    System.IO.File.Delete("\\ekg1\iss\Reports\DataSets\SalesSummary_" + putbacktogether.ToString + ".xml")
                End If
                Dim fs As New StreamWriter("\\ekg1\iss\Reports\DataSets\SalesSummary_" + putbacktogether.ToString + ".xml") '' directory needs to point to a 'iss share'
                '\\ekg1\iss\Reports\DataSets\'Generate a Name for the xml file 
                Dim cmdGET As SqlCommand = New SqlCommand("Select XML_Val from iss.dbo.InsertXMLForReporting", cnn)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGET.ExecuteReader
                fs.WriteLine("<ROOT>") ' had to add additional 'root' tag because SQL doesn't generate it for XML, PATH argument.
                While r1.Read
                    fs.WriteLine(r1.Item("XML_Val").ToString)
                End While
                r1.Close()
                cnn.Close()
                fs.WriteLine("</ROOT>")
                fs.Flush()
                fs.Close()
            Catch ex As Exception
                cnn.Close()
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetXML_TO_TEXT_FILE_DailyPerf")
            End Try

            Try
                Dim cmdDEL As SqlCommand = New SqlCommand("DELETE iss.dbo.InsertXMLForReporting", cnn)

                cnn.Open()
                cmdDEL.ExecuteNonQuery()
                cnn.Close()



            Catch ex As Exception
                cnn.Close()
            End Try
        End Sub
        Public Sub GetSQL_TO_XML_Information_DailyPerf(ByVal ApptDate As String)

            Try
                '' clean up temp table, if it exists.
                '' 
                Try
                    Dim cmdDEL As SqlCommand = New SqlCommand("Drop table iss.dbo.temp", cnn)
                    cnn.Open()
                    cmdDEL.ExecuteNonQuery()
                    cnn.Close()
                Catch ex As Exception
                    cnn.Close()
                    '' blank just to make sure the table is or has been dropped.
                    '' do not need to report this error. it is BY DESIGN
                    '' 
                End Try


                Dim cmdRUN As SqlCommand = New SqlCommand("dbo.SDPReportXML", cnn)
                Dim param1 As SqlParameter = New SqlParameter("@Date", ApptDate)
                cmdRUN.Parameters.Add(param1)
                cmdRUN.CommandType = CommandType.StoredProcedure
                cnn.Open()
                cmdRUN.ExecuteNonQuery()
                cnn.Close()
                '' now get it to a flat file
                GetXML_TO_TEXT_FILE_DailyPerf()
                'Me.Run_DailyPerf(Date.Today)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)

                cnn.Close()
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetSQL_TO_XML_Information_DailyPerf")

            End Try

        End Sub

        Public Function Generate_DATASET(ByVal FileName As String) As DataTable
            Dim dataset As New DataSet
            Try
                dataset.ReadXml(FileName)
                Return dataset.Tables(0)
            Catch ex As Exception
                Return dataset.Tables(0)
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal FileName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Generate_DATASET")
            End Try

        End Function

        Private Function CreateStream(ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean) As Stream
            Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)

            Try
                'Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)
                m_streams.Add(stream)
                Return stream
            Catch ex As Exception
                'MsgBox(ex.Message.ToString)

                Return stream
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CreateStream")
            End Try

        End Function

        'Private Sub Export(ByVal Report As LocalReport)
        '    '' this may or may not need to reset depending on device selected.
        '    '' also, output may be a selectable option, like "PDF" or "EXCEL"
        '    '' for rendering.
        '    '' Depends if client machine will accept it, AND if local rendering supports it.
        '    '' 
        '    Try
        '        Dim deviceInfo As String = _
        '        "<DeviceInfo>" + _
        '        "<OutPutFormat>EMF</OutPutFormat>" + _
        '        "<PageWidth>8.5in</PageWidth>" + _
        '        "<PageHeight>11in</PageHeight>" + _
        '        "<MarginTop>1in</MarginTop>" + _
        '        "<MarginLeft>.25in</MarginLeft>" + _
        '        "<MarginRight>.25in</MarginRight>" + _
        '        "<MarginBottom>1in</MarginBottom>" + _
        '        "</DeviceInfo>"
        '        Dim warnings() As Warning = Nothing
        '        Report.EnableExternalImages = True
        '        m_streams = New List(Of Stream)()
        '        Report.Render("Image", deviceInfo, AddressOf CreateStream, warnings)
        '        Dim stream As Stream
        '        For Each stream In m_streams
        '            stream.Position = 0
        '        Next
        '    Catch ex As Exception
        '        'MsgBox(ex.Message.ToString & " | " & ex.InnerException.ToString)

        '        Dim errp As New ErrorLogFlatFile
        '        errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Report As LocalReport", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Export")
        '    End Try

        'End Sub

        'Private Sub PrintPage(ByVal Sender As Object, ByVal ev As PrintPageEventArgs)
        '    Try
        '        Dim pageImage As New Metafile(m_streams(m_CurrentPageIndex))
        '        ev.Graphics.DrawImage(pageImage, ev.PageBounds)
        '        m_CurrentPageIndex += 1
        '        ev.HasMorePages = (m_CurrentPageIndex < m_streams.Count)
        '    Catch ex As Exception
        '        'MsgBox(ex.Message.ToString)

        '        Dim errp As New ErrorLogFlatFile
        '        errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Sender As Object, ByVal ev As PrintPageEventArgs", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "PrintPage")
        '    End Try

        'End Sub

        Private Sub print()
            'Const printername As String = _
            '" "
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                EditAutoNotesLists.Cursor = Cursors.WaitCursor
                If m_streams Is Nothing Or m_streams.Count = 0 Then
                    Return
                End If

                Dim printDoc As New PrintDocument()
                'printDoc.PrinterSettings.PrinterName = printername


                If Not printDoc.PrinterSettings.IsValid Then
                    'Dim msg As String = String.Format("Can't find printer ""{0}"".", printername)
                    'MsgBox(msg.ToString)
                    ' AddHandler printDoc.PrintPage, AddressOf PrintPage
                    Dim pr1 As New PrintDialog
                    pr1.Document = printDoc
                    pr1.Document.DefaultPageSettings.Landscape = True
                    printDoc.DefaultPageSettings.Margins.Bottom = 1
                    printDoc.DefaultPageSettings.Margins.Top = 1
                    printDoc.DefaultPageSettings.Margins.Left = 1
                    printDoc.DefaultPageSettings.Margins.Right = 1
                    printDoc.PrinterSettings.PrintToFile = True
                    printDoc.DefaultPageSettings.Landscape = True
                    printDoc.DocumentName = "Sales Performance " + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
                    'pr1.ShowDialog()
                    Dim dlg As DialogResult = pr1.ShowDialog
                    Select Case dlg
                        Case Is = DialogResult.OK
                            printDoc.Print()
                            Exit Select
                    End Select
                    'printDoc.Print()
                    'Main.Cursor = Cursors.Default
                    Return
                End If

                '' pipe it to the print preview

                ' AddHandler printDoc.PrintPage, AddressOf PrintPage

                Dim pr As New PrintDialog
                pr.Document = printDoc
                printDoc.DefaultPageSettings.Margins.Bottom = 1
                printDoc.DefaultPageSettings.Margins.Top = 1
                printDoc.DefaultPageSettings.Margins.Left = 1
                printDoc.DefaultPageSettings.Margins.Right = 1
                printDoc.PrinterSettings.PrintToFile = True
                printDoc.DefaultPageSettings.Landscape = True
                printDoc.DocumentName = "Sales Performance " + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
                'printDoc.Print()
                Dim dlg2 As DialogResult = pr.ShowDialog
                Select Case dlg2
                    '' 3 = 'Abort'
                    '' 1 = 'OK'
                    '' 2 = 'Cancel'
                    '' 5 = 'Ignore'
                    '' 7 = 'No'
                    '' 0 = 'None'
                    '' 4 = 'Retry'
                    '' 6 = 'Yes'
                    Case Is = DialogResult.OK

                        Dim cmdDROP As SqlCommand = New SqlCommand("Drop table iss.dbo.temp", cnn)
                        cnn.Open()
                        cmdDROP.ExecuteNonQuery()

                        cnn.Close()
                        printDoc.Print()
                        Exit Select
                End Select
                'Main.Cursor = Cursors.Default
            Catch ex As Exception
                'MsgBox(ex.Message.ToString)

                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Print")

            End Try

        End Sub
        Public Sub Run_DailyPerf(ByVal ApptDate As String)
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                '  Dim report As LocalReport = New LocalReport()
                If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl") = True Then
                    System.IO.File.Delete("\\ekg1\iss\reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl")
                End If

                'Dim y As New GENERATE_RDL ' No. Generate RDL will have to reflect the new query.

                '' will need a new generate RDL for daily performance.
                '' 
                'Me.GenerateRDL()
                Me.GenerateFieldsList_Summary(ApptDate)
                Me.GenerateRDL_SalesSummary()

                'y.GenerateFieldsList(ApptDate)
                'y.GenerateRDL()


                '   report.ReportPath = "\\ekg1\iss\Reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl"
                '' \\ekg1\iss\Reports\RDL Files\'Name of Generated RDL file."
                '  report.DataSources.Add(New ReportDataSource("XML", Me.Generate_DATASET("\\ekg1\iss\Reports\DataSets\SalesSummary_" + putbacktogether + ".xml")))
                '' \\ekg1\iss\Reports\DataSets\'Name of generated XML file'
                '   Export(report)
                m_CurrentPageIndex = 0
                print()
            Catch ex As Exception
                'MsgBox(ex.Message.ToString)
                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Run_DailyPerf")

            End Try

        End Sub


        Public Sub GenerateFieldsList_Summary(ByVal ApptDate As String)
            Try
                Dim cmd1 As SqlCommand = New SqlCommand("create table iss.dbo.temp (id int identity,  rep varchar (max), issued integer, demos integer, " _
    & " nodemos integer, resets integer, nothits integer, sales integer, dollars  money, noresults  integer)", cnn)
                cnn.Open()
                cmd1.ExecuteNonQuery()
                cnn.Close()

                Dim cmdGet As SqlCommand = New SqlCommand("insert into iss.dbo.temp (rep) " _
    & " select distinct rep1 from iss.dbo.leadhistory " _
    & " where rep1 is not null and Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) " _
    & " order by rep1 " _
    & " " _
    & " declare @cnt as integer " _
    & " set @cnt = (select count (id) from iss.dbo.temp) " _
    & " " _
    & " declare @id as integer " _
    & " set @id = 1 " _
    & " " _
    & " declare @rep as varchar (max) " _
    & " set @rep = (select rep from iss.dbo.temp where id = @id ) " _
    & " " _
    & " " _
    & " While @id <= @cnt " _
    & " Begin " _
    & " update iss.dbo.temp " _
    & " set issued = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp " _
    & " set resets = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'Reset' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp " _
    & " set demos = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'Demo/No Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp  " _
    & " set sales = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp " _
    & " set dollars = (select sum (QuotedSold) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp " _
    & " set nothits = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'Not Hit' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp  " _
    & " set nodemos = (select count (sresult) from iss.dbo.leadhistory " _
    & " where rep1 = @rep and sresult = 'No Demo' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where id = @id " _
    & " " _
    & " update iss.dbo.temp " _
    & " set noresults = (select count (id) from enterlead " _
    & " where rep1 = @rep and NeedsSaleResult = 'True' and  " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where id = @id " _
    & " " _
    & " set @id = @id + 1 " _
    & " set @rep = (select rep from iss.dbo.temp where id = @id) " _
    & " " _
    & " End " _
    & "  insert iss.dbo.temp(rep) " _
    & " values ('Total') " _
    & " " _
    & " update iss.dbo.temp " _
    & " set issued = (select count (sresult) from iss.dbo.leadhistory " _
    & " where Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where rep = 'Total' " _
    & " " _
    & "  update iss.dbo.temp " _
    & " set nothits = (select count (sresult) from iss.dbo.leadhistory " _
    & " where sresult = 'Not Hit' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp " _
    & " set demos = (select count (sresult) from iss.dbo.leadhistory " _
    & " where sresult = 'Demo/No Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp " _
    & " set nodemos = (select count (sresult) from iss.dbo.leadhistory " _
    & " where sresult = 'No Demo' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp " _
    & " set sales = (select count (sresult) from iss.dbo.leadhistory  " _
    & " where sresult = 'Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp " _
    & " set resets = (select count (sresult) from iss.dbo.leadhistory " _
    & " where sresult = 'Reset' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp  " _
    & " set noresults = (select count (id) from enterlead " _
    & " where  rep1 is not null and NeedsSaleResult = 'True' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) )" _
    & " where rep = 'Total' " _
    & " " _
    & " update iss.dbo.temp " _
    & " set dollars = (select sum (QuotedSold) from iss.dbo.leadhistory " _
    & " where sresult = 'Sale' and " _
    & " Convert(varchar (12), ApptDate, 101) = convert(varchar (12), @Date, 101) ) " _
    & " where rep = 'Total' " _
    & " " _
    & " select * from iss.dbo.temp ", cnn)



                Dim param1 As SqlParameter = New SqlParameter("@Date", ApptDate)
                cmdGet.Parameters.Add(param1)
                cnn.Open()
                Dim r1 As SqlDataReader
                cmdGet.CommandType = CommandType.Text
                r1 = cmdGet.ExecuteReader(CommandBehavior.SchemaOnly)
                m_fields = New ArrayList
                Dim i As Integer
                For i = 0 To r1.FieldCount - 1
                    m_fields.Add(r1.GetName(i))
                Next
                cnn.Close()



            Catch ex As Exception
                'MsgBox(ex.Message.ToString)

                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateFieldList")

            End Try

        End Sub

        Public Sub GenerateRDL_SalesSummary()
            Try
                'Main.Cursor = Cursors.WaitCursor
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl") = True Then '' point this to an iss share
                    '' \\ekg1\iss\Reports\RDL Files\'Name of RDL Generated'.rdl
                    System.IO.File.Delete("\\ekg1\iss\Reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl") '' point this to an iss share
                End If
                Dim stream As FileStream
                stream = File.OpenWrite("\\ekg1\iss\Reports\RDL Files\SalesSummary_" + putbacktogether + ".rdl") '' point this to an 'iss share'
                Dim writer As New XmlTextWriter(stream, Encoding.UTF8)
                writer.Formatting = Formatting.Indented
                'Report Element
                writer.WriteProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
                writer.WriteStartElement("Report")
                writer.WriteAttributeString("xmlns", Nothing, "http://schemas.microsoft.com/sqlserver/reporting/2003/10/reportdefinition")
                writer.WriteElementString("Width", "6in")

                'DataSource Element
                writer.WriteStartElement("DataSources")
                writer.WriteStartElement("DataSource")
                writer.WriteAttributeString("Name", Nothing, "XML")
                writer.WriteStartElement("ConnectionProperties")
                writer.WriteElementString("DataProvider", "SQL")
                writer.WriteElementString("ConnectString", cnn.ConnectionString.ToString)
                writer.WriteEndElement() ' connection properties
                writer.WriteEndElement() ' datasource
                writer.WriteEndElement() ' datasources
                'dataset element
                writer.WriteStartElement("DataSets")
                writer.WriteStartElement("DataSet")
                writer.WriteAttributeString("Name", Nothing, "XML")
                'query element
                writer.WriteStartElement("Query")
                writer.WriteElementString("DataSourceName", "XML")
                writer.WriteElementString("CommandType", "Text")
                writer.WriteElementString("CommandText", m_commandText)
                writer.WriteElementString("Timeout", "30")
                writer.WriteEndElement() ' query
                writer.WriteStartElement("Fields")
                Dim fieldName As String = ""
                For Each fieldName In m_fields
                    writer.WriteStartElement("Field")
                    writer.WriteAttributeString("Name", Nothing, fieldName)
                    writer.WriteElementString("DataField", Nothing, fieldName)
                    writer.WriteEndElement() ' Field
                Next

                writer.WriteEndElement() 'Fields
                writer.WriteEndElement() 'Dataset
                writer.WriteEndElement() 'Datasets
                'body element
                writer.WriteStartElement("Body")
                '' image properties
                '' 
                writer.WriteElementString("Height", "5in")
                ' Report Items
                writer.WriteStartElement("ReportItems")
                ' Table Element
                writer.WriteStartElement("Table")
                writer.WriteAttributeString("Name", Nothing, "Table1")
                writer.WriteElementString("DataSetName", "XML")
                writer.WriteElementString("Top", ".5in")
                writer.WriteElementString("Left", ".5in")
                writer.WriteElementString("Height", ".5in")

                Dim str As String = CType(m_fields.Count * 1.5, String)
                writer.WriteElementString("Width", str + "in")
                ' table columns
                writer.WriteStartElement("TableColumns")
                For Each fieldName In m_fields
                    writer.WriteStartElement("TableColumn")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteEndElement() ' table column 
                Next
                writer.WriteEndElement() ' table columns
                ' header row
                writer.WriteStartElement("Header")
                writer.WriteStartElement("TableRows")
                writer.WriteStartElement("TableRow")
                writer.WriteElementString("Height", ".25in")
                writer.WriteStartElement("TableCells")

                For Each fieldName In m_fields
                    writer.WriteStartElement("TableCell")
                    writer.WriteStartElement("ReportItems")
                    'textbox
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, "Header" + fieldName)
                    writer.WriteStartElement("Style")
                    writer.WriteElementString("TextDecoration", "Underline")
                    writer.WriteEndElement() ' style
                    writer.WriteElementString("Top", "0in")
                    writer.WriteElementString("Left", "0in")
                    writer.WriteElementString("Height", ".5in")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteElementString("Value", fieldName)
                    writer.WriteEndElement() 'Textbox
                    writer.WriteEndElement() 'Report Items
                    writer.WriteEndElement() ' TableCell
                Next

                writer.WriteEndElement() 'tablecells
                writer.WriteEndElement() 'tablerow
                writer.WriteEndElement() 'tablerows
                writer.WriteEndElement() 'header
                'Details Row
                '
                writer.WriteStartElement("Details")
                writer.WriteStartElement("TableRows")
                writer.WriteStartElement("TableRow")
                writer.WriteElementString("Height", ".25in")
                writer.WriteStartElement("TableCells")
                For Each fieldName In m_fields
                    writer.WriteStartElement("TableCell")
                    writer.WriteStartElement("ReportItems")

                    ' textbox
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, fieldName)
                    writer.WriteStartElement("Style")
                    writer.WriteEndElement() ' style
                    writer.WriteElementString("Top", "0in")
                    writer.WriteElementString("Left", "0in")
                    writer.WriteElementString("Height", ".5in")
                    writer.WriteElementString("Width", "1.5in")
                    writer.WriteElementString("Value", "=Fields!" + fieldName + ".Value")
                    writer.WriteElementString("HideDuplicates", "XML")
                    writer.WriteEndElement() 'TextBox
                    writer.WriteEndElement() 'Report ItemS
                    writer.WriteEndElement() ' TableCells
                Next

                writer.WriteEndElement() 'table cells
                writer.WriteEndElement() 'table row
                writer.WriteEndElement() ' table rows
                writer.WriteEndElement() ' details

                ' end details element and children 
                ' end table element 
                writer.WriteEndElement() 'table
                writer.WriteEndElement() ' report Items
                writer.WriteEndElement() ' body


                '' xtra stuff for images and the like.
                writer.WriteStartElement("PageHeader") ' write start element
                writer.WriteStartElement("ReportItems") ' write start element

                '' textbox 2
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox2")
                writer.WriteElementString("Left", "5.125in") '<Left>5.125in</Left>") ' write element string
                writer.WriteElementString("Top", "0.25in")
                'writer.WriteElementString("rd:DefaultName", "textbox2")
                writer.WriteElementString("ZIndex", "2")
                writer.WriteElementString("Width", "1.25in")
                '' style
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' </Style>
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", Date.Today.ToString)

                writer.WriteEndElement() ' </textbox> || textbox2

                '' textbox1
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox1")
                writer.WriteElementString("Left", "2.25in")
                writer.WriteElementString("Top", "0.25in")
                'writer.WriteElementString("rd:DefaultName", "textbox1")
                writer.WriteElementString("ZIndex", "1")
                writer.WriteElementString("Width", "2.25in")
                '' style
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("FontWeight", "700")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' </Style>
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", "EKG Construction Services, Inc.")
                writer.WriteEndElement() ' </Textbox> || textbox1

                '' image1
                writer.WriteStartElement("Image")
                writer.WriteAttributeString("Name", "image1")
                writer.WriteElementString("Sizing", "Fit")
                writer.WriteElementString("Left", "0.25in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("Width", "0.875in")
                writer.WriteElementString("Source", "External")
                writer.WriteElementString("Height", "0.5in")
                writer.WriteElementString("Value", "file:///C:/Inetpub/wwwroot/copy of outlined for jenn.jpg") '' point this to an IIS webserver with image for 'EXTERNAL'
                '' may be able to switch to a 'EMBEDED' Flag
                '' 
                writer.WriteEndElement() ' image1
                writer.WriteEndElement() ' </reportitems>
                writer.WriteElementString("Height", "0.75in")
                writer.WriteElementString("PrintOnLastPage", "true")
                writer.WriteElementString("PrintOnFirstPage", "true")
                writer.WriteEndElement() ' </PageHeader> 
                '' page footer
                '' 
                '' total of 'issued'
                writer.WriteStartElement("PageFooter")
                writer.WriteElementString("Height", "0.75in")
                writer.WriteElementString("PrintOnLastPage", "true")
                writer.WriteElementString("PrintOnFirstPage", "false")

                writer.WriteStartElement("ReportItems")

                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox10")
                writer.WriteElementString("Left", "5.125in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "9")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total issued ")
                writer.WriteEndElement() ' textobox10
                '' total of 'demos'
                ''
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox9")
                writer.WriteElementString("Left", "4.625in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "8")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total demos ")
                writer.WriteEndElement() ' textbox9
                '' total of 'nodemos'
                ''
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox8")
                writer.WriteElementString("Left", "4.125in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "7")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style") ' style
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total nodemos ")
                writer.WriteEndElement() ' textbox8
                '' total resets
                '' 
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox7")
                writer.WriteElementString("Left", "3.625in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "6")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total resets ")
                writer.WriteEndElement() ' textbox7
                '' total nothits
                ''
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox6")
                writer.WriteElementString("Left", "3.125in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "5")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total nothits ")
                writer.WriteEndElement() ' textbox6
                '' total sales
                '' 
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox5")
                writer.WriteElementString("Left", "3.125in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "5")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total sales ")
                writer.WriteEndElement() ' textbox6
                '' total dollars
                '' 
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox4")
                writer.WriteElementString("Left", "2.625in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "4")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style 
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total dollars ")
                writer.WriteEndElement() ' textbox5
                '' total noresults
                '' 
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox3")
                writer.WriteElementString("Left", "2.125in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "3")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style") ' style
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() 'style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", " total noresults ")
                writer.WriteEndElement() ' textbox4
                '' totals "totals"
                '' 
                writer.WriteStartElement("Textbox")
                writer.WriteAttributeString("Name", Nothing, "textbox11")
                writer.WriteElementString("Left", "1.625in")
                writer.WriteElementString("Top", "0.125in")
                writer.WriteElementString("ZIndex", "2")
                writer.WriteElementString("Width", "0.375in")
                writer.WriteStartElement("Style")
                writer.WriteElementString("PaddingLeft", "2pt")
                writer.WriteElementString("PaddingBottom", "2pt")
                writer.WriteElementString("PaddingRight", "2pt")
                writer.WriteElementString("PaddingTop", "2pt")
                writer.WriteEndElement() ' style
                writer.WriteElementString("CanGrow", "true")
                writer.WriteElementString("Height", "0.25in")
                writer.WriteElementString("Value", "Totals:")
                writer.WriteEndElement() ' textbox3


                writer.WriteEndElement() ' report items
                writer.WriteEndElement() ' page footer

                writer.WriteEndElement() ' report
                '' extra
                ' flush writer and close stream
                writer.Flush()
                stream.Close()

                ''
                '' THE IDEA HERE:
                '' create a class to dynamically make a report file
                '' pass the class as an object (report.rdl) to the background print class (SQL_INFO)
                '' send right to the printer.
                ''

                'Main.Cursor = Cursors.Default
                'MsgBox("Report1.rdl Generated.")
            Catch ex As Exception
                'MsgBox(ex.Message.ToString)

                Dim errp As New ErrorLogFlatFile
                errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateRDL")
            End Try
        End Sub


        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free managed resources when explicitly called
                End If

                ' TODO: free shared unmanaged resources
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class






#End Region
#Region "Sales Scheduled Task Logic"


    Public Class RDL_ScheduledTasks_Sales

        Implements IDisposable

        Private m_CurrentPageIndex As Integer
        Private m_streams As IList(Of Stream)

        Private cnn As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)

        Public Sub GetSQL_TO_XML_Information_Tasks()
            Try
                Dim cmdRUN As SqlCommand = New SqlCommand("dbo.PopSchedTaskXML", cnn)
                Dim strParam As String = "Sales"
                Dim param1 As SqlParameter = New SqlParameter("@Department", strParam)
                cmdRUN.Parameters.Add(param1)
                cmdRUN.CommandType = CommandType.StoredProcedure
                cnn.Open()
                cmdRUN.ExecuteNonQuery()
                cnn.Close()
                '' after this point, the xml data should be in the first row of the test.dbo.InsertXMLData table.
                '' 

                GetXML_TO_TEXT_FILE_Tasks()
                Me.Run_Tasks()

            Catch ex As Exception
                cnn.Close()
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetSQL_TO_XML_Information_NoResults")
            End Try
        End Sub

        Public Sub GetXML_TO_TEXT_FILE_Tasks()
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year
                If System.IO.File.Exists("\\ekg1\iss\Reports\DataSets\ScheduledTasks_" + putbacktogether.ToString + ".xml") = True Then
                    System.IO.File.Delete("\\ekg1\iss\Reports\DataSets\ScheduledTasks_" + putbacktogether.ToString + ".xml")
                End If
                Dim fs As New StreamWriter("\\ekg1\iss\Reports\DataSets\ScheduledTasks_" + putbacktogether.ToString + ".xml") '' directory needs to point to a 'iss share'
                '\\ekg1\iss\Reports\DataSets\'Generate a Name for the xml file 
                Dim cmdGET As SqlCommand = New SqlCommand("Select XML_Val from iss.dbo.InsertXMLForReporting", cnn)
                cnn.Open()
                Dim r1 As SqlDataReader
                r1 = cmdGET.ExecuteReader
                fs.WriteLine("<ROOT>") ' had to add additional 'root' tag because SQL doesn't generate it for XML, PATH argument.
                While r1.Read
                    fs.WriteLine(r1.Item("XML_Val").ToString)
                End While
                r1.Close()
                cnn.Close()
                fs.WriteLine("</ROOT>")
                fs.Flush()
                fs.Close()
            Catch ex As Exception
                cnn.Close()
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GetXML_TO_TEXT_FILE_NoResults")
            End Try

            Try
                Dim cmdDEL As SqlCommand = New SqlCommand("DELETE iss.dbo.InsertXMLForReporting", cnn)

                cnn.Open()
                cmdDEL.ExecuteNonQuery()
                cnn.Close()



            Catch ex As Exception
                cnn.Close()
            End Try


        End Sub

        Public Function Generate_DATASET_Tasks(ByVal FileName As String) As DataTable
            Dim dataset As New DataSet
            Try
                dataset.ReadXml(FileName)
                Return dataset.Tables(0)
            Catch ex As Exception
                Return dataset.Tables(0)
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal FileName as string", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Generate_DATASET")
            End Try

        End Function


        Private Function CreateStreamTasks(ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean) As Stream
            Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)

            Try
                'Dim stream As Stream = New FileStream(Name + "." + FileNameExtension, FileMode.Create, FileAccess.ReadWrite)
                m_streams.Add(stream)
                Return stream
            Catch ex As Exception
                Return stream
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Name As String, ByVal FileNameExtension As String, ByVal encoding As Encoding, ByVal MimeType As String, ByVal WillSeek As Boolean", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "CreateStream")
            End Try

        End Function

        'Private Sub ExportTasks(ByVal Report As LocalReport)
        '    '' this may or may not need to reset depending on device selected.
        '    '' also, output may be a selectable option, like "PDF" or "EXCEL"
        '    '' for rendering.
        '    '' Depends if client machine will accept it, AND if local rendering supports it.
        '    '' 
        '    Try
        '        Dim deviceInfo As String = _
        '        "<DeviceInfo>" + _
        '        "<OutPutFormat>EMF</OutPutFormat>" + _
        '        "<PageWidth>8.5in</PageWidth>" + _
        '        "<PageHeight>11in</PageHeight>" + _
        '        "<MarginTop>1in</MarginTop>" + _
        '        "<MarginLeft>.25in</MarginLeft>" + _
        '        "<MarginRight>.25in</MarginRight>" + _
        '        "<MarginBottom>1in</MarginBottom>" + _
        '        "</DeviceInfo>"
        '        Dim warnings() As Warning = Nothing
        '        Report.EnableExternalImages = True
        '        m_streams = New List(Of Stream)()
        '        Report.Render("Image", deviceInfo, AddressOf CreateStreamTasks, warnings)
        '        Dim stream As Stream
        '        For Each stream In m_streams
        '            stream.Position = 0
        '        Next
        '    Catch ex As Exception
        '        'Dim errp As New ErrorLogFlatFile
        '        'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Report As LocalReport", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Export")
        '    End Try

        'End Sub

        Private Sub PrintPageTasks(ByVal Sender As Object, ByVal ev As PrintPageEventArgs)
            Try
                Dim pageImage As New Metafile(m_streams(m_CurrentPageIndex))
                ev.Graphics.DrawImage(pageImage, ev.PageBounds)
                m_CurrentPageIndex += 1
                ev.HasMorePages = (m_CurrentPageIndex < m_streams.Count)


                '' clean up emf's 
                '' AFTER the print button has been called.
                '' will need another clean up call in 
                '' the select case statement as well. 
                '' 



            Catch ex As Exception
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal Sender As Object, ByVal ev As PrintPageEventArgs", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "PrintPage")
            End Try

        End Sub

        Private Sub printTasks()
            'Const printername As String = _
            '" "
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                EditAutoNotesLists.Cursor = Cursors.WaitCursor
                If m_streams Is Nothing Or m_streams.Count = 0 Then
                    Return
                End If

                Dim printDoc As New PrintDocument()
                'printDoc.PrinterSettings.PrinterName = printername


                If Not printDoc.PrinterSettings.IsValid Then
                    'Dim msg As String = String.Format("Can't find printer ""{0}"".", printername)
                    'MsgBox(msg.ToString)
                    AddHandler printDoc.PrintPage, AddressOf PrintPageTasks
                    Dim pr1 As New PrintDialog
                    pr1.Document = printDoc
                    pr1.Document.DefaultPageSettings.Landscape = True
                    printDoc.DefaultPageSettings.Margins.Bottom = 1
                    printDoc.DefaultPageSettings.Margins.Top = 1
                    printDoc.DefaultPageSettings.Margins.Left = 1
                    printDoc.DefaultPageSettings.Margins.Right = 1
                    printDoc.PrinterSettings.PrintToFile = False
                    printDoc.DefaultPageSettings.Landscape = True
                    printDoc.DocumentName = "Scheduled Tasks" + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
                    'pr1.ShowDialog()
                    Dim dlg As DialogResult = pr1.ShowDialog
                    Select Case dlg
                        Case Is = DialogResult.Cancel
                            'Dim p As FileInfo
                            'Dim dir As DirectoryInfo = New DirectoryInfo("C:\Documents and Settings\xxclayxx\Desktop\Current 7-19-2008 445 PM\New Revision\New Revision\bin\Debug")
                            'For Each p In dir.GetFiles("*.emf", SearchOption.TopDirectoryOnly)
                            '    System.IO.File.Delete(p.FullName)
                            'Next
                            '' cannot clean up here.
                            '' the stream still has the object locked.
                            ''

                            Exit Select
                        Case Is = DialogResult.OK
                            printDoc.Print()
                            Exit Select
                    End Select
                    'printDoc.Print()
                    'Main.Cursor = Cursors.Default
                    Return
                End If

                '' pipe it to the print preview

                AddHandler printDoc.PrintPage, AddressOf PrintPageTasks

                Dim pr As New PrintDialog
                pr.Document = printDoc
                printDoc.DefaultPageSettings.Margins.Bottom = 1
                printDoc.DefaultPageSettings.Margins.Top = 1
                printDoc.DefaultPageSettings.Margins.Left = 1
                printDoc.DefaultPageSettings.Margins.Right = 1
                printDoc.PrinterSettings.PrintToFile = False '' should turn this off.
                printDoc.DefaultPageSettings.Landscape = True
                printDoc.DocumentName = "Scheduled Tasks " + putbacktogether.ToString  ' name of report will be 'No Results' + date ran
                'printDoc.Print()
                Dim dlg2 As DialogResult = pr.ShowDialog
                Select Case dlg2
                    '' 3 = 'Abort'
                    '' 1 = 'OK'
                    '' 2 = 'Cancel'
                    '' 5 = 'Ignore'
                    '' 7 = 'No'
                    '' 0 = 'None'
                    '' 4 = 'Retry'
                    '' 6 = 'Yes'
                    Case Is = DialogResult.Cancel
                        Exit Select
                    Case Is = DialogResult.OK
                        printDoc.Print()

                        Exit Select
                End Select
                'Main.Cursor = Cursors.Default
            Catch ex As Exception
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "None", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Print")

            End Try

        End Sub


        Private Sub Run_Tasks()
            Try
                Dim day As String = Date.Now.Day.ToString
                Dim month As String = Date.Now.Month.ToString
                Dim year As String = Date.Now.Year.ToString
                Dim putbacktogether As String = month & "-" & day & "-" & year

                '  Dim report As LocalReport = New LocalReport()
                If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl") = True Then
                    System.IO.File.Delete("\\ekg1\iss\reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl")
                End If

                Dim y As New GENERATE_RDL
                y.GenerateFieldsList_Tasks()
                y.GenerateRDL_Tasks()


                '  report.ReportPath = "\\ekg1\iss\Reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl"
                '' \\ekg1\iss\Reports\RDL Files\'Name of Generated RDL file."
                ' report.DataSources.Add(New ReportDataSource("XML", Me.Generate_DATASET_Tasks("\\ekg1\iss\Reports\DataSets\ScheduledTasks_" + putbacktogether + ".xml")))
                '' \\ekg1\iss\Reports\DataSets\'Name of generated XML file'
                '    ExportTasks(report)
                m_CurrentPageIndex = 0
                printTasks()


                '' its not the stream object....
                '' its the actual thread handle of the program that is locking them down
                '' after the 'render' call is made....
                '' 

                ' Dim p As FileInfo
                ' Dim dir As DirectoryInfo = New DirectoryInfo("C:\Documents and Settings\xxclayxx\Desktop\Current 7-19-2008 445 PM\New Revision\New Revision\bin\Debug")
                ' For Each p In dir.GetFiles("*.emf", SearchOption.TopDirectoryOnly)
                '    System.IO.File.Delete(p.FullName)
                ' Next

            Catch ex As Exception
                'Dim errp As New ErrorLogFlatFile
                'errp.WriteLog("SALES_REPORTING_LOGIC", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "Run")

            End Try


        End Sub



        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free managed resources when explicitly called
                End If

                ' TODO: free shared unmanaged resources
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Public Class GENERATE_RDL

            ''
            '' the end result of this class will not need a message box stating that the RDL file has been generated.
            '' also, this class and the "SQL_INFO" class will roll right into one method.
            '' IE: call class -> gererate RDL as needed -> generate XML Dataset -> populate report -> Render through local render mode -> select printer -> Print.
            '' complete background printing.
            '' Notes:
            '' when appt time comes out of sql it will have to be formatted to something more legible(?) 
            '' when value of textbox1 is added on render, the time will have to be formatted properly
            '' Field names will will have to be tweaked IE: "Contact1FirstName" -> "Contact Info."
            '' If no data is present to be dropped into the dataset, must have a catch for 'Nulls'
            ''  OR a blank report is generated with header image and datetime variable.
            '' 


            Private m_connection As SqlConnection
            Private m_connectString As String
            Private m_commandText As String = "" '' needs to be reflected toward the scheduled task logic
            Private m_fields As ArrayList

            Private cnn As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)

            Public Sub GenerateFieldsList_Tasks()
                Try
                    Dim cmdGet As SqlCommand = New SqlCommand("Select  scheduledtasks.ID, Completed,Executiondate, Leadnum,Schedaction,Assignedto " _
    & " from iss.dbo.ScheduledTasks " _
    & " Join [Sapreferences] " _
    & " on ScheduledTasks.Department = [sapreferences] . Departmentally " _
    & " " _
    & " where ((Department = @Department)and((Completed = 'True') and (executiondate > getdate() - days)and (show = 'True'))) or ((Department = @Department)and(Completed = 'False')) " _
    & " order by completed desc ,Executiondate desc", cnn)
                    Dim strParam As String = "Sales"
                    Dim param1 As SqlParameter = New SqlParameter("@Department", strParam)
                    cmdGet.Parameters.Add(param1)

                    cnn.Open()
                    Dim r1 As SqlDataReader
                    r1 = cmdGet.ExecuteReader(CommandBehavior.SchemaOnly)
                    m_fields = New ArrayList
                    Dim i As Integer
                    For i = 0 To r1.FieldCount - 1
                        m_fields.Add(r1.GetName(i))
                    Next
                Catch ex As Exception
                    'Dim errp As New ErrorLogFlatFile
                    'errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateFieldList")

                End Try

            End Sub
            Public Sub GenerateRDL_Tasks()
                Try
                    'Main.Cursor = Cursors.WaitCursor
                    Dim day As String = Date.Now.Day.ToString
                    Dim month As String = Date.Now.Month.ToString
                    Dim year As String = Date.Now.Year.ToString
                    Dim putbacktogether As String = month & "-" & day & "-" & year

                    If System.IO.File.Exists("\\ekg1\iss\Reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl") = True Then '' point this to an iss share
                        '' \\ekg1\iss\Reports\RDL Files\'Name of RDL Generated'.rdl
                        System.IO.File.Delete("\\ekg1\iss\Reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl") '' point this to an iss share
                    End If
                    Dim stream As FileStream
                    stream = File.OpenWrite("\\ekg1\iss\Reports\RDL Files\ScheduledTasks_" + putbacktogether + ".rdl") '' point this to an 'iss share'
                    Dim writer As New XmlTextWriter(stream, Encoding.UTF8)
                    writer.Formatting = Formatting.Indented
                    'Report Element
                    writer.WriteProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8""")
                    writer.WriteStartElement("Report")
                    writer.WriteAttributeString("xmlns", Nothing, "http://schemas.microsoft.com/sqlserver/reporting/2003/10/reportdefinition")
                    writer.WriteElementString("Width", "6in")

                    'DataSource Element
                    writer.WriteStartElement("DataSources")
                    writer.WriteStartElement("DataSource")
                    writer.WriteAttributeString("Name", Nothing, "XML")
                    writer.WriteStartElement("ConnectionProperties")
                    writer.WriteElementString("DataProvider", "SQL")
                    writer.WriteElementString("ConnectString", cnn.ConnectionString.ToString)
                    writer.WriteEndElement() ' connection properties
                    writer.WriteEndElement() ' datasource
                    writer.WriteEndElement() ' datasources
                    'dataset element
                    writer.WriteStartElement("DataSets")
                    writer.WriteStartElement("DataSet")
                    writer.WriteAttributeString("Name", Nothing, "XML")
                    'query element
                    writer.WriteStartElement("Query")
                    writer.WriteElementString("DataSourceName", "XML")
                    writer.WriteElementString("CommandType", "Text")
                    writer.WriteElementString("CommandText", m_commandText)
                    writer.WriteElementString("Timeout", "30")
                    writer.WriteEndElement() ' query
                    writer.WriteStartElement("Fields")
                    Dim fieldName As String = ""
                    For Each fieldName In m_fields
                        writer.WriteStartElement("Field")
                        writer.WriteAttributeString("Name", Nothing, fieldName)
                        writer.WriteElementString("DataField", Nothing, fieldName)
                        writer.WriteEndElement() ' Field
                    Next

                    writer.WriteEndElement() 'Fields
                    writer.WriteEndElement() 'Dataset
                    writer.WriteEndElement() 'Datasets
                    'body element
                    writer.WriteStartElement("Body")
                    '' image properties
                    '' 
                    writer.WriteElementString("Height", "5in")
                    ' Report Items
                    writer.WriteStartElement("ReportItems")
                    ' Table Element
                    writer.WriteStartElement("Table")
                    writer.WriteAttributeString("Name", Nothing, "Table1")
                    writer.WriteElementString("DataSetName", "XML")
                    writer.WriteElementString("Top", ".5in")
                    writer.WriteElementString("Left", ".5in")
                    writer.WriteElementString("Height", ".5in")

                    Dim str As String = CType(m_fields.Count * 1.5, String)
                    writer.WriteElementString("Width", str + "in")
                    ' table columns
                    writer.WriteStartElement("TableColumns")
                    For Each fieldName In m_fields
                        writer.WriteStartElement("TableColumn")
                        writer.WriteElementString("Width", "1.5in")
                        writer.WriteEndElement() ' table column 
                    Next
                    writer.WriteEndElement() ' table columns
                    ' header row
                    writer.WriteStartElement("Header")
                    writer.WriteStartElement("TableRows")
                    writer.WriteStartElement("TableRow")
                    writer.WriteElementString("Height", ".25in")
                    writer.WriteStartElement("TableCells")

                    For Each fieldName In m_fields
                        writer.WriteStartElement("TableCell")
                        writer.WriteStartElement("ReportItems")
                        'textbox
                        writer.WriteStartElement("Textbox")
                        writer.WriteAttributeString("Name", Nothing, "Header" + fieldName)
                        writer.WriteStartElement("Style")
                        writer.WriteElementString("TextDecoration", "Underline")
                        writer.WriteEndElement() ' style
                        writer.WriteElementString("Top", "0in")
                        writer.WriteElementString("Left", "0in")
                        writer.WriteElementString("Height", ".5in")
                        writer.WriteElementString("Width", "1.5in")
                        writer.WriteElementString("Value", fieldName)
                        writer.WriteEndElement() 'Textbox
                        writer.WriteEndElement() 'Report Items
                        writer.WriteEndElement() ' TableCell
                    Next

                    writer.WriteEndElement() 'tablecells
                    writer.WriteEndElement() 'tablerow
                    writer.WriteEndElement() 'tablerows
                    writer.WriteEndElement() 'header
                    'Details Row
                    '
                    writer.WriteStartElement("Details")
                    writer.WriteStartElement("TableRows")
                    writer.WriteStartElement("TableRow")
                    writer.WriteElementString("Height", ".25in")
                    writer.WriteStartElement("TableCells")
                    For Each fieldName In m_fields
                        writer.WriteStartElement("TableCell")
                        writer.WriteStartElement("ReportItems")

                        ' textbox
                        writer.WriteStartElement("Textbox")
                        writer.WriteAttributeString("Name", Nothing, fieldName)
                        writer.WriteStartElement("Style")
                        writer.WriteEndElement() ' style
                        writer.WriteElementString("Top", "0in")
                        writer.WriteElementString("Left", "0in")
                        writer.WriteElementString("Height", ".5in")
                        writer.WriteElementString("Width", "1.5in")
                        writer.WriteElementString("Value", "=Fields!" + fieldName + ".Value")
                        writer.WriteElementString("HideDuplicates", "XML")
                        writer.WriteEndElement() 'TextBox
                        writer.WriteEndElement() 'Report ItemS
                        writer.WriteEndElement() ' TableCells
                    Next

                    writer.WriteEndElement() 'table cells
                    writer.WriteEndElement() 'table row
                    writer.WriteEndElement() ' table rows
                    writer.WriteEndElement() ' details
                    ' end details element and children 
                    ' end table element 
                    writer.WriteEndElement() 'table
                    writer.WriteEndElement() ' report Items
                    writer.WriteEndElement() ' body


                    '' xtra stuff for images and the like.
                    writer.WriteStartElement("PageHeader") ' write start element
                    writer.WriteStartElement("ReportItems") ' write start element

                    '' textbox 2
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, "textbox2")
                    writer.WriteElementString("Left", "5.125in") '<Left>5.125in</Left>") ' write element string
                    writer.WriteElementString("Top", "0.25in")
                    'writer.WriteElementString("rd:DefaultName", "textbox2")
                    writer.WriteElementString("ZIndex", "2")
                    writer.WriteElementString("Width", "1.25in")
                    '' style
                    writer.WriteStartElement("Style")
                    writer.WriteElementString("PaddingLeft", "2pt")
                    writer.WriteElementString("PaddingBottom", "2pt")
                    writer.WriteElementString("PaddingRight", "2pt")
                    writer.WriteElementString("PaddingTop", "2pt")
                    writer.WriteEndElement() ' </Style>
                    writer.WriteElementString("CanGrow", "true")
                    writer.WriteElementString("Height", "0.25in")
                    writer.WriteElementString("Value", Date.Today.ToString)

                    writer.WriteEndElement() ' </textbox> || textbox2

                    '' textbox1
                    writer.WriteStartElement("Textbox")
                    writer.WriteAttributeString("Name", Nothing, "textbox1")
                    writer.WriteElementString("Left", "2.25in")
                    writer.WriteElementString("Top", "0.25in")
                    'writer.WriteElementString("rd:DefaultName", "textbox1")
                    writer.WriteElementString("ZIndex", "1")
                    writer.WriteElementString("Width", "2.25in")
                    '' style
                    writer.WriteStartElement("Style")
                    writer.WriteElementString("PaddingLeft", "2pt")
                    writer.WriteElementString("PaddingBottom", "2pt")
                    writer.WriteElementString("FontWeight", "700")
                    writer.WriteElementString("PaddingRight", "2pt")
                    writer.WriteElementString("PaddingTop", "2pt")
                    writer.WriteEndElement() ' </Style>
                    writer.WriteElementString("CanGrow", "true")
                    writer.WriteElementString("Height", "0.25in")
                    writer.WriteElementString("Value", "EKG Construction Services, Inc.")
                    writer.WriteEndElement() ' </Textbox> || textbox1

                    '' image1
                    writer.WriteStartElement("Image")
                    writer.WriteAttributeString("Name", "image1")
                    writer.WriteElementString("Sizing", "Fit")
                    writer.WriteElementString("Left", "0.25in")
                    writer.WriteElementString("Top", "0.125in")
                    writer.WriteElementString("Width", "0.875in")
                    writer.WriteElementString("Source", "External")
                    writer.WriteElementString("Height", "0.5in")
                    writer.WriteElementString("Value", "file:///C:/Inetpub/wwwroot/copy of outlined for jenn.jpg") '' point this to an IIS webserver with image for 'EXTERNAL'
                    '' may be able to switch to a 'EMBEDED' Flag
                    '' 
                    writer.WriteEndElement() ' image1
                    writer.WriteEndElement() ' </reportitems>
                    writer.WriteElementString("Height", "0.75in")
                    writer.WriteElementString("PrintOnLastPage", "true")
                    writer.WriteElementString("PrintOnFirstPage", "true")
                    writer.WriteEndElement() ' </PageHeader> 

                    writer.WriteEndElement() ' report
                    '' extra
                    ' flush writer and close stream
                    writer.Flush()
                    stream.Close()

                    ''
                    '' THE IDEA HERE:
                    '' create a class to dynamically make a report file
                    '' pass the class as an object (report.rdl) to the background print class (SQL_INFO)
                    '' send right to the printer.
                    ''

                    'Main.Cursor = Cursors.Default
                    'MsgBox("Report1.rdl Generated.")
                Catch ex As Exception
                    'Dim errp As New ErrorLogFlatFile
                    'errp.WriteLog("SALES_REPORTING_LOGIC.GENERATE_RDL", "ByVal ApptDate As String", ex.Message.ToString, "Client", STATIC_VARIABLES.CurrentUser & ", " & STATIC_VARIABLES.CurrentForm, "Front_End", "GenerateRDL")
                End Try
            End Sub
        End Class

    End Class

#End Region
End Class
