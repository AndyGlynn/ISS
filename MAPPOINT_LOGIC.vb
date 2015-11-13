
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System
Imports MapPoint
Imports System.IO
Imports System.Text
Imports System.Diagnostics.Process
Imports System.Diagnostics


Public Class MAPPOINT_LOGIC
    Implements IDisposable
    '' mappoint variables
    'Public oApp As MapPoint.Application = New MapPoint.Application
    Public oMap As MapPoint.Map = oApp.ActiveMap
    Public oResults As MapPoint.FindResults
    Public Mappoint_Dataset As MapPoint.DataSet
    '' connection strings 
    Private AccCNN As OleDb.OleDbConnection = New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source='\\Ekg1\database\copy of database --full\EKG.mdb'")
    Private CNN As SqlConnection = New sqlconnection(STATIC_VARIABLES.cnn)
    Private ProgressMaxCount As Integer = 0
    Public Property ProgressMax() As Integer
        Get
            Return ProgressMaxCount
        End Get
        Set(ByVal value As Integer)
            ProgressMaxCount = value
        End Set
    End Property

    Public Sub DoIt(ByVal radius As Double, ByVal StartingZip As String)

        Dim cmdGET As SqlCommand = New SqlCommand("select distinct(zip) from iss.dbo.enterlead", CNN)
        Dim cmdCNT As SqlCommand = New SqlCommand("select Count(distinct(zip)) from iss.dbo.enterlead ", CNN)
        CNN.Open()
        Dim r As SqlDataReader = cmdCNT.ExecuteReader(CommandBehavior.SingleResult)
        r.Read()
        Dim cnt As Integer = r.Item(0)
        r.Close()
        Progress.ProgressBar1.Maximum = cnt * 2 + 3
        Progress.ProgressBar1.Step = 1
        WCaller.pbRadiusSearch.Maximum = cnt * 2 + 3
        WCaller.pbRadiusSearch.Step = 1
        r.Close()
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
        Dim dset_Name As Object = "TEST"
        'If oMap.DataSets.Count > 0 And oMap.DataSets.Item(dset_Name) IsNot Nothing Then
        '    oMap.DataSets(1).Delete()
        'End If
        Dim dataset As MapPoint.DataSet = oMap.DataSets.AddPushpinSet("TEST")
        Dim i As Integer = 0
        While r1.Read
            Dim LOC As MapPoint.Location
            oResults = oMap.FindAddressResults(, , , , r1.Item(0))
            If oResults IsNot Nothing Then
                LOC = oResults.Item(1)
            End If
            If LOC IsNot Nothing Then
                i += 1
                Dim push As MapPoint.Pushpin = oMap.AddPushpin(LOC, "Found")
                push.Note = r1.Item(0)
                push.MoveTo(dataset)
            End If
            Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
            WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        End While
        'Me.ProgressMax = i

        r1.Close()
        CNN.Close()
        Dim Office As MapPoint.Location
        oResults = oMap.FindAddressResults(, , , , StartingZip, GeoCountry.geoCountryUnitedStates)
        Office = oResults.Item(1)
        Office.GoTo()
        Dim shape As MapPoint.Shape = oMap.Shapes.AddShape(GeoAutoShapeType.geoShapeRadius, Office, radius, radius)
        Dim recordset As MapPoint.Recordset = dataset.QueryCircle(Office, radius)
        recordset.MoveFirst()
        'Dim sw As StreamWriter = New StreamWriter("C:\Zips.txt")
        'Form1.ListView1.Items.Clear()
        'Form1.ListView1.Columns(0).Text = "Zip Codes"
        WCaller.lbZipCity.Items.Clear()

        While recordset.EOF = False
            Dim lv As New ListViewItem
            lv.Text = recordset.Pushpin.Note.ToString()
            WCaller.lbZipCity.Items.Add(lv.Text)
            '    Form1.ListView1.Items.Add(lv)
            recordset.MoveNext()
            Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
            WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        End While

        My.Application.DoEvents()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        'sw.Close()
        'oMap.Saved = True
        'oApp.Quit()
        'Form1.btnBegin.Enabled = True
        KillProcess()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        Me.Dispose()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1


        Progress.Close()
        My.Application.DoEvents()
        Progress.Dispose()
    End Sub
    Public Sub DoItCity(ByVal radius As Double, ByVal City As String, ByVal StatePulled As String)
        WCaller.pbRadiusSearch.Maximum = 100
        Dim cmdGET As SqlCommand = New SqlCommand("select city,state from iss.dbo.enterlead group by city,state", CNN)
        Dim cmdCNT As SqlCommand = New SqlCommand("select Count(city) from iss.dbo.enterlead group by city, state", CNN)
        CNN.Open()
        Dim r As SqlDataReader = cmdCNT.ExecuteReader(CommandBehavior.SingleResult)
        r.Read()
        Dim cnt As Integer = r.Item(0)
        r.Close()
        Progress.ProgressBar1.Maximum = cnt * 2 + 3
        Progress.ProgressBar1.Step = 1
        WCaller.pbRadiusSearch.Maximum = cnt * 2 + 3
        WCaller.pbRadiusSearch.Step = 1
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader(CommandBehavior.CloseConnection)
        Dim dset_Name As Object = "TEST"
        'If oMap.DataSets.Count > 0 And oMap.DataSets.Item(dset_Name) IsNot Nothing Then
        '    oMap.DataSets(1).Delete()
        'End If
        Dim dataset As MapPoint.DataSet = oMap.DataSets.AddPushpinSet("TEST")
        Dim ArLocations As New ArrayList
        Dim p As Integer = 0
        '' doesnt pull city 'fayette' ?
        '' was pointing to wrong table.
        '' reflected to point at iss.dbo.tblenterlead with a group by clause on city,state
        ''
        Dim arCity As New ArrayList
        Dim arST As New ArrayList
        While r1.Read
            Dim LOC As MapPoint.Location
            Dim cty As String = ""
            Dim state As String = ""
            cty = r1.Item(0).ToString
            state = r1.Item(1).ToString
            oResults = oMap.FindAddressResults(, cty, , state, ) '' had to fix this code, looking in the postal code region still
            arCity.Add(cty)
            arST.Add(state)

            If oResults IsNot Nothing Then
                '' fails as invalid index item
                '' "requested member of collection doesnt exist. please use valid index of item."
                '' Need a way to get this value out of oresults as string that is human readable.
                '' NOTES:
                '' oresults by itself should return an array of objects (collection)
                '' need to break this code apart and remember how it works....
                '' 
                '' test scenario swanton ohio is 20.1 miles according to mappoint front end.
                '' going to manually add swanton to the table to see if this code still fails,
                '' and I am going to set zip code radius to 21 miles out.
                ''
                '' also need to add a kill process method here or on the object destructor
                '' to clean up mappoint in background.
                '' swanton is already in the list..
                '' no locations found in test scenario....
                '' 4/3/08 Tested 5,15,22 mile radius searches.
                '' WORKS - had to modify code. Slower than piss, may need a progress bar for larger radius searches.
                '' Problem not returning all cities or zips in radius searches.
                '' IE 120 Mile radius search returns roughly 20 cities or zip each time.
                '' Need to find out why. 
                '' 
                '' This Codd does work. It does what it was designed to do.
                '' The requirements are mappoint and a List of either cities or Zip codes to be verified against.
                '' First you have to supply the method with a starting radius
                '' second supply mappoint's dataset with a list of cities or zips to verify against (YOUR data)
                '' in this instance it will be a list of cities supplied from a SQL store (iss.dbo.citypull or iss.dbo.tblZips)
                '' third draw a shape on the map
                '' then have mappoint find out wether the city or zip is within the shape.
                '' --------
                '' The assumed error was that you can pull out map data (IE cities or zips) from mappoint
                '' it does not work like this. You must have a list to begin with.
                '' mappoint doesn't expose its list of cities or zips.
                '' the basic logic is "Does my city/zip fall within this shape?"
                '' Not, "Give me all cities or zips within this shape."
                '' 
                Try
                    LOC = oResults.Item(1)
                    ArLocations.Add(LOC.Name)
                Catch ex As Exception
                    LOC = Nothing
                    MsgBox(ex.ToString)
                End Try
                'Dim c
                'For Each c In oResults
                '    ArLocations.Add(c.ToString)
                'Next
            End If
            If LOC IsNot Nothing Then
                p += 1
                Dim push As MapPoint.Pushpin = oMap.AddPushpin(LOC, "Found")

                push.Note = r1.Item(0)
                push.MoveTo(dataset)
            End If
            Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
            WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        End While

        My.Application.DoEvents()
        r1.Close()
        CNN.Close()
        Dim Office As MapPoint.Location
        oResults = oMap.FindAddressResults(, City, , StatePulled, , GeoCountry.geoCountryUnitedStates)
        Office = oResults.Item(1)
        Office.GoTo()
        Dim r2 As Double = (radius * 2)
        Dim shape As MapPoint.Shape = oMap.Shapes.AddShape(GeoAutoShapeType.geoShapeRadius, Office, radius, radius)
        Dim recordset As MapPoint.Recordset = dataset.QueryCircle(Office, radius)
        recordset.MoveFirst()
        'Dim sw As StreamWriter = New StreamWriter("C:\City.txt")
        WCaller.lbZipCity.Items.Clear()
        'Form1.ListView1.Items.Clear()
        ' Form1.ListView1.Columns(0).Text = "Cities"

        My.Application.DoEvents()
        While recordset.EOF = False
            Dim lv As New ListViewItem
            lv.Text = recordset.Pushpin.Note.ToString
            WCaller.lbZipCity.Items.Add(lv.Text)
            'Form1.ListView1.Items.Add(lv)
            recordset.MoveNext()
            Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
            WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        End While

        My.Application.DoEvents()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        'sw.Close()
        'oMap.Saved = True
        'Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        'oApp.Quit()
        'Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        'Form1.btnBegin.Enabled = True
        KillProcess()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1
        Me.Dispose()
        Progress.ProgressBar1.Value = Progress.ProgressBar1.Value + 1
        WCaller.pbRadiusSearch.Value = WCaller.pbRadiusSearch.Value + 1



        Progress.Close()
        My.Application.DoEvents()
        Progress.Dispose()
        '
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
    Public Sub KillProcess()
        Dim pr() = System.Diagnostics.Process.GetProcessesByName("MapPoint")
        Dim g
        For Each g In pr
            Dim b As System.Diagnostics.Process = g
            If b.HasExited = False Then
                b.Kill()
            End If
        Next
        WCaller.Cursor = Cursors.Default
        WCaller.pbRadiusSearch.Visible = False
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
