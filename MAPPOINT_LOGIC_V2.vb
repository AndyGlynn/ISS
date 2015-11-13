Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports MapPoint


Public Class MAPPOINT_LOGIC_V2

    ''
    '' mappoint symbol reference for plotting 
    '' http://www.mapping-tools.com/howto/mappoint/locations-pushpins/pushpins-for-mappoint-2010/
    '' 

    Public Structure CityAndZip
        Public CityName As String
        Public ZipCode As String
        Public TargetState As String
    End Structure

    Public Structure ZipState
        Public ZipCode As String
        Public TargetState As String
    End Structure

    Public Structure CityState
        Public City As String
        Public TargetState As String
    End Structure

    Public Structure LeadToPlotSales
        Public ID As String
        Public Result As String
        Public ParPrice As String
        Public QuotedSold As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure LeadToPlotDemoNoSale
        Public ID As String
        Public Result As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure LeadToPlotNotHit
        Public ID As String
        Public Result As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure LeadToPlotRecissionCancel
        Public ID As String
        Public Result As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
        Public ParPrice As String
        Public QuotedSold As String
    End Structure

    Public Structure LeadToPlotReset
        Public ID As String
        Public Result As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure LeadToPlotNoDemo
        Public ID As String
        Public Result As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure SingleLeadLookup
        Public ID As String
        Public StAddress As String
        Public City As String
        Public State As String
        Public Zip As String
    End Structure

    Public Structure DriveTimeAndDistance
        Public DriveTime
        Public Distance
    End Structure

    Public arCityAndZips As List(Of CityAndZip)
    Public arCityState As List(Of CityState)
    Public arZipState As List(Of ZipState)

    Private Const sql_cnx As String = "SERVER=192.168.1.2;DATABASE=ISS;User Id=sa;Password=spoken1;"

    Public Function Begin()
        arCityAndZips = GetUniqueCityAndZips()
        Return arCityAndZips
    End Function


    '' split into two initial lists to NOT show duplicates
    '' two different DISTINCT queries
    '' 

    Public Function GetUniqueZipState()
        Dim cnx As New SqlConnection(sql_cnx)
        Dim arRes As New List(Of ZipState)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT DISTINCT(Zip),State FROM EnterLead;", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New ZipState
            g.ZipCode = r1.Item("Zip")
            g.TargetState = r1.Item("State")
            arRes.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arRes
    End Function

    Public Function GetUniqueCityState()
        Dim cnx As New SqlConnection(sql_cnx)
        Dim arRes As New List(Of CityState)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT DISTINCT(City),State FROM EnterLead;", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New CityState
            g.City = r1.Item("City")
            g.TargetState = r1.Item("State")
            arRes.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arRes
    End Function

    Private Function GetUniqueCityAndZips()
        Dim cnx_get As New SqlConnection(sql_cnx)
        Dim arCityZip As New List(Of CityAndZip)
        cnx_get.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT distinct(City),State,Zip from enterlead;", cnx_get)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim CityZip As New CityAndZip
            CityZip.CityName = r1.Item("City")
            CityZip.ZipCode = r1.Item("Zip")
            CityZip.TargetState = r1.Item("State")
            arCityZip.Add(CityZip)
        End While
        r1.Close()
        cnx_get.Close()
        cnx_get = Nothing
        Return arCityZip
    End Function

    Public Function Search_Radius_City(ByVal ListOfCityAndZips As List(Of CityState), ByVal Search_Radius As Double, ByVal StartCity As String, ByVal StartState As String)
        Dim listResults As New List(Of CityState)
        listResults = ListOfCityAndZips
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        oApp.Visible = False
        Dim oMap As MapPoint.Map = oApp.NewMap
        Dim start_loc As MapPoint.Location
        Dim start_res As MapPoint.FindResults = oMap.FindAddressResults(, StartCity, , StartState, , GeoMapRegion.geoMapNorthAmerica)
        start_loc = start_res.Item(1)
        start_loc.GoTo()

        Dim oShape As MapPoint.Shape
        oShape = oMap.Shapes.AddShape(GeoAutoShapeType.geoShapeRadius, start_loc, Search_Radius * 2, Search_Radius * 2)

        Dim oDset As MapPoint.DataSet = oApp.ActiveMap.DataSets.AddPushpinSet("SEARCHSET")

        Dim g As CityState
        For Each g In listResults
            Dim F_Res As MapPoint.FindResults
            Dim F_loc As MapPoint.Location
            F_Res = oMap.FindAddressResults(, g.City, , g.TargetState, , GeoMapRegion.geoMapNorthAmerica)
            F_loc = F_Res.Item(1)
            Dim item_P As MapPoint.Pushpin = oMap.AddPushpin(F_loc, g.City)
            item_P.MoveTo(oDset)
        Next

        oShape.Select()
        Dim oRset As MapPoint.Recordset = oDset.QueryShape(oShape)
        oRset.MoveFirst()
        Dim arCityFoundName As New ArrayList
        Do While Not oRset.EOF
            arCityFoundName.Add(oRset.Location.Name.ToString)
            oRset.MoveNext()
        Loop

        oMap.Saved = True
        oApp.Quit()
        oApp = Nothing
        Return arCityFoundName
    End Function

    Public Function Search_Radius_ZipCode(ByVal ListOfCityAndZips As List(Of ZipState), ByVal Search_Radius As Double, ByVal StartZip As String, ByVal StartState As String)
        Dim listResults As New List(Of ZipState)
        listResults = ListOfCityAndZips
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        oApp.Visible = False
        Dim oMap As MapPoint.Map = oApp.NewMap
        Dim start_loc As MapPoint.Location
        Dim start_res As MapPoint.FindResults = oMap.FindAddressResults(, , , , StartZip)
        start_loc = start_res.Item(1)
        start_loc.GoTo()

        Dim oShape As MapPoint.Shape
        oShape = oMap.Shapes.AddShape(GeoAutoShapeType.geoShapeRadius, start_loc, (Search_Radius * 2), (Search_Radius * 2))

        Dim oDset As MapPoint.DataSet = oApp.ActiveMap.DataSets.AddPushpinSet("SEARCHSET")

        Dim g As ZipState
        For Each g In listResults
            If g.ZipCode.ToString.Length <= 0 Then
                '' dont do anthing with a blank zip code
            ElseIf g.ZipCode.ToString >= 1 Then
                Dim F_Res As MapPoint.FindResults
                Dim F_loc As MapPoint.Location
                F_Res = oMap.FindAddressResults(, , , , g.ZipCode)
                If F_Res.Count <= 0 Then
                ElseIf F_Res.Count >= 1 Then
                    If F_Res IsNot Nothing Then
                        F_loc = F_Res.Item(1)
                    End If
                    If F_loc IsNot Nothing Then
                        Dim item_P As MapPoint.Pushpin = oMap.AddPushpin(F_loc, g.ZipCode)
                        item_P.Symbol = 73
                        item_P.BalloonState = GeoBalloonState.geoDisplayBalloon
                        item_P.Note = g.ZipCode
                        item_P.MoveTo(oDset)
                    End If
                End If
            End If
        Next

        oShape.Select()
        Dim oRset As MapPoint.Recordset = oDset.QueryShape(oShape)
        oRset.MoveFirst()
        Dim arZipFoundName As New ArrayList
        Do While Not oRset.EOF
            arZipFoundName.Add(oRset.Location.Name.ToString)
            oRset.MoveNext()
        Loop

        oMap.Saved = True
        oApp.Quit()
        oApp = Nothing
        Return arZipFoundName
    End Function

    Public Function Get_STATE_ForCity(ByVal City As String)
        Dim cnx_st As New SqlConnection(sql_cnx)
        cnx_st.Open()
        Dim cmdGET As New SqlCommand("SELECT Distinct(State) from enterlead WHERE City ='" & City & "';", cnx_st)
        Dim res As String = cmdGET.ExecuteScalar()
        cnx_st.Close()
        cnx_st = Nothing
        Return res
    End Function
    Public Function Get_STATE_ForZip(ByVal Zip As String)
        Dim cnx_st As New SqlConnection(sql_cnx)
        cnx_st.Open()
        Dim cmdGET As New SqlCommand("SELECT Distinct(State) from enterlead WHERE Zip ='" & Zip & "';", cnx_st)
        Dim res As String = cmdGET.ExecuteScalar()
        cnx_st.Close()
        cnx_st = Nothing
        Return res
    End Function

    Public Function Get_Unique_States()
        Dim cnx_st As New SqlConnection(sql_cnx)
        cnx_st.Open()
        Dim cmdGET As New SqlCommand("SELECT Distinct(State) FROM EnterLead;", cnx_st)
        Dim res As New ArrayList
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            res.Add(r1.Item("State"))
        End While
        r1.Close()
        cnx_st.Close()
        cnx_st = Nothing
        Return res
    End Function


#Region "frmMappoint Stuffs"
    Public Sub Map_Single_Address(ByVal StreetAddress As String, ByVal City As String, ByVal State As String, ByVal Zip As String)

        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        Dim oRes As MapPoint.FindResults
        Dim oLoc As MapPoint.Location
        Dim pushP As MapPoint.Pushpin
        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False
        'oApp.Toolbars(5).Visible = False
        oApp.PaneState = GeoPaneState.geoPaneNone
        oRes = oMap.FindAddressResults(StreetAddress, City, , State, Zip)
        If oRes.Count >= 1 Then
            oLoc = oRes.Item(1)
        End If
        If oLoc IsNot Nothing Then
            pushP = oMap.AddPushpin(oLoc, "Found Result")
            pushP.BalloonState = GeoBalloonState.geoDisplayBalloon
            pushP.Symbol = 73
            pushP.Highlight = True

        End If

        oLoc.GoTo()
        oMap.ZoomOut()

        oMap.Saved = True
        oApp.Visible = True

    End Sub

    Public Sub PointToPointDirections(ByVal BeginStAddress As String, ByVal BeginCity As String, ByVal BeginState As String, ByVal BeginZip As String, ByVal EndStAddress As String, ByVal EndCity As String, ByVal EndState As String, ByVal EndZip As String)

        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False
        oApp.PaneState = GeoPaneState.geoPaneRoutePlanner
        Dim oRes1 As MapPoint.FindResults
        Dim oRes2 As MapPoint.FindResults
        Dim oBeginLoc As MapPoint.Location
        Dim oEndLoc As MapPoint.Location
        Dim pushB As MapPoint.Pushpin
        Dim pushE As MapPoint.Pushpin
        Dim oRoute As MapPoint.Route
        oRoute = oMap.ActiveRoute
        oRoute.Clear()

        oRes1 = oMap.FindAddressResults(BeginStAddress, BeginCity, , BeginState, BeginZip)
        oBeginLoc = oRes1.Item(1)
        pushB = oMap.AddPushpin(oBeginLoc, "Start")
        pushB.Symbol = 73
        pushB.Highlight = True
        pushB.BalloonState = GeoBalloonState.geoDisplayBalloon

        oRes2 = oMap.FindAddressResults(EndStAddress, EndCity, , EndState, EndZip)
        oEndLoc = oRes2.Item(1)
        pushE = oMap.AddPushpin(oEndLoc, "Finish")
        pushE.Symbol = 73
        pushE.Highlight = True
        pushE.BalloonState = GeoBalloonState.geoDisplayBalloon


        oRoute.Waypoints.Add(oBeginLoc, "Start")
        oRoute.Waypoints.Add(oEndLoc, "Finish")
        oRoute.Calculate()

        oApp.Visible = True



    End Sub

    Public Function Get_Unique_Sales_Results()
        Dim arResults As New ArrayList
        Dim cnx_res As New SqlConnection(sql_cnx)
        cnx_res.Open()
        Dim cmdGET As New SqlCommand("SELECT DISTINCT(Result) FROM EnterLead;", cnx_res)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim str As String = r1.Item("Result")
            If str.ToString.Length > 1 Then
                arResults.Add(str)
            End If
        End While
        r1.Close()
        cnx_res.Close()
        cnx_res = Nothing
        Return arResults
    End Function

    Public Function Get_Leads_By_Result_Sales(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotSales)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT ID,ParPrice,QuotedSold,StAddress,City,State,Zip FROM EnterLead WHERE ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "' and Result = 'Sale';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New LeadToPlotSales
            g.ID = r1.Item("ID")
            g.ParPrice = r1.Item("ParPrice")
            g.QuotedSold = r1.Item("QuotedSold")
            g.Result = "Sale"
            g.City = r1.Item("City")
            g.StAddress = r1.Item("StAddress")
            g.Zip = r1.Item("Zip")
            g.State = r1.Item("State")
            arLeadsToPlot.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function Get_Leads_By_Result_DemoNoSale(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotDemoNoSale)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT ID,StAddress,City,State,Zip FROM EnterLead WHERE ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "' and Result = 'Demo/No Sale';", cnx)
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New LeadToPlotDemoNoSale
            g.ID = r1.Item("ID")
            g.Result = "Demo/No Sale"
            g.StAddress = r1.Item("StAddress")
            g.City = r1.Item("City")
            g.State = r1.Item("State")
            g.Zip = r1.Item("Zip")
            arLeadsToPlot.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function Get_Leads_By_Result_NotHit(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotNotHit)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("select ID,Result,StAddress,City,State,Zip from enterlead where Result = 'Not Hit' and ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New LeadToPlotNotHit
            g.ID = r1.Item("ID")
            g.Result = "Not Hit"
            g.StAddress = r1.Item("StAddress")
            g.City = r1.Item("City")
            g.State = r1.Item("State")
            g.Zip = r1.Item("Zip")
            arLeadsToPlot.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function Get_Leads_By_Result_RecissionCancel(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotRecissionCancel)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("select ID,ParPrice,QuotedSold,Result,StAddress,City,State,Zip from enterlead where Result = 'Recission Cancel' and ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim x As New LeadToPlotRecissionCancel
            x.ID = r1.Item("ID")
            x.Result = r1.Item("Result")
            x.StAddress = r1.Item("StAddress")
            x.City = r1.Item("City")
            x.State = r1.Item("State")
            x.Zip = r1.Item("Zip")
            arLeadsToPlot.Add(x)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function Get_Leads_By_Result_Reset(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotReset)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("select ID,Result,StAddress,City,State,Zip from enterlead where Result = 'Reset' and ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New LeadToPlotReset
            g.ID = r1.Item("ID")
            g.Result = r1.Item("Result")
            g.StAddress = r1.Item("StAddress")
            g.City = r1.Item("City")
            g.State = r1.Item("State")
            g.Zip = r1.Item("Zip")
            arLeadsToPlot.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function Get_Leads_By_Result_NoDemo(ByVal StartDate As String, ByVal EndDate As String)
        Dim arLeadsToPlot As New List(Of LeadToPlotNoDemo)
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("select ID,Result,StAddress,City,State,Zip from enterlead where Result = 'No Demo' and ApptDate > '" & StartDate & "' and ApptDate < '" & EndDate & "';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            Dim g As New LeadToPlotNoDemo
            g.ID = r1.Item("ID")
            g.Result = r1.Item("Result")
            g.StAddress = r1.Item("StAddress")
            g.City = r1.Item("City")
            g.State = r1.Item("State")
            g.Zip = r1.Item("Zip")
            arLeadsToPlot.Add(g)
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return arLeadsToPlot
    End Function

    Public Function PlotNoDemo(ByVal ListOfNoDemo As List(Of LeadToPlotNoDemo))
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotNoDemo
        For Each b In ListOfNoDemo
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, b.Result)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString
            push.Symbol = 25
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Function

    Public Sub PlotResets(ByVal ListOfResets As List(Of LeadToPlotReset))
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotReset
        For Each b In ListOfResets
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, b.Result)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString
            push.Symbol = 9
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Sub

    Public Sub PlotRecissionCancel(ByVal ListOfRecissionCancel As List(Of LeadToPlotRecissionCancel))
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotRecissionCancel
        For Each b In ListOfRecissionCancel
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, b.Result)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString & vbCrLf & "Par Price: " & b.ParPrice & vbCrLf & "Quoted / Sold: " & b.QuotedSold
            push.Symbol = 77
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Sub

    Public Sub PlotNotHot(ByVal ListOfNotHit As List(Of LeadToPlotNotHit))
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotNotHit
        For Each b In ListOfNotHit
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, b.Result)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString
            push.Symbol = 78
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Sub

    Public Sub PlotDemoNoSale(ByVal ListOfDemoNoSale As List(Of LeadToPlotDemoNoSale))
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotDemoNoSale
        For Each b In ListOfDemoNoSale
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, b.Result)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString
            push.Symbol = 120
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Sub

    Public Sub PlotSales(ByVal ListOfSales As List(Of LeadToPlotSales))

        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        Dim oMap As MapPoint.Map = oApp.NewMap
        oApp.PaneState = GeoPaneState.geoPaneNone

        oApp.Toolbars(1).Visible = False
        oApp.Toolbars(2).Visible = False
        oApp.Toolbars(3).Visible = False
        oApp.Toolbars(4).Visible = False

        Dim b As LeadToPlotSales
        For Each b In ListOfSales
            Dim oLoc As MapPoint.Location
            Dim oRes As MapPoint.FindResults
            Dim push As MapPoint.Pushpin
            oRes = oMap.FindAddressResults(b.StAddress, b.City, , b.State, b.Zip)
            oLoc = oRes.Item(1)
            oLoc.GoTo()
            oMap.ZoomOut()
            push = oMap.AddPushpin(oLoc, "Sale - " & b.QuotedSold.ToString)
            push.BalloonState = GeoBalloonState.geoDisplayBalloon
            '' or 120 = Green Check Mark
            '' this one has a money sign "$" in the symbol. .. . thought it was more appropriate.
            push.Note = "Lead Num: " & b.ID.ToString
            push.Symbol = 344
            push.Highlight = True
        Next

        oMap.Saved = True
        oApp.Visible = True

    End Sub


    Public Function LookUpLeadNumberAddress(ByVal LeadNumber As String)
        Dim x As New SingleLeadLookup
        Dim cnx As New SqlConnection(sql_cnx)
        cnx.Open()
        Dim cmdGET As New SqlCommand("SELECT StAddress,City,State,Zip FROM EnterLead where ID = '" & LeadNumber & "';", cnx)
        Dim r1 As SqlDataReader = cmdGET.ExecuteReader
        While r1.Read
            x.ID = LeadNumber
            x.StAddress = r1.Item("StAddress")
            x.City = r1.Item("City")
            x.State = r1.Item("State")
            x.Zip = r1.Item("Zip")
        End While
        r1.Close()
        cnx.Close()
        cnx = Nothing
        Return x
    End Function

    Public Function GetDistanceBetweenCities(ByVal StartCity As String, ByVal StartState As String, ByVal EndCity As String, ByVal EndState As String)
        Dim oApp As MapPoint.Application
        oApp = CreateObject("MapPoint.Application")
        oApp.Units = GeoUnits.geoMiles
        Dim oRes1 As MapPoint.FindResults
        Dim oRes2 As MapPoint.FindResults
        Dim oLoc1 As MapPoint.Location
        Dim oLoc2 As MapPoint.Location
        Dim oMap As MapPoint.Map = oApp.NewMap
        Dim oRoute As MapPoint.Route = oMap.ActiveRoute
        oRoute.Clear()
        oRes1 = oMap.FindAddressResults(, StartCity, , StartState)
        oLoc1 = oRes1.Item(1)
        oRes2 = oMap.FindAddressResults(, EndCity, , EndState)
        oLoc2 = oRes2.Item(1)
        oMap.AddPushpin(oLoc1, "Start")
        oMap.AddPushpin(oLoc2, "End")
        oRoute.Waypoints.Add(oLoc1, "Begin City")
        oRoute.Waypoints.Add(oLoc2, "End City")
        oRoute.Calculate()

        Dim dis As Double = oMap.ActiveRoute.Distance
        Dim dTime As Double = oMap.ActiveRoute.DrivingTime
        oMap.Saved = True
        oApp.Quit()
        oApp = Nothing

        Dim DD_Obj As New DriveTimeAndDistance

        DD_Obj.Distance = Math.Round(dis, 2)
        '' reference for correctly displaying drive times
        '' adapted from vb 6
        '' https://msdn.microsoft.com/en-us/library/gg663011.aspx
        DD_Obj.DriveTime = Math.Round(dTime / GeoTimeConstants.geoOneMinute, 0)

        Return DD_Obj

    End Function

#End Region



End Class
