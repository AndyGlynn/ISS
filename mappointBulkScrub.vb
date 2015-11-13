Imports MapPoint

Public Class mappointBulkScrub
    '' will be used instead of static_variables.oApp
    Public oApp As MapPoint.Application
    Public oMap As MapPoint.Map
    Public oRes As MapPoint.FindResults
    Public Sub New()
        oApp = CreateObject("MapPoint.Application")
        oMap = oApp.NewMap()
        oMap.Saved = True
        oApp.Visible = False
    End Sub

    Public Sub VerifyAddress(ByVal obj As ImportData_V2.sqlOperations.Record_And_Address, ByVal DevOrPro As Boolean)
        Try
            oRes = oMap.FindAddressResults(obj.StreetAddress, obj.City, , obj.State, obj.Zip, GeoCountry.geoCountryUnitedStates)
            ''find results quality reference: 
            '' https://msdn.microsoft.com/en-us/library/aa736235.aspx
            Select Case oRes.ResultsQuality
                'Case Is = GeoFindResultsQuality.geoAllResultsValid   DONT USE
                '    '' not used
                Case Is = GeoFindResultsQuality.geoAmbiguousResults
                    '' update table with FIRST address
                    Dim loc As MapPoint.Location = oRes.Item(1) '' Results are NON ZERO BASED INDEXES
                    Dim stAddress As String = loc.StreetAddress.Street
                    Dim City As String = loc.StreetAddress.City
                    Dim zip As String = loc.StreetAddress.PostalCode
                    Dim State As String = loc.StreetAddress.Region
                    Dim Id As String = obj.Id_FromTable
                    Dim t_name As String = "i360__Prospect__c"

                    Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    v.Id_FromTable = Id
                    v.StreetAddress = stAddress
                    v.City = City
                    v.Zip = zip
                    v.State = State
                    Dim b As New ImportData_V2.sqlOperations
                    b.Update_Table(v, DevOrPro)
                    Exit Select
                Case Is = GeoFindResultsQuality.geoFirstResultGood
                    '' update table with address
                    Dim loc As MapPoint.Location = oRes.Item(1) '' Results are NON ZERO BASED INDEXES
                    Dim stAddress As String = loc.StreetAddress.Street
                    Dim City As String = loc.StreetAddress.City
                    Dim zip As String = loc.StreetAddress.PostalCode
                    Dim State As String = loc.StreetAddress.Region
                    Dim Id As String = obj.Id_FromTable
                    Dim t_name As String = "i360__Prospect__c"

                    Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    v.Id_FromTable = Id
                    v.StreetAddress = stAddress
                    v.City = City
                    v.Zip = zip
                    v.State = State
                    Dim b As New ImportData_V2.sqlOperations
                    b.Update_Table(v, DevOrPro)
                    Exit Select
                Case Is = GeoFindResultsQuality.geoNoGoodResult
                    '' dump to proxy table

                    Dim stAddress As String = obj.StreetAddress
                    Dim City As String = obj.City
                    Dim zip As String = obj.Zip
                    Dim State As String = obj.State
                    Dim Id As String = obj.Id_FromTable
                    Dim t_name As String = "i360__Prospect__c"

                    Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    v.Id_FromTable = Id
                    v.StreetAddress = stAddress
                    v.City = City
                    v.Zip = zip
                    v.State = State
                    ' Dim b As New sqlOperations
                    'b.Dump_To_ProxyTable(v, DevOrPro)
                    Exit Select
                Case Is = GeoFindResultsQuality.geoNoResults
                    '' dump to proxy table 
                    Dim stAddress As String = obj.StreetAddress
                    Dim City As String = obj.City
                    Dim zip As String = obj.Zip
                    Dim State As String = obj.State
                    Dim Id As String = obj.Id_FromTable
                    Dim t_name As String = "i360__Prospect__c"

                    Dim v As New ImportData_V2.sqlOperations.Record_And_Address
                    v.Id_FromTable = Id
                    v.StreetAddress = stAddress
                    v.City = City
                    v.Zip = zip
                    v.State = State
                    ' Dim b As New sqlOperations
                    ' b.Dump_To_ProxyTable(v, DevOrPro)
                    Exit Select
            End Select
        Catch ex As Exception
            Dim stAddress As String = obj.StreetAddress
            Dim City As String = obj.City
            Dim zip As String = obj.Zip
            Dim State As String = obj.State
            Dim Id As String = obj.Id_FromTable
            Dim t_name As String = "i360__Prospect__c"

            Dim v As New ImportData_V2.sqlOperations.Record_And_Address
            v.Id_FromTable = Id
            v.StreetAddress = stAddress
            v.City = City
            v.Zip = zip
            v.State = State
            'Dim b As New sqlOperations
            'b.Dump_To_ProxyTable(v, DevOrPro)
        End Try
    End Sub
End Class
