Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Sockets

Public Class networkOperations
    ''
    '' machines 'on system'
    '' machine name
    '' ip address
    '' users 'on system'
    '' network availability
    '' sql server availability
    '' server availability
    '' directory availability
    '' 
    '' 
    Private _CNX_STATE As Boolean = False
    Public _CHECK_CNX As SqlConnection = New SqlConnection("SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1;")
    Private _NETWORK_STATUS As Boolean
    Public MachineName As String = My.Computer.Name.ToString
    Public IPV4_ADDRESS()

    Public Function Check_Connection_Availability(ByVal CNX As SqlConnection)
        Try
            _CHECK_CNX.Open()
            If _CHECK_CNX.State = ConnectionState.Open Then
                _CNX_STATE = True
            ElseIf _CHECK_CNX.State = ConnectionState.Closed Then
                _CNX_STATE = False
            End If
        Catch ex As Exception
            MsgBox(ex.InnerException.ToString)
        End Try
        Return _CNX_STATE
    End Function

    Public Function Check_Network_Availability()
        Try
            If My.Computer.Network.IsAvailable = True Then
                _NETWORK_STATUS = True
            ElseIf My.Computer.Network.IsAvailable = False Then
                _NETWORK_STATUS = False
            End If
        Catch ex As Exception
            MsgBox(ex.InnerException.ToString)
        End Try
        Return _NETWORK_STATUS
    End Function

    Public Function GetIPV4()
        Dim address As IPHostEntry = Dns.GetHostEntry(My.Computer.Name.ToString)
        IPV4_ADDRESS = address.AddressList()
        Return IPV4_ADDRESS
    End Function


End Class
