Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql



Public Class ToggleUndoCandC
#Region "Undo Called and Cancelled"
    '' edit 8-27-2015
    ''
    ''
    '' 1) get the ID of the 'Called and Cancel' flag from [iss.dbo.leadhistory]
    ''    a) or have it supplied as an argument
    '' 2) delete the record of [id]
    '' 3) log that it was 'Undone' 
    '' 4) change enterlead back to 'unconfirmed'
    ''    ???
    '' 5) profit.
    ''
    Public Function Get_ID_OF_CandC(ByVal LeadNum As String)
        Dim res As String = ""
        Dim m_resCNX As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        m_resCNX.Open()
        Dim cmdGET As SqlCommand = New SqlCommand("SELECT max(ID) from leadhistory where LeadNum = '" & LeadNum & "' and MResult = 'Called and Cancelled';", m_resCNX)
        res = cmdGET.ExecuteScalar
        m_resCNX.Close()
        m_resCNX = Nothing

        Return res '' the id of the record to delete 

    End Function

    Public Sub Delete_CandC(ByVal RecID As String, ByVal LeadNum As String, ByVal TriggerDate As String)
        Dim rec_ID As String = RecID.ToString
        Dim m_resCNX As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        m_resCNX.Open()
        Dim cmdDEL As SqlCommand = New SqlCommand("DELETE LeadHistory WHERE ID='" & rec_ID & "';", m_resCNX)
        cmdDEL.ExecuteScalar()
        m_resCNX.Close()
        m_resCNX = Nothing
        Log_CandC_Undone(LeadNum, Date.Now.ToString)
    End Sub

    Public Sub Log_CandC_Undone(ByVal LeadNum As String, ByVal TriggerDate As String)

        ''
        '' what status does it go back to ?
        '' Confirmed, Unconfirmed, ReIssue, Reset ..... ?
        '' Ask Andy. 8-27-2015
        ''
        Dim usr As String = STATIC_VARIABLES.CurrentUser
        Dim status As String = "Unconfirmed"
        Dim m_resCNX As SqlConnection = New SqlConnection(STATIC_VARIABLES.Cnn)
        m_resCNX.Open()
        Dim cmdUP As SqlCommand = New SqlCommand("UPDATE enterlead SET MarketingResults = (Select Top 1 Mresult from LeadHistory Where Leadnum = '" & LeadNum & "' order by id desc) WHERE ID = '" & LeadNum & "';", m_resCNX)
        cmdUP.ExecuteScalar()
        m_resCNX.Close()

  
        m_resCNX = Nothing

    End Sub

#End Region
End Class
