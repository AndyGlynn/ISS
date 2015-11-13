Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text


Public Class ImportExportAdmin

    ''
    '' two databases
    '' two cnx strings
    '' two structures
    '' 
    '' fields don't match up
    '' need to correlate
    '' 

    Private Const ISS_CNX As String = "SERVER=192.168.1.2;Database=ISS;User Id=sa;Password=spoken1;"
    Private Const I360_CNX As String = "SERVER=192.168.1.2;Database=I360_Data_Dump;User Id=sa;Password=spoken1;"

    Public Structure i360Appt
        Public ID As String
        Public OwnerID As String
        Public IsDeleted As String
        Public Name As String        '' when data was exported, commas in name separation threw off column ordering of data so reads like "Nelson Daryl Roof" need to split apart
        Public CreatedDate As String '' when data was exported, commas in name separation threw off column ordering of data so reads like "Nelson Daryl Roof" need to split apart
        Public CreatedById As String
        Public LastModifiedData As String
        Public LastModifiedById As String
        Public SystemModStamp As String
        Public LastActivityDate As String
        Public i360__Address__c As String
        Public i360__Appt_Set_By__c As String
        Public i360__Appt_Set_On__c As String
        Public i360__Cancelation_Email_Sent_Date__c As String
        Public i360__City__c As String
        Public i360__Comments__c As String
        Public i360__Components_1__c As String
        Public i360__Components_2__c As String
        Public i360__Components_3__c As String
        Public i360__Computed_End_DateTime__c As String
        Public i360__Computed_Start_DateTime__c As String
        Public i360__Confirmation_Email_Sent_Date__c As String
        Public i360__Confirmed_By__c As String
        Public i360__Confirmed_On__c As String
        Public i360__Confirmed__c As String
        Public i360__Data_Migration_ID__c As String
        Public i360__Disregard_In_Statistics__c As String
        Public i360__Duration__c As String
        Public i360__Email_Address__c As String
        Public i360__End_Date__c As String
        Public i360__End__c As String
        Public i360__Geocode_Source__c As String
        Public i360__Interests_Summary__c As String
        Public i360__Issued_By__c As String
        Public i360__Issued_On__c As String
        Public i360__Latitude__c As String
        Public i360__Lead_Source_Id__c As String
        Public i360__Lead_Source__c As String
        Public i360__Legacy_ID__c As String
        Public i360__Longitude__c As String
        Public i360__Market_Segment2__c As String
        Public i360__Multi_Component_Sale__c As String
        Public i360__Previous_Appointment__c As String
        Public i360__Price_Given_1__c As String
        Public i360__Price_Given_2__c As String
        Public i360__Price_Given_3__c As String
        Public i360__Prospect_Id__c As String
        Public i360__Prospect__c As String
        Public i360__Quoted_Amount__c As String
        Public i360__Renovate_Right_Pamphlet_Given__c As String
        Public i360__Rep_Split_1__c As String
        Public i360__Rep_Split_2__c As String
        Public i360__Rep_Split_3__c As String
        Public i360__Result_1__c As String
        Public i360__Result_2__c As String
        Public i360__Result_3__c As String
        Public i360__Result_Detail_1__c As String
        Public i360__Result_Detail_2__c As String
        Public i360__Result_Detail_3__c As String
        Public i360__Result_Detail__c As String
        Public i360__Result__c As String
        Public i360__Revision_Number__c As String
        Public i360__Sales_Rep_1__c As String
        Public i360__Sales_Rep_2__c As String
        Public i360__Send_Email_On_Cancelation__c As String
        Public i360__Send_Email_On_Setup__c As String
        Public i360__Sent_Email_To_Sales_Rep__c As String
        Public i360__Source_Type__c As String
        Public i360__Source__c As String
        Public i360__Start_Time__c As String
        Public i360__Start__c As String
        Public i360__State__c As String
        Public i360__Status__c As String
        Public i360__Talked_To__c As String
        Public i360__Time_Block__c As String
        Public i360__Type__c As String
        Public i360__Vendor__c As String
        Public i360__Year_Home_Built__c As String
        Public i360__Zip__c As String
        Public i360__Pending_Action__c As String
        Public i360__Send_Thank_You_Email__c As String
        Public i360__Thank_you_Email_Sent_Date__c As String
        Public i360__External_System_Id__c As String
        Public i360__Google_Calendar_Event_2_ID__c As String
        Public i360__Google_Calendar_Event_ID__c As String
        Public i360__Google_Calendar_Event_Sequence_1__c As String
        Public i360__Google_Calendar_Event_Sequence_2__c As String
        Public i360__County__c As String




    End Structure

    Public Function GetApptCnt()

    End Function

End Class
