Imports System
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class DuplicateRecord
    Public ID As String
    Public MapPointVerified As Boolean
    Public CloseMethod As String = ""
    Private cnn As SqlConnection = New System.Data.SqlClient.SqlConnection(static_variables.cnn)
    Private Sub DuplicateRecord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.pnlDuplicates.Focus()
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        '' update record needs to remember what the switch from enterlead was and continue on with that
        '' IE : SAVE or SAVEANDNEW
        '' Also, update statement needs to update all the information from [Form] enterlead
        '' to the selected selected record ID on [Table] enterlead
        '' Must overwrite information


        '' variable to determine if it is a SAVEANDNEW or a SAVE scenario
        ''
        Dim switch As String = CloseMethod

        If Me.btnUpdate.Text = "Update Record" Then
            MsgBox("You must select a record to update", MsgBoxStyle.Exclamation, "No Record Selected")
            Exit Sub
        End If

        '' Step down method to retrieve just the ID 
        '' that we are working with

        Dim STR As String = Me.btnUpdate.Text
        Dim pieces()
        pieces = Split(STR, " ", 2)
        Dim piece()
        piece = Split(pieces(1), " ", 2)
        Dim ID As String = Trim(piece(1))

        '' ID is the variable to return what ID is selected
        '' also to be used to update [Table] enterlead with the new information 
        '' if UPDATE is chosen.

        '' SP Update Enterlead is going to take a lot of variables,
        '' also will need Unique ID = @ID // ID from btnUpdate
        '' also will need @Mapped flag from mappoint // MapPoint Verified from duplicate record [Form]
        '' 

        ID = Trim(ID)
        Dim marketer As String = Trim(EnterLead.cboMarketer.Text)
        Dim pls As String = Trim(EnterLead.cboPriLead.Text)
        Dim sls As String = Trim(EnterLead.cboSecLead.Text)
        '' leadgen notes: is this field to be restamped with date the lead came through?
        '' or is this field to be stamped with the orginal lead gen date
        '' likewise, is there a marketing result that corresponds to this stating 
        '' that is was 'Regenerated' that needs to be posted somewhere on a table? 
        Dim leadgenon As Date = EnterLead.dtpLeadGen.Value.Date.ToString
        Dim c1f As String = Trim(EnterLead.txtC1FName.Text)
        Dim c1l As String = Trim(EnterLead.txtC1LName.Text)
        Dim c2f As String = Trim(EnterLead.txtC2FName.Text)
        Dim c2l As String = Trim(EnterLead.txtC2LName.Text)
        Dim YearsOwned As String = Trim(EnterLead.txtYearsOwned.Text)
        Dim HomeAge As String = Trim(EnterLead.txtYearsOwned.Text)
        Dim HomeValue As String = Trim(EnterLead.txtHomeVal.Text)
        Dim SpecialInstr As String = Trim(EnterLead.rtfSpecialIns.Text)
        Dim HousePhone As String = Trim(EnterLead.txtHousePhone.Text)
        Dim AltPhone1 As String = Trim(EnterLead.txtAltPhone1.Text)
        Dim altphone2 As String = Trim(EnterLead.txtAltPhone2.Text)
        Dim alt1type As String = Trim(EnterLead.cboAlt1Type.Text)
        Dim alt2type As String = Trim(EnterLead.cboAltPhone2Type.Text)
        Dim sta As String = Trim(EnterLead.txtStAddy.Text)
        Dim cty As String = Trim(EnterLead.txtCity.Text)
        Dim st As String = Trim(EnterLead.txtState.Text)
        Dim zp As String = Trim(EnterLead.txtZip.Text)
        Dim spokewith As String = Trim(EnterLead.cboSpokeWith.Text)
        Dim c1work As String = Trim(EnterLead.cboC1Work.Text)
        Dim c2work As String = Trim(EnterLead.cboC2Work.Text)
        '' this may need to be reworked. I cant remember the call to pull out the date from the control.
        ''
        '' split aptdate apart.
        Dim DateAr()
        Dim aptdate As String = Trim(EnterLead.dtpApptInfo.Value.ToString)
        DateAr = Split(aptdate, " ", 2)
        aptdate = DateAr(0).ToString

        Dim appday As String = Trim(EnterLead.txtApptday.Text)
        Dim appttime As String = Trim(EnterLead.txtApptTime.Text)
        Dim prod1 As String = Trim(EnterLead.cboProduct1.Text)
        Dim prod2 As String = Trim(EnterLead.cboProduct2.Text)
        Dim prod3 As String = Trim(EnterLead.cboProduct3.Text)
        Dim color As String = Trim(EnterLead.txtProdColor.Text)
        Dim qty As String = Trim(EnterLead.txtProdQTY.Text)
        Dim email As String = Trim(EnterLead.txtEmail.Text)
        Dim prodacro1 As String = Trim(EnterLead.txtP1acro.Text)
        Dim prodacro2 As String = Trim(EnterLead.txtP2acro.Text)
        Dim prodacro3 As String = Trim(EnterLead.txtP3acro.Text)
        Dim user As String = STATIC_VARIABLES.CurrentUser
        Dim MarketingManager As String = EnterLead.txtMarketingManager.Text  ''WRONG
        '' what exactly is the description in this case? // where does it come from and where does it need to go?
        ''
        Dim desc As String = ""
        Dim Mapped As Boolean = MapPointVerified

        Dim Description As String = ""

        If EnterLead.cboMarketer.Text = "" And EnterLead.cboSecLead.Text = "" Then
            Description = "Appt. Re-Generated from " & EnterLead.cboPriLead.Text & ", set up with " & EnterLead.cboSpokeWith.Text
        ElseIf EnterLead.cboMarketer.Text <> "" And EnterLead.cboSecLead.Text = "" Then
            Description = "Appt. Re-Generated from " & EnterLead.cboPriLead.Text & " by " & EnterLead.cboMarketer.Text & ", set up with " & EnterLead.cboSpokeWith.Text
        ElseIf EnterLead.cboMarketer.Text <> "" And EnterLead.cboSecLead.Text <> "" Then
            Description = "Appt. Re-Generated from " & EnterLead.cboPriLead.Text & "-" & EnterLead.cboSecLead.Text & " by " & EnterLead.cboMarketer.Text & ", set up with " & EnterLead.cboSpokeWith.Text
        ElseIf EnterLead.cboMarketer.Text = "" And EnterLead.cboSecLead.Text <> "" Then
            Description = "Appt. Re-Generated From " & EnterLead.cboPriLead.Text & "-" & EnterLead.cboSecLead.Text & ", set up with " & EnterLead.cboSpokeWith.Text
        End If

        Dim c As New ENTER_LEAD.UpdateEnterLead(ID, marketer, pls, sls, leadgenon, c1f, c1l, c2f, c2l, YearsOwned, HomeAge, HomeValue, SpecialInstr, HousePhone, AltPhone1, altphone2, _
            alt1type, alt2type, sta, cty, st, zp, spokewith, c1work, c2work, aptdate, appttime, appday, prod1, prod2, prod3, color, qty, email, prodacro1, prodacro2, prodacro3, user, MarketingManager, Description, MapPointVerified)

        Select Case CloseMethod
            Case Is = "SAVEANDNEW"
                EnterLead.Reset()
                Me.Close()
            Case Is = "SAVE"
                EnterLead.Reset()
                EnterLead.Close()
                Me.Close()
        End Select

        Me.Close()
        Me.Dispose()

    End Sub



    Private Sub btnNewRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewRecord.Click
       

        Dim c As New ENTER_LEAD.InsertEnterLead
        c.InsertLead(EnterLead.MapPoint_Verified)
        'Me.lstDupes.Items.Clear()
        Select Case CloseMethod
            Case Is = "SAVE"
                EnterLead.Reset()
                EnterLead.Close()
                Exit Select
            Case Is = "SAVEANDNEW"
                EnterLead.Reset()
                Exit Select
        End Select
        Me.Close()
    End Sub
    Public Function SingleVerify(ByVal LeadNumber As String)
        Dim cmdGET As SqlCommand = New SqlCommand("select Verified from dbo.VerifiedAddress " _
        & " Where LeadNum = '" & LeadNumber & "'", cnn)
        Dim retValue As Boolean
        Dim r1 As SqlDataReader
        r1 = cmdGET.ExecuteReader
        cnn.Open()
        While r1.Read
            retValue = r1.Item(0)
        End While
        r1.Close()
        cnn.Close()
        Return retValue
    End Function
End Class
